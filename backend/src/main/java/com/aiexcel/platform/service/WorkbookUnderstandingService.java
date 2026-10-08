package com.aiexcel.platform.service;

import com.aiexcel.platform.dto.*;
import com.aiexcel.platform.entity.User;
import com.aiexcel.platform.exception.WorkbookUnderstandingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import lombok.extern.slf4j.Slf4j;
import org.springframework.ai.chat.model.ChatModel;
import org.springframework.ai.chat.model.ChatResponse;
import org.springframework.ai.chat.prompt.Prompt;
import org.springframework.ai.converter.BeanOutputConverter;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.stereotype.Service;
import org.springframework.web.client.HttpStatusCodeException;
import org.springframework.ai.chat.messages.SystemMessage;
import org.springframework.ai.chat.messages.UserMessage;

import java.util.regex.Pattern;

/**
 * Generates an AI-powered business understanding of an uploaded Excel workbook.
 * Sends only structured metadata (from ExcelAnalyzerService) to the AI — never
 * the raw file.
 */
@Service
@Slf4j
public class WorkbookUnderstandingService {

    private static final String SYSTEM_PROMPT = """
            You are a concise data analyst. Receive structured workbook metadata and
            return a single, compact JSON object containing only the requested fields.

            RULES:
            - Return EXACTLY one valid JSON object and nothing else.
            - No prose, no markdown, no commentary, no code fences.
            - Use double-quoted strings for all textual values.
            - Follow the requested schema precisely; keep values short and focused.
            """;

    private static final Pattern JSON_FENCE_PATTERN = Pattern.compile("```(?:json)?\\s*(.*?)```",
            Pattern.DOTALL | Pattern.CASE_INSENSITIVE);

    private final ExcelAnalyzerService excelAnalyzerService;
    private final ChatModel chatModel;
    private final ObjectMapper objectMapper = new ObjectMapper();

    public WorkbookUnderstandingService(
            ExcelAnalyzerService excelAnalyzerService,
            @Autowired(required = false) ChatModel chatModel) {
        this.excelAnalyzerService = excelAnalyzerService;
        this.chatModel = chatModel;
    }

    /**
     * Generates an AI understanding of the workbook.
     *
     * @param fileId ID of the uploaded file
     * @param user   authenticated user (must own the file)
     * @return AI-generated business understanding
     */
    public WorkbookUnderstandingResponse understandWorkbook(Long fileId, User user) {
        if (chatModel == null) {
            throw new WorkbookUnderstandingException(
                    "AI understanding is not configured. Set spring.ai.openai.api-key to enable this feature.");
        }

        log.info("Starting workbook understanding for fileId={}, userId={}", fileId, user.getId());

        WorkbookAnalysisResponse analysis = excelAnalyzerService.analyzeWorkbook(fileId, user);
        BeanOutputConverter<WorkbookUnderstandingResponse> converter = new BeanOutputConverter<>(
                WorkbookUnderstandingResponse.class);

        try {
            Prompt primaryPrompt = buildStrictPrompt(analysis, converter);
            String rawContent = callModel(primaryPrompt);
            String sanitizedJson = sanitizeAndNormalizeJson(rawContent, analysis.getFileId(), analysis.getFileName());
            log.debug("Sanitized JSON for fileId={} length={}", fileId, sanitizedJson.length());

            try {
                return converter.convert(sanitizedJson);
            } catch (Exception conversionFailure) {
                log.error("Primary JSON conversion failed for fileId={} class={} message={}",
                        fileId, conversionFailure.getClass().getName(), safeLogMessage(conversionFailure.getMessage()),
                        conversionFailure);
                log.debug("Sanitized JSON that failed conversion: {}", sanitizedJson);
                Prompt retryPrompt = buildRecoveryPrompt(analysis, converter);
                String retryContent = callModel(retryPrompt);
                String retryJson = sanitizeAndNormalizeJson(retryContent, analysis.getFileId(), analysis.getFileName());
                try {
                    return converter.convert(retryJson);
                } catch (Exception retryFailure) {
                    log.error("Retry JSON conversion also failed for fileId={} class={} message={}",
                            fileId, retryFailure.getClass().getName(), safeLogMessage(retryFailure.getMessage()),
                            retryFailure);
                    log.debug("Retry sanitized JSON that failed conversion: {}", retryJson);
                    throw retryFailure;
                }
            }

        } catch (WorkbookUnderstandingException e) {
            throw e;
        } catch (Exception e) {
            HttpStatusCodeException httpEx = extractHttpStatusException(e);
            if (httpEx != null) {
                log.error("AI provider rejected workbook understanding request for fileId={} status={} body={}",
                        fileId, httpEx.getStatusCode(), safeLogMessage(httpEx.getResponseBodyAsString()), httpEx);
                if (HttpStatus.PAYMENT_REQUIRED.equals(httpEx.getStatusCode())
                        || httpEx.getStatusCode().value() == 402) {
                    throw new WorkbookUnderstandingException(
                            "AI provider rejected the request because the OpenRouter budget is exhausted.", e);
                }
                if (HttpStatus.UNAUTHORIZED.equals(httpEx.getStatusCode()) || httpEx.getStatusCode().value() == 401) {
                    throw new WorkbookUnderstandingException(
                            "AI provider rejected the request because the OpenRouter API key is invalid or expired.",
                            e);
                }
                throw new WorkbookUnderstandingException(
                        "AI provider error while generating workbook understanding: " + httpEx.getStatusText(), e);
            }
            log.error("AI call failed for fileId={} class={} message={} cause={}",
                    fileId, e.getClass().getName(), safeLogMessage(e.getMessage()), e.getCause(), e);
            throw new WorkbookUnderstandingException("AI failed to generate workbook understanding", e);
        }
    }

    private String callModel(Prompt prompt) {
        log.debug("Workbook understanding AI request started");
        try {
            ChatResponse response = chatModel.call(prompt);
            String rawContent = response.getResult().getOutput().getText();
            int length = rawContent == null ? 0 : rawContent.length();
            log.debug("Raw AI response length={} characters", length);
            if (rawContent != null && length > 0) {
                int preview = Math.min(1000, length);
                log.debug("Raw AI response preview (first {} chars): {}", preview, rawContent.substring(0, preview));
            }

            // Detect provider-level plain-text errors that may be returned inside a 200
            if (rawContent != null && (rawContent.contains("requires more credits")
                    || rawContent.contains("can only afford")
                    || rawContent.toLowerCase().contains("payment required"))) {
                log.error("AI provider returned an in-band credit/error message: {}", safeLogMessage(rawContent));
                throw new WorkbookUnderstandingException(
                        "AI provider rejected the request because the OpenRouter budget is exhausted.");
            }

            return rawContent;
        } catch (org.springframework.web.client.HttpStatusCodeException httpEx) {
            // Explicitly log the provider response body for diagnostics
            String body = "<empty>";
            try {
                body = safeLogMessage(httpEx.getResponseBodyAsString());
            } catch (Exception ignore) {
            }
            log.error("AI provider HTTP error: status={} body={}", httpEx.getStatusCode(), body, httpEx);
            if (httpEx.getStatusCode().value() == 402
                    || org.springframework.http.HttpStatus.PAYMENT_REQUIRED.equals(httpEx.getStatusCode())) {
                throw new WorkbookUnderstandingException(
                        "AI provider rejected the request because the OpenRouter budget is exhausted.");
            }
            if (httpEx.getStatusCode().value() == 401
                    || org.springframework.http.HttpStatus.UNAUTHORIZED.equals(httpEx.getStatusCode())) {
                throw new WorkbookUnderstandingException(
                        "AI provider rejected the request because the OpenRouter API key is invalid or expired.");
            }
            throw httpEx;
        }
    }

    private Prompt buildStrictPrompt(WorkbookAnalysisResponse analysis,
            BeanOutputConverter<WorkbookUnderstandingResponse> converter) {
        SystemMessage systemMessage = new SystemMessage(SYSTEM_PROMPT);
        UserMessage userMessage = new UserMessage(buildPrompt(analysis) + "\n\n" + converter.getFormat());
        return new Prompt(java.util.List.of(systemMessage, userMessage));
    }

    private Prompt buildRecoveryPrompt(WorkbookAnalysisResponse analysis,
            BeanOutputConverter<WorkbookUnderstandingResponse> converter) {
        String recoverySystemPrompt = "Return ONLY a single valid JSON object for the workbook understanding. " +
                "No markdown fences, no commentary, no leading/trailing text. " +
                "The object must match this schema exactly: \n\n" + converter.getFormat();
        SystemMessage systemMessage = new SystemMessage(recoverySystemPrompt);
        UserMessage userMessage = new UserMessage(buildPrompt(analysis));
        return new Prompt(java.util.List.of(systemMessage, userMessage));
    }

    /**
     * Robust JSON sanitization and normalization.
     * Handles:
     * "Primary JSON conversion failed for fileId={} Reason={}",
     * fileId, safeLogMessage(conversionFailure.getMessage()));
     * log.debug("Sanitized JSON that failed conversion: {}", sanitizedJson);
     * // If the sanitized JSON was empty or very short, prefer not to retry against
     * non-transient errors
     * String retryPrompt = buildRecoveryPrompt(analysis, converter);
     * String retryContent = callModel(retryPrompt);
     * String retryJson = sanitizeAndNormalizeJson(retryContent,
     * analysis.getFileId(), analysis.getFileName());
     * return converter.convert(retryJson);
     * }
     * } catch (org.springframework.web.client.HttpStatusCodeException httpEx) {
     * // bubble through with clearer message already logged in callModel
     * throw httpEx;
     * if (rawText == null || rawText.isBlank()) {
     * throw new WorkbookUnderstandingException("Model returned empty response");
     * }
     * 
     * String jsonCandidate = extractJsonObjectCandidate(rawText);
     * log.debug("JSON candidate after extraction for fileId={}: {}", fileId,
     * jsonCandidate);
     * 
     * String normalized = normalizeUnquotedTokens(jsonCandidate);
     * log.debug("JSON after normalization for fileId={}: {}", fileId, normalized);
     * 
     * try {
     * JsonNode rootNode = objectMapper.readTree(normalized);
     * if (!rootNode.isObject()) {
     * throw new WorkbookUnderstandingException("Model response is not a JSON
     * object");
     * }
     * 
     * ObjectNode objectNode = (ObjectNode) rootNode;
     * if (fileId != null) {
     * objectNode.put("fileId", String.valueOf(fileId));
     * }
     * if (fileName != null) {
     * objectNode.put("fileName", fileName);
     * }
     * 
     * return objectMapper.writeValueAsString(objectNode);
     * } catch (Exception e) {
     * log.error("AI workbook understanding JSON parsing failed for fileId={} after
     * extraction. " +
     * "Normalized JSON was: {} Error: {}",
     * fileId, normalized, safeLogMessage(e.getMessage()), e);
     * throw new WorkbookUnderstandingException("Model returned malformed JSON for
     * workbook understanding");
     * }
     * }
     * 
     * /**
     * Extracts the JSON object portion from a model response that may include
     * markdown fences,
     * explanatory text, or surrounding commentary.
     */
    private String extractJsonObjectCandidate(String rawText) {
        String candidate = rawText.trim();

        if (candidate.startsWith("```")) {
            java.util.regex.Matcher matcher = JSON_FENCE_PATTERN.matcher(candidate);
            if (matcher.find()) {
                candidate = matcher.group(1).trim();
            } else {
                int endFence = candidate.indexOf("```", 3);
                if (endFence > 3) {
                    candidate = candidate.substring(3, endFence).trim();
                }
            }
        }

        int firstBrace = candidate.indexOf('{');
        int lastBrace = candidate.lastIndexOf('}');
        if (firstBrace < 0 || lastBrace <= firstBrace) {
            throw new WorkbookUnderstandingException("Model response does not contain a valid JSON object");
        }

        candidate = candidate.substring(firstBrace, lastBrace + 1);
        return candidate;
    }

    private String sanitizeAndNormalizeJson(String rawText, Long fileId, String fileName) {
        if (rawText == null || rawText.isBlank()) {
            throw new WorkbookUnderstandingException("Model returned empty response");
        }

        String jsonCandidate = extractJsonObjectCandidate(rawText);
        log.debug("JSON candidate after extraction for fileId={}: {}", fileId, jsonCandidate);

        String normalized = normalizeUnquotedTokens(jsonCandidate);
        log.debug("JSON after normalization for fileId={}: {}", fileId, normalized);

        try {
            JsonNode rootNode = objectMapper.readTree(normalized);
            if (!rootNode.isObject()) {
                throw new WorkbookUnderstandingException("Model response is not a JSON object");
            }

            ObjectNode objectNode = (ObjectNode) rootNode;
            if (fileId != null) {
                objectNode.put("fileId", String.valueOf(fileId));
            }
            if (fileName != null) {
                objectNode.put("fileName", fileName);
            }

            return objectMapper.writeValueAsString(objectNode);
        } catch (Exception e) {
            log.error(
                    "AI workbook understanding JSON parsing failed for fileId={} after extraction. Normalized JSON was: {} Error: {}",
                    fileId, normalized, safeLogMessage(e.getMessage()), e);
            throw new WorkbookUnderstandingException("Model returned malformed JSON for workbook understanding");
        }
    }

    /**
     * Fixes malformed bare JSON values without using regex-based destructive
     * rewriting.
     * This safely quotes values like `fileId: 123e...`, `businessDomain: Sales`, or
     * `severity: ERROR` when the model emitted them without JSON quotes.
     */
    private String normalizeUnquotedTokens(String json) {
        StringBuilder result = new StringBuilder(json.length());
        int index = 0;

        while (index < json.length()) {
            char current = json.charAt(index);

            if (current == '"') {
                int stringStart = index;
                index++;
                while (index < json.length()) {
                    if (json.charAt(index) == '\\') {
                        index += 2;
                        continue;
                    }
                    if (json.charAt(index) == '"') {
                        break;
                    }
                    index++;
                }
                if (index >= json.length()) {
                    throw new WorkbookUnderstandingException("Model returned unterminated JSON string");
                }
                index++;
                result.append(json, stringStart, index);
                continue;
            }

            if (current == '{' || current == '}' || current == '[' || current == ']' || current == ','
                    || current == ':') {
                result.append(current);
                index++;
                continue;
            }

            if (Character.isWhitespace(current)) {
                result.append(current);
                index++;
                continue;
            }

            int tokenStart = index;
            while (index < json.length()) {
                char peek = json.charAt(index);
                if (peek == ',' || peek == '}' || peek == ']' || Character.isWhitespace(peek)) {
                    break;
                }
                index++;
            }

            String token = json.substring(tokenStart, index);
            if (isJsonNumber(token) || "true".equalsIgnoreCase(token) || "false".equalsIgnoreCase(token)
                    || "null".equalsIgnoreCase(token)) {
                result.append(token);
            } else {
                result.append('"').append(token).append('"');
            }
        }

        return result.toString();
    }

    private boolean isJsonNumber(String token) {
        return token.matches("-?(?:0|[1-9]\\d*)(?:\\.\\d+)?(?:[eE][+-]?\\d+)?");
    }

    private String safeLogMessage(String message) {
        if (message == null) {
            return "<empty>";
        }
        String sanitized = message.replaceAll("[\\r\\n]+", " ").trim();
        if (sanitized.length() > 200) {
            return sanitized.substring(0, 200) + "...";
        }
        return sanitized;
    }

    private HttpStatusCodeException extractHttpStatusException(Throwable throwable) {
        Throwable current = throwable;
        while (current != null) {
            if (current instanceof HttpStatusCodeException httpStatusCodeException) {
                return httpStatusCodeException;
            }
            current = current.getCause();
        }
        return null;
    }

    private String buildPrompt(WorkbookAnalysisResponse analysis) {
        StringBuilder sb = new StringBuilder();
        sb.append("Summarize the workbook metadata concisely.\n");
        sb.append("File: ").append(analysis.getFileName()).append("\n");
        sb.append("Sheets: ").append(analysis.getSheetCount()).append(" (list follows)\n");

        for (SheetAnalysis sheet : analysis.getSheets()) {
            sb.append("Sheet:").append(sheet.getSheetName())
                    .append(";rows=").append(sheet.getRowCount())
                    .append(";cols=").append(sheet.getColumnCount())
                    .append(";columns=");
            boolean first = true;
            for (ColumnAnalysis col : sheet.getColumns()) {
                if (!first)
                    sb.append("|");
                first = false;
                sb.append(col.getColumnName())
                        .append("(").append(col.getDetectedType()).append(",")
                        .append(col.getMissingPercentage()).append("%")
                        .append(",ex=");
                if (!col.getSampleValues().isEmpty()) {
                    sb.append(col.getSampleValues().stream().limit(2).toList().toString());
                } else {
                    sb.append("[]");
                }
            }
            sb.append("\n");
        }

        sb.append(
                "Return a compact JSON object with these fields exactly: summary (short), businessDomain, sheets (array of {sheetName,description,rowCount,columnCount}), importantColumns (array of {columnName,sheetName,reason}), possibleRelationships (array), dataQualityObservations (array), suggestedAnalyses (array). Keep all text concise (one sentence or phrase). No extra fields.");

        return sb.toString();
    }
}
