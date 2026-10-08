package com.aiexcel.platform.service;

import com.aiexcel.platform.dto.AnalysisRecommendation;
import com.aiexcel.platform.dto.AnalysisRecommendationResponse;
import com.aiexcel.platform.dto.WorkbookAnalysisResponse;
import com.aiexcel.platform.dto.WorkbookUnderstandingResponse;
import com.aiexcel.platform.entity.User;
import com.aiexcel.platform.exception.AnalysisRecommendationException;
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

import java.util.ArrayList;
import java.util.List;

/**
 * Generates AI-powered business analysis recommendations for an uploaded Excel
 * workbook.
 *
 * The AI decides WHICH analyses are useful. The backend stays deterministic:
 * - It only sends structured metadata (deterministic analysis + AI
 * understanding) to the AI.
 * - It never sends the raw workbook.
 * - It never generates or executes Java/SQL/formulas.
 * - It never executes any analysis itself; it only returns the AI's structured
 * suggestions.
 */
@Service
@Slf4j
public class AnalysisRecommendationService {

    private static final String SYSTEM_PROMPT = """
            You are a senior data analyst advising a non-technical business user.
            You receive structured metadata about an Excel workbook plus a prior AI understanding
            of the workbook, and you propose a focused set of business analysis recommendations.

            Hard rules you must follow:
            - NEVER generate executable code, Java, SQL, Python, formulas, or pseudo-code.
            - NEVER include instructions that the reader should "run" or "execute" anything.
            - Only return business-oriented, plain-language recommendations.
            - Only reference sheets and columns that are explicitly present in the provided metadata.
            - If metadata is sparse, recommend fewer, higher-confidence analyses rather than inventing details.
            - Output must be valid JSON matching the requested schema. No prose around the JSON.
            Strict rule: your entire reply MUST be a single valid JSON object that exactly
            matches the requested schema. Do NOT include any explanation, commentary,
            code fences, or additional text. Any deviation will be treated as invalid.
            """;

    private final ExcelAnalyzerService excelAnalyzerService;
    private final WorkbookUnderstandingService workbookUnderstandingService;
    private final ChatModel chatModel;

    public AnalysisRecommendationService(
            ExcelAnalyzerService excelAnalyzerService,
            WorkbookUnderstandingService workbookUnderstandingService,
            @Autowired(required = false) ChatModel chatModel) {
        this.excelAnalyzerService = excelAnalyzerService;
        this.workbookUnderstandingService = workbookUnderstandingService;
        this.chatModel = chatModel;
    }

    /**
     * Generates a structured set of business analysis recommendations for the given
     * file.
     *
     * @param fileId ID of the uploaded file
     * @param user   authenticated user (must own the file)
     * @return structured recommendation response
     */
    public AnalysisRecommendationResponse generateRecommendations(Long fileId, User user) {
        if (chatModel == null) {
            throw new AnalysisRecommendationException(
                    "AI recommendations are not configured. Set spring.ai.openai.api-key to enable this feature.");
        }

        log.info("Starting analysis recommendations for fileId={}, userId={}", fileId, user.getId());

        // 1. Verify ownership + obtain deterministic workbook analysis (throws
        // FileNotFoundException if not owned)
        WorkbookAnalysisResponse analysis = excelAnalyzerService.analyzeWorkbook(fileId, user);

        // 2. Obtain the existing AI understanding of the workbook (re-uses the same AI
        // in a read-only way)
        WorkbookUnderstandingResponse understanding = workbookUnderstandingService.understandWorkbook(fileId, user);

        // 3. Build a structured prompt from structured metadata ONLY (no raw data)
        BeanOutputConverter<AnalysisRecommendationResponse> converter = new BeanOutputConverter<>(
                AnalysisRecommendationResponse.class);

        String userPrompt = buildPrompt(analysis, understanding) + "\n\n" + converter.getFormat();

        // 4. Call AI and parse the structured response
        try {
            SystemMessage systemMessage = new SystemMessage(SYSTEM_PROMPT);
            UserMessage userMessage = new UserMessage(userPrompt);
            Prompt prompt = new Prompt(java.util.List.of(systemMessage, userMessage));
            ChatResponse response = chatModel.call(prompt);

            String content = response.getResult().getOutput().getText();
            String jsonContent = sanitizeAiJson(content, fileId, analysis.getFileName());
            AnalysisRecommendationResponse parsed = converter.convert(jsonContent);

            // Defensive: ensure the DTO is well-formed even if the AI omits fields
            return sanitize(fileId, analysis.getFileName(), parsed);
        } catch (AnalysisRecommendationException e) {
            // Already a clean error; don't wrap again
            throw e;
        } catch (Exception e) {
            HttpStatusCodeException httpEx = extractHttpStatusException(e);
            if (httpEx != null) {
                log.error("AI provider rejected recommendations request for fileId={} status={} body={}",
                        fileId, httpEx.getStatusCode(), httpEx.getResponseBodyAsString(), httpEx);
                if (HttpStatus.PAYMENT_REQUIRED.equals(httpEx.getStatusCode())
                        || httpEx.getStatusCode().value() == 402) {
                    throw new AnalysisRecommendationException(
                            "AI provider rejected the request because the OpenRouter budget is exhausted.");
                }
                if (HttpStatus.UNAUTHORIZED.equals(httpEx.getStatusCode()) || httpEx.getStatusCode().value() == 401) {
                    throw new AnalysisRecommendationException(
                            "AI provider rejected the request because the OpenRouter API key is invalid or expired.");
                }
                throw new AnalysisRecommendationException(
                        "AI provider error while generating analysis recommendations: " + httpEx.getStatusText());
            }
            // Log the full error server-side, but don't expose internals to client
            log.error("AI call failed for recommendations fileId={}", fileId, e);
            throw new AnalysisRecommendationException("AI failed to generate analysis recommendations");
        }
    }

    private String sanitizeAiJson(String rawText, Long fileId, String fileName) {
        if (rawText == null || rawText.isBlank()) {
            return rawText;
        }

        String candidate = rawText.trim();
        if (candidate.startsWith("```")) {
            int end = candidate.indexOf("```", 3);
            if (end > 3) {
                String inside = candidate.substring(3, end).trim();
                if (inside.regionMatches(true, 0, "json", 0, 4)) {
                    inside = inside.substring(4).trim();
                }
                if (!inside.isBlank()) {
                    candidate = inside;
                }
            }
        }

        int firstBrace = candidate.indexOf('{');
        int lastBrace = candidate.lastIndexOf('}');
        if (firstBrace >= 0 && lastBrace > firstBrace) {
            candidate = candidate.substring(firstBrace, lastBrace + 1);
        }

        candidate = quoteBareJsonValues(candidate);

        if (fileId != null && candidate.contains("\"fileId\"")) {
            candidate = forceJsonValue(candidate, "fileId", String.valueOf(fileId));
        }
        if (fileName != null && candidate.contains("\"fileName\"")) {
            String escapedFileName = fileName.replace("\\", "\\\\").replace("\"", "\\\"");
            candidate = forceJsonValue(candidate, "fileName", escapedFileName);
        }

        return candidate;
    }

    private String quoteBareJsonValues(String candidate) {
        StringBuilder sb = new StringBuilder(candidate);
        for (int i = 0; i < sb.length(); i++) {
            char c = sb.charAt(i);
            if (c == '"') {
                int keyEnd = i;
                while (keyEnd < sb.length() && sb.charAt(keyEnd) != '"') {
                    keyEnd++;
                }
                if (keyEnd >= sb.length()) {
                    break;
                }
                int colonIndex = sb.indexOf(":", i + 1);
                if (colonIndex == -1) {
                    continue;
                }
                int valueStart = colonIndex + 1;
                while (valueStart < sb.length() && Character.isWhitespace(sb.charAt(valueStart))) {
                    valueStart++;
                }
                int valueEnd = valueStart;
                if (valueStart < sb.length() && (Character.isLetterOrDigit(sb.charAt(valueStart))
                        || sb.charAt(valueStart) == '-' || sb.charAt(valueStart) == '.')) {
                    while (valueEnd < sb.length() && sb.charAt(valueEnd) != ',' && sb.charAt(valueEnd) != '}') {
                        valueEnd++;
                    }
                    String rawValue = sb.substring(valueStart, valueEnd).trim();
                    if (!rawValue.isEmpty() && !rawValue.startsWith("\"") && !rawValue.equals("true")
                            && !rawValue.equals("false") && !rawValue.equals("null")) {
                        sb.insert(valueEnd, '"');
                        sb.insert(valueStart, '"');
                        i = valueEnd + 1;
                    }
                }
            }
        }
        return sb.toString();
    }

    private String forceJsonValue(String json, String propertyName, String value) {
        String token = "\"" + propertyName + "\"";
        int propertyIndex = json.indexOf(token);
        if (propertyIndex < 0) {
            return json;
        }
        int colonIndex = json.indexOf(':', propertyIndex + token.length());
        if (colonIndex < 0) {
            return json;
        }
        int valueStart = colonIndex + 1;
        while (valueStart < json.length() && Character.isWhitespace(json.charAt(valueStart))) {
            valueStart++;
        }
        int valueEnd = valueStart;
        while (valueEnd < json.length() && json.charAt(valueEnd) != ',' && json.charAt(valueEnd) != '}') {
            valueEnd++;
        }

        return json.substring(0, valueStart) + "\"" + value.replace("\\", "\\\\").replace("\"", "\\\"") + "\""
                + json.substring(valueEnd);
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

    private AnalysisRecommendationResponse sanitize(Long fileId, String fileName,
            AnalysisRecommendationResponse parsed) {
        if (parsed == null) {
            throw new AnalysisRecommendationException("AI returned an empty recommendations response");
        }

        // Always stamp file identity from the verified request, never trust AI to echo
        // it
        parsed.setFileId(fileId);
        if (fileName != null) {
            parsed.setFileName(fileName);
        }

        List<AnalysisRecommendation> recs = parsed.getRecommendations();
        if (recs == null) {
            parsed.setRecommendations(new ArrayList<>());
        }
        return parsed;
    }

    private String buildPrompt(WorkbookAnalysisResponse analysis, WorkbookUnderstandingResponse understanding) {
        StringBuilder sb = new StringBuilder();
        sb.append("Based on the structured workbook metadata and the prior AI understanding below, ")
                .append("propose a focused set of business analysis recommendations.\n\n");

        sb.append("=== Verified workbook identity ===\n");
        sb.append("fileId: ").append(analysis.getFileId()).append("\n");
        sb.append("fileName: ").append(analysis.getFileName()).append("\n\n");

        // --- Deterministic workbook metadata ---
        sb.append("=== Workbook Metadata (deterministic) ===\n");
        sb.append("File: ").append(analysis.getFileName()).append("\n");
        sb.append("Sheets: ").append(analysis.getSheetCount()).append("\n\n");
        for (var sheet : analysis.getSheets()) {
            sb.append("Sheet: ").append(sheet.getSheetName())
                    .append("  rows=").append(sheet.getRowCount())
                    .append("  cols=").append(sheet.getColumnCount()).append("\n");
            for (var col : sheet.getColumns()) {
                sb.append("  - ").append(col.getColumnName())
                        .append(" (").append(col.getDetectedType()).append(")");
                if (col.getMissingPercentage() > 0) {
                    sb.append(", missing=").append(col.getMissingPercentage()).append("%");
                }
                if (col.getUniqueCount() > 0) {
                    sb.append(", unique=").append(col.getUniqueCount());
                }
                sb.append("\n");
            }
            sb.append("\n");
        }

        // --- Prior AI understanding (already structured) ---
        if (understanding != null) {
            sb.append("=== Prior AI Understanding (structured) ===\n");
            if (understanding.getBusinessDomain() != null) {
                sb.append("Domain: ").append(understanding.getBusinessDomain()).append("\n");
            }
            if (understanding.getSummary() != null) {
                sb.append("Summary: ").append(understanding.getSummary()).append("\n");
            }
            if (understanding.getImportantColumns() != null && !understanding.getImportantColumns().isEmpty()) {
                sb.append("Important columns:\n");
                for (var c : understanding.getImportantColumns()) {
                    sb.append("  - ").append(c.getSheetName()).append(".")
                            .append(c.getColumnName()).append(" — ").append(c.getReason()).append("\n");
                }
            }
            if (understanding.getPossibleRelationships() != null
                    && !understanding.getPossibleRelationships().isEmpty()) {
                sb.append("Possible relationships:\n");
                for (var r : understanding.getPossibleRelationships()) {
                    sb.append("  - ").append(r.getEntity1()).append(" <-> ")
                            .append(r.getEntity2()).append(" (").append(r.getRelationshipType())
                            .append("): ").append(r.getDescription()).append("\n");
                }
            }
            if (understanding.getDataQualityObservations() != null
                    && !understanding.getDataQualityObservations().isEmpty()) {
                sb.append("Data quality observations:\n");
                for (var q : understanding.getDataQualityObservations()) {
                    sb.append("  - [").append(q.getSeverity()).append("] ")
                            .append(q.getAffectedEntity()).append(": ")
                            .append(q.getObservation()).append("\n");
                }
            }
            sb.append("\n");
        }

        sb.append(
                """
                        Return a JSON object with exactly these fields:
                        - fileId (number): use the exact verified value above, never invent another value
                        - fileName (string): use the exact verified file name above
                        - overview (string): 1-2 sentence high-level summary of the recommendation set
                        - recommendations (array): each entry must have:
                            * title (string): concise recommendation name
                            * description (string): what this analysis would reveal, in 1-2 sentences
                            * analysisType (string): one of TREND, COMPARISON, DISTRIBUTION, CORRELATION, ANOMALY, SEGMENTATION, SUMMARY
                            * relevantSheets (array of strings): sheet names this analysis would draw from
                            * relevantColumns (array of strings): column names this analysis would focus on
                            * businessValue (string): why this matters to a business user
                        Recommend 3-6 high-value, distinct analyses. Do not duplicate ideas.
                        Only reference sheets and columns that appear in the metadata above.
                        """);

        return sb.toString();
    }
}
