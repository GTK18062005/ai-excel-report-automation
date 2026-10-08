package com.aiexcel.platform.service;

import com.aiexcel.platform.dto.*;
import com.aiexcel.platform.entity.User;
import com.aiexcel.platform.exception.AnalysisIntentException;
import lombok.extern.slf4j.Slf4j;
import org.springframework.ai.chat.model.ChatModel;
import org.springframework.ai.chat.model.ChatResponse;
import org.springframework.ai.chat.prompt.Prompt;
import org.springframework.ai.converter.BeanOutputConverter;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.ArrayList;

/**
 * Parses a natural-language analysis request against an uploaded workbook.
 *
 * The AI interprets WHAT the user wants to do; the backend stays deterministic:
 * - It only sends structured metadata (deterministic analysis + AI
 * understanding) to the AI.
 * - It never sends the raw workbook.
 * - It never generates or executes Java/SQL/Python/formulas.
 * - It never executes any analysis itself; it only returns the AI's structured
 * interpretation.
 */
@Service
@Slf4j
public class AnalysisIntentService {

    private static final String SYSTEM_PROMPT = """
            You are a senior data analyst. A user has described an analysis they want to perform
            on an Excel workbook. Your job is to interpret their intent and return a structured plan.

            Hard rules you must follow:
            - NEVER generate executable code, Java, SQL, Python, formulas, or pseudo-code.
            - NEVER include instructions that the reader should "run" or "execute" anything.
            - Only return a plain-language intent classification and analysis plan.
            - Only reference sheets and columns that are explicitly present in the provided metadata.
            - Output must be valid JSON matching the requested schema. No prose around the JSON.
            """;

    private final ExcelAnalyzerService excelAnalyzerService;
    private final WorkbookUnderstandingService workbookUnderstandingService;
    private final ChatModel chatModel;

    public AnalysisIntentService(
            ExcelAnalyzerService excelAnalyzerService,
            WorkbookUnderstandingService workbookUnderstandingService,
            @Autowired(required = false) ChatModel chatModel) {
        this.excelAnalyzerService = excelAnalyzerService;
        this.workbookUnderstandingService = workbookUnderstandingService;
        this.chatModel = chatModel;
    }

    /**
     * Parses a natural-language analysis instruction against the given file.
     *
     * @param fileId      ID of the uploaded file
     * @param user        authenticated user (must own the file)
     * @param instruction the user's natural-language analysis request
     * @return structured intent interpretation
     */
    public AnalysisIntentResponse parseIntent(Long fileId, User user, String instruction) {
        if (chatModel == null) {
            throw new AnalysisIntentException(
                    "AI intent parsing is not configured. Set spring.ai.openai.api-key to enable this feature.");
        }

        log.info("Parsing analysis intent for fileId={}, userId={}", fileId, user.getId());

        // 1. Verify ownership + obtain deterministic workbook analysis (throws
        // FileNotFoundException if not owned)
        WorkbookAnalysisResponse analysis = excelAnalyzerService.analyzeWorkbook(fileId, user);

        // 2. Obtain the existing AI understanding of the workbook
        WorkbookUnderstandingResponse understanding = workbookUnderstandingService.understandWorkbook(fileId, user);

        // 3. Build structured prompt
        BeanOutputConverter<AnalysisIntentResponse> converter = new BeanOutputConverter<>(AnalysisIntentResponse.class);

        String userPrompt = buildPrompt(instruction, analysis, understanding) + "\n\n" + converter.getFormat();

        // 4. Call AI and parse the structured response
        try {
            Prompt prompt = new Prompt(SYSTEM_PROMPT + "\n\n" + userPrompt);
            ChatResponse response = chatModel.call(prompt);

            String content = response.getResult().getOutput().getText();
            AnalysisIntentResponse parsed = converter.convert(sanitizeAiJson(content));

            return sanitize(fileId, instruction, parsed);
        } catch (AnalysisIntentException e) {
            // Already a clean error; don't wrap again
            throw e;
        } catch (Exception e) {
            log.error("AI call failed for intent parsing fileId={}", fileId, e);
            throw new AnalysisIntentException("AI failed to parse analysis intent");
        }
    }

    private String sanitizeAiJson(String rawText) {
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

        StringBuilder sb = new StringBuilder(candidate);
        for (int i = 0; i < sb.length(); i++) {
            char c = sb.charAt(i);
            if (c == '"') {
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

    private AnalysisIntentResponse sanitize(Long fileId, String instruction, AnalysisIntentResponse parsed) {
        if (parsed == null) {
            throw new AnalysisIntentException("AI returned an empty intent response");
        }

        // Always stamp from the verified request, never trust AI to echo it
        parsed.setFileId(fileId);
        parsed.setInstruction(instruction);

        if (parsed.getRequestedAnalyses() == null) {
            parsed.setRequestedAnalyses(new ArrayList<>());
        }
        if (parsed.getRelevantSheets() == null) {
            parsed.setRelevantSheets(new ArrayList<>());
        }
        if (parsed.getRelevantColumns() == null) {
            parsed.setRelevantColumns(new ArrayList<>());
        }
        if (parsed.getConstraints() == null) {
            parsed.setConstraints(new ArrayList<>());
        }
        return parsed;
    }

    private String buildPrompt(String instruction, WorkbookAnalysisResponse analysis,
            WorkbookUnderstandingResponse understanding) {
        StringBuilder sb = new StringBuilder();
        sb.append("=== User's Analysis Request ===\n");
        sb.append(instruction).append("\n\n");

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

        // --- Prior AI understanding ---
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
        }

        sb.append(
                """
                        Return a JSON object with exactly these fields:
                        - fileId (number): echo the file id
                        - instruction (string): echo the user's request verbatim
                        - intent (string): a short classification of the user's goal (e.g. COMPARISON, TREND_ANALYSIS, SEGMENTATION, ANOMALY_DETECTION)
                        - requestedAnalyses (array): each entry must have {name, analysisType, description}
                          * name: human-readable step name
                          * analysisType: one of COMPARISON, TREND, DISTRIBUTION, CORRELATION, SEGMENTATION, ANOMALY, SUMMARY
                          * description: what this analysis step would produce
                        - relevantSheets (array of strings): sheet names this analysis needs
                        - relevantColumns (array of strings): column names this analysis needs
                        - constraints (array of strings): any constraints or filters the instruction implies (e.g. time range, region, threshold)
                        Only reference sheets and columns that appear in the metadata above.
                        """);

        return sb.toString();
    }
}
