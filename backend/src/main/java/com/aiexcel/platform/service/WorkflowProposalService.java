package com.aiexcel.platform.service;

import com.aiexcel.platform.dto.*;
import com.aiexcel.platform.entity.User;
import com.aiexcel.platform.exception.WorkbookUnderstandingException;
import com.aiexcel.platform.exception.WorkflowProposalException;
import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.extern.slf4j.Slf4j;
import org.springframework.ai.chat.model.ChatModel;
import org.springframework.ai.chat.model.ChatResponse;
import org.springframework.ai.chat.prompt.Prompt;
import org.springframework.ai.converter.BeanOutputConverter;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/**
 * Generates an AI-proposed workflow to fulfil a natural-language analysis
 * request.
 *
 * The AI decides WHAT steps to include; the backend stays deterministic:
 * - It only sends structured metadata (deterministic analysis + AI
 * understanding) to the AI.
 * - It never sends the raw workbook.
 * - It never generates or executes Java/SQL/Python/formulas.
 * - It NEVER executes the proposed workflow.
 *
 * Every operation proposed by the AI is validated against the
 * {@link WorkflowOperation} catalog
 * before the response is returned. Unknown operations cause a clean failure
 * (422).
 */
@Service
@Slf4j
public class WorkflowProposalService {

    private static final String SYSTEM_PROMPT = """
            You are a senior data analyst. A user has described an analysis they want to perform
            on an Excel workbook. Your job is to propose a structured workflow plan.

            Hard rules you must follow:
            - NEVER generate executable code, Java, SQL, Python, formulas, or pseudo-code.
            - NEVER include instructions that the reader should "run" or "execute" anything.
            - You may ONLY propose steps using these exact operation names:
                READ_SHEET, SELECT_COLUMNS, REMOVE_EMPTY_ROWS, REMOVE_DUPLICATES,
                FILTER, SORT, CALCULATE, GROUP_BY, AGGREGATE,
                JOIN_SHEETS, PIVOT, CREATE_SUMMARY, CREATE_CHART,
                VALIDATE, GENERATE_PDF, EXPORT_EXCEL
              If an operation cannot be expressed using these, omit it or describe it
              using a combination of the allowed operations.
            - Only reference sheets and columns that are explicitly present in the provided metadata.
            - Output must be valid JSON matching the requested schema. No prose around the JSON.
            """;

    private static final List<String> ALLOWED_OPERATIONS = List.of(
            "READ_SHEET", "SELECT_COLUMNS", "REMOVE_EMPTY_ROWS", "REMOVE_DUPLICATES",
            "FILTER", "SORT", "CALCULATE", "GROUP_BY", "AGGREGATE",
            "JOIN_SHEETS", "PIVOT", "CREATE_SUMMARY", "CREATE_CHART",
            "VALIDATE", "GENERATE_PDF", "EXPORT_EXCEL");

    private final ExcelAnalyzerService excelAnalyzerService;
    private final WorkbookUnderstandingService workbookUnderstandingService;
    private final ChatModel chatModel;
    private final ObjectMapper objectMapper = new ObjectMapper();

    public WorkflowProposalService(
            ExcelAnalyzerService excelAnalyzerService,
            WorkbookUnderstandingService workbookUnderstandingService,
            @Autowired(required = false) ChatModel chatModel) {
        this.excelAnalyzerService = excelAnalyzerService;
        this.workbookUnderstandingService = workbookUnderstandingService;
        this.chatModel = chatModel;
    }

    /**
     * Proposes a structured workflow for the given analysis request.
     *
     * @param fileId      ID of the uploaded file
     * @param user        authenticated user (must own the file)
     * @param instruction the user's natural-language analysis request
     * @return structured workflow proposal
     */
    public WorkflowProposalResponse proposeWorkflow(Long fileId, User user, String instruction) {
        if (chatModel == null) {
            throw new WorkflowProposalException(
                    "AI workflow proposal is not configured. Set spring.ai.openai.api-key to enable this feature.");
        }

        if (instruction == null || instruction.isBlank()) {
            throw new WorkflowProposalException("Instruction must not be blank");
        }

        log.info("Generating workflow proposal for fileId={}, userId={}", fileId, user.getId());

        // 1. Verify ownership + obtain deterministic analysis (throws
        // FileNotFoundException if not owned)
        WorkbookAnalysisResponse analysis = excelAnalyzerService.analyzeWorkbook(fileId, user);

        // 2. Obtain AI understanding
        WorkbookUnderstandingResponse understanding = workbookUnderstandingService.understandWorkbook(fileId, user);

        // 3. Build structured prompt
        BeanOutputConverter<WorkflowProposalResponse> converter = new BeanOutputConverter<>(
                WorkflowProposalResponse.class);

        String userPrompt = buildPrompt(instruction, analysis, understanding) + "\n\n" + converter.getFormat();
        String fullPrompt = SYSTEM_PROMPT + "\n\n" + userPrompt;
        log.debug("Workflow proposal LLM prompt for fileId={}:\n{}", fileId, fullPrompt);

        // 4. Call AI and parse
        try {
            Prompt prompt = new Prompt(fullPrompt);
            log.info("Workflow proposal AI request start for fileId={}, userId={}", fileId, user.getId());
            ChatResponse response = chatModel.call(prompt);
            log.info("Workflow proposal AI response received for fileId={}, userId={}", fileId, user.getId());

            String rawResponse = response == null || response.getResult() == null
                    || response.getResult().getOutput() == null
                            ? null
                            : response.getResult().getOutput().getText();
            log.debug("Workflow proposal raw AI response for fileId={}: {}", fileId, rawResponse);

            String content = extractJsonPayload(response);
            log.debug("Workflow proposal extracted JSON for fileId={}: {}", fileId, content);

            WorkflowProposalResponse parsed = converter.convert(sanitizeAiJson(content));
            log.info("Workflow proposal conversion successful for fileId={}, workflowName={}, steps={}",
                    fileId, parsed.getWorkflowName(), parsed.getSteps() == null ? 0 : parsed.getSteps().size());

            // 5. Validate all operations against the catalog
            validateOperations(parsed);
            log.info("Workflow proposal validation successful for fileId={}, stepCount={}",
                    fileId, parsed.getSteps() == null ? 0 : parsed.getSteps().size());

            // 6. Sanitize: stamp from verified request, never trust AI for these
            WorkflowProposalResponse finalResponse = sanitize(fileId, instruction, parsed);
            log.info("Workflow proposal finalized for fileId={}, workflowName={}",
                    fileId, finalResponse.getWorkflowName());
            return finalResponse;
        } catch (WorkflowProposalException e) {
            log.error("Workflow proposal validation/parsing failed for fileId={}, message={}", fileId, e.getMessage(),
                    e);
            throw e;
        } catch (WorkbookUnderstandingException e) {
            // Don't wrap WorkbookUnderstandingException - let it propagate so
            // GlobalExceptionHandler
            // can properly detect budget/payment errors and map to 402 vs 422
            log.error("Workflow proposal failed: workbook understanding service error for fileId={}, message={}",
                    fileId, e.getMessage(), e);
            throw e;
        } catch (Exception e) {
            log.error("AI call failed for workflow proposal fileId={} with cause: {}", fileId, e.getMessage(), e);
            throw new WorkflowProposalException("AI failed to generate workflow proposal: " + e.getMessage(), e);
        }
    }

    private String sanitizeAiJson(String content) {
        if (content == null || content.isBlank()) {
            return content;
        }

        String candidate = content.trim();
        if (candidate.startsWith("```")) {
            int codeFenceEnd = candidate.indexOf("```", 3);
            if (codeFenceEnd > 3) {
                String inside = candidate.substring(3, codeFenceEnd).trim();
                if (inside.toLowerCase(Locale.ROOT).startsWith("json")) {
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

    private String extractJsonPayload(ChatResponse response) {
        if (response == null || response.getResult() == null || response.getResult().getOutput() == null) {
            throw new WorkflowProposalException("AI returned an empty workflow proposal");
        }

        String raw = response.getResult().getOutput().getText();
        if (raw == null || raw.isBlank()) {
            throw new WorkflowProposalException("AI returned an empty workflow proposal");
        }

        String candidate = raw.trim();

        // Handle fenced JSON blocks like ```json ... ``` or ``` ... ```
        if (candidate.startsWith("```")) {
            int codeFenceEnd = candidate.indexOf("```", 3);
            if (codeFenceEnd > 3) {
                String inside = candidate.substring(3, codeFenceEnd).trim();
                if (inside.toLowerCase(Locale.ROOT).startsWith("json")) {
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

        if (candidate.isBlank()) {
            throw new WorkflowProposalException("AI returned an empty workflow proposal");
        }

        try {
            objectMapper.readTree(candidate);
        } catch (Exception e) {
            throw new WorkflowProposalException(
                    "AI returned malformed JSON for workflow proposal: " + e.getMessage(), e);
        }

        return candidate;
    }

    /**
     * Validates that every step's operation is in the allowed catalog.
     * Throws WorkflowProposalException with details if any step is invalid.
     */
    private void validateOperations(WorkflowProposalResponse parsed) {
        if (parsed == null || parsed.getSteps() == null) {
            throw new WorkflowProposalException("AI returned an empty or malformed workflow proposal");
        }

        List<String> invalidOps = new ArrayList<>();
        for (int i = 0; i < parsed.getSteps().size(); i++) {
            WorkflowStep step = parsed.getSteps().get(i);
            if (step == null) {
                invalidOps.add("Step " + (i + 1) + ": null step object");
                continue;
            }
            String op = step.getOperation();
            if (!WorkflowOperation.isValid(op)) {
                invalidOps.add("Step " + (i + 1) + ": '" + op + "' is not in the allowed operation catalog");
            }
        }

        if (!invalidOps.isEmpty()) {
            throw new WorkflowProposalException(
                    "AI proposed operations not in the allowed catalog: " + String.join("; ", invalidOps));
        }
    }

    private WorkflowProposalResponse sanitize(Long fileId, String instruction, WorkflowProposalResponse parsed) {
        if (parsed == null) {
            throw new WorkflowProposalException("AI returned an empty workflow proposal");
        }

        parsed.setFileId(fileId);
        parsed.setInstruction(instruction);

        if (parsed.getSteps() == null) {
            parsed.setSteps(new ArrayList<>());
        }

        // Normalize null parameters to an empty map. The LLM can return
        // "parameters": null for steps that need no parameters, but the
        // project's workflow convention is "parameters": {}.
        for (WorkflowStep step : parsed.getSteps()) {
            if (step != null && step.getParameters() == null) {
                step.setParameters(new java.util.HashMap<>());
            }
        }
        return parsed;
    }

    private String buildPrompt(String instruction, WorkbookAnalysisResponse analysis,
            WorkbookUnderstandingResponse understanding) {
        StringBuilder sb = new StringBuilder();
        sb.append("=== User's Analysis Request ===\n");
        sb.append(instruction).append("\n\n");

        // Deterministic metadata
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

        // Prior AI understanding
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
        }

        sb.append("\n=== Allowed Operations ===\n");
        sb.append("You may ONLY use these operation names in the 'operation' field of each step:\n");
        for (String op : ALLOWED_OPERATIONS) {
            sb.append("  - ").append(op).append("\n");
        }

        sb.append("""

                === Parameters by operation ===
                Each operation has a fixed parameter contract. Put every required operation
                argument inside the step's "parameters" object using the keys listed below.
                "inputColumns" describes the columns the step uses; it is NOT a substitute
                for required operation parameters.

                GROUP_BY
                - groupColumns: required, comma-separated column names to group by

                FILTER
                - column: required
                - operator: required
                - value: required

                SORT
                - column: required
                - direction: required
                - type: optional

                AGGREGATE (Use this for math calculations like SUM, AVERAGE, MIN, MAX)
                - groupColumns: optional (omit for overall summary)
                - aggregateColumn: required
                - aggregateFunction: required (SUM, AVERAGE, MIN, MAX, or COUNT)

                PIVOT
                - rowColumn: required
                - pivotColumn: required
                - valueColumn: required
                - aggregationFunction: required

                CALCULATE
                - targetColumn: required
                - expression: required

                JOIN_SHEETS
                - leftSheet: required
                - rightSheet: required
                - joinColumn: required
                - joinType: required

                CREATE_CHART
                - sourceSheet: required
                - chartType: required
                - labelColumn: required
                - dataColumn: required
                - title: optional

                These operations require no parameters:
                READ_SHEET, SELECT_COLUMNS, REMOVE_EMPTY_ROWS, REMOVE_DUPLICATES,
                CREATE_SUMMARY (Only projects columns, does not calculate math), VALIDATE, GENERATE_PDF, EXPORT_EXCEL

                """);

        sb.append(
                """
                        Return a JSON object with exactly these fields:
                        - fileId (number): echo the file id
                        - instruction (string): echo the user's request verbatim
                        - workflowName (string): a short descriptive name for this workflow
                        - description (string): 1-2 sentence plain-language description of what this workflow accomplishes
                        - steps (array): ordered steps, each with:
                            * order (integer): 1-based execution order
                            * operation (string): MUST be one of the allowed operation names above
                            * description (string): what this step does in plain language
                            * inputSheets (array of strings): sheet names this step reads
                            * inputColumns (array of strings): column names this step uses
                            * parameters (object or null): key-value parameters for the operation (e.g. {"sortColumn":"Date","direction":"ASC"})
                        Keep the workflow focused and practical. Do not include more than 10 steps.
                        Only reference sheets and columns that appear in the metadata above.
                        """);

        return sb.toString();
    }
}
