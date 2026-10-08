package com.aiexcel.platform.controller;

import com.aiexcel.platform.dto.AnalysisIntentRequest;
import com.aiexcel.platform.dto.AnalysisIntentResponse;
import com.aiexcel.platform.dto.AnalysisRecommendationResponse;
import com.aiexcel.platform.dto.FileResponse;
import com.aiexcel.platform.dto.WorkbookAnalysisResponse;
import com.aiexcel.platform.dto.WorkbookPreviewResponse;
import com.aiexcel.platform.dto.WorkbookUnderstandingResponse;
import com.aiexcel.platform.dto.WorkflowProposalRequest;
import com.aiexcel.platform.dto.WorkflowProposalResponse;
import com.aiexcel.platform.dto.WorkflowValidationRequest;
import com.aiexcel.platform.dto.WorkflowValidationResponse;
import com.aiexcel.platform.dto.WorkflowApprovalResponse;
import com.aiexcel.platform.dto.WorkflowExecutionResponse;
import com.aiexcel.platform.entity.User;
import com.aiexcel.platform.service.AnalysisIntentService;
import com.aiexcel.platform.service.AnalysisRecommendationService;
import com.aiexcel.platform.service.ExcelAnalyzerService;
import com.aiexcel.platform.service.FileService;
import com.aiexcel.platform.service.WorkbookUnderstandingService;
import com.aiexcel.platform.service.WorkflowExecutionService;
import com.aiexcel.platform.service.WorkflowProposalService;
import com.aiexcel.platform.service.WorkflowValidatorService;
import com.aiexcel.platform.service.WorkflowApprovalService;
import jakarta.validation.Valid;
import lombok.RequiredArgsConstructor;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.security.core.annotation.AuthenticationPrincipal;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.nio.file.Path;
import java.util.List;

@RestController
@RequestMapping("/api/files")
@RequiredArgsConstructor
public class FileController {

    private final FileService fileService;
    private final ExcelAnalyzerService excelAnalyzerService;
    private final WorkbookUnderstandingService workbookUnderstandingService;
    private final AnalysisRecommendationService analysisRecommendationService;
    private final AnalysisIntentService analysisIntentService;
    private final WorkflowProposalService workflowProposalService;
    private final WorkflowValidatorService workflowValidatorService;
    private final WorkflowApprovalService workflowApprovalService;
    private final WorkflowExecutionService workflowExecutionService;

    /**
     * POST /api/files
     * Upload a new file (.xlsx, .xls, or .csv). Max 50MB.
     * Content-Type: multipart/form-data
     * Form field name: "file"
     */
    @PostMapping(consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<FileResponse> uploadFile(
            @RequestParam("file") MultipartFile file,
            @AuthenticationPrincipal User currentUser
    ) {
        FileResponse response = fileService.uploadFile(file, currentUser);
        return ResponseEntity.status(HttpStatus.CREATED).body(response);
    }

    /**
     * GET /api/files
     * List all non-deleted files belonging to the current user.
     */
    @GetMapping
    public ResponseEntity<List<FileResponse>> listFiles(
            @AuthenticationPrincipal User currentUser
    ) {
        return ResponseEntity.ok(fileService.getUserFiles(currentUser));
    }

    /**
     * GET /api/files/{id}
     * Get metadata for a single file. Returns 404 if not found or not owned by current user.
     */
    @GetMapping("/{id}")
    public ResponseEntity<FileResponse> getFile(
            @PathVariable Long id,
            @AuthenticationPrincipal User currentUser
    ) {
        return ResponseEntity.ok(fileService.getFileById(id, currentUser));
    }

    /**
     * DELETE /api/files/{id}
     * Soft-delete: sets status to DELETED. Returns 204 No Content.
     */
    @DeleteMapping("/{id}")
    public ResponseEntity<Void> deleteFile(
            @PathVariable Long id,
            @AuthenticationPrincipal User currentUser
    ) {
        fileService.softDeleteFile(id, currentUser);
        return ResponseEntity.noContent().build();
    }

    /**
     * GET /api/files/{id}/analysis
     * Analyze an uploaded Excel workbook and return structured metadata.
     * Returns 404 if file not found or not owned by user.
     * Returns 422 for invalid/unreadable workbook files.
     */
    @GetMapping("/{id}/analysis")
    public ResponseEntity<WorkbookAnalysisResponse> analyzeWorkbook(
            @PathVariable Long id,
            @AuthenticationPrincipal User currentUser
    ) {
        WorkbookAnalysisResponse response = excelAnalyzerService.analyzeWorkbook(id, currentUser);
        return ResponseEntity.ok(response);
    }

    /**
     * GET /api/files/{id}/preview
     * Return a paginated preview of rows from a specific worksheet.
     * Returns 404 if file not found or not owned by user.
     * Returns 400 if the requested sheet does not exist.
     * Maximum pageSize is 100.
     */
    @GetMapping("/{id}/preview")
    public ResponseEntity<WorkbookPreviewResponse> previewSheet(
            @PathVariable Long id,
            @RequestParam String sheet,
            @RequestParam(defaultValue = "0") int page,
            @RequestParam(defaultValue = "20") int pageSize,
            @AuthenticationPrincipal User currentUser
    ) {
        WorkbookPreviewResponse response = excelAnalyzerService.getSheetPreview(id, currentUser, sheet, page, pageSize);
        return ResponseEntity.ok(response);
    }

    /**
     * GET /api/files/{id}/understanding
     * Generate an AI-powered business understanding of the workbook.
     * Returns 404 if file not found or not owned by user.
     * Returns 422 if AI is not configured or fails.
     */
    @GetMapping("/{id}/understanding")
    public ResponseEntity<WorkbookUnderstandingResponse> understandWorkbook(
            @PathVariable Long id,
            @AuthenticationPrincipal User currentUser
    ) {
        WorkbookUnderstandingResponse response = workbookUnderstandingService.understandWorkbook(id, currentUser);
        return ResponseEntity.ok(response);
    }

    /**
     * GET /api/files/{id}/recommendations
     * Generate AI-powered business analysis recommendations for the workbook.
     * Returns 404 if file not found or not owned by user.
     * Returns 422 if AI is not configured or fails.
     */
    @GetMapping("/{id}/recommendations")
    public ResponseEntity<AnalysisRecommendationResponse> getRecommendations(
            @PathVariable Long id,
            @AuthenticationPrincipal User currentUser
    ) {
        AnalysisRecommendationResponse response = analysisRecommendationService.generateRecommendations(id, currentUser);
        return ResponseEntity.ok(response);
    }

    /**
     * POST /api/files/{id}/analysis-request
     * Parse a natural-language analysis request into a structured intent.
     * Body: { "instruction": "..." }
     * Returns 400 if the instruction is blank.
     * Returns 404 if file not found or not owned by user.
     * Returns 422 if AI is not configured or fails.
     */
    @PostMapping(
            value = "/{id}/analysis-request",
            consumes = MediaType.APPLICATION_JSON_VALUE
    )
    public ResponseEntity<AnalysisIntentResponse> parseAnalysisRequest(
            @PathVariable Long id,
            @Valid @RequestBody AnalysisIntentRequest request,
            @AuthenticationPrincipal User currentUser
    ) {
        AnalysisIntentResponse response = analysisIntentService.parseIntent(
                id, currentUser, request.getInstruction()
        );
        return ResponseEntity.ok(response);
    }

    /**
     * POST /api/files/{id}/workflow-proposal
     * Propose a structured workflow that fulfils the user's analysis request.
     * The AI proposes the steps; the backend does NOT execute them.
     * Body: { "instruction": "..." }
     * Returns 400 if the instruction is blank.
     * Returns 404 if file not found or not owned by user.
     * Returns 422 if AI is not configured, fails, or returns an operation not in the allowed catalog.
     */
    @PostMapping(
            value = "/{id}/workflow-proposal",
            consumes = MediaType.APPLICATION_JSON_VALUE
    )
    public ResponseEntity<WorkflowProposalResponse> proposeWorkflow(
            @PathVariable Long id,
            @Valid @RequestBody WorkflowProposalRequest request,
            @AuthenticationPrincipal User currentUser
    ) {
        WorkflowProposalResponse response = workflowProposalService.proposeWorkflow(
                id, currentUser, request.getInstruction()
        );
        return ResponseEntity.ok(response);
    }

    /**
     * POST /api/files/{id}/workflow-proposal/validate
     * Deterministic validation of a proposed workflow before user approval or execution.
     * Body: { "instruction": "...", "workflowName": "...", "description": "...", "steps": [...] }
     * Returns 200 with { valid: true/false, errors: [...] }.
     * Returns 400 if request body is invalid.
     * Returns 404 if file not found or not owned by user.
     * Returns 422 if the workbook cannot be read.
     * No AI calls are made. No workflow is executed.
     */
    @PostMapping(
            value = "/{id}/workflow-proposal/validate",
            consumes = MediaType.APPLICATION_JSON_VALUE
    )
    public ResponseEntity<WorkflowValidationResponse> validateWorkflowProposal(
            @PathVariable Long id,
            @Valid @RequestBody WorkflowValidationRequest request,
            @AuthenticationPrincipal User currentUser
    ) {
        WorkflowValidationResponse response = workflowValidatorService.validateWorkflow(id, currentUser, request);
        return ResponseEntity.ok(response);
    }

    /**
     * POST /api/files/{id}/workflow-proposal/approve
     * Approve a workflow proposal. The proposal is validated first; only valid
     * proposals in PROPOSED state can be approved. The workflow is NOT executed.
     * Body: { "instruction": "...", "workflowName": "...", "description": "...", "steps": [...] }
     * Returns 200 with the new status on success.
     * Returns 400 if the request body is invalid.
     * Returns 404 if the file is not found or not owned by the user.
     * Returns 409 if the proposal fails validation, or is already approved/rejected.
     */
    @PostMapping(
            value = "/{id}/workflow-proposal/approve",
            consumes = MediaType.APPLICATION_JSON_VALUE
    )
    public ResponseEntity<WorkflowApprovalResponse> approveWorkflowProposal(
            @PathVariable Long id,
            @Valid @RequestBody WorkflowValidationRequest request,
            @AuthenticationPrincipal User currentUser
    ) {
        WorkflowApprovalResponse response = workflowApprovalService.approve(id, currentUser, request);
        return ResponseEntity.ok(response);
    }

    /**
     * POST /api/files/{id}/workflow-proposal/reject
     * Reject a workflow proposal. Transitions the proposal to REJECTED (terminal).
     * The workflow is NOT executed.
     * Body: { "instruction": "...", "workflowName": "...", "description": "...", "steps": [...] }
     * Returns 200 with the new status on success.
     * Returns 400 if the request body is invalid.
     * Returns 404 if the file is not found or not owned by the user.
     * Returns 409 if the proposal is already rejected.
     */
    @PostMapping(
            value = "/{id}/workflow-proposal/reject",
            consumes = MediaType.APPLICATION_JSON_VALUE
    )
    public ResponseEntity<WorkflowApprovalResponse> rejectWorkflowProposal(
            @PathVariable Long id,
            @Valid @RequestBody WorkflowValidationRequest request,
            @AuthenticationPrincipal User currentUser
    ) {
        WorkflowApprovalResponse response = workflowApprovalService.reject(id, currentUser, request);
        return ResponseEntity.ok(response);
    }

    /**
     * POST /api/files/{id}/workflow-proposal/supersede
     * Supersedes an existing approved (or proposed) workflow with a new proposal.
     * Overwrites the persisted workflow data with the new proposal and resets the
     * status to PROPOSED, allowing the new workflow to be validated and approved.
     * The old approved workflow transitions to SUPERSEDED (terminal).
     * Body: { "instruction": "...", "workflowName": "...", "description": "...", "steps": [...] }
     * Returns 200 with the new status on success.
     * Returns 400 if the request body is invalid.
     * Returns 404 if the file is not found or not owned by the user.
     * Returns 409 if the current workflow is already superseded or rejected (terminal).
     */
    @PostMapping(
            value = "/{id}/workflow-proposal/supersede",
            consumes = MediaType.APPLICATION_JSON_VALUE
    )
    public ResponseEntity<WorkflowApprovalResponse> supersedeWorkflowProposal(
            @PathVariable Long id,
            @Valid @RequestBody WorkflowValidationRequest request,
            @AuthenticationPrincipal User currentUser
    ) {
        WorkflowApprovalResponse response = workflowApprovalService.supersede(id, currentUser, request);
        return ResponseEntity.ok(response);
    }

    /**
     * POST /api/files/{id}/workflow/execute
     * Executes the approved workflow for the given file.
     * The workflow must be in APPROVED status; PROPOSED and REJECTED are rejected.
     * The original workbook is never modified; a new output file is created.
     * Returns 200 with the execution result (success or failure details).
     * Returns 404 if the file is not found or not owned by the user.
     * Returns 409 if the workflow is not approved or execution fails.
     */
    @PostMapping("/{id}/workflow/execute")
    public ResponseEntity<WorkflowExecutionResponse> executeWorkflow(
            @PathVariable Long id,
            @AuthenticationPrincipal User currentUser
    ) {
        WorkflowExecutionResponse response = workflowExecutionService.executeWorkflow(id, currentUser);
        return ResponseEntity.ok(response);
    }

    /**
     * GET /api/files/{id}/output
     * Downloads the generated workflow output file for the given uploaded file ID.
     * The file is served as an attachment with the original filename suffixed by "_output".
     * Returns 404 if no workflow output exists for this file.
     * Returns 403/404 if the file is not owned by the current user.
     */
    @GetMapping("/{id}/output")
    public ResponseEntity<byte[]> downloadOutput(
            @PathVariable Long id,
            @AuthenticationPrincipal User currentUser
    ) throws java.io.IOException {
        FileService.OutputFileResult result = fileService.getOutputFile(id, currentUser);
        Path filePath = result.filePath();
        String downloadName = sanitizeFilename(result.downloadFileName());

        String contentType;
        if (downloadName.toLowerCase().endsWith(".xls")
                && !downloadName.toLowerCase().endsWith(".xlsx")) {
            contentType = "application/vnd.ms-excel";
        } else {
            contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        }

        byte[] body = java.nio.file.Files.readAllBytes(filePath);
        return ResponseEntity.ok()
                .contentType(MediaType.parseMediaType(contentType))
                .header("Content-Disposition",
                        "attachment; filename=\"" + downloadName + "\"")
                .body(body);
    }

    /**
     * GET /api/files/{id}/output/pdf
     * Downloads the PDF generated by the GENERATE_PDF workflow operation.
     * The file is served as an attachment with the original filename suffixed by "_output.pdf".
     * Returns 404 if no PDF output exists for this file.
     * Returns 403/404 if the file is not owned by the current user.
     */
    @GetMapping("/{id}/output/pdf")
    public ResponseEntity<byte[]> downloadOutputPdf(
            @PathVariable Long id,
            @AuthenticationPrincipal User currentUser
    ) throws java.io.IOException {
        FileService.OutputFileResult result = fileService.getOutputPdfFile(id, currentUser);
        Path filePath = result.filePath();
        String downloadName = sanitizeFilename(result.downloadFileName());

        byte[] body = java.nio.file.Files.readAllBytes(filePath);
        return ResponseEntity.ok()
                .contentType(MediaType.APPLICATION_PDF)
                .header("Content-Disposition",
                        "attachment; filename=\"" + downloadName + "\"")
                .body(body);
    }

    /**
     * Removes characters that are illegal in a Content-Disposition filename token
     * (specifically double-quote and control characters). The filename originates
     * server-side from the database, not user input, but this guard is still cheap
     * and correct.
     */
    private String sanitizeFilename(String name) {
        if (name == null) return "output.xlsx";
        return name.replaceAll("[\\x00-\\x1F\\x7F\"\\\\]", "_");
    }
}
