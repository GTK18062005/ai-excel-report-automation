package com.aiexcel.platform.service;

import com.aiexcel.platform.dto.*;
import com.aiexcel.platform.entity.*;
import com.aiexcel.platform.exception.FileNotFoundException;
import com.aiexcel.platform.exception.WorkflowApprovalException;
import com.aiexcel.platform.repository.UploadedFileRepository;
import com.aiexcel.platform.repository.WorkflowApprovalRepository;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.List;

/**
 * Workflow approval state machine.
 * <p>
 * Manages the lifecycle of a proposed workflow: PROPOSED → APPROVED | REJECTED.
 * Only one proposal is tracked per (file, user) pair.
 * <p>
 * Approval rules:
 * <ul>
 *   <li>Only workflows in {@link WorkflowStatus#PROPOSED} state can be approved.</li>
 *   <li>The proposal must pass full validation before it can be approved.</li>
 *   <li>Approval is irreversible (APPROVED → REJECTED is the only allowed transition).</li>
 * </ul>
 * <p>
 * Rejection rules:
 * <ul>
 *   <li>Workflows in {@link WorkflowStatus#PROPOSED} or {@link WorkflowStatus#APPROVED}
 *       state can be rejected.</li>
 *   <li>{@link WorkflowStatus#REJECTED} is terminal — no further transitions are allowed.</li>
 * </ul>
 */
@Service
@RequiredArgsConstructor
@Slf4j
public class WorkflowApprovalService {

    private final WorkflowApprovalRepository approvalRepository;
    private final UploadedFileRepository fileRepository;
    private final WorkflowValidatorService validatorService;

    /**
     * Approves a workflow proposal.
     *
     * @param fileId  ID of the uploaded file (ownership is verified)
     * @param user    authenticated user (must own the file)
     * @param request the workflow proposal data
     * @return approval response with current status
     * @throws FileNotFoundException       if the file does not exist or is not owned by the user
     * @throws WorkflowApprovalException   if the proposal is invalid or the state transition is illegal
     */
    @Transactional
    public WorkflowApprovalResponse approve(Long fileId, User user, WorkflowValidationRequest request) {
        log.info("Approving workflow for fileId={}, userId={}", fileId, user.getId());

        // 1. Verify file ownership
        UploadedFile file = resolveOwnedFile(fileId, user);

        // 2. Get or create the approval record (in PROPOSED state)
        WorkflowApproval approval = getOrCreateApproval(file, user, request);

        // 3. Enforce state machine: only PROPOSED can be approved
        if (approval.getStatus() == WorkflowStatus.APPROVED) {
            throw new WorkflowApprovalException(
                    "Workflow is already approved and cannot be approved again");
        }
        if (approval.getStatus() == WorkflowStatus.REJECTED) {
            throw new WorkflowApprovalException(
                    "Workflow has been rejected and cannot be approved again");
        }

        // 4. Run full validation
        WorkflowValidationRequest validationRequest = buildValidationRequest(request);
        WorkflowValidationResponse validation = validatorService.validateWorkflow(fileId, user, validationRequest);

        if (!validation.isValid()) {
            throw new WorkflowApprovalException(
                    "Workflow cannot be approved because it failed validation: " +
                    String.join("; ", validation.getErrors()));
        }

        // 5. Transition to APPROVED
        approval.setStatus(WorkflowStatus.APPROVED);
        approval.setWorkflowName(request.getWorkflowName());
        approval.setDescription(request.getDescription());
        approval.setStepsFromJson(request.getSteps());

        WorkflowApproval saved = approvalRepository.save(approval);
        log.info("Workflow approved for fileId={}, userId={}", fileId, user.getId());

        return WorkflowApprovalResponse.builder()
                .status(saved.getStatus())
                .workflowName(saved.getWorkflowName())
                .message("Workflow approved successfully and is ready for execution")
                .build();
    }

    /**
     * Rejects a workflow proposal.
     *
     * @param fileId  ID of the uploaded file (ownership is verified)
     * @param user    authenticated user (must own the file)
     * @param request the workflow proposal data (used to get or create the record)
     * @return approval response with current status
     * @throws FileNotFoundException     if the file does not exist or is not owned by the user
     * @throws WorkflowApprovalException if the workflow is already rejected
     */
    @Transactional
    public WorkflowApprovalResponse reject(Long fileId, User user, WorkflowValidationRequest request) {
        log.info("Rejecting workflow for fileId={}, userId={}", fileId, user.getId());

        // 1. Verify file ownership
        UploadedFile file = resolveOwnedFile(fileId, user);

        // 2. Get or create the approval record
        WorkflowApproval approval = getOrCreateApproval(file, user, request);

        // 3. Enforce state machine: REJECTED is terminal
        if (approval.getStatus() == WorkflowStatus.REJECTED) {
            throw new WorkflowApprovalException(
                    "Workflow has already been rejected");
        }

        // 4. Transition to REJECTED
        approval.setStatus(WorkflowStatus.REJECTED);
        approval.setWorkflowName(request.getWorkflowName());
        approval.setDescription(request.getDescription());
        approval.setStepsFromJson(request.getSteps());

        WorkflowApproval saved = approvalRepository.save(approval);
        log.info("Workflow rejected for fileId={}, userId={}", fileId, user.getId());

        return WorkflowApprovalResponse.builder()
                .status(saved.getStatus())
                .workflowName(saved.getWorkflowName())
                .message("Workflow rejected")
                .build();
    }

    /**
     * Supersedes an existing approved (or proposed) workflow with a new proposal.
     * <p>
     * Replaces the persisted workflow data (workflowName, description, stepsJson)
     * with the new proposal and resets the status to {@link WorkflowStatus#PROPOSED},
     * allowing the new workflow to go through validation and approval.
     * <p>
     * The old approved workflow becomes {@link WorkflowStatus#SUPERSEDED} — a terminal
     * state that prevents accidental re-execution. Once superseded, the row cannot
     * be approved again or be executed.
     * <p>
     * This is the intended path for iterating on workflows for the same file:
     * propose → validate → supersede → approve → execute.
     *
     * @param fileId  ID of the uploaded file (ownership is verified)
     * @param user    authenticated user (must own the file)
     * @param request the new workflow proposal data
     * @return approval response with current status (PROPOSED)
     * @throws FileNotFoundException       if the file does not exist or is not owned by the user
     * @throws WorkflowApprovalException   if the current workflow is already superseded/rejected
     */
    @Transactional
    public WorkflowApprovalResponse supersede(Long fileId, User user, WorkflowValidationRequest request) {
        log.info("Superseding workflow for fileId={}, userId={}", fileId, user.getId());

        // 1. Verify file ownership
        UploadedFile file = resolveOwnedFile(fileId, user);

        // 2. Get or create the approval record
        WorkflowApproval approval = getOrCreateApproval(file, user, request);

        // 3. SUPERSEDED and REJECTED are terminal — cannot be superseded
        if (approval.getStatus() == WorkflowStatus.SUPERSEDED) {
            throw new WorkflowApprovalException(
                    "Workflow has already been superseded");
        }
        if (approval.getStatus() == WorkflowStatus.REJECTED) {
            throw new WorkflowApprovalException(
                    "Workflow has been rejected and cannot be superseded");
        }

        // 4. Overwrite workflow data with the new proposal and reset to PROPOSED
        approval.setStatus(WorkflowStatus.PROPOSED);
        approval.setWorkflowName(request.getWorkflowName());
        approval.setDescription(request.getDescription());
        approval.setStepsFromJson(request.getSteps());

        WorkflowApproval saved = approvalRepository.save(approval);
        log.info("Workflow superseded for fileId={}, userId={}. New workflow '{}' is now in PROPOSED state.",
                fileId, user.getId(), saved.getWorkflowName());

        return WorkflowApprovalResponse.builder()
                .status(saved.getStatus())
                .workflowName(saved.getWorkflowName())
                .message("Workflow superseded. New proposal is ready for validation and approval.")
                .build();
    }

    // -------------------------------------------------------------------------
    // Private helpers
    // -------------------------------------------------------------------------

    private UploadedFile resolveOwnedFile(Long fileId, User user) {
        return fileRepository
                .findByIdAndUserAndStatusNot(fileId, user, FileStatus.DELETED)
                .orElseThrow(() -> new FileNotFoundException("File not found with id: " + fileId));
    }

    /**
     * Returns the existing approval record for (file, user) or creates a new one
     * in {@link WorkflowStatus#PROPOSED} state.
     * <p>
     * If a record exists, its workflow data (name, description, steps) is updated
     * to match the incoming request before the state transition is applied.
     */
    private WorkflowApproval getOrCreateApproval(UploadedFile file, User user, WorkflowValidationRequest request) {
        return approvalRepository.findByFileAndUser(file, user)
                .orElseGet(() -> {
                    WorkflowApproval wa = WorkflowApproval.builder()
                            .file(file)
                            .user(user)
                            .status(WorkflowStatus.PROPOSED)
                            .workflowName(request.getWorkflowName())
                            .description(request.getDescription())
                            .build();
                    wa.setStepsFromJson(request.getSteps());
                    return wa;
                });
    }

    /**
     * Builds a validation request from the approval request.
     * Reuses the same steps — validation reads the same fields.
     */
    private WorkflowValidationRequest buildValidationRequest(WorkflowValidationRequest request) {
        return WorkflowValidationRequest.builder()
                .instruction(request.getInstruction())
                .workflowName(request.getWorkflowName())
                .description(request.getDescription())
                .steps(request.getSteps())
                .build();
    }
}
