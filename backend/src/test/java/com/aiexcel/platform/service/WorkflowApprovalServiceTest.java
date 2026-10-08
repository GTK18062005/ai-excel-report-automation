package com.aiexcel.platform.service;

import com.aiexcel.platform.dto.*;
import com.aiexcel.platform.entity.*;
import com.aiexcel.platform.exception.FileNotFoundException;
import com.aiexcel.platform.exception.WorkflowApprovalException;
import com.aiexcel.platform.repository.UploadedFileRepository;
import com.aiexcel.platform.repository.WorkflowApprovalRepository;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.junit.jupiter.api.io.TempDir;
import org.mockito.ArgumentCaptor;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.nio.file.Path;
import java.util.*;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

/**
 * Unit tests for {@link WorkflowApprovalService}.
 * Covers: approve valid, approve invalid, approve rejected, approve approved,
 * reject proposal, reject approved, wrong owner, nonexistent file.
 */
@ExtendWith(MockitoExtension.class)
class WorkflowApprovalServiceTest {

    @Mock
    private WorkflowApprovalRepository approvalRepository;

    @Mock
    private UploadedFileRepository fileRepository;

    @Mock
    private WorkflowValidatorService validatorService;

    @TempDir
    Path tempDir;

    private WorkflowApprovalService service;
    private User testUser;
    private UploadedFile testFile;
    private Path xlsxPath;

    @BeforeEach
    void setUp() throws Exception {
        service = new WorkflowApprovalService(approvalRepository, fileRepository, validatorService);

        testUser = User.builder()
                .id(1L)
                .name("Test User")
                .email("test@example.com")
                .password("password")
                .build();

        xlsxPath = createTestWorkbook();

        testFile = UploadedFile.builder()
                .id(1L)
                .user(testUser)
                .originalFileName("sales.xlsx")
                .storedFileName("sales.xlsx")
                .fileType("xlsx")
                .fileSize(1234L)
                .storagePath(xlsxPath.toString())
                .status(FileStatus.UPLOADED)
                .build();
    }

    // -------------------------------------------------------------------------
    // Approve tests
    // -------------------------------------------------------------------------

    @Test
    void approve_validProposal_approvesAndReturnsApproved() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));
        when(approvalRepository.findByFileAndUser(eq(testFile), eq(testUser)))
                .thenReturn(Optional.empty()); // no existing record
        when(validatorService.validateWorkflow(eq(1L), eq(testUser), any()))
                .thenReturn(WorkflowValidationResponse.builder().valid(true).errors(List.of()).build());
        when(approvalRepository.save(any())).thenAnswer(inv -> inv.getArgument(0));

        WorkflowValidationRequest request = buildValidRequest();

        WorkflowApprovalResponse response = service.approve(1L, testUser, request);

        assertEquals(WorkflowStatus.APPROVED, response.getStatus());
        assertEquals("Workflow approved successfully and is ready for execution", response.getMessage());

        ArgumentCaptor<WorkflowApproval> captor = ArgumentCaptor.forClass(WorkflowApproval.class);
        verify(approvalRepository).save(captor.capture());
        assertEquals(WorkflowStatus.APPROVED, captor.getValue().getStatus());
    }

    @Test
    void approve_invalidProposal_throwsConflict() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));
        when(approvalRepository.findByFileAndUser(eq(testFile), eq(testUser)))
                .thenReturn(Optional.empty());
        when(validatorService.validateWorkflow(eq(1L), eq(testUser), any()))
                .thenReturn(WorkflowValidationResponse.builder()
                        .valid(false)
                        .errors(List.of("Step 1: unknown operation 'RUN_PYTHON'"))
                        .build());

        WorkflowValidationRequest request = buildValidRequest();

        WorkflowApprovalException ex = assertThrows(WorkflowApprovalException.class,
                () -> service.approve(1L, testUser, request));

        assertTrue(ex.getMessage().contains("RUN_PYTHON"));
        assertTrue(ex.getMessage().contains("failed validation"));
        verify(approvalRepository, never()).save(any()); // not saved on failed validation
    }

    @Test
    void approve_alreadyApproved_throwsConflict() {
        WorkflowApproval existing = WorkflowApproval.builder()
                .id(5L)
                .file(testFile)
                .user(testUser)
                .status(WorkflowStatus.APPROVED)
                .workflowName("My Workflow")
                .build();

        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));
        when(approvalRepository.findByFileAndUser(eq(testFile), eq(testUser)))
                .thenReturn(Optional.of(existing));

        WorkflowApprovalException ex = assertThrows(WorkflowApprovalException.class,
                () -> service.approve(1L, testUser, buildValidRequest()));

        assertTrue(ex.getMessage().contains("already approved"));
        verify(approvalRepository, never()).save(any());
    }

    @Test
    void approve_alreadyRejected_throwsConflict() {
        WorkflowApproval existing = WorkflowApproval.builder()
                .id(5L)
                .file(testFile)
                .user(testUser)
                .status(WorkflowStatus.REJECTED)
                .workflowName("My Workflow")
                .build();

        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));
        when(approvalRepository.findByFileAndUser(eq(testFile), eq(testUser)))
                .thenReturn(Optional.of(existing));

        WorkflowApprovalException ex = assertThrows(WorkflowApprovalException.class,
                () -> service.approve(1L, testUser, buildValidRequest()));

        assertTrue(ex.getMessage().contains("rejected"));
        verify(approvalRepository, never()).save(any());
    }

    @Test
    void approve_wrongOwner_throwsFileNotFound() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(999L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.empty());

        assertThrows(FileNotFoundException.class,
                () -> service.approve(999L, testUser, buildValidRequest()));

        verify(approvalRepository, never()).findByFileAndUser(any(), any());
    }

    // -------------------------------------------------------------------------
    // Reject tests
    // -------------------------------------------------------------------------

    @Test
    void reject_proposedWorkflow_rejectsAndReturnsRejected() {
        WorkflowApproval existing = WorkflowApproval.builder()
                .id(5L)
                .file(testFile)
                .user(testUser)
                .status(WorkflowStatus.PROPOSED)
                .workflowName("My Workflow")
                .build();

        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));
        when(approvalRepository.findByFileAndUser(eq(testFile), eq(testUser)))
                .thenReturn(Optional.of(existing));
        when(approvalRepository.save(any())).thenAnswer(inv -> inv.getArgument(0));

        WorkflowValidationRequest request = buildValidRequest();

        WorkflowApprovalResponse response = service.reject(1L, testUser, request);

        assertEquals(WorkflowStatus.REJECTED, response.getStatus());
        assertEquals("Workflow rejected", response.getMessage());

        ArgumentCaptor<WorkflowApproval> captor = ArgumentCaptor.forClass(WorkflowApproval.class);
        verify(approvalRepository).save(captor.capture());
        assertEquals(WorkflowStatus.REJECTED, captor.getValue().getStatus());
    }

    @Test
    void reject_approvedWorkflow_rejectsAndReturnsRejected() {
        WorkflowApproval existing = WorkflowApproval.builder()
                .id(5L)
                .file(testFile)
                .user(testUser)
                .status(WorkflowStatus.APPROVED)
                .workflowName("My Workflow")
                .build();

        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));
        when(approvalRepository.findByFileAndUser(eq(testFile), eq(testUser)))
                .thenReturn(Optional.of(existing));
        when(approvalRepository.save(any())).thenAnswer(inv -> inv.getArgument(0));

        WorkflowValidationRequest request = buildValidRequest();

        WorkflowApprovalResponse response = service.reject(1L, testUser, request);

        assertEquals(WorkflowStatus.REJECTED, response.getStatus());
        assertEquals("Workflow rejected", response.getMessage());
    }

    @Test
    void reject_alreadyRejected_throwsConflict() {
        WorkflowApproval existing = WorkflowApproval.builder()
                .id(5L)
                .file(testFile)
                .user(testUser)
                .status(WorkflowStatus.REJECTED)
                .workflowName("My Workflow")
                .build();

        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));
        when(approvalRepository.findByFileAndUser(eq(testFile), eq(testUser)))
                .thenReturn(Optional.of(existing));

        WorkflowApprovalException ex = assertThrows(WorkflowApprovalException.class,
                () -> service.reject(1L, testUser, buildValidRequest()));

        assertTrue(ex.getMessage().contains("already been rejected"));
        verify(approvalRepository, never()).save(any());
    }

    @Test
    void reject_nonexistentFile_throwsFileNotFound() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(999L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.empty());

        assertThrows(FileNotFoundException.class,
                () -> service.reject(999L, testUser, buildValidRequest()));
    }

    @Test
    void reject_newProposal_createsAndRejects() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));
        when(approvalRepository.findByFileAndUser(eq(testFile), eq(testUser)))
                .thenReturn(Optional.empty()); // no existing record
        when(approvalRepository.save(any())).thenAnswer(inv -> inv.getArgument(0));

        WorkflowValidationRequest request = buildValidRequest();

        WorkflowApprovalResponse response = service.reject(1L, testUser, request);

        assertEquals(WorkflowStatus.REJECTED, response.getStatus());

        ArgumentCaptor<WorkflowApproval> captor = ArgumentCaptor.forClass(WorkflowApproval.class);
        verify(approvalRepository).save(captor.capture());
        assertEquals(WorkflowStatus.REJECTED, captor.getValue().getStatus());
    }

    // -------------------------------------------------------------------------
    // Supersede tests
    // -------------------------------------------------------------------------

    @Test
    void supersede_approvedWorkflow_replacesWorkflowDataAndResetsToProposed() {
        // Existing APPROVED workflow with old data
        WorkflowApproval existing = WorkflowApproval.builder()
                .id(5L)
                .file(testFile)
                .user(testUser)
                .status(WorkflowStatus.APPROVED)
                .workflowName("Old Workflow")
                .description("Old description")
                .build();
        existing.setStepsFromJson(List.of(
                WorkflowStep.builder().order(1).operation("EXPORT_EXCEL").build()));

        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));
        when(approvalRepository.findByFileAndUser(eq(testFile), eq(testUser)))
                .thenReturn(Optional.of(existing));
        when(approvalRepository.save(any())).thenAnswer(inv -> inv.getArgument(0));

        WorkflowValidationRequest newRequest = WorkflowValidationRequest.builder()
                .instruction("New instruction")
                .workflowName("New Pivot Workflow")
                .description("New pivot description")
                .steps(List.of(
                        WorkflowStep.builder()
                                .order(1).operation("PIVOT")
                                .description("New pivot step")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of("Region"))
                                .parameters(Map.of())
                                .build()))
                .build();

        WorkflowApprovalResponse response = service.supersede(1L, testUser, newRequest);

        assertEquals(WorkflowStatus.PROPOSED, response.getStatus());
        assertEquals("New Pivot Workflow", response.getWorkflowName());
        assertTrue(response.getMessage().contains("superseded"));

        ArgumentCaptor<WorkflowApproval> captor = ArgumentCaptor.forClass(WorkflowApproval.class);
        verify(approvalRepository).save(captor.capture());
        WorkflowApproval saved = captor.getValue();

        assertEquals(WorkflowStatus.PROPOSED, saved.getStatus());
        assertEquals("New Pivot Workflow", saved.getWorkflowName());
        assertEquals("New pivot description", saved.getDescription());
        // Verify steps were replaced
        List<WorkflowStep> savedSteps = saved.getStepsFromJson();
        assertEquals(1, savedSteps.size());
        assertEquals("PIVOT", savedSteps.get(0).getOperation());
    }

    @Test
    void supersede_proposedWorkflow_replacesWorkflowDataAndResetsToProposed() {
        WorkflowApproval existing = WorkflowApproval.builder()
                .id(5L)
                .file(testFile)
                .user(testUser)
                .status(WorkflowStatus.PROPOSED)
                .workflowName("Old Proposed")
                .build();

        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));
        when(approvalRepository.findByFileAndUser(eq(testFile), eq(testUser)))
                .thenReturn(Optional.of(existing));
        when(approvalRepository.save(any())).thenAnswer(inv -> inv.getArgument(0));

        WorkflowValidationRequest newRequest = WorkflowValidationRequest.builder()
                .workflowName("New Workflow")
                .description("New description")
                .steps(List.of(WorkflowStep.builder()
                        .order(1).operation("READ_SHEET")
                        .inputSheets(List.of("Data"))
                        .inputColumns(List.of())
                        .parameters(Map.of())
                        .build()))
                .build();

        WorkflowApprovalResponse response = service.supersede(1L, testUser, newRequest);

        assertEquals(WorkflowStatus.PROPOSED, response.getStatus());

        ArgumentCaptor<WorkflowApproval> captor = ArgumentCaptor.forClass(WorkflowApproval.class);
        verify(approvalRepository).save(captor.capture());
        assertEquals("New Workflow", captor.getValue().getWorkflowName());
    }

    @Test
    void supersede_alreadySuperseded_throwsConflict() {
        WorkflowApproval existing = WorkflowApproval.builder()
                .id(5L)
                .file(testFile)
                .user(testUser)
                .status(WorkflowStatus.SUPERSEDED)
                .workflowName("Old Workflow")
                .build();

        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));
        when(approvalRepository.findByFileAndUser(eq(testFile), eq(testUser)))
                .thenReturn(Optional.of(existing));

        WorkflowApprovalException ex = assertThrows(WorkflowApprovalException.class,
                () -> service.supersede(1L, testUser, buildValidRequest()));

        assertTrue(ex.getMessage().contains("already been superseded"));
        verify(approvalRepository, never()).save(any());
    }

    @Test
    void supersede_alreadyRejected_throwsConflict() {
        WorkflowApproval existing = WorkflowApproval.builder()
                .id(5L)
                .file(testFile)
                .user(testUser)
                .status(WorkflowStatus.REJECTED)
                .workflowName("Old Workflow")
                .build();

        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));
        when(approvalRepository.findByFileAndUser(eq(testFile), eq(testUser)))
                .thenReturn(Optional.of(existing));

        WorkflowApprovalException ex = assertThrows(WorkflowApprovalException.class,
                () -> service.supersede(1L, testUser, buildValidRequest()));

        assertTrue(ex.getMessage().contains("rejected"));
        verify(approvalRepository, never()).save(any());
    }

    @Test
    void supersede_wrongOwner_throwsFileNotFound() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(999L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.empty());

        assertThrows(FileNotFoundException.class,
                () -> service.supersede(999L, testUser, buildValidRequest()));

        verify(approvalRepository, never()).findByFileAndUser(any(), any());
    }

    @Test
    void supersede_noExistingRecord_createsNewProposedRecord() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));
        when(approvalRepository.findByFileAndUser(eq(testFile), eq(testUser)))
                .thenReturn(Optional.empty()); // no prior record
        when(approvalRepository.save(any())).thenAnswer(inv -> inv.getArgument(0));

        WorkflowValidationRequest newRequest = WorkflowValidationRequest.builder()
                .workflowName("Fresh Proposal")
                .description("A fresh proposal with no prior record")
                .steps(List.of(WorkflowStep.builder()
                        .order(1).operation("READ_SHEET")
                        .inputSheets(List.of("Sheet1"))
                        .inputColumns(List.of())
                        .parameters(Map.of())
                        .build()))
                .build();

        WorkflowApprovalResponse response = service.supersede(1L, testUser, newRequest);

        assertEquals(WorkflowStatus.PROPOSED, response.getStatus());
        assertEquals("Fresh Proposal", response.getWorkflowName());

        ArgumentCaptor<WorkflowApproval> captor = ArgumentCaptor.forClass(WorkflowApproval.class);
        verify(approvalRepository).save(captor.capture());
        assertEquals(WorkflowStatus.PROPOSED, captor.getValue().getStatus());
        assertEquals("Fresh Proposal", captor.getValue().getWorkflowName());
    }

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    private WorkflowValidationRequest buildValidRequest() {
        return WorkflowValidationRequest.builder()
                .instruction("Compare actual vs target")
                .workflowName("Regional Variance")
                .description("Compute variance")
                .steps(List.of(
                        WorkflowStep.builder()
                                .order(1)
                                .operation("READ_SHEET")
                                .description("Load Sales")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build()
                ))
                .build();
    }

    private Path createTestWorkbook() throws Exception {
        org.apache.poi.xssf.usermodel.XSSFWorkbook wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook();
        org.apache.poi.ss.usermodel.Sheet sheet = wb.createSheet("Sales");

        org.apache.poi.ss.usermodel.Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Date");
        header.createCell(1).setCellValue("Region");
        header.createCell(2).setCellValue("Actual");
        header.createCell(3).setCellValue("Target");

        java.io.ByteArrayOutputStream baos = new java.io.ByteArrayOutputStream();
        wb.write(baos);
        wb.close();

        Path filePath = tempDir.resolve("sales.xlsx");
        java.nio.file.Files.write(filePath, baos.toByteArray());
        return filePath;
    }
}
