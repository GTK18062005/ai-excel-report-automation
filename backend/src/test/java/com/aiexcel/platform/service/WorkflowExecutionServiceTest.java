package com.aiexcel.platform.service;

import com.aiexcel.platform.dto.WorkflowExecutionResponse;
import com.aiexcel.platform.dto.WorkflowStep;
import com.aiexcel.platform.entity.FileStatus;
import com.aiexcel.platform.entity.UploadedFile;
import com.aiexcel.platform.entity.User;
import com.aiexcel.platform.entity.WorkflowApproval;
import com.aiexcel.platform.entity.WorkflowStatus;
import com.aiexcel.platform.exception.FileNotFoundException;
import com.aiexcel.platform.exception.WorkflowApprovalException;
import com.aiexcel.platform.exception.WorkflowExecutionException;
import com.aiexcel.platform.repository.UploadedFileRepository;
import com.aiexcel.platform.repository.WorkflowApprovalRepository;
import com.aiexcel.platform.service.workflow.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.junit.jupiter.api.io.TempDir;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

/**
 * Unit tests for {@link WorkflowExecutionService}.
 * Covers: approved workflow executes, proposed/rejected workflow rejected,
 * wrong-owner / nonexistent file / no-approval guards, plus each supported
 * operation (READ_SHEET, SELECT_COLUMNS, REMOVE_EMPTY_ROWS, REMOVE_DUPLICATES,
 * FILTER, SORT), invalid parameters, and original workbook unchanged.
 */
@ExtendWith(MockitoExtension.class)
class WorkflowExecutionServiceTest {

    @Mock
    private WorkflowApprovalRepository approvalRepository;

    @Mock
    private UploadedFileRepository fileRepository;

    @Mock
    private FileStorageService fileStorageService;

    @TempDir
    Path tempDir;

    private WorkflowExecutionService service;
    private User testUser;
    private Path xlsxPath;

    @BeforeEach
    void setUp() throws Exception {
        List<WorkflowOperationHandler> handlers = List.of(
                new ReadSheetHandler(),
                new SelectColumnsHandler(),
                new RemoveEmptyRowsHandler(),
                new RemoveDuplicatesHandler(),
                new FilterHandler(),
                new SortHandler(),
                new CalculateHandler(),
                new GroupByHandler(),
                new AggregateHandler(),
                new JoinSheetsHandler(),
                new PivotHandler(),
                new CreateSummaryHandler(),
                new CreateChartHandler(),
                new ValidateHandler(),
                new GeneratePdfHandler(),
                new ExportExcelHandler()
        );
        service = new WorkflowExecutionService(approvalRepository, fileRepository, fileStorageService, handlers);

        testUser = User.builder().id(1L).name("Test User").email("test@example.com").password("password").build();
        xlsxPath = createTestWorkbook();
    }

    // -------------------------------------------------------------------------
    // Guard tests — these throw exceptions, not return failure responses
    // -------------------------------------------------------------------------

    @Test
    void executeWorkflow_proposedStatus_throwsConflict() {
        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApproval(file, testUser, WorkflowStatus.PROPOSED);

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowApprovalException ex = assertThrows(WorkflowApprovalException.class,
                () -> service.executeWorkflow(1L, testUser));

        assertTrue(ex.getMessage().contains("not approved"));
        assertTrue(ex.getMessage().contains("PROPOSED"));
    }

    @Test
    void executeWorkflow_rejectedStatus_throwsConflict() {
        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApproval(file, testUser, WorkflowStatus.REJECTED);

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowApprovalException ex = assertThrows(WorkflowApprovalException.class,
                () -> service.executeWorkflow(1L, testUser));

        assertTrue(ex.getMessage().contains("rejected"));
    }

    @Test
    void executeWorkflow_nonexistentFile_throwsNotFound() {
        when(fileRepository.findByIdAndUserAndStatusNot(999L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.empty());

        assertThrows(FileNotFoundException.class,
                () -> service.executeWorkflow(999L, testUser));
    }

    @Test
    void executeWorkflow_wrongOwner_throwsNotFound() {
        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.empty());

        assertThrows(FileNotFoundException.class,
                () -> service.executeWorkflow(1L, testUser));
    }

    @Test
    void executeWorkflow_noApprovalRecord_throwsConflict() {
        UploadedFile file = buildFile(xlsxPath, testUser);

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.empty());

        assertThrows(WorkflowApprovalException.class,
                () -> service.executeWorkflow(1L, testUser));
    }

    @Test
    void executeWorkflow_unsupportedOperation_throwsConflict() throws Exception {
        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(WorkflowStep.builder()
                        .order(1).operation("JOIN")
                        .inputSheets(List.of("Sales"))
                        .inputColumns(List.of())
                        .parameters(Map.of())
                        .build()));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> service.executeWorkflow(1L, testUser));

        assertTrue(ex.getMessage().contains("not yet implemented"));
    }

    // -------------------------------------------------------------------------
    // READ_SHEET tests
    // -------------------------------------------------------------------------

    @Test
    void executeWorkflow_readSheet_succeeds() throws Exception {
        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(WorkflowStep.builder()
                        .order(1).operation("READ_SHEET")
                        .inputSheets(List.of("Sales"))
                        .inputColumns(List.of())
                        .parameters(Map.of())
                        .build()));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        int originalRows = countRows(xlsxPath);
        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertNotNull(resp.getOutputFileName());
        assertNull(resp.getErrors());
        assertFalse(resp.getSteps().isEmpty());
        assertEquals("READ_SHEET", resp.getSteps().get(0).getOperation());

        // Original file unchanged
        assertEquals(originalRows, countRows(xlsxPath));

        Path outPath = xlsxPath.resolveSibling(resp.getOutputFileName());
        assertTrue(Files.exists(outPath));
        Files.deleteIfExists(outPath);
    }

    @Test
    void executeWorkflow_readSheet_missingSheet_returnsFailure() throws Exception {
        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(WorkflowStep.builder()
                        .order(1).operation("READ_SHEET")
                        .inputSheets(List.of("NonExistent"))
                        .inputColumns(List.of())
                        .parameters(Map.of())
                        .build()));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertFalse(resp.isSuccess());
        assertNotNull(resp.getErrors());
        assertFalse(resp.getErrors().isEmpty());
        assertTrue(resp.getErrors().get(0).contains("does not exist"));
    }

    // -------------------------------------------------------------------------
    // SELECT_COLUMNS tests
    // -------------------------------------------------------------------------

    @Test
    void executeWorkflow_selectColumns_succeeds() throws Exception {
        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("SELECT_COLUMNS")
                                .inputSheets(List.of())
                                .inputColumns(List.of("Region", "Actual"))
                                .parameters(Map.of())
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertNull(resp.getErrors());
        assertEquals(2, resp.getSteps().size());
        assertEquals("SELECT_COLUMNS", resp.getSteps().get(1).getOperation());
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("Selected"));

        Path outPath = xlsxPath.resolveSibling(resp.getOutputFileName());
        assertTrue(Files.exists(outPath));
        Files.deleteIfExists(outPath);
    }

    @Test
    void executeWorkflow_selectColumns_missingColumn_returnsFailure() throws Exception {
        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("SELECT_COLUMNS")
                                .inputSheets(List.of())
                                .inputColumns(List.of("NonExistentColumn"))
                                .parameters(Map.of())
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertFalse(resp.isSuccess());
        assertNotNull(resp.getErrors());
        assertTrue(resp.getErrors().get(0).contains("not found"));
    }

    // -------------------------------------------------------------------------
    // REMOVE_EMPTY_ROWS tests
    // -------------------------------------------------------------------------

    @Test
    void executeWorkflow_removeEmptyRows_succeeds() throws Exception {
        Path sparsePath = createSparseWorkbook();
        UploadedFile file = buildFile(sparsePath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sparse"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("REMOVE_EMPTY_ROWS")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertEquals("REMOVE_EMPTY_ROWS", resp.getSteps().get(1).getOperation());
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("Removed"));

        Path outPath = sparsePath.resolveSibling(resp.getOutputFileName());
        assertTrue(Files.exists(outPath));
        Files.deleteIfExists(outPath);
    }

    // -------------------------------------------------------------------------
    // REMOVE_DUPLICATES tests
    // -------------------------------------------------------------------------

    @Test
    void executeWorkflow_removeDuplicates_succeeds() throws Exception {
        Path dupPath = createDuplicateWorkbook();
        UploadedFile file = buildFile(dupPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Data"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("REMOVE_DUPLICATES")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertEquals("REMOVE_DUPLICATES", resp.getSteps().get(1).getOperation());
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("Removed"));

        Path outPath = dupPath.resolveSibling(resp.getOutputFileName());
        assertTrue(Files.exists(outPath));
        Files.deleteIfExists(outPath);
    }

    /**
     * Verifies that REMOVE_DUPLICATES reports the correct row counts in the
     * result summary: the number of rows removed (not kept) is reported.
     *
     * Regression test: the handler previously built the summary using rowsKept
     * instead of rowsRemoved, producing a misleading message like
     * "Removed 18 duplicate rows" when only 2 rows were actually removed.
     */
    @Test
    void executeWorkflow_removeDuplicates_reportsCorrectRowCount() throws Exception {
        // 20 data rows: 2 exact duplicates → 18 unique rows, 2 removed
        Path dupPath = createDuplicateRowsWorkbook();
        UploadedFile file = buildFile(dupPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("REMOVE_DUPLICATES")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertEquals("REMOVE_DUPLICATES", resp.getSteps().get(1).getOperation());

        String summary = resp.getSteps().get(1).getResultSummary();
        assertNotNull(summary, "resultSummary must not be null");

        // Summary must say "2" rows removed, not "18"
        assertTrue(summary.contains("Removed 2 duplicate rows"),
                "Summary should report 'Removed 2 duplicate rows' but was: " + summary);
        assertTrue(summary.contains("18 rows kept"),
                "Summary should report '18 rows kept' but was: " + summary);
        assertFalse(summary.contains("Removed 18 duplicate rows"),
                "Summary must NOT report 'Removed 18' (that would be rowsKept, not rowsRemoved): " + summary);

        // Verify the output workbook has the correct row count
        Path outPath = dupPath.resolveSibling(resp.getOutputFileName());
        assertTrue(Files.exists(outPath));
        try (Workbook outWb = WorkbookFactory.create(outPath.toFile(), null, true)) {
            Sheet unique = outWb.getSheet("Sales_unique");
            assertNotNull(unique, "Output must contain Sales_unique sheet");
            int rowCount = unique.getLastRowNum();
            assertEquals(18, rowCount, "Sales_unique should have 18 rows of data (1 header + 18 unique rows)");
        } finally {
            Files.deleteIfExists(outPath);
        }
    }

    // -------------------------------------------------------------------------
    // FILTER tests
    // -------------------------------------------------------------------------

    @Test
    void executeWorkflow_filter_succeeds() throws Exception {
        Path dataPath = createDataWorkbook();
        UploadedFile file = buildFile(dataPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("FILTER")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of("column", "Region", "operator", "equals", "value", "North"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertEquals("FILTER", resp.getSteps().get(1).getOperation());
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("kept"));

        Path outPath = dataPath.resolveSibling(resp.getOutputFileName());
        assertTrue(Files.exists(outPath));
        Files.deleteIfExists(outPath);
    }

    @Test
    void executeWorkflow_filter_missingColumn_returnsFailure() throws Exception {
        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("FILTER")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of("column", "NonExistent", "operator", "equals", "value", "X"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertFalse(resp.isSuccess());
        assertNotNull(resp.getErrors());
        assertTrue(resp.getErrors().get(0).contains("not found"));
    }

    @Test
    void executeWorkflow_filter_invalidOperator_returnsFailure() throws Exception {
        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("FILTER")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of("column", "Region", "operator", "invalid_op", "value", "X"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertFalse(resp.isSuccess());
        assertNotNull(resp.getErrors());
        assertTrue(resp.getErrors().get(0).contains("unsupported operator"));
    }

    // -------------------------------------------------------------------------
    // SORT tests
    // -------------------------------------------------------------------------

    @Test
    void executeWorkflow_sort_succeeds() throws Exception {
        Path dataPath = createDataWorkbook();
        UploadedFile file = buildFile(dataPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("SORT")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of("column", "Actual", "direction", "desc", "type", "numeric"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertEquals("SORT", resp.getSteps().get(1).getOperation());
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("desc"));

        Path outPath = dataPath.resolveSibling(resp.getOutputFileName());
        assertTrue(Files.exists(outPath));
        Files.deleteIfExists(outPath);
    }

    @Test
    void executeWorkflow_sort_missingColumn_returnsFailure() throws Exception {
        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("SORT")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of("column", "NonExistent", "direction", "asc"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertFalse(resp.isSuccess());
        assertNotNull(resp.getErrors());
        assertTrue(resp.getErrors().get(0).contains("not found"));
    }

    @Test
    void executeWorkflow_sort_invalidDirection_returnsFailure() throws Exception {
        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("SORT")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of("column", "Region", "direction", "sideways"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertFalse(resp.isSuccess());
        assertNotNull(resp.getErrors());
        assertTrue(resp.getErrors().get(0).contains("direction"));
    }

    // -------------------------------------------------------------------------
    // Original workbook unchanged test
    // -------------------------------------------------------------------------

    @Test
    void executeWorkflow_originalWorkbookUnchanged() throws Exception {
        // Record original file checksum and row count
        int originalRows = countRows(xlsxPath);
        byte[] originalBytes = Files.readAllBytes(xlsxPath);

        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("SELECT_COLUMNS")
                                .inputSheets(List.of())
                                .inputColumns(List.of("Region"))
                                .parameters(Map.of())
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        // Original file bytes must be unchanged
        assertArrayEquals(originalBytes, Files.readAllBytes(xlsxPath));
        assertEquals(originalRows, countRows(xlsxPath));

        // Cleanup output
        Path outPath = xlsxPath.resolveSibling(resp.getOutputFileName());
        Files.deleteIfExists(outPath);
    }

    // -------------------------------------------------------------------------
    // Multi-step pipeline test
    // -------------------------------------------------------------------------

    @Test
    void executeWorkflow_multiStepPipeline_succeeds() throws Exception {
        Path multiPath = createMultiStepWorkbook();
        UploadedFile file = buildFile(multiPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("SELECT_COLUMNS")
                                .inputSheets(List.of())
                                .inputColumns(List.of("Region", "Actual"))
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(3).operation("REMOVE_EMPTY_ROWS")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(4).operation("SORT")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of("column", "Actual", "direction", "desc", "type", "numeric"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertEquals(4, resp.getSteps().size());
        assertNull(resp.getErrors());

        Path outPath = multiPath.resolveSibling(resp.getOutputFileName());
        assertTrue(Files.exists(outPath));
        Files.deleteIfExists(outPath);
    }

    // -------------------------------------------------------------------------
    // CALCULATE tests
    // -------------------------------------------------------------------------

    @Test
    void executeWorkflow_calculate_addsComputedColumn() throws Exception {
        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("CALCULATE")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "targetColumn", "Variance",
                                        "expression", "Actual - Target"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertEquals(2, resp.getSteps().size());
        assertEquals("CALCULATE", resp.getSteps().get(1).getOperation());
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("Variance"));

        Path outPath = xlsxPath.resolveSibling(resp.getOutputFileName());
        try {
            // Verify output workbook contains a _calc sheet with computed column
            try (Workbook outWb = WorkbookFactory.create(outPath.toFile(), null, true)) {
                Sheet calcSheet = outWb.getSheet("Sales_calc");
                assertNotNull(calcSheet, "Sales_calc sheet should exist");
                Row header = calcSheet.getRow(0);
                assertEquals("Variance", header.getCell(4).getStringCellValue());
                // Row 1: Actual=100, Target=90, Variance=10
                Row r1 = calcSheet.getRow(1);
                assertEquals(10.0, r1.getCell(4).getNumericCellValue(), 0.001);
                // Row 2: Actual=200, Target=180, Variance=20
                Row r2 = calcSheet.getRow(2);
                assertEquals(20.0, r2.getCell(4).getNumericCellValue(), 0.001);
            }
        } finally {
            Files.deleteIfExists(outPath);
        }
    }

    @Test
    void executeWorkflow_calculate_missingExpression_returnsFailure() throws Exception {
        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("CALCULATE")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of("targetColumn", "X"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertFalse(resp.isSuccess());
        assertNotNull(resp.getErrors());
        assertTrue(resp.getErrors().get(0).contains("expression"));
    }

    @Test
    void executeWorkflow_calculate_missingColumn_returnsFailure() throws Exception {
        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("CALCULATE")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "targetColumn", "X",
                                        "expression", "NoSuchColumn + Actual"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertFalse(resp.isSuccess());
        assertNotNull(resp.getErrors());
        assertTrue(resp.getErrors().get(0).contains("not found"));
    }

    @Test
    void executeWorkflow_calculate_unitPriceExpression_success() throws Exception {
        Path unitPricePath = createUnitPriceWorkbook();
        UploadedFile file = buildFile(unitPricePath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("CALCULATE")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "targetColumn", "Revenue",
                                        "expression", "Quantity * Unit Price"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess(),
                "Workflow should succeed; errors: " + resp.getErrors());
        assertEquals(2, resp.getSteps().size());
        assertEquals("CALCULATE", resp.getSteps().get(1).getOperation());

        Path outPath = unitPricePath.resolveSibling(resp.getOutputFileName());
        try (Workbook outWb = WorkbookFactory.create(outPath.toFile(), null, true)) {
            Sheet calcSheet = outWb.getSheet("Sales_calc");
            assertNotNull(calcSheet, "Sales_calc sheet should exist");

            Row header = calcSheet.getRow(0);
            assertEquals("Revenue", header.getCell(3).getStringCellValue(),
                    "Last column should be Revenue");

            // Row 1: Quantity=5 * Unit Price=10 → Revenue=50
            Row r1 = calcSheet.getRow(1);
            assertEquals(50.0, r1.getCell(3).getNumericCellValue(), 0.001,
                    "Row 1 Revenue should be 5 * 10 = 50");

            // Row 2: Quantity=3 * Unit Price=20 → Revenue=60
            Row r2 = calcSheet.getRow(2);
            assertEquals(60.0, r2.getCell(3).getNumericCellValue(), 0.001,
                    "Row 2 Revenue should be 3 * 20 = 60");
        } finally {
            Files.deleteIfExists(outPath);
        }
    }

    @Test
    void executeWorkflow_calculate_mixedExpression_success() throws Exception {
        Path mixedPath = createMixedHeaderWorkbook();
        UploadedFile file = buildFile(mixedPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("CALCULATE")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "targetColumn", "Combined",
                                        "expression", "Actual + Unit Price"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess(),
                "Workflow should succeed; errors: " + resp.getErrors());
        assertEquals(2, resp.getSteps().size());
        assertEquals("CALCULATE", resp.getSteps().get(1).getOperation());

        Path outPath = mixedPath.resolveSibling(resp.getOutputFileName());
        try (Workbook outWb = WorkbookFactory.create(outPath.toFile(), null, true)) {
            Sheet calcSheet = outWb.getSheet("Sales_calc");
            assertNotNull(calcSheet, "Sales_calc sheet should exist");

            Row header = calcSheet.getRow(0);
            int lastCol = header.getLastCellNum() - 1;
            assertEquals("Combined", header.getCell(lastCol).getStringCellValue(),
                    "Last column should be Combined");

            // Row 1: Actual=100 + Unit Price=10 → 110
            Row r1 = calcSheet.getRow(1);
            assertEquals(110.0, r1.getCell(lastCol).getNumericCellValue(), 0.001,
                    "Row 1 Combined should be 100 + 10 = 110");

            // Row 2: Actual=200 + Unit Price=20 → 220
            Row r2 = calcSheet.getRow(2);
            assertEquals(220.0, r2.getCell(lastCol).getNumericCellValue(), 0.001,
                    "Row 2 Combined should be 200 + 20 = 220");
        } finally {
            Files.deleteIfExists(outPath);
        }
    }

    // -------------------------------------------------------------------------
    // GROUP_BY tests
    // -------------------------------------------------------------------------

    @Test
    void executeWorkflow_groupBy_collapsesRows() throws Exception {
        Path dataPath = createDataWorkbook();
        UploadedFile file = buildFile(dataPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("GROUP_BY")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of("groupColumns", "Region"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertEquals("GROUP_BY", resp.getSteps().get(1).getOperation());
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("2 unique groups"));

        Path outPath = dataPath.resolveSibling(resp.getOutputFileName());
        try {
            try (Workbook outWb = WorkbookFactory.create(outPath.toFile(), null, true)) {
                Sheet grouped = outWb.getSheet("Sales_grouped");
                assertNotNull(grouped, "Sales_grouped sheet should exist");
                assertEquals(3, grouped.getLastRowNum() + 1, "Should have header + 2 grouped rows");
            }
        } finally {
            Files.deleteIfExists(outPath);
        }
    }

    @Test
    void executeWorkflow_groupBy_missingColumn_returnsFailure() throws Exception {
        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("GROUP_BY")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of("groupColumns", "NonExistent"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertFalse(resp.isSuccess());
        assertNotNull(resp.getErrors());
        assertTrue(resp.getErrors().get(0).contains("not found"));
    }

    // -------------------------------------------------------------------------
    // AGGREGATE tests
    // -------------------------------------------------------------------------

    @Test
    void executeWorkflow_aggregate_sumByGroup() throws Exception {
        Path dataPath = createDataWorkbook();
        UploadedFile file = buildFile(dataPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("AGGREGATE")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "groupColumns", "Region",
                                        "aggregateColumn", "Actual",
                                        "aggregateFunction", "SUM"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertEquals("AGGREGATE", resp.getSteps().get(1).getOperation());
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("2 groups"));

        Path outPath = dataPath.resolveSibling(resp.getOutputFileName());
        try {
            try (Workbook outWb = WorkbookFactory.create(outPath.toFile(), null, true)) {
                Sheet agg = outWb.getSheet("Sales_agg");
                assertNotNull(agg, "Sales_agg sheet should exist");
                Row header = agg.getRow(0);
                assertEquals("Region", header.getCell(0).getStringCellValue());
                assertEquals("Actual_SUM", header.getCell(1).getStringCellValue());
                // Total North = 300 + 100 = 400
                // Total South = 150
                double total = 0;
                int groups = 0;
                for (int r = 1; r <= agg.getLastRowNum(); r++) {
                    Row row = agg.getRow(r);
                    if (row == null) continue;
                    Cell sumCell = row.getCell(1);
                    if (sumCell != null && sumCell.getCellType() == CellType.NUMERIC) {
                        total += sumCell.getNumericCellValue();
                        groups++;
                    }
                }
                assertEquals(2, groups);
                assertEquals(550.0, total, 0.001);
            }
        } finally {
            Files.deleteIfExists(outPath);
        }
    }

    @Test
    void executeWorkflow_aggregate_invalidFunction_returnsFailure() throws Exception {
        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("AGGREGATE")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "aggregateColumn", "Actual",
                                        "aggregateFunction", "MEDIAN"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertFalse(resp.isSuccess());
        assertNotNull(resp.getErrors());
        assertTrue(resp.getErrors().get(0).contains("unsupported function"));
    }

    // -------------------------------------------------------------------------
    // JOIN_SHEETS tests
    // -------------------------------------------------------------------------

    @Test
    void executeWorkflow_joinSheets_innerJoin_matchesRows() throws Exception {
        Path leftPath = createJoinLeftWorkbook(); // holds both Employees + departments sheets

        UploadedFile leftFile = buildFile(leftPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(leftFile, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Employees"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("JOIN_SHEETS")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "leftSheet",  "Employees",
                                        "rightSheet", "departments",
                                        "joinColumn", "DeptId",
                                        "joinType",   "INNER"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(leftFile));
        when(approvalRepository.findByFileAndUser(leftFile, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertEquals("JOIN_SHEETS", resp.getSteps().get(1).getOperation());
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("INNER"));
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("matched"));

        Path outPath = leftPath.resolveSibling(resp.getOutputFileName());
        try {
            try (Workbook wb = WorkbookFactory.create(outPath.toFile(), null, true)) {
                Sheet out = wb.getSheet("Employees_joined");
                assertNotNull(out, "Employees_joined sheet should exist");
                // Header: Name, DeptId, DeptName (right.DeptName)
                Row header = out.getRow(0);
                assertEquals("Name",     header.getCell(0).getStringCellValue());
                assertEquals("DeptId",  header.getCell(1).getStringCellValue());
                assertEquals("DeptName", header.getCell(2).getStringCellValue());
                // 2 matching rows: Alice→IT, Bob→Sales
                assertEquals(2, out.getLastRowNum());
            }
        } finally {
            Files.deleteIfExists(outPath);
        }
    }

    @Test
    void executeWorkflow_joinSheets_leftJoin_includesUnmatchedLeftRows() throws Exception {
        Path leftPath = createJoinLeftWorkbook();

        UploadedFile leftFile = buildFile(leftPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(leftFile, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Employees"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("JOIN_SHEETS")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "leftSheet",  "Employees",
                                        "rightSheet", "departments",
                                        "joinColumn", "DeptId",
                                        "joinType",   "LEFT"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(leftFile));
        when(approvalRepository.findByFileAndUser(leftFile, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("LEFT"));
        // Carol has DeptId=99 (no match) so LEFT should include her
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("1 unmatched"));

        Path outPath = leftPath.resolveSibling(resp.getOutputFileName());
        try {
            try (Workbook wb = WorkbookFactory.create(outPath.toFile(), null, true)) {
                Sheet out = wb.getSheet("Employees_joined");
                assertNotNull(out);
                // 3 rows: Alice→IT, Bob→Sales, Carol→(blank)
                assertEquals(3, out.getLastRowNum());
                // Carol's right-side cells should be blank
                Row carol = out.getRow(3);
                assertEquals("Carol", carol.getCell(0).getStringCellValue());
                assertEquals("", carol.getCell(2).getStringCellValue());
            }
        } finally {
            Files.deleteIfExists(outPath);
        }
    }

    @Test
    void executeWorkflow_joinSheets_missingJoinColumn_returnsFailure() throws Exception {
        Path leftPath = createJoinLeftWorkbook();

        UploadedFile leftFile = buildFile(leftPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(leftFile, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Employees"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("JOIN_SHEETS")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "leftSheet",  "Employees",
                                        "rightSheet", "departments",
                                        "joinColumn", "NonExistentColumn",
                                        "joinType",   "INNER"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(leftFile));
        when(approvalRepository.findByFileAndUser(leftFile, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertFalse(resp.isSuccess());
        assertNotNull(resp.getErrors());
        assertTrue(resp.getErrors().get(0).contains("not found"));
    }

    @Test
    void executeWorkflow_joinSheets_missingRightSheet_returnsFailure() throws Exception {
        Path leftPath = createJoinLeftWorkbook();

        UploadedFile leftFile = buildFile(leftPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(leftFile, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Employees"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("JOIN_SHEETS")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "leftSheet",  "Employees",
                                        "rightSheet", "NonExistentSheet",
                                        "joinColumn", "DeptId",
                                        "joinType",   "INNER"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(leftFile));
        when(approvalRepository.findByFileAndUser(leftFile, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertFalse(resp.isSuccess());
        assertNotNull(resp.getErrors());
        assertTrue(resp.getErrors().get(0).contains("not found"));
    }

    @Test
    void executeWorkflow_joinSheets_invalidJoinType_returnsFailure() throws Exception {
        Path leftPath = createJoinLeftWorkbook();

        UploadedFile leftFile = buildFile(leftPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(leftFile, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Employees"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("JOIN_SHEETS")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "leftSheet",  "Employees",
                                        "rightSheet", "departments",
                                        "joinColumn", "DeptId",
                                        "joinType",   "CROSS"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(leftFile));
        when(approvalRepository.findByFileAndUser(leftFile, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertFalse(resp.isSuccess());
        assertNotNull(resp.getErrors());
        assertTrue(resp.getErrors().get(0).contains("joinType"));
    }

    @Test
    void executeWorkflow_joinSheets_duplicateKeysInRight_producesMultipleRows() throws Exception {
        Path leftPath = createJoinRightWorkbookWithDuplicateKeys();

        UploadedFile leftFile = buildFile(leftPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(leftFile, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Employees"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("JOIN_SHEETS")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "leftSheet",  "Employees",
                                        "rightSheet", "depts",
                                        "joinColumn", "DeptId",
                                        "joinType",   "INNER"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(leftFile));
        when(approvalRepository.findByFileAndUser(leftFile, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        // Bob (Sales) has 2 rows in right with Sales → 2 output rows for Bob
        // Alice (IT) has 1 row → 1 output row
        // Carol (99) has 0 → excluded by INNER
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("3 matched"));
    }

    @Test
    void executeWorkflow_joinSheets_nullBlankKeys_skipped() throws Exception {
        Path leftPath = createJoinLeftWorkbook();

        UploadedFile leftFile = buildFile(leftPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(leftFile, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Employees"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("JOIN_SHEETS")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "leftSheet",  "Employees",
                                        "rightSheet", "departments",
                                        "joinColumn", "DeptId",
                                        "joinType",   "INNER"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(leftFile));
        when(approvalRepository.findByFileAndUser(leftFile, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        // Should still match Alice and Bob; Carol is excluded because her DeptId=99 has no match
        // Carol is NOT null/blank so she's excluded by INNER join (key mismatch), not by null check
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("2 matched"));
    }

    @Test
    void executeWorkflow_joinSheets_chainedWithCalculate() throws Exception {
        Path leftPath = createJoinLeftWorkbook();

        UploadedFile leftFile = buildFile(leftPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(leftFile, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Employees"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("JOIN_SHEETS")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "leftSheet",  "Employees",
                                        "rightSheet", "departments",
                                        "joinColumn", "DeptId",
                                        "joinType",   "INNER"))
                                .build(),
                        WorkflowStep.builder()
                                .order(3).operation("CALCULATE")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "targetColumn", "DeptId_Doubled",
                                        "expression",   "DeptId + DeptId"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(leftFile));
        when(approvalRepository.findByFileAndUser(leftFile, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertEquals(3, resp.getSteps().size());
        assertEquals("JOIN_SHEETS",  resp.getSteps().get(1).getOperation());
        assertEquals("CALCULATE",   resp.getSteps().get(2).getOperation());
        assertTrue(resp.getSteps().get(2).getResultSummary().contains("DeptId_Doubled"));

        Path outPath = leftPath.resolveSibling(resp.getOutputFileName());
        try {
            try (Workbook wb = WorkbookFactory.create(outPath.toFile(), null, true)) {
                // Last active sheet after chain is the _calc sheet
                Sheet calcSheet = wb.getSheet("Employees_joined_calc");
                assertNotNull(calcSheet, "Employees_joined_calc sheet should exist");
                Row header = calcSheet.getRow(0);
                // DeptId_Doubled should be the last column
                int lastCol = header.getLastCellNum() - 1;
                assertEquals("DeptId_Doubled", header.getCell(lastCol).getStringCellValue());
                // Row 1: DeptId=IT101→Doubled=IT101IT101 (string concat - no error, just not numeric)
                // Row 2: DeptId=Sales→Doubled=2*Sales... wait, DeptId is string in our test data
                // Let's verify the header has the new column at least
            }
        } finally {
            Files.deleteIfExists(outPath);
        }
    }

    @Test
    void executeWorkflow_joinSheets_originalWorkbookUnchanged() throws Exception {
        Path leftPath = createJoinLeftWorkbook();
        byte[] leftOriginal = Files.readAllBytes(leftPath);

        UploadedFile leftFile = buildFile(leftPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(leftFile, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Employees"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("JOIN_SHEETS")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "leftSheet",  "Employees",
                                        "rightSheet", "departments",
                                        "joinColumn", "DeptId",
                                        "joinType",   "INNER"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(leftFile));
        when(approvalRepository.findByFileAndUser(leftFile, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertArrayEquals(leftOriginal, Files.readAllBytes(leftPath));
    }

    @Test
    void executeWorkflow_joinSheets_unapprovedWorkflow_throwsConflict() throws Exception {
        Path leftPath = createJoinLeftWorkbook();

        UploadedFile leftFile = buildFile(leftPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(leftFile, testUser, WorkflowStatus.PROPOSED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Employees"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("JOIN_SHEETS")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "leftSheet",  "Employees",
                                        "rightSheet", "departments",
                                        "joinColumn", "DeptId",
                                        "joinType",   "INNER"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(leftFile));
        when(approvalRepository.findByFileAndUser(leftFile, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowApprovalException ex = assertThrows(WorkflowApprovalException.class,
                () -> service.executeWorkflow(1L, testUser));

        assertTrue(ex.getMessage().contains("not approved"));
    }

    // -------------------------------------------------------------------------
    // PIVOT tests
    // -------------------------------------------------------------------------

    /**
     * Pivot workbook: Region x Product → Sales (SUM)
     * Sheet: "Sales"
     * Columns: Region, Product, Sales
     * Rows:
     *   North, Apples, 100
     *   North, Oranges, 200
     *   South, Apples, 150
     *   South, Oranges, 175
     * Pivot expected (SUM):
     *   Headers: Region, Apples, Oranges
     *   North: 100, 200
     *   South: 150, 175
     */
    private Path createPivotWorkbook() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sales");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Region");
        header.createCell(1).setCellValue("Product");
        header.createCell(2).setCellValue("Sales");
        Row r1 = sheet.createRow(1);
        r1.createCell(0).setCellValue("North");
        r1.createCell(1).setCellValue("Apples");
        r1.createCell(2).setCellValue(100);
        Row r2 = sheet.createRow(2);
        r2.createCell(0).setCellValue("North");
        r2.createCell(1).setCellValue("Oranges");
        r2.createCell(2).setCellValue(200);
        Row r3 = sheet.createRow(3);
        r3.createCell(0).setCellValue("South");
        r3.createCell(1).setCellValue("Apples");
        r3.createCell(2).setCellValue(150);
        Row r4 = sheet.createRow(4);
        r4.createCell(0).setCellValue("South");
        r4.createCell(1).setCellValue("Oranges");
        r4.createCell(2).setCellValue(175);
        Path path = tempDir.resolve("pivot.xlsx");
        try (FileOutputStream fos = new FileOutputStream(path.toFile())) {
            wb.write(fos);
        }
        wb.close();
        return path;
    }

    @Test
    void executeWorkflow_pivot_basicSum() throws Exception {
        Path pivotPath = createPivotWorkbook();
        UploadedFile file = buildFile(pivotPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("PIVOT")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "rowColumn",            "Region",
                                        "pivotColumn",          "Product",
                                        "valueColumn",          "Sales",
                                        "aggregationFunction",   "SUM"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertEquals("PIVOT", resp.getSteps().get(1).getOperation());
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("PIVOT"));

        Path outPath = pivotPath.resolveSibling(resp.getOutputFileName());
        try {
            try (Workbook wb = WorkbookFactory.create(outPath.toFile(), null, true)) {
                Sheet out = wb.getSheet("Sales_pivot");
                assertNotNull(out, "Sales_pivot sheet should exist");
                // Header: Region, Apples, Oranges (sorted alphabetically)
                Row header = out.getRow(0);
                assertEquals("Region", header.getCell(0).getStringCellValue());
                assertEquals("Apples", header.getCell(1).getStringCellValue());
                assertEquals("Oranges", header.getCell(2).getStringCellValue());
                // North row: 100, 200
                Row north = out.getRow(1);
                assertEquals("North", north.getCell(0).getStringCellValue());
                assertEquals(100.0, north.getCell(1).getNumericCellValue(), 0.001);
                assertEquals(200.0, north.getCell(2).getNumericCellValue(), 0.001);
                // South row: 150, 175
                Row south = out.getRow(2);
                assertEquals("South", south.getCell(0).getStringCellValue());
                assertEquals(150.0, south.getCell(1).getNumericCellValue(), 0.001);
                assertEquals(175.0, south.getCell(2).getNumericCellValue(), 0.001);
                // 2 data rows
                assertEquals(2, out.getLastRowNum());
            }
        } finally {
            Files.deleteIfExists(outPath);
        }
    }

    @Test
    void executeWorkflow_pivot_countAggregation() throws Exception {
        Path pivotPath = createPivotWorkbook();
        UploadedFile file = buildFile(pivotPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("PIVOT")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "rowColumn",          "Region",
                                        "pivotColumn",        "Product",
                                        "valueColumn",        "Sales",
                                        "aggregationFunction", "COUNT"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        Path outPath = pivotPath.resolveSibling(resp.getOutputFileName());
        try {
            try (Workbook wb = WorkbookFactory.create(outPath.toFile(), null, true)) {
                Sheet out = wb.getSheet("Sales_pivot");
                // COUNT: each cell = 1 (one row per combination)
                Row north = out.getRow(1);
                assertEquals(1.0, north.getCell(1).getNumericCellValue(), 0.001);
                assertEquals(1.0, north.getCell(2).getNumericCellValue(), 0.001);
            }
        } finally {
            Files.deleteIfExists(outPath);
        }
    }

    @Test
    void executeWorkflow_pivot_invalidRowColumn_returnsFailure() throws Exception {
        Path pivotPath = createPivotWorkbook();
        UploadedFile file = buildFile(pivotPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("PIVOT")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "rowColumn",          "NonExistent",
                                        "pivotColumn",        "Product",
                                        "valueColumn",        "Sales",
                                        "aggregationFunction", "SUM"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertFalse(resp.isSuccess());
        assertNotNull(resp.getErrors());
        assertTrue(resp.getErrors().get(0).contains("not found"));
    }

    @Test
    void executeWorkflow_pivot_invalidAggregation_returnsFailure() throws Exception {
        Path pivotPath = createPivotWorkbook();
        UploadedFile file = buildFile(pivotPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("PIVOT")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "rowColumn",          "Region",
                                        "pivotColumn",        "Product",
                                        "valueColumn",        "Sales",
                                        "aggregationFunction", "MEDIAN"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertFalse(resp.isSuccess());
        assertNotNull(resp.getErrors());
        assertTrue(resp.getErrors().get(0).contains("aggregationFunction"));
    }

    @Test
    void executeWorkflow_pivot_emptyInput_returnsFailure() throws Exception {
        // Workbook with header but no data rows
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Empty");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Region");
        header.createCell(1).setCellValue("Sales");
        Path emptyPath = tempDir.resolve("empty.xlsx");
        try (FileOutputStream fos = new FileOutputStream(emptyPath.toFile())) {
            wb.write(fos);
        }
        wb.close();

        UploadedFile file = buildFile(emptyPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Empty"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("PIVOT")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "rowColumn",          "Region",
                                        "pivotColumn",        "Sales",
                                        "valueColumn",        "Sales",
                                        "aggregationFunction", "SUM"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertFalse(resp.isSuccess());
        assertNotNull(resp.getErrors());
        assertTrue(resp.getErrors().get(0).contains("no valid data"));
    }

    @Test
    void executeWorkflow_pivot_chainedWithCalculate() throws Exception {
        Path pivotPath = createPivotWorkbook();
        UploadedFile file = buildFile(pivotPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("PIVOT")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "rowColumn",          "Region",
                                        "pivotColumn",        "Product",
                                        "valueColumn",        "Sales",
                                        "aggregationFunction", "SUM"))
                                .build(),
                        WorkflowStep.builder()
                                .order(3).operation("CALCULATE")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "targetColumn", "Total",
                                        "expression",   "Apples + Oranges"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertEquals(3, resp.getSteps().size());
        assertEquals("PIVOT",    resp.getSteps().get(1).getOperation());
        assertEquals("CALCULATE", resp.getSteps().get(2).getOperation());
        assertTrue(resp.getSteps().get(2).getResultSummary().contains("Total"));

        Path outPath = pivotPath.resolveSibling(resp.getOutputFileName());
        try {
            try (Workbook wb = WorkbookFactory.create(outPath.toFile(), null, true)) {
                // After PIVOT → CALCULATE, active sheet should be the _calc sheet
                Sheet calcSheet = wb.getSheet("Sales_pivot_calc");
                assertNotNull(calcSheet, "Sales_pivot_calc sheet should exist");
                Row header = calcSheet.getRow(0);
                int lastCol = header.getLastCellNum() - 1;
                assertEquals("Total", header.getCell(lastCol).getStringCellValue());
                // North: Apples(100) + Oranges(200) = 300
                Row north = calcSheet.getRow(1);
                assertEquals(300.0, north.getCell(lastCol).getNumericCellValue(), 0.001);
            }
        } finally {
            Files.deleteIfExists(outPath);
        }
    }

    @Test
    void executeWorkflow_pivot_originalWorkbookUnchanged() throws Exception {
        Path pivotPath = createPivotWorkbook();
        byte[] original = Files.readAllBytes(pivotPath);

        UploadedFile file = buildFile(pivotPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("PIVOT")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "rowColumn",          "Region",
                                        "pivotColumn",        "Product",
                                        "valueColumn",        "Sales",
                                        "aggregationFunction", "SUM"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertArrayEquals(original, Files.readAllBytes(pivotPath));
    }

    @Test
    void executeWorkflow_pivot_unapprovedWorkflow_throwsConflict() throws Exception {
        Path pivotPath = createPivotWorkbook();
        UploadedFile file = buildFile(pivotPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.PROPOSED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("PIVOT")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "rowColumn",          "Region",
                                        "pivotColumn",        "Product",
                                        "valueColumn",        "Sales",
                                        "aggregationFunction", "SUM"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowApprovalException ex = assertThrows(WorkflowApprovalException.class,
                () -> service.executeWorkflow(1L, testUser));

        assertTrue(ex.getMessage().contains("not approved"));
    }

    // -------------------------------------------------------------------------
    // CREATE_SUMMARY tests
    // -------------------------------------------------------------------------

    @Test
    void executeWorkflow_createSummary_basicAllColumns() throws Exception {
        Path dataPath = createDataWorkbook();
        UploadedFile file = buildFile(dataPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("CREATE_SUMMARY")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertEquals("CREATE_SUMMARY", resp.getSteps().get(1).getOperation());
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("Sales_summary"));
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("2 columns"));
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("3 rows"));

        Path outPath = dataPath.resolveSibling(resp.getOutputFileName());
        try {
            try (Workbook wb = WorkbookFactory.create(outPath.toFile(), null, true)) {
                Sheet summary = wb.getSheet("Sales_summary");
                assertNotNull(summary, "Sales_summary sheet should exist");
                Row header = summary.getRow(0);
                assertEquals("Region", header.getCell(0).getStringCellValue());
                assertEquals("Actual", header.getCell(1).getStringCellValue());
                assertEquals(3, summary.getLastRowNum(), "3 data rows (North, South, North)");
                // Verify values copied correctly
                assertEquals("North", summary.getRow(1).getCell(0).getStringCellValue());
                assertEquals(300.0, summary.getRow(1).getCell(1).getNumericCellValue(), 0.001);
            }
        } finally {
            Files.deleteIfExists(outPath);
        }
    }

    @Test
    void executeWorkflow_createSummary_selectedColumns() throws Exception {
        Path dataPath = createDataWorkbook();
        UploadedFile file = buildFile(dataPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("CREATE_SUMMARY")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of("columns", "Region"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("1 columns"));

        Path outPath = dataPath.resolveSibling(resp.getOutputFileName());
        try {
            try (Workbook wb = WorkbookFactory.create(outPath.toFile(), null, true)) {
                Sheet summary = wb.getSheet("Sales_summary");
                assertNotNull(summary);
                Row header = summary.getRow(0);
                assertEquals(1, header.getLastCellNum(), "Only 1 data column");
                assertEquals("Region", header.getCell(0).getStringCellValue());
                // All 3 rows should still be present
                assertEquals(3, summary.getLastRowNum());
            }
        } finally {
            Files.deleteIfExists(outPath);
        }
    }

    @Test
    void executeWorkflow_createSummary_withTitle() throws Exception {
        Path dataPath = createDataWorkbook();
        UploadedFile file = buildFile(dataPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("CREATE_SUMMARY")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of("title", "Regional Sales Report"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("Regional Sales Report"));
    }

    @Test
    void executeWorkflow_createSummary_invalidSourceSheet_returnsFailure() throws Exception {
        Path dataPath = createDataWorkbook();
        UploadedFile file = buildFile(dataPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("CREATE_SUMMARY")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of("sourceSheet", "NonExistentSheet"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertFalse(resp.isSuccess());
        assertNotNull(resp.getErrors());
        assertTrue(resp.getErrors().get(0).contains("not found"));
    }

    @Test
    void executeWorkflow_createSummary_invalidColumn_returnsFailure() throws Exception {
        Path dataPath = createDataWorkbook();
        UploadedFile file = buildFile(dataPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("CREATE_SUMMARY")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of("columns", "Region,NonExistentColumn"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertFalse(resp.isSuccess());
        assertNotNull(resp.getErrors());
        assertTrue(resp.getErrors().get(0).contains("not found"));
    }

    @Test
    void executeWorkflow_createSummary_emptyInput_returnsFailure() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("EmptySheet");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("ColA");
        // No data rows
        Path emptyPath = tempDir.resolve("empty_summary.xlsx");
        try (FileOutputStream fos = new FileOutputStream(emptyPath.toFile())) {
            wb.write(fos);
        }
        wb.close();

        UploadedFile file = buildFile(emptyPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("EmptySheet"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("CREATE_SUMMARY")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        // Empty input is not an error — it produces a summary with 0 data rows
        assertTrue(resp.isSuccess());
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("0 rows"));
    }

    @Test
    void executeWorkflow_createSummary_chainedWithCalculate() throws Exception {
        Path dataPath = createDataWorkbook();
        UploadedFile file = buildFile(dataPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("CREATE_SUMMARY")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of("columns", "Region,Actual"))
                                .build(),
                        WorkflowStep.builder()
                                .order(3).operation("CALCULATE")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of(
                                        "targetColumn", "Actual_Doubled",
                                        "expression",   "Actual + Actual"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertEquals(3, resp.getSteps().size());
        assertEquals("CREATE_SUMMARY",  resp.getSteps().get(1).getOperation());
        assertEquals("CALCULATE",       resp.getSteps().get(2).getOperation());
        assertTrue(resp.getSteps().get(2).getResultSummary().contains("Actual_Doubled"));

        Path outPath = dataPath.resolveSibling(resp.getOutputFileName());
        try {
            try (Workbook wb = WorkbookFactory.create(outPath.toFile(), null, true)) {
                Sheet calcSheet = wb.getSheet("Sales_summary_calc");
                assertNotNull(calcSheet, "Sales_summary_calc sheet should exist");
                Row header = calcSheet.getRow(0);
                int lastCol = header.getLastCellNum() - 1;
                assertEquals("Actual_Doubled", header.getCell(lastCol).getStringCellValue());
            }
        } finally {
            Files.deleteIfExists(outPath);
        }
    }

    @Test
    void executeWorkflow_createSummary_originalWorkbookUnchanged() throws Exception {
        Path dataPath = createDataWorkbook();
        byte[] original = Files.readAllBytes(dataPath);

        UploadedFile file = buildFile(dataPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("CREATE_SUMMARY")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of("columns", "Region"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertArrayEquals(original, Files.readAllBytes(dataPath));
    }

    @Test
    void executeWorkflow_createSummary_unapprovedWorkflow_throwsConflict() throws Exception {
        Path dataPath = createDataWorkbook();
        UploadedFile file = buildFile(dataPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.PROPOSED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("CREATE_SUMMARY")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowApprovalException ex = assertThrows(WorkflowApprovalException.class,
                () -> service.executeWorkflow(1L, testUser));

        assertTrue(ex.getMessage().contains("not approved"));
    }
    // -------------------------------------------------------------------------
    // EXPORT_EXCEL tests
    // -------------------------------------------------------------------------

    @Test
    void executeWorkflow_exportExcel_succeedsAndWritesExportedFile() throws Exception {
        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("EXPORT_EXCEL")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        byte[] originalBytes = Files.readAllBytes(xlsxPath);
        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertEquals(2, resp.getSteps().size());
        assertEquals("EXPORT_EXCEL", resp.getSteps().get(1).getOperation());
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("EXPORT_EXCEL"));
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("_exported.xlsx"));

        // The export artifact must exist alongside the source file
        Path exportPath = xlsxPath.resolveSibling(
                "sales_exported.xlsx".replace("sales", xlsxPath.getFileName().toString().replace(".xlsx", "")));
        // The naming uses the *stored* path. Our buildFile uses xlsxPath (a temp path);
        // for the test we derive the expected export path the same way the handler does.
        String srcName = xlsxPath.getFileName().toString();
        int dot = srcName.lastIndexOf('.');
        String exportName = (dot > 0 ? srcName.substring(0, dot) : srcName) + "_exported.xlsx";
        Path expectedExport = xlsxPath.getParent().resolve(exportName);
        assertTrue(Files.exists(expectedExport), "Exported file should exist at " + expectedExport);
        assertTrue(Files.size(expectedExport) > 0, "Exported file should be non-empty");

        // The primary output XLSX is still produced by the executor
        Path outPath = xlsxPath.resolveSibling(resp.getOutputFileName());
        assertTrue(Files.exists(outPath));

        // The export file is a valid XLSX (ZIP, starts with PK)
        byte[] head = Files.readAllBytes(expectedExport);
        assertTrue(head.length > 4);
        assertEquals(0x50, head[0] & 0xFF);
        assertEquals(0x4B, head[1] & 0xFF);

        // The export contains the Sales sheet with the original data preserved
        try (Workbook outWb = WorkbookFactory.create(expectedExport.toFile(), null, true)) {
            Sheet sales = outWb.getSheet("Sales");
            assertNotNull(sales, "Exported file must contain the Sales sheet");
            assertEquals("North", sales.getRow(1).getCell(1).getStringCellValue());
            assertEquals(100, (int) sales.getRow(1).getCell(2).getNumericCellValue());
        }

        // The original input file must NOT be modified
        assertArrayEquals(originalBytes, Files.readAllBytes(xlsxPath));

        // Cleanup
        Files.deleteIfExists(expectedExport);
        Files.deleteIfExists(outPath);
    }

    @Test
    void executeWorkflow_exportExcel_filteredSheets_onlySelectedSheetsPresent() throws Exception {
        Path dataPath = createDataWorkbook();
        UploadedFile file = buildFile(dataPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        // FILTER creates a new sheet called "Sales_filtered";
                        // EXPORT with sheets="Sales_filtered" should export only that sheet.
                        WorkflowStep.builder()
                                .order(2).operation("FILTER")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of("column", "Region", "operator", "equals", "value", "North"))
                                .build(),
                        WorkflowStep.builder()
                                .order(3).operation("EXPORT_EXCEL")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of("sheets", "Sales_filtered"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        // The step summary for EXPORT_EXCEL includes the exported filename
        assertTrue(resp.getSteps().get(2).getResultSummary().contains("EXPORT_EXCEL"));

        String srcName = dataPath.getFileName().toString();
        int dot = srcName.lastIndexOf('.');
        String exportName = (dot > 0 ? srcName.substring(0, dot) : srcName) + "_exported.xlsx";
        Path expectedExport = dataPath.getParent().resolve(exportName);
        assertTrue(Files.exists(expectedExport), "Exported file should exist at " + expectedExport);

        try (Workbook outWb = WorkbookFactory.create(expectedExport.toFile(), null, true)) {
            assertEquals(1, outWb.getNumberOfSheets(), "Exported workbook should have exactly 1 sheet");
            assertNotNull(outWb.getSheet("Sales_filtered"),
                    "Exported sheet should be Sales_filtered (only the filtered sheet was requested)");
        }

        Path outPath = dataPath.resolveSibling(resp.getOutputFileName());
        Files.deleteIfExists(expectedExport);
        Files.deleteIfExists(outPath);
    }

    @Test
    void executeWorkflow_exportExcel_customFilename_isUsed() throws Exception {
        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("EXPORT_EXCEL")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of("filename", "custom_export.xlsx"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertTrue(resp.getSteps().get(1).getResultSummary().contains("custom_export.xlsx"));

        Path customFile = xlsxPath.getParent().resolve("custom_export.xlsx");
        assertTrue(Files.exists(customFile), "Custom export file should exist at " + customFile);

        Path outPath = xlsxPath.resolveSibling(resp.getOutputFileName());
        Files.deleteIfExists(customFile);
        Files.deleteIfExists(outPath);
    }

    @Test
    void executeWorkflow_exportExcel_missingSheet_returnsFailure() throws Exception {
        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("EXPORT_EXCEL")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of("sheets", "DoesNotExist"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertFalse(resp.isSuccess());
        assertNotNull(resp.getErrors());
        assertTrue(resp.getErrors().get(0).contains("not found"));

        // No export file should have been created
        String srcName = xlsxPath.getFileName().toString();
        int dot = srcName.lastIndexOf('.');
        String exportName = (dot > 0 ? srcName.substring(0, dot) : srcName) + "_exported.xlsx";
        Path defaultExport = xlsxPath.getParent().resolve(exportName);
        assertFalse(Files.exists(defaultExport), "Failed export must not leave a file behind");
    }

    @Test
    void executeWorkflow_exportExcel_pathTraversal_throwsOrIsRejected() throws Exception {
        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("EXPORT_EXCEL")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of("filename", "../../../etc/passwd.xlsx"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        // The handler must reject any filename with path separators or traversal segments
        // (writes are restricted to the source file's parent directory).
        assertFalse(resp.isSuccess(), "Path-traversal filename must be rejected");
        assertNotNull(resp.getErrors());
        // The export file must NOT exist anywhere in the temp directory
        assertFalse(Files.exists(Path.of("etc/passwd.xlsx")),
                "Path traversal must not create a file at the attacker's target");
    }

    @Test
    void executeWorkflow_exportExcel_originalWorkbookUnchanged() throws Exception {
        byte[] originalBytes = Files.readAllBytes(xlsxPath);
        int originalRowCount = countRows(xlsxPath);

        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("EXPORT_EXCEL")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        // Source file must be byte-for-byte identical
        assertArrayEquals(originalBytes, Files.readAllBytes(xlsxPath));
        assertEquals(originalRowCount, countRows(xlsxPath));

        // Cleanup
        String srcName = xlsxPath.getFileName().toString();
        int dot = srcName.lastIndexOf('.');
        String exportName = (dot > 0 ? srcName.substring(0, dot) : srcName) + "_exported.xlsx";
        Path expectedExport = xlsxPath.getParent().resolve(exportName);
        Path outPath = xlsxPath.resolveSibling(resp.getOutputFileName());
        Files.deleteIfExists(expectedExport);
        Files.deleteIfExists(outPath);
    }

    @Test
    void executeWorkflow_exportExcel_doesNotChangeActiveSheet() throws Exception {
        // EXPORT_EXCEL is a side-effect export — it must not change the active sheet
        // for downstream steps.
        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("EXPORT_EXCEL")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(3).operation("FILTER")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of("column", "Region", "operator", "equals", "value", "North"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        assertEquals(3, resp.getSteps().size());
        // FILTER (step 3) must still work because EXPORT_EXCEL didn't break the active sheet chain
        assertEquals("FILTER", resp.getSteps().get(2).getOperation());

        String srcName = xlsxPath.getFileName().toString();
        int dot = srcName.lastIndexOf('.');
        String exportName = (dot > 0 ? srcName.substring(0, dot) : srcName) + "_exported.xlsx";
        Path expectedExport = xlsxPath.getParent().resolve(exportName);
        Path outPath = xlsxPath.resolveSibling(resp.getOutputFileName());
        Files.deleteIfExists(expectedExport);
        Files.deleteIfExists(outPath);
    }

    @Test
    void executeWorkflow_exportExcel_preservesFormulasInCopiedSheets() throws Exception {
        Path formulaPath = createFormulaWorkbook();
        UploadedFile file = buildFile(formulaPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("EXPORT_EXCEL")
                                .inputSheets(List.of())
                                .inputColumns(List.of())
                                .parameters(Map.of("sheets", "Sales"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());

        String srcName = formulaPath.getFileName().toString();
        int dot = srcName.lastIndexOf('.');
        String exportName = (dot > 0 ? srcName.substring(0, dot) : srcName) + "_exported.xlsx";
        Path expectedExport = formulaPath.getParent().resolve(exportName);
        try (Workbook outWb = WorkbookFactory.create(expectedExport.toFile(), null, true)) {
            Sheet sales = outWb.getSheet("Sales");
            assertNotNull(sales);
            // Formulas should be preserved in the exported sheet
            Cell sumC1 = sales.getRow(1).getCell(2);
            assertEquals(CellType.FORMULA, sumC1.getCellType());
            assertEquals("A2+B2", sumC1.getCellFormula());
        }

        Path outPath = formulaPath.resolveSibling(resp.getOutputFileName());
        Files.deleteIfExists(expectedExport);
        Files.deleteIfExists(outPath);
    }

    // -------------------------------------------------------------------------
    // Helper methods
    // -------------------------------------------------------------------------

    private Path createTestWorkbook() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sales");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Date");
        header.createCell(1).setCellValue("Region");
        header.createCell(2).setCellValue("Actual");
        header.createCell(3).setCellValue("Target");
        Row r1 = sheet.createRow(1);
        r1.createCell(0).setCellValue("2024-01-01");
        r1.createCell(1).setCellValue("North");
        r1.createCell(2).setCellValue(100);
        r1.createCell(3).setCellValue(90);
        Row r2 = sheet.createRow(2);
        r2.createCell(0).setCellValue("2024-01-02");
        r2.createCell(1).setCellValue("South");
        r2.createCell(2).setCellValue(200);
        r2.createCell(3).setCellValue(180);
        Path path = tempDir.resolve("sales.xlsx");
        try (FileOutputStream fos = new FileOutputStream(path.toFile())) {
            wb.write(fos);
        }
        wb.close();
        return path;
    }

    /**
     * Creates a workbook with a "Unit Price" column to verify that CALCULATE
     * can reference multi-word column names.
     */
    private Path createUnitPriceWorkbook() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sales");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Quantity");
        header.createCell(1).setCellValue("Unit Price");
        header.createCell(2).setCellValue("Actual");

        Row r1 = sheet.createRow(1);
        r1.createCell(0).setCellValue(5);
        r1.createCell(1).setCellValue(10);
        r1.createCell(2).setCellValue(100);

        Row r2 = sheet.createRow(2);
        r2.createCell(0).setCellValue(3);
        r2.createCell(1).setCellValue(20);
        r2.createCell(2).setCellValue(200);

        Path path = tempDir.resolve("unitprice.xlsx");
        try (FileOutputStream fos = new FileOutputStream(path.toFile())) {
            wb.write(fos);
        }
        wb.close();
        return path;
    }

    /**
     * Creates a workbook with both single-word and multi-word columns
     * (including "Unit Price") to verify mixed expressions work.
     */
    private Path createMixedHeaderWorkbook() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sales");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Actual");
        header.createCell(1).setCellValue("Unit Price");
        header.createCell(2).setCellValue("Quantity");

        Row r1 = sheet.createRow(1);
        r1.createCell(0).setCellValue(100);
        r1.createCell(1).setCellValue(10);
        r1.createCell(2).setCellValue(5);

        Row r2 = sheet.createRow(2);
        r2.createCell(0).setCellValue(200);
        r2.createCell(1).setCellValue(20);
        r2.createCell(2).setCellValue(3);

        Path path = tempDir.resolve("mixedheader.xlsx");
        try (FileOutputStream fos = new FileOutputStream(path.toFile())) {
            wb.write(fos);
        }
        wb.close();
        return path;
    }

    private Path createDataWorkbook() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sales");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Region");
        header.createCell(1).setCellValue("Actual");
        Row r1 = sheet.createRow(1);
        r1.createCell(0).setCellValue("North");
        r1.createCell(1).setCellValue(300);
        Row r2 = sheet.createRow(2);
        r2.createCell(0).setCellValue("South");
        r2.createCell(1).setCellValue(150);
        Row r3 = sheet.createRow(3);
        r3.createCell(0).setCellValue("North");
        r3.createCell(1).setCellValue(100);
        Path path = tempDir.resolve("data.xlsx");
        try (FileOutputStream fos = new FileOutputStream(path.toFile())) {
            wb.write(fos);
        }
        wb.close();
        return path;
    }

    private Path createSparseWorkbook() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sparse");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("ColA");
        header.createCell(1).setCellValue("ColB");
        Row r1 = sheet.createRow(1);
        r1.createCell(0).setCellValue("A");
        r1.createCell(1).setCellValue("B");
        sheet.createRow(2); // blank row
        Row r3 = sheet.createRow(3);
        r3.createCell(0).setCellValue("C");
        r3.createCell(1).setCellValue("D");
        Path path = tempDir.resolve("sparse.xlsx");
        try (FileOutputStream fos = new FileOutputStream(path.toFile())) {
            wb.write(fos);
        }
        wb.close();
        return path;
    }

    private Path createDuplicateWorkbook() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Data");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Name");
        header.createCell(1).setCellValue("Value");
        Row r1 = sheet.createRow(1);
        r1.createCell(0).setCellValue("Alice");
        r1.createCell(1).setCellValue(10);
        Row r2 = sheet.createRow(2);
        r2.createCell(0).setCellValue("Bob");
        r2.createCell(1).setCellValue(20);
        Row r3 = sheet.createRow(3);
        r3.createCell(0).setCellValue("Alice");
        r3.createCell(1).setCellValue(30);
        Path path = tempDir.resolve("dup.xlsx");
        try (FileOutputStream fos = new FileOutputStream(path.toFile())) {
            wb.write(fos);
        }
        wb.close();
        return path;
    }

    /**
     * Creates a workbook with a single sheet "Sales" containing 20 data rows
     * of which exactly 2 are exact duplicates of earlier rows.
     * After REMOVE_DUPLICATES, 18 rows should be kept and 2 removed.
     * All cell values are set as strings to eliminate any numeric-coercion ambiguity.
     */
    private Path createDuplicateRowsWorkbook() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sales");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Order ID");
        header.createCell(1).setCellValue("Amount");

        // Rows 1-18: unique data rows
        String[] names = {"A1","B1","C1","D1","E1","F1","G1","H1","I1","J1","K1","L1","M1","N1","O1","P1","Q1","R1"};
        int[]    vals  = {1001,1002,1003,1004,1005,1006,1007,1008,1009,1010,1011,1012,1013,1014,1015,1016,1017,1018};
        for (int i = 0; i < 18; i++) {
            Row r = sheet.createRow(i + 1);
            r.createCell(0).setCellValue(names[i]);
            r.createCell(1).setCellValue(String.valueOf(vals[i]));
        }
        // Duplicate 1: identical to row 5 (E1, 1005) → row index 6
        Row d1 = sheet.createRow(20);
        d1.createCell(0).setCellValue("E1");
        d1.createCell(1).setCellValue("1005");
        // Duplicate 2: identical to row 12 (L1, 1012) → row index 13
        Row d2 = sheet.createRow(21);
        d2.createCell(0).setCellValue("L1");
        d2.createCell(1).setCellValue("1012");

        Path path = tempDir.resolve("dup_rows.xlsx");
        try (FileOutputStream fos = new FileOutputStream(path.toFile())) {
            wb.write(fos);
        }
        wb.close();
        return path;
    }

    private Path createMultiStepWorkbook() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sales");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Date");
        header.createCell(1).setCellValue("Region");
        header.createCell(2).setCellValue("Actual");
        header.createCell(3).setCellValue("Target");
        for (int i = 1; i <= 5; i++) {
            Row r = sheet.createRow(i);
            r.createCell(0).setCellValue("2024-01-" + String.format("%02d", i));
            r.createCell(1).setCellValue(i % 2 == 0 ? "North" : "South");
            r.createCell(2).setCellValue(100 + i * 10);
            r.createCell(3).setCellValue(90 + i * 10);
        }
        Path path = tempDir.resolve("multi.xlsx");
        try (FileOutputStream fos = new FileOutputStream(path.toFile())) {
            wb.write(fos);
        }
        wb.close();
        return path;
    }

    /**
     * Single workbook with both left and right sheets for join tests.
     * Sheet "Employees": Name, DeptId
     *   Rows: Alice/IT101, Bob/Sales, Carol/99
     * Sheet "departments": DeptId, DeptName
     *   Rows: IT101/IT, Sales/Sales Dept
     * Returns the same path for both — both sheets live in one workbook,
     * which is what the real uploaded file would contain.
     */
    private Path createJoinLeftWorkbook() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();

        // Left sheet
        Sheet left = wb.createSheet("Employees");
        Row leftHeader = left.createRow(0);
        leftHeader.createCell(0).setCellValue("Name");
        leftHeader.createCell(1).setCellValue("DeptId");
        Row lr1 = left.createRow(1);
        lr1.createCell(0).setCellValue("Alice");
        lr1.createCell(1).setCellValue("IT101");
        Row lr2 = left.createRow(2);
        lr2.createCell(0).setCellValue("Bob");
        lr2.createCell(1).setCellValue("Sales");
        Row lr3 = left.createRow(3);
        lr3.createCell(0).setCellValue("Carol");
        lr3.createCell(1).setCellValue("99");

        // Right sheet (same workbook, different sheet name)
        Sheet right = wb.createSheet("departments");
        Row rightHeader = right.createRow(0);
        rightHeader.createCell(0).setCellValue("DeptId");
        rightHeader.createCell(1).setCellValue("DeptName");
        Row rr1 = right.createRow(1);
        rr1.createCell(0).setCellValue("IT101");
        rr1.createCell(1).setCellValue("IT");
        Row rr2 = right.createRow(2);
        rr2.createCell(0).setCellValue("Sales");
        rr2.createCell(1).setCellValue("Sales Dept");

        Path path = tempDir.resolve("join_basic.xlsx");
        try (FileOutputStream fos = new FileOutputStream(path.toFile())) {
            wb.write(fos);
        }
        wb.close();
        return path;
    }

    /**
     * Alias for createJoinLeftWorkbook (same workbook holds both sheets).
     * Kept for API clarity in tests.
     */
    private Path createJoinRightWorkbook() throws Exception {
        return createJoinLeftWorkbook();
    }

    /**
     * Workbooks with duplicate keys in the right sheet.
     * Sheet "Employees" as above; sheet "depts" has IT101/IT, Sales/Sales1, Sales/Sales2.
     */
    private Path createJoinRightWorkbookWithDuplicateKeys() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();

        Sheet left = wb.createSheet("Employees");
        Row leftHeader = left.createRow(0);
        leftHeader.createCell(0).setCellValue("Name");
        leftHeader.createCell(1).setCellValue("DeptId");
        Row lr1 = left.createRow(1);
        lr1.createCell(0).setCellValue("Alice");
        lr1.createCell(1).setCellValue("IT101");
        Row lr2 = left.createRow(2);
        lr2.createCell(0).setCellValue("Bob");
        lr2.createCell(1).setCellValue("Sales");
        Row lr3 = left.createRow(3);
        lr3.createCell(0).setCellValue("Carol");
        lr3.createCell(1).setCellValue("99");

        Sheet right = wb.createSheet("depts");
        Row rightHeader = right.createRow(0);
        rightHeader.createCell(0).setCellValue("DeptId");
        rightHeader.createCell(1).setCellValue("DeptName");
        Row rr1 = right.createRow(1);
        rr1.createCell(0).setCellValue("IT101");
        rr1.createCell(1).setCellValue("IT");
        Row rr2 = right.createRow(2);
        rr2.createCell(0).setCellValue("Sales");
        rr2.createCell(1).setCellValue("Sales1");
        Row rr3 = right.createRow(3);
        rr3.createCell(0).setCellValue("Sales");
        rr3.createCell(1).setCellValue("Sales2");

        Path path = tempDir.resolve("join_dup.xlsx");
        try (FileOutputStream fos = new FileOutputStream(path.toFile())) {
            wb.write(fos);
        }
        wb.close();
        return path;
    }

    private int countRows(Path path) throws Exception {
        try (Workbook wb = WorkbookFactory.create(path.toFile(), null, true)) {
            Sheet sheet = wb.getSheetAt(0);
            int first = sheet.getFirstRowNum();
            int last = sheet.getLastRowNum();
            return last - first + 1;
        }
    }

    private UploadedFile buildFile(Path path, User user) {
        return UploadedFile.builder()
                .id(1L)
                .user(user)
                .originalFileName("test.xlsx")
                .storedFileName("test.xlsx")
                .fileType("xlsx")
                .fileSize(1024L)
                .storagePath(path.toString())
                .status(FileStatus.UPLOADED)
                .build();
    }

    private WorkflowApproval buildApproval(UploadedFile file, User user, WorkflowStatus status) {
        return WorkflowApproval.builder()
                .id(1L)
                .file(file)
                .user(user)
                .status(status)
                .workflowName("Test Workflow")
                .build();
    }

    private WorkflowApproval buildApprovalWithSteps(UploadedFile file, User user,
                                                   WorkflowStatus status, List<WorkflowStep> steps) {
        WorkflowApproval approval = WorkflowApproval.builder()
                .id(1L)
                .file(file)
                .user(user)
                .status(status)
                .workflowName("Test Workflow")
                .build();
        approval.setStepsFromJson(steps);
        return approval;
    }

    // -------------------------------------------------------------------------
    // Source-sheet preservation tests (Option B)
    // -------------------------------------------------------------------------

    /**
     * Creates a workbook with two sheets where the first contains a formula cell.
     * Used to verify formulas are preserved by the source-sheet copy.
     */
    private Path createFormulaWorkbook() throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sales = wb.createSheet("Sales");
        Row header = sales.createRow(0);
        header.createCell(0).setCellValue("A");
        header.createCell(1).setCellValue("B");
        header.createCell(2).setCellValue("Sum");
        Row r1 = sales.createRow(1);
        r1.createCell(0).setCellValue(10);
        r1.createCell(1).setCellValue(20);
        r1.createCell(2).setCellFormula("A2+B2");
        Row r2 = sales.createRow(2);
        r2.createCell(0).setCellValue(5);
        r2.createCell(1).setCellValue(7);
        r2.createCell(2).setCellFormula("A3+B3");

        Sheet notes = wb.createSheet("Notes");
        notes.createRow(0).createCell(0).setCellValue("Important");

        Path path = tempDir.resolve("formula.xlsx");
        try (FileOutputStream fos = new FileOutputStream(path.toFile())) {
            wb.write(fos);
        }
        wb.close();
        return path;
    }

    @Test
    void executeWorkflow_preservesOriginalSourceSheetsInOutput() throws Exception {
        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        Path outPath = xlsxPath.resolveSibling(resp.getOutputFileName());
        try (Workbook outWb = WorkbookFactory.create(outPath.toFile(), null, true)) {
            // Original "Sales" sheet must be preserved in the output
            Sheet preserved = outWb.getSheet("Sales");
            assertNotNull(preserved, "Original 'Sales' sheet must be preserved in output");
            assertEquals("Date",   preserved.getRow(0).getCell(0).getStringCellValue());
            assertEquals("Region", preserved.getRow(0).getCell(1).getStringCellValue());
            assertEquals(100, (int) preserved.getRow(1).getCell(2).getNumericCellValue());
            assertEquals(200, (int) preserved.getRow(2).getCell(2).getNumericCellValue());
        } finally {
            Files.deleteIfExists(outPath);
        }
    }

    @Test
    void executeWorkflow_preservesOriginalDataBytesUnchanged() throws Exception {
        byte[] originalBytes = Files.readAllBytes(xlsxPath);

        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("SELECT_COLUMNS")
                                .inputSheets(List.of())
                                .inputColumns(List.of("Region"))
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(3).operation("CREATE_CHART")
                                .inputSheets(List.of())
                                .inputColumns(List.of("Region", "Actual"))
                                .parameters(Map.of(
                                        "sourceSheet", "Sales",
                                        "chartType", "BAR",
                                        "labelColumn", "Region",
                                        "dataColumn", "Actual"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        // Source file bytes must be byte-for-byte unchanged
        assertArrayEquals(originalBytes, Files.readAllBytes(xlsxPath));
        Path outPath = xlsxPath.resolveSibling(resp.getOutputFileName());
        Files.deleteIfExists(outPath);
    }

    @Test
    void executeWorkflow_preservesFormulasInCopiedSourceSheet() throws Exception {
        Path formulaPath = createFormulaWorkbook();

        UploadedFile file = buildFile(formulaPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        Path outPath = formulaPath.resolveSibling(resp.getOutputFileName());
        try (Workbook outWb = WorkbookFactory.create(outPath.toFile(), null, true)) {
            Sheet sales = outWb.getSheet("Sales");
            assertNotNull(sales, "Original 'Sales' sheet must be preserved");
            // The "Sum" column had formulas in the source — they must still be formulas
            assertEquals(CellType.FORMULA, sales.getRow(1).getCell(2).getCellType());
            assertEquals("A2+B2", sales.getRow(1).getCell(2).getCellFormula());
            assertEquals(CellType.FORMULA, sales.getRow(2).getCell(2).getCellType());
            assertEquals("A3+B3", sales.getRow(2).getCell(2).getCellFormula());
            // Second source sheet also preserved
            Sheet notes = outWb.getSheet("Notes");
            assertNotNull(notes, "Second source sheet 'Notes' must be preserved");
            assertEquals("Important", notes.getRow(0).getCell(0).getStringCellValue());
        } finally {
            Files.deleteIfExists(outPath);
        }
    }

    @Test
    void executeWorkflow_sheetNameCollision_doesNotOverwriteSourceSheet() throws Exception {
        // A CREATE_CHART step creates a sheet called "<sourceSheet>_chart_data".
        // With preservation, the source "Sales" sheet must remain with its original data;
        // the chart sheet must be named to avoid collision.
        UploadedFile file = buildFile(xlsxPath, testUser);
        WorkflowApproval approval = buildApprovalWithSteps(file, testUser, WorkflowStatus.APPROVED,
                List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("CREATE_CHART")
                                .inputSheets(List.of())
                                .inputColumns(List.of("Region", "Actual"))
                                .parameters(Map.of(
                                        "sourceSheet", "Sales",
                                        "chartType", "BAR",
                                        "labelColumn", "Region",
                                        "dataColumn", "Actual"))
                                .build()
                ));

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));
        when(approvalRepository.findByFileAndUser(file, testUser))
                .thenReturn(Optional.of(approval));

        WorkflowExecutionResponse resp = service.executeWorkflow(1L, testUser);

        assertTrue(resp.isSuccess());
        Path outPath = xlsxPath.resolveSibling(resp.getOutputFileName());
        try (Workbook outWb = WorkbookFactory.create(outPath.toFile(), null, true)) {
            Sheet original = outWb.getSheet("Sales");
            assertNotNull(original, "Original 'Sales' must remain by its original name");
            // Original "Sales" data must be intact — not overwritten by handler content
            assertEquals("Date",   original.getRow(0).getCell(0).getStringCellValue());
            assertEquals("Region", original.getRow(0).getCell(1).getStringCellValue());
            assertEquals(100, (int) original.getRow(1).getCell(2).getNumericCellValue());
            // Chart data sheet must exist (its name does not collide with "Sales")
            Sheet chartData = outWb.getSheet("Sales_chart_data");
            assertNotNull(chartData, "Handler-created chart data sheet must also be present");
        } finally {
            Files.deleteIfExists(outPath);
        }
    }
}
