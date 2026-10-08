package com.aiexcel.platform.service;

import com.aiexcel.platform.dto.*;
import com.aiexcel.platform.entity.FileStatus;
import com.aiexcel.platform.entity.UploadedFile;
import com.aiexcel.platform.entity.User;
import com.aiexcel.platform.exception.FileNotFoundException;
import com.aiexcel.platform.repository.UploadedFileRepository;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.junit.jupiter.api.io.TempDir;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.io.ByteArrayOutputStream;
import java.nio.file.Path;
import java.util.*;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.when;

/**
 * Unit tests for {@link WorkflowValidatorService}.
 * Covers: valid workflow, invalid operation, invalid step order,
 * unknown sheet, unknown column, missing required fields, unsafe operation rejection.
 */
@ExtendWith(MockitoExtension.class)
class WorkflowValidatorServiceTest {

    @Mock
    private UploadedFileRepository fileRepository;

    @TempDir
    Path tempDir;

    private WorkflowValidatorService service;
    private User testUser;
    private UploadedFile testFile;

    @BeforeEach
    void setUp() throws Exception {
        service = new WorkflowValidatorService(fileRepository);

        testUser = User.builder()
                .id(1L)
                .name("Test User")
                .email("test@example.com")
                .password("password")
                .build();

        // Create a real workbook on disk: "Sales" sheet with columns Date, Region, Actual, Target
        Path xlsxPath = createTestWorkbook();

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

    @Test
    void validateWorkflow_valid_returnsValid() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));

        WorkflowValidationRequest request = WorkflowValidationRequest.builder()
                .instruction("Compare actual vs target")
                .workflowName("Regional Variance")
                .description("Variance analysis")
                .steps(List.of(
                        WorkflowStep.builder()
                                .order(1)
                                .operation("READ_SHEET")
                                .description("Load Sales")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2)
                                .operation("FILTER")
                                .description("Filter rows")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of("Region"))
                                .parameters(Map.of(
                                        "column", "Region",
                                        "operator", "equals",
                                        "value", "North"))
                                .build()
                ))
                .build();

        WorkflowValidationResponse response = service.validateWorkflow(1L, testUser, request);

        assertNotNull(response);
        assertTrue(response.isValid());
        assertTrue(response.getErrors() == null || response.getErrors().isEmpty());
    }

    @Test
    void validateWorkflow_invalidOperation_returnsErrors() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));

        WorkflowValidationRequest request = WorkflowValidationRequest.builder()
                .instruction("...")
                .workflowName("BadOp")
                .description("Uses unknown operation")
                .steps(List.of(
                        WorkflowStep.builder()
                                .order(1)
                                .operation("RUN_PYTHON")  // not in catalog
                                .description("Run Python")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build()
                ))
                .build();

        WorkflowValidationResponse response = service.validateWorkflow(1L, testUser, request);

        assertNotNull(response);
        assertFalse(response.isValid());
        assertTrue(response.getErrors().stream().anyMatch(e -> e.contains("RUN_PYTHON") && e.contains("unsafe")),
                "Expected an error referencing the unsafe operation, got: " + response.getErrors());
    }

    @Test
    void validateWorkflow_invalidStepOrder_returnsErrors() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));

        WorkflowValidationRequest request = WorkflowValidationRequest.builder()
                .instruction("...")
                .workflowName("BadOrder")
                .description("Step orders skip and are out of sequence")
                .steps(List.of(
                        WorkflowStep.builder()
                                .order(5)  // wrong — must be 1
                                .operation("READ_SHEET")
                                .description("Load")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2)  // wrong — must be 2
                                .operation("FILTER")
                                .description("Filter")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build()
                ))
                .build();

        WorkflowValidationResponse response = service.validateWorkflow(1L, testUser, request);

        assertNotNull(response);
        assertFalse(response.isValid());
        assertTrue(response.getErrors().stream().anyMatch(e -> e.contains("order must be")),
                "Expected an order-related error, got: " + response.getErrors());
    }

    @Test
    void validateWorkflow_unknownSheet_returnsErrors() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));

        WorkflowValidationRequest request = WorkflowValidationRequest.builder()
                .instruction("...")
                .workflowName("BadSheet")
                .description("References a sheet that does not exist")
                .steps(List.of(
                        WorkflowStep.builder()
                                .order(1)
                                .operation("READ_SHEET")
                                .description("Load")
                                .inputSheets(List.of("NotARealSheet"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build()
                ))
                .build();

        WorkflowValidationResponse response = service.validateWorkflow(1L, testUser, request);

        assertNotNull(response);
        assertFalse(response.isValid());
        assertTrue(response.getErrors().stream().anyMatch(e -> e.contains("NotARealSheet") && e.contains("does not exist")),
                "Expected an unknown-sheet error, got: " + response.getErrors());
    }

    @Test
    void validateWorkflow_unknownColumn_returnsErrors() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));

        WorkflowValidationRequest request = WorkflowValidationRequest.builder()
                .instruction("...")
                .workflowName("BadColumn")
                .description("References a column that does not exist")
                .steps(List.of(
                        WorkflowStep.builder()
                                .order(1)
                                .operation("FILTER")
                                .description("Filter")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of("NoSuchColumn"))
                                .parameters(Map.of())
                                .build()
                ))
                .build();

        WorkflowValidationResponse response = service.validateWorkflow(1L, testUser, request);

        assertNotNull(response);
        assertFalse(response.isValid());
        assertTrue(response.getErrors().stream().anyMatch(e -> e.contains("NoSuchColumn") && e.contains("not found")),
                "Expected an unknown-column error, got: " + response.getErrors());
    }

    @Test
    void validateWorkflow_missingRequiredFields_returnsErrors() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));

        // Step with null operation, null inputSheets, null parameters — all required fields missing
        WorkflowStep badStep = WorkflowStep.builder()
                .order(1)
                .operation(null)
                .description("Broken step")
                .inputSheets(null)
                .inputColumns(null)
                .parameters(null)
                .build();

        WorkflowValidationRequest request = WorkflowValidationRequest.builder()
                .instruction("...")
                .workflowName("MissingFields")
                .description("Step missing all required fields")
                .steps(List.of(badStep))
                .build();

        WorkflowValidationResponse response = service.validateWorkflow(1L, testUser, request);

        assertNotNull(response);
        assertFalse(response.isValid());
        List<String> errs = response.getErrors();
        assertTrue(errs.stream().anyMatch(e -> e.contains("operation must not be blank")),
                "Expected operation-required error, got: " + errs);
        assertTrue(errs.stream().anyMatch(e -> e.contains("inputSheets must not be null")),
                "Expected inputSheets-required error, got: " + errs);
        assertTrue(errs.stream().anyMatch(e -> e.contains("inputColumns must not be null")),
                "Expected inputColumns-required error, got: " + errs);
        assertTrue(errs.stream().anyMatch(e -> e.contains("parameters map must not be null")),
                "Expected parameters-required error, got: " + errs);
    }

    @Test
    void validateWorkflow_unsafeOperation_rejected() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));

        // All operations outside the allowed catalog must be rejected
        List<String> unsafe = List.of(
                "DELETE_FILE", "EXECUTE_SQL", "RUN_PYTHON", "SHELL_EXEC", "DROP_TABLE",
                "read_sheet"  // case-sensitive — should be rejected
        );
        List<WorkflowStep> steps = new ArrayList<>();
        int order = 1;
        for (String op : unsafe) {
            steps.add(WorkflowStep.builder()
                    .order(order++)
                    .operation(op)
                    .description("unsafe")
                    .inputSheets(List.of("Sales"))
                    .inputColumns(List.of())
                    .parameters(Map.of())
                    .build());
        }

        WorkflowValidationRequest request = WorkflowValidationRequest.builder()
                .instruction("...")
                .workflowName("Unsafe")
                .description("All unsafe ops")
                .steps(steps)
                .build();

        WorkflowValidationResponse response = service.validateWorkflow(1L, testUser, request);

        assertNotNull(response);
        assertFalse(response.isValid());
        // Each unsafe op should produce at least one error
        for (String op : unsafe) {
            assertTrue(response.getErrors().stream().anyMatch(e -> e.contains(op)),
                    "Expected an error referencing '" + op + "', got: " + response.getErrors());
        }
    }

    @Test
    void validateWorkflow_nullStepsList_returnsInvalid() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));

        WorkflowValidationRequest request = WorkflowValidationRequest.builder()
                .instruction("...")
                .workflowName("NullSteps")
                .description("Null steps")
                .steps(null)
                .build();

        WorkflowValidationResponse response = service.validateWorkflow(1L, testUser, request);

        assertNotNull(response);
        assertFalse(response.isValid());
        assertFalse(response.getErrors().isEmpty());
    }

    @Test
    void validateWorkflow_fileNotOwned_throwsFileNotFound() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(999L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.empty());

        WorkflowValidationRequest request = WorkflowValidationRequest.builder()
                .instruction("...")
                .workflowName("X")
                .description("X")
                .steps(List.of())
                .build();

        assertThrows(FileNotFoundException.class, () -> {
            service.validateWorkflow(999L, testUser, request);
        });
    }

    @Test
    void validateWorkflow_emptySteps_returnsValid() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));

        WorkflowValidationRequest request = WorkflowValidationRequest.builder()
                .instruction("...")
                .workflowName("Empty")
                .description("No steps")
                .steps(List.of())
                .build();

        WorkflowValidationResponse response = service.validateWorkflow(1L, testUser, request);

        assertNotNull(response);
        assertTrue(response.isValid());
    }

    // -------------------------------------------------------------------------
    // Required-parameter validation (V1)
    // -------------------------------------------------------------------------

    @Test
    void validateWorkflow_groupByMissingGroupColumns_returnsInvalid() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));

        // Mirrors the AI-produced payload that previously slipped through validation
        // and failed only at execution time.
        WorkflowValidationRequest request = WorkflowValidationRequest.builder()
                .instruction("Group by product")
                .workflowName("Group by Product")
                .description("Group")
                .steps(List.of(
                        WorkflowStep.builder()
                                .order(5)
                                .operation("GROUP_BY")
                                .description("Group")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of("Product", "Sales Amount"))
                                .parameters(Map.of())
                                .build()
                ))
                .build();

        WorkflowValidationResponse response = service.validateWorkflow(1L, testUser, request);

        assertNotNull(response);
        assertFalse(response.isValid());
        assertTrue(response.getErrors().stream()
                        .anyMatch(e -> e.contains("GROUP_BY") && e.contains("groupColumns")),
                "Expected required-parameter error for GROUP_BY groupColumns, got: "
                        + response.getErrors());
    }

    @Test
    void validateWorkflow_filterMissingRequiredParameter_returnsInvalid() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));

        // 'operator' is required but absent.
        WorkflowValidationRequest request = WorkflowValidationRequest.builder()
                .instruction("Filter")
                .workflowName("Filter")
                .description("Filter")
                .steps(List.of(
                        WorkflowStep.builder()
                                .order(2)
                                .operation("FILTER")
                                .description("Filter")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of("Region"))
                                .parameters(Map.of("column", "Region", "value", "North"))
                                .build()
                ))
                .build();

        WorkflowValidationResponse response = service.validateWorkflow(1L, testUser, request);

        assertFalse(response.isValid());
        assertTrue(response.getErrors().stream()
                        .anyMatch(e -> e.contains("FILTER") && e.contains("'operator'")),
                "Expected missing-'operator' error, got: " + response.getErrors());
    }

    @Test
    void validateWorkflow_sortMissingRequiredParameter_returnsInvalid() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));

        // 'direction' is required but absent.
        WorkflowValidationRequest request = WorkflowValidationRequest.builder()
                .instruction("Sort")
                .workflowName("Sort")
                .description("Sort")
                .steps(List.of(
                        WorkflowStep.builder()
                                .order(3)
                                .operation("SORT")
                                .description("Sort")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of("Region"))
                                .parameters(Map.of("column", "Region"))
                                .build()
                ))
                .build();

        WorkflowValidationResponse response = service.validateWorkflow(1L, testUser, request);

        assertFalse(response.isValid());
        assertTrue(response.getErrors().stream()
                        .anyMatch(e -> e.contains("SORT") && e.contains("'direction'")),
                "Expected missing-'direction' error, got: " + response.getErrors());
    }

    @Test
    void validateWorkflow_aggregateMissingRequiredFunction_returnsInvalid() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));

        // 'aggregateFunction' is required but absent.
        WorkflowValidationRequest request = WorkflowValidationRequest.builder()
                .instruction("Aggregate")
                .workflowName("Aggregate")
                .description("Aggregate")
                .steps(List.of(
                        WorkflowStep.builder()
                                .order(4)
                                .operation("AGGREGATE")
                                .description("Aggregate")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of("Actual"))
                                .parameters(Map.of("aggregateColumn", "Actual"))
                                .build()
                ))
                .build();

        WorkflowValidationResponse response = service.validateWorkflow(1L, testUser, request);

        assertFalse(response.isValid());
        assertTrue(response.getErrors().stream()
                        .anyMatch(e -> e.contains("AGGREGATE") && e.contains("'aggregateFunction'")),
                "Expected missing-'aggregateFunction' error, got: " + response.getErrors());
    }

    @Test
    void validateWorkflow_aggregateMissingRequiredColumn_returnsInvalid() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));

        // 'aggregateColumn' is required but absent.
        WorkflowValidationRequest request = WorkflowValidationRequest.builder()
                .instruction("Aggregate")
                .workflowName("Aggregate")
                .description("Aggregate")
                .steps(List.of(
                        WorkflowStep.builder()
                                .order(4)
                                .operation("AGGREGATE")
                                .description("Aggregate")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of("Actual"))
                                .parameters(Map.of("aggregateFunction", "SUM"))
                                .build()
                ))
                .build();

        WorkflowValidationResponse response = service.validateWorkflow(1L, testUser, request);

        assertFalse(response.isValid());
        assertTrue(response.getErrors().stream()
                        .anyMatch(e -> e.contains("AGGREGATE") && e.contains("'aggregateColumn'")),
                "Expected missing-'aggregateColumn' error, got: " + response.getErrors());
    }

    @Test
    void validateWorkflow_aggregateWithNoGroupColumns_returnsValid() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));

        WorkflowValidationRequest request = WorkflowValidationRequest.builder()
                .instruction("Aggregate total sum")
                .workflowName("Aggregate Total")
                .description("Aggregate without grouping")
                .steps(List.of(
                        WorkflowStep.builder()
                                .order(1)
                                .operation("AGGREGATE")
                                .description("Aggregate")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of("Actual"))
                                .parameters(Map.of(
                                        "aggregateColumn", "Actual",
                                        "aggregateFunction", "SUM"))
                                .build()
                ))
                .build();

        WorkflowValidationResponse response = service.validateWorkflow(1L, testUser, request);

        assertNotNull(response);
        assertTrue(response.isValid(),
                "Expected valid when groupColumns is omitted, got errors: " + response.getErrors());
    }

    @Test
    void validateWorkflow_aggregateWithBlankGroupColumns_returnsInvalid() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));

        WorkflowValidationRequest request = WorkflowValidationRequest.builder()
                .instruction("Aggregate blank group")
                .workflowName("Aggregate Blank")
                .description("Aggregate with blank grouping")
                .steps(List.of(
                        WorkflowStep.builder()
                                .order(1)
                                .operation("AGGREGATE")
                                .description("Aggregate")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of("Actual"))
                                .parameters(Map.of(
                                        "groupColumns", "  ",
                                        "aggregateColumn", "Actual",
                                        "aggregateFunction", "SUM"))
                                .build()
                ))
                .build();

        WorkflowValidationResponse response = service.validateWorkflow(1L, testUser, request);

        assertFalse(response.isValid());
        assertTrue(response.getErrors().stream()
                        .anyMatch(e -> e.contains("AGGREGATE") && e.contains("'groupColumns' must not be blank")),
                "Expected non-blank groupColumns error, got: " + response.getErrors());
    }

    @Test
    void validateWorkflow_createChartMissingRequiredParameter_returnsInvalid() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));

        // 'dataColumn' is required but absent.
        WorkflowValidationRequest request = WorkflowValidationRequest.builder()
                .instruction("Chart")
                .workflowName("Chart")
                .description("Chart")
                .steps(List.of(
                        WorkflowStep.builder()
                                .order(1)
                                .operation("CREATE_CHART")
                                .description("Chart")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of("Product", "Amount"))
                                .parameters(Map.of(
                                        "sourceSheet", "Sales",
                                        "chartType", "BAR",
                                        "labelColumn", "Product"))
                                .build()
                ))
                .build();

        WorkflowValidationResponse response = service.validateWorkflow(1L, testUser, request);

        assertFalse(response.isValid());
        assertTrue(response.getErrors().stream()
                        .anyMatch(e -> e.contains("CREATE_CHART") && e.contains("'dataColumn'")),
                "Expected missing-'dataColumn' error, got: " + response.getErrors());
    }

    @Test
    void validateWorkflow_groupByWithGroupColumns_returnsValid() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));

        WorkflowValidationRequest request = WorkflowValidationRequest.builder()
                .instruction("Group by Region")
                .workflowName("Group by Region")
                .description("Group")
                .steps(List.of(
                        WorkflowStep.builder()
                                .order(1)
                                .operation("GROUP_BY")
                                .description("Group by Region")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of("Region"))
                                .parameters(Map.of("groupColumns", "Region"))
                                .build()
                ))
                .build();

        WorkflowValidationResponse response = service.validateWorkflow(1L, testUser, request);

        assertNotNull(response);
        assertTrue(response.isValid(),
                "Expected valid, got errors: " + response.getErrors());
    }

    @Test
    void validateWorkflow_existingValidWorkflowsRemainValid() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));

        // Reuses the exact pre-existing test pattern from validateWorkflow_valid_returnsValid,
        // extended to include a step from each parameter-requiring operation with all
        // required keys present.
        WorkflowValidationRequest request = WorkflowValidationRequest.builder()
                .instruction("Full pipeline")
                .workflowName("Full pipeline")
                .description("All operations")
                .steps(List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("SELECT_COLUMNS")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of("Region", "Actual"))
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(3).operation("REMOVE_EMPTY_ROWS")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(4).operation("REMOVE_DUPLICATES")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(5).operation("FILTER")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of("Region"))
                                .parameters(Map.of(
                                        "column", "Region",
                                        "operator", "equals",
                                        "value", "North"))
                                .build(),
                        WorkflowStep.builder()
                                .order(6).operation("SORT")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of("Actual"))
                                .parameters(Map.of("column", "Actual", "direction", "desc"))
                                .build(),
                        WorkflowStep.builder()
                                .order(7).operation("CALCULATE")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of("Actual", "Target"))
                                .parameters(Map.of("targetColumn", "Variance", "expression", "Actual - Target"))
                                .build(),
                        WorkflowStep.builder()
                                .order(8).operation("GROUP_BY")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of("Region"))
                                .parameters(Map.of("groupColumns", "Region"))
                                .build(),
                        WorkflowStep.builder()
                                .order(9).operation("AGGREGATE")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of("Region", "Actual"))
                                .parameters(Map.of(
                                        "groupColumns", "Region",
                                        "aggregateColumn", "Actual",
                                        "aggregateFunction", "SUM"))
                                .build(),
                        WorkflowStep.builder()
                                .order(10).operation("CREATE_CHART")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of("Region", "Actual"))
                                .parameters(Map.of(
                                        "sourceSheet", "Sales",
                                        "chartType", "BAR",
                                        "labelColumn", "Region",
                                        "dataColumn", "Actual"))
                                .build()
                ))
                .build();

        WorkflowValidationResponse response = service.validateWorkflow(1L, testUser, request);

        assertNotNull(response);
        assertTrue(response.isValid(),
                "Expected existing valid workflow to remain valid, got errors: "
                        + response.getErrors());
    }

    @Test
    void validateWorkflow_noParameterOperations_validWithEmptyParameters() {
        when(fileRepository.findByIdAndUserAndStatusNot(eq(1L), eq(testUser), eq(FileStatus.DELETED)))
                .thenReturn(Optional.of(testFile));

        // Each operation that requires no parameters must still validate clean
        // when given an empty parameters map.
        WorkflowValidationRequest request = WorkflowValidationRequest.builder()
                .instruction("No-param ops")
                .workflowName("No-param ops")
                .description("No-param ops")
                .steps(List.of(
                        WorkflowStep.builder()
                                .order(1).operation("READ_SHEET")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(2).operation("SELECT_COLUMNS")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of("Region"))
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(3).operation("REMOVE_EMPTY_ROWS")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(4).operation("REMOVE_DUPLICATES")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(5).operation("CREATE_SUMMARY")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build(),
                        WorkflowStep.builder()
                                .order(6).operation("VALIDATE")
                                .inputSheets(List.of("Sales"))
                                .inputColumns(List.of())
                                .parameters(Map.of())
                                .build()
                ))
                .build();

        WorkflowValidationResponse response = service.validateWorkflow(1L, testUser, request);

        assertNotNull(response);
        assertTrue(response.isValid(),
                "Expected no-parameter operations to remain valid with empty parameters, got errors: "
                        + response.getErrors());
    }

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    /**
     * Creates a small xlsx workbook with one sheet "Sales" and 4 columns.
     */
    private Path createTestWorkbook() throws Exception {
        org.apache.poi.xssf.usermodel.XSSFWorkbook wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook();
        org.apache.poi.ss.usermodel.Sheet sheet = wb.createSheet("Sales");

        org.apache.poi.ss.usermodel.Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Date");
        header.createCell(1).setCellValue("Region");
        header.createCell(2).setCellValue("Actual");
        header.createCell(3).setCellValue("Target");

        org.apache.poi.ss.usermodel.Row r1 = sheet.createRow(1);
        r1.createCell(0).setCellValue("2024-01-01");
        r1.createCell(1).setCellValue("North");
        r1.createCell(2).setCellValue(100.0);
        r1.createCell(3).setCellValue(120.0);

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        wb.write(baos);
        wb.close();

        Path filePath = tempDir.resolve("test-sales.xlsx");
        java.nio.file.Files.write(filePath, baos.toByteArray());
        return filePath;
    }
}
