package com.aiexcel.platform.service;

import com.aiexcel.platform.dto.*;
import com.aiexcel.platform.entity.User;
import com.aiexcel.platform.exception.FileNotFoundException;
import com.aiexcel.platform.exception.WorkflowProposalException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.springframework.ai.chat.messages.AssistantMessage;
import org.springframework.ai.chat.model.ChatResponse;
import org.springframework.ai.chat.model.Generation;
import org.springframework.ai.chat.prompt.Prompt;

import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.Mockito.*;
import org.mockito.ArgumentCaptor;

@ExtendWith(MockitoExtension.class)
class WorkflowProposalServiceTest {

    @Mock
    private ExcelAnalyzerService excelAnalyzerService;

    @Mock
    private WorkbookUnderstandingService workbookUnderstandingService;

    @Mock
    private org.springframework.ai.chat.model.ChatModel chatModel;

    private User testUser;

    private WorkbookAnalysisResponse createMockAnalysis() {
        return WorkbookAnalysisResponse.builder()
                .fileId(1L)
                .fileName("sales.xlsx")
                .sheetCount(1)
                .sheets(List.of(
                        SheetAnalysis.builder()
                                .sheetName("Sales")
                                .rowCount(100)
                                .columnCount(4)
                                .columns(List.of(
                                        ColumnAnalysis.builder()
                                                .columnName("Date")
                                                .columnIndex(0)
                                                .detectedType(ColumnDataType.DATE)
                                                .nonEmptyCount(100)
                                                .emptyCount(0)
                                                .missingPercentage(0.0)
                                                .uniqueCount(30)
                                                .duplicateCount(70)
                                                .sampleValues(List.of("2024-01-01"))
                                                .build(),
                                        ColumnAnalysis.builder()
                                                .columnName("Region")
                                                .columnIndex(1)
                                                .detectedType(ColumnDataType.TEXT)
                                                .nonEmptyCount(100)
                                                .emptyCount(0)
                                                .missingPercentage(0.0)
                                                .uniqueCount(5)
                                                .duplicateCount(95)
                                                .sampleValues(List.of("North"))
                                                .build(),
                                        ColumnAnalysis.builder()
                                                .columnName("Actual")
                                                .columnIndex(2)
                                                .detectedType(ColumnDataType.DECIMAL)
                                                .nonEmptyCount(100)
                                                .emptyCount(0)
                                                .missingPercentage(0.0)
                                                .uniqueCount(80)
                                                .duplicateCount(20)
                                                .sampleValues(List.of("100.50"))
                                                .build(),
                                        ColumnAnalysis.builder()
                                                .columnName("Target")
                                                .columnIndex(3)
                                                .detectedType(ColumnDataType.DECIMAL)
                                                .nonEmptyCount(100)
                                                .emptyCount(0)
                                                .missingPercentage(0.0)
                                                .uniqueCount(80)
                                                .duplicateCount(20)
                                                .sampleValues(List.of("120.00"))
                                                .build()))
                                .build()))
                .build();
    }

    private WorkbookUnderstandingResponse createMockUnderstanding() {
        return WorkbookUnderstandingResponse.builder()
                .summary("Sales data with regional breakdown and targets")
                .businessDomain("Sales")
                .importantColumns(List.of())
                .possibleRelationships(List.of())
                .dataQualityObservations(List.of())
                .suggestedAnalyses(List.of())
                .build();
    }

    @BeforeEach
    void setUp() {
        testUser = User.builder()
                .id(1L)
                .name("Test User")
                .email("test@example.com")
                .password("password")
                .build();
    }

    @Test
    void proposeWorkflow_noApiKey_throwsException() {
        WorkflowProposalService service = new WorkflowProposalService(excelAnalyzerService,
                workbookUnderstandingService, null);

        WorkflowProposalException ex = assertThrows(WorkflowProposalException.class, () -> {
            service.proposeWorkflow(1L, testUser, "analyze sales");
        });
        assertEquals(
                "AI workflow proposal is not configured. Set spring.ai.openai.api-key to enable this feature.",
                ex.getMessage());

        verifyNoInteractions(excelAnalyzerService);
        verifyNoInteractions(workbookUnderstandingService);
    }

    @Test
    void proposeWorkflow_validAiResponse_returnsProposal() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser)).thenReturn(createMockAnalysis());
        when(workbookUnderstandingService.understandWorkbook(1L, testUser)).thenReturn(createMockUnderstanding());

        String aiJson = """
                {
                    "fileId": 1,
                    "instruction": "Compare actual vs target by region",
                    "workflowName": "Regional Actual vs Target",
                    "description": "Compute per-region actual-vs-target variance and surface underperforming regions.",
                    "steps": [
                        {
                            "order": 1,
                            "operation": "READ_SHEET",
                            "description": "Load the Sales sheet",
                            "inputSheets": ["Sales"],
                            "inputColumns": [],
                            "parameters": null
                        },
                        {
                            "order": 2,
                            "operation": "GROUP_BY",
                            "description": "Group rows by Region",
                            "inputSheets": ["Sales"],
                            "inputColumns": ["Region"],
                            "parameters": null
                        },
                        {
                            "order": 3,
                            "operation": "AGGREGATE",
                            "description": "Sum Actual and Target per region",
                            "inputSheets": ["Sales"],
                            "inputColumns": ["Actual", "Target"],
                            "parameters": {"functions": "SUM", "groupBy": "Region"}
                        },
                        {
                            "order": 4,
                            "operation": "CALCULATE",
                            "description": "Compute variance (Actual - Target)",
                            "inputSheets": ["Sales"],
                            "inputColumns": ["Actual", "Target"],
                            "parameters": {"expression": "Actual - Target"}
                        },
                        {
                            "order": 5,
                            "operation": "FILTER",
                            "description": "Keep regions below target",
                            "inputSheets": ["Sales"],
                            "inputColumns": ["Actual", "Target"],
                            "parameters": {"where": "variance < 0"}
                        }
                    ]
                }
                """;
        ChatResponse aiResponse = new ChatResponse(List.of(new Generation(new AssistantMessage(aiJson))));
        when(chatModel.call(any(Prompt.class))).thenReturn(aiResponse);

        WorkflowProposalService service = new WorkflowProposalService(excelAnalyzerService,
                workbookUnderstandingService, chatModel);
        WorkflowProposalResponse response = service.proposeWorkflow(1L, testUser, "Compare actual vs target by region");

        assertNotNull(response);
        assertEquals(1L, response.getFileId());
        assertEquals("Compare actual vs target by region", response.getInstruction());
        assertEquals("Regional Actual vs Target", response.getWorkflowName());
        assertEquals(5, response.getSteps().size());

        // Verify every step's operation is in the catalog
        for (WorkflowStep step : response.getSteps()) {
            assertTrue(WorkflowOperation.isValid(step.getOperation()),
                    "Operation " + step.getOperation() + " should be valid");
        }
        assertEquals("READ_SHEET", response.getSteps().get(0).getOperation());
        assertEquals("FILTER", response.getSteps().get(4).getOperation());
    }

    @Test
    void proposeWorkflow_malformedJson_throwsException() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser)).thenReturn(createMockAnalysis());
        when(workbookUnderstandingService.understandWorkbook(1L, testUser)).thenReturn(createMockUnderstanding());

        ChatResponse aiResponse = new ChatResponse(
                List.of(new Generation(new AssistantMessage("This is not JSON"))));
        when(chatModel.call(any(Prompt.class))).thenReturn(aiResponse);

        WorkflowProposalService service = new WorkflowProposalService(excelAnalyzerService,
                workbookUnderstandingService, chatModel);

        assertThrows(WorkflowProposalException.class, () -> {
            service.proposeWorkflow(1L, testUser, "analyze sales");
        });
    }

    @Test
    void proposeWorkflow_aiCallThrows_throwsException() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser)).thenReturn(createMockAnalysis());
        when(workbookUnderstandingService.understandWorkbook(1L, testUser)).thenReturn(createMockUnderstanding());
        when(chatModel.call(any(Prompt.class))).thenThrow(new RuntimeException("AI service unavailable"));

        WorkflowProposalService service = new WorkflowProposalService(excelAnalyzerService,
                workbookUnderstandingService, chatModel);

        WorkflowProposalException ex = assertThrows(WorkflowProposalException.class, () -> {
            service.proposeWorkflow(1L, testUser, "analyze sales");
        });
        assertTrue(ex.getMessage().contains("AI failed to generate workflow proposal"));
        assertTrue(ex.getMessage().contains("AI service unavailable"));
    }

    @Test
    void proposeWorkflow_fileNotOwnedByUser_throwsFileNotFoundException() {
        when(excelAnalyzerService.analyzeWorkbook(999L, testUser))
                .thenThrow(new FileNotFoundException("File not found with id: 999"));

        WorkflowProposalService service = new WorkflowProposalService(excelAnalyzerService,
                workbookUnderstandingService, chatModel);

        assertThrows(FileNotFoundException.class, () -> {
            service.proposeWorkflow(999L, testUser, "analyze sales");
        });

        // AI should never be called
        verify(chatModel, never()).call(any(Prompt.class));
        verifyNoInteractions(workbookUnderstandingService);
    }

    @Test
    void proposeWorkflow_invalidOperation_throwsException() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser)).thenReturn(createMockAnalysis());
        when(workbookUnderstandingService.understandWorkbook(1L, testUser)).thenReturn(createMockUnderstanding());

        // Step 2 uses an operation that's not in the catalog ("RUN_PYTHON")
        String aiJson = """
                {
                    "fileId": 1,
                    "instruction": "do something",
                    "workflowName": "Bad Workflow",
                    "description": "Uses a forbidden operation",
                    "steps": [
                        {
                            "order": 1,
                            "operation": "READ_SHEET",
                            "description": "Load Sales",
                            "inputSheets": ["Sales"],
                            "inputColumns": []
                        },
                        {
                            "order": 2,
                            "operation": "RUN_PYTHON",
                            "description": "Run a Python script (NOT ALLOWED)",
                            "inputSheets": ["Sales"],
                            "inputColumns": []
                        }
                    ]
                }
                """;
        ChatResponse aiResponse = new ChatResponse(List.of(new Generation(new AssistantMessage(aiJson))));
        when(chatModel.call(any(Prompt.class))).thenReturn(aiResponse);

        WorkflowProposalService service = new WorkflowProposalService(excelAnalyzerService,
                workbookUnderstandingService, chatModel);

        WorkflowProposalException ex = assertThrows(WorkflowProposalException.class, () -> {
            service.proposeWorkflow(1L, testUser, "do something");
        });
        assertTrue(ex.getMessage().contains("RUN_PYTHON"));
        assertTrue(ex.getMessage().contains("not in the allowed"));
    }

    @Test
    void proposeWorkflow_partialCatalogViolation_listsAllInvalid() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser)).thenReturn(createMockAnalysis());
        when(workbookUnderstandingService.understandWorkbook(1L, testUser)).thenReturn(createMockUnderstanding());

        // Two invalid operations in different steps
        String aiJson = """
                {
                    "fileId": 1,
                    "instruction": "do something",
                    "workflowName": "Bad Workflow",
                    "description": "Multiple bad ops",
                    "steps": [
                        {"order": 1, "operation": "READ_SHEET", "description": "ok", "inputSheets": ["Sales"], "inputColumns": []},
                        {"order": 2, "operation": "EXECUTE_SQL", "description": "bad", "inputSheets": ["Sales"], "inputColumns": []},
                        {"order": 3, "operation": "FILTER", "description": "ok", "inputSheets": ["Sales"], "inputColumns": []},
                        {"order": 4, "operation": "DELETE_DATABASE", "description": "bad", "inputSheets": ["Sales"], "inputColumns": []}
                    ]
                }
                """;
        ChatResponse aiResponse = new ChatResponse(List.of(new Generation(new AssistantMessage(aiJson))));
        when(chatModel.call(any(Prompt.class))).thenReturn(aiResponse);

        WorkflowProposalService service = new WorkflowProposalService(excelAnalyzerService,
                workbookUnderstandingService, chatModel);

        WorkflowProposalException ex = assertThrows(WorkflowProposalException.class, () -> {
            service.proposeWorkflow(1L, testUser, "do something");
        });
        assertTrue(ex.getMessage().contains("EXECUTE_SQL"));
        assertTrue(ex.getMessage().contains("DELETE_DATABASE"));
    }

    @Test
    void proposeWorkflow_emptyStepsArray_returnsEmptyList() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser)).thenReturn(createMockAnalysis());
        when(workbookUnderstandingService.understandWorkbook(1L, testUser)).thenReturn(createMockUnderstanding());

        // Valid response with empty steps (passes catalog validation, may be a
        // degenerate case)
        String aiJson = """
                {
                    "fileId": 1,
                    "instruction": "do something",
                    "workflowName": "Empty Workflow",
                    "description": "No steps proposed",
                    "steps": []
                }
                """;
        ChatResponse aiResponse = new ChatResponse(List.of(new Generation(new AssistantMessage(aiJson))));
        when(chatModel.call(any(Prompt.class))).thenReturn(aiResponse);

        WorkflowProposalService service = new WorkflowProposalService(excelAnalyzerService,
                workbookUnderstandingService, chatModel);
        WorkflowProposalResponse response = service.proposeWorkflow(1L, testUser, "do something");

        assertNotNull(response);
        assertTrue(response.getSteps().isEmpty());
    }

    @Test
    void workflowOperation_isValid_acceptsAllAllowed() {
        // Sanity check the catalog itself
        assertTrue(WorkflowOperation.isValid("READ_SHEET"));
        assertTrue(WorkflowOperation.isValid("SELECT_COLUMNS"));
        assertTrue(WorkflowOperation.isValid("REMOVE_EMPTY_ROWS"));
        assertTrue(WorkflowOperation.isValid("REMOVE_DUPLICATES"));
        assertTrue(WorkflowOperation.isValid("FILTER"));
        assertTrue(WorkflowOperation.isValid("SORT"));
        assertTrue(WorkflowOperation.isValid("CALCULATE"));
        assertTrue(WorkflowOperation.isValid("GROUP_BY"));
        assertTrue(WorkflowOperation.isValid("AGGREGATE"));
        assertTrue(WorkflowOperation.isValid("JOIN_SHEETS"));
        assertTrue(WorkflowOperation.isValid("PIVOT"));
        assertTrue(WorkflowOperation.isValid("CREATE_SUMMARY"));
        assertTrue(WorkflowOperation.isValid("CREATE_CHART"));
        assertTrue(WorkflowOperation.isValid("VALIDATE"));
        assertTrue(WorkflowOperation.isValid("GENERATE_PDF"));
        assertTrue(WorkflowOperation.isValid("EXPORT_EXCEL"));
    }

    @Test
    void workflowOperation_isValid_rejectsUnknown() {
        assertFalse(WorkflowOperation.isValid("RUN_PYTHON"));
        assertFalse(WorkflowOperation.isValid("EXECUTE_SQL"));
        assertFalse(WorkflowOperation.isValid("read_sheet")); // case-sensitive
        assertFalse(WorkflowOperation.isValid(""));
        assertFalse(WorkflowOperation.isValid(null));
        assertFalse(WorkflowOperation.isValid("   "));
    }

    @Test
    void proposeWorkflow_nullParameters_normalizedToEmptyMap() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser)).thenReturn(createMockAnalysis());
        when(workbookUnderstandingService.understandWorkbook(1L, testUser)).thenReturn(createMockUnderstanding());

        // AI returns "parameters": null on multiple steps. The proposal service
        // must normalize null -> {} so the downstream validator accepts these steps.
        String aiJson = """
                {
                    "fileId": 1,
                    "instruction": "Look at the Sales sheet",
                    "workflowName": "Load and inspect",
                    "description": "Read the sales sheet",
                    "steps": [
                        {
                            "order": 1,
                            "operation": "READ_SHEET",
                            "description": "Load Sales",
                            "inputSheets": ["Sales"],
                            "inputColumns": [],
                            "parameters": null
                        },
                        {
                            "order": 2,
                            "operation": "SELECT_COLUMNS",
                            "description": "Pick columns",
                            "inputSheets": ["Sales"],
                            "inputColumns": ["Region", "Actual"],
                            "parameters": null
                        },
                        {
                            "order": 3,
                            "operation": "REMOVE_EMPTY_ROWS",
                            "description": "Drop blanks",
                            "inputSheets": ["Sales"],
                            "inputColumns": [],
                            "parameters": null
                        },
                        {
                            "order": 4,
                            "operation": "SORT",
                            "description": "Sort by Region",
                            "inputSheets": ["Sales"],
                            "inputColumns": ["Region"],
                            "parameters": null
                        },
                        {
                            "order": 5,
                            "operation": "FILTER",
                            "description": "Filter to North",
                            "inputSheets": ["Sales"],
                            "inputColumns": ["Region"],
                            "parameters": null
                        }
                    ]
                }
                """;
        ChatResponse aiResponse = new ChatResponse(List.of(new Generation(new AssistantMessage(aiJson))));
        when(chatModel.call(any(Prompt.class))).thenReturn(aiResponse);

        WorkflowProposalService service = new WorkflowProposalService(excelAnalyzerService,
                workbookUnderstandingService, chatModel);
        WorkflowProposalResponse response = service.proposeWorkflow(1L, testUser, "Look at the Sales sheet");

        assertNotNull(response);
        assertEquals(5, response.getSteps().size());

        // Every step must have a non-null, empty parameters map.
        for (int i = 0; i < response.getSteps().size(); i++) {
            WorkflowStep step = response.getSteps().get(i);
            assertNotNull(step.getParameters(),
                    "Step " + (i + 1) + " parameters should be normalized to empty map, was null");
            assertTrue(step.getParameters().isEmpty(),
                    "Step " + (i + 1) + " parameters should be empty, was: " + step.getParameters());
        }
    }

    @Test
    void proposeWorkflow_mixedNullAndPresentParameters_normalizesOnlyNull() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser)).thenReturn(createMockAnalysis());
        when(workbookUnderstandingService.understandWorkbook(1L, testUser)).thenReturn(createMockUnderstanding());

        // Step 1 omits parameters (Lombok null), step 2 has parameters, step 3
        // explicitly uses null. Steps 1 and 3 must be normalized; step 2 must
        // keep its values verbatim.
        String aiJson = """
                {
                    "fileId": 1,
                    "instruction": "Mixed params test",
                    "workflowName": "Mixed",
                    "description": "Mixed",
                    "steps": [
                        {
                            "order": 1,
                            "operation": "READ_SHEET",
                            "description": "Load",
                            "inputSheets": ["Sales"],
                            "inputColumns": []
                        },
                        {
                            "order": 2,
                            "operation": "FILTER",
                            "description": "Filter",
                            "inputSheets": ["Sales"],
                            "inputColumns": ["Region"],
                            "parameters": {"column": "Region", "operator": "equals", "value": "North"}
                        },
                        {
                            "order": 3,
                            "operation": "SORT",
                            "description": "Sort",
                            "inputSheets": ["Sales"],
                            "inputColumns": ["Actual"],
                            "parameters": null
                        }
                    ]
                }
                """;
        ChatResponse aiResponse = new ChatResponse(List.of(new Generation(new AssistantMessage(aiJson))));
        when(chatModel.call(any(Prompt.class))).thenReturn(aiResponse);

        WorkflowProposalService service = new WorkflowProposalService(excelAnalyzerService,
                workbookUnderstandingService, chatModel);
        WorkflowProposalResponse response = service.proposeWorkflow(1L, testUser, "Mixed params test");

        // Step 1: parameters was omitted → null in Java → normalized to {}
        assertNotNull(response.getSteps().get(0).getParameters());
        assertTrue(response.getSteps().get(0).getParameters().isEmpty());

        // Step 2: parameters present → kept verbatim
        Map<String, String> step2Params = response.getSteps().get(1).getParameters();
        assertNotNull(step2Params);
        assertEquals(3, step2Params.size());
        assertEquals("Region", step2Params.get("column"));
        assertEquals("equals", step2Params.get("operator"));
        assertEquals("North", step2Params.get("value"));

        // Step 3: explicit null → normalized to {}
        assertNotNull(response.getSteps().get(2).getParameters());
        assertTrue(response.getSteps().get(2).getParameters().isEmpty());
    }

    @Test
    void proposeWorkflow_promptContainsOperationParameterContract() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser)).thenReturn(createMockAnalysis());
        when(workbookUnderstandingService.understandWorkbook(1L, testUser)).thenReturn(createMockUnderstanding());

        String aiJson = """
                {
                    "fileId": 1,
                    "instruction": "Analyze sales",
                    "workflowName": "Sales",
                    "description": "Analyze",
                    "steps": [
                        {"order": 1, "operation": "READ_SHEET", "description": "Load", "inputSheets": ["Sales"], "inputColumns": [], "parameters": null}
                    ]
                }
                """;
        ChatResponse aiResponse = new ChatResponse(List.of(new Generation(new AssistantMessage(aiJson))));
        when(chatModel.call(any(Prompt.class))).thenReturn(aiResponse);

        WorkflowProposalService service = new WorkflowProposalService(excelAnalyzerService,
                workbookUnderstandingService, chatModel);
        service.proposeWorkflow(1L, testUser, "Analyze sales");

        ArgumentCaptor<Prompt> promptCaptor = ArgumentCaptor.forClass(Prompt.class);
        verify(chatModel).call(promptCaptor.capture());

        String fullPrompt = promptCaptor.getValue().getContents().toString();

        // Verify the "Parameters by operation" section is present
        assertTrue(fullPrompt.contains("Parameters by operation"),
                "Prompt should contain 'Parameters by operation' section");

        // Verify key required parameters for each requiring operation are documented
        assertTrue(fullPrompt.contains("groupColumns: required"),
                "Prompt should document GROUP_BY groupColumns parameter");
        assertTrue(fullPrompt.contains("column: required"),
                "Prompt should document required column parameter for FILTER");
        assertTrue(fullPrompt.contains("operator: required"),
                "Prompt should document required operator parameter for FILTER");
        assertTrue(fullPrompt.contains("direction: required"),
                "Prompt should document required direction parameter for SORT");
        assertTrue(fullPrompt.contains("aggregateFunction: required"),
                "Prompt should document required aggregateFunction parameter for AGGREGATE");
        assertTrue(fullPrompt.contains("math calculations"),
                "Prompt should describe AGGREGATE as for math calculations");
        assertTrue(fullPrompt.contains("groupColumns: optional"),
                "Prompt should document AGGREGATE groupColumns parameter as optional");
        assertTrue(fullPrompt.contains("does not calculate math"),
                "Prompt should describe CREATE_SUMMARY as not calculating math");
        assertTrue(fullPrompt.contains("expression: required"),
                "Prompt should document required expression parameter for CALCULATE");
        assertTrue(fullPrompt.contains("joinType: required"),
                "Prompt should document required joinType parameter for JOIN_SHEETS");
        assertTrue(fullPrompt.contains("dataColumn: required"),
                "Prompt should document required dataColumn parameter for CREATE_CHART");

        // Verify no-parameter operations are listed
        assertTrue(fullPrompt.contains("READ_SHEET"),
                "Prompt should mention READ_SHEET");
        assertTrue(fullPrompt.contains("REMOVE_EMPTY_ROWS"),
                "Prompt should mention REMOVE_EMPTY_ROWS");

        // Verify the instruction not to substitute inputColumns for required parameters
        assertTrue(fullPrompt.contains("inputColumns"),
                "Prompt should mention inputColumns");
        assertTrue(fullPrompt.contains("NOT a substitute"),
                "Prompt should warn that inputColumns is not a substitute for required parameters");
    }
}
