package com.aiexcel.platform.service;

import com.aiexcel.platform.dto.*;
import com.aiexcel.platform.entity.User;
import com.aiexcel.platform.exception.AnalysisIntentException;
import com.aiexcel.platform.exception.FileNotFoundException;
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

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
class AnalysisIntentServiceTest {

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
                .sheetCount(2)
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
                                                .build()
                                ))
                                .build()
                ))
                .build();
    }

    private WorkbookUnderstandingResponse createMockUnderstanding() {
        return WorkbookUnderstandingResponse.builder()
                .summary("Sales data with regional breakdown and targets")
                .businessDomain("Sales")
                .importantColumns(List.of(
                        ImportantColumn.builder()
                                .columnName("Region")
                                .sheetName("Sales")
                                .reason("Key segmentation dimension")
                                .build()
                ))
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
    void parseIntent_noApiKey_throwsException() {
        AnalysisIntentService service =
                new AnalysisIntentService(excelAnalyzerService, workbookUnderstandingService, null);

        AnalysisIntentException ex = assertThrows(AnalysisIntentException.class, () -> {
            service.parseIntent(1L, testUser, "any instruction");
        });
        assertEquals(
                "AI intent parsing is not configured. Set spring.ai.openai.api-key to enable this feature.",
                ex.getMessage()
        );

        // Neither service should be called
        verifyNoInteractions(excelAnalyzerService);
        verifyNoInteractions(workbookUnderstandingService);
    }

    @Test
    void parseIntent_validAiResponse_returnsIntent() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser)).thenReturn(createMockAnalysis());
        when(workbookUnderstandingService.understandWorkbook(1L, testUser)).thenReturn(createMockUnderstanding());

        String aiJson = """
                {
                    "fileId": 1,
                    "instruction": "Create a monthly sales report comparing actual sales with targets for each region and highlight regions below target.",
                    "intent": "COMPARISON",
                    "requestedAnalyses": [
                        {
                            "name": "Monthly actual sales aggregation",
                            "analysisType": "SUMMARY",
                            "description": "Aggregate actual sales by month across the Sales sheet."
                        },
                        {
                            "name": "Regional actual vs target",
                            "analysisType": "COMPARISON",
                            "description": "Compare actual vs target by region and identify regions under target."
                        }
                    ],
                    "relevantSheets": ["Sales"],
                    "relevantColumns": ["Date", "Region", "Actual", "Target"],
                    "constraints": ["monthly grain", "per-region", "highlight under-target regions"]
                }
                """;
        ChatResponse aiResponse = new ChatResponse(List.of(new Generation(new AssistantMessage(aiJson))));
        when(chatModel.call(any(Prompt.class))).thenReturn(aiResponse);

        AnalysisIntentService service =
                new AnalysisIntentService(excelAnalyzerService, workbookUnderstandingService, chatModel);
        AnalysisIntentResponse response = service.parseIntent(1L, testUser,
                "Create a monthly sales report comparing actual sales with targets for each region and highlight regions below target.");

        assertNotNull(response);
        assertEquals(1L, response.getFileId());
        assertTrue(response.getInstruction().contains("monthly sales report"));
        assertEquals("COMPARISON", response.getIntent());
        assertEquals(2, response.getRequestedAnalyses().size());
        assertEquals("COMPARISON", response.getRequestedAnalyses().get(1).getAnalysisType());
        assertTrue(response.getRelevantSheets().contains("Sales"));
        assertTrue(response.getRelevantColumns().contains("Region"));
        assertTrue(response.getRelevantColumns().contains("Target"));
        assertTrue(response.getConstraints().contains("per-region"));
    }

    @Test
    void parseIntent_malformedJson_throwsException() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser)).thenReturn(createMockAnalysis());
        when(workbookUnderstandingService.understandWorkbook(1L, testUser)).thenReturn(createMockUnderstanding());

        ChatResponse aiResponse = new ChatResponse(
                List.of(new Generation(new AssistantMessage("This is definitely not JSON")))
        );
        when(chatModel.call(any(Prompt.class))).thenReturn(aiResponse);

        AnalysisIntentService service =
                new AnalysisIntentService(excelAnalyzerService, workbookUnderstandingService, chatModel);

        assertThrows(AnalysisIntentException.class, () -> {
            service.parseIntent(1L, testUser, "some instruction");
        });
    }

    @Test
    void parseIntent_aiCallThrows_throwsException() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser)).thenReturn(createMockAnalysis());
        when(workbookUnderstandingService.understandWorkbook(1L, testUser)).thenReturn(createMockUnderstanding());
        when(chatModel.call(any(Prompt.class))).thenThrow(new RuntimeException("AI service unavailable"));

        AnalysisIntentService service =
                new AnalysisIntentService(excelAnalyzerService, workbookUnderstandingService, chatModel);

        AnalysisIntentException ex = assertThrows(AnalysisIntentException.class, () -> {
            service.parseIntent(1L, testUser, "any instruction");
        });
        assertEquals("AI failed to parse analysis intent", ex.getMessage());
    }

    @Test
    void parseIntent_fileNotOwnedByUser_throwsFileNotFoundException() {
        when(excelAnalyzerService.analyzeWorkbook(999L, testUser))
                .thenThrow(new FileNotFoundException("File not found with id: 999"));

        AnalysisIntentService service =
                new AnalysisIntentService(excelAnalyzerService, workbookUnderstandingService, chatModel);

        assertThrows(FileNotFoundException.class, () -> {
            service.parseIntent(999L, testUser, "any instruction");
        });

        // AI should never be called
        verify(chatModel, never()).call(any(Prompt.class));
        // Understanding service should also not be called (analysis throws first)
        verifyNoInteractions(workbookUnderstandingService);
    }

    @Test
    void parseIntent_blankInstruction_stillInvokesAi() {
        // Service layer: blank instruction is the controller's concern (via @NotBlank).
        // The service-level behavior is to forward whatever string is given.
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser)).thenReturn(createMockAnalysis());
        when(workbookUnderstandingService.understandWorkbook(1L, testUser)).thenReturn(createMockUnderstanding());

        String aiJson = """
                {
                    "fileId": 1,
                    "instruction": "",
                    "intent": "UNKNOWN",
                    "requestedAnalyses": [],
                    "relevantSheets": [],
                    "relevantColumns": [],
                    "constraints": []
                }
                """;
        ChatResponse aiResponse = new ChatResponse(List.of(new Generation(new AssistantMessage(aiJson))));
        when(chatModel.call(any(Prompt.class))).thenReturn(aiResponse);

        AnalysisIntentService service =
                new AnalysisIntentService(excelAnalyzerService, workbookUnderstandingService, chatModel);
        AnalysisIntentResponse response = service.parseIntent(1L, testUser, "");

        assertNotNull(response);
        // The service always stamps the verified request's instruction back into the response
        assertEquals("", response.getInstruction());
        assertTrue(response.getRequestedAnalyses().isEmpty());
    }
}
