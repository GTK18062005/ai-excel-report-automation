package com.aiexcel.platform.service;

import com.aiexcel.platform.dto.*;
import com.aiexcel.platform.entity.User;
import com.aiexcel.platform.exception.FileNotFoundException;
import com.aiexcel.platform.exception.WorkbookUnderstandingException;
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
class WorkbookUnderstandingServiceTest {

    @Mock
    private ExcelAnalyzerService excelAnalyzerService;

    @Mock
    private org.springframework.ai.chat.model.ChatModel chatModel;

    private User testUser;

    private WorkbookAnalysisResponse createMockAnalysis() {
        return WorkbookAnalysisResponse.builder()
                .fileId(1L)
                .fileName("test.xlsx")
                .sheetCount(1)
                .sheets(List.of(
                        SheetAnalysis.builder()
                                .sheetName("Sales")
                                .rowCount(100)
                                .columnCount(3)
                                .columns(List.of(
                                        ColumnAnalysis.builder()
                                                .columnName("Date")
                                                .columnIndex(0)
                                                .detectedType(ColumnDataType.DATE)
                                                .nonEmptyCount(100)
                                                .emptyCount(0)
                                                .missingPercentage(0.0)
                                                .uniqueCount(10)
                                                .duplicateCount(90)
                                                .sampleValues(List.of("2024-01-01"))
                                                .build(),
                                        ColumnAnalysis.builder()
                                                .columnName("Amount")
                                                .columnIndex(1)
                                                .detectedType(ColumnDataType.DECIMAL)
                                                .nonEmptyCount(95)
                                                .emptyCount(5)
                                                .missingPercentage(5.0)
                                                .uniqueCount(50)
                                                .duplicateCount(45)
                                                .sampleValues(List.of("100.50"))
                                                .build()))
                                .build()))
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
    void understandWorkbook_validAiResponse_returnsUnderstanding() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser))
                .thenReturn(createMockAnalysis());

        // Simulate AI returning valid JSON
        String aiJson = """
                {
                    "summary": "A sales data workbook",
                    "businessDomain": "Sales",
                    "sheets": [{"sheetName": "Sales", "description": "Daily sales data", "rowCount": 100, "columnCount": 3}],
                    "importantColumns": [{"columnName": "Amount", "sheetName": "Sales", "reason": "Key metric"}],
                    "possibleRelationships": [],
                    "dataQualityObservations": [{"observation": "5% missing values", "affectedEntity": "Amount", "severity": "WARNING"}],
                    "suggestedAnalyses": [{"title": "Revenue Trend", "description": "Track revenue over time", "relevantEntities": "Amount"}]
                }
                """;
        ChatResponse aiResponse = new ChatResponse(List.of(new Generation(new AssistantMessage(aiJson))));
        when(chatModel.call(any(Prompt.class))).thenReturn(aiResponse);

        WorkbookUnderstandingService service = new WorkbookUnderstandingService(excelAnalyzerService, chatModel);
        WorkbookUnderstandingResponse response = service.understandWorkbook(1L, testUser);

        assertNotNull(response);
        assertEquals("Sales", response.getBusinessDomain());
        assertEquals(1, response.getSheets().size());
        assertEquals("Sales", response.getSheets().get(0).getSheetName());
        assertEquals(1, response.getImportantColumns().size());
        assertEquals(1, response.getDataQualityObservations().size());
        assertEquals(1, response.getSuggestedAnalyses().size());
    }

    @Test
    void understandWorkbook_malformedJson_throwsException() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser))
                .thenReturn(createMockAnalysis());

        ChatResponse aiResponse = new ChatResponse(List.of(new Generation(new AssistantMessage("This is not JSON"))));
        when(chatModel.call(any(Prompt.class))).thenReturn(aiResponse);

        WorkbookUnderstandingService service = new WorkbookUnderstandingService(excelAnalyzerService, chatModel);

        assertThrows(WorkbookUnderstandingException.class, () -> {
            service.understandWorkbook(1L, testUser);
        });
    }

    @Test
    void understandWorkbook_aiCallThrows_throwsException() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser))
                .thenReturn(createMockAnalysis());
        when(chatModel.call(any(Prompt.class)))
                .thenThrow(new RuntimeException("AI service unavailable"));

        WorkbookUnderstandingService service = new WorkbookUnderstandingService(excelAnalyzerService, chatModel);

        assertThrows(WorkbookUnderstandingException.class, () -> {
            service.understandWorkbook(1L, testUser);
        });
    }

    @Test
    void understandWorkbook_fileNotOwnedByUser_throwsFileNotFoundException() {
        when(excelAnalyzerService.analyzeWorkbook(999L, testUser))
                .thenThrow(new FileNotFoundException("File not found with id: 999"));

        WorkbookUnderstandingService service = new WorkbookUnderstandingService(excelAnalyzerService, chatModel);

        assertThrows(FileNotFoundException.class, () -> {
            service.understandWorkbook(999L, testUser);
        });

        // ChatModel should never be called
        verify(chatModel, never()).call(any(Prompt.class));
    }

    @Test
    void understandWorkbook_noApiKey_throwsException() {
        WorkbookUnderstandingService service = new WorkbookUnderstandingService(excelAnalyzerService, null);

        assertThrows(WorkbookUnderstandingException.class, () -> {
            service.understandWorkbook(1L, testUser);
        });

        // ExcelAnalyzerService should never be called
        verify(excelAnalyzerService, never()).analyzeWorkbook(any(), any());
    }

    @Test
    void understandWorkbook_jsonWithMarkdownFences_handlesCorrectly() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser))
                .thenReturn(createMockAnalysis());

        // AI response wrapped in markdown fences
        String aiJson = """
                ```json
                {
                    "summary": "A sales data workbook",
                    "businessDomain": "Sales",
                    "sheets": [{"sheetName": "Sales", "description": "Daily sales data", "rowCount": 100, "columnCount": 3}],
                    "importantColumns": [{"columnName": "Amount", "sheetName": "Sales", "reason": "Key metric"}],
                    "possibleRelationships": [],
                    "dataQualityObservations": [],
                    "suggestedAnalyses": []
                }
                ```
                """;
        ChatResponse aiResponse = new ChatResponse(List.of(new Generation(new AssistantMessage(aiJson))));
        when(chatModel.call(any(Prompt.class))).thenReturn(aiResponse);

        WorkbookUnderstandingService service = new WorkbookUnderstandingService(excelAnalyzerService, chatModel);
        WorkbookUnderstandingResponse response = service.understandWorkbook(1L, testUser);

        assertNotNull(response);
        assertEquals("Sales", response.getBusinessDomain());
        // Verify fileId and fileName were normalized to server values
        assertEquals("1", response.getFileId().toString());
        assertEquals("test.xlsx", response.getFileName());
    }

    @Test
    void understandWorkbook_jsonWithBareUuid_handlesCorrectly() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser))
                .thenReturn(createMockAnalysis());

        // AI response with bare UUID (no quotes) - known issue from the bug report
        String aiJson = """
                {
                    "fileId": 12345,
                    "fileName": "sales_sample.xlsx",
                    "summary": "A sales data workbook",
                    "businessDomain": "Sales",
                    "sheets": [{"sheetName": "Sales", "description": "Daily sales data", "rowCount": 100, "columnCount": 3}],
                    "importantColumns": [{"columnName": "Amount", "sheetName": "Sales", "reason": "Key metric"}],
                    "possibleRelationships": [],
                    "dataQualityObservations": [],
                    "suggestedAnalyses": []
                }
                """;
        ChatResponse aiResponse = new ChatResponse(List.of(new Generation(new AssistantMessage(aiJson))));
        when(chatModel.call(any(Prompt.class))).thenReturn(aiResponse);

        WorkbookUnderstandingService service = new WorkbookUnderstandingService(excelAnalyzerService, chatModel);
        WorkbookUnderstandingResponse response = service.understandWorkbook(1L, testUser);

        assertNotNull(response);
        // Server-side values should override AI values
        assertEquals("1", response.getFileId().toString());
        assertEquals("test.xlsx", response.getFileName());
    }

    @Test
    void understandWorkbook_jsonWithBareEnumValues_handlesCorrectly() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser))
                .thenReturn(createMockAnalysis());

        // AI response with bare enum values (no quotes)
        String aiJson = """
                {
                    "summary": "A sales data workbook",
                    "businessDomain": Sales,
                    "sheets": [{"sheetName": "Sales", "description": "Daily sales data", "rowCount": 100, "columnCount": 3}],
                    "importantColumns": [{"columnName": "Amount", "sheetName": "Sales", "reason": "Key metric"}],
                    "possibleRelationships": [],
                    "dataQualityObservations": [{"observation": "5% missing", "affectedEntity": "Amount", "severity": WARNING}],
                    "suggestedAnalyses": []
                }
                """;
        ChatResponse aiResponse = new ChatResponse(List.of(new Generation(new AssistantMessage(aiJson))));
        when(chatModel.call(any(Prompt.class))).thenReturn(aiResponse);

        WorkbookUnderstandingService service = new WorkbookUnderstandingService(excelAnalyzerService, chatModel);
        WorkbookUnderstandingResponse response = service.understandWorkbook(1L, testUser);

        assertNotNull(response);
        assertEquals("Sales", response.getBusinessDomain());
        assertEquals(1, response.getDataQualityObservations().size());
        assertEquals("WARNING", response.getDataQualityObservations().get(0).getSeverity());
    }

    @Test
    void understandWorkbook_jsonWithSurroundingText_extractsJsonCorrectly() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser))
                .thenReturn(createMockAnalysis());

        // AI response with surrounding explanatory text
        String aiJson = """
                Let me analyze this workbook for you.

                {
                    "summary": "A sales data workbook",
                    "businessDomain": "Sales",
                    "sheets": [{"sheetName": "Sales", "description": "Daily sales data", "rowCount": 100, "columnCount": 3}],
                    "importantColumns": [{"columnName": "Amount", "sheetName": "Sales", "reason": "Key metric"}],
                    "possibleRelationships": [],
                    "dataQualityObservations": [],
                    "suggestedAnalyses": []
                }

                This workbook contains important sales metrics.
                """;
        ChatResponse aiResponse = new ChatResponse(List.of(new Generation(new AssistantMessage(aiJson))));
        when(chatModel.call(any(Prompt.class))).thenReturn(aiResponse);

        WorkbookUnderstandingService service = new WorkbookUnderstandingService(excelAnalyzerService, chatModel);
        WorkbookUnderstandingResponse response = service.understandWorkbook(1L, testUser);

        assertNotNull(response);
        assertEquals("Sales", response.getBusinessDomain());
    }

    @Test
    void understandWorkbook_emptyAiResponse_throwsException() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser))
                .thenReturn(createMockAnalysis());

        ChatResponse aiResponse = new ChatResponse(List.of(new Generation(new AssistantMessage(""))));
        when(chatModel.call(any(Prompt.class))).thenReturn(aiResponse);

        WorkbookUnderstandingService service = new WorkbookUnderstandingService(excelAnalyzerService, chatModel);

        assertThrows(WorkbookUnderstandingException.class, () -> {
            service.understandWorkbook(1L, testUser);
        });
    }

    @Test
    void understandWorkbook_jsonWithoutBraces_throwsExpectiveException() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser))
                .thenReturn(createMockAnalysis());

        ChatResponse aiResponse = new ChatResponse(
                List.of(new Generation(new AssistantMessage("This is completely invalid: no JSON at all"))));
        when(chatModel.call(any(Prompt.class))).thenReturn(aiResponse);

        WorkbookUnderstandingService service = new WorkbookUnderstandingService(excelAnalyzerService, chatModel);

        assertThrows(WorkbookUnderstandingException.class, () -> {
            service.understandWorkbook(1L, testUser);
        });
    }
}
