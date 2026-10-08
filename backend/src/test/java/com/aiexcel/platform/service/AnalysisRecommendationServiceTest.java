package com.aiexcel.platform.service;

import com.aiexcel.platform.dto.*;
import com.aiexcel.platform.entity.User;
import com.aiexcel.platform.exception.AnalysisRecommendationException;
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
class AnalysisRecommendationServiceTest {

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
                                                .build(),
                                        ColumnAnalysis.builder()
                                                .columnName("Product")
                                                .columnIndex(2)
                                                .detectedType(ColumnDataType.TEXT)
                                                .nonEmptyCount(100)
                                                .emptyCount(0)
                                                .missingPercentage(0.0)
                                                .uniqueCount(20)
                                                .duplicateCount(80)
                                                .sampleValues(List.of("Widget"))
                                                .build()
                                ))
                                .build()
                ))
                .build();
    }

    private WorkbookUnderstandingResponse createMockUnderstanding() {
        return WorkbookUnderstandingResponse.builder()
                .summary("Sales data for Q1 2024")
                .businessDomain("Sales")
                .importantColumns(List.of(
                        ImportantColumn.builder()
                                .columnName("Amount")
                                .sheetName("Sales")
                                .reason("Primary revenue metric")
                                .build()
                ))
                .possibleRelationships(List.of())
                .dataQualityObservations(List.of(
                        DataQualityObservation.builder()
                                .observation("5% missing values")
                                .affectedEntity("Amount")
                                .severity("WARNING")
                                .build()
                ))
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
    void generateRecommendations_noApiKey_throwsException() {
        AnalysisRecommendationService service =
                new AnalysisRecommendationService(excelAnalyzerService, workbookUnderstandingService, null);

        AnalysisRecommendationException ex = assertThrows(AnalysisRecommendationException.class, () -> {
            service.generateRecommendations(1L, testUser);
        });
        assertEquals(
                "AI recommendations are not configured. Set spring.ai.openai.api-key to enable this feature.",
                ex.getMessage()
        );

        // Neither service should be called
        verifyNoInteractions(excelAnalyzerService);
        verifyNoInteractions(workbookUnderstandingService);
    }

    @Test
    void generateRecommendations_validAiResponse_returnsRecommendations() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser)).thenReturn(createMockAnalysis());
        when(workbookUnderstandingService.understandWorkbook(1L, testUser)).thenReturn(createMockUnderstanding());

        String aiJson = """
                {
                    "fileId": 1,
                    "fileName": "sales.xlsx",
                    "overview": "Focus on revenue trends and top products for the period.",
                    "recommendations": [
                        {
                            "title": "Monthly Revenue Trend",
                            "description": "Plot total revenue over time to identify growth patterns and seasonal dips.",
                            "analysisType": "TREND",
                            "relevantSheets": ["Sales"],
                            "relevantColumns": ["Date", "Amount"],
                            "businessValue": "Enables the business to understand revenue trajectory and plan accordingly."
                        },
                        {
                            "title": "Top Products by Revenue",
                            "description": "Rank products by total sales amount to identify the most profitable offerings.",
                            "analysisType": "COMPARISON",
                            "relevantSheets": ["Sales"],
                            "relevantColumns": ["Product", "Amount"],
                            "businessValue": "Helps focus marketing and inventory resources on high-value products."
                        }
                    ]
                }
                """;
        ChatResponse aiResponse = new ChatResponse(List.of(new Generation(new AssistantMessage(aiJson))));
        when(chatModel.call(any(Prompt.class))).thenReturn(aiResponse);

        AnalysisRecommendationService service =
                new AnalysisRecommendationService(excelAnalyzerService, workbookUnderstandingService, chatModel);
        AnalysisRecommendationResponse response = service.generateRecommendations(1L, testUser);

        assertNotNull(response);
        assertEquals(1L, response.getFileId());
        assertEquals("sales.xlsx", response.getFileName());
        assertEquals(2, response.getRecommendations().size());

        AnalysisRecommendation first = response.getRecommendations().get(0);
        assertEquals("Monthly Revenue Trend", first.getTitle());
        assertEquals("TREND", first.getAnalysisType());
        assertEquals(List.of("Sales"), first.getRelevantSheets());
        assertTrue(first.getRelevantColumns().contains("Date"));
    }

    @Test
    void generateRecommendations_malformedJson_throwsException() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser)).thenReturn(createMockAnalysis());
        when(workbookUnderstandingService.understandWorkbook(1L, testUser)).thenReturn(createMockUnderstanding());

        ChatResponse aiResponse = new ChatResponse(
                List.of(new Generation(new AssistantMessage("This is definitely not JSON")))
        );
        when(chatModel.call(any(Prompt.class))).thenReturn(aiResponse);

        AnalysisRecommendationService service =
                new AnalysisRecommendationService(excelAnalyzerService, workbookUnderstandingService, chatModel);

        assertThrows(AnalysisRecommendationException.class, () -> {
            service.generateRecommendations(1L, testUser);
        });
    }

    @Test
    void generateRecommendations_aiCallThrows_throwsException() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser)).thenReturn(createMockAnalysis());
        when(workbookUnderstandingService.understandWorkbook(1L, testUser)).thenReturn(createMockUnderstanding());
        when(chatModel.call(any(Prompt.class))).thenThrow(new RuntimeException("AI service unavailable"));

        AnalysisRecommendationService service =
                new AnalysisRecommendationService(excelAnalyzerService, workbookUnderstandingService, chatModel);

        AnalysisRecommendationException ex = assertThrows(AnalysisRecommendationException.class, () -> {
            service.generateRecommendations(1L, testUser);
        });
        assertEquals("AI failed to generate analysis recommendations", ex.getMessage());
    }

    @Test
    void generateRecommendations_fileNotOwnedByUser_throwsFileNotFoundException() {
        when(excelAnalyzerService.analyzeWorkbook(999L, testUser))
                .thenThrow(new FileNotFoundException("File not found with id: 999"));

        AnalysisRecommendationService service =
                new AnalysisRecommendationService(excelAnalyzerService, workbookUnderstandingService, chatModel);

        assertThrows(FileNotFoundException.class, () -> {
            service.generateRecommendations(999L, testUser);
        });

        // AI should never be called
        verify(chatModel, never()).call(any(Prompt.class));
        // Understanding service should also not be called (analysis throws first)
        verifyNoInteractions(workbookUnderstandingService);
    }

    @Test
    void generateRecommendations_aiReturnsEmptyRecommendations_returnsEmptyList() {
        when(excelAnalyzerService.analyzeWorkbook(1L, testUser)).thenReturn(createMockAnalysis());
        when(workbookUnderstandingService.understandWorkbook(1L, testUser)).thenReturn(createMockUnderstanding());

        String aiJson = """
                {
                    "fileId": 1,
                    "fileName": "sales.xlsx",
                    "overview": "No specific recommendations could be generated.",
                    "recommendations": []
                }
                """;
        ChatResponse aiResponse = new ChatResponse(List.of(new Generation(new AssistantMessage(aiJson))));
        when(chatModel.call(any(Prompt.class))).thenReturn(aiResponse);

        AnalysisRecommendationService service =
                new AnalysisRecommendationService(excelAnalyzerService, workbookUnderstandingService, chatModel);
        AnalysisRecommendationResponse response = service.generateRecommendations(1L, testUser);

        assertNotNull(response);
        assertEquals(1L, response.getFileId());
        assertTrue(response.getRecommendations().isEmpty());
    }
}
