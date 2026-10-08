package com.aiexcel.platform.service.workflow;

import com.aiexcel.platform.dto.StepExecutionResult;
import com.aiexcel.platform.dto.WorkflowStep;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.util.HashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

class AggregateHandlerTest {

    private AggregateHandler handler;
    private ExecutionContext context;
    private Workbook inputWorkbook;
    private Workbook outputWorkbook;
    private Sheet sourceSheet;

    @BeforeEach
    void setUp() {
        handler = new AggregateHandler();
        inputWorkbook = new XSSFWorkbook();
        outputWorkbook = new XSSFWorkbook();
        context = new ExecutionContext(inputWorkbook, outputWorkbook, "test.xlsx", "test_output.xlsx", "TestSheet", new java.util.ArrayList<>());
        
        sourceSheet = inputWorkbook.createSheet("TestSheet");
        Row header = sourceSheet.createRow(0);
        header.createCell(0).setCellValue("Name");
        header.createCell(1).setCellValue("Value");
        
        Row r1 = sourceSheet.createRow(1);
        r1.createCell(0).setCellValue("Item1");
        r1.createCell(1).setCellValue(50.0);
        
        Row r2 = sourceSheet.createRow(2);
        r2.createCell(0).setCellValue("Item2");
        r2.createCell(1).setCellValue(50.5);
        
        context.setActiveSheet("TestSheet");
    }

    @Test
    void execute_noGroupColumns_formatsCorrectly() {
        WorkflowStep step = new WorkflowStep();
        step.setOrder(1);
        step.setOperation("AGGREGATE");
        Map<String, String> params = new HashMap<>();
        params.put("aggregateColumn", "Value");
        params.put("aggregateFunction", "SUM");
        // groupColumns is intentionally omitted
        step.setParameters(params);

        StepExecutionResult result = handler.execute(step, context);

        assertEquals("TestSheet_agg", result.getSheetName());
        
        Sheet outSheet = outputWorkbook.getSheet("TestSheet_agg");
        assertNotNull(outSheet);
        
        Row outHeader = outSheet.getRow(0);
        assertNotNull(outHeader);
        
        // 1. exactly one header cell is created
        assertEquals(1, outHeader.getLastCellNum());
        assertEquals("Value_SUM", outHeader.getCell(0).getStringCellValue());
        
        Row outData = outSheet.getRow(1);
        assertNotNull(outData);
        
        // 2. exactly one data cell is created
        assertEquals(1, outData.getLastCellNum());
        
        // 3. the numeric result is in that data cell
        assertEquals(CellType.NUMERIC, outData.getCell(0).getCellType());
        assertEquals(100.5, outData.getCell(0).getNumericCellValue(), 0.001);
    }
}
