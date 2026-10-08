package com.aiexcel.platform.service.workflow;

import com.aiexcel.platform.dto.StepExecutionResult;
import com.aiexcel.platform.dto.WorkflowStep;
import com.aiexcel.platform.exception.WorkflowExecutionException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Handler for {@code CALCULATE}.
 * <p>
 * Adds a computed column to each row using a simple arithmetic expression.
 * Supported parameters in {@code step.parameters}:
 * <ul>
 *   <li>{@code targetColumn} — required: name for the new column</li>
 *   <li>{@code expression} — required: expression referencing existing columns by name,
 *       e.g. {@code Actual - Target} or {@code Actual / Target * 100}</li>
 *   <li>{@code format} — optional: {@code integer} (rounds), {@code percent} (×100),
 *       or {@code decimal} (default)</li>
 * </ul>
 * <p>
 * Supported operators: +, -, *, /. Parentheses are NOT supported.
 * Column references in the expression are matched by exact header name.
 * Cells whose referenced cells are non-numeric produce an empty result.
 */
@Component
@Slf4j
public class CalculateHandler implements WorkflowOperationHandler {

    private static final Pattern TOKEN = Pattern.compile("([A-Za-z_][A-Za-z0-9_]*(?:\\s+[A-Za-z_][A-Za-z0-9_]*)*)|([0-9]+\\.?[0-9]*)|([+\\-*/])|(\\S)");

    @Override
    public String getSupportedOperation() {
        return "CALCULATE";
    }

    @Override
    public StepExecutionResult execute(WorkflowStep step, ExecutionContext context) {
        Workbook output = context.getOutputWorkbook();

        String sourceSheetName = ExecutionUtils.resolveSourceSheet(step, context);
        Sheet sourceSheet = ExecutionUtils.findSheet(context, sourceSheetName, step.getOrder());

        Map<String, String> params = step.getParameters();
        if (params == null) {
            throw new WorkflowExecutionException(
                    "CALCULATE step " + step.getOrder() + ": parameters must be provided");
        }

        String targetColumn = params.get("targetColumn");
        if (targetColumn == null || targetColumn.isBlank()) {
            throw new WorkflowExecutionException(
                    "CALCULATE step " + step.getOrder() + ": parameter 'targetColumn' is required");
        }

        String expression = params.get("expression");
        if (expression == null || expression.isBlank()) {
            throw new WorkflowExecutionException(
                    "CALCULATE step " + step.getOrder() + ": parameter 'expression' is required");
        }

        String format = params.getOrDefault("format", "decimal").trim().toLowerCase();
        if (!ALLOWED_FORMATS.contains(format)) {
            throw new WorkflowExecutionException(
                    "CALCULATE step " + step.getOrder() + ": parameter 'format' must be one of "
                            + String.join(", ", ALLOWED_FORMATS) + ", got '" + format + "'");
        }

        // Parse expression to find column references and validate them
        List<String> referencedColumns = parseReferencedColumns(expression);
        if (referencedColumns.isEmpty()) {
            throw new WorkflowExecutionException(
                    "CALCULATE step " + step.getOrder() + ": expression '" + expression
                            + "' does not reference any column names");
        }

        // Read header and build column name -> index map
        int firstRow = sourceSheet.getFirstRowNum();
        if (firstRow < 0) {
            throw new WorkflowExecutionException(
                    "CALCULATE step " + step.getOrder() + ": source sheet '" + sourceSheetName + "' is empty");
        }
        Row header = sourceSheet.getRow(firstRow);
        if (header == null) {
            throw new WorkflowExecutionException(
                    "CALCULATE step " + step.getOrder() + ": source sheet '" + sourceSheetName + "' has no header row");
        }

        int colCount = header.getLastCellNum();
        List<String> headers = new ArrayList<>();
        for (int c = 0; c < colCount; c++) {
            Cell hc = header.getCell(c);
            headers.add(hc != null && hc.getCellType() == CellType.STRING
                    ? hc.getStringCellValue().trim() : "");
        }

        // Validate all referenced columns exist
        for (String refCol : referencedColumns) {
            if (!headers.contains(refCol)) {
                throw new WorkflowExecutionException(
                        "CALCULATE step " + step.getOrder() + ": referenced column '" + refCol
                                + "' not found in sheet '" + sourceSheetName + "'");
            }
        }

        int targetColIndex = colCount; // append at end
        String outSheetName = sourceSheetName + "_calc";
        int suffix = 1;
        while (output.getSheet(outSheetName) != null) {
            outSheetName = sourceSheetName + "_calc_" + (++suffix);
        }
        Sheet outSheet = output.createSheet(outSheetName);

        // Copy header + new column name
        Row outHeader = outSheet.createRow(0);
        for (int c = 0; c < colCount; c++) {
            Cell srcCell = header.getCell(c);
            Cell newCell = outHeader.createCell(c);
            if (srcCell != null) copyCellValue(srcCell, newCell);
        }
        outHeader.createCell(targetColIndex).setCellValue(targetColumn);

        // Compile expression into tokens for evaluation
        List<Token> tokens = tokenize(expression);
        int[] colIndices = new int[referencedColumns.size()];
        for (int i = 0; i < referencedColumns.size(); i++) {
            colIndices[i] = headers.indexOf(referencedColumns.get(i));
        }

        int lastRow = sourceSheet.getLastRowNum();
        int computedRows = 0;
        int errorRows = 0;

        for (int r = firstRow + 1; r <= lastRow; r++) {
            Row srcRow = sourceSheet.getRow(r);
            if (srcRow == null) continue;

            Row outRow = outSheet.createRow(outSheet.getLastRowNum() + 1);

            // Copy source cells
            for (int c = 0; c < colCount; c++) {
                Cell srcCell = srcRow.getCell(c);
                if (srcCell == null) continue;
                Cell newCell = outRow.createCell(c);
                copyCellValue(srcCell, newCell);
            }

            // Evaluate expression for this row
            double[] values = new double[referencedColumns.size()];
            boolean allNumeric = true;
            for (int i = 0; i < referencedColumns.size(); i++) {
                Cell cell = srcRow.getCell(colIndices[i]);
                if (cell == null || cell.getCellType() != CellType.NUMERIC) {
                    allNumeric = false;
                    break;
                }
                values[i] = cell.getNumericCellValue();
            }

            Cell resultCell = outRow.createCell(targetColIndex);
            if (allNumeric) {
                try {
                    double result = evaluate(tokens, values);
                    switch (format) {
                        case "integer" -> resultCell.setCellValue(Math.round(result));
                        case "percent" -> resultCell.setCellValue(result * 100.0);
                        default -> resultCell.setCellValue(result);
                    }
                    computedRows++;
                } catch (ArithmeticException e) {
                    errorRows++;
                } catch (Exception e) {
                    errorRows++;
                }
            }
            // else leave cell blank
        }

        String summary = String.format(
                "Computed '%s' = %s → sheet '%s' (%d rows computed, %d errors)",
                targetColumn, expression, outSheetName, computedRows, errorRows);
        log.debug("CALCULATE step {}: {}", step.getOrder(), summary);

        context.setActiveSheet(outSheetName);

        return StepExecutionResult.builder()
                .order(step.getOrder())
                .operation(step.getOperation())
                .sheetName(outSheetName)
                .rowCount(computedRows)
                .summary(summary)
                .nextActiveSheet(outSheetName)
                .build();
    }

    private List<String> parseReferencedColumns(String expr) {
        List<String> cols = new ArrayList<>();
        Matcher m = TOKEN.matcher(expr);
        while (m.find()) {
            if (m.group(1) != null) { // identifier = column name
                String name = m.group(1);
                if (!OPERATORS.contains(name) && !KEYWORDS.contains(name)) {
                    if (!cols.contains(name)) cols.add(name);
                }
            }
        }
        return cols;
    }

    private List<Token> tokenize(String expr) {
        List<Token> tokens = new ArrayList<>();
        Matcher m = TOKEN.matcher(expr);
        while (m.find()) {
            if (m.group(1) != null) {
                tokens.add(new Token(TokenType.COLUMN, m.group(1)));
            } else if (m.group(2) != null) {
                tokens.add(new Token(TokenType.NUMBER, m.group(2)));
            } else if (m.group(3) != null) {
                tokens.add(new Token(TokenType.OPERATOR, m.group(3)));
            }
            // group(4) is invalid token — skip
        }
        return tokens;
    }

    private double evaluate(List<Token> tokens, double[] columnValues) {
        // Simple recursive-descent evaluator for left-to-right precedence: + and - then * and /
        if (tokens.isEmpty()) return 0;
        return evalAddSub(tokens, columnValues, new int[]{0});
    }

    private double evalAddSub(List<Token> tokens, double[] values, int[] pos) {
        double left = evalMulDiv(tokens, values, pos);
        while (pos[0] < tokens.size()) {
            Token op = tokens.get(pos[0]);
            if (op.type != TokenType.OPERATOR || (!op.text.equals("+") && !op.text.equals("-"))) break;
            pos[0]++;
            double right = evalMulDiv(tokens, values, pos);
            left = op.text.equals("+") ? left + right : left - right;
        }
        return left;
    }

    private double evalMulDiv(List<Token> tokens, double[] values, int[] pos) {
        double left = evalAtom(tokens, values, pos);
        while (pos[0] < tokens.size()) {
            Token op = tokens.get(pos[0]);
            if (op.type != TokenType.OPERATOR || (!op.text.equals("*") && !op.text.equals("/"))) break;
            pos[0]++;
            double right = evalAtom(tokens, values, pos);
            if (op.text.equals("/") && right == 0.0) throw new ArithmeticException("division by zero");
            left = op.text.equals("*") ? left * right : left / right;
        }
        return left;
    }

    private double evalAtom(List<Token> tokens, double[] values, int[] pos) {
        if (pos[0] >= tokens.size()) return 0;
        Token t = tokens.get(pos[0]);
        pos[0]++;
        if (t.type == TokenType.NUMBER) {
            return Double.parseDouble(t.text);
        }
        if (t.type == TokenType.COLUMN) {
            // Find column index in the original parse order
            // tokens were built in parse order, so we map column name to its position
            int colIdx = columnIndexByName(t.text, tokens, values);
            return values[colIdx];
        }
        return 0;
    }

    private int columnIndexByName(String name, List<Token> tokens, double[] values) {
        int idx = 0;
        for (Token t : tokens) {
            if (t.type == TokenType.COLUMN) {
                if (t.text.equals(name)) return idx;
                idx++;
            }
        }
        return 0;
    }

    private void copyCellValue(Cell source, Cell target) {
        switch (source.getCellType()) {
            case STRING -> target.setCellValue(source.getStringCellValue());
            case NUMERIC -> target.setCellValue(source.getNumericCellValue());
            case BOOLEAN -> target.setCellValue(source.getBooleanCellValue());
            case BLANK -> target.setCellType(CellType.BLANK);
            case FORMULA -> {
                try { target.setCellFormula(source.getCellFormula()); }
                catch (Exception ignored) { target.setCellType(CellType.BLANK); }
            }
            default -> {}
        }
    }

    private enum TokenType { COLUMN, NUMBER, OPERATOR }

    private record Token(TokenType type, String text) {}

    private static final Set<String> ALLOWED_FORMATS = Set.of("decimal", "integer", "percent");
    private static final Set<String> OPERATORS = Set.of("+", "-", "*", "/");
    private static final Set<String> KEYWORDS = Set.of();
}
