package com.aiexcel.platform.service.workflow;

import com.aiexcel.platform.dto.StepExecutionResult;
import com.aiexcel.platform.dto.WorkflowStep;
import com.aiexcel.platform.exception.WorkflowExecutionException;
import lombok.extern.slf4j.Slf4j;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Handler for {@code GENERATE_PDF}.
 * <p>
 * A deterministic workflow step that exports the current workflow output
 * workbook to a PDF file. The PDF is written as a sibling of the existing
 * XLSX output, with the same base name but a {@code .pdf} extension:
 * <pre>
 *   data.xlsx           → data_output.xlsx   (workflow XLSX)
 *                          data_output.pdf    (this handler)
 * </pre>
 * <p>
 * The original input workbook is NEVER modified and the existing
 * {@code _output.xlsx} is NEVER modified. The handler produces a new
 * file on disk and reports its path in the step summary.
 * <p>
 * Supported parameters in {@code step.parameters}:
 * <ul>
 *   <li>{@code sourceSheet} — optional: name of the source sheet; defaults to
 *       the context's active sheet (same convention as CREATE_SUMMARY)</li>
 *   <li>{@code title} — optional: report title; defaults to
 *       "Workflow Report: {sheetName}"</li>
 *   <li>{@code pageSize} — optional: rows per page (default 30, max 200)</li>
 * </ul>
 * <p>
 * V1 design notes:
 * <ul>
 *   <li>Tabular layout only — no charts, images, or styling beyond bold headers.</li>
 *   <li>Pagination is row-based; columns wider than the page are truncated.</li>
 *   <li>No AI; rendering is fully deterministic in Java using PDFBox.</li>
 * </ul>
 */
@Component
@Slf4j
public class GeneratePdfHandler implements WorkflowOperationHandler {

    private static final int DEFAULT_ROWS_PER_PAGE = 30;
    private static final int MAX_ROWS_PER_PAGE = 200;
    private static final int MAX_COLUMNS = 12;
    private static final int MAX_CELL_CHARS = 60;

    // Layout constants (PDF user-space units, 72/inch)
    private static final float MARGIN = 36f;       // 0.5 inch
    private static final float TITLE_FONT_SIZE = 16f;
    private static final float HEADER_FONT_SIZE = 10f;
    private static final float BODY_FONT_SIZE = 9f;
    private static final float META_FONT_SIZE = 9f;
    private static final float ROW_HEIGHT = 14f;
    private static final float COL_WIDTH = 60f;    // narrow columns to fit many on a page

    @Override
    public String getSupportedOperation() {
        return "GENERATE_PDF";
    }

    @Override
    public StepExecutionResult execute(WorkflowStep step, ExecutionContext context) {
        Map<String, String> params = step.getParameters() == null
                ? Map.of() : step.getParameters();

        // ── Resolve source sheet ─────────────────────────────────────────────
        String explicitSource = params.get("sourceSheet");
        String sourceSheetName = (explicitSource != null && !explicitSource.isBlank())
                ? explicitSource.trim()
                : ExecutionUtils.resolveSourceSheet(step, context);
        Sheet sourceSheet = ExecutionUtils.findSheet(context, sourceSheetName, step.getOrder());

        // ── Resolve output PDF path ─────────────────────────────────────────
        String sourcePath = context.getSourcePath();
        if (sourcePath == null || sourcePath.isBlank()) {
            throw new WorkflowExecutionException(
                    "GENERATE_PDF step " + step.getOrder()
                            + ": source path is not available in context");
        }
        Path pdfPath = derivePdfSiblingPath(Paths.get(sourcePath), step.getOrder());

        // ── Resolve rows-per-page ───────────────────────────────────────────
        int rowsPerPage = DEFAULT_ROWS_PER_PAGE;
        String pageSizeRaw = params.get("pageSize");
        if (pageSizeRaw != null && !pageSizeRaw.isBlank()) {
            try {
                int parsed = Integer.parseInt(pageSizeRaw.trim());
                if (parsed <= 0) {
                    throw new WorkflowExecutionException(
                            "GENERATE_PDF step " + step.getOrder()
                                    + ": 'pageSize' must be a positive integer, got " + pageSizeRaw);
                }
                rowsPerPage = Math.min(parsed, MAX_ROWS_PER_PAGE);
            } catch (NumberFormatException e) {
                throw new WorkflowExecutionException(
                        "GENERATE_PDF step " + step.getOrder()
                                + ": 'pageSize' must be a numeric value, got '" + pageSizeRaw + "'");
            }
        }

        // ── Read header row ─────────────────────────────────────────────────
        int firstRow = sourceSheet.getFirstRowNum();
        int lastRow = sourceSheet.getLastRowNum();
        List<String> headers = new ArrayList<>();
        if (firstRow >= 0) {
            Row headerRow = sourceSheet.getRow(firstRow);
            if (headerRow != null) {
                int colCount = Math.min(headerRow.getLastCellNum(), MAX_COLUMNS);
                for (int c = 0; c < colCount; c++) {
                    Cell hc = headerRow.getCell(c);
                    headers.add(hc != null && hc.getCellType() == CellType.STRING
                            ? hc.getStringCellValue().trim() : "");
                }
            }
        }

        // ── Read data rows ──────────────────────────────────────────────────
        List<List<String>> dataRows = new ArrayList<>();
        if (firstRow >= 0 && lastRow > firstRow) {
            for (int r = firstRow + 1; r <= lastRow; r++) {
                Row srcRow = sourceSheet.getRow(r);
                if (srcRow == null) continue;
                List<String> cells = new ArrayList<>();
                for (int c = 0; c < headers.size(); c++) {
                    Cell cell = srcRow.getCell(c);
                    cells.add(formatCell(cell));
                }
                dataRows.add(cells);
            }
        }

        // ── Determine title ─────────────────────────────────────────────────
        String title = params.get("title");
        if (title == null || title.isBlank()) {
            title = "Workflow Report: " + sourceSheetName;
        }

        // ── Render PDF ──────────────────────────────────────────────────────
        try {
            renderPdf(pdfPath, title, sourceSheetName, step, context, headers, dataRows, rowsPerPage);
        } catch (IOException e) {
            throw new WorkflowExecutionException(
                    "GENERATE_PDF step " + step.getOrder()
                            + ": failed to write PDF — " + e.getMessage(), e);
        }

        int pageCount = computePageCount(dataRows.size(), rowsPerPage);
        String summary = String.format(
                "GENERATE_PDF → '%s' (%d columns, %d rows, %d page(s)) — title: %s",
                pdfPath.getFileName().toString(), headers.size(), dataRows.size(), pageCount, title);
        log.debug("GENERATE_PDF step {}: {}", step.getOrder(), summary);

        // Do NOT change activeSheet — this is a side-effect operation only.
        return StepExecutionResult.builder()
                .order(step.getOrder())
                .operation(step.getOperation())
                .sheetName(sourceSheetName)
                .columns(headers)
                .rowCount(dataRows.size())
                .summary(summary)
                .nextActiveSheet(sourceSheetName)
                .build();
    }

    // -------------------------------------------------------------------------
    // Path derivation
    // -------------------------------------------------------------------------

    /**
     * Derives a sibling PDF path from the source file path using the same
     * "_output" naming rule as {@link ExecutionUtils#deriveOutputFileName(Path)},
     * but with a {@code .pdf} extension. The parent directory is preserved.
     */
    private Path derivePdfSiblingPath(Path sourcePath, int stepOrder) {
        if (sourcePath == null) {
            throw new WorkflowExecutionException(
                    "GENERATE_PDF step " + stepOrder + ": cannot derive PDF path from null source");
        }
        String name = sourcePath.getFileName().toString();
        int dot = name.lastIndexOf('.');
        if (dot > 0) {
            name = name.substring(0, dot) + "_output.pdf";
        } else {
            name = name + "_output.pdf";
        }
        Path parent = sourcePath.getParent();
        return parent != null ? parent.resolve(name) : Path.of(name);
    }

    // -------------------------------------------------------------------------
    // PDF rendering
    // -------------------------------------------------------------------------

    private void renderPdf(Path pdfPath,
                           String title,
                           String sourceSheetName,
                           WorkflowStep step,
                           ExecutionContext context,
                           List<String> headers,
                           List<List<String>> dataRows,
                           int rowsPerPage) throws IOException {

        // Ensure the parent directory exists (uploads dir always does, but be safe)
        Path parent = pdfPath.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }

        PDFont titleFont = PDType1Font.HELVETICA_BOLD;
        PDFont headerFont = PDType1Font.HELVETICA_BOLD;
        PDFont bodyFont = PDType1Font.HELVETICA;
        PDFont metaFont = PDType1Font.HELVETICA;

        try (PDDocument document = new PDDocument()) {
            int totalDataRows = dataRows.size();
            int totalPages = computePageCount(totalDataRows, rowsPerPage);

            if (totalPages == 0) {
                // No data → still emit one page with title and "no data" note
                drawCoverPage(document, titleFont, title, sourceSheetName, step);
            } else {
                for (int pageIdx = 0; pageIdx < totalPages; pageIdx++) {
                    int start = pageIdx * rowsPerPage;
                    int end = Math.min(start + rowsPerPage, totalDataRows);
                    List<List<String>> pageRows = dataRows.subList(start, end);
                    drawDataPage(document, pageIdx + 1, totalPages, titleFont, headerFont, bodyFont, metaFont,
                            title, sourceSheetName, step, context, headers, pageRows);
                }
            }

            document.save(pdfPath.toFile());
        }
    }

    private void drawCoverPage(PDDocument document,
                               PDFont titleFont,
                               String title,
                               String sourceSheetName,
                               WorkflowStep step) throws IOException {
        PDPage page = new PDPage(PDRectangle.LETTER);
        document.addPage(page);
        PDRectangle mediaBox = page.getMediaBox();
        float pageWidth = mediaBox.getWidth();

        try (PDPageContentStream cs = new PDPageContentStream(document, page)) {
            writeCenteredLine(cs, titleFont, TITLE_FONT_SIZE, title,
                    pageWidth / 2f, mediaBox.getHeight() - MARGIN - TITLE_FONT_SIZE);
            writeCenteredLine(cs, PDType1Font.HELVETICA, BODY_FONT_SIZE,
                    "Source sheet: " + sourceSheetName,
                    pageWidth / 2f, mediaBox.getHeight() - MARGIN - TITLE_FONT_SIZE - 24f);
            writeCenteredLine(cs, PDType1Font.HELVETICA, BODY_FONT_SIZE,
                    "No data rows in source sheet.",
                    pageWidth / 2f, mediaBox.getHeight() - MARGIN - TITLE_FONT_SIZE - 42f);
        }
    }

    private void drawDataPage(PDDocument document,
                              int pageNumber,
                              int totalPages,
                              PDFont titleFont,
                              PDFont headerFont,
                              PDFont bodyFont,
                              PDFont metaFont,
                              String title,
                              String sourceSheetName,
                              WorkflowStep step,
                              ExecutionContext context,
                              List<String> headers,
                              List<List<String>> pageRows) throws IOException {
        PDPage page = new PDPage(PDRectangle.LETTER);
        document.addPage(page);
        PDRectangle mediaBox = page.getMediaBox();
        float pageWidth = mediaBox.getWidth();
        float pageHeight = mediaBox.getHeight();

        try (PDPageContentStream cs = new PDPageContentStream(document, page)) {
            float y = pageHeight - MARGIN;

            // Title (left aligned)
            cs.beginText();
            cs.setFont(titleFont, TITLE_FONT_SIZE);
            cs.newLineAtOffset(MARGIN, y - TITLE_FONT_SIZE);
            cs.showText(safeForPdf(title));
            cs.endText();
            y -= (TITLE_FONT_SIZE + 6f);

            // Meta line
            cs.beginText();
            cs.setFont(metaFont, META_FONT_SIZE);
            cs.newLineAtOffset(MARGIN, y - META_FONT_SIZE);
            cs.showText(safeForPdf("Source sheet: " + sourceSheetName
                    + " | Step: " + step.getOrder() + " (GENERATE_PDF)"));
            cs.endText();
            y -= (META_FONT_SIZE + 10f);

            // Table header
            int visibleCols = Math.min(headers.size(), MAX_COLUMNS);
            float tableWidth = COL_WIDTH * visibleCols;
            float tableX = MARGIN;

            cs.setLineWidth(0.5f);
            cs.addRect(tableX, y - ROW_HEIGHT, tableWidth, ROW_HEIGHT);
            cs.stroke();

            // Draw header text
            for (int c = 0; c < visibleCols; c++) {
                String headerText = truncate(headers.get(c), MAX_CELL_CHARS);
                writeCellText(cs, headerFont, HEADER_FONT_SIZE, headerText,
                        tableX + c * COL_WIDTH, y, COL_WIDTH);
            }
            y -= ROW_HEIGHT;

            // Data rows
            for (List<String> row : pageRows) {
                // Skip drawing the row if we have no vertical room — but with
                // 30 rows per page on letter size we always have enough space.
                cs.addRect(tableX, y - ROW_HEIGHT, tableWidth, ROW_HEIGHT);
                cs.stroke();

                for (int c = 0; c < visibleCols; c++) {
                    String val = c < row.size() ? truncate(row.get(c), MAX_CELL_CHARS) : "";
                    writeCellText(cs, bodyFont, BODY_FONT_SIZE, val,
                            tableX + c * COL_WIDTH, y, COL_WIDTH);
                }
                y -= ROW_HEIGHT;
            }

            // Footer
            float footerY = MARGIN / 2f + 8f;
            String footer = "Page " + pageNumber + " of " + totalPages
                    + " | Generated at " + new java.util.Date();
            cs.beginText();
            cs.setFont(metaFont, META_FONT_SIZE);
            cs.newLineAtOffset(MARGIN, footerY);
            cs.showText(safeForPdf(footer));
            cs.endText();
        }
    }

    private void writeCenteredLine(PDPageContentStream cs, PDFont font, float fontSize,
                                   String text, float centerX, float baselineY) throws IOException {
        float width = font.getStringWidth(text) / 1000f * fontSize;
        cs.beginText();
        cs.setFont(font, fontSize);
        cs.newLineAtOffset(centerX - width / 2f, baselineY);
        cs.showText(safeForPdf(text));
        cs.endText();
    }

    private void writeCellText(PDPageContentStream cs, PDFont font, float fontSize,
                               String text, float cellX, float topY, float cellWidth) throws IOException {
        // Truncate to fit the cell width
        String s = truncateForWidth(font, fontSize, text, cellWidth - 4f);
        cs.beginText();
        cs.setFont(font, fontSize);
        cs.newLineAtOffset(cellX + 2f, topY - ROW_HEIGHT + 3f);
        cs.showText(safeForPdf(s));
        cs.endText();
    }

    private String truncateForWidth(PDFont font, float fontSize, String text, float maxWidth) {
        if (text == null) return "";
        // Approximate: 0.5 * fontSize per char width for Helvetica
        int maxChars = Math.max(1, (int) (maxWidth / (fontSize * 0.5f)));
        if (text.length() <= maxChars) return text;
        return text.substring(0, Math.max(1, maxChars - 1)) + "…";
    }

    /**
     * Removes characters that are illegal or unsafe in the default
     * Helvetica PDF font (any non-WinAnsi character would throw at draw time).
     */
    private String safeForPdf(String s) {
        if (s == null) return "";
        // Replace any non-WinAnsi chars (code points > 0xFF) with '?'
        StringBuilder sb = new StringBuilder(s.length());
        for (int i = 0; i < s.length(); i++) {
            char ch = s.charAt(i);
            if (ch <= 0xFF) {
                sb.append(ch);
            } else {
                sb.append('?');
            }
        }
        return sb.toString();
    }

    private String truncate(String s, int max) {
        if (s == null) return "";
        if (s.length() <= max) return s;
        return s.substring(0, Math.max(1, max - 1)) + "…";
    }

    private String formatCell(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> {
                String v = cell.getStringCellValue();
                yield v != null ? v : "";
            }
            case NUMERIC -> {
                double d = cell.getNumericCellValue();
                if (d == Math.floor(d) && !Double.isInfinite(d)
                        && d >= Long.MIN_VALUE && d <= Long.MAX_VALUE) {
                    yield String.valueOf((long) d);
                }
                yield String.valueOf(d);
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case BLANK, _NONE -> "";
            case FORMULA -> {
                try {
                    String f = cell.getCellFormula();
                    yield f != null ? "=" + f : "";
                } catch (Exception e) {
                    yield "";
                }
            }
            default -> "";
        };
    }

    private int computePageCount(int dataRowCount, int rowsPerPage) {
        if (dataRowCount == 0) return 0;
        return (dataRowCount + rowsPerPage - 1) / rowsPerPage;
    }
}
