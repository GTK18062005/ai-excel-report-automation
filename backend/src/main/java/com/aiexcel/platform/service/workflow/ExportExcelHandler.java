package com.aiexcel.platform.service.workflow;

import com.aiexcel.platform.dto.StepExecutionResult;
import com.aiexcel.platform.dto.WorkflowStep;
import com.aiexcel.platform.exception.WorkflowExecutionException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.springframework.stereotype.Component;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;

/**
 * Handler for {@code EXPORT_EXCEL}.
 * <p>
 * A deterministic workflow step that writes the current output workbook
 * (the in-memory workbook built by preceding steps) to a named XLSX export file.
 * <p>
 * The export file is distinct from the workflow's primary output file:
 * <ul>
 *   <li>Primary output: {@code basename_output.xlsx} — written by the executor
 *       at the end of a successful workflow run.</li>
 *   <li>Export file: {@code basename_exported.xlsx} (or a custom name) —
 *       written by this handler.</li>
 * </ul>
 * <p>
 * EXPORT_EXCEL can be placed anywhere in the workflow and is useful for
 * producing a focused or filtered view of the current output workbook state,
 * even if the workflow ultimately fails (in which case only the export file
 * may exist, depending on placement).
 * <p>
 * Supported parameters in {@code step.parameters}:
 * <ul>
 *   <li>{@code sheets} — optional: comma-separated list of sheet names to include
 *       in the export; if omitted, all sheets currently in the output workbook
 *       are exported</li>
 *   <li>{@code filename} — optional: base name for the export file
 *       (without path); defaults to the source file's base name with
 *       {@code _exported.xlsx} appended</li>
 * </ul>
 * <p>
 * This handler does NOT modify the input workbook or the in-memory output
 * workbook. It writes a new file to the same storage directory.
 */
@Component
@Slf4j
public class ExportExcelHandler implements WorkflowOperationHandler {

    private static final String EXPORT_SUFFIX = "_exported.xlsx";

    @Override
    public String getSupportedOperation() {
        return "EXPORT_EXCEL";
    }

    @Override
    public StepExecutionResult execute(WorkflowStep step, ExecutionContext context) {
        Map<String, String> params = step.getParameters() == null
                ? Map.of() : step.getParameters();

        Workbook output = context.getOutputWorkbook();
        if (output == null) {
            throw new WorkflowExecutionException(
                    "EXPORT_EXCEL step " + step.getOrder()
                            + ": output workbook is not available");
        }

        // ── Resolve which sheets to export ──────────────────────────────────
        List<String> sheetsToExport = resolveSheetsToExport(output, params, step.getOrder());

        // ── Resolve export filename ─────────────────────────────────────────
        String filenameParam = params.get("filename");
        Path exportPath = deriveExportPath(context.getSourcePath(), filenameParam);

        // ── Copy the selected sheets to a new workbook ──────────────────────
        try {
            writeExportFile(output, sheetsToExport, exportPath, step.getOrder());
        } catch (IOException e) {
            throw new WorkflowExecutionException(
                    "EXPORT_EXCEL step " + step.getOrder()
                            + ": failed to write export file — " + e.getMessage(), e);
        }

        // ── Build summary ──────────────────────────────────────────────────
        int totalRows = 0;
        for (String name : sheetsToExport) {
            Sheet s = output.getSheet(name);
            if (s != null) {
                totalRows += Math.max(0, s.getLastRowNum() - s.getFirstRowNum());
            }
        }
        String summary = String.format(
                "EXPORT_EXCEL → '%s' (%d sheets, ~%d rows exported)",
                exportPath.getFileName().toString(), sheetsToExport.size(), totalRows);
        log.debug("EXPORT_EXCEL step {}: {}", step.getOrder(), summary);

        return StepExecutionResult.builder()
                .order(step.getOrder())
                .operation(step.getOperation())
                .rowCount(totalRows)
                .summary(summary)
                .nextActiveSheet(context.getActiveSheet())
                .build();
    }

    // -------------------------------------------------------------------------
    // Sheet resolution
    // -------------------------------------------------------------------------

    private List<String> resolveSheetsToExport(Workbook output, Map<String, String> params, int stepOrder) {
        String sheetsRaw = params.get("sheets");
        if (sheetsRaw == null || sheetsRaw.isBlank()) {
            // No filter → export all sheets currently in the output workbook
            List<String> all = new ArrayList<>();
            for (int i = 0; i < output.getNumberOfSheets(); i++) {
                String name = output.getSheetAt(i).getSheetName();
                if (name != null) all.add(name);
            }
            return all;
        }

        // Parse the comma-separated sheet list
        List<String> requested = new ArrayList<>();
        for (String part : sheetsRaw.split(",")) {
            String trimmed = part.trim();
            if (!trimmed.isEmpty()) requested.add(trimmed);
        }

        if (requested.isEmpty()) {
            throw new WorkflowExecutionException(
                    "EXPORT_EXCEL step " + stepOrder
                            + ": 'sheets' parameter is empty");
        }

        // Verify each requested sheet exists in the output workbook
        List<String> missing = new ArrayList<>();
        List<String> result = new ArrayList<>();
        for (String name : requested) {
            if (output.getSheet(name) != null) {
                result.add(name);
            } else {
                missing.add(name);
            }
        }

        if (!missing.isEmpty()) {
            throw new WorkflowExecutionException(
                    "EXPORT_EXCEL step " + stepOrder
                            + ": sheet(s) not found in output workbook: "
                            + String.join(", ", missing));
        }

        return result;
    }

    // -------------------------------------------------------------------------
    // Path derivation
    // -------------------------------------------------------------------------

    private Path deriveExportPath(String sourcePath, String filenameParam) {
        if (sourcePath == null || sourcePath.isBlank()) {
            throw new WorkflowExecutionException(
                    "EXPORT_EXCEL: source path is not available in context");
        }
        Path src = Paths.get(sourcePath);

        String name;
        if (filenameParam != null && !filenameParam.isBlank()) {
            // Path-safety: the user-supplied filename must be a bare name, not a path.
            // We reject any input that would let the file escape the source directory.
            String trimmed = filenameParam.trim();
            if (trimmed.contains("/") || trimmed.contains("\\")
                    || trimmed.contains("..")
                    || Paths.get(trimmed).isAbsolute()) {
                throw new WorkflowExecutionException(
                        "EXPORT_EXCEL: invalid 'filename' parameter — must be a bare filename "
                                + "without path separators or '..' segments");
            }
            // Use the provided filename, preserving .xlsx extension if absent
            name = trimmed;
            if (!name.toLowerCase().endsWith(".xlsx") && !name.toLowerCase().endsWith(".xls")) {
                name = name + ".xlsx";
            }
        } else {
            // Default: replace the source extension with _exported.xlsx
            String srcName = src.getFileName().toString();
            int dot = srcName.lastIndexOf('.');
            name = (dot > 0 ? srcName.substring(0, dot) : srcName) + EXPORT_SUFFIX;
        }

        Path parent = src.getParent();
        Path resolved = parent != null ? parent.resolve(name) : Path.of(name);
        // Final guard: after resolution the export must still be in the source's
        // parent directory (defence in depth against any future logic changes).
        if (parent != null && !resolved.normalize().startsWith(parent.normalize())) {
            throw new WorkflowExecutionException(
                    "EXPORT_EXCEL: resolved export path '" + resolved
                            + "' escapes the source directory");
        }
        return resolved;
    }

    // -------------------------------------------------------------------------
    // Workbook write
    // -------------------------------------------------------------------------

    private void writeExportFile(Workbook source, List<String> sheetNames,
                                 Path exportPath, int stepOrder) throws IOException {
        // Ensure the parent directory exists
        Path parent = exportPath.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }

        String sourceFileType = detectFileType(source);
        Workbook export = newXSSFWorkbook(sourceFileType);

        try {
            for (String sheetName : sheetNames) {
                Sheet srcSheet = source.getSheet(sheetName);
                if (srcSheet == null) continue; // already verified above

                if (export instanceof XSSFWorkbook && srcSheet instanceof XSSFSheet) {
                    copyXssfSheet((XSSFWorkbook) export, (XSSFSheet) srcSheet, sheetName);
                } else {
                    copySheetValuesManually(srcSheet, export, sheetName);
                }
            }

            try (FileOutputStream fos = new FileOutputStream(exportPath.toFile())) {
                export.write(fos);
            }
        } finally {
            export.close();
        }
    }

    private String detectFileType(Workbook wb) {
        // The output workbook is always XSSF in the current executor,
        // but match the source file type for consistency
        if (wb instanceof org.apache.poi.hssf.usermodel.HSSFWorkbook) {
            return "xls";
        }
        return "xlsx";
    }

    private Workbook newXSSFWorkbook(String fileType) {
        if ("xls".equalsIgnoreCase(fileType)) {
            return new org.apache.poi.hssf.usermodel.HSSFWorkbook();
        }
        return new XSSFWorkbook();
    }

    /**
     * Copies a single XSSF sheet into the target XSSF workbook: cell values,
     * formulas, merged regions, column widths, and row heights are preserved.
     * Matches the logic in {@code WorkflowExecutionService.copyXssfSheet}.
     */
    private void copyXssfSheet(XSSFWorkbook tgtWb, XSSFSheet src, String tgtName) {
        Sheet tgt = tgtWb.createSheet(tgtName);

        // Column widths
        int srcLastCol = computeLastColumn(src);
        for (int c = 0; c <= srcLastCol; c++) {
            int w = src.getColumnWidth(c);
            if (w > 0) tgt.setColumnWidth(c, w);
        }

        int firstRow = src.getFirstRowNum();
        int lastRow = src.getLastRowNum();
        for (int r = firstRow; r <= lastRow; r++) {
            Row srcRow = src.getRow(r);
            if (srcRow == null) continue;
            Row tgtRow = tgt.createRow(r);
            tgtRow.setHeight(srcRow.getHeight());
            int firstCell = srcRow.getFirstCellNum();
            int lastCell = srcRow.getLastCellNum();
            for (int c = firstCell; c < lastCell; c++) {
                Cell srcCell = srcRow.getCell(c);
                if (srcCell == null) continue;
                Cell tgtCell = tgtRow.createCell(c);
                copyCellValue(srcCell, tgtCell);
            }
        }

        for (int m = 0; m < src.getNumMergedRegions(); m++) {
            tgt.addMergedRegion(src.getMergedRegion(m));
        }
    }

    /**
     * Copies a sheet using only the public POI cell-copy API (for non-XSSF sheets
     * or when source is HSSF). Values, formulas, merged regions, column widths,
     * and row heights are preserved.
     */
    private void copySheetValuesManually(Sheet src, Workbook tgtWb, String tgtName) {
        Sheet tgt = tgtWb.createSheet(tgtName);

        int srcLastCol = computeLastColumn(src);
        for (int c = 0; c <= srcLastCol; c++) {
            int w = src.getColumnWidth(c);
            if (w > 0) tgt.setColumnWidth(c, w);
        }

        int firstRow = src.getFirstRowNum();
        int lastRow = src.getLastRowNum();
        for (int r = firstRow; r <= lastRow; r++) {
            Row srcRow = src.getRow(r);
            if (srcRow == null) continue;
            Row tgtRow = tgt.createRow(r);
            tgtRow.setHeight(srcRow.getHeight());
            int firstCell = srcRow.getFirstCellNum();
            int lastCell = srcRow.getLastCellNum();
            for (int c = firstCell; c < lastCell; c++) {
                Cell srcCell = srcRow.getCell(c);
                if (srcCell == null) continue;
                Cell tgtCell = tgtRow.createCell(c);
                copyCellValue(srcCell, tgtCell);
            }
        }

        for (int m = 0; m < src.getNumMergedRegions(); m++) {
            tgt.addMergedRegion(src.getMergedRegion(m));
        }
    }

    private int computeLastColumn(Sheet src) {
        int max = -1;
        int firstRow = src.getFirstRowNum();
        int lastRow = src.getLastRowNum();
        for (int r = firstRow; r <= lastRow; r++) {
            Row row = src.getRow(r);
            if (row != null) {
                int rowLast = row.getLastCellNum() - 1;
                if (rowLast > max) max = rowLast;
            }
        }
        return max;
    }

    private void copyCellValue(Cell src, Cell tgt) {
        switch (src.getCellType()) {
            case STRING -> tgt.setCellValue(src.getStringCellValue());
            case NUMERIC -> tgt.setCellValue(src.getNumericCellValue());
            case BOOLEAN -> tgt.setCellValue(src.getBooleanCellValue());
            case FORMULA -> {
                try { tgt.setCellFormula(src.getCellFormula()); }
                catch (Exception ignored) { tgt.setCellType(CellType.BLANK); }
            }
            case ERROR -> tgt.setCellErrorValue(src.getErrorCellValue());
            case BLANK, _NONE -> tgt.setCellType(CellType.BLANK);
        }
    }
}
