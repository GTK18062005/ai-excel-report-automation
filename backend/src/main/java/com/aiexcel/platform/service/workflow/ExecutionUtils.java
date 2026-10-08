package com.aiexcel.platform.service.workflow;

import com.aiexcel.platform.dto.WorkflowStep;
import com.aiexcel.platform.exception.WorkflowExecutionException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.nio.file.Path;
import java.util.List;

/**
 * Static helpers shared by operation handlers and the file service.
 * <p>
 * These utilities must be deterministic and side-effect free; the executor
 * and the download endpoint both rely on the same naming rules so the file
 * the executor writes and the file the endpoint resolves always match.
 */
public final class ExecutionUtils {

    private ExecutionUtils() {}

    /**
     * Resolves which sheet the step should read from. Priority:
     *  1. First sheet named in {@code step.inputSheets}
     *  2. The context's {@code activeSheet}
     *  3. Throws {@link WorkflowExecutionException} if neither is available
     */
    public static String resolveSourceSheet(WorkflowStep step, ExecutionContext context) {
        List<String> inputSheets = step.getInputSheets();
        if (inputSheets != null && !inputSheets.isEmpty()) {
            String first = inputSheets.get(0);
            if (first != null && !first.isBlank()) {
                return first;
            }
        }
        if (context.getActiveSheet() != null && !context.getActiveSheet().isBlank()) {
            return context.getActiveSheet();
        }
        throw new WorkflowExecutionException(
                "Step " + step.getOrder() + " (" + step.getOperation()
                + "): cannot determine source sheet — provide inputSheets or run READ_SHEET first");
    }

    /**
     * Looks up a sheet by name, checking the input workbook first, then the
     * output workbook. This lets downstream steps continue to read sheets
     * produced by earlier steps without re-reading the original.
     *
     * @return the resolved sheet
     * @throws WorkflowExecutionException if not found in either workbook
     */
    static Sheet findSheet(ExecutionContext context, String sheetName, int stepOrder) {
        if (sheetName == null || sheetName.isBlank()) {
            throw new WorkflowExecutionException("Step " + stepOrder + ": sheet name is blank");
        }
        if (context.getInputWorkbook() != null && context.getInputWorkbook().getSheet(sheetName) != null) {
            return context.getInputWorkbook().getSheet(sheetName);
        }
        if (context.getOutputWorkbook() != null && context.getOutputWorkbook().getSheet(sheetName) != null) {
            return context.getOutputWorkbook().getSheet(sheetName);
        }
        throw new WorkflowExecutionException(
                "Step " + stepOrder + ": sheet '" + sheetName + "' not found in input or output workbook");
    }

    /**
     * Derives the workflow output filename from a stored filename string by
     * inserting {@code _output} before the file extension.
     * <p>
     * Examples:
     * <pre>
     *   d329be04-9a09-4ecf-baf1-1aabedc461aa.xlsx  →  d329be04-9a09-4ecf-baf1-1aabedc461aa_output.xlsx
     *   sales_sample.xlsx                          →  sales_sample_output.xlsx
     *   no-extension                               →  no-extension_output
     *   "archive.tar.gz"                           →  "archive.tar_output.gz"
     * </pre>
     * The rule is "last dot wins" — only the extension after the last {@code .}
     * is treated as the extension. This mirrors the convention used by the
     * workflow executor so the file written at execution time matches the
     * filename the download endpoint resolves.
     *
     * @param name source filename (may be a bare name or a path)
     * @return derived output filename, or {@code null} if {@code name} is null
     */
    public static String deriveOutputFileName(String name) {
        if (name == null) {
            return null;
        }
        // Strip any directory prefix so we always operate on the bare filename
        String base = name;
        int slash = Math.max(base.lastIndexOf('/'), base.lastIndexOf('\\'));
        if (slash >= 0) {
            base = base.substring(slash + 1);
        }
        int dot = base.lastIndexOf('.');
        if (dot > 0) {
            return base.substring(0, dot) + "_output" + base.substring(dot);
        }
        return base + "_output";
    }

    /**
     * Path-based overload of {@link #deriveOutputFileName(String)}: derives
     * the output filename from the source file's {@code Path} and returns the
     * full sibling path. The parent directory of the source file is preserved
     * so the output is always written alongside the input (even for absolute paths).
     * Used by the executor to compute the write target.
     *
     * @param sourcePath the source file's path (must not be null)
     * @return sibling path with the same parent directory and an output-suffixed name
     */
    public static Path deriveOutputFileName(Path sourcePath) {
        String name = sourcePath.getFileName().toString();
        int dot = name.lastIndexOf('.');
        if (dot > 0) {
            name = name.substring(0, dot) + "_output" + name.substring(dot);
        } else {
            name = name + "_output";
        }
        // Preserve the parent directory of the source path (unlike resolveSibling
        // which uses the CWD when the sibling has no directory component).
        Path parent = sourcePath.getParent();
        return parent != null ? parent.resolve(name) : Path.of(name);
    }
}
