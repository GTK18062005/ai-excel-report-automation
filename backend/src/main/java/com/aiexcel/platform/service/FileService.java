package com.aiexcel.platform.service;

import com.aiexcel.platform.dto.FileResponse;
import com.aiexcel.platform.entity.FileStatus;
import com.aiexcel.platform.entity.UploadedFile;
import com.aiexcel.platform.entity.User;
import com.aiexcel.platform.exception.FileNotFoundException;
import com.aiexcel.platform.repository.UploadedFileRepository;
import com.aiexcel.platform.service.workflow.ExecutionUtils;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;

import java.nio.file.Path;
import java.util.List;

@Service
@RequiredArgsConstructor
public class FileService {

    private final UploadedFileRepository fileRepository;
    private final FileStorageService fileStorageService;

    @Transactional
    public FileResponse uploadFile(MultipartFile file, User user) {
        FileStorageService.StoredFileMetadata stored = fileStorageService.store(file);

        UploadedFile entity = UploadedFile.builder()
                .user(user)
                .originalFileName(file.getOriginalFilename())
                .storedFileName(stored.storedFileName())
                .fileType(stored.fileType())
                .fileSize(stored.fileSize())
                .storagePath(stored.storagePath())
                .status(FileStatus.UPLOADED)
                .build();

        return toResponse(fileRepository.save(entity));
    }

    @Transactional(readOnly = true)
    public List<FileResponse> getUserFiles(User user) {
        return fileRepository
                .findByUserAndStatusNotOrderByUploadedAtDesc(user, FileStatus.DELETED)
                .stream()
                .map(this::toResponse)
                .toList();
    }

    @Transactional(readOnly = true)
    public FileResponse getFileById(Long id, User user) {
        UploadedFile file = findOwnedFile(id, user);
        return toResponse(file);
    }

    @Transactional
    public void softDeleteFile(Long id, User user) {
        UploadedFile file = findOwnedFile(id, user);
        file.setStatus(FileStatus.DELETED);
        fileRepository.save(file);
        // Physical file is left on disk — can be cleaned up by a scheduled job later
    }

    /**
     * Returns the workflow output file for the given uploaded file ID, scoped
     * to the owning user.
     * <p>
     * The output file name is derived deterministically from the source file
     * name (same convention as {@code WorkflowExecutionService}):
     * <pre>
     *   d329be04-9a09-4ecf-baf1-1aabedc461aa.xlsx
     *     → d329be04-9a09-4ecf-baf1-1aabedc461aa_output.xlsx
     * </pre>
     * <p>
     * The {@code downloadFileName} in the result is based on the original
     * user-supplied filename, not the UUID stored name (more meaningful for the
     * browser save dialog): {@code sales_sample_output.xlsx}.
     *
     * @throws FileNotFoundException if the file is not owned by the user,
     *         has been soft-deleted, or no output file exists on disk.
     */
    @Transactional(readOnly = true)
    public OutputFileResult getOutputFile(Long id, User user) {
        UploadedFile file = findOwnedFile(id, user);

        // Derive the output file name from the stored (UUID-based) source filename.
        // WorkflowExecutionService uses the same formula, so this is guaranteed
        // to match the file written at execution time.
        String outputStoredName = ExecutionUtils.deriveOutputFileName(file.getStoragePath());
        if (outputStoredName == null) {
            throw new FileNotFoundException(
                    "Cannot determine output file name for file id: " + id);
        }

        Path outputPath = fileStorageService.load(outputStoredName);
        if (outputPath == null) {
            throw new FileNotFoundException(
                    "No workflow output exists for file id: " + id
                            + " (file may not have been executed yet)");
        }

        // Build a meaningful download filename from the original name
        String downloadName = ExecutionUtils.deriveOutputFileName(file.getOriginalFileName());
        if (downloadName == null) {
            // Fallback: extract bare filename from the storage path
            String bare = outputStoredName;
            int slash = Math.max(bare.lastIndexOf('/'), bare.lastIndexOf('\\'));
            if (slash >= 0) {
                bare = bare.substring(slash + 1);
            }
            downloadName = bare;
        }

        return new OutputFileResult(outputPath, downloadName);
    }

    /**
     * Result of {@link #getOutputFile(Long, User)}: the resolved on-disk path
     * to the generated output and a user-friendly filename for download.
     */
    public record OutputFileResult(Path filePath, String downloadFileName) {}

    /**
     * Returns the PDF workflow output file for the given uploaded file ID,
     * scoped to the owning user.
     * <p>
     * The PDF is written by the {@code GENERATE_PDF} handler as a sibling of
     * the XLSX output, with the source extension replaced by {@code .pdf}:
     * <pre>
     *   d329be04-9a09-4ecf-baf1-1aabedc461aa.xlsx
     *     → d329be04-9a09-4ecf-baf1-1aabedc461aa_output.pdf
     * </pre>
     * Ownership is verified exactly as in {@link #getOutputFile(Long, User)}.
     * The path is resolved through {@link FileStorageService#load(String)},
     * which preserves the existing path-traversal protections.
     *
     * @throws FileNotFoundException if the file is not owned by the user,
     *         has been soft-deleted, or no PDF output exists on disk.
     */
    @Transactional(readOnly = true)
    public OutputFileResult getOutputPdfFile(Long id, User user) {
        UploadedFile file = findOwnedFile(id, user);

        // Derive the stored PDF filename from the stored (UUID-based) source filename.
        // This mirrors the deterministic naming used by the GENERATE_PDF handler.
        String storedSourceName = extractStoredFileName(file.getStoragePath());
        if (storedSourceName == null) {
            throw new FileNotFoundException(
                    "Cannot determine PDF output file name for file id: " + id);
        }
        String outputPdfName = derivePdfSiblingName(storedSourceName);

        Path outputPath = fileStorageService.load(outputPdfName);
        if (outputPath == null) {
            throw new FileNotFoundException(
                    "No PDF workflow output exists for file id: " + id
                            + " (GENERATE_PDF may not have been executed yet)");
        }

        // Build a meaningful download filename from the original name
        String downloadName = derivePdfSiblingName(file.getOriginalFileName());
        if (downloadName == null) {
            downloadName = outputPdfName;
        }

        return new OutputFileResult(outputPath, downloadName);
    }

    /**
     * Extracts the bare filename from a stored file path (which may be
     * absolute or relative). Returns null if the path has no filename.
     */
    private String extractStoredFileName(String storagePath) {
        if (storagePath == null || storagePath.isBlank()) return null;
        int slash = Math.max(storagePath.lastIndexOf('/'), storagePath.lastIndexOf('\\'));
        return slash >= 0 ? storagePath.substring(slash + 1) : storagePath;
    }

    /**
     * Replaces the extension of a stored filename with {@code .pdf} and
     * inserts {@code _output} before the new extension — matches the
     * deterministic naming used by {@code GENERATE_PDF}.
     * <p>
     * Examples:
     * <pre>
     *   d329be04-...xlsx        → d329be04-..._output.pdf
     *   sales_sample.xlsx       → sales_sample_output.pdf
     *   no-extension            → no-extension_output.pdf
     * </pre>
     */
    private String derivePdfSiblingName(String name) {
        if (name == null) return null;
        String base = name;
        int slash = Math.max(base.lastIndexOf('/'), base.lastIndexOf('\\'));
        if (slash >= 0) base = base.substring(slash + 1);
        int dot = base.lastIndexOf('.');
        if (dot > 0) {
            return base.substring(0, dot) + "_output.pdf";
        }
        return base + "_output.pdf";
    }

    // -------------------------------------------------------------------------
    // Private helpers
    // -------------------------------------------------------------------------

    private UploadedFile findOwnedFile(Long id, User user) {
        return fileRepository
                .findByIdAndUserAndStatusNot(id, user, FileStatus.DELETED)
                .orElseThrow(() -> new FileNotFoundException("File not found with id: " + id));
    }

    private FileResponse toResponse(UploadedFile file) {
        return FileResponse.builder()
                .id(file.getId())
                .originalFileName(file.getOriginalFileName())
                .fileType(file.getFileType())
                .fileSize(file.getFileSize())
                .status(file.getStatus())
                .uploadedAt(file.getUploadedAt())
                .build();
    }
}
