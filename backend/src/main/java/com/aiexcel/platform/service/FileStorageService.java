package com.aiexcel.platform.service;

import com.aiexcel.platform.exception.InvalidFileTypeException;
import jakarta.annotation.PostConstruct;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.Set;
import java.util.UUID;

@Service
public class FileStorageService {

    private static final Set<String> ALLOWED_EXTENSIONS = Set.of("xlsx", "xls", "csv");

    private final String uploadDir;

    private Path uploadPath;

    public FileStorageService(@Value("${file.upload-dir:./uploads}") String uploadDir) {
        this.uploadDir = uploadDir;
    }

    @PostConstruct
    public void init() {
        try {
            uploadPath = Paths.get(uploadDir).toAbsolutePath().normalize();
            Files.createDirectories(uploadPath);
        } catch (IOException e) {
            throw new RuntimeException("Could not initialize upload directory: " + uploadDir, e);
        }
    }

    /**
     * Stores a multipart file securely and returns metadata about the stored file.
     * - Extension is validated against the allowlist.
     * - The stored filename is UUID-based — original name is never used on disk.
     * - Path traversal is prevented by normalizing and verifying the resolved path.
     */
    public StoredFileMetadata store(MultipartFile file) {
        String originalFilename = StringUtils.cleanPath(
                file.getOriginalFilename() != null ? file.getOriginalFilename() : "unknown"
        );

        String extension = extractExtension(originalFilename);

        if (!ALLOWED_EXTENSIONS.contains(extension)) {
            throw new InvalidFileTypeException(
                    "File type '." + extension + "' is not allowed. Accepted types: .xlsx, .xls, .csv"
            );
        }

        String storedFileName = UUID.randomUUID() + "." + extension;

        Path targetPath = uploadPath.resolve(storedFileName).normalize();

        // Guard against path traversal
        if (!targetPath.startsWith(uploadPath)) {
            throw new RuntimeException("Illegal storage path detected — aborting upload");
        }

        try {
            Files.copy(file.getInputStream(), targetPath, StandardCopyOption.REPLACE_EXISTING);
        } catch (IOException e) {
            throw new RuntimeException("Failed to store file '" + originalFilename + "'", e);
        }

        return new StoredFileMetadata(storedFileName, targetPath.toString(), extension, file.getSize());
    }

    /**
     * Deletes the physical file from disk. Path traversal guarded.
     * Silently ignores missing files since the entity soft-delete still proceeds.
     */
    public void delete(String storedFileName) {
        try {
            Path targetPath = uploadPath.resolve(storedFileName).normalize();
            if (targetPath.startsWith(uploadPath)) {
                Files.deleteIfExists(targetPath);
            }
        } catch (IOException e) {
            // Log-only: soft delete in DB proceeds regardless
        }
    }

    /**
     * Resolves a stored filename to its absolute Path inside the upload directory.
     * Returns {@code null} if the file does not exist or if the path would
     * escape the upload root (path traversal guard).
     * <p>
     * The input must be a bare filename, not an absolute or relative path —
     * any path separators in the input cause resolution to be rejected.
     */
    public Path load(String storedFileName) {
        if (storedFileName == null || storedFileName.isBlank()) {
            return null;
        }
        // Reject any path-like input: only bare filenames are allowed.
        if (storedFileName.contains("/") || storedFileName.contains("\\")
                || storedFileName.contains("..")) {
            return null;
        }
        Path resolved = uploadPath.resolve(storedFileName).normalize();
        if (!resolved.startsWith(uploadPath)) {
            return null;
        }
        if (!Files.exists(resolved)) {
            return null;
        }
        return resolved;
    }

    private String extractExtension(String filename) {
        int dotIndex = filename.lastIndexOf('.');
        if (dotIndex < 0 || dotIndex == filename.length() - 1) {
            throw new InvalidFileTypeException("File has no valid extension. Accepted types: .xlsx, .xls, .csv");
        }
        return filename.substring(dotIndex + 1).toLowerCase();
    }

    public record StoredFileMetadata(
            String storedFileName,
            String storagePath,
            String fileType,
            long fileSize
    ) {}
}
