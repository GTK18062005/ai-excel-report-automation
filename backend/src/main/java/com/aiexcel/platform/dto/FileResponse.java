package com.aiexcel.platform.dto;

import com.aiexcel.platform.entity.FileStatus;
import lombok.Builder;
import lombok.Data;

import java.time.LocalDateTime;

@Data
@Builder
public class FileResponse {

    private Long id;
    private String originalFileName;
    private String fileType;
    private Long fileSize;
    private FileStatus status;
    private LocalDateTime uploadedAt;
}
