package com.aiexcel.platform.repository;

import com.aiexcel.platform.entity.FileStatus;
import com.aiexcel.platform.entity.UploadedFile;
import com.aiexcel.platform.entity.User;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.Optional;

@Repository
public interface UploadedFileRepository extends JpaRepository<UploadedFile, Long> {

    List<UploadedFile> findByUserAndStatusNotOrderByUploadedAtDesc(User user, FileStatus status);

    Optional<UploadedFile> findByIdAndUserAndStatusNot(Long id, User user, FileStatus status);
}
