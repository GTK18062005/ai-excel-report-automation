package com.aiexcel.platform.repository;

import com.aiexcel.platform.entity.UploadedFile;
import com.aiexcel.platform.entity.User;
import com.aiexcel.platform.entity.WorkflowApproval;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.Optional;

@Repository
public interface WorkflowApprovalRepository extends JpaRepository<WorkflowApproval, Long> {

    /**
     * Returns the single approval record for the given file + user pair, if any.
     * The {@code (file_id, user_id)} unique constraint guarantees at most one row.
     */
    Optional<WorkflowApproval> findByFileAndUser(UploadedFile file, User user);
}
