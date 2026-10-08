package com.aiexcel.platform.entity;

import com.aiexcel.platform.dto.WorkflowStep;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import jakarta.persistence.*;
import lombok.*;
import org.springframework.data.annotation.CreatedDate;
import org.springframework.data.annotation.LastModifiedDate;
import org.springframework.data.jpa.domain.support.AuditingEntityListener;

import java.time.LocalDateTime;
import java.util.List;

/**
 * Tracks the approval state of a workflow proposal for a specific file + user.
 * <p>
 * Only one active proposal per file per user is stored at a time. The state machine is:
 * PROPOSED → APPROVED (valid only) | REJECTED; APPROVED → REJECTED; REJECTED → terminal.
 * <p>
 * The {@code steps} list is persisted as JSON in {@link #stepsJson} via the
 * helper methods {@link #setStepsFromJson(List)} and {@link #getStepsFromJson()}.
 */
@Entity
@Table(name = "workflow_approvals",
        uniqueConstraints = @UniqueConstraint(columnNames = {"file_id", "user_id"}))
@EntityListeners(AuditingEntityListener.class)
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class WorkflowApproval {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @ManyToOne(fetch = FetchType.LAZY)
    @JoinColumn(name = "file_id", nullable = false)
    private UploadedFile file;

    @ManyToOne(fetch = FetchType.LAZY)
    @JoinColumn(name = "user_id", nullable = false)
    private User user;

    @Enumerated(EnumType.STRING)
    @Column(nullable = false)
    private WorkflowStatus status;

    @Column(length = 500)
    private String workflowName;

    @Column(length = 1000)
    private String description;

    /**
     * Serialised JSON list of {@link WorkflowStep}. Persisted as TEXT in the database.
     */
    @Column(columnDefinition = "TEXT")
    private String stepsJson;

    @CreatedDate
    @Column(nullable = false, updatable = false)
    private LocalDateTime createdAt;

    @LastModifiedDate
    @Column(nullable = false)
    private LocalDateTime updatedAt;

    // -------------------------------------------------------------------------
    // JSON helpers
    // -------------------------------------------------------------------------

    private static final ObjectMapper MAPPER = new ObjectMapper();

    /**
     * Deserialises {@link #stepsJson} into a list of steps.
     * Returns an empty list if the JSON is null or empty.
     */
    public List<WorkflowStep> getStepsFromJson() {
        if (stepsJson == null || stepsJson.isBlank()) return List.of();
        try {
            return MAPPER.readValue(stepsJson, new TypeReference<List<WorkflowStep>>() {});
        } catch (JsonProcessingException e) {
            return List.of();
        }
    }

    /**
     * Serialises the given list of steps into {@link #stepsJson}.
     */
    public void setStepsFromJson(List<WorkflowStep> stepList) {
        if (stepList == null) {
            this.stepsJson = null;
            return;
        }
        try {
            this.stepsJson = MAPPER.writeValueAsString(stepList);
        } catch (JsonProcessingException e) {
            this.stepsJson = "[]";
        }
    }
}
