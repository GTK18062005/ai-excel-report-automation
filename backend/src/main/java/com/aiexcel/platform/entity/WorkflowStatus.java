package com.aiexcel.platform.entity;

/**
 * Lifecycle state of a proposed workflow.
 * <p>
 * Transitions:
 * <ul>
 *   <li>{@code PROPOSED} → {@code APPROVED} (only when validation succeeds)</li>
 *   <li>{@code PROPOSED} → {@code REJECTED}</li>
 *   <li>{@code APPROVED} → {@code REJECTED} (undo approval)</li>
 *   <li>{@code APPROVED} → {@code SUPERSEDED} (replaced by a new proposal)</li>
 *   <li>{@code REJECTED} → (terminal; cannot be approved again)</li>
 *   <li>{@code SUPERSEDED} → (terminal; superseded by a newer proposal)</li>
 * </ul>
 */
public enum WorkflowStatus {
    PROPOSED,
    APPROVED,
    REJECTED,
    SUPERSEDED
}
