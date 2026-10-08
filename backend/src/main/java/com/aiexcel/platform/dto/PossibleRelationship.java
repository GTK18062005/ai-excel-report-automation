package com.aiexcel.platform.dto;

import lombok.Builder;
import lombok.Data;

/**
 * A possible relationship between sheets or columns identified by the AI.
 */
@Data
@Builder
public class PossibleRelationship {

    /** Description of the relationship */
    private String description;

    /** First entity involved */
    private String entity1;

    /** Second entity involved */
    private String entity2;

    /** Type of relationship (e.g., one-to-many, join-key, hierarchy) */
    private String relationshipType;
}
