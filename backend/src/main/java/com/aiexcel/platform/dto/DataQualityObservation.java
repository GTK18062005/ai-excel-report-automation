package com.aiexcel.platform.dto;

import lombok.Builder;
import lombok.Data;

/**
 * A data quality observation identified by the AI.
 */
@Data
@Builder
public class DataQualityObservation {

    /** What was observed */
    private String observation;

    /** The sheet or column affected */
    private String affectedEntity;

    /** Severity level: INFO, WARNING, ERROR */
    private String severity;
}
