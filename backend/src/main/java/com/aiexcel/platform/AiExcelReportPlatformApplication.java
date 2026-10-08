package com.aiexcel.platform;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.data.jpa.repository.config.EnableJpaAuditing;

@SpringBootApplication
@EnableJpaAuditing
public class AiExcelReportPlatformApplication {

    public static void main(String[] args) {
        SpringApplication.run(AiExcelReportPlatformApplication.class, args);
    }
}
