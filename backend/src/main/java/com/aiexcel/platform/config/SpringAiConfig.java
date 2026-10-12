package com.aiexcel.platform.config;

import io.micrometer.observation.ObservationRegistry;
import org.springframework.ai.chat.model.ChatModel;
import org.springframework.ai.openai.OpenAiChatModel;
import org.springframework.ai.openai.OpenAiChatOptions;
import org.springframework.ai.openai.api.OpenAiApi;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.autoconfigure.condition.ConditionalOnExpression;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

/**
 * Manual Spring AI configuration for OpenAI / Groq.
 * Only activated when spring.ai.openai.api-key is set and not empty.
 */
@Configuration
@ConditionalOnExpression("!'${spring.ai.openai.api-key:}'.trim().isEmpty()")
public class SpringAiConfig {

    @Value("${spring.ai.openai.api-key}")
    private String apiKey;

    @Value("${spring.ai.openai.base-url:https://api.openai.com}")
    private String baseUrl;

    @Value("${spring.ai.openai.chat.options.model:gpt-4o-mini}")
    private String model;

    @Value("${spring.ai.openai.chat.options.temperature:0.2}")
    private double temperature;

    @Value("${spring.ai.openai.chat.options.max-tokens:4000}")
    private Integer maxTokens;

    @Bean
    public OpenAiApi openAiApi() {
        // OpenRouter is OpenAI-compatible. The configured base-url already includes
        // /v1,
        // so we override completionsPath to /chat/completions to avoid a duplicate /v1.
        String completionsPath = baseUrl.endsWith("/v1") ? "/chat/completions" : "/v1/chat/completions";
        return OpenAiApi.builder()
                .apiKey(apiKey)
                .baseUrl(baseUrl)
                .completionsPath(completionsPath)
                .build();
    }

    @Bean
    public ChatModel chatModel(OpenAiApi openAiApi) {
        OpenAiChatOptions options = OpenAiChatOptions.builder()
                .model(model)
                .temperature(temperature)
                .maxTokens(maxTokens)
                .build();
        return OpenAiChatModel.builder()
                .openAiApi(openAiApi)
                .defaultOptions(options)
                .observationRegistry(ObservationRegistry.NOOP)
                .build();
    }
}
