package com.aiexcel.platform.config;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.AutoConfigureMockMvc;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.ActiveProfiles;
import org.springframework.test.web.servlet.MockMvc;

import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.options;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.header;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.status;

/**
 * Focused tests verifying that the CorsConfigurationSource in {@link SecurityConfig}
 * correctly grants cross-origin access to http://localhost:5173 (Vite dev server)
 * and denies it to unlisted origins.
 *
 * Uses the full Spring Boot context with the "test" profile so the real
 * SecurityFilterChain processes requests.
 */
@SpringBootTest
@AutoConfigureMockMvc
@ActiveProfiles("test")
class CorsConfigTest {

    @Autowired
    private MockMvc mockMvc;

    // -----------------------------------------------------------------------
    // Preflight — allowed origin
    // -----------------------------------------------------------------------

    @Test
    @DisplayName("OPTIONS preflight from Vite origin returns 200 with CORS headers")
    void preflight_fromAllowedOrigin_returns200WithCorsHeaders() throws Exception {
        mockMvc.perform(options("/api/auth/login")
                        .header("Origin", "http://localhost:5173")
                        .header("Access-Control-Request-Method", "POST")
                        .header("Access-Control-Request-Headers", "Authorization,Content-Type"))
                .andExpect(status().isOk())
                .andExpect(header().string("Access-Control-Allow-Origin", "http://localhost:5173"))
                .andExpect(header().exists("Access-Control-Allow-Methods"))
                .andExpect(header().exists("Access-Control-Allow-Headers"));
    }

    @Test
    @DisplayName("OPTIONS preflight from Vite origin allows GET method")
    void preflight_allowsGetMethod() throws Exception {
        mockMvc.perform(options("/api/files")
                        .header("Origin", "http://localhost:5173")
                        .header("Access-Control-Request-Method", "GET")
                        .header("Access-Control-Request-Headers", "Authorization"))
                .andExpect(status().isOk())
                .andExpect(header().string("Access-Control-Allow-Origin", "http://localhost:5173"));
    }

    @Test
    @DisplayName("OPTIONS preflight from Vite origin allows DELETE method")
    void preflight_allowsDeleteMethod() throws Exception {
        mockMvc.perform(options("/api/files/1")
                        .header("Origin", "http://localhost:5173")
                        .header("Access-Control-Request-Method", "DELETE")
                        .header("Access-Control-Request-Headers", "Authorization"))
                .andExpect(status().isOk())
                .andExpect(header().string("Access-Control-Allow-Origin", "http://localhost:5173"));
    }

    // -----------------------------------------------------------------------
    // Preflight — disallowed origin must NOT receive CORS headers
    // -----------------------------------------------------------------------

    @Test
    @DisplayName("OPTIONS preflight from unlisted origin does NOT receive Allow-Origin header")
    void preflight_fromDisallowedOrigin_noAllowOriginHeader() throws Exception {
        mockMvc.perform(options("/api/auth/login")
                        .header("Origin", "http://localhost:3000")
                        .header("Access-Control-Request-Method", "POST")
                        .header("Access-Control-Request-Headers", "Content-Type"))
                .andExpect(header().doesNotExist("Access-Control-Allow-Origin"));
    }
}
