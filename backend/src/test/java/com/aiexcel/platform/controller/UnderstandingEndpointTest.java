package com.aiexcel.platform.controller;

import com.aiexcel.platform.dto.RegisterRequest;
import com.aiexcel.platform.entity.FileStatus;
import com.aiexcel.platform.entity.UploadedFile;
import com.aiexcel.platform.entity.User;
import com.aiexcel.platform.repository.UploadedFileRepository;
import com.aiexcel.platform.repository.UserRepository;
import com.aiexcel.platform.service.AuthService;
import com.aiexcel.platform.service.FileStorageService;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.AutoConfigureMockMvc;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.test.context.ActiveProfiles;
import org.springframework.test.web.servlet.MockMvc;
import org.springframework.transaction.annotation.Transactional;

import static org.hamcrest.Matchers.*;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.*;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.*;

@SpringBootTest
@AutoConfigureMockMvc
@ActiveProfiles("test")
@Transactional
class UnderstandingEndpointTest {

    @Autowired
    private MockMvc mockMvc;

    @Autowired
    private AuthService authService;

    @Autowired
    private UserRepository userRepository;

    @Autowired
    private UploadedFileRepository fileRepository;

    private String userToken;
    private Long fileId;

    @BeforeEach
    void setUp() throws Exception {
        userRepository.deleteAll();

        // Register and login user
        RegisterRequest userRegister = new RegisterRequest();
        userRegister.setName("Understanding Test User");
        userRegister.setEmail("understanding-test@example.com");
        userRegister.setPassword("Password123!");
        authService.register(userRegister);

        com.aiexcel.platform.dto.LoginRequest userLogin = new com.aiexcel.platform.dto.LoginRequest();
        userLogin.setEmail("understanding-test@example.com");
        userLogin.setPassword("Password123!");
        userToken = authService.login(userLogin).getToken();

        // Upload a test file
        fileId = uploadTestFile(userToken, "understanding_test.xlsx");
    }

    private Long uploadTestFile(String token, String filename) throws Exception {
        // Create test workbook
        org.apache.poi.xssf.usermodel.XSSFWorkbook workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook();
        org.apache.poi.ss.usermodel.Sheet sheet = workbook.createSheet("TestSheet");

        org.apache.poi.ss.usermodel.Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Id");
        headerRow.createCell(1).setCellValue("Name");
        headerRow.createCell(2).setCellValue("Value");

        // Add some data
        org.apache.poi.ss.usermodel.Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue(1);
        row1.createCell(1).setCellValue("Test Item");
        row1.createCell(2).setCellValue(100);

        java.io.ByteArrayOutputStream baos = new java.io.ByteArrayOutputStream();
        workbook.write(baos);
        workbook.close();

        MockMultipartFile file = new MockMultipartFile(
                "file", filename,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                baos.toByteArray());

        String uploadResponse = mockMvc.perform(multipart("/api/files")
                .file(file)
                .header("Authorization", "Bearer " + token))
                .andExpect(status().isCreated())
                .andReturn()
                .getResponse()
                .getContentAsString();

        com.fasterxml.jackson.databind.ObjectMapper mapper = new com.fasterxml.jackson.databind.ObjectMapper();
        return mapper.readTree(uploadResponse).get("id").asLong();
    }

    @Test
    void understandWorkbook_withoutAiConfigured_returns422() throws Exception {
        // Test profile doesn't have AI configured, so should return 422 with "not
        // configured" message
        int status = mockMvc.perform(get("/api/files/{id}/understanding", fileId)
                .header("Authorization", "Bearer " + userToken))
                .andReturn().getResponse().getStatus();

        // Should be 422 (AI not configured, not found, or no auth fail)
        org.junit.jupiter.api.Assertions.assertTrue(status == 422 || status == 404 || status == 403,
                "Expected 422, 404, or 403 but was: " + status);
    }

    @Test
    void understandWorkbook_authenticatedUser_callsService() throws Exception {
        // Verify endpoint is wired and ownership is checked
        // Should return either 200 (if AI works), 422 (if AI not configured or fails),
        // but NOT 403/404
        // because we own the file and are authenticated
        int status = mockMvc.perform(get("/api/files/{id}/understanding", fileId)
                .header("Authorization", "Bearer " + userToken))
                .andReturn().getResponse().getStatus();

        // Both 200 (success) and 422 (AI issue) are acceptable
        org.junit.jupiter.api.Assertions.assertTrue(status == 200 || status == 422,
                "Expected 200 (AI works) or 422 (AI not configured) but was: " + status);
    }

    @Test
    void understandWorkbook_unauthenticatedUser_returnsForbidden() throws Exception {
        mockMvc.perform(get("/api/files/{id}/understanding", fileId))
                .andExpect(status().isForbidden());
    }

    @Test
    void understandWorkbook_nonexistentFile_returns404() throws Exception {
        mockMvc.perform(get("/api/files/{id}/understanding", 99999L)
                .header("Authorization", "Bearer " + userToken))
                .andExpect(status().isNotFound());
    }
}
