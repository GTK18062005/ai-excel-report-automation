package com.aiexcel.platform.controller;

import com.aiexcel.platform.dto.RegisterRequest;
import com.aiexcel.platform.repository.UploadedFileRepository;
import com.aiexcel.platform.repository.UserRepository;
import com.aiexcel.platform.service.AuthService;
import com.aiexcel.platform.service.FileStorageService;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.AutoConfigureMockMvc;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.test.context.ActiveProfiles;
import org.springframework.test.web.servlet.MockMvc;
import org.springframework.transaction.annotation.Transactional;

import java.io.ByteArrayOutputStream;

import static org.assertj.core.api.Assertions.assertThat;
import static org.hamcrest.Matchers.*;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.*;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.*;

/**
 * Integration tests for the GET /api/files/{id}/output endpoint.
 * Verifies authentication, ownership checks, and successful download behaviour.
 */
@SpringBootTest
@AutoConfigureMockMvc
@ActiveProfiles("test")
@Transactional
class FileControllerOutputTest {

    @Autowired
    private MockMvc mockMvc;

    @Autowired
    private AuthService authService;

    @Autowired
    private UserRepository userRepository;

    @Autowired
    private UploadedFileRepository fileRepository;

    @Autowired
    private FileStorageService fileStorageService;

    private String user1Token;
    private String user2Token;
    private Long user1FileId;

    @BeforeEach
    void setUp() throws Exception {
        userRepository.deleteAll();

        // Register and login user1
        RegisterRequest user1Register = new RegisterRequest();
        user1Register.setName("Output Owner");
        user1Register.setEmail("owner@test.com");
        user1Register.setPassword("Password123!");
        authService.register(user1Register);
        var user1Login = new com.aiexcel.platform.dto.LoginRequest();
        user1Login.setEmail("owner@test.com");
        user1Login.setPassword("Password123!");
        user1Token = authService.login(user1Login).getToken();

        // Register and login user2
        RegisterRequest user2Register = new RegisterRequest();
        user2Register.setName("Other User");
        user2Register.setEmail("other@test.com");
        user2Register.setPassword("Password123!");
        authService.register(user2Register);
        var user2Login = new com.aiexcel.platform.dto.LoginRequest();
        user2Login.setEmail("other@test.com");
        user2Login.setPassword("Password123!");
        user2Token = authService.login(user2Login).getToken();

        // Upload a file for user1 and execute a workflow to produce output
        user1FileId = uploadAndExecuteWorkflow(user1Token, "sales_sample.xlsx");
    }

    // -------------------------------------------------------------------------
    // Happy-path tests
    // -------------------------------------------------------------------------

    @Test
    void downloadOutput_authenticatedOwner_canDownload() throws Exception {
        var result = mockMvc.perform(get("/api/files/{id}/output", user1FileId)
                        .header("Authorization", "Bearer " + user1Token))
                .andExpect(status().isOk())
                .andExpect(header().string("Content-Disposition",
                        containsString("attachment")))
                .andExpect(header().string("Content-Disposition",
                        containsString("_output")))
                .andExpect(content().contentType(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .andReturn();
        assertThat(result.getResponse().getContentAsByteArray().length).isGreaterThan(0);
    }

    @Test
    void downloadOutput_contentTypeIsXlsx() throws Exception {
        mockMvc.perform(get("/api/files/{id}/output", user1FileId)
                        .header("Authorization", "Bearer " + user1Token))
                .andExpect(status().isOk())
                .andExpect(content().contentType(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
    }

    // -------------------------------------------------------------------------
    // Auth / ownership tests
    // -------------------------------------------------------------------------

    @Test
    void downloadOutput_unauthenticated_returns403() throws Exception {
        // Spring Security returns 403 Forbidden when no credentials are provided
        mockMvc.perform(get("/api/files/{id}/output", user1FileId))
                .andExpect(status().isForbidden());
    }

    @Test
    void downloadOutput_wrongUser_returns404() throws Exception {
        // findOwnedFile returns empty → FileNotFoundException → 404
        mockMvc.perform(get("/api/files/{id}/output", user1FileId)
                        .header("Authorization", "Bearer " + user2Token))
                .andExpect(status().isNotFound());
    }

    @Test
    void downloadOutput_nonexistentId_returns404() throws Exception {
        mockMvc.perform(get("/api/files/{id}/output", 99999L)
                        .header("Authorization", "Bearer " + user1Token))
                .andExpect(status().isNotFound());
    }

    // -------------------------------------------------------------------------
    // No-output file exists (workflow never executed)
    // -------------------------------------------------------------------------

    @Test
    void downloadOutput_noOutputFile_returns404() throws Exception {
        // Upload a fresh file for user1 — no workflow executed yet
        Long freshFileId = uploadTestFile(user1Token, "unexecuted.xlsx");

        var result = mockMvc.perform(get("/api/files/{id}/output", freshFileId)
                        .header("Authorization", "Bearer " + user1Token))
                .andReturn();

        assertThat(result.getResponse().getStatus()).isEqualTo(404);
        assertThat(result.getResponse().getContentAsString())
                .containsIgnoringCase("no workflow output");
    }

    // -------------------------------------------------------------------------
    // Path-traversal / ID-confusion tests
    // -------------------------------------------------------------------------

    @Test
    void downloadOutput_pathTraversal_returns404() throws Exception {
        // A file ID that doesn't belong to user1 (and therefore can't be found
        // by findOwnedFile) returns 404 — there is no path a client can
        // inject because the path is computed server-side from DB metadata.
        // Passing user2's token with user1's fileId is the practical equivalent
        // of a path-traversal / ID-confusion attempt.
        mockMvc.perform(get("/api/files/{id}/output", user1FileId)
                        .header("Authorization", "Bearer " + user2Token))
                .andExpect(status().isNotFound());
    }

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    /**
     * Creates a simple one-step workflow, approves it, executes it, and returns
     * the file ID. This produces the output file on disk that the download
     * endpoint will serve.
     */
    private Long uploadAndExecuteWorkflow(String token, String fileName) throws Exception {
        // 1. Upload file
        Long fileId = uploadTestFile(token, fileName);

        // 2. Propose, validate, and approve a simple READ_SHEET workflow
        String validSteps = """
                {
                    "instruction": "Read the sheet",
                    "workflowName": "Read Sheet",
                    "description": "Load and read the sheet",
                    "steps": [
                        {
                            "order": 1,
                            "operation": "READ_SHEET",
                            "description": "Load TestSheet",
                            "inputSheets": ["TestSheet"],
                            "inputColumns": [],
                            "parameters": {}
                        }
                    ]
                }
                """;

        mockMvc.perform(post("/api/files/{id}/workflow-proposal/approve", fileId)
                        .header("Authorization", "Bearer " + token)
                        .contentType("application/json")
                        .content(validSteps))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.status").value("APPROVED"));

        // 3. Execute the workflow — this writes the output file to disk
        mockMvc.perform(post("/api/files/{id}/workflow/execute", fileId)
                        .header("Authorization", "Bearer " + token))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.success").value(true));

        return fileId;
    }

    private Long uploadTestFile(String token, String fileName) throws Exception {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("TestSheet");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Id");
            header.createCell(1).setCellValue("Name");
            Row data = sheet.createRow(1);
            data.createCell(0).setCellValue(1);
            data.createCell(1).setCellValue("Test Item");
            wb.write(baos);
        }

        MockMultipartFile file = new MockMultipartFile(
                "file",
                fileName,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                baos.toByteArray()
        );

        String response = mockMvc.perform(multipart("/api/files")
                        .file(file)
                        .header("Authorization", "Bearer " + token))
                .andExpect(status().isCreated())
                .andReturn()
                .getResponse()
                .getContentAsString();

        return com.fasterxml.jackson.databind.json.JsonMapper.builder()
                .build()
                .readTree(response)
                .get("id")
                .asLong();
    }
}
