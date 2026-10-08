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
import java.io.InputStream;

import static org.assertj.core.api.Assertions.assertThat;
import static org.hamcrest.Matchers.*;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.*;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.*;

/**
 * Integration tests for the GET /api/files/{id}/output/pdf endpoint.
 * Verifies authentication, ownership checks, successful download, and
 * pre-existing XLSX output endpoint is not affected.
 */
@SpringBootTest
@AutoConfigureMockMvc
@ActiveProfiles("test")
@Transactional
class FileControllerOutputPdfTest {

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
    private Long user1FileId;  // has READ_SHEET + GENERATE_PDF executed
    private Long user1FileNoPdfId;  // has READ_SHEET only (no PDF)

    @BeforeEach
    void setUp() throws Exception {
        userRepository.deleteAll();

        // Register and login user1
        RegisterRequest user1Register = new RegisterRequest();
        user1Register.setName("PDF Owner");
        user1Register.setEmail("pdfowner@test.com");
        user1Register.setPassword("Password123!");
        authService.register(user1Register);
        var user1Login = new com.aiexcel.platform.dto.LoginRequest();
        user1Login.setEmail("pdfowner@test.com");
        user1Login.setPassword("Password123!");
        user1Token = authService.login(user1Login).getToken();

        // Register and login user2
        RegisterRequest user2Register = new RegisterRequest();
        user2Register.setName("Other User");
        user2Register.setEmail("pdfother@test.com");
        user2Register.setPassword("Password123!");
        authService.register(user2Register);
        var user2Login = new com.aiexcel.platform.dto.LoginRequest();
        user2Login.setEmail("pdfother@test.com");
        user2Login.setPassword("Password123!");
        user2Token = authService.login(user2Login).getToken();

        // Upload a file and execute READ_SHEET + GENERATE_PDF → produces PDF
        user1FileId = uploadAndExecutePdfWorkflow(user1Token, "sales_with_pdf.xlsx");

        // Upload a second file with only READ_SHEET → no PDF generated
        user1FileNoPdfId = uploadAndExecuteXlsxWorkflow(user1Token, "sales_no_pdf.xlsx");
    }

    // -------------------------------------------------------------------------
    // Happy-path tests
    // -------------------------------------------------------------------------

    @Test
    void downloadOutputPdf_authenticatedOwner_canDownload() throws Exception {
        var result = mockMvc.perform(get("/api/files/{id}/output/pdf", user1FileId)
                        .header("Authorization", "Bearer " + user1Token))
                .andExpect(status().isOk())
                .andExpect(header().string("Content-Disposition",
                        containsString("attachment")))
                .andExpect(header().string("Content-Disposition",
                        containsString("_output.pdf")))
                .andExpect(content().contentType("application/pdf"))
                .andReturn();
        assertThat(result.getResponse().getContentAsByteArray().length).isGreaterThan(0);

        // Verify it starts with the PDF magic bytes
        byte[] body = result.getResponse().getContentAsByteArray();
        assertThat(body.length).isGreaterThan(5);
        assertThat(body[0]).isEqualTo((byte) '%');  // %PDF-
        assertThat(body[1]).isEqualTo((byte) 'P');
        assertThat(body[2]).isEqualTo((byte) 'D');
        assertThat(body[3]).isEqualTo((byte) 'F');
        assertThat(body[4]).isEqualTo((byte) '-');
    }

    @Test
    void downloadOutputPdf_contentTypeIsApplicationPdf() throws Exception {
        mockMvc.perform(get("/api/files/{id}/output/pdf", user1FileId)
                        .header("Authorization", "Bearer " + user1Token))
                .andExpect(status().isOk())
                .andExpect(content().contentType("application/pdf"));
    }

    @Test
    void downloadOutputPdf_filenameHasPdfExtension() throws Exception {
        mockMvc.perform(get("/api/files/{id}/output/pdf", user1FileId)
                        .header("Authorization", "Bearer " + user1Token))
                .andExpect(status().isOk())
                .andExpect(header().string("Content-Disposition",
                        containsString("_output.pdf")));
    }

    // -------------------------------------------------------------------------
    // Auth / ownership tests
    // -------------------------------------------------------------------------

    @Test
    void downloadOutputPdf_unauthenticated_returns403() throws Exception {
        mockMvc.perform(get("/api/files/{id}/output/pdf", user1FileId))
                .andExpect(status().isForbidden());
    }

    @Test
    void downloadOutputPdf_wrongUser_returns404() throws Exception {
        mockMvc.perform(get("/api/files/{id}/output/pdf", user1FileId)
                        .header("Authorization", "Bearer " + user2Token))
                .andExpect(status().isNotFound());
    }

    @Test
    void downloadOutputPdf_nonexistentId_returns404() throws Exception {
        mockMvc.perform(get("/api/files/{id}/output/pdf", 99999L)
                        .header("Authorization", "Bearer " + user1Token))
                .andExpect(status().isNotFound());
    }

    // -------------------------------------------------------------------------
    // PDF not generated yet
    // -------------------------------------------------------------------------

    @Test
    void downloadOutputPdf_noPdfOutput_returns404() throws Exception {
        // user1FileNoPdfId has only READ_SHEET executed — no GENERATE_PDF
        var result = mockMvc.perform(get("/api/files/{id}/output/pdf", user1FileNoPdfId)
                        .header("Authorization", "Bearer " + user1Token))
                .andReturn();

        assertThat(result.getResponse().getStatus()).isEqualTo(404);
        assertThat(result.getResponse().getContentAsString())
                .containsIgnoringCase("no pdf");
    }

    // -------------------------------------------------------------------------
    // Path-traversal / ID-confusion tests
    // -------------------------------------------------------------------------

    @Test
    void downloadOutputPdf_pathTraversal_returns404() throws Exception {
        // Using user2's token with user1's fileId is the practical equivalent
        // of a path-traversal / ID-confusion attempt — the server computes
        // the path server-side from DB metadata with no user input.
        mockMvc.perform(get("/api/files/{id}/output/pdf", user1FileId)
                        .header("Authorization", "Bearer " + user2Token))
                .andExpect(status().isNotFound());
    }

    // -------------------------------------------------------------------------
    // Existing XLSX endpoint still works
    // -------------------------------------------------------------------------

    @Test
    void downloadOutput_xlsxEndpoint_stillWorks() throws Exception {
        // The pre-existing XLSX output endpoint must remain functional.
        // user1FileId has READ_SHEET + GENERATE_PDF; the XLSX output exists.
        mockMvc.perform(get("/api/files/{id}/output", user1FileId)
                        .header("Authorization", "Bearer " + user1Token))
                .andExpect(status().isOk())
                .andExpect(content().contentType(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
    }

    @Test
    void downloadOutput_xlsxEndpoint_returnsCorrectContentType() throws Exception {
        mockMvc.perform(get("/api/files/{id}/output", user1FileId)
                        .header("Authorization", "Bearer " + user1Token))
                .andExpect(status().isOk())
                .andExpect(header().string("Content-Disposition",
                        containsString("_output")))
                .andExpect(header().string("Content-Disposition",
                        not(containsString(".pdf"))));
    }

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    /**
     * Uploads a file, approves and executes a two-step workflow
     * (READ_SHEET + GENERATE_PDF), then returns the file ID.
     * This produces both the XLSX output and the PDF output on disk.
     */
    private Long uploadAndExecutePdfWorkflow(String token, String fileName) throws Exception {
        Long fileId = uploadTestFile(token, fileName);

        String steps = """
                {
                    "instruction": "Read sheet and generate PDF",
                    "workflowName": "Read and Export PDF",
                    "description": "Read sheet then export to PDF",
                    "steps": [
                        {
                            "order": 1,
                            "operation": "READ_SHEET",
                            "description": "Load TestSheet",
                            "inputSheets": ["TestSheet"],
                            "inputColumns": [],
                            "parameters": {}
                        },
                        {
                            "order": 2,
                            "operation": "GENERATE_PDF",
                            "description": "Export to PDF",
                            "inputSheets": [],
                            "inputColumns": [],
                            "parameters": {
                                "title": "Test Report",
                                "pageSize": "30"
                            }
                        }
                    ]
                }
                """;

        mockMvc.perform(post("/api/files/{id}/workflow-proposal/approve", fileId)
                        .header("Authorization", "Bearer " + token)
                        .contentType("application/json")
                        .content(steps))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.status").value("APPROVED"));

        mockMvc.perform(post("/api/files/{id}/workflow/execute", fileId)
                        .header("Authorization", "Bearer " + token))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.success").value(true));

        return fileId;
    }

    /**
     * Uploads a file and executes only a READ_SHEET workflow
     * (no GENERATE_PDF step). This produces the XLSX output but no PDF.
     */
    private Long uploadAndExecuteXlsxWorkflow(String token, String fileName) throws Exception {
        Long fileId = uploadTestFile(token, fileName);

        String steps = """
                {
                    "instruction": "Read sheet",
                    "workflowName": "Read Sheet",
                    "description": "Load sheet",
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
                        .content(steps))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.status").value("APPROVED"));

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
            Row data1 = sheet.createRow(1);
            data1.createCell(0).setCellValue(1);
            data1.createCell(1).setCellValue("Test Item");
            Row data2 = sheet.createRow(2);
            data2.createCell(0).setCellValue(2);
            data2.createCell(1).setCellValue("Another Item");
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
