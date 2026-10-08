package com.aiexcel.platform.service;

import com.aiexcel.platform.dto.LoginRequest;
import com.aiexcel.platform.dto.RegisterRequest;
import com.aiexcel.platform.exception.InvalidCredentialsException;
import com.aiexcel.platform.repository.UserRepository;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.ActiveProfiles;
import org.springframework.transaction.annotation.Transactional;

import static org.junit.jupiter.api.Assertions.*;

@SpringBootTest
@ActiveProfiles("test")
@Transactional
class AuthServiceTest {

    @Autowired
    private AuthService authService;

    @Autowired
    private UserRepository userRepository;

    @BeforeEach
    void setUp() {
        userRepository.deleteAll();
    }

    @Test
    void registerAndLoginWithSamePassword_shouldSucceed() {
        String email = "authtest@example.com";
        String password = "TestPassword123!";

        RegisterRequest registerRequest = new RegisterRequest();
        registerRequest.setName("Auth Test User");
        registerRequest.setEmail(email);
        registerRequest.setPassword(password);
        authService.register(registerRequest);

        LoginRequest loginRequest = new LoginRequest();
        loginRequest.setEmail(email);
        loginRequest.setPassword(password);
        var response = authService.login(loginRequest);

        assertNotNull(response.getToken());
        assertNotNull(response.getUserId());
        assertEquals(email, response.getEmail());
    }

    @Test
    void loginWithWrongPassword_shouldFail() {
        String email = "wrongpass@example.com";
        String correctPassword = "CorrectPassword123!";
        String wrongPassword = "WrongPassword123!";

        RegisterRequest registerRequest = new RegisterRequest();
        registerRequest.setName("Wrong Pass User");
        registerRequest.setEmail(email);
        registerRequest.setPassword(correctPassword);
        authService.register(registerRequest);

        LoginRequest loginRequest = new LoginRequest();
        loginRequest.setEmail(email);
        loginRequest.setPassword(wrongPassword);

        assertThrows(InvalidCredentialsException.class, () -> {
            authService.login(loginRequest);
        });
    }
}
