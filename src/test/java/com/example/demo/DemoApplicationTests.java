package com.example.demo;

import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.IOException;

@SpringBootTest
class DemoApplicationTests {

    @Autowired
    DemoApplication demoApplication;

    @Test
    void contextLoads() throws IOException {
        demoApplication.testDemo();
    }

}
