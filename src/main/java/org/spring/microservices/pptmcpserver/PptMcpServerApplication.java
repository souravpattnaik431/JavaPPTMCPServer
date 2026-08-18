package org.spring.microservices.pptmcpserver;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.builder.SpringApplicationBuilder;
import org.springframework.boot.web.servlet.support.SpringBootServletInitializer;

/**
 * Spring Boot MCP Server Application Entrypoint.
 * Configured as a WAR package supporting both standalone execution and external servlet container deployment (Tomcat/WildFly/etc).
 */
@SpringBootApplication
public class PptMcpServerApplication extends SpringBootServletInitializer {

    @Override
    protected SpringApplicationBuilder configure(SpringApplicationBuilder application) {
        return application.sources(PptMcpServerApplication.class);
    }

    public static void main(String[] args) {
        SpringApplication.run(PptMcpServerApplication.class, args);
    }

}
