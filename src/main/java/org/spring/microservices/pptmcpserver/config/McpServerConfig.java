package org.spring.microservices.pptmcpserver.config;

import io.modelcontextprotocol.server.McpSyncServerExchange;
import io.modelcontextprotocol.spec.McpSchema;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.spring.microservices.pptmcpserver.state.PresentationStateManager;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import java.util.List;
import java.util.function.BiConsumer;

/**
 * Spring AI MCP Server configuration.
 * Registers roots change handler to capture and track active client workspace roots from VS Code/Copilot.
 */
@Slf4j
@Configuration
@RequiredArgsConstructor
public class McpServerConfig {

    private final PresentationStateManager stateManager;

    @Bean
    public BiConsumer<McpSyncServerExchange, List<McpSchema.Root>> syncRootsChangeConsumer() {
        return (exchange, roots) -> {
            log.info("Received sync MCP roots notification with {} root(s)", roots != null ? roots.size() : 0);
            stateManager.updateWorkspaceRoots(roots);
        };
    }
}
