package org.spring.microservices.pptmcpserver.controller;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.spring.microservices.pptmcpserver.state.PresentationStateManager;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.ByteArrayOutputStream;
import java.io.File;

/**
 * REST Download Controller for PowerPoint presentations.
 * Allows remote clients (VS Code, Claude Desktop, web browsers, curl) to download
 * generated .pptx files directly when the MCP server is deployed remotely as a WAR or container.
 */
@Slf4j
@RestController
@RequestMapping("/api/presentations")
@RequiredArgsConstructor
public class PresentationDownloadController {

    private final PresentationStateManager stateManager;

    @GetMapping("/{id}/download")
    public ResponseEntity<byte[]> downloadPresentation(@PathVariable String id) {
        log.info("HTTP Download requested for presentation ID: '{}'", id);
        try {
            XMLSlideShow ppt = stateManager.getPresentation(id);
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            ppt.write(baos);
            byte[] bytes = baos.toByteArray();

            String filePath = stateManager.getFilePath(id);
            String fileName = (filePath != null) ? new File(filePath).getName() : "presentation.pptx";
            if (!fileName.endsWith(".pptx")) {
                fileName += ".pptx";
            }

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + fileName + "\"")
                    .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.presentationml.presentation"))
                    .body(bytes);
        } catch (Exception e) {
            log.error("Failed to stream download for presentation ID '{}': {}", id, e.getMessage());
            return ResponseEntity.notFound().build();
        }
    }

    @GetMapping("/download")
    public ResponseEntity<byte[]> downloadActivePresentation() {
        String activeId = stateManager.getCurrentPresentationId();
        if (activeId == null) {
            return ResponseEntity.badRequest().body("No active presentation loaded in memory.".getBytes());
        }
        return downloadPresentation(activeId);
    }
}
