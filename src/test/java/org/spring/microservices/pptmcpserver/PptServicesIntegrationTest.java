package org.spring.microservices.pptmcpserver;

import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.spring.microservices.pptmcpserver.constants.Literals;
import org.spring.microservices.pptmcpserver.service.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * End-to-end integration test validating presentation lifecycle, content insertion,
 * structural elements, design styling, compound executive templates, and image rendering.
 */
@Slf4j
@SpringBootTest
class PptServicesIntegrationTest {

    

    @Autowired
    private PresentationService presentationService;

    @Autowired
    private SlideContentService slideContentService;

    @Autowired
    private StructuralElementService structuralElementService;

    @Autowired
    private DesignService designService;

    @Autowired
    private LayoutTemplateService layoutTemplateService;

    @Autowired
    private RenderService renderService;

    @Autowired
    private GuardrailService guardrailService;

    @Test
    void testEndToEndPptWorkflow(@TempDir Path tempDir) {
        log.info("--- Starting End-to-End PPT Workflow Integration Test ---");

        // 1. Diagnostics
        Map<String, Object> serverInfo = renderService.getServerInfo();
        assertEquals("PPTMCPServer", serverInfo.get("serverName"));
        assertEquals(43, serverInfo.get("totalTools"));
        log.info("Server diagnostics verified: 43 tools registered");

        // 2. Create Presentation
        Map<String, Object> createRes = presentationService.createPresentation("AI & Cloud Architecture 2026", 960.0, 540.0);
        assertEquals(Literals.SUCCESS, createRes.get(Literals.STATUS));
        String presId = (String) createRes.get(Literals.PRESENTATION_ID);
        assertNotNull(presId);
        log.info("Created test presentation ID: {}", presId);

        // 3. Add Slide 1 (Title slide)
        Map<String, Object> slide1 = slideContentService.addSlide(presId, null, null, "AI Innovations 2026", "Enterprise Transformation Strategy");
        assertEquals(Literals.SUCCESS, slide1.get(Literals.STATUS));
        assertEquals(0, slide1.get("slideIndex"));

        // 4. Add Formatted Textbox on Slide 1
        Map<String, Object> tbRes = slideContentService.addTextbox(
                presId, 0, "Confidential - For Internal Use Only",
                50.0, 460.0, 400.0, 30.0, 12.0, "#64748B", "Arial", false, true, false, "LEFT"
        );
        assertEquals(Literals.SUCCESS, tbRes.get(Literals.STATUS));

        // 5. Add Slide 2 (Content slide with Table & Shape)
        Map<String, Object> slide2 = slideContentService.addSlide(presId, null, null, "Performance Metrics", null);
        assertEquals(Literals.SUCCESS, slide2.get(Literals.STATUS));
        assertEquals(1, slide2.get("slideIndex"));

        // Set Slide 2 Solid Background
        Map<String, Object> bgRes = designService.setSlideBackground(presId, 1, "#F8FAFC", null, null, null);
        assertEquals(Literals.SUCCESS, bgRes.get(Literals.STATUS));

        // Add Table to Slide 2
        Map<String, Object> tableRes = structuralElementService.addTable(
                presId, 1, 3, 3,
                50.0, 120.0, 860.0, 200.0,
                List.of(
                        List.of("Throughput", "15k req/s", "+35%"),
                        List.of("Latency (p99)", "18 ms", "-42%")
                ),
                List.of("Metric", "Value", "Delta YoY"),
                "#1E3A8A", "#FFFFFF", "#F1F5F9", "#FFFFFF"
        );
        assertEquals(Literals.SUCCESS, tableRes.get(Literals.STATUS));

        // Format a Table Cell
        Map<String, Object> cellRes = structuralElementService.formatTableCell(
                presId, 1, null, 1, 2, "+35% (Record)", "#DCFCE7", "#166534", 12.0, "Arial", true, false, "CENTER"
        );
        assertEquals(Literals.SUCCESS, cellRes.get(Literals.STATUS));

        // Add AutoShape (Star)
        Map<String, Object> shapeRes = structuralElementService.addShape(
                presId, 1, "star",
                850.0, 40.0, 50.0, 50.0,
                "#F59E0B", "#D97706", 1.5, null, null, null, null
        );
        assertEquals(Literals.SUCCESS, shapeRes.get(Literals.STATUS));

        // Add Connector
        Map<String, Object> connRes = structuralElementService.addConnector(
                presId, 1, "STRAIGHT", 50.0, 360.0, 910.0, 360.0, "#94A3B8", 1.0
        );
        assertEquals(Literals.SUCCESS, connRes.get(Literals.STATUS));

        // 6. Add Slide 3 (Content Slide with Rounded Rectangle Card)
        Map<String, Object> slide3 = slideContentService.addSlide(presId, null, null, "Quarterly Adoption", "Enterprise adoption metrics across regions");
        assertEquals(Literals.SUCCESS, slide3.get(Literals.STATUS));

        Map<String, Object> shapeRes3 = structuralElementService.addShape(
                presId, 2, "round_rect",
                50.0, 150.0, 860.0, 250.0,
                "#EFF6FF", "#BFDBFE", 1.5,
                "Adoption Highlights:\n• 78k Enterprise Users Active\n• +140% YoY Throughput Scaling\n• 99.99% Availability SLA Maintained",
                "#1E3A8A", 18.0, false
        );
        assertEquals(Literals.SUCCESS, shapeRes3.get(Literals.STATUS));

        // 7. Verify Text Extraction
        Map<String, Object> textRes = slideContentService.extractSlideText(presId, 1);
        assertEquals(Literals.SUCCESS, textRes.get(Literals.STATUS));
        assertTrue(((String) textRes.get("fullText")).contains("Throughput"));

        // 8. Render Slide 1 to PNG image
        File renderImg = tempDir.resolve("slide_0.png").toFile();
        Map<String, Object> renderRes = renderService.renderSlideToImage(presId, 0, renderImg.getAbsolutePath(), 1.0, "PNG");
        assertEquals(Literals.SUCCESS, renderRes.get(Literals.STATUS));
        assertTrue(renderImg.exists());
        assertTrue(renderImg.length() > 0);

        // 9. Add Split-Screen Executive Slide (Text Left + Image Right)
        Map<String, Object> splitRes = layoutTemplateService.addSplitContentSlide(
                presId, "System Architecture Deep-Dive", "Enterprise Overview",
                List.of(
                        "Microservices decoupled via Apache Kafka message brokers",
                        "Spring AI MCP endpoints handle subagent LLM invocations",
                        "Zero-latency in-memory multi-presentation caching"
                ),
                renderImg.getAbsolutePath(), "RIGHT", "Figure 1: Generated Slide Preview",
                "#1E3A8A", "#FFFFFF"
        );
        assertEquals(Literals.SUCCESS, splitRes.get(Literals.STATUS));

        // 10. Add Metric KPI Dashboard Slide
        Map<String, Object> metricRes = layoutTemplateService.addMetricCardsSlide(
                presId, "Executive Financial & Technical KPIs", "Q3 2026 Results",
                List.of(
                        Map.of("label", "ARR", "value", "$48.5M", "trend", "+32% YoY", "trendPositive", true, "detail", "Exceeded FY26 target"),
                        Map.of("label", "p99 Latency", "value", "12 ms", "trend", "-45%", "trendPositive", true, "detail", "Optimized with POI 5.5"),
                        Map.of("label", "Active Subagents", "value", "1,250", "trend", "+110%", "trendPositive", true, "detail", "Enterprise scale")
                ),
                "#1E3A8A", "#F8FAFC"
        );
        assertEquals(Literals.SUCCESS, metricRes.get(Literals.STATUS));

        // 11. Add Comparison Slide (3 Columns)
        Map<String, Object> compRes = layoutTemplateService.addComparisonSlide(
                presId, "Framework Comparison: Java MCP vs Python MCP", "Architecture evaluation",
                List.of(
                        Map.of("header", "Java PPT MCP", "badge", "Selected", "points", List.of("Native Graphics2D rendering", "Strong type safety", "In-memory state manager", "30+ production tools"), "color", "#1E3A8A"),
                        Map.of("header", "Python PPT MCP", "badge", "Legacy", "points", List.of("No native rendering", "Requires LibreOffice", "Single-threaded GIL"), "color", "#64748B")
                )
        );
        assertEquals(Literals.SUCCESS, compRes.get(Literals.STATUS));

        // 12. Add Process Flow Slide
        Map<String, Object> flowRes = layoutTemplateService.addProcessFlowSlide(
                presId, "Presentation Generation Pipeline", "End-to-end execution",
                List.of(
                        Map.of("step", "1", "title", "Schema Plan", "description", "LLM designs layout JSON", "duration", "0.5s"),
                        Map.of("step", "2", "title", "POI Build", "description", "Java MCP generates shapes & charts", "duration", "0.2s"),
                        Map.of("step", "3", "title", "Render & Save", "description", "Graphics2D exports PNG preview", "duration", "0.3s")
                ),
                "#1E3A8A"
        );
        assertEquals(Literals.SUCCESS, flowRes.get(Literals.STATUS));

        // 13. Add Speaker Notes
        Map<String, Object> notesRes = slideContentService.setSpeakerNotes(presId, 0, "Welcome stakeholders and introduce agenda.");
        assertEquals(Literals.SUCCESS, notesRes.get(Literals.STATUS));

        // 14. Add Corporate Header/Footer across slides
        Map<String, Object> hfRes = designService.addHeaderFooter(presId, "ENTERPRISE CONFIDENTIAL", "© 2026 Enterprise PPT MCP Server", true, 0);
        assertEquals(Literals.SUCCESS, hfRes.get(Literals.STATUS));

        // 15. Duplicate Slide
        Map<String, Object> dupRes = layoutTemplateService.duplicateSlide(presId, 0);
        assertEquals(Literals.SUCCESS, dupRes.get(Literals.STATUS));

        // 15b. Guardrail Presentation Quality Audit & Rule Validation
        Map<String, Object> auditRes = guardrailService.auditPresentationQuality(presId);
        assertEquals(Literals.SUCCESS, auditRes.get(Literals.STATUS));
        assertTrue((int) auditRes.get(Literals.QUALITY_SCORE) > 0);
        assertNotNull(auditRes.get(Literals.QUALITY_GRADE));

        Map<String, Object> rulesRes = guardrailService.validatePresentationRules(
                presId, 5, 20, null, List.of("AI Innovations"), null, 8.0, 15, 300, null, "16:9");
        assertEquals(Literals.SUCCESS, rulesRes.get(Literals.STATUS));
        assertTrue((Boolean) rulesRes.get(Literals.PASSED));

        // 16. Save Presentation to PPTX file
        File pptxFile = tempDir.resolve("final_output.pptx").toFile();
        Map<String, Object> saveRes = presentationService.savePresentation(presId, pptxFile.getAbsolutePath(), null);
        assertEquals(Literals.SUCCESS, saveRes.get(Literals.STATUS));
        assertTrue(pptxFile.exists());
        assertTrue(pptxFile.length() > 0);

        // 17. Batch Render all slides to images
        File batchDir = tempDir.resolve("rendered_batch").toFile();
        Map<String, Object> batchRenderRes = renderService.renderAllSlidesToImages(presId, batchDir.getAbsolutePath(), 1.0, "PNG");
        assertEquals(Literals.SUCCESS, batchRenderRes.get(Literals.STATUS));
        assertTrue((Integer) batchRenderRes.get("slideCount") >= 7);

        // 18. Create new presentation from saved template
        Map<String, Object> tplRes = presentationService.createPresentationFromTemplate(pptxFile.getAbsolutePath(), "Cloned from Template", false);
        assertEquals(Literals.SUCCESS, tplRes.get(Literals.STATUS));
        assertEquals(0, (Integer) tplRes.get("slideCount"));
        assertFalse(((List<?>) tplRes.get("availableLayouts")).isEmpty());

        // 19. Verify List Presentations
        Map<String, Object> listRes = presentationService.listPresentations();
        assertEquals(Literals.SUCCESS, listRes.get(Literals.STATUS));
        assertTrue((Integer) listRes.get("count") >= 2);

        log.info("--- End-to-End PPT Workflow Integration Test COMPLETED SUCCESSFULLY ---");
    }
}
