package org.spring.microservices.pptmcpserver;

import org.junit.jupiter.api.Test;
import org.spring.microservices.pptmcpserver.constants.Literals;
import org.spring.microservices.pptmcpserver.service.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.util.List;
import java.util.Map;

@SpringBootTest
public class CumulativeCheckTest {

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

    @Test
    void testCumulative() {
        File dir = new File("test_output/cum_check");
        dir.mkdirs();

        Map<String, Object> createRes = presentationService.createPresentation("Cum Check", 960.0, 540.0);
        String presId = (String) createRes.get(Literals.PRESENTATION_ID);

        // Slide 1
        slideContentService.addSlide(presId, 0, "Title Slide", "Executive Technology Strategy 2026", "Modernizing Enterprise Infrastructure with Spring AI & MCP");
        designService.setSlideBackground(presId, 0, null, "diagonal", "#0F172A", "#1E3A8A");
        slideContentService.setSpeakerNotes(presId, 0, "Welcome everyone. Today we review our Q3 architecture modernization roadmap.");
        presentationService.savePresentation(presId, new File(dir, "cum_1.pptx").getAbsolutePath(), null);

        // Slide 2
        layoutTemplateService.addMetricCardsSlide(
                presId, "Executive Key Performance Indicators", "FY26 Technical & Financial Milestones",
                List.of(
                        Map.of("label", "ARR Growth", "value", "$52.4M", "trend", "+38% YoY", "trendPositive", true, "detail", "Exceeded FY26 budget"),
                        Map.of("label", "API Latency (p99)", "value", "8.5 ms", "trend", "-62%", "trendPositive", true, "detail", "Zero JVM GC pauses"),
                        Map.of("label", "Active MCP Tools", "value", "40 Tools", "trend", "100% Validated", "trendPositive", true, "detail", "Enterprise Grade"),
                        Map.of("label", "Throughput", "value", "45k req/s", "trend", "+140%", "trendPositive", true, "detail", "Distributed caching")
                ),
                "#1E3A8A", "#F8FAFC"
        );
        presentationService.savePresentation(presId, new File(dir, "cum_2.pptx").getAbsolutePath(), null);

        // Slide 3
        File slide1Img = new File(dir, "preview.png");
        renderService.renderSlideToImage(presId, 0, slide1Img.getAbsolutePath(), 1.0, "PNG");
        layoutTemplateService.addSplitContentSlide(
                presId, "Microservices & MCP Architecture", "Platform Overview & Technical Capabilities",
                List.of(
                        "Decoupled Spring Boot microservices with Spring AI 2.0.0 integration",
                        "High-throughput in-memory PresentationStateManager with concurrent thread-safety",
                        "Native Java Graphics2D high-DPI slide rendering with zero external OS dependencies",
                        "Support for all 16 POI AutoShapes, XDDF charts, tables, and hyperlinks"
                ),
                slide1Img.getAbsolutePath(), "RIGHT", "Figure 1: Generated Presentation Title Preview",
                "#1E3A8A", "#FFFFFF"
        );
        presentationService.savePresentation(presId, new File(dir, "cum_3.pptx").getAbsolutePath(), null);

        // Slide 4
        slideContentService.addSlide(presId, null, null, "Performance Benchmarks & SLAs", null);
        structuralElementService.addTable(
                presId, 3, 4, 4,
                50.0, 110.0, 860.0, 220.0,
                List.of(
                        List.of("Presentation Generation", "850 ms", "120 ms", "-85% (Optimal)"),
                        List.of("Slide-to-Image Rendering", "1,200 ms", "180 ms", "-85% (Native 2D)"),
                        List.of("Memory Footprint", "512 MB", "96 MB", "-81% (Efficient)")
                ),
                List.of("Operation", "Legacy Python", "Java POI 5.5", "Improvement"),
                "#1E3A8A", "#FFFFFF", "#F1F5F9", "#FFFFFF"
        );
        structuralElementService.addShape(presId, 3, "star", 860.0, 40.0, 45.0, 45.0, "#F59E0B", "#D97706", 1.0, null, null, null, null);
        structuralElementService.addShape(presId, 3, "round_rect", 50.0, 360.0, 860.0, 70.0, "#EFF6FF", "#BFDBFE", 1.0, "Key Takeaway: Java POI reduces execution latency by over 80% with native Graphics2D rendering.", "#1E3A8A", 13.0, true);
        presentationService.savePresentation(presId, new File(dir, "cum_4.pptx").getAbsolutePath(), null);

        // Slide 5
        layoutTemplateService.addMetricCardsSlide(
                presId, "Quarterly Adoption & Throughput", "Platform growth and performance acceleration",
                List.of(
                        Map.of("label", "Q1 Adoption", "value", "12.5k req/s", "trend", "+45%", "trendPositive", true, "detail", "Initial rollout"),
                        Map.of("label", "Q2 Growth", "value", "24.8k req/s", "trend", "+98%", "trendPositive", true, "detail", "Multi-region"),
                        Map.of("label", "Q3 Peak", "value", "38.2k req/s", "trend", "+54%", "trendPositive", true, "detail", "Production scale"),
                        Map.of("label", "Q4 Forecast", "value", "54.0k req/s", "trend", "+41%", "trendPositive", true, "detail", "Target SLA")
                ),
                "#1E3A8A", "#F8FAFC"
        );
        presentationService.savePresentation(presId, new File(dir, "cum_5.pptx").getAbsolutePath(), null);

        // Slide 6
        layoutTemplateService.addComparisonSlide(
                presId, "Architectural Comparison: Java MCP vs Python PPT", "Key Technology Drivers",
                List.of(
                        Map.of("header", "Java Spring AI MCP", "badge", "Recommended", "points", List.of("Native Graphics2D slide rendering"), "color", "#1E3A8A"),
                        Map.of("header", "Python PPT MCP", "badge", "Legacy", "points", List.of("Cannot render slides"), "color", "#64748B")
                )
        );
        presentationService.savePresentation(presId, new File(dir, "cum_6.pptx").getAbsolutePath(), null);

        // Slide 7
        layoutTemplateService.addProcessFlowSlide(
                presId, "Rollout & Execution Roadmap", "Phase-wise deployment plan",
                List.of(
                        Map.of("step", "1", "title", "Schema Design", "description", "LLM designs presentation plan", "duration", "Week 1"),
                        Map.of("step", "2", "title", "POI Build", "description", "MCP builds shapes, tables, charts", "duration", "Week 2-3")
                ),
                "#1E3A8A"
        );
        presentationService.savePresentation(presId, new File(dir, "cum_7.pptx").getAbsolutePath(), null);

        // Header & Footer
        designService.addHeaderFooter(presId, "CONFIDENTIAL - FOR INTERNAL USE ONLY", "Copyright 2026 Enterprise PPT MCP Server", true, 1);
        presentationService.savePresentation(presId, new File(dir, "cum_8_hf.pptx").getAbsolutePath(), null);
    }
}
