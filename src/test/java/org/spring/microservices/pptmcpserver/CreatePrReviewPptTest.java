package org.spring.microservices.pptmcpserver;

import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;
import org.spring.microservices.pptmcpserver.constants.Literals;
import org.spring.microservices.pptmcpserver.service.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Creates the client-facing 3-slide presentation on "AI PR Review GitHub Copilot Agent for Selenium & Playwright"
 * using the PPT MCP Server tools, saves it to the workspace root, and exports slide preview PNGs.
 */
@Slf4j
@SpringBootTest
public class CreatePrReviewPptTest {

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
    void createPrReviewPresentation() {
        File outputDir = new File("test_output/ai_pr_review");
        if (!outputDir.exists()) {
            outputDir.mkdirs();
        }

        String targetPptxPath = new File("AI_PR_Review_Copilot_Selenium_Playwright.pptx").getAbsolutePath();
        log.info("Generating 3-Slide AI PR Review Presentation at: {}", targetPptxPath);

        // 1. Create 16:9 Presentation
        String presTitle = "AI PR Review Copilot Agent for Selenium & Playwright";
        Map<String, Object> createRes = presentationService.createPresentation(presTitle, 960.0, 540.0, targetPptxPath);
        String presId = (String) createRes.get(Literals.PRESENTATION_ID);
        assertNotNull(presId, "Presentation ID must not be null");

        // 2. Set Metadata Properties
        presentationService.setCoreProperties(
                presId,
                "AI PR Review Copilot Agent: Selenium & Playwright",
                "Quality Engineering & AI Center of Excellence",
                "Pull Request Review AI Agent Use Case",
                "GitHub Copilot, PR Review, Selenium, Playwright, Test Automation, AI Agent, QA",
                "Client-facing architecture and business value pitch for AI-driven PR review in test automation."
        );

        // =========================================================================
        // SLIDE 1: Title & Executive Overview
        // =========================================================================
        slideContentService.addSlide(presId, 0, "Blank", null, null);
        designService.setSlideBackground(presId, 0, "#F8FAFC", null, null, null);

        // Top Category Pill Badge
        structuralElementService.addShape(
                presId, 0, "round_rect",
                300.0, 30.0, 360.0, 28.0,
                "#EFF6FF", "#93C5FD", 1.0,
                "AI-POWERED QUALITY ENGINEERING & TEST GOVERNANCE",
                "#1E40AF", 10.5, true
        );

        // Main Title
        slideContentService.addTextbox(
                presId, 0,
                "AI PR Review Copilot Agent",
                50.0, 68.0, 860.0, 48.0,
                32.0, "#0F172A", "Segoe UI", true, false, false, "CENTER"
        );

        // Secondary Title
        slideContentService.addTextbox(
                presId, 0,
                "Intelligent Pull Request Review for Selenium & Playwright Automation Suites",
                50.0, 118.0, 860.0, 32.0,
                18.0, "#1E3A8A", "Segoe UI", true, false, false, "CENTER"
        );

        // Subtitle Context
        slideContentService.addTextbox(
                presId, 0,
                "Autonomous Test Script Review, Locator Healing & CI/CD Quality Gates for Enterprise Test Suites",
                50.0, 155.0, 860.0, 30.0,
                13.5, "#64748B", "Segoe UI", false, false, false, "CENTER"
        );

        // 3 Cards on Slide 1
        structuralElementService.addShape(
                presId, 0, "round_rect",
                50.0, 205.0, 266.0, 280.0,
                "#FFFFFF", "#CBD5E1", 1.5,
                null, null, null, null
        );
        structuralElementService.addShape(
                presId, 0, "round_rect",
                50.0, 205.0, 266.0, 45.0,
                "#1E3A8A", "#1E3A8A", 1.0,
                "Framework Coverage",
                "#FFFFFF", 14.0, true
        );
        slideContentService.addTextbox(
                presId, 0,
                "● SUPPORTED FRAMEWORKS\n\n• Selenium 4.x (Java, Python, C#)\n• Playwright (TypeScript, Java)\n• Cypress & WebdriverIO suites\n• Hybrid BDD / Cucumber steps\n• Unified Page Object Model (POM)",
                65.0, 260.0, 236.0, 210.0,
                12.5, "#334155", "Segoe UI", false, false, false, "LEFT"
        );

        structuralElementService.addShape(
                presId, 0, "round_rect",
                346.0, 205.0, 266.0, 280.0,
                "#FFFFFF", "#CBD5E1", 1.5,
                null, null, null, null
        );
        structuralElementService.addShape(
                presId, 0, "round_rect",
                346.0, 205.0, 266.0, 45.0,
                "#0D9488", "#0D9488", 1.0,
                "AI Agent Capabilities",
                "#FFFFFF", 14.0, true
        );
        slideContentService.addTextbox(
                presId, 0,
                "● GITHUB COPILOT AGENT\n\n• Smart locator healing suggestions\n• Detects hardcoded sleep() & waits\n• Flags missing / weak assertions\n• Validates async/await safety\n• Automated inline PR diff reviews",
                361.0, 260.0, 236.0, 210.0,
                12.5, "#334155", "Segoe UI", false, false, false, "LEFT"
        );

        structuralElementService.addShape(
                presId, 0, "round_rect",
                642.0, 205.0, 266.0, 280.0,
                "#FFFFFF", "#CBD5E1", 1.5,
                null, null, null, null
        );
        structuralElementService.addShape(
                presId, 0, "round_rect",
                642.0, 205.0, 266.0, 45.0,
                "#7C3AED", "#7C3AED", 1.0,
                "Target Client Outcomes",
                "#FFFFFF", 14.0, true
        );
        slideContentService.addTextbox(
                presId, 0,
                "● ENTERPRISE VALUE\n\n• 70% Faster PR approval cycles\n• 85% Reduction in test flakiness\n• Zero unhandled race conditions\n• Continuous SDET code governance\n• Seamless CI/CD pipeline gating",
                657.0, 260.0, 236.0, 210.0,
                12.5, "#334155", "Segoe UI", false, false, false, "LEFT"
        );

        slideContentService.setSpeakerNotes(
                presId,
                0,
                "Welcome. In this presentation, we explore how deploying a specialized GitHub Copilot PR Review Agent transforms Selenium and Playwright test automation workflows by automatically identifying flaky locators, catching missing assertions, and accelerating release cycles."
        );

        // =========================================================================
        // SLIDE 2: Why It Is Needed (Problem Statement & Automation Challenges)
        // =========================================================================
        layoutTemplateService.addComparisonSlide(
                presId,
                "Why AI PR Review for Test Automation?",
                "Key Industry Challenges & Bottlenecks in Selenium & Playwright Test Pull Requests",
                List.of(
                        Map.of(
                                "header", "Brittle Selectors & Flakiness",
                                "badge", "High Test Fragility",
                                "color", "#DC2626",
                                "points", List.of(
                                        "Over-reliance on brittle absolute XPaths and auto-generated CSS classes",
                                        "Minor UI cosmetic changes break tests, creating false alarm alerts",
                                        "Flaky retries mask actual regression bugs and erode trust in QA",
                                        "Developers spend 15+ hours weekly investigating false test failures"
                                )
                        ),
                        Map.of(
                                "header", "Anti-Patterns & Code Smells",
                                "badge", "Maintainability Debt",
                                "color", "#D97706",
                                "points", List.of(
                                        "Hardcoded Thread.sleep() and unhandled async/await race conditions",
                                        "Weak or missing assertions leading to silent false-positive passes",
                                        "Duplicate locators violating Page Object Model (POM) DRY principles",
                                        "Inconsistent migration practices from Selenium to modern Playwright"
                                )
                        ),
                        Map.of(
                                "header", "High Review Latency",
                                "badge", "Delivery Bottleneck",
                                "color", "#2563EB",
                                "points", List.of(
                                        "Senior SDETs spend 40%+ bandwidth reviewing repetitive PR boilerplate",
                                        "2 to 3 day review turnaround delays sprint release velocity",
                                        "Subjective code reviews lead to uneven quality across global teams",
                                        "Test suite maintenance overhead scales faster than feature delivery"
                                )
                        )
                )
        );
        slideContentService.setSpeakerNotes(
                presId,
                1,
                "Slide 2 details the core industry pain points: brittle selectors causing test flakiness, widespread anti-patterns like hardcoded sleeps, and review bottlenecks where senior engineers waste time on routine checks."
        );

        // =========================================================================
        // SLIDE 3: Key Advantages & Client Business Value (ROI & Outcomes)
        // =========================================================================
        layoutTemplateService.addMetricCardsSlide(
                presId,
                "Key Advantages & Measurable Client Impact",
                "Enterprise ROI, Velocity Gains & Quality Governance Delivered by AI Copilot Agent",
                List.of(
                        Map.of(
                                "label", "Review Latency",
                                "value", "-70%",
                                "trend", "3.5x Speedup",
                                "trendPositive", true,
                                "detail", "Automated inline reviews delivered in < 60s directly on GitHub PRs"
                        ),
                        Map.of(
                                "label", "Flaky Test Reduction",
                                "value", "85%",
                                "trend", "High Stability",
                                "trendPositive", true,
                                "detail", "Auto-recommends robust user-facing locators (getByRole, getByTestId)"
                        ),
                        Map.of(
                                "label", "Guardrail Enforcement",
                                "value", "100%",
                                "trend", "Zero Anti-Patterns",
                                "trendPositive", true,
                                "detail", "Instantly flags missing awaits, sleep() calls, & improper assertions"
                        ),
                        Map.of(
                                "label", "Annual Cost Savings",
                                "value", "$120k+",
                                "trend", "High ROI",
                                "trendPositive", true,
                                "detail", "Reclaims 20+ engineering hours per SDET/month for core test innovation"
                        )
                ),
                "#1E3A8A",
                "#F8FAFC"
        );
        slideContentService.setSpeakerNotes(
                presId,
                2,
                "Slide 3 highlights the quantifiable business impact: 70% reduction in review cycle times, 85% fewer flaky tests via smart locator healing, 100% compliance with automation standards, and significant cost savings."
        );

        // Add Executive Takeaway Banner on Slide 3 with compact height & centered text
        structuralElementService.addShape(
                presId, 2, "round_rect",
                50.0, 425.0, 860.0, 50.0,
                "#EFF6FF", "#3B82F6", 1.5,
                "Client Value Summary: Transforms QA from a reactive bottleneck into an autonomous, high-velocity quality engine powered by GitHub Copilot AI.",
                "#1E3A8A", 12.5, true
        );

        // 3. Persist presentation to disk
        Map<String, Object> saveRes = presentationService.savePresentation(presId, targetPptxPath, null);
        log.info("Presentation saved result: {}", saveRes);

        File pptxFile = new File(targetPptxPath);
        assertTrue(pptxFile.exists(), "Target PPTX file must exist on disk");
        assertTrue(pptxFile.length() > 0, "Target PPTX file must not be empty");
        log.info("Successfully generated PPTX file at {} (size: {} bytes)", pptxFile.getAbsolutePath(), pptxFile.length());

        // 4. Render slide images to verify visuals
        for (int i = 0; i < 3; i++) {
            File imgFile = new File(outputDir, "slide_" + (i + 1) + ".png");
            Map<String, Object> renderRes = renderService.renderSlideToImage(presId, i, imgFile.getAbsolutePath(), 1.5, "PNG");
            log.info("Rendered slide {} to {}: {}", i + 1, imgFile.getAbsolutePath(), renderRes);
            assertTrue(imgFile.exists(), "Slide image must exist: " + imgFile.getAbsolutePath());
        }
    }
}
