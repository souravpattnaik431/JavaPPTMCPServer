package org.spring.microservices.pptmcpserver;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.spring.microservices.pptmcpserver.constants.Literals;
import org.spring.microservices.pptmcpserver.service.GuardrailService;
import org.spring.microservices.pptmcpserver.service.PresentationService;
import org.spring.microservices.pptmcpserver.service.SlideContentService;
import org.spring.microservices.pptmcpserver.state.PresentationStateManager;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.awt.geom.Rectangle2D;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Comprehensive integration tests for GuardrailService (5 tools).
 */
@Slf4j
@SpringBootTest
class GuardrailServiceTest {

    @Autowired
    private GuardrailService guardrailService;

    @Autowired
    private PresentationService presentationService;

    @Autowired
    private SlideContentService slideContentService;

    @Autowired
    private PresentationStateManager stateManager;

    private String presentationId;

    @BeforeEach
    void setUp() {
        Map<String, Object> createRes = presentationService.createPresentation("Guardrail Test Suite", 960.0, 540.0);
        presentationId = (String) createRes.get(Literals.PRESENTATION_ID);
        assertNotNull(presentationId);
    }

    @Test
    void testGetGuardrailGuidelines() {
        Map<String, Object> result = guardrailService.getGuardrailGuidelines();
        assertEquals(Literals.SUCCESS, result.get(Literals.STATUS));
        assertNotNull(result.get("guidelines"));

        @SuppressWarnings("unchecked")
        Map<String, Object> guidelines = (Map<String, Object>) result.get("guidelines");
        assertTrue(guidelines.containsKey("canvasDimensions"));
        assertTrue(guidelines.containsKey("layoutZones"));
        assertTrue(guidelines.containsKey("typographyScale"));
        assertTrue(guidelines.containsKey("colorAndContrastRules"));
        assertTrue(guidelines.containsKey("antiPatternsToAvoid"));
    }

    @Test
    void testAuditCleanPresentation() {
        // Add title slide and content slide
        slideContentService.addSlide(presentationId, 0, "Title Slide", "Executive Strategy 2026", "Quarterly Operations Review");
        slideContentService.addSlide(presentationId, 1, "Title and Content", "Cloud Modernization Pillars", null);
        slideContentService.addTextbox(presentationId, 1, "• Pillar 1: Hybrid Cloud Architecture\n• Pillar 2: Zero-Trust Security\n• Pillar 3: AI Automation",
                60.0, 140.0, 700.0, 200.0, 16.0, "#0F172A", null, false, false, false, null);

        Map<String, Object> audit = guardrailService.auditPresentationQuality(presentationId);
        log.info("Audit result in testAuditCleanPresentation: {}", audit);
        assertEquals(Literals.SUCCESS, audit.get(Literals.STATUS));
        assertEquals(2, audit.get(Literals.SLIDE_COUNT));

        int score = (int) audit.get(Literals.QUALITY_SCORE);
        assertTrue(score >= 80, "Expected high quality score for clean slides, but was: " + score);
        assertEquals(0, audit.get("totalErrors"), "Errors found: " + audit.get(Literals.SLIDE_AUDITS));
    }

    @Test
    void testAuditDefectivePresentationDetectsIssues() {
        // Add slide
        slideContentService.addSlide(presentationId, 1, "Title and Content", "Defects Demo Slide", null);

        // Inject out-of-bounds overflowing shape (x=1000, y=600 is way off 960x540 canvas)
        XMLSlideShow ppt = stateManager.getPresentation(presentationId);
        XSLFSlide slide = ppt.getSlides().getFirst();
        XSLFTextBox overflowBox = slide.createTextBox();
        overflowBox.setAnchor(new Rectangle2D.Double(850.0, 480.0, 300.0, 150.0));
        overflowBox.setText("This box overflows the bottom-right corner");

        // Inject unresolved placeholder text
        XSLFTextBox placeholderBox = slide.createTextBox();
        placeholderBox.setAnchor(new Rectangle2D.Double(100.0, 200.0, 300.0, 50.0));
        placeholderBox.setText("[TODO] Insert revenue financial data here");

        // Inject tiny font text
        XSLFTextBox tinyBox = slide.createTextBox();
        tinyBox.setAnchor(new Rectangle2D.Double(100.0, 300.0, 300.0, 50.0));
        var run = tinyBox.addNewTextParagraph().addNewTextRun();
        run.setText("Microscopic illegible text");
        run.setFontSize(7.0);

        Map<String, Object> audit = guardrailService.auditPresentationQuality(presentationId);
        assertEquals(Literals.SUCCESS, audit.get(Literals.STATUS));

        int errorCount = (int) audit.get("totalErrors");
        int warningCount = (int) audit.get("totalWarnings");

        assertTrue(errorCount >= 2, "Expected at least 2 errors for overflow and [TODO] placeholder, got: " + errorCount);
        assertTrue(warningCount >= 1, "Expected at least 1 warning for tiny font size, got: " + warningCount);

        // Test single slide validation tool
        Map<String, Object> slideVal = guardrailService.validateSlideQuality(presentationId, 0);
        assertEquals(Literals.SUCCESS, slideVal.get(Literals.STATUS));
        assertEquals(0, slideVal.get(Literals.SLIDE_INDEX));
        assertEquals(errorCount, slideVal.get("errorCount"));
    }

    @Test
    void testValidatePresentationRulesCustomInstructions() {
        slideContentService.addSlide(presentationId, 0, "Title Slide", "Company Overview", "Annual Investor Pitch");
        slideContentService.addSlide(presentationId, 1, "Title and Content", "Product Roadmap", null);
        slideContentService.addTextbox(presentationId, 1, "• Feature A: [TBD] pricing\n• Feature B: Live streaming",
                50.0, 120.0, 600.0, 150.0, 10.0, "#000000", null, false, false, false, null);

        // 1. Validate rules that should fail
        Map<String, Object> failResult = guardrailService.validatePresentationRules(
                presentationId,
                4, // minSlides (we have 2)
                null,
                null,
                List.of("Company Overview", "Financials"), // "Financials" is missing
                List.of("TBD", "TODO"), // "TBD" is present on slide 2
                14.0, // minFontSize 14.0 (slide 2 has 10.0)
                5,
                100,
                null,
                "16:9"
        );

        assertEquals(Literals.SUCCESS, failResult.get(Literals.STATUS));
        assertFalse((Boolean) failResult.get(Literals.PASSED));

        @SuppressWarnings("unchecked")
        List<Map<String, Object>> violations = (List<Map<String, Object>>) failResult.get(Literals.VIOLATIONS);
        assertTrue(violations.size() >= 3, "Expected multiple violations, got: " + violations.size());

        // 2. Validate rules that should pass
        Map<String, Object> passResult = guardrailService.validatePresentationRules(
                presentationId,
                1,
                5,
                2,
                List.of("Company Overview", "Product Roadmap"),
                List.of("FORBIDDEN_KEYWORD_NOT_PRESENT"),
                8.0,
                10,
                200,
                null,
                "16:9"
        );

        assertEquals(Literals.SUCCESS, passResult.get(Literals.STATUS));
        assertTrue((Boolean) passResult.get(Literals.PASSED));
        assertEquals(100, passResult.get(Literals.COMPLIANCE_PERCENTAGE));
    }

    @Test
    void testAutofixPresentationIssues() {
        slideContentService.addSlide(presentationId, 1, "Title and Content", "Auto-Fix Demonstration", null);

        XMLSlideShow ppt = stateManager.getPresentation(presentationId);
        XSLFSlide slide = ppt.getSlides().getFirst();

        // Add overflowing shape (x=900, y=500, w=200, h=100 overflows 960x540)
        XSLFTextBox overflowBox = slide.createTextBox();
        overflowBox.setAnchor(new Rectangle2D.Double(900.0, 500.0, 200.0, 100.0));
        var run = overflowBox.addNewTextParagraph().addNewTextRun();
        run.setText("Reposition me onto slide canvas");
        run.setFontSize(8.0); // undersized font

        // Add empty ghost placeholder
        XSLFTextBox ghostBox = slide.createTextBox();
        ghostBox.setAnchor(new Rectangle2D.Double(200.0, 200.0, 100.0, 50.0));
        ghostBox.setText("   ");

        // Run auto-fix
        Map<String, Object> fixResult = guardrailService.autofixPresentationIssues(
                presentationId,
                "Calibri",
                12.0,
                20.0
        );

        assertEquals(Literals.SUCCESS, fixResult.get(Literals.STATUS));
        @SuppressWarnings("unchecked")
        Map<String, Object> fixes = (Map<String, Object>) fixResult.get(Literals.FIXES_APPLIED);
        assertTrue((int) fixes.get("boundsClampedShapes") >= 1);
        assertTrue((int) fixes.get("placeholdersRemoved") >= 1);
        assertTrue((int) fixes.get("fontSizesAdjusted") >= 1);

        // Verify shape anchor is now strictly within bounds
        Rectangle2D newAnchor = overflowBox.getAnchor();
        assertTrue(newAnchor.getX() >= 20.0);
        assertTrue(newAnchor.getY() >= 20.0);
        assertTrue(newAnchor.getX() + newAnchor.getWidth() <= 940.0 + 1.0);
        assertTrue(newAnchor.getY() + newAnchor.getHeight() <= 520.0 + 1.0);
        assertEquals(12.0, run.getFontSize());
        assertEquals("Calibri", run.getFontFamily());
    }
}
