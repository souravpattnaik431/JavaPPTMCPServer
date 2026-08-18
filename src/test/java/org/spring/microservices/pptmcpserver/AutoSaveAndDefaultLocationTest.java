package org.spring.microservices.pptmcpserver;

import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.spring.microservices.pptmcpserver.service.LayoutTemplateService;
import org.spring.microservices.pptmcpserver.service.PresentationService;
import org.spring.microservices.pptmcpserver.service.SlideContentService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertTrue;

@Slf4j
@SpringBootTest
class AutoSaveAndDefaultLocationTest {

    @Autowired
    private PresentationService presentationService;

    @Autowired
    private SlideContentService slideContentService;

    @Autowired
    private LayoutTemplateService layoutTemplateService;

    @Autowired
    private org.spring.microservices.pptmcpserver.state.PresentationStateManager stateManager;

    @org.junit.jupiter.api.BeforeEach
    void setUp() {
        stateManager.clearAll();
    }

    @Test
    @DisplayName("Test 1: User gives no path and no filename - Auto-names and auto-saves without save_presentation call")
    void testAutoSaveWithTitleOnly() {
        String promptTitle = "Sauce Labs vs TestMu (LambdaTest) Playwright Comparison";
        Map<String, Object> pres = presentationService.createPresentation(promptTitle, 960.0, 540.0, null);
        String pId = (String) pres.get("presentationId");
        String mappedPath = (String) pres.get("outputPath");

        log.info("Mapped automatic output path: {}", mappedPath);
        assertTrue(mappedPath.contains("Sauce_Labs_vs_TestMu__LambdaTest__Playwright_Comparison.pptx") || mappedPath.endsWith(".pptx"));

        // Add slides (Notice: NEVER calling save_presentation)
        slideContentService.addSlide(pId, null, null, "Introduction", "Overview of Playwright test runners");
        layoutTemplateService.addComparisonSlide(pId, "Feature Breakdown", "Head to head", List.of(
                Map.of("header", "Sauce Labs", "points", List.of("Enterprise grid", "Real devices")),
                Map.of("header", "TestMu (LambdaTest)", "points", List.of("HyperExecute speed", "SmartUI"))
        ));

        // Verify the file was auto-saved to disk immediately!
        File generatedFile = new File(mappedPath);
        assertTrue(generatedFile.exists(), "The .pptx file must be auto-created on disk without needing save_presentation");
        assertTrue(generatedFile.length() > 0, "The .pptx file must not be empty");
        log.info("Verified auto-saved presentation on disk: {} (size: {} bytes)", generatedFile.getAbsolutePath(), generatedFile.length());
    }

    @Test
    @DisplayName("Test 2: User specifies a relative folder/path - Auto-saves into that specific folder")
    void testAutoSaveWithCustomFolder() {
        String customFolder = "test_output" + File.separator + "custom_folder";
        Map<String, Object> pres = presentationService.createPresentation("DevOps Pipeline", 960.0, 540.0, customFolder + File.separator);
        String pId = (String) pres.get("presentationId");
        String mappedPath = (String) pres.get("outputPath");

        layoutTemplateService.addProcessFlowSlide(pId, "CI/CD Gates", "4 Steps", List.of(
                Map.of("step", "1", "title", "Lint", "description", "SonarQube gate")
        ), "#1E3A8A");

        File generatedFile = new File(mappedPath);
        assertTrue(generatedFile.exists(), "File must exist in custom folder");
        assertTrue(generatedFile.getAbsolutePath().contains("custom_folder"));
    }

    @Test
    @DisplayName("Test 3: Copilot calls addSlide directly without createPresentation or savePresentation — Auto-saves to workspace")
    void testDirectAddSlideAutoSavesWithoutCreateOrSave() {
        // Direct call with null presentationId, null createPresentation, no savePresentation
        Map<String, Object> slideRes = slideContentService.addSlide(
                null, 0, "Title Slide", "Copilot Automated PR Review", "AI-Powered Quality Engineering"
        );
        String presId = (String) slideRes.get("presentationId");
        log.info("Direct addSlide presentation ID: {}", presId);

        // Add a second slide using layout template
        layoutTemplateService.addMetricCardsSlide(
                presId, "Key Performance Gains", "Review Speedup",
                List.of(Map.of("label", "Review Time", "value", "-70%", "trend", "3.5x", "trendPositive", true)),
                "#1E3A8A", "#F8FAFC"
        );

        // Verify the file was auto-named based on title and auto-saved to disk!
        File autoSavedFile = new File("Copilot_Automated_PR_Review.pptx");
        assertTrue(autoSavedFile.exists(), "File must be auto-created on disk based on title even when createPresentation was never called");
        assertTrue(autoSavedFile.length() > 0, "Auto-saved file must not be empty");
        log.info("Verified direct Copilot workflow auto-saved file: {} (size: {} bytes)", autoSavedFile.getAbsolutePath(), autoSavedFile.length());
        
        // Cleanup test artifact
        autoSavedFile.deleteOnExit();
    }
}
