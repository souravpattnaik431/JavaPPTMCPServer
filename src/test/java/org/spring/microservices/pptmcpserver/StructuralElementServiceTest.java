package org.spring.microservices.pptmcpserver;

import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;
import org.spring.microservices.pptmcpserver.constants.Literals;
import org.spring.microservices.pptmcpserver.service.PresentationService;
import org.spring.microservices.pptmcpserver.service.RenderService;
import org.spring.microservices.pptmcpserver.service.SlideContentService;
import org.spring.microservices.pptmcpserver.service.StructuralElementService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Dedicated unit and integration tests for StructuralElementService and shape types.
 */
@Slf4j
@SpringBootTest
public class StructuralElementServiceTest {

    @Autowired
    private PresentationService presentationService;

    @Autowired
    private SlideContentService slideContentService;

    @Autowired
    private StructuralElementService structuralElementService;

    @Autowired
    private RenderService renderService;

    @Test
    void testVariousShapeTypesAndAliases() {
        log.info("--- Testing StructuralElementService Shape Types ---");

        // 1. Create a test presentation
        Map<String, Object> createRes = presentationService.createPresentation("Shape Gallery Showcase", 960.0, 540.0);
        String presId = (String) createRes.get(Literals.PRESENTATION_ID);
        assertNotNull(presId);

        // 2. Add a slide for testing shapes
        Map<String, Object> slideRes = slideContentService.addSlide(presId, null, null, "All PowerPoint Shapes Showcase", "Comprehensive shape support test");
        assertEquals(Literals.SUCCESS, slideRes.get(Literals.STATUS));

        // 3. Test list of various shapes (aliases + POI enums)
        List<ShapeTestCase> testShapes = List.of(
                new ShapeTestCase("rectangle", "RECT", 50, 100, "Rectangle"),
                new ShapeTestCase("round_rect", "ROUND_RECT", 180, 100, "Round Rect"),
                new ShapeTestCase("circle", "ELLIPSE", 310, 100, "Circle"),
                new ShapeTestCase("diamond", "DIAMOND", 440, 100, "Diamond"),
                new ShapeTestCase("triangle", "TRIANGLE", 570, 100, "Triangle"),
                new ShapeTestCase("trapezoid", "TRAPEZOID", 700, 100, "Trapezoid"),
                new ShapeTestCase("cylinder", "CAN", 830, 100, "Cylinder"),

                new ShapeTestCase("cube", "CUBE", 50, 220, "Cube"),
                new ShapeTestCase("donut", "DONUT", 180, 220, "Donut"),
                new ShapeTestCase("chevron", "CHEVRON", 310, 220, "Chevron"),
                new ShapeTestCase("star_8", "STAR_8", 440, 220, "8-Pt Star"),
                new ShapeTestCase("heart", "HEART", 570, 220, "Heart"),
                new ShapeTestCase("lightning", "LIGHTNING_BOLT", 700, 220, "Lightning"),
                new ShapeTestCase("cloud", "CLOUD_CALLOUT", 830, 220, "Cloud"),

                new ShapeTestCase("flow_process", "FLOW_CHART_PROCESS", 50, 340, "Process"),
                new ShapeTestCase("flow_decision", "FLOW_CHART_DECISION", 180, 340, "Decision"),
                new ShapeTestCase("SNIP_2_DIAG_RECT", "SNIP_2_DIAG_RECT", 310, 340, "Diag Snip"),
                new ShapeTestCase("parallelogram", "PARALLELOGRAM", 440, 340, "Parallelogram"),
                new ShapeTestCase("right_arrow", "RIGHT_ARROW", 570, 340, "Arrow"),
                new ShapeTestCase("banner", "RIBBON", 700, 340, "Banner"),
                new ShapeTestCase("unknown_shape_fallback", "RECT", 830, 340, "Default Rect")
        );

        for (ShapeTestCase tc : testShapes) {
            Map<String, Object> shapeRes = structuralElementService.addShape(
                    presId, 0, tc.inputName,
                    (double) tc.x, (double) tc.y, 100.0, 80.0,
                    "#3B82F6", "#1D4ED8", 1.5,
                    tc.label, "#FFFFFF", 10.0, true
            );
            assertEquals(Literals.SUCCESS, shapeRes.get(Literals.STATUS), "Failed adding shape: " + tc.inputName);
            assertEquals(tc.expectedShapeType, shapeRes.get(Literals.SHAPE_TYPE), "Mismatch shape type for: " + tc.inputName);
        }

        // 4. Save test PPT to test_output/Shapes_Showcase.pptx
        File outDir = new File("test_output");
        if (!outDir.exists()) {
            outDir.mkdirs();
        }
        File pptFile = new File(outDir, "Shapes_Showcase.pptx");
        Map<String, Object> saveRes = presentationService.savePresentation(presId, pptFile.getAbsolutePath(), null);
        assertEquals(Literals.SUCCESS, saveRes.get(Literals.STATUS));
        assertTrue(pptFile.exists());

        // 5. Render slide to image for visual verification
        File previewPng = new File(outDir, "shapes_showcase_preview.png");
        Map<String, Object> renderRes = renderService.renderSlideToImage(presId, 0, previewPng.getAbsolutePath(), 1.5, "PNG");
        assertEquals(Literals.SUCCESS, renderRes.get(Literals.STATUS));
        assertTrue(previewPng.exists());

        log.info("Successfully generated and verified Shapes_Showcase.pptx and preview image!");
    }

    private record ShapeTestCase(String inputName, String expectedShapeType, int x, int y, String label) {}
}
