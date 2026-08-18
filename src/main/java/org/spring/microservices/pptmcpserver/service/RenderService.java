package org.spring.microservices.pptmcpserver.service;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.spring.microservices.pptmcpserver.state.PresentationStateManager;
import org.springframework.ai.mcp.annotation.McpTool;
import org.springframework.ai.mcp.annotation.McpToolParam;
import org.springframework.stereotype.Service;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.util.*;
import java.util.List;

import static org.spring.microservices.pptmcpserver.constants.Literals.*;

/**
 * Service 6: Slide-to-Image Rendering & Server Diagnostics (3 MCP Tools).
 * Provides high-DPI Graphics2D slide rendering directly to PNG/JPEG image files and runtime diagnostic metadata.
 */
@Slf4j
@Service
@RequiredArgsConstructor
public class RenderService {

    private final PresentationStateManager stateManager;

    /**
     * Tool 1: Render a PowerPoint slide into a high-resolution PNG or JPEG image.
     *
     * @param presentationId Optional presentation ID.
     * @param slideIndex     0-based slide index to render.
     * @param outputPath     Target image file path.
     * @param scale          Resolution scale factor (default: 1.5).
     * @param format         Image format ('PNG' or 'JPEG').
     * @return Map with outputPath, image dimensions, and status.
     */
    @McpTool(name = "render_slide_to_image", description = "Render a slide to a PNG/JPEG image using native Java Graphics2D high-DPI rendering.")
    public Map<String, Object> renderSlideToImage(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "0-based slide index to render") int slideIndex,
            @McpToolParam(description = "Output image path (default: slide_{index}.png)", required = false) String outputPath,
            @McpToolParam(description = "Resolution scale factor (e.g. 1.0, 1.5, 2.0; default: 1.5 for sharp output)", required = false) Double scale,
            @McpToolParam(description = "Image format: 'PNG' or 'JPEG' (default: 'PNG')", required = false) String format) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Rendering slide {} to image for presentation ID '{}': scale={}, format='{}'", slideIndex, resolvedId, scale, format);
        XMLSlideShow ppt = stateManager.getPresentation(resolvedId);
        XSLFSlide slide = stateManager.getSlide(resolvedId, slideIndex);

        double sc = (scale != null && scale > 0) ? scale : 1.5;
        String fmt = (format != null && !format.trim().isEmpty()) ? format.trim().toUpperCase() : "PNG";

        String finalPath = (outputPath != null && !outputPath.trim().isEmpty())
                ? outputPath.trim()
                : "slide_" + slideIndex + "." + fmt.toLowerCase();

        File outputFile = new File(finalPath);
        File parentDir = outputFile.getParentFile();
        if (parentDir != null && !parentDir.exists()) {
            parentDir.mkdirs();
        }

        Dimension pgSize = ppt.getPageSize();
        int width = (int) Math.round(pgSize.width * sc);
        int height = (int) Math.round(pgSize.height * sc);

        int imageType = "JPEG".equalsIgnoreCase(fmt) || "JPG".equalsIgnoreCase(fmt)
                ? BufferedImage.TYPE_INT_RGB
                : BufferedImage.TYPE_INT_ARGB;

        BufferedImage img = new BufferedImage(width, height, imageType);
        Graphics2D graphics = img.createGraphics();

        graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        graphics.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
        graphics.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
        graphics.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);
        graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);

        graphics.setPaint(Color.WHITE);
        graphics.fillRect(0, 0, width, height);

        graphics.scale(sc, sc);
        slide.draw(graphics);
        graphics.dispose();

        try {
            ImageIO.write(img, fmt, outputFile);
            log.info("Successfully rendered slide {} to '{}' ({}x{}, scale={})", slideIndex, outputFile.getAbsolutePath(), width, height, sc);

            Map<String, Object> result = new LinkedHashMap<>();
            result.put(STATUS, SUCCESS);
            result.put(PRESENTATION_ID, resolvedId);
            result.put(SLIDE_INDEX, slideIndex);
            result.put(OUTPUT_PATH, outputFile.getAbsolutePath());
            result.put(IMAGE_WIDTH, width);
            result.put(IMAGE_HEIGHT, height);
            result.put(SCALE, sc);
            result.put(FORMAT, fmt);
            result.put(MESSAGE, "Slide successfully rendered to: " + outputFile.getAbsolutePath());
            return result;
        } catch (Exception e) {
            log.error("Failed to render slide {} image: {}", slideIndex, e.getMessage(), e);
            throw new RuntimeException("Failed to render slide image: " + e.getMessage(), e);
        }
    }

    /**
     * Tool 2: Render all PowerPoint slides into high-resolution images.
     *
     * @param presentationId Optional presentation ID.
     * @param outputDir      Target directory.
     * @param scale          Resolution scale factor.
     * @param format         Image format ('PNG' or 'JPEG').
     * @return Map with outputDir, slideCount, and list of rendered image paths.
     */
    @McpTool(name = "render_all_slides_to_images", description = "Render all slides in the presentation to a directory of PNG/JPEG images.")
    public Map<String, Object> renderAllSlidesToImages(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "Target directory to save rendered slide images", required = false) String outputDir,
            @McpToolParam(description = "Resolution scale factor (default: 1.5)", required = false) Double scale,
            @McpToolParam(description = "Image format: 'PNG' or 'JPEG' (default: 'PNG')", required = false) String format) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Batch rendering all slides in presentation ID '{}' to outputDir='{}'", resolvedId, outputDir);
        XMLSlideShow ppt = stateManager.getPresentation(resolvedId);
        List<XSLFSlide> slides = ppt.getSlides();

        String targetDir = (outputDir != null && !outputDir.trim().isEmpty()) ? outputDir.trim() : "rendered_slides";
        File dir = new File(targetDir);
        if (!dir.exists()) {
            dir.mkdirs();
        }

        String fmt = (format != null && !format.trim().isEmpty()) ? format.trim().toUpperCase() : "PNG";
        List<String> renderedPaths = new ArrayList<>();

        for (int i = 0; i < slides.size(); i++) {
            File slideImg = new File(dir, "slide_" + (i + 1) + "." + fmt.toLowerCase());
            renderSlideToImage(resolvedId, i, slideImg.getAbsolutePath(), scale, fmt);
            renderedPaths.add(slideImg.getAbsolutePath());
        }

        log.info("Successfully batch rendered {} slides to '{}'", slides.size(), dir.getAbsolutePath());
        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(SLIDE_COUNT, slides.size());
        result.put(OUTPUT_DIR, dir.getAbsolutePath());
        result.put(RENDERED_IMAGES, renderedPaths);
        result.put(MESSAGE, String.format("Rendered %d slides to directory: %s", slides.size(), dir.getAbsolutePath()));
        return result;
    }

    /**
     * Tool 3: Get server diagnostic information and feature capabilities.
     *
     * @return Map containing server metadata, total tool count, and active capabilities.
     */
    @McpTool(name = "get_server_info", description = "Retrieve PPT MCP Server runtime information, active service list, and total tool count.")
    public Map<String, Object> getServerInfo() {
        log.info("Retrieving server info & diagnostics");
        Map<String, Object> info = new LinkedHashMap<>();
        info.put("serverName", "PPTMCPServer");
        info.put("version", "1.1.0");
        info.put("framework", "Spring Boot 4.1.0 + Spring AI 2.0.0 + Apache POI 5.5.1");
        info.put("protocol", "STREAMABLE_HTTP (SYNC)");
        info.put("endpoint", "/mcp");
        info.put("totalTools", 43);
        info.put("activePresentations", stateManager.listAll().size());
        info.put("services", List.of(
                "PresentationService (9 tools: create, create_from_template, open, save, get_info, set_properties, list, switch, merge)",
                "SlideContentService (11 tools: add_slide, delete_slide, reorder_slide, get_slide_info, extract_text, add_textbox, format_text, add_image, add_hyperlink, set_speaker_notes, get_speaker_notes)",
                "StructuralElementService (5 tools: add_table, format_table_cell, add_shape, add_connector, get_slide_masters)",
                "DesignService (5 tools: set_slide_background, get_color_schemes, format_shape, set_slide_dimensions, add_header_footer)",
                "LayoutTemplateService (5 tools: add_split_content_slide, add_metric_cards_slide, add_comparison_slide, add_process_flow_slide, duplicate_slide)",
                "RenderService (3 tools: render_slide_to_image, render_all_slides_to_images, get_server_info)",
                "GuardrailService (5 tools: audit_presentation_quality, validate_slide_quality, validate_presentation_rules, autofix_presentation_issues, get_guardrail_guidelines)"
        ));
        info.put("superpowers", List.of(
                "Executive split-screen compound slides (text + image side-by-side) with auto-fitting",
                "Executive KPI metric dashboard slides with trend badges",
                "Multi-column architecture/solution comparison slides",
                "Sequential process flow and roadmap slides",
                "Native high-DPI Graphics2D Slide-to-PNG/JPEG rendering (single + batch)",
                "Slide duplication & cloning across presentations",
                "Presenter speaker notes management",
                "Corporate headers, footers, confidentiality notices, and slide numbering",
                "Branded PPTX template cloning and customization",
                "Automated Presentation Quality & Visual Boundary Overflow Auditing (0-100% scoring)",
                "User Instruction & Corporate Rule Compliance Verification",
                "One-Click Automated Issue Remediation (bounds clamping, ghost placeholder cleanup, font normalization)"
        ));
        return info;
    }
}
