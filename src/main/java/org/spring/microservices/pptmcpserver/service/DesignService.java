package org.spring.microservices.pptmcpserver.service;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.xslf.usermodel.*;
import org.spring.microservices.pptmcpserver.state.PresentationStateManager;
import org.spring.microservices.pptmcpserver.util.ColorHelper;
import org.springframework.ai.mcp.annotation.McpTool;
import org.springframework.ai.mcp.annotation.McpToolParam;
import org.springframework.stereotype.Service;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.Point2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.util.*;
import java.util.List;

import static org.spring.microservices.pptmcpserver.constants.Literals.*;

/**
 * Service 4: Visual Design & Styling (5 MCP Tools).
 * Manages background colors, Java2D gradients, curated color palettes, shape formatting, aspect ratios, and corporate headers/footers.
 */
@Slf4j
@Service
@RequiredArgsConstructor
public class DesignService {

    private final PresentationStateManager stateManager;

    /**
     * Tool 1: Set solid or gradient slide background.
     *
     * @param presentationId Optional presentation ID.
     * @param slideIndex     0-based slide index.
     * @param colorHex       Solid fill color hex/name.
     * @param gradientType   Gradient orientation: 'horizontal', 'vertical', 'diagonal'.
     * @param startColorHex  Gradient start color.
     * @param endColorHex    Gradient end color.
     * @return Map confirming background application.
     */
    @McpTool(name = "set_slide_background", description = "Set a solid color or gradient background on a slide.")
    public Map<String, Object> setSlideBackground(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "0-based slide index") int slideIndex,
            @McpToolParam(description = "Solid fill color hex/name (e.g. '#0F172A', 'white')", required = false) String colorHex,
            @McpToolParam(description = "Gradient orientation: 'horizontal', 'vertical', 'diagonal'", required = false) String gradientType,
            @McpToolParam(description = "Gradient start color hex/name", required = false) String startColorHex,
            @McpToolParam(description = "Gradient end color hex/name", required = false) String endColorHex) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Setting background on slide {} in presentation ID '{}': solid='{}', gradient='{}' ({} -> {})",
                slideIndex, resolvedId, colorHex, gradientType, startColorHex, endColorHex);
        XMLSlideShow ppt = stateManager.getPresentation(resolvedId);
        XSLFSlide slide = stateManager.getSlide(resolvedId, slideIndex);

        if (startColorHex != null && endColorHex != null) {
            // Apply gradient background
            Color c1 = ColorHelper.parseColor(startColorHex, new Color(15, 23, 42));
            Color c2 = ColorHelper.parseColor(endColorHex, new Color(59, 130, 246));

            int imgW = 1920;
            int imgH = 1080;
            BufferedImage gradientImg = new BufferedImage(imgW, imgH, BufferedImage.TYPE_INT_RGB);
            Graphics2D g2d = gradientImg.createGraphics();

            String gType = (gradientType != null) ? gradientType.trim().toLowerCase() : "diagonal";
            Point2D startPt = new Point2D.Double(0, 0);
            Point2D endPt = switch (gType) {
                case "vertical" -> new Point2D.Double(0, imgH);
                case "horizontal" -> new Point2D.Double(imgW, 0);
                default -> new Point2D.Double(imgW, imgH);
            };

            g2d.setPaint(new GradientPaint(startPt, c1, endPt, c2));
            g2d.fillRect(0, 0, imgW, imgH);
            g2d.dispose();

            try {
                ByteArrayOutputStream baos = new ByteArrayOutputStream();
                ImageIO.write(gradientImg, "png", baos);
                XSLFPictureData pd = ppt.addPicture(baos.toByteArray(), PictureData.PictureType.PNG);
                XSLFPictureShape pic = slide.createPicture(pd);
                pic.setAnchor(new Rectangle2D.Double(0, 0, ppt.getPageSize().getWidth(), ppt.getPageSize().getHeight()));

                log.info("Successfully applied {} gradient background to slide {}", gType, slideIndex);
                Map<String, Object> result = new LinkedHashMap<>();
                result.put(STATUS, SUCCESS);
                result.put(PRESENTATION_ID, resolvedId);
                result.put(SLIDE_INDEX, slideIndex);
                result.put(TYPE, "GRADIENT");
                result.put(GRADIENT_TYPE, gType);
                result.put(MESSAGE, "Gradient background applied successfully.");
                stateManager.autoSave(resolvedId);
                return result;
            } catch (Exception e) {
                log.error("Failed to apply gradient background on slide {}: {}", slideIndex, e.getMessage(), e);
                throw new RuntimeException("Failed to apply gradient background: " + e.getMessage(), e);
            }
        } else {
            // Solid background
            Color color = ColorHelper.parseColor(colorHex, Color.WHITE);
            XSLFBackground bg = slide.getBackground();
            bg.setFillColor(color);

            log.info("Successfully applied solid background ({}) to slide {}", colorHex, slideIndex);
            Map<String, Object> result = new LinkedHashMap<>();
            result.put(STATUS, SUCCESS);
            result.put(PRESENTATION_ID, resolvedId);
            result.put(SLIDE_INDEX, slideIndex);
            result.put(TYPE, "SOLID");
            result.put(COLOR, colorHex);
            result.put(MESSAGE, "Solid background color applied successfully.");
            stateManager.autoSave(resolvedId);
            return result;
        }
    }

    /**
     * Tool 2: Get curated professional color schemes.
     *
     * @return Map containing design color palettes.
     */
    @McpTool(name = "get_color_schemes", description = "Get curated designer color palettes for corporate, tech, modern, and dark presentation themes.")
    public Map<String, Object> getColorSchemes() {
        log.info("Retrieving curated designer color schemes");
        Map<String, Object> schemes = new LinkedHashMap<>();

        schemes.put("modern_blue", Map.of(
                PRIMARY, HEX_1E3A8A,
                SECONDARY, "#3B82F6",
                ACCENT, "#60A5FA",
                BACKGROUND, HEX_WHITE,
                CARD_BG, HEX_F8FAFC,
                TEXT_PRIMARY, HEX_0F172A,
                TEXT_SECONDARY, "#64748B"
        ));

        schemes.put("corporate_slate", Map.of(
                PRIMARY, HEX_0F172A,
                SECONDARY, "#334155",
                ACCENT, "#0EA5E9",
                BACKGROUND, HEX_F8FAFC,
                CARD_BG, HEX_WHITE,
                TEXT_PRIMARY, HEX_0F172A,
                TEXT_SECONDARY, "#64748B"
        ));

        schemes.put("emerald_growth", Map.of(
                PRIMARY, "#064E3B",
                SECONDARY, "#059669",
                ACCENT, "#10B981",
                BACKGROUND, "#F0FDF4",
                CARD_BG, HEX_WHITE,
                TEXT_PRIMARY, "#064E3B",
                TEXT_SECONDARY, "#047857"
        ));

        schemes.put("midnight_purple", Map.of(
                PRIMARY, "#4C1D95",
                SECONDARY, "#7C3AED",
                ACCENT, "#A855F7",
                BACKGROUND, HEX_0F172A,
                CARD_BG, "#1E293B",
                TEXT_PRIMARY, HEX_F8FAFC,
                TEXT_SECONDARY, "#94A3B8"
        ));

        schemes.put("sunset_warmth", Map.of(
                PRIMARY, "#9A3412",
                SECONDARY, "#EA580C",
                ACCENT, "#F97316",
                BACKGROUND, "#FFF7ED",
                CARD_BG, HEX_WHITE,
                TEXT_PRIMARY, "#431407",
                TEXT_SECONDARY, "#9A3412"
        ));

        schemes.put("cyber_neon", Map.of(
                PRIMARY, "#00F0FF",
                SECONDARY, "#7000FF",
                ACCENT, "#FF007F",
                BACKGROUND, "#0A0A0F",
                CARD_BG, "#14141F",
                TEXT_PRIMARY, HEX_WHITE,
                TEXT_SECONDARY, "#A0A0B0"
        ));

        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(SCHEMES_COUNT, schemes.size());
        result.put(SCHEMES, schemes);
        return result;
    }

    /**
     * Tool 3: Format an existing shape's fill color, line color, border width, or rotation.
     *
     * @param presentationId Optional presentation ID.
     * @param slideIndex     0-based slide index.
     * @param shapeIndex     0-based shape index.
     * @param shapeName      Optional shape name.
     * @param fillColor      Fill color hex/name.
     * @param lineColor      Line border color hex/name.
     * @param lineWidth      Border width in points.
     * @param rotation       Rotation angle in degrees.
     * @return Map confirming formatting.
     */
    @McpTool(name = "format_shape", description = "Format fill color, border line, and rotation of an existing shape on a slide.")
    public Map<String, Object> formatShape(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "0-based slide index") int slideIndex,
            @McpToolParam(description = "0-based shape index", required = false) Integer shapeIndex,
            @McpToolParam(description = "Shape name", required = false) String shapeName,
            @McpToolParam(description = "Fill color hex/name", required = false) String fillColor,
            @McpToolParam(description = "Line border color hex/name", required = false) String lineColor,
            @McpToolParam(description = "Border width in points", required = false) Double lineWidth,
            @McpToolParam(description = "Rotation angle in degrees (0 - 360)", required = false) Double rotation) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Formatting shape on slide {} (index={}, name='{}') in presentation ID '{}'", slideIndex, shapeIndex, shapeName, resolvedId);
        XSLFSlide slide = stateManager.getSlide(resolvedId, slideIndex);

        XSLFSimpleShape shape = findSimpleShape(slide, shapeIndex, shapeName);
        if (shape == null) {
            throw new IllegalArgumentException(String.format(
                    "No customizable shape found on slide %d matching shapeIndex=%s, shapeName='%s'. Action Hint: Use 'add_shape' or 'add_textbox' first to create a shape, or inspect slide shapes with 'inspect_slide_elements'.",
                    slideIndex, shapeIndex, shapeName));
        }

        if (fillColor != null) {
            Color fc = ColorHelper.parseColor(fillColor);
            if (fc != null) shape.setFillColor(fc);
        }
        if (lineColor != null) {
            Color lc = ColorHelper.parseColor(lineColor);
            if (lc != null) shape.setLineColor(lc);
        }
        if (lineWidth != null) {
            shape.setLineWidth(lineWidth);
        }
        if (rotation != null) {
            shape.setRotation(rotation);
        }

        log.info("Successfully formatted shape '{}' on slide {}", shape.getShapeName(), slideIndex);
        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(SLIDE_INDEX, slideIndex);
        result.put(SHAPE_NAME, shape.getShapeName());
        result.put(MESSAGE, "Shape formatted successfully.");
        stateManager.autoSave(resolvedId);
        return result;
    }

    /**
     * Tool 4: Set presentation slide dimensions / aspect ratio.
     *
     * @param presentationId Optional presentation ID.
     * @param preset         Preset: '16:9', '4:3', 'A4', 'custom'.
     * @param width          Custom width in points.
     * @param height         Custom height in points.
     * @return Map with width, height, and status.
     */
    @McpTool(name = "set_slide_dimensions", description = "Set slide aspect ratio ('16:9', '4:3', 'A4') or custom width and height in points.")
    public Map<String, Object> setSlideDimensions(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "Preset name: '16:9', '4:3', 'A4', 'custom'", required = false) String preset,
            @McpToolParam(description = "Custom width in points (if preset is 'custom')", required = false) Double width,
            @McpToolParam(description = "Custom height in points (if preset is 'custom')", required = false) Double height) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Setting slide dimensions for presentation ID '{}': preset='{}', width={}, height={}", resolvedId, preset, width, height);
        XMLSlideShow ppt = stateManager.getPresentation(resolvedId);

        double w;
        double h;
        String chosenPreset = (preset != null) ? preset.trim().toLowerCase() : "16:9";

        switch (chosenPreset) {
            case "4:3" -> {
                w = 720.0;
                h = 540.0;
            }
            case "a4" -> {
                w = 842.0;
                h = 595.0;
            }
            case "custom" -> {
                w = (width != null && width > 0) ? width : 960.0;
                h = (height != null && height > 0) ? height : 540.0;
            }
            default -> { // 16:9
                w = 960.0;
                h = 540.0;
                chosenPreset = "16:9";
            }
        }

        ppt.setPageSize(new Dimension((int) Math.round(w), (int) Math.round(h)));
        log.info("Successfully set slide dimensions to {}x{} ({}) for presentation '{}'", w, h, chosenPreset, resolvedId);

        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(PRESET, chosenPreset);
        result.put(WIDTH, w);
        result.put(HEIGHT, h);
        result.put(MESSAGE, String.format("Slide dimensions set to %.0f x %.0f (%s)", w, h, chosenPreset));
        stateManager.autoSave(resolvedId);
        return result;
    }

    /**
     * Tool 5: Apply consistent header, footer, confidentiality notice, and slide numbers across slides.
     *
     * @param presentationId     Optional presentation ID.
     * @param headerText         Header banner text.
     * @param footerText         Footer text.
     * @param includePageNumber  Whether to include slide numbers.
     * @param startSlideIndex    0-based start slide index.
     * @return Map with modified slides count.
     */
    @McpTool(name = "add_header_footer", description = "Apply corporate headers, footers, confidentiality notices, and slide numbers across all slides.")
    public Map<String, Object> addHeaderFooter(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "Header banner text (e.g., 'COMPANY CONFIDENTIAL')", required = false) String headerText,
            @McpToolParam(description = "Footer text (e.g., '© 2026 Enterprise Corp. All Rights Reserved.')", required = false) String footerText,
            @McpToolParam(description = "Include slide number in footer (default: true)", required = false) Boolean includePageNumber,
            @McpToolParam(description = "0-based start slide index (e.g., 1 to skip title slide; default: 0)", required = false) Integer startSlideIndex) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Applying headers/footers to presentation ID '{}': header='{}', footer='{}', startSlideIndex={}",
                resolvedId, headerText, footerText, startSlideIndex);
        XMLSlideShow ppt = stateManager.getPresentation(resolvedId);
        List<XSLFSlide> slides = ppt.getSlides();

        int start = (startSlideIndex != null && startSlideIndex >= 0) ? startSlideIndex : 0;
        boolean pageNum = includePageNumber == null || includePageNumber;

        Dimension pgSize = ppt.getPageSize();
        double w = pgSize.getWidth();
        double h = pgSize.getHeight();

        int modifiedCount = 0;
        for (int i = start; i < slides.size(); i++) {
            XSLFSlide slide = slides.get(i);

            // Header banner
            if (headerText != null && !headerText.trim().isEmpty()) {
                XSLFTextBox hBox = slide.createTextBox();
                hBox.setAnchor(new Rectangle2D.Double(50, 10, w - 100, 20));
                hBox.clearText();
                XSLFTextParagraph hp = hBox.addNewTextParagraph();
                hp.setTextAlign(org.apache.poi.sl.usermodel.TextParagraph.TextAlign.RIGHT);
                XSLFTextRun hr = hp.addNewTextRun();
                hr.setText(headerText.trim());
                hr.setFontSize(9.0);
                hr.setFontColor(new Color(148, 163, 184));
            }

            // Footer text
            if (footerText != null && !footerText.trim().isEmpty()) {
                XSLFTextBox fBox = slide.createTextBox();
                fBox.setAnchor(new Rectangle2D.Double(50, h - 30, w - 160, 20));
                fBox.clearText();
                XSLFTextParagraph fp = fBox.addNewTextParagraph();
                XSLFTextRun fr = fp.addNewTextRun();
                fr.setText(footerText.trim());
                fr.setFontSize(9.0);
                fr.setFontColor(new Color(148, 163, 184));
            }

            // Slide number
            if (pageNum) {
                XSLFTextBox pBox = slide.createTextBox();
                pBox.setAnchor(new Rectangle2D.Double(w - 100, h - 30, 50, 20));
                pBox.clearText();
                XSLFTextParagraph pp = pBox.addNewTextParagraph();
                pp.setTextAlign(org.apache.poi.sl.usermodel.TextParagraph.TextAlign.RIGHT);
                XSLFTextRun pr = pp.addNewTextRun();
                pr.setText(String.valueOf(i + 1));
                pr.setFontSize(9.0);
                pr.setFontColor(new Color(148, 163, 184));
            }

            modifiedCount++;
        }

        log.info("Successfully applied headers/footers to {} slides in presentation '{}'", modifiedCount, resolvedId);
        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(SLIDES_UPDATED, modifiedCount);
        result.put(HEADER_TEXT, headerText);
        result.put(FOOTER_TEXT, footerText);
        result.put(MESSAGE, "Headers/footers applied to " + modifiedCount + " slides.");
        stateManager.autoSave(resolvedId);
        return result;
    }

    /**
     * find simple shape
     * @param slide mention XSLFSlide slide
     * @param shapeIndex mention shapeIndex
     * @param shapeName mention shapeName
     * @return XSLFSimpleShape
     */
    private XSLFSimpleShape findSimpleShape(XSLFSlide slide, Integer shapeIndex, String shapeName) {
        List<XSLFShape> shapes = slide.getShapes();
        if (shapeIndex != null && shapeIndex >= 0 && shapeIndex < shapes.size()) {
            XSLFShape s = shapes.get(shapeIndex);
            if (s instanceof XSLFSimpleShape) return (XSLFSimpleShape) s;
        }
        if (shapeName != null && !shapeName.trim().isEmpty()) {
            for (XSLFShape s : shapes) {
                if (s.getShapeName().equalsIgnoreCase(shapeName.trim()) && s instanceof XSLFSimpleShape) {
                    return (XSLFSimpleShape) s;
                }
            }
        }
        for (XSLFShape s : shapes) {
            if (s instanceof XSLFSimpleShape) return (XSLFSimpleShape) s;
        }
        return null;
    }
}
