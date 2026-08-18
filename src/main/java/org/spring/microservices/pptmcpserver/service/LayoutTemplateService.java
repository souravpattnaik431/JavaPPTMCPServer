package org.spring.microservices.pptmcpserver.service;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.xslf.usermodel.*;
import org.spring.microservices.pptmcpserver.state.PresentationStateManager;
import org.spring.microservices.pptmcpserver.util.ColorHelper;
import org.spring.microservices.pptmcpserver.util.PictureHelper;
import org.springframework.ai.mcp.annotation.McpTool;
import org.springframework.ai.mcp.annotation.McpToolParam;
import org.springframework.stereotype.Service;

import javax.imageio.ImageIO;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.nio.file.Files;
import java.util.*;

import static org.spring.microservices.pptmcpserver.constants.Literals.*;

/**
 * Service 5: Executive Compound Slide Layouts & Business Templates (5 MCP Tools).
 * Provides high-level composite layouts tailored for executive pitches, business reviews, and time-constrained presentations.
 */
@Slf4j
@Service
@RequiredArgsConstructor
public class LayoutTemplateService {

    private final PresentationStateManager stateManager;

    /**
     * Tool 1: Add a Split-Screen Slide (Text on Left / Image on Right, or vice versa).
     * Ideal for executive business calls where attendees need to read key takeaways while visualizing diagrams/screenshots simultaneously.
     *
     * @param presentationId  Optional presentation ID.
     * @param title           Slide title.
     * @param subtitle        Optional subtitle or context banner.
     * @param bulletPoints    List of text bullet points or summary paragraphs.
     * @param imagePath       File path to the image/diagram to display.
     * @param imagePosition   Image placement: 'RIGHT' (default) or 'LEFT'.
     * @param imageCaption    Optional caption below the image.
     * @param primaryColor    Theme primary color hex.
     * @param backgroundColor Background color hex.
     * @return Map with slideIndex, layout, and status.
     */
    @McpTool(name = "add_split_content_slide", description = "Create a high-impact content slide with text takeaways/bullet points. ALWAYS call save_presentation after adding all slides to save the .pptx file in the user's active workspace.")
    public Map<String, Object> addSplitContentSlide(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "Slide title") String title,
            @McpToolParam(description = "Optional subtitle or context banner", required = false) String subtitle,
            @McpToolParam(description = "List of text bullet points or summary paragraphs") List<String> bulletPoints,
            @McpToolParam(description = "File path to the image/diagram to display (optional; if omitted, full-width card is used)", required = false) String imagePath,
            @McpToolParam(description = "Image placement: 'RIGHT' (default) or 'LEFT'", required = false) String imagePosition,
            @McpToolParam(description = "Optional caption below the image", required = false) String imageCaption,
            @McpToolParam(description = "Theme primary color hex (default: '#1E3A8A')", required = false) String primaryColor,
            @McpToolParam(description = "Background color hex (default: '#FFFFFF')", required = false) String backgroundColor) {

        Object[] created = stateManager.getOrCreatePresentation(presentationId);
        String resolvedId = (String) created[0];
        XMLSlideShow ppt = (XMLSlideShow) created[1];
        log.info("Adding split-content slide to presentation ID '{}': title='{}', imagePosition='{}'", resolvedId, title, imagePosition);
        XSLFSlide slide = createCleanSlide(ppt);
        int slideIndex = ppt.getSlides().size() - 1;

        Dimension pgSize = ppt.getPageSize();
        double totalW = pgSize.getWidth();
        double totalH = pgSize.getHeight();

        Color primary = ColorHelper.parseColor(primaryColor, new Color(30, 58, 138));
        Color bg = ColorHelper.parseColor(backgroundColor, Color.WHITE);
        slide.getBackground().setFillColor(bg);

        double titleH = (subtitle != null && !subtitle.trim().isEmpty()) ? 65.0 : 45.0;
        createSlideHeader(slide, title, subtitle, primary, totalW, 30, titleH);

        double contentY = 30 + titleH + 15;
        double contentH = totalH - contentY - 40;

        boolean hasImage = (imagePath != null && !imagePath.trim().isEmpty() && new File(imagePath.trim()).exists() && new File(imagePath.trim()).isFile());
        boolean imgOnRight = !"LEFT".equalsIgnoreCase(imagePosition);
        String imgStatus;

        if (hasImage) {
            double colW = (totalW - 130) / 2.0;
            double textX = imgOnRight ? 50 : 50 + colW + 30;
            double imgX = imgOnRight ? 50 + colW + 30 : 50;

            renderTextCardColumn(slide, textX, contentY, colW, contentH, bulletPoints);
            imgStatus = renderImageColumn(ppt, slide, imagePath, imageCaption, imgX, contentY, colW, contentH);
        } else {
            // Full-width card layout: allocates full widescreen width (860pt) so bullet points do not wrap unnecessarily
            double textX = 50;
            double colW = totalW - 100;
            renderTextCardColumn(slide, textX, contentY, colW, contentH, bulletPoints);
            imgStatus = NO_IMAGE;
        }

        log.info("Successfully created content slide at index {} with imageStatus='{}'", slideIndex, imgStatus);
        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(SLIDE_INDEX, slideIndex);
        result.put(SLIDE_NUMBER, slideIndex + 1);
        result.put(LAYOUT, hasImage ? "SPLIT_SCREEN" : "FULL_WIDTH_CARD");
        result.put(IMAGE_PLACEMENT, imgOnRight ? "RIGHT" : "LEFT");
        result.put(IMAGE_STATUS, imgStatus);
        result.put(MESSAGE, (hasImage ? "Split-screen executive slide" : "Full-width content card slide") + " created successfully at index " + slideIndex);
        if (title != null && !title.trim().isEmpty()) {
            stateManager.updateFilePathFromTitleIfDefault(resolvedId, title);
        }
        stateManager.autoSave(resolvedId);
        return result;
    }

    /**
     * Renders a text card column containing bullet points on the slide.
     *
     * @param slide        The target slide.
     * @param textX        The X position for the card.
     * @param contentY     The Y position for the card.
     * @param colW         The column width.
     * @param contentH     The column height.
     * @param bulletPoints The list of bullet point texts.
     */
    private void renderTextCardColumn(XSLFSlide slide, double textX, double contentY, double colW, double contentH, List<String> bulletPoints) {
        XSLFAutoShape textCard = slide.createAutoShape();
        textCard.setShapeType(ShapeType.ROUND_RECT);
        textCard.setAnchor(new Rectangle2D.Double(textX, contentY, colW, contentH));
        textCard.setFillColor(new Color(248, 250, 252));
        textCard.setLineColor(new Color(226, 232, 240));
        textCard.setLineWidth(1.0);

        XSLFTextBox textBox = slide.createTextBox();
        textBox.setAnchor(new Rectangle2D.Double(textX + 25, contentY + 25, colW - 50, contentH - 50));
        textBox.clearText();

        if (bulletPoints != null) {
            for (String bp : bulletPoints) {
                if (bp != null && !bp.trim().isEmpty()) {
                    XSLFTextParagraph p = textBox.addNewTextParagraph();
                    p.setIndentLevel(0);
                    p.setSpaceAfter(16.0);
                    p.setSpaceBefore(2.0);
                    XSLFTextRun r = p.addNewTextRun();
                    r.setText("•  " + bp.trim());
                    r.setFontSize(15.0);
                    r.setFontFamily(FONT_SEGOE_UI);
                    r.setFontColor(new Color(30, 41, 59));
                }
            }
        }
        if (textBox.getTextParagraphs().isEmpty()) {
            textBox.addNewTextParagraph().addNewTextRun().setText("");
        }
    }

    /**
     * Loads, scales, and places an image with an optional caption in the designated image column.
     *
     * @param ppt          The presentation instance.
     * @param slide        The target slide.
     * @param imagePath    The file path of the image.
     * @param imageCaption The optional caption text.
     * @param imgX         The X position for the image column.
     * @param contentY     The Y position for the image column.
     * @param colW         The column width.
     * @param contentH     The column height.
     * @return Status string indicating whether image was added, missing, or encountered an error.
     */
    private String renderImageColumn(XMLSlideShow ppt, XSLFSlide slide, String imagePath, String imageCaption, double imgX, double contentY, double colW, double contentH) {
        if (imagePath == null || imagePath.trim().isEmpty()) {
            return NO_IMAGE;
        }
        File imgFile = new File(imagePath.trim());
        if (!imgFile.exists() || !imgFile.isFile()) {
            return NO_IMAGE;
        }

        try {
            byte[] bytes = Files.readAllBytes(imgFile.toPath());
            PictureData.PictureType picType = PictureHelper.determinePictureType(imgFile.getName());
            XSLFPictureData pd = ppt.addPicture(bytes, picType);
            XSLFPictureShape pic = slide.createPicture(pd);

            double maxImgH = (imageCaption != null && !imageCaption.trim().isEmpty()) ? contentH - 40 : contentH;
            BufferedImage bImg = ImageIO.read(imgFile);
            double origW = (bImg != null) ? bImg.getWidth() : colW;
            double origH = (bImg != null) ? bImg.getHeight() : maxImgH;

            double scale = Math.min(colW / origW, maxImgH / origH);
            double fittedW = origW * scale;
            double fittedH = origH * scale;

            double fX = imgX + (colW - fittedW) / 2.0;
            double fY = contentY + (maxImgH - fittedH) / 2.0;
            pic.setAnchor(new Rectangle2D.Double(fX, fY, fittedW, fittedH));

            if (imageCaption != null && !imageCaption.trim().isEmpty()) {
                XSLFTextBox capBox = slide.createTextBox();
                capBox.setAnchor(new Rectangle2D.Double(imgX, contentY + maxImgH + 5, colW, 30));
                XSLFTextParagraph cp = capBox.addNewTextParagraph();
                cp.setTextAlign(TextAlign.CENTER);
                XSLFTextRun cr = cp.addNewTextRun();
                cr.setText(imageCaption.trim());
                cr.setFontSize(11.0);
                cr.setFontColor(new Color(100, 116, 139));
                cr.setItalic(true);
            }
            return "IMAGE_ADDED";
        } catch (Exception e) {
            log.warn("Failed to load/place image on split slide: {}", e.getMessage());
            return "ERROR_ADDING_IMAGE: " + e.getMessage();
        }
    }

    /**
     * Tool 2: Add a Metric / KPI Cards Slide.
     * Displays 2 to 4 prominent metric highlight cards with numbers, labels, and trend badges.
     *
     * @param presentationId  Optional presentation ID.
     * @param title           Slide title.
     * @param subtitle        Optional subtitle.
     * @param metrics         List of metric card definitions.
     * @param primaryColor    Primary theme color hex.
     * @param backgroundColor Background color hex.
     * @return Map with slideIndex, cardCount, and status.
     */
    @McpTool(name = "add_metric_cards_slide", description = "Add an executive KPI dashboard slide with metric cards. ALWAYS call save_presentation after adding all slides to save the .pptx file in the user's active workspace.")
    public Map<String, Object> addMetricCardsSlide(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "Slide title (e.g. 'Executive Performance Summary')") String title,
            @McpToolParam(description = "Optional subtitle", required = false) String subtitle,
            @McpToolParam(description = "List of metric card maps: [{'value': '$45.2M', 'label': 'Revenue', 'trend': '+24% YoY', 'trendPositive': true, 'detail': 'Exceeded Q3 plan'}]") List<Map<String, Object>> metrics,
            @McpToolParam(description = "Primary theme color hex (default: '#1E3A8A')", required = false) String primaryColor,
            @McpToolParam(description = "Background color hex (default: '#F8FAFC')", required = false) String backgroundColor) {

        if (metrics == null || metrics.isEmpty()) {
            throw new IllegalArgumentException("metrics list cannot be empty. Action Hint: Provide at least 1 metric card map (e.g. [{'value': '99.9%', 'label': 'Uptime', 'trend': '+0.5% YoY', 'trendPositive': true}]).");
        }

        Object[] created = stateManager.getOrCreatePresentation(presentationId);
        String resolvedId = (String) created[0];
        XMLSlideShow ppt = (XMLSlideShow) created[1];
        log.info("Adding KPI metric cards slide to presentation ID '{}': title='{}', count={}", resolvedId, title, metrics.size());
        XSLFSlide slide = createCleanSlide(ppt);
        int slideIndex = ppt.getSlides().size() - 1;

        Dimension pgSize = ppt.getPageSize();
        double totalW = pgSize.getWidth();

        Color primary = ColorHelper.parseColor(primaryColor, new Color(30, 58, 138));
        Color bg = ColorHelper.parseColor(backgroundColor, new Color(248, 250, 252));
        slide.getBackground().setFillColor(bg);

        double titleH = (subtitle != null && !subtitle.trim().isEmpty()) ? 65.0 : 45.0;
        createSlideHeader(slide, title, subtitle, primary, totalW, 35, titleH);

        int count = Math.min(metrics.size(), 4);
        renderMetricCards(slide, metrics, count, primary, totalW, 35 + titleH + 30);

        log.info("Successfully created metric cards slide at index {} with {} cards", slideIndex, count);
        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(SLIDE_INDEX, slideIndex);
        result.put(CARD_COUNT, count);
        result.put(MESSAGE, "Metric KPI cards slide created successfully at index " + slideIndex);
        if (title != null && !title.trim().isEmpty()) {
            stateManager.updateFilePathFromTitleIfDefault(resolvedId, title);
        }
        stateManager.autoSave(resolvedId);
        return result;
    }

    /**
     * Renders the grid of metric cards across the slide width.
     *
     * @param slide    The target slide.
     * @param metrics  The list of metric maps.
     * @param count    The number of cards to render.
     * @param primary  The theme primary color.
     * @param totalW   The total slide width.
     * @param startY   The Y position where cards start.
     */
    private void renderMetricCards(XSLFSlide slide, List<Map<String, Object>> metrics, int count, Color primary, double totalW, double startY) {
        double cardH = 260.0;
        double gap = 20.0;
        double availableW = totalW - 100 - (gap * (count - 1));
        double cardW = availableW / count;

        for (int i = 0; i < count; i++) {
            double cardX = 50 + i * (cardW + gap);
            renderSingleMetricCard(slide, metrics.get(i), primary, cardX, startY, cardW, cardH);
        }
    }

    /**
     * Renders an individual KPI metric card with value, label, trend badge, and detail description.
     *
     * @param slide   The target slide.
     * @param m       The map containing metric data.
     * @param primary The theme primary color.
     * @param cardX   The X position of the card.
     * @param startY  The Y position of the card.
     * @param cardW   The card width.
     * @param cardH   The card height.
     */
    private void renderSingleMetricCard(XSLFSlide slide, Map<String, Object> m, Color primary, double cardX, double startY, double cardW, double cardH) {
        String val = String.valueOf(m.getOrDefault("value", ""));
        String label = String.valueOf(m.getOrDefault("label", ""));
        String trend = String.valueOf(m.getOrDefault("trend", ""));
        boolean trendPos = Boolean.parseBoolean(String.valueOf(m.getOrDefault("trendPositive", "true")));
        String detail = String.valueOf(m.getOrDefault("detail", ""));

        XSLFAutoShape card = slide.createAutoShape();
        card.setShapeType(ShapeType.ROUND_RECT);
        card.setAnchor(new Rectangle2D.Double(cardX, startY, cardW, cardH));
        card.setFillColor(Color.WHITE);
        card.setLineColor(new Color(226, 232, 240));
        card.setLineWidth(1.5);

        XSLFTextBox cBox = slide.createTextBox();
        cBox.setAnchor(new Rectangle2D.Double(cardX + 15, startY + 20, cardW - 30, cardH - 40));
        cBox.clearText();

        XSLFTextParagraph lp = cBox.addNewTextParagraph();
        XSLFTextRun lr = lp.addNewTextRun();
        lr.setText(label.toUpperCase());
        lr.setFontSize(12.0);
        lr.setFontColor(new Color(100, 116, 139));
        lr.setBold(true);

        XSLFTextParagraph vp = cBox.addNewTextParagraph();
        vp.setSpaceBefore(10.0);
        vp.setSpaceAfter(8.0);
        XSLFTextRun vr = vp.addNewTextRun();
        vr.setText(val);
        vr.setFontSize(30.0);
        vr.setBold(true);
        vr.setFontColor(primary);

        if (trend != null && !trend.trim().isEmpty() && !"null".equalsIgnoreCase(trend)) {
            XSLFTextParagraph trP = cBox.addNewTextParagraph();
            trP.setSpaceAfter(10.0);
            XSLFTextRun trR = trP.addNewTextRun();
            trR.setText((trendPos ? "▲ " : "▼ ") + trend.trim());
            trR.setFontSize(13.0);
            trR.setBold(true);
            trR.setFontColor(trendPos ? new Color(22, 101, 52) : new Color(153, 27, 27));
        }

        if (detail != null && !detail.trim().isEmpty() && !"null".equalsIgnoreCase(detail)) {
            XSLFTextParagraph dp = cBox.addNewTextParagraph();
            XSLFTextRun dr = dp.addNewTextRun();
            dr.setText(detail.trim());
            dr.setFontSize(11.0);
            dr.setFontColor(new Color(71, 85, 105));
        }
    }

    /**
     * Tool 3: Add a Side-by-Side Comparison Slide (2 or 3 Columns).
     * Ideal for comparing architectures, vendor options, pros & cons, or before-and-after scenarios.
     *
     * @param presentationId Optional presentation ID.
     * @param title          Slide title.
     * @param subtitle       Optional subtitle.
     * @param columns        List of comparison column definitions.
     * @return Map with slideIndex, columnCount, and status.
     */
    @McpTool(name = "add_comparison_slide", description = "Create a multi-column comparison slide. ALWAYS call save_presentation after adding all slides to save the .pptx file in the user's active workspace.")
    public Map<String, Object> addComparisonSlide(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "Slide title (e.g. 'Solution Architecture Comparison')") String title,
            @McpToolParam(description = "Optional subtitle", required = false) String subtitle,
            @McpToolParam(description = "List of comparison column maps: [{'header': 'Option A', 'badge': 'Current', 'points': ['Point 1', 'Point 2'], 'color': '#1E3A8A'}]") List<Map<String, Object>> columns) {

        if (columns == null || columns.isEmpty()) {
            throw new IllegalArgumentException("columns list cannot be empty. Action Hint: Provide at least 2 comparison column maps (e.g. [{'header': 'Option A', 'badge': 'Recommended', 'points': ['Benefit 1', 'Benefit 2']}]).");
        }

        Object[] created = stateManager.getOrCreatePresentation(presentationId);
        String resolvedId = (String) created[0];
        XMLSlideShow ppt = (XMLSlideShow) created[1];
        log.info("Adding comparison slide to presentation ID '{}': title='{}', cols={}", resolvedId, title, columns.size());
        XSLFSlide slide = createCleanSlide(ppt);
        int slideIndex = ppt.getSlides().size() - 1;

        Dimension pgSize = ppt.getPageSize();
        double totalW = pgSize.getWidth();
        double totalH = pgSize.getHeight();

        slide.getBackground().setFillColor(new Color(248, 250, 252));

        double titleH = (subtitle != null && !subtitle.trim().isEmpty()) ? 65.0 : 45.0;
        createSlideHeader(slide, title, subtitle, new Color(15, 23, 42), totalW, 30, titleH);

        int colCount = Math.min(columns.size(), 3);
        double startY = 30 + titleH + 20;
        double colH = totalH - startY - 40;
        double gap = 20.0;
        double colW = (totalW - 100 - (gap * (colCount - 1))) / colCount;

        for (int i = 0; i < colCount; i++) {
            double colX = 50 + i * (colW + gap);
            renderSingleComparisonColumn(slide, columns.get(i), i, colX, startY, colW, colH);
        }

        log.info("Successfully created comparison slide at index {} with {} columns", slideIndex, colCount);
        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(SLIDE_INDEX, slideIndex);
        result.put(COLUMN_COUNT, colCount);
        result.put(MESSAGE, "Comparison slide created successfully at index " + slideIndex);
        if (title != null && !title.trim().isEmpty()) {
            stateManager.updateFilePathFromTitleIfDefault(resolvedId, title);
        }
        stateManager.autoSave(resolvedId);
        return result;
    }

    /**
     * Renders a single comparison column card with header banner, badge, and bullet points.
     *
     * @param slide   The target slide.
     * @param col     The comparison column configuration map.
     * @param index   The 0-based column index.
     * @param colX    The X position of the column.
     * @param startY  The Y position of the column.
     * @param colW    The column width.
     * @param colH    The column height.
     */
    private void renderSingleComparisonColumn(XSLFSlide slide, Map<String, Object> col, int index, double colX, double startY, double colW, double colH) {
        String header = String.valueOf(col.getOrDefault("header", "Column " + (index + 1)));
        String badge = (String) col.get("badge");
        String colColorHex = (String) col.get("color");
        Color colColor = ColorHelper.parseColor(colColorHex, (index == 0) ? new Color(30, 58, 138) : new Color(59, 130, 246));

        XSLFAutoShape baseCard = slide.createAutoShape();
        baseCard.setShapeType(ShapeType.ROUND_RECT);
        baseCard.setAnchor(new Rectangle2D.Double(colX, startY, colW, colH));
        baseCard.setFillColor(Color.WHITE);
        baseCard.setLineColor(new Color(226, 232, 240));
        baseCard.setLineWidth(1.5);

        XSLFAutoShape headBanner = slide.createAutoShape();
        headBanner.setShapeType(ShapeType.ROUND_RECT);
        headBanner.setAnchor(new Rectangle2D.Double(colX, startY, colW, 55));
        headBanner.setFillColor(colColor);
        headBanner.setLineColor(colColor);

        XSLFTextBox headText = slide.createTextBox();
        headText.setAnchor(new Rectangle2D.Double(colX + 10, startY + 12, colW - 20, 35));
        headText.clearText();
        XSLFTextParagraph hp = headText.addNewTextParagraph();
        hp.setTextAlign(TextAlign.CENTER);
        XSLFTextRun hr = hp.addNewTextRun();
        hr.setText(header);
        hr.setFontSize(16.0);
        hr.setBold(true);
        hr.setFontColor(Color.WHITE);

        XSLFTextBox bodyBox = slide.createTextBox();
        bodyBox.setAnchor(new Rectangle2D.Double(colX + 15, startY + 65, colW - 30, colH - 75));
        bodyBox.clearText();

        if (badge != null && !badge.trim().isEmpty()) {
            XSLFTextParagraph bp = bodyBox.addNewTextParagraph();
            bp.setSpaceAfter(8.0);
            XSLFTextRun br = bp.addNewTextRun();
            br.setText("● " + badge.trim().toUpperCase());
            br.setFontSize(11.0);
            br.setFontFamily(FONT_SEGOE_UI);
            br.setBold(true);
            br.setFontColor(colColor);
        }

        Object pointsObj = col.get("points");
        if (pointsObj instanceof List<?> points) {
            for (Object pt : points) {
                if (pt != null) {
                    XSLFTextParagraph p = bodyBox.addNewTextParagraph();
                    p.setSpaceAfter(8.0);
                    XSLFTextRun r = p.addNewTextRun();
                    r.setText("• " + pt);
                    r.setFontSize(13.0);
                    r.setFontColor(new Color(51, 65, 85));
                }
            }
        }
        if (bodyBox.getTextParagraphs().isEmpty()) {
            bodyBox.addNewTextParagraph().addNewTextRun().setText("");
        }
    }

    /**
     * Tool 4: Add a Horizontal Process / Roadmap Flow Slide.
     * Displays a sequential series of steps (Step 1 -> Step 2 -> Step 3 -> Step 4) with numbered badges and descriptions.
     *
     * @param presentationId Optional presentation ID.
     * @param title          Slide title.
     * @param subtitle       Optional subtitle.
     * @param steps          List of sequential step definitions.
     * @param primaryColor   Primary theme color hex.
     * @return Map with slideIndex, stepCount, and status.
     */
    @McpTool(name = "add_process_flow_slide", description = "Create a sequential roadmap/process flow slide. ALWAYS call save_presentation after adding all slides to save the .pptx file in the user's active workspace.")
    public Map<String, Object> addProcessFlowSlide(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "Slide title (e.g. 'Implementation Roadmap')") String title,
            @McpToolParam(description = "Optional subtitle", required = false) String subtitle,
            @McpToolParam(description = "List of step maps: [{'step': '1', 'title': 'Discovery', 'description': 'Assess requirements', 'duration': 'Week 1-2'}]") List<Map<String, String>> steps,
            @McpToolParam(description = "Primary theme color hex (default: '#1E3A8A')", required = false) String primaryColor) {

        if (steps == null || steps.isEmpty()) {
            throw new IllegalArgumentException("steps list cannot be empty. Action Hint: Provide at least 2 process step maps (e.g. [{'step': '1', 'title': 'Phase 1', 'description': 'Initial Setup'}]).");
        }

        Object[] created = stateManager.getOrCreatePresentation(presentationId);
        String resolvedId = (String) created[0];
        XMLSlideShow ppt = (XMLSlideShow) created[1];
        log.info("Adding process flow slide to presentation ID '{}': title='{}', steps={}", resolvedId, title, steps.size());
        XSLFSlide slide = createCleanSlide(ppt);
        int slideIndex = ppt.getSlides().size() - 1;

        Dimension pgSize = ppt.getPageSize();
        double totalW = pgSize.getWidth();

        Color primary = ColorHelper.parseColor(primaryColor, new Color(30, 58, 138));
        slide.getBackground().setFillColor(new Color(248, 250, 252));

        double titleH = (subtitle != null && !subtitle.trim().isEmpty()) ? 65.0 : 45.0;
        createSlideHeader(slide, title, subtitle, primary, totalW, 35, titleH);

        int stepCount = Math.min(steps.size(), 5);
        double startY = 35 + titleH + 40;
        double cardH = 220.0;
        double gap = 20.0;
        double stepW = (totalW - 100 - (gap * (stepCount - 1))) / stepCount;

        for (int i = 0; i < stepCount; i++) {
            double stepX = 50 + i * (stepW + gap);
            renderSingleProcessStep(slide, steps.get(i), i, primary, stepX, startY, stepW, cardH);
        }

        log.info("Successfully created process flow slide at index {} with {} steps", slideIndex, stepCount);
        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(SLIDE_INDEX, slideIndex);
        result.put(STEP_COUNT, stepCount);
        result.put(MESSAGE, "Process flow slide created successfully at index " + slideIndex);
        if (title != null && !title.trim().isEmpty()) {
            stateManager.updateFilePathFromTitleIfDefault(resolvedId, title);
        }
        stateManager.autoSave(resolvedId);
        return result;
    }

    /**
     * Renders a single process step card with a step number circle badge, title, duration, and description.
     *
     * @param slide   The target slide.
     * @param step    The step data map.
     * @param index   The 0-based step index.
     * @param primary The theme primary color.
     * @param stepX   The X position of the step card.
     * @param startY  The Y position of the step card.
     * @param stepW   The step card width.
     * @param cardH   The step card height.
     */
    private void renderSingleProcessStep(XSLFSlide slide, Map<String, String> step, int index, Color primary, double stepX, double startY, double stepW, double cardH) {
        String stepNum = step.getOrDefault("step", String.valueOf(index + 1));
        String stepTitle = step.getOrDefault("title", "Step " + stepNum);
        String desc = step.getOrDefault("description", "");
        String dur = step.getOrDefault("duration", "");

        XSLFAutoShape card = slide.createAutoShape();
        card.setShapeType(ShapeType.ROUND_RECT);
        card.setAnchor(new Rectangle2D.Double(stepX, startY, stepW, cardH));
        card.setFillColor(Color.WHITE);
        card.setLineColor(new Color(226, 232, 240));
        card.setLineWidth(1.5);

        XSLFAutoShape badge = slide.createAutoShape();
        badge.setShapeType(ShapeType.ELLIPSE);
        badge.setAnchor(new Rectangle2D.Double(stepX + 15, startY + 15, 36, 36));
        badge.setFillColor(primary);
        badge.setLineColor(primary);

        XSLFTextBox bText = slide.createTextBox();
        bText.setAnchor(new Rectangle2D.Double(stepX + 15, startY + 20, 36, 30));
        bText.clearText();
        XSLFTextParagraph bp = bText.addNewTextParagraph();
        bp.setTextAlign(TextAlign.CENTER);
        XSLFTextRun br = bp.addNewTextRun();
        br.setText(stepNum);
        br.setFontSize(14.0);
        br.setBold(true);
        br.setFontColor(Color.WHITE);

        XSLFTextBox contentBox = slide.createTextBox();
        contentBox.setAnchor(new Rectangle2D.Double(stepX + 15, startY + 60, stepW - 30, cardH - 75));
        contentBox.clearText();

        XSLFTextParagraph stP = contentBox.addNewTextParagraph();
        XSLFTextRun stR = stP.addNewTextRun();
        stR.setText(stepTitle);
        stR.setFontSize(15.0);
        stR.setBold(true);
        stR.setFontColor(new Color(15, 23, 42));

        if (dur != null && !dur.trim().isEmpty()) {
            XSLFTextParagraph durP = contentBox.addNewTextParagraph();
            durP.setSpaceAfter(6.0);
            XSLFTextRun durR = durP.addNewTextRun();
            durR.setText("Timeline: " + dur.trim());
            durR.setFontSize(11.0);
            durR.setFontFamily(FONT_SEGOE_UI);
            durR.setFontColor(new Color(100, 116, 139));
        }

        if (desc != null && !desc.trim().isEmpty()) {
            XSLFTextParagraph dP = contentBox.addNewTextParagraph();
            XSLFTextRun dR = dP.addNewTextRun();
            dR.setText(desc.trim());
            dR.setFontSize(12.0);
            dR.setFontColor(new Color(71, 85, 105));
        }
        if (contentBox.getTextParagraphs().isEmpty()) {
            contentBox.addNewTextParagraph().addNewTextRun().setText("");
        }
    }

    /**
     * Creates a standardized slide header with title and optional subtitle.
     *
     * @param slide    The target slide.
     * @param title    The main title text.
     * @param subtitle The optional subtitle text.
     * @param primary  The theme primary color for the title.
     * @param totalW   The total slide width.
     * @param startY   The Y position for the header box.
     * @param titleH   The height allocated for the header box.
     */
    private void createSlideHeader(XSLFSlide slide, String title, String subtitle, Color primary, double totalW, double startY, double titleH) {
        XSLFTextBox titleBox = slide.createTextBox();
        titleBox.setAnchor(new Rectangle2D.Double(50, startY, totalW - 100, titleH));
        titleBox.clearText();

        XSLFTextParagraph tp = titleBox.addNewTextParagraph();
        XSLFTextRun tr = tp.addNewTextRun();
        tr.setText(title);
        tr.setFontSize(26.0);
        tr.setBold(true);
        tr.setFontColor(primary);

        if (subtitle != null && !subtitle.trim().isEmpty()) {
            XSLFTextParagraph sp = titleBox.addNewTextParagraph();
            XSLFTextRun sr = sp.addNewTextRun();
            sr.setText(subtitle.trim());
            sr.setFontSize(14.0);
            sr.setFontColor(new Color(100, 116, 139));
        }
    }

    /**
     * Tool 5: Duplicate an existing slide in the presentation.
     *
     * @param presentationId Optional presentation ID.
     * @param slideIndex     0-based slide index to duplicate.
     * @return Map with source and new slide indices.
     */
    @McpTool(name = "duplicate_slide", description = "Duplicate an existing slide and append the clone to the presentation.")
    public Map<String, Object> duplicateSlide(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "0-based slide index to duplicate") int slideIndex) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Duplicating slide {} in presentation ID '{}'", slideIndex, resolvedId);
        XMLSlideShow ppt = stateManager.getPresentation(resolvedId);
        XSLFSlide srcSlide = stateManager.getSlide(resolvedId, slideIndex);

        XSLFSlide newSlide = ppt.createSlide();
        newSlide.importContent(srcSlide);
        int newIndex = ppt.getSlides().size() - 1;

        log.info("Successfully duplicated slide {} to new slide index {}", slideIndex, newIndex);
        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put("sourceSlideIndex", slideIndex);
        result.put("newSlideIndex", newIndex);
        result.put(MESSAGE, String.format("Slide at index %d duplicated to new slide index %d", slideIndex, newIndex));
        stateManager.autoSave(resolvedId);
        return result;
    }

    /**
     * Creates a clean blank slide for custom composite layouts,
     * stripping any inherited layout placeholder text shapes to prevent ghost shapes.
     *
     * @param ppt Presentation instance.
     * @return Clean XSLFSlide instance.
     */
    private XSLFSlide createCleanSlide(XMLSlideShow ppt) {
        XSLFSlideLayout blankLayout = null;
        if (!ppt.getSlideMasters().isEmpty()) {
            for (XSLFSlideLayout l : ppt.getSlideMasters().getFirst().getSlideLayouts()) {
                if ("Blank".equalsIgnoreCase(l.getName()) || "11_Blank".equalsIgnoreCase(l.getName()) || l.getName().toLowerCase().contains("blank")) {
                    blankLayout = l;
                    break;
                }
            }
        }
        XSLFSlide slide = (blankLayout != null) ? ppt.createSlide(blankLayout) : ppt.createSlide();
        List<XSLFShape> toRemove = new ArrayList<>();
        for (XSLFShape s : slide.getShapes()) {
            if (s instanceof XSLFTextShape) {
                toRemove.add(s);
            }
        }
        for (XSLFShape s : toRemove) {
            slide.removeShape(s);
        }
        return slide;
    }
}
