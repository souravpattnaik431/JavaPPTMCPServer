package org.spring.microservices.pptmcpserver.service;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.sl.usermodel.Placeholder;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.xslf.usermodel.*;
import org.jspecify.annotations.NonNull;
import org.jspecify.annotations.Nullable;
import org.spring.microservices.pptmcpserver.state.PresentationStateManager;
import org.spring.microservices.pptmcpserver.util.ColorHelper;
import org.spring.microservices.pptmcpserver.util.PictureHelper;
import org.spring.microservices.pptmcpserver.util.TextHelper;
import org.springframework.ai.mcp.annotation.McpTool;
import org.springframework.ai.mcp.annotation.McpToolParam;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.nio.file.Files;
import java.util.*;

import static org.spring.microservices.pptmcpserver.constants.Literals.*;

/**
 * Service 2: Slides, Text, Images, Hyperlinks & Speaker Notes (11 MCP Tools).
 * Manages slide lifecycle, typography, media placement, links, and presenter notes.
 */
@Slf4j
@Service
@RequiredArgsConstructor
public class SlideContentService {

    private final PresentationStateManager stateManager;

    /**
     * Tool 1: Add a new slide to the presentation.
     *
     * @param presentationId Optional presentation ID.
     * @param layoutIndex    Layout index from master.
     * @param layoutName     Layout name.
     * @param title          Optional slide title.
     * @param subtitle       Optional slide subtitle / body text.
     * @return Map containing slideIndex, layoutName, and status.
     */
    @McpTool(name = "add_slide", description = "Add a new slide to the presentation with an optional layout, title, and subtitle. ALWAYS call save_presentation after adding all slides to save the .pptx file in the user's active workspace.")
    public Map<String, Object> addSlide(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "Layout index from master (0: Title, 1: Title & Content, etc.)", required = false) Integer layoutIndex,
            @McpToolParam(description = "Layout name (e.g., 'Title Slide', 'Title and Content', 'Blank')", required = false) String layoutName,
            @McpToolParam(description = "Optional slide title", required = false) String title,
            @McpToolParam(description = "Optional slide subtitle / content text", required = false) String subtitle) {

        // Atomically get the existing presentation or create one shared instance.
        // This prevents concurrent add_slide calls from each creating their own separate presentation.
        // The auto-created presentation stays in memory only — no file is written until save_presentation is called.
        Object[] created = stateManager.getOrCreatePresentation(presentationId);
        String resolvedId = (String) created[0];
        XMLSlideShow ppt = (XMLSlideShow) created[1];
        log.info("Adding slide to presentation ID '{}': title='{}', layoutName='{}', layoutIndex={}", resolvedId, title, layoutName, layoutIndex);

        XSLFSlideLayout selectedLayout = getSelectedLayout(layoutIndex, layoutName, ppt);
        XSLFSlide slide = (selectedLayout != null) ? ppt.createSlide(selectedLayout) : ppt.createSlide();
        int slideIndex = ppt.getSlides().size() - 1;
        boolean isTitleSlide = (slideIndex == 0);

        applySlideTitle(slide, ppt, title, isTitleSlide);
        applySlideSubtitle(slide, ppt, subtitle, isTitleSlide);
        TextHelper.cleanupEmptyPlaceholders(slide);

        log.info("Successfully added slide at index {} to presentation '{}'", slideIndex, resolvedId);
        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(SLIDE_INDEX, slideIndex);
        result.put(SLIDE_NUMBER, slideIndex + 1);
        result.put(LAYOUT_NAME, (slide.getSlideLayout() != null) ? slide.getSlideLayout().getName() : "Default");
        result.put(TITLE, title);
        result.put(MESSAGE, "Slide added successfully at index " + slideIndex);
        if (title != null && !title.trim().isEmpty()) {
            stateManager.updateFilePathFromTitleIfDefault(resolvedId, title);
        }
        stateManager.autoSave(resolvedId);
        return result;
    }

    /**
     * Applies the slide title to an existing title placeholder or creates a new top title text box.
     * Sets bold styling, professional primary color (#1E3A8A), and responsive sizing.
     *
     * @param slide        The target slide.
     * @param ppt          The presentation instance.
     * @param title        The title text to set.
     * @param isTitleSlide Whether this is the first (introduction/title) slide.
     */
    private void applySlideTitle(XSLFSlide slide, XMLSlideShow ppt, String title, boolean isTitleSlide) {
        if (title == null || title.trim().isEmpty()) {
            return;
        }
        Color titleColor = new Color(30, 58, 138); // Bold primary navy blue #1E3A8A
        double fontSize = isTitleSlide ? 36.0 : 26.0;

        XSLFTextShape titleShape = findPlaceholderShape(slide, Placeholder.TITLE, Placeholder.CENTERED_TITLE);
        if (titleShape != null) {
            titleShape.clearText();
            XSLFTextParagraph p = titleShape.addNewTextParagraph();
            if (isTitleSlide) {
                p.setTextAlign(TextAlign.CENTER);
            }
            XSLFTextRun r = p.addNewTextRun();
            r.setText(title.trim());
            r.setFontSize(fontSize);
            r.setBold(true);
            r.setFontColor(titleColor);
            r.setFontFamily(FONT_SEGOE_UI);
        } else {
            XSLFTextBox tb = slide.createTextBox();
            double titleY = isTitleSlide ? 160.0 : 35.0;
            double titleH = isTitleSlide ? 90.0 : 55.0;
            tb.setAnchor(new Rectangle2D.Double(50, titleY, ppt.getPageSize().getWidth() - 100, titleH));
            tb.clearText();
            XSLFTextParagraph p = tb.addNewTextParagraph();
            if (isTitleSlide) {
                p.setTextAlign(TextAlign.CENTER);
            }
            XSLFTextRun r = p.addNewTextRun();
            r.setText(title.trim());
            r.setFontSize(fontSize);
            r.setBold(true);
            r.setFontColor(titleColor);
            r.setFontFamily(FONT_SEGOE_UI);
        }
    }

    /**
     * Applies the slide subtitle/content to an existing body/subtitle placeholder or creates a text box.
     * Sets clean slate-gray styling (#475569) and clears empty placeholders if no subtitle is given.
     *
     * @param slide        The target slide.
     * @param ppt          The presentation instance.
     * @param subtitle     The subtitle text to set.
     * @param isTitleSlide Whether this is the first (introduction/title) slide.
     */
    private void applySlideSubtitle(XSLFSlide slide, XMLSlideShow ppt, String subtitle, boolean isTitleSlide) {
        Color subColor = new Color(71, 85, 105); // Slate gray #475569
        double fontSize = isTitleSlide ? 18.0 : 14.0;

        XSLFTextShape subShape = findPlaceholderShape(slide, Placeholder.BODY, Placeholder.SUBTITLE);
        if (subShape != null) {
            if (subtitle != null && !subtitle.trim().isEmpty()) {
                populateSubtitleParagraphs(subShape, subtitle, isTitleSlide, fontSize, subColor);
            } else {
                slide.removeShape(subShape);
            }
        } else if (subtitle != null && !subtitle.trim().isEmpty()) {
            createSubtitleTextBox(slide, ppt, subtitle, isTitleSlide, fontSize, subColor);
        }
    }

    private void createSubtitleTextBox(XSLFSlide slide, XMLSlideShow ppt, String subtitle, boolean isTitleSlide, double fontSize, Color subColor) {
        double subY = isTitleSlide ? 270.0 : 95.0;
        double subH = isTitleSlide ? 80.0 : ppt.getPageSize().getHeight() - subY - 50.0;

        XSLFTextBox subTb = slide.createTextBox();
        subTb.setAnchor(new Rectangle2D.Double(50, subY, ppt.getPageSize().getWidth() - 100, subH));
        populateSubtitleParagraphs(subTb, subtitle, isTitleSlide, fontSize, subColor);
    }

    private void populateSubtitleParagraphs(XSLFTextShape shape, String subtitle, boolean isTitleSlide, double fontSize, Color subColor) {
        shape.clearText();
        String[] lines = subtitle.trim().split("\r?\n");
        for (String line : lines) {
            String trimmed = line.trim();
            if (trimmed.isEmpty()) {
                continue;
            }
            XSLFTextParagraph sp = shape.addNewTextParagraph();
            if (isTitleSlide) {
                sp.setTextAlign(TextAlign.CENTER);
            }
            if (lines.length > 1) {
                sp.setSpaceAfter(10.0);
                sp.setSpaceBefore(2.0);
            }
            XSLFTextRun sr = sp.addNewTextRun();
            sr.setText(trimmed);
            sr.setFontSize(fontSize);
            sr.setFontColor(subColor);
            sr.setFontFamily(FONT_SEGOE_UI);
        }
    }

    /**
     * Finds a text placeholder shape matching any of the specified placeholder types.
     *
     * @param slide The slide to search.
     * @param types The matching placeholder types.
     * @return The matched XSLFTextShape, or null if none found.
     */
    private XSLFTextShape findPlaceholderShape(XSLFSlide slide, Placeholder... types) {
        Set<Placeholder> set = Set.of(types);
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTextShape ts) {
                Placeholder pt = ts.getTextType() != null ? ts.getTextType() : ts.getPlaceholder();
                if (pt != null && set.contains(pt)) {
                    return ts;
                }
            }
        }
        return null;
    }

    /**
     * Resolves the requested slide layout from the master by name or index.
     *
     * @param layoutIndex 0-based layout index.
     * @param layoutName  Layout name string.
     * @param ppt         The presentation instance.
     * @return The resolved XSLFSlideLayout, or null if not found.
     */
    private static @Nullable XSLFSlideLayout getSelectedLayout(Integer layoutIndex, String layoutName, XMLSlideShow ppt) {
        List<XSLFSlideMaster> masters = ppt.getSlideMasters();
        if (masters == null || masters.isEmpty()) {
            return null;
        }
        XSLFSlideLayout[] layouts = masters.getFirst().getSlideLayouts();

        if (layoutName != null && !layoutName.trim().isEmpty()) {
            for (XSLFSlideLayout l : layouts) {
                if (l.getName().equalsIgnoreCase(layoutName.trim())) {
                    return l;
                }
            }
        }

        if (layoutIndex != null && layoutIndex >= 0 && layoutIndex < layouts.length) {
            return layouts[layoutIndex];
        }
        return null;
    }

    /**
     * Tool 2: Delete a slide by its 0-based index.
     *
     * @param presentationId Optional presentation ID.
     * @param slideIndex     0-based slide index to remove.
     * @return Map with deletedIndex, remainingSlideCount, and status.
     */
    @McpTool(name = "delete_slide", description = "Delete a slide from the presentation by its 0-based slide index.")
    public Map<String, Object> deleteSlide(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "0-based slide index to remove") int slideIndex) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Deleting slide at index {} from presentation ID '{}'", slideIndex, resolvedId);
        XMLSlideShow ppt = stateManager.getPresentation(resolvedId);

        if (slideIndex < 0 || slideIndex >= ppt.getSlides().size()) {
            throw new IllegalArgumentException(String.format(
                    "Slide index %d out of bounds (total slides: %d). Action Hint: Slide indices are 0-based (0 to %d).",
                    slideIndex, ppt.getSlides().size(), Math.max(0, ppt.getSlides().size() - 1)));
        }

        ppt.removeSlide(slideIndex);
        log.info("Successfully deleted slide {}. Remaining slides: {}", slideIndex, ppt.getSlides().size());

        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put("deletedIndex", slideIndex);
        result.put("remainingSlideCount", ppt.getSlides().size());
        result.put(MESSAGE, "Successfully deleted slide " + slideIndex);
        stateManager.autoSave(resolvedId);
        return result;
    }

    /**
     * Tool 3: Reorder a slide from one position to another.
     *
     * @param presentationId Optional presentation ID.
     * @param slideIndex     Current 0-based index.
     * @param newIndex       Target 0-based index.
     * @return Map confirming reorder.
     */
    @McpTool(name = "reorder_slide", description = "Move a slide from one position to another 0-based index.")
    public Map<String, Object> reorderSlide(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "Current 0-based index of the slide") int slideIndex,
            @McpToolParam(description = "Target 0-based index to move the slide to") int newIndex) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Reordering slide in presentation ID '{}': {} -> {}", resolvedId, slideIndex, newIndex);
        XMLSlideShow ppt = stateManager.getPresentation(resolvedId);

        XSLFSlide slide = getSlide(slideIndex, newIndex, ppt);
        ppt.setSlideOrder(slide, newIndex);
        log.info("Successfully moved slide from {} to {}", slideIndex, newIndex);

        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put("oldIndex", slideIndex);
        result.put("newIndex", newIndex);
        result.put(MESSAGE, "Slide reordered successfully.");
        stateManager.autoSave(resolvedId);
        return result;
    }

    /**
     * get Slide
     * @param slideIndex mention slide index
     * @param newIndex mention new index
     * @param ppt mention XMLSlideShow ppt object
     * @return slide
     */
    private static XSLFSlide getSlide(int slideIndex, int newIndex, XMLSlideShow ppt) {
        List<XSLFSlide> slides = ppt.getSlides();
        if (slideIndex < 0 || slideIndex >= slides.size()) {
            throw new IllegalArgumentException(String.format(
                    "slideIndex %d out of bounds (total slides: %d). Action Hint: Slide index must be between 0 and %d.",
                    slideIndex, slides.size(), Math.max(0, slides.size() - 1)));
        }
        if (newIndex < 0 || newIndex >= slides.size()) {
            throw new IllegalArgumentException(String.format(
                    "newIndex %d out of bounds (total slides: %d). Action Hint: Destination index must be between 0 and %d.",
                    newIndex, slides.size(), Math.max(0, slides.size() - 1)));
        }

        return slides.get(slideIndex);
    }

    /**
     * Tool 4: Get detailed info about a specific slide.
     *
     * @param presentationId Optional presentation ID.
     * @param slideIndex     0-based slide index.
     * @return Map with shapes, placeholders, positions, and layout details.
     */
    @McpTool(name = "get_slide_info", description = "Retrieve detailed information about a slide: shapes, placeholders, positions, and layout.")
    public Map<String, Object> getSlideInfo(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "0-based slide index") int slideIndex) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Retrieving slide info for slide {} in presentation ID '{}'", slideIndex, resolvedId);
        XSLFSlide slide = stateManager.getSlide(resolvedId, slideIndex);

        List<Map<String, Object>> shapeList = new ArrayList<>();
        List<XSLFShape> shapes = slide.getShapes();
        for (int i = 0; i < shapes.size(); i++) {
            XSLFShape s = shapes.get(i);
            Map<String, Object> sMap = createSMap(i, s);
            shapeList.add(sMap);
        }

        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(SLIDE_INDEX, slideIndex);
        result.put(LAYOUT_NAME, (slide.getSlideLayout() != null) ? slide.getSlideLayout().getName() : "Unknown");
        result.put(SHAPE_COUNT, shapes.size());
        result.put(SHAPES, shapeList);
        return result;
    }

    /**
     * Constructs a structured metadata map for a single shape on a slide,
     * capturing its index, name, class type, bounding box anchor coordinates,
     * and type-specific properties (text content/placeholder for text shapes,
     * row/column counts for tables, or picture format type for picture shapes).
     *
     * @param i 0-based shape index on the slide.
     * @param s The XSLFShape instance to inspect.
     * @return Map containing shape attributes and dimensions.
     */
    private static @NonNull Map<String, Object> createSMap(int i, XSLFShape s) {
        Map<String, Object> sMap = new LinkedHashMap<>();
        sMap.put("index", i);
        sMap.put(SHAPE_NAME, s.getShapeName());
        sMap.put(SHAPE_CLASS, s.getClass().getSimpleName());

        Rectangle2D anchor = s.getAnchor();
        if (anchor != null) {
            Map<String, Object> pos = new LinkedHashMap<>();
            pos.put(X, anchor.getX());
            pos.put(Y, anchor.getY());
            pos.put(WIDTH, anchor.getWidth());
            pos.put(HEIGHT, anchor.getHeight());
            sMap.put(ANCHOR, pos);
        }

        switch (s) {
            case XSLFTextShape ts -> {
                sMap.put(TEXT, ts.getText());
                sMap.put("isPlaceholder", ts.getTextType() != null);
                if (ts.getTextType() != null) {
                    sMap.put("placeholderType", ts.getTextType().name());
                }
            }
            case XSLFTable tbl -> {
                sMap.put("tableRows", tbl.getNumberOfRows());
                sMap.put("tableColumns", tbl.getNumberOfColumns());
            }
            case XSLFPictureShape pic -> sMap.put("pictureType", pic.getPictureData().getType().name());
            default -> {
                // Non-text, non-table, non-picture shapes (e.g. connectors, groups) have no additional properties
            }
        }
        return sMap;
    }

    /**
     * Tool 5: Extract all text content from a slide.
     *
     * @param presentationId Optional presentation ID.
     * @param slideIndex     0-based slide index.
     * @return Map with paragraph list and fullText string.
     */
    @McpTool(name = "extract_slide_text", description = "Extract all text content from shapes, text boxes, and tables on a specific slide.")
    public Map<String, Object> extractSlideText(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "0-based slide index") int slideIndex) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Extracting text from slide {} in presentation ID '{}'", slideIndex, resolvedId);
        XSLFSlide slide = stateManager.getSlide(resolvedId, slideIndex);

        List<String> paragraphs = new ArrayList<>();
        StringBuilder fullText = new StringBuilder();

        for (XSLFShape shape : slide.getShapes()) {
            extractShapeText(shape, paragraphs, fullText);
        }

        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(SLIDE_INDEX, slideIndex);
        result.put("paragraphCount", paragraphs.size());
        result.put("paragraphs", paragraphs);
        result.put("fullText", fullText.toString().trim());
        return result;
    }

    /**
     * Extracts text from an individual shape (text shape or table) and appends to paragraph list and fullText.
     *
     * @param shape      The shape to inspect.
     * @param paragraphs The collection of extracted paragraphs.
     * @param fullText   The aggregated string builder.
     */
    private void extractShapeText(XSLFShape shape, List<String> paragraphs, StringBuilder fullText) {
        if (shape instanceof XSLFTextShape ts) {
            extractTextShapeText(ts, paragraphs, fullText);
        } else if (shape instanceof XSLFTable table) {
            extractTableText(table, paragraphs, fullText);
        }
    }

    /**
     * Extracts paragraphs from a text shape.
     *
     * @param ts         The text shape.
     * @param paragraphs The collection of extracted paragraphs.
     * @param fullText   The aggregated string builder.
     */
    private void extractTextShapeText(XSLFTextShape ts, List<String> paragraphs, StringBuilder fullText) {
        for (XSLFTextParagraph p : ts.getTextParagraphs()) {
            String pText = p.getText();
            if (pText != null && !pText.trim().isEmpty()) {
                paragraphs.add(pText.trim());
                fullText.append(pText.trim()).append("\n");
            }
        }
    }

    /**
     * Extracts text cell-by-cell from a table shape.
     *
     * @param table      The table shape.
     * @param paragraphs The collection of extracted paragraphs.
     * @param fullText   The aggregated string builder.
     */
    private void extractTableText(XSLFTable table, List<String> paragraphs, StringBuilder fullText) {
        for (int r = 0; r < table.getNumberOfRows(); r++) {
            for (int c = 0; c < table.getNumberOfColumns(); c++) {
                XSLFTableCell cell = table.getCell(r, c);
                if (cell != null && cell.getText() != null && !cell.getText().trim().isEmpty()) {
                    paragraphs.add(cell.getText().trim());
                    fullText.append(cell.getText().trim()).append("\t");
                }
            }
            fullText.append("\n");
        }
    }

    /**
     * Tool 6: Add a formatted text box to a slide.
     *
     * @param presentationId Optional presentation ID.
     * @param slideIndex     0-based slide index.
     * @param text           Text content.
     * @param x              X coordinate in points.
     * @param y              Y coordinate in points.
     * @param width          Width in points.
     * @param height         Height in points.
     * @param fontSize       Font size in points.
     * @param fontColor      Font color hex/name.
     * @param fontFamily     Font family name.
     * @param bold           Bold flag.
     * @param italic         Italic flag.
     * @param underline      Underline flag.
     * @param alignment      Alignment: LEFT, CENTER, RIGHT, JUSTIFY.
     * @return Map with shapeName and status.
     */
    @McpTool(name = "add_textbox", description = "Add a customizable text box with position, size, font family, size, color, and alignment.")
    public Map<String, Object> addTextbox(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "0-based slide index") int slideIndex,
            @McpToolParam(description = "Text content to insert") String text,
            @McpToolParam(description = "X position in points (default: 50)", required = false) Double x,
            @McpToolParam(description = "Y position in points (default: 50)", required = false) Double y,
            @McpToolParam(description = "Width in points (default: 400)", required = false) Double width,
            @McpToolParam(description = "Height in points (default: 100)", required = false) Double height,
            @McpToolParam(description = "Font size in points (default: 18)", required = false) Double fontSize,
            @McpToolParam(description = "Font color hex or name (e.g., '#333333', 'blue')", required = false) String fontColor,
            @McpToolParam(description = "Font family (e.g., 'Calibri', 'Segoe UI', 'Arial')", required = false) String fontFamily,
            @McpToolParam(description = "Bold text flag", required = false) Boolean bold,
            @McpToolParam(description = "Italic text flag", required = false) Boolean italic,
            @McpToolParam(description = "Underline text flag", required = false) Boolean underline,
            @McpToolParam(description = "Text alignment: LEFT, CENTER, RIGHT, JUSTIFY", required = false) String alignment) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Adding textbox to slide {} in presentation ID '{}': '{}' at ({}, {})", slideIndex, resolvedId, text, x, y);
        XSLFSlide slide = stateManager.getSlide(resolvedId, slideIndex);
        TextHelper.cleanupEmptyPlaceholders(slide);

        double posX = (x != null) ? x : 50.0;
        double posY = (y != null) ? y : 50.0;
        double defaultW = Math.max(400.0, stateManager.getPresentation(resolvedId).getPageSize().getWidth() - posX - 50.0);
        double posW = (width != null && width > 0) ? width : defaultW;
        double posH = (height != null && height > 0) ? height : 100.0;

        XSLFTextBox targetBox = slide.createTextBox();

        targetBox.setAnchor(new Rectangle2D.Double(posX, posY, posW, posH));
        targetBox.clearText();

        Double effFontSize = (fontSize != null && fontSize > 0) ? fontSize : 15.0;
        String effFontFamily = (fontFamily != null && !fontFamily.trim().isEmpty()) ? fontFamily.trim() : FONT_SEGOE_UI;
        String effFontColor = (fontColor != null && !fontColor.trim().isEmpty()) ? fontColor.trim() : "#1E293B";
        TextAlign align = TextHelper.parseTextAlign(alignment);

        populateTextBoxParagraphs(targetBox, text, align, effFontSize, effFontColor, effFontFamily, bold, italic, underline);

        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(SLIDE_INDEX, slideIndex);
        result.put(SHAPE_NAME, targetBox.getShapeName());
        result.put(X, posX);
        result.put(Y, posY);
        result.put(WIDTH, posW);
        result.put(HEIGHT, posH);
        result.put(MESSAGE, "Text box added successfully.");
        stateManager.autoSave(resolvedId);
        return result;
    }

    private void populateTextBoxParagraphs(XSLFTextBox targetBox, String text, TextAlign align, Double effFontSize, String effFontColor, String effFontFamily, Boolean bold, Boolean italic, Boolean underline) {
        String rawText = (text != null) ? text : "";
        String[] lines = rawText.split("\r?\n");
        boolean multiParagraph = lines.length > 1;

        for (String line : lines) {
            String trimmed = line.trim();
            if (trimmed.isEmpty()) {
                continue;
            }
            XSLFTextParagraph p = targetBox.addNewTextParagraph();
            if (align != null) {
                p.setTextAlign(align);
            }
            if (multiParagraph) {
                p.setSpaceAfter(12.0);
                p.setSpaceBefore(2.0);
            }
            XSLFTextRun r = p.addNewTextRun();
            r.setText(trimmed);
            applyFontProperties(r, effFontSize, effFontColor, effFontFamily, bold, italic, underline);
        }

        if (targetBox.getTextParagraphs().isEmpty()) {
            XSLFTextParagraph p = targetBox.addNewTextParagraph();
            XSLFTextRun r = p.addNewTextRun();
            r.setText(rawText);
            applyFontProperties(r, effFontSize, effFontColor, effFontFamily, bold, italic, underline);
        }
    }

    /**
     * Applies styling properties (size, color, family, weight, style) to a text run.
     *
     * @param r          The text run to style.
     * @param fontSize   Font size in points.
     * @param fontColor  Font color hex/name.
     * @param fontFamily Font family name.
     * @param bold       Bold flag.
     * @param italic     Italic flag.
     * @param underline  Underline flag.
     */
    private void applyFontProperties(XSLFTextRun r, Double fontSize, String fontColor, String fontFamily, Boolean bold, Boolean italic, Boolean underline) {
        if (fontSize != null && fontSize > 0) {
            r.setFontSize(fontSize);
        }
        if (fontColor != null) {
            Color c = ColorHelper.parseColor(fontColor);
            if (c != null) r.setFontColor(c);
        }
        if (fontFamily != null && !fontFamily.trim().isEmpty()) {
            r.setFontFamily(fontFamily.trim());
        }
        if (bold != null) r.setBold(bold);
        if (italic != null) r.setItalic(italic);
        if (underline != null) r.setUnderlined(underline);
    }

    /**
     * Tool 7: Format text in an existing text shape.
     *
     * @param presentationId Optional presentation ID.
     * @param slideIndex     0-based slide index.
     * @param shapeIndex     0-based shape index.
     * @param shapeName      Optional shape name.
     * @param fontSize       Font size in points.
     * @param fontColor      Font color hex/name.
     * @param fontFamily     Font family name.
     * @param bold           Bold flag.
     * @param italic         Italic flag.
     * @param underline      Underline flag.
     * @param alignment      Alignment: LEFT, CENTER, RIGHT, JUSTIFY.
     * @return Map confirming formatting.
     */
    @McpTool(name = "format_text", description = "Format font properties of an existing text box or shape on a slide.")
    public Map<String, Object> formatText(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "0-based slide index") int slideIndex,
            @McpToolParam(description = "0-based shape index on the slide", required = false) Integer shapeIndex,
            @McpToolParam(description = "Shape name to search for (if shapeIndex not provided)", required = false) String shapeName,
            @McpToolParam(description = "Font size in points", required = false) Double fontSize,
            @McpToolParam(description = "Font color hex or name", required = false) String fontColor,
            @McpToolParam(description = "Font family name", required = false) String fontFamily,
            @McpToolParam(description = "Bold flag", required = false) Boolean bold,
            @McpToolParam(description = "Italic flag", required = false) Boolean italic,
            @McpToolParam(description = "Underline flag", required = false) Boolean underline,
            @McpToolParam(description = "Alignment: LEFT, CENTER, RIGHT, JUSTIFY", required = false) String alignment) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Formatting text on slide {} shapeIndex={} in presentation ID '{}'", slideIndex, shapeIndex, resolvedId);
        XSLFSlide slide = stateManager.getSlide(resolvedId, slideIndex);

        XSLFTextShape target = findTextShape(slide, shapeIndex, shapeName);
        if (target == null) {
            throw new IllegalArgumentException(String.format(
                    "No text shape found on slide %d matching criteria (shapeIndex=%s, shapeName='%s'). Action Hint: Call 'add_textbox' first to add text, or use 'inspect_slide_elements' to see existing shape indices.",
                    slideIndex, shapeIndex, shapeName));
        }

        TextAlign align = TextHelper.parseTextAlign(alignment);

        for (XSLFTextParagraph p : target.getTextParagraphs()) {
            if (align != null) {
                p.setTextAlign(align);
            }
            for (XSLFTextRun r : p.getTextRuns()) {
                applyFontProperties(r, fontSize, fontColor, fontFamily, bold, italic, underline);
            }
        }

        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(SLIDE_INDEX, slideIndex);
        result.put(SHAPE_NAME, target.getShapeName());
        result.put(MESSAGE, "Text formatting applied successfully.");
        stateManager.autoSave(resolvedId);
        return result;
    }

    /**
     * Tool 8: Add an image to a slide.
     *
     * @param presentationId Optional presentation ID.
     * @param slideIndex     0-based slide index.
     * @param imagePath      File path to the image.
     * @param x              X coordinate in points.
     * @param y              Y coordinate in points.
     * @param width          Width in points.
     * @param height         Height in points.
     * @return Map with shapeName and status.
     */
    @McpTool(name = "add_image", description = "Add an image (PNG, JPEG, GIF, BMP, SVG) to a slide from a local file path.")
    public Map<String, Object> addImage(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "0-based slide index") int slideIndex,
            @McpToolParam(description = "File path to the image") String imagePath,
            @McpToolParam(description = "X position in points (default: 50)", required = false) Double x,
            @McpToolParam(description = "Y position in points (default: 50)", required = false) Double y,
            @McpToolParam(description = "Width in points (default: 300)", required = false) Double width,
            @McpToolParam(description = "Height in points (default: 200)", required = false) Double height) {

        if (imagePath == null || imagePath.trim().isEmpty()) {
            throw new IllegalArgumentException("imagePath cannot be empty. Action Hint: Provide a valid file path to an image file (e.g. 'images/architecture.png').");
        }

        File imgFile = new File(imagePath.trim());
        if (!imgFile.exists() || !imgFile.isFile()) {
            throw new IllegalArgumentException(String.format(
                    "Image file does not exist at '%s'. Action Hint: Ensure the image file exists on disk before adding it to the slide.",
                    imgFile.getAbsolutePath()));
        }

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Adding image '{}' to slide {} in presentation ID '{}'", imgFile.getName(), slideIndex, resolvedId);
        XMLSlideShow ppt = stateManager.getPresentation(resolvedId);
        XSLFSlide slide = stateManager.getSlide(resolvedId, slideIndex);

        PictureData.PictureType picType = PictureHelper.determinePictureType(imgFile.getName());

        try {
            byte[] bytes = Files.readAllBytes(imgFile.toPath());
            XSLFPictureData pd = ppt.addPicture(bytes, picType);
            XSLFPictureShape picShape = slide.createPicture(pd);

            double posX = (x != null) ? x : 50.0;
            double posY = (y != null) ? y : 50.0;
            double posW = (width != null && width > 0) ? width : 300.0;
            double posH = (height != null && height > 0) ? height : 200.0;

            picShape.setAnchor(new Rectangle2D.Double(posX, posY, posW, posH));
            log.info("Successfully added picture shape '{}' ({}) to slide {}", picShape.getShapeName(), picType, slideIndex);

            Map<String, Object> result = new LinkedHashMap<>();
            result.put(STATUS, SUCCESS);
            result.put(PRESENTATION_ID, resolvedId);
            result.put(SLIDE_INDEX, slideIndex);
            result.put(SHAPE_NAME, picShape.getShapeName());
            result.put("pictureType", picType.name());
            result.put(X, posX);
            result.put(Y, posY);
            result.put(WIDTH, posW);
            result.put(HEIGHT, posH);
            result.put(MESSAGE, "Image added successfully.");
            stateManager.autoSave(resolvedId);
            return result;
        } catch (Exception e) {
            log.error("Failed to add image '{}' to slide {}: {}", imagePath, slideIndex, e.getMessage(), e);
            throw new RuntimeException("Failed to add image: " + e.getMessage(), e);
        }
    }

    /**
     * Tool 9: Add a hyperlink to text in a shape.
     *
     * @param presentationId Optional presentation ID.
     * @param slideIndex     0-based slide index.
     * @param shapeIndex     0-based shape index.
     * @param shapeName      Optional shape name.
     * @param url            Target web URL.
     * @param linkText       Optional target link text substring.
     * @return Map confirming hyperlink.
     */
    @McpTool(name = "add_hyperlink", description = "Add a clickable web URL hyperlink to text in an existing shape.")
    public Map<String, Object> addHyperlink(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "0-based slide index") int slideIndex,
            @McpToolParam(description = "0-based shape index on the slide", required = false) Integer shapeIndex,
            @McpToolParam(description = "Shape name", required = false) String shapeName,
            @McpToolParam(description = "Target URL (e.g. 'https://example.com')") String url,
            @McpToolParam(description = "Optional text to attach the link to. If omitted, applies to all text in the shape.", required = false) String linkText) {

        if (url == null || url.trim().isEmpty()) {
            throw new IllegalArgumentException("URL cannot be empty. Action Hint: Provide a valid URL starting with https:// or https://.");
        }

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Adding hyperlink '{}' to slide {} in presentation ID '{}'", url, slideIndex, resolvedId);
        XSLFSlide slide = stateManager.getSlide(resolvedId, slideIndex);

        XSLFTextShape target = findTextShape(slide, shapeIndex, shapeName);
        if (target == null) {
            throw new IllegalArgumentException(String.format(
                    "No text shape found on slide %d matching criteria (shapeIndex=%s, shapeName='%s'). Action Hint: Add a textbox first with 'add_textbox', or inspect slide shapes with 'inspect_slide_elements'.",
                    slideIndex, shapeIndex, shapeName));
        }

        boolean applied = false;
        for (XSLFTextParagraph p : target.getTextParagraphs()) {
            for (XSLFTextRun r : p.getTextRuns()) {
                if (linkText == null || linkText.trim().isEmpty() || (r.getRawText() != null && r.getRawText().contains(linkText.trim()))) {
                    XSLFHyperlink link = r.createHyperlink();
                    link.setAddress(url.trim());
                    applied = true;
                }
            }
        }

        if (!applied) {
            XSLFTextParagraph p = target.addNewTextParagraph();
            XSLFTextRun r = p.addNewTextRun();
            r.setText((linkText != null) ? linkText : url);
            XSLFHyperlink link = r.createHyperlink();
            link.setAddress(url.trim());
        }

        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(SLIDE_INDEX, slideIndex);
        result.put(URL, url.trim());
        result.put(MESSAGE, "Hyperlink attached successfully.");
        stateManager.autoSave(resolvedId);
        return result;
    }

    /**
     * Tool 10: Set presenter speaker notes on a slide.
     *
     * @param presentationId Optional presentation ID.
     * @param slideIndex     0-based slide index.
     * @param notesText      Speaker notes content.
     * @return Map confirming notes update.
     */
    @McpTool(name = "set_speaker_notes", description = "Add or update presenter speaker notes for a slide.")
    public Map<String, Object> setSpeakerNotes(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "0-based slide index") int slideIndex,
            @McpToolParam(description = "Speaker notes text for presenter") String notesText) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Setting speaker notes for slide {} in presentation ID '{}'", slideIndex, resolvedId);
        XMLSlideShow ppt = stateManager.getPresentation(resolvedId);
        XSLFSlide slide = stateManager.getSlide(resolvedId, slideIndex);

        XSLFNotes notes = ppt.getNotesSlide(slide);
        XSLFTextShape bodyShape = findSpeakerNotesBodyShape(notes);
        if (bodyShape != null) {
            bodyShape.setText(notesText != null ? notesText : "");
        }

        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(SLIDE_INDEX, slideIndex);
        result.put(NOTES, notesText);
        result.put(MESSAGE, "Speaker notes updated for slide " + slideIndex);
        stateManager.autoSave(resolvedId);
        return result;
    }

    /**
     * Tool 11: Get speaker notes from a slide.
     *
     * @param presentationId Optional presentation ID.
     * @param slideIndex     0-based slide index.
     * @return Map containing speaker notes text.
     */
    @McpTool(name = "get_speaker_notes", description = "Retrieve speaker notes attached to a slide.")
    public Map<String, Object> getSpeakerNotes(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "0-based slide index") int slideIndex) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Retrieving speaker notes for slide {} in presentation ID '{}'", slideIndex, resolvedId);
        XMLSlideShow ppt = stateManager.getPresentation(resolvedId);
        XSLFSlide slide = stateManager.getSlide(resolvedId, slideIndex);

        XSLFNotes notes = ppt.getNotesSlide(slide);
        XSLFTextShape bodyShape = findSpeakerNotesBodyShape(notes);
        String notesText = (bodyShape != null) ? bodyShape.getText() : "";

        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(SLIDE_INDEX, slideIndex);
        result.put(NOTES, notesText);
        return result;
    }

    /**
     * Finds the main body text shape on a speaker notes slide.
     *
     * @param notes The speaker notes slide.
     * @return The body text shape, or null if not found.
     */
    private XSLFTextShape findSpeakerNotesBodyShape(XSLFNotes notes) {
        if (notes == null) {
            return null;
        }
        for (XSLFShape s : notes.getShapes()) {
            if (s instanceof XSLFTextShape ts) {
                Placeholder ph = ts.getTextType() != null ? ts.getTextType() : ts.getPlaceholder();
                if (isBodyPlaceholder(ph)) {
                    return ts;
                }
            }
        }
        for (XSLFShape s : notes.getShapes()) {
            if (s instanceof XSLFTextShape ts) {
                return ts;
            }
        }
        return null;
    }

    /**
     * Determines whether a placeholder type corresponds to a body or content area.
     *
     * @param type The placeholder type.
     * @return True if body, none, or null.
     */
    private boolean isBodyPlaceholder(Placeholder type) {
        return type == Placeholder.BODY || type == Placeholder.NONE || type == null;
    }

    /**
     * Finds a text shape on a slide by 0-based index, shape name, or first available text shape.
     *
     * @param slide      The slide to search.
     * @param shapeIndex Optional 0-based shape index.
     * @param shapeName  Optional shape name.
     * @return The matched XSLFTextShape, or null if none found.
     */
    private XSLFTextShape findTextShape(XSLFSlide slide, Integer shapeIndex, String shapeName) {
        List<XSLFShape> shapes = slide.getShapes();
        if (shapeIndex != null && shapeIndex >= 0 && shapeIndex < shapes.size()) {
            XSLFShape s = shapes.get(shapeIndex);
            if (s instanceof XSLFTextShape) return (XSLFTextShape) s;
        }
        if (shapeName != null && !shapeName.trim().isEmpty()) {
            for (XSLFShape s : shapes) {
                if (s.getShapeName().equalsIgnoreCase(shapeName.trim()) && s instanceof XSLFTextShape) {
                    return (XSLFTextShape) s;
                }
            }
        }
        for (XSLFShape s : shapes) {
            if (s instanceof XSLFTextShape) return (XSLFTextShape) s;
        }
        return null;
    }
}
