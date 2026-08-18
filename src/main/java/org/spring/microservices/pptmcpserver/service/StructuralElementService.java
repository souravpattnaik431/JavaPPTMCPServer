package org.spring.microservices.pptmcpserver.service;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.xslf.usermodel.*;
import org.jspecify.annotations.NonNull;
import org.spring.microservices.pptmcpserver.state.PresentationStateManager;
import org.spring.microservices.pptmcpserver.util.ColorHelper;
import org.spring.microservices.pptmcpserver.util.TextHelper;
import org.springframework.ai.mcp.annotation.McpTool;
import org.springframework.ai.mcp.annotation.McpToolParam;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.geom.Rectangle2D;
import java.util.*;

import static org.spring.microservices.pptmcpserver.constants.Literals.*;

/**
 * Service 3: Tables, Shapes, Connectors, Charts & Slide Masters (7 MCP Tools).
 * Provides structural visual components, vector shapes, XDDF interactive charts, and slide layout inspections.
 */
@Slf4j
@Service
@RequiredArgsConstructor
public class StructuralElementService {

    private final PresentationStateManager stateManager;

    /**
     * Tool 1: Add a styled table to a slide.
     *
     * @param presentationId  Optional presentation ID.
     * @param slideIndex      0-based slide index.
     * @param rows            Number of rows.
     * @param cols            Number of columns.
     * @param x               X coordinate in points.
     * @param y               Y coordinate in points.
     * @param width           Table width in points.
     * @param height          Table height in points.
     * @param data            2D list of row cell values.
     * @param headers         List of header column titles.
     * @param headerBgColor   Header background color hex/name.
     * @param headerTextColor Header text color hex/name.
     * @param rowBgColor1     Alternating row 1 background color.
     * @param rowBgColor2     Alternating row 2 background color.
     * @return Map with table shape info and status.
     */
    @McpTool(name = "add_table", description = "Add a styled table to a slide with rows, columns, headers, and optional cell data.")
    public Map<String, Object> addTable(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "0-based slide index") int slideIndex,
            @McpToolParam(description = "Number of rows (including header if present)") int rows,
            @McpToolParam(description = "Number of columns") int cols,
            @McpToolParam(description = "X position in points (default: 50)", required = false) Double x,
            @McpToolParam(description = "Y position in points (default: 100)", required = false) Double y,
            @McpToolParam(description = "Table width in points (default: 860)", required = false) Double width,
            @McpToolParam(description = "Table height in points (default: 300)", required = false) Double height,
            @McpToolParam(description = "Optional 2D array of row data: [[cell_0_0, cell_0_1], ...]", required = false) List<List<String>> data,
            @McpToolParam(description = "Optional list of header column titles", required = false) List<String> headers,
            @McpToolParam(description = "Header background color hex/name (e.g. '#1E3A8A')", required = false) String headerBgColor,
            @McpToolParam(description = "Header text color hex/name (default: '#FFFFFF')", required = false) String headerTextColor,
            @McpToolParam(description = "Alternating row 1 background color (default: '#F8FAFC')", required = false) String rowBgColor1,
            @McpToolParam(description = "Alternating row 2 background color (default: '#FFFFFF')", required = false) String rowBgColor2) {

        int numRows = Math.max(1, rows);
        int numCols = Math.max(1, cols);

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Adding {}x{} table to slide {} in presentation ID '{}'", numRows, numCols, slideIndex, resolvedId);
        XSLFSlide slide = stateManager.getSlide(resolvedId, slideIndex);
        TextHelper.cleanupEmptyPlaceholders(slide);

        double posX = (x != null) ? x : 50.0;
        double posY = (y != null) ? y : 100.0;
        double posW = (width != null && width > 0) ? width : 860.0;
        double posH = (height != null && height > 0) ? height : 300.0;

        XSLFTable table = slide.createTable(numRows, numCols);
        table.setAnchor(new Rectangle2D.Double(posX, posY, posW, posH));

        double colWidth = posW / numCols;
        for (int c = 0; c < numCols; c++) {
            table.setColumnWidth(c, colWidth);
        }

        Color headBg = ColorHelper.parseColor(headerBgColor, new Color(30, 58, 138));
        Color headText = ColorHelper.parseColor(headerTextColor, Color.WHITE);
        Color rBg1 = ColorHelper.parseColor(rowBgColor1, new Color(248, 250, 252));
        Color rBg2 = ColorHelper.parseColor(rowBgColor2, Color.WHITE);

        int startDataRow = populateTableHeaders(table, headers, numCols, headBg, headText);
        populateTableData(table, data, startDataRow, numRows, numCols, rBg1, rBg2);

        log.info("Successfully added table '{}' ({} rows, {} cols) to slide {}", table.getShapeName(), numRows, numCols, slideIndex);
        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(SLIDE_INDEX, slideIndex);
        result.put(SHAPE_NAME, table.getShapeName());
        result.put("rows", numRows);
        result.put("columns", numCols);
        result.put(MESSAGE, "Table created successfully.");
        stateManager.autoSave(resolvedId);
        return result;
    }

    /**
     * Populates the header row of a table with column header titles and styles.
     *
     * @param table    The table to populate.
     * @param headers  The list of header strings.
     * @param numCols  The number of columns in the table.
     * @param headBg   The header background fill color.
     * @param headText The header text font color.
     * @return 1 if headers were added (so data starts on row 1), 0 otherwise.
     */
    private int populateTableHeaders(XSLFTable table, List<String> headers, int numCols, Color headBg, Color headText) {
        if (headers == null || headers.isEmpty()) {
            return 0;
        }
        int maxHeaderCols = Math.min(headers.size(), numCols);
        for (int c = 0; c < maxHeaderCols; c++) {
            formatHeaderCell(table.getCell(0, c), headers.get(c), headBg, headText);
        }
        return 1;
    }

    /**
     * Formats a single header cell with text, background color, center alignment, and bold white styling.
     *
     * @param cell     The table cell to format.
     * @param text     The header text.
     * @param headBg   The background fill color.
     * @param headText The text font color.
     */
    private void formatHeaderCell(XSLFTableCell cell, String text, Color headBg, Color headText) {
        if (cell == null) return;
        cell.setText(text);
        cell.setFillColor(headBg);
        for (XSLFTextParagraph p : cell.getTextParagraphs()) {
            p.setTextAlign(TextAlign.CENTER);
            for (XSLFTextRun r : p.getTextRuns()) {
                r.setFontColor(headText);
                r.setBold(true);
                r.setFontSize(14.0);
            }
        }
    }

    /**
     * Populates table body rows with data values and alternating row colors.
     *
     * @param table        The table to populate.
     * @param data         The 2D list of row cell values.
     * @param startDataRow The starting row index for data.
     * @param numRows      The total number of rows.
     * @param numCols      The total number of columns.
     * @param rBg1         Alternating background color 1.
     * @param rBg2         Alternating background color 2.
     */
    private void populateTableData(XSLFTable table, List<List<String>> data, int startDataRow, int numRows, int numCols, Color rBg1, Color rBg2) {
        if (data == null) {
            return;
        }
        for (int rIdx = 0; rIdx < data.size() && (rIdx + startDataRow) < numRows; rIdx++) {
            int actualRow = rIdx + startDataRow;
            Color rowBg = (rIdx % 2 == 0) ? rBg1 : rBg2;
            populateTableRow(table, data.get(rIdx), actualRow, numCols, rowBg);
        }
    }

    /**
     * Populates a single table data row across its columns.
     *
     * @param table     The target table.
     * @param rowData   The list of cell values for this row.
     * @param actualRow The row index in the table.
     * @param numCols   The number of columns.
     * @param rowBg     The background color for this row.
     */
    private void populateTableRow(XSLFTable table, List<String> rowData, int actualRow, int numCols, Color rowBg) {
        int maxCols = Math.min(rowData.size(), numCols);
        for (int cIdx = 0; cIdx < maxCols; cIdx++) {
            formatDataCell(table.getCell(actualRow, cIdx), rowData.get(cIdx), rowBg);
        }
    }

    /**
     * Formats a single data cell with text, background fill, and font styling.
     *
     * @param cell  The table cell to format.
     * @param text  The cell text content.
     * @param rowBg The cell background color.
     */
    private void formatDataCell(XSLFTableCell cell, String text, Color rowBg) {
        if (cell == null) return;
        cell.setText(text);
        cell.setFillColor(rowBg);
        for (XSLFTextParagraph p : cell.getTextParagraphs()) {
            for (XSLFTextRun r : p.getTextRuns()) {
                r.setFontSize(12.0);
                r.setFontColor(new Color(30, 41, 59));
            }
        }
    }

    /**
     * Tool 2: Format a specific table cell.
     *
     * @param presentationId Optional presentation ID.
     * @param slideIndex     0-based slide index.
     * @param shapeIndex     0-based table shape index.
     * @param row            0-based row index.
     * @param col            0-based column index.
     * @param text           Text content.
     * @param bgColor        Cell background color hex/name.
     * @param textColor      Text color hex/name.
     * @param fontSize       Font size in points.
     * @param fontFamily     Font family name.
     * @param bold           Bold flag.
     * @param italic         Italic flag.
     * @param alignment      Alignment: LEFT, CENTER, RIGHT.
     * @return Map confirming cell update.
     */
    @McpTool(name = "format_table_cell", description = "Format text, background color, font, or alignment for a specific table cell.")
    public Map<String, Object> formatTableCell(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "0-based slide index") int slideIndex,
            @McpToolParam(description = "0-based table shape index", required = false) Integer shapeIndex,
            @McpToolParam(description = "0-based row index") int row,
            @McpToolParam(description = "0-based column index") int col,
            @McpToolParam(description = "Text content to set", required = false) String text,
            @McpToolParam(description = "Cell background color hex/name", required = false) String bgColor,
            @McpToolParam(description = "Text color hex/name", required = false) String textColor,
            @McpToolParam(description = "Font size in points", required = false) Double fontSize,
            @McpToolParam(description = "Font family name", required = false) String fontFamily,
            @McpToolParam(description = "Bold flag", required = false) Boolean bold,
            @McpToolParam(description = "Italic flag", required = false) Boolean italic,
            @McpToolParam(description = "Text alignment: LEFT, CENTER, RIGHT", required = false) String alignment) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Formatting table cell [{}, {}] on slide {} in presentation ID '{}'", row, col, slideIndex, resolvedId);
        XSLFSlide slide = stateManager.getSlide(resolvedId, slideIndex);

        XSLFTable table = findTable(slide, shapeIndex);
        if (table == null) {
            throw new IllegalArgumentException(String.format(
                    "No table found on slide %d (shapeIndex=%s). Action Hint: Call 'add_table' first to create a table on this slide before formatting cells.",
                    slideIndex, shapeIndex));
        }

        if (row < 0 || row >= table.getNumberOfRows() || col < 0 || col >= table.getNumberOfColumns()) {
            throw new IllegalArgumentException(String.format(
                    "Cell [%d, %d] is out of table bounds [rows: %d, cols: %d] on slide %d. Action Hint: Row and column indices are 0-based (row: 0..%d, col: 0..%d).",
                    row, col, table.getNumberOfRows(), table.getNumberOfColumns(), slideIndex,
                    Math.max(0, table.getNumberOfRows() - 1), Math.max(0, table.getNumberOfColumns() - 1)));
        }

        XSLFTableCell cell = table.getCell(row, col);
        if (text != null) {
            cell.setText(text);
        }
        if (bgColor != null) {
            Color bg = ColorHelper.parseColor(bgColor);
            if (bg != null) cell.setFillColor(bg);
        }

        formatCellTextProperties(cell, textColor, fontSize, fontFamily, bold, italic, alignment);

        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(SLIDE_INDEX, slideIndex);
        result.put("row", row);
        result.put("column", col);
        result.put(MESSAGE, "Table cell updated successfully.");
        stateManager.autoSave(resolvedId);
        return result;
    }

    /**
     * Formats the text properties and alignment across all paragraphs in a table cell.
     *
     * @param cell       The table cell.
     * @param textColor  Text color hex/name.
     * @param fontSize   Font size in points.
     * @param fontFamily Font family name.
     * @param bold       Bold flag.
     * @param italic     Italic flag.
     * @param alignment  Alignment name string.
     */
    private void formatCellTextProperties(XSLFTableCell cell, String textColor, Double fontSize, String fontFamily, Boolean bold, Boolean italic, String alignment) {
        Color tColor = (textColor != null) ? ColorHelper.parseColor(textColor) : null;
        TextAlign align = TextHelper.parseTextAlign(alignment);

        for (XSLFTextParagraph p : cell.getTextParagraphs()) {
            if (align != null) {
                p.setTextAlign(align);
            }
            for (XSLFTextRun r : p.getTextRuns()) {
                formatParagraphTextRun(r, tColor, fontSize, fontFamily, bold, italic);
            }
        }
    }

    /**
     * Applies styling to a paragraph text run.
     *
     * @param r          The text run.
     * @param tColor     Text color.
     * @param fontSize   Font size.
     * @param fontFamily Font family name.
     * @param bold       Bold flag.
     * @param italic     Italic flag.
     */
    private void formatParagraphTextRun(XSLFTextRun r, Color tColor, Double fontSize, String fontFamily, Boolean bold, Boolean italic) {
        if (tColor != null) r.setFontColor(tColor);
        if (fontSize != null && fontSize > 0) r.setFontSize(fontSize);
        if (fontFamily != null && !fontFamily.trim().isEmpty()) r.setFontFamily(fontFamily.trim());
        if (bold != null) r.setBold(bold);
        if (italic != null) r.setItalic(italic);
    }

    /**
     * Tool 3: Add an AutoShape to a slide.
     *
     * @param presentationId Optional presentation ID.
     * @param slideIndex     0-based slide index.
     * @param shapeType      Shape type name (e.g. 'rectangle', 'round_rect', 'star', 'arrow').
     * @param x              X coordinate in points.
     * @param y              Y coordinate in points.
     * @param width          Width in points.
     * @param height         Height in points.
     * @param fillColor      Fill color hex/name.
     * @param lineColor      borderline color hex/name.
     * @param lineWidth      Border width in points.
     * @param text           Text inside shape.
     * @param textColor      Text color hex/name.
     * @param fontSize       Font size in points.
     * @param bold           Bold flag.
     * @return Map with shapeName and status.
     */
    @McpTool(name = "add_shape", description = "Add a geometric shape (rectangle, round_rect, ellipse, diamond, triangle, arrow, star, heart, etc.) with custom fill, border, and optional text.")
    public Map<String, Object> addShape(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "0-based slide index") int slideIndex,
            @McpToolParam(description = "Shape type name (e.g. 'rectangle', 'round_rect', 'ellipse', 'diamond', 'triangle', 'arrow', 'star', 'heart', 'cloud')") String shapeType,
            @McpToolParam(description = "X position in points (default: 100)", required = false) Double x,
            @McpToolParam(description = "Y position in points (default: 100)", required = false) Double y,
            @McpToolParam(description = "Width in points (default: 200)", required = false) Double width,
            @McpToolParam(description = "Height in points (default: 150)", required = false) Double height,
            @McpToolParam(description = "Fill color hex or name (e.g. '#2563EB', 'blue')", required = false) String fillColor,
            @McpToolParam(description = "Line border color hex or name", required = false) String lineColor,
            @McpToolParam(description = "Border line width in points (default: 1.0)", required = false) Double lineWidth,
            @McpToolParam(description = "Text inside shape", required = false) String text,
            @McpToolParam(description = "Text color hex or name", required = false) String textColor,
            @McpToolParam(description = "Font size in points", required = false) Double fontSize,
            @McpToolParam(description = "Bold flag for text inside shape", required = false) Boolean bold) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Adding shape '{}' to slide {} in presentation ID '{}' at ({}, {})", shapeType, slideIndex, resolvedId, x, y);
        XSLFSlide slide = stateManager.getSlide(resolvedId, slideIndex);
        TextHelper.cleanupEmptyPlaceholders(slide);

        ShapeType st = mapShapeType(shapeType);

        double posX = (x != null) ? x : 100.0;
        double posY = (y != null) ? y : 100.0;
        double posW = (width != null && width > 0) ? width : 200.0;
        double posH = (height != null && height > 0) ? height : 150.0;

        XSLFAutoShape shape = slide.createAutoShape();
        shape.setShapeType(st);
        shape.setAnchor(new Rectangle2D.Double(posX, posY, posW, posH));

        configureShapeStyle(shape, fillColor, lineColor, lineWidth);
        formatShapeText(shape, text, textColor, fontSize, bold);

        log.info("Successfully added shape '{}' ({}) to slide {}", shape.getShapeName(), st, slideIndex);
        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(SLIDE_INDEX, slideIndex);
        result.put(SHAPE_NAME, shape.getShapeName());
        result.put(SHAPE_TYPE, st.name());
        result.put(X, posX);
        result.put(Y, posY);
        result.put(WIDTH, posW);
        result.put(HEIGHT, posH);
        result.put(MESSAGE, "Shape added successfully.");
        stateManager.autoSave(resolvedId);
        return result;
    }

    /**
     * Configures fill color, line border color, and border thickness on an AutoShape.
     *
     * @param shape     The AutoShape to configure.
     * @param fillColor Fill color hex/name.
     * @param lineColor borderline color hex/name.
     * @param lineWidth borderline thickness in points.
     */
    private void configureShapeStyle(XSLFAutoShape shape, String fillColor, String lineColor, Double lineWidth) {
        if (fillColor != null) {
            Color c = ColorHelper.parseColor(fillColor);
            if (c != null) shape.setFillColor(c);
        }
        if (lineColor != null) {
            Color lc = ColorHelper.parseColor(lineColor);
            if (lc != null) shape.setLineColor(lc);
        }
        if (lineWidth != null) {
            shape.setLineWidth(lineWidth);
        }
    }

    /**
     * Adds and formats text content centered inside an AutoShape.
     *
     * @param shape     The AutoShape.
     * @param text      Text content.
     * @param textColor Text font color hex/name.
     * @param fontSize  Font size in points.
     * @param bold      Bold flag.
     */
    private void formatShapeText(XSLFAutoShape shape, String text, String textColor, Double fontSize, Boolean bold) {
        if (text == null || text.trim().isEmpty()) {
            return;
        }
        shape.setText(text.trim());
        Color tCol = (textColor != null) ? ColorHelper.parseColor(textColor) : Color.WHITE;
        for (XSLFTextParagraph p : shape.getTextParagraphs()) {
            p.setTextAlign(TextAlign.CENTER);
            for (XSLFTextRun r : p.getTextRuns()) {
                formatShapeTextRun(r, tCol, fontSize, bold);
            }
        }
    }

    /**
     * Styles an individual text run inside an AutoShape.
     *
     * @param r        The text run.
     * @param tCol     Font color.
     * @param fontSize Font size in points.
     * @param bold     Bold flag.
     */
    private void formatShapeTextRun(XSLFTextRun r, Color tCol, Double fontSize, Boolean bold) {
        if (tCol != null) r.setFontColor(tCol);
        if (fontSize != null && fontSize > 0) r.setFontSize(fontSize);
        if (bold != null) r.setBold(bold);
    }

    /**
     * Tool 4: Add a connector line between two points.
     *
     * @param presentationId Optional presentation ID.
     * @param slideIndex     0-based slide index.
     * @param connectorType  Connector style: STRAIGHT, ELBOW, CURVE.
     * @param startX         Start X coordinate.
     * @param startY         Start Y coordinate.
     * @param endX           End X coordinate.
     * @param endY           End Y coordinate.
     * @param lineColor      Line color hex/name.
     * @param lineWidth      Line thickness in points.
     * @return Map with connector info.
     */
    @McpTool(name = "add_connector", description = "Add a connector line between two points with custom color and line width.")
    public Map<String, Object> addConnector(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "0-based slide index") int slideIndex,
            @McpToolParam(description = "Connector style: STRAIGHT, ELBOW, CURVE", required = false) String connectorType,
            @McpToolParam(description = "Start X coordinate in points") double startX,
            @McpToolParam(description = "Start Y coordinate in points") double startY,
            @McpToolParam(description = "End X coordinate in points") double endX,
            @McpToolParam(description = "End Y coordinate in points") double endY,
            @McpToolParam(description = "Line color hex/name (default: '#475569')", required = false) String lineColor,
            @McpToolParam(description = "Line thickness in points (default: 2.0)", required = false) Double lineWidth) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Adding connector line on slide {} in presentation ID '{}' from ({},{}) to ({},{})", slideIndex, resolvedId, startX, startY, endX, endY);
        XSLFSlide slide = stateManager.getSlide(resolvedId, slideIndex);

        XSLFConnectorShape conn = slide.createConnector();
        double minX = Math.min(startX, endX);
        double minY = Math.min(startY, endY);
        double w = Math.max(Math.abs(endX - startX), 1.0);
        double h = Math.max(Math.abs(endY - startY), 1.0);

        conn.setAnchor(new Rectangle2D.Double(minX, minY, w, h));
        Color c = ColorHelper.parseColor(lineColor, new Color(71, 85, 105));
        conn.setLineColor(c);
        conn.setLineWidth((lineWidth != null && lineWidth > 0) ? lineWidth : 2.0);

        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(SLIDE_INDEX, slideIndex);
        result.put(SHAPE_NAME, conn.getShapeName());
        result.put(MESSAGE, "Connector line added successfully.");
        stateManager.autoSave(resolvedId);
        return result;
    }



    /**
     * Tool 7: List slide masters and layout placeholders.
     *
     * @param presentationId Optional presentation ID.
     * @return Map listing all master layouts and indices.
     */
    @McpTool(name = "get_slide_masters", description = "Inspect all slide masters, layouts, and their available placeholders in the presentation. If no presentation exists yet, one is automatically created (16:9 widescreen).")
    public Map<String, Object> getSlideMasters(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Inspecting slide masters for presentation ID '{}'", resolvedId);
        XMLSlideShow ppt;
        try {
            ppt = stateManager.getPresentation(resolvedId);
        } catch (IllegalStateException | IllegalArgumentException _) {
            log.warn("No presentation found for ID '{}' — auto-creating a blank 16:9 presentation.", resolvedId);
            ppt = new XMLSlideShow();
            ppt.setPageSize(new Dimension(960, 540));
            resolvedId = stateManager.store(ppt, null);
            stateManager.setFilePath(resolvedId, DEFAULT_PRESENTATION_FILE_NAME);
            stateManager.autoSave(resolvedId);
            log.info("Auto-created presentation with ID '{}'", resolvedId);
        }

        List<Map<String, Object>> mastersList = new ArrayList<>();
        for (int m = 0; m < ppt.getSlideMasters().size(); m++) {
            Map<String, Object> masterMap = getMasterMap(ppt, m);
            mastersList.add(masterMap);
        }

        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put("masterCount", ppt.getSlideMasters().size());
        result.put("masters", mastersList);
        return result;
    }

    /**
     * Generates a metadata map describing a slide master and all its associated layouts.
     *
     * @param ppt The presentation instance.
     * @param m   The 0-based slide master index.
     * @return Non-null Map of master index and layout list.
     */
    private static @NonNull Map<String, Object> getMasterMap(XMLSlideShow ppt, int m) {
        XSLFSlideMaster master = ppt.getSlideMasters().get(m);
        Map<String, Object> masterMap = new LinkedHashMap<>();
        masterMap.put("masterIndex", m);

        List<Map<String, Object>> layoutsList = new ArrayList<>();
        XSLFSlideLayout[] layouts = master.getSlideLayouts();
        for (int l = 0; l < layouts.length; l++) {
            XSLFSlideLayout layout = layouts[l];
            Map<String, Object> lMap = new LinkedHashMap<>();
            lMap.put("layoutIndex", l);
            lMap.put(LAYOUT_NAME, layout.getName());
            lMap.put("layoutType", (layout.getType() != null) ? layout.getType().name() : "CUSTOM");
            layoutsList.add(lMap);
        }
        masterMap.put("layouts", layoutsList);
        return masterMap;
    }

    /**
     * Finds a table shape on a slide by 0-based index or the first available table shape.
     *
     * @param slide      The slide to search.
     * @param shapeIndex Optional 0-based shape index.
     * @return The matched XSLFTable, or null if not found.
     */
    private XSLFTable findTable(XSLFSlide slide, Integer shapeIndex) {
        List<XSLFShape> shapes = slide.getShapes();
        if (shapeIndex != null && shapeIndex >= 0 && shapeIndex < shapes.size()) {
            XSLFShape s = shapes.get(shapeIndex);
            if (s instanceof XSLFTable) return (XSLFTable) s;
        }
        for (XSLFShape s : shapes) {
            if (s instanceof XSLFTable) return (XSLFTable) s;
        }
        return null;
    }



    /**
     * Maps a shape name string to an Apache POI ShapeType enum value.
     *
     * @param stName The shape type name string (e.g. 'rectangle', 'round_rect', 'ellipse', 'star', etc.).
     * @return The corresponding ShapeType, default is RECT.
     */
    private ShapeType mapShapeType(String stName) {
        if (stName == null) return ShapeType.RECT;
        String s = stName.trim().toLowerCase().replace("-", "_");
        return switch (s) {
            case "round_rect", "rounded_rectangle", "round_rectangle" -> ShapeType.ROUND_RECT;
            case "ellipse", "oval", "circle" -> ShapeType.ELLIPSE;
            case "diamond" -> ShapeType.DIAMOND;
            case "triangle", "isosceles_triangle" -> ShapeType.TRIANGLE;
            case "rt_triangle", "right_triangle" -> ShapeType.RT_TRIANGLE;
            case "pentagon" -> ShapeType.PENTAGON;
            case "hexagon" -> ShapeType.HEXAGON;
            case "octagon" -> ShapeType.OCTAGON;
            case "star", "star_5", "5_point_star" -> ShapeType.STAR_5;
            case "arrow", "right_arrow" -> ShapeType.RIGHT_ARROW;
            case "left_arrow" -> ShapeType.LEFT_ARROW;
            case "up_arrow" -> ShapeType.UP_ARROW;
            case "down_arrow" -> ShapeType.DOWN_ARROW;
            case "heart" -> ShapeType.HEART;
            case "lightning_bolt", "lightning" -> ShapeType.LIGHTNING_BOLT;
            case "cloud", "cloud_callout" -> ShapeType.CLOUD_CALLOUT;
            case "flow_process", "flowchart_process" -> ShapeType.FLOW_CHART_PROCESS;
            case "flow_decision", "flowchart_decision" -> ShapeType.FLOW_CHART_DECISION;
            default -> ShapeType.RECT;
        };
    }
}
