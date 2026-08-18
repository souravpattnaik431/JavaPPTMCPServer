package org.spring.microservices.pptmcpserver.util;

import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;

/**
 * Utility helper for text operations and alignment parsing.
 */
public final class TextHelper {

    private TextHelper() {
        throw new UnsupportedOperationException("TextHelper is a utility class and cannot be instantiated.");
    }

    /**
     * Safely parse a text alignment string into Apache POI TextAlign enum.
     *
     * @param alignment Alignment string ('LEFT', 'CENTER', 'RIGHT', 'JUSTIFY').
     * @return TextAlign enum or null if invalid or null.
     */
    public static TextAlign parseTextAlign(String alignment) {
        if (alignment == null || alignment.trim().isEmpty()) {
            return null;
        }
        try {
            return TextAlign.valueOf(alignment.trim().toUpperCase());
        } catch (IllegalArgumentException | NullPointerException _) {
            // Ignore unrecognized alignment string and return null
            return null;
        }
    }

    /**
     * Removes empty/unfilled body or subtitle placeholders from a slide
     * to prevent PowerPoint from displaying ghost text like "Click to edit Master text styles"
     * and avoid schema-invalid zero-paragraph text bodies.
     *
     * @param slide Target slide.
     */
    public static void cleanupEmptyPlaceholders(org.apache.poi.xslf.usermodel.XSLFSlide slide) {
        if (slide == null) return;
        java.util.List<org.apache.poi.xslf.usermodel.XSLFShape> toRemove = new java.util.ArrayList<>();
        for (org.apache.poi.xslf.usermodel.XSLFShape s : slide.getShapes()) {
            if (s instanceof org.apache.poi.xslf.usermodel.XSLFTextShape ts) {
                String txt = ts.getText();
                if (txt != null && txt.toLowerCase().contains("click to edit")) {
                    toRemove.add(s);
                }
            }
        }
        for (org.apache.poi.xslf.usermodel.XSLFShape s : toRemove) {
            slide.removeShape(s);
        }
    }
}

