package org.spring.microservices.pptmcpserver.util;

import lombok.extern.slf4j.Slf4j;

import java.awt.Color;
import java.util.HashMap;
import java.util.Map;

/**
 * Utility for parsing and resolving color inputs (hex strings, RGB strings, named colors).
 */
@Slf4j
public final class ColorHelper {

    private static final Map<String, Color> COLOR_NAMES = new HashMap<>();

    static {
        COLOR_NAMES.put("black", Color.BLACK);
        COLOR_NAMES.put("white", Color.WHITE);
        COLOR_NAMES.put("red", Color.RED);
        COLOR_NAMES.put("green", new Color(34, 139, 34)); // Forest green
        COLOR_NAMES.put("blue", new Color(30, 144, 255)); // Dodger blue
        COLOR_NAMES.put("navy", new Color(0, 0, 128));
        COLOR_NAMES.put("yellow", Color.YELLOW);
        COLOR_NAMES.put("cyan", Color.CYAN);
        COLOR_NAMES.put("magenta", Color.MAGENTA);
        COLOR_NAMES.put("orange", Color.ORANGE);
        COLOR_NAMES.put("pink", Color.PINK);
        COLOR_NAMES.put("gray", Color.GRAY);
        COLOR_NAMES.put("grey", Color.GRAY);
        COLOR_NAMES.put("lightgray", Color.LIGHT_GRAY);
        COLOR_NAMES.put("lightgrey", Color.LIGHT_GRAY);
        COLOR_NAMES.put("darkgray", Color.DARK_GRAY);
        COLOR_NAMES.put("darkgrey", Color.DARK_GRAY);
        COLOR_NAMES.put("purple", new Color(128, 0, 128));
        COLOR_NAMES.put("teal", new Color(0, 128, 128));
        COLOR_NAMES.put("maroon", new Color(128, 0, 0));
        COLOR_NAMES.put("olive", new Color(128, 128, 0));
        COLOR_NAMES.put("coral", new Color(255, 127, 80));
        COLOR_NAMES.put("gold", new Color(255, 215, 0));
        COLOR_NAMES.put("silver", new Color(192, 192, 192));
        COLOR_NAMES.put("transparent", new Color(0, 0, 0, 0));
    }

    private ColorHelper() {
    }

    /**
     * Parse a color string into a java.awt.Color.
     * Supports "#RRGGBB", "#RGB", "0xRRGGBB", "rgb(r,g,b)", or named colors.
     *
     * @param colorStr     Color string representation.
     * @param defaultColor Fallback color if parsing fails.
     * @return Resolved Color instance.
     */
    public static Color parseColor(String colorStr, Color defaultColor) {
        if (colorStr == null || colorStr.trim().isEmpty()) {
            return defaultColor;
        }
        String s = colorStr.trim().toLowerCase();

        if (COLOR_NAMES.containsKey(s)) {
            return COLOR_NAMES.get(s);
        }

        try {
            if (s.startsWith("#")) {
                s = s.substring(1);
            } else if (s.startsWith("0x")) {
                s = s.substring(2);
            }

            if (s.length() == 3) {
                // Expand short hex like #fff -> #ffffff
                char r = s.charAt(0);
                char g = s.charAt(1);
                char b = s.charAt(2);
                s = "" + r + r + g + g + b + b;
            }

            if (s.length() == 6) {
                int rgb = Integer.parseInt(s, 16);
                return new Color(rgb);
            } else if (s.length() == 8) {
                int rgba = (int) Long.parseLong(s, 16);
                return new Color(rgba, true);
            }
        } catch (Exception e) {
            log.debug("Could not parse color '{}', falling back to default: {}", colorStr, e.getMessage());
        }

        return defaultColor;
    }

    /**
     * Parse a color string with no fallback (returns null if unparseable).
     *
     * @param colorStr Color string representation.
     * @return Resolved Color or null.
     */
    public static Color parseColor(String colorStr) {
        return parseColor(colorStr, null);
    }
}
