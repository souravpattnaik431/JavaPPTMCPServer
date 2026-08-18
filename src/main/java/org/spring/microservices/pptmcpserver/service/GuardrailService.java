package org.spring.microservices.pptmcpserver.service;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.sl.usermodel.Placeholder;
import org.apache.poi.xslf.usermodel.*;
import org.jspecify.annotations.NonNull;
import org.spring.microservices.pptmcpserver.state.PresentationStateManager;
import org.springframework.ai.mcp.annotation.McpTool;
import org.springframework.ai.mcp.annotation.McpToolParam;
import org.springframework.stereotype.Service;

import java.awt.Dimension;
import java.awt.geom.Rectangle2D;
import java.util.*;

import static org.spring.microservices.pptmcpserver.constants.Literals.*;

/**
 * Service 7: Quality Assurance, Validation & AI Guardrails (5 MCP Tools).
 * Provides automated presentation linting, visual boundary/overflow checks,
 * user-instruction custom rule compliance, auto-remediation, and design guidelines.
 */
@Slf4j
@Service
@RequiredArgsConstructor
public class GuardrailService {

    private final PresentationStateManager stateManager;

    private static final List<String> PLACEHOLDER_PATTERNS = List.of(
            "click to edit",
            "[insert title]",
            "[insert text]",
            "[todo]",
            "[tbd]",
            "lorem ipsum",
            "undefined",
            "null",
            "[slide title]",
            "[placeholder]"
    );

    /**
     * Tool 1: Comprehensive Presentation Quality Audit.
     * Checks all slides for boundary overflow, shape overlap, empty placeholders,
     * typography readability, missing titles, and contrast issues.
     *
     * @param presentationId Optional presentation ID.
     * @return Detailed audit report with score, letter grade, and categorized findings.
     */
    @McpTool(name = "audit_presentation_quality", description = "Perform a comprehensive automated quality and visual audit across all slides in the presentation. Detects overflow, collisions, empty placeholders, tiny fonts, and layout flaws.")
    public Map<String, Object> auditPresentationQuality(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Auditing presentation quality for ID: '{}'", resolvedId);
        XMLSlideShow ppt = stateManager.getPresentation(resolvedId);

        Dimension pageSize = ppt.getPageSize();
        double slideWidth = pageSize.getWidth();
        double slideHeight = pageSize.getHeight();
        List<XSLFSlide> slides = ppt.getSlides();

        int totalErrors = 0;
        int totalWarnings = 0;
        int totalInfos = 0;
        List<Map<String, Object>> slideAudits = new ArrayList<>();

        if (slides.isEmpty()) {
            Map<String, Object> result = new LinkedHashMap<>();
            result.put(STATUS, SUCCESS);
            result.put(PRESENTATION_ID, resolvedId);
            result.put(SLIDE_COUNT, 0);
            result.put(QUALITY_SCORE, 0);
            result.put(QUALITY_GRADE, "F");
            result.put(MESSAGE, "Presentation is empty (0 slides). Add slides to begin.");
            result.put(ERRORS, List.of("Presentation has no slides."));
            return result;
        }

        for (int i = 0; i < slides.size(); i++) {
            XSLFSlide slide = slides.get(i);
            SlideAuditResult slideAudit = auditSingleSlide(slide, i, slideWidth, slideHeight);
            totalErrors += slideAudit.errorCount;
            totalWarnings += slideAudit.warningCount;
            totalInfos += slideAudit.infoCount;

            Map<String, Object> slideMap = new LinkedHashMap<>();
            slideMap.put(SLIDE_INDEX, i);
            slideMap.put(SLIDE_NUMBER, i + 1);
            slideMap.put(TITLE, slideAudit.title);
            slideMap.put(SHAPE_COUNT, slide.getShapes().size());
            slideMap.put(QUALITY_SCORE, slideAudit.score);
            slideMap.put(ISSUES, slideAudit.issues);
            slideAudits.add(slideMap);
        }

        // Calculate overall quality score (weighted penalty deduction)
        int penalty = (totalErrors * 15) + (totalWarnings * 5) + (totalInfos);
        int overallScore = (int) Math.clamp(100L - (penalty / Math.max(1, slides.size())), 0L, 100L);
        String grade = computeGrade(overallScore);

        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(SLIDE_COUNT, slides.size());
        result.put(WIDTH, slideWidth);
        result.put(HEIGHT, slideHeight);
        result.put(QUALITY_SCORE, overallScore);
        result.put(QUALITY_GRADE, grade);
        result.put("totalErrors", totalErrors);
        result.put("totalWarnings", totalWarnings);
        result.put("totalInfos", totalInfos);
        result.put(SLIDE_AUDITS, slideAudits);
        result.put(MESSAGE, String.format("Presentation audit complete. Quality Score: %d/100 (Grade: %s) with %d error(s) and %d warning(s).",
                overallScore, grade, totalErrors, totalWarnings));

        return result;
    }

    /**
     * Tool 2: Single Slide Quality & Layout Check.
     * Evaluates a specific slide in real-time right after generation so the AI can correct issues immediately.
     *
     * @param presentationId Optional presentation ID.
     * @param slideIndex     0-based index of the slide to validate.
     * @return Map containing slide-specific score, bounding boxes, issues, and fix suggestions.
     */
    @McpTool(name = "validate_slide_quality", description = "Validate visual formatting, bounds overflow, font readability, and placeholder integrity for a single specific slide.")
    public Map<String, Object> validateSlideQuality(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "0-based slide index to validate") int slideIndex) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Validating single slide {} for presentation ID: '{}'", slideIndex, resolvedId);
        XMLSlideShow ppt = stateManager.getPresentation(resolvedId);
        XSLFSlide slide = stateManager.getSlide(resolvedId, slideIndex);

        Dimension pageSize = ppt.getPageSize();
        SlideAuditResult audit = auditSingleSlide(slide, slideIndex, pageSize.getWidth(), pageSize.getHeight());

        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(SLIDE_INDEX, slideIndex);
        result.put(SLIDE_NUMBER, slideIndex + 1);
        result.put(TITLE, audit.title);
        result.put(QUALITY_SCORE, audit.score);
        result.put(QUALITY_GRADE, computeGrade(audit.score));
        result.put(SHAPE_COUNT, slide.getShapes().size());
        result.put("errorCount", audit.errorCount);
        result.put("warningCount", audit.warningCount);
        result.put(ISSUES, audit.issues);
        result.put(MESSAGE, String.format("Slide %d validated: Score %d/100 (%s) with %d issue(s).",
                slideIndex + 1, audit.score, computeGrade(audit.score), audit.issues.size()));
        return result;
    }

    /**
     * Tool 3: User Instruction & Custom Rule Validator.
     * Verifies that the presentation satisfies specific user prompt requirements or corporate guidelines.
     *
     * @param presentationId       Optional presentation ID.
     * @param minSlides            Minimum required slide count.
     * @param maxSlides            Maximum allowed slide count.
     * @param expectedSlideCount   Exact expected slide count.
     * @param requiredSlideTitles  List of required slide title keywords (e.g., ['Agenda', 'Architecture', 'Conclusion']).
     * @param forbiddenKeywords    List of forbidden words or placeholder tokens (e.g., ['Draft', 'TBD', 'Lorem']).
     * @param minFontSize          Minimum allowable font size across all text.
     * @param maxBulletsPerSlide   Maximum bullet items allowed per slide.
     * @param maxWordsPerSlide     Maximum total words allowed on a single slide.
     * @param requireFooter        Whether corporate footer/slide numbers are mandatory.
     * @param requiredAspectRatio  Expected aspect ratio (e.g. '16:9' or '4:3').
     * @return Compliance report indicating pass/fail status and specific rule violations.
     */
    @McpTool(name = "validate_presentation_rules", description = "Validate the presentation against custom user instructions, guidelines, and design constraints (e.g., slide count, required sections, forbidden words, font size limits).")
    public Map<String, Object> validatePresentationRules(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "Minimum required slide count", required = false) Integer minSlides,
            @McpToolParam(description = "Maximum allowed slide count", required = false) Integer maxSlides,
            @McpToolParam(description = "Exact expected slide count", required = false) Integer expectedSlideCount,
            @McpToolParam(description = "List of required title keywords (e.g. ['Agenda', 'Architecture', 'Summary'])", required = false) List<String> requiredSlideTitles,
            @McpToolParam(description = "List of forbidden terms or placeholder words (e.g. ['Draft', 'TBD', 'Lorem ipsum'])", required = false) List<String> forbiddenKeywords,
            @McpToolParam(description = "Minimum allowable font size in points (e.g. 12.0 or 14.0)", required = false) Double minFontSize,
            @McpToolParam(description = "Maximum allowed bullet points per slide (e.g. 6)", required = false) Integer maxBulletsPerSlide,
            @McpToolParam(description = "Maximum allowed words per slide (e.g. 100)", required = false) Integer maxWordsPerSlide,
            @McpToolParam(description = "Require corporate footer or slide numbering on all content slides", required = false) Boolean requireFooter,
            @McpToolParam(description = "Expected aspect ratio ('16:9' or '4:3')", required = false) String requiredAspectRatio) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Validating custom rules for presentation ID: '{}'", resolvedId);
        XMLSlideShow ppt = stateManager.getPresentation(resolvedId);
        List<XSLFSlide> slides = ppt.getSlides();
        Dimension pageSize = ppt.getPageSize();

        List<Map<String, Object>> violations = new ArrayList<>();

        // 1. Slide Count Rules
        validateSlideCounts(slides.size(), expectedSlideCount, minSlides, maxSlides, violations);

        // 2. Aspect Ratio Check
        validateAspectRatio(pageSize, requiredAspectRatio, violations);

        // 3. Collect all slide titles and slide texts for content checks
        List<String> titlesFound = new ArrayList<>();
        validateSlidesContent(slides, pageSize, forbiddenKeywords, minFontSize, maxBulletsPerSlide, maxWordsPerSlide, requireFooter, violations, titlesFound);

        // 4. Check Required Slide Titles
        validateRequiredTitles(titlesFound, requiredSlideTitles, violations);

        boolean passed = violations.isEmpty();
        int totalRulesChecked = 5 + (requiredSlideTitles != null ? requiredSlideTitles.size() : 0) + (forbiddenKeywords != null ? forbiddenKeywords.size() : 0);
        int complianceScore = passed ? 100 : Math.max(0, (int) Math.round(100.0 * (totalRulesChecked - violations.size()) / totalRulesChecked));

        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(PASSED, passed);
        result.put(COMPLIANCE_PERCENTAGE, complianceScore);
        result.put("violationCount", violations.size());
        result.put(VIOLATIONS, violations);
        result.put(MESSAGE, passed
                ? "All user guidelines and custom rules passed successfully! (100% compliance)"
                : String.format("Rule validation completed with %d violation(s). Compliance: %d%%.", violations.size(), complianceScore));

        return result;
    }

    private void validateSlideCounts(int slideCount, Integer expectedSlideCount, Integer minSlides, Integer maxSlides, List<Map<String, Object>> violations) {
        if (expectedSlideCount != null && slideCount != expectedSlideCount) {
            violations.add(createViolation("SLIDE_COUNT_MISMATCH",
                    String.format("Expected exactly %d slide(s), but found %d.", expectedSlideCount, slideCount), PRESENTATION));
        }
        if (minSlides != null && slideCount < minSlides) {
            violations.add(createViolation("INSUFFICIENT_SLIDES",
                    String.format("Presentation has %d slide(s), which is less than minimum required (%d).", slideCount, minSlides), PRESENTATION));
        }
        if (maxSlides != null && slideCount > maxSlides) {
            violations.add(createViolation("EXCESS_SLIDES",
                    String.format("Presentation has %d slide(s), which exceeds maximum allowed (%d).", slideCount, maxSlides), PRESENTATION));
        }
    }

    private void validateAspectRatio(Dimension pageSize, String requiredAspectRatio, List<Map<String, Object>> violations) {
        if (requiredAspectRatio == null || requiredAspectRatio.trim().isEmpty()) {
            return;
        }
        String targetRatio = requiredAspectRatio.trim().replace(" ", "");
        double actualRatio = pageSize.getWidth() / pageSize.getHeight();
        if ("16:9".equalsIgnoreCase(targetRatio) && Math.abs(actualRatio - (16.0 / 9.0)) > 0.05) {
            violations.add(createViolation("ASPECT_RATIO_MISMATCH",
                    String.format("Expected 16:9 widescreen ratio, but slide dimensions are %.0fx%.0f (ratio %.2f:1).",
                            pageSize.getWidth(), pageSize.getHeight(), actualRatio), PRESENTATION));
        } else if ("4:3".equalsIgnoreCase(targetRatio) && Math.abs(actualRatio - (4.0 / 3.0)) > 0.05) {
            violations.add(createViolation("ASPECT_RATIO_MISMATCH",
                    String.format("Expected 4:3 standard ratio, but slide dimensions are %.0fx%.0f.",
                            pageSize.getWidth(), pageSize.getHeight()), PRESENTATION));
        }
    }

    private void validateSlidesContent(List<XSLFSlide> slides, Dimension pageSize, List<String> forbiddenKeywords, Double minFontSize,
                                       Integer maxBulletsPerSlide, Integer maxWordsPerSlide, Boolean requireFooter,
                                       List<Map<String, Object>> violations, List<String> titlesFound) {
        for (int i = 0; i < slides.size(); i++) {
            XSLFSlide slide = slides.get(i);
            String title = extractSlideTitle(slide);
            if (title != null && !title.trim().isEmpty()) {
                titlesFound.add(title.toLowerCase());
            }
            validateSingleSlideRules(slide, i, pageSize, forbiddenKeywords, minFontSize, maxBulletsPerSlide, maxWordsPerSlide, requireFooter, violations);
        }
    }

    private void validateSingleSlideRules(XSLFSlide slide, int slideIndex, Dimension pageSize, List<String> forbiddenKeywords,
                                          Double minFontSize, Integer maxBulletsPerSlide, Integer maxWordsPerSlide,
                                          Boolean requireFooter, List<Map<String, Object>> violations) {
        int bulletCount = 0;
        int wordCount = 0;
        boolean hasFooter = false;

        for (XSLFShape s : slide.getShapes()) {
            if (s instanceof XSLFTextShape ts) {
                wordCount += countWordsAndCheckKeywords(ts, slideIndex, forbiddenKeywords, violations);
                bulletCount += countBulletsAndCheckFonts(ts, slideIndex, minFontSize, violations);
                Rectangle2D anchor = ts.getAnchor();
                if (anchor != null && anchor.getY() >= pageSize.getHeight() - 60) {
                    hasFooter = true;
                }
            }
        }

        String slideScope = SLIDE_PREFIX + (slideIndex + 1);
        if (maxBulletsPerSlide != null && bulletCount > maxBulletsPerSlide) {
            violations.add(createViolation("TOO_MANY_BULLETS",
                    String.format("Slide %d has %d bullet points, exceeding limit of %d.", slideIndex + 1, bulletCount, maxBulletsPerSlide), slideScope));
        }
        if (maxWordsPerSlide != null && wordCount > maxWordsPerSlide) {
            violations.add(createViolation("TOO_MANY_WORDS",
                    String.format("Slide %d has %d words, exceeding word limit of %d.", slideIndex + 1, wordCount, maxWordsPerSlide), slideScope));
        }
        if (Boolean.TRUE.equals(requireFooter) && slideIndex > 0 && !hasFooter) {
            violations.add(createViolation("MISSING_FOOTER",
                    String.format("Slide %d is missing required corporate footer or slide numbering.", slideIndex + 1), slideScope));
        }
    }

    private int countWordsAndCheckKeywords(XSLFTextShape ts, int slideIndex, List<String> forbiddenKeywords, List<Map<String, Object>> violations) {
        String fullText = ts.getText();
        if (fullText == null || fullText.trim().isEmpty()) {
            return 0;
        }
        String[] words = fullText.trim().split("\\s+");
        if (forbiddenKeywords != null && !forbiddenKeywords.isEmpty()) {
            String lower = fullText.toLowerCase();
            for (String kw : forbiddenKeywords) {
                if (kw != null && !kw.trim().isEmpty() && lower.contains(kw.trim().toLowerCase())) {
                    violations.add(createViolation("FORBIDDEN_KEYWORD",
                            String.format("Slide %d contains forbidden keyword/phrase '%s'.", slideIndex + 1, kw), SLIDE_PREFIX + (slideIndex + 1)));
                }
            }
        }
        return words.length;
    }

    private int countBulletsAndCheckFonts(XSLFTextShape ts, int slideIndex, Double minFontSize, List<Map<String, Object>> violations) {
        int bulletCount = 0;
        for (XSLFTextParagraph p : ts.getTextParagraphs()) {
            if (p.isBullet() || (p.getText() != null && (p.getText().trim().startsWith("•") || p.getText().trim().startsWith("-")))) {
                bulletCount++;
            }
            if (minFontSize != null && minFontSize > 0) {
                checkParagraphFontSizes(p, slideIndex, minFontSize, violations);
            }
        }
        return bulletCount;
    }

    private void checkParagraphFontSizes(XSLFTextParagraph p, int slideIndex, double minFontSize, List<Map<String, Object>> violations) {
        for (XSLFTextRun r : p.getTextRuns()) {
            Double sz = r.getFontSize();
            if (sz != null && sz > 0 && sz < minFontSize && r.getRawText() != null && !r.getRawText().trim().isEmpty()) {
                violations.add(createViolation("FONT_SIZE_TOO_SMALL",
                        String.format("Slide %d has text with font size %.1fpt, below minimum required %.1fpt (text: '%s').",
                                slideIndex + 1, sz, minFontSize, truncate(r.getRawText(), 30)), SLIDE_PREFIX + (slideIndex + 1)));
            }
        }
    }

    private void validateRequiredTitles(List<String> titlesFound, List<String> requiredSlideTitles, List<Map<String, Object>> violations) {
        if (requiredSlideTitles == null || requiredSlideTitles.isEmpty()) {
            return;
        }
        for (String reqTitle : requiredSlideTitles) {
            if (reqTitle != null && !reqTitle.trim().isEmpty()) {
                String cleanReq = reqTitle.trim().toLowerCase();
                boolean found = titlesFound.stream().anyMatch(t -> t.contains(cleanReq));
                if (!found) {
                    violations.add(createViolation("MISSING_REQUIRED_SECTION",
                            String.format("Required slide section '%s' was not found in any slide title.", reqTitle), PRESENTATION));
                }
            }
        }
    }

    /**
     * Tool 4: Automated Presentation Issue Remediation (Auto-Fix).
     * Automatically fixes common AI mistakes:
     * - Clamps out-of-bounds coordinates back within slide margins.
     * - Deletes empty ghost placeholders.
     * - Bumps tiny fonts to minimum legible size (>= 12pt).
     * - Standardizes font families across all text.
     *
     * @param presentationId   Optional presentation ID.
     * @param targetFontFamily Optional font family to standardize on (e.g., 'Calibri', 'Arial', 'Inter').
     * @param minFontSize      Minimum font size to enforce (default: 12.0).
     * @param marginPadding    Safety margin from slide boundary in points (default: 20.0).
     * @return Map detailing the exact fixes applied across the presentation.
     */
    @McpTool(name = "autofix_presentation_issues", description = "Automatically repair common presentation defects: clamp out-of-bounds overflowing shapes back onto the slide, remove empty placeholders, and standardize font sizes/families.")
    public Map<String, Object> autofixPresentationIssues(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "Font family to standardize across all slides (e.g. 'Calibri', 'Arial', 'Segoe UI')", required = false) String targetFontFamily,
            @McpToolParam(description = "Minimum font size in points (default: 12.0)", required = false) Double minFontSize,
            @McpToolParam(description = "Safety edge margin in points (default: 20.0)", required = false) Double marginPadding) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Auto-fixing presentation defects for ID: '{}'", resolvedId);
        XMLSlideShow ppt = stateManager.getPresentation(resolvedId);

        Dimension pageSize = ppt.getPageSize();
        double slideWidth = pageSize.getWidth();
        double slideHeight = pageSize.getHeight();
        double margin = (marginPadding != null && marginPadding >= 0) ? marginPadding : 20.0;
        double minSize = (minFontSize != null && minFontSize > 0) ? minFontSize : 12.0;

        int[] fixCounts = new int[4];

        for (XSLFSlide slide : ppt.getSlides()) {
            autofixSingleSlide(slide, slideWidth, slideHeight, margin, minSize, targetFontFamily, fixCounts);
        }

        stateManager.autoSave(resolvedId);

        return getResult(fixCounts, resolvedId);
    }

    /**
     * get result
     * @param fixCounts mention fixCounts
     * @param resolvedId mention resolvedId
     * @return result
     */
    private static @NonNull Map<String, Object> getResult(int[] fixCounts, String resolvedId) {
        Map<String, Object> fixes = new LinkedHashMap<>();
        fixes.put("boundsClampedShapes", fixCounts[0]);
        fixes.put("placeholdersRemoved", fixCounts[1]);
        fixes.put("fontSizesAdjusted", fixCounts[2]);
        fixes.put("fontFamiliesStandardized", fixCounts[3]);

        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(FIXES_APPLIED, fixes);
        result.put(MESSAGE, String.format("Auto-fix complete: clamped %d out-of-bounds shape(s), removed %d ghost placeholder(s), adjusted %d undersized font(s).",
                fixCounts[0], fixCounts[1], fixCounts[2]));
        return result;
    }

    private void autofixSingleSlide(XSLFSlide slide, double slideWidth, double slideHeight, double margin, double minSize, String targetFontFamily, int[] fixCounts) {
        List<XSLFShape> toRemove = new ArrayList<>();
        for (XSLFShape s : slide.getShapes()) {
            if (isGhostPlaceholder(s)) {
                toRemove.add(s);
            } else {
                if (s instanceof XSLFTextShape ts) {
                    adjustTextShapeFonts(ts, minSize, targetFontFamily, fixCounts);
                }
                clampShapeBounds(s, s.getAnchor(), slideWidth, slideHeight, margin, fixCounts);
            }
        }
        for (XSLFShape s : toRemove) {
            slide.removeShape(s);
            fixCounts[1]++;
        }
    }

    private boolean isGhostPlaceholder(XSLFShape s) {
        if (s instanceof XSLFTextShape ts) {
            String raw = ts.getText();
            return raw == null || raw.trim().isEmpty() || (ts.getTextType() == Placeholder.BODY && raw.toLowerCase().contains("click to edit"));
        }
        return false;
    }

    private void adjustTextShapeFonts(XSLFTextShape ts, double minSize, String targetFontFamily, int[] fixCounts) {
        for (XSLFTextParagraph p : ts.getTextParagraphs()) {
            for (XSLFTextRun r : p.getTextRuns()) {
                Double sz = r.getFontSize();
                if (sz != null && sz > 0 && sz < minSize) {
                    r.setFontSize(minSize);
                    fixCounts[2]++;
                }
                if (targetFontFamily != null && !targetFontFamily.trim().isEmpty()) {
                    r.setFontFamily(targetFontFamily.trim());
                    fixCounts[3]++;
                }
            }
        }
    }

    private void clampShapeBounds(XSLFShape s, Rectangle2D anchor, double slideWidth, double slideHeight, double margin, int[] fixCounts) {
        if (anchor == null) {
            return;
        }
        double x = anchor.getX();
        double y = anchor.getY();
        double w = anchor.getWidth();
        double h = anchor.getHeight();

        if (w >= slideWidth * 0.95 && h >= slideHeight * 0.95) {
            return;
        }

        boolean modified = false;
        if (x < margin) {
            x = margin;
            modified = true;
        }
        if (y < margin) {
            y = margin;
            modified = true;
        }
        if (x + w > slideWidth - margin) {
            w = Math.min(w, slideWidth - 2 * margin);
            x = (w > slideWidth - 2 * margin) ? margin : slideWidth - margin - w;
            modified = true;
        }
        if (y + h > slideHeight - margin) {
            h = Math.min(h, slideHeight - 2 * margin);
            y = (h > slideHeight - 2 * margin) ? margin : slideHeight - margin - h;
            modified = true;
        }

        if (modified && s instanceof org.apache.poi.sl.usermodel.PlaceableShape<?, ?> ps) {
            ps.setAnchor(new Rectangle2D.Double(x, y, w, h));
            fixCounts[0]++;
        }
    }

    /**
     * Tool 5: Presentation Design Guidelines & Best Practices.
     * Returns structured coordinates, typography scales, contrast rules, and layout best practices.
     *
     * @return Map containing design principles and golden layout rules for AI models.
     */
    @McpTool(name = "get_guardrail_guidelines", description = "Retrieve expert PowerPoint design principles, coordinate grid systems, typography scales, and visual guardrails to ensure clean presentation generation.")
    public Map<String, Object> getGuardrailGuidelines() {
        log.info("Retrieving presentation design guardrail guidelines");

        Map<String, Object> guidelines = new LinkedHashMap<>();
        guidelines.put("canvasDimensions", Map.of(
                "standardWidescreen", "960 x 540 pt (16:9 ratio, default)",
                "coordinateOrigin", "Top-left is (x=0, y=0), bottom-right is (x=960, y=540)",
                "safeMargins", "Left/Right: 40pt, Top: 30pt, Bottom: 30pt (Safe usable area: x: 40-920, y: 30-510)"
        ));

        guidelines.put("layoutZones", Map.of(
                "titleZone", "x: 40, y: 30, width: 880, height: 60 (Title text: 24-32pt Bold, Subtitle: 14-16pt)",
                "mainContentZone", "x: 40, y: 100, width: 880, height: 380 (Usable area for cards, tables, split content)",
                "splitLeftColumn", "x: 40, y: 110, width: 420, height: 360",
                "splitRightColumn", "x: 490, y: 110, width: 430, height: 360",
                "threeCardGrid", "Card 1 (x: 40, w: 270), Card 2 (x: 345, w: 270), Card 3 (x: 650, w: 270), y: 120, h: 340",
                "footerZone", "x: 40, y: 495, width: 880, height: 30 (Text: 9-11pt, subtle gray #64748B)"
        ));

        guidelines.put("typographyScale", Map.of(
                "heroPresentationTitle", "36 - 44 pt (Bold)",
                "slideHeaderTitle", "24 - 30 pt (Bold)",
                "sectionSubtitle", "14 - 18 pt (Semi-Bold / Regular)",
                "cardHeader", "16 - 20 pt (Bold)",
                "bodyContent", "13 - 16 pt (Regular / Medium)",
                "captionOrFooter", "9 - 11 pt (Regular)",
                "minimumAllowedBodyFont", "12 pt (anything smaller is illegible on screen displays)"
        ));

        guidelines.put("colorAndContrastRules", List.of(
                "Dark Navy / Slate theme: Dark background (#0F172A), white title (#FFFFFF), slate body (#94A3B8), cyan/teal accent (#06B6D4 / #14B8A6)",
                "Clean Executive Light theme: Light background (#F8FAFC), dark navy headers (#1E3A8A / #0F172A), dark gray body (#334155), pure white card containers (#FFFFFF)",
                "NEVER use white text on white backgrounds or black text on black backgrounds",
                "Ensure minimum contrast ratio of 4.5:1 between text and background"
        ));

        guidelines.put("antiPatternsToAvoid", List.of(
                "Do NOT overlap text boxes or place body text directly on top of charts/tables",
                "Do NOT allow elements to exceed slide width (x + width > 960) or slide height (y + height > 540)",
                "Do NOT leave default placeholder strings like '[Insert Title]' or 'Click to edit'",
                "Do NOT exceed 6-7 bullet points or 100 words per slide (chunk into multi-card columns instead)",
                "Do NOT use raw bright red (#FF0000) or neon green for entire card backgrounds (use refined corporate hexes)"
        ));

        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put("guidelines", guidelines);
        result.put(MESSAGE, "Presentation guardrail guidelines and layout rules loaded successfully.");
        return result;
    }

    // ==========================================
    // Internal Helper Methods for Quality Audit
    // ==========================================

    private record SlideAuditResult(
            String title,
            int score,
            int errorCount,
            int warningCount,
            int infoCount,
            List<Map<String, Object>> issues
    ) {}

    private SlideAuditResult auditSingleSlide(XSLFSlide slide, int slideIndex, double slideWidth, double slideHeight) {
        List<Map<String, Object>> issues = new ArrayList<>();
        int[] counts = new int[3]; // [0: errorCount, 1: warningCount, 2: infoCount]

        String slideTitle = extractSlideTitle(slide);
        List<XSLFShape> shapes = slide.getShapes();

        // 1. Check for completely empty slide
        if (shapes.isEmpty()) {
            issues.add(createIssue(SEVERITY_ERROR, "EMPTY_SLIDE", "Slide contains no shapes or content.", "Add content, text boxes, or a layout template."));
            counts[0]++;
            return new SlideAuditResult(slideTitle, 0, counts[0], counts[1], counts[2], issues);
        }

        // 2. Check for missing slide title (except on title slide if large text exists)
        if ((slideTitle == null || slideTitle.trim().isEmpty()) && slideIndex > 0) {
            issues.add(createIssue(SEVERITY_WARNING, "MISSING_TITLE", "Slide does not have a recognizable title or header.", "Add a clear header text box at y <= 100pt."));
            counts[1]++;
        }

        // 3. Shape Inspection & Collision Loop
        List<Rectangle2D> candidateBounds = new ArrayList<>();
        auditShapes(shapes, slideWidth, slideHeight, candidateBounds, issues, counts);
        auditElementCollisions(candidateBounds, issues, counts);

        // 4. Visual Clutter / Density Check
        if (shapes.size() > 25) {
            issues.add(createIssue(SEVERITY_INFO, "HIGH_ELEMENT_DENSITY",
                    String.format("Slide has %d shapes, which may cause cognitive overload.", shapes.size()),
                    "Consider splitting content across multiple slides."));
            counts[2]++;
        }

        int penalty = (counts[0] * 25) + (counts[1] * 8) + (counts[2] * 2);
        int slideScore = (int) Math.clamp(100L - penalty, 0L, 100L);

        return new SlideAuditResult(slideTitle != null ? slideTitle : "Untitled Slide", slideScore, counts[0], counts[1], counts[2], issues);
    }

    private void auditShapes(List<XSLFShape> shapes, double slideWidth, double slideHeight, List<Rectangle2D> candidateBounds, List<Map<String, Object>> issues, int[] counts) {
        for (int sIdx = 0; sIdx < shapes.size(); sIdx++) {
            XSLFShape shape = shapes.get(sIdx);
            Rectangle2D anchor = shape.getAnchor();
            if (anchor == null) {
                continue;
            }

            double x = anchor.getX();
            double y = anchor.getY();
            double w = anchor.getWidth();
            double h = anchor.getHeight();
            boolean isFullBackground = (w >= slideWidth * 0.95 && h >= slideHeight * 0.95);

            if (!isFullBackground && (x < -2 || y < -2 || (x + w) > slideWidth + 4 || (y + h) > slideHeight + 4)) {
                issues.add(createIssue(SEVERITY_ERROR, "SLIDE_OVERFLOW",
                        String.format("Shape '%s' (bounds: x=%.0f, y=%.0f, w=%.0f, h=%.0f) overflows slide canvas (%.0fx%.0f).",
                                getShapeDisplayName(shape, sIdx), x, y, w, h, slideWidth, slideHeight),
                        "Clamp coordinates within x:[30, " + (slideWidth - 30) + "] and y:[30, " + (slideHeight - 30) + "]."));
                counts[0]++;
            }

            if (shape instanceof XSLFTextShape ts) {
                auditTextShape(ts, issues, counts);
            }

            if (!isFullBackground) {
                candidateBounds.add(anchor);
            }
        }
    }

    private void auditTextShape(XSLFTextShape ts, List<Map<String, Object>> issues, int[] counts) {
        String fullText = ts.getText();
        if (fullText != null) {
            String lower = fullText.toLowerCase();
            for (String pattern : PLACEHOLDER_PATTERNS) {
                if (lower.contains(pattern)) {
                    issues.add(createIssue(SEVERITY_ERROR, "UNRESOLVED_PLACEHOLDER",
                            String.format("Text shape contains unresolved placeholder text: '%s'.", truncate(fullText.trim(), 40)),
                            "Replace placeholder text with actual presentation content."));
                    counts[0]++;
                    break;
                }
            }
        }

        for (XSLFTextParagraph p : ts.getTextParagraphs()) {
            for (XSLFTextRun r : p.getTextRuns()) {
                Double sz = r.getFontSize();
                if (sz != null && sz > 0 && sz < 11.0 && r.getRawText() != null && !r.getRawText().trim().isEmpty()) {
                    issues.add(createIssue(SEVERITY_WARNING, "TINY_FONT_SIZE",
                            String.format("Font size %.1fpt in text '%s' is too small for standard display presentation.", sz, truncate(r.getRawText(), 30)),
                            "Increase font size to at least 13-16pt for body text or 10-11pt for footers."));
                    counts[1]++;
                    return;
                }
            }
        }
    }

    private void auditElementCollisions(List<Rectangle2D> candidateBounds, List<Map<String, Object>> issues, int[] counts) {
        for (int i = 0; i < candidateBounds.size(); i++) {
            Rectangle2D b1 = candidateBounds.get(i);
            for (int j = i + 1; j < candidateBounds.size(); j++) {
                Rectangle2D b2 = candidateBounds.get(j);
                checkPairCollision(b1, b2, issues, counts);
            }
        }
    }

    private void checkPairCollision(Rectangle2D b1, Rectangle2D b2, List<Map<String, Object>> issues, int[] counts) {
        if (!b1.intersects(b2)) {
            return;
        }
        Rectangle2D intersection = b1.createIntersection(b2);
        double overlapArea = intersection.getWidth() * intersection.getHeight();
        double area1 = b1.getWidth() * b1.getHeight();
        double area2 = b2.getWidth() * b2.getHeight();
        double minArea = Math.min(area1, area2);

        if (minArea > 0 && (overlapArea / minArea) > 0.25 && !isContainerRelationship(b1, b2, area1, area2)) {
            issues.add(createIssue(SEVERITY_WARNING, "ELEMENT_COLLISION",
                    String.format("Potential visual overlap detected between elements at (%.0f, %.0f, w:%.0f, h:%.0f) and (%.0f, %.0f, w:%.0f, h:%.0f).",
                            b1.getX(), b1.getY(), b1.getWidth(), b1.getHeight(),
                            b2.getX(), b2.getY(), b2.getWidth(), b2.getHeight()),
                    "Adjust element coordinates or use multi-card column spacing."));
            counts[1]++;
        }
    }

    private boolean isContainerRelationship(Rectangle2D b1, Rectangle2D b2, double area1, double area2) {
        return (area1 > area2 * 2.5 && b1.contains(b2)) || (area2 > area1 * 2.5 && b2.contains(b1));
    }

    private static String extractSlideTitle(XSLFSlide slide) {
        String title = slide.getTitle();
        if (title != null && !title.trim().isEmpty()) {
            return title.trim();
        }
        for (XSLFShape s : slide.getShapes()) {
            if (s instanceof XSLFTextShape ts) {
                String candidate = extractTitleCandidate(ts);
                if (candidate != null) {
                    return candidate;
                }
            }
        }
        return null;
    }

    private static String extractTitleCandidate(XSLFTextShape ts) {
        Placeholder pt = ts.getTextType();
        if (pt == Placeholder.TITLE || pt == Placeholder.CENTERED_TITLE) {
            String text = ts.getText();
            if (text != null && !text.trim().isEmpty()) {
                return text.trim();
            }
        }
        Rectangle2D anchor = ts.getAnchor();
        if (anchor != null && anchor.getY() <= 200 && anchor.getHeight() <= 120) {
            String text = ts.getText();
            if (text != null && !text.trim().isEmpty()) {
                return text.trim();
            }
        }
        return null;
    }

    private static String getShapeDisplayName(XSLFShape shape, int index) {
        String shapeName = shape.getShapeName();
        if (shapeName != null && !shapeName.trim().isEmpty()) {
            return shapeName;
        }
        return shape.getClass().getSimpleName() + " #" + (index + 1);
    }

    private static Map<String, Object> createIssue(String severity, String code, String description, String suggestion) {
        Map<String, Object> issue = new LinkedHashMap<>();
        issue.put("severity", severity);
        issue.put("code", code);
        issue.put("description", description);
        issue.put("suggestion", suggestion);
        return issue;
    }

    private static Map<String, Object> createViolation(String ruleCode, String description, String scope) {
        Map<String, Object> violation = new LinkedHashMap<>();
        violation.put("ruleCode", ruleCode);
        violation.put("description", description);
        violation.put("scope", scope);
        return violation;
    }
    private static String computeGrade(int score) {
        if (score >= 90) return "A";
        if (score >= 80) return "B";
        if (score >= 70) return "C";
        if (score >= 60) return "D";
        return "F";
    }

    private static String truncate(String s, int maxLen) {
        if (s == null) return "";
        String trimmed = s.replaceAll("\\s+", " ").trim();
        if (trimmed.length() <= maxLen) return trimmed;
        return trimmed.substring(0, Math.max(0, maxLen - 3)) + "...";
    }
}
