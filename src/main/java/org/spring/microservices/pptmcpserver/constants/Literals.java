package org.spring.microservices.pptmcpserver.constants;

/**
 * Common string literals and constants used across PPT MCP services and models.
 * Centralized to eliminate duplication and adhere to SonarQube clean code rules.
 */
public final class Literals {

    // Status & Result Keys
    public static final String STATUS = "status";
    public static final String SUCCESS = "SUCCESS";
    public static final String MESSAGE = "message";
    public static final String PRESENTATION_ID = "presentationId";
    public static final String CURRENT_PRESENTATION_ID = "currentPresentationId";
    public static final String SLIDE_INDEX = "slideIndex";
    public static final String SLIDE_COUNT = "slideCount";
    public static final String SLIDE_NUMBER = "slideNumber";
    public static final String TITLE = "title";
    public static final String SUBTITLE = "subtitle";
    public static final String WIDTH = "width";
    public static final String HEIGHT = "height";
    public static final String X = "x";
    public static final String Y = "y";
    public static final String SHAPE_NAME = "shapeName";
    public static final String SHAPE_CLASS = "shapeClass";
    public static final String SHAPE_TYPE = "shapeType";
    public static final String SHAPES = "shapes";
    public static final String SHAPE_COUNT = "shapeCount";
    public static final String ANCHOR = "anchor";
    public static final String TEXT = "text";
    public static final String OUTPUT_PATH = "outputPath";
    public static final String OUTPUT_DIR = "outputDir";
    public static final String FILE_PATH = "filePath";
    public static final String FILE_NAME = "fileName";
    public static final String DOWNLOAD_URL = "downloadUrl";
    public static final String TEMPLATE_PATH = "templatePath";
    public static final String LAYOUT_NAME = "layoutName";
    public static final String AVAILABLE_LAYOUTS = "availableLayouts";
    public static final String CORE_PROPERTIES = "coreProperties";
    public static final String PRESENTATIONS = "presentations";
    public static final String COUNT = "count";
    public static final String URL = "url";
    public static final String NOTES = "notes";
    public static final String FORMAT = "format";
    public static final String SCALE = "scale";
    public static final String IMAGE_WIDTH = "imageWidth";
    public static final String IMAGE_HEIGHT = "imageHeight";
    public static final String RENDERED_IMAGES = "renderedImages";
    public static final String HEADER_TEXT = "headerText";
    public static final String FOOTER_TEXT = "footerText";
    public static final String SLIDES_UPDATED = "slidesUpdated";
    public static final String PRESET = "preset";
    public static final String TYPE = "type";
    public static final String GRADIENT_TYPE = "gradientType";
    public static final String COLOR = "color";
    public static final String SCHEMES = "schemes";
    public static final String SCHEMES_COUNT = "schemesCount";
    public static final String CARD_COUNT = "cardCount";
    public static final String COLUMN_COUNT = "columnCount";
    public static final String STEP_COUNT = "stepCount";
    public static final String LAYOUT = "layout";
    public static final String IMAGE_PLACEMENT = "imagePlacement";
    public static final String IMAGE_STATUS = "imageStatus";

    // Theme Color Scheme Keys
    public static final String PRIMARY = "primary";
    public static final String SECONDARY = "secondary";
    public static final String ACCENT = "accent";
    public static final String BACKGROUND = "background";
    public static final String CARD_BG = "cardBg";
    public static final String TEXT_PRIMARY = "textPrimary";
    public static final String TEXT_SECONDARY = "textSecondary";

    // Common Hex Color Values
    public static final String HEX_WHITE = "#FFFFFF";
    public static final String HEX_F8FAFC = "#F8FAFC";
    public static final String HEX_0F172A = "#0F172A";
    public static final String HEX_1E3A8A = "#1E3A8A";

    // Guardrail & Validation Keys
    public static final String QUALITY_SCORE = "qualityScore";
    public static final String QUALITY_GRADE = "qualityGrade";
    public static final String ISSUES = "issues";
    public static final String ISSUE_COUNT = "issueCount";
    public static final String ERRORS = "errors";
    public static final String WARNINGS = "warnings";
    public static final String VIOLATIONS = "violations";
    public static final String PASSED = "passed";
    public static final String FIXES_APPLIED = "fixesApplied";
    public static final String SLIDE_AUDITS = "slideAudits";
    public static final String COMPLIANCE_PERCENTAGE = "compliancePercentage";

    // General Presentation, Layout & Typography Constants
    public static final String PRESENTATION = "Presentation";
    public static final String PPTX_EXTENSION = ".pptx";
    public static final String DEFAULT_PRESENTATION_FILE_NAME = "presentation.pptx";
    public static final String NO_IMAGE = "NO_IMAGE";
    public static final String FONT_SEGOE_UI = "Segoe UI";
    public static final String SEVERITY_ERROR = "ERROR";
    public static final String SEVERITY_WARNING = "WARNING";
    public static final String SEVERITY_INFO = "INFO";
    public static final String SLIDE_PREFIX = "Slide ";

    private Literals() {
        throw new UnsupportedOperationException("Literals is a constant utility class and cannot be instantiated.");
    }
}
