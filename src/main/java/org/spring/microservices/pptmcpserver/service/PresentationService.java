package org.spring.microservices.pptmcpserver.service;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.jspecify.annotations.NonNull;
import org.spring.microservices.pptmcpserver.state.PresentationStateManager;
import org.springframework.ai.mcp.annotation.McpTool;
import org.springframework.ai.mcp.annotation.McpToolParam;
import org.springframework.stereotype.Service;

import java.awt.Dimension;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

import static org.spring.microservices.pptmcpserver.constants.Literals.*;

/**
 * Service 1: Presentation Lifecycle Management (9 MCP Tools).
 * Handles creation, opening from templates, saving, metadata, switching, and merging PowerPoint presentations.
 */
@Slf4j
@Service
@RequiredArgsConstructor
public class PresentationService {

    private final PresentationStateManager stateManager;

    /**
     * Tool 1: Create a new blank presentation with optional custom dimensions and title.
     *
     * @param title  Optional presentation title.
     * @param width  Width in points (default: 960 for 16:9 widescreen).
     * @param height Height in points (default: 540 for 16:9 widescreen).
     * @return Map with presentationId, dimensions, and status.
     */
    @McpTool(name = "create_presentation", description = "Create and initialize a new PowerPoint presentation. Provide outputPath with the active workspace file path (e.g., 'presentation.pptx' or 'output/deck.pptx') so the file is saved directly in the user's workspace.")
    public Map<String, Object> createPresentation(
            @McpToolParam(description = "Optional title for the presentation", required = false) String title,
            @McpToolParam(description = "Width in points (default: 960 for 16:9 widescreen)", required = false) Double width,
            @McpToolParam(description = "Height in points (default: 540 for 16:9 widescreen)", required = false) Double height,
            @McpToolParam(description = "Destination file path in the user's active workspace (e.g. 'deck.pptx' or 'output/deck.pptx')", required = false) String outputPath) {

        log.info("Creating new presentation with title='{}', width={}, height={}, outputPath='{}'", title, width, height, outputPath);
        XMLSlideShow ppt = new XMLSlideShow();
        double w = (width != null && width > 0) ? width : 960.0;
        double h = (height != null && height > 0) ? height : 540.0;
        ppt.setPageSize(new Dimension((int) Math.round(w), (int) Math.round(h)));

        if (title != null && !title.trim().isEmpty()) {
            POIXMLProperties.CoreProperties coreProps = ppt.getProperties().getCoreProperties();
            coreProps.setTitle(title.trim());
        }

        String presId = stateManager.store(ppt, null);
        String resolvedPath = resolveDefaultFilePath(title, outputPath);
        stateManager.setFilePath(presId, resolvedPath);
        stateManager.autoSave(presId);

        log.info("Successfully created blank presentation with ID: '{}', mapped to disk path: '{}'", presId, resolvedPath);

        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, presId);
        result.put(WIDTH, w);
        result.put(HEIGHT, h);
        result.put(TITLE, title);
        result.put(OUTPUT_PATH, resolvedPath);
        result.put(MESSAGE, "Created new blank presentation with ID: " + presId + " (mapped to: " + resolvedPath + ")");
        return result;
    }

    /**
     * Overloaded helper method to create a new blank presentation without specifying an output path.
     *
     * @param title  Optional presentation title.
     * @param width  Width in points (default: 960 for 16:9 widescreen).
     * @param height Height in points (default: 540 for 16:9 widescreen).
     * @return Map with presentationId, dimensions, and status.
     */
    public Map<String, Object> createPresentation(String title, Double width, Double height) {
        return createPresentation(title, width, height, null);
    }

    /**
     * Resolves the default destination file path for saving a presentation on disk.
     *
     * @param title      The presentation title.
     * @param outputPath The user-specified output path or directory.
     * @return The resolved file path string ending with .pptx.
     */
    private String resolveDefaultFilePath(String title, String outputPath) {
        if (outputPath != null && !outputPath.trim().isEmpty()) {
            String trimmed = outputPath.trim();
            File out = new File(trimmed);
            if (trimmed.endsWith("/") || trimmed.endsWith("\\") || (out.exists() && out.isDirectory())) {
                String fileName = (title != null && !title.trim().isEmpty())
                        ? sanitizeFileName(title) + PPTX_EXTENSION
                        : DEFAULT_PRESENTATION_FILE_NAME;
                return stateManager.resolveDestinationPath(new File(out, fileName).getPath());
            }
            return stateManager.resolveDestinationPath(trimmed.endsWith(PPTX_EXTENSION) ? trimmed : trimmed + PPTX_EXTENSION);
        }
        String fileName = (title != null && !title.trim().isEmpty())
                ? sanitizeFileName(title) + PPTX_EXTENSION
                : DEFAULT_PRESENTATION_FILE_NAME;
        return stateManager.resolveDestinationPath(fileName);
    }

    /**
     * Sanitizes a string to ensure it forms a valid filename across file systems.
     *
     * @param input The input string.
     * @return The sanitized filename safe string.
     */
    private String sanitizeFileName(String input) {
        return input.trim()
                .replaceAll("[\\\\/:*?\"<>|]", "_")
                .replaceAll("\\s+", "_")
                .replaceAll("_+", "_");
    }

    /**
     * Tool 2: Create a new presentation initialized from an existing PPTX template.
     *
     * @param templatePath File path to the .pptx template.
     * @param title        Optional title for the new presentation.
     * @return Map with presentationId, availableLayouts, and status.
     */
    @McpTool(name = "create_presentation_from_template", description = "Create a new presentation initialized from an existing branded .pptx template file. Preserves all master layouts, color themes, and formatting. By default, clears old sample slides so you start with 0 slides and can add new slides using the template's layouts.")
    public Map<String, Object> createPresentationFromTemplate(
            @McpToolParam(description = "Absolute or relative file path to the template .pptx file") String templatePath,
            @McpToolParam(description = "Optional title for the new presentation", required = false) String title,
            @McpToolParam(description = "Whether to keep existing content slides from the template (default: false - starts clean with 0 slides while preserving all master layouts and themes)", required = false) Boolean keepExistingSlides) {

        log.info("Creating presentation from template: '{}', title='{}', keepExistingSlides={}", templatePath, title, keepExistingSlides);
        if (templatePath == null || templatePath.trim().isEmpty()) {
            throw new IllegalArgumentException("templatePath cannot be empty. Action Hint: Provide a valid file path to a .pptx template (e.g. 'templates/corporate.pptx').");
        }

        File file = new File(templatePath.trim());
        if (!file.exists() || !file.isFile()) {
            throw new IllegalArgumentException(String.format(
                    "Template file does not exist at '%s'. Action Hint: Ensure the template path is correct, or use 'create_presentation' to create a blank deck instead.",
                    file.getAbsolutePath()));
        }

        try (FileInputStream fis = new FileInputStream(file)) {
            XMLSlideShow ppt = new XMLSlideShow(fis);

            if (!Boolean.TRUE.equals(keepExistingSlides)) {
                while (!ppt.getSlides().isEmpty()) {
                    ppt.removeSlide(0);
                }
                log.info("Cleared existing content slides from template '{}' to start clean with {} master layout(s)",
                        file.getName(), ppt.getSlideMasters().isEmpty() ? 0 : ppt.getSlideMasters().getFirst().getSlideLayouts().length);
            }

            String presId = stateManager.store(ppt, null);

            if (title != null && !title.trim().isEmpty()) {
                POIXMLProperties.CoreProperties coreProps = ppt.getProperties().getCoreProperties();
                coreProps.setTitle(title.trim());
            }

            String resolvedPath = resolveDefaultFilePath(title, null);
            stateManager.setFilePath(presId, resolvedPath);
            stateManager.autoSave(presId);

            log.info("Successfully created presentation from template '{}' with ID: {}", file.getName(), presId);
            return getTemplateResult(ppt, presId, file);
        } catch (Exception e) {
            log.error("Failed to create presentation from template: {}", e.getMessage(), e);
            throw new RuntimeException("Failed to create presentation from template: " + e.getMessage(), e);
        }
    }

    /**
     * Builds the result map containing template information and available slide layout names.
     *
     * @param ppt    The loaded XMLSlideShow presentation instance.
     * @param presId The generated presentation ID.
     * @param file   The template File object.
     * @return Non-null Map populated with template details.
     */
    private static @NonNull Map<String, Object> getTemplateResult(XMLSlideShow ppt, String presId, File file) {
        List<String> layoutNames = new ArrayList<>();
        if (!ppt.getSlideMasters().isEmpty()) {
            for (XSLFSlideLayout l : ppt.getSlideMasters().getFirst().getSlideLayouts()) {
                layoutNames.add(l.getName());
            }
        }

        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, presId);
        result.put(TEMPLATE_PATH, file.getAbsolutePath());
        result.put(SLIDE_COUNT, ppt.getSlides().size());
        result.put(AVAILABLE_LAYOUTS, layoutNames);
        result.put(MESSAGE, "Created new presentation from template with ID: " + presId);
        return result;
    }

    /**
     * Tool 3: Open an existing PPTX presentation file.
     *
     * @param filePath File path to the .pptx file.
     * @return Map with presentationId, slideCount, and status.
     */
    @McpTool(name = "open_presentation", description = "Open an existing PowerPoint (.pptx) file from disk into active memory.")
    public Map<String, Object> openPresentation(
            @McpToolParam(description = "Absolute or relative file path to the .pptx file") String filePath) {

        log.info("Opening presentation from file: {}", filePath);
        if (filePath == null || filePath.trim().isEmpty()) {
            throw new IllegalArgumentException("filePath cannot be empty. Action Hint: Provide the file path of the .pptx file you want to open.");
        }

        File file = new File(filePath.trim());
        if (!file.exists() || !file.isFile()) {
            throw new IllegalArgumentException(String.format(
                    "Presentation file does not exist at '%s'. Action Hint: Verify the file path, or call 'create_presentation' to initialize a new presentation.",
                    file.getAbsolutePath()));
        }

        try (FileInputStream fis = new FileInputStream(file)) {
            XMLSlideShow ppt = new XMLSlideShow(fis);
            String presId = stateManager.store(ppt, null);
            stateManager.setFilePath(presId, file.getAbsolutePath());

            log.info("Successfully opened presentation '{}' with {} slide(s), assigned ID: {}", file.getName(), ppt.getSlides().size(), presId);
            Map<String, Object> result = new LinkedHashMap<>();
            result.put(STATUS, SUCCESS);
            result.put(PRESENTATION_ID, presId);
            result.put(FILE_PATH, file.getAbsolutePath());
            result.put(SLIDE_COUNT, ppt.getSlides().size());
            result.put(MESSAGE, "Successfully opened presentation with " + ppt.getSlides().size() + " slide(s).");
            return result;
        } catch (Exception e) {
            log.error("Failed to open presentation from '{}': {}", filePath, e.getMessage(), e);
            throw new RuntimeException("Failed to open presentation: " + e.getMessage(), e);
        }
    }

    /**
     * Tool 4: Save presentation to a file on disk.
     *
     * @param presentationId Optional presentation ID.
     * @param outputPath     Target output file path.
     * @param baseDir        Optional base directory.
     * @return Map with outputPath, slideCount, and status.
     */
    @McpTool(name = "save_presentation", description = "Save and persist the PowerPoint presentation. ALWAYS call this tool after creating slides to ensure the .pptx file is written to the user's active workspace.")
    public Map<String, Object> savePresentation(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "Destination file path in the user's active workspace (e.g. 'deck.pptx' or 'output/deck.pptx')", required = false) String outputPath,
            @McpToolParam(description = "Optional base directory path", required = false) String baseDir) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Saving presentation ID '{}' to outputPath='{}', baseDir='{}'", resolvedId, outputPath, baseDir);
        XMLSlideShow ppt = stateManager.getPresentation(resolvedId);

        String finalPath;
        if (outputPath != null && !outputPath.trim().isEmpty()) {
            finalPath = stateManager.resolveDestinationPath(outputPath.trim());
        } else {
            String savedPath = stateManager.getFilePath(resolvedId);
            finalPath = (savedPath != null) ? savedPath : stateManager.resolveDestinationPath(DEFAULT_PRESENTATION_FILE_NAME);
        }

        if (baseDir != null && !baseDir.trim().isEmpty() && !new File(finalPath).isAbsolute()) {
            finalPath = stateManager.resolveDestinationPath(new File(baseDir.trim(), finalPath).getPath());
        }

        File targetFile = new File(finalPath);
        File parentDir = targetFile.getParentFile();
        if (parentDir != null && !parentDir.exists()) {
            parentDir.mkdirs();
        }

        try (FileOutputStream fos = new FileOutputStream(targetFile)) {
            ppt.write(fos);
            stateManager.setFilePath(resolvedId, targetFile.getAbsolutePath());
            log.info("Successfully saved presentation '{}' ({} slides) to '{}'", resolvedId, ppt.getSlides().size(), targetFile.getAbsolutePath());

            return createResult(resolvedId, targetFile, ppt);
        } catch (Exception e) {
            log.error("Failed to save presentation '{}' to '{}': {}", resolvedId, finalPath, e.getMessage(), e);
            throw new RuntimeException("Failed to save presentation: " + e.getMessage(), e);
        }
    }

    /**
     * create result
     * @param resolvedId mention resolveId
     * @param targetFile mention targetFile
     * @param ppt mention XMLSlideShow ppt object
     * @return result
     */
    private static @NonNull Map<String, Object> createResult(String resolvedId, File targetFile, XMLSlideShow ppt) {
        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(OUTPUT_PATH, targetFile.getAbsolutePath());
        result.put(FILE_NAME, targetFile.getName());
        result.put("fileSizeBytes", targetFile.length());
        result.put(DOWNLOAD_URL, "/api/presentations/" + resolvedId + "/download");
        result.put(SLIDE_COUNT, ppt.getSlides().size());
        result.put(MESSAGE, "Presentation successfully saved to: " + targetFile.getAbsolutePath() + " (download at: /api/presentations/" + resolvedId + "/download)");
        return result;
    }

    /**
     * Tool 5: Export presentation as Base64 binary and HTTP download link.
     * Essential for remote MCP servers deployed in cloud/WAR to stream files into local VS Code or Claude workspaces.
     *
     * @param presentationId Optional presentation ID.
     * @return Map with presentationId, fileName, fileSizeBytes, downloadUrl, base64Data, and status.
     */
    @McpTool(name = "export_presentation", description = "Export the presentation as Base64 binary data and HTTP download link. Essential for remote MCP servers deployed in cloud/WAR to save files directly into local VS Code or Claude workspaces.")
    public Map<String, Object> exportPresentation(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Exporting presentation ID '{}' as Base64 binary", resolvedId);
        XMLSlideShow ppt = stateManager.getPresentation(resolvedId);

        try (java.io.ByteArrayOutputStream baos = new java.io.ByteArrayOutputStream()) {
            ppt.write(baos);
            byte[] bytes = baos.toByteArray();
            String base64 = Base64.getEncoder().encodeToString(bytes);

            String filePath = stateManager.getFilePath(resolvedId);
            String fileName = (filePath != null) ? new File(filePath).getName() : DEFAULT_PRESENTATION_FILE_NAME;
            if (!fileName.endsWith(PPTX_EXTENSION)) {
                fileName += PPTX_EXTENSION;
            }

            Map<String, Object> result = new LinkedHashMap<>();
            result.put(STATUS, SUCCESS);
            result.put(PRESENTATION_ID, resolvedId);
            result.put(FILE_NAME, fileName);
            result.put("fileSizeBytes", bytes.length);
            result.put(DOWNLOAD_URL, "/api/presentations/" + resolvedId + "/download");
            result.put("base64Data", base64);
            result.put(MESSAGE, "Exported presentation '" + fileName + "' (" + bytes.length + " bytes). Remote clients can save this base64 data to their local workspace or download via downloadUrl.");
            return result;
        } catch (Exception e) {
            log.error("Failed to export presentation '{}': {}", resolvedId, e.getMessage(), e);
            throw new RuntimeException("Failed to export presentation: " + e.getMessage(), e);
        }
    }

    /**
     * Tool 5: Get presentation info and metadata.
     *
     * @param presentationId Optional presentation ID.
     * @return Map containing slide count, dimensions, master layouts, and core metadata properties.
     */
    @McpTool(name = "get_presentation_info", description = "Retrieve metadata, slide count, dimensions, master layouts, and core properties of a presentation.")
    public Map<String, Object> getPresentationInfo(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Retrieving presentation info for ID: '{}'", resolvedId);
        XMLSlideShow ppt = stateManager.getPresentation(resolvedId);

        Dimension pageSize = ppt.getPageSize();
        List<String> layoutNames = new ArrayList<>();
        List<XSLFSlideMaster> masters = ppt.getSlideMasters();
        if (masters != null && !masters.isEmpty()) {
            for (XSLFSlideLayout layout : masters.getFirst().getSlideLayouts()) {
                layoutNames.add(layout.getName());
            }
        }

        Map<String, Object> corePropsMap = new LinkedHashMap<>();
        try {
            POIXMLProperties.CoreProperties props = ppt.getProperties().getCoreProperties();
            corePropsMap.put(TITLE, props.getTitle());
            corePropsMap.put("creator", props.getCreator());
            corePropsMap.put("subject", props.getSubject());
            corePropsMap.put("keywords", props.getKeywords());
            corePropsMap.put("description", props.getDescription());
            corePropsMap.put("created", props.getCreated() != null ? props.getCreated().toString() : null);
            corePropsMap.put("modified", props.getModified() != null ? props.getModified().toString() : null);
            corePropsMap.put("lastModifiedByUser", props.getLastModifiedByUser());
        } catch (Exception e) {
            log.debug("Core properties read warning: {}", e.getMessage());
        }

        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(SLIDE_COUNT, ppt.getSlides().size());
        result.put(WIDTH, pageSize.getWidth());
        result.put(HEIGHT, pageSize.getHeight());
        result.put(AVAILABLE_LAYOUTS, layoutNames);
        result.put(CORE_PROPERTIES, corePropsMap);
        return result;
    }

    /**
     * Tool 6: Set core metadata properties on the presentation.
     *
     * @param presentationId Optional presentation ID.
     * @param title          Presentation title.
     * @param author         Author / creator name.
     * @param subject        Subject / topic.
     * @param keywords       Keywords separated by comma.
     * @param description    Description or comments.
     * @return Map with updated properties.
     */
    @McpTool(name = "set_core_properties", description = "Set title, author, subject, keywords, or description in presentation metadata.")
    public Map<String, Object> setCoreProperties(
            @McpToolParam(description = "Presentation ID (optional, defaults to active presentation)", required = false) String presentationId,
            @McpToolParam(description = "Presentation title", required = false) String title,
            @McpToolParam(description = "Author / Creator name", required = false) String author,
            @McpToolParam(description = "Subject / Topic", required = false) String subject,
            @McpToolParam(description = "Keywords separated by comma or semicolon", required = false) String keywords,
            @McpToolParam(description = "Description or comments", required = false) String description) {

        String resolvedId = stateManager.resolveId(presentationId);
        log.info("Updating core properties for presentation ID '{}': title='{}', author='{}'", resolvedId, title, author);
        XMLSlideShow ppt = stateManager.getPresentation(resolvedId);

        POIXMLProperties.CoreProperties coreProps = ppt.getProperties().getCoreProperties();
        if (title != null) coreProps.setTitle(title.trim());
        if (author != null) coreProps.setCreator(author.trim());
        if (subject != null) coreProps.setSubjectProperty(subject.trim());
        if (keywords != null) coreProps.setKeywords(keywords.trim());
        if (description != null) coreProps.setDescription(description.trim());

        return createResult(resolvedId, coreProps);
    }

    /**
     * Helper to assemble core properties result map.
     *
     * @param resolvedId Target presentation ID.
     * @param coreProps  CoreProperties object.
     * @return Result map.
     */
    private static @NonNull Map<String, Object> createResult(String resolvedId, POIXMLProperties.CoreProperties coreProps) {
        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(PRESENTATION_ID, resolvedId);
        result.put(TITLE, coreProps.getTitle());
        result.put("creator", coreProps.getCreator());
        result.put("subject", coreProps.getSubject());
        result.put("keywords", coreProps.getKeywords());
        result.put("description", coreProps.getDescription());
        result.put(MESSAGE, "Core properties updated successfully.");
        return result;
    }

    /**
     * Tool 7: List all loaded presentations in memory.
     *
     * @return Map containing all loaded presentation metadata.
     */
    @McpTool(name = "list_presentations", description = "List all presentation instances currently loaded into memory with their ID and slide counts.")
    public Map<String, Object> listPresentations() {
        log.info("Listing all active presentations in memory");
        List<Map<String, Object>> list = stateManager.listAll();
        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(COUNT, list.size());
        result.put(CURRENT_PRESENTATION_ID, stateManager.getCurrentPresentationId());
        result.put(PRESENTATIONS, list);
        return result;
    }

    /**
     * Tool 8: Switch active presentation in memory.
     *
     * @param presentationId The target presentation ID to switch to.
     * @return Map confirming switch.
     */
    @McpTool(name = "switch_presentation", description = "Switch the active presentation in memory to a given presentation ID.")
    public Map<String, Object> switchPresentation(
            @McpToolParam(description = "The presentation ID to switch to") String presentationId) {

        log.info("Switching active presentation to: '{}'", presentationId);
        stateManager.setCurrentPresentationId(presentationId);
        Map<String, Object> result = new LinkedHashMap<>();
        result.put(STATUS, SUCCESS);
        result.put(CURRENT_PRESENTATION_ID, presentationId);
        result.put(MESSAGE, "Switched active presentation to: " + presentationId);
        return result;
    }

    /**
     * Tool 9: Merge slides from another PPTX file into the target presentation.
     *
     * @param sourceFilePath       Path to external source .pptx file.
     * @param targetPresentationId Target presentation ID.
     * @return Map with imported count and total slides.
     */
    @McpTool(name = "merge_presentations", description = "Merge all slides from an external .pptx file into the active or specified presentation.")
    public Map<String, Object> mergePresentations(
            @McpToolParam(description = "Source .pptx file path to merge from") String sourceFilePath,
            @McpToolParam(description = "Target presentation ID (optional, defaults to active presentation)", required = false) String targetPresentationId) {

        if (sourceFilePath == null || sourceFilePath.trim().isEmpty()) {
            throw new IllegalArgumentException("sourceFilePath cannot be empty. Action Hint: Provide a valid .pptx file path to import slides from.");
        }

        File srcFile = new File(sourceFilePath.trim());
        if (!srcFile.exists() || !srcFile.isFile()) {
            throw new IllegalArgumentException(String.format(
                    "Source file to merge does not exist at '%s'. Action Hint: Ensure the source .pptx file path is correct.",
                    srcFile.getAbsolutePath()));
        }

        String resolvedId = stateManager.resolveId(targetPresentationId);
        log.info("Merging slides from '{}' into presentation ID '{}'", srcFile.getName(), resolvedId);
        XMLSlideShow targetPpt = stateManager.getPresentation(resolvedId);

        int importedCount = 0;
        try (FileInputStream fis = new FileInputStream(srcFile);
             XMLSlideShow srcPpt = new XMLSlideShow(fis)) {

            for (XSLFSlide srcSlide : srcPpt.getSlides()) {
                XSLFSlide newSlide = targetPpt.createSlide();
                newSlide.importContent(srcSlide);
                importedCount++;
            }

            log.info("Successfully imported {} slides from '{}' into '{}'", importedCount, srcFile.getName(), resolvedId);
            Map<String, Object> result = new LinkedHashMap<>();
            result.put(STATUS, SUCCESS);
            result.put("targetPresentationId", resolvedId);
            result.put("sourceFilePath", srcFile.getAbsolutePath());
            result.put("importedSlideCount", importedCount);
            result.put("totalSlideCount", targetPpt.getSlides().size());
            result.put(MESSAGE, String.format("Successfully merged %d slides from %s into presentation %s",
                    importedCount, srcFile.getName(), resolvedId));
            return result;
        } catch (Exception e) {
            log.error("Failed to merge presentation '{}': {}", sourceFilePath, e.getMessage(), e);
            throw new RuntimeException("Failed to merge presentations: " + e.getMessage(), e);
        }
    }
}
