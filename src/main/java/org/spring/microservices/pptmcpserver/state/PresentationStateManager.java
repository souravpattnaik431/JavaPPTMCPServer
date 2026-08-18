package org.spring.microservices.pptmcpserver.state;

import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.springframework.stereotype.Component;

import java.awt.Dimension;
import java.io.File;
import java.io.IOException;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.atomic.AtomicInteger;

import static org.spring.microservices.pptmcpserver.constants.Literals.*;

/**
 * Shared state manager for PowerPoint presentations across all MCP services.
 * Thread-safe in-memory cache supporting multiple active presentations concurrently.
 */
@Slf4j
@Component
public class PresentationStateManager {

    private final ConcurrentHashMap<String, XMLSlideShow> presentations = new ConcurrentHashMap<>();
    private final ConcurrentHashMap<String, String> presentationFilePaths = new ConcurrentHashMap<>();
    private final AtomicInteger idCounter = new AtomicInteger(0);
    /**
     * -- GETTER --
     *  Get the ID of the current active presentation.
     *
     */
    @Getter
    private volatile String currentPresentationId = null;

    /**
     * Active client workspace root directory detected from MCP roots notifications.
     */
    @Getter
    @Setter
    private volatile String workspaceRoot = null;

    /**
     * Updates the active client workspace root from MCP roots notification.
     *
     * @param roots List of workspace roots provided by client.
     */
    public void updateWorkspaceRoots(List<io.modelcontextprotocol.spec.McpSchema.Root> roots) {
        if (roots == null || roots.isEmpty()) {
            return;
        }
        String currentServerDir = resolveCurrentServerDir();
        String detectedRoot = findValidWorkspaceRoot(roots, currentServerDir);
        if (detectedRoot != null) {
            this.workspaceRoot = detectedRoot;
            log.info("Active client workspace root detected: '{}'", detectedRoot);
            return;
        }
        applyFallbackWorkspaceRoot(roots.getFirst());
    }

    private String resolveCurrentServerDir() {
        try {
            return new File(".").getCanonicalPath();
        } catch (Exception _) {
            return new File(".").getAbsolutePath();
        }
    }

    private String findValidWorkspaceRoot(List<io.modelcontextprotocol.spec.McpSchema.Root> roots, String currentServerDir) {
        for (io.modelcontextprotocol.spec.McpSchema.Root r : roots) {
            if (r.uri() != null && r.uri().startsWith("file:")) {
                try {
                    File f = new File(java.net.URI.create(r.uri()));
                    if (f.exists() && f.isDirectory()) {
                        String abs = f.getCanonicalPath();
                        if (!abs.equalsIgnoreCase(currentServerDir)) {
                            return abs;
                        }
                    }
                } catch (Exception e) {
                    log.warn("Failed to parse root URI '{}': {}", r.uri(), e.getMessage());
                }
            }
        }
        return null;
    }

    private void applyFallbackWorkspaceRoot(io.modelcontextprotocol.spec.McpSchema.Root firstRoot) {
        try {
            File f = new File(java.net.URI.create(firstRoot.uri()));
            this.workspaceRoot = f.getCanonicalPath();
            log.info("Client workspace root set to: '{}'", this.workspaceRoot);
        } catch (Exception _) {
            // Ignored as fallback failsafe when root URI cannot be converted to local path
        }
    }

    /**
     * Resolves a destination file path against the active client workspace root if relative.
     *
     * @param path The input file path (relative or absolute).
     * @return The fully resolved absolute file path string.
     */
    public String resolveDestinationPath(String path) {
        if (path == null || path.trim().isEmpty()) {
            path = DEFAULT_PRESENTATION_FILE_NAME;
        }
        File f = new File(path.trim());
        if (f.isAbsolute()) {
            return f.getAbsolutePath();
        }
        if (workspaceRoot != null && !workspaceRoot.trim().isEmpty()) {
            return new File(workspaceRoot.trim(), path.trim()).getAbsolutePath();
        }
        return f.getAbsolutePath();
    }

    /**
     * Generate a unique presentation ID.
     *
     * @return Generated ID string (e.g., 'presentation_1').
     */
    public String generateId() {
        return "presentation_" + idCounter.incrementAndGet();
    }

    /**
     * Store a presentation in the state manager.
     *
     * @param ppt The XMLSlideShow instance.
     * @param id  Optional ID. If null or blank, an ID will be generated.
     * @return The ID assigned to this presentation.
     */
    public String store(XMLSlideShow ppt, String id) {
        if (id == null || id.trim().isEmpty()) {
            id = generateId();
        }
        presentations.put(id, ppt);
        if (currentPresentationId == null) {
            currentPresentationId = id;
        }
        log.debug("Stored presentation with ID: '{}' (Total in memory: {})", id, presentations.size());
        return id;
    }

    /**
     * Associate a file path with a presentation ID.
     * Automatically resolves against the client workspace root if relative.
     *
     * @param id       Presentation ID.
     * @param filePath File path on disk.
     */
    public void setFilePath(String id, String filePath) {
        if (id != null && filePath != null) {
            String resolved = resolveDestinationPath(filePath);
            presentationFilePaths.put(id, resolved);
            log.info("Associated file path '{}' with presentation ID '{}'", resolved, id);
        }
    }

    /**
     * Get the file path associated with a presentation ID.
     *
     * @param id Presentation ID.
     * @return Associated file path or null.
     */
    public String getFilePath(String id) {
        return presentationFilePaths.get(id);
    }

    /**
     * Automatically persists the presentation to disk if a file path is associated with it.
     *
     * @param id Presentation ID.
     */
    public void autoSave(String id) {
        if (id == null) {
            return;
        }
        String filePath = presentationFilePaths.get(id);
        if (filePath == null || filePath.trim().isEmpty()) {
            return;
        }
        XMLSlideShow ppt = presentations.get(id);
        if (ppt == null) {
            return;
        }

        try {
            java.io.File targetFile = new java.io.File(filePath);
            java.io.File parentDir = targetFile.getParentFile();
            if (parentDir != null && !parentDir.exists()) {
                parentDir.mkdirs();
            }
            try (java.io.FileOutputStream fos = new java.io.FileOutputStream(targetFile)) {
                ppt.write(fos);
            }
            log.debug("Auto-saved presentation '{}' to '{}'", id, targetFile.getAbsolutePath());
        } catch (java.io.IOException e) {
            log.warn("Auto-save failed for presentation '{}': {}", id, e.getMessage());
        }
    }

    /**
     * Retrieve a presentation by ID. If ID is null or blank, retrieves the current active presentation.
     *
     * @param id Presentation ID.
     * @return Active XMLSlideShow instance.
     */
    public XMLSlideShow getPresentation(String id) {
        if (id == null || id.trim().isEmpty()) {
            return getCurrentPresentation();
        }
        XMLSlideShow ppt = presentations.get(id);
        if (ppt == null) {
            log.error("Presentation not found for ID: '{}'", id);
            throw new IllegalArgumentException(String.format(
                    "Presentation not found for ID '%s'. Active presentation IDs: %s. Action Hint: Omit 'presentationId' to use the current active presentation, or call 'create_presentation' / 'open_presentation' first.",
                    id, presentations.keySet()));
        }
        return ppt;
    }

    /**
     * Get the active/current presentation.
     *
     * @return Current XMLSlideShow instance.
     */
    public XMLSlideShow getCurrentPresentation() {
        if (currentPresentationId != null) {
            XMLSlideShow ppt = presentations.get(currentPresentationId);
            if (ppt != null) {
                return ppt;
            }
        }
        if (!presentations.isEmpty()) {
            Map.Entry<String, XMLSlideShow> first = presentations.entrySet().iterator().next();
            currentPresentationId = first.getKey();
            return first.getValue();
        }
        log.error("No presentation currently open or loaded in memory.");
        throw new IllegalStateException("No presentation is currently open or loaded. Action Hint: Call 'create_presentation(title=\"...\")' or 'open_presentation(filePath=\"...\")' first before adding slides.");
    }

    /**
     * Set the current active presentation ID.
     *
     * @param id Presentation ID to switch to.
     */
    public void setCurrentPresentationId(String id) {
        if (id == null || !presentations.containsKey(id)) {
            log.error("Cannot switch to non-existent presentation ID: '{}'", id);
            throw new IllegalArgumentException(String.format(
                    "Cannot switch to non-existent presentation ID '%s'. Available presentation IDs: %s. Action Hint: Call 'list_presentations' to see all open presentations.",
                    id, presentations.keySet()));
        }
        this.currentPresentationId = id;
        log.info("Active presentation ID switched to: '{}'", id);
    }

    /**
     * Remove and close a presentation by ID.
     *
     * @param id Presentation ID to remove.
     */
    public void remove(String id) {
        XMLSlideShow removed = presentations.remove(id);
        presentationFilePaths.remove(id);
        if (removed != null) {
            try {
                removed.close();
            } catch (IOException e) {
                log.warn("Error closing presentation '{}': {}", id, e.getMessage());
            }
        }
        if (Objects.equals(currentPresentationId, id)) {
            currentPresentationId = presentations.isEmpty() ? null : presentations.keySet().iterator().next();
        }
        log.info("Removed presentation ID '{}' from memory", id);
    }

    /**
     * Clears all open presentations and reset active ID.
     */
    public void clearAll() {
        for (String id : new ArrayList<>(presentations.keySet())) {
            remove(id);
        }
        currentPresentationId = null;
    }

    /**
     * Safely retrieve a slide from a presentation by slide index (0-indexed).
     *
     * @param presId     Presentation ID.
     * @param slideIndex 0-based slide index.
     * @return Target XSLFSlide instance.
     */
    public XSLFSlide getSlide(String presId, int slideIndex) {
        XMLSlideShow ppt = getPresentation(presId);
        List<XSLFSlide> slides = ppt.getSlides();
        if (slideIndex < 0 || slideIndex >= slides.size()) {
            log.error("Slide index {} out of bounds for presentation '{}' (total: {})", slideIndex, presId, slides.size());
            throw new IllegalArgumentException(String.format(
                    "Slide index %d out of bounds for presentation '%s' (total slides: %d). Action Hint: Slide indices are 0-based (0 to %d). Call 'add_slide' to create a slide first.",
                    slideIndex, presId, slides.size(), Math.max(0, slides.size() - 1)));
        }
        return slides.get(slideIndex);
    }

    /**
     * List all loaded presentations and their metadata.
     *
     * @return List of presentation metadata maps.
     */
    public List<Map<String, Object>> listAll() {
        List<Map<String, Object>> list = new ArrayList<>();
        for (Map.Entry<String, XMLSlideShow> entry : presentations.entrySet()) {
            String id = entry.getKey();
            XMLSlideShow ppt = entry.getValue();
            Dimension pageSize = ppt.getPageSize();
            Map<String, Object> meta = new LinkedHashMap<>();
            meta.put(PRESENTATION_ID, id);
            meta.put(SLIDE_COUNT, ppt.getSlides().size());
            meta.put(WIDTH, pageSize.getWidth());
            meta.put(HEIGHT, pageSize.getHeight());
            meta.put(FILE_PATH, presentationFilePaths.get(id));
            meta.put("isCurrent", Objects.equals(id, currentPresentationId));
            list.add(meta);
        }
        return list;
    }

    /**
     * Resolve effective presentation ID (uses passed id or current active ID).
     *
     * @param id Optional ID.
     * @return Resolved effective presentation ID.
     */
    public String resolveId(String id) {
        if (id != null && !id.trim().isEmpty()) {
            return id.trim();
        }
        if (currentPresentationId != null) {
            return currentPresentationId;
        }
        if (!presentations.isEmpty()) {
            return presentations.keySet().iterator().next();
        }
        return null;
    }

    /**
     * Updates the presentation file path from the slide title if it's currently using the generic default name or unmapped.
     * Automatically resolves the destination path in the active client workspace.
     *
     * @param id    Presentation ID.
     * @param title Title of the slide.
     */
    public void updateFilePathFromTitleIfDefault(String id, String title) {
        if (id == null || title == null || title.trim().isEmpty()) {
            return;
        }
        String currentPath = presentationFilePaths.get(id);
        if (currentPath != null && !currentPath.endsWith(DEFAULT_PRESENTATION_FILE_NAME)) {
            return;
        }
        String sanitized = sanitizeTitleToFileName(title);
        if (sanitized.isEmpty()) {
            return;
        }
        String newPath = resolveDestinationPath(sanitized + PPTX_EXTENSION);
        if (newPath.equalsIgnoreCase(currentPath)) {
            return;
        }
        if (currentPath != null) {
            deleteOldDefaultFile(currentPath);
        }
        setFilePath(id, newPath);
        log.info("Auto-renamed presentation '{}' file path based on title to: '{}'", id, newPath);
    }

    private static String sanitizeTitleToFileName(String title) {
        return title.trim()
                .replaceAll("[\\\\/:*?\"<>|]", "_")
                .replaceAll("\\s+", "_")
                .replaceAll("_+", "_");
    }

    private void deleteOldDefaultFile(String currentPath) {
        try {
            File oldFile = new File(currentPath);
            if (oldFile.exists() && oldFile.length() < 10000) {
                boolean deleted = oldFile.delete();
                if (!deleted) {
                    log.debug("Old presentation file '{}' could not be deleted", oldFile.getAbsolutePath());
                }
            }
        } catch (Exception _) {
            // Best-effort cleanup of previous default presentation file
        }
    }

    /**
     * Atomically get-or-create a presentation by ID (or the active/default one if id is null/blank).
     * Uses synchronized access to prevent concurrent threads from each creating their own
     * separate presentation when no presentation exists yet.
     * Automatically binds the presentation to the active client workspace root so files are
     * persisted to disk immediately as slides are added.
     *
     * @param id Presentation ID (nullable/blank → use current active or create new).
     * @return A two-element Object array: [resolvedId (String), ppt (XMLSlideShow)].
     */
    public synchronized Object[] getOrCreatePresentation(String id) {
        // Re-resolve inside the synchronized block to pick up any presentation
        // created by a concurrently-entering thread that got the lock first.
        String resolvedId = resolveId(id);
        if (resolvedId != null) {
            XMLSlideShow existing = presentations.get(resolvedId);
            if (existing != null) {
                return new Object[]{resolvedId, existing};
            }
        }
        log.warn("No presentation found for ID '{}' — auto-creating a blank 16:9 presentation in workspace.", resolvedId);
        XMLSlideShow ppt = new XMLSlideShow();
        ppt.setPageSize(new java.awt.Dimension(960, 540));
        resolvedId = store(ppt, null);
        String defaultPath = resolveDestinationPath(DEFAULT_PRESENTATION_FILE_NAME);
        setFilePath(resolvedId, defaultPath);
        autoSave(resolvedId);
        log.info("Auto-created presentation with ID '{}' (initialized on disk at: '{}')", resolvedId, defaultPath);
        return new Object[]{resolvedId, ppt};
    }
}
