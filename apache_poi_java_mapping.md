# Apache POI 5.5.1 Java Method Mapping — Complete Analysis

Complete reverse-engineering of all `python-pptx` methods used across **every file** in the Python PPT MCP Server, mapped to their **exact Java Apache POI equivalent** with **JAR origins**.

---

## 1. Total Dependencies You Need in Java

### ✅ Required JARs (Maven `pom.xml`)

```xml
<!-- CORE: This single dependency pulls in everything for .pptx -->
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>5.5.1</version>
</dependency>
```

> [!IMPORTANT]
> `poi-ooxml:5.5.1` **automatically pulls** these transitive JARs via Maven:
>
> | Transitive JAR                               | What it provides                                                        |
> |:---------------------------------------------|:------------------------------------------------------------------------|
> | `org.apache.poi:poi:5.5.1`                   | Core POI engine — `XMLSlideShow`, shapes, colors, fonts, units          |
> | `org.apache.poi:poi-ooxml-lite:5.5.1`        | Pre-compiled OpenXML type bindings (`CT*` classes for XML manipulation) |
> | `org.apache.xmlbeans:xmlbeans:5.3.0`         | Low-level XML parser for `.pptx` package internals                      |
> | `org.apache.commons:commons-compress:1.27.1` | ZIP extraction for `.pptx` file format                                  |
> | `commons-io:commons-io:2.18.0`               | Stream/file I/O utilities                                               |

### ❌ You Do NOT Need `poi-ooxml-full`

`poi-ooxml-full` replaces `poi-ooxml-lite` and gives you **all** `CT*` XMLBeans classes (e.g., `CTSlideTransition`, `CTShadow`, `CTReflection`). You only need it if you do **raw XML-level manipulation** (like injecting transition XML directly).

For all 32 tools in the Python MCP server, **`poi-ooxml` alone is sufficient**. The only operations that need raw XML in the Python code (transitions, shadows, glow, reflections) are **placeholder stubs** that don't actually do anything.

### Extra JARs (replacing Python's `Pillow` and `fonttools`)

| Python Library          | What it does in the MCP server                                                     | Java Equivalent                                                                                                                          | Extra JAR needed?                                                                                                          |
|:------------------------|:-----------------------------------------------------------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------|:---------------------------------------------------------------------------------------------------------------------------|
| **`Pillow` (PIL)**      | Image enhancement (brightness, contrast, blur, filters), gradient image generation | **`java.awt.*` + `javax.imageio.*`** (JDK built-in)                                                                                      | ❌ No extra JAR — JDK `java.desktop` module has `BufferedImage`, `Graphics2D`, `RenderingHints`, `ConvolveOp`, `RescaleOp` |
| **`fonttools`**         | Font analysis (`TTFont`), font subsetting (`Subsetter`)                            | No direct equivalent needed. POI uses system-installed fonts. If you need subsetting: **Apache FOP** `fop-core` or **HarfBuzz** bindings | ⚠️ Optional — only if you need font analysis/subsetting                                                                    |
| **`base64` (stdlib)**   | Decoding base64-encoded image data                                                 | **`java.util.Base64`** (JDK built-in)                                                                                                    | ❌ No extra JAR                                                                                                            |
| **`tempfile` (stdlib)** | Temporary files for image processing                                               | **`java.io.File.createTempFile()`** (JDK built-in)                                                                                       | ❌ No extra JAR                                                                                                            |

---

## 2. Method-by-Method Mapping: Python → Java

### A. Presentation Management (`presentation_utils.py`)

| #  | Python (`python-pptx`)                        | Java (`Apache POI 5.5.1`)                                             | JAR                 |
|:---|:----------------------------------------------|:----------------------------------------------------------------------|:--------------------|
| 1  | `Presentation()` — create new                 | `new XMLSlideShow()`                                                  | `poi-ooxml`         |
| 2  | `Presentation(file_path)` — open existing     | `new XMLSlideShow(new FileInputStream(path))`                         | `poi-ooxml`         |
| 3  | `Presentation(template_path)` — from template | `new XMLSlideShow(new FileInputStream(templatePath))`                 | `poi-ooxml`         |
| 4  | `presentation.save(file_path)`                | `ppt.write(new FileOutputStream(path))`                               | `poi-ooxml`         |
| 5  | `len(presentation.slides)`                    | `ppt.getSlides().size()`                                              | `poi-ooxml`         |
| 6  | `presentation.slide_layouts`                  | `ppt.getSlideMasters().get(0).getSlideLayouts()`                      | `poi-ooxml`         |
| 7  | `layout.name`                                 | `layout.getName()`                                                    | `poi-ooxml`         |
| 8  | `len(layout.placeholders)`                    | `layout.getPlaceholders().length` *(iterate via `getPlaceholders()`)* | `poi-ooxml`         |
| 9  | `presentation.slide_width`                    | `ppt.getPageSize().width`                                             | `poi-ooxml` + `poi` |
| 10 | `presentation.slide_height`                   | `ppt.getPageSize().height`                                            | `poi-ooxml` + `poi` |
| 11 | `prs.slide_width = Inches(13.333)`            | `ppt.setPageSize(new Dimension(width_pt, height_pt))`                 | `poi-ooxml`         |

---

### B. Core Properties (`presentation_utils.py` lines 169-217)

| #  | Python (`python-pptx`)         | Java (`Apache POI 5.5.1`)                 | JAR         |
|:---|:-------------------------------|:------------------------------------------|:------------|
| 12 | `presentation.core_properties` | `ppt.getProperties().getCoreProperties()` | `poi-ooxml` |
| 13 | `core_props.title = "..."`     | `coreProps.setTitle("...")`               | `poi-ooxml` |
| 14 | `core_props.subject = "..."`   | `coreProps.setSubjectProperty("...")`     | `poi-ooxml` |
| 15 | `core_props.author = "..."`    | `coreProps.setCreator("...")`             | `poi-ooxml` |
| 16 | `core_props.keywords = "..."`  | `coreProps.setKeywords("...")`            | `poi-ooxml` |
| 17 | `core_props.comments = "..."`  | `coreProps.setDescription("...")`         | `poi-ooxml` |
| 18 | `core_props.created`           | `coreProps.getCreated()`                  | `poi-ooxml` |
| 19 | `core_props.modified`          | `coreProps.getModified()`                 | `poi-ooxml` |
| 20 | `core_props.last_modified_by`  | `coreProps.getLastModifiedByUser()`       | `poi-ooxml` |

---

### C. Slide Management (`content_utils.py` lines 17-77)

| #  | Python (`python-pptx`)                  | Java (`Apache POI 5.5.1`)                                     | JAR              |
|:---|:----------------------------------------|:--------------------------------------------------------------|:-----------------|
| 21 | `presentation.slide_layouts[index]`     | `ppt.getSlideMasters().get(0).getSlideLayouts()[index]`       | `poi-ooxml`      |
| 22 | `presentation.slides.add_slide(layout)` | `ppt.createSlide(layout)`                                     | `poi-ooxml`      |
| 23 | `pres.slides[index]`                    | `ppt.getSlides().get(index)`                                  | `poi-ooxml`      |
| 24 | `slide.slide_layout.name`               | `slide.getSlideLayout().getName()`                            | `poi-ooxml`      |
| 25 | `slide.placeholders`                    | `slide.getPlaceholders()` *(iterate)*                         | `poi-ooxml`      |
| 26 | `placeholder.placeholder_format.idx`    | `placeholder.getTextBody()` + XML parsing for ph idx          | `poi-ooxml`      |
| 27 | `placeholder.placeholder_format.type`   | Check `CTPlaceholder.getType()` via XMLBeans                  | `poi-ooxml-lite` |
| 28 | `slide.shapes`                          | `slide.getShapes()` → returns `List<XSLFShape>`               | `poi-ooxml`      |
| 29 | `shape.name`                            | `shape.getShapeName()`                                        | `poi-ooxml`      |
| 30 | `shape.shape_type`                      | `shape.getShapeType()` → `ShapeType` enum                     | `poi`            |
| 31 | `shape.left, shape.top`                 | `shape.getAnchor().getX(), shape.getAnchor().getY()`          | `poi`            |
| 32 | `shape.width, shape.height`             | `shape.getAnchor().getWidth(), shape.getAnchor().getHeight()` | `poi`            |

---

### D. Text Boxes & Formatting (`content_utils.py` lines 121-283)

| #  | Python (`python-pptx`)                      | Java (`Apache POI 5.5.1`)                                                            | JAR                            |
|:---|:--------------------------------------------|:-------------------------------------------------------------------------------------|:-------------------------------|
| 33 | `slide.shapes.add_textbox(left, top, w, h)` | `slide.createTextBox()` then `textBox.setAnchor(new Rectangle2D.Double(x, y, w, h))` | `poi-ooxml`                    |
| 34 | `textbox.text_frame.text = "..."`           | `textBox.setText("...")`                                                             | `poi-ooxml`                    |
| 35 | `text_frame.word_wrap = True`               | `textBox.setWordWrap(true)`                                                          | `poi-ooxml`                    |
| 36 | `text_frame.paragraphs`                     | `textBox.getTextParagraphs()` → `List<XSLFTextParagraph>`                            | `poi-ooxml`                    |
| 37 | `text_frame.add_paragraph()`                | `textBox.addNewTextParagraph()`                                                      | `poi-ooxml`                    |
| 38 | `text_frame.clear()`                        | `textBox.clearText()`                                                                | `poi-ooxml`                    |
| 39 | `paragraph.text = "..."`                    | `paragraph.setText("...")` *(via `addNewTextRun().setText(...)`)*                    | `poi-ooxml`                    |
| 40 | `paragraph.level = 0`                       | `paragraph.setLevel(0)`                                                              | `poi-ooxml`                    |
| 41 | `paragraph.alignment = PP_ALIGN.CENTER`     | `paragraph.setTextAlign(TextParagraph.TextAlign.CENTER)`                             | `poi`                          |
| 42 | `paragraph.runs`                            | `paragraph.getTextRuns()` → `List<XSLFTextRun>`                                      | `poi-ooxml`                    |
| 43 | `paragraph.add_run()`                       | `paragraph.addNewTextRun()`                                                          | `poi-ooxml`                    |
| 44 | `run.text = "..."`                          | `run.setText("...")`                                                                 | `poi-ooxml`                    |
| 45 | `run.font.size = Pt(24)`                    | `run.setFontSize(24.0)`                                                              | `poi-ooxml`                    |
| 46 | `run.font.name = "Segoe UI"`                | `run.setFontFamily("Segoe UI")`                                                      | `poi-ooxml`                    |
| 47 | `run.font.bold = True`                      | `run.setBold(true)`                                                                  | `poi-ooxml`                    |
| 48 | `run.font.italic = True`                    | `run.setItalic(true)`                                                                | `poi-ooxml`                    |
| 49 | `run.font.underline = True`                 | `run.setUnderlined(true)`                                                            | `poi-ooxml`                    |
| 50 | `run.font.color.rgb = RGBColor(r,g,b)`      | `run.setFontColor(new Color(r, g, b))`                                               | `poi-ooxml` + `java.awt.Color` |

> [!NOTE]
> Python's `Inches(value)` → In Java, use `Units.toEMU(value * 72)` or set anchors directly in points. POI uses EMU (English Metric Units: 1 inch = 914400 EMU). The `Rectangle2D.Double` anchor uses points.

---

### E. Title & Placeholder (`content_utils.py` lines 79-118)

| #  | Python (`python-pptx`)            | Java (`Apache POI 5.5.1`)                     | JAR         |
|:---|:----------------------------------|:----------------------------------------------|:------------|
| 51 | `slide.shapes.title`              | `slide.getTitle()` *(may return null)*        | `poi-ooxml` |
| 52 | `slide.shapes.title.text = "..."` | `slide.getTitle().setText("...")`             | `poi-ooxml` |
| 53 | `slide.placeholders[idx]`         | `slide.getPlaceholder(idx)`                   | `poi-ooxml` |
| 54 | `placeholder.text = "..."`        | `placeholder.setText("...")`                  | `poi-ooxml` |
| 55 | `placeholder.text_frame`          | `placeholder.getTextBody()` → `XSLFTextShape` | `poi-ooxml` |

---

### F. Images (`content_utils.py` lines 285-315)

| #  | Python (`python-pptx`)                            | Java (`Apache POI 5.5.1`)                                                                  | JAR                     |
|:---|:--------------------------------------------------|:-------------------------------------------------------------------------------------------|:------------------------|
| 56 | `slide.shapes.add_picture(path, left, top, w, h)` | Step 1: `XSLFPictureData pd = ppt.addPicture(new File(path), PictureData.PictureType.PNG)` | `poi-ooxml`             |
|    |                                                   | Step 2: `XSLFPictureShape pic = slide.createPicture(pd)`                                   | `poi-ooxml`             |
|    |                                                   | Step 3: `pic.setAnchor(new Rectangle2D.Double(x, y, w, h))`                                | `poi`                   |
| 57 | `Inches(value)` for positioning                   | `value * 72` (points) or `value * 914400` (EMU)                                            | `poi` (`Units.toEMU()`) |

---

### G. Tables (`content_utils.py` lines 318-373)

| #  | Python (`python-pptx`)                                | Java (`Apache POI 5.5.1`)                                | JAR         |
|:---|:------------------------------------------------------|:---------------------------------------------------------|:------------|
| 58 | `slide.shapes.add_table(rows, cols, left, top, w, h)` | `slide.createTable(rows, cols)` then set anchor          | `poi-ooxml` |
| 59 | `table.cell(row, col).text = "..."`                   | `table.getCell(row, col).setText("...")`                 | `poi-ooxml` |
| 60 | `cell.text_frame`                                     | `cell.getTextBody()` → `XSLFTextShape` methods           | `poi-ooxml` |
| 61 | `cell.fill.solid()`                                   | `cell.setFillColor(new Color(r, g, b))`                  | `poi-ooxml` |
| 62 | `cell.fill.fore_color.rgb = RGBColor(...)`            | `cell.setFillColor(new Color(r, g, b))`                  | `poi-ooxml` |
| 63 | `table.rows` / `table.columns`                        | `table.getNumberOfRows()` / `table.getNumberOfColumns()` | `poi-ooxml` |

---

### H. Charts (Status: Excluded / Deprecated in MCP Toolset)

> [!WARNING]
> **Why the Chart Feature (`add_chart`, `update_chart_data`) is Excluded from the MCP Server:**
>
> 1. **OpenXML Packaging & Relationship Fragility**:
>    - In OpenXML PresentationML, creating an embedded chart requires synchronized coordination across multiple interrelated package parts:
>      - `ppt/charts/chart1.xml` (the DrawingML chart space and plot area)
>      - `ppt/charts/_rels/chart1.xml.rels` (package relationships)
>      - `ppt/embeddings/Microsoft_Excel_Worksheet*.xlsx` (the embedded binary Excel OLE workbook)
>      - `ppt/slides/slide*.xml` (the `<p:graphicFrame>` container and coordinate transform `<p:xfrm>`)
>    - Apache POI 5.5.1's XDDF chart subsystem automatically generates external data references (`<c:externalData r:id="rId1"/>`) pointing to embedded Excel workbooks. However, POI's implementation does not reliably manage all internal OLE package streams, theme overrides, and DrawingML namespaces across various chart types.
>
> 2. **PowerPoint Strict OPC Parser Corruption Warnings**:
>    - Microsoft PowerPoint employs a strict Open Packaging Conventions (OPC) schema validator. Even minor discrepancies in embedded workbook headers, plot area boundaries, or coordinate transforms (`Units.toEMU`) cause PowerPoint to flag the presentation as corrupted and display the *"PowerPoint found a problem with content... PowerPoint can attempt to repair the presentation"* dialog.
>
> 3. **Production Design Alternatives for AI & MCP Agents**:
>    - Rather than relying on fragile OpenXML chart objects, enterprise presentations generated by AI agents achieve far higher visual fidelity, full responsiveness, and **100% corruption-free reliability** using:
>      - **Executive KPI Metric Dashboards** (`add_metric_cards_slide`): Structured value, delta trend, and highlight badges.
>      - **Formatted Benchmarking & SLA Tables** (`add_table`, `format_table_cell`): High-contrast, stylized tabular matrices.
>      - **Multi-Column Architecture & Solution Cards** (`add_comparison_slide`): Feature and tradeoff comparison columns.
>      - **Horizontal Process Flow Roadmaps** (`add_process_flow_slide`): Step-by-step phased rollout visualizations.

| Python (`python-pptx`)     | Java (`Apache POI 5.5.1`)                              | Status in Java MCP Server                                 |
|:---------------------------|:-------------------------------------------------------|:----------------------------------------------------------|
| `CategoryChartData()`      | `XDDFChartData` + `XDDFCategoryAxis` + `XDDFValueAxis` | ❌ Excluded (causes PowerPoint repair dialogs)            |
| `slide.shapes.add_chart()` | `slide.createChart()` + `slide.addChart()`             | ❌ Excluded — use `add_metric_cards_slide` or `add_table` |
| `chart.replace_data()`     | `chart.plot(chartData)`                                | ❌ Excluded                                               |

---

### I. Shapes & AutoShapes (`ppt_mcp_server.py` lines 123-188, `design_utils.py`)

| #  | Python (`python-pptx`)                                            | Java (`Apache POI 5.5.1`)                                                                            | JAR                 |
|:---|:------------------------------------------------------------------|:-----------------------------------------------------------------------------------------------------|:--------------------|
| 74 | `slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)` | `slide.createAutoShape()` then `shape.setShapeType(ShapeType.ROUND_RECT)` and `shape.setAnchor(...)` | `poi-ooxml` + `poi` |
| 75 | `shape.fill.solid()`                                              | `shape.setFillColor(new Color(r, g, b))`                                                             | `poi-ooxml`         |
| 76 | `shape.fill.fore_color.rgb = RGBColor(...)`                       | `shape.setFillColor(new Color(r, g, b))`                                                             | `poi-ooxml`         |
| 77 | `shape.line.color.rgb = RGBColor(...)`                            | `shape.setLineColor(new Color(r, g, b))`                                                             | `poi-ooxml`         |
| 78 | `shape.line.width = Pt(2)`                                        | `shape.setLineWidth(2.0)`                                                                            | `poi-ooxml`         |
| 79 | `shape.line.fill.background()` (no line)                          | `shape.setLineColor(null)` or `shape.setLineWidth(0)`                                                | `poi-ooxml`         |
| 80 | `shape.text_frame.text = "..."`                                   | `shape.setText("...")`                                                                               | `poi-ooxml`         |
| 81 | `shape.rotation = 45`                                             | `shape.setRotation(45.0)`                                                                            | `poi-ooxml`         |

**Shape Type Mapping (Python integer → Java `ShapeType` enum):**

| Python `shape_type_map` key | Java `ShapeType` enum           | From JAR |
|:----------------------------|:--------------------------------|:---------|
| `rectangle` (1)             | `ShapeType.RECT`                | `poi`    |
| `rounded_rectangle` (5)     | `ShapeType.ROUND_RECT`          | `poi`    |
| `oval` (9)                  | `ShapeType.ELLIPSE`             | `poi`    |
| `diamond` (4)               | `ShapeType.DIAMOND`             | `poi`    |
| `triangle` (7)              | `ShapeType.TRIANGLE`            | `poi`    |
| `right_triangle` (8)        | `ShapeType.RT_TRIANGLE`         | `poi`    |
| `pentagon` (51)             | `ShapeType.PENTAGON`            | `poi`    |
| `hexagon` (10)              | `ShapeType.HEXAGON`             | `poi`    |
| `octagon` (6)               | `ShapeType.OCTAGON`             | `poi`    |
| `star` (92)                 | `ShapeType.STAR_5`              | `poi`    |
| `arrow` (33)                | `ShapeType.RIGHT_ARROW`         | `poi`    |
| `cloud` (179)               | `ShapeType.CLOUD_CALLOUT`       | `poi`    |
| `heart` (21)                | `ShapeType.HEART`               | `poi`    |
| `lightning_bolt` (22)       | `ShapeType.LIGHTNING_BOLT`      | `poi`    |
| `flowchart_process` (61)    | `ShapeType.FLOW_CHART_PROCESS`  | `poi`    |
| `flowchart_decision` (63)   | `ShapeType.FLOW_CHART_DECISION` | `poi`    |

---

### J. Connectors (`connector_tools.py`)

| #  | Python (`python-pptx`)                             | Java (`Apache POI 5.5.1`)                                      | JAR         |
|:---|:---------------------------------------------------|:---------------------------------------------------------------|:------------|
| 82 | `MSO_CONNECTOR.STRAIGHT`                           | `ConnectorType.STRAIGHT` *(via XMLBeans `CTConnector`)*        | `poi-ooxml` |
| 83 | `MSO_CONNECTOR.ELBOW`                              | `ConnectorType.ELBOW`                                          | `poi-ooxml` |
| 84 | `MSO_CONNECTOR.CURVED`                             | `ConnectorType.CURVED`                                         | `poi-ooxml` |
| 85 | `slide.shapes.add_connector(type, x1, y1, x2, y2)` | `slide.createConnector(x1, y1, x2, y2)` → `XSLFConnectorShape` | `poi-ooxml` |
| 86 | `connector.line.width = Pt(1)`                     | `connector.setLineWidth(1.0)`                                  | `poi-ooxml` |
| 87 | `connector.line.color.rgb = RGBColor(...)`         | `connector.setLineColor(new Color(r, g, b))`                   | `poi-ooxml` |

---

### K. Hyperlinks (`hyperlink_tools.py`)

| #  | Python (`python-pptx`)                  | Java (`Apache POI 5.5.1`)                                   | JAR         |
|:---|:----------------------------------------|:------------------------------------------------------------|:------------|
| 88 | `run.hyperlink.address` (get)           | `run.getHyperlink().getAddress()`                           | `poi-ooxml` |
| 89 | `run.hyperlink.address = "url"` (set)   | `run.createHyperlink().setAddress("url")`                   | `poi-ooxml` |
| 90 | `run.hyperlink.address = None` (remove) | `run.getHyperlink().setAddress(null)` *(or remove via XML)* | `poi-ooxml` |

---

### L. Slide Masters (`master_tools.py`)

| #  | Python (`python-pptx`) | Java (`Apache POI 5.5.1`)                                               | JAR         |
|:---|:-----------------------|:------------------------------------------------------------------------|:------------|
| 91 | `pres.slide_masters`   | `ppt.getSlideMasters()` → `List<XSLFSlideMaster>`                       | `poi-ooxml` |
| 92 | `master.slide_layouts` | `master.getSlideLayouts()` → `XSLFSlideLayout[]`                        | `poi-ooxml` |
| 93 | `layout.name`          | `layout.getName()`                                                      | `poi-ooxml` |
| 94 | `layout.placeholders`  | iterate `layout.getPlaceholders()` *(no direct getter, iterate shapes)* | `poi-ooxml` |

---

### M. Slide Background (`design_utils.py` lines 270-326)

| #  | Python (`python-pptx`)                                 | Java (`Apache POI 5.5.1`)                                                                                  | JAR                      |
|:---|:-------------------------------------------------------|:-----------------------------------------------------------------------------------------------------------|:-------------------------|
| 95 | `slide.background.fill.solid()`                        | `slide.getBackground().setFillColor(new Color(r, g, b))`                                                   | `poi-ooxml`              |
| 96 | `slide.background.fill.fore_color.rgb = RGBColor(...)` | `slide.getBackground().setFillColor(...)`                                                                  | `poi-ooxml`              |
| 97 | Gradient background (via adding picture at 0,0)        | Same approach: create `BufferedImage` gradient → `addPicture()` → `createPicture()` at (0,0,full_w,full_h) | `poi-ooxml` + `java.awt` |

---

### N. Image Enhancement (Python `Pillow` → Java `java.awt`)

These functions from `design_utils.py` use **Pillow (PIL)**, which maps to **JDK built-in** `java.awt` and `javax.imageio`:

| #   | Python (`Pillow`)                            | Java (`java.awt.*`) — NO extra JAR                          | 
|:----|:---------------------------------------------|:------------------------------------------------------------|
| 98  | `Image.open(path)`                           | `ImageIO.read(new File(path))`                              |
| 99  | `Image.new('RGB', (w, h))`                   | `new BufferedImage(w, h, BufferedImage.TYPE_INT_RGB)`       |
| 100 | `ImageDraw.Draw(img)`                        | `img.createGraphics()` → `Graphics2D`                       |
| 101 | `draw.line([(x1,y1),(x2,y2)], fill=(r,g,b))` | `g2d.setColor(new Color(r,g,b)); g2d.drawLine(x1,y1,x2,y2)` |
| 102 | `img.putpixel((x,y), (r,g,b))`               | `img.setRGB(x, y, new Color(r,g,b).getRGB())`               |
| 103 | `img.save(path, 'PNG')`                      | `ImageIO.write(img, "png", new File(path))`                 |
| 104 | `ImageEnhance.Brightness(img).enhance(1.1)`  | `new RescaleOp(1.1f, 0, null).filter(img, null)`            |
| 105 | `ImageEnhance.Contrast(img).enhance(1.15)`   | Custom `RescaleOp` with contrast formula                    |
| 106 | `ImageFilter.GaussianBlur(radius=2)`         | `new ConvolveOp(new Kernel(...))` with Gaussian kernel      |
| 107 | `ImageFilter.SHARPEN`                        | `new ConvolveOp(new Kernel(3, 3, sharpenKernel))`           |

---

### O. Font Analysis (Python `fonttools` → Optional in Java)

| #   | Python (`fonttools`)                         | Java Equivalent                                           | Extra JAR?      |
|:----|:---------------------------------------------|:----------------------------------------------------------|:----------------|
| 108 | `TTFont(font_path)` — read font file         | `java.awt.Font.createFont(Font.TRUETYPE_FONT, file)`      | ❌ JDK built-in |
| 109 | `font['name'].names` — read font metadata    | `Font.getName()`, `Font.getFamily()`                      | ❌ JDK built-in |
| 110 | `Subsetter().subset(font)` — font subsetting | Not available in POI or JDK. Use **Apache FOP** if needed | ⚠️ Optional     |

---

### P. Units & Colors (used everywhere)

| #   | Python (`python-pptx`) | Java (`Apache POI`)                                       | JAR          |
|:----|:-----------------------|:----------------------------------------------------------|:-------------|
| 111 | `Inches(value)`        | `Units.toEMU(value * 72)` or multiply by `914400`         | `poi`        |
| 112 | `Pt(value)`            | `Units.toEMU(value)` or just pass `double` to setFontSize | `poi`        |
| 113 | `RGBColor(r, g, b)`    | `new java.awt.Color(r, g, b)`                             | JDK built-in |

---

## 3. Summary: Final `pom.xml`

```xml
<dependencies>
    <!-- ONLY dependency needed for ALL 32 tools -->
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi-ooxml</artifactId>
        <version>5.5.1</version>
    </dependency>
    
    <!-- Your Spring AI MCP dependency (for MCP server framework) -->
    <!-- <dependency>
        <groupId>org.springframework.ai</groupId>
        <artifactId>spring-ai-mcp-server-spring-boot-starter</artifactId>
    </dependency> -->
</dependencies>
```

> [!TIP]
> **Everything else** (image processing, gradient generation, font handling, base64, temp files) comes from the **JDK itself** (`java.awt`, `javax.imageio`, `java.util.Base64`, `java.io`). You don't need any additional JAR beyond `poi-ooxml:5.5.1`.

---

## 4. Quick Reference: Python Import → Java Import

```diff
- from pptx import Presentation
+ import org.apache.poi.xslf.usermodel.XMLSlideShow;

- from pptx.util import Inches, Pt
+ import org.apache.poi.util.Units;

- from pptx.dml.color import RGBColor
+ import java.awt.Color;

- from pptx.enum.text import PP_ALIGN
+ import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;

- from pptx.enum.shapes import MSO_SHAPE
+ import org.apache.poi.sl.usermodel.ShapeType;

- from pptx.enum.shapes import MSO_CONNECTOR
+ // Use XMLBeans CTConnector or XSLFConnectorShape directly

- from pptx.chart.data import CategoryChartData
- from pptx.enum.chart import XL_CHART_TYPE
+ // Charts excluded in favor of Executive Metric Cards & Tabular Dashboards

- from PIL import Image, ImageEnhance, ImageFilter, ImageDraw
+ import java.awt.image.BufferedImage;
+ import javax.imageio.ImageIO;
+ import java.awt.Graphics2D;

- from fontTools.ttLib import TTFont
+ import java.awt.Font; // JDK built-in (no fonttools equivalent needed)
```

---

## 5. Service 7: Quality Assurance, Validation & AI Guardrails (5 MCP Tools)

| Tool Name                     | Purpose & Functionality                                                                                                                                                                | Key Java / POI APIs                                                                                                 |
|:------------------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|:--------------------------------------------------------------------------------------------------------------------|
| `audit_presentation_quality`  | Automated full-deck quality audit. Calculates 0-100% Quality Score and Letter Grade. Flags coordinate overflow, collisions, empty placeholders, tiny fonts, and layout density issues. | `XMLSlideShow.getSlides()`, `XSLFShape.getAnchor()`, `Rectangle2D.intersects()`, `Rectangle2D.createIntersection()` |
| `validate_slide_quality`      | Real-time single slide quality and bounding box validation. Allows AI agents to verify slide correctness immediately after adding content.                                             | `XSLFSlide.getShapes()`, `XSLFTextShape.getTextParagraphs()`, `XSLFTextRun.getFontSize()`                           |
| `validate_presentation_rules` | User-instruction & custom constraint validator (min/max slide count, required section titles, forbidden placeholder keywords like `[TODO]`/`[TBD]`, font size floors, bullet limits).  | Case-insensitive title searching, regex word counting, bullet paragraph inspection                                  |
| `autofix_presentation_issues` | Automated remediation engine: clamps out-of-bounds overflowing shapes back onto the slide canvas with safe margins, removes ghost placeholders, standardizes font sizes and families.  | `PlaceableShape.setAnchor()`, `XSLFSlide.removeShape()`, `XSLFTextRun.setFontSize()`, `XSLFTextRun.setFontFamily()` |
| `get_guardrail_guidelines`    | Returns structured design guidelines for LLM agents (16:9 canvas grids, safe margin bounds, typography scales, contrast ratios, and anti-patterns to avoid).                           | Structured metadata dictionary and layout coordinates                                                               |

