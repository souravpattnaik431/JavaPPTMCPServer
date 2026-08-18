# 🚀 Java PowerPoint (PPT) MCP Server

An enterprise-grade **Model Context Protocol (MCP) Server** built in **Java** using **Spring AI** and **Apache POI** that enables AI assistants (GitHub Copilot, Google Antigravity, Claude Desktop, Cursor) to autonomously generate, structure, style, and audit rich PowerPoint presentations (`.pptx`).

---

## 💡 The Origin Story: Why Build This?

### 1. Inspired by Python, Supercharged with Java & Apache POI
The initial inspiration came from [GongRzhe/Office-PowerPoint-MCP-Server](https://github.com/GongRzhe/Office-PowerPoint-MCP-Server), which demonstrated how effective an MCP server can be for slide creation using Python's `python-pptx`. 

However, Python solutions often lack the low-level rendering depth, shape manipulation fidelity, and battle-tested OOXML precision that **Apache POI** provides in the Java ecosystem. The question arose: ***Why not build a comprehensive, production-grade MCP server in Java using Apache POI and Spring AI?***

### 2. Eliminating Expensive Microsoft 365 Copilot Add-On Licenses
Creating presentations with Microsoft 365 Copilot requires an expensive **extra add-on license** (often $30/user/month) on top of already paid enterprise Office 365 subscriptions (like Microsoft 365 E3 or E5). 

This project bridges that gap: it gives developers, consultants, and teams the ability to create structured, multi-slide decks directly from their favorite coding environment and AI assistant without paying for specialized office-copilot add-ons.

---

## 🛠️ The Build Journey: Foundational Architecture & "Vibe Coding"

* **Architecture First:** The foundational Spring AI MCP transport, state management, and packaging (WAR/JAR) were architected and implemented upfront.
* **Vibe Coding & Iterative Refinement:** All 44 individual tools, coordinate calculations, Apache POI shape hierarchies, custom layouts (KPI metric cards, comparison tables, roadmap flows), SonarQube-clean refactorings, and auto-save mechanisms were developed through **vibe coding**—collaborating with advanced AI via iterative prompt engineering.
* **Stress-Testing Cycle:** The server was tested against 10–15 diverse enterprise presentation prompts (e.g., Cloud Architecture, AI PR Reviews, Key Vault Security, Executive Strategies), reviewing the rendered slides visually, and improving the tool designs and error handlers cycle by cycle.

---

## 🧠 Model Recommendations & Real-World Observations

### ⚠️ GitHub Copilot "Auto" Mode vs. Frontier / Reasoning Models

During testing across various models, we observed distinct behaviors:

| Model Tier                                                                                                    | Behavior & Reliability                                                                                                                           | Recommendation                                                                                                           |
|:--------------------------------------------------------------------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------|:-------------------------------------------------------------------------------------------------------------------------|
| **GitHub Copilot Free / "Auto" Mode** *(lightweight models like GPT-4o-mini)*                                 | **~30% failure rate.** Tends to hallucinate parameters, miss required fields, or call tools out of order.                                        | ❌ Not recommended for complex multi-slide generation.                                                                   |
| **Frontier / Reasoning Models** *(Claude 4.5 / 4.6 Sonnet, GPT-5, Google Antigravity / Gemini 3.7 Thinking)* | **~95–100% precision.** Flawlessly orchestrates multi-step tool calls, plans layout variety, formats cards, and applies cohesive color palettes. | ✅ **Strongly Recommended.** Switch your Copilot chat dropdown from **Auto** to **Claude 4.5/4.6 Sonnet** or **GPT-5**. |

### 🎯 Setting Realistic Expectations (What This Is & Isn't)

* **What it is:** A high-speed **Draft v1 Presentation Accelerator**. It takes a raw idea or topic and generates 5–10 well-structured slides with tables, KPI metric cards, comparison columns, and bulleted takeaways in under 5 minutes.
* **What it is NOT:** A one-click replacement for final client deliverables. AI can still make minor aesthetic imperfections (e.g., slight font size discrepancies or color contrast variations). 
* **The Workflow:** Let the MCP server build the 80% foundation (layouts, cards, content hierarchy, tables), and you spend 30–60 minutes reviewing and fine-tuning the polish.

### ⏱️ Time Savings Matrix

| Task                                  |    Manual Creation    |        With Java PPT MCP Server         |
|:--------------------------------------|:---------------------:|:---------------------------------------:|
| Planning layout & structure           |        30 mins        |        **Instant (AI-planned)**         |
| Creating KPI cards & comparison grids |        45 mins        |             **10 seconds**              |
| Building formatted data tables        |        30 mins        |             **10 seconds**              |
| Researching & drafting bullet points  |        60 mins        |              **1 minute**               |
| Human review & final design polish    |        45 mins        |             **30–45 mins**              |
| **Total Time per 5-Slide Deck**       | **~3.5 to 4.5 Hours** | **~35 to 50 Minutes (~80% time saved)** |

---

## 📦 What's Inside? (44 MCP Tools)

The server exposes 44 dedicated tools across 7 modular service domains:

```
├── 1. Presentation Lifecycle (9 tools)      → Create, open from template, merge, export Base64, save, switch active decks
├── 2. Slide Content & Media (11 tools)       → Add slides, rich text formatting, local images, hyperlinks, speaker notes
├── 3. Custom Layout Templates (5 tools)     → Executive KPI metric cards, comparison cards, sequential roadmap flows
├── 4. Structural Elements & Tables (6 tools) → Multi-row tables, cell borders/colors, text alignment, shape connectors
├── 5. Design & Theming (5 tools)            → Gradient backgrounds, corporate color themes, headers/footers, slide numbers
├── 6. Guardrails & Quality Audits (5 tools) → Density checks, placeholder detection, canvas overflow audit, auto-fix
└── 7. Slide Rendering & Preview (3 tools)   → High-res PNG/JPEG slide rendering (960x540 / 1440x810) for visual feedback
```

---

## 💻 Tech Stack & Prerequisites

* **Java:** Java 25+ (tested on Java 25)
* **Framework:** Spring Boot 4.1.x
* **AI Protocol:** Spring AI MCP Server 2.0.0 (`spring-ai-starter-mcp-server-webmvc` / Streamable HTTP & SSE)
* **PPT Engine:** Apache POI 5.5.x (`poi-ooxml`)
* **Utilities:** Lombok, Maven Wrapper

---

## 🚀 How to Run and Use

### Option 1: Run Locally for VS Code GitHub Copilot

1. **Clone and Build the Server:**
   ```bash
   git clone <your-repo-url>
   cd PPTMCPServer
   ./mvnw spring-boot:run
   ```
   *(The server starts on `http://localhost:8080`)*

2. **Configure VS Code MCP (`.vscode/mcp.json`):**
   Create or edit `.vscode/mcp.json` inside your project root or workspace:
   ```json
   {
     "servers": {
       "java-ppt-mcp-server": {
         "type": "http",
         "url": "http://localhost:8080/mcp"
       }
     }
   }
   ```

3. **Use in GitHub Copilot Chat:**
   * Open VS Code Chat (`Ctrl+Alt+I` or `Cmd+Alt+I`).
   * Select **Claude 3.7 Sonnet**, **Claude 3.5 Sonnet**, or **GPT-4o** from the model dropdown.
   * Prompt: *"Using the java ppt mcp tools, create a 5-slide presentation on Zero-Trust Architecture with KPI metric cards and a comparison table."*

> 💡 **Prompting Tip:** Always include phrases like **`using java ppt mcp server`** or **`using ppt mcp tools`** in your prompt. Without this, some AI models may default to writing standalone Python (`python-pptx`) scripts instead of calling your live MCP server tools.

---

### Option 2: Deploy to Remote Server / Tomcat (HTTPS)

For team-wide or cloud deployments, package as a WAR/JAR and host behind a reverse proxy (e.g., NGINX / Tomcat with SSL):

1. **Build WAR package:**
   ```bash
   ./mvnw clean package -DskipTests
   ```
2. **Deploy to Tomcat / Container:** Deploy `target/PPTMCPServer-0.0.1-SNAPSHOT.war`.
3. **Configure `.vscode/mcp.json` with HTTPS:**
   ```json
   {
     "servers": {
       "enterprise-ppt-mcp": {
         "type": "http",
         "url": "https://mcp.yourdomain.com/mcp"
       }
     }
   }
   ```

---

### Option 3: Use with Claude Desktop

To connect Claude Desktop to your local Java MCP server:

1. Open `claude_desktop_config.json`:
   * **Windows:** `%APPDATA%\Claude\claude_desktop_config.json`
   * **macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`

2. Add the server entry:
   ```json
   {
     "mcpServers": {
       "java-ppt-mcp": {
         "command": "npx",
         "args": [
           "-y",
           "@modelcontextprotocol/server-sse",
           "http://localhost:8080/mcp"
         ]
       }
     }
   }
   ```
3. Restart Claude Desktop.

---

## 📝 Example Prompts to Try

> ### 💡 Why explicitly mentioning *"Using the Java PPT MCP tools"* is important:
> Many LLMs (including GPT-4o, Claude, and Gemini) have strong pre-training biases toward generating standalone Python scripts (`python-pptx`) when given generic prompts like *"Create a presentation on X"*. 
> 
> Explicitly including phrases like **`"Using the java ppt mcp server..."`** or **`"Using the ppt mcp tools..."`** accomplishes three crucial things:
> 1. **Forces MCP Tool Activation:** It directs the LLM's attention mechanism to select and invoke your registered Java MCP tools instead of outputting raw Python code.
> 2. **Leverages Apache POI Fidelity:** Ensures slides are constructed with native OOXML shape precision, theme palettes, and executive layouts rather than generic script outputs.
> 3. **Enables Real-Time Auto-Save:** Automatically persists the deck directly into your active workspace and runs quality audit guardrails in the background.

* **Executive KPI Deck:**
  > *"Using the java ppt mcp server, create an executive quarterly business review deck. Include a title slide, a 4-card KPI metric dashboard, an architecture split-screen slide, and a 4-step rollout roadmap."*
* **Technology Comparison:**
  > *"Using the ppt mcp tools, create a 5-slide comparison deck between Microservices and Monolithic architectures. Use a 2-column comparison layout with pros/cons and a summary table."*
* **Template-Based Deck:**
  > *"Using the java ppt mcp server, open the corporate template at 'templates/brand_template.pptx' and generate 4 new slides preserving the master layout and color palette."*
* **Quality Audit & Auto-Fix:**
  > *"Using the ppt mcp tools, audit the current presentation for density, contrast, or canvas overflows, and autofix detected layout defects."*

---

## 🧪 Robust Unit & Integration Test Coverage

The repository includes an extensive JUnit 5 & Spring Boot test suite covering all 44 tools and end-to-end slide generation pipelines:

* **`AutoSaveAndDefaultLocationTest`:** Validates atomic presentation creation, auto-saving directly to active client workspace roots, and title-based auto-renaming.
* **`GuardrailServiceTest`:** Tests quality auditing, visual density scoring, contrast compliance, and auto-fix remediation.
* **`PptServicesIntegrationTest`:** Tests table manipulation, custom shape connectors, typography styling, gradients, and hyperlinks.
* **`VisualVerificationTest`:** Renders complete multi-slide executive decks to high-resolution PNG images (`960x540` / `1440x810`) and verifies output rendering across all tools.
* **`CumulativeCheckTest`:** End-to-end simulation of multistep presentation assembly.

To run the complete test suite:
```bash
./mvnw test
```

---

## 🛡️ License & Acknowledgements

* **License:** MIT License.
* **Inspiration:** Thanks to [GongRzhe/Office-PowerPoint-MCP-Server](https://github.com/GongRzhe/Office-PowerPoint-MCP-Server) for inspiring the exploration of MCP PowerPoint tooling.
* **Powered by:** [Spring AI](https://spring.io/projects/spring-ai) & [Apache POI](https://poi.apache.org/).
