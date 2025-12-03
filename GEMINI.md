# Gemini Context: PowerPoint Agent Tools

## Project Overview

This project, "PowerPoint Agent Tools," is a comprehensive suite of stateless Command-Line Interface (CLI) utilities designed to enable AI agents to programmatically create, modify, validate, and export Microsoft PowerPoint (`.pptx`) presentations.

The core design philosophy is to provide a robust, predictable, and machine-friendly interface for AI agents to interact with PowerPoint files, abstracting away the complexities of the underlying file format and libraries.

**Key Characteristics:**

*   **Primary Language:** Python
*   **Core Dependencies:** `python-pptx`, `Pillow`
*   **Optional Dependencies:** `LibreOffice` (for PDF/Image export).
*   **Architecture:** The project follows a "Hub-and-Spoke" model:
    *   **Hub:** A central `core/powerpoint_agent_core.py` module encapsulates the complex business logic for file manipulation, including safety checks, low-level XML modifications, and atomic file operations.
    *   **Spokes:** A collection of over 40 individual `ppt_*.py` scripts in the `tools/` directory act as the public CLI interface. Each tool is responsible for a single, discrete action (e.g., adding a slide, inserting a shape).
*   **Interaction Model:** All tools are designed to be called from the command line, parsing arguments and producing structured `JSON` output on `stdout`. This makes the output reliable and easy for an AI agent to parse and consume. `stderr` is intentionally suppressed to maintain output purity.

## Building and Running

The project relies on Python and its dependencies, which are listed in `requirements.txt`. The recommended way to run the tools is using a Python environment manager like `uv` or `venv`.

### Installation

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/anthropics/powerpoint-agent-tools.git
    cd powerpoint-agent-tools
    ```
2.  **Install dependencies:**
    ```bash
    # Using uv (recommended)
    uv pip install -r requirements.txt

    # Or using standard pip
    pip install -r requirements.txt
    ```

### Running Tools

All tools are executed as standalone Python scripts from the `tools/` directory. The primary pattern is:

```bash
uv run tools/<tool_name>.py [ARGUMENTS] --json
```
or
```bash
python tools/<tool_name>.py [ARGUMENTS] --json
```

**Example:** Creating a new presentation and adding a slide.

```bash
# 1. Create a new presentation
uv run tools/ppt_create_new.py --output test.pptx --json

# 2. Add a slide to it
uv run tools/ppt_add_slide.py --file test.pptx --layout "Blank" --json

# 3. Inspect the result
uv run tools/ppt_get_info.py --file test.pptx --json
```

## Development Conventions

*   **Stateless and Atomic:** Each tool execution is self-contained. It opens the presentation, performs one action, saves, and closes. The state is not preserved in memory between tool calls.
*   **JSON I/O:** All tools **must** support the `--json` flag to produce machine-readable output on `stdout`. Error messages are also structured as JSON.
*   **Output Hygiene:** The `stderr` stream is redirected to `/dev/null` within each tool to prevent library "noise" from polluting the JSON output.
*   **Safety First:** The system has a multi-layered safety model that includes:
    *   **Clone-Before-Edit:** Agents are advised to work on copies of presentations.
    *   **Approval Tokens:** Destructive operations (e.g., `ppt_delete_slide.py`) require a security token to prevent accidental data loss.
    *   **Probing:** Agents are expected to use "probe" tools (`ppt_get_info.py`, `ppt_get_slide_info.py`) to inspect the presentation's state *before* performing an operation. This is critical because actions like adding or removing shapes can change the index of other shapes on a slide.
*   **Tool Template:** New tools must follow a standard template that includes the hygiene block, `sys.path` modification for standalone execution, `argparse` for CLI arguments, and a `try/except` block to ensure all output (success or failure) is valid JSON.
