# 📚 AGENT SYSTEM REFERENCE (v3.1.1)

> **PowerPoint Agent Tools** - Enabling AI agents to engineer presentations with precision, safety, and visual intelligence.

**Document Version:** 1.2.0
**Project Version:** 3.1.1
**Last Updated:** December 2025

---

## 🚀 Quick Start Guide

**Get up and running in 60 seconds**

```bash
# 1. Clone the repository
git clone https://github.com/anthropics/powerpoint-agent-tools.git
cd powerpoint-agent-tools

# 2. Install dependencies (uv recommended)
# Note: Requires python-pptx >= 0.6.23 and Pillow >= 12.0.0
uv pip install -r requirements.txt
uv pip install -r requirements-dev.txt

# 3. Create a test presentation
uv run tools/ppt_create_new.py --output test.pptx --json
uv run tools/ppt_add_slide.py --file test.pptx --layout "Blank" --json

# 4. Inspect the presentation (Probe)
uv run tools/ppt_get_info.py --file test.pptx --json

# 5. Add a semi-transparent overlay (Overlay Pattern)
uv run tools/ppt_add_shape.py --file test.pptx --slide 0 --shape rectangle \
  --position '{"left":"0%","top":"0%"}' --size '{"width":"100%","height":"100%"}' \
  --fill-color "#FFFFFF" --fill-opacity 0.15 --json

# 6. Validate result
uv run tools/ppt_validate_presentation.py --file test.pptx --policy standard --json
```

### 🔑 Key Concepts to Remember

| Concept | Rule | Why It Matters |
|---------|------|----------------|
| 🔒 **Clone Before Edit** | Never modify source files directly | Prevents accidental data loss |
| 🔍 **Probe Before Operate** | Always inspect slide structure first | Avoids layout guessing errors |
| 🔄 **Refresh Indices** | Re-query after structural operations | Shape indices shift after changes |
| 📊 **JSON-First I/O** | All tools output structured JSON | Enables machine parsing |
| 🤫 **Output Hygiene** | `stderr` suppressed for clean JSON | Prevents parsing errors |
| 👮 **Governance** | Destructive ops require tokens | Prevents unauthorized deletion |

---

## ✨ What's New in v3.1.1

| Feature | Description |
|---------|-------------|
| 🧩 **Merge Tool** | `ppt_merge_presentations.py` combines slides from multiple decks |
| 🔎 **Content Search** | `ppt_search_content.py` finds text/regex across slides/notes |
| 🔌 **JSON Adapter** | `ppt_json_adapter.py` normalizes and validates tool outputs |
| 🛡️ **Validation Policies** | `lenient`, `standard`, and `strict` validation profiles |
| 🎨 **Opacity Support** | Native `fill_opacity` (0.0-1.0) replacing legacy transparency |
| 📝 **Smart Notes** | `ppt_add_notes.py` supports append, prepend, and overwrite modes |

---

## 📋 Table of Contents

1. [🎯 Project Identity & Mission](#1-project-identity--mission)
2. [🏗️ Architecture Overview](#2-architecture-overview)
3. [🏛️ Design Philosophy](#3-design-philosophy)
4. [🛠️ Programming Model](#4-programming-model)
5. [📏 Code Standards](#5-code-standards)
6. [⚠️ Critical Patterns & Gotchas](#6-critical-patterns--gotchas)
7. [🧪 Testing Requirements](#7-testing-requirements)
8. [📤 Contribution Workflow](#8-contribution-workflow)
9. [📖 Quick Reference](#9-quick-reference)
10. [🔧 Troubleshooting](#10-troubleshooting)

---

## 1. 🎯 Project Identity & Mission

### Core Mission

**"Enabling AI agents to engineer presentations with precision, safety, and visual intelligence"**

PowerPoint Agent Tools is a suite of **42 stateless CLI utilities** designed for AI agents to programmatically create, modify, and validate PowerPoint (`.pptx`) files.

### Compatibility Matrix

| Component | Version | Notes |
|-----------|---------|-------|
| Python | 3.8+ | 3.10+ recommended |
| python-pptx | **0.6.23** | Required for latest features |
| Pillow | **>= 12.0.0** | For image processing/compression |
| LibreOffice | 7.4+ | Required for PDF/Image export |
| uv | Latest | Recommended package manager |

---

## 2. 🏗️ Architecture Overview

### Hub-and-Spoke Model

```
                         ┌─────────────────────────┐
                         │   AI Agent / Human      │
                         │   (Orchestration Layer) │
                         └───────────┬─────────────┘
                                     │
                    ┌────────────────┼────────────────┐
                    │                │                │
                    ▼                ▼                ▼
           ┌───────────────┐ ┌───────────────┐ ┌───────────────┐
           │ ppt_add_      │ │ ppt_merge_    │ │ ppt_validate_ │
           │ shape.py      │ │ slides.py     │ │ presentation  │
           │   (SPOKE)     │ │   (SPOKE)     │ │   (SPOKE)     │
           └───────┬───────┘ └───────┬───────┘ └───────┬───────┘
                   │                 │                 │
                   └─────────────────┼─────────────────┘
                                     │
                                     ▼
                    ┌─────────────────────────────────┐
                    │   powerpoint_agent_core.py      │
                    │            (HUB)                │
                    │                                 │
                    │   • PowerPointAgent (Context)   │
                    │   • Atomic File Locking         │
                    │   • XML Manipulation (Opacity)  │
                    │   • Geometry-Aware Versioning   │
                    │   • Approval Token Validation   │
                    └─────────────────────────────────┘
```

### Exit Code Matrix (v3.1.1)

The tools utilize a standardized exit code system for robust error handling:

| Code | Meaning | Recovery Action |
|------|---------|-----------------|
| `0` | **Success** | Proceed to next step. |
| `1` | **Usage Error** | Check arguments and file paths. |
| `2` | **Validation Error** | JSON Schema validation failed. Check input format. |
| `3` | **Transient Error** | Timeout or Lock Error. Retry with backoff. |
| `4` | **Permission Error** | Approval token missing or invalid scope. |
| `5` | **Internal Error** | Unexpected crash. Check logs. |

---

## 3. 🏛️ Design Philosophy

### The Five-Level Safety Hierarchy

1.  **Clone-Before-Edit**: Modify only working copies (`/work/` or cloned files).
2.  **Approval Tokens**: Destructive actions (`delete_slide`, `remove_shape`) strictly require HMAC tokens.
3.  **Output Hygiene**: `sys.stderr` is redirected to `/dev/null` to guarantee JSON purity on stdout.
4.  **Version Hashing**: Geometry-aware hashing prevents race conditions by detecting layout shifts.
5.  **Accessibility**: WCAG 2.1 checks are built-in (Alt Text, Contrast).

### The Statelessness Contract

```python
# ✅ CORRECT: Atomic Context
with PowerPointAgent(path) as agent:
    agent.open(path)
    agent.add_slide(...)
    agent.save() 
# File unlocked, memory cleared

# ❌ WRONG: Persistent State assumption
agent = PowerPointAgent(path) # No context manager
agent.add_slide(...) # May fail or leave locks
```

---

## 4. 🛠️ Programming Model

### Tool Development Template

Use this template for all new tools to ensure v3.1.1 compliance.

```python
#!/usr/bin/env python3
"""
PowerPoint [Action] Tool v3.1.1
[Description]

Usage: uv run tools/ppt_[name].py ... --json
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null immediately to prevent library noise.
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any

# Allow importing core without package installation
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError
)

__version__ = "3.1.1"

def do_action(filepath: Path, args) -> Dict[str, Any]:
    # Validate file existence
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")

    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture state before
        version_before = agent.get_presentation_version()
        
        # Perform Operation
        # result = agent.some_method(...)
        
        agent.save()
        
        # Capture state after
        version_after = agent.get_presentation_version()
        
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--file', required=True, type=Path)
    parser.add_argument('--json', action='store_true', default=True)
    args = parser.parse_args()
    
    try:
        result = do_action(args.file, args)
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        # Standard Error Response
        print(json.dumps({
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }, indent=2))
        sys.exit(1)

if __name__ == "__main__":
    main()
```

---

## 5. 📏 Code Standards

### Naming Conventions & Structure
*   **Files**: `ppt_<verb>_<noun>.py` (e.g., `ppt_add_slide.py`).
*   **Imports**: `sys.path.insert` block required for standalone execution.
*   **Hygiene**: `sys.stderr` suppression block is **mandatory** at top of file.

### Data Structures

**Position Dictionary:**
```json
{"left": "10%", "top": "20%"} 
{"anchor": "center", "offset_x": 0, "offset_y": -1.0}
{"grid_row": 2, "grid_col": 3, "grid_size": 12}
```

**Size Dictionary:**
```json
{"width": "50%", "height": "auto"}
{"width": 5.0, "height": 3.0}
```

---

## 6. ⚠️ Critical Patterns & Gotchas

### 1. Shape Index Shift
**The Rule:** Structural operations (Add, Remove, Z-Order, Group) invalidate shape indices.
**The Fix:** ALWAYS re-query `ppt_get_slide_info.py` after these operations. Do not cache indices.

### 2. The Overlay Pattern
**Goal:** Make text readable over images.
**Steps:**
1. Add Shape: `ppt_add_shape.py` with `fill_opacity=0.15`.
2. **Refresh Indices**: `ppt_get_slide_info.py` (New shape is at top).
3. Send to Back: `ppt_set_z_order.py --action send_to_back`.
4. **Refresh Indices**: Indices have shifted again!

### 3. Chart Updates
**Limitation:** `python-pptx` cannot reliably update charts if the schema changes.
**Strategy:**
1. Delete old chart (`ppt_remove_shape.py` + Token).
2. Create new chart (`ppt_add_chart.py`).

### 4. Validation Policies
The `ppt_validate_presentation.py` tool supports three strictness levels:
*   `lenient`: Drafts. Allows empty slides, some missing alt text.
*   `standard`: Internal use. Balanced checks.
*   `strict`: Production. Zero tolerance for accessibility/structure issues.

---

## 9. 📖 Quick Reference: Tool Catalog (42 Tools)

### Creation & Composition
| Tool | Description |
|------|-------------|
| `ppt_create_new.py` | Create blank presentation |
| `ppt_create_from_template.py` | Create from .pptx template |
| `ppt_create_from_structure.py` | Build full deck from JSON definition |
| `ppt_clone_presentation.py` | **Safety**: Create working copy |
| `ppt_merge_presentations.py` | **New**: Combine slides from multiple decks |

### Slide Management
| Tool | Description |
|------|-------------|
| `ppt_add_slide.py` | Add slide with specific layout |
| `ppt_delete_slide.py` | **Destructive**: Remove slide (Requires Token) |
| `ppt_duplicate_slide.py` | Clone existing slide |
| `ppt_reorder_slides.py` | Move slide position |
| `ppt_set_slide_layout.py` | Change layout of existing slide |
| `ppt_set_footer.py` | Configure footer/numbers |
| `ppt_set_background.py` | Set color or image background |

### Content & Text
| Tool | Description |
|------|-------------|
| `ppt_add_text_box.py` | Add text container |
| `ppt_add_bullet_list.py` | Add list with 6x6 rule validation |
| `ppt_set_title.py` | Set title/subtitle |
| `ppt_format_text.py` | Style text (font, size, color) |
| `ppt_replace_text.py` | Find & replace (global/scoped) |
| `ppt_add_notes.py` | Add speaker notes |

### Visuals & Shapes
| Tool | Description |
|------|-------------|
| `ppt_add_shape.py` | Add geometry with opacity |
| `ppt_format_shape.py` | Style shapes (fill, line) |
| `ppt_remove_shape.py` | **Destructive**: Remove shape (Requires Token) |
| `ppt_set_z_order.py` | Layer management (Front/Back) |
| `ppt_add_connector.py` | Connect two shapes |

### Images & Media
| Tool | Description |
|------|-------------|
| `ppt_insert_image.py` | Add image with auto-sizing |
| `ppt_replace_image.py` | Swap image, keep position |
| `ppt_crop_image.py` | Crop visual extent |
| `ppt_set_image_properties.py`| Set Alt Text (Accessibility) |

### Data Visualization
| Tool | Description |
|------|-------------|
| `ppt_add_chart.py` | Create charts (Bar, Line, Pie...) |
| `ppt_update_chart_data.py` | Refresh chart data |
| `ppt_format_chart.py` | Set title, legend position |
| `ppt_add_table.py` | Create data tables |
| `ppt_format_table.py` | Style table rows/cells |

### Inspection & Discovery
| Tool | Description |
|------|-------------|
| `ppt_get_info.py` | Metadata, version, dimensions |
| `ppt_get_slide_info.py` | Deep slide inspection (shapes, text) |
| `ppt_capability_probe.py` | **Critical**: Detect layouts & placeholders |
| `ppt_extract_notes.py` | Dump speaker notes |
| `ppt_search_content.py` | **New**: Regex search across deck |

### Validation & Pipeline
| Tool | Description |
|------|-------------|
| `ppt_validate_presentation.py`| Policy-based validation |
| `ppt_check_accessibility.py` | WCAG 2.1 compliance check |
| `ppt_export_images.py` | Render slides to PNG/JPG |
| `ppt_export_pdf.py` | Convert deck to PDF |
| `ppt_json_adapter.py` | **New**: Normalize tool output |

---

## 10. 🔧 Troubleshooting

### Common Errors

| Error | Code | Cause | Fix |
|-------|------|-------|-----|
| `ApprovalTokenError` | 4 | Destructive op without token | Generate token with correct scope |
| `ShapeNotFoundError` | 1 | Stale shape index | Run `ppt_get_slide_info.py` |
| `FileLockError` | 3 | File in use | Wait or force-clear `.lock` file |
| `JSONDecodeError` | 5 | Stdout pollution | Check hygiene block in tool |
| `SlideNotFoundError` | 1 | Index out of bounds | Check slide count |

### Recovery
If a presentation becomes corrupted or stuck:
1. **Delete the working copy**: `rm work.pptx`
2. **Re-clone**: `ppt_clone_presentation.py` from source.
3. **Re-run** operations from manifest.

