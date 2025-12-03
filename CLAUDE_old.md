# 📚 AGENT SYSTEM REFERENCE

> **PowerPoint Agent Tools** - Enabling AI agents to engineer presentations with precision, safety, and visual intelligence.

**Document Version:** 1.1.0  
**Project Version:** 3.1.0  
**Last Updated:** November 2025

---

## 🚀 Quick Start Guide

**Get up and running in 60 seconds**

```bash
# 1. Clone the repository
git clone https://github.com/anthropics/powerpoint-agent-tools.git
cd powerpoint-agent-tools

# 2. Install dependencies (uv recommended)
uv pip install -r requirements.txt
uv pip install -r requirements-dev.txt

# 3. Create a test presentation
uv run tools/ppt_create_new.py --output test.pptx --json
uv run tools/ppt_add_slide.py --file test.pptx --layout "Blank" --json

# 4. Inspect the presentation
uv run tools/ppt_get_info.py --file test.pptx --json

# 5. Add a semi-transparent overlay shape
uv run tools/ppt_add_shape.py --file test.pptx --slide 0 --shape rectangle \
  --position '{"left":"0%","top":"0%"}' --size '{"width":"100%","height":"100%"}' \
  --fill-color "#FFFFFF" --fill-opacity 0.15 --json

# 6. Run tests to verify installation
pytest tests/ -v
```

### 🔑 Key Concepts to Remember

| Concept | Rule | Why It Matters |
|---------|------|----------------|
| 🔒 **Clone Before Edit** | Never modify source files directly | Prevents accidental data loss |
| 🔍 **Probe Before Operate** | Always inspect slide structure first | Avoids layout guessing errors |
| 🔄 **Refresh Indices** | Re-query after structural operations | Shape indices shift after changes |
| 📊 **JSON-First I/O** | All tools output structured JSON | Enables machine parsing |
| ♿ **Accessibility First** | Validate contrast and alt text | Creates inclusive presentations |

---

## ✨ What's New in v3.1.0

| Feature | Description |
|---------|-------------|
| 🎨 **Opacity Support** | New `fill_opacity` and `line_opacity` parameters (0.0-1.0) |
| 📦 **Overlay Mode** | `--overlay` preset for quick background overlays |
| 🔧 **format_shape() Fix** | Now properly supports transparency via XML manipulation |
| ⚠️ **Deprecation** | `transparency` parameter deprecated (use `fill_opacity` instead) |
| 📋 **Enhanced Returns** | Core methods return detailed `styling` and `changes_detail` dicts |

---

## 📋 Table of Contents

1. [🎯 Project Identity & Mission](#1--project-identity--mission)
2. [🏗️ Architecture Overview](#2-️-architecture-overview)
3. [🏛️ Design Philosophy](#3-️-design-philosophy)
4. [🛠️ Programming Model](#4-️-programming-model)
5. [📏 Code Standards](#5--code-standards)
6. [⚠️ Critical Patterns & Gotchas](#6-️-critical-patterns--gotchas)
7. [🧪 Testing Requirements](#7--testing-requirements)
8. [📤 Contribution Workflow](#8--contribution-workflow)
9. [📖 Quick Reference](#9--quick-reference)
10. [🔧 Troubleshooting](#10--troubleshooting)
11. [📋 Appendix: Cheat Sheet](#11--appendix-cheat-sheet)

---

## 1. 🎯 Project Identity & Mission

### Core Mission

**"Enabling AI agents to engineer presentations with precision, safety, and visual intelligence"**

PowerPoint Agent Tools is a suite of **39 stateless CLI utilities** designed for AI agents to programmatically create, modify, and validate PowerPoint (`.pptx`) files.

### Key Features

| Feature | Description | Benefit |
|---------|-------------|---------|
| **Stateless Architecture** | Each tool call is independent | Reliable in distributed environments |
| **Atomic Operations** | Open → Modify → Save → Close | Predictable, recoverable workflows |
| **Design Intelligence** | Typography, color theory, density rules | Professional outputs |
| **Accessibility First** | WCAG 2.1 compliance checking | Inclusive presentations |
| **JSON-First I/O** | Structured machine-readable output | Easy AI integration |
| **Clone-Before-Edit** | Automatic file safety | Zero risk to source materials |

### Target Audience

- **AI Presentation Architects** — LLM-based agents that generate/modify presentations
- **Automation Engineers** — Building CI/CD pipelines for report generation
- **Human Developers** — Creating presentation automation workflows
- **Accessibility Specialists** — Ensuring WCAG compliance

### Compatibility Matrix

| Component | Minimum Version | Recommended | Notes |
|-----------|-----------------|-------------|-------|
| Python | 3.8 | 3.10+ | Type hints require 3.8+ |
| python-pptx | 0.6.21 | Latest | Core dependency |
| PowerPoint | 2016 | 2019+ | For viewing output |
| LibreOffice | 7.0 | 7.4+ | For PDF export |
| uv | 0.1.0 | Latest | Package manager |

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
           │ ppt_add_      │ │ ppt_get_      │ │ ppt_validate_ │
           │ shape.py      │ │ slide_info.py │ │ presentation  │
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
                    │   • PowerPointAgent class       │
                    │   • All XML manipulation        │
                    │   • File locking                │
                    │   • Position/Size resolution    │
                    │   • Color helpers               │
                    └─────────────────────────────────┘
                                     │
                                     ▼
                    ┌─────────────────────────────────┐
                    │          python-pptx            │
                    │      (Underlying Library)       │
                    └─────────────────────────────────┘
```

### Directory Structure

```
powerpoint-agent-tools/
├── core/
│   ├── __init__.py                    # Public API exports
│   ├── powerpoint_agent_core.py       # THE HUB - all core logic
│   └── strict_validator.py            # JSON Schema validation
├── tools/                             # THE SPOKES - 37+ CLI utilities
│   ├── ppt_add_shape.py
│   ├── ppt_get_info.py
│   ├── ppt_capability_probe.py
│   └── ... (34+ more tools)
├── schemas/                           # JSON Schemas for validation
│   ├── manifest.schema.json
│   └── tool_output_schemas/
├── tests/                             # Comprehensive test suite
│   ├── test_core.py
│   ├── test_shape_opacity.py
│   ├── conftest.py
│   └── assets/
├── AGENT_SYSTEM_PROMPT.md             # System prompt for AI agents
├── CONTRIBUTING_TOOLS.md              # Tool creation guide
└── requirements.txt                   # Dependencies
```

### Key Components

| Component | Location | Responsibility |
|-----------|----------|----------------|
| **PowerPointAgent** | `core/powerpoint_agent_core.py` | Context manager class; all operations |
| **CLI Tools** | `tools/ppt_*.py` | Thin wrappers; argparse + JSON output |
| **Strict Validator** | `core/strict_validator.py` | JSON Schema validation with caching |
| **PathValidator** | `core/powerpoint_agent_core.py` | Security-hardened path validation |
| **Position/Size** | `core/powerpoint_agent_core.py` | Resolve %, inches, anchor, grid |
| **ColorHelper** | `core/powerpoint_agent_core.py` | Hex parsing, contrast calculation |

### Validation Module

The `strict_validator.py` module provides production-grade JSON Schema validation:

**Supported Drafts:** Draft-07, Draft-2019-09, Draft-2020-12

**Custom Format Checkers:**

| Format | Description | Example |
|--------|-------------|---------|
| `hex-color` | Validates #RRGGBB | `#0070C0` |
| `percentage` | Validates N% | `50%` |
| `file-path` | Valid path string | `/path/to/file` |
| `absolute-path` | Absolute path only | `/absolute/path` |
| `slide-index` | Non-negative integer | `0`, `5` |
| `shape-index` | Non-negative integer | `0`, `3` |

**Usage:**

```python
from core.strict_validator import validate_against_schema, validate_dict

# Simple validation (raises on error)
validate_against_schema(data, "schemas/manifest.schema.json")

# Rich validation (returns result object)
result = validate_dict(data, schema_path="schemas/manifest.schema.json")
if not result.is_valid:
    for error in result.errors:
        print(f"{error.path}: {error.message}")
```

---

## 3. 🏛️ Design Philosophy

### The Four Pillars

```
┌─────────────────────────────────────────────────────────────┐
│                      DESIGN PILLARS                          │
├──────────────┬──────────────┬──────────────┬────────────────┤
│  STATELESS   │    ATOMIC    │  COMPOSABLE  │   ACCESSIBLE   │
├──────────────┼──────────────┼──────────────┼────────────────┤
│ Each call    │ Open→Modify  │ Tools can be │ WCAG 2.1       │
│ independent  │ →Save→Close  │ chained      │ compliance     │
├──────────────┼──────────────┼──────────────┼────────────────┤
│ No memory of │ One action   │ Output feeds │ Alt text,      │
│ previous     │ per call     │ next input   │ contrast,      │
│ calls        │              │              │ reading order  │
└──────────────┴──────────────┴──────────────┴────────────────┘
                              │
                              ▼
                    ┌──────────────────┐
                    │  VISUAL-AWARE    │
                    │                  │
                    │ Typography scales│
                    │ Color theory     │
                    │ Content density  │
                    │ Layout systems   │
                    └──────────────────┘
```

### Core Principles

| Principle | Implementation | Why It Matters |
|-----------|----------------|----------------|
| **Clone Before Edit** | `ppt_clone_presentation.py` first | Prevents data loss |
| **Probe Before Operate** | `ppt_capability_probe.py --deep` | Avoids guessing errors |
| **JSON-First I/O** | Structured JSON to stdout | Machine parsing |
| **Fail Safely** | Incomplete > corrupted | Production reliability |
| **Refresh After Changes** | Re-query shape indices | Prevents stale references |
| **Accessibility Default** | Built-in WCAG validation | Inclusive outputs |

### The Statelessness Contract

```python
# ✅ CORRECT: Each call is independent and self-contained
with PowerPointAgent(filepath) as agent:
    agent.open(filepath)
    agent.add_shape(
        slide_index=0,
        shape_type="rectangle",
        position={"left": "10%", "top": "10%"},
        size={"width": "20%", "height": "20%"},
        fill_color="#0070C0",
        fill_opacity=0.15
    )
    agent.save()
# File is closed, lock released, no state retained

# ❌ WRONG: Assuming state persists between calls
agent.add_shape(...)  # Will fail - no file open
```

**Why Statelessness Matters:**

1. AI agents may lose context between calls
2. Prevents race conditions in parallel execution
3. Enables pipeline composition
4. Simplifies error recovery
5. Makes the system predictable and deterministic

---

## 4. 🛠️ Programming Model

### Adding a New Tool — Template

```python
#!/usr/bin/env python3
"""
PowerPoint [Action] [Object] Tool v3.x.x
[One-line description of tool purpose]

Author: PowerPoint Agent Team
License: MIT
Version: 3.x.x

Usage:
    uv run tools/ppt_[verb]_[noun].py --file deck.pptx [args] --json

Exit Codes:
    0: Success
    1: Error (check error_type in JSON for details)
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any, Optional

# 1. PATH SETUP (required for imports without package install)
sys.path.insert(0, str(Path(__file__).parent.parent))

# 2. IMPORTS FROM CORE
from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
    ShapeNotFoundError,
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.x.x"

# ============================================================================
# MAIN LOGIC FUNCTION
# ============================================================================

def do_action(
    filepath: Path,
    slide_index: int,
    # ... other typed parameters
) -> Dict[str, Any]:
    """
    Perform the action on the PowerPoint file.
    
    Args:
        filepath: Path to PowerPoint file (absolute path required)
        slide_index: Target slide index (0-based)
        
    Returns:
        Dict with operation results
        
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If slide index is invalid
        
    Example:
        >>> result = do_action(Path("/path/to/deck.pptx"), slide_index=0)
        >>> print(result["shape_index"])
        5
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Open, operate, save - STATELESS pattern
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # ... perform operations via agent methods ...
        result = agent.some_method(slide_index=slide_index)
        
        agent.save()
    
    # Return structured result (no "status" key - that's added by CLI layer)
    return {
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "tool_version": __version__,
        # ... action-specific fields from result ...
    }

# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Tool description",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    uv run tools/ppt_xxx.py --file deck.pptx --slide 0 --json
        """
    )
    
    # Required arguments
    parser.add_argument(
        "--file", 
        required=True, 
        type=Path, 
        help="PowerPoint file path"
    )
    parser.add_argument(
        "--slide",
        required=True,
        type=int,
        help="Slide index (0-based)"
    )
    
    # Standard arguments
    parser.add_argument(
        "--json", 
        action="store_true", 
        default=True, 
        help="JSON output (default: true)"
    )
    
    args = parser.parse_args()
    
    try:
        result = do_action(filepath=args.file, slide_index=args.slide)
        output = {"status": "success", **result}
        print(json.dumps(output, indent=2))
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible"
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": e.message,
            "error_type": "SlideNotFoundError",
            "details": e.details,
            "suggestion": "Use ppt_get_info.py to check available slides"
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": e.message,
            "error_type": type(e).__name__,
            "details": e.details
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
```

### Data Structures Reference

#### Position Dictionary

**Percentage (Recommended — responsive):**

```json
{"left": "10%", "top": "20%"}
```

**Inches (Absolute — precise positioning):**

```json
{"left": 1.5, "top": 2.0}
```

**Anchor-based (Layout-aware):**

```json
{"anchor": "center", "offset_x": 0, "offset_y": -1.0}
```

**Anchor options:** `top_left`, `top_center`, `top_right`, `center_left`, `center`, `center_right`, `bottom_left`, `bottom_center`, `bottom_right`

**Grid-based (12-column system):**

```json
{"grid_row": 2, "grid_col": 3, "grid_size": 12}
```

#### Size Dictionary

**Percentage (Responsive):**

```json
{"width": "50%", "height": "40%"}
```

**Inches (Fixed):**

```json
{"width": 5.0, "height": 3.0}
```

**Auto (Preserve aspect ratio):**

```json
{"width": "50%", "height": "auto"}
```

#### Color System

```python
# Hex format (with or without #)
"#0070C0"
"0070C0"

# Preset semantic names (if tool supports)
"primary"    # #0070C0 - Main brand color
"accent"     # #ED7D31 - Secondary emphasis  
"success"    # #70AD47 - Positive indicators
"warning"    # #FFC000 - Caution items
"danger"     # #C00000 - Critical errors
```

### Exit Code Convention

| Code | Meaning | Details |
|------|---------|---------|
| `0` | Success | Operation completed |
| `1` | Error | Check `error_type` in JSON output |

> **Note:** For granular error classification, check the `error_type` field in the JSON output. The field contains the exception class name (e.g., `SlideNotFoundError`, `ValueError`) for programmatic handling.

### JSON Output Standards

**Success Response:**

```json
{
  "status": "success",
  "file": "/absolute/path/to/file.pptx",
  "slide_index": 0,
  "shape_index": 5,
  "tool_version": "3.1.0"
}
```

**Warning Response:**

```json
{
  "status": "warning",
  "file": "/absolute/path/to/file.pptx",
  "warnings": [
    "Low contrast ratio detected (3.8:1)"
  ],
  "result": {
    "shape_index": 5
  }
}
```

**Error Response:**

```json
{
  "status": "error",
  "error": "Slide index 5 out of range (0-4)",
  "error_type": "SlideNotFoundError",
  "details": {
    "requested": 5,
    "available": 5
  },
  "suggestion": "Use ppt_get_info.py to check available slides"
}
```

---

## 5. 📏 Code Standards

### Style Requirements

| Aspect | Requirement |
|--------|-------------|
| **Python Version** | 3.8+ |
| **Type Hints** | Mandatory for all function signatures |
| **Docstrings** | Required for modules, classes, functions |
| **Line Length** | 100 characters (soft limit) |
| **Formatting** | `black` with default settings |
| **Linting** | `ruff` with no errors |
| **Imports** | Grouped: stdlib → third-party → local |

### Naming Conventions

| Element | Convention | Example |
|---------|------------|---------|
| **Tool files** | `ppt_<verb>_<noun>.py` | `ppt_add_shape.py` |
| **Core methods** | `snake_case` | `add_shape()` |
| **Private methods** | `_snake_case` | `_set_fill_opacity()` |
| **Constants** | `UPPER_SNAKE_CASE` | `AVAILABLE_SHAPES` |
| **Classes** | `PascalCase` | `PowerPointAgent` |
| **Type aliases** | `PascalCase` | `PositionDict` |

### Documentation Standards

```python
def method_name(
    self,
    required_param: str,
    optional_param: Optional[int] = None
) -> Dict[str, Any]:
    """
    Short one-line description ending with period.
    
    Longer description explaining behavior and edge cases.
    
    Args:
        required_param: Description of parameter
        optional_param: Description with default noted (Default: None)
            
    Returns:
        Dict with the following keys:
            - key1 (str): Description
            - key2 (int): Description
            
    Raises:
        ValueError: When required_param is empty
        SlideNotFoundError: When slide doesn't exist
        
    Example:
        >>> result = agent.method_name("value", optional_param=42)
        >>> print(result["key1"])
        'expected output'
    """
```

---

## 6. ⚠️ Critical Patterns & Gotchas

### 1. The Shape Index Problem

**Problem:** Shape indices are positional and shift after structural operations.

```python
# ❌ WRONG - indices become stale after structural changes
result1 = agent.add_shape(slide_index=0, ...)  # Returns shape_index: 5
result2 = agent.add_shape(slide_index=0, ...)  # Returns shape_index: 6
agent.remove_shape(slide_index=0, shape_index=5)
agent.format_shape(slide_index=0, shape_index=6, ...)  # ❌ Now index 5!

# ✅ CORRECT - re-query after structural changes
result1 = agent.add_shape(slide_index=0, ...)
result2 = agent.add_shape(slide_index=0, ...)
agent.remove_shape(slide_index=0, shape_index=result1["shape_index"])

# IMMEDIATELY refresh indices
slide_info = agent.get_slide_info(slide_index=0)

# Find target shape by characteristics
for shape in slide_info["shapes"]:
    if shape["name"] == "target_shape":
        agent.format_shape(slide_index=0, shape_index=shape["index"], ...)
```

**Operations that invalidate indices:**

| Operation | Effect |
|-----------|--------|
| `add_shape()` | Adds new index at end |
| `remove_shape()` | Shifts subsequent indices down |
| `set_z_order()` | Reorders indices |
| `delete_slide()` | Invalidates all indices on slide |

### 2. The Probe-First Pattern

**Problem:** Template layouts are unpredictable.

```python
# ❌ WRONG - guessing layout names
agent.add_slide(layout_name="Title and Content")  # Might not exist!

# ✅ CORRECT - probe first
with PowerPointAgent(filepath) as agent:
    agent.open(filepath)
    probe_result = agent.get_capabilities(deep=True)
    
    available_layouts = probe_result["layouts"]
    print(f"Available: {available_layouts}")
    
    # Use discovered layout
    agent.add_slide(layout_name=available_layouts[1])
    agent.save()
```

**The Deep Probe Innovation:** The capability probe creates a transient slide in memory to measure actual placeholder geometry, then discards it. This is the only reliable way to know exact positioning.

### 3. The Overlay Pattern

```python
# ✅ Complete overlay workflow for text readability

# 1. Add overlay shape with opacity
result = agent.add_shape(
    slide_index=0,
    shape_type="rectangle",
    position={"left": "0%", "top": "0%"},
    size={"width": "100%", "height": "100%"},
    fill_color="#FFFFFF",
    fill_opacity=0.15  # 15% opaque
)
overlay_index = result["shape_index"]

# 2. IMMEDIATELY refresh indices
slide_info = agent.get_slide_info(slide_index=0)

# 3. Send overlay to back
agent.set_z_order(
    slide_index=0,
    shape_index=overlay_index,
    action="send_to_back"
)

# 4. IMMEDIATELY refresh indices again
slide_info = agent.get_slide_info(slide_index=0)
```

### 4. Opacity vs Transparency

```
OPACITY (Modern - use this):
0.0 ◄──────────────────────────► 1.0
Invisible                    Fully visible

TRANSPARENCY (Deprecated):
1.0 ◄──────────────────────────► 0.0
Invisible                    Fully visible

CONVERSION: opacity = 1.0 - transparency
```

```python
# ✅ MODERN (preferred)
agent.add_shape(fill_color="#0070C0", fill_opacity=0.15)

# ⚠️ DEPRECATED (backward compatible but logs warning)
agent.format_shape(transparency=0.85)  # Converts to fill_opacity=0.15
```

### 5. File Handling Safety

```python
# ✅ ALWAYS use absolute paths
filepath = Path(filepath).resolve()

# ✅ ALWAYS validate existence
if not filepath.exists():
    raise FileNotFoundError(f"File not found: {filepath}")

# ✅ ALWAYS use context manager
with PowerPointAgent(filepath) as agent:
    agent.open(filepath)
    # ... operations ...
    agent.save()

# ✅ ALWAYS clone before editing
agent.clone_presentation(
    source=Path("/source/template.pptx"),
    output=Path("/work/modified.pptx")
)
```

### 6. Presentation Version Tracking

The system tracks a `presentation_version` hash to detect concurrent modifications:

```python
# Version is returned from operations
result = agent.add_shape(...)
version = result.get("presentation_version")  # e.g., "a1b2c3d4"

# Detect external changes
info1 = agent.get_presentation_info()
# ... potential external modification ...
info2 = agent.get_presentation_info()

if info1["presentation_version"] != info2["presentation_version"]:
    print("⚠️ File was modified externally!")
```

### 7. Chart Update Limitations

**⚠️ Important:** `python-pptx` has limited chart update support.

```python
# ❌ RISKY: Updating existing chart data
agent.update_chart_data(slide_index=0, chart_index=0, data=new_data)
# May fail if schema doesn't match exactly

# ✅ PREFERRED: Delete and recreate
agent.remove_shape(slide_index=0, shape_index=chart_index)
agent.add_chart(
    slide_index=0, 
    chart_type="column", 
    data=new_data,
    position={"left": "10%", "top": "20%"},
    size={"width": "80%", "height": "60%"}
)
```

### 8. XML Manipulation (Advanced)

When python-pptx doesn't expose a feature:

```python
from lxml import etree
from pptx.oxml.ns import qn

# Access shape XML
spPr = shape._sp.spPr

# Find or create elements
solidFill = spPr.find(qn('a:solidFill'))
if solidFill is None:
    solidFill = etree.SubElement(spPr, qn('a:solidFill'))

color_elem = solidFill.find(qn('a:srgbClr'))
if color_elem is None:
    color_elem = etree.SubElement(solidFill, qn('a:srgbClr'))

# Set opacity (OOXML scale: 0-100000)
alpha_elem = etree.SubElement(color_elem, qn('a:alpha'))
alpha_elem.set('val', str(int(0.15 * 100000)))  # 15% opacity
```

**OOXML Alpha Scale:** 0 = invisible, 100000 = fully opaque

---

## 7. 🧪 Testing Requirements

### Test Structure

```
tests/
├── test_core.py                  # Core library unit tests
├── test_shape_opacity.py         # Feature-specific tests
├── test_tools/                   # CLI tool integration tests
│   ├── test_ppt_add_shape.py
│   └── ...
├── conftest.py                   # Shared fixtures
├── test_utils.py                 # Helper functions
└── assets/                       # Test files
    ├── sample.pptx
    └── template.pptx
```

### Required Test Coverage

| Category | What to Test |
|----------|--------------|
| **Happy Path** | Normal usage succeeds |
| **Edge Cases** | Boundary values (0, 1, max, empty) |
| **Error Cases** | Invalid inputs raise correct exceptions |
| **Validation** | Invalid ranges/formats rejected |
| **Backward Compat** | Deprecated features still work |
| **CLI Integration** | Tool produces valid JSON |

### Test Pattern

```python
import pytest
from pathlib import Path

@pytest.fixture
def test_presentation(tmp_path):
    """Create a test presentation with blank slide."""
    pptx_path = tmp_path / "test.pptx"
    with PowerPointAgent() as agent:
        agent.create_new()
        agent.add_slide(layout_name="Blank")
        agent.save(pptx_path)
    return pptx_path

class TestAddShapeOpacity:
    """Tests for add_shape() opacity functionality."""
    
    def test_opacity_applied(self, test_presentation):
        """Test shape with valid opacity value."""
        with PowerPointAgent(test_presentation) as agent:
            agent.open(test_presentation)
            result = agent.add_shape(
                slide_index=0,
                shape_type="rectangle",
                position={"left": "10%", "top": "10%"},
                size={"width": "20%", "height": "20%"},
                fill_color="#0070C0",
                fill_opacity=0.5
            )
            agent.save()
        
        # Core method returns dict with styling details
        assert "shape_index" in result
        assert result["styling"]["fill_opacity"] == 0.5
        assert result["styling"]["fill_opacity_applied"] is True
    
    def test_opacity_boundary_zero(self, test_presentation):
        """Test opacity=0.0 (fully transparent)."""
        with PowerPointAgent(test_presentation) as agent:
            agent.open(test_presentation)
            result = agent.add_shape(
                slide_index=0,
                shape_type="rectangle",
                position={"left": "10%", "top": "10%"},
                size={"width": "20%", "height": "20%"},
                fill_color="#0070C0",
                fill_opacity=0.0
            )
            agent.save()
        
        assert result["styling"]["fill_opacity"] == 0.0
    
    def test_opacity_invalid_raises(self, test_presentation):
        """Test that invalid opacity raises ValueError."""
        with PowerPointAgent(test_presentation) as agent:
            agent.open(test_presentation)
            
            with pytest.raises(ValueError) as excinfo:
                agent.add_shape(
                    slide_index=0,
                    shape_type="rectangle",
                    position={"left": "10%", "top": "10%"},
                    size={"width": "20%", "height": "20%"},
                    fill_color="#0070C0",
                    fill_opacity=1.5  # Invalid
                )
            
            assert "must be between 0.0 and 1.0" in str(excinfo.value)
    
    def test_transparency_backward_compat(self, test_presentation):
        """Test deprecated transparency parameter."""
        with PowerPointAgent(test_presentation) as agent:
            agent.open(test_presentation)
            
            # Add a shape first
            add_result = agent.add_shape(
                slide_index=0,
                shape_type="rectangle",
                position={"left": "10%", "top": "10%"},
                size={"width": "20%", "height": "20%"},
                fill_color="#0070C0"
            )
            
            # Format with deprecated parameter
            result = agent.format_shape(
                slide_index=0,
                shape_index=add_result["shape_index"],
                transparency=0.85  # Should convert to fill_opacity=0.15
            )
            agent.save()
        
        assert "transparency_converted_to_opacity" in result["changes_applied"]
        assert result["changes_detail"]["converted_opacity"] == 0.15
```

### Running Tests

```bash
# Run all tests
pytest tests/ -v

# Run specific file
pytest tests/test_shape_opacity.py -v

# Run with coverage
pytest tests/ --cov=core --cov-report=html

# Run parallel (faster)
pytest tests/ -v -n auto

# Stop on first failure
pytest tests/ -v -x
```

---

## 8. 📤 Contribution Workflow

### Before Starting

1. **Read this document** — Understand the architecture
2. **Check existing tools** — Don't duplicate functionality
3. **Review system prompt** — Understand AI agent usage
4. **Set up environment:**

```bash
uv pip install -r requirements.txt
uv pip install -r requirements-dev.txt
```

### PR Checklist

#### Code Quality

- [ ] Type hints on all function signatures
- [ ] Docstrings on all public functions
- [ ] Follows naming conventions
- [ ] `black` formatted
- [ ] `ruff` passes

#### For New Tools

- [ ] File named `ppt_<verb>_<noun>.py`
- [ ] Uses standard template structure
- [ ] Outputs valid JSON to stdout only
- [ ] Exit code 0 for success, 1 for error
- [ ] Validates paths with `pathlib.Path`
- [ ] All exceptions converted to JSON

#### For Core Changes

- [ ] Complete docstring with example
- [ ] Appropriate typed exceptions
- [ ] Documented return Dict structure
- [ ] Backward compatible or deprecation path

#### Testing

- [ ] Happy path tests
- [ ] Edge case tests
- [ ] Error case tests
- [ ] All tests pass: `pytest tests/ -v`

### Common Mistakes

| Mistake | Problem | Fix |
|---------|---------|-----|
| Non-JSON to stdout | Breaks parsing | Use stderr for logs |
| Stale shape indices | Operations fail | Re-query after changes |
| No input validation | Cryptic errors | Validate early |
| Missing context manager | Resource leaks | Always use `with` |
| Hardcoded paths | Platform issues | Use `pathlib.Path` |

---

## 9. 📖 Quick Reference

### Tool Catalog (39 Tools)

| Domain | Tools |
|--------|-------|
| **Creation** | `ppt_create_new`, `ppt_create_from_template`, `ppt_create_from_structure`, `ppt_clone_presentation` |
| **Slides** | `ppt_add_slide`, `ppt_delete_slide`, `ppt_duplicate_slide`, `ppt_reorder_slides`, `ppt_set_slide_layout`, `ppt_set_footer` |
| **Content** | `ppt_set_title`, `ppt_add_text_box`, `ppt_add_bullet_list`, `ppt_format_text`, `ppt_replace_text`, `ppt_add_notes` |
| **Images** | `ppt_insert_image`, `ppt_replace_image`, `ppt_crop_image`, `ppt_set_image_properties` |
| **Shapes** | `ppt_add_shape`, `ppt_format_shape`, `ppt_add_connector`, `ppt_set_background`, `ppt_set_z_order`, `ppt_remove_shape` |
| **Data Viz** | `ppt_add_chart`, `ppt_update_chart_data`, `ppt_format_chart`, `ppt_add_table` |
| **Inspection** | `ppt_get_info`, `ppt_get_slide_info`, `ppt_extract_notes`, `ppt_capability_probe` |
| **Validation** | `ppt_validate_presentation`, `ppt_check_accessibility`, `ppt_export_images`, `ppt_export_pdf` |

### Core Exceptions

| Exception | When Raised | Recovery |
|-----------|-------------|----------|
| `PowerPointAgentError` | Base exception | Handle subclasses |
| `SlideNotFoundError` | Invalid slide index | Check with `ppt_get_info.py` |
| `ShapeNotFoundError` | Invalid shape index | Refresh with `ppt_get_slide_info.py` |
| `LayoutNotFoundError` | Layout doesn't exist | Use probe to discover |
| `ValidationError` | Schema validation failed | Fix input data |

### Key Constants

```python
# Slide dimensions (16:9 widescreen - default)
SLIDE_WIDTH_INCHES = 13.333
SLIDE_HEIGHT_INCHES = 7.5

# Alternative dimensions (4:3 standard)
SLIDE_WIDTH_4_3_INCHES = 10.0
SLIDE_HEIGHT_4_3_INCHES = 7.5

# Content density (6×6 rule)
MAX_BULLETS_PER_SLIDE = 6
MAX_WORDS_PER_BULLET = 6

# Accessibility (WCAG 2.1 AA)
MIN_CONTRAST_RATIO = 4.5
MIN_FONT_SIZE_PT = 10

# Overlay defaults
OVERLAY_OPACITY = 0.15
```

### Common Commands

```bash
# Clone before editing
uv run tools/ppt_clone_presentation.py \
  --source original.pptx --output work.pptx --json

# Probe template capabilities
uv run tools/ppt_capability_probe.py --file work.pptx --deep --json

# Add semi-transparent overlay
uv run tools/ppt_add_shape.py --file work.pptx --slide 0 --shape rectangle \
  --position '{"left":"0%","top":"0%"}' --size '{"width":"100%","height":"100%"}' \
  --fill-color "#FFFFFF" --fill-opacity 0.15 --json

# Refresh shape indices
uv run tools/ppt_get_slide_info.py --file work.pptx --slide 0 --json

# Send overlay to back
uv run tools/ppt_set_z_order.py --file work.pptx --slide 0 \
  --shape 5 --action send_to_back --json

# Validate accessibility
uv run tools/ppt_check_accessibility.py --file work.pptx --json

# Export to PDF
uv run tools/ppt_export_pdf.py --file work.pptx --output work.pdf --json
```

---

## 10. 🔧 Troubleshooting

### Common Errors

| Error | Cause | Fix |
|-------|-------|-----|
| `SlideNotFoundError` | Index out of range | Run `ppt_get_info.py` to check `slide_count` |
| `LayoutNotFoundError` | Layout name wrong | Use probe to discover actual names |
| `ShapeNotFoundError` | Stale index | Refresh with `ppt_get_slide_info.py` |
| `FileNotFoundError` | Path doesn't exist | Use absolute paths |
| `JSONDecodeError` | Malformed JSON | Validate JSON syntax |

### Debugging Tips

1. **Enable verbose logging** — Set `LOG_LEVEL=DEBUG`
2. **Check file permissions** — Ensure read/write access
3. **Validate JSON inputs** — Use online validators
4. **Test with samples** — Start with `assets/sample.pptx`
5. **Check disk space** — Ensure ≥100MB free

### Recovery Commands

```bash
# Restore from backup
cp presentation_backup.pptx presentation.pptx

# Recreate work copy
uv run tools/ppt_clone_presentation.py \
  --source original.pptx --output work.pptx --json

# Validate file integrity
uv run tools/ppt_validate_presentation.py --file work.pptx --json
```

---

## 11. 📋 Appendix: Cheat Sheet

### Essential Commands

```bash
# 🔒 Clone (always first)
ppt_clone_presentation.py --source X.pptx --output Y.pptx

# 🔍 Probe (discover template)
ppt_capability_probe.py --file Y.pptx --deep --json

# 🔄 Refresh (after structural changes)
ppt_get_slide_info.py --file Y.pptx --slide N --json

# 🎨 Overlay (for readability)
ppt_add_shape.py --file Y.pptx --slide N --shape rectangle \
  --position '{"left":"0%","top":"0%"}' \
  --size '{"width":"100%","height":"100%"}' \
  --fill-color "#FFFFFF" --fill-opacity 0.15 --json

# ♿ Validate (before delivery)
ppt_check_accessibility.py --file Y.pptx --json
```

### Five Golden Rules

1. 🔒 **Clone** — Never edit source files
2. 🔍 **Probe** — Discover before guessing
3. 🔄 **Refresh** — Indices shift after changes
4. 📊 **JSON** — stdout is for JSON only
5. ♿ **Validate** — Check accessibility always

### Quick Opacity Reference

| Opacity | Effect | Use Case |
|---------|--------|----------|
| `0.0` | Invisible | Hidden elements |
| `0.15` | Subtle | Text overlays |
| `0.5` | Half | Watermarks |
| `1.0` | Solid | Normal shapes |

---

## 🏁 Final Directive

**You are a Presentation Architect—not a slide typist.**

Your mission: Engineer presentations that communicate with clarity, persuade with evidence, delight with thoughtful design, and remain accessible to all.

**Every slide:** Accessible, aligned, validated, documented.

**Every operation:** Probed, tracked, refreshed, logged.

**Every decision:** Deliberate, documented, reversible.

**Every delivery:** Summarized, validated, recommended.

---

**Document Version:** 1.1.0  
**Project Version:** 3.1.0  
**Last Updated:** November 2025
