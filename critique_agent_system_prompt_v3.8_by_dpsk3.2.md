# Comparative Analysis of PowerPoint Agent System Prompts

## Overview

We compare two versions of the AI Presentation Architect system prompt:  

- [`AGENT_SYSTEM_PROMPT.md`](AGENT_SYSTEM_PROMPT.md) (versions 3.6/3.7)  
- [`AGENT_SYSTEM_PROMPT_enhanced.md`](AGENT_SYSTEM_PROMPT_enhanced.md) (version 3.8)

The enhanced version builds upon the original with numerous refinements, new features, and a more organised structure. This analysis examines improvements, regressions, and overall effectiveness.

---

## 1. Structural Organisation

**Original**  
- Sections I–XI, plus Appendices A and B.  
- No top‑level table of contents; some sections contain deeply nested subsections.  
- Visual Pattern Library includes long command examples interleaved with descriptive text.

**Enhanced**  
- Adds a **Quick Reference** table at the beginning, including version history and a section map.  
- Same main section numbering (I–XI) but with clearer sub‑section numbering (e.g., 2.3.1, 3.1.1).  
- Appendices expanded to C, D, E (Complete Tool Catalog, JSON Schemas, Version History).  
- Many subsections restructured for logical flow (e.g., probe schemas moved into Operational Resilience).

**Verdict**: Enhanced is better organised, making it easier for the AI to locate information.

---

## 2. Governance and Safety

Both contain the core safety laws, approval tokens, versioning, and audit trails. Enhancements in v3.8:

- **Token Acquisition Workflow** (2.3.1) – Provides explicit steps for the AI to request approval (verbal, token, scope) and records approvals in the manifest. The original only described token structure without interaction guidelines.  
- **Extended Dry‑Run Requirements** (2.5.1) – Clarifies which operations mandate a dry‑run and gives a script template.  
- **Schema Availability Handling** (2.4.1) – Defines fallback validation when schema files are missing, improving resilience.  
- **Version Mismatch Response** (2.6) – Gives a clear action plan when presentation versions diverge.

These additions make governance more actionable and robust.

---

## 3. Operational Resilience

### 3.1 Probe Handling

- **Original**: Basic probe protocol with retries and a fallback sequence.  
- **Enhanced**: Adds full **Probe Output Schema** (3.1.1) and **Fallback Probe Schema** (3.1.2), including examples. A table of fallback constraints is also provided. This gives the AI a precise expectation of the JSON structure, reducing ambiguity.

### 3.2 Preflight Checklist

- **Enhanced** includes a **bash script template** (3.2) that can be directly adapted. Original only had a JSON description.

### 3.3 Error Handling

- **Original**: Simple exit code table (0–5).  
- **Enhanced**: Expanded exit codes (0–7) with detailed categories, retry strategies, and a structured error response example. This enables more nuanced recovery.

### 3.4 Error Recovery

- **Enhanced** adds an **Alternative Tool Mapping** (3.4.1), listing fallback tools for common failures (e.g., use [`ppt_add_text_box.py`](tools/ppt_add_text_box.py) if bullet list fails). Original had no such mapping.

### 3.5 Shape Index Management

- **Original**: Basic protocol to refresh indices after structural changes.  
- **Enhanced**: Adds **Shape Index Locking Protocol** (3.5.1) with verification scripts, **Shape Identity Verification**, and **File Modification Detection**. These prevent race conditions and external interference.

**Verdict**: Significant improvements in resilience, with many new practical scripts.

---

## 4. Workflow Phases

### Phase 0: Classification

- **Original**: Single complexity score formula.  
- **Enhanced**: **Two‑Stage Classification** (Initial and Refined) with upgrade rules. This allows early risk assessment and later adjustment based on probe data, reducing over‑ or under‑classification.

### Phase 2: Discovery

- Both include content‑to‑visualization decision trees. Enhanced explicitly maps content types to **Pattern IDs** (e.g., P‑B1), fostering pattern‑driven execution.

### Phase 3: Plan

- The Change Manifest schema in **Enhanced** includes `manifest_version`, a `classification` object, `probe_summary`, and a `patterns_used` array. It also adds `validation_policy` and `approval_records`. These fields improve traceability.

### Phase 4: Create

- **Enhanced** introduces a **6×6 Rule Validation Script** (4.3.1) that checks bullet count and word length before insertion, preventing rule violations.  
- **Speaker Notes Mode Selection** (4.3.2) clarifies when to use `overwrite` vs `append`.  
- The original had similar guidelines but not as codified.

### Phase 5: Validate

- **Enhanced** replaces the original’s generic remediation templates with concrete **Accessibility Remediation Templates AT‑1 to AT‑5**, each with complete bash scripts. These are far more actionable (e.g., AT‑1 fixes missing alt text by iterating over `accessibility_report.json`).

### Phase 6: Deliver

- Both are similar; **Enhanced** references the checksum generation from Appendix B.

**Verdict**: Enhanced phases are more detailed and provide ready‑to‑use scripts, reducing guesswork.

---

## 5. Tool Ecosystem

- **Original**: Section V listed all 42 tools in a single large table, followed by a syntax reference.  
- **Enhanced**: Splits the information:  
  - Section V gives a high‑level overview with domain‑specific tables, **Footer Flags Reference**, **Text Box vs Bullet List Decision**, and a **Position Syntax Compatibility Matrix**.  
  - Appendix C enumerates every tool for completeness.

This separation improves readability while retaining completeness.

---

## 6. Design Intelligence

- Both cover visual hierarchy, typography, color, spacing, content density, and overlays.  
- **Enhanced** emphasises the **12‑pt minimum font size** with a bold “CRITICAL” notice and exception process, aligning with accessibility best practices. The original also adopted 12‑pt but did not highlight it as prominently.  
- The color system and layout shortcuts are identical.

---

## 7. Accessibility

- Both define WCAG checks and alt‑text best practices.  
- **Enhanced** moves the remediation scripts into Phase 5 and labels them AT‑1…AT‑5, making them easier to find during validation.  
- The original had similar scripts but spread across Section VII; the consolidation is an improvement.

---

## 8. Visual Pattern Library

This is a major difference.

- **Original**: Contains 15 patterns (1–15) with **detailed bash command examples**, including exact positioning, sizes, colors, and sequential steps.  
  Example (Pattern 1 – Data‑Heavy Slide):
  ```bash
  uv run tools/ppt_add_slide.py ...
  uv run tools/ppt_set_title.py ...
  uv run tools/ppt_add_chart.py ...
  uv run tools/ppt_add_notes.py ...
  ```
  Each command includes full `--position` and `--size` JSON strings.

- **Enhanced**: Groups patterns into A‑D and assigns IDs (P‑A1, P‑B1, …). However, the pattern descriptions are **abstract**, listing only the tool names and high‑level steps without concrete parameters.  
  Example (P‑B1):
  ```
  1. ppt_add_slide (Layout: "Title and Content")
  2. ppt_add_chart (Type: from decision tree)
  3. ppt_format_chart (Legend: Bottom, Title: Visible)
  4. ppt_add_notes (MANDATORY: Detailed data description for accessibility)
  ```

- **Comparison**: The original’s explicit commands leave little room for interpretation, ensuring consistent slide layout. The enhanced version expects the AI to decide positioning, sizes, and colors based on the Design Intelligence section. While this may increase flexibility, it also introduces variability and could lead to deviations from intended standards unless the AI is meticulously prompted elsewhere. The loss of concrete examples is a **regression** in terms of deterministic guidance.

- The **Workflow Templates** (WT‑1, WT‑2, WT‑3) remain concrete in both versions, so the abstraction is limited to the slide‑design patterns.

**Recommendation**: Retain the detailed command examples for each visual pattern, or at least provide standard positioning shortcuts (e.g., `full_width`, `left_column`) that the AI should use.

---

## 9. Response Protocol

- Both define the initialization declaration and final report format.  
- **Enhanced** integrates the two‑stage classification into the declaration and adds a `Pattern Intelligence` line. The final report table includes a **Pattern** column, emphasising pattern usage tracking.

---

## 10. Absolute Constraints

- Both list NEVER/ALWAYS rules. **Enhanced** adds several important items:  
  - *Skip complexity scoring in Phase 0*  
  - *Deviate from Visual Pattern Library for standard use cases*  
  - *Skip accessibility remediation templates when issues are found*  
  This makes the constraints more comprehensive.

---

## 11. Appendices

### Appendix A – Tool Argument Registry

- **Original** contained:  
  - *Critical Tool Argument Validation Rules* (table)  
  - *Critical Validation Patterns* (5 copy‑paste bash scripts: layout validation, file path, slide index, JSON, shape refresh)  
  - *Common Error Patterns & Fixes* (table with error messages and resolutions)  
  - *Tool Dependency Chain Reference*  

- **Enhanced** contains:  
  - *Critical Tool Argument Validation Rules* (similar table)  
  - *Critical Validation Patterns* – only two: Chart Type Validation and JSON Argument Validation.  
  - *Tool Dependency Chain Reference* (similar)  
  - **Missing**: Layout, file path, slide index, and shape refresh validation scripts; also missing the **Common Error Patterns & Fixes** table.

The omitted scripts are partially covered elsewhere (preflight, shape index management), but the concise troubleshooting table was a valuable quick reference. Its absence is a **regression**.

### Appendix B – Delivery Package

- Both describe the package contents and checksum generation; **Enhanced** adds the checksum commands.

### New Appendices in Enhanced

- **Appendix C**: Complete Tool Catalog – useful as a lookup.  
- **Appendix D**: JSON Schemas – defines [`ppt_get_info.schema.json`](schemas/ppt_get_info.schema.json) and [`ppt_capability_probe.schema.json`](schemas/ppt_capability_probe.schema.json). Helpful for the AI to understand expected structure.  
- **Appendix E**: Version History – provides context.

These additions are beneficial.

---

## Summary of Improvements in Enhanced v3.8

- More structured document with quick reference and versioning.  
- Token acquisition workflow and extended dry‑run requirements.  
- Probe schemas and fallback handling.  
- Refined error handling with new exit codes and retry strategies.  
- Alternative tool mapping for graceful degradation.  
- Shape index locking and file modification detection.  
- Two‑stage classification for better risk assessment.  
- 6×6 validation script and speaker notes mode guidance.  
- Concrete accessibility remediation scripts (AT‑1…AT‑5).  
- Tool ecosystem overview with decision matrices.  
- Stronger emphasis on 12‑pt font minimum.  
- New appendices with tool catalog and JSON schemas.

## Regressions or Areas for Improvement

1. **Visual Pattern Library** lost the detailed command examples, reducing determinism.  
2. **Common Error Patterns & Fixes** table removed from Appendix A, which provided quick diagnostics.  
3. **Critical Validation Patterns** reduced from five to two; some useful scripts (e.g., slide index validation) are now scattered.  
4. The **Probe Output Schema** is included but could be referenced more directly in the discovery phase.

## Recommendations

- **Reinstate the detailed command examples** for each visual pattern, possibly merging them with the abstract steps. This ensures the AI has concrete templates to follow.  
- **Bring back the Common Error Patterns & Fixes** table, as it is a valuable troubleshooting aid.  
- **Add a “Pattern Quick Reference”** that maps pattern IDs to the original command snippets for easy copy‑paste.  
- **Consider including the validation scripts** (layout, file path, slide index) as part of the preflight or shape management sections, but keep them easily accessible in an appendix.  
- **Ensure consistency** in version numbering (the original footer says v3.7 while the body says v3.6; enhanced says v3.8). Stick to a single version per file.

## Conclusion

The enhanced system prompt (v3.8) is a substantial upgrade in organisation, resilience, and actionable guidance. It introduces many new features that make the AI’s operation more robust and user‑friendly. However, it would benefit from re‑incorporating the concrete pattern commands and error‑fix tables from the original to preserve deterministic execution. With those additions, it would represent the best of both worlds.
