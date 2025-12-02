# Comprehensive Analysis: PowerPoint AI Agent System Prompt v3.7

## Executive Summary

I have completed a meticulous, multi-dimensional review of your PowerPoint AI Agent system prompt. This is an **impressively comprehensive** specification that demonstrates sophisticated thinking about agent governance, safety, accessibility, and design intelligence. However, I've identified **10 critical issues**, **10 medium-severity concerns**, and **12 minor improvements** that should be addressed before production deployment.

**Overall Assessment**: 🟡 **Near Production-Ready** — Requires targeted fixes to resolve contradictions and gaps.

---

## Analysis Methodology

```
┌─────────────────────────────────────────────────────────────────────┐
│  ANALYSIS DIMENSIONS                                                │
├─────────────────────────────────────────────────────────────────────┤
│  1. Structural Integrity    → Organization, flow, cross-references │
│  2. Technical Accuracy      → Tool specs, syntax, integrations     │
│  3. Governance & Safety     → Hierarchy, tokens, audit trails      │
│  4. Operational Feasibility → LLM executability, cognitive load    │
│  5. Accessibility Compliance→ WCAG coverage, enforcement           │
│  6. Consistency Audit       → Contradictions, version mismatches   │
│  7. Gap Analysis            → Missing scenarios, undefined edges   │
│  8. Pattern Library Review  → Completeness, coherence, usability   │
└─────────────────────────────────────────────────────────────────────┘
```

---

## SECTION 1: Critical Issues (Immediate Action Required)

### 🔴 Critical Issue #1: Version Number Inconsistency

**Symptom**: Multiple conflicting version references throughout the document.

| Location | Version Stated |
|----------|---------------|
| Section VIII Header | "NEW v3.6" |
| Manifest Schema | `$schema: "presentation-architect/manifest-v3.6"` |
| Section V Header | "v3.6" |
| Final Directive | "v3.7" |
| Document Footer | "v3.7" |
| Initialization Declaration (Section X) | "v3.6" |

**Impact**: Agent may report inconsistent versions; creates confusion about authoritative state.

**Recommendation**:
```diff
- All references should be unified to v3.7
- Create explicit version history section:

+ ## VERSION HISTORY
+ | Version | Date | Changes |
+ |---------|------|---------|
+ | v3.5 | 2024-11-01 | Base tool catalog (42 tools) |
+ | v3.6 | 2024-11-15 | Added Visual Pattern Library |
+ | v3.7 | 2024-12-01 | Enhanced governance, checksums |
```

---

### 🔴 Critical Issue #2: Font Size Contradiction

**Symptom**: Direct contradiction between two sections on minimum font size.

**Section 7.1 (Accessibility Requirements)**:
> "Font size | No text below 10pt, prefer ≥12pt"

**Section 6.2 (Typography System)**:
> "**NO EXCEPTIONS** | **12pt** | - | - | **10pt font size no longer permitted**"

**Impact**: Agent cannot determine correct minimum font size; will either over-correct or under-correct.

**Recommendation**:
```diff
# Section 7.1 - Update to match Section 6.2:
- | Font size | No text below 10pt, prefer ≥12pt | Manual verification |
+ | Font size | Minimum 12pt for all text | ppt_format_text + Manual verification |

# Add rationale note:
+ **Note**: 12pt minimum aligns with accessibility best practices for 
+ projected presentations. 10pt is only readable at close distances.
```

---

### 🔴 Critical Issue #3: Tool Catalog Incompleteness

**Symptom**: Document claims 42 tools but never provides complete list. Multiple tools are referenced but never defined.

**Missing Tool Definitions** (referenced but undefined):

| Tool Name | Where Referenced | Purpose |
|-----------|------------------|---------|
| `ppt_json_adapter.py` | Section 2.4, Phase ALL | Schema validation |
| `ppt_set_image_properties.py` | Template 1 (Accessibility) | Alt-text remediation |
| `ppt_replace_image.py` | Template 9.3 | Logo replacement |
| `ppt_add_connector.py` | Pattern 4 (Process Flow) | Shape connections |
| `ppt_format_chart.py` | Section 4.3 | Chart formatting |
| `ppt_format_shape.py` | Appendix A.4 | Shape styling |
| `ppt_export_images.py` | Section 6.2 (Deliver) | Image export |
| `ppt_extract_notes.py` | Section 6.2, 9.1 | Notes extraction |
| `ppt_export_pdf.py` | Section 6.2 | PDF export |

**Impact**: Agent will attempt to use non-existent tools or hallucinate syntax.

**Recommendation**:
```markdown
## APPENDIX C: COMPLETE TOOL CATALOG (42 Tools)

### C.1 File Operations (6 tools)
| Tool | Purpose | Required Args | Risk Level |
|------|---------|---------------|------------|
| ppt_clone_presentation.py | Create working copy | --source, --output | 🟢 Low |
| ppt_create_new.py | New blank presentation | --output, --slides | 🟢 Low |
| ppt_create_from_template.py | From template | --template, --output | 🟢 Low |
| ppt_create_from_structure.py | From JSON spec | --structure, --output | 🟢 Low |
| ppt_merge_presentations.py | Combine decks | --files, --output | 🟡 Medium |
| ppt_export_pdf.py | PDF conversion | --file, --output | 🟢 Low |

### C.2 Inspection Tools (5 tools)
[... complete catalog for all 42 tools ...]
```

---

### 🔴 Critical Issue #4: Approval Token Acquisition Gap

**Symptom**: Token system is well-defined for validation but has no actionable path for token acquisition.

**Current State**:
- Token structure defined ✅
- Token scopes defined ✅
- Enforcement protocol defined ✅
- **Token acquisition mechanism**: ❌ UNDEFINED

**The prompt says**:
> "If destructive operation requested without token → REFUSE"
> "Provide token generation instructions with required scope"

But no actual instructions exist for how a user obtains a valid token.

**Impact**: Agent will refuse operations with no path forward for the user.

**Recommendation**:
```markdown
### 2.3.1 Token Acquisition Workflow

**For Users (Human Workflow)**:
When agent requests approval token:

1. **Review the operation request** in agent's response
2. **Generate token** using your organization's authorization system
3. **Provide token** to agent in format: `--approval-token "apt-YYYYMMDD-NNN"`

**Example Agent Request**:
```
⚠️ APPROVAL REQUIRED

Operation: Delete slide 5 from presentation
Required Token Scope: delete:slide
Risk Level: 🔴 Critical

To proceed, provide an approval token:
  --approval-token "apt-YYYYMMDD-NNN"

Or confirm verbally: "Approved: delete slide 5"
```

**Simplified Mode (for trusted environments)**:
If operating in trusted mode, user may provide verbal confirmation:
- "Approved: [operation description]"
- Agent records approval in manifest with user attribution
```

---

### 🔴 Critical Issue #5: Complexity Score Chicken-and-Egg Problem

**Symptom**: Complexity score requires counting `destructive_ops` and `accessibility_issues` before Phase 0 classification, but these values are only known after analysis.

**Formula**:
```
Score = (slide_count × 0.3) + (destructive_ops × 2.0) + (accessibility_issues × 1.5)
```

**Problem**: 
- `slide_count`: Known immediately ✅
- `destructive_ops`: Only known after planning ❌
- `accessibility_issues`: Only known after accessibility probe ❌

**Impact**: Agent cannot accurately classify requests at intake.

**Recommendation**:
```markdown
### Phase 0: Two-Stage Classification

**Stage 0a: Initial Classification (Pre-Analysis)**
Score = (estimated_slide_count × 0.3) + (has_destructive_keywords × 3.0)

Destructive keywords: "delete", "remove", "replace all", "mass", "global"

| Initial Score | Classification | Action |
|---------------|----------------|--------|
| < 3.0 | SIMPLE (tentative) | Streamlined workflow |
| 3.0-10.0 | STANDARD (tentative) | Full manifest |
| > 10.0 | COMPLEX (tentative) | Phased delivery |

**Stage 0b: Refined Classification (Post-Discovery)**
After probe and accessibility baseline:
Score = (slide_count × 0.3) + (destructive_ops × 2.0) + (accessibility_issues × 1.5)

May upgrade classification (never downgrade for safety):
- SIMPLE → STANDARD: If destructive ops detected
- STANDARD → COMPLEX: If accessibility issues > 5
```

---

### 🔴 Critical Issue #6: Schema Files Not Provided

**Symptom**: Document references schema files for validation but schemas are not included.

**Referenced Schemas**:
- `ppt_get_info.schema.json`
- `ppt_capability_probe.schema.json`
- `ppt_get_slide_info.schema.json`
- Tool-specific schemas

**Impact**: Agent cannot perform the "MANDATORY" validation steps in Section 2.4.

**Recommendation**: Either:

**Option A**: Include schemas in appendix
```json
// APPENDIX D: JSON SCHEMAS
// D.1 ppt_get_info.schema.json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "type": "object",
  "required": ["tool_version", "schema_version", "presentation_version", "slide_count"],
  "properties": {
    "tool_version": { "type": "string", "pattern": "^\\d+\\.\\d+\\.\\d+$" },
    "schema_version": { "type": "string" },
    "presentation_version": { "type": "string", "pattern": "^[a-f0-9]{16}$" },
    "slide_count": { "type": "integer", "minimum": 0 }
  }
}
```

**Option B**: Make validation conditional
```markdown
### 2.4 JSON Schema Validation Framework

**If schema files available**: Execute full validation pipeline
**If schema files unavailable**: Perform structural validation:
- Verify JSON parses successfully
- Check required fields exist
- Validate data types match expected
```

---

### 🔴 Critical Issue #7: Probe Output Structure Undefined

**Symptom**: Multiple sections depend on probe output structure, but the complete output schema is never defined.

**Sections Depending on Probe Output**:
- Phase 2 (Discovery) expects: `layouts_available`, `theme.colors`, `theme.fonts`
- Phase 3 (Plan) references: `probe.capabilities`
- Design Intelligence references: `theme.colors.accent1`, etc.

**Problem Example** (Section 6.3):
> "Extract theme.colors from probe"
> "Map theme colors to semantic roles: accent1 → primary actions"

But probe output structure is never shown.

**Recommendation**:
```markdown
### 3.1.1 Probe Output Schema

**ppt_capability_probe.py --deep --json** returns:
```json
{
  "tool_version": "1.2.0",
  "schema_version": "probe-v2",
  "probe_timestamp": "2024-12-01T10:30:00Z",
  "probe_type": "full",
  "file": "/path/to/presentation.pptx",
  "presentation_version": "a1b2c3d4e5f6g7h8",
  
  "dimensions": {
    "width_pt": 720,
    "height_pt": 540,
    "aspect_ratio": "4:3"
  },
  
  "slide_count": 12,
  
  "layouts_available": [
    "Title Slide",
    "Title and Content",
    "Two Content",
    "Comparison",
    "Blank",
    "Picture with Caption",
    "Section Header"
  ],
  
  "theme": {
    "name": "Office Theme",
    "colors": {
      "accent1": "#0070C0",
      "accent2": "#ED7D31",
      "accent3": "#FFC000",
      "accent4": "#70AD47",
      "background1": "#FFFFFF",
      "background2": "#F2F2F2",
      "text1": "#000000",
      "text2": "#595959"
    },
    "fonts": {
      "heading": "Calibri Light",
      "body": "Calibri"
    }
  },
  
  "capabilities": {
    "supports_charts": true,
    "supports_tables": true,
    "supports_smartart": false,
    "supports_3d": false
  }
}
```

---

### 🔴 Critical Issue #8: Pattern Organization Confusion

**Symptom**: Patterns are numbered 1-15 but introduced in mixed order, with groups defined mid-section.

**Current Structure**:
```
8.1 Pattern Selection Decision Tree (Groups A-D defined here)
8.2 Pattern 1: Data-Heavy (actually Group B)
8.3 Pattern 2: Executive Summary (Group A)
8.4 Pattern 3: Comparison (Group B)
8.5 Pattern 4: Process Flow (Group D)
8.6 Pattern 5: Image Showcase (Group C)
--- GROUP A HEADER ---
8.7 Pattern 6: Quote Impact
8.8 Pattern 13: Testimonial
8.9 Pattern 15: Q&A Closing
--- GROUP B HEADER ---
8.10 Pattern 10: Financial Summary
...
```

**Problems**:
1. Patterns 1-5 appear before their group headers
2. Patterns within groups are non-sequential (Group A: 2, 6, 13, 15)
3. No Pattern 7, 8, 9 in Group A but they exist in other groups

**Impact**: Agent cannot reliably map content type to pattern using the decision tree.

**Recommendation**:
```markdown
## SECTION VIII: VISUAL PATTERN LIBRARY

### 8.1 Pattern Index

| Pattern ID | Name | Group | Use Case |
|------------|------|-------|----------|
| P-A1 | Executive Summary | A: Narrative | Key points, thesis |
| P-A2 | Quote Impact | A: Narrative | Quotes, testimonials |
| P-A3 | Testimonial | A: Narrative | Customer validation |
| P-A4 | Q&A Closing | A: Narrative | Presentation close |
| P-B1 | Data-Heavy Slide | B: Analytics | Charts, data viz |
| P-B2 | Comparison Slide | B: Analytics | Side-by-side analysis |
| P-B3 | Financial Summary | B: Analytics | KPIs, financial data |
| P-B4 | SWOT Analysis | B: Analytics | Strategic frameworks |
| P-B5 | Risk Matrix | B: Analytics | Risk assessment |
| P-C1 | Image Showcase | C: Visual | Photo-focused slides |
| P-C2 | Technical Detail | C: Visual | Code, architecture |
| P-C3 | Timeline | C: Visual | Milestones, roadmaps |
| P-C4 | Product Showcase | C: Visual | Product marketing |
| P-D1 | Process Flow | D: Structure | Workflows, procedures |
| P-D2 | Team Bio | D: Structure | Personnel highlights |

### 8.2 GROUP A: Narrative & Impact Patterns
[Patterns P-A1 through P-A4 in sequence]

### 8.3 GROUP B: Data & Analytics Patterns
[Patterns P-B1 through P-B5 in sequence]
...
```

---

### 🔴 Critical Issue #9: Race Condition Risk in Shape Index Management

**Symptom**: Shape index management relies on sequential refresh, but no protection against concurrent modifications.

**Current Protocol**:
```bash
# Add shape → Refresh → Modify → Refresh
uv run tools/ppt_add_shape.py ...
uv run tools/ppt_get_slide_info.py ...  # Refresh
uv run tools/ppt_set_z_order.py --shape 7 ...  # Uses refreshed index
```

**Problem**: If another process modifies the file between refresh and use, shape index 7 may reference wrong shape.

**Impact**: Could result in modifying or deleting wrong shapes.

**Recommendation**:
```markdown
### 3.5.1 Shape Index Locking Protocol

**Single-User Mode** (default):
Standard refresh protocol applies.

**Multi-User/Concurrent Mode**:
1. **Acquire lock** before structural operations
2. **Refresh indices** immediately before each shape-targeting operation  
3. **Validate shape identity** using name property when available
4. **Release lock** after operation sequence complete

**Shape Identity Verification**:
```bash
# Prefer shape name over index when available
SHAPE_INFO=$(uv run tools/ppt_get_slide_info.py --file work.pptx --slide 2 --json)
SHAPE_INDEX=$(echo "$SHAPE_INFO" | jq '.shapes[] | select(.name == "OverlayRect") | .index')

# If name not available, verify shape type matches expectation
SHAPE_TYPE=$(echo "$SHAPE_INFO" | jq -r ".shapes[$INDEX].type")
if [ "$SHAPE_TYPE" != "rectangle" ]; then
  echo "⚠️ Shape type mismatch - indices may have shifted"
  # Re-probe and recalculate
fi
```

**File Modification Detection**:
Before each operation, verify presentation_version matches expected:
```bash
CURRENT_VERSION=$(uv run tools/ppt_get_info.py --file work.pptx --json | jq -r '.presentation_version')
if [ "$CURRENT_VERSION" != "$EXPECTED_VERSION" ]; then
  echo "⚠️ Presentation modified externally - aborting for safety"
  exit 5
fi
```
```

---

### 🔴 Critical Issue #10: Accessibility Remediation Workflow Syntax Errors

**Symptom**: Bash scripts in Workflow 1 and 2 (Section 7.4) have syntax issues that would fail in execution.

**Issue in Workflow 1**:
```bash
for issue in $(echo "$MISSING_ALT_IMAGES" | jq -c '.'); do
  SLIDE=$(echo "$issue" | jq -r '.slide')
  # ...
done
```

**Problems**:
1. `jq -c '.'` on an array produces individual objects that break in bash `for`
2. Should use `jq -c '.[]'` to iterate array elements
3. Quoting issues with JSON in shell variables

**Recommendation**:
```bash
# CORRECTED Workflow 1: Missing Alt Text Remediation

# 1. Run accessibility check
uv run tools/ppt_check_accessibility.py --file work.pptx --json > accessibility_report.json

# 2. Extract images without alt text (corrected jq syntax)
ISSUE_COUNT=$(jq '[.issues[] | select(.type == "missing_alt_text")] | length' accessibility_report.json)

# 3. Iterate safely using jq indices
for i in $(seq 0 $((ISSUE_COUNT - 1))); do
  SLIDE=$(jq -r ".issues[$i].slide" accessibility_report.json)
  SHAPE=$(jq -r ".issues[$i].shape" accessibility_report.json)
  
  # Apply remediation template
  uv run tools/ppt_set_image_properties.py \
    --file work.pptx \
    --slide "$SLIDE" \
    --shape "$SHAPE" \
    --alt-text "Descriptive text for image on slide $((SLIDE + 1))" \
    --json
done

# 4. Re-validate
uv run tools/ppt_check_accessibility.py --file work.pptx --json
```

---

## SECTION 2: Medium-Severity Issues

### 🟡 Issue #11: 6×6 Rule Not Enforced by Tools

**Problem**: The 6×6 rule is stated as a constraint but `ppt_add_bullet_list.py` doesn't validate it.

**Current**: Agent must manually count words and bullets.

**Recommendation**: Add explicit validation step:
```markdown
### 4.3.1 6×6 Rule Validation Protocol

Before calling ppt_add_bullet_list.py:

```bash
# Validate 6x6 rule before execution
ITEMS="Item one here,Item two,Item three is longer than expected"
ITEM_COUNT=$(echo "$ITEMS" | tr ',' '\n' | wc -l)
MAX_WORDS=$(echo "$ITEMS" | tr ',' '\n' | while read line; do echo "$line" | wc -w; done | sort -rn | head -1)

if [ "$ITEM_COUNT" -gt 6 ]; then
  echo "⚠️ 6×6 VIOLATION: $ITEM_COUNT items exceeds maximum 6"
  echo "💡 Split across multiple slides or summarize"
  exit 1
fi

if [ "$MAX_WORDS" -gt 6 ]; then
  echo "⚠️ 6×6 VIOLATION: $MAX_WORDS words exceeds maximum 6 per bullet"
  echo "💡 Shorten bullet text or use speaker notes for detail"
  exit 1
fi
```
```

---

### 🟡 Issue #12: Error Recovery Level 2 Lacks Tool Mapping

**Problem**: Level 2 says "Use alternative tool for same goal" but provides no alternative mapping.

**Recommendation**:
```markdown
### 3.4.1 Alternative Tool Mapping

| Primary Tool | Alternative | Use Case |
|--------------|-------------|----------|
| ppt_add_bullet_list.py | ppt_add_text_box.py + manual bullets | When bullet tool fails |
| ppt_add_chart.py | ppt_insert_image.py (chart as image) | When chart rendering fails |
| ppt_set_background.py | ppt_add_shape.py (full-slide shape) | When background tool fails |
| ppt_add_connector.py | ppt_add_shape.py (line) | When connector tool fails |
| ppt_format_table.py | Multiple ppt_format_text.py calls | When table tool fails |
```

---

### 🟡 Issue #13: Fallback Probe Output Undefined

**Problem**: Section 3.1 mentions fallback probe returns "minimal metadata JSON with probe_fallback: true" but format is not specified.

**Recommendation**:
```markdown
### 3.1.2 Fallback Probe Output Schema

When primary probe fails, merge outputs from ppt_get_info and ppt_get_slide_info:

```json
{
  "probe_type": "fallback",
  "probe_fallback": true,
  "probe_timestamp": "2024-12-01T10:30:00Z",
  "fallback_reason": "Primary probe timeout after 3 retries",
  
  "file": "/path/to/presentation.pptx",
  "presentation_version": "a1b2c3d4e5f6g7h8",
  "slide_count": 12,
  
  "layouts_available": null,  // Unknown in fallback mode
  
  "theme": {
    "colors": null,  // Unknown in fallback mode
    "fonts": null    // Unknown in fallback mode
  },
  
  "sampled_slides": [
    { "index": 0, "shape_count": 5 },
    { "index": 1, "shape_count": 8 }
  ],
  
  "capabilities": {
    "full_probe_available": false
  }
}
```

**Fallback Mode Constraints**:
- Cannot use theme-extracted colors → Use canonical fallback palettes
- Cannot verify layout names → Use only "Title and Content" and "Blank"
- Must sample first 3 slides for structure inference
```

---

### 🟡 Issue #14: Speaker Notes Mode Guidance Missing

**Problem**: Two modes available (`--mode overwrite` and `--mode append`) but no guidance on when to use each.

**Recommendation**:
```markdown
### 4.3.6 Speaker Notes Mode Selection

| Scenario | Recommended Mode | Rationale |
|----------|------------------|-----------|
| New slide creation | `overwrite` | Start fresh |
| Adding accessibility descriptions | `append` | Preserve existing content |
| Correcting errors | `overwrite` | Replace incorrect content |
| Adding supplementary info | `append` | Build on existing notes |
| Template-based creation | `overwrite` | Replace placeholder text |

**Default**: Use `append` unless explicitly replacing all notes.
```

---

### 🟡 Issue #15: Position/Size Syntax Support Per Tool Unknown

**Problem**: Section 5.2 shows 4 syntax types (percentage, inches, anchor, grid) but doesn't specify which tools support which.

**Recommendation**:
```markdown
### 5.2.1 Position Syntax Compatibility Matrix

| Tool | Percentage | Inches | Anchor | Grid |
|------|------------|--------|--------|------|
| ppt_add_shape.py | ✅ | ✅ | ❌ | ❌ |
| ppt_add_text_box.py | ✅ | ✅ | ❌ | ❌ |
| ppt_add_chart.py | ✅ | ✅ | ❌ | ❌ |
| ppt_insert_image.py | ✅ | ✅ | ✅ | ❌ |
| ppt_add_table.py | ✅ | ✅ | ❌ | ❌ |

**Default Recommendation**: Use percentage-based positioning for maximum compatibility.
```

---

### 🟡 Issue #16: Dry-Run Coverage Inconsistent

**Problem**: Dry-run required for text replacement but not other risky operations.

**Recommendation**:
```markdown
### 2.5.1 Extended Dry-Run Requirements

| Operation | Dry-Run Required | Rationale |
|-----------|------------------|-----------|
| ppt_replace_text.py | ✅ MANDATORY | Mass text changes |
| ppt_set_background.py --all-slides | ✅ MANDATORY | Global visual change |
| ppt_format_text.py --all-shapes | ✅ RECOMMENDED | Multi-shape formatting |
| ppt_remove_shape.py | ✅ RECOMMENDED | Destructive operation |
| ppt_delete_slide.py | ⚠️ No dry-run (use clone) | Full backup required |
```

---

### 🟡 Issue #17: Template Naming Conflict

**Problem**: "Template 3" refers to both:
- Section 5.3 Accessibility Remediation Template (Complex Visual Description)
- Section 9.3 Workflow Template (Surgical Rebranding)

**Recommendation**: Rename to distinguish:
```markdown
- Accessibility Templates: AT-1, AT-2, AT-3, AT-4, AT-5
- Workflow Templates: WT-1, WT-2, WT-3
```

---

### 🟡 Issue #18: Manifest Schema Evolution Undefined

**Problem**: Schema version changed from 3.6 to 3.7 but no migration guidance.

**Recommendation**:
```markdown
### 3.1.3 Manifest Schema Migration

**v3.6 → v3.7 Changes**:
| Field | v3.6 | v3.7 | Migration |
|-------|------|------|-----------|
| `$schema` | `manifest-v3.6` | `manifest-v3.7` | Update string |
| `diff_summary.checksums` | ❌ | ✅ Added | Add empty object |
| `operations[].pattern_reference` | Optional | Recommended | Add if missing |

**Backward Compatibility**: v3.6 manifests can be read by v3.7 agents.
```

---

### 🟡 Issue #19: Exit Code 3 Overloaded

**Problem**: Exit code 3 covers "Timeout, I/O, network" but these have different retry semantics.

**Recommendation**:
```markdown
### 3.3.1 Refined Exit Codes

| Code | Category | Meaning | Retryable | Max Retries |
|------|----------|---------|-----------|-------------|
| 3 | Timeout | Operation timed out | Yes | 3 |
| 6 | I/O Error | File read/write failed | Maybe | 1 |
| 7 | Network Error | Remote resource unavailable | Yes | 5 |

**Retry Strategy by Code**:
- Code 3 (Timeout): Exponential backoff (2s, 4s, 8s)
- Code 6 (I/O): Single retry after 1s, then escalate
- Code 7 (Network): Linear backoff (5s intervals)
```

---

### 🟡 Issue #20: Chart Type Validation Missing

**Problem**: Chart types use underscores (`line_markers`) but no validation that tool accepts this exact string.

**Recommendation**:
```markdown
### 5.3.1 Chart Type Validation

Before calling ppt_add_chart.py, validate chart type:

```bash
VALID_TYPES="column column_stacked bar bar_stacked line line_markers area pie doughnut scatter"
REQUESTED_TYPE="line_markers"

if [[ ! " $VALID_TYPES " =~ " $REQUESTED_TYPE " ]]; then
  echo "❌ Invalid chart type: $REQUESTED_TYPE"
  echo "Valid types: $VALID_TYPES"
  exit 1
fi
```
```

---

## SECTION 3: Minor Issues and Polish

| # | Issue | Location | Recommendation |
|---|-------|----------|----------------|
| 21 | Duplicate Response Protocol | Section X.2 + Final Directive | Remove duplicate; keep Final Directive version |
| 22 | "auto" height undefined | Position examples | Define: `"auto"` calculates proportional height from width |
| 23 | Percentage quotes inconsistent | `"10%"` vs `10` | Standardize: Always use strings with `%` for percentages |
| 24 | Accessibility report output undefined | Bash scripts use `jq` | Provide sample output schema |
| 25 | Shape type vocabulary missing | `--shape rectangle` | List all valid shape types |
| 26 | Text box vs. bullet list criteria | Content population | Add decision guidance |
| 27 | Footer `--show-number` vs `--show-date` | Section 9.3 | Document when to use each |
| 28 | Grid-based positioning unused | Section 5.2 | Either remove or show example usage |
| 29 | Cognitive group A-D diagram | Section 8.1 | Make groups sequential in subsections |
| 30 | HMAC example security warning | Section 2.3 | Consider removing code entirely |
| 31 | LibreOffice dependency | PDF export | Document installation requirement |
| 32 | Checksum file format | Appendix B.2 | Use standard format (`sha256sum` compatible) |

---

## SECTION 4: Structural Strengths (What Works Well)

### ✅ Excellence Areas

| Aspect | Assessment | Notes |
|--------|------------|-------|
| **Safety Hierarchy** | Excellent | Clear precedence, memorable laws |
| **Audit Trail Design** | Excellent | Comprehensive logging specification |
| **Accessibility Focus** | Excellent | WCAG 2.1 AA throughout |
| **Pattern Library Concept** | Excellent | Reduces hallucination risk |
| **Error Handling Matrix** | Very Good | Clear exit codes and actions |
| **Validation Gates** | Very Good | Multi-stage quality assurance |
| **Design Intelligence** | Very Good | Visual hierarchy, typography, color systems |
| **Phase Structure** | Good | Logical workflow progression |
| **Manifest Schema** | Good | Comprehensive change tracking |

### ✅ Standout Features

1. **Three Inviolable Laws**: Memorable, enforceable, and safety-critical
2. **Probe-Before-Populate**: Prevents layout mismatches
3. **Version Tracking**: Prevents race conditions and conflicts
4. **Remediation Templates**: Concrete fixes for common issues
5. **6×6 Rule**: Enforces content discipline
6. **Overlay Safety Protocol**: Prevents content obscuring

---

## SECTION 5: Recommendations Summary

### Priority Matrix

```
┌─────────────────────────────────────────────────────────────────────┐
│                    PRIORITY MATRIX                                  │
├─────────────────────────────────────────────────────────────────────┤
│                                                                     │
│  URGENT + IMPORTANT          │  IMPORTANT (Schedule)               │
│  (Fix before deployment)     │                                     │
│  ────────────────────────    │  ────────────────────────           │
│  • Version unification (#1)  │  • Complete tool catalog (#3)       │
│  • Font size fix (#2)        │  • Probe schema definition (#7)     │
│  • Token acquisition (#4)    │  • Pattern reorganization (#8)      │
│  • Complexity scoring (#5)   │  • Exit code refinement (#19)       │
│                              │                                     │
├──────────────────────────────┼─────────────────────────────────────┤
│  URGENT (Quick fixes)        │  NICE TO HAVE (Backlog)             │
│  ────────────────────────    │  ────────────────────────           │
│  • Bash syntax fixes (#10)   │  • Alternative tool mapping (#12)   │
│  • Schema availability (#6)  │  • Position syntax matrix (#15)     │
│  • Race condition note (#9)  │  • Template renaming (#17)          │
│                              │  • Minor polish (#21-32)            │
│                              │                                     │
└─────────────────────────────────────────────────────────────────────┘
```

### Implementation Checklist

```markdown
## Pre-Deployment Fixes

### Critical (Block deployment)
- [ ] Unify all version references to v3.7
- [ ] Fix font size contradiction (12pt minimum everywhere)
- [ ] Add token acquisition workflow for users
- [ ] Fix complexity score formula (two-stage approach)
- [ ] Fix bash syntax errors in accessibility workflows
- [ ] Add conditional handling for missing schemas

### High Priority (First sprint)
- [ ] Create complete tool catalog (all 42 tools)
- [ ] Define probe output schema
- [ ] Reorganize pattern library with consistent numbering
- [ ] Add shape identity verification for race condition prevention

### Medium Priority (Second sprint)
- [ ] Add 6×6 rule validation script
- [ ] Create alternative tool mapping table
- [ ] Define fallback probe output schema
- [ ] Add speaker notes mode guidance
- [ ] Create position syntax compatibility matrix
- [ ] Extend dry-run requirements
- [ ] Rename templates to avoid conflicts
- [ ] Add manifest migration guidance
- [ ] Refine exit codes (split code 3)
- [ ] Add chart type validation

### Low Priority (Polish)
- [ ] Remove duplicate response protocol
- [ ] Define "auto" height behavior
- [ ] Standardize percentage syntax
- [ ] Provide accessibility report schema
- [ ] List valid shape types
- [ ] Add text box vs. bullet list guidance
- [ ] Document footer flags
- [ ] Remove or document grid positioning
- [ ] Remove HMAC code example
- [ ] Document LibreOffice dependency
- [ ] Standardize checksum file format
```

---

## SECTION 6: Revised Sections (Ready to Copy)

### 6.1 Corrected Section 7.1 (Accessibility)

```markdown
### 7.1 WCAG 2.1 AA Mandatory Checks

| Check | Requirement | Tool | Remediation |
|-------|-------------|------|-------------|
| Alt text | All images must have descriptive alt text | ppt_check_accessibility | AT-1: ppt_set_image_properties |
| Color contrast | Text ≥4.5:1 (body), ≥3:1 (large/bold) | ppt_check_accessibility | AT-2: ppt_format_text |
| Reading order | Logical tab order for screen readers | ppt_check_accessibility | AT-4: Shape repositioning |
| Font size | **Minimum 12pt for all text** | ppt_check_accessibility + Manual | AT-5: ppt_format_text |
| Color independence | Information not conveyed by color alone | Manual verification | Add patterns/labels/text |
```

### 6.2 Corrected Complexity Score (Section IV, Phase 0)

```markdown
### Phase 0: REQUEST INTAKE & CLASSIFICATION

**Two-Stage Classification Protocol**

**Stage A: Initial Classification (Pre-Analysis)**
```
Initial Score = (estimated_slides × 0.3) + (destructive_keywords × 3.0)
```

Destructive keywords detected in request:
- "delete", "remove" → +3.0 each
- "replace all", "mass", "global" → +3.0 each
- "merge", "combine" → +1.5 each

| Initial Score | Tentative Class | Workflow |
|---------------|-----------------|----------|
| < 3.0 | 🟢 SIMPLE | Streamlined |
| 3.0–10.0 | 🟡 STANDARD | Full manifest |
| > 10.0 | 🔴 COMPLEX | Phased delivery |

**Stage B: Refined Classification (After Discovery)**
```
Final Score = (slide_count × 0.3) + (destructive_ops × 2.0) + (accessibility_issues × 1.5)
```

Classification may only **upgrade** (never downgrade):
- SIMPLE → STANDARD if destructive_ops > 0
- STANDARD → COMPLEX if accessibility_issues > 5 or slide_count > 20
```

---

## Final Assessment

| Dimension | Score | Notes |
|-----------|-------|-------|
| Comprehensiveness | 9/10 | Exceptional coverage; minor gaps in tool catalog |
| Consistency | 6/10 | Version conflicts and contradictions need resolution |
| Executability | 7/10 | Mostly executable; some syntax errors and undefined references |
| Safety Design | 9/10 | Excellent governance framework |
| Accessibility | 9/10 | Strong WCAG focus with minor enforcement gaps |
| Maintainability | 7/10 | Good structure; needs version management |
| LLM Suitability | 8/10 | Clear patterns reduce hallucination; complexity may challenge smaller models |

### Overall Verdict

**🟡 Near Production-Ready**

This is an exceptionally well-designed system prompt that demonstrates sophisticated thinking about AI agent governance, safety, and design intelligence. The critical issues identified are all **resolvable** without fundamental restructuring. After addressing the version inconsistencies, font size contradiction, token acquisition gap, and tool catalog completeness, this prompt will be **production-ready**.

The Visual Pattern Library is a particularly innovative feature that should significantly reduce hallucination risk for less capable LLMs by providing concrete, copy-paste execution patterns.

---

**Ready to assist with implementing any of these recommendations.** Would you like me to:

1. Draft the complete tool catalog (Appendix C)?
2. Create the probe output schema specification?
3. Reorganize the pattern library with consistent numbering?
4. Develop the two-stage complexity scoring implementation?
5. Any other specific section revision?
