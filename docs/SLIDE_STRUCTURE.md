# Slide Structure Documentation

This document provides a detailed breakdown of each slide in the Escalation to Client Details Report PowerPoint presentation.

## Slide Overview

| # | Section | Title | Primary Content |
|---|---------|-------|-----------------|
| 1 | Executive Summary | Title Slide | Client name, tier, period, report date |
| 2 | Executive Summary | Key Metrics Overview | Alerts triaged hero metric |
| 3 | Executive Summary | Period Highlights | Achievement summary |
| 4 | Value Delivered | Cost Avoidance Hero | $7.55M modeled cost avoided |
| 5 | Value Delivered | Value Breakdown | Security Ops, Coverage, Threat Prevention |
| 6 | Value Delivered | Detection Sources | Source breakdown table |
| 7 | Protection Achieved | Performance Trends | MTTR/MTTD/FP% trend chart |
| 8 | Protection Achieved | Industry Comparison | Comparison table |
| 9 | Protection Achieved | Target Metrics | P90, Priority-based MTTR |
| 10 | Threat Landscape | MITRE Tactics | Stacked bar chart |
| 11 | Threat Landscape | Severity Alignment | Sankey diagram |
| 12 | Threat Landscape | Escalation Methods | Auto vs manual breakdown |
| 13 | Threat Landscape | Additional Context | Reserved |
| 14 | Insights | Prioritized Improvement Plan | 3 improvement items |
| 15 | Insights | Operational Insights | Metrics summary |
| 16 | Additional Content | After-Hours Notifications | After-hours breakdown |
| 17 | Additional Content | Response Efficiency | Priority-based response |
| 18 | Additional Content | Collaboration Quality | Participation metrics |
| 19 | Additional Content | Detection Quality | Quality breakdown |
| 20 | Additional Content | Security Outcomes | Comprehensive summary |
| 21 | Forward Direction | Looking Ahead | Next steps, recommendations |

---

## Master Slide Layout (CRITICALSTART Branding)

Per the CRITICALSTART Slide Branding guidelines, all slides follow a master template with:

### Slide Dimensions

```
┌─────────────────────────────────────────────────────────────┐
│ ESCALATION REPORT    CRITICAL START              1         │ ← Header (Footer typography, 9pt)
│                                                             │
│ ┌─────────────────────────────────────────────────────────┐ │
│ │                                                         │ │
│ │  Slide Title (H3 - 48pt)                    [Logo]     │ │
│ │                                                         │ │
│ │                   Content Area                          │ │
│ │              (0.5" margins on sides)                    │ │
│ │                                                         │ │
│ │                                                         │ │
│ └─────────────────────────────────────────────────────────┘ │
│ DECEMBER 2025       ©2025 CRITICAL START           EBR     │ ← Footer (Footer typography, 9pt)
└─────────────────────────────────────────────────────────────┘
                 10 × 5.625 inches (16:9)
```

### Header Elements (Transparent Background)

| Position | Content | Typography |
|----------|---------|------------|
| Left | Short Presentation Title (e.g., "ESCALATION REPORT") | Footer (9pt) |
| Center | "CRITICAL START" | Footer (9pt) |
| Right | Slide number | Footer (9pt) |

**Note:** Title slide (Slide 1) has NO header—footer only.

### Footer Elements (Transparent Background)

| Position | Content | Typography |
|----------|---------|------------|
| Left | Date (e.g., "DECEMBER 2025") | Footer (9pt) |
| Center | Copyright (e.g., "©2025 CRITICAL START") | Footer (9pt) |
| Right | Intent summary (e.g., "EBR") | Footer (9pt) |

---

## Typography Scale (CRITICALSTART Branding)

| Style | Size (pt) | Size (px) | Line Height | Usage |
|-------|-----------|-----------|-------------|-------|
| H1 | 114 | 152 | 152px | Hero/title slides only |
| H2 | 72 | 96 | 104px | Section headers, large metrics |
| H3 | 48 | 64 | 72px | Slide titles |
| H4 | 27 | 36 | 38px | Subheadings |
| H5 | 18 | 24 | 32px | Card titles, bullet text |
| H6 | 12 | 16 | 24px | Labels, small text |
| Paragraph | 12 | 16 | 24px | Body text |
| Footer | 9 | 12 | 16px | Header/footer elements |

**Font Family:** Roboto (Regular, Medium, Bold) with Arial as fallback.

---

## Spacing Constants

| Constant | Value | Usage |
|----------|-------|-------|
| `MARGIN_STANDARD` | 0.5" | Side margins, minimum spacing |
| `MARGIN_CONTENT` | 1.0" | Content area padding |
| `CARD_SPACING` | 0.2" | Space between cards |
| `HEADER_HEIGHT` | 0.4" | Transparent header area |
| `FOOTER_HEIGHT` | 0.35" | Transparent footer area |

---

## Color Usage

| Context | Color | Hex | RGB |
|---------|-------|-----|-----|
| Primary brand | CS_BLUE | #009CDE | (0, 156, 222) |
| Headlines/accents | CS_NAVY | #004C97 | (0, 76, 151) |
| Body text | CS_SLATE | #343741 | (52, 55, 65) |
| Critical alerts | CS_RED | #EF3340 | (239, 51, 64) |
| Warnings/CTAs | CS_ORANGE | #FF6A14 | (255, 106, 20) |
| Comparisons | CS_VIOLET | #702F8A | (112, 47, 138) |

---

## Detailed Slide Specifications

### Slide 1: Title Slide (Special Layout)

**Background:** Blue sweep gradient (#009CDE → #004C97)
**Header:** NONE (title slides have no header)
**Footer:** Yes (date, copyright, intent)

**Elements:**
- Logo (top-left, 1.6" × 0.35")
- Main title: "ESCALATION TO CLIENT DETAILS REPORT" (H1 - 114pt)
- Subtitle: Tier name (H4 - 27pt)
- Client name (H5 - 18pt)
- Period: "August 1-31, 2025 (31 days)" (Paragraph - 12pt)

**Layout:**
```
┌─────────────────────────────────────────────────────────────┐
│ [LOGO]                                                      │
│                                                             │
│ ESCALATION TO CLIENT                                        │  ← H1 (114pt)
│ DETAILS REPORT                                              │
│                                                             │
│ Signature Tier                                              │  ← H4 (27pt)
│ Lennar Corporation                                          │  ← H5 (18pt)
│ August 1-31, 2025 (31 days)                                │  ← Paragraph (12pt)
│                                                             │
│ DECEMBER 2025       ©2025 CRITICAL START           EBR     │  ← Footer
└─────────────────────────────────────────────────────────────┘
```

### Slide 2: Key Metrics Overview

**Header:** Transparent with title/CRITICAL START/slide#
**Title:** "Executive Summary" (H3 - 48pt, navy)

**Elements:**
- Hero metric: Total incidents escalated (H2 - 72pt)
- Subtext: Daily average context (H5 - 18pt)
- Context: Alert triage summary (Paragraph - 12pt)

### Slide 3: Period Highlights

**Header:** Transparent with title/CRITICAL START/slide#
**Title:** "Period Highlights" (H3 - 48pt, navy)

**Elements:**
- 2×2 grid of highlight cards
- Each card: Title (H5) + value text (Paragraph)

### Content Slides Pattern (Slides 4+)

All content slides follow this consistent structure:

```
┌─────────────────────────────────────────────────────────────┐
│ ESCALATION REPORT    CRITICAL START              N         │
│                                                             │
│ Slide Title (H3)                                  [Logo]   │
│                                                             │
│ ┌─────────────────────────────────────────────────────────┐ │
│ │                                                         │ │
│ │              Main Content Area                          │ │
│ │        (Cards, charts, tables, metrics)                 │ │
│ │                                                         │ │
│ └─────────────────────────────────────────────────────────┘ │
│                                                             │
│ DECEMBER 2025       ©2025 CRITICAL START           EBR     │
└─────────────────────────────────────────────────────────────┘
```

---

## Adding New Slides

### Step 1: Define Data Model

Add required fields to `ReportData` in `report_data.py`:

```python
# New Metric Category
new_metric_value: int = 0
new_metric_context: str = ""
```

### Step 2: Create Slide Builder

Use the `setup_content_slide` helper for consistent branding:

```python
def build_new_slide(prs, data):
    """Create the New Slide.
    
    Per CRITICALSTART branding guidelines:
    - All slides have transparent header and footer
    - Uses H1-H6 typography scale
    
    Args:
        prs (Presentation): The presentation object.
        data (ReportData): The report data object.
    """
    # Use helper for consistent header/footer/title
    slide, content_top = setup_content_slide(prs, "New Slide Title")
    
    # Add content below content_top
    # ...
    
    # Add insight callout (optional)
    add_insight_callout(
        slide, "What This Means", 
        "Explanation text...",
        top=prs.slide_height - FOOTER_HEIGHT - Inches(1.0),
        height=Inches(0.8),
        prs=prs
    )
```

### Step 3: Integrate in Main

Add to the `main()` function:

```python
logger.info("  Building new slide...")
build_new_slide(prs, data)
```

---

## Helper Functions

### `setup_content_slide(prs, title_text, include_title=True)`

Creates a new content slide with standard branding:
- Adds transparent header (title, CRITICAL START, slide number)
- Adds transparent footer (date, copyright, intent)
- Adds logo at top right
- Adds slide title using H3 typography
- Returns `(slide, content_top)` tuple

### `add_master_slide_elements(slide, prs, slide_number, include_header, include_footer)`

Adds header and footer elements to any slide.

### `add_master_slide_header(slide, prs, slide_number, include_header)`

Adds header elements only (title, CRITICAL START, slide #).

### `add_master_slide_footer(slide, prs, date_text, include_footer)`

Adds footer elements only (date, copyright, intent).

### `add_insight_callout(slide, title_text, body_text, top, height, prs)`

Creates "What This Means" insight box with consistent styling.

### `create_gradient_background(prs, slide, gradient_type)`

Applies gradient background to slide (used for title slide).

### `add_logo(slide, position, prs)`

Places Critical Start logo at specified position.

---

## Revision History

| Date | Version | Author | Summary |
|------|---------|--------|---------|
| 2025-12-22 | 1.0 | AI (Claude) | Initial slide structure documentation |
| 2025-12-22 | 2.0 | AI (Claude) | Updated for CRITICALSTART branding guidelines (transparent header/footer, typography scale) |
