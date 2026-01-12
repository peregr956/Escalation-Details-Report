# Slide Structure Documentation

This document provides a detailed breakdown of each slide in the Escalation to Client Details Report PowerPoint presentation.

## Narrative Flow

The presentation follows a clear narrative arc designed for C-suite executives and security leadership:

**Current State → Value Delivered → Performance → Threats → Opportunities**

The presentation uses **6 section title cards** to guide the narrative and **13 content slides** for a total of **19 slides**.

## Slide Overview (Final 19-Slide Structure)

| # | Type | Section | Title | Primary Content |
|---|------|---------|-------|-----------------|
| 1 | Content | Opening | Title Slide | Client name, tier, period, report date |
| 2 | **Section Card** | | Executive Summary | "A comprehensive view of your security posture and the value delivered this period" |
| 3 | Content | Executive Summary | Executive Dashboard | 6 key metrics in 2×3 grid (threats, response, closed E2E, alerts, after-hours, FP rate) |
| 4 | Content | Executive Summary | AI Accelerated Security | CORR Platform funnel: Security Events → Potential Threats → Alerts → Response Actions |
| 5 | **Section Card** | | Value Delivered | "Quantifying the business impact: cost avoidance, operational efficiency, and breach prevention" |
| 6 | Content | Value | Value Delivered | Hero cost avoidance + 3 breakdown cards (ops, coverage, prevention) |
| 7 | **Section Card** | | Protection Achieved | "Measurable outcomes that exceed industry standards and reduce your breach risk" |
| 8 | Content | Performance | Critical Start's Performance | Split layout: MTTR by severity cards (left) + trend chart with industry benchmarks (right) |
| 9 | Content | Performance | Response & Detection | Dual-panel: Response Efficiency (3 cards) + Detection Quality (3 cards with CORR context) |
| 10 | **Section Card** | | Threat Landscape | "Understanding the threats we're seeing and how we're adapting to protect your organization" |
| 11 | Content | Threats | Severity Alignment | Sankey diagram showing vendor → CS severity flows with discrete counts (MARQUEE VISUALIZATION) |
| 12 | Content | Threats | Threat & Detection Sources | Split: MITRE tactics chart + Detection sources cards |
| 13 | **Section Card** | | Insights & Continuous Improvement | "How we're evolving our partnership to advance your security maturity" |
| 14 | Content | Insights | Prioritized Improvements | 3 improvement items with priority badges + insight bar |
| 15 | Content | Insights | Operational Coverage | Dual-panel: After-Hours hero + Collaboration metrics |
| 16 | **Section Card** | | Forward Direction | "Strategic recommendations to strengthen your security posture and maximize partnership value" |
| 17 | Content | Closing | Key Takeaways | 4 executive summary bullets + insight bar |
| 18 | Content | Closing | Looking Ahead | Next period targets, strategic focus, partnership |
| 19 | Content | Closing | Contact | Thank you + contact information |

---

## Changes from Previous 15-Slide Structure

### Section Cards Restored
- **6 section title cards** have been restored to guide the narrative flow
- Section cards are gray background slides with section title and narrative subtitle

### Content Redundancy Removed
| Metric | Now Appears In | Removed From |
|--------|---------------|--------------|
| $7.55M Cost Avoided | Value Delivered (slide 5) only | Executive Dashboard, Key Takeaways |
| Closed End-to-End | Executive Dashboard | Security Outcomes (now pipeline) |

### Visual Improvements
| Slide | Change |
|-------|--------|
| Security Outcomes | Converted from bullet list to **infographic pipeline flow** (Inputs → Actions → Outcomes) |
| Executive Dashboard | Added insight bar; replaced Cost Avoided with Closed E2E metric |
| Industry Comparison | Added insight bar |
| Prioritized Improvements | Added insight bar |
| Key Takeaways | Added insight bar; removed exact cost figure |
| Looking Ahead | Added insight bar |

### Insight Bars Added
Every content slide now has a "blue bar" insight callout providing the "bottom line up front":

| Slide | Insight Title | Purpose |
|-------|--------------|---------|
| Executive Dashboard | "Your Security Posture at a Glance" | Summary of dashboard metrics |
| Security Operations Pipeline | "The Pipeline in Action" | Raw alerts → zero breaches message |
| Critical Start's Performance | (existing) | Trend interpretation |
| Industry Comparison | "Outperforming Industry Benchmarks" | Competitive advantage message |
| Response & Detection | (existing) | Detection quality insight |
| Severity Alignment | (existing) | Severity flow interpretation |
| Threat & Detection Sources | (existing) | Sources insight |
| Prioritized Improvements | "Focus Areas for Continuous Improvement" | Improvement priorities |
| Operational Coverage | (existing) | Coverage insight |
| Key Takeaways | "Bottom Line" | Executive summary |
| Looking Ahead | "Strategic Path Forward" | Recommendations summary |

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

### Section Card Layout

```
┌─────────────────────────────────────────────────────────────┐
│                                                             │
│                                                             │
│                  SECTION TITLE                              │ ← H2 (72pt, Navy)
│                                                             │
│    Narrative subtitle describing what this section covers   │ ← Paragraph (18pt, Slate)
│                                                             │
│                                                             │
└─────────────────────────────────────────────────────────────┘
           Gray background (#F5F5F5), centered text
```

### Header Elements (Transparent Background)

| Position | Content | Typography |
|----------|---------|------------|
| Left | Short Presentation Title (e.g., "ESCALATION REPORT") | Footer (9pt) |
| Center | "CRITICAL START" | Footer (9pt) |
| Right | Slide number | Footer (9pt) |

**Note:** Title slide (Slide 1) and Section Cards have NO header.

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
| Positive outcomes | CS_GREEN | #28A745 | (40, 167, 69) |

---

## Content Ownership Matrix

This matrix ensures each key metric appears only once (or in carefully controlled contexts):

| Metric | Primary Slide | Notes |
|--------|--------------|-------|
| $7.55M Cost Avoided | Value Delivered (slide 5) | Hero metric with breakdown |
| Threats Contained | Executive Dashboard | Count and rate |
| Response Advantage % | Executive Dashboard | With MTTR comparison |
| Alerts Triaged | Pipeline Infographic (slide 6) | Part of INPUTS stage |
| Closed End-to-End | Executive Dashboard | Replaced Cost Avoided |
| After-Hours Count | Executive Dashboard + Operational Coverage | Dashboard shows count; Coverage shows breakdown |
| Industry Comparison | Industry Comparison (slide 9) | Dedicated comparison table |

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

### Slide 3: Executive Dashboard

**Title:** "Executive Dashboard" (H3 - 48pt, navy)

**Elements:**
- 2×3 grid of metric cards
- Top row (blue border): Threats Contained, MTTR (with % faster than industry), **Closed End-to-End**
- Bottom row (gray border): Alerts Triaged, After-Hours, False Positive Rate
- Each card: Title (10pt) + Metric (42pt) + Detail (11pt)
- **Note:** Trend arrows removed from metric cards per Jan 2026 feedback. Border color alone indicates status.
- **Insight bar at bottom:** "Your Security Posture at a Glance"

### Slide 4: AI Accelerated Security (CORR Funnel)

**Title:** "AI Accelerated, Human Validated Security" (H3 - 48pt, navy)

**Subtitle:** "CORR Is Our Superpower to Finding the Right Alerts to Prevent Incidents" (14pt, slate)

**Elements:**
- Horizontal funnel visualization with 4 overlapping trapezoid-shaped cards (funnel effect: long end on left)
- Left anchor: Light gray rounded rectangle with "Security Event In-Flow CORR Platform" text (stacked, 3 lines)
- Stage 1 (Orange): 1.2B Security Events
- Stage 2 (Blue): 2m Potential Threats - with "TBR Agent / 99% Resolution" pill badge above
- Stage 3 (Purple): 150k Alerts - with "Case Agent / 92% Consolidation" pill badge above
- Stage 4 (Red): 72k Response Actions - with "SOC / AI Investigation" pill badge above
- Gray chevron arrow at end indicating flow direction
- Cards overlap horizontally by ~15-20% with proper z-index stacking (later cards in front)
- Subtle drop shadows on cards for depth
- Rendered as SVG-based HTML template and inserted as PNG image

**Visual Design:**
- Trapezoid cards with rounded corners (~20px radius)
- Large bold white numbers (prominent, largest element)
- White labels below numbers
- Agent labels: Dark gray text on light gray rounded pill backgrounds

**Colors (Updated):**
| Stage | Color | Hex |
|-------|-------|-----|
| Security Events | Orange | #F7941D |
| Potential Threats | Blue | #0077C8 |
| Alerts | Purple | #6B2D7B |
| Response Actions | Red | #E63946 |
| Left Anchor/Badges | Gray | #E8E8E8 / #F0F0F0 |

### Slide 6: Value Delivered (Hero + Breakdown)

**Title:** "Value Delivered" (H3 - 48pt, navy)

**Elements:**
- Left panel: Hero cost avoidance (~$X.XM, 72pt) with label and subtitle
- Right panel: 3 stacked breakdown cards (Ops, Coverage, Prevention)
- **Expanded methodology footnote** at bottom explaining calculation methodology:
  - Operations = analyst hours × $85/hr avg. rate
  - Coverage = 24/7 monitoring × market SOC rates
  - Prevention = contained threats × Ponemon breach cost model

### Slide 8: Critical Start's Performance (Split Layout with Industry Benchmarks)

**Title:** "Critical Start's Performance" (H3 - 48pt, navy)

**Layout:** Split panel design with metrics on left, trend chart with industry benchmarks on right

**Left Panel - MTTR by Severity:**
- Panel header: "MTTR BY SEVERITY" (navy background, white text)
- 3 stacked metric cards:
  - Critical & High (Vendor): 67 min (red border)
  - Medium & Low (Vendor): 52 min (orange border)
  - P90 All Severities: 87 min (blue border)

**Right Panel - Trend Chart with Industry Benchmarks (enlarged):**
- Performance trends chart placeholder (MTTR/MTTD/FP% over 3 periods)
- **NEW:** Industry benchmark dashed lines for MTTR and MTTD comparison
- Chart height increased to 2.65" for better visibility
- Legend: "MTTR (blue) | MTTD (navy) | FP% (red) | Industry benchmarks (dashed)"
- **Insight bar:** Explains industry benchmark outperformance, not just restating metrics

**Note:** Industry Comparison table (former slide 9) merged into this slide as benchmark lines on chart.

### Slide 9: Response & Detection (Restructured)

**Title:** "Response & Detection Quality" (H3 - 48pt, navy)

**Layout:**
```
┌─────────────────────┐  ┌─────────────────┬─────────────────┐
│   Remediation Rate  │  │ True Positive   │ False Positive  │
│        92.5%        │  │     Rate        │     Rate        │
├─────────────────────┤  ├─────────────────┴─────────────────┤
│ Playbook Automation │  │         Alert Reduction           │
│        86.1%        │  │    2M → 267 (99.99% resolved)     │
├─────────────────────┤  └───────────────────────────────────┘
│  Human Review Rate  │
│        13.1%        │
└─────────────────────┘
```

**Elements:**
- Left half: Response Efficiency (3 vertical cards)
  - **Remediation Rate** (renamed from Containment Rate) - % of alerts with remediation actions
  - Playbook Automation
  - **Human Review Rate** (renamed from Analyst Escalation) - % requiring manual analyst review
- Right half: Detection Quality (2-over-1 layout)
  - Top row: True Positive Rate, False Positive Rate (side by side)
  - Bottom row: **Alert Reduction** (renamed from Threat Reduction) - full width, shows CORR funnel context
- **Insight box:** "Of 2 million potential threats detected, only X alerts required your attention—Critical Start resolved 99.99% before escalation."

### Slide 11: Severity Alignment (Marquee Visualization with Discrete Counts)

**Title:** "Severity Alignment Flow" (H3 - 48pt, navy)

**Layout:** Split panel design with narrative cards on left, Sankey diagram on right

```
┌───────────────┬─────────────────────────────────────────────────┐
│   UPGRADED    │                                                 │
│  12.3% (33)   │                                                 │
│   (value+)    │                                                 │
├───────────────┤             SANKEY DIAGRAM                      │
│  DOWNGRADED   │    (with legend showing counts per severity)    │
│  24.7% (66)   │                                                 │
│   (time+)     │                                                 │
├───────────────┤                                                 │
│   ALIGNED     │                                                 │
│  63.0% (168)  │                                                 │
└───────────────┴─────────────────────────────────────────────────┘
```

**Elements:**
- Left column: 3 stacked narrative cards (~2" wide)
  - **Upgraded** (X.X% (N)) - "Added Value Beyond Vendor Detection" (red border)
  - **Downgraded** (X.X% (N)) - "Analyst Time Returned to Client" (blue border) *(renamed from De-escalated)*
  - **Aligned** (X.X% (N)) - "Vendor Assessment Confirmed" (navy border)
  - **NEW:** Cards now show both percentage AND discrete count (e.g., "12.3% (33)")
- Right area: Larger Sankey chart placeholder (~7" wide, 3.3" tall)
  - **NEW:** Sankey legend includes counts per CS severity level (e.g., "Critical (9)")
- **Note:** Removed "X escalations" chip per Jan 2026 feedback (redundant with left cards)

**Narrative Storytelling:**
- Upgraded alerts = CS provided more value than the vendor's original detection
- Downgraded alerts = CS returned analyst time back to the client
- Aligned alerts = Vendor assessment confirmed by CS analysis

**Note:** This slide is kept standalone as a marquee visualization.

### Slide 13: Threat & Detection Sources (Merged)

**Title:** "Threat & Detection Sources" (H3 - 48pt, navy)

**Elements:**
- Left half (55% width): MITRE ATT&CK stacked bar chart placeholder (enlarged to 3.0" height)
  - **Note:** Removed redundant legend and insight text boxes per Jan 2026 feedback
- Right half (45% width): 3 detection source cards (narrower) with FP rates
- Summary insight at bottom

### Slide 15: Prioritized Improvements

**Title:** "Prioritized Improvement Plan" (H3 - 48pt, navy)

**Elements:**
- 3 improvement item cards with priority badges (HIGH/MEDIUM/LOW)
- Each card: Priority badge, title, description, owner, target, expected impact
- **Insight bar:** "Focus Areas for Continuous Improvement"

### Slide 16: Operational Coverage (Merged)

**Title:** "Operational Coverage" (H3 - 48pt, navy)

**Elements:**
- Left panel: After-Hours hero (64pt) + weeknight/weekend breakdown + **business hours definition** (e.g., "9AM-5PM EST, Mon-Fri")
- Right panel: Collaboration metrics (3 rows with value + label)
  - Average Touches "per alert" (updated terminology)
- Insight box at bottom: **Rewritten** to explain significance ("Critical Start handled X alerts when your team was offline...")

### Slide 18: Key Takeaways

**Title:** "Key Takeaways: This Period" (H3 - 48pt, navy)

**Elements:**
- 4 takeaway cards with checkmark icons
- Each card: Checkmark + takeaway text (18pt)
- **Note:** Cost figure replaced with "Millions in modeled cost exposure avoided"
- **Insight bar:** "Bottom Line" summary

### Slide 19: Looking Ahead

**Title:** "Looking Ahead" (H3 - 48pt, navy)

**Elements:**
- Section 1: Next Period Targets (2 bullet items)
- Section 2: Strategic Focus (2 bullet items)
- Section 3: Your Partnership (contact CTA)
- **NO insight bar** - This slide IS the narrative content (per Jan 2026 feedback)

### Slide 20: Contact (Closing)

**Title:** "Thank You" (60pt, navy, centered)

**Elements:**
- Subtitle: "Questions? We're here to help." (24pt)
- Contact card with **dynamic CSM name and email** from ReportData
- Website: www.criticalstart.com
- Report date footer

---

## Helper Functions

### `setup_content_slide(prs, title_text, include_title=True)`

Creates a new content slide with standard branding:
- Adds transparent header (title, CRITICAL START, slide number)
- Adds transparent footer (date, copyright, intent)
- Adds logo at top right
- Adds slide title using H3 typography
- Returns `(slide, content_top)` tuple

### `create_section_header_layout(prs, title_text, narrative_text)`

Creates a gray section title card with:
- Gray background (#F5F5F5)
- Centered title (H2 - 72pt, Navy)
- Narrative subtitle (18pt, Slate)
- Footer only (no header)

### `add_insight_callout(slide, prs, title_text, body_text, top, height)`

Adds a standardized blue insight bar with:
- Light blue background with blue border
- Title text (bold, 14pt)
- Body text (12pt)

### `add_master_slide_elements(slide, prs, slide_number, include_header, include_footer)`

Adds header and footer elements to any slide.

### `add_logo(slide, position, prs)`

Places Critical Start logo at specified position.

---

## Revision History

| Date | Version | Author | Summary |
|------|---------|--------|---------|
| 2025-12-22 | 1.0 | AI (Claude) | Initial slide structure documentation |
| 2025-12-22 | 2.0 | AI (Claude) | Updated for CRITICALSTART branding guidelines |
| 2025-12-23 | 3.0 | AI (Claude) | Added gray section title cards with narrative subtitles |
| 2026-01-09 | 4.0 | AI (Claude) | Consolidated to 15-slide structure: Removed section cards, merged duplicate slides, added Contact slide |
| 2026-01-09 | 5.0 | AI (Claude) | **Revised 20-slide structure**: Restored 6 section cards, removed content redundancy (cost avoided), converted Security Outcomes to pipeline infographic, added insight bars to all content slides |
| 2026-01-09 | 6.0 | AI (Claude) | **21-slide structure**: Added CORR Platform funnel slide (AI Accelerated Security) after Executive Dashboard in Executive Summary section |
| 2026-01-09 | 7.0 | AI (Claude) | **20-slide structure**: Removed redundant Security Operations Pipeline slide (now covered by CORR Funnel) |
| 2026-01-09 | 8.0 | AI (Claude) | Redesigned slide 8: Renamed to "Critical Start's Performance", added split layout with MTTR by severity metrics panel |
| 2026-01-12 | 9.0 | AI (Claude) | **Jan 2026 Stakeholder Feedback**: Removed trend arrows from metric cards (rely on border color); standardized terminology (alerts, true positives, downgraded, end-to-end); fixed vertical overflow on slides 15, 18, 19; removed insight box from slide 19; updated slide 3 labels (MTTR, CLOSED END-TO-END); added cost methodology footnote to slide 6; updated slide 8 labels and enlarged chart; updated slide 12 terminology (Downgraded); cleaned up slide 13; added business hours definition to slide 16; added dynamic CSM name to slide 20; rewrote insight boxes to explain significance |
| 2026-01-12 | 10.0 | AI (Claude) | **Jan 2026 Cleanup**: Merged Industry Comparison (slide 9) into Performance slide as benchmark lines on trend chart; restructured Response & Detection to remove inverse metrics and add CORR context (Threat Reduction); fixed slide 18 footer overlap; added discrete counts to slide 12 narrative cards and Sankey legend. Presentation now 19 slides. |
| 2026-01-12 | 11.0 | AI (Claude) | **Slide 9 Metrics Cleanup**: Changed right panel to 2-over-1 layout (True Positive + False Positive on top, Alert Reduction spanning bottom); renamed metrics for clarity: Containment Rate → Remediation Rate, Analyst Escalation → Human Review Rate, Threat Reduction → Alert Reduction. |
