# Current State Documentation

This document provides a comprehensive overview of the PowerPoint generation system for the Escalation to Client Details Report.

## System Architecture

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                         PowerPoint Generation Pipeline                       │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│   ┌──────────────────┐    ┌──────────────────┐    ┌──────────────────┐     │
│   │  report_data.py  │───▶│generate_presen-  │───▶│   output/*.pptx  │     │
│   │   (ReportData)   │    │   tation.py      │    │                  │     │
│   └──────────────────┘    └────────┬─────────┘    └──────────────────┘     │
│                                    │                                        │
│                                    ▼                                        │
│                           ┌──────────────────┐                              │
│                           │chart_renderer.py │                              │
│                           │   (Playwright)   │                              │
│                           └────────┬─────────┘                              │
│                                    │                                        │
│                                    ▼                                        │
│                           ┌──────────────────┐                              │
│                           │ chart_templates/ │                              │
│                           │  - pie_chart.html│                              │
│                           │  - trend_chart   │                              │
│                           │  - stacked_bar   │                              │
│                           │  - sankey_chart  │                              │
│                           └──────────────────┘                              │
│                                                                              │
└─────────────────────────────────────────────────────────────────────────────┘
```

### Core Components

#### 1. `generate_presentation.py` (3,030 lines)

The main orchestration script that creates the PowerPoint presentation.

**Key Functions:**
| Function | Purpose |
|----------|---------|
| `create_presentation()` | Initialize 16:9 widescreen presentation |
| `apply_branding(prs)` | Set up slide masters with brand colors |
| `add_logo(slide, position, prs)` | Add Critical Start logo to slides |
| `create_title_slide_layout()` | Create title slide with gradient background |
| `create_section_header_layout()` | Create section divider slides |
| `create_content_slide_layout()` | Create standard content slides |
| `build_executive_summary_slides()` | Slides 1-3: Title and executive summary |
| `build_value_delivered_slides()` | Slides 4-5: Cost avoidance and value breakdown |
| `build_protection_achieved_slides()` | Slides 6-8: Performance trends and metrics |
| `build_threat_landscape_slides()` | Slides 9-12: MITRE tactics, severity flows |
| `build_insights_slides()` | Slides 13-15: Improvement plan and insights |
| `build_forward_direction_slide()` | Slide 16: Next steps and recommendations |
| `insert_chart_image()` | Insert rendered chart images into slides |

**Dependencies:**
- `python-pptx` - PowerPoint file generation
- `PIL (Pillow)` - Image processing
- `pathlib` - Path handling
- `argparse` - Command-line arguments
- `logging` - Logging and diagnostics

#### 2. `report_data.py`

Data model containing all report metrics.

**ReportData Dataclass Fields:**

```python
# Client Info
client_name: str          # "Lennar Corporation"
tier: str                 # "Signature Tier"
period_start: str         # "August 1, 2025"
period_end: str           # "August 31, 2025"
period_days: int          # 31
report_date: str          # "November 5, 2025"

# Hero Metrics
alerts_triaged: int       # 2110
client_touch_decisions: int  # 1690
closed_end_to_end: int    # 420
true_threats_contained: int  # 11
response_advantage_percent: float  # 34.0
mttr_minutes: int         # 126
p90_minutes: int          # 87
industry_median_minutes: int  # 192
after_hours_escalations: int  # 158
coverage_hours: int       # 744
automation_percent: float # 86.0

# Executive Summary
incidents_escalated: int  # 267
incidents_per_day: float  # 8.9
false_positive_rate: float  # 9.0

# Cost Avoidance
total_modeled: int        # 7550000
analyst_hours: int        # 452
analyst_cost_equivalent: int  # 38000
coverage_cost_equivalent: int  # 163000
breach_exposure_avoided: int  # 7340000

# Performance Metrics
critical_high_mttr: int   # 67
medium_low_mttr: int      # 52
mttd_minutes: int         # 4
containment_rate: float   # 98.0

# Detection Quality
true_threat_precision: float  # 31.4
signal_fidelity: float    # 91.0
client_validated: float   # 86.9

# Complex Data Structures
industry_comparison: List[Dict]   # MTTR, MTTD, Incidents/Day comparisons
detection_sources: List[Dict]     # Source name, incidents, percent, fp_rate
playbook_auto: Dict               # count, percent for automated escalations
analyst_escalation: Dict          # count, percent for manual escalations
mttr_trend: List[int]             # Historical MTTR values
mttd_trend: List[int]             # Historical MTTD values
fp_trend: List[float]             # Historical FP% values
period_labels: List[str]          # Period labels for trend charts
tactics: List[str]                # MITRE tactic names
high_severity: List[int]          # High severity counts by tactic
medium_severity: List[int]        # Medium severity counts by tactic
low_severity: List[int]           # Low severity counts by tactic
info_severity: List[int]          # Info severity counts by tactic
severity_flows: List[Dict]        # Sankey flow data (from, to, flow)
improvement_items: List[Dict]     # Priority, title, owner, target, description
avg_touches: float                # 2.3
client_participation: str         # "72%"
client_led_closures: str          # "21%"
```

#### 3. `chart_renderer.py`

Chart generation module using Playwright and Chart.js.

**Supported Chart Types:**
| Type | Template | Purpose |
|------|----------|---------|
| `pie` | `pie_chart.html` | Operational load distribution |
| `trend` | `trend_chart.html` | MTTR/MTTD/FP% trends over time |
| `stacked_bar` | `stacked_bar_chart.html` | MITRE tactics by severity |
| `sankey` | `sankey_chart.html` | Severity flow from vendor to CS |

**Key Functions:**
- `get_template_path(template_name)` - Resolve template file path
- `inject_chart_data(page, data)` - Inject data via JavaScript
- `wait_for_chart_ready(page, canvas_selector, timeout)` - Wait for rendering
- `capture_chart_image(page, output_path, width, height)` - Screenshot chart
- `render_chart(template_name, data, output_path, width, height)` - Full pipeline

---

## Current Slide Structure

### Slide Breakdown

| Slide # | Section | Title | Content |
|---------|---------|-------|---------|
| 1 | Executive Summary | Title Slide | Client name, tier, period, report date |
| 2 | Executive Summary | Key Metrics Overview | Alerts triaged (hero metric) |
| 3 | Executive Summary | Period Highlights | Key achievements summary |
| 4 | Value Delivered | Cost Avoidance Hero | $7.55M modeled cost avoided |
| 5 | Value Delivered | Value Breakdown | 3 cards: Security Ops, Coverage, Threat Prevention |
| 6 | Value Delivered | Detection Sources | Table with source breakdown |
| 7 | Protection Achieved | Performance Trends | Chart: MTTR, MTTD, FP% trends |
| 8 | Protection Achieved | Industry Comparison | Table comparing metrics to industry |
| 9 | Protection Achieved | Target Metrics | 4 cards: P90, Critical/High MTTR, Medium/Low MTTR, MTTD |
| 10 | Threat Landscape | MITRE Tactics | Stacked bar chart by severity |
| 11 | Threat Landscape | Severity Alignment | Sankey diagram: Vendor → CS severity |
| 12 | Threat Landscape | Escalation Methods | Playbook auto vs analyst escalation |
| 13 | Threat Landscape | Section placeholder | Additional threat context |
| 14 | Insights | Prioritized Improvement Plan | 3 improvement items with priority badges |
| 15 | Insights | Operational Insights | Detection quality metrics |
| 16 | Forward Direction | Looking Ahead | Next steps and recommendations |

### Visual Layout Standards (Current)

**Slide Dimensions:**
- Width: 10 inches (16:9 widescreen)
- Height: 5.625 inches

**Header Bar:**
- Height: 0.8 inches
- Background: CS_NAVY (#004C97)
- Title font: Calibri, 28pt, bold, white

**Content Area:**
- Top margin: 1.0 inches (below header)
- Side margins: 0.5 inches
- Bottom margin: 0.3 inches

**Logo Placement:**
- Position: Top right (content slides) or Top left (title slides)
- Size: 1.6" × 0.35"
- Margin from edge: 0.5" horizontal, 0.3" vertical

---

## Content Gaps Analysis

### Missing Content (Present in HTML, Absent from PowerPoint)

#### 1. After-Hours Customer Notifications
**HTML Location:** `escalation_report_enhanced-3.html` lines 1069-1114
**Content:**
- 158 total after-hours escalations
- Weeknight vs weekend breakdown
- Notification methods
- Response coverage details

#### 2. Response Efficiency
**HTML Location:** `escalation_report_enhanced-3.html` lines 1117-1144
**Content:**
- P90 response time details
- Priority-based response breakdown
- Efficiency metrics by severity

#### 3. Collaboration Quality
**HTML Location:** `escalation_report_enhanced-3.html` lines 1146-1174
**Content:**
- Average touches per incident (2.3)
- Client participation rate (72%)
- Client-led closures (21%)
- Collaboration quality indicators

#### 4. Detection Quality (Detailed)
**HTML Location:** `escalation_report_enhanced-3.html` lines 1175-1230
**Content:**
- True threat precision (31.4%)
- Signal fidelity (91%)
- Client validated outcomes (86.9%)
- Detailed breakdown with context

#### 5. Security Outcomes This Period
**HTML Location:** `escalation_report_enhanced-3.html` lines 1390-1470
**Content:**
- Comprehensive summary of all outcomes
- Visual cards for key achievements
- Consolidated view for executives

#### 6. Operational Load (Detailed)
**HTML Location:** `escalation_report_enhanced-3.html` lines 1032-1066
**Content:**
- Business hours: 51%
- After hours: 41%
- Weekends: 8%
- CS SOC alert resolution breakdown

### Partially Covered Content

#### Detection Sources
**Current:** Basic table with source, incidents, percent
**Missing:** False positive rate per source, tuning recommendations

#### Industry Comparison
**Current:** Simple table comparison
**Missing:** Context on benchmark sources, trend over time

---

## Current Aesthetic State

### Brand Colors Implementation

```python
# Current color constants in generate_presentation.py
CS_BLUE = RGBColor(0, 156, 222)    # #009CDE - Primary 1
CS_NAVY = RGBColor(0, 76, 151)    # #004C97 - Primary 2
CS_SLATE = RGBColor(52, 55, 65)   # #343741 - Primary 3
CS_RED = RGBColor(239, 51, 64)    # #EF3340 - Secondary 2
CS_ORANGE = RGBColor(255, 106, 20) # #FF6A14 - Secondary 3
```

**Missing Colors:**
- Violet (#702F8A) - Secondary 1, for highlights and graphics
- Gradient support (#009CDE → #004C97)

### Typography (Current vs Brand Standard)

| Element | Current | Brand Standard | Status |
|---------|---------|----------------|--------|
| Title Font | Calibri | Roboto Bold | ❌ Non-compliant |
| Body Font | Calibri | Roboto Regular | ❌ Non-compliant |
| Fallback Headline | N/A | Arial Black | ⚠️ Not implemented |
| Fallback Body | N/A | Arial Narrow | ⚠️ Not implemented |
| Title Size | 36pt | Variable by context | ✓ Acceptable |
| Body Size | 16pt | 16px minimum | ✓ Compliant |
| Metric Size | 60pt | Variable | ✓ Acceptable |

### Layout Issues

1. **Inconsistent Margins**
   - Some slides use 0.5" margins, others use 1.0"
   - No standardized grid system

2. **Spacing Variations**
   - Card spacing ranges from 0.15" to 0.4"
   - Inconsistent vertical rhythm

3. **White Space**
   - Hero metrics lack breathing room
   - Dense content areas need better separation

### Visual Hierarchy Issues

1. **Font Weight Usage**
   - Limited use of font weights for emphasis
   - Headers and body text sometimes similar weight

2. **Color for Emphasis**
   - Underutilized secondary colors (Violet, Orange)
   - Critical states not consistently highlighted

3. **Chart Integration**
   - Charts rendered as external images
   - Placeholder text visible when charts missing
   - Sizing could be optimized

---

## Technical Debt

### Code Organization Issues

1. **Large Functions**
   - `build_executive_summary_slides()` - 320+ lines
   - `build_protection_achieved_slides()` - 500+ lines
   - `build_threat_landscape_slides()` - 500+ lines
   - `build_insights_slides()` - 400+ lines

2. **Duplicated Patterns**
   - Header creation repeated in each section
   - Card layouts reimplemented multiple times
   - Text formatting code duplicated

3. **Missing Abstractions**
   - No reusable card component
   - No standardized metric display
   - No layout grid helpers

### Documentation Gaps

1. No inline comments for complex calculations
2. Limited docstrings on helper functions
3. No developer guide for adding new slides
4. No chart data format documentation

---

## File Inventory

### Source Files

| File | Lines | Purpose |
|------|-------|---------|
| `generate_presentation.py` | 3,030 | Main presentation generator |
| `report_data.py` | 262 | Data model and sample data |
| `chart_renderer.py` | 332 | Chart rendering with Playwright |

### Chart Templates

| File | Purpose |
|------|---------|
| `chart_templates/pie_chart.html` | Operational load pie chart |
| `chart_templates/trend_chart.html` | MTTR/MTTD/FP% trend lines |
| `chart_templates/stacked_bar_chart.html` | MITRE tactics by severity |
| `chart_templates/sankey_chart.html` | Severity flow visualization |

### Assets

| File | Purpose |
|------|---------|
| `assets/critical-start-logo.svg` | Primary logo (vector) |
| `assets/critical-start-logo.png` | Primary logo (raster, preferred for PPTX) |

### Output

| File | Purpose |
|------|---------|
| `output/escalation_report_*.pptx` | Generated presentations |

---

## Dependencies

### Python Packages

```
python-pptx==1.0.2     # PowerPoint generation
pillow==12.0.0         # Image processing
playwright==1.57.0     # Browser automation for charts
lxml==6.0.2            # XML processing (python-pptx dependency)
xlsxwriter==3.2.9      # Excel support (optional)
```

### External Resources

- Chart.js 4.4.0 (CDN in chart templates)
- chartjs-chart-sankey (CDN in sankey template)
- Google Fonts: Roboto (in chart templates)
- Font Awesome 6.5.1 (in HTML report)

---

## Known Issues

1. **Chart Rendering Failures**
   - Playwright may timeout on slow systems
   - Missing charts show placeholder text

2. **Logo File Handling**
   - Falls back to SVG if PNG not found
   - SVG may not render correctly in all PowerPoint versions

3. **Data Hardcoding**
   - Sample data hardcoded in `report_data.py`
   - No dynamic data loading mechanism

4. **Gradient Limitations**
   - python-pptx has limited gradient support
   - Using solid colors as workaround

---

## Revision History

| Date | Version | Author | Summary |
|------|---------|--------|---------|
| 2025-12-22 | 1.0 | AI (Claude) | Initial current state documentation |


