# Escalation to Client Details Report

Interactive, single-page business review experience that quantifies the work Critical Start's Managed Detection & Response (MDR) team delivers for executive stakeholders. This repo stores the production-ready HTML artifact, sample data, and the Markdown-based brand system that replaced the legacy DOCX package.

## Overview

The Escalation to Client Details Report complements the always-on ROI dashboard by providing a deeper, semi-annual/annual walkthrough of operational outcomes, risk posture, and measurable protection. It is usually narrated by the Customer Success Manager (CSM) with C-suite, CISO, and SOC leaders but can also be generated on demand for custom date ranges.

## Objectives
- **Drive renewals:** Quantify escalations handled, threats contained, efficiency gains, and compliance against contracted SLOs.
- **Improve operations:** Feed clients actionable insights on alert quality, automation gaps, and tuning opportunities—directly mapped to CORR-derived metrics.

## Audience
- **Primary:** Client executives, CISOs, and security leadership.
- **Secondary:** SOC managers/analysts who action the recommendations.

## Repository Map

| Path | Description |
| --- | --- |
| `generate_presentation.py` | PowerPoint presentation generator—creates branded slides from report data. |
| `report_data.py` | Data model (ReportData dataclass) and dynamic data loading from Excel. |
| `data_parser.py` | Excel file parser—loads 45-column incident data into typed records. |
| `metrics_calculator.py` | Metrics aggregation—computes all ReportData fields from incident data. |
| `insight_generator.py` | Rule-based insight engine—generates recommendations and narratives. |
| `config_loader.py` | Configuration loader—parses client YAML configuration files. |
| `client_config.yaml` | Template configuration with industry benchmarks and thresholds. |
| `chart_renderer.py` | Chart rendering module—uses Playwright to render Chart.js charts to PNG. |
| `constants.py` | Brand colors, fonts, typography scale, and layout constants. |
| `helpers.py` | Reusable layout helper functions for slide creation. |
| `chart_templates/` | HTML templates for chart rendering (pie, trend, stacked bar, sankey). |
| `examples/` | Sample HTML reports (`escalation_report_enhanced-3.html`, `your_team_performance.html`). |
| `data/` | Data files including `Report Extract.xlsx` for validation. |
| `tests/` | Test scripts for branding and slide generation. |
| `scripts/` | Utility scripts (e.g., `convert_logo_to_png.py`). |
| `brand/` | Brand guidelines (`brand-guidelines.md`, `visual-identity.md`). |
| `docs/` | Technical documentation (`CURRENT_STATE.md`, `SLIDE_STRUCTURE.md`, `DEVELOPER_GUIDE.md`). |
| `assets/` | Logo and image assets (`Critical-Start-Stacked-Logo_0-2.png`). |
| `output/` | Generated PowerPoint presentations (gitignored). |
| `ROADMAP.md` | Phased improvement plan for the presentation system. |
| `CHANGELOG.md` | Version history and release notes. |

## Tech Stack

### HTML Report
- Static HTML5 document with inline CSS for easier PDF exports.
- Vanilla JavaScript + Chart.js (line, pie, stacked bar, and Sankey via `chartjs-chart-sankey`).
- Google Fonts (Roboto) and Font Awesome icons.
- Fully responsive layout with print styles tuned for "Save as PDF".

### PowerPoint Generator
- Python 3.10+ with `python-pptx` for presentation generation.
- Playwright for headless browser chart rendering.
- PIL/Pillow for image processing.
- openpyxl for Excel data parsing.
- PyYAML for configuration file handling.
- Brand-compliant typography (Roboto/Arial fallback) and color palette.

## Quick Start

### Preview the HTML report locally
```bash
cd examples
python3 -m http.server 4173
# Navigate to http://localhost:4173/escalation_report_enhanced-3.html
```

### Generate PowerPoint Presentation

1. Set up the virtual environment:
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   pip install -r requirements.txt
   playwright install chromium
   ```

2. Generate the presentation from Excel data:
   ```bash
   # From a single period Excel file
   python generate_presentation.py --data report.xlsx --config client_config.yaml

   # From multiple periods (for trend charts)
   python generate_presentation.py --data aug.xlsx sep.xlsx oct.xlsx --config client_config.yaml

   # Using static sample data (for demos/testing)
   python generate_presentation.py
   ```

3. The presentation will be saved to `output/escalation_report_YYYY-MM-DD.pptx`

**Command-line options:**
```bash
python generate_presentation.py --help
python generate_presentation.py --data report.xlsx --config client.yaml  # Load from Excel
python generate_presentation.py --validate                                # Validate data only
python generate_presentation.py --output-dir ./custom_output
python generate_presentation.py --no-threat-landscape                     # Exclude threat slides
python generate_presentation.py --keep-charts                             # Keep temp chart images
```

### Export or share
- **HTML Report:** Use the browser's `Print → Save as PDF` workflow with "Background graphics" enabled.
- **PowerPoint:** Open the generated `.pptx` file in PowerPoint or Google Slides.
- For executive walkthroughs, keep the tab in full-screen (1920×1080) or export to PDF before sending.

## Data-Driven Generation

The PowerPoint generator now supports fully dynamic data loading from Excel files:

### Input Files
- **Excel Data (1-3 files):** Standard 45-column incident export from CORR. Provide multiple files for trend comparison—the last file is treated as the current period.
- **Config File (YAML):** Client-specific settings including tier, industry benchmarks, and SLA targets. Copy `client_config.yaml` as a template.

### What Gets Computed
All 120+ ReportData fields are automatically computed from the incident data:
- Volume metrics (alerts triaged, incidents escalated, true threats)
- Response times (MTTR, MTTD, P90) from TTR/TTD columns
- Detection sources grouped by Product column
- MITRE tactics data from tactic/priority columns
- After-hours analysis from timestamps
- Trend data across multiple period files

### Auto-Generated Insights
The insight engine generates recommendations based on threshold comparisons:
- Improvement items when metrics exceed thresholds
- Key achievements highlighting positive outcomes
- Executive summary narrative with dynamic values

### Legacy: HTML Report
For the HTML report, update the `.hero-section` blocks and Chart.js configs directly. Use `data/Report Extract.xlsx` as the source reference.

## Sample configuration

See `examples/escalation_report_enhanced-3.html` for a complete sample:
- Client: **Example Corporation**
- Tier: **Signature Tier MDR**
- Period: **Aug 1–31, 2025** (31 days / 744 hours of coverage)
- Incidents escalated: **267** (8.9 per day)
- True positives contained: **11**; zero breaches reported
- Response speed advantage: **34% faster** than peers

## Development

See `docs/DEVELOPER_GUIDE.md` for:
- Project structure and module overview
- Adding new slides
- Code style guidelines
- Testing instructions

Install development dependencies:
```bash
pip install -r requirements-dev.txt
```

## Brand Workspace

Markdown files under `brand/` are the authoritative branding system:
- Keep `brand-guidelines.md` and `visual-identity.md` up to date.
- Link any binary asset in `assets/` and document usage constraints.
- Follow `brand/README.md` for editing etiquette.

## Design, Automation & UX Guardrails

- Embrace Human-Centered Design: high contrast, responsive breakpoints, legible typography.
- Every number must trace back to CORR or the sanctioned Excel export.
- Recommendations must come from rule-driven thresholds so reports stay auto-generatable.
- Print/PDF parity is required—avoid browser-only interactions that break when exported.

## Working with AI Copilots

- Recognize this repo as part of Critical Start's client-facing analytics suite before suggesting changes.
- Preserve HCD alignment and Critical Start voice; never introduce off-brand colors, fonts, or assets.
- Maintain data lineage—if a figure cannot be derived from CORR or the approved export, don't surface it.
- Prioritize reusable snippets and automation-friendly approaches when contributing code or copy.

## License & Confidentiality

This project is proprietary to Critical Start. All datasets, code, and visualizations remain confidential and should only be used by authorized internal teams and approved clients. No public redistribution.
