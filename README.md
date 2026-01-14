# Escalation to Client Details Report

Executive business review deck generator that quantifies the work Critical Start's Managed Detection & Response (MDR) team delivers for executive stakeholders. This repo stores the PowerPoint generator, anonymized sample data, and the Markdown-based brand system that replaced the legacy DOCX package.

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
| `clients/` | Client configuration files (use `sample.yaml` as template). |
| `data/` | Data files including `alert-details-template.xlsx` (headers only). |
| `tests/` | Test scripts for branding and slide generation. |
| `scripts/` | Utility scripts (e.g., `convert_logo_to_png.py`). |
| `brand/` | Brand guidelines (`brand-guidelines.md`, `visual-identity.md`). |
| `docs/` | Technical documentation (`CURRENT_STATE.md`, `SLIDE_STRUCTURE.md`, `DEVELOPER_GUIDE.md`). |
| `assets/` | Logo and image assets (`Critical-Start-Stacked-Logo_0-2.png`). |
| `output/` | Generated PowerPoint presentations (gitignored). |
| `ROADMAP.md` | Phased improvement plan for the presentation system. |
| `CHANGELOG.md` | Version history and release notes. |

## Tech Stack

### PowerPoint Generator
- Python 3.10+ with `python-pptx` for presentation generation.
- Playwright for headless browser chart rendering.
- PIL/Pillow for image processing.
- openpyxl for Excel data parsing.
- PyYAML for configuration file handling.
- Brand-compliant typography (Roboto/Arial fallback) and color palette.

## Quick Start

### Generate PowerPoint Presentation

1. Set up the virtual environment:
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   pip install -r requirements.txt
   playwright install chromium
   ```

2. Generate the presentation:
   ```bash
   # Using sample configuration (Acme Co.)
   python generate_presentation.py --config clients/sample.yaml

   # Using real client data (local only)
   python generate_presentation.py --data data/alert-details-export-*.xlsx --config clients/[client].yaml
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
- **PowerPoint:** Open the generated `.pptx` file in PowerPoint or Google Slides.
- For executive walkthroughs, keep the tab in full-screen (1920×1080) or export to PDF before sending.

## Data Sensitivity

- Client data files are gitignored (`data/*-export-*.xlsx`).
- Client-specific configs are gitignored (`clients/*.yaml`) except `clients/sample.yaml`.
- Use `clients/sample.yaml` as the template for new client configurations.

## Data-Driven Generation

The PowerPoint generator now supports fully dynamic data loading from Excel files:

### Input Files
- **Excel Data (1-3 files):** Standard 45-column incident export from CORR. Provide multiple files for trend comparison—the last file is treated as the current period.
- **Config File (YAML):** Client-specific settings including tier, industry benchmarks, and SLA targets. Copy `clients/sample.yaml` as a template or start from `client_config.yaml`.

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

## Sample configuration

`clients/sample.yaml` provides an anonymized baseline:
- Client: **Acme Co.**
- Tier: **Signature Tier MDR**
- Period: **Jan 1–Dec 31, 2025** (365 days / 8,760 hours of coverage)
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
