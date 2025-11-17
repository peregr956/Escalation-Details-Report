# Escalation to Client Details Report

Interactive, single-page business review experience that quantifies the work Critical Start’s Managed Detection & Response (MDR) team delivers for executive stakeholders. This repo stores the production-ready HTML artifact, sample data, and the Markdown-based brand system that replaced the legacy DOCX package.

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
| `escalation_report_enhanced-3.html` | Signature Tier sample of the interactive report. Pure HTML+CSS+JS with Chart.js 4.4.0 and Font Awesome 6.5.1. |
| `Report Extract.xlsx` | Sanitized slice of the Escalation to Client Details export used to validate numbers/columns. Keep aligned with CORR. |
| `brand/brand-guidelines.md` | Master narrative, voice, governance prompts, and revision log. |
| `brand/visual-identity.md` | Logo usage, palette (#009CDE, #004C97, #EF3340, #FF6A14), typography, accessibility, and asset inventory. |
| `brand/README.md` | How-to for maintaining the Markdown guidelines. |
| `assets/critical-start-logo.svg` | Approved vector lockup embedded directly in the HTML header. |

## Tech Stack
- Static HTML5 document with inline CSS for easier PDF exports.
- Vanilla JavaScript + Chart.js (line, pie, stacked bar, and Sankey via `chartjs-chart-sankey`).
- Google Fonts (Roboto) and Font Awesome icons.
- Fully responsive layout with print styles tuned for “Save as PDF”.

## Quick Start

### Preview the dashboard locally
1. `cd /workspace`
2. `python3 -m http.server 4173`
3. Navigate to `http://localhost:4173/escalation_report_enhanced-3.html`
4. Stop the server with `Ctrl+C` when finished. (You can also open the file directly in a browser, but a local server keeps font/CDN calls consistent.)

### Export or share
- Use the browser’s `Print → Save as PDF` workflow with “Background graphics” enabled.
- For executive walkthroughs, keep the tab in full-screen (1920×1080) or export to PDF before sending.

## Updating Data & Narrative
- **Source data & QA:** 95% of the metrics come from the Escalation to Client Details Excel Export in CORR. Use `Report Extract.xlsx` as the working copy and keep column references intact (e.g., Incident Id col 1, CS SOC Verdict col 12, Escalated Datetime col 20, CS SOC Time to Respond col 37, MITRE Tactic col 42, Vendor Severity col 45, Current Priority col 13).
- **Hero + Executive Summary:** Update the `.hero-section` blocks and the first card (`Executive Summary`) together. Values (hours, alerts resolved, coverage, compliance, incidents escalated, true positives) should match the same date range and appear again in “Security Outcomes This Period”.
- **Charts & automation:** Near the bottom of the HTML you’ll find the Chart.js configs. Update the numerical arrays only—styling is shared.
  - `operationalLoadChart`: After-hours vs business-hours distribution (pie).
  - `trendChart`: MTTR, MTTD, FP% line graph (labels `['Period -2', 'Period -1', 'Current']`).
  - `severityFlows`: Array of `{ from, to, flow }` objects powering the Sankey chart and all severity insight copy.
  - `mitreChart`: Provides the stacked bar by tactic & severity.
- **Severity alignment flow:** When regenerating `severityFlows`, keep `Vendor`/`CS` prefixes and ensure totals reconcile to the escalated incident count. The script auto-computes upgraded/downgraded copies.
- **Prioritized improvement plan:** Thresholds appear in the “Recommendation Engine” list (false positive 10%, manual escalation 12%, containment 95%). Adjust copy, owners, and targets as the automation rules or client remediation plans evolve.
- **Value Delivered & Security Outcomes:** Refresh hours saved, CS SOC touches, target compliance, and forward-looking commitments so the hero, outcomes grid, and improvement plan tell the same story.

## Sample configuration (`escalation_report_enhanced-3.html`)
- Client: **Example Corporation**
- Tier: **Signature Tier MDR**
- Period: **Aug 1–31, 2025** (31 days / 744 hours of coverage)
- Incidents escalated: **267** (8.9 per day); alerts resolved without escalation: **1,843** (59.5 per day)
- True positives contained: **11**; zero breaches reported
- Target compliance achieved: **88%** with **87‑minute** P90 response; MTTR trending down to 126 minutes
- False positive rate: **9.0%** (Palo Alto Cortex XDR at 11.2% noted in tuning plan)
- Manual escalations: **14%** vs 12% target—driving the automation recommendation
- Report generated: **5 Nov 2025**; data sources: CORR + Escalation export (45 columns)

## Brand Workspace
Markdown files under `brand/` are now the authoritative branding system:
- Keep `brand-guidelines.md` and `visual-identity.md` up to date, including revision histories and asset tables.
- Link any binary asset in `assets/` (e.g., `assets/critical-start-logo.svg`) or an approved CDN, and document usage constraints.
- Follow `brand/README.md` for editing etiquette (tables aligned, assets referenced, change log updated).

## Design, Automation & UX Guardrails
- Embrace Human-Centered Design: high contrast, responsive breakpoints, legible typography.
- Every number must trace back to CORR or the sanctioned Excel export; note the column reference where space allows.
- Recommendations must come from rule-driven thresholds so reports stay auto-generatable with minimal manual touch.
- Print/PDF parity is required—avoid browser-only interactions that break when exported.

## Working with ChatGPT (or other copilots)
- Recognize this repo as part of Critical Start’s client-facing analytics suite before suggesting changes.
- Preserve HCD alignment and Critical Start voice; never introduce off-brand colors, fonts, or assets.
- Maintain data lineage—if a figure cannot be derived from CORR or the approved export, don’t surface it.
- Prioritize reusable snippets and automation-friendly approaches when contributing code or copy.
- Avoid new data sources unless explicitly cleared as industry benchmarks.

## License & Confidentiality
This project is proprietary to Critical Start. All datasets, code, and visualizations remain confidential and should only be used by authorized internal teams and approved clients. No public redistribution.
