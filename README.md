# Escalation to Client Details Report

## Overview

The **Escalation to Client Details Report** is an interactive dashboard developed by Critical Start to deliver comprehensive visibility into Managed Detection and Response (MDR) performance for enterprise clients. Unlike the ROI Dashboard, which focuses on continuous value demonstration, this report provides a semi-annual or annual deep-dive into operational outcomes, risk posture, and measurable protection results.

The report is designed to be guided by the Customer Success Manager (CSM) during business reviews with client security leadership and executives. Clients may also self-generate reports on demand using a custom date range picker to review specific time periods.

## Purpose

The Escalation to Client Details Report serves two main objectives:

1. **Drive Renewals** – By quantifying the operational value and threat mitigation achieved through Critical Start's MDR service.
2. **Inform Operational Improvements** – By providing data-driven insight into alert handling, detection accuracy, and response efficiency.

## Audience

- **Primary**: Client C-suite executives, CISOs, and Security Leadership
- **Secondary**: Security Operations Center (SOC) Managers and Analysts

## Key Sections

### 1. Hero (Executive Summary Header)

**Title**: "What We Delivered This Period — Real Protection, Measurable Outcomes"

**Highlights:**
- Total hours of expert analysis provided
- Number of alerts resolved without escalation
- Coverage level (e.g., 24×7)
- Target compliance across incident severities

### 2. Executive Summary

Sample highlights include:
> "Your security posture remains strong. We maintained 88% target compliance across all incident severities, with a 90th percentile response time of 87 minutes."

- Number of incidents escalated, true positives identified, and false positive rate
- P90 response times (Critical/High vs. Medium/Low), Mean Time to Detect (MTTD)
- **Industry Comparison**: Benchmarked against external standards rather than internal data to ensure meaningful context

### 3. Value Delivered

- Labor hours saved and their equivalent dollar value (loaded/unloaded cost)
- Threat prevention modeled via conservative breach cost estimates
- Visualized cost avoidance and productivity gain metrics

### 4. Operational Metrics

- **After-Hours Notifications**: Volume of off-hours alerts, weekend calls, and customer notifications
- **Response Efficiency**: Average touches per incident, client participation rate, and collaboration indicators
- **Detection Accuracy**: False positive rate, severity distribution, and source quality

### 5. Threat Landscape

- Summary of incidents by MITRE ATT&CK tactic and severity
- Prioritized improvement plan auto-generated from threshold analysis

**Example:**
- **Detection Tuning**: False positive rate (9%) exceeds threshold; Palo Alto Cortex XDR identified as primary driver.
- **Automation Opportunity**: Manual escalations (14%) exceed target; recommend expanding playbook coverage.

### 6. Security Outcomes

- Clear linkage between MDR investment and quantifiable security outcomes
- Demonstrated reduction in potential business impact from detected threats

## Data Source

The report is built on data from **CORR** (Cyber Operations Risk Response Platform) and draws 95% of its metrics from the **Escalation to Client Details Excel Export**.

An extract of the Escalation to Client Details Excel Export is provided in the Github Repo as "Report Extract"

Each section references its corresponding Excel column to ensure transparency and traceability.

## Design and Automation Requirements

- Must comply with **Human-Centered Design (HCD)** principles for clarity, accessibility, and usability.
- Fully auto-generatable from CORR data with minimal human intervention.
- Recommendations and improvement plans are automated based on rule-driven threshold logic.
- Designed for seamless integration into quarterly or semi-annual business review workflows.

## Current Focus (Development Phase)

- **Refactoring & Documentation**: Streamlining structure and layout for readability and HCD alignment.
- **Automation Validation**: Ensuring all metrics map directly to CORR-based data sources.
- **UX Review**: Optimizing data hierarchy and executive storytelling to meet client expectations and renewal goals.

## For ChatGPT Assistance

When connected to this repository, ChatGPT should:

- Recognize this project as part of Critical Start's client-facing analytics suite.
- Preserve HCD alignment in all UI/UX or structural recommendations.
- Ensure data lineage integrity — every figure should trace back to the CORR export.
- Prioritize automation, scalability, and consistency across reporting periods.
- Provide concise, explainable, and reusable code for metrics generation, visualization, or layout refactoring.
- Avoid introducing any data source outside CORR or approved industry benchmark datasets.

## Example Usage

Users (CSMs or clients) can:

- Select a custom date range via the dashboard interface.
- Generate a complete, auto-formatted PDF or web-based report.
- Walk through results with the CSM for contextual interpretation.

## Technical Details

### Technologies Used
- Pure HTML5 with embedded CSS and JavaScript
- Chart.js (v4.4.0) for interactive visualizations
- Responsive design with print-friendly layouts
- Google Fonts (Roboto) for typography
- Critical Start brand colors (#009CDE, #004C97, #EF3340, #FF6A14)

### Report Structure
1. **Tier 1**: Hero impact metrics (most important client-facing value)
2. **Tier 2**: Performance & trust zone (trend analysis, benchmarks)
3. **Tier 3**: Coverage & capability zone (value delivered, detection sources)
4. **Tier 4**: Technical detail zone (response efficiency, collaboration, threat landscape)

### Visualizations
- **3-Month Trend Line Chart**: MTTR, MTTD, and false positive rate trends
- **Operational Load Pie Chart**: Business hours vs. after-hours distribution
- **MITRE ATT&CK Stacked Bar Chart**: Threat tactics by severity level

### Current Report Configuration

**File**: `escalation_report_enhanced-3.html`

- **Client**: Example Corporation
- **Tier**: Signature Tier MDR
- **Period**: August 1-31, 2025 (31 days)
- **Total Incidents Escalated**: 267
- **Alerts Resolved by CS SOC**: 1,843
- **Status**: Target Met (88% compliance)
- **Generated**: November 5, 2025

## Brand Guidelines Workspace

The bulky DOCX guidelines have been replaced with Markdown templates under `brand/`. Use them as the living source of truth so future edits stay version-controlled:

- `brand/brand-guidelines.md` – master narrative, messaging, and governance prompts
- `brand/visual-identity.md` – logo usage, colors, typography, accessibility, and asset inventory
- `brand/README.md` – instructions for updating the templates

Populate these files once the approved copy is ready and keep assets (like `assets/critical-start-logo.svg`) linked from this directory for easy reference.

## License and Confidentiality

This project is proprietary to Critical Start and not for public redistribution.

All datasets, code, and visualizations are confidential and intended solely for authorized internal and client use.
