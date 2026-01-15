# Changelog

All notable changes to the Executive Business Review project are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [Unreleased]

### Added
- `constants.py` - Extracted brand colors, fonts, typography, and layout constants
- `helpers.py` - Extracted reusable layout helper functions
- `examples/` directory for sample HTML reports
- `tests/` directory for test scripts
- `scripts/` directory for utility scripts
- `data/` directory for data files
- `CHANGELOG.md` for tracking version history
- `requirements-dev.txt` for development dependencies
- `docs/DEVELOPER_GUIDE.md` for contributor documentation

### Changed
- Restructured repository for better organization
- Updated `.gitignore` to exclude generated output files
- Moved test files to `tests/` directory
- Moved HTML example files to `examples/` directory
- Moved utility scripts to `scripts/` directory
- Moved data files to `data/` directory
- Refactored `generate_presentation.py` to use modular imports

### Removed
- Tracked `.pptx` output files from git (now gitignored)

---

## [1.4.0] - 2026-01-14

### Changed
- Sample data updated with anonymized placeholder client ("Acme Co.")
- Client configurations moved to `clients/` directory

### Removed
- Deprecated HTML report examples (`examples/` folder)
- HTML report documentation and Quick Start instructions
- Legacy sample data file

### Added
- Sample client configuration (`clients/sample.yaml`)
- Headers-only data template for engineering reference
- Data sensitivity handling (client data and configs gitignored)

---

## [1.3.0] - 2026-01-09

### Changed
- Revised slide structure to 20 slides: 14 content + 6 section cards
- Improved executive dashboard layout

---

## [1.2.0] - 2025-12-23

### Added
- Gray section title cards with narrative subtitles
- Transparent header/footer elements per branding guidelines
- CRITICALSTART Slide Branding compliance

### Changed
- Typography scale adjusted for presentation readability
- Native PowerPoint gradient fills for title slides

---

## [1.1.0] - 2025-12-22

### Added
- `ROADMAP.md` for phased improvement planning
- `docs/CURRENT_STATE.md` for system architecture documentation
- `docs/SLIDE_STRUCTURE.md` for slide-by-slide specifications
- Detailed breakdown slides for after-hours, response efficiency, collaboration, detection quality
- Extended `ReportData` class with additional metrics

### Changed
- Improved image insertion with aspect ratio preservation
- Enhanced logging for chart rendering

---

## [1.0.0] - 2025-12-19

### Added
- Initial PowerPoint generation system
- Chart rendering with Playwright and Chart.js
- Pie chart, trend chart, stacked bar, and Sankey templates
- Brand color palette and typography constants
- `report_data.py` data model
- `chart_renderer.py` for headless chart rendering
- Interactive HTML report (`escalation_report_enhanced-3.html`)
- Brand guidelines documentation (`brand/`)

---

## Version History from Git

| Date | Commit | Summary |
|------|--------|---------|
| 2026-01-09 | 3d29252 | Add .gitignore and remove cached/old files |
| 2026-01-09 | e6db29a | Revised slide structure: 14 content + 6 section cards = 20 slides |
| 2025-12-23 | 4d4dbc3 | Refactor presentation typography and enhance header/footer |
| 2025-12-23 | b10e0ae | Add gray section title cards with narrative subtitles |
| 2025-12-22 | c819c0a | Add detailed breakdown slides and documentation |
| 2025-12-22 | 320e298 | Enhance image insertion with aspect ratio preservation |
| 2025-12-22 | d1a5366 | Add PowerPoint generation guide |
| 2025-12-22 | 10742bb | Add chart data preparation and image insertion |
| 2025-12-22 | bfa3f1d | Refactor chart rendering with multiple chart types |
| 2025-12-22 | 5e8f218 | Add chart templates (Pie, Sankey, Stacked Bar, Trend) |
