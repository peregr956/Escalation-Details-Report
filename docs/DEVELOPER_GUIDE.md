# Developer Guide

This guide explains how to contribute to the Escalation to Client Details Report project.

## Project Structure

```
Escalation-Details-Report/
├── assets/                  # Logo and image assets
├── brand/                   # Brand guidelines documentation
├── chart_templates/         # Chart.js HTML templates
├── data/                    # Sample data files
├── docs/                    # Technical documentation
├── examples/                # Sample HTML reports
├── output/                  # Generated presentations (gitignored)
├── scripts/                 # Utility scripts
├── tests/                   # Test files
├── constants.py             # Brand colors, fonts, layout constants
├── helpers.py               # Reusable layout helper functions
├── generate_presentation.py # Main presentation generator
├── report_data.py           # Data model and sample data
├── chart_renderer.py        # Playwright chart rendering
├── requirements.txt         # Production dependencies
└── requirements-dev.txt     # Development dependencies
```

## Setup Development Environment

```bash
# Clone the repository
git clone https://github.com/peregr956/Escalation-Details-Report.git
cd Escalation-Details-Report

# Create virtual environment
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements-dev.txt

# Install Playwright browsers
playwright install chromium
```

## Module Overview

### `constants.py`

Contains all brand-related constants:

- **Colors**: `CS_BLUE`, `CS_NAVY`, `CS_SLATE`, `CS_VIOLET`, `CS_RED`, `CS_ORANGE`, `CS_GREEN`
- **Typography**: `H1_FONT_SIZE` through `H6_FONT_SIZE`, `TITLE_FONT_NAME`, `BODY_FONT_NAME`
- **Layout**: `MARGIN_STANDARD`, `CARD_SPACING`, `HEADER_HEIGHT`, `FOOTER_HEIGHT`
- **Section Narratives**: `SECTION_NARRATIVES` dict for section title cards

### `helpers.py`

Reusable functions for slide creation:

| Function | Purpose |
|----------|---------|
| `create_presentation()` | Initialize 16:9 widescreen presentation |
| `apply_branding(prs)` | Set up slide master with brand colors |
| `setup_content_slide(prs, title)` | Create standard content slide with header/footer |
| `add_logo(slide, position, prs)` | Add Critical Start logo |
| `add_master_slide_elements(...)` | Add header and footer elements |
| `add_insight_callout(...)` | Add blue insight callout box |
| `create_metric_card(...)` | Create metric display card |
| `create_data_table(...)` | Create standardized data table |

### `generate_presentation.py`

Main slide builder functions:

| Function | Slides |
|----------|--------|
| `build_executive_summary_slides()` | Title, Executive Dashboard |
| `build_value_delivered_slides()` | Cost Avoidance, Value Breakdown |
| `build_protection_achieved_slides()` | Trends, Industry Comparison |
| `build_threat_landscape_slides()` | MITRE Tactics, Severity Alignment |
| `build_insights_slides()` | Improvement Plan, Operational Coverage |
| `build_forward_direction_slide()` | Looking Ahead, Key Takeaways |
| `build_contact_slide()` | Contact information |

### `report_data.py`

Data model for report metrics:

```python
from report_data import ReportData, get_report_data

# Get sample data
data = get_report_data()

# Access metrics
print(data.client_name)           # "Lennar Corporation"
print(data.incidents_escalated)   # 267
print(data.mttr_minutes)          # 126
```

## Adding a New Slide

1. **Add data fields** (if needed) to `ReportData` in `report_data.py`:

```python
@dataclass
class ReportData:
    # ... existing fields ...
    new_metric: int = 0
```

2. **Create the slide builder** in `generate_presentation.py`:

```python
def build_my_new_slide(prs, data):
    """Build the new slide."""
    slide, content_top = setup_content_slide(prs, "My New Slide Title")
    
    # Add content using helpers
    create_metric_card(
        slide,
        left=Inches(0.5),
        top=content_top,
        width=Inches(2.5),
        height=Inches(1.5),
        value=str(data.new_metric),
        label="New Metric"
    )
    
    return slide
```

3. **Call the builder** from `main()`:

```python
def main():
    # ... existing code ...
    build_my_new_slide(prs, data)
```

## Chart Templates

Chart templates are HTML files in `chart_templates/` that use Chart.js:

| Template | Chart Type |
|----------|------------|
| `pie_chart.html` | Pie/Doughnut |
| `trend_chart.html` | Line chart |
| `stacked_bar_chart.html` | Stacked bar |
| `sankey_chart.html` | Sankey diagram |

Charts are rendered to PNG using Playwright and inserted into slides.

## Running Tests

```bash
# From project root
cd tests
python test_branding.py
python test_executive_summary.py

# With pytest (when available)
pytest tests/
```

## Code Style

- Follow PEP 8 guidelines
- Use type hints for function arguments
- Add docstrings to all public functions
- Keep functions under 100 lines when possible

```bash
# Format code
black *.py

# Check linting
flake8 *.py

# Sort imports
isort *.py
```

## Brand Guidelines

All visual elements must comply with Critical Start brand guidelines:

- **Colors**: Use only approved palette from `constants.py`
- **Typography**: Use Roboto (primary) or Arial (fallback)
- **Logo**: Use `add_logo()` helper for consistent placement
- **Spacing**: Use `MARGIN_STANDARD` and `CARD_SPACING` constants

See `brand/visual-identity.md` for complete guidelines.

## Commit Messages

Use descriptive commit messages:

```
Add new slide for collaboration metrics

- Create build_collaboration_slide() function
- Add collaboration_score field to ReportData
- Update SLIDE_STRUCTURE.md documentation
```

## Pull Request Process

1. Create a feature branch: `git checkout -b feature/my-feature`
2. Make changes and test locally
3. Update documentation if needed
4. Submit PR with clear description
5. Wait for review and approval
