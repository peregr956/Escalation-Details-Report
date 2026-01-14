"""Slide builders package for Escalation Details Report.

This package provides a modular structure for slide builder functions,
organized by section of the Executive Business Review presentation.

Package Structure:
    slides/
    ├── __init__.py           # This file - package documentation
    ├── common.py             # Shared slide layouts (title, section, content)
    ├── executive_summary.py  # Title slide and executive dashboard
    ├── value_delivered.py    # Cost avoidance and value breakdown
    ├── protection_achieved.py # Performance trends and response metrics
    ├── threat_landscape.py   # MITRE tactics and severity alignment
    ├── insights.py           # Improvement plans and operational coverage
    └── forward_direction.py  # Key takeaways, looking ahead, contact

Current Implementation:
    All slide builder functions are currently implemented in the main
    generate_presentation.py file. This package establishes the target
    structure for incremental refactoring.

Usage:
    The main generate_presentation.py uses these functions directly.
    Future refactoring can move implementations into this package,
    then update generate_presentation.py to import from here.

    # Future usage pattern:
    # from slides import build_executive_summary_slides
    # from slides.common import create_section_header_layout

For the current implementation, import directly from generate_presentation:
    from generate_presentation import (
        build_executive_summary_slides,
        create_section_header_layout,
        # etc.
    )
"""

# Package version
__version__ = '0.1.0'

# Note: Imports are intentionally not included here to avoid circular
# dependencies with generate_presentation.py. When implementations are
# moved to this package, imports can be added.

