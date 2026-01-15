"""Slide builders package for Escalation Details Report.

==============================================================================
IMPORTANT: PLANNED REFACTORING - NOT YET IMPLEMENTED
==============================================================================

This package contains STUB FILES ONLY. The actual slide builder implementations
are currently in the main `generate_presentation.py` file (~3,000 lines).

This package exists to establish the target modular structure for future
refactoring. See ROADMAP.md Phase 4 for the planned refactoring timeline.

STATUS: Placeholder stubs - do NOT import from this package directly.

To use slide builders, import from generate_presentation.py:

    from generate_presentation import (
        build_executive_summary_slides,
        build_value_delivered_slides,
        build_protection_achieved_slides,
        build_threat_landscape_slides,
        build_insights_slides,
        build_forward_direction_slide,
        create_section_header_layout,
    )

==============================================================================
PLANNED PACKAGE STRUCTURE (Target State)
==============================================================================

    slides/
    ├── __init__.py           # This file - package exports
    ├── common.py             # Shared slide layouts (title, section, content)
    ├── executive_summary.py  # Title slide and executive dashboard
    ├── value_delivered.py    # Cost avoidance and value breakdown
    ├── protection_achieved.py # Performance trends and response metrics
    ├── threat_landscape.py   # MITRE tactics and severity alignment
    ├── insights.py           # Improvement plans and operational coverage
    └── forward_direction.py  # Key takeaways, looking ahead, contact

Each module documents the functions that will be moved from generate_presentation.py.

REFACTORING GUIDANCE:
1. Move one section at a time (e.g., start with executive_summary.py)
2. Update generate_presentation.py imports after each move
3. Add unit tests for the moved functions
4. Ensure no circular imports (use helpers.py for shared utilities)
"""

# Package version
__version__ = '0.1.0'

# Note: Imports are intentionally not included here.
# This package contains stubs only - actual implementations are in
# generate_presentation.py. When refactoring is complete, add exports here.

# DO NOT ADD IMPORTS until implementations are moved to this package.
# Doing so will cause ImportErrors.

