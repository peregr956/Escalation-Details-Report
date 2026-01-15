"""Slide builders package for Executive Business Review.

==============================================================================
REFACTORING IN PROGRESS - HANDOFF READY
==============================================================================

This package provides a modular structure for slide building. The foundation
has been established to enable gradual migration of code from the monolithic
`generate_presentation.py` (~3,000 lines).

CURRENT STATUS (January 2026):
- base.py: SlideBuilder and SectionBuilder base classes - IMPLEMENTED
- common.py: Shared layouts with re-exports from generate_presentation.py - PARTIAL
- Other modules: Stub files documenting target functions - STUBS ONLY

The main `generate_presentation.py` remains the source of truth for now.
Engineers can gradually migrate functions to this package.

==============================================================================
USAGE
==============================================================================

For now, import from generate_presentation.py:

    from generate_presentation import (
        build_executive_summary_slides,
        build_value_delivered_slides,
        build_protection_achieved_slides,
        build_threat_landscape_slides,
        build_insights_slides,
        build_forward_direction_slide,
        create_section_header_layout,
    )

As functions are migrated, they can be imported from this package:

    from slides.common import create_section_header_layout
    from slides.base import SlideBuilder

==============================================================================
PACKAGE STRUCTURE
==============================================================================

    slides/
    ├── __init__.py           # This file - package exports and docs
    ├── base.py               # SlideBuilder base class (IMPLEMENTED)
    ├── common.py             # Shared layouts (PARTIAL - with re-exports)
    ├── executive_summary.py  # Title slide and executive dashboard (STUB)
    ├── value_delivered.py    # Cost avoidance and value breakdown (STUB)
    ├── protection_achieved.py # Performance trends and metrics (STUB)
    ├── threat_landscape.py   # MITRE tactics and severity (STUB)
    ├── insights.py           # Improvement plans (STUB)
    └── forward_direction.py  # Key takeaways, contact (STUB)

==============================================================================
MIGRATION GUIDE FOR ENGINEERS
==============================================================================

To migrate a function (e.g., build_value_delivered_slides):

1. Create the implementation in the appropriate module (slides/value_delivered.py)
2. Import dependencies from helpers.py and constants.py
3. Add unit tests in tests/test_slides_value_delivered.py
4. Update generate_presentation.py to import from slides package
5. Remove the old function from generate_presentation.py
6. Run full test suite to verify

Key concerns:
- Avoid circular imports (helpers.py should not import from slides/)
- Maintain backward compatibility during transition
- Keep function signatures identical when moving
"""

__version__ = '1.0.0'

# Note: Imports commented out to avoid circular import issues during transition.
# Uncomment as modules are fully implemented.

# from slides.base import SlideBuilder, SectionBuilder
# from slides.common import (
#     create_content_slide_layout,
#     insert_chart_image,
# )

__all__ = [
    # 'SlideBuilder',
    # 'SectionBuilder',
    # 'create_content_slide_layout',
    # 'insert_chart_image',
]
