"""Common slide layouts and utilities.

This module will contain shared slide creation functions used across sections:
- Title slide layout (gradient background, centered text)
- Section header layout (gray background, narrative subtitle)
- Content slide layout (standard header/footer with content area)
- Chart data preparation (format data for Chart.js rendering)
- Chart image insertion (place rendered chart PNGs into slides)

Current Implementation:
    These functions are currently in generate_presentation.py.
    To use them, import from the main module:
    
        from generate_presentation import (
            create_title_slide_layout,
            create_section_header_layout,
            create_content_slide_layout,
            prepare_chart_data,
            insert_chart_image,
        )

Target Functions (to be moved here):
    - create_title_slide_layout(prs, title_text, subtitle_text=None)
    - create_section_header_layout(prs, section_title, narrative_subtitle=None)
    - create_content_slide_layout(prs, slide_title, content_items=None)
    - prepare_chart_data(data: ReportData) -> dict
    - insert_chart_image(slide, placeholder_id, image_path, ...) -> bool
"""

# Placeholder - implementations to be moved from generate_presentation.py
__all__ = [
    'create_title_slide_layout',
    'create_section_header_layout',
    'create_content_slide_layout',
    'prepare_chart_data',
    'insert_chart_image',
]
