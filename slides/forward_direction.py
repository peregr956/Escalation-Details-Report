"""Forward Direction slides builder.

This module will build the closing slides of the Executive Business Review:
- Key takeaways summary (4 bullet cards with checkmarks)
- Looking ahead (next period targets, strategic focus)
- Contact information (dynamic CSM name/email)

Slides in this section:
16. Forward Direction section card - Gray with narrative
17. Key Takeaways - 4 achievement cards + bottom line insight
18. Looking Ahead - Next steps and strategic focus
19. Contact - Thank you + CSM contact info

Current Implementation:
    These functions are currently in generate_presentation.py.
    
        from generate_presentation import (
            build_key_takeaways_slide,
            build_forward_direction_slide,
            build_contact_slide,
            build_additional_content_slides,
        )

Target Functions:
    - build_key_takeaways_slide(prs, section_title, takeaways, data=None)
    - build_forward_direction_slide(prs, data) -> next steps slide
    - build_contact_slide(prs, data) -> contact info slide
    - build_additional_content_slides(prs, data) -> any extra content
"""

# Placeholder - implementations to be moved from generate_presentation.py
__all__ = [
    'build_key_takeaways_slide',
    'build_forward_direction_slide',
    'build_contact_slide',
    'build_additional_content_slides',
]
