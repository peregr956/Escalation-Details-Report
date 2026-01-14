"""Executive Summary slides builder.

This module will build the opening slides of the Executive Business Review:
- Title slide with client info and period
- Executive Dashboard with key metrics (6 cards in 2x3 grid)
- CORR Platform funnel visualization

Slides in this section:
1. Title Slide - Navy gradient background, client name, tier, period
2. Executive Summary section card - Gray with narrative
3. Executive Dashboard - 6 metric cards (threats, MTTR, closed E2E, alerts, after-hours, FP rate)
4. AI Accelerated Security (CORR Funnel) - Security events -> alerts pipeline

Current Implementation:
    These functions are currently in generate_presentation.py.
    
        from generate_presentation import (
            build_executive_summary_slides,
            build_corr_funnel_slide,
        )

Target Functions:
    - build_executive_summary_slides(prs, data) -> builds slides 1-4
    - build_corr_funnel_slide(prs, funnel_image_path=None) -> CORR funnel
"""

# Placeholder - implementations to be moved from generate_presentation.py
__all__ = [
    'build_executive_summary_slides',
    'build_corr_funnel_slide',
]
