"""Brand constants for Critical Start PowerPoint presentations.

This module contains all brand colors, typography settings, and layout constants
per the brand/visual-identity.md guidelines.
"""
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor


# =============================================================================
# BRAND COLORS (per brand/visual-identity.md)
# =============================================================================

# Primary Colors
CS_BLUE = RGBColor(0, 156, 222)      # Primary 1 - Critical Start Blue #009CDE
CS_NAVY = RGBColor(0, 76, 151)       # Primary 2 - Deep Navy #004C97
CS_SLATE = RGBColor(52, 55, 65)      # Primary 3 - Charcoal #343741

# Secondary Colors
CS_VIOLET = RGBColor(112, 47, 138)   # Secondary 1 - Violet #702F8A
CS_RED = RGBColor(239, 51, 64)       # Secondary 2 - Red #EF3340
CS_ORANGE = RGBColor(255, 106, 20)   # Secondary 3 - Orange #FF6A14

# Status Colors
CS_GREEN = RGBColor(40, 167, 69)     # Success Green - for positive outcomes


# =============================================================================
# FONT CONSTANTS (per brand/visual-identity.md)
# =============================================================================

# Primary: Roboto (Regular, Medium, Bold)
# Fallback Headline: Arial Black
# Fallback Body: Arial Narrow
TITLE_FONT_NAME = "Roboto"
BODY_FONT_NAME = "Roboto"
FALLBACK_TITLE_FONT = "Arial Black"
FALLBACK_BODY_FONT = "Arial Narrow"


# =============================================================================
# TYPOGRAPHY SCALE (adjusted for slide readability)
# =============================================================================

# Reduced from brand guidelines to fit presentation layouts appropriately
# Note: H1-H3 are for content titles, NOT the header/footer elements
H1_FONT_SIZE = Pt(72)        # Hero/title slides - reduced from 114pt
H2_FONT_SIZE = Pt(48)        # Section headers, large metrics - reduced from 72pt
H3_FONT_SIZE = Pt(36)        # Slide titles - reduced from 48pt
H4_FONT_SIZE = Pt(20)        # Subheadings - reduced from 27pt
H5_FONT_SIZE = Pt(14)        # Card titles, bullet text - reduced from 18pt
H6_FONT_SIZE = Pt(10)        # Labels, small text - reduced from 12pt
PARAGRAPH_FONT_SIZE = Pt(10) # Body text - reduced from 12pt
FOOTER_FONT_SIZE = Pt(9)     # Footer AND header elements - unchanged

# Legacy aliases for backward compatibility
TITLE_FONT_SIZE = H3_FONT_SIZE       # Slide titles use H3
SUBTITLE_FONT_SIZE = H4_FONT_SIZE    # Subtitles use H4
HEADING_FONT_SIZE = H5_FONT_SIZE     # Headings use H5
BODY_FONT_SIZE = PARAGRAPH_FONT_SIZE
CAPTION_FONT_SIZE = FOOTER_FONT_SIZE
METRIC_FONT_SIZE = H2_FONT_SIZE      # Large metrics use H2
HERO_METRIC_FONT_SIZE = H1_FONT_SIZE # Hero metrics use H1


# =============================================================================
# LAYOUT CONSTANTS (standardized spacing)
# =============================================================================

MARGIN_STANDARD = Inches(0.5)
MARGIN_CONTENT = Inches(1.0)
CARD_SPACING = Inches(0.2)
HEADER_HEIGHT = Inches(0.4)   # Reduced for transparent header
FOOTER_HEIGHT = Inches(0.35)  # Footer area height


# =============================================================================
# MASTER SLIDE CONSTANTS
# =============================================================================

PRESENTATION_TITLE = "ESCALATION REPORT"
PRESENTATION_INTENT = "EBR"  # Executive Business Review
COPYRIGHT_TEXT = "©2025 CRITICAL START"


# =============================================================================
# SECTION NARRATIVE CONSTANTS
# =============================================================================

# These narrative subtitles appear on gray section title slides to guide
# executives through a clear story arc:
# Current State → Value Delivered → Performance → Threats → Opportunities
SECTION_NARRATIVES = {
    "executive_summary": {
        "title": "Executive Summary",
        "narrative": "A comprehensive view of your security posture and the value delivered this period"
    },
    "value_delivered": {
        "title": "Value Delivered",
        "narrative": "Quantifying the business impact: cost avoidance, operational efficiency, and breach prevention"
    },
    "protection_achieved": {
        "title": "Protection Achieved",
        "narrative": "Measurable outcomes that exceed industry standards and reduce your breach risk"
    },
    "threat_landscape": {
        "title": "Threat Landscape",
        "narrative": "Understanding the threats we're seeing and how we're adapting to protect your organization"
    },
    "insights": {
        "title": "Insights & Continuous Improvement",
        "narrative": "How we're evolving our partnership to advance your security maturity"
    },
    "forward_direction": {
        "title": "Forward Direction",
        "narrative": "Strategic recommendations to strengthen your security posture and maximize partnership value"
    }
}


def get_brand_colors():
    """Return dict of RGBColor objects for brand colors.
    
    Returns:
        dict: Dictionary with color names as keys and RGBColor objects as values.
    """
    return {
        'blue': CS_BLUE,
        'navy': CS_NAVY,
        'slate': CS_SLATE,
        'violet': CS_VIOLET,
        'red': CS_RED,
        'orange': CS_ORANGE,
        'green': CS_GREEN
    }
