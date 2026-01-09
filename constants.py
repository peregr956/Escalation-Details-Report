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

# Typography scale per Executive Business Review specifications
# Note: H1-H3 are for content titles, NOT the header/footer elements
H1_FONT_SIZE = Pt(48)        # Hero/title slides
H2_FONT_SIZE = Pt(28)        # Section headers, large metrics
H3_FONT_SIZE = Pt(22)        # Slide titles
H4_FONT_SIZE = Pt(18)        # Subheadings
H5_FONT_SIZE = Pt(16)        # Card titles, bullet text
H6_FONT_SIZE = Pt(11)        # Labels, small text
PARAGRAPH_FONT_SIZE = Pt(12) # Body text
FOOTER_FONT_SIZE = Pt(9)     # Footer AND header elements

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

PRESENTATION_TITLE = "EXECUTIVE BUSINESS REVIEW"
PRESENTATION_INTENT = "EBR"  # Executive Business Review
COPYRIGHT_TEXT = "©2025 CRITICAL START"


# =============================================================================
# METRIC THRESHOLDS (for status indicators)
# =============================================================================

# Each metric has:
#   - direction: "lower_is_better" or "higher_is_better"
#   - good_threshold: boundary for "good" status (green)
#   - warning_threshold: boundary for "warning" status (orange)
#   - Values beyond warning_threshold are considered "bad" (red)
#
# Arrow logic:
#   - higher_is_better + good = ↑ green, higher_is_better + bad = ↓ red
#   - lower_is_better + good = ↓ green, lower_is_better + bad = ↑ red
#   - warning = → orange (neutral direction)
# Uses outlined Unicode arrows per FA6 brand guidelines

METRIC_THRESHOLDS = {
    # Response Time Metrics (lower is better)
    "mttr": {
        "direction": "lower_is_better",
        "good_threshold": 150,      # <= 150 min is good
        "warning_threshold": 200,   # 150-200 min is warning, > 200 is bad
        "unit": "min",
        "label": "Mean Time to Respond"
    },
    "mttd": {
        "direction": "lower_is_better",
        "good_threshold": 45,       # <= 45 min is good
        "warning_threshold": 60,    # 45-60 min is warning, > 60 is bad
        "unit": "min",
        "label": "Mean Time to Detect"
    },
    "critical_high_mttr": {
        "direction": "lower_is_better",
        "good_threshold": 60,       # <= 60 min is good
        "warning_threshold": 90,    # 60-90 min is warning, > 90 is bad
        "unit": "min",
        "label": "Critical/High MTTR"
    },
    "medium_low_mttr": {
        "direction": "lower_is_better",
        "good_threshold": 60,       # <= 60 min is good
        "warning_threshold": 120,   # 60-120 min is warning, > 120 is bad
        "unit": "min",
        "label": "Medium/Low MTTR"
    },
    "p90_minutes": {
        "direction": "lower_is_better",
        "good_threshold": 90,       # <= 90 min is good
        "warning_threshold": 120,   # 90-120 min is warning, > 120 is bad
        "unit": "min",
        "label": "P90 Response Time"
    },
    
    # Quality Metrics (lower is better for FP, higher for others)
    "false_positive_rate": {
        "direction": "lower_is_better",
        "good_threshold": 10.0,     # <= 10% is good
        "warning_threshold": 15.0,  # 10-15% is warning, > 15% is bad
        "unit": "%",
        "label": "False Positive Rate"
    },
    
    # Performance Metrics (higher is better)
    "containment_rate": {
        "direction": "higher_is_better",
        "good_threshold": 95.0,     # >= 95% is good
        "warning_threshold": 90.0,  # 90-95% is warning, < 90% is bad
        "unit": "%",
        "label": "Containment Rate"
    },
    "response_advantage_percent": {
        "direction": "higher_is_better",
        "good_threshold": 20.0,     # >= 20% better is good
        "warning_threshold": 10.0,  # 10-20% is warning, < 10% is bad
        "unit": "%",
        "label": "Response Advantage"
    },
    "sla_compliance_rate": {
        "direction": "higher_is_better",
        "good_threshold": 95.0,     # >= 95% is good
        "warning_threshold": 90.0,  # 90-95% is warning, < 90% is bad
        "unit": "%",
        "label": "SLA Compliance"
    },
    "automation_percent": {
        "direction": "higher_is_better",
        "good_threshold": 80.0,     # >= 80% is good
        "warning_threshold": 70.0,  # 70-80% is warning, < 70% is bad
        "unit": "%",
        "label": "Automation Rate"
    },
    
    # Detection Quality Metrics (higher is better)
    "true_threat_precision": {
        "direction": "higher_is_better",
        "good_threshold": 30.0,     # >= 30% is good
        "warning_threshold": 20.0,  # 20-30% is warning, < 20% is bad
        "unit": "%",
        "label": "True Threat Precision"
    },
    "signal_fidelity": {
        "direction": "higher_is_better",
        "good_threshold": 90.0,     # >= 90% is good
        "warning_threshold": 80.0,  # 80-90% is warning, < 80% is bad
        "unit": "%",
        "label": "Signal Fidelity"
    },
    "client_validated": {
        "direction": "higher_is_better",
        "good_threshold": 85.0,     # >= 85% is good
        "warning_threshold": 75.0,  # 75-85% is warning, < 75% is bad
        "unit": "%",
        "label": "Client Validated"
    },
    
    # Volume Metrics (context-dependent, generally higher is informational)
    "threats_contained": {
        "direction": "higher_is_better",
        "good_threshold": 1,        # >= 1 threat contained is good (doing our job)
        "warning_threshold": 0,     # 0 could be warning (nothing to show?)
        "unit": "",
        "label": "Threats Contained"
    },
    "closed_end_to_end": {
        "direction": "higher_is_better",
        "good_threshold": 100,      # >= 100 is good (active engagement)
        "warning_threshold": 50,    # 50-100 is warning
        "unit": "",
        "label": "Closed End-to-End"
    }
}


# =============================================================================
# SECTION NARRATIVE CONSTANTS
# =============================================================================

# These narrative subtitles appear on gray section title slides to guide
# executives through a clear story arc:
# Current State → Value Delivered → Performance → Threats → Opportunities
#
# Each narrative now includes:
# - A connector phrase referencing prior sections (story thread)
# - Outcome-focused language that frames what the section demonstrates
SECTION_NARRATIVES = {
    "executive_summary": {
        "title": "Executive Summary",
        "narrative": "Your security at a glance—key metrics that show where we stand and where we're winning",
        "story_thread": None  # First section, no prior context
    },
    "value_delivered": {
        "title": "Value Delivered",
        "narrative": "Translating those outcomes into business impact: cost avoided, time saved, and breaches prevented",
        "story_thread": "Building on the metrics above"
    },
    "protection_achieved": {
        "title": "Protection Achieved",
        "narrative": "How we delivered that value—response speed and detection quality that outpace industry standards",
        "story_thread": "Behind the cost savings"
    },
    "threat_landscape": {
        "title": "Threat Landscape",
        "narrative": "The threats driving our performance metrics—and how we're calibrating our defenses to match",
        "story_thread": "What we protected against"
    },
    "insights": {
        "title": "Insights & Continuous Improvement",
        "narrative": "From threat patterns to action—the specific improvements that will elevate next period's outcomes",
        "story_thread": "Learning from this period"
    },
    "forward_direction": {
        "title": "Forward Direction",
        "narrative": "Strategic next steps to build on this period's wins and address areas needing attention",
        "story_thread": "Charting the path ahead"
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
