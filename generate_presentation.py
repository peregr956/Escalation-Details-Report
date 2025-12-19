from pptx import Presentation
from pptx.util import Inches, Pt, RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pathlib import Path

# Brand Color Constants
CS_BLUE = RGBColor(0, 156, 222)
CS_NAVY = RGBColor(0, 76, 151)
CS_SLATE = RGBColor(52, 55, 65)
CS_RED = RGBColor(239, 51, 64)
CS_ORANGE = RGBColor(255, 106, 20)

# Font Constants
TITLE_FONT_NAME = "Calibri"
BODY_FONT_NAME = "Calibri"
TITLE_FONT_SIZE = Pt(36)
BODY_FONT_SIZE = Pt(16)
METRIC_FONT_SIZE = Pt(60)


def get_brand_colors():
    """Return dict of RGBColor objects for brand colors.
    
    Returns:
        dict: Dictionary with color names as keys and RGBColor objects as values.
    """
    return {
        'blue': CS_BLUE,
        'navy': CS_NAVY,
        'slate': CS_SLATE,
        'red': CS_RED,
        'orange': CS_ORANGE
    }


def create_presentation():
    """Initialize 16:9 widescreen presentation.
    
    Returns:
        Presentation: A new Presentation object with 16:9 aspect ratio.
    """
    prs = Presentation()
    # Set slide size to 16:9 widescreen (10 x 5.625 inches)
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)
    return prs


def apply_branding(prs):
    """Set up slide masters with brand colors.
    
    Args:
        prs (Presentation): The presentation to apply branding to.
    """
    # Get the slide master
    slide_master = prs.slide_master
    
    # Set background color to white (default)
    background = slide_master.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    # Note: python-pptx has limited support for modifying slide masters directly.
    # Most branding will be applied per-slide using layout helpers.


def add_logo(slide, position='top_right', prs=None):
    """Add Critical Start logo to slides.
    
    Uses the approved PNG logo file (assets/critical-start-logo.png) if available,
    otherwise falls back to SVG format.
    
    Args:
        slide: The slide object to add the logo to.
        position (str): Position for logo placement. Options:
            - 'top_left': Top left corner
            - 'top_right': Top right corner (default)
            - 'bottom_left': Bottom left corner
            - 'bottom_right': Bottom right corner
        prs (Presentation, optional): Presentation object for getting dimensions.
    
    Returns:
        Picture: The picture shape object, or None if logo file not found.
    """
    # Use approved PNG logo (preferred)
    logo_path = Path("assets/critical-start-logo.png")
    
    # Fallback to SVG if PNG doesn't exist
    if not logo_path.exists():
        logo_path = Path("assets/critical-start-logo.svg")
        if not logo_path.exists():
            print(f"Warning: Logo file not found.")
            print(f"  Expected: assets/critical-start-logo.png (preferred)")
            print(f"  Fallback: assets/critical-start-logo.svg")
            return None
    
    # Define logo size (maintain aspect ratio)
    logo_width = Inches(1.6)  # ~160px at 96 DPI
    logo_height = Inches(0.35)  # ~35px at 96 DPI
    
    # Get slide dimensions (default to 16:9 widescreen)
    if prs:
        slide_width = prs.slide_width
        slide_height = prs.slide_height
    else:
        # Default to 16:9 widescreen dimensions
        slide_width = Inches(10)
        slide_height = Inches(5.625)
    
    if position == 'top_left':
        left = Inches(0.5)
        top = Inches(0.3)
    elif position == 'top_right':
        left = slide_width - logo_width - Inches(0.5)
        top = Inches(0.3)
    elif position == 'bottom_left':
        left = Inches(0.5)
        top = slide_height - logo_height - Inches(0.3)
    elif position == 'bottom_right':
        left = slide_width - logo_width - Inches(0.5)
        top = slide_height - logo_height - Inches(0.3)
    else:
        # Default to top_right
        left = slide_width - logo_width - Inches(0.5)
        top = Inches(0.3)
    
    try:
        # Add picture to slide
        picture = slide.shapes.add_picture(str(logo_path), left, top, 
                                          width=logo_width, height=logo_height)
        return picture
    except Exception as e:
        print(f"Error adding logo: {e}")
        return None


def create_title_slide_layout(prs, title_text, subtitle_text=None):
    """Create title slide with logo and gradient background.
    
    Args:
        prs (Presentation): The presentation object.
        title_text (str): Main title text.
        subtitle_text (str, optional): Subtitle text.
    
    Returns:
        Slide: The created slide object.
    """
    # Use blank layout for full control
    blank_slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(blank_slide_layout)
    
    # Add gradient background (blue to navy)
    # Note: python-pptx gradient support is limited, so we'll use a solid color
    # with a shape overlay for gradient effect
    background_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, 
        prs.slide_width, prs.slide_height
    )
    fill = background_shape.fill
    fill.solid()
    fill.fore_color.rgb = CS_NAVY  # Use navy as base
    background_shape.line.fill.background()  # No border
    
    # Add logo at top
    add_logo(slide, position='top_left', prs=prs)
    
    # Add title
    title_left = Inches(1)
    title_top = Inches(2)
    title_width = prs.slide_width - Inches(2)
    title_height = Inches(1.5)
    
    title_box = slide.shapes.add_textbox(title_left, title_top, 
                                        title_width, title_height)
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    title_paragraph = title_frame.paragraphs[0]
    title_paragraph.text = title_text
    title_paragraph.font.name = TITLE_FONT_NAME
    title_paragraph.font.size = TITLE_FONT_SIZE
    title_paragraph.font.bold = True
    title_paragraph.font.color.rgb = RGBColor(255, 255, 255)  # White text
    title_paragraph.alignment = PP_ALIGN.LEFT
    
    # Add subtitle if provided
    if subtitle_text:
        subtitle_left = Inches(1)
        subtitle_top = Inches(3.5)
        subtitle_width = prs.slide_width - Inches(2)
        subtitle_height = Inches(1)
        
        subtitle_box = slide.shapes.add_textbox(subtitle_left, subtitle_top,
                                               subtitle_width, subtitle_height)
        subtitle_frame = subtitle_box.text_frame
        subtitle_frame.word_wrap = True
        subtitle_paragraph = subtitle_frame.paragraphs[0]
        subtitle_paragraph.text = subtitle_text
        subtitle_paragraph.font.name = BODY_FONT_NAME
        subtitle_paragraph.font.size = Pt(20)
        subtitle_paragraph.font.color.rgb = RGBColor(255, 255, 255)  # White text
        subtitle_paragraph.alignment = PP_ALIGN.LEFT
    
    return slide


def create_section_header_layout(prs, section_title):
    """Create section divider slide.
    
    Args:
        prs (Presentation): The presentation object.
        section_title (str): Section title text.
    
    Returns:
        Slide: The created slide object.
    """
    blank_slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(blank_slide_layout)
    
    # Add gradient background
    background_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0,
        prs.slide_width, prs.slide_height
    )
    fill = background_shape.fill
    fill.solid()
    fill.fore_color.rgb = CS_BLUE  # Use blue as base
    background_shape.line.fill.background()
    
    # Add logo
    add_logo(slide, position='top_right', prs=prs)
    
    # Add section title (centered)
    title_left = Inches(1)
    title_top = Inches(2)
    title_width = prs.slide_width - Inches(2)
    title_height = Inches(1.5)
    
    title_box = slide.shapes.add_textbox(title_left, title_top,
                                        title_width, title_height)
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    title_paragraph = title_frame.paragraphs[0]
    title_paragraph.text = section_title
    title_paragraph.font.name = TITLE_FONT_NAME
    title_paragraph.font.size = Pt(44)
    title_paragraph.font.bold = True
    title_paragraph.font.color.rgb = RGBColor(255, 255, 255)  # White text
    title_paragraph.alignment = PP_ALIGN.CENTER
    
    return slide


def create_content_slide_layout(prs, slide_title, content_items=None):
    """Create standard content slide with header.
    
    Args:
        prs (Presentation): The presentation object.
        slide_title (str): Slide title text.
        content_items (list, optional): List of content strings for bullet points.
    
    Returns:
        Slide: The created slide object.
    """
    blank_slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(blank_slide_layout)
    
    # Add logo at top right
    add_logo(slide, position='top_right', prs=prs)
    
    # Add title header with brand color background
    header_height = Inches(0.8)
    header_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0,
        prs.slide_width, header_height
    )
    fill = header_shape.fill
    fill.solid()
    fill.fore_color.rgb = CS_NAVY
    header_shape.line.fill.background()
    
    # Add title text on header
    title_left = Inches(0.5)
    title_top = Inches(0.1)
    title_width = prs.slide_width - Inches(2.5)  # Leave space for logo
    title_height = Inches(0.6)
    
    title_box = slide.shapes.add_textbox(title_left, title_top,
                                        title_width, title_height)
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    title_paragraph = title_frame.paragraphs[0]
    title_paragraph.text = slide_title
    title_paragraph.font.name = TITLE_FONT_NAME
    title_paragraph.font.size = Pt(28)
    title_paragraph.font.bold = True
    title_paragraph.font.color.rgb = RGBColor(255, 255, 255)  # White text
    title_paragraph.alignment = PP_ALIGN.LEFT
    
    # Add content area if items provided
    if content_items:
        content_left = Inches(0.5)
        content_top = header_height + Inches(0.3)
        content_width = prs.slide_width - Inches(1)
        content_height = prs.slide_height - content_top - Inches(0.3)
        
        content_box = slide.shapes.add_textbox(content_left, content_top,
                                             content_width, content_height)
        content_frame = content_box.text_frame
        content_frame.word_wrap = True
        content_frame.margin_left = Inches(0.3)
        content_frame.margin_right = Inches(0.3)
        content_frame.margin_top = Inches(0.2)
        content_frame.margin_bottom = Inches(0.2)
        
        # Add content items as bullet points
        for i, item in enumerate(content_items):
            if i == 0:
                paragraph = content_frame.paragraphs[0]
            else:
                paragraph = content_frame.add_paragraph()
            
            paragraph.text = item
            paragraph.font.name = BODY_FONT_NAME
            paragraph.font.size = BODY_FONT_SIZE
            paragraph.font.color.rgb = CS_SLATE
            paragraph.level = 0
            paragraph.space_after = Pt(12)
    
    return slide


def main():
    """Main function to orchestrate slide generation."""
    pass


def build_executive_summary_slides(prs, data):
    """Create the Title slide and Executive Summary slides (Slides 1-3).
    
    Args:
        prs (Presentation): The presentation object.
        data (ReportData): The report data object containing all metrics.
    """
    # Slide 1 - Title Slide
    blank_slide_layout = prs.slide_layouts[6]  # Blank layout
    slide1 = prs.slides.add_slide(blank_slide_layout)
    
    # Add navy background
    background_shape = slide1.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0,
        prs.slide_width, prs.slide_height
    )
    fill = background_shape.fill
    fill.solid()
    fill.fore_color.rgb = CS_NAVY
    background_shape.line.fill.background()
    
    # Add logo at top left
    add_logo(slide1, position='top_left', prs=prs)
    
    # Add main title
    title_left = Inches(1)
    title_top = Inches(1.8)
    title_width = prs.slide_width - Inches(2)
    title_height = Inches(0.8)
    
    title_box = slide1.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    title_paragraph = title_frame.paragraphs[0]
    title_paragraph.text = "Escalation to Client Details Report"
    title_paragraph.font.name = TITLE_FONT_NAME
    title_paragraph.font.size = Pt(44)
    title_paragraph.font.bold = True
    title_paragraph.font.color.rgb = RGBColor(255, 255, 255)
    title_paragraph.alignment = PP_ALIGN.LEFT
    
    # Add subtitle
    subtitle_left = Inches(1)
    subtitle_top = Inches(2.6)
    subtitle_width = prs.slide_width - Inches(2)
    subtitle_height = Inches(0.5)
    
    subtitle_box = slide1.shapes.add_textbox(subtitle_left, subtitle_top, subtitle_width, subtitle_height)
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.word_wrap = True
    subtitle_paragraph = subtitle_frame.paragraphs[0]
    subtitle_paragraph.text = data.tier
    subtitle_paragraph.font.name = BODY_FONT_NAME
    subtitle_paragraph.font.size = Pt(24)
    subtitle_paragraph.font.color.rgb = RGBColor(255, 255, 255)
    subtitle_paragraph.alignment = PP_ALIGN.LEFT
    
    # Add client name
    client_left = Inches(1)
    client_top = Inches(3.3)
    client_width = prs.slide_width - Inches(2)
    client_height = Inches(0.5)
    
    client_box = slide1.shapes.add_textbox(client_left, client_top, client_width, client_height)
    client_frame = client_box.text_frame
    client_frame.word_wrap = True
    client_paragraph = client_frame.paragraphs[0]
    client_paragraph.text = data.client_name
    client_paragraph.font.name = BODY_FONT_NAME
    client_paragraph.font.size = Pt(20)
    client_paragraph.font.color.rgb = RGBColor(255, 255, 255)
    client_paragraph.alignment = PP_ALIGN.LEFT
    
    # Add period
    # Format: "August 1-31, 2025 (31 days)"
    start_month_day = data.period_start.split(',')[0].strip()  # "August 1"
    end_month_day = data.period_end.split(',')[0].strip()  # "August 31"
    year = data.period_end.split(',')[1].strip()  # "2025"
    
    # Extract month and day numbers
    start_parts = start_month_day.split()
    end_parts = end_month_day.split()
    
    if len(start_parts) == 2 and len(end_parts) == 2 and start_parts[0] == end_parts[0]:
        # Same month, format as "August 1-31, 2025"
        period_text = f"{start_parts[0]} {start_parts[1]}-{end_parts[1]}, {year} ({data.period_days} days)"
    else:
        # Different months, use full format
        period_text = f"{start_month_day} - {end_month_day}, {year} ({data.period_days} days)"
    period_left = Inches(1)
    period_top = Inches(3.9)
    period_width = prs.slide_width - Inches(2)
    period_height = Inches(0.5)
    
    period_box = slide1.shapes.add_textbox(period_left, period_top, period_width, period_height)
    period_frame = period_box.text_frame
    period_frame.word_wrap = True
    period_paragraph = period_frame.paragraphs[0]
    period_paragraph.text = period_text
    period_paragraph.font.name = BODY_FONT_NAME
    period_paragraph.font.size = Pt(18)
    period_paragraph.font.color.rgb = RGBColor(255, 255, 255)
    period_paragraph.alignment = PP_ALIGN.LEFT
    
    # Add report date at bottom
    report_date_left = Inches(1)
    report_date_top = prs.slide_height - Inches(0.6)
    report_date_width = prs.slide_width - Inches(2)
    report_date_height = Inches(0.4)
    
    report_date_box = slide1.shapes.add_textbox(report_date_left, report_date_top, report_date_width, report_date_height)
    report_date_frame = report_date_box.text_frame
    report_date_frame.word_wrap = True
    report_date_paragraph = report_date_frame.paragraphs[0]
    report_date_paragraph.text = f"Report Date: {data.report_date}"
    report_date_paragraph.font.name = BODY_FONT_NAME
    report_date_paragraph.font.size = Pt(14)
    report_date_paragraph.font.color.rgb = RGBColor(255, 255, 255)
    report_date_paragraph.alignment = PP_ALIGN.LEFT
    
    # Slide 2 - Key Metrics Overview
    slide2 = prs.slides.add_slide(blank_slide_layout)
    
    # Add logo at top right
    add_logo(slide2, position='top_right', prs=prs)
    
    # Add title header
    header_height = Inches(0.8)
    header_shape = slide2.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0,
        prs.slide_width, header_height
    )
    fill = header_shape.fill
    fill.solid()
    fill.fore_color.rgb = CS_NAVY
    header_shape.line.fill.background()
    
    # Add title text on header
    slide2_title_left = Inches(0.5)
    slide2_title_top = Inches(0.1)
    slide2_title_width = prs.slide_width - Inches(2.5)
    slide2_title_height = Inches(0.6)
    
    slide2_title_box = slide2.shapes.add_textbox(slide2_title_left, slide2_title_top, slide2_title_width, slide2_title_height)
    slide2_title_frame = slide2_title_box.text_frame
    slide2_title_frame.word_wrap = True
    slide2_title_paragraph = slide2_title_frame.paragraphs[0]
    slide2_title_paragraph.text = "Executive Summary"
    slide2_title_paragraph.font.name = TITLE_FONT_NAME
    slide2_title_paragraph.font.size = Pt(28)
    slide2_title_paragraph.font.bold = True
    slide2_title_paragraph.font.color.rgb = RGBColor(255, 255, 255)
    slide2_title_paragraph.alignment = PP_ALIGN.LEFT
    
    # Add large headline metric
    metric_left = Inches(1)
    metric_top = header_height + Inches(0.8)
    metric_width = prs.slide_width - Inches(2)
    metric_height = Inches(1.2)
    
    metric_box = slide2.shapes.add_textbox(metric_left, metric_top, metric_width, metric_height)
    metric_frame = metric_box.text_frame
    metric_frame.word_wrap = True
    metric_paragraph = metric_frame.paragraphs[0]
    metric_paragraph.text = f"{data.incidents_escalated} Incidents Escalated"
    metric_paragraph.font.name = TITLE_FONT_NAME
    metric_paragraph.font.size = Pt(56)
    metric_paragraph.font.bold = True
    metric_paragraph.font.color.rgb = CS_NAVY
    metric_paragraph.alignment = PP_ALIGN.LEFT
    
    # Add subtext
    subtext_left = Inches(1)
    subtext_top = metric_top + Inches(1.3)
    subtext_width = prs.slide_width - Inches(2)
    subtext_height = Inches(0.6)
    
    subtext_box = slide2.shapes.add_textbox(subtext_left, subtext_top, subtext_width, subtext_height)
    subtext_frame = subtext_box.text_frame
    subtext_frame.word_wrap = True
    subtext_paragraph = subtext_frame.paragraphs[0]
    subtext_paragraph.text = f"Average {data.incidents_per_day} per day requiring your team's attention"
    subtext_paragraph.font.name = BODY_FONT_NAME
    subtext_paragraph.font.size = Pt(20)
    subtext_paragraph.font.color.rgb = CS_SLATE
    subtext_paragraph.alignment = PP_ALIGN.LEFT
    
    # Add additional context
    context_left = Inches(1)
    context_top = subtext_top + Inches(0.8)
    context_width = prs.slide_width - Inches(2)
    context_height = Inches(1)
    
    context_box = slide2.shapes.add_textbox(context_left, context_top, context_width, context_height)
    context_frame = context_box.text_frame
    context_frame.word_wrap = True
    context_paragraph = context_frame.paragraphs[0]
    client_touch_percent = int((data.client_touch_decisions / data.alerts_triaged) * 100) if data.alerts_triaged > 0 else 0
    context_paragraph.text = f"CS SOC triaged {data.alerts_triaged:,} alerts this period—{data.client_touch_decisions:,} ({client_touch_percent}%) guided with your team and {data.closed_end_to_end:,} closed end-to-end"
    context_paragraph.font.name = BODY_FONT_NAME
    context_paragraph.font.size = Pt(16)
    context_paragraph.font.color.rgb = CS_SLATE
    context_paragraph.alignment = PP_ALIGN.LEFT
    
    # Slide 3 - Period Highlights
    slide3 = prs.slides.add_slide(blank_slide_layout)
    
    # Add logo at top right
    add_logo(slide3, position='top_right', prs=prs)
    
    # Add title header
    header_shape3 = slide3.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0,
        prs.slide_width, header_height
    )
    fill = header_shape3.fill
    fill.solid()
    fill.fore_color.rgb = CS_NAVY
    header_shape3.line.fill.background()
    
    # Add title text on header
    slide3_title_box = slide3.shapes.add_textbox(slide2_title_left, slide2_title_top, slide2_title_width, slide2_title_height)
    slide3_title_frame = slide3_title_box.text_frame
    slide3_title_frame.word_wrap = True
    slide3_title_paragraph = slide3_title_frame.paragraphs[0]
    slide3_title_paragraph.text = "Period Highlights"
    slide3_title_paragraph.font.name = TITLE_FONT_NAME
    slide3_title_paragraph.font.size = Pt(28)
    slide3_title_paragraph.font.bold = True
    slide3_title_paragraph.font.color.rgb = RGBColor(255, 255, 255)
    slide3_title_paragraph.alignment = PP_ALIGN.LEFT
    
    # Calculate after-hours percentage
    after_hours_percent = int((data.after_hours_escalations / data.incidents_escalated) * 100) if data.incidents_escalated > 0 else 0
    
    # Define 4 highlight cards
    highlights = [
        {
            "title": "Threat Outcomes",
            "value": f"{data.true_threats_contained} threats contained, zero breaches"
        },
        {
            "title": "Response Advantage",
            "value": f"{int(data.response_advantage_percent)}% faster MTTR vs peers ({data.mttr_minutes}m vs {data.industry_median_minutes}m)"
        },
        {
            "title": "24/7 Protection",
            "value": f"{after_hours_percent}% of escalations handled after-hours"
        },
        {
            "title": "Detection Quality",
            "value": f"{data.false_positive_rate}% false positive rate"
        }
    ]
    
    # Position cards in a 2x2 grid
    card_width = (prs.slide_width - Inches(2.5)) / 2  # Two columns with margins
    card_height = Inches(1.5)
    card_spacing = Inches(0.3)
    
    start_left = Inches(0.5)
    start_top = header_height + Inches(0.5)
    
    for i, highlight in enumerate(highlights):
        row = i // 2
        col = i % 2
        
        card_left = start_left + col * (card_width + Inches(0.3))
        card_top = start_top + row * (card_height + card_spacing)
        
        # Create card background with light blue tint
        card_shape = slide3.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, card_left, card_top,
            card_width, card_height
        )
        fill = card_shape.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(240, 248, 255)  # Light blue background
        line = card_shape.line
        line.color.rgb = CS_BLUE
        line.width = Pt(2)
        
        # Add card title
        card_title_left = card_left + Inches(0.2)
        card_title_top = card_top + Inches(0.15)
        card_title_width = card_width - Inches(0.4)
        card_title_height = Inches(0.4)
        
        card_title_box = slide3.shapes.add_textbox(card_title_left, card_title_top, card_title_width, card_title_height)
        card_title_frame = card_title_box.text_frame
        card_title_frame.word_wrap = True
        card_title_paragraph = card_title_frame.paragraphs[0]
        card_title_paragraph.text = highlight["title"]
        card_title_paragraph.font.name = TITLE_FONT_NAME
        card_title_paragraph.font.size = Pt(18)
        card_title_paragraph.font.bold = True
        card_title_paragraph.font.color.rgb = CS_NAVY
        card_title_paragraph.alignment = PP_ALIGN.LEFT
        
        # Add card value
        card_value_left = card_left + Inches(0.2)
        card_value_top = card_title_top + Inches(0.45)
        card_value_width = card_width - Inches(0.4)
        card_value_height = Inches(0.8)
        
        card_value_box = slide3.shapes.add_textbox(card_value_left, card_value_top, card_value_width, card_value_height)
        card_value_frame = card_value_box.text_frame
        card_value_frame.word_wrap = True
        card_value_paragraph = card_value_frame.paragraphs[0]
        card_value_paragraph.text = highlight["value"]
        card_value_paragraph.font.name = BODY_FONT_NAME
        card_value_paragraph.font.size = Pt(14)
        card_value_paragraph.font.color.rgb = CS_SLATE
        card_value_paragraph.alignment = PP_ALIGN.LEFT


def create_executive_summary_slide(prs, report_data):
    """Create the Executive Summary slide."""
    pass


def build_value_delivered_slides(prs, data):
    """Create the Value Delivered section (Slides 4-6).
    
    Args:
        prs (Presentation): The presentation object.
        data (ReportData): The report data object containing all metrics.
    """
    blank_slide_layout = prs.slide_layouts[6]  # Blank layout
    
    # Slide 4 - Cost Avoidance Hero
    slide4 = prs.slides.add_slide(blank_slide_layout)
    
    # Add logo at top right
    add_logo(slide4, position='top_right', prs=prs)
    
    # Add title
    title_left = Inches(0.5)
    title_top = Inches(0.3)
    title_width = prs.slide_width - Inches(2.5)
    title_height = Inches(0.6)
    
    title_box = slide4.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    title_paragraph = title_frame.paragraphs[0]
    title_paragraph.text = "Modeled Cost Exposure Avoided"
    title_paragraph.font.name = TITLE_FONT_NAME
    title_paragraph.font.size = Pt(32)
    title_paragraph.font.bold = True
    title_paragraph.font.color.rgb = CS_NAVY
    title_paragraph.alignment = PP_ALIGN.CENTER
    
    # Add hero number (~$7.55M)
    # Format total_modeled: 7550000 -> $7.55M
    total_millions = data.total_modeled / 1000000
    hero_value = f"~${total_millions:.2f}M"
    
    hero_left = Inches(0.5)
    hero_top = Inches(1.5)
    hero_width = prs.slide_width - Inches(1)
    hero_height = Inches(1.5)
    
    hero_box = slide4.shapes.add_textbox(hero_left, hero_top, hero_width, hero_height)
    hero_frame = hero_box.text_frame
    hero_frame.word_wrap = True
    hero_paragraph = hero_frame.paragraphs[0]
    hero_paragraph.text = hero_value
    hero_paragraph.font.name = TITLE_FONT_NAME
    hero_paragraph.font.size = Pt(80)
    hero_paragraph.font.bold = True
    hero_paragraph.font.color.rgb = CS_NAVY
    hero_paragraph.alignment = PP_ALIGN.CENTER
    
    # Add subtitle
    subtitle_left = Inches(0.5)
    subtitle_top = Inches(3.2)
    subtitle_width = prs.slide_width - Inches(1)
    subtitle_height = Inches(0.6)
    
    subtitle_box = slide4.shapes.add_textbox(subtitle_left, subtitle_top, subtitle_width, subtitle_height)
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.word_wrap = True
    subtitle_paragraph = subtitle_frame.paragraphs[0]
    subtitle_paragraph.text = "Modeled cost exposure avoided across operations, coverage, and threat containment"
    subtitle_paragraph.font.name = BODY_FONT_NAME
    subtitle_paragraph.font.size = Pt(20)
    subtitle_paragraph.font.color.rgb = CS_SLATE
    subtitle_paragraph.alignment = PP_ALIGN.CENTER
    
    # Add note at bottom
    note_left = Inches(0.5)
    note_top = prs.slide_height - Inches(0.8)
    note_width = prs.slide_width - Inches(1)
    note_height = Inches(0.4)
    
    note_box = slide4.shapes.add_textbox(note_left, note_top, note_width, note_height)
    note_frame = note_box.text_frame
    note_frame.word_wrap = True
    note_paragraph = note_frame.paragraphs[0]
    note_paragraph.text = "Illustrative impact only; not redeployable budget"
    note_paragraph.font.name = BODY_FONT_NAME
    note_paragraph.font.size = Pt(14)
    note_paragraph.font.color.rgb = CS_SLATE
    note_paragraph.alignment = PP_ALIGN.CENTER
    note_paragraph.font.italic = True
    
    # Slide 5 - Value Breakdown
    slide5 = prs.slides.add_slide(blank_slide_layout)
    
    # Add logo at top right
    add_logo(slide5, position='top_right', prs=prs)
    
    # Add title header
    header_height = Inches(0.8)
    header_shape = slide5.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0,
        prs.slide_width, header_height
    )
    fill = header_shape.fill
    fill.solid()
    fill.fore_color.rgb = CS_NAVY
    header_shape.line.fill.background()
    
    # Add title text on header
    slide5_title_left = Inches(0.5)
    slide5_title_top = Inches(0.1)
    slide5_title_width = prs.slide_width - Inches(2.5)
    slide5_title_height = Inches(0.6)
    
    slide5_title_box = slide5.shapes.add_textbox(slide5_title_left, slide5_title_top, slide5_title_width, slide5_title_height)
    slide5_title_frame = slide5_title_box.text_frame
    slide5_title_frame.word_wrap = True
    slide5_title_paragraph = slide5_title_frame.paragraphs[0]
    slide5_title_paragraph.text = "Value Delivered - Breakdown"
    slide5_title_paragraph.font.name = TITLE_FONT_NAME
    slide5_title_paragraph.font.size = Pt(28)
    slide5_title_paragraph.font.bold = True
    slide5_title_paragraph.font.color.rgb = RGBColor(255, 255, 255)
    slide5_title_paragraph.alignment = PP_ALIGN.LEFT
    
    # Create 3-column layout with cards
    card_width = (prs.slide_width - Inches(2.2)) / 3  # Three columns with margins
    card_height = Inches(3.2)
    card_spacing = Inches(0.2)
    
    start_left = Inches(0.5)
    start_top = header_height + Inches(0.4)
    
    # Format values
    analyst_k = data.analyst_cost_equivalent / 1000
    coverage_k = data.coverage_cost_equivalent / 1000
    breach_m = data.breach_exposure_avoided / 1000000
    
    # Define 3 value cards
    value_cards = [
        {
            "title": "Security Operations",
            "value1": f"{data.analyst_hours} analyst hours delivered",
            "value2": f"~${analyst_k:.0f}K equivalent"
        },
        {
            "title": "24/7 Coverage",
            "value1": f"{data.coverage_hours} hours of monitoring",
            "value2": f"~${coverage_k:.0f}K equivalent"
        },
        {
            "title": "Threat Prevention",
            "value1": f"{data.true_threats_contained} intrusions contained",
            "value2": f"~${breach_m:.2f}M breach exposure avoided"
        }
    ]
    
    for i, card in enumerate(value_cards):
        card_left = start_left + i * (card_width + card_spacing)
        card_top = start_top
        
        # Create card background with light blue tint
        card_shape = slide5.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, card_left, card_top,
            card_width, card_height
        )
        fill = card_shape.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(240, 248, 255)  # Light blue background
        line = card_shape.line
        line.color.rgb = CS_BLUE
        line.width = Pt(2)
        
        # Add card title
        card_title_left = card_left + Inches(0.2)
        card_title_top = card_top + Inches(0.3)
        card_title_width = card_width - Inches(0.4)
        card_title_height = Inches(0.5)
        
        card_title_box = slide5.shapes.add_textbox(card_title_left, card_title_top, card_title_width, card_title_height)
        card_title_frame = card_title_box.text_frame
        card_title_frame.word_wrap = True
        card_title_paragraph = card_title_frame.paragraphs[0]
        card_title_paragraph.text = card["title"]
        card_title_paragraph.font.name = TITLE_FONT_NAME
        card_title_paragraph.font.size = Pt(20)
        card_title_paragraph.font.bold = True
        card_title_paragraph.font.color.rgb = CS_NAVY
        card_title_paragraph.alignment = PP_ALIGN.CENTER
        
        # Add value 1
        card_value1_left = card_left + Inches(0.2)
        card_value1_top = card_title_top + Inches(0.7)
        card_value1_width = card_width - Inches(0.4)
        card_value1_height = Inches(0.8)
        
        card_value1_box = slide5.shapes.add_textbox(card_value1_left, card_value1_top, card_value1_width, card_value1_height)
        card_value1_frame = card_value1_box.text_frame
        card_value1_frame.word_wrap = True
        card_value1_paragraph = card_value1_frame.paragraphs[0]
        card_value1_paragraph.text = card["value1"]
        card_value1_paragraph.font.name = BODY_FONT_NAME
        card_value1_paragraph.font.size = Pt(16)
        card_value1_paragraph.font.color.rgb = CS_SLATE
        card_value1_paragraph.alignment = PP_ALIGN.CENTER
        
        # Add value 2 (cost equivalent)
        card_value2_left = card_left + Inches(0.2)
        card_value2_top = card_value1_top + Inches(1.0)
        card_value2_width = card_width - Inches(0.4)
        card_value2_height = Inches(0.6)
        
        card_value2_box = slide5.shapes.add_textbox(card_value2_left, card_value2_top, card_value2_width, card_value2_height)
        card_value2_frame = card_value2_box.text_frame
        card_value2_frame.word_wrap = True
        card_value2_paragraph = card_value2_frame.paragraphs[0]
        card_value2_paragraph.text = card["value2"]
        card_value2_paragraph.font.name = TITLE_FONT_NAME
        card_value2_paragraph.font.size = Pt(24)
        card_value2_paragraph.font.bold = True
        card_value2_paragraph.font.color.rgb = CS_NAVY
        card_value2_paragraph.alignment = PP_ALIGN.CENTER
    
    # Add footer methodology note
    footer_left = Inches(0.5)
    footer_top = start_top + card_height + Inches(0.3)
    footer_width = prs.slide_width - Inches(1)
    footer_height = Inches(0.4)
    
    footer_box = slide5.shapes.add_textbox(footer_left, footer_top, footer_width, footer_height)
    footer_frame = footer_box.text_frame
    footer_frame.word_wrap = True
    footer_paragraph = footer_frame.paragraphs[0]
    footer_paragraph.text = "Methodology: Cost equivalents based on industry-standard rates for SOC analyst time, 24/7 monitoring infrastructure, and average breach cost per incident."
    footer_paragraph.font.name = BODY_FONT_NAME
    footer_paragraph.font.size = Pt(11)
    footer_paragraph.font.color.rgb = CS_SLATE
    footer_paragraph.alignment = PP_ALIGN.CENTER
    footer_paragraph.font.italic = True
    
    # Slide 6 - Security Outcomes
    slide6 = prs.slides.add_slide(blank_slide_layout)
    
    # Add logo at top right
    add_logo(slide6, position='top_right', prs=prs)
    
    # Add title header
    header_shape6 = slide6.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0,
        prs.slide_width, header_height
    )
    fill = header_shape6.fill
    fill.solid()
    fill.fore_color.rgb = CS_NAVY
    header_shape6.line.fill.background()
    
    # Add title text on header
    slide6_title_box = slide6.shapes.add_textbox(slide5_title_left, slide5_title_top, slide5_title_width, slide5_title_height)
    slide6_title_frame = slide6_title_box.text_frame
    slide6_title_frame.word_wrap = True
    slide6_title_paragraph = slide6_title_frame.paragraphs[0]
    slide6_title_paragraph.text = "Security Outcomes This Period"
    slide6_title_paragraph.font.name = TITLE_FONT_NAME
    slide6_title_paragraph.font.size = Pt(28)
    slide6_title_paragraph.font.bold = True
    slide6_title_paragraph.font.color.rgb = RGBColor(255, 255, 255)
    slide6_title_paragraph.alignment = PP_ALIGN.LEFT
    
    # Create checklist/bullet format with outcomes
    outcomes = [
        f"{data.alerts_triaged:,} Alerts Triaged",
        f"{data.client_touch_decisions:,} Client-Touch Decisions",
        f"{data.closed_end_to_end:,} Closed End-to-End",
        f"{data.true_threats_contained} True Threats Contained",
        f"{int(data.response_advantage_percent)}% Faster Than Industry",
        f"{data.p90_minutes}-Minute P90 Response",
        f"{data.after_hours_escalations} After-Hours Escalations",
        f"{data.coverage_hours} Hours of Coverage"
    ]
    
    # Create two-column layout for outcomes
    content_left = Inches(0.8)
    content_top = header_height + Inches(0.5)
    content_width = prs.slide_width - Inches(1.6)
    content_height = prs.slide_height - content_top - Inches(0.3)
    
    # Use textbox with bullet points
    content_box = slide6.shapes.add_textbox(content_left, content_top, content_width, content_height)
    content_frame = content_box.text_frame
    content_frame.word_wrap = True
    content_frame.margin_left = Inches(0.3)
    content_frame.margin_right = Inches(0.3)
    content_frame.margin_top = Inches(0.2)
    content_frame.margin_bottom = Inches(0.2)
    
    # Add outcomes as bullet points
    for i, outcome in enumerate(outcomes):
        if i == 0:
            paragraph = content_frame.paragraphs[0]
        else:
            paragraph = content_frame.add_paragraph()
        
        # Add bullet character to text
        paragraph.text = "• " + outcome
        paragraph.font.name = BODY_FONT_NAME
        paragraph.font.size = Pt(18)
        paragraph.font.color.rgb = CS_SLATE
        paragraph.level = 0
        paragraph.space_after = Pt(12)
        paragraph.font.bold = False


def create_protection_achieved_slide(prs, report_data):
    """Create the How Protection Was Achieved slide."""
    pass


def create_threat_landscape_slide(prs, report_data):
    """Create the Threat Landscape slide (optional)."""
    pass


def create_insights_opportunities_slide(prs, report_data):
    """Create the Insights and Opportunities slide."""
    pass


def create_forward_direction_slide(prs, report_data):
    """Create the Forward Direction slide."""
    pass


if __name__ == "__main__":
    main()

