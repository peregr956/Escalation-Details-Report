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


def create_executive_summary_slide(prs, report_data):
    """Create the Executive Summary slide."""
    pass


def create_value_delivered_slide(prs, report_data):
    """Create the Value Delivered slide."""
    pass


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

