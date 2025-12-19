from pptx import Presentation
from pptx.util import Inches, Pt, RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pathlib import Path
import argparse
import logging
import shutil
from datetime import datetime
from typing import Dict, Optional

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


def prepare_chart_data(data) -> Dict:
    """Prepare chart data dictionary from ReportData for chart rendering.
    
    Args:
        data: ReportData instance containing all metrics
    
    Returns:
        Dictionary with chart data structured for different chart types
    """
    chart_data = {}
    
    # Trend chart data (MTTR, MTTD, FP% trends)
    chart_data['trend'] = {
        'labels': data.period_labels,
        'datasets': [
            {
                'label': 'MTTR (min)',
                'data': data.mttr_trend,
                'borderColor': '#009CDE',
                'backgroundColor': 'rgba(0, 156, 222, 0.1)',
                'borderWidth': 3,
                'tension': 0.4,
                'pointRadius': 4,
                'pointHoverRadius': 6
            },
            {
                'label': 'MTTD (min)',
                'data': data.mttd_trend,
                'borderColor': '#004C97',
                'backgroundColor': 'rgba(0, 76, 151, 0.1)',
                'borderWidth': 3,
                'tension': 0.4,
                'pointRadius': 4,
                'pointHoverRadius': 6
            },
            {
                'label': 'FP %',
                'data': data.fp_trend,
                'borderColor': '#EF3340',
                'backgroundColor': 'rgba(239, 51, 64, 0.1)',
                'borderWidth': 3,
                'borderDash': [8, 4],
                'tension': 0.4,
                'pointRadius': 4,
                'pointHoverRadius': 6,
                'yAxisID': 'y1'
            }
        ]
    }
    
    # Pie chart data (Operational Load)
    chart_data['pie'] = {
        'labels': ['Business Hours', 'After Hours', 'Weekend'],
        'data': [
            int(data.business_hours_percent),
            int(data.after_hours_percent),
            int(data.weekend_percent)
        ],
        'backgroundColor': ['#009CDE', '#702F8A', '#EF3340']
    }
    
    # Sankey chart data (Severity Flow)
    chart_data['sankey'] = {
        'flows': data.severity_flows
    }
    
    # Stacked bar chart data (MITRE ATT&CK)
    chart_data['stacked_bar'] = {
        'labels': data.tactics,
        'datasets': [
            {
                'label': 'High',
                'data': data.high_severity,
                'backgroundColor': '#EF3340'
            },
            {
                'label': 'Medium',
                'data': data.medium_severity,
                'backgroundColor': '#FF6A14'
            },
            {
                'label': 'Low',
                'data': data.low_severity,
                'backgroundColor': '#009CDE'
            },
            {
                'label': 'Info',
                'data': data.info_severity,
                'backgroundColor': '#702F8A'
            }
        ]
    }
    
    return chart_data


def insert_chart_image(slide, placeholder_id: Optional[str], image_path: str):
    """Replace a placeholder shape with a chart image.
    
    Args:
        slide: The slide object containing the placeholder
        placeholder_id: Optional ID string to identify the placeholder (e.g., "severity_sankey")
                       If None, searches for first placeholder with "[Chart:" text
        image_path: Path to the image file to insert
    
    Returns:
        bool: True if placeholder was found and replaced, False otherwise
    """
    image_path_obj = Path(image_path)
    if not image_path_obj.exists():
        logging.warning(f"Chart image not found: {image_path}")
        return False
    
    # Find the placeholder shape
    placeholder_shape = None
    
    for shape in slide.shapes:
        if hasattr(shape, 'text_frame') and shape.text_frame:
            text = shape.text_frame.text
            if "[Chart:" in text:
                # If placeholder_id is specified, check if it matches
                if placeholder_id:
                    if f"ID: {placeholder_id}" in text:
                        placeholder_shape = shape
                        break
                else:
                    # Use first placeholder found
                    placeholder_shape = shape
                    break
    
    if not placeholder_shape:
        logging.warning(f"Placeholder not found on slide (ID: {placeholder_id})")
        return False
    
    # Get placeholder position and size
    left = placeholder_shape.left
    top = placeholder_shape.top
    width = placeholder_shape.width
    height = placeholder_shape.height
    
    # Remove the placeholder shape
    slide.shapes._spTree.remove(placeholder_shape._element)
    
    # Insert the image
    try:
        slide.shapes.add_picture(str(image_path), left, top, width, height)
        logging.info(f"Inserted chart image: {image_path} (ID: {placeholder_id})")
        return True
    except Exception as e:
        logging.error(f"Failed to insert chart image {image_path}: {e}")
        return False


def main():
    """Main function to orchestrate slide generation.
    
    This function:
    1. Loads report data
    2. Renders all charts
    3. Creates presentation
    4. Builds all slides in order
    5. Inserts chart images
    6. Saves presentation
    """
    # Set up logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    logger = logging.getLogger(__name__)
    
    # Parse command-line arguments
    parser = argparse.ArgumentParser(
        description='Generate Escalation Details Report PowerPoint presentation'
    )
    parser.add_argument(
        '--client-name',
        type=str,
        help='Client name to use in the report (overrides data default)'
    )
    parser.add_argument(
        '--no-threat-landscape',
        action='store_true',
        help='Exclude threat landscape slides from the presentation'
    )
    parser.add_argument(
        '--output-dir',
        type=str,
        default='output',
        help='Output directory for the presentation file (default: output)'
    )
    parser.add_argument(
        '--keep-charts',
        action='store_true',
        help='Keep temporary chart images after generation (default: delete)'
    )
    
    args = parser.parse_args()
    
    logger.info("=" * 60)
    logger.info("Starting Escalation Details Report Generation")
    logger.info("=" * 60)
    
    # Step 1: Import and load data
    logger.info("Step 1: Loading report data...")
    try:
        from report_data import get_report_data, ReportData
        data = get_report_data()
        
        # Override client name if provided
        if args.client_name:
            data.client_name = args.client_name
            logger.info(f"Using client name: {args.client_name}")
        else:
            logger.info(f"Using client name from data: {data.client_name}")
        
        logger.info(f"✓ Loaded data for {data.client_name}")
        logger.info(f"  Period: {data.period_start} to {data.period_end}")
        logger.info(f"  Incidents escalated: {data.incidents_escalated}")
    except Exception as e:
        logger.error(f"Failed to load report data: {e}")
        return 1
    
    # Step 2: Prepare chart data
    logger.info("Step 2: Preparing chart data...")
    try:
        chart_data = prepare_chart_data(data)
        logger.info("✓ Chart data prepared")
        logger.info(f"  Charts to render: trend, pie, sankey, stacked_bar")
    except Exception as e:
        logger.error(f"Failed to prepare chart data: {e}")
        return 1
    
    # Step 3: Render charts
    logger.info("Step 3: Rendering charts...")
    try:
        from chart_renderer import render_charts_sync
        
        temp_charts_dir = "temp_charts"
        chart_images = render_charts_sync(chart_data, output_dir=temp_charts_dir)
        
        # Log rendered charts
        rendered_count = sum(1 for v in chart_images.values() if v is not None)
        logger.info(f"✓ Rendered {rendered_count}/{len(chart_images)} charts")
        for chart_name, path in chart_images.items():
            if path:
                logger.info(f"  - {chart_name}: {path}")
            else:
                logger.warning(f"  - {chart_name}: FAILED")
    except Exception as e:
        logger.error(f"Failed to render charts: {e}")
        return 1
    
    # Step 4: Create presentation
    logger.info("Step 4: Creating presentation...")
    try:
        prs = create_presentation()
        apply_branding(prs)
        logger.info("✓ Presentation created")
    except Exception as e:
        logger.error(f"Failed to create presentation: {e}")
        return 1
    
    # Step 5: Build slides
    logger.info("Step 5: Building slides...")
    try:
        # Build slides in order
        logger.info("  Building executive summary slides (1-3)...")
        build_executive_summary_slides(prs, data)
        
        logger.info("  Building value delivered slides (4-5)...")
        build_value_delivered_slides(prs, data)
        
        logger.info("  Building protection achieved slides (6-8)...")
        build_protection_achieved_slides(prs, data)
        
        if not args.no_threat_landscape:
            logger.info("  Building threat landscape slides (9-12)...")
            build_threat_landscape_slides(prs, data, include=True)
        else:
            logger.info("  Skipping threat landscape slides (--no-threat-landscape)")
            build_threat_landscape_slides(prs, data, include=False)
        
        logger.info("  Building insights slides (13-15)...")
        build_insights_slides(prs, data)
        
        logger.info("  Building forward direction slide (16)...")
        build_forward_direction_slide(prs, data)
        
        logger.info(f"✓ Built {len(prs.slides)} slides")
    except Exception as e:
        logger.error(f"Failed to build slides: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return 1
    
    # Step 6: Insert chart images
    logger.info("Step 6: Inserting chart images...")
    try:
        charts_inserted = 0
        
        # Search through all slides for chart placeholders
        for slide_idx, slide in enumerate(prs.slides):
            # Check all shapes for chart placeholders
            for shape in slide.shapes:
                if hasattr(shape, 'text_frame') and shape.text_frame:
                    text = shape.text_frame.text
                    
                    # Trend chart - "Performance Trends"
                    if chart_images.get('trend') and "[Chart: Performance Trends]" in text:
                        if insert_chart_image(slide, None, chart_images['trend']):
                            charts_inserted += 1
                            logger.info(f"  Inserted trend chart into slide {slide_idx + 1}")
                            break  # Move to next slide
                    
                    # Sankey chart - "Severity Alignment Sankey"
                    if chart_images.get('sankey') and 'severity_sankey' in text:
                        if insert_chart_image(slide, 'severity_sankey', chart_images['sankey']):
                            charts_inserted += 1
                            logger.info(f"  Inserted sankey chart into slide {slide_idx + 1}")
                            break  # Move to next slide
                    
                    # Stacked bar chart - "MITRE ATT&CK Stacked Bar"
                    if chart_images.get('stacked_bar') and 'mitre_stacked_bar' in text:
                        if insert_chart_image(slide, 'mitre_stacked_bar', chart_images['stacked_bar']):
                            charts_inserted += 1
                            logger.info(f"  Inserted stacked bar chart into slide {slide_idx + 1}")
                            break  # Move to next slide
        
        logger.info(f"✓ Inserted {charts_inserted} chart images")
    except Exception as e:
        logger.error(f"Failed to insert chart images: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return 1
    
    # Step 7: Save presentation
    logger.info("Step 7: Saving presentation...")
    try:
        # Create output directory
        output_dir = Path(args.output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Generate filename with date
        date_str = datetime.now().strftime("%Y-%m-%d")
        output_filename = f"escalation_report_{date_str}.pptx"
        output_path = output_dir / output_filename
        
        prs.save(str(output_path))
        logger.info(f"✓ Presentation saved")
        logger.info(f"  Path: {output_path}")
        logger.info(f"  Total slides: {len(prs.slides)}")
    except Exception as e:
        logger.error(f"Failed to save presentation: {e}")
        return 1
    
    # Step 8: Cleanup (optional)
    if not args.keep_charts:
        logger.info("Step 8: Cleaning up temporary files...")
        try:
            temp_charts_path = Path(temp_charts_dir)
            if temp_charts_path.exists():
                shutil.rmtree(temp_charts_path)
                logger.info(f"✓ Deleted {temp_charts_dir}/")
        except Exception as e:
            logger.warning(f"Failed to cleanup temp files: {e}")
    else:
        logger.info(f"Step 8: Keeping temporary charts in {temp_charts_dir}/")
    
    logger.info("=" * 60)
    logger.info("✓ Presentation generation completed successfully!")
    logger.info("=" * 60)
    
    return 0


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


def build_protection_achieved_slides(prs, data):
    """Create the Protection section (Slides 7-10).
    
    Args:
        prs (Presentation): The presentation object.
        data (ReportData): The report data object containing all metrics.
    """
    blank_slide_layout = prs.slide_layouts[6]  # Blank layout
    header_height = Inches(0.8)
    
    # Slide 7 - Performance Trends
    slide7 = prs.slides.add_slide(blank_slide_layout)
    
    # Add logo at top right
    add_logo(slide7, position='top_right', prs=prs)
    
    # Add title header
    header_shape7 = slide7.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0,
        prs.slide_width, header_height
    )
    fill = header_shape7.fill
    fill.solid()
    fill.fore_color.rgb = CS_NAVY
    header_shape7.line.fill.background()
    
    # Add title text on header
    title_left = Inches(0.5)
    title_top = Inches(0.1)
    title_width = prs.slide_width - Inches(2.5)
    title_height = Inches(0.6)
    
    title_box7 = slide7.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_frame7 = title_box7.text_frame
    title_frame7.word_wrap = True
    title_paragraph7 = title_frame7.paragraphs[0]
    title_paragraph7.text = "Getting Better Every Period"
    title_paragraph7.font.name = TITLE_FONT_NAME
    title_paragraph7.font.size = Pt(28)
    title_paragraph7.font.bold = True
    title_paragraph7.font.color.rgb = RGBColor(255, 255, 255)
    title_paragraph7.alignment = PP_ALIGN.LEFT
    
    # Add chart placeholder (rectangle with border)
    chart_left = Inches(0.8)
    chart_top = header_height + Inches(0.4)
    chart_width = prs.slide_width - Inches(1.6)
    chart_height = Inches(2.8)
    
    chart_placeholder = slide7.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, chart_left, chart_top,
        chart_width, chart_height
    )
    fill = chart_placeholder.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)  # White background
    line = chart_placeholder.line
    line.color.rgb = CS_SLATE
    line.width = Pt(1)
    
    # Add placeholder text
    placeholder_text = chart_placeholder.text_frame
    placeholder_text.text = "[Chart: Performance Trends]"
    placeholder_text.paragraphs[0].font.name = BODY_FONT_NAME
    placeholder_text.paragraphs[0].font.size = Pt(14)
    placeholder_text.paragraphs[0].font.color.rgb = CS_SLATE
    placeholder_text.paragraphs[0].alignment = PP_ALIGN.CENTER
    placeholder_text.vertical_anchor = 1  # Middle
    
    # Add legend below chart
    legend_top = chart_top + chart_height + Inches(0.2)
    legend_left = chart_left
    legend_width = chart_width
    legend_height = Inches(0.5)
    
    legend_box = slide7.shapes.add_textbox(legend_left, legend_top, legend_width, legend_height)
    legend_frame = legend_box.text_frame
    legend_frame.word_wrap = True
    legend_paragraph = legend_frame.paragraphs[0]
    legend_paragraph.text = "MTTR (blue) | MTTD (navy) | FP% (red dashed)"
    legend_paragraph.font.name = BODY_FONT_NAME
    legend_paragraph.font.size = Pt(14)
    legend_paragraph.font.color.rgb = CS_SLATE
    legend_paragraph.alignment = PP_ALIGN.CENTER
    
    # Add insight box
    insight_left = Inches(0.8)
    insight_top = legend_top + legend_height + Inches(0.2)
    insight_width = chart_width
    insight_height = Inches(0.8)
    
    insight_shape = slide7.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, insight_left, insight_top,
        insight_width, insight_height
    )
    fill = insight_shape.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(240, 248, 255)  # Light blue background
    line = insight_shape.line
    line.color.rgb = CS_BLUE
    line.width = Pt(2)
    
    insight_text = insight_shape.text_frame
    insight_text.text = "MTTR decreased 25% to 126 minutes, MTTD improved 22% to 42 minutes"
    insight_text.paragraphs[0].font.name = BODY_FONT_NAME
    insight_text.paragraphs[0].font.size = Pt(16)
    insight_text.paragraphs[0].font.color.rgb = CS_NAVY
    insight_text.paragraphs[0].font.bold = True
    insight_text.paragraphs[0].alignment = PP_ALIGN.CENTER
    insight_text.vertical_anchor = 1  # Middle
    
    # Slide 8 - Industry Comparison
    slide8 = prs.slides.add_slide(blank_slide_layout)
    
    # Add logo at top right
    add_logo(slide8, position='top_right', prs=prs)
    
    # Add title header
    header_shape8 = slide8.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0,
        prs.slide_width, header_height
    )
    fill = header_shape8.fill
    fill.solid()
    fill.fore_color.rgb = CS_NAVY
    header_shape8.line.fill.background()
    
    # Add title text on header
    title_box8 = slide8.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_frame8 = title_box8.text_frame
    title_frame8.word_wrap = True
    title_paragraph8 = title_frame8.paragraphs[0]
    title_paragraph8.text = "Industry Comparison"
    title_paragraph8.font.name = TITLE_FONT_NAME
    title_paragraph8.font.size = Pt(28)
    title_paragraph8.font.bold = True
    title_paragraph8.font.color.rgb = RGBColor(255, 255, 255)
    title_paragraph8.alignment = PP_ALIGN.LEFT
    
    # Create table for comparison
    table_top = header_height + Inches(0.5)
    table_left = Inches(1)
    table_width = prs.slide_width - Inches(2)
    table_height = Inches(2.5)
    
    # Calculate row and column dimensions
    num_rows = len(data.industry_comparison) + 1  # +1 for header
    num_cols = 4
    row_height = table_height / num_rows
    col_width = table_width / num_cols
    
    # Create table using shapes (python-pptx doesn't have native table support)
    # Create header row background
    header_row_shape = slide8.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, table_left, table_top,
        table_width, row_height
    )
    fill = header_row_shape.fill
    fill.solid()
    fill.fore_color.rgb = CS_NAVY
    header_row_shape.line.fill.background()
    
    # Add header text
    headers = ["Metric", "Your Performance", "Industry Average", "Difference"]
    for col_idx, header_text in enumerate(headers):
        header_cell_left = table_left + col_idx * col_width
        header_cell_width = col_width
        header_cell_top = table_top
        
        header_cell_box = slide8.shapes.add_textbox(
            header_cell_left + Inches(0.1), header_cell_top + Inches(0.05),
            header_cell_width - Inches(0.2), row_height - Inches(0.1)
        )
        header_cell_frame = header_cell_box.text_frame
        header_cell_frame.word_wrap = True
        header_cell_paragraph = header_cell_frame.paragraphs[0]
        header_cell_paragraph.text = header_text
        header_cell_paragraph.font.name = BODY_FONT_NAME
        header_cell_paragraph.font.size = Pt(14)
        header_cell_paragraph.font.bold = True
        header_cell_paragraph.font.color.rgb = RGBColor(255, 255, 255)
        header_cell_paragraph.alignment = PP_ALIGN.LEFT
    
    # Add data rows
    for row_idx, comparison in enumerate(data.industry_comparison):
        row_top = table_top + (row_idx + 1) * row_height
        
        # Alternate row background color
        if row_idx % 2 == 0:
            row_bg = slide8.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, table_left, row_top,
                table_width, row_height
            )
            fill = row_bg.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(250, 250, 250)  # Light gray
            row_bg.line.fill.background()
        
        # Metric name
        metric_cell_box = slide8.shapes.add_textbox(
            table_left + Inches(0.1), row_top + Inches(0.05),
            col_width - Inches(0.2), row_height - Inches(0.1)
        )
        metric_cell_frame = metric_cell_box.text_frame
        metric_cell_frame.word_wrap = True
        metric_cell_paragraph = metric_cell_frame.paragraphs[0]
        metric_cell_paragraph.text = f"{comparison['metric']} (minutes)" if comparison['metric'] != "Incidents/Day" else comparison['metric']
        metric_cell_paragraph.font.name = BODY_FONT_NAME
        metric_cell_paragraph.font.size = Pt(13)
        metric_cell_paragraph.font.color.rgb = CS_SLATE
        metric_cell_paragraph.alignment = PP_ALIGN.LEFT
        
        # Your Performance
        your_cell_box = slide8.shapes.add_textbox(
            table_left + col_width + Inches(0.1), row_top + Inches(0.05),
            col_width - Inches(0.2), row_height - Inches(0.1)
        )
        your_cell_frame = your_cell_box.text_frame
        your_cell_frame.word_wrap = True
        your_cell_paragraph = your_cell_frame.paragraphs[0]
        your_cell_paragraph.text = str(comparison['yours'])
        your_cell_paragraph.font.name = BODY_FONT_NAME
        your_cell_paragraph.font.size = Pt(13)
        your_cell_paragraph.font.color.rgb = CS_SLATE
        your_cell_paragraph.alignment = PP_ALIGN.LEFT
        
        # Industry Average
        industry_cell_box = slide8.shapes.add_textbox(
            table_left + 2 * col_width + Inches(0.1), row_top + Inches(0.05),
            col_width - Inches(0.2), row_height - Inches(0.1)
        )
        industry_cell_frame = industry_cell_box.text_frame
        industry_cell_frame.word_wrap = True
        industry_cell_paragraph = industry_cell_frame.paragraphs[0]
        industry_cell_paragraph.text = str(comparison['industry'])
        industry_cell_paragraph.font.name = BODY_FONT_NAME
        industry_cell_paragraph.font.size = Pt(13)
        industry_cell_paragraph.font.color.rgb = CS_SLATE
        industry_cell_paragraph.alignment = PP_ALIGN.LEFT
        
        # Difference (with blue badge)
        diff_cell_left = table_left + 3 * col_width + Inches(0.1)
        diff_cell_top = row_top + Inches(0.05)
        diff_cell_width = col_width - Inches(0.2)
        diff_cell_height = row_height - Inches(0.1)
        
        # Create blue badge background
        badge_shape = slide8.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, diff_cell_left, diff_cell_top,
            diff_cell_width, diff_cell_height
        )
        fill = badge_shape.fill
        fill.solid()
        fill.fore_color.rgb = CS_BLUE
        badge_shape.line.fill.background()
        
        diff_cell_box = slide8.shapes.add_textbox(
            diff_cell_left, diff_cell_top,
            diff_cell_width, diff_cell_height
        )
        diff_cell_frame = diff_cell_box.text_frame
        diff_cell_frame.word_wrap = True
        diff_cell_paragraph = diff_cell_frame.paragraphs[0]
        diff_cell_paragraph.text = comparison['difference']
        diff_cell_paragraph.font.name = BODY_FONT_NAME
        diff_cell_paragraph.font.size = Pt(13)
        diff_cell_paragraph.font.bold = True
        diff_cell_paragraph.font.color.rgb = RGBColor(255, 255, 255)
        diff_cell_paragraph.alignment = PP_ALIGN.CENTER
        diff_cell_frame.vertical_anchor = 1  # Middle
    
    # Slide 9 - Response Efficiency
    slide9 = prs.slides.add_slide(blank_slide_layout)
    
    # Add logo at top right
    add_logo(slide9, position='top_right', prs=prs)
    
    # Add title header
    header_shape9 = slide9.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0,
        prs.slide_width, header_height
    )
    fill = header_shape9.fill
    fill.solid()
    fill.fore_color.rgb = CS_NAVY
    header_shape9.line.fill.background()
    
    # Add title text on header
    title_box9 = slide9.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_frame9 = title_box9.text_frame
    title_frame9.word_wrap = True
    title_paragraph9 = title_frame9.paragraphs[0]
    title_paragraph9.text = "Response Efficiency"
    title_paragraph9.font.name = TITLE_FONT_NAME
    title_paragraph9.font.size = Pt(28)
    title_paragraph9.font.bold = True
    title_paragraph9.font.color.rgb = RGBColor(255, 255, 255)
    title_paragraph9.alignment = PP_ALIGN.LEFT
    
    # Create 3 metric cards
    card_width = (prs.slide_width - Inches(2.2)) / 3
    card_height = Inches(3)
    card_spacing = Inches(0.2)
    cards_start_left = Inches(0.5)
    cards_start_top = header_height + Inches(0.5)
    
    # Calculate containment count (round to nearest integer)
    containment_count = round((data.containment_rate / 100) * data.incidents_escalated)
    
    efficiency_cards = [
        {
            "title": "Containment Rate",
            "value": f"{data.containment_rate}%",
            "subtitle": f"({containment_count} of {data.incidents_escalated})",
            "accent_color": CS_BLUE
        },
        {
            "title": "Playbook Automation",
            "value": f"{data.playbook_auto['percent']}%",
            "subtitle": f"({data.playbook_auto['count']} incidents)",
            "accent_color": CS_BLUE
        },
        {
            "title": "Analyst Escalation",
            "value": f"{data.analyst_escalation['percent']}%",
            "subtitle": f"({data.analyst_escalation['count']} incidents)",
            "accent_color": CS_ORANGE
        }
    ]
    
    for i, card in enumerate(efficiency_cards):
        card_left = cards_start_left + i * (card_width + card_spacing)
        card_top = cards_start_top
        
        # Create card background
        card_shape = slide9.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, card_left, card_top,
            card_width, card_height
        )
        fill = card_shape.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(240, 248, 255)  # Light blue background
        line = card_shape.line
        line.color.rgb = card["accent_color"]
        line.width = Pt(3)
        
        # Add card title
        card_title_box = slide9.shapes.add_textbox(
            card_left + Inches(0.2), card_top + Inches(0.3),
            card_width - Inches(0.4), Inches(0.5)
        )
        card_title_frame = card_title_box.text_frame
        card_title_frame.word_wrap = True
        card_title_paragraph = card_title_frame.paragraphs[0]
        card_title_paragraph.text = card["title"]
        card_title_paragraph.font.name = TITLE_FONT_NAME
        card_title_paragraph.font.size = Pt(18)
        card_title_paragraph.font.bold = True
        card_title_paragraph.font.color.rgb = CS_NAVY
        card_title_paragraph.alignment = PP_ALIGN.CENTER
        
        # Add card value
        card_value_box = slide9.shapes.add_textbox(
            card_left + Inches(0.2), card_top + Inches(1.0),
            card_width - Inches(0.4), Inches(1.0)
        )
        card_value_frame = card_value_box.text_frame
        card_value_frame.word_wrap = True
        card_value_paragraph = card_value_frame.paragraphs[0]
        card_value_paragraph.text = card["value"]
        card_value_paragraph.font.name = TITLE_FONT_NAME
        card_value_paragraph.font.size = Pt(48)
        card_value_paragraph.font.bold = True
        card_value_paragraph.font.color.rgb = card["accent_color"]
        card_value_paragraph.alignment = PP_ALIGN.CENTER
        
        # Add card subtitle
        card_subtitle_box = slide9.shapes.add_textbox(
            card_left + Inches(0.2), card_top + Inches(2.2),
            card_width - Inches(0.4), Inches(0.6)
        )
        card_subtitle_frame = card_subtitle_box.text_frame
        card_subtitle_frame.word_wrap = True
        card_subtitle_paragraph = card_subtitle_frame.paragraphs[0]
        card_subtitle_paragraph.text = card["subtitle"]
        card_subtitle_paragraph.font.name = BODY_FONT_NAME
        card_subtitle_paragraph.font.size = Pt(14)
        card_subtitle_paragraph.font.color.rgb = CS_SLATE
        card_subtitle_paragraph.alignment = PP_ALIGN.CENTER
    
    # Slide 10 - Detection Quality
    slide10 = prs.slides.add_slide(blank_slide_layout)
    
    # Add logo at top right
    add_logo(slide10, position='top_right', prs=prs)
    
    # Add title header
    header_shape10 = slide10.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0,
        prs.slide_width, header_height
    )
    fill = header_shape10.fill
    fill.solid()
    fill.fore_color.rgb = CS_NAVY
    header_shape10.line.fill.background()
    
    # Add title text on header
    title_box10 = slide10.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_frame10 = title_box10.text_frame
    title_frame10.word_wrap = True
    title_paragraph10 = title_frame10.paragraphs[0]
    title_paragraph10.text = "Detection Quality"
    title_paragraph10.font.name = TITLE_FONT_NAME
    title_paragraph10.font.size = Pt(28)
    title_paragraph10.font.bold = True
    title_paragraph10.font.color.rgb = RGBColor(255, 255, 255)
    title_paragraph10.alignment = PP_ALIGN.LEFT
    
    # Create 2x2 grid of metric cards
    grid_card_width = (prs.slide_width - Inches(2.2)) / 2
    grid_card_height = Inches(1.5)
    grid_spacing = Inches(0.2)
    grid_start_left = Inches(0.5)
    grid_start_top = header_height + Inches(0.5)
    
    quality_cards = [
        {
            "title": "True Threat Precision",
            "value": f"{data.true_threat_precision}%",
            "accent_color": CS_RED
        },
        {
            "title": "Signal Fidelity",
            "value": f"{data.signal_fidelity}%",
            "accent_color": CS_BLUE
        },
        {
            "title": "False Positive Rate",
            "value": f"{data.false_positive_rate}%",
            "accent_color": CS_ORANGE
        },
        {
            "title": "Client-Validated",
            "value": f"{data.client_validated}%",
            "accent_color": CS_BLUE
        }
    ]
    
    for i, card in enumerate(quality_cards):
        row = i // 2
        col = i % 2
        
        card_left = grid_start_left + col * (grid_card_width + grid_spacing)
        card_top = grid_start_top + row * (grid_card_height + grid_spacing)
        
        # Create card background
        card_shape = slide10.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, card_left, card_top,
            grid_card_width, grid_card_height
        )
        fill = card_shape.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(240, 248, 255)  # Light blue background
        line = card_shape.line
        line.color.rgb = card["accent_color"]
        line.width = Pt(3)
        
        # Add card title
        card_title_box = slide10.shapes.add_textbox(
            card_left + Inches(0.15), card_top + Inches(0.2),
            grid_card_width - Inches(0.3), Inches(0.4)
        )
        card_title_frame = card_title_box.text_frame
        card_title_frame.word_wrap = True
        card_title_paragraph = card_title_frame.paragraphs[0]
        card_title_paragraph.text = card["title"]
        card_title_paragraph.font.name = TITLE_FONT_NAME
        card_title_paragraph.font.size = Pt(16)
        card_title_paragraph.font.bold = True
        card_title_paragraph.font.color.rgb = CS_NAVY
        card_title_paragraph.alignment = PP_ALIGN.LEFT
        
        # Add card value
        card_value_box = slide10.shapes.add_textbox(
            card_left + Inches(0.15), card_top + Inches(0.7),
            grid_card_width - Inches(0.3), Inches(0.6)
        )
        card_value_frame = card_value_box.text_frame
        card_value_frame.word_wrap = True
        card_value_paragraph = card_value_frame.paragraphs[0]
        card_value_paragraph.text = card["value"]
        card_value_paragraph.font.name = TITLE_FONT_NAME
        card_value_paragraph.font.size = Pt(36)
        card_value_paragraph.font.bold = True
        card_value_paragraph.font.color.rgb = card["accent_color"]
        card_value_paragraph.alignment = PP_ALIGN.LEFT
    
    # Add key insight box at bottom
    insight_box_top = grid_start_top + 2 * (grid_card_height + grid_spacing) + Inches(0.3)
    insight_box_left = grid_start_left
    insight_box_width = prs.slide_width - Inches(1)
    insight_box_height = Inches(0.7)
    
    insight_box_shape = slide10.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, insight_box_left, insight_box_top,
        insight_box_width, insight_box_height
    )
    fill = insight_box_shape.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(240, 248, 255)  # Light blue background
    line = insight_box_shape.line
    line.color.rgb = CS_BLUE
    line.width = Pt(2)
    
    insight_box_text = insight_box_shape.text_frame
    insight_box_text.text = "Signal quality improved: false positives fell from 10.8% to 9.0%"
    insight_box_text.paragraphs[0].font.name = BODY_FONT_NAME
    insight_box_text.paragraphs[0].font.size = Pt(16)
    insight_box_text.paragraphs[0].font.color.rgb = CS_NAVY
    insight_box_text.paragraphs[0].font.bold = True
    insight_box_text.paragraphs[0].alignment = PP_ALIGN.CENTER
    insight_box_text.vertical_anchor = 1  # Middle


def build_threat_landscape_slides(prs, data, include=True):
    """Create the optional Threat Landscape section (Slides 11-13).
    
    Args:
        prs (Presentation): The presentation object.
        data (ReportData): The report data object containing all metrics.
        include (bool): Whether to include this section. Default True.
    """
    if not include:
        return
    
    blank_slide_layout = prs.slide_layouts[6]  # Blank layout
    header_height = Inches(0.8)
    title_left = Inches(0.5)
    title_top = Inches(0.1)
    title_width = prs.slide_width - Inches(2.5)
    title_height = Inches(0.6)
    
    # Calculate severity alignment stats from severity_flows
    severity_order = ['Informational', 'Low', 'Medium', 'High', 'Critical']
    
    def get_severity_index(label):
        """Extract severity level from label and return its index."""
        parts = label.split(' ')
        severity_label = parts[-1]  # Get last part (e.g., 'Critical' from 'Vendor Critical')
        return severity_order.index(severity_label) if severity_label in severity_order else -1
    
    upgraded_count = 0
    de_escalated_count = 0
    aligned_count = 0
    
    for flow in data.severity_flows:
        source_idx = get_severity_index(flow['from'])
        target_idx = get_severity_index(flow['to'])
        flow_count = flow['flow']
        
        if target_idx > source_idx:
            upgraded_count += flow_count
        elif target_idx < source_idx:
            de_escalated_count += flow_count
        else:
            aligned_count += flow_count
    
    total_escalations = upgraded_count + de_escalated_count + aligned_count
    
    # Slide 11 - Severity Alignment Flow
    slide11 = prs.slides.add_slide(blank_slide_layout)
    
    # Add logo at top right
    add_logo(slide11, position='top_right', prs=prs)
    
    # Add title header
    header_shape11 = slide11.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0,
        prs.slide_width, header_height
    )
    fill = header_shape11.fill
    fill.solid()
    fill.fore_color.rgb = CS_NAVY
    header_shape11.line.fill.background()
    
    # Add title text on header
    title_box11 = slide11.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_frame11 = title_box11.text_frame
    title_frame11.word_wrap = True
    title_paragraph11 = title_frame11.paragraphs[0]
    title_paragraph11.text = "Severity Alignment Flow"
    title_paragraph11.font.name = TITLE_FONT_NAME
    title_paragraph11.font.size = Pt(28)
    title_paragraph11.font.bold = True
    title_paragraph11.font.color.rgb = RGBColor(255, 255, 255)
    title_paragraph11.alignment = PP_ALIGN.LEFT
    
    # Add subtitle
    subtitle_left = Inches(0.5)
    subtitle_top = header_height + Inches(0.1)
    subtitle_width = prs.slide_width - Inches(2.5)
    subtitle_height = Inches(0.4)
    
    subtitle_box = slide11.shapes.add_textbox(subtitle_left, subtitle_top, subtitle_width, subtitle_height)
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.word_wrap = True
    subtitle_paragraph = subtitle_frame.paragraphs[0]
    subtitle_paragraph.text = "Vendor-reported criticality vs. Critical Start adjudication"
    subtitle_paragraph.font.name = BODY_FONT_NAME
    subtitle_paragraph.font.size = Pt(14)
    subtitle_paragraph.font.color.rgb = CS_SLATE
    subtitle_paragraph.alignment = PP_ALIGN.LEFT
    
    # Add chip with total escalations
    chip_left = prs.slide_width - Inches(2.5)
    chip_top = header_height + Inches(0.1)
    chip_width = Inches(1.8)
    chip_height = Inches(0.4)
    
    chip_shape = slide11.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, chip_left, chip_top,
        chip_width, chip_height
    )
    fill = chip_shape.fill
    fill.solid()
    fill.fore_color.rgb = CS_BLUE
    chip_shape.line.fill.background()
    
    chip_text = chip_shape.text_frame
    chip_text.text = f"{total_escalations} escalations"
    chip_text.paragraphs[0].font.name = BODY_FONT_NAME
    chip_text.paragraphs[0].font.size = Pt(12)
    chip_text.paragraphs[0].font.bold = True
    chip_text.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    chip_text.paragraphs[0].alignment = PP_ALIGN.CENTER
    chip_text.vertical_anchor = 1  # Middle
    
    # Add 3 stat boxes
    stat_box_width = (prs.slide_width - Inches(2.2)) / 3
    stat_box_height = Inches(1.2)
    stat_spacing = Inches(0.2)
    stat_start_left = Inches(0.5)
    stat_start_top = subtitle_top + subtitle_height + Inches(0.3)
    
    stat_boxes = [
        {"label": "Upgraded", "value": str(upgraded_count), "color": CS_RED},
        {"label": "De-escalated", "value": str(de_escalated_count), "color": CS_BLUE},
        {"label": "Aligned", "value": str(aligned_count), "color": CS_NAVY}
    ]
    
    for i, stat in enumerate(stat_boxes):
        stat_left = stat_start_left + i * (stat_box_width + stat_spacing)
        stat_top = stat_start_top
        
        # Create stat box background
        stat_shape = slide11.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, stat_left, stat_top,
            stat_box_width, stat_box_height
        )
        fill = stat_shape.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(240, 248, 255)  # Light blue background
        line = stat_shape.line
        line.color.rgb = stat["color"]
        line.width = Pt(3)
        
        # Add stat value
        stat_value_box = slide11.shapes.add_textbox(
            stat_left + Inches(0.1), stat_top + Inches(0.2),
            stat_box_width - Inches(0.2), Inches(0.5)
        )
        stat_value_frame = stat_value_box.text_frame
        stat_value_frame.word_wrap = True
        stat_value_paragraph = stat_value_frame.paragraphs[0]
        stat_value_paragraph.text = stat["value"]
        stat_value_paragraph.font.name = TITLE_FONT_NAME
        stat_value_paragraph.font.size = Pt(36)
        stat_value_paragraph.font.bold = True
        stat_value_paragraph.font.color.rgb = stat["color"]
        stat_value_paragraph.alignment = PP_ALIGN.CENTER
        
        # Add stat label
        stat_label_box = slide11.shapes.add_textbox(
            stat_left + Inches(0.1), stat_top + Inches(0.7),
            stat_box_width - Inches(0.2), Inches(0.4)
        )
        stat_label_frame = stat_label_box.text_frame
        stat_label_frame.word_wrap = True
        stat_label_paragraph = stat_label_frame.paragraphs[0]
        stat_label_paragraph.text = stat["label"]
        stat_label_paragraph.font.name = BODY_FONT_NAME
        stat_label_paragraph.font.size = Pt(14)
        stat_label_paragraph.font.color.rgb = CS_SLATE
        stat_label_paragraph.alignment = PP_ALIGN.CENTER
    
    # Add Sankey chart placeholder
    chart_left = Inches(0.8)
    chart_top = stat_start_top + stat_box_height + Inches(0.3)
    chart_width = prs.slide_width - Inches(1.6)
    chart_height = Inches(2.0)
    
    chart_placeholder = slide11.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, chart_left, chart_top,
        chart_width, chart_height
    )
    fill = chart_placeholder.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)  # White background
    line = chart_placeholder.line
    line.color.rgb = CS_SLATE
    line.width = Pt(1)
    
    # Add placeholder text with ID for later replacement
    placeholder_text = chart_placeholder.text_frame
    placeholder_text.text = "[Chart: Severity Alignment Sankey - ID: severity_sankey]"
    placeholder_text.paragraphs[0].font.name = BODY_FONT_NAME
    placeholder_text.paragraphs[0].font.size = Pt(14)
    placeholder_text.paragraphs[0].font.color.rgb = CS_SLATE
    placeholder_text.paragraphs[0].alignment = PP_ALIGN.CENTER
    placeholder_text.vertical_anchor = 1  # Middle
    
    # Add key insight bullets
    insight_left = Inches(0.8)
    insight_top = chart_top + chart_height + Inches(0.2)
    insight_width = chart_width
    insight_height = Inches(0.8)
    
    insight_box = slide11.shapes.add_textbox(insight_left, insight_top, insight_width, insight_height)
    insight_frame = insight_box.text_frame
    insight_frame.word_wrap = True
    insight_frame.margin_left = Inches(0.2)
    insight_frame.margin_right = Inches(0.2)
    
    # Calculate percentages
    upgrade_pct = (upgraded_count / total_escalations * 100) if total_escalations > 0 else 0
    downgrade_pct = (de_escalated_count / total_escalations * 100) if total_escalations > 0 else 0
    
    insight_text = f"• {upgraded_count} escalations upgraded ({upgrade_pct:.1f}%) - CS raised severity based on business impact\n"
    insight_text += f"• {de_escalated_count} escalations de-escalated ({downgrade_pct:.1f}%) - CS reduced noise from vendor over-classification"
    
    insight_paragraph = insight_frame.paragraphs[0]
    insight_paragraph.text = insight_text
    insight_paragraph.font.name = BODY_FONT_NAME
    insight_paragraph.font.size = Pt(13)
    insight_paragraph.font.color.rgb = CS_SLATE
    insight_paragraph.alignment = PP_ALIGN.LEFT
    
    # Add footnote
    footnote_left = Inches(0.5)
    footnote_top = prs.slide_height - Inches(0.5)
    footnote_width = prs.slide_width - Inches(1)
    footnote_height = Inches(0.3)
    
    footnote_box = slide11.shapes.add_textbox(footnote_left, footnote_top, footnote_width, footnote_height)
    footnote_frame = footnote_box.text_frame
    footnote_frame.word_wrap = True
    footnote_paragraph = footnote_frame.paragraphs[0]
    footnote_paragraph.text = "Source: Vendor Severity (col 45) vs. Current Priority (col 13)"
    footnote_paragraph.font.name = BODY_FONT_NAME
    footnote_paragraph.font.size = Pt(10)
    footnote_paragraph.font.color.rgb = CS_SLATE
    footnote_paragraph.alignment = PP_ALIGN.LEFT
    footnote_paragraph.font.italic = True
    
    # Slide 12 - MITRE ATT&CK Landscape
    slide12 = prs.slides.add_slide(blank_slide_layout)
    
    # Add logo at top right
    add_logo(slide12, position='top_right', prs=prs)
    
    # Add title header
    header_shape12 = slide12.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0,
        prs.slide_width, header_height
    )
    fill = header_shape12.fill
    fill.solid()
    fill.fore_color.rgb = CS_NAVY
    header_shape12.line.fill.background()
    
    # Add title text on header
    title_box12 = slide12.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_frame12 = title_box12.text_frame
    title_frame12.word_wrap = True
    title_paragraph12 = title_frame12.paragraphs[0]
    title_paragraph12.text = "Threat Landscape by Tactic & Severity"
    title_paragraph12.font.name = TITLE_FONT_NAME
    title_paragraph12.font.size = Pt(28)
    title_paragraph12.font.bold = True
    title_paragraph12.font.color.rgb = RGBColor(255, 255, 255)
    title_paragraph12.alignment = PP_ALIGN.LEFT
    
    # Add stacked bar chart placeholder
    chart_left12 = Inches(0.8)
    chart_top12 = header_height + Inches(0.4)
    chart_width12 = prs.slide_width - Inches(1.6)
    chart_height12 = Inches(3.0)
    
    chart_placeholder12 = slide12.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, chart_left12, chart_top12,
        chart_width12, chart_height12
    )
    fill = chart_placeholder12.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)  # White background
    line = chart_placeholder12.line
    line.color.rgb = CS_SLATE
    line.width = Pt(1)
    
    # Add placeholder text with ID
    placeholder_text12 = chart_placeholder12.text_frame
    placeholder_text12.text = "[Chart: MITRE ATT&CK Stacked Bar - ID: mitre_stacked_bar]"
    placeholder_text12.paragraphs[0].font.name = BODY_FONT_NAME
    placeholder_text12.paragraphs[0].font.size = Pt(14)
    placeholder_text12.paragraphs[0].font.color.rgb = CS_SLATE
    placeholder_text12.paragraphs[0].alignment = PP_ALIGN.CENTER
    placeholder_text12.vertical_anchor = 1  # Middle
    
    # Add legend
    legend_top = chart_top12 + chart_height12 + Inches(0.2)
    legend_left = chart_left12
    legend_width = chart_width12
    legend_height = Inches(0.4)
    
    legend_box = slide12.shapes.add_textbox(legend_left, legend_top, legend_width, legend_height)
    legend_frame = legend_box.text_frame
    legend_frame.word_wrap = True
    legend_paragraph = legend_frame.paragraphs[0]
    legend_paragraph.text = "High (red) | Medium (orange) | Low (blue) | Info (gray)"
    legend_paragraph.font.name = BODY_FONT_NAME
    legend_paragraph.font.size = Pt(12)
    legend_paragraph.font.color.rgb = CS_SLATE
    legend_paragraph.alignment = PP_ALIGN.CENTER
    
    # Add key insight
    insight_left12 = Inches(0.8)
    insight_top12 = legend_top + legend_height + Inches(0.2)
    insight_width12 = chart_width12
    insight_height12 = Inches(0.6)
    
    # Calculate Persistence stats (first tactic)
    if len(data.tactics) > 0 and len(data.high_severity) > 0:
        persistence_total = (data.high_severity[0] + data.medium_severity[0] + 
                           data.low_severity[0] + data.info_severity[0])
        persistence_high = data.high_severity[0]
        persistence_high_pct = (persistence_high / persistence_total * 100) if persistence_total > 0 else 0
        insight_text12 = f"Persistence tactics generated {persistence_total} escalations with {persistence_high} high-severity cases ({persistence_high_pct:.1f}%)"
    else:
        insight_text12 = "Persistence tactics generated 77 escalations with 12 high-severity cases (15.6%)"
    
    insight_shape12 = slide12.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, insight_left12, insight_top12,
        insight_width12, insight_height12
    )
    fill = insight_shape12.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(240, 248, 255)  # Light blue background
    line = insight_shape12.line
    line.color.rgb = CS_BLUE
    line.width = Pt(2)
    
    insight_text_frame12 = insight_shape12.text_frame
    insight_text_frame12.text = insight_text12
    insight_text_frame12.paragraphs[0].font.name = BODY_FONT_NAME
    insight_text_frame12.paragraphs[0].font.size = Pt(14)
    insight_text_frame12.paragraphs[0].font.color.rgb = CS_NAVY
    insight_text_frame12.paragraphs[0].font.bold = True
    insight_text_frame12.paragraphs[0].alignment = PP_ALIGN.CENTER
    insight_text_frame12.vertical_anchor = 1  # Middle
    
    # Add footnote
    footnote_left12 = Inches(0.5)
    footnote_top12 = prs.slide_height - Inches(0.5)
    footnote_width12 = prs.slide_width - Inches(1)
    footnote_height12 = Inches(0.3)
    
    footnote_box12 = slide12.shapes.add_textbox(footnote_left12, footnote_top12, footnote_width12, footnote_height12)
    footnote_frame12 = footnote_box12.text_frame
    footnote_frame12.word_wrap = True
    footnote_paragraph12 = footnote_frame12.paragraphs[0]
    footnote_paragraph12.text = "Source: MITRE ATT&CK Tactic (col X) and Current Priority (col 13)"
    footnote_paragraph12.font.name = BODY_FONT_NAME
    footnote_paragraph12.font.size = Pt(10)
    footnote_paragraph12.font.color.rgb = CS_SLATE
    footnote_paragraph12.alignment = PP_ALIGN.LEFT
    footnote_paragraph12.font.italic = True
    
    # Slide 13 - Detection Sources
    slide13 = prs.slides.add_slide(blank_slide_layout)
    
    # Add logo at top right
    add_logo(slide13, position='top_right', prs=prs)
    
    # Add title header
    header_shape13 = slide13.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0,
        prs.slide_width, header_height
    )
    fill = header_shape13.fill
    fill.solid()
    fill.fore_color.rgb = CS_NAVY
    header_shape13.line.fill.background()
    
    # Add title text on header
    title_box13 = slide13.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_frame13 = title_box13.text_frame
    title_frame13.word_wrap = True
    title_paragraph13 = title_frame13.paragraphs[0]
    title_paragraph13.text = "Detection Sources & Quality"
    title_paragraph13.font.name = TITLE_FONT_NAME
    title_paragraph13.font.size = Pt(28)
    title_paragraph13.font.bold = True
    title_paragraph13.font.color.rgb = RGBColor(255, 255, 255)
    title_paragraph13.alignment = PP_ALIGN.LEFT
    
    # Create 3 source cards
    source_card_width = (prs.slide_width - Inches(2.2)) / 3
    source_card_height = Inches(3.5)
    source_card_spacing = Inches(0.2)
    source_start_left = Inches(0.5)
    source_start_top = header_height + Inches(0.4)
    
    fp_threshold = 10.0  # 10% threshold
    
    for i, source in enumerate(data.detection_sources):
        card_left = source_start_left + i * (source_card_width + source_card_spacing)
        card_top = source_start_top
        
        # Determine FP rate color (orange if above threshold, blue if below)
        fp_rate = source.get('fp_rate', 0)
        fp_color = CS_ORANGE if fp_rate > fp_threshold else CS_BLUE
        
        # Create card background
        card_shape = slide13.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, card_left, card_top,
            source_card_width, source_card_height
        )
        fill = card_shape.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(240, 248, 255)  # Light blue background
        line = card_shape.line
        line.color.rgb = fp_color
        line.width = Pt(3)
        
        # Add source name
        source_name_left = card_left + Inches(0.15)
        source_name_top = card_top + Inches(0.3)
        source_name_width = source_card_width - Inches(0.3)
        source_name_height = Inches(0.6)
        
        source_name_box = slide13.shapes.add_textbox(
            source_name_left, source_name_top, source_name_width, source_name_height
        )
        source_name_frame = source_name_box.text_frame
        source_name_frame.word_wrap = True
        source_name_paragraph = source_name_frame.paragraphs[0]
        source_name_paragraph.text = source['source']
        source_name_paragraph.font.name = TITLE_FONT_NAME
        source_name_paragraph.font.size = Pt(16)
        source_name_paragraph.font.bold = True
        source_name_paragraph.font.color.rgb = CS_NAVY
        source_name_paragraph.alignment = PP_ALIGN.LEFT
        
        # Add incidents count
        incidents_left = source_name_left
        incidents_top = source_name_top + source_name_height + Inches(0.2)
        incidents_width = source_name_width
        incidents_height = Inches(0.5)
        
        incidents_box = slide13.shapes.add_textbox(
            incidents_left, incidents_top, incidents_width, incidents_height
        )
        incidents_frame = incidents_box.text_frame
        incidents_frame.word_wrap = True
        incidents_paragraph = incidents_frame.paragraphs[0]
        incidents_paragraph.text = f"{source['incidents']} incidents ({source['percent']}%)"
        incidents_paragraph.font.name = BODY_FONT_NAME
        incidents_paragraph.font.size = Pt(14)
        incidents_paragraph.font.color.rgb = CS_SLATE
        incidents_paragraph.alignment = PP_ALIGN.LEFT
        
        # Add FP Rate with visual indicator
        fp_left = incidents_left
        fp_top = incidents_top + incidents_height + Inches(0.3)
        fp_width = source_name_width
        fp_height = Inches(1.0)
        
        # Create FP rate badge
        fp_badge_shape = slide13.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, fp_left, fp_top,
            fp_width, fp_height
        )
        fill = fp_badge_shape.fill
        fill.solid()
        fill.fore_color.rgb = fp_color
        fp_badge_shape.line.fill.background()
        
        fp_badge_text = fp_badge_shape.text_frame
        fp_badge_text.text = f"FP Rate: {fp_rate}%"
        fp_badge_text.paragraphs[0].font.name = TITLE_FONT_NAME
        fp_badge_text.paragraphs[0].font.size = Pt(20)
        fp_badge_text.paragraphs[0].font.bold = True
        fp_badge_text.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
        fp_badge_text.paragraphs[0].alignment = PP_ALIGN.CENTER
        fp_badge_text.vertical_anchor = 1  # Middle
        
        # Add status indicator text below badge
        status_left = fp_left
        status_top = fp_top + fp_height + Inches(0.1)
        status_width = fp_width
        status_height = Inches(0.4)
        
        status_box = slide13.shapes.add_textbox(
            status_left, status_top, status_width, status_height
        )
        status_frame = status_box.text_frame
        status_frame.word_wrap = True
        status_paragraph = status_frame.paragraphs[0]
        if fp_rate > fp_threshold:
            status_paragraph.text = "Above threshold"
        else:
            status_paragraph.text = "Good"
        status_paragraph.font.name = BODY_FONT_NAME
        status_paragraph.font.size = Pt(12)
        status_paragraph.font.color.rgb = CS_SLATE
        status_paragraph.alignment = PP_ALIGN.CENTER


def build_insights_slides(prs, data):
    """Create the Insights section (Slides 14-15).
    
    Args:
        prs (Presentation): The presentation object.
        data (ReportData): The report data object containing all metrics.
    """
    blank_slide_layout = prs.slide_layouts[6]  # Blank layout
    header_height = Inches(0.8)
    title_left = Inches(0.5)
    title_top = Inches(0.1)
    title_width = prs.slide_width - Inches(2.5)
    title_height = Inches(0.6)
    
    # Slide 14 - Prioritized Improvement Plan
    slide14 = prs.slides.add_slide(blank_slide_layout)
    
    # Add logo at top right
    add_logo(slide14, position='top_right', prs=prs)
    
    # Add title header
    header_shape14 = slide14.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0,
        prs.slide_width, header_height
    )
    fill = header_shape14.fill
    fill.solid()
    fill.fore_color.rgb = CS_NAVY
    header_shape14.line.fill.background()
    
    # Add title text on header
    title_box14 = slide14.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_frame14 = title_box14.text_frame
    title_frame14.word_wrap = True
    title_paragraph14 = title_frame14.paragraphs[0]
    title_paragraph14.text = "Prioritized Improvement Plan"
    title_paragraph14.font.name = TITLE_FONT_NAME
    title_paragraph14.font.size = Pt(28)
    title_paragraph14.font.bold = True
    title_paragraph14.font.color.rgb = RGBColor(255, 255, 255)
    title_paragraph14.alignment = PP_ALIGN.LEFT
    
    # Create improvement items cards
    card_start_top = header_height + Inches(0.4)
    card_width = prs.slide_width - Inches(1.0)
    card_height = Inches(1.3)
    card_spacing = Inches(0.2)
    card_left = Inches(0.5)
    
    # Map priority to colors
    priority_colors = {
        "HIGH": CS_RED,
        "MEDIUM": CS_ORANGE,
        "LOW": CS_BLUE
    }
    
    # Expected impact text for each item
    expected_impacts = [
        "Reduce escalations by ~20%",
        "Reduce manual review burden",
        "Improved detection of advanced threats"
    ]
    
    # Concise descriptions as per requirements
    concise_descriptions = [
        "Palo Alto Cortex XDR false positive rate is 11.2%, exceeding the 10.0% threshold",
        "Manual escalations at 14% exceed 12% target. 38 incidents required analyst judgment",
        "Persistence + Defense Evasion account for 67% of high-severity incidents"
    ]
    
    for i, item in enumerate(data.improvement_items):
        card_top = card_start_top + i * (card_height + card_spacing)
        
        # Determine priority color
        priority = item.get("priority", "MEDIUM")
        border_color = priority_colors.get(priority, CS_ORANGE)
        
        # Create card background with colored left border
        card_shape = slide14.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, card_left, card_top,
            card_width, card_height
        )
        fill = card_shape.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(240, 248, 255)  # Light blue background
        line = card_shape.line
        line.color.rgb = border_color
        line.width = Pt(4)  # Thicker left border
        
        # Add priority badge at top left
        badge_width = Inches(1.0)
        badge_height = Inches(0.35)
        badge_left = card_left + Inches(0.2)
        badge_top = card_top + Inches(0.15)
        
        badge_shape = slide14.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, badge_left, badge_top,
            badge_width, badge_height
        )
        fill = badge_shape.fill
        fill.solid()
        fill.fore_color.rgb = border_color
        badge_shape.line.fill.background()
        
        badge_text = badge_shape.text_frame
        badge_text.text = priority
        badge_text.paragraphs[0].font.name = TITLE_FONT_NAME
        badge_text.paragraphs[0].font.size = Pt(12)
        badge_text.paragraphs[0].font.bold = True
        badge_text.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
        badge_text.paragraphs[0].alignment = PP_ALIGN.CENTER
        badge_text.vertical_anchor = 1  # Middle
        
        # Add item title next to badge
        item_title_left = badge_left + badge_width + Inches(0.2)
        item_title_top = badge_top
        item_title_width = card_width - badge_width - Inches(0.6)
        item_title_height = badge_height
        
        item_title_box = slide14.shapes.add_textbox(
            item_title_left, item_title_top, item_title_width, item_title_height
        )
        item_title_frame = item_title_box.text_frame
        item_title_frame.word_wrap = True
        item_title_paragraph = item_title_frame.paragraphs[0]
        item_title_paragraph.text = f"Item {i+1} - {item['title']}"
        item_title_paragraph.font.name = TITLE_FONT_NAME
        item_title_paragraph.font.size = Pt(16)
        item_title_paragraph.font.bold = True
        item_title_paragraph.font.color.rgb = CS_NAVY
        item_title_paragraph.alignment = PP_ALIGN.LEFT
        
        # Add description
        desc_left = card_left + Inches(0.2)
        desc_top = badge_top + badge_height + Inches(0.15)
        desc_width = card_width - Inches(0.4)
        desc_height = Inches(0.4)
        
        desc_box = slide14.shapes.add_textbox(desc_left, desc_top, desc_width, desc_height)
        desc_frame = desc_box.text_frame
        desc_frame.word_wrap = True
        desc_paragraph = desc_frame.paragraphs[0]
        desc_paragraph.text = concise_descriptions[i]
        desc_paragraph.font.name = BODY_FONT_NAME
        desc_paragraph.font.size = Pt(12)
        desc_paragraph.font.color.rgb = CS_SLATE
        desc_paragraph.alignment = PP_ALIGN.LEFT
        
        # Add metadata (Owner, Target, Expected Impact) in two rows
        meta_left = desc_left
        meta_top = desc_top + desc_height + Inches(0.1)
        meta_width = card_width - Inches(0.4)
        meta_height = Inches(0.35)
        
        # Row 1: Owner and Target
        meta_text1 = f"Owner: {item['owner']} | Target: {item['target']}"
        meta_box1 = slide14.shapes.add_textbox(meta_left, meta_top, meta_width, meta_height)
        meta_frame1 = meta_box1.text_frame
        meta_frame1.word_wrap = True
        meta_paragraph1 = meta_frame1.paragraphs[0]
        meta_paragraph1.text = meta_text1
        meta_paragraph1.font.name = BODY_FONT_NAME
        meta_paragraph1.font.size = Pt(11)
        meta_paragraph1.font.color.rgb = CS_SLATE
        meta_paragraph1.alignment = PP_ALIGN.LEFT
        
        # Row 2: Expected Impact
        meta_top2 = meta_top + meta_height + Inches(0.05)
        meta_text2 = f"Expected Impact: {expected_impacts[i]}"
        meta_box2 = slide14.shapes.add_textbox(meta_left, meta_top2, meta_width, meta_height)
        meta_frame2 = meta_box2.text_frame
        meta_frame2.word_wrap = True
        meta_paragraph2 = meta_frame2.paragraphs[0]
        meta_paragraph2.text = meta_text2
        meta_paragraph2.font.name = BODY_FONT_NAME
        meta_paragraph2.font.size = Pt(11)
        meta_paragraph2.font.bold = True
        meta_paragraph2.font.color.rgb = CS_NAVY
        meta_paragraph2.alignment = PP_ALIGN.LEFT
    
    # Slide 15 - Operational Insights
    slide15 = prs.slides.add_slide(blank_slide_layout)
    
    # Add logo at top right
    add_logo(slide15, position='top_right', prs=prs)
    
    # Add title header
    header_shape15 = slide15.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0,
        prs.slide_width, header_height
    )
    fill = header_shape15.fill
    fill.solid()
    fill.fore_color.rgb = CS_NAVY
    header_shape15.line.fill.background()
    
    # Add title text on header
    title_box15 = slide15.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_frame15 = title_box15.text_frame
    title_frame15.word_wrap = True
    title_paragraph15 = title_frame15.paragraphs[0]
    title_paragraph15.text = "Operational Insights"
    title_paragraph15.font.name = TITLE_FONT_NAME
    title_paragraph15.font.size = Pt(28)
    title_paragraph15.font.bold = True
    title_paragraph15.font.color.rgb = RGBColor(255, 255, 255)
    title_paragraph15.alignment = PP_ALIGN.LEFT
    
    # Create two-column layout
    column_width = (prs.slide_width - Inches(1.8)) / 2  # Two columns with spacing
    column_spacing = Inches(0.3)
    left_column_left = Inches(0.5)
    right_column_left = left_column_left + column_width + column_spacing
    column_top = header_height + Inches(0.4)
    column_height = Inches(3.5)
    
    # Left column - After-Hours Coverage
    left_card_shape = slide15.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left_column_left, column_top,
        column_width, column_height
    )
    fill = left_card_shape.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(240, 248, 255)  # Light blue background
    line = left_card_shape.line
    line.color.rgb = CS_BLUE
    line.width = Pt(2)
    
    # Left column title
    left_title_left = left_column_left + Inches(0.2)
    left_title_top = column_top + Inches(0.2)
    left_title_width = column_width - Inches(0.4)
    left_title_height = Inches(0.4)
    
    left_title_box = slide15.shapes.add_textbox(
        left_title_left, left_title_top, left_title_width, left_title_height
    )
    left_title_frame = left_title_box.text_frame
    left_title_frame.word_wrap = True
    left_title_paragraph = left_title_frame.paragraphs[0]
    left_title_paragraph.text = "After-Hours Coverage"
    left_title_paragraph.font.name = TITLE_FONT_NAME
    left_title_paragraph.font.size = Pt(18)
    left_title_paragraph.font.bold = True
    left_title_paragraph.font.color.rgb = CS_NAVY
    left_title_paragraph.alignment = PP_ALIGN.LEFT
    
    # After Hours Calls: 18% (136 of 756)
    after_hours_left = left_title_left
    after_hours_top = left_title_top + left_title_height + Inches(0.3)
    after_hours_width = left_title_width
    after_hours_height = Inches(0.5)
    
    after_hours_box = slide15.shapes.add_textbox(
        after_hours_left, after_hours_top, after_hours_width, after_hours_height
    )
    after_hours_frame = after_hours_box.text_frame
    after_hours_frame.word_wrap = True
    after_hours_paragraph = after_hours_frame.paragraphs[0]
    after_hours_paragraph.text = "After Hours Calls: 18% (136 of 756)"
    after_hours_paragraph.font.name = BODY_FONT_NAME
    after_hours_paragraph.font.size = Pt(14)
    after_hours_paragraph.font.color.rgb = CS_SLATE
    after_hours_paragraph.alignment = PP_ALIGN.LEFT
    
    # Weekend Calls: 15% (22 of 147)
    weekend_left = after_hours_left
    weekend_top = after_hours_top + after_hours_height + Inches(0.2)
    weekend_width = after_hours_width
    weekend_height = after_hours_height
    
    weekend_box = slide15.shapes.add_textbox(
        weekend_left, weekend_top, weekend_width, weekend_height
    )
    weekend_frame = weekend_box.text_frame
    weekend_frame.word_wrap = True
    weekend_paragraph = weekend_frame.paragraphs[0]
    weekend_paragraph.text = "Weekend Calls: 15% (22 of 147)"
    weekend_paragraph.font.name = BODY_FONT_NAME
    weekend_paragraph.font.size = Pt(14)
    weekend_paragraph.font.color.rgb = CS_SLATE
    weekend_paragraph.alignment = PP_ALIGN.LEFT
    
    # Total Off-Hours Calls: 158
    total_left = weekend_left
    total_top = weekend_top + weekend_height + Inches(0.2)
    total_width = weekend_width
    total_height = after_hours_height
    
    total_box = slide15.shapes.add_textbox(
        total_left, total_top, total_width, total_height
    )
    total_frame = total_box.text_frame
    total_frame.word_wrap = True
    total_paragraph = total_frame.paragraphs[0]
    total_paragraph.text = f"Total Off-Hours Calls: {data.after_hours_escalations}"
    total_paragraph.font.name = BODY_FONT_NAME
    total_paragraph.font.size = Pt(14)
    total_paragraph.font.bold = True
    total_paragraph.font.color.rgb = CS_NAVY
    total_paragraph.alignment = PP_ALIGN.LEFT
    
    # Right column - Collaboration Quality
    right_card_shape = slide15.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, right_column_left, column_top,
        column_width, column_height
    )
    fill = right_card_shape.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(240, 248, 255)  # Light blue background
    line = right_card_shape.line
    line.color.rgb = CS_BLUE
    line.width = Pt(2)
    
    # Right column title
    right_title_left = right_column_left + Inches(0.2)
    right_title_top = column_top + Inches(0.2)
    right_title_width = column_width - Inches(0.4)
    right_title_height = Inches(0.4)
    
    right_title_box = slide15.shapes.add_textbox(
        right_title_left, right_title_top, right_title_width, right_title_height
    )
    right_title_frame = right_title_box.text_frame
    right_title_frame.word_wrap = True
    right_title_paragraph = right_title_frame.paragraphs[0]
    right_title_paragraph.text = "Collaboration Quality"
    right_title_paragraph.font.name = TITLE_FONT_NAME
    right_title_paragraph.font.size = Pt(18)
    right_title_paragraph.font.bold = True
    right_title_paragraph.font.color.rgb = CS_NAVY
    right_title_paragraph.alignment = PP_ALIGN.LEFT
    
    # Average Touches per Incident: 2.3
    touches_left = right_title_left
    touches_top = right_title_top + right_title_height + Inches(0.3)
    touches_width = right_title_width
    touches_height = after_hours_height
    
    touches_box = slide15.shapes.add_textbox(
        touches_left, touches_top, touches_width, touches_height
    )
    touches_frame = touches_box.text_frame
    touches_frame.word_wrap = True
    touches_paragraph = touches_frame.paragraphs[0]
    touches_paragraph.text = f"Average Touches per Incident: {data.avg_touches}"
    touches_paragraph.font.name = BODY_FONT_NAME
    touches_paragraph.font.size = Pt(14)
    touches_paragraph.font.color.rgb = CS_SLATE
    touches_paragraph.alignment = PP_ALIGN.LEFT
    
    # Active Client Participation: 72%
    participation_left = touches_left
    participation_top = touches_top + touches_height + Inches(0.2)
    participation_width = touches_width
    participation_height = touches_height
    
    participation_box = slide15.shapes.add_textbox(
        participation_left, participation_top, participation_width, participation_height
    )
    participation_frame = participation_box.text_frame
    participation_frame.word_wrap = True
    participation_paragraph = participation_frame.paragraphs[0]
    participation_paragraph.text = f"Active Client Participation: {data.client_participation}"
    participation_paragraph.font.name = BODY_FONT_NAME
    participation_paragraph.font.size = Pt(14)
    participation_paragraph.font.color.rgb = CS_SLATE
    participation_paragraph.alignment = PP_ALIGN.LEFT
    
    # Client-Led Closures: 21%
    closures_left = participation_left
    closures_top = participation_top + participation_height + Inches(0.2)
    closures_width = participation_width
    closures_height = touches_height
    
    closures_box = slide15.shapes.add_textbox(
        closures_left, closures_top, closures_width, closures_height
    )
    closures_frame = closures_box.text_frame
    closures_frame.word_wrap = True
    closures_paragraph = closures_frame.paragraphs[0]
    closures_paragraph.text = f"Client-Led Closures: {data.client_led_closures}"
    closures_paragraph.font.name = BODY_FONT_NAME
    closures_paragraph.font.size = Pt(14)
    closures_paragraph.font.bold = True
    closures_paragraph.font.color.rgb = CS_NAVY
    closures_paragraph.alignment = PP_ALIGN.LEFT
    
    # Add insight box at bottom spanning both columns
    insight_box_top = column_top + column_height + Inches(0.3)
    insight_box_left = left_column_left
    insight_box_width = prs.slide_width - Inches(1.0)
    insight_box_height = Inches(0.7)
    
    insight_box_shape = slide15.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, insight_box_left, insight_box_top,
        insight_box_width, insight_box_height
    )
    fill = insight_box_shape.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(240, 248, 255)  # Light blue background
    line = insight_box_shape.line
    line.color.rgb = CS_BLUE
    line.width = Pt(2)
    
    insight_text = insight_box_shape.text_frame
    insight_text.text = "82.5% of after-hours alerts handled by CS SOC without requiring customer notification"
    insight_text.paragraphs[0].font.name = BODY_FONT_NAME
    insight_text.paragraphs[0].font.size = Pt(15)
    insight_text.paragraphs[0].font.bold = True
    insight_text.paragraphs[0].font.color.rgb = CS_NAVY
    insight_text.paragraphs[0].alignment = PP_ALIGN.CENTER
    insight_text.vertical_anchor = 1  # Middle


def build_forward_direction_slide(prs, data):
    """Create the Forward Direction slide (Slide 16).
    
    Args:
        prs (Presentation): The presentation object.
        data (ReportData): The report data object containing all metrics.
    """
    blank_slide_layout = prs.slide_layouts[6]  # Blank layout
    header_height = Inches(0.8)
    title_left = Inches(0.5)
    title_top = Inches(0.1)
    title_width = prs.slide_width - Inches(2.5)
    title_height = Inches(0.6)
    
    # Slide 16 - Forward Direction
    slide16 = prs.slides.add_slide(blank_slide_layout)
    
    # Add logo at top right
    add_logo(slide16, position='top_right', prs=prs)
    
    # Add title header
    header_shape16 = slide16.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0,
        prs.slide_width, header_height
    )
    fill = header_shape16.fill
    fill.solid()
    fill.fore_color.rgb = CS_NAVY
    header_shape16.line.fill.background()
    
    # Add title text on header
    title_box16 = slide16.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_frame16 = title_box16.text_frame
    title_frame16.word_wrap = True
    title_paragraph16 = title_frame16.paragraphs[0]
    title_paragraph16.text = "Looking Ahead"
    title_paragraph16.font.name = TITLE_FONT_NAME
    title_paragraph16.font.size = Pt(28)
    title_paragraph16.font.bold = True
    title_paragraph16.font.color.rgb = RGBColor(255, 255, 255)
    title_paragraph16.alignment = PP_ALIGN.LEFT
    
    # Section 1 - Next Period Targets
    section1_top = header_height + Inches(0.4)
    section1_left = Inches(0.5)
    section1_width = prs.slide_width - Inches(1.0)
    section1_height = Inches(1.2)
    
    # Section 1 title
    section1_title_box = slide16.shapes.add_textbox(
        section1_left, section1_top, section1_width, Inches(0.4)
    )
    section1_title_frame = section1_title_box.text_frame
    section1_title_frame.word_wrap = True
    section1_title_paragraph = section1_title_frame.paragraphs[0]
    section1_title_paragraph.text = "Next Period Targets"
    section1_title_paragraph.font.name = TITLE_FONT_NAME
    section1_title_paragraph.font.size = Pt(20)
    section1_title_paragraph.font.bold = True
    section1_title_paragraph.font.color.rgb = CS_NAVY
    section1_title_paragraph.alignment = PP_ALIGN.LEFT
    
    # Section 1 content (bullets)
    section1_content_top = section1_top + Inches(0.5)
    section1_content_box = slide16.shapes.add_textbox(
        section1_left + Inches(0.3), section1_content_top,
        section1_width - Inches(0.3), section1_height - Inches(0.5)
    )
    section1_content_frame = section1_content_box.text_frame
    section1_content_frame.word_wrap = True
    section1_content_frame.margin_left = Inches(0.2)
    section1_content_frame.margin_right = Inches(0.2)
    section1_content_frame.margin_top = Inches(0.1)
    section1_content_frame.margin_bottom = Inches(0.1)
    
    # Get Palo Alto XDR FP rate from detection_sources
    palo_alto_fp_rate = 11.2
    for source in data.detection_sources:
        if "Palo Alto" in source.get('source', ''):
            palo_alto_fp_rate = source.get('fp_rate', 11.2)
            break
    
    # Calculate reduction (11.2% to 10% = ~4 fewer escalations)
    # Assuming 189 incidents at 11.2% FP rate, reducing to 10% would save ~2-4 escalations
    target1_text = f"Trim Palo Alto XDR false positives from {palo_alto_fp_rate}% to 10% threshold (~4 fewer escalations)"
    target2_text = f"Reduce manual escalations from {data.analyst_escalation['count']} ({data.analyst_escalation['percent']}%) to 32 or fewer (12% target)"
    
    # Add bullet points with checkmark-like styling (using blue color)
    para1 = section1_content_frame.paragraphs[0]
    para1.text = "✓ " + target1_text
    para1.font.name = BODY_FONT_NAME
    para1.font.size = Pt(14)
    para1.font.color.rgb = CS_BLUE
    para1.level = 0
    para1.space_after = Pt(8)
    
    para2 = section1_content_frame.add_paragraph()
    para2.text = "✓ " + target2_text
    para2.font.name = BODY_FONT_NAME
    para2.font.size = Pt(14)
    para2.font.color.rgb = CS_BLUE
    para2.level = 0
    para2.space_after = Pt(8)
    
    # Section 2 - Strategic Focus
    section2_top = section1_top + section1_height + Inches(0.3)
    section2_height = Inches(1.0)
    
    # Section 2 title
    section2_title_box = slide16.shapes.add_textbox(
        section1_left, section2_top, section1_width, Inches(0.4)
    )
    section2_title_frame = section2_title_box.text_frame
    section2_title_frame.word_wrap = True
    section2_title_paragraph = section2_title_frame.paragraphs[0]
    section2_title_paragraph.text = "Strategic Focus"
    section2_title_paragraph.font.name = TITLE_FONT_NAME
    section2_title_paragraph.font.size = Pt(20)
    section2_title_paragraph.font.bold = True
    section2_title_paragraph.font.color.rgb = CS_NAVY
    section2_title_paragraph.alignment = PP_ALIGN.LEFT
    
    # Section 2 content
    section2_content_top = section2_top + Inches(0.5)
    section2_content_box = slide16.shapes.add_textbox(
        section1_left + Inches(0.3), section2_content_top,
        section1_width - Inches(0.3), section2_height - Inches(0.5)
    )
    section2_content_frame = section2_content_box.text_frame
    section2_content_frame.word_wrap = True
    section2_content_frame.margin_left = Inches(0.2)
    section2_content_frame.margin_right = Inches(0.2)
    section2_content_frame.margin_top = Inches(0.1)
    section2_content_frame.margin_bottom = Inches(0.1)
    
    focus1_text = "Proactive hunts targeting Persistence and Defense Evasion tactics"
    focus2_text = "Additional playbook coverage for analyst-escalated scenarios"
    
    para3 = section2_content_frame.paragraphs[0]
    para3.text = "✓ " + focus1_text
    para3.font.name = BODY_FONT_NAME
    para3.font.size = Pt(14)
    para3.font.color.rgb = CS_BLUE
    para3.level = 0
    para3.space_after = Pt(8)
    
    para4 = section2_content_frame.add_paragraph()
    para4.text = "✓ " + focus2_text
    para4.font.name = BODY_FONT_NAME
    para4.font.size = Pt(14)
    para4.font.color.rgb = CS_BLUE
    para4.level = 0
    para4.space_after = Pt(8)
    
    # Section 3 - Your Partnership
    section3_top = section2_top + section2_height + Inches(0.3)
    section3_height = Inches(0.8)
    
    # Section 3 title
    section3_title_box = slide16.shapes.add_textbox(
        section1_left, section3_top, section1_width, Inches(0.4)
    )
    section3_title_frame = section3_title_box.text_frame
    section3_title_frame.word_wrap = True
    section3_title_paragraph = section3_title_frame.paragraphs[0]
    section3_title_paragraph.text = "Your Partnership"
    section3_title_paragraph.font.name = TITLE_FONT_NAME
    section3_title_paragraph.font.size = Pt(20)
    section3_title_paragraph.font.bold = True
    section3_title_paragraph.font.color.rgb = CS_NAVY
    section3_title_paragraph.alignment = PP_ALIGN.LEFT
    
    # Section 3 content
    section3_content_top = section3_top + Inches(0.5)
    section3_content_box = slide16.shapes.add_textbox(
        section1_left, section3_content_top,
        section1_width, section3_height - Inches(0.5)
    )
    section3_content_frame = section3_content_box.text_frame
    section3_content_frame.word_wrap = True
    section3_content_paragraph = section3_content_frame.paragraphs[0]
    section3_content_paragraph.text = "Questions? Your Customer Success Manager is here to help."
    section3_content_paragraph.font.name = BODY_FONT_NAME
    section3_content_paragraph.font.size = Pt(16)
    section3_content_paragraph.font.color.rgb = CS_SLATE
    section3_content_paragraph.alignment = PP_ALIGN.LEFT
    
    # Footer - Report generated date and data sources
    footer_top = prs.slide_height - Inches(0.5)
    footer_left = Inches(0.5)
    footer_width = prs.slide_width - Inches(1.0)
    footer_height = Inches(0.3)
    
    footer_box = slide16.shapes.add_textbox(footer_left, footer_top, footer_width, footer_height)
    footer_frame = footer_box.text_frame
    footer_frame.word_wrap = True
    footer_paragraph = footer_frame.paragraphs[0]
    footer_paragraph.text = f"Report generated: {data.report_date} | Data sources: Escalation Details Report"
    footer_paragraph.font.name = BODY_FONT_NAME
    footer_paragraph.font.size = Pt(10)
    footer_paragraph.font.color.rgb = CS_SLATE
    footer_paragraph.alignment = PP_ALIGN.LEFT
    footer_paragraph.font.italic = True


if __name__ == "__main__":
    main()

