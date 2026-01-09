"""Escalation to Client Details Report - PowerPoint Presentation Generator

This module generates branded PowerPoint presentations from report data.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pathlib import Path
import argparse
import logging
import shutil
from datetime import datetime
from typing import Dict, Optional
from PIL import Image

# Import constants from dedicated module
from constants import (
    CS_BLUE, CS_NAVY, CS_SLATE, CS_VIOLET, CS_RED, CS_ORANGE, CS_GREEN,
    TITLE_FONT_NAME, BODY_FONT_NAME, FALLBACK_TITLE_FONT, FALLBACK_BODY_FONT,
    H1_FONT_SIZE, H2_FONT_SIZE, H3_FONT_SIZE, H4_FONT_SIZE,
    H5_FONT_SIZE, H6_FONT_SIZE, PARAGRAPH_FONT_SIZE, FOOTER_FONT_SIZE,
    TITLE_FONT_SIZE, SUBTITLE_FONT_SIZE, HEADING_FONT_SIZE,
    BODY_FONT_SIZE, CAPTION_FONT_SIZE, METRIC_FONT_SIZE, HERO_METRIC_FONT_SIZE,
    MARGIN_STANDARD, MARGIN_CONTENT, CARD_SPACING, HEADER_HEIGHT, FOOTER_HEIGHT,
    PRESENTATION_TITLE, PRESENTATION_INTENT, COPYRIGHT_TEXT,
    SECTION_NARRATIVES, get_brand_colors
)

# Import helper functions from dedicated module
from helpers import (
    create_presentation, apply_branding, get_slide_number,
    add_master_slide_header, add_master_slide_footer, add_master_slide_elements,
    add_logo, setup_content_slide, create_gradient_background,
    add_header_bar, add_insight_callout, create_metric_card, create_data_table
)


# =============================================================================
# SLIDE BUILDER FUNCTIONS
# =============================================================================

def build_key_takeaways_slide(prs, section_title, takeaways, data=None):
    """Create a Key Takeaways slide for executive summary of a section.
    
    Per CRITICALSTART branding guidelines:
    - All slides have transparent header and footer
    - Uses H1-H6 typography scale
    
    Args:
        prs (Presentation): The presentation object.
        section_title (str): The section this summarizes (e.g., "Value Delivered")
        takeaways (list): List of takeaway strings (3-4 recommended)
        data (ReportData, optional): Data object for dynamic values
    
    Returns:
        The slide object.
    """
    # Use setup_content_slide helper for consistent branding
    slide, content_top = setup_content_slide(prs, f"Key Takeaways: {section_title}")
    
    # Create takeaway cards
    card_top = content_top + Inches(0.1)
    card_height = Inches(0.9)
    card_spacing = Inches(0.15)
    card_width = prs.slide_width - MARGIN_STANDARD * 2
    card_left = MARGIN_STANDARD
    
    for i, takeaway in enumerate(takeaways):
        current_top = card_top + i * (card_height + card_spacing)
        
        # Card background
        card_shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, card_left, current_top,
            card_width, card_height
        )
        fill = card_shape.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(240, 248, 255)
        line = card_shape.line
        line.color.rgb = CS_BLUE
        line.width = Pt(2)
        
        # Checkmark icon (using text)
        check_box = slide.shapes.add_textbox(
            card_left + Inches(0.15), current_top + Inches(0.25),
            Inches(0.5), Inches(0.5)
        )
        check_frame = check_box.text_frame
        check_para = check_frame.paragraphs[0]
        check_para.text = "✓"
        check_para.font.name = TITLE_FONT_NAME
        check_para.font.size = Pt(28)
        check_para.font.bold = True
        check_para.font.color.rgb = CS_BLUE
        check_para.alignment = PP_ALIGN.CENTER
        
        # Takeaway text
        text_box = slide.shapes.add_textbox(
            card_left + Inches(0.7), current_top + Inches(0.2),
            card_width - Inches(0.9), card_height - Inches(0.3)
        )
        text_frame = text_box.text_frame
        text_frame.word_wrap = True
        text_para = text_frame.paragraphs[0]
        text_para.text = takeaway
        text_para.font.name = BODY_FONT_NAME
        text_para.font.size = Pt(18)
        text_para.font.color.rgb = CS_NAVY
        text_para.alignment = PP_ALIGN.LEFT
    
    # Add insight bar at bottom
    insight_top = card_top + len(takeaways) * (card_height + card_spacing) + Inches(0.1)
    add_insight_callout(
        slide, prs,
        "Bottom Line",
        "Faster response, complete containment, measurable value—your security investment is delivering.",
        insight_top,
        height=Inches(0.7)
    )
    
    return slide


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
    
    # Note: Logo removed from title slide per branding guidelines
    
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


def create_section_header_layout(prs, section_title, narrative_subtitle=None):
    """Create section divider slide with gray background and optional narrative subtitle.
    
    Per brand guidelines, section title slides use Charcoal gray background (#343741)
    with white text. These slides serve as narrative transitions that guide executives
    through the presentation story.
    
    Args:
        prs (Presentation): The presentation object.
        section_title (str): Section title text (H2 typography, 72pt).
        narrative_subtitle (str, optional): Narrative subtitle text (H4 typography, 27pt).
    
    Returns:
        Slide: The created slide object.
    """
    blank_slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(blank_slide_layout)
    slide_number = get_slide_number(prs)
    
    # Add gray background (CS_SLATE - Charcoal #343741)
    background_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0,
        prs.slide_width, prs.slide_height
    )
    fill = background_shape.fill
    fill.solid()
    fill.fore_color.rgb = CS_SLATE  # Use charcoal gray per brand guidelines
    background_shape.line.fill.background()
    
    # Add master slide elements (header and footer) with white text for WCAG AA compliance
    # Gray background (#343741) requires white text for sufficient contrast
    add_master_slide_elements(slide, prs, slide_number=slide_number,
                               include_header=True, include_footer=True,
                               text_color=RGBColor(255, 255, 255))
    
    # Note: Logo removed from section cards per branding guidelines
    
    # Calculate vertical positioning
    # If we have a subtitle, move title up slightly to accommodate both
    if narrative_subtitle:
        title_top = Inches(1.8)
    else:
        title_top = Inches(2.2)
    
    # Add section title (centered) - H2 typography (72pt)
    title_left = Inches(1)
    title_width = prs.slide_width - Inches(2)
    title_height = Inches(1.2)
    
    title_box = slide.shapes.add_textbox(title_left, title_top,
                                        title_width, title_height)
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    title_paragraph = title_frame.paragraphs[0]
    title_paragraph.text = section_title
    title_paragraph.font.name = TITLE_FONT_NAME
    title_paragraph.font.size = H2_FONT_SIZE  # 72pt per typography scale
    title_paragraph.font.bold = True
    title_paragraph.font.color.rgb = RGBColor(255, 255, 255)  # White text
    title_paragraph.alignment = PP_ALIGN.CENTER
    
    # Add narrative subtitle if provided - H4 typography (27pt)
    if narrative_subtitle:
        subtitle_top = title_top + Inches(1.3)  # 0.3" below title
        subtitle_height = Inches(0.8)
        
        subtitle_box = slide.shapes.add_textbox(title_left, subtitle_top,
                                                title_width, subtitle_height)
        subtitle_frame = subtitle_box.text_frame
        subtitle_frame.word_wrap = True
        subtitle_paragraph = subtitle_frame.paragraphs[0]
        subtitle_paragraph.text = narrative_subtitle
        subtitle_paragraph.font.name = BODY_FONT_NAME
        subtitle_paragraph.font.size = H4_FONT_SIZE  # 27pt per typography scale
        subtitle_paragraph.font.bold = False
        subtitle_paragraph.font.color.rgb = RGBColor(255, 255, 255)  # White text
        subtitle_paragraph.alignment = PP_ALIGN.CENTER
    
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


def insert_chart_image(slide, placeholder_id: Optional[str], image_path: str, 
                       padding: float = 0.1, min_fill_ratio: float = 0.85):
    """Replace a placeholder shape with a chart image.
    
    This function locates a placeholder shape (identified by "[Chart:" text)
    and replaces it with a chart image while preserving aspect ratio and
    centering within the placeholder bounds.
    
    Args:
        slide: The slide object containing the placeholder
        placeholder_id: Optional ID string to identify the placeholder (e.g., "severity_sankey")
                       If None, searches for first placeholder with "[Chart:" text
        image_path: Path to the image file to insert
        padding: Padding around the image in inches (default 0.1")
        min_fill_ratio: Minimum ratio of placeholder area to fill (0.0-1.0, default 0.85)
    
    Returns:
        bool: True if placeholder was found and replaced, False otherwise
    
    Note:
        Charts are rendered at 2x DPI for clarity. The function automatically
        handles aspect ratio preservation and centering.
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
                    if f"ID: {placeholder_id}" in text or placeholder_id in text:
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
    placeholder_width = placeholder_shape.width
    placeholder_height = placeholder_shape.height
    
    # Get actual image dimensions to preserve aspect ratio
    # Charts are rendered at scale=2, so we need to account for that
    # The images are rendered at 2x DPI, so physical size is half the pixel dimensions
    try:
        with Image.open(image_path) as img:
            img_width_px, img_height_px = img.size
            # Convert pixels to inches (assuming 96 DPI base, but rendered at 2x scale = 192 DPI effective)
            # So 1 pixel = 1/192 inches at the rendered scale
            # But since we're fitting to placeholder, we just need aspect ratio
            img_aspect_ratio = img_width_px / img_height_px
    except Exception as e:
        logging.warning(f"Could not read image dimensions for {image_path}: {e}. Using placeholder dimensions.")
        img_aspect_ratio = placeholder_width / placeholder_height
    
    # Calculate placeholder aspect ratio
    placeholder_aspect_ratio = placeholder_width / placeholder_height
    
    # Adjust dimensions to preserve image aspect ratio while fitting within placeholder bounds
    # Work with the aspect ratios and fit to the constraining dimension
    if img_aspect_ratio > placeholder_aspect_ratio:
        # Image is wider relative to placeholder - fit to placeholder width, adjust height
        width = placeholder_width
        height = placeholder_width / img_aspect_ratio
        # Center vertically if height is less than placeholder
        if height < placeholder_height:
            top = top + (placeholder_height - height) / 2
    else:
        # Image is taller relative to placeholder - fit to placeholder height, adjust width
        height = placeholder_height
        width = placeholder_height * img_aspect_ratio
        # Center horizontally if width is less than placeholder
        if width < placeholder_width:
            left = left + (placeholder_width - width) / 2
    
    # Remove the placeholder shape
    slide.shapes._spTree.remove(placeholder_shape._element)
    
    # Insert the image with preserved aspect ratio
    try:
        slide.shapes.add_picture(str(image_path), left, top, width, height)
        logging.info(f"Inserted chart image: {image_path} (ID: {placeholder_id}) - Size: {width}x{height}, Aspect ratio preserved")
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
        # Build slides in order with section title cards for narrative flow
        # Narrative Arc: Current State → Value Delivered → Performance → Threats → Opportunities
        
        logger.info("  Building executive summary slides (includes title slide and section card)...")
        build_executive_summary_slides(prs, data)
        
        # Insert Value Delivered section title card
        logger.info("  Inserting 'Value Delivered' section card...")
        create_section_header_layout(
            prs,
            SECTION_NARRATIVES["value_delivered"]["title"],
            SECTION_NARRATIVES["value_delivered"]["narrative"]
        )
        
        logger.info("  Building value delivered slides...")
        build_value_delivered_slides(prs, data)
        
        # Insert Protection Achieved section title card
        logger.info("  Inserting 'Protection Achieved' section card...")
        create_section_header_layout(
            prs,
            SECTION_NARRATIVES["protection_achieved"]["title"],
            SECTION_NARRATIVES["protection_achieved"]["narrative"]
        )
        
        logger.info("  Building protection achieved slides...")
        build_protection_achieved_slides(prs, data)
        
        if not args.no_threat_landscape:
            # Insert Threat Landscape section title card
            logger.info("  Inserting 'Threat Landscape' section card...")
            create_section_header_layout(
                prs,
                SECTION_NARRATIVES["threat_landscape"]["title"],
                SECTION_NARRATIVES["threat_landscape"]["narrative"]
            )
            
            logger.info("  Building threat landscape slides...")
            build_threat_landscape_slides(prs, data, include=True)
        else:
            logger.info("  Skipping threat landscape slides (--no-threat-landscape)")
            build_threat_landscape_slides(prs, data, include=False)
        
        # Insert Insights section title card
        logger.info("  Inserting 'Insights & Continuous Improvement' section card...")
        create_section_header_layout(
            prs,
            SECTION_NARRATIVES["insights"]["title"],
            SECTION_NARRATIVES["insights"]["narrative"]
        )
        
        logger.info("  Building insights slides...")
        build_insights_slides(prs, data)
        
        logger.info("  Building additional content slides...")
        build_additional_content_slides(prs, data)
        
        # Insert Forward Direction section title card
        logger.info("  Inserting 'Forward Direction' section card...")
        create_section_header_layout(
            prs,
            SECTION_NARRATIVES["forward_direction"]["title"],
            SECTION_NARRATIVES["forward_direction"]["narrative"]
        )
        
        logger.info("  Building key takeaways slide...")
        # Build executive summary key takeaways
        takeaways = [
            f"{data.response_advantage_percent}% faster response than industry peers—threats are contained before spreading",
            f"100% threat containment with zero breaches this period across {data.true_threats_contained} true positive incidents",
            "Millions in modeled cost exposure avoided through proactive security operations",
            f"{data.after_hours_escalations} after-hours escalations handled seamlessly with {int(data.automation_percent)}% automation"
        ]
        build_key_takeaways_slide(prs, "This Period", takeaways, data)
        
        logger.info("  Building forward direction slide...")
        build_forward_direction_slide(prs, data)
        
        logger.info("  Building contact slide...")
        build_contact_slide(prs, data)
        
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
    
    Per CRITICALSTART branding guidelines:
    - Title slide: No header, footer only, H1 typography for main title
    - Content slides: Header and footer with transparent background
    
    Args:
        prs (Presentation): The presentation object.
        data (ReportData): The report data object containing all metrics.
    """
    # Slide 1 - Title Slide (NO HEADER per branding guidelines)
    blank_slide_layout = prs.slide_layouts[6]  # Blank layout
    slide1 = prs.slides.add_slide(blank_slide_layout)
    
    # Add gradient background (blue sweep: #009CDE → #004C97)
    create_gradient_background(prs, slide1, 'blue_sweep')
    
    # Note: Logo removed from title slide per branding guidelines
    
    # Add main title using H1 typography (114pt per branding guide)
    title_left = Inches(1)
    title_top = Inches(1.5)
    title_width = prs.slide_width - Inches(2)
    title_height = Inches(1.5)
    
    title_box = slide1.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    title_paragraph = title_frame.paragraphs[0]
    title_paragraph.text = "EXECUTIVE BUSINESS\nREVIEW"
    title_paragraph.font.name = TITLE_FONT_NAME
    title_paragraph.font.size = H1_FONT_SIZE  # 114pt per branding guide
    title_paragraph.font.bold = True
    title_paragraph.font.color.rgb = RGBColor(255, 255, 255)
    title_paragraph.alignment = PP_ALIGN.LEFT
    
    # Add client name using H5 typography
    client_left = Inches(1)
    client_top = Inches(3.2)
    client_width = prs.slide_width - Inches(2)
    client_height = Inches(0.5)
    
    client_box = slide1.shapes.add_textbox(client_left, client_top, client_width, client_height)
    client_frame = client_box.text_frame
    client_frame.word_wrap = True
    client_paragraph = client_frame.paragraphs[0]
    client_paragraph.text = data.client_name
    client_paragraph.font.name = BODY_FONT_NAME
    client_paragraph.font.size = H5_FONT_SIZE  # 16pt
    client_paragraph.font.color.rgb = RGBColor(255, 255, 255)
    client_paragraph.alignment = PP_ALIGN.LEFT
    
    # Add period using paragraph typography
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
    period_top = Inches(3.7)
    period_width = prs.slide_width - Inches(2)
    period_height = Inches(0.5)
    
    period_box = slide1.shapes.add_textbox(period_left, period_top, period_width, period_height)
    period_frame = period_box.text_frame
    period_frame.word_wrap = True
    period_paragraph = period_frame.paragraphs[0]
    period_paragraph.text = period_text
    period_paragraph.font.name = BODY_FONT_NAME
    period_paragraph.font.size = PARAGRAPH_FONT_SIZE  # 12pt
    period_paragraph.font.color.rgb = RGBColor(255, 255, 255)
    period_paragraph.alignment = PP_ALIGN.LEFT
    
    # Add footer only (no header on title slide per branding guidelines)
    add_master_slide_elements(slide1, prs, slide_number=None, 
                               include_header=False, include_footer=True)
    
    # Note: Report date is now in the footer per branding guidelines
    report_date_left = Inches(1)
    report_date_top = prs.slide_height - Inches(0.9)  # Moved up to avoid footer
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
    
    # Insert Executive Summary section title card
    create_section_header_layout(
        prs, 
        SECTION_NARRATIVES["executive_summary"]["title"],
        SECTION_NARRATIVES["executive_summary"]["narrative"]
    )
    
    # Slide 3 - Executive Dashboard (Merged BLUF)
    # Combines Security Posture + Value metrics into single consolidated dashboard
    slide2 = prs.slides.add_slide(blank_slide_layout)
    slide2_number = get_slide_number(prs)
    
    # Add master slide elements (header and footer with transparent background)
    add_master_slide_elements(slide2, prs, slide_number=slide2_number,
                               include_header=True, include_footer=True)
    
    # Add logo at top right
    add_logo(slide2, position='top_right', prs=prs)
    
    # Content area starts below header
    content_top = HEADER_HEIGHT + Inches(0.3)
    
    # Add slide title using H3 typography
    slide2_title_left = MARGIN_STANDARD
    slide2_title_top = content_top
    slide2_title_width = prs.slide_width - Inches(2)
    slide2_title_height = Inches(0.45)
    
    slide2_title_box = slide2.shapes.add_textbox(slide2_title_left, slide2_title_top, slide2_title_width, slide2_title_height)
    slide2_title_frame = slide2_title_box.text_frame
    slide2_title_frame.word_wrap = True
    slide2_title_paragraph = slide2_title_frame.paragraphs[0]
    slide2_title_paragraph.text = "Executive Dashboard"
    slide2_title_paragraph.font.name = TITLE_FONT_NAME
    slide2_title_paragraph.font.size = H3_FONT_SIZE
    slide2_title_paragraph.font.bold = True
    slide2_title_paragraph.font.color.rgb = CS_NAVY
    slide2_title_paragraph.alignment = PP_ALIGN.LEFT
    
    # Calculate values for display
    total_millions = data.total_modeled / 1000000
    after_hours_percent = int((data.after_hours_escalations / data.incidents_escalated) * 100) if data.incidents_escalated > 0 else 0
    
    # Consolidated 6 key metrics in 2 rows x 3 columns
    dashboard_metrics = [
        # Row 1: Core Security Outcomes
        {
            "title": "THREATS CONTAINED",
            "metric": f"{data.true_threats_contained}",
            "detail": "100% contained, zero breaches"
        },
        {
            "title": "RESPONSE ADVANTAGE",
            "metric": f"{int(data.response_advantage_percent)}%",
            "detail": f"Faster than industry ({data.mttr_minutes}m vs {data.industry_median_minutes}m)"
        },
        {
            "title": "CLOSED E2E",
            "metric": f"{data.closed_end_to_end:,}",
            "detail": "Incidents resolved without client action"
        },
        # Row 2: Operations & Coverage
        {
            "title": "ALERTS TRIAGED",
            "metric": f"{data.alerts_triaged:,}",
            "detail": f"{data.client_touch_decisions:,} guided decisions"
        },
        {
            "title": "AFTER-HOURS",
            "metric": f"{data.after_hours_escalations}",
            "detail": f"{after_hours_percent}% of escalations"
        },
        {
            "title": "FALSE POSITIVE RATE",
            "metric": f"{data.false_positive_rate}%",
            "detail": f"{int(data.automation_percent)}% auto-routed"
        }
    ]
    
    # Card dimensions for 2x3 grid
    card_width = (prs.slide_width - Inches(1.4)) / 3
    card_height = Inches(1.4)
    card_start_left = Inches(0.5)
    card_start_top = content_top + Inches(0.5)
    card_spacing_h = Inches(0.2)
    card_spacing_v = Inches(0.15)
    
    # Draw 6 cards in 2x3 grid
    for idx, card in enumerate(dashboard_metrics):
        row = idx // 3
        col = idx % 3
        card_left = card_start_left + col * (card_width + card_spacing_h)
        card_top = card_start_top + row * (card_height + card_spacing_v)
        
        # Create card background
        card_shape = slide2.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, card_left, card_top,
            card_width, card_height
        )
        fill = card_shape.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(240, 248, 255) if row == 0 else RGBColor(248, 250, 252)
        line = card_shape.line
        line.color.rgb = CS_BLUE if row == 0 else CS_SLATE
        line.width = Pt(2) if row == 0 else Pt(1)
        
        # Add card title (small label)
        title_box = slide2.shapes.add_textbox(
            card_left + Inches(0.12), card_top + Inches(0.08),
            card_width - Inches(0.24), Inches(0.22)
        )
        title_frame = title_box.text_frame
        title_para = title_frame.paragraphs[0]
        title_para.text = card["title"]
        title_para.font.name = TITLE_FONT_NAME
        title_para.font.size = Pt(10)
        title_para.font.bold = True
        title_para.font.color.rgb = CS_SLATE
        title_para.alignment = PP_ALIGN.LEFT
        
        # Add large metric
        metric_box = slide2.shapes.add_textbox(
            card_left + Inches(0.12), card_top + Inches(0.28),
            card_width - Inches(0.24), Inches(0.6)
        )
        metric_frame = metric_box.text_frame
        metric_para = metric_frame.paragraphs[0]
        metric_para.text = card["metric"]
        metric_para.font.name = TITLE_FONT_NAME
        metric_para.font.size = Pt(42)
        metric_para.font.bold = True
        metric_para.font.color.rgb = CS_NAVY
        metric_para.alignment = PP_ALIGN.LEFT
        
        # Add detail text
        detail_box = slide2.shapes.add_textbox(
            card_left + Inches(0.12), card_top + Inches(0.95),
            card_width - Inches(0.24), Inches(0.4)
        )
        detail_frame = detail_box.text_frame
        detail_frame.word_wrap = True
        detail_para = detail_frame.paragraphs[0]
        detail_para.text = card["detail"]
        detail_para.font.name = BODY_FONT_NAME
        detail_para.font.size = Pt(11)
        detail_para.font.color.rgb = CS_SLATE
        detail_para.alignment = PP_ALIGN.LEFT
    
    # Add insight bar at bottom of dashboard
    insight_top = card_start_top + 2 * (card_height + card_spacing_v) + Inches(0.2)
    add_insight_callout(
        slide2, prs,
        "Your Security Posture at a Glance",
        "Every metric meeting or exceeding targets—your organization is well-protected.",
        insight_top,
        height=Inches(0.7)
    )
    
    # Slide 4 - CORR Platform Funnel (AI Accelerated Security)
    build_corr_funnel_slide(prs)


def create_executive_summary_slide(prs, report_data):
    """Create the Executive Summary slide."""
    pass


def build_corr_funnel_slide(prs):
    """Create the CORR Platform funnel slide showing security event flow.
    
    This slide visualizes the AI-accelerated security pipeline with 4 stages:
    Security Events → Potential Threats → Alerts → Response Actions
    
    The funnel uses overlapping rounded rectangles of decreasing height to show
    how the CORR platform filters and processes security events.
    
    Args:
        prs (Presentation): The presentation object.
    
    Returns:
        Slide: The created slide object.
    """
    # Create slide with standard branding
    slide, content_top = setup_content_slide(prs, "AI Accelerated, Human Validated Security")
    
    # Add subtitle
    subtitle_box = slide.shapes.add_textbox(
        MARGIN_STANDARD, content_top,
        prs.slide_width - Inches(2), Inches(0.4)
    )
    subtitle_frame = subtitle_box.text_frame
    subtitle_para = subtitle_frame.paragraphs[0]
    subtitle_para.text = "CORR Is Our Superpower to Finding the Right Alerts to Prevent Incidents"
    subtitle_para.font.name = BODY_FONT_NAME
    subtitle_para.font.size = Pt(14)
    subtitle_para.font.color.rgb = CS_SLATE
    subtitle_para.alignment = PP_ALIGN.LEFT
    
    # =========================================================================
    # Funnel Visualization Parameters
    # =========================================================================
    funnel_top = content_top + Inches(0.6)
    funnel_left = Inches(0.4)
    funnel_width = prs.slide_width - Inches(0.8)
    
    # Define the 4 funnel stages with decreasing heights
    # Heights decrease to create the funnel effect (tallest on left, shortest on right)
    funnel_stages = [
        {
            "value": "1.2B",
            "label": "Security Events",
            "color": CS_ORANGE,
            "height": Inches(2.8),
            "width": Inches(1.8),
            "has_shield": True,
            "shield_label": "Security Event\nIn-Flow CORR\nPlatform"
        },
        {
            "value": "2m",
            "label": "Potential Threats",
            "color": CS_BLUE,
            "height": Inches(2.3),
            "width": Inches(1.6),
            "agent_label": "TBR Agent",
            "agent_metric": "99% Resolution"
        },
        {
            "value": "150k",
            "label": "Alerts",
            "color": CS_VIOLET,
            "height": Inches(1.9),
            "width": Inches(1.4),
            "agent_label": "Case Agent",
            "agent_metric": "92% Consolidation"
        },
        {
            "value": "72k",
            "label": "Response\nActions",
            "color": CS_RED,
            "height": Inches(1.5),
            "width": Inches(1.2),
            "agent_label": "SOC / AI",
            "agent_metric": "Investigation"
        }
    ]
    
    # Calculate positions for overlapping effect
    # Each stage overlaps the previous one slightly
    overlap = Inches(0.15)
    total_width = sum(stage["width"] for stage in funnel_stages) - overlap * (len(funnel_stages) - 1)
    start_left = funnel_left + (funnel_width - total_width - Inches(1.0)) / 2  # Center with room for arrow
    
    # Maximum height for vertical alignment (align bottoms)
    max_height = max(stage["height"] for stage in funnel_stages)
    funnel_bottom = funnel_top + max_height
    
    # Draw gray background arrow (the flow indicator)
    arrow_left = start_left - Inches(0.3)
    arrow_width = total_width + Inches(1.5)
    arrow_height = Inches(1.8)
    arrow_top = funnel_bottom - arrow_height / 2 - Inches(0.3)
    
    # Create arrow shape using pentagon/chevron
    arrow_shape = slide.shapes.add_shape(
        MSO_SHAPE.CHEVRON, 
        start_left + total_width + Inches(0.2),
        funnel_bottom - Inches(0.9),
        Inches(0.8), Inches(0.8)
    )
    arrow_shape.fill.solid()
    arrow_shape.fill.fore_color.rgb = RGBColor(180, 180, 180)  # Gray arrow
    arrow_shape.line.fill.background()
    
    # Draw each funnel stage
    current_left = start_left
    
    for idx, stage in enumerate(funnel_stages):
        stage_height = stage["height"]
        stage_width = stage["width"]
        stage_top = funnel_bottom - stage_height  # Align bottoms
        
        # Draw agent label above the stage (if present)
        if "agent_label" in stage:
            label_top = stage_top - Inches(0.55)
            
            # Agent label (bold)
            agent_label_box = slide.shapes.add_textbox(
                current_left, label_top,
                stage_width, Inches(0.25)
            )
            agent_tf = agent_label_box.text_frame
            agent_para = agent_tf.paragraphs[0]
            agent_para.text = stage["agent_label"]
            agent_para.font.name = TITLE_FONT_NAME
            agent_para.font.size = Pt(11)
            agent_para.font.bold = True
            agent_para.font.color.rgb = CS_NAVY
            agent_para.alignment = PP_ALIGN.CENTER
            
            # Agent metric (smaller, below)
            metric_label_box = slide.shapes.add_textbox(
                current_left, label_top + Inches(0.2),
                stage_width, Inches(0.2)
            )
            metric_tf = metric_label_box.text_frame
            metric_para = metric_tf.paragraphs[0]
            metric_para.text = stage["agent_metric"]
            metric_para.font.name = BODY_FONT_NAME
            metric_para.font.size = Pt(9)
            metric_para.font.color.rgb = CS_SLATE
            metric_para.alignment = PP_ALIGN.CENTER
        
        # Draw the rounded rectangle for this stage
        rect_shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            current_left, stage_top,
            stage_width, stage_height
        )
        rect_shape.fill.solid()
        rect_shape.fill.fore_color.rgb = stage["color"]
        rect_shape.line.fill.background()
        
        # Adjust corner radius (adjustments property)
        if hasattr(rect_shape, 'adjustments') and len(rect_shape.adjustments) > 0:
            rect_shape.adjustments[0] = 0.15  # Rounded corners
        
        # Add shield icon and label for first stage
        if stage.get("has_shield"):
            # Shield icon (using text-based shield character)
            shield_box = slide.shapes.add_textbox(
                current_left + Inches(0.3), stage_top + Inches(0.3),
                stage_width - Inches(0.6), Inches(0.6)
            )
            shield_tf = shield_box.text_frame
            shield_para = shield_tf.paragraphs[0]
            shield_para.text = "🛡"
            shield_para.font.size = Pt(32)
            shield_para.alignment = PP_ALIGN.CENTER
            
            # Shield label text
            shield_label_box = slide.shapes.add_textbox(
                current_left + Inches(0.1), stage_top + Inches(0.85),
                stage_width - Inches(0.2), Inches(0.8)
            )
            shield_label_tf = shield_label_box.text_frame
            shield_label_tf.word_wrap = True
            shield_label_para = shield_label_tf.paragraphs[0]
            shield_label_para.text = stage["shield_label"]
            shield_label_para.font.name = TITLE_FONT_NAME
            shield_label_para.font.size = Pt(12)
            shield_label_para.font.bold = True
            shield_label_para.font.color.rgb = RGBColor(255, 255, 255)
            shield_label_para.alignment = PP_ALIGN.CENTER
        
        # Add value (large number) - positioned in center of shape
        value_top = stage_top + stage_height * 0.35
        if stage.get("has_shield"):
            value_top = stage_top + stage_height * 0.55
            
        value_box = slide.shapes.add_textbox(
            current_left, value_top,
            stage_width, Inches(0.6)
        )
        value_tf = value_box.text_frame
        value_para = value_tf.paragraphs[0]
        value_para.text = stage["value"]
        value_para.font.name = TITLE_FONT_NAME
        value_para.font.size = Pt(36)
        value_para.font.bold = True
        value_para.font.color.rgb = RGBColor(255, 255, 255)
        value_para.alignment = PP_ALIGN.CENTER
        
        # Add label below value
        label_box = slide.shapes.add_textbox(
            current_left, value_top + Inches(0.45),
            stage_width, Inches(0.5)
        )
        label_tf = label_box.text_frame
        label_tf.word_wrap = True
        label_para = label_tf.paragraphs[0]
        label_para.text = stage["label"]
        label_para.font.name = BODY_FONT_NAME
        label_para.font.size = Pt(11)
        label_para.font.bold = True
        label_para.font.color.rgb = RGBColor(255, 255, 255)
        label_para.alignment = PP_ALIGN.CENTER
        
        # Move to next position with overlap
        current_left += stage_width - overlap
    
    return slide


def build_value_delivered_slides(prs, data):
    """Create the Value Delivered section (consolidated to 2 slides).
    
    Per CRITICALSTART branding guidelines:
    - All slides have header and footer with transparent background
    - Uses H1-H6 typography scale
    
    Args:
        prs (Presentation): The presentation object.
        data (ReportData): The report data object containing all metrics.
    """
    blank_slide_layout = prs.slide_layouts[6]  # Blank layout
    content_top = HEADER_HEIGHT + Inches(0.3)
    
    # Slide 3 - Value Delivered (Merged: Hero + Breakdown)
    slide_value = prs.slides.add_slide(blank_slide_layout)
    slide_value_number = get_slide_number(prs)
    
    # Add master slide elements
    add_master_slide_elements(slide_value, prs, slide_number=slide_value_number,
                               include_header=True, include_footer=True)
    
    # Add logo at top right
    add_logo(slide_value, position='top_right', prs=prs)
    
    # Add title using H3 typography
    title_left = MARGIN_STANDARD
    title_top = content_top
    title_width = prs.slide_width - Inches(2)
    title_height = Inches(0.5)
    
    title_box = slide_value.shapes.add_textbox(title_left, title_top, title_width, title_height)
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    title_paragraph = title_frame.paragraphs[0]
    title_paragraph.text = "Value Delivered"
    title_paragraph.font.name = TITLE_FONT_NAME
    title_paragraph.font.size = H3_FONT_SIZE
    title_paragraph.font.bold = True
    title_paragraph.font.color.rgb = CS_NAVY
    title_paragraph.alignment = PP_ALIGN.LEFT
    
    # Hero section (left side) - Cost Avoidance
    total_millions = data.total_modeled / 1000000
    hero_value = f"~${total_millions:.1f}M"
    
    hero_left = Inches(0.5)
    hero_top = content_top + Inches(0.55)
    hero_width = Inches(3.5)
    hero_height = Inches(2.8)
    
    # Hero background
    hero_bg = slide_value.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, hero_left, hero_top,
        hero_width, hero_height
    )
    fill = hero_bg.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(240, 248, 255)
    line = hero_bg.line
    line.color.rgb = CS_BLUE
    line.width = Pt(2)
    
    # Hero value
    hero_box = slide_value.shapes.add_textbox(hero_left + Inches(0.2), hero_top + Inches(0.3), 
                                               hero_width - Inches(0.4), Inches(1.2))
    hero_frame = hero_box.text_frame
    hero_paragraph = hero_frame.paragraphs[0]
    hero_paragraph.text = hero_value
    hero_paragraph.font.name = TITLE_FONT_NAME
    hero_paragraph.font.size = Pt(72)
    hero_paragraph.font.bold = True
    hero_paragraph.font.color.rgb = CS_NAVY
    hero_paragraph.alignment = PP_ALIGN.CENTER
    
    # Hero label
    hero_label = slide_value.shapes.add_textbox(hero_left + Inches(0.2), hero_top + Inches(1.5), 
                                                 hero_width - Inches(0.4), Inches(0.4))
    label_frame = hero_label.text_frame
    label_para = label_frame.paragraphs[0]
    label_para.text = "Cost Exposure Avoided"
    label_para.font.name = TITLE_FONT_NAME
    label_para.font.size = Pt(18)
    label_para.font.bold = True
    label_para.font.color.rgb = CS_NAVY
    label_para.alignment = PP_ALIGN.CENTER
    
    # Hero subtitle
    hero_sub = slide_value.shapes.add_textbox(hero_left + Inches(0.2), hero_top + Inches(2.0), 
                                               hero_width - Inches(0.4), Inches(0.6))
    sub_frame = hero_sub.text_frame
    sub_frame.word_wrap = True
    sub_para = sub_frame.paragraphs[0]
    sub_para.text = "Modeled operational, coverage, and breach avoidance"
    sub_para.font.name = BODY_FONT_NAME
    sub_para.font.size = Pt(12)
    sub_para.font.color.rgb = CS_SLATE
    sub_para.alignment = PP_ALIGN.CENTER
    
    # Right side - 3 breakdown cards (stacked vertically)
    card_left = Inches(4.2)
    card_width = prs.slide_width - card_left - Inches(0.5)
    card_height = Inches(0.85)
    card_spacing = Inches(0.1)
    card_start_top = hero_top
    
    # Format values
    analyst_k = data.analyst_cost_equivalent / 1000
    coverage_k = data.coverage_cost_equivalent / 1000
    breach_m = data.breach_exposure_avoided / 1000000
    
    breakdown_cards = [
        {"title": "Security Operations", "metric": f"~${analyst_k:.0f}K", "detail": f"{data.analyst_hours} analyst hours"},
        {"title": "24/7 Coverage", "metric": f"~${coverage_k:.0f}K", "detail": f"{data.coverage_hours} hours monitoring"},
        {"title": "Threat Prevention", "metric": f"~${breach_m:.1f}M", "detail": f"{data.true_threats_contained} threats contained"}
    ]
    
    for i, card in enumerate(breakdown_cards):
        card_top = card_start_top + i * (card_height + card_spacing)
        
        # Card background
        card_shape = slide_value.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, card_left, card_top,
            card_width, card_height
        )
        fill = card_shape.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(248, 250, 252)
        line = card_shape.line
        line.color.rgb = CS_SLATE
        line.width = Pt(1)
        
        # Card title (left)
        title_box = slide_value.shapes.add_textbox(
            card_left + Inches(0.15), card_top + Inches(0.12),
            Inches(2.2), Inches(0.3)
        )
        t_frame = title_box.text_frame
        t_para = t_frame.paragraphs[0]
        t_para.text = card["title"]
        t_para.font.name = TITLE_FONT_NAME
        t_para.font.size = Pt(14)
        t_para.font.bold = True
        t_para.font.color.rgb = CS_SLATE
        t_para.alignment = PP_ALIGN.LEFT
        
        # Card detail (left, below title)
        detail_box = slide_value.shapes.add_textbox(
            card_left + Inches(0.15), card_top + Inches(0.45),
            Inches(2.5), Inches(0.3)
        )
        d_frame = detail_box.text_frame
        d_para = d_frame.paragraphs[0]
        d_para.text = card["detail"]
        d_para.font.name = BODY_FONT_NAME
        d_para.font.size = Pt(11)
        d_para.font.color.rgb = CS_SLATE
        d_para.alignment = PP_ALIGN.LEFT
        
        # Card metric (right)
        metric_box = slide_value.shapes.add_textbox(
            card_left + card_width - Inches(1.8), card_top + Inches(0.15),
            Inches(1.6), Inches(0.55)
        )
        m_frame = metric_box.text_frame
        m_para = m_frame.paragraphs[0]
        m_para.text = card["metric"]
        m_para.font.name = TITLE_FONT_NAME
        m_para.font.size = Pt(28)
        m_para.font.bold = True
        m_para.font.color.rgb = CS_NAVY
        m_para.alignment = PP_ALIGN.RIGHT
    
    # Methodology note at bottom
    note_top = hero_top + hero_height + Inches(0.15)
    note_box = slide_value.shapes.add_textbox(Inches(0.5), note_top, 
                                               prs.slide_width - Inches(1), Inches(0.3))
    note_frame = note_box.text_frame
    note_para = note_frame.paragraphs[0]
    note_para.text = "Illustrative impact only; not redeployable budget. Based on industry-standard rates."
    note_para.font.name = BODY_FONT_NAME
    note_para.font.size = Pt(10)
    note_para.font.color.rgb = CS_SLATE
    note_para.font.italic = True
    note_para.alignment = PP_ALIGN.CENTER


def build_protection_achieved_slides(prs, data):
    """Create the Protection section (Slides 7-10).
    
    Args:
        prs (Presentation): The presentation object.
        data (ReportData): The report data object containing all metrics.
    """
    # Slide 7 - Critical Start's Performance (Split Layout)
    slide7, content_top_7 = setup_content_slide(prs, "Critical Start's Performance")
    
    # =========================================================================
    # LEFT PANEL: Response Metrics (MTTR by Severity)
    # =========================================================================
    panel_left = Inches(0.5)
    panel_top = content_top_7 + Inches(0.1)
    panel_width = Inches(2.3)
    panel_height = Inches(3.2)
    
    # Panel background
    panel_bg = slide7.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, panel_left, panel_top,
        panel_width, panel_height
    )
    panel_bg.fill.solid()
    panel_bg.fill.fore_color.rgb = RGBColor(248, 250, 252)  # Very light gray
    panel_bg.line.color.rgb = CS_SLATE
    panel_bg.line.width = Pt(1)
    
    # Panel header
    header_height = Inches(0.4)
    header_shape = slide7.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, panel_left, panel_top,
        panel_width, header_height
    )
    header_shape.fill.solid()
    header_shape.fill.fore_color.rgb = CS_NAVY
    header_shape.line.fill.background()
    
    header_text = slide7.shapes.add_textbox(panel_left, panel_top, panel_width, header_height)
    header_tf = header_text.text_frame
    header_tf.paragraphs[0].text = "RESPONSE METRICS"
    header_tf.paragraphs[0].font.name = TITLE_FONT_NAME
    header_tf.paragraphs[0].font.size = Pt(12)
    header_tf.paragraphs[0].font.bold = True
    header_tf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    header_tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # Metric cards within panel
    metrics = [
        {"value": f"{data.critical_high_mttr} min", "label": "Critical/High", "color": CS_RED},
        {"value": f"{data.medium_low_mttr} min", "label": "Medium/Low", "color": CS_ORANGE},
        {"value": f"{data.p90_minutes} min", "label": "P90 Response", "color": CS_BLUE},
    ]
    
    card_top = panel_top + header_height + Inches(0.15)
    card_height = Inches(0.85)
    card_spacing = Inches(0.1)
    card_margin = Inches(0.1)
    card_width = panel_width - card_margin * 2
    
    for idx, metric in enumerate(metrics):
        current_top = card_top + idx * (card_height + card_spacing)
        
        # Card background
        card_shape = slide7.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            panel_left + card_margin, current_top,
            card_width, card_height
        )
        card_shape.fill.solid()
        card_shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
        card_shape.line.color.rgb = metric["color"]
        card_shape.line.width = Pt(2)
        
        # Metric value (large)
        value_box = slide7.shapes.add_textbox(
            panel_left + card_margin, current_top + Inches(0.1),
            card_width, Inches(0.45)
        )
        value_tf = value_box.text_frame
        value_tf.paragraphs[0].text = metric["value"]
        value_tf.paragraphs[0].font.name = TITLE_FONT_NAME
        value_tf.paragraphs[0].font.size = Pt(28)
        value_tf.paragraphs[0].font.bold = True
        value_tf.paragraphs[0].font.color.rgb = metric["color"]
        value_tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # Metric label
        label_box = slide7.shapes.add_textbox(
            panel_left + card_margin, current_top + Inches(0.5),
            card_width, Inches(0.3)
        )
        label_tf = label_box.text_frame
        label_tf.paragraphs[0].text = metric["label"]
        label_tf.paragraphs[0].font.name = BODY_FONT_NAME
        label_tf.paragraphs[0].font.size = Pt(11)
        label_tf.paragraphs[0].font.color.rgb = CS_SLATE
        label_tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # =========================================================================
    # RIGHT PANEL: Trend Chart
    # =========================================================================
    chart_left = panel_left + panel_width + Inches(0.2)
    chart_top = content_top_7 + Inches(0.1)
    chart_width = prs.slide_width - chart_left - Inches(0.5)
    chart_height = Inches(2.3)
    
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
    legend_top = chart_top + chart_height + Inches(0.15)
    legend_left = chart_left
    legend_width = chart_width
    legend_height = Inches(0.35)
    
    legend_box = slide7.shapes.add_textbox(legend_left, legend_top, legend_width, legend_height)
    legend_frame = legend_box.text_frame
    legend_frame.word_wrap = True
    legend_paragraph = legend_frame.paragraphs[0]
    legend_paragraph.text = "MTTR (blue) | MTTD (navy) | FP% (red dashed)"
    legend_paragraph.font.name = BODY_FONT_NAME
    legend_paragraph.font.size = Pt(11)
    legend_paragraph.font.color.rgb = CS_SLATE
    legend_paragraph.alignment = PP_ALIGN.CENTER
    
    # Add insight box below legend (spanning full width of right panel)
    insight_top = legend_top + legend_height + Inches(0.1)
    insight_height = Inches(0.65)
    
    insight_shape = slide7.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, chart_left, insight_top,
        chart_width, insight_height
    )
    fill = insight_shape.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(240, 248, 255)  # Light blue background
    line = insight_shape.line
    line.color.rgb = CS_BLUE
    line.width = Pt(2)
    
    insight_text = insight_shape.text_frame
    insight_text.text = "MTTR decreased 25% to 126 min | MTTD improved 22% to 42 min"
    insight_text.paragraphs[0].font.name = BODY_FONT_NAME
    insight_text.paragraphs[0].font.size = Pt(13)
    insight_text.paragraphs[0].font.color.rgb = CS_NAVY
    insight_text.paragraphs[0].font.bold = True
    insight_text.paragraphs[0].alignment = PP_ALIGN.CENTER
    insight_text.vertical_anchor = 1  # Middle
    
    # Slide 8 - Industry Comparison
    slide8, content_top_8 = setup_content_slide(prs, "Industry Comparison")
    
    # Create table for comparison
    table_top = content_top_8 + Inches(0.1)
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
    
    # Add insight bar for Industry Comparison
    industry_insight_top = table_top + table_height + Inches(0.3)
    add_insight_callout(
        slide8, prs,
        "Outperforming Industry Benchmarks",
        "Faster response, better detection, and lower risk—exceeding standards across all metrics.",
        industry_insight_top,
        height=Inches(0.7)
    )
    
    # Slide 7 - Response & Detection Quality (Merged dual-panel)
    slide_resp_det, content_top_rd = setup_content_slide(prs, "Response & Detection Quality")
    
    # Left panel: Response Efficiency (3 vertical cards)
    left_panel_left = Inches(0.5)
    left_panel_width = (prs.slide_width - Inches(1.2)) / 2
    card_height = Inches(0.9)
    card_spacing = Inches(0.1)
    cards_start_top = content_top_rd + Inches(0.1)
    
    # Calculate containment count (round to nearest integer)
    containment_count = round((data.containment_rate / 100) * data.incidents_escalated)
    
    efficiency_cards = [
        {"title": "Containment Rate", "value": f"{data.containment_rate}%", "detail": f"{containment_count} of {data.incidents_escalated}", "color": CS_BLUE},
        {"title": "Playbook Automation", "value": f"{data.playbook_auto['percent']}%", "detail": f"{data.playbook_auto['count']} incidents", "color": CS_BLUE},
        {"title": "Analyst Escalation", "value": f"{data.analyst_escalation['percent']}%", "detail": f"{data.analyst_escalation['count']} incidents", "color": CS_ORANGE}
    ]
    
    for i, card in enumerate(efficiency_cards):
        card_top = cards_start_top + i * (card_height + card_spacing)
        
        # Create card background
        card_shape = slide_resp_det.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, left_panel_left, card_top,
            left_panel_width, card_height
        )
        fill = card_shape.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(240, 248, 255)
        line = card_shape.line
        line.color.rgb = card["color"]
        line.width = Pt(2)
        
        # Card title (left)
        title_box = slide_resp_det.shapes.add_textbox(
            left_panel_left + Inches(0.15), card_top + Inches(0.15),
            Inches(2.2), Inches(0.3)
        )
        t_frame = title_box.text_frame
        t_para = t_frame.paragraphs[0]
        t_para.text = card["title"]
        t_para.font.name = TITLE_FONT_NAME
        t_para.font.size = Pt(14)
        t_para.font.bold = True
        t_para.font.color.rgb = CS_NAVY
        t_para.alignment = PP_ALIGN.LEFT
        
        # Detail (left, below title)
        detail_box = slide_resp_det.shapes.add_textbox(
            left_panel_left + Inches(0.15), card_top + Inches(0.48),
            Inches(2.2), Inches(0.3)
        )
        d_frame = detail_box.text_frame
        d_para = d_frame.paragraphs[0]
        d_para.text = card["detail"]
        d_para.font.name = BODY_FONT_NAME
        d_para.font.size = Pt(11)
        d_para.font.color.rgb = CS_SLATE
        d_para.alignment = PP_ALIGN.LEFT
        
        # Value (right)
        value_box = slide_resp_det.shapes.add_textbox(
            left_panel_left + left_panel_width - Inches(1.5), card_top + Inches(0.15),
            Inches(1.3), Inches(0.6)
        )
        v_frame = value_box.text_frame
        v_para = v_frame.paragraphs[0]
        v_para.text = card["value"]
        v_para.font.name = TITLE_FONT_NAME
        v_para.font.size = Pt(32)
        v_para.font.bold = True
        v_para.font.color.rgb = card["color"]
        v_para.alignment = PP_ALIGN.RIGHT
    
    # Right panel: Detection Quality (2x2 grid)
    right_panel_left = left_panel_left + left_panel_width + Inches(0.2)
    right_panel_width = left_panel_width
    grid_card_width = (right_panel_width - Inches(0.1)) / 2
    grid_card_height = Inches(1.35)
    grid_spacing = Inches(0.1)
    
    quality_cards = [
        {"title": "True Threat", "value": f"{data.true_threat_precision}%", "color": CS_RED},
        {"title": "Signal Fidelity", "value": f"{data.signal_fidelity}%", "color": CS_BLUE},
        {"title": "False Positive", "value": f"{data.false_positive_rate}%", "color": CS_ORANGE},
        {"title": "Client-Validated", "value": f"{data.client_validated}%", "color": CS_BLUE}
    ]
    
    for i, card in enumerate(quality_cards):
        row = i // 2
        col = i % 2
        card_left = right_panel_left + col * (grid_card_width + grid_spacing)
        card_top = cards_start_top + row * (grid_card_height + grid_spacing)
        
        # Create card background
        card_shape = slide_resp_det.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, card_left, card_top,
            grid_card_width, grid_card_height
        )
        fill = card_shape.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(248, 250, 252)
        line = card_shape.line
        line.color.rgb = card["color"]
        line.width = Pt(2)
        
        # Title
        title_box = slide_resp_det.shapes.add_textbox(
            card_left + Inches(0.1), card_top + Inches(0.12),
            grid_card_width - Inches(0.2), Inches(0.3)
        )
        t_frame = title_box.text_frame
        t_para = t_frame.paragraphs[0]
        t_para.text = card["title"]
        t_para.font.name = TITLE_FONT_NAME
        t_para.font.size = Pt(12)
        t_para.font.bold = True
        t_para.font.color.rgb = CS_SLATE
        t_para.alignment = PP_ALIGN.LEFT
        
        # Value
        value_box = slide_resp_det.shapes.add_textbox(
            card_left + Inches(0.1), card_top + Inches(0.5),
            grid_card_width - Inches(0.2), Inches(0.7)
        )
        v_frame = value_box.text_frame
        v_para = v_frame.paragraphs[0]
        v_para.text = card["value"]
        v_para.font.name = TITLE_FONT_NAME
        v_para.font.size = Pt(36)
        v_para.font.bold = True
        v_para.font.color.rgb = card["color"]
        v_para.alignment = PP_ALIGN.LEFT
    
    # Insight box at bottom
    insight_top = cards_start_top + 3 * (card_height + card_spacing) + Inches(0.1)
    insight_shape = slide_resp_det.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.5), insight_top,
        prs.slide_width - Inches(1), Inches(0.6)
    )
    fill = insight_shape.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(240, 248, 255)
    line = insight_shape.line
    line.color.rgb = CS_BLUE
    line.width = Pt(2)
    
    insight_text = insight_shape.text_frame
    insight_text.text = f"Signal quality improved: FP rate at {data.false_positive_rate}% with {data.containment_rate}% containment"
    insight_text.paragraphs[0].font.name = BODY_FONT_NAME
    insight_text.paragraphs[0].font.size = Pt(14)
    insight_text.paragraphs[0].font.color.rgb = CS_NAVY
    insight_text.paragraphs[0].font.bold = True
    insight_text.paragraphs[0].alignment = PP_ALIGN.CENTER
    insight_text.vertical_anchor = 1


def build_threat_landscape_slides(prs, data, include=True):
    """Create the optional Threat Landscape section (Slides 11-13).
    
    Per CRITICALSTART branding guidelines:
    - All slides have transparent header and footer
    - Uses H1-H6 typography scale
    
    Args:
        prs (Presentation): The presentation object.
        data (ReportData): The report data object containing all metrics.
        include (bool): Whether to include this section. Default True.
    """
    if not include:
        return
    
    # Note: setup_content_slide is used for each slide now
    
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
    slide11, content_top_11 = setup_content_slide(prs, "Severity Alignment Flow")
    
    # Add subtitle using H6 typography
    subtitle_left = MARGIN_STANDARD
    subtitle_top = content_top_11 - Inches(0.6)
    subtitle_width = prs.slide_width - Inches(2.5)
    subtitle_height = Inches(0.4)
    
    subtitle_box = slide11.shapes.add_textbox(subtitle_left, subtitle_top, subtitle_width, subtitle_height)
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.word_wrap = True
    subtitle_paragraph = subtitle_frame.paragraphs[0]
    subtitle_paragraph.text = "Vendor-reported criticality vs. Critical Start adjudication"
    subtitle_paragraph.font.name = BODY_FONT_NAME
    subtitle_paragraph.font.size = H6_FONT_SIZE  # 12pt
    subtitle_paragraph.font.color.rgb = CS_SLATE
    subtitle_paragraph.alignment = PP_ALIGN.LEFT
    
    # Add chip with total escalations
    chip_left = prs.slide_width - Inches(2.5)
    chip_top = content_top_11 - Inches(0.6)
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
    
    # Slide 9 - Threat & Detection Sources (Merged: MITRE Tactics + Detection Sources)
    slide_threats, content_top_threats = setup_content_slide(prs, "Threat & Detection Sources")
    
    # Left half: MITRE ATT&CK chart placeholder (compact)
    chart_left = Inches(0.5)
    chart_top = content_top_threats + Inches(0.1)
    chart_width = (prs.slide_width - Inches(1.2)) / 2
    chart_height = Inches(2.3)
    
    chart_placeholder = slide_threats.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, chart_left, chart_top,
        chart_width, chart_height
    )
    fill = chart_placeholder.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    line = chart_placeholder.line
    line.color.rgb = CS_SLATE
    line.width = Pt(1)
    
    placeholder_text = chart_placeholder.text_frame
    placeholder_text.text = "[Chart: MITRE ATT&CK Stacked Bar - ID: mitre_stacked_bar]"
    placeholder_text.paragraphs[0].font.name = BODY_FONT_NAME
    placeholder_text.paragraphs[0].font.size = Pt(12)
    placeholder_text.paragraphs[0].font.color.rgb = CS_SLATE
    placeholder_text.paragraphs[0].alignment = PP_ALIGN.CENTER
    placeholder_text.vertical_anchor = 1
    
    # Chart legend below
    legend_top = chart_top + chart_height + Inches(0.1)
    legend_box = slide_threats.shapes.add_textbox(chart_left, legend_top, chart_width, Inches(0.3))
    legend_frame = legend_box.text_frame
    legend_para = legend_frame.paragraphs[0]
    legend_para.text = "High | Medium | Low | Info"
    legend_para.font.name = BODY_FONT_NAME
    legend_para.font.size = Pt(10)
    legend_para.font.color.rgb = CS_SLATE
    legend_para.alignment = PP_ALIGN.CENTER
    
    # MITRE insight below legend
    if len(data.tactics) > 0 and len(data.high_severity) > 0:
        persistence_total = (data.high_severity[0] + data.medium_severity[0] + 
                           data.low_severity[0] + data.info_severity[0])
        persistence_high = data.high_severity[0]
        mitre_insight = f"Persistence: {persistence_total} escalations, {persistence_high} high-severity"
    else:
        mitre_insight = "Persistence: 77 escalations, 12 high-severity"
    
    insight_top = legend_top + Inches(0.35)
    insight_box = slide_threats.shapes.add_textbox(chart_left, insight_top, chart_width, Inches(0.4))
    insight_frame = insight_box.text_frame
    insight_para = insight_frame.paragraphs[0]
    insight_para.text = mitre_insight
    insight_para.font.name = BODY_FONT_NAME
    insight_para.font.size = Pt(11)
    insight_para.font.bold = True
    insight_para.font.color.rgb = CS_NAVY
    insight_para.alignment = PP_ALIGN.CENTER
    
    # Right half: Detection Sources (3 compact cards stacked)
    right_left = chart_left + chart_width + Inches(0.2)
    right_width = chart_width
    source_card_height = Inches(1.0)
    source_spacing = Inches(0.1)
    fp_threshold = 10.0
    
    for i, source in enumerate(data.detection_sources[:3]):  # Limit to 3
        card_top = content_top_threats + Inches(0.1) + i * (source_card_height + source_spacing)
        fp_rate = source.get('fp_rate', 0)
        fp_color = CS_ORANGE if fp_rate > fp_threshold else CS_BLUE
        
        # Card background
        card_shape = slide_threats.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, right_left, card_top,
            right_width, source_card_height
        )
        fill = card_shape.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(248, 250, 252)
        line = card_shape.line
        line.color.rgb = fp_color
        line.width = Pt(2)
        
        # Source name (left)
        name_box = slide_threats.shapes.add_textbox(
            right_left + Inches(0.12), card_top + Inches(0.12),
            Inches(2.5), Inches(0.35)
        )
        n_frame = name_box.text_frame
        n_para = n_frame.paragraphs[0]
        n_para.text = source['source']
        n_para.font.name = TITLE_FONT_NAME
        n_para.font.size = Pt(13)
        n_para.font.bold = True
        n_para.font.color.rgb = CS_NAVY
        n_para.alignment = PP_ALIGN.LEFT
        
        # Incidents (left, below name)
        inc_box = slide_threats.shapes.add_textbox(
            right_left + Inches(0.12), card_top + Inches(0.5),
            Inches(2.5), Inches(0.35)
        )
        i_frame = inc_box.text_frame
        i_para = i_frame.paragraphs[0]
        i_para.text = f"{source['incidents']} incidents ({source['percent']}%)"
        i_para.font.name = BODY_FONT_NAME
        i_para.font.size = Pt(11)
        i_para.font.color.rgb = CS_SLATE
        i_para.alignment = PP_ALIGN.LEFT
        
        # FP Rate (right side)
        fp_box = slide_threats.shapes.add_textbox(
            right_left + right_width - Inches(1.3), card_top + Inches(0.25),
            Inches(1.2), Inches(0.5)
        )
        f_frame = fp_box.text_frame
        f_para = f_frame.paragraphs[0]
        f_para.text = f"FP: {fp_rate}%"
        f_para.font.name = TITLE_FONT_NAME
        f_para.font.size = Pt(18)
        f_para.font.bold = True
        f_para.font.color.rgb = fp_color
        f_para.alignment = PP_ALIGN.RIGHT
    
    # Summary insight at bottom
    summary_top = content_top_threats + Inches(0.1) + 3 * (source_card_height + source_spacing) + Inches(0.15)
    summary_shape = slide_threats.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.5), summary_top,
        prs.slide_width - Inches(1), Inches(0.5)
    )
    fill = summary_shape.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(240, 248, 255)
    line = summary_shape.line
    line.color.rgb = CS_BLUE
    line.width = Pt(2)
    
    summary_text = summary_shape.text_frame
    # Find source with highest FP rate
    high_fp_source = max(data.detection_sources, key=lambda x: x.get('fp_rate', 0))
    summary_text.text = f"Tuning opportunity: {high_fp_source['source']} at {high_fp_source['fp_rate']}% FP rate"
    summary_text.paragraphs[0].font.name = BODY_FONT_NAME
    summary_text.paragraphs[0].font.size = Pt(13)
    summary_text.paragraphs[0].font.color.rgb = CS_NAVY
    summary_text.paragraphs[0].font.bold = True
    summary_text.paragraphs[0].alignment = PP_ALIGN.CENTER
    summary_text.vertical_anchor = 1


def build_insights_slides(prs, data):
    """Create the Insights section (Slides 14-15).
    
    Per CRITICALSTART branding guidelines:
    - All slides have transparent header and footer
    - Uses H1-H6 typography scale
    
    Args:
        prs (Presentation): The presentation object.
        data (ReportData): The report data object containing all metrics.
    """
    # Slide 14 - Prioritized Improvement Plan
    slide14, content_top_14 = setup_content_slide(prs, "Prioritized Improvement Plan")
    
    # Create improvement items cards
    card_start_top = content_top_14 + Inches(0.1)
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
    
    # Add insight bar for Prioritized Improvements
    improvements_insight_top = card_start_top + 3 * (card_height + card_spacing) - Inches(0.1)
    add_insight_callout(
        slide14, prs,
        "Focus Areas for Continuous Improvement",
        "Three targeted actions to reduce noise, improve detection, and strengthen your security posture.",
        improvements_insight_top,
        height=Inches(0.7)
    )
    
    # Slide 11 - Operational Coverage (Merged: After-Hours + Operational Insights)
    slide_ops, content_top_ops = setup_content_slide(prs, "Operational Coverage")
    
    # Left panel: After-Hours metrics (hero + breakdown)
    left_width = (prs.slide_width - Inches(1.2)) / 2
    left_left = Inches(0.5)
    panel_top = content_top_ops + Inches(0.1)
    panel_height = Inches(3.0)
    
    # Left panel background
    left_panel = slide_ops.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left_left, panel_top,
        left_width, panel_height
    )
    fill = left_panel.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(240, 248, 255)
    line = left_panel.line
    line.color.rgb = CS_BLUE
    line.width = Pt(2)
    
    # After-Hours Hero number
    hero_box = slide_ops.shapes.add_textbox(left_left + Inches(0.2), panel_top + Inches(0.2), 
                                             Inches(2.5), Inches(1.0))
    hero_frame = hero_box.text_frame
    hero_para = hero_frame.paragraphs[0]
    hero_para.text = str(data.after_hours_escalations)
    hero_para.font.name = TITLE_FONT_NAME
    hero_para.font.size = Pt(64)
    hero_para.font.bold = True
    hero_para.font.color.rgb = CS_NAVY
    hero_para.alignment = PP_ALIGN.LEFT
    
    # After-Hours label
    label_box = slide_ops.shapes.add_textbox(left_left + Inches(0.2), panel_top + Inches(1.15), 
                                              Inches(3.0), Inches(0.35))
    label_frame = label_box.text_frame
    label_para = label_frame.paragraphs[0]
    label_para.text = "After-Hours Escalations"
    label_para.font.name = TITLE_FONT_NAME
    label_para.font.size = Pt(16)
    label_para.font.bold = True
    label_para.font.color.rgb = CS_NAVY
    label_para.alignment = PP_ALIGN.LEFT
    
    # Breakdown stats
    weeknight_count = getattr(data, 'after_hours_weeknight', int(data.after_hours_escalations * 0.82))
    weekend_count = getattr(data, 'after_hours_weekend', int(data.after_hours_escalations * 0.18))
    
    stats_box = slide_ops.shapes.add_textbox(left_left + Inches(0.2), panel_top + Inches(1.6), 
                                              left_width - Inches(0.4), Inches(1.2))
    stats_frame = stats_box.text_frame
    stats_frame.word_wrap = True
    stats_para = stats_frame.paragraphs[0]
    stats_para.text = f"{weeknight_count} weeknights · {weekend_count} weekends"
    stats_para.font.name = BODY_FONT_NAME
    stats_para.font.size = Pt(14)
    stats_para.font.color.rgb = CS_SLATE
    stats_para.alignment = PP_ALIGN.LEFT
    
    stats_para2 = stats_frame.add_paragraph()
    stats_para2.text = f"{data.coverage_hours} hours coverage · {int(data.automation_percent)}% auto-routed"
    stats_para2.font.name = BODY_FONT_NAME
    stats_para2.font.size = Pt(13)
    stats_para2.font.color.rgb = CS_SLATE
    stats_para2.alignment = PP_ALIGN.LEFT
    
    # Right panel: Collaboration metrics
    right_left = left_left + left_width + Inches(0.2)
    right_width = left_width
    
    right_panel = slide_ops.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, right_left, panel_top,
        right_width, panel_height
    )
    fill = right_panel.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(248, 250, 252)
    line = right_panel.line
    line.color.rgb = CS_SLATE
    line.width = Pt(2)
    
    # Collaboration title
    collab_title = slide_ops.shapes.add_textbox(right_left + Inches(0.2), panel_top + Inches(0.2), 
                                                 right_width - Inches(0.4), Inches(0.35))
    c_frame = collab_title.text_frame
    c_para = c_frame.paragraphs[0]
    c_para.text = "Collaboration Quality"
    c_para.font.name = TITLE_FONT_NAME
    c_para.font.size = Pt(16)
    c_para.font.bold = True
    c_para.font.color.rgb = CS_NAVY
    c_para.alignment = PP_ALIGN.LEFT
    
    # Collaboration metrics (3 compact rows)
    collab_data = [
        {"label": "Average Touches", "value": str(data.avg_touches), "detail": "per incident"},
        {"label": "Client Participation", "value": data.client_participation, "detail": "with client input"},
        {"label": "Client-Led Closures", "value": data.client_led_closures, "detail": "closed by team"}
    ]
    
    row_height = Inches(0.8)
    row_start = panel_top + Inches(0.6)
    
    for i, c in enumerate(collab_data):
        row_top = row_start + i * row_height
        
        # Value (left aligned)
        v_box = slide_ops.shapes.add_textbox(right_left + Inches(0.2), row_top, Inches(1.5), Inches(0.6))
        v_frame = v_box.text_frame
        v_para = v_frame.paragraphs[0]
        v_para.text = c["value"]
        v_para.font.name = TITLE_FONT_NAME
        v_para.font.size = Pt(28)
        v_para.font.bold = True
        v_para.font.color.rgb = CS_NAVY
        v_para.alignment = PP_ALIGN.LEFT
        
        # Label + detail (right of value)
        l_box = slide_ops.shapes.add_textbox(right_left + Inches(1.7), row_top + Inches(0.05), Inches(2.5), Inches(0.6))
        l_frame = l_box.text_frame
        l_para = l_frame.paragraphs[0]
        l_para.text = c["label"]
        l_para.font.name = TITLE_FONT_NAME
        l_para.font.size = Pt(12)
        l_para.font.bold = True
        l_para.font.color.rgb = CS_SLATE
        l_para.alignment = PP_ALIGN.LEFT
        
        d_para = l_frame.add_paragraph()
        d_para.text = c["detail"]
        d_para.font.name = BODY_FONT_NAME
        d_para.font.size = Pt(10)
        d_para.font.color.rgb = CS_SLATE
        d_para.alignment = PP_ALIGN.LEFT
    
    # Insight box at bottom
    insight_top = panel_top + panel_height + Inches(0.15)
    insight_shape = slide_ops.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.5), insight_top,
        prs.slide_width - Inches(1), Inches(0.55)
    )
    fill = insight_shape.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(240, 248, 255)
    line = insight_shape.line
    line.color.rgb = CS_BLUE
    line.width = Pt(2)
    
    insight_text = insight_shape.text_frame
    insight_text.text = f"Continuous protection: {data.after_hours_escalations} off-hours incidents handled with {data.client_participation} client engagement"
    insight_text.paragraphs[0].font.name = BODY_FONT_NAME
    insight_text.paragraphs[0].font.size = Pt(13)
    insight_text.paragraphs[0].font.color.rgb = CS_NAVY
    insight_text.paragraphs[0].font.bold = True
    insight_text.paragraphs[0].alignment = PP_ALIGN.CENTER
    insight_text.vertical_anchor = 1


def build_additional_content_slides(prs, data):
    """Create additional content slides for comprehensive coverage.
    
    Per CRITICALSTART branding guidelines:
    - All slides have transparent header and footer
    - Uses H1-H6 typography scale
    
    This function creates slides for:
    - After-Hours Customer Notifications
    - Response Efficiency
    - Collaboration Quality
    - Detection Quality (Detailed)
    - Security Outcomes Summary
    
    Args:
        prs (Presentation): The presentation object.
        data (ReportData): The report data object containing all metrics.
    """
    # NOTE: After-Hours and Collaboration slides removed - merged into Operational Coverage in Insights section
    # NOTE: Response Efficiency slide removed - covered in Protection Achieved section
    # NOTE: Detection Quality slide removed - covered in Protection Achieved section
    # NOTE: Security Outcomes Summary slide removed - redundant cost avoided metric (now shown only in Value Delivered)
    
    # This function is now deprecated - all slides have been consolidated into their respective sections.
    # Keeping the function for potential future use.
    pass


def build_forward_direction_slide(prs, data):
    """Create the Forward Direction slide (Slide 16).
    
    Per CRITICALSTART branding guidelines:
    - All slides have transparent header and footer
    - Uses H1-H6 typography scale
    
    Args:
        prs (Presentation): The presentation object.
        data (ReportData): The report data object containing all metrics.
    """
    # Forward Direction slide
    slide16, content_top_16 = setup_content_slide(prs, "Looking Ahead")
    
    # Section 1 - Next Period Targets
    section1_top = content_top_16 + Inches(0.1)
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
    
    # Add insight bar
    insight_top = section3_top + section3_height + Inches(0.2)
    add_insight_callout(
        slide16, prs,
        "Strategic Path Forward",
        "Targeted recommendations to strengthen your security posture and maximize partnership value.",
        insight_top,
        height=Inches(0.7)
    )


def build_contact_slide(prs, data):
    """Create the closing Contact/Thank You slide.
    
    Args:
        prs (Presentation): The presentation object.
        data (ReportData): The report data object containing all metrics.
    """
    blank_slide_layout = prs.slide_layouts[6]
    slide_contact = prs.slides.add_slide(blank_slide_layout)
    slide_number = get_slide_number(prs)
    
    # Add master slide elements
    add_master_slide_elements(slide_contact, prs, slide_number=slide_number,
                               include_header=True, include_footer=True)
    
    # Add logo at top right
    add_logo(slide_contact, position='top_right', prs=prs)
    
    # Center content area
    content_top = HEADER_HEIGHT + Inches(0.8)
    
    # Thank You message
    thank_you_box = slide_contact.shapes.add_textbox(
        Inches(0.5), content_top,
        prs.slide_width - Inches(1), Inches(1.0)
    )
    thank_you_frame = thank_you_box.text_frame
    thank_you_para = thank_you_frame.paragraphs[0]
    thank_you_para.text = "Thank You"
    thank_you_para.font.name = TITLE_FONT_NAME
    thank_you_para.font.size = Pt(60)
    thank_you_para.font.bold = True
    thank_you_para.font.color.rgb = CS_NAVY
    thank_you_para.alignment = PP_ALIGN.CENTER
    
    # Subtitle
    subtitle_box = slide_contact.shapes.add_textbox(
        Inches(0.5), content_top + Inches(1.2),
        prs.slide_width - Inches(1), Inches(0.6)
    )
    subtitle_frame = subtitle_box.text_frame
    subtitle_para = subtitle_frame.paragraphs[0]
    subtitle_para.text = "Questions? We're here to help."
    subtitle_para.font.name = BODY_FONT_NAME
    subtitle_para.font.size = Pt(24)
    subtitle_para.font.color.rgb = CS_SLATE
    subtitle_para.alignment = PP_ALIGN.CENTER
    
    # Contact card
    card_width = Inches(6)
    card_height = Inches(1.5)
    card_left = (prs.slide_width - card_width) / 2
    card_top = content_top + Inches(2.2)
    
    card_shape = slide_contact.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, card_left, card_top,
        card_width, card_height
    )
    fill = card_shape.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(240, 248, 255)
    line = card_shape.line
    line.color.rgb = CS_BLUE
    line.width = Pt(2)
    
    # Contact info
    contact_box = slide_contact.shapes.add_textbox(
        card_left + Inches(0.3), card_top + Inches(0.3),
        card_width - Inches(0.6), card_height - Inches(0.6)
    )
    contact_frame = contact_box.text_frame
    contact_frame.word_wrap = True
    
    contact_para1 = contact_frame.paragraphs[0]
    contact_para1.text = "Your Customer Success Manager"
    contact_para1.font.name = TITLE_FONT_NAME
    contact_para1.font.size = Pt(18)
    contact_para1.font.bold = True
    contact_para1.font.color.rgb = CS_NAVY
    contact_para1.alignment = PP_ALIGN.CENTER
    
    contact_para2 = contact_frame.add_paragraph()
    contact_para2.text = "support@criticalstart.com"
    contact_para2.font.name = BODY_FONT_NAME
    contact_para2.font.size = Pt(16)
    contact_para2.font.color.rgb = CS_BLUE
    contact_para2.alignment = PP_ALIGN.CENTER
    
    contact_para3 = contact_frame.add_paragraph()
    contact_para3.text = "www.criticalstart.com"
    contact_para3.font.name = BODY_FONT_NAME
    contact_para3.font.size = Pt(14)
    contact_para3.font.color.rgb = CS_SLATE
    contact_para3.alignment = PP_ALIGN.CENTER
    
    # Report date footer
    report_date_box = slide_contact.shapes.add_textbox(
        Inches(0.5), prs.slide_height - Inches(0.7),
        prs.slide_width - Inches(1), Inches(0.3)
    )
    report_date_frame = report_date_box.text_frame
    report_date_para = report_date_frame.paragraphs[0]
    report_date_para.text = f"Report Date: {data.report_date}"
    report_date_para.font.name = BODY_FONT_NAME
    report_date_para.font.size = Pt(11)
    report_date_para.font.color.rgb = CS_SLATE
    report_date_para.font.italic = True
    report_date_para.alignment = PP_ALIGN.CENTER


if __name__ == "__main__":
    main()

