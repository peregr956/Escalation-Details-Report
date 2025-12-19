"""Test script for build_executive_summary_slides function."""

from generate_presentation import create_presentation, build_executive_summary_slides
from report_data import get_report_data

def test_executive_summary_slides():
    """Test the executive summary slides generation."""
    # Create presentation
    prs = create_presentation()
    
    # Get report data
    data = get_report_data()
    
    # Build executive summary slides
    build_executive_summary_slides(prs, data)
    
    # Save presentation
    output_path = "output/test_executive_summary.pptx"
    prs.save(output_path)
    print(f"✓ Executive Summary slides created successfully!")
    print(f"  Saved to: {output_path}")
    print(f"  Total slides: {len(prs.slides)}")
    print(f"  Expected: 3 slides (Title, Executive Summary, Period Highlights)")

if __name__ == "__main__":
    test_executive_summary_slides()

