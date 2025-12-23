"""Test script to verify branding functions work correctly."""
from generate_presentation import (
    create_presentation,
    apply_branding,
    add_logo,
    get_brand_colors,
    create_title_slide_layout,
    create_section_header_layout,
    create_content_slide_layout
)

def test_branding():
    """Test all branding functions."""
    print("Testing Critical Start PowerPoint branding...")
    
    # Test color constants
    colors = get_brand_colors()
    print(f"✓ Brand colors: {list(colors.keys())}")
    
    # Create presentation
    prs = create_presentation()
    print(f"✓ Presentation created: {prs.slide_width} x {prs.slide_height}")
    
    # Apply branding
    apply_branding(prs)
    print("✓ Branding applied to presentation")
    
    # Create title slide
    title_slide = create_title_slide_layout(
        prs,
        "Escalation to Client Details Report",
        "Signature Tier - Test Period"
    )
    print("✓ Title slide created")
    
    # Create section header
    section_slide = create_section_header_layout(prs, "Executive Summary")
    print("✓ Section header slide created")
    
    # Create content slide
    content_slide = create_content_slide_layout(
        prs,
        "Key Metrics",
        [
            "2,110 Alerts Triaged & Routed",
            "11 True Threats Contained",
            "34% Faster Response Than Peers",
            "158 After-Hours Escalations Absorbed"
        ]
    )
    print("✓ Content slide created")
    
    # Save test presentation
    output_path = "output/test_branding.pptx"
    prs.save(output_path)
    print(f"✓ Test presentation saved to {output_path}")
    print("\nAll tests passed! ✓")

if __name__ == "__main__":
    test_branding()


