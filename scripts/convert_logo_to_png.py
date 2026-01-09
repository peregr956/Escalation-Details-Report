"""Convert Critical Start logo from SVG to PNG for PowerPoint use."""
import base64
import re
from pathlib import Path
import io

try:
    import cairosvg
    HAS_CAIROSVG = True
except ImportError:
    HAS_CAIROSVG = False
    try:
        from svglib.svglib import svg2rlg
        from reportlab.graphics import renderPM
        HAS_SVGLIB = True
    except ImportError:
        HAS_SVGLIB = False


def extract_logo_from_html():
    """Extract base64 logo from HTML file."""
    html_path = Path("escalation_report_enhanced-3.html")
    if not html_path.exists():
        return None
    
    with open(html_path, 'r', encoding='utf-8') as f:
        html_content = f.read()
    
    # Find base64 encoded SVG
    match = re.search(r'data:image/svg\+xml;base64,([A-Za-z0-9+/=]+)', html_content)
    if match:
        svg_base64 = match.group(1)
        svg_data = base64.b64decode(svg_base64)
        return svg_data
    return None


def convert_svg_to_png(svg_data, output_path, width=480, height=104):
    """Convert SVG data to PNG."""
    output_path = Path(output_path)
    
    if HAS_CAIROSVG:
        # Use cairosvg (preferred method)
        cairosvg.svg2png(bytestring=svg_data, write_to=str(output_path), 
                        output_width=width, output_height=height)
        return True
    elif HAS_SVGLIB:
        # Use svglib + reportlab
        drawing = svg2rlg(io.BytesIO(svg_data))
        if drawing:
            renderPM.drawToFile(drawing, str(output_path), fmt='PNG', 
                              configPIL={'width': width, 'height': height})
            return True
    else:
        # Fallback: try to use the existing SVG file
        svg_path = Path("assets/critical-start-logo.svg")
        if svg_path.exists():
            print("Warning: cairosvg and svglib not available. Using SVG file directly.")
            print("For best results, install cairosvg: pip install cairosvg")
            # Copy SVG as fallback (PowerPoint can sometimes use SVG)
            with open(svg_path, 'rb') as f:
                svg_content = f.read()
            # Try basic conversion with PIL (won't work for SVG, but we'll try)
            # Actually, let's just inform the user
            return False
    
    return False


def main():
    """Main conversion function."""
    output_path = Path("assets/critical-start-logo.png")
    
    # Check if PNG already exists
    if output_path.exists():
        print(f"✓ Logo PNG already exists at {output_path}")
        return
    
    # Try to extract from HTML first
    svg_data = extract_logo_from_html()
    
    # Fallback to existing SVG file
    if not svg_data:
        svg_path = Path("assets/critical-start-logo.svg")
        if svg_path.exists():
            with open(svg_path, 'rb') as f:
                svg_data = f.read()
        else:
            print("Error: Could not find logo file")
            return
    
    # Convert to PNG
    success = convert_svg_to_png(svg_data, output_path, width=480, height=104)
    
    if success:
        print(f"✓ Logo converted to {output_path}")
    else:
        print("⚠ Could not convert SVG to PNG automatically.")
        print("\nTo convert manually:")
        print("1. Open assets/critical-start-logo.svg in a browser or image editor")
        print("2. Export/Save as PNG at 480x104 pixels (or maintain aspect ratio)")
        print("3. Save as assets/critical-start-logo.png")
        print("\nOr install cairosvg in a virtual environment:")
        print("  python3 -m venv venv")
        print("  source venv/bin/activate")
        print("  pip install cairosvg")
        print("  python convert_logo_to_png.py")


if __name__ == "__main__":
    main()

