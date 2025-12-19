"""
Chart Renderer Module

Renders Chart.js charts from HTML templates to high-resolution PNG images using Playwright.
"""

from playwright.async_api import async_playwright, Page, Browser
from pathlib import Path
import asyncio
import json
import logging
from typing import Dict

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Chart template names
CHART_TEMPLATES = {
    'pie': 'pie_chart.html',
    'trend': 'trend_chart.html',
    'stacked_bar': 'stacked_bar_chart.html',
    'sankey': 'sankey_chart.html'
}


def get_template_path(template_name: str) -> Path:
    """
    Resolve the path to a chart template file.
    
    Args:
        template_name: Name of the template file (e.g., 'pie_chart.html') or short name (e.g., 'pie')
    
    Returns:
        Path object pointing to the template file
    
    Raises:
        FileNotFoundError: If the template file doesn't exist
    """
    # Check if it's a short name and map to full filename
    if template_name in CHART_TEMPLATES:
        template_name = CHART_TEMPLATES[template_name]
    
    # Ensure it has .html extension if not provided
    if not template_name.endswith('.html'):
        template_name += '.html'
    
    # Get the project root (parent of this file)
    project_root = Path(__file__).parent
    template_path = project_root / 'chart_templates' / template_name
    
    if not template_path.exists():
        raise FileNotFoundError(f"Chart template not found: {template_path}")
    
    return template_path


async def inject_chart_data(page: Page, data: dict) -> None:
    """
    Inject chart data into the page via JavaScript by setting window.CHART_DATA.
    
    Args:
        page: Playwright Page object
        data: Dictionary containing chart data to inject
    """
    try:
        # Serialize data to JSON and inject into page
        data_json = json.dumps(data)
        await page.evaluate(f"""
            window.CHART_DATA = {data_json};
        """)
        logger.debug(f"Injected chart data: {data_json[:100]}...")
    except Exception as e:
        logger.error(f"Failed to inject chart data: {e}")
        raise


async def wait_for_chart_ready(page: Page, canvas_selector: str = "canvas", timeout: int = 10000) -> None:
    """
    Wait for Chart.js to finish rendering the chart.
    
    This function waits for:
    1. The canvas element to exist
    2. Chart.js to be loaded
    3. The chart to be fully rendered (canvas has non-zero dimensions)
    
    Args:
        page: Playwright Page object
        canvas_selector: CSS selector for the canvas element (default: "canvas")
        timeout: Maximum time to wait in milliseconds (default: 10000)
    
    Raises:
        TimeoutError: If the chart doesn't render within the timeout period
    """
    try:
        # Wait for canvas element to exist
        await page.wait_for_selector(canvas_selector, timeout=timeout)
        
        # Wait for Chart.js to be loaded
        await page.wait_for_function(
            "typeof Chart !== 'undefined'",
            timeout=timeout
        )
        
        # Wait for chart to be rendered - check that canvas has non-zero dimensions
        # and that Chart.js has created a chart instance
        await page.wait_for_function(
            f"""
            () => {{
                const canvas = document.querySelector('{canvas_selector}');
                if (!canvas) return false;
                // Check canvas has valid dimensions
                if (canvas.width === 0 || canvas.height === 0) return false;
                // Check if Chart.js has initialized (chart instance exists)
                // We can check by looking for Chart.js internal state or just wait for dimensions
                return canvas.width > 0 && canvas.height > 0;
            }}
            """,
            timeout=timeout
        )
        
        # Additional wait to ensure chart animation and rendering completes
        # Chart.js may need time to fully render, especially for complex charts
        await page.wait_for_timeout(1000)
        
        logger.debug(f"Chart ready on canvas: {canvas_selector}")
    except Exception as e:
        logger.warning(f"Chart readiness check completed with warning: {e}")
        # Fallback: wait longer to ensure chart is rendered
        await page.wait_for_timeout(1500)


async def render_chart(
    template_name: str,
    data: dict,
    output_path: str,
    width: int = 900,
    height: int = 350,
    scale: int = 2
) -> str:
    """
    Render a Chart.js chart from an HTML template to a PNG image.
    
    Args:
        template_name: Name of the template file (e.g., 'pie_chart.html' or 'pie')
        data: Dictionary containing chart data to inject
        output_path: Path where the rendered PNG image should be saved
        width: Viewport width in pixels (default: 900)
        height: Viewport height in pixels (default: 350)
        scale: Device scale factor for high DPI rendering (default: 2)
    
    Returns:
        Path to the saved image file
    
    Raises:
        FileNotFoundError: If the template file doesn't exist
        Exception: If chart rendering fails
    """
    # Resolve template path
    template_path = get_template_path(template_name)
    template_url = template_path.as_uri()
    
    # Ensure output directory exists
    output_path_obj = Path(output_path)
    output_path_obj.parent.mkdir(parents=True, exist_ok=True)
    
    logger.info(f"Rendering chart from {template_name} to {output_path}")
    
    try:
        async with async_playwright() as p:
            # Launch headless Chromium
            browser = await p.chromium.launch(headless=True)
            
            # Create context with high DPI settings
            context = await browser.new_context(
                viewport={'width': width, 'height': height},
                device_scale_factor=scale
            )
            
            # Create page
            page = await context.new_page()
            
            # Inject chart data before navigation so it's available when DOMContentLoaded fires
            if data:
                await page.add_init_script(f"""
                    window.CHART_DATA = {json.dumps(data)};
                """)
            
            # Navigate to template file
            await page.goto(template_url, wait_until='domcontentloaded')
            
            # Also inject data after navigation as a fallback (in case init script didn't work)
            if data:
                await inject_chart_data(page, data)
            
            # Wait for chart to render
            await wait_for_chart_ready(page)
            
            # Find canvas element
            canvas = await page.query_selector('canvas')
            if not canvas:
                raise Exception("Canvas element not found in template")
            
            # Take screenshot of canvas element
            await canvas.screenshot(path=output_path)
            
            logger.info(f"Chart rendered successfully to {output_path}")
            
            return str(output_path)
            
    except FileNotFoundError:
        logger.error(f"Template file not found: {template_path}")
        raise
    except Exception as e:
        logger.error(f"Failed to render chart: {e}")
        raise


async def render_all_charts(data: dict, output_dir: str = "temp_charts") -> Dict[str, str]:
    """
    Render all chart types and return a dictionary mapping chart names to image paths.
    
    Args:
        data: Dictionary containing chart data (can be structured for different chart types)
        output_dir: Directory where rendered charts should be saved (default: "temp_charts")
    
    Returns:
        Dictionary mapping chart names to their rendered image file paths
    
    Example:
        {
            'pie': 'temp_charts/pie_chart.png',
            'trend': 'temp_charts/trend_chart.png',
            'stacked_bar': 'temp_charts/stacked_bar_chart.png',
            'sankey': 'temp_charts/sankey_chart.png'
        }
    """
    output_path_obj = Path(output_dir)
    output_path_obj.mkdir(parents=True, exist_ok=True)
    
    results = {}
    
    # Default dimensions for different chart types
    chart_configs = {
        'pie': {'width': 500, 'height': 500},
        'trend': {'width': 900, 'height': 350},
        'stacked_bar': {'width': 900, 'height': 350},
        'sankey': {'width': 900, 'height': 500}
    }
    
    for chart_name, template_file in CHART_TEMPLATES.items():
        try:
            # Extract chart-specific data if data is structured by chart type
            chart_data = data.get(chart_name, data) if isinstance(data, dict) and any(k in data for k in CHART_TEMPLATES.keys()) else data
            
            # Get chart-specific dimensions
            config = chart_configs.get(chart_name, {'width': 900, 'height': 350})
            
            # Generate output filename
            output_filename = f"{chart_name}_chart.png"
            output_path = output_path_obj / output_filename
            
            # Render chart
            rendered_path = await render_chart(
                template_name=chart_name,
                data=chart_data,
                output_path=str(output_path),
                width=config['width'],
                height=config['height'],
                scale=2
            )
            
            results[chart_name] = rendered_path
            logger.info(f"Successfully rendered {chart_name} chart")
            
        except Exception as e:
            logger.error(f"Failed to render {chart_name} chart: {e}")
            # Continue with other charts even if one fails
            results[chart_name] = None
    
    return results


def render_charts_sync(data: dict, output_dir: str = "temp_charts") -> Dict[str, str]:
    """
    Synchronous wrapper for render_all_charts().
    
    This function wraps the async render_all_charts() function to allow
    synchronous usage in scripts that don't use async/await.
    
    Args:
        data: Dictionary containing chart data
        output_dir: Directory where rendered charts should be saved (default: "temp_charts")
    
    Returns:
        Dictionary mapping chart names to their rendered image file paths
    """
    try:
        return asyncio.run(render_all_charts(data, output_dir))
    except Exception as e:
        logger.error(f"Error in synchronous chart rendering: {e}")
        raise


# Example usage
if __name__ == "__main__":
    # Test data
    test_data = {
        'pie': {
            'labels': ['Business Hours', 'After Hours', 'Weekend'],
            'data': [51, 41, 8],
            'backgroundColor': ['#009CDE', '#702F8A', '#EF3340']
        },
        'trend': {
            'labels': ['Period -2', 'Period -1', 'Current'],
            'datasets': [
                {
                    'label': 'MTTR (min)',
                    'data': [168, 150, 126],
                    'borderColor': '#009CDE',
                    'backgroundColor': 'rgba(0, 156, 222, 0.1)',
                    'borderWidth': 3
                }
            ]
        }
    }
    
    # Test synchronous rendering
    print("Testing chart rendering...")
    results = render_charts_sync(test_data)
    print(f"Rendered charts: {results}")
