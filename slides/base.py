"""Base classes and utilities for slide builders.

This module provides the foundation for modular slide building:
- SlideBuilder base class for consistent slide creation
- Shared utility functions for layout and styling
- Type hints and documentation patterns

Note: The actual slide builder implementations are currently in
generate_presentation.py. This module establishes the pattern
for gradual migration to the slides/ package.
"""
from abc import ABC, abstractmethod
from typing import Optional, Tuple
from pptx import Presentation
from pptx.slide import Slide


class SlideBuilder(ABC):
    """Abstract base class for slide section builders.

    Each section (Executive Summary, Value Delivered, etc.) should
    have its own concrete implementation of this class.

    Example:
        class ExecutiveSummaryBuilder(SlideBuilder):
            def build(self):
                self._build_title_slide()
                self._build_dashboard_slide()
                self._build_funnel_slide()

            def _build_title_slide(self):
                # Implementation...
                pass
    """

    def __init__(self, prs: Presentation, data):
        """Initialize the builder with presentation and data.

        Args:
            prs: The PowerPoint presentation object
            data: ReportData instance containing all metrics
        """
        self.prs = prs
        self.data = data

    @abstractmethod
    def build(self) -> None:
        """Build all slides for this section.

        Implementations should create slides in the correct order
        and handle any section-specific logic.
        """
        pass

    def setup_content_slide(self, title: str) -> Tuple[Slide, float]:
        """Create a standard content slide with header/footer.

        This is a convenience wrapper around helpers.setup_content_slide.

        Args:
            title: The slide title

        Returns:
            Tuple of (slide, content_top) where content_top is the
            Y coordinate for placing content below the header.
        """
        from helpers import setup_content_slide
        return setup_content_slide(self.prs, title)

    def get_slide_number(self) -> int:
        """Get the current slide number (1-indexed)."""
        from helpers import get_slide_number
        return get_slide_number(self.prs)

    def add_logo(self, slide: Slide, position: str = 'top_right') -> None:
        """Add the Critical Start logo to a slide.

        Args:
            slide: The slide to add the logo to
            position: Logo position ('top_right', 'top_left', 'bottom_right')
        """
        from helpers import add_logo
        add_logo(slide, position=position, prs=self.prs)


class SectionBuilder(SlideBuilder):
    """Base class for section builders that include a section header.

    Section headers are the gray transition cards between major
    sections (Executive Summary, Value Delivered, etc.).
    """

    @property
    @abstractmethod
    def section_title(self) -> str:
        """The title shown on the section header card."""
        pass

    @property
    @abstractmethod
    def section_narrative(self) -> str:
        """The subtitle/narrative on the section header card."""
        pass

    def build_section_header(self) -> Slide:
        """Create the section header card.

        Returns:
            The created section header slide
        """
        # Import here to avoid circular imports
        # In the final refactored state, create_section_header_layout
        # would be in slides/common.py
        import sys
        sys.path.insert(0, str(__file__).rsplit('/', 2)[0])
        from generate_presentation import create_section_header_layout

        return create_section_header_layout(
            self.prs,
            self.section_title,
            self.section_narrative
        )
