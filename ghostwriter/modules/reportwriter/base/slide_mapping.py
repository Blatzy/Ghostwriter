"""Slide mapping manager for flexible PPTX template configuration."""

import logging
from typing import Dict, List, Optional, Tuple

from pptx import Presentation

logger = logging.getLogger(__name__)


class SlideConfig:
    """Represents configuration for a single slide type."""

    def __init__(
        self,
        type: str,
        layout_index: int,
        mode: str,
        enabled: bool,
        position: int,
    ):
        self.type = type
        self.layout_index = layout_index
        self.mode = mode  # 'static' or 'dynamic'
        self.enabled = enabled
        self.position = position

    def to_dict(self) -> dict:
        """Convert to dictionary for JSON serialization."""
        return {
            "type": self.type,
            "layout_index": self.layout_index,
            "mode": self.mode,
            "enabled": self.enabled,
            "position": self.position,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "SlideConfig":
        """Create from dictionary."""
        return cls(
            type=data["type"],
            layout_index=data["layout_index"],
            mode=data["mode"],
            enabled=data.get("enabled", True),
            position=data["position"],
        )


class SlideMappingManager:
    """Manages slide mapping configuration for PPTX templates."""

    # Slide type definitions
    SLIDE_TYPES = {
        # Project slides
        "title": "Title Slide",
        "agenda": "Agenda",
        "introduction": "Team Introduction",
        "assessment_details": "Assessment Details",
        "methodology": "Methodology",
        "timeline": "Assessment Timeline",
        "attack_path": "Attack Path Overview",
        # Report slides
        "observations_overview": "Positive Observations Overview",
        "observation": "Individual Observation Slide",
        "findings_overview": "Findings Overview",
        "finding": "Individual Finding Slide",
        "recommendations": "Recommendations",
        "next_steps": "Next Steps",
        "final": "Final/Closing Slide",
    }

    # Default mapping for backwards compatibility
    DEFAULT_MAPPING = {
        "version": 1,
        "slides": [
            {"type": "title", "layout_index": 0, "mode": "dynamic", "enabled": True, "position": 1},
            {"type": "agenda", "layout_index": 1, "mode": "dynamic", "enabled": True, "position": 2},
            {"type": "introduction", "layout_index": 1, "mode": "dynamic", "enabled": True, "position": 3},
            {"type": "assessment_details", "layout_index": 1, "mode": "dynamic", "enabled": True, "position": 4},
            {"type": "methodology", "layout_index": 1, "mode": "dynamic", "enabled": True, "position": 5},
            {"type": "timeline", "layout_index": 1, "mode": "dynamic", "enabled": True, "position": 6},
            {"type": "attack_path", "layout_index": 1, "mode": "dynamic", "enabled": True, "position": 7},
            {"type": "observations_overview", "layout_index": 1, "mode": "dynamic", "enabled": True, "position": 8},
            {"type": "observation", "layout_index": 1, "mode": "dynamic", "enabled": True, "position": 9},
            {"type": "findings_overview", "layout_index": 1, "mode": "dynamic", "enabled": True, "position": 10},
            {"type": "finding", "layout_index": 1, "mode": "dynamic", "enabled": True, "position": 11},
            {"type": "recommendations", "layout_index": 1, "mode": "dynamic", "enabled": True, "position": 12},
            {"type": "next_steps", "layout_index": 1, "mode": "dynamic", "enabled": True, "position": 13},
            {"type": "final", "layout_index": 12, "mode": "dynamic", "enabled": True, "position": 14},
        ],
    }

    def __init__(
        self,
        mapping_data: Optional[dict] = None,
        presentation: Optional[Presentation] = None,
    ):
        """
        Initialize with mapping data from ReportTemplate.slide_mapping.

        Args:
            mapping_data: Dictionary from JSONField or None for defaults
            presentation: python-pptx Presentation object for validation
        """
        self.presentation = presentation

        # Validate and sanitize mapping_data
        if mapping_data is None or not isinstance(mapping_data, dict):
            logger.warning("Invalid or missing slide mapping data, using defaults")
            self.mapping_data = self.DEFAULT_MAPPING.copy()
        else:
            self.mapping_data = mapping_data

        try:
            self.slides = self._parse_slides()
        except Exception as e:
            logger.exception("Failed to parse slide mapping, using defaults: %s", e)
            self.mapping_data = self.DEFAULT_MAPPING.copy()
            self.slides = self._parse_slides()

    def _parse_slides(self) -> List[SlideConfig]:
        """Parse slides from mapping data."""
        slides_data = self.mapping_data.get("slides", [])
        slides = []
        for s in slides_data:
            try:
                slides.append(SlideConfig.from_dict(s))
            except (KeyError, TypeError, ValueError) as e:
                logger.warning("Failed to parse slide config: %s. Skipping. Error: %s", s, e)
                continue
        return slides

    def get_slide_config(self, slide_type: str) -> Optional[SlideConfig]:
        """Get configuration for a specific slide type."""
        for slide in self.slides:
            if slide.type == slide_type:
                return slide
        return None

    def get_layout_index(self, slide_type: str, fallback: int = 1) -> int:
        """
        Get layout index for a slide type with fallback.

        Args:
            slide_type: The slide type to look up
            fallback: Default layout index if not found or invalid

        Returns:
            Layout index to use
        """
        config = self.get_slide_config(slide_type)
        if not config or not config.enabled:
            return fallback

        # Validate layout exists in presentation
        if self.presentation:
            try:
                layout_count = len(self.presentation.slide_layouts)
                if config.layout_index >= layout_count:
                    logger.warning(
                        "Layout index %d for slide type '%s' exceeds available layouts (%d). Falling back to layout %d.",
                        config.layout_index,
                        slide_type,
                        layout_count,
                        fallback,
                    )
                    return fallback
            except Exception as e:
                logger.warning("Error validating layout index: %s. Using fallback.", e)
                return fallback

        return config.layout_index

    def is_slide_enabled(self, slide_type: str) -> bool:
        """Check if a slide type is enabled."""
        config = self.get_slide_config(slide_type)
        return config.enabled if config else True

    def get_slides_by_position(self) -> List[SlideConfig]:
        """Get all enabled slides sorted by position."""
        enabled = [s for s in self.slides if s.enabled]
        return sorted(enabled, key=lambda s: s.position)

    def validate(self) -> Tuple[List[str], List[str]]:
        """
        Validate the slide mapping configuration.

        Returns:
            Tuple of (warnings, errors)
        """
        warnings = []
        errors = []

        # Check for duplicate positions
        positions = [s.position for s in self.slides if s.enabled]
        if len(positions) != len(set(positions)):
            warnings.append("Duplicate position values found in slide mapping")

        # Check for invalid slide types
        for slide in self.slides:
            if slide.type not in self.SLIDE_TYPES and not slide.type.startswith("custom_"):
                warnings.append(f"Unknown slide type: {slide.type}")

        # Check for invalid modes
        for slide in self.slides:
            if slide.mode not in ("static", "dynamic"):
                errors.append(f"Invalid mode '{slide.mode}' for slide type {slide.type}")

        # Validate layout indices if presentation is available
        if self.presentation:
            try:
                layout_count = len(self.presentation.slide_layouts)
                for slide in self.slides:
                    if slide.enabled and slide.layout_index >= layout_count:
                        errors.append(
                            f"Layout index {slide.layout_index} for slide type '{slide.type}' "
                            f"exceeds available layouts (0-{layout_count-1})"
                        )
            except Exception as e:
                logger.warning("Error validating layouts: %s", e)

        # Check for required slide types
        required_types = ["title", "final"]
        for req_type in required_types:
            config = self.get_slide_config(req_type)
            if not config or not config.enabled:
                warnings.append(f"Required slide type '{req_type}' is not enabled")

        return warnings, errors

    def to_dict(self) -> dict:
        """Export mapping to dictionary for JSON storage."""
        return {
            "version": self.mapping_data.get("version", 1),
            "slides": [s.to_dict() for s in self.slides],
        }

    @classmethod
    def extract_layouts_from_pptx(cls, pptx_path: str) -> List[Dict[str, any]]:
        """
        Extract layout information from a PPTX file.

        Args:
            pptx_path: Path to PPTX template file

        Returns:
            List of dicts with layout info: [{'index': 0, 'name': 'Title Slide'}, ...]
        """
        try:
            prs = Presentation(pptx_path)
            layouts = []
            for idx, layout in enumerate(prs.slide_layouts):
                layouts.append(
                    {
                        "index": idx,
                        "name": layout.name,
                    }
                )
            return layouts
        except Exception as e:
            logger.exception("Failed to extract layouts from %s: %s", pptx_path, e)
            return []
