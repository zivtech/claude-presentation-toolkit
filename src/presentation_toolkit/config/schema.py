"""
Configuration Schema for Presentation Toolkit

Pydantic models defining the brand configuration structure.
All brand-specific values (colors, fonts, slide catalog) are defined here.
"""

from typing import Dict, List, Optional, Tuple
from pydantic import BaseModel, Field, field_validator
import re


class ColorPalette(BaseModel):
    """Color palette with named colors and hex values."""
    primary: Dict[str, str] = Field(default_factory=dict, description="Primary brand colors")
    secondary: Dict[str, str] = Field(default_factory=dict, description="Secondary brand colors")
    tertiary: Dict[str, str] = Field(default_factory=dict, description="Tertiary/accent colors")

    @field_validator('primary', 'secondary', 'tertiary', mode='before')
    @classmethod
    def normalize_hex_colors(cls, v):
        """Normalize hex colors to lowercase without # prefix."""
        if isinstance(v, dict):
            return {k: v_val.lower().lstrip('#') for k, v_val in v.items()}
        return v

    def all_colors(self) -> Dict[str, str]:
        """Get all colors as a single dict."""
        all_colors = {}
        all_colors.update(self.primary)
        all_colors.update(self.secondary)
        all_colors.update(self.tertiary)
        return all_colors


class FontConfig(BaseModel):
    """Font configuration for brand compliance."""
    brand: List[str] = Field(description="List of approved brand fonts")
    replace: List[str] = Field(
        default_factory=lambda: ["Arial", "Calibri", "Helvetica"],
        description="Fonts that should be replaced with brand fonts"
    )
    mapping: Dict[str, str] = Field(
        default_factory=dict,
        description="Mapping of usage context to specific font (e.g., 'headlines': 'Brand Bold')"
    )


class TemplateConfig(BaseModel):
    """Template file configuration."""
    default: Optional[str] = Field(None, description="Path to default template PPTX")
    extended: Optional[str] = Field(None, description="Path to extended template with additional layouts")


class SlideCategory(BaseModel):
    """Configuration for a slide category."""
    indices: List[int] = Field(description="Template slide indices for this category")
    description: Optional[str] = Field(None, description="Description of when to use this category")


class TextCapacity(BaseModel):
    """Text capacity configuration for a layout."""
    title_max_chars: int = Field(150, description="Maximum characters for title")
    body_max_chars: int = Field(500, description="Maximum characters for body")
    title_font_size: int = Field(2400, description="Title font size in hundredths of a point")
    body_font_size: int = Field(1400, description="Body font size in hundredths of a point")

    @classmethod
    def from_tuple(cls, values: Tuple[int, int, int, int]) -> "TextCapacity":
        """Create from a tuple of (title_max, body_max, title_pt, body_pt)."""
        return cls(
            title_max_chars=values[0],
            body_max_chars=values[1],
            title_font_size=values[2],
            body_font_size=values[3]
        )


class ContentPattern(BaseModel):
    """Patterns for content type detection."""
    title_patterns: List[str] = Field(default_factory=list, description="Regex patterns to match in title")
    body_patterns: List[str] = Field(default_factory=list, description="Regex patterns to match in body")
    anywhere_patterns: List[str] = Field(default_factory=list, description="Patterns to match anywhere")
    keywords: List[str] = Field(default_factory=list, description="Keywords to match (case-insensitive)")


class OrientationConfig(BaseModel):
    """Orientation tracking configuration."""
    left: List[str] = Field(default_factory=list, description="Categories with content on left")
    right: List[str] = Field(default_factory=list, description="Categories with content on right")
    center: List[str] = Field(default_factory=list, description="Categories with centered content")


class BrandConfig(BaseModel):
    """Complete brand configuration for presentation toolkit."""

    version: str = Field("1.0", description="Configuration schema version")
    brand_name: str = Field(description="Name of the brand")

    # Required configurations
    colors: ColorPalette = Field(description="Brand color palette")
    fonts: FontConfig = Field(description="Font configuration")

    # Optional configurations
    template: TemplateConfig = Field(default_factory=TemplateConfig, description="Template paths")

    slide_catalog: Dict[str, SlideCategory] = Field(
        default_factory=dict,
        description="Mapping of content types to template slide indices"
    )

    text_capacity: Dict[str, TextCapacity] = Field(
        default_factory=dict,
        description="Text capacity limits per layout category"
    )

    content_patterns: Dict[str, ContentPattern] = Field(
        default_factory=dict,
        description="Patterns for content type detection"
    )

    orientations: OrientationConfig = Field(
        default_factory=OrientationConfig,
        description="Orientation configuration for layout selection"
    )

    gui_colors: List[str] = Field(
        default_factory=lambda: ["blue", "coral", "yellow", "navy"],
        description="GUI block color rotation order"
    )

    def get_slide_indices(self, category: str) -> List[int]:
        """Get template slide indices for a category."""
        if category in self.slide_catalog:
            return self.slide_catalog[category].indices
        return []

    def get_text_capacity(self, category: str) -> TextCapacity:
        """Get text capacity for a layout category."""
        if category in self.text_capacity:
            return self.text_capacity[category]
        # Return default
        return self.text_capacity.get('default', TextCapacity())

    def is_brand_color(self, hex_color: str) -> bool:
        """Check if a color is a brand color."""
        if hex_color is None:
            return True  # Theme colors are OK
        normalized = hex_color.lower().lstrip('#')
        return normalized in self.colors.all_colors().values()

    def is_bad_font(self, font_name: str) -> bool:
        """Check if a font needs replacement."""
        if font_name is None:
            return False
        font_lower = font_name.lower()
        # Check if it's already a good font
        if any(good.lower() in font_lower for good in self.fonts.brand):
            return False
        # Check if it's a known bad font
        return any(bad.lower() in font_lower for bad in self.fonts.replace)

    def get_all_slide_indices(self) -> Dict[str, List[int]]:
        """Get all slide indices as a simple dict for backward compatibility."""
        return {k: v.indices for k, v in self.slide_catalog.items()}

    def get_all_text_capacities(self) -> Dict[str, Tuple[int, int, int, int]]:
        """Get all text capacities as tuples for backward compatibility."""
        return {
            k: (v.title_max_chars, v.body_max_chars, v.title_font_size, v.body_font_size)
            for k, v in self.text_capacity.items()
        }


# Default content patterns (can be overridden in config)
DEFAULT_CONTENT_PATTERNS = {
    'statistic': ContentPattern(
        title_patterns=[
            r'^\d+%',                    # Starts with percentage
            r'^\$[\d,]+',               # Dollar amounts
            r'^[\d,]+\+?\s*(million|billion|users|websites|developers|organizations)',
            r'^\d+x',                    # Multipliers like "10x"
            r'^\d+/\d+',                # Fractions
            r'^#\d+',                   # Rankings like "#1"
        ],
        anywhere_patterns=[
            r'\b\d{2,}%\b',             # Any percentage 10%+
            r'\b\d+\s*(million|billion)\b',
            r'\bover\s+\d+',
            r'\bmore than\s+\d+',
        ]
    ),
    'quote': ContentPattern(
        title_patterns=[
            r'^["""]',                   # Starts with quote mark
        ],
        body_patterns=[
            r'^\s*—',                    # Attribution dash
            r'said\s+\w+',              # "said [Name]"
        ]
    ),
    'numbered_step': ContentPattern(
        title_patterns=[
            r'^(step\s*)?\d+[.:\)]',    # "Step 1:" or "1." or "1)"
            r'^(first|second|third|fourth|fifth)',
            r'\bphase\s*\d+',
            r'\bpart\s*\d+',
        ]
    ),
    'bullet_list': ContentPattern(
        body_patterns=[
            r'^\s*[-•▪]\s+',            # Bullet markers
            r'\n\s*[-•▪]\s+',           # Multiple bullets
        ]
    ),
    'comparison': ContentPattern(
        anywhere_patterns=[
            r'\bvs\.?\b',               # "vs" or "vs."
            r'\bbefore\b.*\bafter\b',
            r'\bpros?\b.*\bcons?\b',
        ]
    ),
    'section_header': ContentPattern(
        keywords=['overview', 'introduction', 'summary', 'agenda', 'contents',
                 'why', 'what is', 'features', 'benefits', 'resources']
    ),
    'case_study': ContentPattern(
        keywords=['customer', 'client', 'case study', 'success story',
                 'testimonial', 'partner', 'rebuilt', 'transformed',
                 'organization', 'implemented'],
        title_patterns=[
            r'\bhow\s+\w+\s+(rebuilt|transformed|migrated|updated|launched)',
        ]
    ),
    'feature': ContentPattern(
        keywords=['feature', 'capability', 'integration', 'benefit',
                 'advantage', 'solution', 'powerful', 'flexible']
    ),
}


# Default slide catalog (generic, can be overridden)
DEFAULT_SLIDE_CATALOG = {
    'title_opening': SlideCategory(indices=[0], description="Opening/title slides"),
    'hero_photo': SlideCategory(indices=[1], description="Hero image with text overlay"),
    'statement_center': SlideCategory(indices=[3], description="Centered statement"),
    'section_divider': SlideCategory(indices=[4], description="Section divider"),
    'content_image_left': SlideCategory(indices=[5], description="Image left, text right"),
    'content_image_right': SlideCategory(indices=[6], description="Image right, text left"),
    'feature_default': SlideCategory(indices=[7], description="Feature with bullets"),
    'stat_default': SlideCategory(indices=[8], description="Statistics display"),
    'quote_default': SlideCategory(indices=[9], description="Quote slide"),
    'two_column': SlideCategory(indices=[10], description="Two-column layout"),
    'closing_cta': SlideCategory(indices=[11], description="Closing/CTA slide"),
}


# Default text capacity (generic, can be overridden)
DEFAULT_TEXT_CAPACITY = {
    'title_opening': TextCapacity(title_max_chars=200, body_max_chars=150, title_font_size=3600, body_font_size=1800),
    'hero_photo': TextCapacity(title_max_chars=150, body_max_chars=300, title_font_size=3600, body_font_size=1400),
    'statement_center': TextCapacity(title_max_chars=200, body_max_chars=300, title_font_size=3200, body_font_size=1600),
    'section_divider': TextCapacity(title_max_chars=150, body_max_chars=300, title_font_size=4000, body_font_size=1800),
    'content_image_left': TextCapacity(title_max_chars=150, body_max_chars=500, title_font_size=2400, body_font_size=1400),
    'content_image_right': TextCapacity(title_max_chars=150, body_max_chars=500, title_font_size=2400, body_font_size=1400),
    'feature_default': TextCapacity(title_max_chars=150, body_max_chars=500, title_font_size=2400, body_font_size=1400),
    'stat_default': TextCapacity(title_max_chars=50, body_max_chars=500, title_font_size=7200, body_font_size=1400),
    'quote_default': TextCapacity(title_max_chars=300, body_max_chars=200, title_font_size=2400, body_font_size=1600),
    'two_column': TextCapacity(title_max_chars=100, body_max_chars=600, title_font_size=2400, body_font_size=1200),
    'closing_cta': TextCapacity(title_max_chars=200, body_max_chars=200, title_font_size=3200, body_font_size=1800),
    'default': TextCapacity(title_max_chars=150, body_max_chars=500, title_font_size=2400, body_font_size=1400),
}
