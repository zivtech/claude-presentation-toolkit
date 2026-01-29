"""Configuration module for Presentation Toolkit."""

from .schema import (
    BrandConfig,
    ColorPalette,
    FontConfig,
    TemplateConfig,
    SlideCategory,
    TextCapacity,
    ContentPattern,
    OrientationConfig,
    DEFAULT_CONTENT_PATTERNS,
    DEFAULT_SLIDE_CATALOG,
    DEFAULT_TEXT_CAPACITY,
)
from .loader import load_config, save_config, parse_config, create_minimal_config

__all__ = [
    'BrandConfig',
    'ColorPalette',
    'FontConfig',
    'TemplateConfig',
    'SlideCategory',
    'TextCapacity',
    'ContentPattern',
    'OrientationConfig',
    'DEFAULT_CONTENT_PATTERNS',
    'DEFAULT_SLIDE_CATALOG',
    'DEFAULT_TEXT_CAPACITY',
    'load_config',
    'save_config',
    'parse_config',
    'create_minimal_config',
]
