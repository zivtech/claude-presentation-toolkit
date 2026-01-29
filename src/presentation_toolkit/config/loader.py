"""
Configuration Loader for Presentation Toolkit

Loads brand configuration from YAML or JSON files.
"""

import json
from pathlib import Path
from typing import Union

import yaml

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


def load_config(config_path: Union[str, Path]) -> BrandConfig:
    """Load brand configuration from YAML or JSON file.

    Args:
        config_path: Path to configuration file (.yaml, .yml, or .json)

    Returns:
        BrandConfig instance

    Raises:
        FileNotFoundError: If config file doesn't exist
        ValueError: If config format is unsupported or invalid
    """
    config_path = Path(config_path)

    if not config_path.exists():
        raise FileNotFoundError(f"Configuration file not found: {config_path}")

    suffix = config_path.suffix.lower()

    with open(config_path, 'r', encoding='utf-8') as f:
        if suffix in ['.yaml', '.yml']:
            data = yaml.safe_load(f)
        elif suffix == '.json':
            data = json.load(f)
        else:
            raise ValueError(f"Unsupported config format: {suffix}. Use .yaml, .yml, or .json")

    return parse_config(data)


def parse_config(data: dict) -> BrandConfig:
    """Parse configuration data into BrandConfig model.

    Handles backward compatibility and applies defaults where needed.

    Args:
        data: Raw configuration dictionary

    Returns:
        BrandConfig instance
    """
    # Process colors
    if 'colors' in data:
        colors_data = data['colors']
        # Handle flat color dict (backward compat)
        if all(isinstance(v, str) for v in colors_data.values()):
            data['colors'] = ColorPalette(primary=colors_data)
        elif not isinstance(colors_data, ColorPalette):
            data['colors'] = ColorPalette(**colors_data)

    # Process fonts
    if 'fonts' in data and not isinstance(data['fonts'], FontConfig):
        data['fonts'] = FontConfig(**data['fonts'])

    # Process template
    if 'template' in data and not isinstance(data['template'], TemplateConfig):
        data['template'] = TemplateConfig(**data['template'])

    # Process slide_catalog - handle both simple dict and full objects
    if 'slide_catalog' in data:
        catalog = {}
        for key, value in data['slide_catalog'].items():
            if isinstance(value, dict):
                if 'indices' in value:
                    catalog[key] = SlideCategory(**value)
                else:
                    # Handle {"category": [1, 2, 3]} format
                    catalog[key] = SlideCategory(indices=list(value.values())[0] if value else [])
            elif isinstance(value, list):
                # Handle {"category": [1, 2, 3]} format
                catalog[key] = SlideCategory(indices=value)
            elif isinstance(value, SlideCategory):
                catalog[key] = value
        data['slide_catalog'] = catalog
    else:
        # Apply defaults
        data['slide_catalog'] = DEFAULT_SLIDE_CATALOG

    # Process text_capacity - handle both tuple format and full objects
    if 'text_capacity' in data:
        capacity = {}
        for key, value in data['text_capacity'].items():
            if isinstance(value, (list, tuple)) and len(value) == 4:
                capacity[key] = TextCapacity.from_tuple(tuple(value))
            elif isinstance(value, dict):
                capacity[key] = TextCapacity(**value)
            elif isinstance(value, TextCapacity):
                capacity[key] = value
        data['text_capacity'] = capacity
    else:
        # Apply defaults
        data['text_capacity'] = DEFAULT_TEXT_CAPACITY

    # Process content_patterns
    if 'content_patterns' in data:
        patterns = {}
        for key, value in data['content_patterns'].items():
            if isinstance(value, dict):
                patterns[key] = ContentPattern(**value)
            elif isinstance(value, ContentPattern):
                patterns[key] = value
        data['content_patterns'] = patterns
    else:
        # Apply defaults
        data['content_patterns'] = DEFAULT_CONTENT_PATTERNS

    # Process orientations
    if 'orientations' in data and not isinstance(data['orientations'], OrientationConfig):
        data['orientations'] = OrientationConfig(**data['orientations'])

    return BrandConfig(**data)


def save_config(config: BrandConfig, output_path: Union[str, Path]) -> None:
    """Save brand configuration to YAML or JSON file.

    Args:
        config: BrandConfig instance to save
        output_path: Path for output file
    """
    output_path = Path(output_path)
    suffix = output_path.suffix.lower()

    # Convert to dict
    data = config.model_dump(exclude_none=True)

    with open(output_path, 'w', encoding='utf-8') as f:
        if suffix in ['.yaml', '.yml']:
            yaml.dump(data, f, default_flow_style=False, sort_keys=False, allow_unicode=True)
        elif suffix == '.json':
            json.dump(data, f, indent=2)
        else:
            raise ValueError(f"Unsupported config format: {suffix}")


def create_minimal_config(brand_name: str, brand_colors: dict, brand_fonts: list) -> BrandConfig:
    """Create a minimal brand configuration.

    Args:
        brand_name: Name of the brand
        brand_colors: Dictionary of color name -> hex value
        brand_fonts: List of approved brand fonts

    Returns:
        BrandConfig instance with defaults applied
    """
    return BrandConfig(
        brand_name=brand_name,
        colors=ColorPalette(primary=brand_colors),
        fonts=FontConfig(brand=brand_fonts),
        slide_catalog=DEFAULT_SLIDE_CATALOG,
        text_capacity=DEFAULT_TEXT_CAPACITY,
        content_patterns=DEFAULT_CONTENT_PATTERNS,
    )
