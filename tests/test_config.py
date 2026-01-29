"""Tests for configuration loading and schema."""

import pytest
from pathlib import Path
import tempfile
import yaml

from presentation_toolkit.config import (
    BrandConfig,
    ColorPalette,
    FontConfig,
    load_config,
    save_config,
    create_minimal_config,
)


def test_color_palette_normalization():
    """Test that colors are normalized to lowercase without #."""
    palette = ColorPalette(
        primary={"blue": "#009CDE", "navy": "12285F"},
        secondary={"white": "FFFFFF"}
    )
    assert palette.primary["blue"] == "009cde"
    assert palette.primary["navy"] == "12285f"
    assert palette.secondary["white"] == "ffffff"


def test_color_palette_all_colors():
    """Test that all_colors returns combined dict."""
    palette = ColorPalette(
        primary={"blue": "009cde"},
        secondary={"white": "ffffff"},
        tertiary={"yellow": "ffc423"}
    )
    all_colors = palette.all_colors()
    assert "blue" in all_colors
    assert "white" in all_colors
    assert "yellow" in all_colors


def test_font_config_defaults():
    """Test FontConfig default values."""
    config = FontConfig(brand=["Brand Font"])
    assert "Arial" in config.replace
    assert "Calibri" in config.replace


def test_brand_config_is_brand_color():
    """Test brand color checking."""
    config = BrandConfig(
        brand_name="Test",
        colors=ColorPalette(primary={"blue": "009CDE"}),
        fonts=FontConfig(brand=["Test Font"])
    )
    assert config.is_brand_color("009cde") is True
    assert config.is_brand_color("#009CDE") is True
    assert config.is_brand_color("ff0000") is False
    assert config.is_brand_color(None) is True  # Theme colors


def test_brand_config_is_bad_font():
    """Test bad font detection."""
    config = BrandConfig(
        brand_name="Test",
        colors=ColorPalette(primary={"blue": "009CDE"}),
        fonts=FontConfig(
            brand=["ZT Gatha", "Noto Sans"],
            replace=["Arial", "Calibri"]
        )
    )
    assert config.is_bad_font("Arial") is True
    assert config.is_bad_font("arial") is True
    assert config.is_bad_font("Calibri Light") is True
    assert config.is_bad_font("ZT Gatha Bold") is False
    assert config.is_bad_font("Noto Sans") is False
    assert config.is_bad_font(None) is False


def test_load_config_yaml():
    """Test loading config from YAML file."""
    config_data = {
        "version": "1.0",
        "brand_name": "Test Brand",
        "colors": {
            "primary": {"blue": "0066CC"},
            "secondary": {"white": "FFFFFF"}
        },
        "fonts": {
            "brand": ["Test Font"],
            "replace": ["Arial"]
        }
    }

    with tempfile.NamedTemporaryFile(mode='w', suffix='.yaml', delete=False) as f:
        yaml.dump(config_data, f)
        temp_path = f.name

    try:
        config = load_config(temp_path)
        assert config.brand_name == "Test Brand"
        assert config.colors.primary["blue"] == "0066cc"
        assert "Test Font" in config.fonts.brand
    finally:
        Path(temp_path).unlink()


def test_create_minimal_config():
    """Test minimal config creation."""
    config = create_minimal_config(
        brand_name="My Brand",
        brand_colors={"blue": "0066CC", "white": "FFFFFF"},
        brand_fonts=["My Font"]
    )
    assert config.brand_name == "My Brand"
    assert "blue" in config.colors.primary
    assert "My Font" in config.fonts.brand
    # Should have default slide catalog
    assert len(config.slide_catalog) > 0


def test_save_and_load_config():
    """Test round-trip save and load."""
    config = create_minimal_config(
        brand_name="Round Trip",
        brand_colors={"blue": "0066CC"},
        brand_fonts=["Test"]
    )

    with tempfile.NamedTemporaryFile(mode='w', suffix='.yaml', delete=False) as f:
        temp_path = f.name

    try:
        save_config(config, temp_path)
        loaded = load_config(temp_path)
        assert loaded.brand_name == config.brand_name
        assert loaded.colors.primary == config.colors.primary
    finally:
        Path(temp_path).unlink()
