"""
Claude Presentation Toolkit

A generic, brand-agnostic toolkit for migrating PowerPoint presentations
to branded templates with intelligent content-type detection and layout variety.
"""

__version__ = "0.1.0"

from .config import (
    BrandConfig,
    load_config,
    save_config,
    create_minimal_config,
)

from .migrate import (
    migrate_presentation,
    detect_and_parse,
    detect_content_type,
    LayoutSelector,
)

from .analyze import (
    analyze_presentation,
    get_analysis_json,
)

from .extract import (
    extract_pptx_to_markdown,
)

__all__ = [
    # Config
    'BrandConfig',
    'load_config',
    'save_config',
    'create_minimal_config',
    # Migration
    'migrate_presentation',
    'detect_and_parse',
    'detect_content_type',
    'LayoutSelector',
    # Analysis
    'analyze_presentation',
    'get_analysis_json',
    # Extraction
    'extract_pptx_to_markdown',
]
