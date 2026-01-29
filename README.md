# Claude Presentation Toolkit

A generic, brand-agnostic toolkit for migrating PowerPoint presentations to branded templates. Supports intelligent content-type detection, layout variety tracking, and brand compliance analysis.

## Brand-Specific Skills

This toolkit provides the engine for brand-specific presentation skills:

- **[drupal-brand-skill](https://github.com/zivtech/drupal-brand-skill)** - Drupal brand guidelines with presentation migration support

Brand skills use this toolkit's generic capabilities with their specific configurations (colors, fonts, layouts, templates).

## Features

- **Content Extraction**: Extract text and images from PPTX/PDF files
- **Intelligent Migration**: Auto-detect content types (stats, quotes, bullets, etc.) and map to appropriate template layouts
- **Layout Variety**: Prevents consecutive layout repeats for visual interest
- **Brand Compliance Analysis**: Scan presentations for off-brand fonts and colors
- **Configurable**: All brand-specific values (colors, fonts, slide catalog) loaded from YAML config

## Installation

```bash
pip install -e .
```

Or use directly:

```bash
python -m presentation_toolkit.cli migrate input.pptx output.pptx --config brand.yaml
```

## Quick Start

### 1. Create a Brand Configuration

See `examples/sample_brand_config.yaml` for a complete example.

```yaml
version: "1.0"
brand_name: "My Brand"

colors:
  primary:
    brand_blue: "0066CC"
  secondary:
    white: "FFFFFF"
    black: "000000"

fonts:
  brand: ["Brand Font", "Noto Sans"]
  replace: ["Arial", "Calibri", "Helvetica"]
```

### 2. Migrate a Presentation

```bash
# Using CLI
pptx-migrate input.pptx output.pptx --config my-brand.yaml --template template.pptx

# Using Python
from presentation_toolkit import migrate_presentation, load_config

config = load_config("my-brand.yaml")
migrate_presentation("input.pptx", "output.pptx", config, "template.pptx")
```

### 3. Analyze Brand Compliance

```bash
pptx-analyze presentation.pptx --config my-brand.yaml
```

### 4. Extract Content

```bash
pptx-extract input.pptx --output content.md --images
```

## Configuration Reference

See the [Configuration Schema](src/presentation_toolkit/config/schema.py) for full details.

### Required Fields

- `brand_name`: Name of the brand
- `colors`: Color palette with hex values (without #)
- `fonts.brand`: List of approved brand fonts
- `fonts.replace`: List of fonts to flag for replacement

### Optional Fields

- `template.default`: Path to default template PPTX
- `slide_catalog`: Mapping of content types to template slide indices
- `text_capacity`: Character limits and font sizes per layout type
- `content_patterns`: Regex patterns for content type detection

## CLI Commands

### `pptx-migrate`

Migrate a presentation to a branded template.

```bash
pptx-migrate input.pptx output.pptx --config brand.yaml [options]

Options:
  --template PATH    Path to template PPTX (overrides config)
  --no-images        Skip image extraction/insertion
  --verbose          Show detailed progress
```

### `pptx-analyze`

Analyze a presentation for brand compliance.

```bash
pptx-analyze presentation.pptx --config brand.yaml [options]

Options:
  --json             Output results as JSON
  --strict           Fail on any compliance issue
```

### `pptx-extract`

Extract content from a presentation.

```bash
pptx-extract input.pptx [options]

Options:
  --output PATH      Output markdown file (default: input.md)
  --images           Also extract images to folder
```

## Python API

```python
from presentation_toolkit import (
    load_config,
    migrate_presentation,
    analyze_presentation,
    extract_pptx_to_markdown,
)

# Load brand configuration
config = load_config("brand.yaml")

# Migrate a presentation
migrate_presentation(
    input_path="source.pptx",
    output_path="branded.pptx",
    config=config,
    template_path="template.pptx",
    insert_images=True
)

# Analyze for brand compliance
issues = analyze_presentation("deck.pptx", config)
for slide in issues:
    print(f"Slide {slide['num']}: {slide['issues']}")

# Extract content
slides = extract_pptx_to_markdown("source.pptx", "content.md", extract_images=True)
```

## License

MIT
