"""
Content Extraction Utility

Extracts content from PPTX files into markdown format that can be
edited and then migrated to a branded template.

Includes slide-to-image mapping for proper image migration.
"""

import re
import json
import shutil
import zipfile
import tempfile
from pathlib import Path
from datetime import datetime
from typing import Optional, List, Dict, Any, Union

from lxml import etree

from .pptx_utils import NSMAP, clean_text


def get_layout_name(rels_content: bytes, layout_map: Dict[int, str]) -> str:
    """Extract layout name from relationships XML."""
    if not rels_content:
        return "UNKNOWN"

    rels_tree = etree.fromstring(rels_content)
    for rel in rels_tree:
        target = rel.get('Target', '')
        if 'slideLayout' in target:
            match = re.search(r'slideLayout(\d+)', target)
            if match:
                layout_num = int(match.group(1))
                return layout_map.get(layout_num, f"LAYOUT_{layout_num}")
    return "UNKNOWN"


def get_slide_images(rels_content: bytes) -> List[str]:
    """Extract image references from slide relationships XML.

    Returns list of image filenames used by this slide.
    """
    if not rels_content:
        return []

    images = []
    rels_tree = etree.fromstring(rels_content)

    for rel in rels_tree:
        rel_type = rel.get('Type', '')
        target = rel.get('Target', '')

        if 'image' in rel_type.lower() and target:
            img_name = Path(target).name
            images.append(img_name)

    return images


def extract_pptx_to_markdown(
    pptx_path: Union[str, Path],
    output_path: Optional[Union[str, Path]] = None,
    extract_images: bool = False
) -> List[Dict[str, Any]]:
    """Extract PPTX content to markdown format.

    Args:
        pptx_path: Path to source PPTX
        output_path: Path for output markdown (default: same name with .md)
        extract_images: If True, also extract images to a folder

    Returns:
        slides_data: List of slide dictionaries with content and image mappings
    """
    pptx_path = Path(pptx_path)

    if output_path is None:
        output_path = pptx_path.with_suffix('.md')
    else:
        output_path = Path(output_path)

    print(f"Extracting: {pptx_path.name}")

    with tempfile.TemporaryDirectory() as tmpdir:
        work_dir = Path(tmpdir)

        with zipfile.ZipFile(pptx_path, 'r') as zf:
            zf.extractall(work_dir)

        slides_dir = work_dir / 'ppt/slides'
        rels_dir = slides_dir / '_rels'

        # Build layout name map from slideLayouts
        layout_map = {}
        layouts_dir = work_dir / 'ppt/slideLayouts'
        if layouts_dir.exists():
            for layout_file in layouts_dir.glob('slideLayout*.xml'):
                num = int(re.search(r'slideLayout(\d+)', layout_file.name).group(1))
                tree = etree.parse(str(layout_file))
                root = tree.getroot()

                name_attr = root.get('matchingName') or root.get('name')
                if name_attr:
                    name = name_attr.upper().replace(' ', '_').replace('-', '_')
                    name = re.sub(r'[^A-Z0-9_]', '', name)
                    layout_map[num] = name
                else:
                    layout_map[num] = f"LAYOUT_{num}"

        # Get all slide files sorted by number
        slide_files = sorted(
            [f for f in slides_dir.glob('slide*.xml') if f.is_file()],
            key=lambda x: int(re.search(r'slide(\d+)', x.name).group(1))
        )

        slides_data = []
        image_mapping = {}

        for slide_file in slide_files:
            num = int(re.search(r'slide(\d+)', slide_file.name).group(1))

            tree = etree.parse(str(slide_file))
            root = tree.getroot()

            rels_file = rels_dir / f'slide{num}.xml.rels'
            layout = "DEFAULT"
            images = []

            if rels_file.exists():
                with open(rels_file, 'rb') as f:
                    rels_content = f.read()
                    layout = get_layout_name(rels_content, layout_map)
                    images = get_slide_images(rels_content)

            if images:
                image_mapping[num] = images

            # Extract all text
            texts = []
            for t in root.xpath('.//a:t', namespaces=NSMAP):
                if t.text:
                    texts.append(clean_text(t.text))

            # Deduplicate adjacent identical texts
            unique_texts = []
            prev = None
            for t in texts:
                if t != prev and t:
                    unique_texts.append(t)
                    prev = t

            title = unique_texts[0] if unique_texts else ""
            body = '\n'.join(unique_texts[1:]) if len(unique_texts) > 1 else ""

            slides_data.append({
                'number': num,
                'layout': layout,
                'title': title,
                'body': body,
                'images': images,
                'image_count': len(images),
            })

        # Extract images if requested
        images_dir = None
        if extract_images:
            images_dir = output_path.parent / f"{output_path.stem}-images"
            images_dir.mkdir(parents=True, exist_ok=True)

            media_dir = work_dir / 'ppt/media'
            if media_dir.exists():
                for img in media_dir.iterdir():
                    if img.is_file():
                        shutil.copy(img, images_dir / img.name)
                print(f"Extracted {len(list(images_dir.iterdir()))} images to {images_dir.name}/")

            mapping_file = output_path.parent / f"{output_path.stem}-image-mapping.json"
            with open(mapping_file, 'w', encoding='utf-8') as f:
                json.dump({
                    'source': pptx_path.name,
                    'images_dir': str(images_dir.name),
                    'slide_images': {str(k): v for k, v in image_mapping.items()},
                    'total_images': sum(len(v) for v in image_mapping.values()),
                    'slides_with_images': len(image_mapping),
                }, f, indent=2)
            print(f"Created image mapping: {mapping_file.name}")

    # Generate markdown
    md_content = generate_markdown(slides_data, pptx_path.name)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(md_content)

    print(f"Created: {output_path}")
    print(f"Slides: {len(slides_data)}")
    print(f"Slides with images: {len(image_mapping)}")

    return slides_data


def generate_markdown(slides_data: List[Dict[str, Any]], source_name: str) -> str:
    """Generate markdown from slides data."""
    slides_with_images = sum(1 for s in slides_data if s.get('images'))
    total_images = sum(s.get('image_count', 0) for s in slides_data)

    lines = [
        f"# {Path(source_name).stem} - Content Catalog",
        "",
        f"**Source:** {source_name}",
        f"**Total Slides:** {len(slides_data)}",
        f"**Slides with Images:** {slides_with_images}",
        f"**Total Image References:** {total_images}",
        f"**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "",
        "---",
        "",
    ]

    for slide in slides_data:
        lines.extend([
            f"## Slide {slide['number']}",
            f"**Layout:** {slide['layout']}",
            f"**Title:** {slide['title']}",
        ])

        if slide.get('images'):
            lines.append(f"**Images:** {', '.join(slide['images'])}")
        else:
            lines.append("**Images:** (none)")

        lines.extend([
            "",
            "### Content",
            slide['body'] if slide['body'] else "(No additional content)",
            "",
            "---",
            "",
        ])

    return '\n'.join(lines)
