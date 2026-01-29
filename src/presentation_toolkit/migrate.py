"""
Presentation Migration Engine

Intelligently maps content to varied template layouts with:
- Content-type detection (stats, quotes, features, bullets, etc.)
- Layout variety tracking (no consecutive repeats)
- Left/right orientation alternation
- GUI block color rotation

Supports:
- PPTX input (auto-extracts content)
- Markdown input (structured slide content)
- CSV input (spreadsheet format)

Outputs brand-compliant PPTX using a provided template.
"""

import re
import os
import csv
import sys
import shutil
import struct
import zipfile
import tempfile
from pathlib import Path
from collections import deque
from typing import Optional, List, Dict, Any, Tuple, Union

from lxml import etree

from .config import BrandConfig, TextCapacity
from .pptx_utils import (
    NSMAP,
    clean_text,
    find_placeholder,
    find_text_boxes,
    find_shape_by_name,
    get_placeholder_width,
    calculate_font_size,
    set_font_size,
    replace_text_in_shape,
    replace_text_in_placeholder,
    replace_text_in_named_shape,
    find_largest_picture,
    get_next_rid,
)


# ============================================================
# CONTENT TYPE DETECTION
# ============================================================

def detect_content_type(slide: Dict[str, Any], config: BrandConfig) -> str:
    """Analyze slide content to determine the best content type.

    Args:
        slide: Slide dict with 'title', 'body', 'number'
        config: Brand configuration with content patterns

    Returns:
        Content type string
    """
    title = slide.get('title', '').strip()
    body = slide.get('body', '').strip()
    combined = f"{title}\n{body}".lower()
    slide_num = slide.get('number', 0)

    patterns = config.content_patterns

    # Check for MULTIPLE statistics (stats dashboard - 4+ stat-like values)
    stat_number_patterns = [
        r'\b\d+%',
        r'\b\d+[KMB]\+?\b',
        r'\b\d{2,}k\b',
        r'\b\$[\d,]+',
        r'\b\d+x\b',
    ]
    stat_count = 0
    for pattern in stat_number_patterns:
        stat_count += len(re.findall(pattern, combined, re.IGNORECASE))

    if stat_count >= 4:
        return 'stats_dashboard'

    # Check statistic patterns
    if 'statistic' in patterns:
        stat_pattern = patterns['statistic']
        for pattern in stat_pattern.title_patterns:
            if re.search(pattern, title, re.IGNORECASE):
                return 'statistic'
        for pattern in stat_pattern.anywhere_patterns:
            if re.search(pattern, combined, re.IGNORECASE):
                return 'statistic'

    # Check quote patterns
    if 'quote' in patterns:
        quote_pattern = patterns['quote']
        for pattern in quote_pattern.title_patterns:
            if re.search(pattern, title):
                return 'quote'
        for pattern in quote_pattern.body_patterns:
            if re.search(pattern, body):
                return 'quote'

    # Check numbered step patterns
    if 'numbered_step' in patterns:
        step_pattern = patterns['numbered_step']
        for pattern in step_pattern.title_patterns:
            if re.search(pattern, title, re.IGNORECASE):
                return 'numbered_step'

    # Check comparison patterns
    if 'comparison' in patterns:
        comp_pattern = patterns['comparison']
        for pattern in comp_pattern.anywhere_patterns:
            if re.search(pattern, combined):
                return 'comparison'

    # Count bullets
    bullet_count = len(re.findall(r'^\s*[-•▪]', body, re.MULTILINE))

    # Check for FULL case study
    has_bullets = bullet_count >= 2
    has_quote = (
        bool(re.search(r'["""].*["""]', combined)) or
        '"' in combined or
        title.strip().startswith('"') or
        title.strip().startswith('"')
    )
    has_why_section = 'why' in combined and ('chosen' in combined or any(
        kw in combined for kw in config.content_patterns.get('case_study', {}).keywords if hasattr(config.content_patterns.get('case_study', {}), 'keywords')
    ))

    # Case study keywords
    case_study_keywords = []
    if 'case_study' in patterns:
        case_study_keywords = patterns['case_study'].keywords

    is_case_study = any(kw in combined for kw in case_study_keywords)

    quote_chars = ['\u201c', '\u201d', '"', "'", '\u2018', '\u2019']
    title_is_quote = any(title.strip().startswith(q) for q in quote_chars)

    if has_bullets and (has_quote or has_why_section):
        return 'case_study_full'

    if title_is_quote and is_case_study:
        return 'case_study_full'

    if is_case_study:
        return 'case_study'

    # Bullet list check
    if bullet_count >= 3:
        return 'bullet_list'

    # Section header check
    if 'section_header' in patterns:
        section_keywords = patterns['section_header'].keywords
        title_lower = title.lower()
        if title_lower.startswith('how ') and len(title) < 30:
            return 'section_header'
        if any(kw in title_lower for kw in section_keywords) and len(title) < 80:
            return 'section_header'

    # Statement check
    if len(title) > 20 and len(title) < 80 and len(body) < 100:
        return 'statement'

    # Feature check
    if 'feature' in patterns:
        feature_keywords = patterns['feature'].keywords
        if any(kw in combined for kw in feature_keywords):
            return 'feature'

    # Default based on content length
    if len(body) > 200:
        content_types_for_long = ['detailed_content', 'feature', 'bullet_list']
        return content_types_for_long[slide_num % len(content_types_for_long)]

    return 'feature'


def detect_slide_position(slide_num: int, total_slides: int) -> str:
    """Determine if slide is opening, closing, or middle."""
    if slide_num <= 2:
        return 'opening'
    elif slide_num >= total_slides - 2:
        return 'closing'
    elif slide_num % 10 == 0 or slide_num % 10 == 1:
        return 'section_break'
    else:
        return 'middle'


# ============================================================
# INTELLIGENT LAYOUT SELECTION
# ============================================================

class LayoutSelector:
    """Selects varied layouts based on content type and recent history."""

    def __init__(self, config: BrandConfig):
        self.config = config
        self.recent_slides = deque(maxlen=3)
        self.last_orientation = None
        self.gui_color_index = 0
        self.gui_colors = config.gui_colors
        self.step_counter = 2

    def get_opposite_orientation(self) -> str:
        """Return opposite of last orientation."""
        if self.last_orientation == 'left':
            return 'right'
        elif self.last_orientation == 'right':
            return 'left'
        return 'center'

    def rotate_gui_color(self) -> str:
        """Get next GUI color in rotation."""
        color = self.gui_colors[self.gui_color_index]
        self.gui_color_index = (self.gui_color_index + 1) % len(self.gui_colors)
        return color

    def select_layout(self, slide: Dict[str, Any], slide_num: int, total_slides: int) -> Tuple[int, str]:
        """Select the best template slide index for this content."""
        content_type = detect_content_type(slide, self.config)
        position = detect_slide_position(slide_num, total_slides)

        candidates = self._get_candidates(content_type, position)
        slide_catalog = self.config.get_all_slide_indices()

        # Filter out recently used slides
        available = []
        for cat in candidates:
            for idx in slide_catalog.get(cat, []):
                if idx not in self.recent_slides:
                    available.append((cat, idx))

        if not available:
            for cat in candidates:
                for idx in slide_catalog.get(cat, []):
                    available.append((cat, idx))

        # Priority for specialized layouts
        specialized_layouts = ['stats_dashboard', 'case_study_full']
        if available and candidates and candidates[0] in specialized_layouts:
            cat, idx = available[0]
            if cat == candidates[0]:
                for orient, cats in self._get_orientations().items():
                    if cat in cats:
                        self._record_selection(idx, orient)
                        return idx, cat

        # Select based on orientation preference
        preferred_orientation = self.get_opposite_orientation()

        for cat, idx in available:
            for orient, cats in self._get_orientations().items():
                if cat in cats and orient == preferred_orientation:
                    self._record_selection(idx, orient)
                    return idx, cat

        if available:
            cat, idx = available[0]
            for orient, cats in self._get_orientations().items():
                if cat in cats:
                    self._record_selection(idx, orient)
                    return idx, cat
            self._record_selection(idx, 'center')
            return idx, cat

        # Ultimate fallback
        self._record_selection(3, 'center')
        return 3, 'default'

    def _get_orientations(self) -> Dict[str, List[str]]:
        """Get orientation mappings."""
        return {
            'left': self.config.orientations.left,
            'right': self.config.orientations.right,
            'center': self.config.orientations.center,
        }

    def _get_candidates(self, content_type: str, position: str) -> List[str]:
        """Get candidate categories for content type and position."""
        slide_catalog = self.config.get_all_slide_indices()

        # Priority for specialized multi-zone layouts
        if content_type == 'stats_dashboard':
            return ['stats_dashboard', 'stat_outline_gui', 'stat_default']

        if content_type == 'case_study_full':
            return ['case_study_full', 'content_image_right']

        # Opening slides
        if position == 'opening':
            if content_type == 'statistic':
                return ['stat_outline_gui', 'stat_coral_filled', 'stat_default']
            return ['title_opening', 'hero_photo', 'statement_center']

        # Closing slides
        if position == 'closing':
            return ['closing_cta', 'closing_statement', 'statement_center']

        # Section breaks
        if position == 'section_break':
            return ['section_divider', 'statement_center']

        # Content-based selection
        if content_type == 'statistic':
            gui_color = self.rotate_gui_color()
            color_mappings = {
                'coral': ['stat_coral_filled', 'stat_outline_gui', 'stat_default'],
                'navy': ['stat_navy_filled', 'stat_photo_bg', 'stat_default'],
                'yellow': ['stat_outline_gui', 'stat_photo_left', 'stat_default'],
            }
            return color_mappings.get(gui_color, ['stat_photo_bg', 'stat_coral_filled', 'stat_default'])

        if content_type == 'quote':
            return ['quote_navy_bg', 'quote_centered', 'quote_default']

        if content_type == 'numbered_step':
            step_map = {2: 'numbered_02', 3: 'numbered_03', 4: 'numbered_04', 5: 'numbered_05'}
            step = self.step_counter
            self.step_counter = (self.step_counter % 4) + 2
            return [step_map.get(step, 'numbered_02'), 'feature_default']

        if content_type == 'comparison':
            return ['two_column']

        if content_type == 'section_header':
            return ['section_divider', 'statement_center', 'hero_photo']

        if content_type == 'case_study':
            gui_color = self.rotate_gui_color()
            if gui_color == 'coral':
                return ['stat_coral_filled', 'feature_coral_gui', 'feature_default']
            elif gui_color == 'navy':
                return ['photo_text_right', 'content_image_left']
            else:
                return ['quote_navy_bg', 'photo_text_left', 'content_image_right']

        if content_type == 'bullet_list':
            gui_color = self.rotate_gui_color()
            color_mappings = {
                'yellow': ['feature_yellow_gui', 'feature_yellow_bg', 'feature_default'],
                'coral': ['feature_coral_gui', 'feature_blue_bg', 'feature_default'],
                'navy': ['content_image_left', 'feature_white_bg', 'feature_default'],
            }
            return color_mappings.get(gui_color, ['feature_blue_bg', 'feature_white_bg', 'feature_default'])

        if content_type == 'statement':
            gui_color = self.rotate_gui_color()
            if gui_color == 'navy':
                return ['hero_photo', 'photo_text_right', 'statement_center']
            elif gui_color == 'coral':
                return ['statement_center', 'closing_statement']
            else:
                return ['statement_center', 'photo_text_left', 'hero_photo']

        if content_type == 'detailed_content':
            gui_color = self.rotate_gui_color()
            color_mappings = {
                'yellow': ['feature_yellow_bg', 'content_image_right', 'feature_default'],
                'coral': ['feature_coral_gui', 'content_image_left', 'feature_default'],
                'navy': ['content_image_left', 'photo_text_right', 'feature_default'],
            }
            return color_mappings.get(gui_color, ['content_image_right', 'feature_white_bg', 'feature_default'])

        # Default feature slides with rotation
        gui_color = self.rotate_gui_color()
        color_mappings = {
            'yellow': ['feature_yellow_bg', 'feature_yellow_gui', 'content_image_right', 'feature_default'],
            'coral': ['feature_coral_gui', 'stat_coral_filled', 'numbered_02', 'feature_default'],
            'navy': ['content_image_left', 'photo_text_right', 'feature_blue_bg', 'feature_default'],
        }
        return color_mappings.get(gui_color, ['content_image_right', 'feature_white_bg', 'feature_blue_bg', 'feature_default'])

    def _record_selection(self, idx: int, orientation: str) -> None:
        """Record the selected slide for history tracking."""
        self.recent_slides.append(idx)
        self.last_orientation = orientation


# ============================================================
# INPUT PARSERS
# ============================================================

def parse_markdown(md_path: Path) -> List[Dict[str, Any]]:
    """Parse markdown file to extract slides."""
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()

    slides = []

    # Try format 1: "## Slide X"
    pattern1 = r'## Slide (\d+)\n(.*?)(?=\n## Slide \d+|\n*$)'
    matches1 = re.findall(pattern1, content, re.DOTALL)

    # Try format 2: "### Slide X - Name"
    pattern2 = r'### Slide (\d+)[^\n]*\n(.*?)(?=\n### Slide \d+|\n---\n*$|\Z)'
    matches2 = re.findall(pattern2, content, re.DOTALL)

    matches = matches1 if len(matches1) >= len(matches2) else matches2

    for num, slide_content in matches:
        slide = {'number': int(num), 'layout': 'DEFAULT', 'title': '', 'body': ''}

        layout_match = re.search(r'\*\*Layout:\*\*\s*(\w+(?:_\w+)*)', slide_content)
        if layout_match:
            slide['layout'] = layout_match.group(1).upper()

        title_match = re.search(r'\*\*Title:\*\*\s*([^\n]*?)(?=\n|$)', slide_content)
        if title_match:
            title_text = title_match.group(1).strip()
            if not title_text.startswith('**') and title_text not in ['0', '1', '2']:
                slide['title'] = clean_text(title_text)

        body_text = ''

        content_match = re.search(r'### Content\n(.*?)(?=\n---|\Z)', slide_content, re.DOTALL)
        if content_match:
            body_text = content_match.group(1).strip()

        if not body_text:
            text_match = re.search(r'\*\*Text:\*\*\s*(.*?)(?=\n\*\*[A-Z]|\n---|\Z)', slide_content, re.DOTALL)
            if text_match:
                body_text = text_match.group(1).strip()

        if not body_text:
            bullets = re.findall(r'^\s*[-•]\s+(.+)$', slide_content, re.MULTILINE)
            if bullets:
                body_text = '\n'.join(bullets)

        if body_text:
            lines = [clean_text(l) for l in body_text.split('\n') if l.strip()]
            title_line = slide['title'].split('\n')[0].strip() if slide['title'] else ''
            filtered_lines = []
            seen = set()

            skip_phrases = ['Photo:', 'Image:', 'https://', 'http://']

            for l in lines:
                if l.startswith('**') or l in ['0', '1', '2']:
                    continue
                if l == title_line:
                    continue
                if any(skip in l for skip in skip_phrases):
                    continue
                if l in seen:
                    continue
                if len(l) < 3:
                    continue
                seen.add(l)
                filtered_lines.append(l)

            slide['body'] = '\n'.join(filtered_lines[:15])

        if not slide['title'] and slide['body']:
            body_lines = slide['body'].split('\n')
            if body_lines:
                slide['title'] = body_lines[0]
                slide['body'] = '\n'.join(body_lines[1:])

        slides.append(slide)

    return slides


def parse_csv(csv_path: Path) -> List[Dict[str, Any]]:
    """Parse CSV file to extract slides."""
    slides = []
    with open(csv_path, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            slide = {
                'number': int(row.get('slide_number', len(slides) + 1)),
                'layout': row.get('layout', 'DEFAULT').upper(),
                'title': clean_text(row.get('title', '')),
                'body': clean_text(row.get('body', '')),
            }
            slides.append(slide)
    return slides


def parse_pptx(pptx_path: Path, image_output_dir: Optional[Path] = None) -> List[Dict[str, Any]]:
    """Extract content and images from PPTX file."""
    slides = []

    if image_output_dir:
        image_output_dir = Path(image_output_dir)
        image_output_dir.mkdir(parents=True, exist_ok=True)

    total_images_extracted = 0

    with tempfile.TemporaryDirectory() as tmpdir:
        work_dir = Path(tmpdir)

        with zipfile.ZipFile(pptx_path, 'r') as zf:
            zf.extractall(work_dir)

        slides_dir = work_dir / 'ppt/slides'
        media_dir = work_dir / 'ppt/media'

        slide_files = sorted(
            [f for f in slides_dir.glob('slide*.xml')],
            key=lambda x: int(re.search(r'slide(\d+)', x.name).group(1))
        )

        for slide_file in slide_files:
            num = int(re.search(r'slide(\d+)', slide_file.name).group(1))
            tree = etree.parse(str(slide_file))
            root = tree.getroot()

            texts = []
            for t in root.xpath('.//a:t', namespaces=NSMAP):
                if t.text:
                    texts.append(clean_text(t.text))

            layout = 'DEFAULT'
            extracted_images = []
            rels_file = slides_dir / f'_rels/slide{num}.xml.rels'

            if rels_file.exists():
                rels_tree = etree.parse(str(rels_file))
                for rel in rels_tree.getroot():
                    target = rel.get('Target', '')
                    rel_type = rel.get('Type', '').split('/')[-1]

                    if 'slideLayout' in target:
                        layout_num = re.search(r'slideLayout(\d+)', target)
                        if layout_num:
                            layout = 'DEFAULT'

                    if rel_type == 'image' and image_output_dir:
                        image_name = os.path.basename(target)
                        source_path = media_dir / image_name

                        if source_path.exists():
                            try:
                                with open(source_path, 'rb') as f:
                                    data = f.read(32)

                                width, height = 0, 0
                                ext = source_path.suffix.lower()

                                if ext == '.png' and data[:8] == b'\x89PNG\r\n\x1a\n':
                                    width = struct.unpack('>I', data[16:20])[0]
                                    height = struct.unpack('>I', data[20:24])[0]
                                elif ext in ['.jpg', '.jpeg']:
                                    width, height = 800, 600

                                if width >= 100 and height >= 100:
                                    dest_filename = f"slide{num}_img{len(extracted_images)}{ext}"
                                    dest_path = image_output_dir / dest_filename
                                    shutil.copy2(source_path, dest_path)

                                    extracted_images.append({
                                        'path': str(dest_path),
                                        'width': width,
                                        'height': height,
                                        'ext': ext[1:]
                                    })
                                    total_images_extracted += 1
                            except Exception:
                                pass

            slide = {
                'number': num,
                'layout': layout,
                'title': texts[0] if texts else '',
                'body': '\n'.join(texts[1:5]) if len(texts) > 1 else '',
                'images': extracted_images,
                'image_count': len(extracted_images)
            }
            slides.append(slide)

    if image_output_dir:
        print(f"  Extracted {total_images_extracted} images to {image_output_dir}")

    return slides


def parse_pdf(pdf_path: Path, image_output_dir: Optional[Path] = None) -> List[Dict[str, Any]]:
    """Extract content and images from PDF file."""
    try:
        import fitz  # PyMuPDF
    except ImportError:
        raise ImportError("PyMuPDF (fitz) is required for PDF parsing. Install with: pip install PyMuPDF")

    slides = []
    doc = fitz.open(pdf_path)

    print(f"  PDF has {len(doc)} pages")

    if image_output_dir:
        image_output_dir = Path(image_output_dir)
        image_output_dir.mkdir(parents=True, exist_ok=True)

    total_images_extracted = 0

    for page_num, page in enumerate(doc, 1):
        text = page.get_text().strip()
        lines = [line.strip() for line in text.split('\n') if line.strip()]

        title = ''
        body_lines = []

        for i, line in enumerate(lines):
            if not title and len(line) > 2:
                title = line
            elif title:
                body_lines.append(line)

        body = '\n'.join(body_lines)

        images = page.get_images()
        image_count = len(images)
        extracted_images = []

        if image_output_dir and images:
            for img_idx, img in enumerate(images):
                xref = img[0]
                try:
                    base_image = doc.extract_image(xref)
                    if base_image:
                        image_bytes = base_image["image"]
                        image_ext = base_image["ext"]

                        width = base_image.get("width", 0)
                        height = base_image.get("height", 0)

                        if width >= 100 and height >= 100:
                            image_filename = f"page{page_num}_img{img_idx}.{image_ext}"
                            image_path = image_output_dir / image_filename

                            with open(image_path, "wb") as img_file:
                                img_file.write(image_bytes)

                            extracted_images.append({
                                'path': str(image_path),
                                'width': width,
                                'height': height,
                                'ext': image_ext
                            })
                            total_images_extracted += 1
                except Exception:
                    pass

        text_chars = len(text)
        has_minimal_text = text_chars < 50

        slide = {
            'number': page_num,
            'layout': 'DEFAULT',
            'title': clean_text(title),
            'body': clean_text(body),
            'image_count': image_count,
            'images': extracted_images,
            '_extraction_notes': []
        }

        if has_minimal_text and image_count > 0:
            slide['_extraction_notes'].append(
                f"WARNING: Only {text_chars} chars extracted but {image_count} images found."
            )

        if not title and not body:
            slide['_extraction_notes'].append(
                "WARNING: No text extracted. This may be a title slide with image-based text."
            )

        slides.append(slide)

    doc.close()

    if image_output_dir:
        print(f"  Extracted {total_images_extracted} images to {image_output_dir}")

    return slides


def detect_and_parse(input_path: Union[str, Path], image_output_dir: Optional[Path] = None) -> List[Dict[str, Any]]:
    """Detect input format and parse accordingly."""
    path = Path(input_path)
    suffix = path.suffix.lower()

    if suffix == '.md':
        print(f"Detected Markdown input: {path.name}")
        return parse_markdown(path)
    elif suffix == '.csv':
        print(f"Detected CSV input: {path.name}")
        return parse_csv(path)
    elif suffix in ['.pptx', '.ppt']:
        print(f"Detected PPTX input: {path.name}")
        return parse_pptx(path, image_output_dir)
    elif suffix == '.pdf':
        print(f"Detected PDF input: {path.name}")
        return parse_pdf(path, image_output_dir)
    else:
        raise ValueError(f"Unsupported input format: {suffix}")


# ============================================================
# MULTI-ZONE LAYOUT POPULATION
# ============================================================

def parse_stats_content(title: str, body: str) -> List[Dict[str, str]]:
    """Parse content into stats dashboard zones."""
    stats = []
    combined = f"{title}\n{body}"

    stat_patterns = [
        r'^[\d,.$]+[%KMBkmb+]*$',
        r'^\d+[\d,.]*\s*[%KMBkmb+]',
        r'^(Millions?|Billions?|Thousands?|Hundreds?)\b',
    ]

    lines = combined.split('\n')
    current_number = None
    current_labels = []

    for line in lines:
        line = line.strip()
        if not line:
            continue

        is_number = False
        for pattern in stat_patterns:
            if re.match(pattern, line, re.IGNORECASE):
                is_number = True
                break

        if not is_number and re.match(r'^\d', line) and len(line) < 20:
            is_number = True

        if is_number:
            if current_number:
                label = ' '.join(current_labels) if current_labels else ''
                stats.append({'number': current_number, 'label': label})
                current_labels = []
            current_number = line
        elif current_number:
            if len(line) < 50 and not any(c in line for c in '.!?'):
                current_labels.append(line)
            else:
                label = ' '.join(current_labels) if current_labels else ''
                stats.append({'number': current_number, 'label': label})
                current_number = None
                current_labels = []

    if current_number:
        label = ' '.join(current_labels) if current_labels else ''
        stats.append({'number': current_number, 'label': label})

    return stats[:6]


def parse_case_study_content(title: str, body: str) -> Dict[str, str]:
    """Parse content into case study zones."""
    zones = {
        'company_name': title,
        'description': '',
        'bullets': '',
        'quote': '',
        'attribution': ''
    }

    lines = body.split('\n')
    in_bullets = False
    description_lines = []
    bullet_lines = []

    for line in lines:
        line = line.strip()
        if not line:
            continue

        if line.startswith('"') or line.startswith('"') or line.startswith("'"):
            quote_text = line
            if '—' in quote_text or ' - ' in quote_text:
                parts = re.split(r'\s*[—-]\s*', quote_text, 1)
                zones['quote'] = parts[0].strip(' ""\'"')
                if len(parts) > 1:
                    zones['attribution'] = '— ' + parts[1]
            else:
                zones['quote'] = quote_text.strip(' ""\'"')
            continue

        if line.startswith('—') or line.startswith('- '):
            zones['attribution'] = line
            continue

        if 'why' in line.lower():
            in_bullets = True
            continue

        if line.startswith('•') or line.startswith('-') or line.startswith('*'):
            in_bullets = True
            bullet_lines.append(line)
            continue

        if not in_bullets:
            description_lines.append(line)
        else:
            if len(line) < 100:
                bullet_lines.append('• ' + line)

    zones['description'] = ' '.join(description_lines)
    zones['bullets'] = '\n'.join(bullet_lines)

    return zones


def populate_stats_dashboard(slide_path: Path, title: str, body: str) -> etree._ElementTree:
    """Populate a stats dashboard slide with parsed statistics."""
    tree = etree.parse(str(slide_path))
    root = tree.getroot()

    stats = parse_stats_content(title, body)

    replace_text_in_placeholder(root, 'title', title, 3600)

    for i, stat in enumerate(stats, 1):
        number_name = f'Stat{i}_Number'
        label_name = f'Stat{i}_Label'

        replace_text_in_named_shape(root, number_name, stat['number'], 7200)
        replace_text_in_named_shape(root, label_name, stat['label'], 1800)

    return tree


def populate_case_study_full(slide_path: Path, title: str, body: str) -> etree._ElementTree:
    """Populate a case study slide with parsed content zones."""
    tree = etree.parse(str(slide_path))
    root = tree.getroot()

    zones = parse_case_study_content(title, body)

    replace_text_in_placeholder(root, 'title', zones['company_name'], 3600)
    replace_text_in_placeholder(root, 'body', zones['description'], 1600, idx='1')
    replace_text_in_placeholder(root, 'body', zones['bullets'], 1400, idx='2')

    if zones['quote']:
        quote_text = f'"{zones["quote"]}"'
        replace_text_in_named_shape(root, 'Quote', quote_text, 2400)

    if zones['attribution']:
        replace_text_in_named_shape(root, 'Attribution', zones['attribution'], 1400)

    return tree


# ============================================================
# IMAGE INSERTION
# ============================================================

def insert_image_in_slide(
    slide_xml_path: Path,
    slide_rels_path: Path,
    image_path: Union[str, Path],
    media_dir: Path,
    next_rid: str
) -> bool:
    """Insert an image into a slide by replacing the largest existing picture."""
    tree = etree.parse(str(slide_xml_path))
    root = tree.getroot()

    pic, old_rid, area = find_largest_picture(root)

    if pic is None:
        return False

    image_path = Path(image_path)
    new_image_name = f"image_user_{next_rid}.{image_path.suffix.lstrip('.')}"
    dest_path = media_dir / new_image_name
    shutil.copy2(image_path, dest_path)

    blip = pic.find('.//a:blip', namespaces=NSMAP)
    if blip is not None:
        blip.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', next_rid)

    rels_tree = etree.parse(str(slide_rels_path))
    rels_root = rels_tree.getroot()

    new_rel = etree.SubElement(rels_root, 'Relationship')
    new_rel.set('Id', next_rid)
    new_rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
    new_rel.set('Target', f'../media/{new_image_name}')

    tree.write(str(slide_xml_path), xml_declaration=True, encoding='UTF-8', standalone=True)
    rels_tree.write(str(slide_rels_path), xml_declaration=True, encoding='UTF-8', standalone=True)

    return True


# ============================================================
# TEXT REPLACEMENT
# ============================================================

def replace_text_in_slide(
    slide_path: Path,
    new_title: str,
    new_body: str,
    title_font_size: int = 2400,
    body_font_size: int = 1400
) -> etree._ElementTree:
    """Replace text in slide XML with proper placeholder targeting."""
    tree = etree.parse(str(slide_path))
    root = tree.getroot()

    title_shape = find_placeholder(root, 'title')
    body_shape = find_placeholder(root, 'body')

    if body_shape is None:
        body_shape = find_placeholder(root, 'body', idx='1')

    text_boxes = None
    if title_shape is None or body_shape is None:
        text_boxes = find_text_boxes(root)
        if text_boxes:
            if title_shape is None and len(text_boxes) >= 1:
                title_shape = text_boxes[0]
            if body_shape is None and len(text_boxes) >= 2:
                body_shape = text_boxes[1]

    if title_shape is not None and new_title:
        title_width = get_placeholder_width(title_shape)
        if title_width:
            title_font_size = calculate_font_size(new_title, title_width, title_font_size, 1800)

    if body_shape is not None and new_body:
        body_width = get_placeholder_width(body_shape)
        if body_width:
            body_font_size = calculate_font_size(new_body, body_width, body_font_size, 1200)

    title_replaced = False
    body_replaced = False

    if new_title:
        title_replaced = replace_text_in_placeholder(root, 'title', new_title, title_font_size)
        if not title_replaced and text_boxes and len(text_boxes) >= 1:
            title_replaced = replace_text_in_shape(text_boxes[0], new_title, title_font_size)
        if not title_replaced:
            for t_elem in root.xpath('.//a:t', namespaces=NSMAP):
                if t_elem.text and len(t_elem.text.strip()) > 2:
                    t_elem.text = new_title
                    set_font_size(t_elem, title_font_size)
                    title_replaced = True
                    break

    if new_body:
        body_replaced = replace_text_in_placeholder(root, 'body', new_body, body_font_size, idx='1')
        if not body_replaced:
            body_replaced = replace_text_in_placeholder(root, 'body', new_body, body_font_size)
        if not body_replaced and text_boxes and len(text_boxes) >= 2:
            body_replaced = replace_text_in_shape(text_boxes[1], new_body, body_font_size)
        if not body_replaced:
            count = 0
            for t_elem in root.xpath('.//a:t', namespaces=NSMAP):
                if t_elem.text and len(t_elem.text.strip()) > 2:
                    count += 1
                    if count == 2:
                        t_elem.text = new_body
                        set_font_size(t_elem, body_font_size)
                        body_replaced = True
                        break

    return tree


# ============================================================
# MIGRATION ENGINE
# ============================================================

def migrate_presentation(
    slides: List[Dict[str, Any]],
    output_path: Union[str, Path],
    config: BrandConfig,
    template_path: Union[str, Path],
    insert_images: bool = True
) -> Path:
    """Create migrated presentation from slides data with intelligent layout selection.

    Args:
        slides: List of slide dicts with 'title', 'body', and optionally 'images'
        output_path: Path for output PPTX file
        config: Brand configuration
        template_path: Path to template PPTX
        insert_images: Whether to insert extracted images into slides

    Returns:
        Path to created presentation
    """
    template_path = Path(template_path)
    output_path = Path(output_path)

    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")

    print(f"\nMigrating {len(slides)} slides...")
    print(f"Template: {template_path}")
    print(f"Output: {output_path}")

    selector = LayoutSelector(config)

    work_dir = Path(tempfile.mkdtemp())
    template_dir = work_dir / 'template'
    output_dir = work_dir / 'output'

    layout_assignments = []

    try:
        with zipfile.ZipFile(template_path, 'r') as zf:
            zf.extractall(template_dir)

        shutil.copytree(template_dir, output_dir)

        slides_dir = output_dir / 'ppt/slides'
        rels_dir = output_dir / 'ppt/slides/_rels'

        total_slides = len(slides)
        for i, slide in enumerate(slides):
            new_num = i + 1

            template_slide_num, category = selector.select_layout(slide, new_num, total_slides)

            content_type = detect_content_type(slide, config)
            layout_assignments.append({
                'slide': new_num,
                'template': template_slide_num,
                'category': category,
                'content_type': content_type,
                'title_preview': slide['title'][:40] + '...' if len(slide['title']) > 40 else slide['title']
            })

            src_slide = template_dir / f'ppt/slides/slide{template_slide_num}.xml'
            src_rels = template_dir / f'ppt/slides/_rels/slide{template_slide_num}.xml.rels'

            if not src_slide.exists():
                src_slide = template_dir / 'ppt/slides/slide3.xml'
                src_rels = template_dir / 'ppt/slides/_rels/slide3.xml.rels'
                print(f"  Warning: Template slide {template_slide_num} not found, using fallback")

            dst_slide = slides_dir / f'slide{new_num}.xml'
            dst_rels = rels_dir / f'slide{new_num}.xml.rels'

            shutil.copy(src_slide, dst_slide)
            if src_rels.exists():
                shutil.copy(src_rels, dst_rels)

            capacity = config.get_text_capacity(category)
            title_max = capacity.title_max_chars
            body_max = capacity.body_max_chars
            title_pt = capacity.title_font_size
            body_pt = capacity.body_font_size

            if content_type == 'stats_dashboard' and category == 'stats_dashboard':
                tree = populate_stats_dashboard(dst_slide, slide['title'], slide['body'])
                tree.write(str(dst_slide), xml_declaration=True, encoding='UTF-8', standalone=True)
            elif content_type == 'case_study_full' and category == 'case_study_full':
                tree = populate_case_study_full(dst_slide, slide['title'], slide['body'])
                tree.write(str(dst_slide), xml_declaration=True, encoding='UTF-8', standalone=True)
            else:
                new_title = slide['title'].replace('\n', ' ')[:title_max]
                new_body = slide['body'].replace('\n', ' ')[:body_max]

                if len(slide['body']) > body_max:
                    body_pt = max(1200, body_pt - 200)

                tree = replace_text_in_slide(dst_slide, new_title, new_body, title_pt, body_pt)
                tree.write(str(dst_slide), xml_declaration=True, encoding='UTF-8', standalone=True)

            slide_images = slide.get('images', [])
            if insert_images and slide_images:
                largest_image = max(slide_images, key=lambda x: x.get('width', 0) * x.get('height', 0))
                media_dir = output_dir / 'ppt/media'

                if dst_rels.exists():
                    next_rid = get_next_rid(dst_rels)
                    image_inserted = insert_image_in_slide(
                        dst_slide, dst_rels, largest_image['path'], media_dir, next_rid
                    )
                    if image_inserted:
                        layout_assignments[-1]['image_inserted'] = True

            if new_num % 20 == 0:
                print(f"  Processed {new_num}/{total_slides} slides...")

        update_package_structure(output_dir, len(slides))

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root_path, dirs, files in os.walk(output_dir):
                for file in files:
                    file_path = Path(root_path) / file
                    arc_path = file_path.relative_to(output_dir)
                    zf.write(file_path, arc_path)

        print(f"\nMigration complete: {output_path}")

        print_layout_report(layout_assignments)

        return output_path

    finally:
        shutil.rmtree(work_dir)


def print_layout_report(assignments: List[Dict[str, Any]]) -> None:
    """Print a summary of layout variety used."""
    print("\n" + "=" * 60)
    print("LAYOUT VARIETY REPORT")
    print("=" * 60)

    template_counts = {}
    category_counts = {}
    content_type_counts = {}

    for a in assignments:
        template_counts[a['template']] = template_counts.get(a['template'], 0) + 1
        category_counts[a['category']] = category_counts.get(a['category'], 0) + 1
        content_type_counts[a['content_type']] = content_type_counts.get(a['content_type'], 0) + 1

    print(f"\nUnique templates used: {len(template_counts)}")
    print(f"Unique categories used: {len(category_counts)}")

    print("\nTemplate distribution:")
    for idx, count in sorted(template_counts.items(), key=lambda x: -x[1])[:10]:
        print(f"  Slide {idx:2d}: {count} times")

    print("\nContent types detected:")
    for ct, count in sorted(content_type_counts.items(), key=lambda x: -x[1]):
        print(f"  {ct}: {count}")

    print("\nFirst 10 slide assignments:")
    for a in assignments[:10]:
        print(f"  Slide {a['slide']:3d} -> Template {a['template']:2d} ({a['category']}) - {a['title_preview']}")

    consecutive_repeats = 0
    for i in range(1, len(assignments)):
        if assignments[i]['template'] == assignments[i-1]['template']:
            consecutive_repeats += 1

    print(f"\nConsecutive template repeats: {consecutive_repeats}")
    if consecutive_repeats == 0:
        print("  Excellent variety!")
    elif consecutive_repeats < 5:
        print("  Good variety with minor repeats")
    else:
        print("  Consider improving content type detection")


def update_package_structure(output_dir: Path, num_slides: int) -> None:
    """Update PPTX internal structure for new slide count."""
    ct_path = output_dir / '[Content_Types].xml'
    ct_tree = etree.parse(str(ct_path))
    ct_root = ct_tree.getroot()

    ns = {'ct': 'http://schemas.openxmlformats.org/package/2006/content-types'}
    for override in ct_root.xpath('.//ct:Override[contains(@PartName, "/ppt/slides/slide")]', namespaces=ns):
        ct_root.remove(override)

    for i in range(1, num_slides + 1):
        override = etree.SubElement(ct_root, '{http://schemas.openxmlformats.org/package/2006/content-types}Override')
        override.set('PartName', f'/ppt/slides/slide{i}.xml')
        override.set('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml')

    ct_tree.write(str(ct_path), xml_declaration=True, encoding='UTF-8', standalone=True)

    rels_path = output_dir / 'ppt/_rels/presentation.xml.rels'
    rels_tree = etree.parse(str(rels_path))
    rels_root = rels_tree.getroot()

    rels_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'

    for rel in list(rels_root):
        if 'slides/slide' in rel.get('Target', ''):
            rels_root.remove(rel)

    max_rid = max(int(rel.get('Id', 'rId0').replace('rId', '')) for rel in rels_root)

    for i in range(1, num_slides + 1):
        rel = etree.SubElement(rels_root, f'{{{rels_ns}}}Relationship')
        rel.set('Id', f'rId{max_rid + i}')
        rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide')
        rel.set('Target', f'slides/slide{i}.xml')

    rels_tree.write(str(rels_path), xml_declaration=True, encoding='UTF-8', standalone=True)

    pres_path = output_dir / 'ppt/presentation.xml'
    pres_tree = etree.parse(str(pres_path))
    pres_root = pres_tree.getroot()

    pres_ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
    sld_id_lst = pres_root.find('.//p:sldIdLst', pres_ns)

    if sld_id_lst is not None:
        for child in list(sld_id_lst):
            sld_id_lst.remove(child)

        for i in range(1, num_slides + 1):
            sld_id = etree.SubElement(sld_id_lst, f'{{{pres_ns["p"]}}}sldId')
            sld_id.set('id', str(255 + i))
            sld_id.set(f'{{{NSMAP["r"]}}}id', f'rId{max_rid + i}')

    pres_tree.write(str(pres_path), xml_declaration=True, encoding='UTF-8', standalone=True)
