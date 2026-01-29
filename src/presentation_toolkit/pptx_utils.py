"""
PPTX XML Manipulation Utilities

Low-level utilities for working with PowerPoint XML structures.
"""

from lxml import etree
from pathlib import Path
from typing import Optional, List, Tuple, Dict, Any

# OOXML namespaces
NSMAP = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}

# EMU conversion constants
EMU_PER_INCH = 914400
EMU_PER_PT = 12700


def clean_text(text: str) -> str:
    """Remove control characters that break XML.

    Args:
        text: Input text string

    Returns:
        Cleaned text string
    """
    if not text:
        return ""
    cleaned = ''.join(c if ord(c) >= 32 or c in '\n\t' else ' ' for c in str(text))
    return cleaned.strip()


def find_placeholder(root: etree._Element, ph_type: str, idx: Optional[str] = None) -> Optional[etree._Element]:
    """Find shape by placeholder type.

    Args:
        root: XML root element
        ph_type: Placeholder type ('title', 'body', 'subTitle', etc.)
        idx: Optional placeholder index

    Returns:
        Shape element or None
    """
    if idx is not None:
        xpath = f'.//p:sp[.//p:ph[@type="{ph_type}" and @idx="{idx}"]]'
    else:
        xpath = f'.//p:sp[.//p:ph[@type="{ph_type}"]]'

    shapes = root.xpath(xpath, namespaces=NSMAP)
    return shapes[0] if shapes else None


def find_text_boxes(root: etree._Element) -> List[etree._Element]:
    """Find all text box shapes (p:sp with txBox="1").

    Returns list of shapes sorted by position (top-to-bottom, left-to-right).
    """
    xpath = './/p:sp[p:nvSpPr/p:cNvSpPr[@txBox="1"]]'
    shapes = root.xpath(xpath, namespaces=NSMAP)

    def get_position(shape):
        xfrm = shape.find('.//a:xfrm', namespaces=NSMAP)
        if xfrm is not None:
            off = xfrm.find('a:off', namespaces=NSMAP)
            if off is not None:
                return (int(off.get('y', 0)), int(off.get('x', 0)))
        return (0, 0)

    return sorted(shapes, key=get_position)


def find_shape_by_name(root: etree._Element, name_pattern: str) -> Optional[etree._Element]:
    """Find a shape by its name (or partial match).

    Args:
        root: XML root element
        name_pattern: Name to search for (case-insensitive partial match)

    Returns:
        Shape element or None
    """
    pattern = name_pattern.lower()
    for shape in root.xpath('.//p:sp', namespaces=NSMAP):
        nvSpPr = shape.find('.//p:nvSpPr', namespaces=NSMAP)
        if nvSpPr is not None:
            cNvPr = nvSpPr.find('p:cNvPr', namespaces=NSMAP)
            if cNvPr is not None:
                name = cNvPr.get('name', '').lower()
                if pattern in name:
                    return shape
    return None


def get_placeholder_width(shape: Optional[etree._Element]) -> Optional[int]:
    """Get placeholder width in EMUs for font scaling."""
    if shape is None:
        return None

    xfrm = shape.find('.//a:xfrm', namespaces=NSMAP)
    if xfrm is None:
        return None

    ext = xfrm.find('a:ext', namespaces=NSMAP)
    if ext is None:
        return None

    return int(ext.get('cx', 0))


def get_placeholder_dimensions(shape: Optional[etree._Element]) -> Optional[Dict[str, Any]]:
    """Extract dimensions from a shape element.

    Returns:
        dict with x, y, width, height in EMUs
    """
    if shape is None:
        return None

    xfrm = shape.find('.//a:xfrm', namespaces=NSMAP)
    if xfrm is None:
        return None

    off = xfrm.find('a:off', namespaces=NSMAP)
    ext = xfrm.find('a:ext', namespaces=NSMAP)

    if off is None or ext is None:
        return None

    return {
        'x': int(off.get('x', 0)),
        'y': int(off.get('y', 0)),
        'width': int(ext.get('cx', 0)),
        'height': int(ext.get('cy', 0)),
        'width_inches': int(ext.get('cx', 0)) / EMU_PER_INCH,
        'height_inches': int(ext.get('cy', 0)) / EMU_PER_INCH,
    }


def get_text_from_shape(shape: Optional[etree._Element]) -> str:
    """Extract all text from a shape."""
    if shape is None:
        return ""

    texts = []
    for t in shape.xpath('.//a:t', namespaces=NSMAP):
        if t.text:
            texts.append(t.text)
    return ' '.join(texts)


def get_font_size_from_shape(shape: Optional[etree._Element]) -> Optional[int]:
    """Get the first font size found in a shape (in hundredths of a point)."""
    if shape is None:
        return None

    # Check run properties
    for rPr in shape.xpath('.//a:rPr[@sz]', namespaces=NSMAP):
        return int(rPr.get('sz'))

    # Check default run properties
    for defRPr in shape.xpath('.//a:defRPr[@sz]', namespaces=NSMAP):
        return int(defRPr.get('sz'))

    return None


def calculate_font_size(text: str, placeholder_width_emu: int, max_size: int, min_size: int) -> int:
    """Determine font size that fits text in placeholder.

    Args:
        text: The text to fit
        placeholder_width_emu: Width in EMUs
        max_size: Maximum font size in hundredths of a point
        min_size: Minimum font size in hundredths of a point

    Returns:
        Appropriate font size in hundredths of a point
    """
    if not text or not placeholder_width_emu:
        return max_size

    # Approximate: 1 character at 100pt ≈ 70000 EMUs wide (depends on font)
    char_width_at_100pt = 60000  # Conservative estimate

    for size in range(max_size, min_size - 1, -200):  # Step by 2pt
        scale = size / 10000  # Convert hundredths to points ratio
        estimated_width = len(text) * char_width_at_100pt * scale
        if estimated_width <= placeholder_width_emu * 0.95:  # 5% margin
            return size

    return min_size


def set_font_size(text_elem: etree._Element, size_hundredths: int) -> None:
    """Set font size on a text element.

    Args:
        text_elem: The <a:t> element
        size_hundredths: Size in hundredths of a point (e.g., 1400 = 14pt)
    """
    parent = text_elem.getparent()  # This is <a:r>
    if parent is not None:
        rPr = parent.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
        if rPr is None:
            rPr = etree.Element('{http://schemas.openxmlformats.org/drawingml/2006/main}rPr')
            parent.insert(0, rPr)
        rPr.set('sz', str(size_hundredths))


def replace_text_in_shape(shape: Optional[etree._Element], new_text: str, font_size: Optional[int] = None) -> bool:
    """Replace text in any shape element.

    Args:
        shape: Shape element to modify
        new_text: Text to insert
        font_size: Optional font size in hundredths of a point

    Returns:
        True if text was replaced, False otherwise
    """
    if shape is None:
        return False

    text_runs = shape.xpath('.//a:t', namespaces=NSMAP)
    if not text_runs:
        return False

    text_runs[0].text = new_text
    if font_size:
        set_font_size(text_runs[0], font_size)

    # Clear subsequent text runs
    for t in text_runs[1:]:
        t.text = ""

    return True


def replace_text_in_placeholder(root: etree._Element, ph_type: str, new_text: str,
                                 font_size: Optional[int] = None, idx: Optional[str] = None) -> bool:
    """Replace text in a specific placeholder by type.

    Args:
        root: XML root element
        ph_type: Placeholder type ('title', 'body')
        new_text: Text to insert
        font_size: Optional font size in hundredths of a point
        idx: Optional placeholder index

    Returns:
        True if replacement was successful
    """
    shape = find_placeholder(root, ph_type, idx)
    return replace_text_in_shape(shape, new_text, font_size)


def replace_text_in_named_shape(root: etree._Element, shape_name: str, new_text: str,
                                 font_size: Optional[int] = None) -> bool:
    """Replace text in a shape identified by name.

    Args:
        root: XML root element
        shape_name: Name of shape to find
        new_text: Text to insert
        font_size: Optional font size in hundredths of point

    Returns:
        True if text was replaced, False otherwise
    """
    shape = find_shape_by_name(root, shape_name)
    return replace_text_in_shape(shape, new_text, font_size)


def find_largest_picture(root: etree._Element) -> Tuple[Optional[etree._Element], Optional[str], int]:
    """Find the largest p:pic element in a slide (likely the content image area).

    Returns:
        Tuple of (pic_element, rId, area) or (None, None, 0)
    """
    pics = root.xpath('.//p:pic', namespaces=NSMAP)
    largest = None
    largest_rid = None
    largest_area = 0

    for pic in pics:
        spPr = pic.find('p:spPr', namespaces=NSMAP)
        if spPr is None:
            continue

        xfrm = spPr.find('a:xfrm', namespaces=NSMAP)
        if xfrm is None:
            continue

        ext = xfrm.find('a:ext', namespaces=NSMAP)
        if ext is None:
            continue

        cx = int(ext.get('cx', 0))
        cy = int(ext.get('cy', 0))
        area = cx * cy

        if area > largest_area:
            largest_area = area
            largest = pic

            blip = pic.find('.//a:blip', namespaces=NSMAP)
            if blip is not None:
                largest_rid = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')

    return largest, largest_rid, largest_area


def get_next_rid(rels_path: Path) -> str:
    """Get the next available relationship ID from a rels file.

    Returns:
        String like 'rId10'
    """
    try:
        tree = etree.parse(str(rels_path))
        root = tree.getroot()

        max_id = 0
        for rel in root:
            rid = rel.get('Id', '')
            if rid.startswith('rId'):
                try:
                    num = int(rid[3:])
                    max_id = max(max_id, num)
                except ValueError:
                    pass

        return f'rId{max_id + 1}'
    except Exception:
        return 'rId100'  # Safe fallback


def rgb_to_hex(rgb) -> Optional[str]:
    """Convert RGBColor to hex string."""
    if rgb is None:
        return None
    return f'{rgb.red:02x}{rgb.green:02x}{rgb.blue:02x}'.lower()
