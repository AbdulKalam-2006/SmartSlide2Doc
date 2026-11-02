from dataclasses import dataclass, field
from typing import List, Optional, Tuple

@dataclass
class TextRunData:
    text: str
    bold: bool = False
    italic: bool = False
    underline: bool = False
    font_size: Optional[float] = None
    rgb_color: Optional[Tuple[int, int, int]] = None

@dataclass
class ParagraphData:
    runs: List[TextRunData] = field(default_factory=list)
    level: int = 0
    is_bullet: bool = False

@dataclass
class SlideElement:
    element_type: str  # 'title', 'text', 'image', 'table', 'chart', 'smartart'
    left_ratio: float  # 0.0 to 1.0 relative to slide width
    top_ratio: float   # 0.0 to 1.0 relative to slide height
    width_ratio: float
    height_ratio: float
    paragraphs: List[ParagraphData] = field(default_factory=list)
    image_bytes: Optional[bytes] = None
    image_size_px: Optional[Tuple[int, int]] = None
    ocr_text: Optional[str] = None
    table_data: Optional[List[List[str]]] = None
    chart_title: Optional[str] = None

@dataclass
class LayoutRow:
    """Represents either a full-width single element or a 2-column side-by-side row."""
    is_two_column: bool = False
    left_elements: List[SlideElement] = field(default_factory=list)
    right_elements: List[SlideElement] = field(default_factory=list)
    single_element: Optional[SlideElement] = None

@dataclass
class AnalyzedSlide:
    slide_number: int
    orientation: str  # 'portrait' or 'landscape'
    title: Optional[str] = None
    rows: List[LayoutRow] = field(default_factory=list)
