import io
from typing import List, Optional
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PIL import Image
from core.models import SlideElement, ParagraphData, TextRunData
from core.ocr_engine import OCREngine
from core.chart_smartart import VectorChartAndSmartArtRenderer

class PPTXExtractor:
    def __init__(
        self,
        pptx_source,
        enable_ocr: bool = True,
        render_charts: bool = True,
        tesseract_cmd: Optional[str] = None
    ):
        """pptx_source can be a file path (str) or a BytesIO stream."""
        self.prs = Presentation(pptx_source)
        self.slide_w = float(self.prs.slide_width)
        self.slide_h = float(self.prs.slide_height)
        self.ocr_engine = OCREngine(enabled=enable_ocr, tesseract_cmd=tesseract_cmd)
        self.render_charts = render_charts

    def extract_all_slides(self) -> List[List[SlideElement]]:
        all_slides_elements = []
        for slide in self.prs.slides:
            elements = []
            title_shape = slide.shapes.title
            for shape in slide.shapes:
                self._extract_shape(shape, title_shape, elements)
            all_slides_elements.append(elements)
        return all_slides_elements

    @staticmethod
    def _try_get_image_blob(shape) -> Optional[bytes]:
        """
        Extracts binary image data whether the shape is a standard PICTURE (13),
        a Picture/Content PLACEHOLDER (14), or a picture-filled shape.
        """
        try:
            if hasattr(shape, "image") and shape.image is not None:
                return shape.image.blob
        except Exception:
            pass
        return None

    def _extract_shape(self, shape, title_shape, elements: List[SlideElement], offset_x=0.0, offset_y=0.0):
        # 1. Unpack grouped shapes recursively
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for sub_shape in shape.shapes:
                self._extract_shape(sub_shape, title_shape, elements)
            return

        left_r = max(0.0, (float(getattr(shape, "left", 0)) + offset_x) / self.slide_w)
        top_r = max(0.0, (float(getattr(shape, "top", 0)) + offset_y) / self.slide_h)
        width_r = max(0.05, float(getattr(shape, "width", 0)) / self.slide_w)
        height_r = max(0.05, float(getattr(shape, "height", 0)) / self.slide_h)

        # 2. Extract Native Vector Charts
        if getattr(shape, "has_chart", False) and self.render_charts:
            png_bytes, px_size, chart_table, chart_title = VectorChartAndSmartArtRenderer.extract_chart(shape)
            if png_bytes or chart_table:
                elements.append(SlideElement(
                    element_type="chart",
                    left_ratio=left_r, top_ratio=top_r,
                    width_ratio=width_r, height_ratio=height_r,
                    image_bytes=png_bytes,
                    image_size_px=px_size,
                    table_data=chart_table,
                    chart_title=chart_title
                ))
                return

        # 3. Extract SmartArt Diagrams
        smartart_paras = VectorChartAndSmartArtRenderer.extract_smartart_paragraphs(shape)
        if smartart_paras:
            elements.append(SlideElement(
                element_type="smartart",
                left_ratio=left_r, top_ratio=top_r,
                width_ratio=width_r, height_ratio=height_r,
                paragraphs=smartart_paras
            ))
            return

        # 4. Extract Tables
        if getattr(shape, "has_table", False):
            table_matrix = []
            for row in shape.table.rows:
                row_cells = [cell.text.strip() for cell in row.cells]
                if any(row_cells):
                    table_matrix.append(row_cells)
            if table_matrix:
                elements.append(SlideElement(
                    element_type="table",
                    left_ratio=left_r, top_ratio=top_r,
                    width_ratio=width_r, height_ratio=height_r,
                    table_data=table_matrix
                ))
                return

        # 5. Extract Pictures (Works for BOTH MSO_SHAPE_TYPE.PICTURE and PLACEHOLDER images!)
        img_bytes = self._try_get_image_blob(shape)
        if img_bytes is not None:
            try:
                with Image.open(io.BytesIO(img_bytes)) as img:
                    if img.format not in ["PNG", "JPEG", "GIF", "BMP"]:
                        buf = io.BytesIO()
                        img.convert("RGB").save(buf, format="PNG")
                        img_bytes = buf.getvalue()
                    px_size = img.size

                ocr_extracted = self.ocr_engine.extract_text_from_image(img_bytes)
                elements.append(SlideElement(
                    element_type="image",
                    left_ratio=left_r, top_ratio=top_r,
                    width_ratio=width_r, height_ratio=height_r,
                    image_bytes=img_bytes,
                    image_size_px=px_size,
                    ocr_text=ocr_extracted
                ))
                return
            except Exception:
                pass

        # 6. Extract Standard Text Frames
        if getattr(shape, "has_text_frame", False):
            paragraphs_data = []
            for p in shape.text_frame.paragraphs:
                full_text = p.text.strip()
                if not full_text:
                    continue
                runs_data = []
                for r in p.runs:
                    if not r.text:
                        continue
                    rgb = None
                    try:
                        if r.font.color and r.font.color.rgb:
                            c = r.font.color.rgb
                            rgb = (c[0], c[1], c[2])
                    except Exception:
                        rgb = None

                    f_size = r.font.size.pt if r.font.size else None
                    runs_data.append(TextRunData(
                        text=r.text,
                        bold=bool(r.font.bold),
                        italic=bool(r.font.italic),
                        underline=bool(r.font.underline),
                        font_size=f_size,
                        rgb_color=rgb
                    ))

                if not runs_data:
                    runs_data.append(TextRunData(text=full_text))

                paragraphs_data.append(ParagraphData(
                    runs=runs_data,
                    level=p.level,
                    is_bullet=(p.level > 0 or full_text.startswith(("•", "-", "▪")))
                ))

            if paragraphs_data:
                el_type = "title" if (title_shape and shape == title_shape) else "text"
                elements.append(SlideElement(
                    element_type=el_type,
                    left_ratio=left_r, top_ratio=top_r,
                    width_ratio=width_r, height_ratio=height_r,
                    paragraphs=paragraphs_data
                ))
