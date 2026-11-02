import io
from typing import List
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from core.models import AnalyzedSlide, SlideElement, LayoutRow

class CompactDocxBuilder:
    def __init__(
        self,
        max_image_height_in: float = 3.2,
        show_slide_dividers: bool = True,
        include_chart_tables: bool = True,
        include_ocr_captions: bool = True
    ):
        self.doc = Document()
        self.max_image_height_in = max_image_height_in
        self.show_slide_dividers = show_slide_dividers
        self.include_chart_tables = include_chart_tables
        self.include_ocr_captions = include_ocr_captions
        self.current_orientation = None
        self._configure_base_styles()

    def _configure_base_styles(self):
        normal_style = self.doc.styles["Normal"]
        normal_style.font.name = "Calibri"
        normal_style.font.size = Pt(10.5)
        normal_style.paragraph_format.space_before = Pt(1)
        normal_style.paragraph_format.space_after = Pt(2)
        normal_style.paragraph_format.line_spacing = 1.12

    def build(self, analyzed_slides: List[AnalyzedSlide]) -> bytes:
        for idx, slide in enumerate(analyzed_slides):
            self._ensure_orientation(slide.orientation, is_first=(idx == 0))

            if self.show_slide_dividers or slide.title:
                self._render_slide_header(slide.slide_number, slide.title)

            avail_width_in = 9.5 if slide.orientation == "landscape" else 7.0

            for row in slide.rows:
                if row.is_two_column:
                    self._render_two_column_row(row, avail_width_in)
                elif row.single_element:
                    self._render_element_into_container(
                        self.doc, row.single_element, max_width_in=avail_width_in
                    )

        output_stream = io.BytesIO()
        self.doc.save(output_stream)
        return output_stream.getvalue()

    def _ensure_orientation(self, target_orientation: str, is_first: bool):
        if is_first:
            section = self.doc.sections[0]
            self._apply_section_geometry(section, target_orientation)
            self.current_orientation = target_orientation
            return

        if target_orientation != self.current_orientation:
            section = self.doc.add_section(WD_SECTION.NEW_PAGE)
            self._apply_section_geometry(section, target_orientation)
            self.current_orientation = target_orientation

    @staticmethod
    def _apply_section_geometry(section, orientation: str):
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

        if orientation == "landscape":
            section.orientation = WD_ORIENT.LANDSCAPE
            section.page_width = Inches(11.0)
            section.page_height = Inches(8.5)
        else:
            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = Inches(8.5)
            section.page_height = Inches(11.0)

    def _render_slide_header(self, slide_num: int, title: str):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True

        if self.show_slide_dividers:
            badge = p.add_run(f"[Slide {slide_num}] ")
            badge.font.size = Pt(8.5)
            badge.font.bold = True
            badge.font.color.rgb = RGBColor(110, 110, 110)

        if title:
            t_run = p.add_run(title)
            t_run.font.size = Pt(13)
            t_run.font.bold = True
            t_run.font.color.rgb = RGBColor(25, 55, 95)

    def _render_two_column_row(self, row: LayoutRow, total_width_in: float):
        tbl = self.doc.add_table(rows=1, cols=2)
        tbl.autofit = True
        self._remove_table_borders(tbl)

        col_width = (total_width_in / 2.0) - 0.15
        left_cell = tbl.cell(0, 0)
        right_cell = tbl.cell(0, 1)
        left_cell.width = Inches(col_width)
        right_cell.width = Inches(col_width)

        self._populate_cell(left_cell, row.left_elements, col_width)
        self._populate_cell(right_cell, row.right_elements, col_width)

    def _populate_cell(self, cell, elements: List[SlideElement], max_w_in: float):
        for el in elements:
            self._render_element_into_container(cell, el, max_width_in=max_w_in)
        if len(cell.paragraphs) > 1 and not cell.paragraphs[0].text and not cell.paragraphs[0].runs:
            p_element = cell.paragraphs[0]._element
            p_element.getparent().remove(p_element)

    def _render_element_into_container(self, container, el: SlideElement, max_width_in: float):
        if el.element_type in ("text", "title", "smartart"):
            if el.element_type == "smartart":
                tag_p = container.add_paragraph()
                tag_p.paragraph_format.space_before = Pt(2)
                tag_p.paragraph_format.space_after = Pt(1)
                tag_run = tag_p.add_run("SmartArt Diagram Nodes:")
                tag_run.font.size = Pt(9.0)
                tag_run.font.italic = True
                tag_run.font.color.rgb = RGBColor(80, 110, 140)

            for p_data in el.paragraphs:
                p = container.add_paragraph()
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(2)

                if p_data.is_bullet or p_data.level > 0:
                    indent_level = max(1, p_data.level)
                    p.paragraph_format.left_indent = Inches(0.22 * indent_level)
                    bullet_prefix = "• " if indent_level == 1 else "◦ "
                    first_text = p_data.runs[0].text.lstrip("•-▪ ")
                    p_data.runs[0].text = bullet_prefix + first_text

                for r_data in p_data.runs:
                    run = p.add_run(r_data.text)
                    run.bold = r_data.bold
                    run.italic = r_data.italic
                    run.underline = r_data.underline
                    if r_data.font_size:
                        scaled_pt = max(9.0, min(12.0, r_data.font_size * 0.55))
                        run.font.size = Pt(scaled_pt)
                    else:
                        run.font.size = Pt(10.5)

                    if r_data.rgb_color:
                        run.font.color.rgb = RGBColor(*r_data.rgb_color)

        elif el.element_type == "image" and el.image_bytes:
            self._insert_constrained_picture(container, el.image_bytes, el.image_size_px, el.width_ratio, max_width_in)
            if self.include_ocr_captions and el.ocr_text:
                ocr_p = container.add_paragraph()
                ocr_p.paragraph_format.space_before = Pt(1)
                ocr_p.paragraph_format.space_after = Pt(3)
                ocr_p.paragraph_format.left_indent = Inches(0.2)
                lbl = ocr_p.add_run("[OCR Extracted Text]: ")
                lbl.font.size = Pt(8.5)
                lbl.font.bold = True
                lbl.font.color.rgb = RGBColor(100, 100, 100)
                txt = ocr_p.add_run(el.ocr_text)
                txt.font.size = Pt(8.5)
                txt.font.italic = True
                txt.font.color.rgb = RGBColor(90, 90, 90)

        elif el.element_type == "chart":
            if el.image_bytes:
                self._insert_constrained_picture(container, el.image_bytes, el.image_size_px, el.width_ratio, max_width_in)
            if self.include_chart_tables and el.table_data:
                self._render_table(container, el.table_data)

        elif el.element_type == "table" and el.table_data:
            self._render_table(container, el.table_data)

    def _insert_constrained_picture(self, container, img_bytes: bytes, size_px, width_ratio: float, max_width_in: float):
        p = container.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        run = p.add_run()

        w_px, h_px = size_px or (400, 300)
        aspect = w_px / max(1.0, float(h_px))

        target_w = min(max_width_in * 0.92, max(1.8, max_width_in * width_ratio))
        target_h = target_w / aspect
        if target_h > self.max_image_height_in:
            target_h = self.max_image_height_in
            target_w = target_h * aspect

        run.add_picture(io.BytesIO(img_bytes), width=Inches(target_w))

    @staticmethod
    def _render_table(container, table_data: List[List[str]]):
        rows_cnt = len(table_data)
        cols_cnt = max(len(r) for r in table_data)
        if rows_cnt > 0 and cols_cnt > 0:
            tbl = container.add_table(rows=rows_cnt, cols=cols_cnt)
            tbl.style = "Table Grid"
            for r_i, row_vals in enumerate(table_data):
                for c_i, val in enumerate(row_vals):
                    cell = tbl.cell(r_i, c_i)
                    cell.text = val
                    for cell_p in cell.paragraphs:
                        cell_p.paragraph_format.space_before = Pt(1)
                        cell_p.paragraph_format.space_after = Pt(1)
                        for cell_run in cell_p.runs:
                            cell_run.font.size = Pt(9.5)
                            if r_i == 0:
                                cell_run.font.bold = True

    @staticmethod
    def _remove_table_borders(table):
        tbl_pr = table._tbl.tblPr
        borders = OxmlElement("w:tblBorders")
        for border_name in ("top", "left", "bottom", "right", "insideH", "insideV"):
            border = OxmlElement(f"w:{border_name}")
            border.set(qn("w:val"), "none")
            borders.append(border)
        tbl_pr.append(borders)
