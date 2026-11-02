from typing import List
from core.models import SlideElement, LayoutRow, AnalyzedSlide

class SpatialLayoutAnalyzer:
    """
    Analyzes slide element coordinates to:
    1. Determine optimal page orientation ('portrait' vs 'landscape').
    2. Automatically group side-by-side elements into 2-column rows.
    """
    def __init__(self, orientation_policy: str = "both"):
        # 'both' (auto-switch), 'portrait', or 'landscape'
        self.orientation_policy = orientation_policy.lower()

    def analyze_presentation(self, raw_slides: List[List[SlideElement]]) -> List[AnalyzedSlide]:
        analyzed = []
        for idx, elements in enumerate(raw_slides, start=1):
            if not elements:
                continue
            analyzed.append(self._analyze_single_slide(idx, elements))
        return analyzed

    def _analyze_single_slide(self, slide_num: int, elements: List[SlideElement]) -> AnalyzedSlide:
        title_text = None
        body_elements: List[SlideElement] = []

        for el in elements:
            if el.element_type == "title" and title_text is None:
                title_text = " ".join(
                    "".join(r.text for r in p.runs) for p in el.paragraphs
                ).strip()
            else:
                body_elements.append(el)

        # Sort top-to-bottom, then left-to-right
        body_elements.sort(key=lambda e: (round(e.top_ratio, 2), round(e.left_ratio, 2)))

        orientation = self._decide_orientation(body_elements)
        rows = self._cluster_into_rows(body_elements)

        return AnalyzedSlide(
            slide_number=slide_num,
            orientation=orientation,
            title=title_text,
            rows=rows
        )

    def _decide_orientation(self, elements: List[SlideElement]) -> str:
        if self.orientation_policy in ("portrait", "landscape"):
            return self.orientation_policy

        for el in elements:
            # Criterion 1: Wide data table (>= 6 columns)
            if el.element_type in ("table", "chart") and el.table_data:
                num_cols = max(len(r) for r in el.table_data)
                if num_cols >= 6:
                    return "landscape"

            # Criterion 2: Ultra-wide panoramic diagram/chart taking >=80% slide width
            if el.element_type == "image" and el.image_size_px:
                w_px, h_px = el.image_size_px
                aspect = w_px / max(1, h_px)
                if el.width_ratio >= 0.80 and aspect >= 2.2:
                    return "landscape"

        return "portrait"

    def _cluster_into_rows(self, elements: List[SlideElement]) -> List[LayoutRow]:
        """Pairs left-column and right-column elements when they overlap vertically."""
        rows: List[LayoutRow] = []
        used = [False] * len(elements)

        for i, el_a in enumerate(elements):
            if used[i]:
                continue

            a_center_x = el_a.left_ratio + (el_a.width_ratio / 2.0)
            is_narrow_enough = el_a.width_ratio <= 0.62

            partner_idx = None
            if is_narrow_enough:
                for j in range(i + 1, len(elements)):
                    if used[j]:
                        continue
                    el_b = elements[j]
                    if el_b.width_ratio > 0.62:
                        continue

                    b_center_x = el_b.left_ratio + (el_b.width_ratio / 2.0)

                    opposite_halves = (
                        (a_center_x < 0.52 and b_center_x >= 0.48) or
                        (b_center_x < 0.52 and a_center_x >= 0.48)
                    )
                    if opposite_halves and self._has_vertical_overlap(el_a, el_b):
                        partner_idx = j
                        break

            if partner_idx is not None:
                used[i] = True
                used[partner_idx] = True
                el_b = elements[partner_idx]
                left_el, right_el = (el_a, el_b) if el_a.left_ratio <= el_b.left_ratio else (el_b, el_a)
                rows.append(LayoutRow(
                    is_two_column=True,
                    left_elements=[left_el],
                    right_elements=[right_el]
                ))
            else:
                used[i] = True
                rows.append(LayoutRow(is_two_column=False, single_element=el_a))

        return rows

    @staticmethod
    def _has_vertical_overlap(a: SlideElement, b: SlideElement) -> bool:
        a_top, a_bot = a.top_ratio, a.top_ratio + a.height_ratio
        b_top, b_bot = b.top_ratio, b.top_ratio + b.height_ratio
        overlap = min(a_bot, b_bot) - max(a_top, b_top)
        min_height = min(a.height_ratio, b.height_ratio)
        tops_aligned = abs(a_top - b_top) <= 0.08
        return tops_aligned or (overlap > 0.25 * max(0.05, min_height))
