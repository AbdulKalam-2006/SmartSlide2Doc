import io
import xml.etree.ElementTree as ET
from typing import List, Optional, Tuple
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from pptx.enum.chart import XL_CHART_TYPE
from core.models import ParagraphData, TextRunData

class VectorChartAndSmartArtRenderer:
    """
    Handles native PowerPoint Charts and SmartArt (IGX diagrams):
    1. Renders native PPTX chart data into high-DPI images AND extracts a structured data table.
    2. Parses OpenXML SmartArt (<dgm:relIds> / drawingML text bodies) so all SmartArt nodes
       become fully editable, hierarchical bullet lists in Word.
    """

    @staticmethod
    def extract_chart(shape) -> Tuple[Optional[bytes], Optional[Tuple[int, int]], List[List[str]], Optional[str]]:
        """
        Returns (rendered_png_bytes, (width_px, height_px), fallback_data_table, chart_title)
        """
        if not getattr(shape, "has_chart", False):
            return None, None, [], None

        chart = shape.chart
        chart_title = None
        try:
            if chart.has_title and chart.chart_title.has_text_frame:
                chart_title = chart.chart_title.text_frame.text.strip()
        except Exception:
            chart_title = None

        categories = []
        try:
            plot = chart.plots[0]
            categories = [str(c) for c in plot.categories]
        except Exception:
            pass

        series_list = []
        for s in chart.series:
            s_name = str(getattr(s, "name", "Series"))
            try:
                vals = [float(v) if v is not None else 0.0 for v in s.values]
            except Exception:
                vals = []
            series_list.append((s_name, vals))

        if not categories and series_list and series_list[0][1]:
            categories = [f"Item {i+1}" for i in range(len(series_list[0][1]))]

        # Build editable data table representation of the chart
        table_matrix: List[List[str]] = []
        if categories and series_list:
            header = ["Category"] + [s_name for s_name, _ in series_list]
            table_matrix.append(header)
            for idx, cat in enumerate(categories):
                row = [cat]
                for _, vals in series_list:
                    val_str = f"{vals[idx]:g}" if idx < len(vals) else ""
                    row.append(val_str)
                table_matrix.append(row)

        # Render high-DPI Matplotlib visual of the chart
        png_bytes, px_size = VectorChartAndSmartArtRenderer._render_chart_matplotlib(
            chart, chart_title, categories, series_list
        )
        return png_bytes, px_size, table_matrix, chart_title

    @staticmethod
    def _render_chart_matplotlib(chart, title: Optional[str], categories: List[str], series_list: List[Tuple[str, List[float]]]):
        if not series_list:
            return None, None

        fig, ax = plt.subplots(figsize=(6.2, 3.4), dpi=160)
        chart_type = None
        try:
            chart_type = chart.chart_type
        except Exception:
            pass

        palette = ["#1f4e78", "#2e75b6", "#ed7d31", "#70ad47", "#ffc000", "#5b9bd5"]

        try:
            if chart_type in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.PIE_EXPLODED, XL_CHART_TYPE.DOUGHNUT) and len(series_list) == 1:
                _, vals = series_list[0]
                labels = categories if len(categories) == len(vals) else [f"Slice {i+1}" for i in range(len(vals))]
                ax.pie(vals, labels=labels, autopct="%1.1f%%", startangle=140, colors=palette[:len(vals)])
            elif chart_type in (XL_CHART_TYPE.LINE, XL_CHART_TYPE.LINE_MARKERS, XL_CHART_TYPE.LINE_STACKED):
                x = range(len(categories))
                for idx, (s_name, vals) in enumerate(series_list):
                    if len(vals) == len(categories):
                        ax.plot(x, vals, marker="o", linewidth=2, label=s_name, color=palette[idx % len(palette)])
                ax.set_xticks(list(x))
                ax.set_xticklabels(categories, rotation=20, ha="right", fontsize=8.5)
                ax.grid(True, linestyle="--", alpha=0.4)
                if len(series_list) > 1:
                    ax.legend(frameon=False, fontsize=8.5)
            else:
                # Grouped bar chart default
                n_series = max(1, len(series_list))
                x = list(range(len(categories)))
                total_bar_width = 0.72
                bar_w = total_bar_width / n_series
                for idx, (s_name, vals) in enumerate(series_list):
                    if len(vals) == len(categories):
                        offsets = [pos - (total_bar_width / 2) + (idx + 0.5) * bar_w for pos in x]
                        ax.bar(offsets, vals, width=bar_w, label=s_name, color=palette[idx % len(palette)])
                ax.set_xticks(x)
                ax.set_xticklabels(categories, rotation=20, ha="right", fontsize=8.5)
                ax.grid(axis="y", linestyle="--", alpha=0.4)
                if len(series_list) > 1:
                    ax.legend(frameon=False, fontsize=8.5)

            if title:
                ax.set_title(title, fontsize=10.5, fontweight="bold", pad=8)

            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            fig.tight_layout()

            buf = io.BytesIO()
            fig.savefig(buf, format="png", dpi=160)
            plt.close(fig)
            png_data = buf.getvalue()
            return png_data, (992, 544)
        except Exception:
            plt.close(fig)
            return None, None

    @staticmethod
    def extract_smartart_paragraphs(shape) -> List[ParagraphData]:
        """
        Extracts hierarchical text nodes from PowerPoint SmartArt graphic frames
        by inspecting both the shape XML and linked SmartArt data/drawing XML parts.
        """
        paragraphs: List[ParagraphData] = []
        try:
            element = shape._element
            xml_bytes = ET.tostring(element)
            root = ET.fromstring(xml_bytes)

            # Check if this graphicFrame contains a SmartArt diagram reference (<dgm:relIds>)
            dgm_ns = "{http://schemas.openxmlformats.org/drawingml/2006/diagram}"
            a_ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
            r_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

            rel_ids = root.find(f".//{dgm_ns}relIds")
            if rel_ids is None:
                return []

            dm_rId = rel_ids.attrib.get(f"{r_ns}dm")
            if dm_rId and hasattr(shape, "part") and dm_rId in shape.part.rels:
                data_part = shape.part.rels[dm_rId].target_part
                data_root = ET.fromstring(data_part.blob)

                # Extract text from all <dgm:pt> data model points
                for pt in data_root.findall(f".//{dgm_ns}pt"):
                    t_elems = pt.findall(f".//{a_ns}t")
                    node_text = "".join(t.text for t in t_elems if t.text).strip()
                    if node_text:
                        paragraphs.append(ParagraphData(
                            runs=[TextRunData(text=node_text, bold=False)],
                            level=1,
                            is_bullet=True
                        ))
        except Exception:
            pass

        return paragraphs
