from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from PIL import Image, ImageDraw
import io
from core.extractor import PPTXExtractor
from core.analyzer import SpatialLayoutAnalyzer
from core.builder import CompactDocxBuilder

def create_sample_pptx(filename="sample/demo_presentation.pptx"):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Slide 1: Title + Side-by-Side Comparison (2 columns)
    blank_layout = prs.slide_layouts[6]
    s1 = prs.slides.add_slide(blank_layout)
    t1 = s1.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.5), Inches(0.9))
    t1.text_frame.text = "System Architecture & Comparison"
    t1.text_frame.paragraphs[0].runs[0].font.bold = True
    t1.text_frame.paragraphs[0].runs[0].font.size = Pt(28)

    left_box = s1.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(5.4), Inches(4.5))
    tf_l = left_box.text_frame
    tf_l.text = "Standard PPT-to-Word Converters:"
    for item in ["Forces 1 slide = 1 page", "Leaves 60% blank space on sparse slides", "Breaks alignment on portrait pages"]:
        p = tf_l.add_paragraph()
        p.text = item
        p.level = 1

    right_box = s1.shapes.add_textbox(Inches(6.8), Inches(1.6), Inches(5.4), Inches(4.5))
    tf_r = right_box.text_frame
    tf_r.text = "SmartSlide2Doc Engine:"
    for item in ["Streams multiple slides onto Page 1", "Preserves 2-column side-by-side blocks", "Auto-switches Portrait & Landscape"]:
        p = tf_r.add_paragraph()
        p.text = item
        p.level = 1

    # Slide 2: Native Chart + Summary Notes (Flows right onto Page 1 below Slide 1!)
    s2 = prs.slides.add_slide(blank_layout)
    t2 = s2.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.5), Inches(0.8))
    t2.text_frame.text = "Quarterly Throughput Metrics (Vector Chart)"
    t2.text_frame.paragraphs[0].runs[0].font.bold = True

    chart_data = CategoryChartData()
    chart_data.categories = ["Q1", "Q2", "Q3", "Q4"]
    chart_data.add_series("Baseline", (42, 48, 51, 55))
    chart_data.add_series("SmartSlide2Doc", (78, 86, 92, 97))
    s2.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(0.8), Inches(1.5), Inches(6.0), Inches(4.2),
        chart_data
    )

    notes_box = s2.shapes.add_textbox(Inches(7.2), Inches(1.8), Inches(5.2), Inches(3.8))
    ntf = notes_box.text_frame
    ntf.text = "Key Observations:"
    for pt in ["Page count reduced by 64% across 50-slide decks", "All chart values extracted as editable Word tables", "Zero visual clipping"]:
        p = ntf.add_paragraph()
        p.text = pt
        p.level = 1

    # Slide 3: Wide 7-Column Data Table (Triggers Auto-Landscape Section Break!)
    s3 = prs.slides.add_slide(blank_layout)
    t3 = s3.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.5), Inches(0.8))
    t3.text_frame.text = "Regional Deployment Matrix (Auto-Landscape Trigger)"
    tbl_shape = s3.shapes.add_table(4, 7, Inches(0.8), Inches(1.6), Inches(11.5), Inches(3.5))
    headers = ["Region", "Nodes", "Latency", "Uptime", "OCR Pass", "Chart Mode", "Status"]
    rows = [
        headers,
        ["South-1", "128", "11ms", "99.98%", "Enabled", "Vector+Table", "Active"],
        ["West-2", "96", "14ms", "99.95%", "Enabled", "Vector+Table", "Active"],
        ["North-1", "64", "18ms", "99.91%", "Enabled", "Vector+Table", "Standby"],
    ]
    for r_idx, r_vals in enumerate(rows):
        for c_idx, val in enumerate(r_vals):
            tbl_shape.table.cell(r_idx, c_idx).text = val

    prs.save(filename)
    print(f"Created sample presentation: {filename}")

if __name__ == "__main__":
    create_sample_pptx()
    extractor = PPTXExtractor("sample/demo_presentation.pptx")
    raw = extractor.extract_all_slides()
    analyzer = SpatialLayoutAnalyzer(orientation_policy="both")
    analyzed = analyzer.analyze_presentation(raw)
    builder = CompactDocxBuilder()
    docx_bytes = builder.build(analyzed)
    with open("sample/demo_output_compact.docx", "wb") as f:
        f.write(docx_bytes)
    print("Generated sample output: sample/demo_output_compact.docx")
