import io
import streamlit as st
from core.extractor import PPTXExtractor
from core.analyzer import SpatialLayoutAnalyzer
from core.builder import CompactDocxBuilder

st.set_page_config(page_title="SmartSlide2Doc Pro", page_icon="📄", layout="wide")

st.title("📄 SmartSlide2Doc — Zero-Wastage PPTX to Word Engine")
st.markdown(
    "Convert presentations into **continuous, editable Word documents** with automatic "
    "2-column spatial packing, OCR text extraction from images, vector chart rendering, "
    "SmartArt node unpacking, and hybrid Portrait/Landscape switching."
)

with st.sidebar:
    st.header("⚙️ Conversion Settings")
    orientation_mode = st.selectbox(
        "Page Orientation Policy",
        options=["both", "portrait", "landscape"],
        index=0,
        help="'both' uses Portrait by default and switches to Landscape automatically for wide tables or panoramic charts."
    )
    max_img_h = st.slider(
        "Max Image/Chart Height (Inches)",
        min_value=1.5, max_value=5.5, value=3.0, step=0.25,
        help="Constrains tall images so multiple slides pack cleanly onto a single page."
    )
    show_dividers = st.checkbox("Show Compact Slide Badges ([Slide #])", value=True)
    enable_ocr = st.checkbox("Enable Image OCR (Searchable Text)", value=True)
    render_charts = st.checkbox("Render Native Charts + Data Tables", value=True)

uploaded_file = st.file_uploader("Upload a PowerPoint file (.pptx)", type=["pptx"])

if uploaded_file is not None:
    if st.button("🚀 Convert to Compact Word Document", type="primary"):
        with st.spinner("Extracting shapes, running OCR & chart rendering, and building DOCX..."):
            pptx_bytes = io.BytesIO(uploaded_file.read())

            extractor = PPTXExtractor(
                pptx_bytes,
                enable_ocr=enable_ocr,
                render_charts=render_charts
            )
            raw_slides = extractor.extract_all_slides()

            analyzer = SpatialLayoutAnalyzer(orientation_policy=orientation_mode)
            analyzed_slides = analyzer.analyze_presentation(raw_slides)

            builder = CompactDocxBuilder(
                max_image_height_in=max_img_h,
                show_slide_dividers=show_dividers,
                include_chart_tables=render_charts,
                include_ocr_captions=enable_ocr
            )
            docx_bytes = builder.build(analyzed_slides)

            portrait_cnt = sum(1 for s in analyzed_slides if s.orientation == "portrait")
            landscape_cnt = sum(1 for s in analyzed_slides if s.orientation == "landscape")
            two_col_cnt = sum(sum(1 for r in s.rows if r.is_two_column) for s in analyzed_slides)

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total Slides Processed", len(analyzed_slides))
            col2.metric("Portrait Slides", portrait_cnt)
            col3.metric("Landscape Slides (Auto)", landscape_cnt)
            col4.metric("Side-by-Side 2-Col Blocks", two_col_cnt)

            out_name = uploaded_file.name.rsplit(".", 1)[0] + "_compact.docx"
            st.download_button(
                label="📥 Download Editable Word (.docx)",
                data=docx_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
