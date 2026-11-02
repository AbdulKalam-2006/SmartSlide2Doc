# SmartSlide2Doc Pro

Zero-wastage PowerPoint (`.pptx`) to Microsoft Word (`.docx`) transformation engine with:
- **Editable Conversion Mode**: Preserves typography, hierarchical bullets, native tables, SmartArt diagram nodes, and chart data tables.
- **Auto Slide Arrangement**: Streams consecutive slides continuously onto the same Word page with zero wasted vertical whitespace, and automatically pairs side-by-side shapes into borderless 2-column layouts.
- **Dynamic Hybrid Orientation (`both`)**: Uses Portrait by default and inserts dynamic Landscape section breaks only for wide tables (>=6 columns) or panoramic visuals.
- **OCR & Vector Chart/SmartArt Support**:
  - **Image OCR**: Uses `pytesseract` (when Tesseract is installed) to index text inside raster images and append searchable captions.
  - **Vector Charts**: Renders native PowerPoint charts via Matplotlib at 160 DPI and appends an editable Word table of the underlying series data.
  - **SmartArt Unpacking**: Parses OpenXML `<dgm:relIds>` data models so SmartArt graphics become editable bullet lists.

## Quick Start

1. **Install Python dependencies:**
   ```bash
   pip install -r requirements.txt
   ```
   *(Optional: Install Tesseract OCR on your system and add it to PATH for raster image text extraction).*

2. **Launch the Streamlit Web UI:**
   ```bash
   streamlit run app.py
   ```

3. **Run via Command Line:**
   ```bash
   python cli.py sample/demo_presentation.pptx -o sample/output.docx --orientation both
   ```

4. **Regenerate Sample Demo Files:**
   ```bash
   python generate_demo_pptx.py
   ```
