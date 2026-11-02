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

# 📄 SmartSlide2Doc — Zero-Wastage PPTX to Word Engine

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)](https://www.python.org/)
[![Streamlit](https://img.shields.io/badge/Streamlit-Web_UI-FF4B4B?logo=streamlit&logoColor=white)](https://streamlit.io/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)

**SmartSlide2Doc** is an intelligent document transformation engine that converts fixed-canvas PowerPoint presentations (`.pptx`) into compact, continuous, and fully editable Microsoft Word documents (`.docx`) with **zero page wastage**.

Unlike standard PPT-to-Word converters that force a `1 Slide = 1 Page` layout—leaving huge blank gaps, misaligned images, and awkward landscape pages—**SmartSlide2Doc** extracts semantic content, preserves side-by-side spatial layouts, scales media proportionally, and flows consecutive slides seamlessly onto the same page.

---

## ✨ Key Features

- **🔄 Continuous Zero-Wastage Page Flow**: Eliminates forced slide page breaks. If Slide 1 only uses half of Word Page 1, Slide 2 flows immediately into the remaining space on Page 1.
- **📐 Automatic 2-Column Spatial Clustering**: Analyzes the 2D bounding-box coordinates (`left`, `top`, `width`, `height`) of every shape on a slide. Side-by-side text boxes or image-text pairs are automatically locked into borderless 2-column Word tables.
- **📄 Hybrid Portrait & Landscape Switching (`both` mode)**: Uses space-efficient **Portrait** orientation by default and dynamically inserts Word Section Breaks to switch to **Landscape** only when a slide contains wide tables ($\ge 6$ columns) or panoramic diagrams.
- **📊 Vector Chart Renderer + Editable Data Tables**: Renders native PowerPoint charts (`Bar`, `Column`, `Line`, `Pie`, `Doughnut`) at 160 DPI via Matplotlib and generates an editable Word table containing the exact series data right below the visual.
- **🧩 OpenXML SmartArt Unpacker**: Parses underlying `<dgm:relIds>` drawingML data models to convert locked SmartArt graphics into hierarchical, editable Word bullet lists.
- **🔍 Embedded Image OCR**: Integrates `Pillow` preprocessing and `Tesseract OCR` (`pytesseract`) to extract text trapped inside raster screenshots and diagrams, appending searchable captions underneath.

---

## 🏗️ System Architecture

```text
+-------------------------------------------------------------------------+
|                        SmartSlide2Doc Pipeline                          |
+-------------------------------------------------------------------------+
       |                        |                        |
       v                        v                        v
+--------------+      +-------------------+      +------------------------+
| 1. Extractor | ---> | 2. Auto-Layout    | ---> | 3. Section Manager     |
| - Text/Runs  |      |    Classifier     |      | - Dynamic Portrait vs  |
| - Images/OCR |      | - 1D Vertical     |      |   Landscape switching  |
| - Charts/Tbl |      | - 2D Side-by-Side |      | - Continuous flow      |
| - SmartArt   |      | - Wide Visuals    |      |   compaction           |
+--------------+      +-------------------+      +------------------------+
                                                            |
                                                            v
                                                 +------------------------+
                                                 | 4. DOCX Renderer       |
                                                 | - Invisible Grid Tables|
                                                 | - Aspect-Ratio Scaling |
                                                 | - Native Editable Text |
                                                 +------------------------+