import argparse
from pathlib import Path
from core.extractor import PPTXExtractor
from core.analyzer import SpatialLayoutAnalyzer
from core.builder import CompactDocxBuilder

def main():
    parser = argparse.ArgumentParser(description="SmartSlide2Doc: Zero-wastage PPTX to DOCX converter.")
    parser.add_argument("input", help="Path to input .pptx file")
    parser.add_argument("-o", "--output", help="Path to output .docx file", default=None)
    parser.add_argument(
        "--orientation", choices=["both", "portrait", "landscape"], default="both",
        help="Orientation strategy: 'both' switches dynamically per slide content."
    )
    parser.add_argument("--max-img-height", type=float, default=3.0, help="Max image height in inches.")
    parser.add_argument("--no-ocr", action="store_true", help="Disable OCR extraction on embedded images.")
    parser.add_argument("--no-charts", action="store_true", help="Disable vector chart rendering.")
    args = parser.parse_args()

    in_path = Path(args.input)
    out_path = Path(args.output) if args.output else in_path.with_name(f"{in_path.stem}_compact.docx")

    extractor = PPTXExtractor(
        str(in_path),
        enable_ocr=not args.no_ocr,
        render_charts=not args.no_charts
    )
    raw_slides = extractor.extract_all_slides()

    analyzer = SpatialLayoutAnalyzer(orientation_policy=args.orientation)
    analyzed_slides = analyzer.analyze_presentation(raw_slides)

    builder = CompactDocxBuilder(
        max_image_height_in=args.max_img_height,
        include_chart_tables=not args.no_charts,
        include_ocr_captions=not args.no_ocr
    )
    docx_bytes = builder.build(analyzed_slides)

    out_path.write_bytes(docx_bytes)
    print(f"[SUCCESS] Converted {len(analyzed_slides)} slides -> {out_path}")

if __name__ == "__main__":
    main()
