import io
import shutil
from typing import Optional
from PIL import Image, ImageOps, ImageFilter

try:
    import pytesseract
    HAS_PYTESSERACT = True
except ImportError:
    HAS_PYTESSERACT = False

class OCREngine:
    """
    Extracts embedded text from raster images and diagrams so that
    screenshots, scanned diagrams, and flattened graphics become
    searchable and editable inside the generated Word document.
    """
    def __init__(self, enabled: bool = True, tesseract_cmd: Optional[str] = None):
        self.enabled = enabled
        if tesseract_cmd and HAS_PYTESSERACT:
            pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
        self.available = self._check_availability()

    def _check_availability(self) -> bool:
        if not self.enabled or not HAS_PYTESSERACT:
            return False
        # Check if tesseract binary is on PATH or configured
        if shutil.which("tesseract") is not None:
            return True
        try:
            pytesseract.get_tesseract_version()
            return True
        except Exception:
            return False

    def extract_text_from_image(self, image_bytes: bytes, min_chars: int = 8) -> Optional[str]:
        if not self.available or not image_bytes:
            return None
        try:
            with Image.open(io.BytesIO(image_bytes)) as img:
                # Skip tiny icons/bullets
                if img.width < 90 or img.height < 35:
                    return None
                # Preprocess: grayscale + contrast enhancement for crisp OCR
                gray = ImageOps.grayscale(img)
                sharpened = gray.filter(ImageFilter.SHARPEN)
                raw_text = pytesseract.image_to_string(sharpened, config="--psm 6")
                cleaned = " ".join(
                    line.strip() for line in raw_text.splitlines() if line.strip()
                )
                if len(cleaned) >= min_chars:
                    return cleaned
        except Exception:
            return None
        return None
