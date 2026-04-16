"""app/ocr/engine.py — Tesseract OCR engine initialisation and image pre-processing."""
import os
import shutil
import logging
from io import BytesIO

from PIL import Image, ImageOps, ImageEnhance
try:
    import pytesseract
    from pytesseract import Output
except Exception:
    pytesseract = None
    Output = None

from ..config import TESSERACT_DIR

MONTHS_MAP = {
    'january': 1, 'february': 2, 'march': 3, 'april': 4, 'may': 5, 'june': 6,
    'july': 7, 'august': 8, 'september': 9, 'october': 10, 'november': 11, 'december': 12,
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'jun': 6, 'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
}

def _preprocess_for_ocr(img: Image.Image) -> Image.Image:
    # ensure it's RGB (no alpha) before processing
    rgb = img.convert('RGB')
    gray = ImageOps.grayscale(rgb)
    gray = ImageEnhance.Contrast(gray).enhance(2.0)
    gray = gray.resize((gray.size[0] * 2, gray.size[1] * 2))
    # Original threshold (170) works better for general readability
    bw = gray.point(lambda x: 0 if x < 170 else 255, '1')
    return bw

