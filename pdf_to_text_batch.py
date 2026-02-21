#!/usr/bin/env python3
"""
Batch PDF to Text Converter for PhD Applications
Converts a folder of PDFs (text-based + scanned) to individual .txt files
Uses PyMuPDF for text-based pages and pytesseract OCR for scanned pages

Requirements:
    pip install pymupdf pytesseract pillow
    Also install Tesseract OCR engine:
        Mac:    brew install tesseract
        Linux:  sudo apt install tesseract-ocr
        Windows: https://github.com/UB-Mannheim/tesseract/wiki
"""

import fitz  # PyMuPDF
import os
import sys
from pathlib import Path
from PIL import Image
import io

# Try to import pytesseract for OCR fallback on scanned pages
try:
    import pytesseract
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False
    print("⚠️  pytesseract not found — scanned pages will be skipped.")
    print("   Install with: pip install pytesseract")


def extract_text_from_pdf(pdf_path: Path, ocr_threshold: int = 50) -> str:
    """
    Extract text from a PDF, using OCR for scanned pages.
    
    Args:
        pdf_path: Path to the PDF file
        ocr_threshold: If a page has fewer than this many characters,
                       treat it as scanned and attempt OCR
    
    Returns:
        Extracted text as a string
    """
    doc = fitz.open(str(pdf_path))
    full_text = []

    for page_num, page in enumerate(doc, start=1):
        # Try direct text extraction first
        text = page.get_text().strip()

        if len(text) >= ocr_threshold:
            # Good text-based page
            full_text.append(f"--- Page {page_num} ---\n{text}")
        elif OCR_AVAILABLE:
            # Likely a scanned page — render to image and OCR
            mat = fitz.Matrix(2.0, 2.0)  # 2x zoom for better OCR accuracy
            pix = page.get_pixmap(matrix=mat)
            img_data = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_data))
            ocr_text = pytesseract.image_to_string(img).strip()
            if ocr_text:
                full_text.append(f"--- Page {page_num} [OCR] ---\n{ocr_text}")
            else:
                full_text.append(f"--- Page {page_num} [OCR: no text detected] ---")
        else:
            full_text.append(f"--- Page {page_num} [scanned, OCR not available] ---")

    doc.close()
    return "\n\n".join(full_text)


def batch_convert(input_folder: str, output_folder: str = None):
    """
    Convert all PDFs in input_folder to .txt files in output_folder.
    
    Args:
        input_folder: Path to folder containing PDFs
        output_folder: Path to save .txt files (defaults to input_folder/text_output)
    """
    input_path = Path(input_folder)
    
    if not input_path.exists():
        print(f"❌ Input folder not found: {input_folder}")
        sys.exit(1)

    # Set up output folder
    if output_folder is None:
        output_path = input_path / "text_output"
    else:
        output_path = Path(output_folder)
    output_path.mkdir(parents=True, exist_ok=True)

    # Find all PDFs
    pdf_files = sorted(input_path.glob("*.pdf"))
    if not pdf_files:
        print(f"⚠️  No PDF files found in: {input_folder}")
        sys.exit(0)

    print(f"📂 Found {len(pdf_files)} PDF(s) in: {input_folder}")
    print(f"💾 Output folder: {output_path}\n")

    success, failed = 0, []

    for i, pdf_file in enumerate(pdf_files, start=1):
        out_file = output_path / (pdf_file.stem + ".txt")
        print(f"[{i}/{len(pdf_files)}] Processing: {pdf_file.name} ...", end=" ", flush=True)
        
        try:
            text = extract_text_from_pdf(pdf_file)
            out_file.write_text(text, encoding="utf-8")
            char_count = len(text)
            print(f"✅ ({char_count:,} chars → {out_file.name})")
            success += 1
        except Exception as e:
            print(f"❌ FAILED: {e}")
            failed.append((pdf_file.name, str(e)))

    # Summary
    print(f"\n{'='*50}")
    print(f"✅ Successfully converted: {success}/{len(pdf_files)}")
    if failed:
        print(f"❌ Failed ({len(failed)}):")
        for name, err in failed:
            print(f"   • {name}: {err}")
    print(f"\nText files saved to: {output_path.resolve()}")


if __name__ == "__main__":
    # Usage: python pdf_to_text_batch.py /path/to/pdfs [/path/to/output]
    if len(sys.argv) < 2:
        print("Usage: python pdf_to_text_batch.py <input_folder> [output_folder]")
        print("Example: python pdf_to_text_batch.py ~/Downloads/applicants")
        sys.exit(1)

    input_dir = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    batch_convert(input_dir, output_dir)