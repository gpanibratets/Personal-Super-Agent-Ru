#!/usr/bin/env python3
"""Extract text from PDF files for medical records processing."""

import sys
import os

def extract_pdf_text(pdf_path):
    """Extract text from PDF file."""
    try:
        import pypdf
        with open(pdf_path, 'rb') as file:
            reader = pypdf.PdfReader(file)
            text = ""
            for page in reader.pages:
                text += page.extract_text() + "\n"
            return text
    except ImportError:
        try:
            import PyPDF2
            with open(pdf_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                text = ""
                for page in reader.pages:
                    text += page.extract_text() + "\n"
                return text
        except ImportError:
            print("Error: Need to install pypdf or PyPDF2: pip install pypdf", file=sys.stderr)
            return None
    except Exception as e:
        print(f"Error reading PDF {pdf_path}: {e}", file=sys.stderr)
        return None

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 extract_pdf_text.py <pdf_file>")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    if not os.path.exists(pdf_path):
        print(f"Error: File not found: {pdf_path}", file=sys.stderr)
        sys.exit(1)
    
    text = extract_pdf_text(pdf_path)
    if text:
        print(text)
    else:
        sys.exit(1)

