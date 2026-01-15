#!/usr/bin/env python3
"""
Script to analyze construction detail PDFs
"""
import PyPDF2
import sys

def extract_text_from_pdf(pdf_path):
    """Extract text from a PDF file"""
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            text = ""
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                text += f"\n--- Page {page_num + 1} ---\n"
                text += page.extract_text()
            return text
    except Exception as e:
        return f"Error reading PDF: {str(e)}"

def main():
    pdf1_path = "src/lieke/6.1 Detail + analyse Sven de Raad (1).pdf"
    pdf2_path = "src/lieke/EVE6.1_Groep 7.1 Detail Lieke.pdf"

    print("=" * 80)
    print("ANALYZING FILE 1: Sven de Raad Detail + Analysis")
    print("=" * 80)
    text1 = extract_text_from_pdf(pdf1_path)
    print(text1)

    print("\n\n")
    print("=" * 80)
    print("ANALYZING FILE 2: Lieke Detail")
    print("=" * 80)
    text2 = extract_text_from_pdf(pdf2_path)
    print(text2)

if __name__ == "__main__":
    main()

