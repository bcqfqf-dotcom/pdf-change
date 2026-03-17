import fitz
import sys

def analyze(path):
    doc = fitz.open(path)
    for i, page in enumerate(doc):
        rect = page.rect
        print(f"Page {i}: {rect.width:.2f} x {rect.height:.2f} pts")
    doc.close()

if __name__ == "__main__":
    if len(sys.argv) > 1:
        analyze(sys.argv[1])
    else:
        print("Usage: python analyze_pdf.py <path_to_pdf>")
