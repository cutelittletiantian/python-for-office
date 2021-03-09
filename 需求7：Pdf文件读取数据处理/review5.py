import os

pdfListPath = "./resources/review5"

pdfDocs = os.listdir(pdfListPath)


def sort_pdf(pdf_file):
    pdf_name = os.path.splitext(p=pdf_file)[0]
    return int(pdf_name)


pdfDocs.sort(key=sort_pdf, reverse=False)
print(pdfDocs)