# Pdf-Image-to-text-converter

The PDF/Image Converter is a Python script designed to extract text from PDF files and convert it into structured data in Excel format for further analysis. This script utilizes various libraries such as PyPDF2, pathlib, openpyxl, fitz, pytesseract, and PIL.Image to accomplish its tasks.

The process involves the script searching for PDF files in a designated directory, creating a list of these files, and then using the extracttext() method to extract the text data from each PDF file. The extracted data is then added to an Excel workbook to produce a structured and organized result.
