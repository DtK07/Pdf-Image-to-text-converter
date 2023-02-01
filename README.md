# Pdf-Image-to-text-converter

Pdf/Image coverter is a python script intended to convert the text in the pdf to excel format so that data can be structured and used for further analysis.

Libraries used:
PyPDF2, pathlib, openpyxl, fitz, pytesseract, PIL.Image

Process:
Once the code is run, it will look for the pdf files in a mentioned directory and create a list of files. Then creates a pdf object and uses extracttext() method to pull out the data and then create an excel workbook and update the extracted data.
