# scanned_invoice_extractor
Python-based OCR system that extracts key information from scanned invoices (PDF/images) using Tesseract and OpenCV, and organizes the data into structured Excel output.

# Scanned Invoice Extractor (OCR)

A Python-based OCR system that extracts important information from scanned invoices (PDF/images) and converts it into structured Excel data. The project uses Tesseract OCR and image preprocessing techniques to improve text extraction accuracy and automate manual invoice data entry.

## Features

* Extracts text from scanned invoice PDFs and images
* Uses OCR to automatically identify invoice details
* Applies image preprocessing using OpenCV to improve OCR accuracy
* Extracts key fields such as invoice number, date, vendor, and total amount
* Stores extracted data in structured Excel format
* Reduces manual data entry and improves efficiency

## Tech Stack

* Python
* Tesseract OCR
* OpenCV
* PyTesseract
* Pandas
* OpenPyXL

## How It Works

1. A scanned invoice (PDF or image) is provided as input.
2. The system preprocesses the image using OpenCV to enhance text clarity.
3. Tesseract OCR extracts the textual content from the invoice.
4. Important invoice details are identified from the extracted text.
5. The extracted information is organized into structured format.
6. The final data is exported into an Excel file.

## Project Structure

Scanned-Invoice-Extractor
│
├── input_invoices/        # Folder containing invoice PDFs/images
├── output/                # Extracted Excel files
├── main2.py   # Main script for processing invoices
└── README.md              # Project documentation

## Applications

* Automating invoice data entry
* Financial record management
* Document digitization
* Business process automation

## Future Improvements

* Support multiple invoice formats
* Improve field detection accuracy
* Build a web interface for uploading invoices
* Implement machine learning for better document understanding

## Author

Developed as part of an OCR-based document processing project using Python.
