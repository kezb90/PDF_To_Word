PDF to Word Converter
This project is a Python application that converts PDF files to Word documents, preserving both text and images. It leverages PyPDF2 for text extraction, PyMuPDF (fitz) for image extraction, and python-docx for generating the Word document. This setup ensures high fidelity in transferring content from PDF to Word, making it ideal for documents with mixed text and images.

Features
Extract Text: Retrieves and organizes text from each PDF page.
Extract Images: Captures images from the PDF, including embedded graphics, and places them in the Word document.
Create Word Document: Generates a Word file (.docx) with extracted text and images, with each page separated and formatted appropriately.
Installation
Requirements
Python 3.6+
Required Libraries: Install the dependencies with the following command:
bash
Copy code
pip install PyPDF2 PyMuPDF python-docx Pillow
Libraries Overview
PyPDF2: Extracts text content from PDF pages.
PyMuPDF (fitz): Extracts images from PDF pages, supporting a variety of embedded image formats.
python-docx: Creates and formats Word documents, inserting text and images.
Pillow (PIL): Converts images to ensure compatibility with python-docx.
Usage
Clone the repository or download the source code.
Place the PDF file you want to convert in the same directory as the script.
Run the script with:
bash
Copy code
python main.py
The program will create an output Word document (output_word_file.docx) with the extracted content from your PDF.
Example
python
Copy code
# Running the conversion function
pdf_to_word("PDF.pdf", "output_word_file.docx")
Output
The program saves a Word document with:

Page Headers: Each page from the PDF is marked and separated in the Word document.
Images: Images are inserted in their respective positions.
Page Breaks: Separates each page's content, matching the PDF structure.
Code Overview
extract_text_with_pypdf2(pdf_path): Uses PyPDF2 to extract and structure text content from the PDF.
extract_images_with_pymupdf(pdf_path): Uses PyMuPDF to extract images, converting them to PNG format for compatibility.
pdf_to_word(pdf_path, word_path): The main function that combines extracted text and images into a Word document.
Error Handling
The code uses try-except blocks to manage errors related to unsupported image formats and missing content.
Future Enhancements
Customizable Output: Add support for user-defined page layouts and image sizes.
Enhanced Error Handling: Add notifications for unsupported formats or missing data.
Metadata Extraction: Include PDF metadata (author, title, etc.) in the Word document.
License
This project is open-source under the MIT License. You are free to use, modify, and distribute it as long as attribution is provided.
