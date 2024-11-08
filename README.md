# 📄 PDF to Word Converter

A Python application that converts PDF files to Word documents, preserving both text and images. This tool is ideal for transferring content from PDFs with a mix of text and images into editable Word documents.

## ✨ Features

- 🔹 **Text Extraction**: Retrieves and organizes text content from each PDF page.
- 🔹 **Image Extraction**: Captures embedded images from PDFs, saving them within the Word document.
- 🔹 **Word Document Generation**: Creates a `.docx` file, formatting the extracted text and images in a readable, organized structure.

---

## 📦 Installation

### Prerequisites

Ensure you have [Python 3.6+](https://www.python.org/downloads/) installed.

### Required Libraries

Install dependencies using pip:

```bash
pip install PyPDF2 PyMuPDF python-docx Pillow
```
PyPDF2: Extracts text from PDF pages.
PyMuPDF (fitz): Extracts images from PDF pages, including complex image formats.
python-docx: Allows creation and formatting of Word documents.
Pillow: Converts images to compatible formats for python-docx.

---

# 🚀 Usage
Steps

* 1-Clone the Repository:

```bash
git clone https://github.com/your-username/pdf-to-word-converter.git
cd pdf-to-word-converter
```

* 2-Add Your PDF:

  Place the PDF file you wish to convert in the project directory.

* 3-Run the Script:

```bash
Copy code
python main.py
```

* 4-Result:

  The program will create an output Word document (output_word_file.docx) with the extracted content.

# Example Code Usage

```
# Running the conversion function
pdf_to_word("PDF.pdf", "output_word_file.docx")
```

# 📄 Output
Text: All text from the PDF is extracted and organized page by page.
Images: Each image is saved and inserted in its respective location.
Page Layout: The output Word document has page headers, images, and page breaks to mimic the PDF's original structure.

# 📁 Project Structure

```
pdf-to-word-converter/
├── main.py            # Main script for conversion
├── README.md          # Project README
└── requirements.txt   # List of required packages
```

# 📝 Code Overview

 ` 1. extract_text_with_pypdf2(pdf_path)`

 * Uses PyPDF2 to extract and organize text from each page in the PDF.

 ` 2. extract_images_with_pymupdf(pdf_path)`

 * Uses PyMuPDF to extract images from each page, converting them to a compatible format (PNG).

 ` 3. pdf_to_word(pdf_path, word_path)`

 * Combines extracted text and images into a .docx file, structuring content by page and adding page breaks.

# ⚙️ Error Handling

* Image Format Compatibility: Converts images to PNG format to prevent compatibility issues with python-docx.
* Missing Content Handling: Adds page headers indicating missing text or images if any are unavailable on a page.
  
# 📈 Future Enhancements
* Customizable Layout: Add options to specify page layouts and image sizes.
* Enhanced Metadata: Capture and insert PDF metadata (author, title) in the Word document.
* Progress Indicators: Add progress bars for large PDF files to improve user experience.
# 📜 License
* This project is open-source under the MIT License. Feel free to use, modify, and distribute it with proper attribution.

# 💬 Contact
For any questions or suggestions, please reach out:

* Email: m90.rahmati@gmail.com
* GitHub: @kezb90
