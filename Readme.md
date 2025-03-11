# Document Processor

## Overview
This project provides a document processing tool that extracts text, links, images, and tables from PDF, DOCX, and PPTX files. The extracted data can be stored in files or an SQLite database.

## Features
- Extract text, including headings and font styles
- Extract hyperlinks from documents
- Extract images from PDFs, Word, and PowerPoint files
- Extract tables from PDFs and Word files
- Store extracted data in files or an SQLite database

## Installation

### Prerequisites
Ensure you have **Python 3.7+** installed along with the required dependencies.

### Install Dependencies
Run the following command to install required packages:

```bash
pip install -r requirements.txt
```

## Usage

### Extract Data from a Document
To extract data from a file and store the results, use the following:

```python
from main import PDFLoader, PPTLoader, DOCXLoader, DataExtractor, FileStorage, SQLStorage

# Choose a file to process
file_path = "path/to/your/document.pdf"
loader = PDFLoader(file_path)  # Change to DOCXLoader or PPTLoader accordingly

# Extract data
extractor = DataExtractor(loader)
data = {
    "text": extractor.extract_text(),
    "links": extractor.extract_links(),
    "images": extractor.extract_images(),
    "tables": extractor.extract_tables()
}

# Store the extracted data
storage = FileStorage()
storage.save(data, "output_folder")

# Save to database
storage = SQLStorage("documents.db")
storage.save(data, "output_folder")
```

## Running Tests
To ensure the code is working correctly, run the unit tests:

```bash
python -m unittest test_main.py
```

## Directory Structure
```
project-root/
│── main.py            # Main script for document processing
│── test_main.py       # Unit tests
│── requirements.txt   # Required Python packages
│── README.md          # Documentation
│── test_files/        # Sample files for testing
│── output/            # Output directory for extracted data
```

## Contributing
Feel free to contribute to this project by submitting pull requests or reporting issues!

## License
This project is licensed under the MIT License.

