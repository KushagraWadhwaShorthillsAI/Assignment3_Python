import unittest
import os
from main import PDFLoader, PPTLoader, DOCXLoader, DataExtractor

class TestFileLoaders(unittest.TestCase):
    def setUp(self):
        """Setup test files paths"""
        self.pdf_file = "test_files/Report merged.pdf"
        self.ppt_file = "test_files/Chatfolio.pptx"
        self.docx_file = "test_files/cover page.docx"
        self.empty_pdf = "test_files/empty.pdf"
        self.complex_table_docx = "test_files/complex_table.docx"
        self.no_links_ppt = "test_files/no_links.pptx"
    
    def test_pdf_loader(self):
        """Test PDFLoader for text, links, images, and tables"""
        loader = PDFLoader(self.pdf_file)
        extractor = DataExtractor(loader)
        
        text_data = extractor.extract_text()
        self.assertIn("text", text_data)
        
        links_data = extractor.extract_links()
        self.assertIsInstance(links_data, list)
        
        images_data = extractor.extract_images()
        self.assertIsInstance(images_data, list)
        
        tables_data = extractor.extract_tables()
        self.assertIsInstance(tables_data, list)

    def test_ppt_loader(self):
        """Test PPTLoader for text, links, images, and tables"""
        loader = PPTLoader(self.ppt_file)
        extractor = DataExtractor(loader)
    
        self.assertIn("text", extractor.extract_text())
        self.assertIsInstance(extractor.extract_links(), list)
        self.assertIsInstance(extractor.extract_images(), list)
        self.assertIsInstance(extractor.extract_tables(), list)

    def test_docx_loader(self):
        """Test DOCXLoader for text, links, images, and tables"""
        loader = DOCXLoader(self.docx_file)
        extractor = DataExtractor(loader)

        self.assertIn("text", extractor.extract_text())
        self.assertIsInstance(extractor.extract_links(), list)
        self.assertIsInstance(extractor.extract_images(), list)
        self.assertIsInstance(extractor.extract_tables(), list)

    def test_pdf_loader_empty(self):
        """Test PDFLoader with an empty file"""
        loader = PDFLoader(self.empty_pdf)
        extractor = DataExtractor(loader)
        
        try:
            text_output = extractor.extract_text()["text"]
        except Exception as e:
            text_output = "Error: " + str(e)
        
        self.assertIn(text_output, [{1: []}, {}])  # Expecting an empty list or empty dict
        self.assertEqual(extractor.extract_links(), [])
        self.assertEqual(extractor.extract_images(), [])
        self.assertEqual(extractor.extract_tables(), [])

    def test_ppt_loader_no_links(self):
        """Test PPTLoader with a file containing no hyperlinks"""
        loader = PPTLoader(self.no_links_ppt)
        extractor = DataExtractor(loader)

        self.assertEqual(extractor.extract_links(), [])

    def test_docx_loader_complex_table(self):
        """Test DOCXLoader with a document containing complex tables"""
        loader = DOCXLoader(self.complex_table_docx)
        extractor = DataExtractor(loader)

        tables = extractor.extract_tables()
        self.assertGreater(len(tables), 0)  # Ensure at least one table is extracted
        self.assertTrue(any(len(row) > 1 for table in tables for row in table["table"]))  # Ensure tables are structured

if __name__ == "__main__":
    unittest.main()
