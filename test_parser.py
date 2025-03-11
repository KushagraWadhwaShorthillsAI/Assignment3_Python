import unittest
import os
from unittest.mock import MagicMock, patch

from main import (
    FileLoader, PDFLoader, DOCXLoader, PPTLoader,
    DataExtractor, FileStorage, SQLStorage
)

class TestFileLoader(unittest.TestCase):
    """Basic tests for FileLoader implementations."""
    
    def test_pdf_loader(self):
        """Test PDF loader initialization."""
        with patch('fitz.open'), patch('pdfplumber.open'):
            # Create a simple loader
            loader = PDFLoader("test.pdf")
            
            # Check basic attributes
            self.assertEqual(loader.file_path, "test.pdf")
            self.assertEqual(loader.metadata['file_name'], "test.pdf")
            self.assertEqual(loader.metadata['file_type'], ".pdf")
    
    def test_docx_loader(self):
        """Test DOCX loader initialization."""
        with patch('docx.Document'):
            # Create a simple loader
            loader = DOCXLoader("test.docx")
            
            # Check basic attributes
            self.assertEqual(loader.file_path, "test.docx")
            self.assertEqual(loader.metadata['file_name'], "test.docx")
            self.assertEqual(loader.metadata['file_type'], ".docx")
    
    def test_ppt_loader(self):
        """Test PPT loader initialization."""
        with patch('pptx.Presentation'):
            # Create a simple loader
            loader = PPTLoader("test.pptx")
            
            # Check basic attributes
            self.assertEqual(loader.file_path, "test.pptx")
            self.assertEqual(loader.metadata['file_name'], "test.pptx")
            self.assertEqual(loader.metadata['file_type'], ".pptx")


class TestPDFLoader(unittest.TestCase):
    """Basic tests for PDFLoader methods."""
    
    def setUp(self):
        """Set up test environment."""
        # Create patches
        self.fitz_patch = patch('fitz.open')
        self.pdfplumber_patch = patch('pdfplumber.open')
        
        # Start patches
        self.mock_fitz = self.fitz_patch.start()
        self.mock_pdfplumber = self.pdfplumber_patch.start()
        
        # Setup mock document and page
        self.mock_doc = MagicMock()
        self.mock_page = MagicMock()
        
        # Configure mocks
        self.mock_fitz.return_value = self.mock_doc
        self.mock_doc.__iter__.return_value = [self.mock_page]
        
        # Mock text data
        self.mock_page.get_text.return_value = {
            "blocks": [
                {"lines": [{"spans": [{"text": "Test Text"}]}]}
            ]
        }
        
        # Create loader
        self.loader = PDFLoader("test.pdf")
    
    def tearDown(self):
        """Clean up test patches."""
        self.fitz_patch.stop()
        self.pdfplumber_patch.stop()
    
    def test_extract_text(self):
        """Test PDF text extraction."""
        result = self.loader.extract_text()
        
        # Check result structure
        self.assertIn("text", result)
        self.assertIn("metadata", result)
        
        # Check text content
        self.assertIn(1, result["text"])  # Page 1 exists
    
    def test_extract_links(self):
        """Test PDF link extraction."""
        # Mock links
        self.mock_page.get_links.return_value = [
            {"uri": "http://example.com", "page": 1}
        ]
        
        result = self.loader.extract_links()
        
        # Check results
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["url"], "http://example.com")


class TestDataExtractor(unittest.TestCase):
    """Basic tests for DataExtractor."""
    
    def setUp(self):
        """Create a mock loader and extractor."""
        self.mock_loader = MagicMock(spec=FileLoader)
        self.extractor = DataExtractor(self.mock_loader)
        
        # Set up mock return values
        self.mock_loader.extract_text.return_value = {"text": {1: ["Test text"]}}
        self.mock_loader.extract_links.return_value = [{"url": "http://example.com"}]
        self.mock_loader.extract_images.return_value = [{"image_path": "image.png"}]
        self.mock_loader.extract_tables.return_value = [{"table": [["A", "B"]]}]
    
    def test_extract_text(self):
        """Test text extraction."""
        result = self.extractor.extract_text()
        
        # Check result
        self.mock_loader.extract_text.assert_called_once()
        self.assertEqual(result, {"text": {1: ["Test text"]}})
    
    def test_extract_links(self):
        """Test link extraction."""
        result = self.extractor.extract_links()
        
        # Check result
        self.mock_loader.extract_links.assert_called_once()
        self.assertEqual(result, [{"url": "http://example.com"}])


class TestFileStorage(unittest.TestCase):
    """Basic tests for FileStorage."""
    
    def setUp(self):
        """Set up test environment."""
        self.storage = FileStorage()
        
        # Mock test data
        self.test_data = {
            "text": {"text": {1: ["Test content"]}},
            "links": [{"url": "http://example.com"}]
        }
    
    @patch('os.makedirs')
    @patch('builtins.open', create=True)
    def test_save(self, mock_open, mock_makedirs):
        """Test saving data to files."""
        # Mock file handles
        mock_file = MagicMock()
        mock_open.return_value.__enter__.return_value = mock_file
        
        # Save data
        self.storage.save(self.test_data, "test_file")
        
        # Check if directories were created
        mock_makedirs.assert_called()
        
        # Check if file was opened for writing
        mock_open.assert_called()
        
        # Check if data was written
        mock_file.write.assert_called()


class TestSQLStorage(unittest.TestCase):
    """Basic tests for SQLStorage."""
    
    def setUp(self):
        """Set up test environment."""
        # Use in-memory database
        self.storage = SQLStorage(":memory:")
        
        # Test data
        self.test_data = {
            "text": {
                "text": {1: ["Test content"]},
                "metadata": {"file_path": "test.pdf"}
            },
            "links": [{"url": "http://example.com"}]
        }
    
    def test_save_and_query(self):
        """Test saving and querying data."""
        # Save test data
        self.storage.save(self.test_data, "test_file")
        
        # Query the data
        result = self.storage.query_document(file_name="test_file")
        
        # Check results
        self.assertIn("document", result)
        self.assertEqual(result["document"]["file_name"], "test_file")
    
    def test_list_documents(self):
        """Test listing documents."""
        # Save test data
        self.storage.save(self.test_data, "test_file1")
        self.storage.save(self.test_data, "test_file2")
        
        # List documents
        documents = self.storage.list_documents()
        
        # Check results
        self.assertEqual(len(documents), 2)
        self.assertIn("test_file1", [doc["file_name"] for doc in documents])
        self.assertIn("test_file2", [doc["file_name"] for doc in documents])
    
    def test_delete_document(self):
        """Test deleting a document."""
        # Save test data
        self.storage.save(self.test_data, "test_file")
        
        # Get document ID
        documents = self.storage.list_documents()
        doc_id = documents[0]["id"]
        
        # Delete document
        self.storage.delete_document(doc_id)
        
        # Check if deleted
        documents_after = self.storage.list_documents()
        self.assertEqual(len(documents_after), 0)


class TestIntegration(unittest.TestCase):
    """Simple integration test."""
    
    @patch('pptx.Presentation')
    def test_basic_workflow(self, mock_presentation):
        """Test basic workflow: load, extract, save."""
        # Mock presentation
        mock_pres = MagicMock()
        mock_presentation.return_value = mock_pres
        
        # Mock slide with text
        mock_slide = MagicMock()
        mock_slide.shapes.title.text = "Test Title"
        mock_pres.slides = [mock_slide]
        
        # Create components
        loader = PPTLoader("test.pptx")
        extractor = DataExtractor(loader)
        storage = SQLStorage(":memory:")
        
        # Extract data
        data = {
            "text": extractor.extract_text(),
            "links": extractor.extract_links()
        }
        
        # Save data
        storage.save(data, "test")
        
        # Query data
        result = storage.query_document(file_name="test")
        
        # Check results
        self.assertEqual(result["document"]["file_name"], "test")


if __name__ == '__main__':
    unittest.main()