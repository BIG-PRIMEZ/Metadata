Document Metadata Extractor
A Python application that extracts metadata from various document types, computes a hash of the metadata, and stores it in an SQL database with timestamps. The system includes a graphical user interface for easy document selection, metadata viewing, and database operations.

Show Image
<!-- Add a screenshot of your application -->

Features
Multi-format Document Support:
PDF (.pdf)
Word Documents (.docx)
Excel Spreadsheets (.xlsx)
PowerPoint Presentations (.pptx)
Images (.jpg, .jpeg, .png)
Email Files (.eml)
CSV Files (.csv)
Comprehensive Metadata Extraction:
Author, creation date, and modification information
Document-specific properties (page count, slide count, etc.)
EXIF data from images
Content previews where applicable
Security & Verification:
SHA-256 hashing of metadata for integrity checks
Timestamps for audit trails
Database Integration:
SQLite storage for easy deployment
View historical metadata records
User-Friendly Interface:
Intuitive document selection
Formatted metadata display
Hash visualization
Database record browsing
Installation
Prerequisites
Python 3.7 or higher
Setup Instructions
Clone the repository:
bash
git clone https://github.com/yourusername/document-metadata-extractor.git
cd document-metadata-extractor
Run the setup script to create a virtual environment and install dependencies:
bash
python setup.py
This script will:
Create a virtual environment in the venv directory
Install all required dependencies
Alternative manual setup:
bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
Usage
Running the Application
With virtual environment already set up:
bash
# On Windows:
.\venv\Scripts\python.exe main.py

# On macOS/Linux:
./venv/bin/python main.py
Alternative (with activated virtual environment):
bash
# Activate virtual environment first
# On Windows:
.\venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate

# Then run
python main.py
Using the Application
Select a document:
Click the "Browse" button and select a supported file type
Extract metadata:
Click "Extract Metadata" to process the selected file
View the extracted metadata in the central text area
The metadata hash will appear in the field below
Save to database:
Click "Save to Database" to store the metadata with a timestamp
View records:
Click "View All Records" to see previously saved metadata
Select any record to view its detailed metadata
Project Structure
document-metadata-extractor/
├── main.py                         # Main entry point
├── setup.py                        # Virtual environment setup script
├── requirements.txt                # Dependencies list
├── document_metadata_extractor.py  # Core application code
└── venv/                           # Virtual environment directory (created by setup.py)
Dependencies
PyPDF2 - PDF file handling
python-docx - Word document processing
python-pptx - PowerPoint presentation processing
Pillow - Image file handling and EXIF extraction
openpyxl - Excel spreadsheet processing
tkinter - GUI framework (included in Python standard library)
How It Works
Metadata Extraction Process
Document Selection: The user selects a document through the GUI
Format Detection: The system identifies the document type based on file extension
Extraction: Document-specific extractors pull relevant metadata
Hashing: A SHA-256 hash is generated from the metadata JSON
Storage: Metadata, hash, filename, and timestamp are saved to SQLite
Core Components
DocumentMetadataExtractor: Contains format-specific extraction methods
MetadataHasher: Creates cryptographic hashes of metadata
DatabaseManager: Handles SQLite operations
MetadataExtractorGUI: Provides the user interface
Troubleshooting
Common Issues
Missing dependencies:
ModuleNotFoundError: No module named 'xyz'
Solution: Ensure you're running the application from the virtual environment or run setup.py again
Unsupported file type: If you receive an "Unsupported file type" error, check that your file has one of the supported extensions
Database errors: If you encounter database issues, try deleting the document_metadata.db file and restart the application
Getting Help
For bugs or feature requests, please open an issue on GitHub.

Future Improvements
Support for additional file formats (e.g., audio, video, archives)
Enhanced metadata extraction depth
Bulk document processing
Export capabilities for metadata records
Cloud database integration
License
This project is licensed under the MIT License - see the LICENSE file for details.

Acknowledgements
Thanks to all the open-source libraries that made this project possible
Inspiration from digital forensics tools and document management systems
