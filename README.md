# 📄 PDF Updater

![PDF Updater](https://img.shields.io/badge/PDF-Updater-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Python](https://img.shields.io/badge/python-3.7%2B-blue)
![Flask](https://img.shields.io/badge/flask-2.0%2B-red)

**PDF Updater** is a platform-independent web application that allows you to **insert content from Word documents into existing PDF files**, with full **formatting, image, and layout preservation**.

> 🌐 **Live Demo**: [/akhil.up.railway.app](https://akhil.up.railway.app//)

---

## 🚀 Features

- ✅ **Cross-Platform**: Works on Windows, macOS, and Linux
- 📌 **Position Control**: Add content to the **beginning**, **end**, or **replace** PDF pages
- 📂 **Batch Processing**: Upload and process **multiple PDFs** at once
- 🖼️ **Preserve Formatting**: Maintains original **text styles**, **images**, and **layouts**
- 📊 **Progress Tracking**: See live progress during file processing
- 📃 **Detailed Reports**: Summary of successful and failed conversions
- 🔐 **Secure Handling**: Uses **temporary storage** and removes files after use

---

## Screenshots

![PDF Updater Interface](https://github.com/user-attachments/assets/38ee0eeb-768b-4fb2-b290-4a220956452e)
)

## Installation

### Prerequisites

- Python 3.7 or higher
- One of the following for Word-to-PDF conversion:
  - **Windows**: Microsoft Word (recommended)
  - **All Platforms**: LibreOffice (recommended for cross-platform)
  - **Fallback**: Python `docx2pdf` package

### Basic Installation

1. Clone the repository:
```bash
git clone https://github.com/Akhil4826/PDF-akhil.git
cd pdf-updater
```

2. Create a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

### Configuration

No additional configuration is required for basic usage. The application will automatically use available conversion methods based on your platform.

## Usage

1. Start the application:
```bash
python app.py
```

2. Open your web browser and navigate to `http://localhost:5000`

3. Upload a Word document and PDF files through the interface

4. Select your desired positioning option:
   - **Append**: Add Word content to the end of each PDF
   - **Prepend**: Add Word content to the beginning of each PDF
   - **Replace**: Replace PDF content with the Word content

5. Click "Process Files" and wait for the operation to complete

6. Download the results as a ZIP file

## API Reference

### Upload Files

```
POST /api/upload
```

**Form Parameters:**
- `wordDocument`: Word document file (.doc or .docx)
- `pdfFiles[]`: One or more PDF files
- `contentPosition`: One of "append", "prepend", or "replace"

**Response:**
```json
{
  "job_id": "unique-job-identifier",
  "message": "Files uploaded successfully. Processing started."
}
```

### Check Job Status

```
GET /api/job/{job_id}
```

**Response:**
```json
{
  "status": "Processing: 3/10 - document.pdf",
  "progress": 45,
  "state": "processing"
}
```

When completed:
```json
{
  "status": "Completed: 10/10 PDFs processed successfully",
  "progress": 100,
  "state": "completed",
  "download_url": "/api/download/{job_id}"
}
```

### Download Results

```
GET /api/download/{job_id}
```

Returns a ZIP file containing the processed PDFs.

### Check Requirements

```
GET /api/check-requirements
```

**Response:**
```json
{
  "requirements": {
    "pymupdf": true,
    "word_conversion": true
  },
  "platform": "Windows",
  "python_version": "3.9.7"
}
```

## Technical Details

### Word-to-PDF Conversion Methods

The application attempts to use the most reliable method available on your platform:

1. **Windows**: Uses the `win32com` client to control Microsoft Word
2. **Any Platform**: Attempts to use LibreOffice command-line interface
3. **Fallback**: Uses the `docx2pdf` Python package

### PDF Processing

- The application uses PyMuPDF (fitz) for all PDF operations
- All content, including text, images, formatting, tables, and vector graphics, is preserved in the resulting PDFs
- PDF integrity is verified before and after processing

## System Requirements

- Python 3.7+
- 50MB+ available memory
- One of the following for Word document conversion:
  - Microsoft Word (Windows)
  - LibreOffice (any platform)
  - `docx2pdf` package (requires Microsoft Word on Windows)

## Development

### Project Structure

```
pdf-updater/
├── app.py              # Main application file
├── requirements.txt    # Python dependencies
├── static/             # Static assets (JS, CSS)
├── templates/          # HTML templates
│   └── index.html      # Main application interface
└── README.md           # This file
```

### Requirements

Required Python packages (see `requirements.txt`):
```
Flask>=2.0.1
PyMuPDF>=1.19.0
Werkzeug>=2.0.1
pywin32>=301;platform_system=="Windows"
docx2pdf>=0.1.7
```

## Troubleshooting

### Word to PDF Conversion Issues

- Ensure Microsoft Word or LibreOffice is properly installed
- On Windows, make sure your Office installation is not in a "Safe Mode" or "Protected View"
- Try converting the Word document manually first to check for compatibility issues

### PDF Processing Issues

- Ensure PDFs are not password-protected or encrypted
- Check that PDFs are valid and can be opened in a standard PDF viewer
- Large PDFs with complex layouts may take longer to process

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

Distributed under the MIT License. See `LICENSE` for more information.

## Contact

Your Name - [Akhil] - akhilteotia19@gmail.com

Project Link: [https://github.com/yourusername/pdf-updater](https://github.com/akhil4826/pdf-updater)

## Acknowledgements

- [PyMuPDF](https://github.com/pymupdf/PyMuPDF)
- [Flask](https://flask.palletsprojects.com/)
- [docx2pdf](https://github.com/AlJohri/docx2pdf)
