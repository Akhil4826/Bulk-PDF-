# PDF Updater - Documentation

## Table of Contents

1. [Introduction](#introduction)
2. [Installation](#installation)
   - [System Requirements](#system-requirements)
   - [Dependencies](#dependencies)
   - [Platform-Specific Setup](#platform-specific-setup)
3. [Architecture](#architecture)
   - [Component Overview](#component-overview)
   - [File Processing Workflow](#file-processing-workflow)
4. [API Reference](#api-reference)
   - [REST Endpoints](#rest-endpoints)
   - [Job Status Lifecycle](#job-status-lifecycle)
5. [User Guide](#user-guide)
   - [Web Interface](#web-interface)
   - [Processing Options](#processing-options)
   - [Handling Large Files](#handling-large-files)
6. [Development Guide](#development-guide)
   - [Setting Up a Development Environment](#setting-up-a-development-environment)
   - [Adding Features](#adding-features)
   - [Code Organization](#code-organization)
7. [Troubleshooting](#troubleshooting)
   - [Common Issues](#common-issues)
   - [Error Messages](#error-messages)
   - [Logging](#logging)
8. [Technical Deep Dive](#technical-deep-dive)
   - [Word to PDF Conversion](#word-to-pdf-conversion)
   - [PDF Content Manipulation](#pdf-content-manipulation)
   - [Performance Considerations](#performance-considerations)
9. [Security Considerations](#security-considerations)
   - [File Handling](#file-handling)
   - [Temporary Files](#temporary-files)
10. [Deployment](#deployment)
    - [Development Server](#development-server)
    - [Production Deployment](#production-deployment)
    - [Docker Deployment](#docker-deployment)

## Introduction

PDF Updater is a cross-platform web application designed to modify PDF files by adding content from Microsoft Word documents. It allows users to append, prepend, or replace PDF content while preserving all formatting, images, and layout in the resulting PDFs.

The application is built with Python and Flask, using PyMuPDF for PDF operations and various methods for Word-to-PDF conversion depending on the platform.

## Installation

### System Requirements

- **Operating System**: Windows, macOS, or Linux
- **Python**: Version 3.7 or higher
- **Memory**: At least 50MB available RAM, more recommended for large documents
- **Disk Space**: At least 100MB free space for installation, plus additional space for document processing
- **Word Processing Software**: 
  - Windows: Microsoft Word (recommended) or LibreOffice
  - macOS/Linux: LibreOffice

### Dependencies

The application requires the following primary Python packages:

- **Flask**: Web framework
- **PyMuPDF**: PDF manipulation
- **pywin32**: Windows-specific Office automation (Windows only)
- **docx2pdf**: Fallback Word-to-PDF conversion

To install all dependencies:

```bash
pip install -r requirements.txt
```

Contents of `requirements.txt`:

```
Flask>=2.0.1
PyMuPDF>=1.19.0
Werkzeug>=2.0.1
pywin32>=301;platform_system=="Windows"
docx2pdf>=0.1.7
```

### Platform-Specific Setup

#### Windows

For optimal performance on Windows, ensure Microsoft Word is installed. The application will use Word's COM interface for the most reliable conversion.

Alternatively, LibreOffice can be installed:
1. Download and install from [LibreOffice website](https://www.libreoffice.org/download/download/)
2. Ensure the executable is in your system PATH

#### macOS

On macOS, LibreOffice is recommended:
1. Download and install from [LibreOffice website](https://www.libreoffice.org/download/download/)
2. The default installation path is usually detected automatically

#### Linux

On Linux, install LibreOffice using your distribution's package manager:

For Debian/Ubuntu:
```bash
sudo apt update
sudo apt install libreoffice
```

For Fedora/RHEL:
```bash
sudo dnf install libreoffice
```

## Architecture

### Component Overview

The PDF Updater consists of the following main components:

1. **Web Interface**: HTML/CSS/JavaScript frontend for user interaction
2. **Flask Server**: Python web server handling requests and responses
3. **File Processor**: Core functionality for converting and manipulating documents
4. **Job Manager**: Handles background processing and status tracking

![Architecture Diagram](architecture.png)

### File Processing Workflow

The application follows this workflow for processing files:

1. User uploads a Word document and one or more PDF files
2. Files are saved to a temporary directory
3. A unique job ID is generated
4. Processing begins in a background thread:
   - Word document is converted to PDF
   - For each target PDF, the content is merged according to the selected position
   - Results are verified for integrity
   - A ZIP file is created with the processed files
5. User can check job status and download results when complete
6. Temporary files are cleaned up after a defined period (default 1 hour)

## API Reference

### REST Endpoints

#### Upload Files
- **URL**: `/api/upload`
- **Method**: `POST`
- **Content-Type**: `multipart/form-data`
- **Parameters**:
  - `wordDocument`: Word document file (.doc or .docx)
  - `pdfFiles[]`: One or more PDF files
  - `contentPosition`: One of "append", "prepend", or "replace"
- **Success Response**:
  - **Code**: `200 OK`
  - **Content**: 
    ```json
    {
      "job_id": "uuid-string",
      "message": "Files uploaded successfully. Processing started."
    }
    ```
- **Error Response**:
  - **Code**: `400 Bad Request`
  - **Content**: 
    ```json
    {
      "error": "Error message describing the issue"
    }
    ```

#### Check Job Status
- **URL**: `/api/job/{job_id}`
- **Method**: `GET`
- **Success Response**:
  - **Code**: `200 OK`
  - **Content**: 
    ```json
    {
      "status": "Processing: 3/10 - document.pdf",
      "progress": 45,
      "state": "processing"
    }
    ```
  - When completed:
    ```json
    {
      "status": "Completed: 10/10 PDFs processed successfully",
      "progress": 100,
      "state": "completed",
      "download_url": "/api/download/{job_id}"
    }
    ```
- **Error Response**:
  - **Code**: `404 Not Found`
  - **Content**: 
    ```json
    {
      "error": "Job not found"
    }
    ```

#### Download Results
- **URL**: `/api/download/{job_id}`
- **Method**: `GET`
- **Success Response**:
  - **Code**: `200 OK`
  - **Content**: Binary ZIP file
- **Error Response**:
  - **Code**: `404 Not Found`
  - **Content**: 
    ```json
    {
      "error": "Results not found"
    }
    ```

#### Check Requirements
- **URL**: `/api/check-requirements`
- **Method**: `GET`
- **Success Response**:
  - **Code**: `200 OK`
  - **Content**: 
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

### Job Status Lifecycle

Job status follows this lifecycle:

1. **Initializing**: Job created, files uploaded
2. **Converting Word document to PDF**: First stage of processing
3. **Processing: X/Y - filename.pdf**: Processing individual PDFs
4. **Creating ZIP archive**: Packaging results
5. **Completed: X/Y PDFs processed successfully**: Processing finished
6. **Error: [error message]**: If processing fails

## User Guide

### Web Interface

The web interface is designed to be intuitive with the following sections:

1. **Requirements Check**: Displays warnings if system requirements are not met
2. **Word Document Upload**: Drag and drop or select a Word document
3. **PDF Files Upload**: Drag and drop or select PDF files for processing
4. **Content Position**: Choose where to place the Word content in the PDFs
5. **Process Button**: Start processing the files
6. **Job Status**: View progress and download results when complete

### Processing Options

The application offers three options for combining Word content with PDFs:

1. **Append**: Adds the Word content after the existing PDF content
   - Useful for adding terms and conditions, appendices, or supplementary information
   - Original PDF pagination is preserved

2. **Prepend**: Adds the Word content before the existing PDF content
   - Useful for adding cover pages, introductions, or table of contents
   - Original PDF begins after the Word content

3. **Replace**: Replaces the PDF content with the Word content
   - Useful for updating outdated documents with new content
   - Original PDF content is discarded

### Handling Large Files

When working with large files, keep in mind:

- **File Size Limit**: The default limit is 50MB per file
- **Processing Time**: Large or complex documents take longer to process
- **Memory Usage**: Processing large files requires more system memory
- **Batch Size**: For very large batches, consider processing in smaller groups

## Development Guide

### Setting Up a Development Environment

1. Clone the repository:
```bash
git clone https://github.com/yourusername/pdf-updater.git
cd pdf-updater
```

2. Create and activate a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Run the development server:
```bash
python app.py
```

### Adding Features

To add new features to the application:

1. **Backend Modifications**:
   - Update `app.py` to add new routes or functionality
   - Follow existing patterns for error handling and job management

2. **Frontend Modifications**:
   - Update `templates/index.html` for UI changes
   - Add JavaScript functions to handle new interactions

3. **Testing**:
   - Test with various document types and sizes
   - Verify cross-platform compatibility if applicable

### Code Organization

The codebase is organized as follows:

```
pdf-updater/
├── app.py              # Main application with Flask routes and core logic
├── requirements.txt    # Python dependencies
├── static/             # Static assets
│   ├── css/            # Stylesheets
│   ├── js/             # JavaScript files
│   └── img/            # Images
├── templates/          # HTML templates
│   └── index.html      # Main application interface
└── tests/              # Test files (if applicable)
```

## Troubleshooting

### Common Issues

#### Word to PDF Conversion Failures

**Symptoms**: 
- "Failed to convert Word document to PDF" error
- Conversion starts but never completes

**Possible Solutions**:
1. Verify Word or LibreOffice is properly installed
2. Check for macros or complex features in the Word document
3. Try simplifying the document if it's very complex
4. Ensure proper permissions for accessing the application and files

#### PDF Processing Errors

**Symptoms**:
- Processing completes but output PDFs are corrupted
- "Error processing PDF" messages

**Possible Solutions**:
1. Ensure PDFs are not encrypted or password-protected
2. Check for unusual PDF features like forms or 3D content
3. Verify PDF compliance with standard PDF versions
4. Try with a simpler PDF to isolate the issue

#### Memory or Performance Issues

**Symptoms**:
- Application crashes with large files
- Very slow processing times

**Possible Solutions**:
1. Increase available system memory
2. Process fewer files in one batch
3. Simplify complex documents
4. Check for other resource-intensive applications running concurrently

### Error Messages

| Error Message | Possible Cause | Solution |
|---------------|----------------|----------|
| "No Word document provided" | Missing Word file in upload | Ensure a Word document is selected |
| "Invalid Word document format" | Uploaded file isn't a .doc or .docx | Use a proper Word document format |
| "Failed to convert Word document to PDF" | Word conversion issues | Check Word installation and document complexity |
| "Word document converted, but resulted in invalid PDF" | Conversion produced corrupted output | Simplify the Word document and try again |
| "Job not found" | Accessing an expired or invalid job ID | Start a new processing job |
| "File too large" | File exceeds size limit | Reduce file size or increase the limit in config |

### Logging

The application logs important events to help with troubleshooting:

- Runtime errors are printed to standard output when running in debug mode
- Convert and processing failures are logged with details
- For additional debugging, enable Flask's debug mode in development

## Technical Deep Dive

### Word to PDF Conversion

The application uses a hierarchical approach to Word-to-PDF conversion:

1. **Windows COM Automation** (Windows only):
   - Uses `win32com.client` to control Microsoft Word
   - Most reliable method on Windows with Word installed
   - Preserves all formatting and elements

2. **LibreOffice Command Line**:
   - Uses LibreOffice's headless mode via subprocess calls
   - Cross-platform solution for Windows, macOS, and Linux
   - Good preservation of formatting and elements
   - Searches multiple common installation paths for compatibility

3. **docx2pdf Package** (fallback):
   - Python package that uses various methods for conversion
   - On Windows, requires Microsoft Word
   - On other platforms, requires LibreOffice

### PDF Content Manipulation

PDF manipulation uses PyMuPDF (a Python binding for MuPDF) with these approaches:

1. **PDF Insertion** (for append/prepend):
   - Uses `insert_pdf()` function to combine documents
   - Maintains bookmarks, links, and other PDF features
   - Preserves original formatting and layout

2. **Quality Preservation**:
   - Uses optimal compression settings with `garbage=4, deflate=True`
   - Cleans up redundant data with `clean=True`
   - Maintains image quality in the output files

3. **PDF Integrity Verification**:
   - Checks source and result PDFs to ensure they can be opened
   - Reports page counts and validation status
   - Helps identify issues before delivering results to users

### Performance Considerations

The application includes several optimizations for performance:

1. **Background Processing**:
   - Uses threading to handle jobs in the background
   - Allows multiple concurrent jobs without blocking the web interface
   - Provides progress updates via the status API

2. **Resource Management**:
   - Temporary files are stored in system temp directories
   - Automatic cleanup after a configurable period (default 1 hour)
   - Proper closing of document objects to release memory

3. **Platform-Specific Optimizations**:
   - Uses the most efficient conversion method for each platform
   - Implements platform detection to choose appropriate paths

## Security Considerations

### File Handling

The application incorporates these security measures for file handling:

1. **Filename Sanitization**:
   - Uses `secure_filename()` from Werkzeug to sanitize uploaded filenames
   - Prevents directory traversal attacks
   - Removes potentially harmful characters

2. **File Type Validation**:
   - Validates file extensions for both Word documents and PDFs
   - Checks file content types when possible
   - Rejects files that don't match expected formats

3. **Size Limitations**:
   - Default 50MB limit prevents resource exhaustion
   - Configurable through Flask's `MAX_CONTENT_LENGTH`
   - Clear error messages for oversized files

### Temporary Files

Management of temporary files follows these practices:

1. **Isolation**:
   - Each job gets its own directory with a UUID identifier
   - Prevents interference between concurrent jobs
   - Simplifies cleanup

2. **Automatic Cleanup**:
   - Background thread removes temporary files after a set period
   - Prevents accumulation of temporary files on the server
   - Occurs regardless of job completion status

3. **Access Control**:
   - Job IDs are randomly generated UUIDs
   - Direct file system access is not exposed through the web interface
   - Downloads only available for valid job IDs

## Deployment

### Development Server

For development and testing:

```bash
python app.py
```

This starts a development server on `http://localhost:5000`

### Production Deployment

For production, use a proper WSGI server:

#### Example with Gunicorn

1. Install Gunicorn:
```bash
pip install gunicorn
```

2. Run with Gunicorn:
```bash
gunicorn -w 4 -b 0.0.0.0:8000 app:app
```

#### Example with uWSGI

1. Install uWSGI:
```bash
pip install uwsgi
```

2. Create a uwsgi.ini file:
```ini
[uwsgi]
module = app:app
master = true
processes = 4
socket = 0.0.0.0:8000
vacuum = true
die-on-term = true
```

3. Run with uWSGI:
```bash
uwsgi --ini uwsgi.ini
```

### Docker Deployment

For containerized deployment, use the provided Dockerfile:

1. Build the Docker image:
```bash
docker build -t pdf-updater .
```

2. Run the container:
```bash
docker run -p 8000:8000 pdf-updater
```

#### Dockerfile Example

```dockerfile
FROM python:3.9-slim

RUN apt-get update && apt-get install -y \
    libreoffice \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV FLASK_APP=app.py
ENV FLASK_ENV=production

EXPOSE 8000

CMD ["gunicorn", "--bind", "0.0.0.0:8000", "app:app"]
```

---

This documentation provides a comprehensive overview of the PDF Updater application, its installation, usage, architecture, and technical details. For further assistance or to report issues, please contact the developer or submit issues through the project's GitHub repository.
