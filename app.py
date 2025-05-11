from flask import Flask, render_template, request, jsonify, send_file
import os
import tempfile
import uuid
import shutil
import fitz  # PyMuPDF
import pythoncom  # For COM initialization
import platform
import threading
import time
import zipfile
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Configuration
UPLOAD_FOLDER = os.path.join(tempfile.gettempdir(), 'pdf_updater_uploads')
RESULTS_FOLDER = os.path.join(tempfile.gettempdir(), 'pdf_updater_results')
ALLOWED_EXTENSIONS_WORD = {'docx', 'doc'}
ALLOWED_EXTENSIONS_PDF = {'pdf'}
MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB max upload size

# Create upload and results directories if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULTS_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['RESULTS_FOLDER'] = RESULTS_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

# Store processing jobs
active_jobs = {}

def allowed_file(filename, allowed_extensions):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions

def word_to_pdf_windows(word_path, output_path):
    """Convert Word document to PDF using COM on Windows"""
    try:
        import win32com.client
        pythoncom.CoInitialize()
        word = win32com.client.Dispatch('Word.Application')
        doc = word.Documents.Open(word_path)
        doc.SaveAs(output_path, FileFormat=17)  # 17 is PDF format
        doc.Close()
        word.Quit()
        return True
    except Exception as e:
        print(f"Error converting Word to PDF using Windows method: {str(e)}")
        return False
    finally:
        pythoncom.CoUninitialize()

def word_to_pdf_libreoffice(word_path, output_path):
    """Convert Word document to PDF using LibreOffice"""
    try:
        import subprocess
        # Try to detect the LibreOffice executable
        soffice_paths = [
            'soffice',  # Linux/macOS default if in PATH
            '/usr/bin/soffice',
            '/usr/lib/libreoffice/program/soffice',
            '/Applications/LibreOffice.app/Contents/MacOS/soffice',  # macOS
            'C:\\Program Files\\LibreOffice\\program\\soffice.exe',  # Windows
            'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe'
        ]
        
        soffice_path = None
        for path in soffice_paths:
            try:
                subprocess.run([path, '--version'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False)
                soffice_path = path
                break
            except (FileNotFoundError, subprocess.SubprocessError):
                continue
                
        if not soffice_path:
            raise Exception("LibreOffice not found. Please install LibreOffice.")
        
        result = subprocess.run([
            soffice_path,
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', os.path.dirname(output_path),
            word_path
        ], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
        
        # LibreOffice saves with the same name but .pdf extension
        temp_output = os.path.join(
            os.path.dirname(output_path),
            os.path.splitext(os.path.basename(word_path))[0] + '.pdf'
        )
        
        # Rename to the desired output path if needed
        if temp_output != output_path and os.path.exists(temp_output):
            shutil.move(temp_output, output_path)
            
        return os.path.exists(output_path)
    except Exception as e:
        print(f"Error converting Word to PDF using LibreOffice method: {str(e)}")
        return False

def word_to_pdf(word_path, output_path):
    """Cross-platform Word to PDF conversion"""
    # First try using platform-specific methods
    if platform.system() == 'Windows':
        # Try Windows-specific method first
        if word_to_pdf_windows(word_path, output_path):
            return True
    
    # Fall back to LibreOffice (works on all platforms if installed)
    if word_to_pdf_libreoffice(word_path, output_path):
        return True
    
    # If all else fails, try docx2pdf as a last resort
    try:
        from docx2pdf import convert
        convert(word_path, output_path)
        return os.path.exists(output_path)
    except Exception as e:
        print(f"Error converting Word to PDF using docx2pdf: {str(e)}")
        return False

def process_pdf(target_pdf, overlay_pdf, output_path, position='append'):
    """Process a single PDF file with the content from another PDF,
    preserving all content including images and formatting"""
    try:
        target_document = fitz.open(target_pdf)
        overlay_document = fitz.open(overlay_pdf)
        
        if position == 'append':
            # Append the overlay document to the target
            target_document.insert_pdf(overlay_document)
        elif position == 'prepend':
            # Create a new document with the overlay first, then the target
            result_document = fitz.open()
            result_document.insert_pdf(overlay_document)
            result_document.insert_pdf(target_document)
            target_document.close()
            target_document = result_document
        else:  # replace
            # Just use the overlay document directly
            target_document.close()
            target_document = overlay_document
            
        # Save the resulting document with full compression to preserve quality
        target_document.save(output_path, garbage=4, deflate=True, clean=True)
        target_document.close()
        
        if position != 'replace':
            overlay_document.close()
            
        return True
    except Exception as e:
        print(f"Error processing PDF: {str(e)}")
        return False

def check_pdf_integrity(pdf_path):
    """Verify that a PDF can be opened and is not corrupted"""
    try:
        doc = fitz.open(pdf_path)
        page_count = len(doc)
        doc.close()
        return True, page_count
    except Exception as e:
        return False, str(e)

def process_files_task(job_id, word_file_path, pdf_files, content_position):
    """Background task to process PDF files"""
    job_info = active_jobs[job_id]
    result_folder = os.path.join(app.config['RESULTS_FOLDER'], job_id)
    os.makedirs(result_folder, exist_ok=True)
    
    try:
        # Step 1: Convert Word to PDF
        job_info['status'] = "Converting Word document to PDF"
        job_info['progress'] = 5
        
        temp_pdf = os.path.join(result_folder, "converted_word.pdf")
        if not word_to_pdf(os.path.abspath(word_file_path), os.path.abspath(temp_pdf)):
            job_info['status'] = "Failed to convert Word document to PDF"
            job_info['state'] = "error"
            return
        
        # Verify the converted PDF
        is_valid, page_count = check_pdf_integrity(temp_pdf)
        if not is_valid:
            job_info['status'] = "Word document converted, but resulted in invalid PDF"
            job_info['state'] = "error"
            return
            
        # Step 2: Process each PDF file
        total_files = len(pdf_files)
        processed = 0
        successful = 0
        failed_files = []
        
        for i, pdf_path in enumerate(pdf_files):
            pdf_filename = os.path.basename(pdf_path)
            output_path = os.path.join(result_folder, pdf_filename)
            
            job_info['status'] = f"Processing: {i+1}/{total_files} - {pdf_filename}"
            job_info['progress'] = 10 + int((i / total_files) * 80)
            
            # Verify source PDF
            is_valid, _ = check_pdf_integrity(pdf_path)
            if not is_valid:
                failed_files.append(f"{pdf_filename} (Invalid PDF)")
                processed += 1
                continue
            
            if process_pdf(pdf_path, temp_pdf, output_path, content_position):
                # Verify result PDF
                is_valid, _ = check_pdf_integrity(output_path)
                if is_valid:
                    successful += 1
                else:
                    failed_files.append(pdf_filename)
            else:
                failed_files.append(pdf_filename)
                
            processed += 1
            
        # Step 3: Create ZIP archive of results
        job_info['status'] = "Creating ZIP archive"
        job_info['progress'] = 90
        
        zip_path = os.path.join(app.config['RESULTS_FOLDER'], f"{job_id}.zip")
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for root, dirs, files in os.walk(result_folder):
                for file in files:
                    if file != "converted_word.pdf":  # Skip the temporary conversion file
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, result_folder)
                        zip_file.write(file_path, arcname)
        
        # Create a report file
        report_path = os.path.join(result_folder, "processing_report.txt")
        with open(report_path, 'w') as report_file:
            report_file.write(f"PDF Processing Report\n")
            report_file.write(f"===================\n\n")
            report_file.write(f"Word document: {os.path.basename(word_file_path)}\n")
            report_file.write(f"Position: {content_position}\n")
            report_file.write(f"Total files: {total_files}\n")
            report_file.write(f"Successfully processed: {successful}\n")
            report_file.write(f"Failed: {len(failed_files)}\n\n")
            
            if failed_files:
                report_file.write("Failed files:\n")
                for failed in failed_files:
                    report_file.write(f"- {failed}\n")
        
        # Add report to zip
        with zipfile.ZipFile(zip_path, 'a', zipfile.ZIP_DEFLATED) as zip_file:
            zip_file.write(report_path, "processing_report.txt")
        
        # Update job status to completed
        job_info['status'] = f"Completed: {successful}/{total_files} PDFs processed successfully"
        job_info['progress'] = 100
        job_info['state'] = "completed"
        job_info['result_path'] = zip_path
        
    except Exception as e:
        job_info['status'] = f"Error: {str(e)}"
        job_info['state'] = "error"
        
    finally:
        # Clean up after some time (keep results for download)
        def cleanup_job():
            time.sleep(3600)  # Keep files for 1 hour
            if os.path.exists(result_folder):
                shutil.rmtree(result_folder)
            if os.path.exists(os.path.join(app.config['UPLOAD_FOLDER'], job_id)):
                shutil.rmtree(os.path.join(app.config['UPLOAD_FOLDER'], job_id))
            if job_id in active_jobs:
                del active_jobs[job_id]
                
        cleanup_thread = threading.Thread(target=cleanup_job)
        cleanup_thread.daemon = True
        cleanup_thread.start()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/upload', methods=['POST'])
def upload_files():
    if 'wordDocument' not in request.files:
        return jsonify({'error': 'No Word document provided'}), 400
        
    word_file = request.files['wordDocument']
    if word_file.filename == '':
        return jsonify({'error': 'No Word document selected'}), 400
        
    if not allowed_file(word_file.filename, ALLOWED_EXTENSIONS_WORD):
        return jsonify({'error': 'Invalid Word document format (must be .doc or .docx)'}), 400
    
    if 'pdfFiles[]' not in request.files:
        return jsonify({'error': 'No PDF files provided'}), 400
    
    pdf_files = request.files.getlist('pdfFiles[]')
    if len(pdf_files) == 0 or pdf_files[0].filename == '':
        return jsonify({'error': 'No PDF files selected'}), 400
    
    # Check PDF file formats
    for pdf_file in pdf_files:
        if not allowed_file(pdf_file.filename, ALLOWED_EXTENSIONS_PDF):
            return jsonify({'error': f'Invalid file format: {pdf_file.filename} (must be .pdf)'}), 400
    
    # Get content position option
    content_position = request.form.get('contentPosition', 'append')
    if content_position not in ['append', 'prepend', 'replace']:
        content_position = 'append'
    
    # Create job ID and save uploaded files
    job_id = str(uuid.uuid4())
    job_folder = os.path.join(app.config['UPLOAD_FOLDER'], job_id)
    os.makedirs(job_folder, exist_ok=True)
    
    # Save Word document
    word_filename = secure_filename(word_file.filename)
    word_path = os.path.join(job_folder, word_filename)
    word_file.save(word_path)
    
    # Save PDF files
    pdf_paths = []
    for pdf_file in pdf_files:
        pdf_filename = secure_filename(pdf_file.filename)
        pdf_path = os.path.join(job_folder, pdf_filename)
        pdf_file.save(pdf_path)
        pdf_paths.append(pdf_path)
    
    # Create job status and start processing
    active_jobs[job_id] = {
        'status': 'Initializing',
        'progress': 0,
        'state': 'processing',
        'files_count': len(pdf_paths),
        'word_filename': word_filename
    }
    
    # Start processing in background
    processing_thread = threading.Thread(
        target=process_files_task,
        args=(job_id, word_path, pdf_paths, content_position)
    )
    processing_thread.daemon = True
    processing_thread.start()
    
    return jsonify({
        'job_id': job_id,
        'message': 'Files uploaded successfully. Processing started.'
    })

@app.route('/api/job/<job_id>', methods=['GET'])
def get_job_status(job_id):
    if job_id not in active_jobs:
        return jsonify({'error': 'Job not found'}), 404
        
    job_info = active_jobs[job_id]
    response = {
        'status': job_info['status'],
        'progress': job_info['progress'],
        'state': job_info['state']
    }
    
    if job_info['state'] == 'completed':
        response['download_url'] = f'/api/download/{job_id}'
        
    return jsonify(response)

@app.route('/api/download/<job_id>', methods=['GET'])
def download_results(job_id):
    if job_id not in active_jobs or 'result_path' not in active_jobs[job_id]:
        return jsonify({'error': 'Results not found'}), 404
        
    job_info = active_jobs[job_id]
    zip_path = job_info['result_path']
    
    if not os.path.exists(zip_path):
        return jsonify({'error': 'Results file not found'}), 404
        
    return send_file(
        zip_path,
        mimetype='application/zip',
        as_attachment=True,
        download_name='pdf_updater_results.zip'
    )

@app.route('/api/check-requirements', methods=['GET'])
def check_requirements():
    """Check if all required components are available"""
    requirements = {
        'pymupdf': True,  # We assume PyMuPDF is installed since it's imported
        'word_conversion': False
    }
    
    # Check if any Word to PDF conversion method is available
    if platform.system() == 'Windows':
        try:
            import win32com.client
            requirements['word_conversion'] = True
        except ImportError:
            pass
    
    # Check for LibreOffice
    if not requirements['word_conversion']:
        soffice_paths = [
            'soffice',
            '/usr/bin/soffice',
            '/usr/lib/libreoffice/program/soffice',
            '/Applications/LibreOffice.app/Contents/MacOS/soffice',
            'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
            'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe'
        ]
        
        for path in soffice_paths:
            try:
                import subprocess
                subprocess.run([path, '--version'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False)
                requirements['word_conversion'] = True
                break
            except (FileNotFoundError, subprocess.SubprocessError, ImportError):
                continue
    
    # Check for docx2pdf as last resort
    if not requirements['word_conversion']:
        try:
            import docx2pdf
            requirements['word_conversion'] = True
        except ImportError:
            pass
    
    return jsonify({
        'requirements': requirements,
        'platform': platform.system(),
        'python_version': platform.python_version()
    })

@app.errorhandler(413)
def request_entity_too_large(error):
    return jsonify({'error': 'File too large. Maximum size is 50MB.'}), 413

if __name__ == '__main__':
    app.run(debug=True)
