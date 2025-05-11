from flask import Flask, render_template, request, jsonify, send_file
import os
import tempfile
import uuid
import shutil
import fitz  # PyMuPDF
from docx2pdf import convert
from werkzeug.utils import secure_filename
import zipfile
import threading
import time

app = Flask(__name__)

# Configuration
UPLOAD_FOLDER = os.path.join(tempfile.gettempdir(), 'pdf_updater_uploads')
RESULTS_FOLDER = os.path.join(tempfile.gettempdir(), 'pdf_updater_results')
ALLOWED_EXTENSIONS_WORD = {'docx'}
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

def word_to_pdf(word_path, output_path):
    """Convert Word document to PDF"""
    try:
        convert(word_path, output_path)
        return True
    except Exception as e:
        print(f"Error converting Word to PDF: {str(e)}")
        return False

def process_pdf(target_pdf, overlay_pdf, output_path, position='append'):
    """Process a single PDF file with the content from another PDF"""
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
            
        # Save the resulting document
        target_document.save(output_path)
        target_document.close()
        
        if position != 'replace':
            overlay_document.close()
            
        return True
    except Exception as e:
        print(f"Error processing PDF: {str(e)}")
        return False

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
        if not word_to_pdf(word_file_path, temp_pdf):
            job_info['status'] = "Failed to convert Word document to PDF"
            job_info['state'] = "error"
            return
            
        # Step 2: Process each PDF file
        total_files = len(pdf_files)
        processed = 0
        successful = 0
        
        for i, pdf_path in enumerate(pdf_files):
            pdf_filename = os.path.basename(pdf_path)
            output_path = os.path.join(result_folder, pdf_filename)
            
            job_info['status'] = f"Processing: {i+1}/{total_files} - {pdf_filename}"
            job_info['progress'] = 10 + int((i / total_files) * 80)
            
            if process_pdf(pdf_path, temp_pdf, output_path, content_position):
                successful += 1
                
            processed += 1
            
        # Step 3: Create ZIP archive of results
        job_info['status'] = "Creating ZIP archive"
        job_info['progress'] = 90
        
        zip_path = os.path.join(app.config['RESULTS_FOLDER'], f"{job_id}.zip")
        with zipfile.ZipFile(zip_path, 'w') as zip_file:
            for root, dirs, files in os.walk(result_folder):
                for file in files:
                    if file != "converted_word.pdf":  # Skip the temporary conversion file
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, result_folder)
                        zip_file.write(file_path, arcname)
        
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
        return jsonify({'error': 'Invalid Word document format (must be .docx)'}), 400
    
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

@app.errorhandler(413)
def request_entity_too_large(error):
    return jsonify({'error': 'File too large. Maximum size is 50MB.'}), 413

if __name__ == '__main__':
    app.run(debug=True)