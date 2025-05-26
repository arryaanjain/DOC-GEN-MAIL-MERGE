from flask import Blueprint, request, send_file, after_this_request
from werkzeug.utils import secure_filename
import os
import shutil
from app.core.document_processor import DocumentProcessor

ob = DocumentProcessor(debug=True)

routes = Blueprint('routes', __name__)

@routes.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return {'error': 'No file part'}, 400

    file = request.files['file']
    processing_date = request.form.get('date')
    
    if not processing_date:
        return {'error': 'Processing date is required'}, 400
    
    if file.filename == '':
        return {'error': 'No selected file'}, 400

    if file and file.filename.endswith('.docx'):
        filename = secure_filename(file.filename)
        tmp_dir = os.path.join(os.getcwd(), 'tmp')  # Use a 'tmp' folder in the current working directory
        os.makedirs(tmp_dir, exist_ok=True)  # Ensure the directory exists
        docx_path = os.path.join(tmp_dir, filename)
        file.save(docx_path)

        xlsx_filename = filename.rsplit('.', 1)[0] + '.xlsx'
        xlsx_path = os.path.join(tmp_dir, xlsx_filename)

        try:
            ob.convert_docx_to_xlsx(docx_path, xlsx_path,  processing_date=processing_date)

            # Debug: Confirm cleanup registration
            print(f"🧹 Registering cleanup for directory: {tmp_dir}")

            # Schedule cleanup after the response is sent
            @after_this_request
            def cleanup(response):
                print(f"🧹 Cleaning up directory: {tmp_dir}")
                shutil.rmtree(tmp_dir, ignore_errors=True)
                return response

            # Debug: Confirm file is being sent
            print(f"📤 Sending file: {xlsx_path}")
            return send_file(xlsx_path, as_attachment=True, max_age=0)
        except Exception as e:
            # Cleanup in case of an error
            print(f"⚠️ Exception occurred: {e}")
            shutil.rmtree(tmp_dir, ignore_errors=True)
            return {'error': str(e)}, 500

    return {'error': 'Invalid file type. Only DOCX files are accepted.'}, 400