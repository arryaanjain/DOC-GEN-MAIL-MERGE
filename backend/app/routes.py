from flask import Blueprint, request, send_file
from werkzeug.utils import secure_filename
import os
import shutil
from .utils import convert_docx_to_xlsx

routes = Blueprint('routes', __name__)

@routes.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return {'error': 'No file part'}, 400

    file = request.files['file']
    
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
            convert_docx_to_xlsx(docx_path, xlsx_path)
            response = send_file(xlsx_path, as_attachment=True)
        finally:
            # Cleanup: Remove the tmp directory and its contents
            shutil.rmtree(tmp_dir, ignore_errors=True)

        return response
    return {'error': 'Invalid file type. Only DOCX files are accepted.'}, 400