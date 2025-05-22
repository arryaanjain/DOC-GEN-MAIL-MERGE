from flask import Blueprint, request, send_file
from werkzeug.utils import secure_filename
import os
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
        docx_path = os.path.join('/tmp', filename)
        file.save(docx_path)

        xlsx_filename = filename.rsplit('.', 1)[0] + '.xlsx'
        xlsx_path = os.path.join('/tmp', xlsx_filename)

        convert_docx_to_xlsx(docx_path, xlsx_path)

        return send_file(xlsx_path, as_attachment=True)
    return {'error': 'Invalid file type. Only DOCX files are accepted.'}, 400