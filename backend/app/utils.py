def convert_docx_to_xlsx(docx_file_path, xlsx_file_path):
    from docx import Document
    import pandas as pd

    # Read the DOCX file
    doc = Document(docx_file_path)
    data = []

    # Extract text from each paragraph in the DOCX file
    for para in doc.paragraphs:
        data.append([para.text])

    # Create a DataFrame and write to an XLSX file
    df = pd.DataFrame(data, columns=["Text"])
    df.to_excel(xlsx_file_path, index=False)

def save_uploaded_file(file, upload_folder):
    import os
    file_path = os.path.join(upload_folder, file.filename)
    file.save(file_path)
    return file_path