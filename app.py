from flask import Flask, request, render_template, send_file
import os
import json
import pythoncom
from excelParsing import ExcelVBAProcessor
from openaiClient import OPENAI_CLIENT

app = Flask(__name__)

def load_config():
    with open("config.json", 'r') as file:
        config = json.load(file)
    return config

@app.route('/')
def upload_file():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return "No file part"
    
    file = request.files['file']
    if file.filename == '':
        return "No selected file"
    
    if file:
        uploads_dir = "uploads"
        if not os.path.exists(uploads_dir):
            os.makedirs(uploads_dir)
        file_path = os.path.join(uploads_dir, file.filename)
        file.save(file_path)
        config = load_config()
        openaiClient = OPENAI_CLIENT(
            azure_endpoint=config['openaiEndPoint'],
            api_key=config['openaiKey'],
            api_version=config['openaiVersion'],
            model=config['model']
        )

        # Initialize COM
        pythoncom.CoInitialize()  

        try:
            processor = ExcelVBAProcessor(file_path, openaiClient)
            processor.process_excel_file()
            
            # After processing, show a message with options
            return render_template('conversion_complete.html', download_file=processor.output_path)
        finally:
            pythoncom.CoUninitialize() 

@app.route('/download/<path:filename>', methods=['GET'])
def download(filename):
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
