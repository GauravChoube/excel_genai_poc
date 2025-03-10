import os
import json
import pandas as pd
from flask import Flask, request, jsonify, render_template
from excelParsing import ExcelVBAProcessor
from openaiClient import OPENAI_CLIENT
import pythoncom  # Import pythoncom for COM initialization
import importlib.util  # For dynamic import
import logging

# Configuration
CONFIG_PATH = "config.json"
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Set up logging
logging.basicConfig(level=logging.INFO)

app = Flask(__name__)

# Function to read the configuration
def read_config():
    try:
        with open(CONFIG_PATH, 'r') as file:
            config = json.load(file)
        
        if not config:
            raise ValueError("ERROR: Invalid config file. Please provide OpenAI credentials.")
        if 'openaiKey' not in config or not config['openaiKey']:
            raise ValueError("ERROR: Missing OpenAI Key in config file.")
        if 'openaiEndPoint' not in config or not config['openaiEndPoint']:
            raise ValueError("ERROR: Missing OpenAI Endpoint in config file.")
        if 'openaiVersion' not in config or not config['openaiVersion']:
            raise ValueError("ERROR: Missing OpenAI Version in config file.")
        
        return config
    except Exception as e:
        logging.error(f"❌ Config Error: {e}")
        return None

config = read_config()
openaiClient = None
if config:
    openaiClient = OPENAI_CLIENT(
        azure_endpoint=config['openaiEndPoint'],
        api_key=config['openaiKey'],
        api_version=config['openaiVersion'],
        model=config.get('model', 'gpt-4')
    )

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
    if file:
        file_path = os.path.abspath(os.path.join(UPLOAD_FOLDER, file.filename))
        file.save(file_path)

        if not os.path.exists(file_path):
            return jsonify({"error": f"File {file.filename} was not saved at {file_path}"}), 500

        # Process the Excel file
        try:
            pythoncom.CoInitialize()
            processor = ExcelVBAProcessor(file_path=file_path, openAIClient=openaiClient)
            processor.process_excel_file()
            pythoncom.CoUninitialize()
            return jsonify({"message": "File uploaded and processed successfully", "file_path": file_path}), 200
        except Exception as e:
            pythoncom.CoUninitialize()
            return jsonify({"error": str(e)}), 500

# @app.route('/view_data', methods=['POST'])
# def view_data():
#     file = request.files.get('file')
#     if not file:
#         return jsonify({"error": "File is required."}), 400

#     try:
#         excel_file_path = os.path.join(UPLOAD_FOLDER, file.filename)
#         file.save(excel_file_path)

#         # Dynamically import the ConvertedExcelMacros class
#         spec = importlib.util.spec_from_file_location("ConvertedExcelMacros", "converted_macros.py")
#         converted_macros_module = importlib.util.module_from_spec(spec)
#         spec.loader.exec_module(converted_macros_module)

#         # Create an instance of the ConvertedExcelMacros class
#         macros_instance = converted_macros_module.ConvertedExcelMacros()
#         data, sheets = macros_instance.view_data(excel_file_path)
#         return jsonify({"data": data, "sheets": sheets})
#     except Exception as e:
#         logging.error(f"Error in view_data: {e}")
#         return jsonify({"error": "Failed to read data. See server logs for details."}), 500
@app.route('/view_data', methods=['POST'])
def view_data():
    file = request.files.get('file')
    if not file:
        return jsonify({"error": "File is required."}), 400

    try:
        excel_file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(excel_file_path)

        # Read the Excel file using pandas
        excel_data = pd.ExcelFile(excel_file_path)
        data = {}
        for sheet_name in excel_data.sheet_names:
            data[sheet_name] = excel_data.parse(sheet_name).to_dict(orient='records')

        logging.info(f"Data retrieved from {file.filename}: {data}")  # Log the data for debugging
        return jsonify({"data": data, "sheets": excel_data.sheet_names}), 200
    except Exception as e:
        logging.error(f"Error in view_data: {e}")
        return jsonify({"error": "Failed to read data. See server logs for details."}), 500

    
@app.route('/existing_macros', methods=['POST'])
def existing_macros():
    file = request.files.get('file')
    if not file:
        return jsonify({"error": "File is required."}), 400

    try:
        excel_file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(excel_file_path)

        # Dynamically import the ConvertedExcelMacros class
        spec = importlib.util.spec_from_file_location("ConvertedExcelMacros", "converted_macros.py")
        converted_macros_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(converted_macros_module)

        # Create an instance of the ConvertedExcelMacros class
        macros_instance = converted_macros_module.ConvertedExcelMacros()
        macros = macros_instance.existing_macros(excel_file_path)
        return jsonify({"macros": macros})
    except Exception as e:
        logging.error(f"Error in existing_macros: {e}")
        return jsonify({"error": "Failed to get existing macros. See server logs for details."}), 500

@app.route('/execute_macro', methods=['POST'])
def execute_macro():
    file = request.files.get('file')
    macro_name = request.form.get('macro_name')
    if not file or not macro_name:
        return jsonify({"error": "File and macro name are required."}), 400

    try:
        excel_file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(excel_file_path)

        # Dynamically import the ConvertedExcelMacros class
        spec = importlib.util.spec_from_file_location("ConvertedExcelMacros", "converted_macros.py")
        converted_macros_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(converted_macros_module)

        # Create an instance of the ConvertedExcelMacros class
        macros_instance = converted_macros_module.ConvertedExcelMacros()
        result = macros_instance.execute_macro(excel_file_path, macro_name)
        return jsonify(result)
    except Exception as e:
        logging.error(f"Error in execute_macro: {e}")
        return jsonify({"error": "Failed to execute macro. See server logs for details."}), 500

if __name__ == '__main__':
    app.run(debug=True)