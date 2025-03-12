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
# @app.route('/view_data', methods=['POST'])
# def view_data():
#     # Get the filename from the request
#     filename = request.form.get('filename')
#     if not filename:
#         return jsonify({"error": "Filename is required."}), 400

#     try:
#         excel_file_path = os.path.join(UPLOAD_FOLDER, filename)

#         # Read the Excel file using pandas
#         excel_data = pd.ExcelFile(excel_file_path)
#         data = {}
#         for sheet_name in excel_data.sheet_names:
#             data[sheet_name] = excel_data.parse(sheet_name).to_dict(orient='records')

#         logging.info(f"Data retrieved from {filename}: {data}")  # Log the data for debugging
#         return jsonify({"data": data, "sheets": excel_data.sheet_names}), 200
#     except Exception as e:
#         logging.error(f"Error in view_data: {e}")
#         return jsonify({"error": "Failed to read data. See server logs for details."}), 500
@app.route('/view_data', methods=['POST'])
def view_data():
    filename = request.form.get('filename')
    file_path = os.path.join(UPLOAD_FOLDER, filename)
    try:
        excel_data = pd.ExcelFile(file_path)
        data = {}
        for sheet_name in excel_data.sheet_names:
            data[sheet_name] = excel_data.parse(sheet_name).to_dict(orient='records')
        return jsonify({"sheets": excel_data.sheet_names, "data": data}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/add_row', methods=['POST'])
def add_row():
    filename = request.form['filename']
    new_row = json.loads(request.form['newRow'])  # Ensure this is a valid JSON string
    file_path = os.path.join(UPLOAD_FOLDER, filename)

    try:
        # Load the existing data
        excel_data = pd.ExcelFile(file_path)
        sheet_name = excel_data.sheet_names[0]  # Assuming you want to add to the first sheet
        df = excel_data.parse(sheet_name)

        # Create a DataFrame for the new row
        new_row_df = pd.DataFrame([new_row])  # Create a DataFrame from the new row dictionary

        # Concatenate the new row DataFrame with the existing DataFrame
        df = pd.concat([df, new_row_df], ignore_index=True)

        # Save the updated DataFrame back to the Excel file
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        return jsonify({"message": "Row added successfully"}), 200
    except Exception as e:
        print(f"Error adding row: {e}")  # Log the error for debugging
        return jsonify({"error": str(e)}), 500

@app.route('/delete_row', methods=['POST'])
def delete_row():
    filename = request.form['filename']
    row_index = int(request.form['rowIndex'])
    file_path = os.path.join(UPLOAD_FOLDER, filename)

    try:
        # Load the existing data
        excel_data = pd.ExcelFile(file_path)
        sheet_name = excel_data.sheet_names[0]  # Assuming you want to delete from the first sheet
        df = excel_data.parse(sheet_name)

        # Drop the specified row
        df = df.drop(index=row_index)

        # Save the updated DataFrame back to the Excel file
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        return jsonify({"message": "Row deleted successfully"}), 200
    except Exception as e:
        print(f"Error deleting row: {e}")  # Log the error for debugging
        return jsonify({"error": str(e)}), 500

@app.route('/update_row', methods=['POST'])
def update_row():
    filename = request.form['filename']
    row_index = int(request.form['rowIndex'])
    updated_row = json.loads(request.form['updatedRow'])
    file_path = os.path.join(UPLOAD_FOLDER, filename)

    try:
        excel_data = pd.ExcelFile(file_path)
        sheet_name = excel_data.sheet_names[0]
        df = excel_data.parse(sheet_name)

        for key, value in updated_row.items():
            df.at[row_index, key] = value

        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        return jsonify({"message": "Row updated successfully"}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500


    
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