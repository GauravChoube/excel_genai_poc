import os
import json
import pandas as pd
from flask import Flask, request, jsonify, render_template
from excelParsing import ExcelVBAProcessor
from openaiClient import OPENAI_CLIENT
import pythoncom
import importlib.util
import logging

# Configuration
CONFIG_PATH = "config.json"
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Set up logging
logging.basicConfig(level=logging.INFO)

app = Flask(__name__)

# Global variables to store DataFrame and VBA macros
data_frame = None
vba_macros = None

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
    global data_frame, vba_macros  # Declare global variables
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
            data_frame = processor.read_excel_data()  # Store DataFrame
            logging.info(f"DataFrame created: {data_frame}")  # Log the DataFrame
            vba_macros = processor.extract_vba_macros()  # Store extracted macros
            processor.convert_vba_to_python()
            processor.save_python_class()  # Save the converted macros
            pythoncom.CoUninitialize()
            return jsonify({"message": "File uploaded and processed successfully", "file_path": file_path}), 200
        except Exception as e:
            pythoncom.CoUninitialize()
            logging.error(f"Error processing file: {e}")  # Log the error for debugging
            return jsonify({"error": str(e)}), 500

@app.route('/view_data', methods=['POST'])
def view_data():
    global data_frame  # Declare global variable
    if data_frame is None:
        return jsonify({"error": "No data available. Please upload a file first."}), 400

    data = {}
    for sheet_name in data_frame.keys():
        data[sheet_name] = data_frame[sheet_name].fillna("N/A").to_dict(orient='records')
    
    return jsonify({"sheets": list(data_frame.keys()), "data": data}), 200

@app.route('/add_row', methods=['POST'])
def add_row():
    global data_frame  # Declare global variable
    if data_frame is None:
        return jsonify({"error": "No data available. Please upload a file first."}), 400

    new_row = json.loads(request.form['newRow'])  # Ensure this is a valid JSON string
    sheet_name = list(data_frame.keys())[0]  # Assuming you want to add to the first sheet

    try:
        new_row_df = pd.DataFrame([new_row])  # Create a DataFrame from the new row
        data_frame[sheet_name] = pd.concat([data_frame[sheet_name], new_row_df], ignore_index=True)  # Concatenate the new row DataFrame with the existing DataFrame
        logging.info(f"New row added to '{sheet_name}': {new_row}")  # Log the new row added
        return jsonify({"message": "Row added successfully"}), 200
    except Exception as e:
        logging.error(f"Error adding row: {e}")  
        return jsonify({"error": str(e)}), 500


@app.route('/update_row', methods=['POST'])
def update_row():
    global data_frame  # Declare global variable
    if data_frame is None:
        return jsonify({"error": "No data available. Please upload a file first."}), 400

    sheet_name = request.json.get('sheet_name')
    row_index = request.json.get('row_index')
    updated_row = request.json.get('updated_row')

    try:
        # Update the specified row with new data
        for column, value in updated_row.items():
            data_frame[sheet_name].at[row_index, column] = value  # Update the DataFrame
        logging.info(f"Row updated in '{sheet_name}': {updated_row}")  # Log the updated row
        return jsonify({"message": "Row updated successfully"}), 200
    except Exception as e:
        logging.error(f"Error updating row: {e}")  
        return jsonify({"error": str(e)}), 500

@app.route('/delete_row', methods=['POST'])
def delete_row():
    global data_frame  # Declare global variable
    if data_frame is None:
        return jsonify({"error": "No data available. Please upload a file first."}), 400

    row_index = request.json.get('row_index')
    sheet_name = list(data_frame.keys())[0]  # Assuming you want to delete from the first sheet

    try:
        # Check if the row index is valid
        if row_index is not None and 0 <= row_index < len(data_frame[sheet_name]):
            data_frame[sheet_name].drop(index=row_index, inplace=True)  # Delete the row
            data_frame[sheet_name].reset_index(drop=True, inplace=True)  # Reset index after deletion
            logging.info(f"Row {row_index} deleted from '{sheet_name}'")  # Log the deletion
            return jsonify({"message": "Row deleted successfully"}), 200
        else:
            return jsonify({"error": "Invalid row index"}), 400
    except Exception as e:
        logging.error(f"Error deleting row: {e}")  
        return jsonify({"error": str(e)}), 500

@app.route('/add_column', methods=['POST'])
def add_column():
    global data_frame  # Declare global variable
    if data_frame is None:
        return jsonify({"error": "No data available. Please upload a file first."}), 400

    column_name = request.json.get('column_name')
    sheet_name = list(data_frame.keys())[0]  # Assuming you want to add to the first sheet

    try:
        data_frame[sheet_name][column_name] = None  # Add a new column with None values
        logging.info(f"Column '{column_name}' added to '{sheet_name}'")  # Log the addition
        return jsonify({"message": "Column added successfully"}), 200
    except Exception as e:
        logging.error(f"Error adding column: {e}")  
        return jsonify({"error": str(e)}), 500
    
@app.route('/existing_macros', methods=['POST'])
def existing_macros():
    global vba_macros  # Declare global variable
    if vba_macros is None:
        return jsonify({"error": "No macros available. Please upload a file first."}), 400

    # Collect all macro names from the extracted macros
    macro_names = []
    for module in vba_macros.values():
        macro_names.extend(module["macros"])

    return jsonify({"macros": macro_names}), 200

@app.route('/execute_macro', methods=['POST'])
def execute_macro():
    global data_frame  # Declare global variable
    if data_frame is None:
        return jsonify({"error": "No data available. Please upload a file first."}), 400

    macro_name = request.json.get('macro_name')
    if not macro_name:
        return jsonify({"error": "Macro name is required."}), 400

    try:
        # Dynamically import the ConvertedExcelMacros class
        converted_macros = importlib.import_module('converted_macros')
        macro_class = converted_macros.ConvertedExcelMacros(data_frame)  # Use the latest DataFrame
        result = macro_class.execute_macro(macro_name)  # Execute the macro
        return jsonify(result), 200
    except Exception as e:
        logging.error(f"Error executing macro '{macro_name}': {e}")
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    app.run(debug=False)