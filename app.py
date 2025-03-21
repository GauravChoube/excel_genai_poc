import os
import json
import pandas as pd
from flask import Flask, request, jsonify, render_template
from excelParsing import ExcelVBAProcessor
from openaiClient import OPENAI_CLIENT
import pythoncom
import importlib.util
import logging
import openpyxl
import time
import xlwings as xw

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

# def dataframe_to_excel_with_dataframe_formulas(df, output_file, sheet_name="Sheet1", formula_column="formula"):
#     """
#     Writes a pandas DataFrame to Excel, handling formulas stored within the DataFrame, and keeping the headers.

#     Args:
#         df (pd.DataFrame): The DataFrame to write.
#         output_file (str): The path to the output Excel file.
#         sheet_name (str): The name of the sheet.
#         formula_column (str): The name of the column containing formulas (or None if no such column exists).
#     """
#     try:
#         # Create a copy of the dataframe so as not to modify the original.
#         df_copy = df.copy()

#         # Extract formulas and remove formula column from the copy.
#         formula_cells = {}
#         if formula_column in df_copy.columns:
#             for index, row in df_copy.iterrows():
#                 formula = row[formula_column]
#                 if pd.notna(formula) and isinstance(formula, str) and formula.startswith("="):
#                     col_letter = openpyxl.utils.get_column_letter(df_copy.columns.get_loc(formula_column))
#                     cell_address = f"{openpyxl.utils.get_column_letter(df_copy.columns.get_loc(formula_column)+1)}{index + 2}"
#                     formula_cells[cell_address] = formula
#             df_copy = df_copy.drop(formula_column, axis=1)

#         # Write the DataFrame to Excel (without formula column), keeping headers
#         with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
#             df_copy.to_excel(writer, sheet_name=sheet_name, index=False) #Now write the headers.

#         # Load the workbook and add formulas
#         workbook = openpyxl.load_workbook(output_file)
#         sheet = workbook[sheet_name]
#         for cell_address, formula in formula_cells.items():
#             sheet[cell_address] = formula

#         # Save the workbook
#         workbook.save(output_file)

#         print(f"DataFrame written to '{output_file}-{sheet_name}' with DataFrame formulas and headers.")

#     except Exception as e:
#         print(f"An error occurred: {e}")

def excel_to_excel_with_dataframe_formulas(excel_data, output_file, formula_column="formula"):
    """
    Reads an Excel file with multiple sheets, handles formulas stored in DataFrames, and writes to a new Excel file.

    Args:
        input_file (str): Path to the input Excel file.
        output_file (str): Path to the output Excel file.
        formula_column (str): The name of the column containing formulas (or None if no such column exists).
    """
    try:
        # # Read all sheets from the input Excel file
        # excel_data = pd.read_excel(input_file, sheet_name=None)

        # Create an Excel writer for the output file
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name, df in excel_data.items():
                # Create a copy of the dataframe so as not to modify the original.
                df_copy = df.copy()

                # Extract formulas and remove formula column from the copy.
                formula_cells = {}
                if formula_column in df_copy.columns:
                    for index, row in df_copy.iterrows():
                        formula = row[formula_column]
                        if pd.notna(formula) and isinstance(formula, str) and formula.startswith("="):
                            cell_address = f"{openpyxl.utils.get_column_letter(df_copy.columns.get_loc(formula_column)+1)}{index + 2}"
                            formula_cells[cell_address] = formula
                    df_copy = df_copy.drop(formula_column, axis=1)

                # Write the DataFrame to the output sheet
                df_copy.to_excel(writer, sheet_name=sheet_name, index=False)

                # Load the sheet and add formulas
                workbook = writer.book
                sheet = workbook[sheet_name]
                for cell_address, formula in formula_cells.items():
                    sheet[cell_address] = formula
        
         # Force recalculation and save using xlwings
        app = xw.App(visible=False)  # Run Excel in the background
        wb = app.books.open(output_file)
        app.calculate()  # Force recalculation
        wb.save()
        wb.close()
        app.quit()

        print(f"Excel data  copied to '{output_file}' with DataFrame formulas.")

    except Exception as e:
        print(f"An error occurred: {e}")



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
            data_frame = processor.read_excel_with_formulas_all_sheets()  # Store DataFrame
            # data_frame = processor.read_excel_data()  # Store DataFrame

            logging.info(f"DataFrame created: {data_frame}")  # Log the DataFrame
            logging.info(f"Type of data_frame: {type(data_frame)}") 
            if isinstance(data_frame, dict):
                for key, df in data_frame.items():
                    logging.info(f"DataFrame '{key}' shape: {df.shape}")  # Log the shape of each DataFrame
            else:
                logging.error("data_frame is not a dictionary.")

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
    tmpExcelFile = "tmpExcel.xlsx"
    '''
        Execute the all formula with lastest data.
        Following the this approach 
            1. Create excel file with current datafram
            2. read with open
    '''  
    excel_to_excel_with_dataframe_formulas(excel_data=data_frame,output_file=tmpExcelFile)

    updateDf = pd.read_excel(tmpExcelFile,sheet_name=None)
    if updateDf is None:
        return jsonify({"error": "Formula execution is failed. Data not updated"}), 400
    
    if os.path.exists(tmpExcelFile):
        os.remove(tmpExcelFile)


    data = {}
    for sheet_name in updateDf.keys():
        data[sheet_name] = updateDf[sheet_name].fillna("N/A").to_dict(orient='records')

        logging.info(f"{sheet_name}:\n{updateDf[sheet_name]}\n after N/A")
        logging.info(f"{sheet_name}:\n{data[sheet_name]}")
    
    logging.info(f"sheet: {updateDf.keys()} data:{data}")
    return jsonify({"sheets": list(updateDf.keys()), "data": data}), 200


@app.route('/update_row', methods=['POST'])
def update_row():
    global data_frame
    if data_frame is None:
        return jsonify({"error": "No data available. Please upload a file first."}), 400

    sheet_name = request.json.get('sheet_name')
    row_index = request.json.get('row_index')
    updated_row = request.json.get('updated_row')

    try:
        # Update the specified row with new data
        for column, value in updated_row.items():
            # Ensure the value matches the existing data type
            if column in data_frame[sheet_name].columns:
                existing_dtype = data_frame[sheet_name][column].dtype
                updated_value = pd.Series(value).astype(existing_dtype).iloc[0]  # Convert to existing dtype
                data_frame[sheet_name].at[row_index, column] = updated_value  # Update the DataFrame
        logging.info(f"Row updated in '{sheet_name}': {updated_row}")
        return jsonify({"message": "Row updated successfully"}), 200
    except Exception as e:
        logging.error(f"Error updating row: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/add_row', methods=['POST'])
def add_row():
    global data_frame
    if data_frame is None:
        return jsonify({"error": "No data available. Please upload a file first."}), 400

    sheet_name = request.json.get('sheet_name')
    new_row = request.json.get('newRow')

    try:
        # Create a DataFrame from the new row
        new_row_df = pd.DataFrame([new_row])

        # Ensure the new row matches the data types of the existing DataFrame
        for column in data_frame[sheet_name].columns:
            if column in new_row_df.columns:
                new_row_df[column] = new_row_df[column].astype(data_frame[sheet_name][column].dtype)

        # Concatenate the new row DataFrame with the existing DataFrame
        data_frame[sheet_name] = pd.concat([data_frame[sheet_name], new_row_df], ignore_index=True)
        logging.info(f"New row added to '{sheet_name}': {new_row}")
        return jsonify({"message": "Row added successfully"}), 200
    except Exception as e:
        logging.error(f"Error adding row: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/delete_row', methods=['POST'])
def delete_row():
    global data_frame
    if data_frame is None:
        return jsonify({"error": "No data available. Please upload a file first."}), 400

    sheet_name = request.json.get('sheet_name')
    row_index = request.json.get('row_index')

    try:
        if row_index is not None and 0 <= row_index < len(data_frame[sheet_name]):
            data_frame[sheet_name].drop(index=row_index, inplace=True)
            data_frame[sheet_name].reset_index(drop=True, inplace=True)
            logging.info(f"Row {row_index} deleted from '{sheet_name}'")
            return jsonify({"message": "Row deleted successfully"}), 200
        else:
            return jsonify({"error": "Invalid row index"}), 400
    except Exception as e:
        logging.error(f"Error deleting row: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/add_column', methods=['POST'])
def add_column():
    global data_frame
    if data_frame is None:
        return jsonify({"error": "No data available. Please upload a file first."}), 400

    sheet_name = request.json.get('sheet_name')
    column_name = request.json.get('column_name')

    try:
        data_frame[sheet_name][column_name] = None
        logging.info(f"Column '{column_name}' added to '{sheet_name}'")
        return jsonify({"message": "Column added successfully"}), 200
    except Exception as e:
        logging.error(f"Error adding column: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/delete_column', methods=['POST'])
def delete_column():
    global data_frame
    if data_frame is None:
        return jsonify({"error": "No data available. Please upload a file first."}), 400

    sheet_name = request.json.get('sheet_name')
    column_name = request.json.get('column_name')

    try:
        if column_name in data_frame[sheet_name].columns:
            data_frame[sheet_name].drop(columns=[column_name], inplace=True)
            logging.info(f"Column '{column_name}' deleted from '{sheet_name}'")
            return jsonify({"message": "Column deleted successfully"}), 200
        else:
            return jsonify({"error": "Column not found"}), 400
    except Exception as e:
        logging.error(f"Error deleting column: {e}")
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
        logging.info(f"Type of data_frame before passing: {type(data_frame)}")  # Log the type
        if isinstance(data_frame, dict):
            for key, df in data_frame.items():
                logging.info(f"DataFrame '{key}' shape: {df.shape}")  # Log the shape of each DataFrame
        else:
            logging.error("data_frame is not a dictionary.")
        # Dynamically import the ConvertedExcelMacros class
        converted_macros = importlib.import_module('converted_macros')
        macro_class = converted_macros.ConvertedExcelMacros(data_frame)  # Use the latest DataFrame
        result = macro_class.execute_macro(macro_name)  # Execute the macro
        return jsonify(result), 200
    except Exception as e:
        # logging.error(f"Error executing macro '{macro_name}': {e}")
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    app.run(debug=False)