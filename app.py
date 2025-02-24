import os
import json
import pythoncom
import pandas as pd  # Import pandas for reading Excel files
from flask import Flask, request, jsonify, render_template, send_from_directory, session
from flask_socketio import SocketIO, emit
from excelParsing import ExcelVBAProcessor
from openaiClient import OPENAI_CLIENT
import subprocess

app = Flask(__name__)
socketio = SocketIO(app)

# Configuration
CONFIG_PATH = "config.json"
UPLOAD_FOLDER = os.path.abspath("uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
OUTPUT_FOLDER = os.path.abspath(".")



app.secret_key = 'your_secret_key' 

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
        print(f"❌ Config Error: {e}")
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
    socketio.emit('log', {'message': "🔄 Starting file upload..."})
    if 'file' not in request.files:
        socketio.emit('log', {'message': "❌ No file part in the request."})
        return jsonify({"error": "No file part"}), 400

    file = request.files['file']
    if file.filename == '':
        socketio.emit('log', {'message': "❌ No selected file."})
        return jsonify({"error": "No selected file"}), 400

    file_path = os.path.abspath(os.path.join(UPLOAD_FOLDER, file.filename))
    file.save(file_path)

    if not os.path.exists(file_path):
        socketio.emit('log', {'message': f"❌ File {file.filename} was not saved at {file_path}."})
        return jsonify({"error": f"File {file.filename} was not saved at {file_path}"}), 500

    socketio.emit('log', {'message': f"✅ File saved successfully at: {file_path}"})

    # Store the uploaded filename in the session
    session['uploaded_file'] = file.filename

    # Process the Excel file
    try:
        pythoncom.CoInitialize()
        socketio.emit('log', {'message': "🔄 Processing the Excel file..."})

        if not openaiClient:
            socketio.emit('log', {'message': "❌ OpenAI Client not initialized."})
            return jsonify({"error": "OpenAI Client not initialized"}), 500

        processor = ExcelVBAProcessor(file_path=file_path, openAIClient=openaiClient)
        result = processor.process_excel_file()
        socketio.emit('log', {'message': "✅ File processed successfully."})

        pythoncom.CoUninitialize()
        
        return jsonify({"success": True, "redirect": "/result"}) 
    except Exception as e:
        pythoncom.CoUninitialize()
        socketio.emit('log', {'message': f"❌ Error during processing: {str(e)}"})
        return jsonify({"error": str(e)}), 500

# Result Route
@app.route('/result')
def result():
    return render_template('result.html')

# Execute Route
@app.route('/execute', methods=['POST'])
def execute_file():
    print("Execute route hit")  # Debugging line
    file_path = os.path.join(OUTPUT_FOLDER, "converted_macros.py")
   
    
    if not os.path.exists(file_path):
        print("File not found:", file_path)  # Debugging line
        return jsonify({"error": "File not found"}), 404

    # Execute the Python file
    try:
        print("Executing file at:", file_path)  # Debugging line
        result = subprocess.run(['python', file_path], capture_output=True, text=True)
        output = result.stdout.strip()  # Get the standard output
        error = result.stderr.strip()    # Get the standard error

        # Log output and error
        print("Output:", output)
        print("Error:", error)

        # Return both output and error
        return jsonify({"output": output, "error": error}), 200
    except Exception as e:
        print("Execution error:", str(e))  # Debugging line
        return jsonify({"error": str(e)}), 500

# Visualization Route
@app.route('/visualize', methods=['GET'])
def visualize_data():
    if 'uploaded_file' not in session:
        return jsonify({"error": "No uploaded file found"}), 404

    file_name = session['uploaded_file']
    file_path = os.path.join(UPLOAD_FOLDER, file_name)

    if not os.path.exists(file_path):
        return jsonify({"error": "File not found"}), 404

    # Read the Excel file
    try:
        xls = pd.ExcelFile(file_path)
        sheets = {sheet_name: xls.parse(sheet_name).to_html(classes='data') for sheet_name in xls.sheet_names}
        return render_template('visualize.html', sheets=sheets)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# Download Route
@app.route('/download', methods=['GET'])
def download_file():
    file_name = "converted_macros.py"
    file_path = os.path.join(OUTPUT_FOLDER, file_name)

    if not os.path.exists(file_path):
        return jsonify({"error": "File not found"}), 404

    return send_from_directory(OUTPUT_FOLDER, file_name, as_attachment=True)

if __name__ == '__main__':
    socketio.run(app, debug=True)

