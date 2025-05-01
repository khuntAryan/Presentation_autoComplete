from flask import Flask, request, jsonify, send_file
import os
from preprocess import detect_and_rename_placeholders
import tempfile
import uuid

app = Flask(__name__)

TEMPLATES_DIR = os.path.join(os.path.dirname(__file__), 'templates')
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), 'output')

# Ensure directories exist
os.makedirs(TEMPLATES_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

@app.route('/health', methods=['GET'])
def health_check():
    return jsonify({"status": "ok"})

@app.route('/api/preprocess', methods=['POST'])
def preprocess_template():
    # Check if file was uploaded
    if 'file' not in request.files:
        return jsonify({"error": "No file provided"}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No file selected"}), 400
    
    # Save uploaded file temporarily
    temp_filename = f"{uuid.uuid4()}.pptx"
    input_path = os.path.join(TEMPLATES_DIR, temp_filename)
    file.save(input_path)
    
    # Process the file
    output_filename = f"processed_{temp_filename}"
    output_path = os.path.join(OUTPUT_DIR, output_filename)
    
    try:
        success = detect_and_rename_placeholders(input_path, output_path)
        if not success:
            return jsonify({"error": "Failed to process template"}), 500
        
        # Return processed file
        return send_file(output_path, as_attachment=True, 
                         download_name="processed_template.pptx")
    except Exception as e:
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500
    finally:
        # Clean up temporary files
        if os.path.exists(input_path):
            os.remove(input_path)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
