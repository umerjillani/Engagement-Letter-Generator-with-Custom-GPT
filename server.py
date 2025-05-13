from flask import Flask, request, jsonify, send_file
from werkzeug.utils import secure_filename
import os
import json
import uuid as uuid_lib
from datetime import datetime, timedelta
import threading
import time
import shutil

# Import the document generation function
from document_generator import create_engagement_letter

app = Flask(__name__)
app.config['JSON_SORT_KEYS'] = False

# Create temp directory for files
TEMP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp')
os.makedirs(TEMP_DIR, exist_ok=True)

# Create templates directory
TEMPLATES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'templates')
os.makedirs(TEMPLATES_DIR, exist_ok=True)

# Dictionary to store temporary file info with expiration
temp_files = {}

# Template JSON path
TEMPLATE_PATH = os.path.join(TEMPLATES_DIR, 'structured_document.json')

# Ensure template exists by copying the provided template if needed
if not os.path.exists(TEMPLATE_PATH):
    # Create a template file with basic structure if none exists
    # In production, you would want to have this template pre-configured
    with open(TEMPLATE_PATH, 'w') as f:
        json.dump({
            "header": {
                "company_details": {
                    "phone": "(480) 947-3321",
                    "address": "6750 E. Camelback Road, Suite 103",
                    "city_state_zip": "Scottsdale, AZ 85251"
                }
            },
            # Add basic template structure
            "sections": {
                "introduction": {
                    "description": "Sample introduction for engagement letter."
                }
            }
        }, f, indent=4)

@app.route('/api/generate-document', methods=['POST'])
def generate_document():
    try:
        # Get user data from request
        user_data = request.json
        if not user_data:
            return jsonify({"error": "No user data provided"}), 400
        
        # Generate unique file ID
        file_id = str(uuid_lib.uuid4())
        
        # Create file paths
        user_data_path = os.path.join(TEMP_DIR, f'userdata_{file_id}.json')
        output_path = os.path.join(TEMP_DIR, f'engagement_letter_{file_id}.docx')
        
        # Write user data to file
        with open(user_data_path, 'w') as f:
            json.dump(user_data, f, indent=4)
        
        # Generate document
        create_engagement_letter(TEMPLATE_PATH, user_data_path, output_path)
        
        # Clean up user data file
        if os.path.exists(user_data_path):
            os.remove(user_data_path)
        
        # Set expiration (1 hour from now)
        expiration_time = datetime.now() + timedelta(hours=24)
        
        # Store file info
        temp_files[file_id] = {
            "path": output_path,
            "expires_at": expiration_time
        }
        
        # Generate download URL
        host_url = request.host_url.rstrip('/')
        download_url = f"{host_url}/api/download/{file_id}"
        
        return jsonify({
            "message": "Document generated successfully",
            "file_id": file_id,
            "download_url": download_url,
            "expires_at": expiration_time.isoformat()
        }), 200
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/download/<file_id>', methods=['GET'])
def download_file(file_id):
    # Check if file exists and has not expired
    if file_id not in temp_files:
        return jsonify({"error": "File not found or expired"}), 404
    
    file_info = temp_files[file_id]
    file_path = file_info["path"]
    
    if not os.path.exists(file_path):
        # Remove from temp_files if file doesn't exist
        del temp_files[file_id]
        return jsonify({"error": "File not found"}), 404
    
    # Check if file has expired
    if datetime.now() > file_info["expires_at"]:
        # Clean up expired file
        if os.path.exists(file_path):
            os.remove(file_path)
        del temp_files[file_id]
        return jsonify({"error": "File has expired"}), 404
    
    return send_file(
        file_path,
        as_attachment=True,
        download_name="engagement_letter.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# Initialize template file from provided JSON if needed
@app.route('/api/setup-template', methods=['POST'])
def setup_template():
    try:
        # Allow template to be provided directly in the request body
        template_data = request.json
        if not template_data:
            return jsonify({"error": "No template data provided"}), 400
        
        # Write template to file
        with open(TEMPLATE_PATH, 'w') as f:
            json.dump(template_data, f, indent=4)
        
        return jsonify({
            "message": "Template set up successfully",
            "template_path": TEMPLATE_PATH
        }), 200
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# Function to periodically clean up expired files
def cleanup_expired_files():
    while True:
        current_time = datetime.now()
        
        # Make a copy of keys to avoid dictionary changed size during iteration error
        file_ids = list(temp_files.keys())
        
        for file_id in file_ids:
            if file_id in temp_files and current_time > temp_files[file_id]["expires_at"]:
                file_path = temp_files[file_id]["path"]
                if os.path.exists(file_path):
                    try:
                        os.remove(file_path)
                        print(f"Deleted expired file: {file_path}")
                    except Exception as e:
                        print(f"Error deleting file {file_path}: {e}")
                
                # Remove from temp_files dictionary
                del temp_files[file_id]
        
        # Sleep for 15 minutes before next cleanup
        time.sleep(15 * 60)

if __name__ == '__main__':
    # Start cleanup thread
    cleanup_thread = threading.Thread(target=cleanup_expired_files, daemon=True)
    cleanup_thread.start()
    
    # Start Flask server
    app.run(host='0.0.0.0', port=5000, debug=True)