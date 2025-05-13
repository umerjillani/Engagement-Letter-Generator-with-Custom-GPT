# Document Generator API

## 1. Purpose & Overview

This system is designed to automate the creation of client-specific engagement letters using a structured template and user-provided data. It streamlines the generation of professional DOCX documents for financial advisory.

### Key Components

| Component | Description |
|-----------|-------------|
| server.py | Hosts the Flask-based REST API. Manages HTTP endpoints for document generation and download, as well as template initialization. Also handles expiration-based cleanup of temporary files. |
| document_generator.py | Contains the logic to dynamically populate a Word (DOCX) document using user inputs and a JSON template structure. |
| structured_document.json | A template definition file describing the layout, placeholders, and content structure for the engagement letter. |
| temp/ | Temporary directory used to store generated DOCX files. Each document is retained for 24 hours after generation and cleaned up via a background process. |

### Process Flow

1. **User Input Submission**
   - A user (such as a Custom GPT or frontend app) sends a POST request to `/api/generate-document` with user-specific data in JSON format.

2. **Document Generation**
   - The API saves the input, invokes `create_engagement_letter()` to generate a DOCX file using the JSON template, and deletes the input file after use.

3. **Download Endpoint**
   - The client receives a download_url and can retrieve the document via a GET request to `/api/download/<file_id>` within 24 hours.

## 2. Environment Preparation

### System Requirements
- Python 3.8 or higher
- pip (Python package manager)
- ngrok (for exposing local server to the internet)
- Open API(Chat GPT) premium subscription for creating the Custom GPT

### Installation

1. Clone the repository:
```bash
git clone <your-repository-url>
cd <repository-directory>
```

2. Create a virtual environment:
```bash
python -m venv venv
```

3. Activate the virtual environment:
- Windows:
```bash
venv\Scripts\activate
```
- Linux/Mac:
```bash
source venv/bin/activate
```

4. Install required packages:
```bash
pip install -r requirements.txt
```

## 3. Project Structure

```
project_root/
├── document_generator.py    # Main document generation logic
├── server.py               # Flask API server
├── templates/              # Template files
│   └── FA.json            # Financial Advisor template
├── temp/                   # Temporary files directory
├── requirements.txt        # Python dependencies
└── README.md              # This file
```

## 4. Running the Flask Server

### 4.1. Initial Startup
```bash
python server.py
```
The server runs on: http://127.0.0.1:5000

A structured template file (structured_document.json) should be placed under templates/ on first launch if it does not exist.

### 4.2. API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| /api/generate-document | POST | Accepts user data, generates DOCX, returns link |
| /api/download/<file_id> | GET | Downloads the generated document |
| /api/setup-template | POST | Replaces or sets up the structured template |

## 5. Ngrok Setup and Exposure

1. Download ngrok:
   - Visit https://ngrok.com/download
   - Sign up for a free account
   - Download the appropriate version for your OS

2. Install ngrok:
   - Windows: Extract the downloaded zip file
   - Linux/Mac: 
     ```bash
     unzip /path/to/ngrok.zip
     ```

3. Connect your ngrok account:
   - Get your authtoken from ngrok dashboard
   - In the dashboard, go to the "Getting Started" or "Setup & Installation" section
   - You'll see a command like this:
     ```bash
     ngrok config add-authtoken 2P7exampleTokenX5yHqDkEXAMPLE
     ```
   - Copy the full command or just the token part
   - In your terminal or command prompt, paste the above command
   - This saves the token in your local ngrok config file (usually at ~/.ngrok2/ngrok.yml)

4. Run ngrok:
```bash
ngrok http 5000
```
This will generate a public HTTPS URL like:
```
https://xxxx-xx-xx-xxx-xx.ngrok-free.app
```

## 6. Creating and Configuring a Custom GPT

### Visit the GPT Creation Portal
1. Go to: https://chat.openai.com/gpts
2. Click "Create a GPT" in the top right corner

### Name and Identity
- Name: Engagement Letter Generator
- Description: Customised professional engagement letters Generator by collecting and enhancing user-provided information. Delivers a downloadable document via secure API integration.

### Instructions (System Message)
[Detailed instructions for the Custom GPT configuration are provided in the system message section]

### Conversation Starters
- "Create an engagement letter for my financial advisory"

### Custom GPT's Knowledge Base
- Upload structured_document.json as reference data in the GPT's knowledge panel
- This JSON file will serve as a reference-only static file

### Custom GPT Actions – Configuration Details

#### Authentication
- Set the Auth Type to: None
- Ensure the ngrok URL is publicly accessible

#### OpenAPI Schema
1. Copy the contents of openapi.yaml
2. Paste it into the "Paste OpenAPI Schema" input box in ChatGPT
3. Update the Base URL in the servers section with your active ngrok HTTPS URL:
```yaml
servers:
  - url: https://your-ngrok-url.ngrok-free.app
```

### Final Check
1. Test the GPT by providing engagement letter inputs
2. Verify that it successfully returns a downloadable link

## 7. Security Considerations

1. The application uses temporary file storage with automatic cleanup
2. Files expire after 24 hours
3. Implement proper authentication in production
4. Use HTTPS in production
5. Consider rate limiting for API endpoints

## 8. Troubleshooting

1. If ngrok connection fails:
   - Check if ngrok is running
   - Verify your authtoken
   - Ensure port 5000 is not blocked

2. If document generation fails:
   - Check template file format
   - Verify user data structure
   - Check file permissions in temp directory

## 9. Production Deployment

For production deployment:
1. Use a production-grade WSGI server (e.g., Gunicorn)
2. Set up proper logging
3. Implement authentication
4. Use environment variables for configuration
5. Set up proper error handling
6. Use a production-grade database for file management

## 10. License

[Your License Information]

## 11. Support

[Your Support Information] 