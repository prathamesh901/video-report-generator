from flask import Flask, request, send_file, jsonify
from flask_cors import CORS  # Import CORS
from pymongo import MongoClient
from io import BytesIO
import zipfile
from mongo_pdf_finall import create_document 
import os # Import your function

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

# Connect to MongoDB
client = MongoClient('mongodb://localhost:27017/')
db = client['videoDB']
collection = db['video_processing']

@app.route('/download/<video_hash>', methods=['GET'])
def download_report(video_hash):
    """Fetch video metadata, generate Word & PDF, and return as ZIP."""
    
    # Fetch video metadata from MongoDB
    video_data = collection.find_one({"video_hash": video_hash})
    if not video_data:
        return jsonify({"error": "Video metadata not found"}), 404
    
    try:
        # Generate Word & PDF blobs
        word_blob, pdf_blob = create_document(video_data)
        
        # Create ZIP in memory
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            zip_file.writestr("Video_Report.docx", word_blob)
            zip_file.writestr("Video_Report.pdf", pdf_blob)
        
        zip_buffer.seek(0)

        # Send ZIP file as a response
        return send_file(zip_buffer, 
                         mimetype='application/zip',
                         as_attachment=True,
                         download_name="Video_Report.zip")
    
    except Exception as e:
        print(f"Error: {e}")  # Print error in terminal for debugging
        return jsonify({"error": "Internal Server Error", "details": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))  # Get Render's dynamic port
    app.run(host="0.0.0.0", port=port, debug=True)

