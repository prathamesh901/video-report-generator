from flask import Flask, jsonify, send_file
from pymongo import MongoClient
import os
import zipfile
from io import BytesIO

app = Flask(__name__)

# Connect to MongoDB
client = MongoClient("mongodb+srv://prathameshhh902:Bvit%402002@cluster0.q2fnv.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0")
db = client["videoDB"]  # Make sure your database name is correct
collection = db["video_processing"]  # Collection name

@app.route('/')
def home():
    return """
        <h2>✅ Video Report Generator</h2>
        <button onclick="window.location.href='/download-latest'">Download Latest Report</button>
    """

@app.route('/download-latest')
def download_latest():
    """Finds the latest video hash and downloads the report"""
    latest_video = collection.find_one({}, sort=[("_id", -1)])  # Get the most recent video
    if not latest_video:
        return jsonify({"error": "No video metadata found"}), 404

    video_hash = latest_video["video_hash"]
    return download_report(video_hash)

@app.route('/download/<video_hash>')
def download_report(video_hash):
    """Fetches and downloads the report for a given video hash"""
    video_data = collection.find_one({"video_hash": video_hash})
    if not video_data:
        return jsonify({"error": "No metadata found"}), 404

    # Generate the reports (Replace this with your actual function)
    word_blob = generate_word_report(video_data)
    pdf_blob = generate_pdf_report(video_data)

    # Create a ZIP file in memory
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr("Video_Report.docx", word_blob)
        zip_file.writestr("Video_Report.pdf", pdf_blob)

    zip_buffer.seek(0)
    return send_file(zip_buffer, mimetype="application/zip", as_attachment=True, download_name="Video_Report.zip")

def generate_word_report(data):
    """Simulated function to generate a Word file (Replace with actual logic)"""
    return b"Dummy Word content"

def generate_pdf_report(data):
    """Simulated function to generate a PDF file (Replace with actual logic)"""
    return b"Dummy PDF content"



if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))  # Use PORT from Render, default to 5000
    app.run(host="0.0.0.0", port=port, debug=True)

