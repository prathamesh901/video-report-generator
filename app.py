from flask import Flask, send_file, render_template, jsonify
from pymongo import MongoClient
from mongo_pdf_finall import create_document
import os

app = Flask(__name__, static_folder="static", template_folder="templates")

# MongoDB connection
client = MongoClient("mongodb+srv://prathameshhh902:Bvit%402002@cluster0.q2fnv.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0")
db = client['videoDB']
collection = db['video_processing']

# Function to convert MongoDB documents to JSON serializable format
def convert_to_json(doc):
    doc["_id"] = str(doc["_id"])  # Convert ObjectId to string
    return doc

# Serve the main page
@app.route("/")
def home():
    return render_template("index.html")

# Fetch all videos from MongoDB for selection dropdown
@app.route("/get_videos", methods=["GET"])
def get_videos():
    videos = collection.find({}, {"video_hash": 1, "original_filename": 1})  # Fetch video hash & name
    return jsonify([convert_to_json(video) for video in videos])

# Download Word or PDF for a specific video
@app.route("/download/<video_hash>/<file_type>", methods=["GET"])
def download_report(video_hash, file_type):
    """Fetch JSON metadata for the given video_hash and generate a report."""
    video_data = collection.find_one({"video_hash": video_hash})

    if not video_data:
        return jsonify({"error": "Video not found"}), 404

    # Generate Word & PDF documents
    word_blob, pdf_blob = create_document(video_data)

    # Define file paths
    word_path = f"Video_Report_{video_hash}.docx"
    pdf_path = f"Video_Report_{video_hash}.pdf"

    # Save files temporarily
    with open(word_path, "wb") as f:
        f.write(word_blob)
    with open(pdf_path, "wb") as f:
        f.write(pdf_blob)

    try:
        # Return requested file
        if file_type == "word":
            return send_file(word_path, as_attachment=True)
        elif file_type == "pdf":
            return send_file(pdf_path, as_attachment=True)
        else:
            return jsonify({"error": "Invalid file type"}), 400
    finally:
        # Clean up files after sending
        if os.path.exists(word_path):
            os.remove(word_path)
        if os.path.exists(pdf_path):
            os.remove(pdf_path)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
