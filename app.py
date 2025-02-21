from flask import Flask, send_file, render_template
from pymongo import MongoClient
from mongo_pdf_finall import create_document
import os

app = Flask(__name__, static_folder="static", template_folder="templates")

# MongoDB connection
client = MongoClient("mongodb+srv://prathameshhh902:Bvit%402002@cluster0.q2fnv.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0")
db = client['videoDB']
collection = db['video_processing']

@app.route("/")
def home():
    """Serve the main page with the download button."""
    return render_template("index.html")

def generate_report():
    """Fetch the latest video metadata and generate Word & PDF reports."""
    video_data = collection.find_one(sort=[("upload_timestamp", -1)])

    if not video_data:
        return None, None

    word_blob, pdf_blob = create_document(video_data)

    word_path = "Video_Report.docx"
    pdf_path = "Video_Report.pdf"

    with open(word_path, "wb") as f:
        f.write(word_blob)
    with open(pdf_path, "wb") as f:
        f.write(pdf_blob)

    return word_path, pdf_path

@app.route("/download_word", methods=["GET"])
def download_word():
    """Serve the Word document."""
    word_path, _ = generate_report()
    if not word_path:
        return {"error": "No video data found"}, 404
    return send_file(word_path, as_attachment=True)

@app.route("/download_pdf", methods=["GET"])
def download_pdf():
    """Serve the PDF document."""
    _, pdf_path = generate_report()
    if not pdf_path:
        return {"error": "No video data found"}, 404
    return send_file(pdf_path, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
