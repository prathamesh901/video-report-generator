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

@app.route("/download", methods=["GET"])
def download_report():
    """Fetch the latest video metadata and generate a report."""
    video_data = collection.find_one(sort=[("upload_timestamp", -1)])

    if not video_data:
        return {"error": "No video data found"}, 404

    # Generate Word & PDF documents
    word_blob, pdf_blob = create_document(video_data)

    # Save files
    word_path = "Video_Report.docx"
    pdf_path = "Video_Report.pdf"
    with open(word_path, "wb") as f:
        f.write(word_blob)
    with open(pdf_path, "wb") as f:
        f.write(pdf_blob)

    # Zip both files
    zip_path = "Video_Report.zip"
    os.system(f"zip -r {zip_path} {word_path} {pdf_path}")

    return send_file(zip_path, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
