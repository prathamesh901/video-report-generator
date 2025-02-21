import os
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from datetime import datetime
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from pymongo import MongoClient

def create_document(data):
    # === Word Document Creation ===
    doc = Document()
    title = doc.add_heading('Video Analysis Report', 0)
    title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    # === PDF Creation ===
    pdf_buffer = BytesIO()
    pdf_doc = SimpleDocTemplate(pdf_buffer, pagesize=letter)
    styles = getSampleStyleSheet()
    pdf_content = [Paragraph('Video Analysis Report', styles['Title']), Spacer(1, 12)]
    
    def add_section(heading, text):
        """Helper function to add sections to both Word and PDF."""
        doc.add_heading(heading, level=1)
        doc.add_paragraph(text)
        pdf_content.append(Paragraph(heading, styles['Heading1']))
        pdf_content.append(Paragraph(text, styles['Normal']))
        pdf_content.append(Spacer(1, 12))
    
    # Project Information
    add_section('Project Information', f"Project ID: {data.get('project_id', 'Not available')}")
    add_section('Project Name', f"{data.get('project_name', 'Not available')}")
    add_section('Privacy', f"{data.get('privacy', 'Not available')}")
    upload_time = data.get('upload_timestamp', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    add_section('Upload Timestamp', upload_time)
    add_section('Video Status', f"{data.get('video_upload_status', 'Not available')}")
    
    # Video Metadata
    metadata = data.get('metadata', {})
    add_section('File Name', f"{data.get('original_filename', 'Not available')}")
    add_section('Video Hash', f"{data.get('video_hash', 'Not available')}")
    if metadata:
        add_section('File Type', f"{metadata.get('file_type', 'Not available')}")
        add_section('File Size', f"{metadata.get('file_size_MB', 'Not available')} MB")
        add_section('Duration', f"{metadata.get('duration_sec', 'Not available')} seconds")
        add_section('Frame Rate', f"{metadata.get('frame_rate', 'Not available')} fps")
        add_section('Resolution', f"{metadata.get('video_resolution', 'Not available')}")
    
        # Transcription
        if 'transcription_with_timestamps' in metadata:
            doc.add_heading('Transcription', level=1)
            pdf_content.append(Paragraph('Transcription', styles['Heading1']))
            for entry in metadata.get('transcription_with_timestamps', []):
                text = f"[{entry.get('start_time', 0):.2f} - {entry.get('end_time', 0):.2f}] {entry.get('speaker', 'Unknown')}: {entry.get('text', 'No text')}"
                doc.add_paragraph(text)
                pdf_content.append(Paragraph(text, styles['Normal']))
            pdf_content.append(Spacer(1, 12))
    
        # Other Metadata Sections
        for key in ['summary', 'keywords_and_topics', 'brands_and_locations', 'sentiment', 'people', 'objects', 'ai_voice_location']:
            value = metadata.get(key, 'Not available')
            if isinstance(value, dict):
                text = ', '.join(value.get('keywords', []) + value.get('topics', [])) if key == 'keywords_and_topics' else ', '.join(value.get('brands', []) + value.get('locations', []))
            elif isinstance(value, list):
                text = '\n'.join([f"{p.get('name', 'Unknown')} - {p.get('designation', 'Not specified')}" for p in value]) if key == 'people' else ', '.join(value)
            else:
                text = value
            add_section(key.replace('_', ' ').title(), text)
    
    # Footer
    timestamp = f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    pdf_content.append(Spacer(1, 36))
    pdf_content.append(Paragraph(timestamp, styles['Normal']))
    doc.sections[0].footer.paragraphs[0].text = timestamp
    doc.sections[0].footer.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    # Save Word Document as Blob
    word_buffer = BytesIO()
    doc.save(word_buffer)
    word_blob = word_buffer.getvalue()
    word_buffer.close()
    
    # Build PDF
    pdf_doc.build(pdf_content)
    pdf_blob = pdf_buffer.getvalue()
    pdf_buffer.close()
    
    return word_blob, pdf_blob

def get_video_metadata(video_hash: str):
    mongo_uri = os.getenv("MONGO_URI")
    client = MongoClient(mongo_uri)
    db = client['videoDB']
    collection = db['video_processing']
    video_data = collection.find_one({"video_hash": video_hash})
    client.close()
    return video_data if video_data else None

def main(video_hash):
    result = get_video_metadata(video_hash)
    if result:
        return create_document(result)
    else:
        print("No metadata found for the given video hash")
        return None, None

if __name__ == "__main__":
    sample_hash = "ecd71914b736e3664412b10dd127c3447156d2f539181016a90b97cf27fa4e6f"
    word_blob, pdf_blob = main(sample_hash)
    if word_blob and pdf_blob:
        print(f"Word document size: {len(word_blob)} bytes")
        print(f"PDF document size: {len(pdf_blob)} bytes")
        with open("test_output.docx", "wb") as f:
            f.write(word_blob)
        with open("test_output.pdf", "wb") as f:
            f.write(pdf_blob)
