import os
import io
import re
from flask import Flask, render_template, request, send_file, jsonify
import google.generativeai as genai
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)

# Configure Gemini
GENAI_API_KEY = os.getenv("GEMINI_API_KEY")
genai.configure(api_key=GENAI_API_KEY)

def clean_text(text):
    """Removes Markdown symbols like ** and ## from the text."""
    text = text.replace('**', '').replace('__', '').replace('##', '')
    text = re.sub(r'^-{3,}', '', text, flags=re.MULTILINE)
    return text.strip()

def generate_paper_content(subject, book, chapters, difficulty, grade, time, marks, prompt_text):
    model = genai.GenerativeModel('gemini-pro')
    
    # Updated Prompt for Multiple Chapters
    full_prompt = f"""
    Create a formal exam question paper for CBSE Class {grade}.
    
    Subject: {subject}
    Book/Part: {book}
    Target Chapters/Units: {chapters}
    Difficulty Level: {difficulty}
    Time: {time}
    Max Marks: {marks}
    
    Specific Instructions: {prompt_text}
    
    IMPORTANT INSTRUCTIONS:
    1. Ensure questions are distributed among the selected chapters: {chapters}.
    2. Do NOT use markdown symbols like asterisks (**), hashes (##), or dashes (---).
    3. Structure the paper clearly with "SECTION A", "SECTION B" as headings.
    4. Start directly with "SECTION A".
    """
    
    response = model.generate_content(full_prompt)
    return clean_text(response.text)

def create_word_doc(header_info, content):
    doc = Document()
    
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(11)

    def add_centered_line(text, size=11, is_bold=False):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        run.bold = is_bold
        run.font.size = Pt(size)

    add_centered_line("BHARATIYA VIDYA BHAVAN’S RESIDENTIAL PUBLIC SCHOOL", 14, True)
    add_centered_line("(Sponsored by rice millers’ education and cultural society, W.G.Dt)", 10, False)
    add_centered_line("VIDHYASHRAM – BHIMAVARAM-534201", 12, True)
    
    add_centered_line(f"{header_info['test_name']} – {header_info['academic_year']}", 12, True)

    doc.add_paragraph()

    table = doc.add_table(rows=2, cols=2)
    table.autofit = True
    
    table.cell(0, 0).text = f"Sub: {header_info['subject']}"
    cell_r1 = table.cell(0, 1)
    cell_r1.text = f"Time: {header_info['time']}"
    cell_r1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    table.cell(1, 0).text = f"Class: {header_info['grade']}"
    cell_r2 = table.cell(1, 1)
    cell_r2.text = f"Marks: {header_info['marks']}"
    cell_r2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    run.bold = True

    p_line = doc.add_paragraph("_" * 80)
    p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph() 

    lines = content.split('\n')
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        p = doc.add_paragraph()
        run = p.add_run(line)
        
        if (line.upper().startswith("SECTION") or line.upper().startswith("PART")) and len(line) < 40:
            run.bold = True
            run.font.size = Pt(12)
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(6)

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    data = request.json
    try:
        text = generate_paper_content(
            data.get('subject'),
            data.get('book'),
            data.get('chapters'),  # This will now be a string of multiple chapters
            data.get('difficulty'),
            data.get('grade'),
            data.get('time'),
            data.get('marks'),
            data.get('prompt')
        )
        return jsonify({'success': True, 'content': text})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/download', methods=['POST'])
def download():
    data = request.form
    header_info = {
        'subject': data.get('subject'),
        'grade': data.get('grade'),
        'test_name': data.get('testName'),
        'academic_year': data.get('year'),
        'time': data.get('time'),
        'marks': data.get('marks')
    }
    content = data.get('content')
    
    file_stream = create_word_doc(header_info, content)
    
    return send_file(
        file_stream,
        as_attachment=True,
        download_name=f"{data.get('subject')}_Test.docx",
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
