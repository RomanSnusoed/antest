from flask import Flask, render_template, request, redirect, url_for
import os
from cleaner import extract_text_from_pdf, extract_text_from_image, clean_and_remove_personal_info
import spacy

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'pdf', 'jpg', 'jpeg', 'png', "docx"}

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/')
def index():
    return render_template('index.html')

from cleaner import extract_text_from_pdf, extract_text_from_image, extract_text_from_docx, clean_and_remove_personal_info

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(request.url)
    
    file = request.files['file']
    if file and allowed_file(file.filename):
        filename = file.filename
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        # Extract text based on file type
        if filename.endswith('.pdf'):
            extracted_text = extract_text_from_pdf(filepath)
        elif filename.endswith(('.jpg', '.jpeg', '.png')):
            extracted_text = extract_text_from_image(filepath)
        elif filename.endswith('.docx'):
            extracted_text = extract_text_from_docx(filepath)
        else:
            return redirect(url_for('index'))

        # Process text
        nlp = spacy.load("en_core_web_trf")
        doc = nlp(extracted_text)
        spacy_entities = [(ent.text, ent.label_) for ent in doc.ents]
        regex_entities = clean_and_remove_personal_info(extracted_text, return_regex_matches=True)
        cleaned_text = clean_and_remove_personal_info(extracted_text)

        # Render results
        return render_template(
            'result.html',
            extracted_text=extracted_text,
            spacy_entities=spacy_entities,
            spacy_entity_count=len(spacy_entities),
            regex_entities=regex_entities,
            regex_entity_count=len(regex_entities),
            cleaned_text=cleaned_text
        )
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True, port=8080)