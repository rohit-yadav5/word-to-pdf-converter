import subprocess
from flask import Flask, request, send_file
import os

app = Flask(__name__)

@app.route('/')
def home():
    return '''
    <h2>Upload Word file to convert to PDF</h2>
    <form method="POST" action="/convert" enctype="multipart/form-data">
      <input type="file" name="file" accept=".docx" required />
      <input type="submit" value="Convert" />
    </form>
    '''

@app.route('/convert', methods=['POST'])
def convert_file():
    try:
        file = request.files['file']
        if not file:
            return "No file uploaded", 400

        input_docx = "temp.docx"
        output_pdf = "temp.pdf"

        # Save the uploaded Word file
        file.save(input_docx)

        # Run LibreOffice to convert .docx to .pdf
        # The output PDF will be saved in the same directory as input_docx
        subprocess.run([
            '/Applications/LibreOffice.app/Contents/MacOS/soffice',
            '--headless',
            '--convert-to',
            'pdf',
            input_docx
        ], check=True)

        # Check if PDF was created
        if not os.path.exists("temp.pdf"):
            return "PDF conversion failed", 500

        # Send the converted PDF file to the user
        return send_file("temp.pdf", as_attachment=True)

    except Exception as e:
        return f"<h3>Error during conversion:</h3><pre>{e}</pre>", 500

if __name__ == "__main__":
    app.run(debug=True)
