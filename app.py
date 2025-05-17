import subprocess
from flask import Flask, request, send_file
import os
import tempfile

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

        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = os.path.join(tmpdir, "input.docx")
            output_path = os.path.join(tmpdir, "input.pdf")

            # Save uploaded file to temp folder
            file.save(input_path)

            # Convert using LibreOffice
            subprocess.run([
                '/Applications/LibreOffice.app/Contents/MacOS/soffice',
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', tmpdir,
                input_path
            ], check=True)

            # Send converted PDF to user
            return send_file(output_path, as_attachment=True)

    except Exception as e:
        return f"<h3>Error during conversion:</h3><pre>{e}</pre>", 500

if __name__ == "__main__":
    app.run(debug=True)
