from flask import Flask, request, send_file
from docx2pdf import convert
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
    file = request.files['file']
    if file:
        try:
            file.save("temp.docx")
            convert("temp.docx", "output.pdf")
            os.remove("temp.docx")
            return send_file("output.pdf", as_attachment=True)
        except Exception as e:
            return f"Error during conversion: {e}", 500
    return "No file uploaded", 400

if __name__ == "__main__":
    app.run(debug=True)
