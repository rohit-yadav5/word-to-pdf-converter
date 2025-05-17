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
    try:
        file = request.files['file']
        if not file:
            return "No file uploaded", 400

        input_path = "temp.docx"
        output_path = "output.pdf"

        # Save uploaded file
        file.save(input_path)

        # Try converting the file
        print("Starting conversion...")
        convert(input_path, output_path)
        print("Conversion successful.")

        # Clean up the input file
        os.remove(input_path)

        return send_file(output_path, as_attachment=True)

    except Exception as e:
        print("Error during conversion:", e)
        return f'''
            <h3>❌ Error during conversion</h3>
            <p>{str(e)}</p>
            <a href="/">Go back</a>
        ''', 500

if __name__ == "__main__":
    app.run(debug=True)
