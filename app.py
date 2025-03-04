from flask import Flask, request, send_file, render_template, jsonify
import pytesseract
from PIL import Image
import pandas as pd
import io
import cv2
import numpy as np

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_image():
    if 'image' not in request.files:
        return 'No file uploaded', 400

    file = request.files['image']
    image = Image.open(file.stream)

    # Convert image to grayscale and enhance contrast
    image = np.array(image)  # Convert PIL image to NumPy array
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)  # Convert to grayscale
    _, binary = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)  # Binarize

    # Use Tesseract to extract text
    extracted_text = pytesseract.image_to_string(binary, config='--psm 6')

    # Convert extracted text to a table
    rows = extracted_text.split('\n')
    table_data = [row.split() for row in rows]

    # Create a DataFrame and save as Excel
    df = pd.DataFrame(table_data)
    excel_file = io.BytesIO();
    df.to_excel(excel_file, index=False)
    excel_file.seek(0)

    return send_file(
        excel_file,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='converted_table.xlsx'
    )

@app.route('/extract', methods=['POST'])
def extract_content():
    if 'image' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['image']
    image = Image.open(file.stream)

    # Convert image to grayscale and enhance contrast
    image = np.array(image)  # Convert PIL image to NumPy array
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)  # Convert to grayscale
    _, binary = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)  # Binarize

    # Use Tesseract to extract text
    extracted_text = pytesseract.image_to_string(binary, config='--psm 6')

    return jsonify({'content': extracted_text})

if __name__ == '__main__':
    app.run(debug=True)