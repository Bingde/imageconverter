from flask import Flask, session, request, send_file, render_template, redirect, url_for, jsonify, session
from datetime import datetime, timedelta
from dotenv import load_dotenv
import threading
import time
import pytesseract
from PIL import Image
import pandas as pd
import re
import io
import cv2
import os
import numpy as np
import sys
# Add the virtual environment site-packages directory to sys.path
load_dotenv()  # Load environment variables from .env file
venv_path = os.getenv("VENEPATH")
if venv_path not in sys.path:
    sys.path.append(venv_path)
import fitz  # PyMuPDF
import secrets
import uuid
import subprocess
import logging
from logging.handlers import RotatingFileHandler


# Generate a secure secret key (only do this once and store it securely)
if not os.environ.get("FLASK_SECRET_KEY"):
    os.environ["FLASK_SECRET_KEY"] = secrets.token_hex(32)

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY")

# Set up logging
if not app.debug:
    file_handler = RotatingFileHandler(
        "app.log",  # Log file name
        maxBytes=1024 * 1024,  # 1 MB per file
        backupCount=10  # Keep up to 10 log files
    )
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(logging.Formatter(
        "%(asctime)s %(levelname)s: %(message)s [in %(pathname)s:%(lineno)d]"
    ))
    app.logger.addHandler(file_handler)


# Set up folders
uploads = os.getenv("UPLOAD_FOLDER")

compressed = os.getenv("COMPRESSED_FOLDER")

exceled = os.getenv("EXCEL_FOLDER")


# Ensure folders exist
os.makedirs(uploads, exist_ok=True)
os.makedirs(compressed, exist_ok=True)
os.makedirs(exceled, exist_ok=True)

def get_file_size(file_path):
    """
    Get the size of a file in megabytes (MB).
    :param file_path: Path to the file.
    :return: File size in MB as a float.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"The file {file_path} does not exist.")
    
    size_bytes = os.path.getsize(file_path)
    size_mb = size_bytes / (1024 * 1024)  # Convert bytes to MB
    return round(size_mb, 3)  # Round to 2 decimal places

def compress_pdf(input_path, output_path, image_quality):
    """
    Compress a PDF file using PyMuPDF and Ghostscript.
    :param input_path: Path to the input PDF file.
    :param output_path: Path to save the compressed PDF file.
    :param image_quality: Image quality (1-100). Lower values mean higher compression.
    """
    # Step 1: Recompress images using PyMuPDF
    pdf_document = fitz.open(input_path)

    # Iterate through each page and recompress images
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        image_list = page.get_images(full=True)

        for img in image_list:
            xref = img[0]  # Image XREF
            base_image = pdf_document.extract_image(xref)
            image_bytes = base_image["image"]

            # Open the image using PIL (Pillow)
            image = Image.open(io.BytesIO(image_bytes))

            # Recompress the image (reduce quality)
            if image.mode == "L":  # Grayscale
                img_format = "JPEG" if base_image["bpc"] == 8 else "PNG"
            else:  # Color
                img_format = "JPEG"

            # Save the recompressed image to a bytes buffer
            buffer = io.BytesIO()
            image.save(buffer, format=img_format, quality=image_quality)
            buffer.seek(0)

            # Get the original image's position and size
            image_rect = page.get_image_bbox(img)  # Get the bounding box of the image

            # Reinsert the recompressed image at the same location and size
            page.insert_image(
                image_rect,  # Use the original image's bounding box
                stream=buffer.read(),  # Recompressed image data
            )

    # Save the intermediate PDF
    intermediate_path = os.path.join(os.path.dirname(output_path), "intermediate.pdf")
    pdf_document.save(intermediate_path, garbage=4, deflate=True)
    pdf_document.close()

    # Step 2: Use Ghostscript for advanced compression
    ghostscript_command = [
        "gs",
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        "-dPDFSETTINGS=/screen",  # Lowest quality, smallest size
        "-dNOPAUSE",
        "-dQUIET",
        "-dBATCH",
        f"-sOutputFile={output_path}",
        intermediate_path,
    ]

    try:
        subprocess.run(ghostscript_command, check=True)
    except subprocess.CalledProcessError as e:
        print(f"Ghostscript compression failed: {e}")
    finally:
        # Clean up the intermediate file
        if os.path.exists(intermediate_path):
            os.remove(intermediate_path)

# Image-to-Excel Function
def image_to_excel(image_path, output_path):
    # Convert image to grayscale and enhance contrast
    image = np.array(Image.open(image_path))  # Convert PIL image to NumPy array
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)  # Convert to grayscale
    _, binary = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)  # Binarize

    # Use Tesseract to extract text
    extracted_text = pytesseract.image_to_string(binary, config='--psm 6')

    # Convert extracted text to a table
    rows = extracted_text.split('\n')
    table_data = [row.split() for row in rows]

    # Create a DataFrame and save as Excel
    df = pd.DataFrame(table_data)
    df.to_excel(output_path, index=False)

def sanitize_filename(filename):
    """
    Sanitize the filename to remove special characters.
    :param filename: The original filename.
    :return: A sanitized filename.
    """
    # Replace spaces and special characters with underscores
    sanitized = re.sub(r"[^\w.-]", "_", filename)
    return sanitized


def calculate_reduction_percentage(original_size, compressed_size):
    """
    Calculate the percentage reduction in file size.
    :param original_size: Original file size in MB.
    :param compressed_size: Compressed file size in MB.
    :return: Percentage reduction as a float.
    """
    if original_size == 0:
        return 0  # Avoid division by zero
    reduction = ((original_size - compressed_size) / original_size) * 100
    return round(reduction, 2)  # Round to 2 decimal places

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/compress", methods=["GET", "POST"])
def compress():
    if request.method == "POST":
        if "file" not in request.files:
            return redirect(request.url)
        
        file = request.files["file"]
        if file.filename == "":
            return redirect(request.url)
        
        if file and file.filename.endswith(".pdf"):
            # Sanitize the file name
            sanitized_filename = sanitize_filename(file.filename)
            
            # Save the uploaded file
            upload_path = os.path.join(uploads, sanitized_filename)
            file.save(upload_path)
            
            # Verify the file was saved
            if not os.path.exists(upload_path):
                return "Error: File was not saved correctly.", 400
            
            # Get the original file size
            original_size = get_file_size(upload_path)
            
            # Get the selected compression quality
            quality = int(request.form.get("quality", 50))  # Default to 50 if not provided
            
            # Map quality to compression level
            if quality == 10:
                compress_level = "Low"
            elif quality == 50:
                compress_level = "Medium"
            elif quality == 90:
                compress_level = "High"
            else:
                compress_level = "Unknown"  # Fallback for unexpected values
            filename_without_ext = file.filename.rstrip(".pdf")
            # Generate a unique name for the compressed file
            filename_without_ext = file.filename.rstrip(".pdf")
            compressed_filename = f"compressed_{compress_level}_{filename_without_ext}.pdf"
            compressed_path = os.path.join(compressed, compressed_filename)
            
            # Compress the PDF with the selected quality
            compress_pdf(upload_path, compressed_path, image_quality=quality)

            # Get the compressed file size
            compressed_size = get_file_size(compressed_path)
            
            # Calculate the percentage reduction
            reduction_percentage = calculate_reduction_percentage(original_size, compressed_size)
            
            
            # Render the template with file sizes, reduction percentage, and download link
            return render_template(
                "compress.html",
                original_size=original_size,
                compressed_size=compressed_size,
                reduction_percentage=reduction_percentage,
                download_link=url_for("download_file", filename=compressed_filename)
            )
    
    return render_template("compress.html")

@app.route("/image-to-excel", methods=["GET", "POST"])
def image_to_excel_route():
    if request.method == "POST":
        if "image" not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files["image"]
        if file.filename == "":
            return redirect(request.url)
        
        if file and file.filename.lower().endswith((".png", ".jpg", ".jpeg")):
            # Save the uploaded image
            upload_path = os.path.join(uploads, file.filename)
            file.save(upload_path)
            
            # Generate a unique name for the Excel file
            excel_filename = f"output_{uuid.uuid4().hex}.xlsx"
            excel_path = os.path.join(exceled, excel_filename)
            
            # Convert image to Excel
            image_to_excel(upload_path, excel_path)
            
            
            # Redirect to the download page
            return redirect(url_for("download", filename=excel_filename, type="excel"))
    
    return render_template("image_to_excel.html")

@app.route("/download/<filename>")
def download_file(filename):
    compressed_path = os.path.join(compressed, filename)
    return send_file(
        compressed_path,
        as_attachment=True,
        download_name=filename,
        mimetype="application/pdf"
    )

@app.route("/download/<type>/<filename>")
def download(type, filename):
    if type == "pdf":
        file_path = os.path.join(compressed, filename)
    elif type == "excel":
        file_path = os.path.join(exceled, filename)
    else:
        return "Invalid file type."
    
    return send_file(file_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=False)