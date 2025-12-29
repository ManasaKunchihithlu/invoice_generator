"""
Flask Web Application for Invoice Generator
Allows users to upload Excel files and generate PDF invoices
"""

from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
import os
import logging
import traceback
from invoice_generator import InvoiceGenerator

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this'  # Change this to a random secret key
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls'}

# Create necessary folders
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs('generated_invoices', exist_ok=True)

# Set up basic logging to a file for debugging web errors
logging.basicConfig(
    filename='invoice_generator.log',
    level=logging.INFO,
    format='%(asctime)s %(levelname)s %(message)s'
)


def allowed_file(filename):
    """Check if the uploaded file has an allowed extension"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']


@app.route('/')
def index():
    """Main page with upload form"""
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and generate invoices"""
    if 'file' not in request.files:
        flash('No file selected', 'error')
        return redirect(url_for('index'))
    
    file = request.files['file']
    
    if file.filename == '':
        flash('No file selected', 'error')
        return redirect(url_for('index'))
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        try:
            # Generate invoices using the uploaded file
            # InvoiceGenerator expects a config file path (default 'config.json'),
            # so create the generator without passing the excel file.
            generator = InvoiceGenerator()
            generated_files = generator.process_excel_file(filepath)
            count = len(generated_files)

            flash(f'Success! Generated {count} invoices in the generated_invoices folder.', 'success')
            return render_template('success.html', count=count)

        except Exception as e:
            # Log full traceback for debugging
            tb = traceback.format_exc()
            logging.error("Error generating invoices: %s", tb)
            # Surface a simple message to the user
            flash('Error generating invoices. The error has been logged for investigation.', 'error')
            return redirect(url_for('index'))
    
    else:
        flash('Invalid file type. Please upload an Excel file (.xlsx or .xls)', 'error')
        return redirect(url_for('index'))


@app.route('/download/<invoice_number>')
def download_invoice(invoice_number):
    """Download a specific invoice"""
    filepath = os.path.join('generated_invoices', f'{invoice_number}.pdf')
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    else:
        flash('Invoice not found', 'error')
        return redirect(url_for('index'))


if __name__ == '__main__':
    print("Starting Invoice Generator Web Application...")
    print("Access the application at: http://localhost:5000")
    app.run(debug=True, host='0.0.0.0', port=5000)
