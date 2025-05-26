from flask import Flask, render_template, request, redirect, url_for, send_file, flash, session
import os
from pdf2image import convert_from_path
from PIL import Image, ImageChops
import shutil
from werkzeug.utils import secure_filename
import pandas as pd
import xlsxwriter
from datetime import datetime
import math
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from flask import send_from_directory
import atexit
from threading import Thread
from time import sleep
import requests
import subprocess
import tempfile

# Initialize Flask app
app = Flask(__name__, static_folder='static', template_folder='templates')
app.secret_key = 'your-secret-key-123'  # Change this for production

# Configuration
app.config.update(
    UPLOAD_FOLDER_PDF='static/uploads/pdfs',
    OUTPUT_FOLDER_PNG='static/outputs/pngs',
    UPLOAD_FOLDER_EXCEL='static/uploads/excels',
    OUTPUT_FOLDER_REPORTS='static/outputs/reports',
    UPLOAD_FOLDER_SALES_DATA='static/uploads/sales_data',
    OUTPUT_FOLDER_SALES_REPORTS='static/outputs/sales_reports',
    UPLOAD_FOLDER_VBA='static/uploads/vba_pdfs',
    VBA_OUTPUT_FOLDER='static/outputs/vba_output',
    VBA_TEMPLATE_PATH='static/templates/VBAPdfExportTemplate.xlsm',
    ALLOWED_EXTENSIONS_PDF={'pdf'},
    ALLOWED_EXTENSIONS_EXCEL={'xlsx', 'xls', 'xlsm'},
    MAX_CONTENT_LENGTH=16 * 1024 * 1024  # 16MB max upload
)

def clear_all_folders():
    """Clear all upload and output folders"""
    folders_to_clear = [
        app.config['UPLOAD_FOLDER_PDF'],
        app.config['OUTPUT_FOLDER_PNG'],
        app.config['UPLOAD_FOLDER_EXCEL'],
        app.config['OUTPUT_FOLDER_REPORTS'],
        app.config['UPLOAD_FOLDER_SALES_DATA'],
        app.config['OUTPUT_FOLDER_SALES_REPORTS'],
        app.config['UPLOAD_FOLDER_VBA'],
        app.config['VBA_OUTPUT_FOLDER']
    ]

    for folder in folders_to_clear:
        try:
            for filename in os.listdir(folder):
                file_path = os.path.join(folder, filename)
                try:
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        os.unlink(file_path)
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)
                except Exception as e:
                    print(f'Failed to delete {file_path}. Reason: {e}')
        except FileNotFoundError:
            os.makedirs(folder, exist_ok=True)

# Clear folders on application start
clear_all_folders()

# Register cleanup function to run when application exits
atexit.register(clear_all_folders)

# Create folders if they don't exist
for folder in [
    app.config['UPLOAD_FOLDER_PDF'],
    app.config['OUTPUT_FOLDER_PNG'],
    app.config['UPLOAD_FOLDER_EXCEL'],
    app.config['OUTPUT_FOLDER_REPORTS'],
    app.config['UPLOAD_FOLDER_SALES_DATA'],
    app.config['OUTPUT_FOLDER_SALES_REPORTS'],
    app.config['UPLOAD_FOLDER_VBA'],
    app.config['VBA_OUTPUT_FOLDER'],
    os.path.dirname(app.config['VBA_TEMPLATE_PATH'])
]:
    os.makedirs(folder, exist_ok=True)
# Clear folders on application start
clear_all_folders()

# Register cleanup function to run when application exits
atexit.register(clear_all_folders)

# Create folders if they don't exist
for folder in [
    app.config['UPLOAD_FOLDER_PDF'],
    app.config['OUTPUT_FOLDER_PNG'],
    app.config['UPLOAD_FOLDER_EXCEL'],
    app.config['OUTPUT_FOLDER_REPORTS'],
    app.config['UPLOAD_FOLDER_SALES_DATA'],
    app.config['OUTPUT_FOLDER_SALES_REPORTS'],
    app.config['UPLOAD_FOLDER_VBA'],
    os.path.dirname(app.config['VBA_TEMPLATE_PATH'])
]:
    os.makedirs(folder, exist_ok=True)

def allowed_file(filename, file_type='pdf'):
    ext = filename.rsplit('.', 1)[1].lower() if '.' in filename else ''
    if file_type == 'pdf':
        return ext in app.config['ALLOWED_EXTENSIONS_PDF']
    return ext in app.config['ALLOWED_EXTENSIONS_EXCEL']

def trim_whitespace(image):
    bg = Image.new(image.mode, image.size, (255, 255, 255))
    diff = ImageChops.difference(image, bg)
    bbox = diff.getbbox()
    return image.crop(bbox) if bbox else image

def has_png_files():
    return any(fname.endswith('.png') for fname in os.listdir(app.config['OUTPUT_FOLDER_PNG']))

def generate_png_name(base_name, output_folder):
    """Generate unique PNG filename with incremental suffix if needed"""
    counter = 1
    name = f"{base_name}.png"
    while os.path.exists(os.path.join(output_folder, name)):
        name = f"{base_name}_{counter}.png"
        counter += 1
    return name

def safe_divide(numerator, denominator):
    """Safe division that handles division by zero"""
    if denominator == 0:
        return float('inf') if numerator > 0 else float('-inf') if numerator < 0 else 0
    return numerator / denominator

def format_change(current, previous, is_amount=False):
    try:
        if is_amount:
            current = round(current)
            previous = round(previous)
            cur_str = f"{current:,}"
            prev_str = f"{previous:,}"
        else:
            cur_str = f"{current}"
            prev_str = f"{previous}"

        if current == previous:
            return f"{cur_str} (No change)"

        if previous == 0:
            return f"{cur_str} (New data)"

        percent = round(safe_divide((current - previous), abs(previous)) * 100)

        if current > previous:
            return f"{cur_str} (↑ {percent}% from {prev_str})"
        else:
            return f"{cur_str} (↓ {abs(percent)}% from {prev_str})"
    except Exception:
        return f"{current} (Error calculating change)"

def validate_excel_data(df):
    """Validate the Excel data structure and content"""
    required_columns = [
        'Sales Person',
        'Current Bid participate',
        'Last Bid participate',
        'Current Bid Visitor',
        'Last Bid Visitor',
        'Current Bid Winner',
        'Last Bid Winner',
        'Curr Bid Amt(Cnf)',
        'Last Bid Amt(Cnf)',
        'Current Bid Amt(All)',
        'Last Bid Amt(All)'
    ]

    # Check if all required columns exist
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")

    # Convert numeric columns to appropriate types
    numeric_columns = [
        'Current Bid participate', 'Last Bid participate',
        'Current Bid Visitor', 'Last Bid Visitor',
        'Current Bid Winner', 'Last Bid Winner',
        'Curr Bid Amt(Cnf)', 'Last Bid Amt(Cnf)',
        'Current Bid Amt(All)', 'Last Bid Amt(All)'
    ]

    for col in numeric_columns:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    return df

def generate_excel_report(input_path):
    try:
        # Read and validate the Excel file
        df = pd.read_excel(input_path)
        df = validate_excel_data(df)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(app.config['OUTPUT_FOLDER_REPORTS'], f"Bids_summary_{timestamp}.xlsx")

        workbook = xlsxwriter.Workbook(output_file)
        worksheet = workbook.add_worksheet("Summary")

        # Formats
        bold = workbook.add_format({'bold': True})
        normal = workbook.add_format()
        wrap = workbook.add_format({'text_wrap': True})
        error_format = workbook.add_format({'font_color': 'red'})

        row = 0
        for _, r in df.iterrows():
            try:
                name = str(r['Sales Person'])
                name_line = name if name.strip().lower().startswith("mr.") else f"Mr. {name}"

                summary_lines = [
                    [bold, f"{name_line}\n"],
                    [bold, "• Bid Participation: "], normal,
                    format_change(r['Current Bid participate'], r['Last Bid participate']) + "\n",
                    [bold, "• Bid Visitors: "], normal,
                    format_change(r['Current Bid Visitor'], r['Last Bid Visitor']) + "\n",
                    [bold, "• Bid Wins: "], normal, format_change(r['Current Bid Winner'], r['Last Bid Winner']) + "\n",
                    [bold, "• Confirmed Bid Amount: "], normal,
                    format_change(r['Curr Bid Amt(Cnf)'], r['Last Bid Amt(Cnf)'], is_amount=True) + "\n",
                    [bold, "• Total Bid Amount (All): "], normal,
                    format_change(r['Current Bid Amt(All)'], r['Last Bid Amt(All)'], is_amount=True),
                ]

                rich_text = []
                for i in summary_lines:
                    if isinstance(i, list):
                        rich_text.append(i[0])
                        rich_text.append(i[1])
                    else:
                        rich_text.append(i)

                worksheet.write_rich_string(row, 0, *rich_text, wrap)
                row += 1
            except Exception as e:
                worksheet.write_string(row, 0, f"Error processing row {row + 1}: {str(e)}", error_format)
                row += 1

        worksheet.set_column(0, 0, 90)
        workbook.close()
        return output_file
    except Exception as e:
        raise Exception(f"Error generating report: {str(e)}")

def generate_sales_performance_report(input_path):
    try:
        # Read Excel file
        df = pd.read_excel(input_path, sheet_name="Sheet2")

        # Set the first column as index (metrics)
        df.set_index(df.columns[0], inplace=True)

        # Clean column names
        df.columns = [col.strip() for col in df.columns]

        # Function to format values
        def format_value(value, metric):
            if '%' in metric or 'Ratio' in metric:
                return f"{float(value):.2%}"
            elif metric in ['New Gain Customers', 'Total Business', 'Adding Bid', 'Avg Days of stone sold ']:
                return str(int(float(value)))
            elif isinstance(value, float):
                return f"{value:.2f}"
            return str(value)

        # Function to generate clean reports
        def generate_report(person_name):
            relevant_metrics = [
                'New Gain Customers',
                'Total Business',
                'Adding Bid',
                'Customer Active Ratio',
                'Sale Amount % Through Bid ',
                '% of Sale from top 10 customer',
                'Avg Days of stone sold ',
                'Goal Achvd %'
            ]

            report_lines = [
                f"Performance Summary - {person_name}",
                "-" * 40
            ]

            for metric in relevant_metrics:
                if metric in df.index:
                    value = df.loc[metric, person_name]
                    formatted_value = format_value(value, metric)
                    report_lines.append(f"{metric}: {formatted_value}")

            return "\n".join(report_lines)

        # Generate all reports
        all_reports = {person: generate_report(person) for person in df.columns[1:]}

        # Create output filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(app.config['OUTPUT_FOLDER_SALES_REPORTS'],
                                 f"Sales_Performance_Reports_{timestamp}.xlsx")

        # Create a DataFrame for Excel output
        output_data = []
        for person, report in all_reports.items():
            output_data.append([person, report])

        output_df = pd.DataFrame(output_data, columns=['Salesperson', 'Performance Report'])

        # Export to single Excel sheet with perfect formatting
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            output_df.to_excel(writer, sheet_name='Performance Reports', index=False)

            # Get the worksheet and apply formatting
            worksheet = writer.sheets['Performance Reports']

            # Set column widths
            worksheet.column_dimensions['A'].width = 20  # Salesperson column
            worksheet.column_dimensions['B'].width = 60  # Wider report column

            # Apply formatting to all cells
            for row in worksheet.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
                    cell.font = Font(name='Calibri', size=11)

            # Freeze header row
            worksheet.freeze_panes = "A2"

        return output_file
    except Exception as e:
        raise Exception(f"Error generating sales performance report: {str(e)}")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/dashboard')
def dashboard():
    # Check for PDF files
    pdf_files = []
    pdf_folder = app.config['UPLOAD_FOLDER_PDF']
    if os.path.exists(pdf_folder):
        pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')]

    # Check for Excel files
    excel_files = []
    excel_folder = app.config['UPLOAD_FOLDER_EXCEL']
    if os.path.exists(excel_folder):
        excel_files = [f for f in os.listdir(excel_folder) if f.lower().endswith(('.xlsx', '.xls'))]

    # Check for report files
    report_files = []
    report_folder = app.config['OUTPUT_FOLDER_REPORTS']
    if os.path.exists(report_folder):
        report_files = [f for f in os.listdir(report_folder) if f.lower().endswith('.xlsx')]

    # Check for VBA PDFs
    vba_pdfs = []
    vba_folder = app.config['UPLOAD_FOLDER_VBA']
    if os.path.exists(vba_folder):
        vba_pdfs = [f for f in os.listdir(vba_folder) if f.lower().endswith('.pdf')]
        
    # Check for VBA generated PDFs
    vba_generated_pdfs = []
    vba_output_folder = os.path.join(app.config['VBA_OUTPUT_FOLDER'], 'temp')
    if os.path.exists(vba_output_folder):
        vba_generated_pdfs = [f for f in os.listdir(vba_output_folder) if f.lower().endswith('.pdf')]

    # Check for sales data files
    sales_data_files = []
    sales_data_folder = app.config['UPLOAD_FOLDER_SALES_DATA']
    if os.path.exists(sales_data_folder):
        sales_data_files = [f for f in os.listdir(sales_data_folder) if f.lower().endswith(('.xlsx', '.xls'))]

    # Check for sales report files
    sales_report_files = []
    sales_report_folder = app.config['OUTPUT_FOLDER_SALES_REPORTS']
    if os.path.exists(sales_report_folder):
        sales_report_files = [f for f in os.listdir(sales_report_folder) if f.lower().endswith('.xlsx')]

    # Check if any PNGs exist
    has_pngs = False
    png_folder = app.config['OUTPUT_FOLDER_PNG']
    if os.path.exists(png_folder):
        has_pngs = any(f.lower().endswith('.png') for f in os.listdir(png_folder))

    return render_template('dashboard.html',
                         pdf_files=pdf_files,
                         excel_files=excel_files,
                         report_files=report_files,
                         sales_data_files=sales_data_files,
                         sales_report_files=sales_report_files,
                         vba_pdfs=vba_pdfs,
                         vba_generated_pdfs=vba_generated_pdfs,  # Add this line
                         has_pngs=has_pngs)

@app.route('/upload-pdf', methods=['POST'])
def upload_pdf():
    if 'file' not in request.files:
        flash('No file part', 'error')
        return redirect(url_for('dashboard'))

    # Clear upload folder first
    for filename in os.listdir(app.config['UPLOAD_FOLDER_PDF']):
        file_path = os.path.join(app.config['UPLOAD_FOLDER_PDF'], filename)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
        except Exception as e:
            flash(f'Error clearing upload folder: {e}', 'error')

    files = request.files.getlist('file')
    file_count = 0

    for file in files:
        if file.filename == '':
            continue

        if file and allowed_file(file.filename, 'pdf'):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER_PDF'], filename))
            file_count += 1

    if file_count > 0:
        flash(f'Successfully uploaded {file_count} PDF file(s)', 'success')
    else:
        flash('No valid PDF files uploaded', 'error')

    return redirect(url_for('dashboard'))

@app.route('/convert-pdf', methods=['POST'])
def convert_pdf():
    # Clear output folder first
    for filename in os.listdir(app.config['OUTPUT_FOLDER_PNG']):
        file_path = os.path.join(app.config['OUTPUT_FOLDER_PNG'], filename)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
        except Exception as e:
            flash(f'Error clearing output folder: {e}', 'error')

    pdf_files = os.listdir(app.config['UPLOAD_FOLDER_PDF'])
    success_count = 0

    for filename in pdf_files:
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER_PDF'], filename)
            try:
                images = convert_from_path(pdf_path, dpi=300)
                base_name = os.path.splitext(filename)[0]

                for img in images:
                    trimmed = trim_whitespace(img)
                    png_name = generate_png_name(base_name, app.config['OUTPUT_FOLDER_PNG'])
                    trimmed.save(os.path.join(app.config['OUTPUT_FOLDER_PNG'], png_name))

                success_count += 1
                os.remove(pdf_path)
            except Exception as e:
                flash(f'Error converting {filename}: {e}', 'error')

    if success_count > 0:
        flash(f'Successfully converted {success_count} PDF file(s) to PNGs', 'success')

    return redirect(url_for('dashboard'))

@app.route('/download-pngs')
def download_pngs():
    zip_filename = 'converted_images.zip'
    try:
        shutil.make_archive("converted_images", 'zip', app.config['OUTPUT_FOLDER_PNG'])
        return send_file(zip_filename, as_attachment=True)
    except Exception as e:
        flash(f'Error creating ZIP file: {e}', 'error')
        return redirect(url_for('dashboard'))

@app.route('/upload-excel', methods=['POST'])
def upload_excel():
    if 'file' not in request.files:
        flash('No file part', 'error')
        return redirect(url_for('dashboard'))

    # Clear upload folder first
    for filename in os.listdir(app.config['UPLOAD_FOLDER_EXCEL']):
        file_path = os.path.join(app.config['UPLOAD_FOLDER_EXCEL'], filename)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
        except Exception as e:
            flash(f'Error clearing upload folder: {e}', 'error')

    file = request.files['file']
    if file.filename == '':
        flash('No selected file', 'error')
        return redirect(url_for('dashboard'))

    if file and allowed_file(file.filename, 'excel'):
        filename = secure_filename(file.filename)
        file.save(os.path.join(app.config['UPLOAD_FOLDER_EXCEL'], filename))
        flash('Excel file uploaded successfully!', 'success')
    else:
        flash('Invalid file type. Only Excel files are allowed', 'error')

    return redirect(url_for('dashboard'))

@app.route('/generate-report', methods=['POST'])
def generate_report():
    if not os.listdir(app.config['UPLOAD_FOLDER_EXCEL']):
        flash('No Excel files uploaded to generate report', 'error')
        return redirect(url_for('dashboard'))

    try:
        excel_file = os.listdir(app.config['UPLOAD_FOLDER_EXCEL'])[0]
        input_path = os.path.join(app.config['UPLOAD_FOLDER_EXCEL'], excel_file)
        output_file = generate_excel_report(input_path)
        flash('Report generated successfully!', 'success')
        session['latest_report'] = output_file

        # Remove the Excel file after successful report generation
        os.remove(input_path)
    except Exception as e:
        flash(f'Error generating report: {str(e)}', 'error')

    return redirect(url_for('dashboard'))

@app.route('/download-report')
def download_report():
    if 'latest_report' not in session or not os.path.exists(session['latest_report']):
        flash('No report available for download', 'error')
        return redirect(url_for('dashboard'))

    return send_file(session['latest_report'], as_attachment=True)

@app.route('/upload-sales-data', methods=['POST'])
def upload_sales_data():
    if 'file' not in request.files:
        flash('No file part', 'error')
        return redirect(url_for('dashboard'))

    # Clear upload folder first
    for filename in os.listdir(app.config['UPLOAD_FOLDER_SALES_DATA']):
        file_path = os.path.join(app.config['UPLOAD_FOLDER_SALES_DATA'], filename)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
        except Exception as e:
            flash(f'Error clearing upload folder: {e}', 'error')

    file = request.files['file']
    if file.filename == '':
        flash('No selected file', 'error')
        return redirect(url_for('dashboard'))

    if file and allowed_file(file.filename, 'excel'):
        filename = secure_filename(file.filename)
        file.save(os.path.join(app.config['UPLOAD_FOLDER_SALES_DATA'], filename))
        flash('Sales data file uploaded successfully!', 'success')
    else:
        flash('Invalid file type. Only Excel files are allowed', 'error')

    return redirect(url_for('dashboard'))

@app.route('/generate-sales-report', methods=['POST'])
def generate_sales_report():
    if not os.listdir(app.config['UPLOAD_FOLDER_SALES_DATA']):
        flash('No sales data files uploaded to generate report', 'error')
        return redirect(url_for('dashboard'))

    try:
        sales_data_file = os.listdir(app.config['UPLOAD_FOLDER_SALES_DATA'])[0]
        input_path = os.path.join(app.config['UPLOAD_FOLDER_SALES_DATA'], sales_data_file)
        output_file = generate_sales_performance_report(input_path)
        flash('Sales performance report generated successfully!', 'success')
        session['latest_sales_report'] = output_file

        # Remove the sales data file after successful report generation
        os.remove(input_path)
    except Exception as e:
        flash(f'Error generating sales performance report: {str(e)}', 'error')

    return redirect(url_for('dashboard'))

@app.route('/download-sales-report')
def download_sales_report():
    if 'latest_sales_report' not in session or not os.path.exists(session['latest_sales_report']):
        flash('No sales report available for download', 'error')
        return redirect(url_for('dashboard'))

    return send_file(session['latest_sales_report'], as_attachment=True)

@app.route('/process-vba-excel', methods=['POST'])
def process_vba_excel():
    if 'file' not in request.files:
        flash('No file uploaded', 'error')
        return redirect(url_for('dashboard'))

    file = request.files['file']
    if file.filename == '':
        flash('No selected file', 'error')
        return redirect(url_for('dashboard'))

    if file and allowed_file(file.filename, 'excel'):
        try:
            # Clear previous outputs
            temp_output = os.path.join(app.config['VBA_OUTPUT_FOLDER'], 'temp')
            if os.path.exists(temp_output):
                shutil.rmtree(temp_output)
            os.makedirs(temp_output, exist_ok=True)

            # Save the uploaded file
            filename = secure_filename(file.filename)
            input_path = os.path.join(app.config['UPLOAD_FOLDER_VBA'], filename)
            file.save(input_path)
            
            # Process the Excel file
            from openpyxl import load_workbook
            from openpyxl.utils import range_boundaries
            from reportlab.lib.pagesizes import letter, landscape
            from reportlab.pdfgen import canvas
            from reportlab.lib.units import inch
            from reportlab.platypus import Table, TableStyle
            from reportlab.lib import colors

            wb = load_workbook(input_path)
            ws = wb.active

            generated_pdfs = []
            
            # Determine which processing method to use based on sheet structure
            if ws.title == "Screen Shot":  # Bid Summary format
                ranges = [
                    ("A1:BM2", "Summary_Header"),
                    ("A10:BM11", ws['B10'].value),
                    ("A13:BM14", ws['B13'].value),
                    # Add all your ranges here as in your VBA macro
                ]
                
                for rng, name in ranges:
                    if not name:
                        continue
                        
                    pdf_name = f"{name}.pdf"
                    pdf_path = os.path.join(temp_output, pdf_name)
                    
                    # Get the actual cell values from the range
                    min_col, min_row, max_col, max_row = range_boundaries(rng)
                    data = []
                    for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                                          min_col=min_col, max_col=max_col):
                        data.append([cell.value for cell in row])
                    
                    # Create PDF with proper formatting
                    c = canvas.Canvas(pdf_path, pagesize=landscape(letter) if max_col-min_col > 10 else letter)
                    
                    # Create a table with the data
                    table = Table(data)
                    
                    # Style the table to match Excel appearance
                    table.setStyle(TableStyle([
                        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),  # Header row
                        ('FONTSIZE', (0,0), (-1,0), 12),
                        ('BOTTOMPADDING', (0,0), (-1,0), 12),
                        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                    ]))
                    
                    # Draw the table on the PDF
                    table.wrapOn(c, letter[0]-2*inch, letter[1]-2*inch)
                    table.drawOn(c, inch, letter[1]-inch-table._height)
                    
                    c.save()
                    generated_pdfs.append(pdf_name)
            
            else:  # Sales Person format
                for col in range(3, 23):  # Columns C to V
                    header = ws.cell(row=1, column=col).value
                    if not header:
                        continue
                        
                    pdf_name = f"{header}.pdf"
                    pdf_path = os.path.join(temp_output, pdf_name)
                    
                    # Get all data from this column
                    data = []
                    for row in range(1, ws.max_row + 1):
                        cell = ws.cell(row=row, column=col)
                        if cell.value:
                            data.append([cell.value])
                    
                    # Create PDF
                    c = canvas.Canvas(pdf_path, pagesize=letter)
                    
                    # Add title
                    c.setFont("Helvetica-Bold", 16)
                    c.drawString(inch, letter[1]-inch, str(header))
                    
                    # Add data
                    c.setFont("Helvetica", 12)
                    y_position = letter[1] - 1.5*inch
                    for value in data[1:]:  # Skip header
                        c.drawString(inch, y_position, str(value[0]))
                        y_position -= 0.25*inch
                        if y_position < inch:
                            c.showPage()
                            y_position = letter[1] - inch
                    
                    c.save()
                    generated_pdfs.append(pdf_name)
            
            # Create ZIP file immediately
            zip_filename = 'generated_pdfs.zip'
            zip_path = os.path.join(app.config['VBA_OUTPUT_FOLDER'], zip_filename)
            
            # Remove old zip if exists
            if os.path.exists(zip_path):
                os.remove(zip_path)
            
            # Create new zip
            shutil.make_archive(
                os.path.join(app.config['VBA_OUTPUT_FOLDER'], 'generated_pdfs'), 
                'zip', 
                temp_output
            )
            
            # Store only the zip info in session
            session['vba_zip_file'] = zip_filename
            session['vba_generated_count'] = len(generated_pdfs)
            
            # Clean up
            wb.close()
            os.remove(input_path)
            shutil.rmtree(temp_output)
            
            flash(f'Successfully generated {len(generated_pdfs)} PDF file(s). Ready to download.', 'success')
            
        except Exception as e:
            flash(f'Error processing file: {str(e)}', 'error')
            app.logger.error(f"Error in process_vba_excel: {str(e)}")
    else:
        flash('Invalid file type. Only Excel files allowed', 'error')

    return redirect(url_for('dashboard'))

@app.route('/download-vba-pdf/<filename>')
def download_vba_pdf(filename):
    try:
        return send_from_directory(
            os.path.join(app.config['VBA_OUTPUT_FOLDER'], 'temp'),
            filename,
            as_attachment=True
        )
    except FileNotFoundError:
        flash('Requested PDF file not found', 'error')
        return redirect(url_for('dashboard'))

@app.route('/download-all-vba-pdfs')
def download_all_vba_pdfs():
    if 'vba_zip_file' not in session:
        flash('No PDFs available for download', 'error')
        return redirect(url_for('dashboard'))
    
    zip_path = os.path.join(app.config['VBA_OUTPUT_FOLDER'], session['vba_zip_file'])
    
    if not os.path.exists(zip_path):
        flash('PDF files no longer available', 'error')
        return redirect(url_for('dashboard'))
    
    try:
        return send_file(zip_path, as_attachment=True)
    except Exception as e:
        flash(f'Error downloading ZIP file: {e}', 'error')
        return redirect(url_for('dashboard'))

def ping_self():
    while True:
        try:
            # Ping your own Render URL every 5 minutes
            requests.get("https://your-app-name.onrender.com")
            sleep(300)  # 300 seconds = 5 minutes
        except:
            sleep(60)  # If error, wait 1 minute and retry

# Start keep-alive thread when not in debug mode
if not app.debug and os.environ.get("WERKZEUG_RUN_MAIN") != "true":
    t = Thread(target=ping_self)
    t.daemon = True
    t.start()

if __name__ == '__main__':
    app.run(debug=True)
