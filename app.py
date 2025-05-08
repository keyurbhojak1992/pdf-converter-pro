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
import atexit

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
    ALLOWED_EXTENSIONS_PDF={'pdf'},
    ALLOWED_EXTENSIONS_EXCEL={'xlsx', 'xls'},
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
        app.config['OUTPUT_FOLDER_SALES_REPORTS']
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
    app.config['OUTPUT_FOLDER_SALES_REPORTS']
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
    # Ensure folders are empty on new session
    if 'session_started' not in session:
        clear_all_folders()
        session['session_started'] = True

    pdf_files = [f for f in os.listdir(app.config['UPLOAD_FOLDER_PDF']) if f.endswith('.pdf')]
    excel_files = [f for f in os.listdir(app.config['UPLOAD_FOLDER_EXCEL']) if f.endswith(('.xlsx', '.xls'))]
    has_pngs = has_png_files()
    report_files = [f for f in os.listdir(app.config['OUTPUT_FOLDER_REPORTS']) if f.endswith('.xlsx')]
    sales_data_files = [f for f in os.listdir(app.config['UPLOAD_FOLDER_SALES_DATA']) if f.endswith(('.xlsx', '.xls'))]
    sales_report_files = [f for f in os.listdir(app.config['OUTPUT_FOLDER_SALES_REPORTS']) if f.endswith('.xlsx')]

    # Prepare PNG files list
    png_files = []
    if has_pngs:
        png_files = [f for f in os.listdir(app.config['OUTPUT_FOLDER_PNG']) if f.endswith('.png')]

    return render_template('dashboard.html',
                           pdf_files=pdf_files,
                           excel_files=excel_files,
                           has_pngs=has_pngs,
                           report_files=report_files,
                           png_files=png_files,
                           sales_data_files=sales_data_files,
                           sales_report_files=sales_report_files)


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
  
# ===== KEEP-ALIVE SETUP =====
from threading import Thread
from time import sleep
import requests

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
# ===== END KEEP-ALIVE =====
if __name__ == '__main__':
    app.run(debug=True)
