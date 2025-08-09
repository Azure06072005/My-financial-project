from flask import Flask, request, jsonify
from flask_cors import CORS
import pandas as pd
import os
from werkzeug.utils import secure_filename
import json

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

# Configuration
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# Create upload directory if it doesn't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def process_excel_sheet(excel_path, sheet_name):
    """Process Excel sheet and return DataFrame"""
    try:
        print(f"Reading sheet {sheet_name} from {os.path.basename(excel_path)}")
        
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=5, index_col=0)
        
        df.dropna(how='all', axis=0, inplace=True)
        df.dropna(how='all', axis=1, inplace=True)
        
        df.index.name = 'Chỉ tiêu'
        df.columns.name = 'Quý'
        
        print(f"Successfully processed sheet {sheet_name}")
        return df
    
    except ValueError as e:
        print(f"Error: Sheet {sheet_name} not found in {excel_path}")
        return None
    except Exception as e:
        print(f"An unexpected error occurred while processing sheet {sheet_name}: {e}")
        return None

def dataframe_to_json(df):
    """Convert DataFrame to JSON format suitable for frontend"""
    if df is None:
        return None
    
    # Reset index to make it a column
    df_reset = df.reset_index()
    
    # Convert to dict format
    return {
        'columns': df_reset.columns.tolist(),
        'data': df_reset.values.tolist(),
        'shape': df.shape
    }

@app.route('/api/upload', methods=['POST'])
def upload_file():
    """Handle file upload and process Excel sheets"""
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        
        # Process the Excel file
        sheet_mapping = {
            'balance_sheet': 'CDKT',
            'income_statement': 'KQKD', 
            'financial_ratios': 'CSTC'
        }
        
        results = {}
        processed_sheets = []
        
        for df_key, sheet_name in sheet_mapping.items():
            df = process_excel_sheet(file_path, sheet_name)
            if df is not None:
                results[df_key] = dataframe_to_json(df)
                processed_sheets.append(sheet_name)
        
        # Clean up uploaded file
        try:
            os.remove(file_path)
        except Exception as e:
            print(f"Warning: Could not delete file {file_path}: {e}")
        
        if not results:
            return jsonify({'error': 'No valid sheets found in the Excel file'}), 400
        
        return jsonify({
            'message': f'Successfully processed {len(processed_sheets)} sheets',
            'processed_sheets': processed_sheets,
            'data': results
        })
    
    return jsonify({'error': 'Invalid file type. Please upload .xlsx or .xls files'}), 400

@app.route('/api/sheets', methods=['GET'])
def get_available_sheets():
    """Get available sheet types"""
    return jsonify({
        'sheets': [
            {'key': 'balance_sheet', 'name': 'Balance Sheet (CDKT)', 'display': 'Cân đối kế toán'},
            {'key': 'income_statement', 'name': 'Income Statement (KQKD)', 'display': 'Kết quả kinh doanh'},
            {'key': 'financial_ratios', 'name': 'Financial Ratios (CSTC)', 'display': 'Chỉ số tài chính'}
        ]
    })

@app.route('/api/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    return jsonify({'status': 'healthy', 'message': 'Flask server is running'})

if __name__ == '__main__':
    print("Starting Flask server on http://localhost:5000")
    print("Health check: http://localhost:5000/api/health")
    print("Upload endpoint: http://localhost:5000/api/upload")
    app.run(debug=True, host='127.0.0.1', port=5000)