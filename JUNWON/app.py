from flask import Flask, render_template, request, jsonify, send_from_directory
import os
import pandas as pd
import openpyxl
import re
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///vpr_reports.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# Database Model
class VPRReport(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    project_name = db.Column(db.String(200))
    content_html = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

with app.app_context():
    db.create_all()

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/validate', methods=['POST'])
def validate_file():
    if 'file' not in request.files:
        return jsonify({'status': 'NG', 'message': 'No file'})
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'status': 'NG', 'message': 'Empty filename'})

    try:
        # Simple check if file can be opened as excel
        pd.read_excel(file, nrows=1)
        return jsonify({'status': 'OK'})
    except Exception as e:
        return jsonify({'status': 'NG', 'message': str(e)})

@app.route('/extract', methods=['POST'])
def extract_data():
    # Placeholder for Step 2
    return jsonify({'status': 'success', 'message': 'Extraction logic pending Step 2'})

if __name__ == '__main__':
    app.run(debug=True, port=5000)
