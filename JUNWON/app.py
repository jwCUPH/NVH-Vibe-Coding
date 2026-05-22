import os, sys, warnings, json, re, openpyxl
from flask import Flask, render_template, request, jsonify, Response
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

# Suppress openpyxl warnings & init Flask
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
base_dir = os.path.abspath(os.path.dirname(__file__))

app = Flask(__name__, template_folder=os.path.join(base_dir, 'templates'), static_folder=os.path.join(base_dir, 'static'))
app.config.update(
    UPLOAD_FOLDER=os.path.join(base_dir, 'uploads'),
    SQLALCHEMY_DATABASE_URI='sqlite:///' + os.path.join(base_dir, 'vpr_reports.db'),
    SQLALCHEMY_TRACK_MODIFICATIONS=False
)

db = SQLAlchemy(app)

# Database Model
class VPRReport(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    req_no = db.Column(db.String(100))
    project = db.Column(db.String(200))
    date = db.Column(db.String(50))
    content_html = db.Column(db.Text)
    raw_data = db.Column(db.Text) 
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

with app.app_context():
    db.create_all()
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# --- Routes ---
@app.route('/')
def index(): return render_template('index.html')

@app.route('/validate', methods=['POST'])
def validate_file():
    file = request.files.get('file')
    if not file or file.filename == '': return jsonify({'status': 'NG'})
    try:
        openpyxl.load_workbook(file, read_only=True)
        return jsonify({'status': 'OK'})
    except: return jsonify({'status': 'NG'})

# Helper functions for Parsing
def excel_coord_to_indices(coord):
    match = re.match(r"([A-Z]+)([0-9]+)", coord)
    if not match: return None, None
    c_s, r_s = match.groups()
    col = 0
    for char in c_s: col = col * 26 + (ord(char) - ord('A') + 1)
    return int(r_s), col

def get_val(sheet, coord):
    r, c = excel_coord_to_indices(coord)
    if not r: return ""
    v = sheet.cell(row=r, column=c).value
    return str(v) if v is not None else ""

def get_min_max(sheet, range_str):
    match = re.match(r"([A-Z]+)([0-9]+)~([A-Z]+)([0-9]+)", range_str)
    if not match: return "N/A"
    sc_s, sr, ec_s, er = match.groups()
    _, sc = excel_coord_to_indices(sc_s + sr)
    _, ec = excel_coord_to_indices(ec_s + er)
    vals = [v for c in range(sc, ec + 1) if isinstance(v := sheet.cell(row=int(sr), column=c).value, (int, float))]
    return f"{min(vals)}~{max(vals)}" if vals else "N/A"

def get_unique_vals(sheet, range_str):
    match = re.match(r"([A-Z]+)([0-9]+)~([A-Z]+)([0-9]+)", range_str)
    if not match: return "N/A"
    sc_s, sr, ec_s, er = match.groups()
    _, sc = excel_coord_to_indices(sc_s + sr)
    _, ec = excel_coord_to_indices(ec_s + er)
    vals = [str(v).strip() for c in range(sc, ec + 1) if (v := sheet.cell(row=int(sr), column=c).value) is not None]
    unique = list(dict.fromkeys(vals))
    return ", ".join(unique) if unique else "N/A"

@app.route('/extract', methods=['POST'])
def extract_data():
    files = request.files.getlist('files')
    if not files: return jsonify({'status': 'error'})
    
    saved_files = []
    for i, f in enumerate(files):
        p = os.path.join(app.config['UPLOAD_FOLDER'], f"temp_{i}_{f.filename}")
        f.save(p); saved_files.append((p, f.filename))

    def generate():
        res = []
        total = len(saved_files)
        try:
            for i, (path, name) in enumerate(saved_files):
                pct = 5 + (i/total)*90
                step = f"[{i+1}/{total}]"
                yield f"data: {json.dumps({'status': 'progress', 'percent': int(pct), 'message': f'{step} {name} 분석 중...'})}\n\n"
                
                wb = openpyxl.load_workbook(path, data_only=True)
                db_s = wb['DB'] if 'DB' in wb.sheetnames else None
                if db_s:
                    res.append({
                        'day': i+1, 'client_info': get_val(db_s, 'G4'), 'project': get_val(db_s, 'G3'),
                        'req_no': get_val(db_s, 'G2'), 'date': get_val(db_s, 'F15'), 'testers': f"{get_val(db_s, 'G13')}, {get_val(db_s, 'G14')}",
                        'vehicle': f"{get_val(db_s, 'G22')} {get_val(db_s, 'G23')}",
                        'tire_ref': {'size': get_val(db_s, 'G17'), 'pattern': get_val(db_s, 'G18'), 'brand': get_val(db_s, 'G19'), 'marking': get_val(db_s, 'G20')},
                        'tire_sample': {'size': get_unique_vals(db_s, 'H17~Q17'), 'pattern': get_unique_vals(db_s, 'H18~Q18'), 'brand': get_unique_vals(db_s, 'H19~Q19'), 'marking': get_unique_vals(db_s, 'H20~Q20')},
                        'temp_air': get_min_max(db_s, 'G39~Q39'), 'temp_road': get_min_max(db_s, 'G40~Q40')
                    })
                wb.close(); os.remove(path)
            yield f"data: {json.dumps({'status': 'success', 'data': res, 'file_names': [f[1] for f in saved_files]})}\n\n"
        except Exception as e: yield f"data: {json.dumps({'status': 'error', 'message': str(e)})}\n\n"
    return Response(generate(), mimetype='text/event-stream')

@app.route('/save', methods=['POST'])
def save_report():
    d = request.json
    try:
        new_r = VPRReport(req_no=d.get('req_no'), project=d.get('project'), date=d.get('date'), content_html=d.get('html_content'), raw_data=json.dumps(d.get('raw_data')))
        db.session.add(new_r); db.session.commit()
        return jsonify({'status': 'success'})
    except Exception as e: return jsonify({'status': 'error', 'message': str(e)})

@app.route('/reports')
def list_reports():
    rs = VPRReport.query.order_by(VPRReport.created_at.desc()).all()
    return jsonify([{'id': r.id, 'req_no': r.req_no, 'project': r.project, 'date': r.date, 'created_at': r.created_at.strftime('%Y-%m-%d %H:%M')} for r in rs])

@app.route('/report/<int:report_id>')
def get_report(report_id):
    r = VPRReport.query.get_or_404(report_id)
    return jsonify({'content': r.content_html, 'raw_data': json.loads(r.raw_data) if r.raw_data else None})

if __name__ == '__main__': app.run(debug=True, port=5000)
