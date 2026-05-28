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
    # Use iter_rows for faster access in read_only mode even for single cell
    for row in sheet.iter_rows(min_row=r, max_row=r, min_col=c, max_col=c, values_only=True):
        v = row[0]
        return str(v) if v is not None else ""
    return ""

def get_min_max(sheet, range_str):
    match = re.match(r"([A-Z]+)([0-9]+)~([A-Z]+)([0-9]+)", range_str)
    if not match: return "N/A"
    sc_s, sr, ec_s, er = match.groups()
    _, sc = excel_coord_to_indices(sc_s + sr)
    _, ec = excel_coord_to_indices(ec_s + er)
    
    vals = []
    for row in sheet.iter_rows(min_row=int(sr), max_row=int(sr), min_col=sc, max_col=ec, values_only=True):
        vals = [v for v in row if isinstance(v, (int, float))]
    return f"{min(vals)}~{max(vals)}" if vals else "N/A"

def get_unique_vals(sheet, range_str):
    match = re.match(r"([A-Z]+)([0-9]+)~([A-Z]+)([0-9]+)", range_str)
    if not match: return "N/A"
    sc_s, sr, ec_s, er = match.groups()
    _, sc = excel_coord_to_indices(sc_s + sr)
    _, ec = excel_coord_to_indices(ec_s + er)
    
    vals = []
    for row in sheet.iter_rows(min_row=int(sr), max_row=int(sr), min_col=sc, max_col=ec, values_only=True):
        vals = [str(v).strip() for v in row if v is not None]
    unique = list(dict.fromkeys(vals))
    return ", ".join(unique) if unique else "N/A"

def get_table_data(sheet, start_coord, end_coord):
    sr, sc = excel_coord_to_indices(start_coord)
    er, ec = excel_coord_to_indices(end_coord)
    data = []
    for row in sheet.iter_rows(min_row=sr, max_row=er, min_col=sc, max_col=ec, values_only=True):
        if any(v is not None for v in row):
            # Keep raw values (int/float) for delta calculation in frontend
            data.append(list(row))
    return data

def get_spectrum_data(sheet, cols_groups):
    # cols_groups: list of column indices like [1, 15, 29, 43]
    # frequency is in column 1
    # samples are in cols [g+1, g+2, g+3, g+4, g+5]
    
    # Pre-fetch all needed data in one go (up to 4000Hz, max col 50ish)
    all_data = list(sheet.iter_rows(min_row=1, max_row=2003, min_col=1, max_col=50, values_only=True))
    if not all_data: return {}

    row1 = all_data[0]
    row2 = all_data[1]
    
    data = {}
    for g in cols_groups:
        cat_name = row1[g-1]
        if not cat_name: continue
        
        samples = []
        for c in range(g + 1, g + 6):
            sname = row2[c-1]
            if sname: samples.append({'name': sname, 'col_idx': c-1})
        
        freqs = []
        values = {s['name']: [] for s in samples}
        
        for r_idx in range(2, len(all_data)):
            row = all_data[r_idx]
            f = row[0]
            if f is None: break
            freqs.append(f)
            for s in samples:
                val = row[s['col_idx']]
                values[s['name']].append(val if val is not None else 0)
        
        data[cat_name] = {'freqs': freqs, 'samples': values}
    return data

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
                base_pct = (i / total) * 100
                file_pct = 100 / total
                step = f"[{i+1}/{total}]"
                
                yield f"data: {json.dumps({'status': 'progress', 'percent': int(base_pct + file_pct * 0.05), 'message': f'{step} {name} 엑셀 파일 로드 중...'})}\n\n"
                
                wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
                db_s = wb['DB'] if 'DB' in wb.sheetnames else None
                rep_s = wb['Report'] if 'Report' in wb.sheetnames else None
                rr_s = wb['RR'] if 'RR' in wb.sheetnames else None
                sr_s = wb['SR'] if 'SR' in wb.sheetnames else None

                day_data = {'day': i+1}
                
                yield f"data: {json.dumps({'status': 'progress', 'percent': int(base_pct + file_pct * 0.2), 'message': f'{step} {name} 기본 정보(DB 시트) 추출 중...'})}\n\n"
                if db_s:
                    g22_val = get_val(db_s, 'G22')
                    g17_val = get_val(db_s, 'G17')
                    wheel_size_match = re.search(r'R(\d{2})', g22_val + g17_val)
                    wheel_size = wheel_size_match.group(1) if wheel_size_match else ""
                    
                    # Extract markings list using iter_rows
                    markings = []
                    for row in db_s.iter_rows(min_row=20, max_row=20, min_col=8, max_col=17, values_only=True):
                        markings = [v for v in row if v is not None]
                    
                    day_data.update({
                        'client_info': get_val(db_s, 'G7'), 'project': get_val(db_s, 'G3'),
                        'req_no': get_val(db_s, 'G4'), 'date': get_val(db_s, 'F15'), 
                        'testers': f"{get_val(db_s, 'G13')}, {get_val(db_s, 'G14')}",
                        'vehicle': f"{get_val(db_s, 'G22')}, {get_val(db_s, 'G23')}",
                        'tire_ref': {'size': get_val(db_s, 'G17'), 'pattern': get_val(db_s, 'G18'), 'brand': get_val(db_s, 'G19'), 'marking': get_val(db_s, 'G20')},
                        'tire_sample': {
                            'size': get_unique_vals(db_s, 'H17~Q17'), 
                            'pattern': get_unique_vals(db_s, 'H18~Q18'), 
                            'brand': get_unique_vals(db_s, 'H19~Q19'), 
                            'marking': get_unique_vals(db_s, 'H20~Q20'),
                            'markings_list': markings
                        },
                        'pressure': f"전륜 {get_val(db_s, 'G30')}{get_val(db_s, 'G32')}/{get_val(db_s, 'G33')}Jx{wheel_size}, 후륜-{get_val(db_s, 'G31')}{get_val(db_s, 'G32')}/,{get_val(db_s, 'G34')}Jx{wheel_size}",
                        'temp_air': get_min_max(db_s, 'G39~Q39'), 'temp_road': get_min_max(db_s, 'G40~Q40')
                    })
                
                yield f"data: {json.dumps({'status': 'progress', 'percent': int(base_pct + file_pct * 0.4), 'message': f'{step} {name} 정량치(Report 시트) 추출 중...'})}\n\n"
                if rep_s:
                    day_data['table_3_4_1'] = get_table_data(rep_s, 'B22', 'O36')
                    day_data['table_3_4_2'] = get_table_data(rep_s, 'B38', 'O52')
                
                yield f"data: {json.dumps({'status': 'progress', 'percent': int(base_pct + file_pct * 0.6), 'message': f'{step} {name} 로드노이즈(RR 시트) 추출 중...'})}\n\n"
                if rr_s:
                    day_data['spectrum_rr'] = get_spectrum_data(rr_s, [1, 15, 29, 43])
                    
                yield f"data: {json.dumps({'status': 'progress', 'percent': int(base_pct + file_pct * 0.8), 'message': f'{step} {name} 패턴노이즈(SR 시트) 추출 중...'})}\n\n"
                if sr_s:
                    day_data['spectrum_sr'] = get_spectrum_data(sr_s, [1, 15, 29, 43])

                yield f"data: {json.dumps({'status': 'progress', 'percent': int(base_pct + file_pct * 0.95), 'message': f'{step} {name} 분석 마무리 중...'})}\n\n"

                res.append(day_data)
                wb.close(); os.remove(path)
            yield f"data: {json.dumps({'status': 'success', 'data': res, 'file_names': [f[1] for f in saved_files]})}\n\n"
        except Exception as e: 
            import traceback
            print(traceback.format_exc())
            yield f"data: {json.dumps({'status': 'error', 'message': str(e)})}\n\n"
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
