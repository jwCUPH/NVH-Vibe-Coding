from flask import Flask, render_template, jsonify
import random
from datetime import datetime, timedelta

app = Flask(__name__)

# --- Mock Data Generation ---
def generate_mock_data():
    vehicles = ["CN8", "NX5", "SX3"]
    sizes = ["17인치", "18인치", "19인치"]
    rounds = ["1s", "2s", "3s"]
    
    data = {}
    
    for v in vehicles:
        data[v] = {}
        for s in sizes:
            data[v][s] = {}
            for r in rounds:
                date = (datetime.now() - timedelta(days=random.randint(0, 365))).strftime("%Y-%m-%d")
                num_specs = random.randint(4, 8)
                spec_prefix = f"E{r[0]}"
                specs = [f"{spec_prefix}{chr(65+i)}" for i in range(num_specs)]
                
                struct_data = []
                for i, spec in enumerate(specs):
                    struct_data.append({
                        "Spec": spec,
                        "C/C 구조": random.choice(["A-Type", "B-Type", "C-Type"]),
                        "Belt 구조": random.choice(["Steel", "Hybrid", "Nylon"]),
                        "Belt Angle": random.choice([20, 22, 24, 26]),
                        "B/F Height": random.choice([15, 18, 20, 22]),
                        "보강벨트 재질": random.choice(["Nylon", "Aramid", "PET"]),
                        "보강벨트 구조": random.choice(["Full Cap", "Edge Cover", "Spiral"])
                    })
                
                nvh_results = []
                for spec in specs:
                    nvh_results.append({
                        "Spec": spec,
                        "Booming": round(random.uniform(50, 70), 1),
                        "Cavity Peak": round(random.uniform(50, 70), 1),
                        "Cavity": round(random.uniform(50, 70), 1),
                        "Rumble": round(random.uniform(50, 70), 1),
                        "Pattern Noise": round(random.uniform(50, 70), 1)
                    })
                
                comp_results = []
                for spec in specs:
                    comp_results.append({
                        "Spec": spec,
                        "ABN": round(random.uniform(20, 40), 1),
                        "SBN": round(random.uniform(15, 30), 1),
                        "Modal": round(random.uniform(100, 200), 1),
                        "Cleat Impact": round(random.uniform(0.5, 1.5), 2),
                        "F/T": round(random.uniform(80, 120), 1)
                    })
                
                data[v][s][r] = {
                    "Date": date,
                    "Specs": specs,
                    "Structural": struct_data,
                    "NVH": nvh_results,
                    "Component": comp_results
                }
    return data

MOCK_DATA = generate_mock_data()

@app.route('/')
def index():
    return render_template('index.html', data=MOCK_DATA)

@app.route('/api/data/<v>/<s>/<r>')
def get_data(v, s, r):
    return jsonify(MOCK_DATA.get(v, {}).get(s, {}).get(r, {}))

if __name__ == '__main__':
    app.run(debug=True)
