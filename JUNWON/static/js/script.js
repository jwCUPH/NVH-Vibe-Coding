document.addEventListener('DOMContentLoaded', () => {
    let rowCount = 1;
    let currentExtractedData = null;
    let currentFileNames = [];
    let charts = [];

    // --- Upload Section Logic ---
    window.addRow = function(btn) {
        rowCount++;
        const container = document.getElementById('upload-container');
        const newRow = document.createElement('div');
        newRow.className = 'upload-row';
        newRow.id = `row-${rowCount}`;
        newRow.innerHTML = `
            <span class="day-label">Day ${rowCount}</span>
            <input type="file" name="file-${rowCount}" onchange="validateFile(${rowCount})">
            <span class="status" id="status-${rowCount}"></span>
            <div class="controls">
                <button class="btn-plus" onclick="addRow(this)">+</button>
                <button class="btn-minus" onclick="removeRow(${rowCount})">-</button>
            </div>
        `;
        container.appendChild(newRow);
    };

    window.removeRow = function(id) {
        const row = document.getElementById(`row-${id}`);
        if (row) row.remove();
        reorderDays();
    };

    function reorderDays() {
        const rows = document.querySelectorAll('.upload-row');
        rowCount = rows.length;
        rows.forEach((row, index) => {
            row.querySelector('.day-label').textContent = `Day ${index + 1}`;
        });
    }

    window.validateFile = async function(id) {
        const input = document.querySelector(`input[name="file-${id}"]`);
        const statusSpan = document.getElementById(`status-${id}`);
        if (!input.files.length) return;
        const formData = new FormData();
        formData.append('file', input.files[0]);
        try {
            statusSpan.textContent = '...';
            const response = await fetch('/validate', { method: 'POST', body: formData });
            const data = await response.json();
            statusSpan.textContent = data.status;
            statusSpan.className = 'status ' + data.status.toLowerCase();
        } catch (error) {
            statusSpan.textContent = 'NG';
            statusSpan.className = 'status ng';
        }
    };

    // --- Auto Dash Logic for Content Editable ---
    window.handleAutoDash = function(e) {
        if (e.key === 'Enter') {
            e.preventDefault();
            const selection = window.getSelection();
            const range = selection.getRangeAt(0);
            
            const br = document.createElement('br');
            const textNode = document.createTextNode('- ');
            
            range.deleteContents();
            range.insertNode(br);
            range.setStartAfter(br);
            range.setEndAfter(br);
            range.insertNode(textNode);
            range.setStartAfter(textNode);
            range.setEndAfter(textNode);
            
            selection.removeAllRanges();
            selection.addRange(range);
        }
    };

    // --- Dynamic Row Addition (Outside) ---
    window.addDynamicRowOutside = function(containerId, prefix) {
        const container = document.getElementById(containerId);
        const div = document.createElement('div');
        div.className = 'dynamic-row-item';
        div.style.marginBottom = '2px';
        
        const index = container.children.length + 1;
        div.innerHTML = `
            <span class="row-num" style="display:inline-block; min-width: 35px;">&nbsp; ${prefix}${index})</span>
            <span class="editable-text" contenteditable="true" style="min-width: 300px;"></span>
            <button class="btn-remove-small" onclick="this.parentElement.remove(); updateNumbersOutside('${containerId}', '${prefix}')" style="margin-left: 5px; background: #e74c3c; color: white; border: none; cursor: pointer; padding: 2px 5px; font-size: 10px;">×</button>
        `;
        container.appendChild(div);
    };

    window.updateNumbersOutside = function(containerId, prefix) {
        const container = document.getElementById(containerId);
        Array.from(container.children).forEach((child, i) => {
            child.querySelector('.row-num').innerHTML = `&nbsp; ${prefix}${i + 1})`;
        });
    };

    // --- Extraction Logic ---
    window.extractVPR = async function() {
        const inputs = document.querySelectorAll('input[type="file"]');
        const formData = new FormData();
        let validFiles = 0;
        inputs.forEach(input => {
            if (input.files.length > 0) {
                const row = input.closest('.upload-row');
                const status = row.querySelector('.status').textContent;
                if (status === 'OK') {
                    formData.append('files', input.files[0]);
                    validFiles++;
                }
            }
        });
        if (validFiles === 0) { alert('정상적인(OK) 엑셀 파일을 업로드해주세요.'); return; }

        const btn = document.querySelector('.btn-extract');
        const overlay = document.getElementById('loading-overlay');
        const loadingText = document.getElementById('loading-text');
        const progressBar = document.getElementById('progress-bar');
        const progressPercent = document.getElementById('progress-percent');
        
        btn.disabled = true;
        overlay.style.display = 'flex';
        
        try {
            const response = await fetch('/extract', { method: 'POST', body: formData });
            const reader = response.body.getReader();
            const decoder = new TextDecoder();
            let buffer = '';
            while (true) {
                const { value, done } = await reader.read();
                if (done) break;
                buffer += decoder.decode(value, { stream: true });
                const lines = buffer.split('\n\n');
                buffer = lines.pop();
                for (const line of lines) {
                    if (line.startsWith('data: ')) {
                        const data = JSON.parse(line.substring(6));
                        if (data.status === 'progress') {
                            loadingText.textContent = data.message;
                            progressBar.style.width = data.percent + '%';
                            progressPercent.textContent = data.percent + '%';
                        } else if (data.status === 'success') {
                            displayManualInput(data.data, data.file_names);
                        } else if (data.status === 'error') { alert('에러: ' + data.message); }
                    }
                }
            }
        } catch (error) { alert('통신 에러: ' + error.message); }
        finally { btn.disabled = false; setTimeout(() => overlay.style.display = 'none', 500); }
    };

    function displayManualInput(data, fileNames) {
        currentExtractedData = data;
        currentFileNames = fileNames;
        const section = document.getElementById('manual-input-section');
        const form = document.getElementById('manual-form');
        section.style.display = 'block';
        
        const common = data[0];
        const multiDay = data.length > 1;
        
        let html = `
<div class="edit-controls">
    <button class="floating-btn" onclick="addExtraPlace()">2-4) 평가장소 추가</button>
    <button class="floating-btn" onclick="addDynamicRowOutside('m-analysis-container', '-')">3-5) 구조민감도 추가</button>
    <button class="floating-btn" onclick="addDynamicRowOutside('m-summary-container', '4-')">4. 요약 추가</button>
</div>

<div>
<p><b>■ 의뢰자 정보 :</b> ${common.client_info}</p>
<p><br></p>
<p><b>1. Project</b> : ${common.project}</p>
<p><br></p>
<p><b>2. Tire & Test Information</b></p>
<p>&nbsp; 2-1) 의뢰 번호 : ${common.req_no}</p>
<p>&nbsp; 2-2) 평가 일자 : ${common.date}</p>
<p>&nbsp; 2-3) 평가자 : ${common.testers}</p>
<p>&nbsp; 2-4) 평가 장소/노면 :</p>
<div style="padding-left: 20px;">
    <p>- Road Noise : <span class="editable-text" contenteditable="true" style="min-width: 300px;">@모형로 Rough Asphalt 10초(60kph, D drive)</span></p>
    <p>- Pattern Noise : <span class="editable-text" contenteditable="true" style="min-width: 300px;">@범용로 Smooth Asphalt 10초(80kph, D drive)</span></p>
    <div id="extra-places-container"></div>
</div>

<p>&nbsp; 2-5) 평가 타이어 :</p>
<p>&nbsp; &nbsp; &nbsp;- Ref. : 사이즈(${common.tire_ref.size}), 패턴(${common.tire_ref.pattern}), 브랜드(${common.tire_ref.brand}), 마킹(${common.tire_ref.marking})</p>
<p>&nbsp; &nbsp; &nbsp;- Sample : 사이즈(${common.tire_sample.size}), 패턴(${common.tire_sample.pattern}), 브랜드(${common.tire_sample.brand}), 마킹(${common.tire_sample.marking})</p>

<p>&nbsp; 2-6) 평가 공기압/휠 : ${common.pressure}</p>
<p>&nbsp; 2-7) 평가 차량 : ${common.vehicle}</p>
<p>&nbsp; 2-8) 평가 온도 :</p>
${data.map(d => `<p>&nbsp; &nbsp;2-8-${d.day}) Day${d.day} : 대기 ${d.temp_air}, 노면 ${d.temp_road}</p>`).join('')}
<p><br></p>

<p><b>3. Test Results</b></p>
<p>&nbsp; 3-1) 제조품질 성능검증결과</p>
    <table class="result-table">
        <tr>
            <th style="width: 25%;">Spec.</th>
            <th style="width: 15%;">공진음 수준</th>
            <th style="width: 15%;">진동 수준</th>
            <th style="width: 25%;">비고 (발생속도 및 기타)</th>
            <th rowspan="${(common.tire_sample.markings_list.length || 0) + 2}" class="note-cell">
                ※ 제조품질 성능수준<br>
                수준1 : 양호~경미 (거의 신경 쓰이지 않음)<br>
                수준2 : 중간 (인지 수준이나, 불쾌감 적음)<br>
                수준3 : 강함 (명확히 인지되며, 불쾌감 유발)
            </th>
        </tr>
        <tr>
            <td>Ref.</td>
            <td>${qualitySelect()}</td>
            <td>${qualitySelect()}</td>
            <td><span class="editable-text" contenteditable="true" style="width: 90%;">-</span></td>
        </tr>
        ${(common.tire_sample.markings_list || []).map(m => `
            <tr>
                <td><span class="editable-text" contenteditable="true" style="width: 90%;">${m}</span></td>
                <td>${qualitySelect()}</td>
                <td>${qualitySelect()}</td>
                <td><span class="editable-text" contenteditable="true" style="width: 90%;">-</span></td>
            </tr>
        `).join('')}
    </table>
<p><br></p>

<p>&nbsp; 3-2) 로드노이즈 <span style="color: blue; font-weight: bold;">[Target : 부밍 ${targetSelect('m-target-booming')}dB(A), 캐비티 ${targetSelect('m-target-cavity')}dB(A), 럼블 ${targetSelect('m-target-rumble')}dB(A)]</span></p>
<div class="editable-block" contenteditable="true" onkeydown="handleAutoDash(event)" style="margin-left: 15px;">- </div>
${data.map(d => `
    ${multiDay ? `<p>&nbsp; &nbsp; 3-2-${d.day}) Day${d.day}</p>` : ''}
    <div class="chart-grid">
        <div class="chart-wrapper"><canvas id="chart-rr-flw-${d.day}"></canvas></div>
        <div class="chart-wrapper"><canvas id="chart-rr-fli-${d.day}"></canvas></div>
        <div class="chart-wrapper"><canvas id="chart-rr-rcc-${d.day}"></canvas></div>
        <div class="chart-wrapper"><canvas id="chart-rr-rrw-${d.day}"></canvas></div>
    </div>
`).join('\n')}

<p>&nbsp; 3-3) 패턴노이즈 <span style="color: blue; font-weight: bold;">[Target : 패턴 ${targetSelect('m-target-pattern')}dB(A)]</span></p>
<div class="editable-block" contenteditable="true" onkeydown="handleAutoDash(event)" style="margin-left: 15px;">- </div>
${data.map(d => `
    ${multiDay ? `<p>&nbsp; &nbsp; 3-3-${d.day}) Day${d.day}</p>` : ''}
    <div class="chart-grid">
        <div class="chart-wrapper"><canvas id="chart-sr-flw-${d.day}"></canvas></div>
        <div class="chart-wrapper"><canvas id="chart-sr-fli-${d.day}"></canvas></div>
    </div>
`).join('\n')}

<p>&nbsp; 3-4) 실내소음 정량치 결과</p>
${data.map(d => `
    ${multiDay ? `<p>&nbsp; &nbsp; 3-4-${d.day}) Day${d.day}</p>` : ''}
    <div class="table-container">${renderReportTable(d.table_3_4_1)}</div>
    <div class="table-container">${renderReportTable(d.table_3_4_2)}</div>
`).join('\n')}

<p>&nbsp; 3-5) 추가 분석 : 구조 민감도 정리</p>
<div id="m-analysis-container"></div>
<p><br></p>

<p><b><span style="color: blue;">4. Test Summary</span></b></p>
<div id="m-summary-container">
    <div class="dynamic-row-item" style="margin-bottom: 2px;">
        <span class="row-num" style="display:inline-block; min-width: 35px;">&nbsp; 4-1)</span>
        <span class="editable-text" contenteditable="true" style="min-width: 300px;"></span>
    </div>
</div>
<p><br></p>

<p><b><span style="color: blue;">5. 첨부</span></b></p>
<p>&nbsp; 5-1) NVH 실차평가 보고서 ---------------------------------- ${data.length}부</p>
<div id="m-attach-container"></div>

</div>
        `;
        form.innerHTML = html;
        
        setTimeout(() => {
            data.forEach(d => {
                if (d.spectrum_rr) {
                    renderChart(`chart-rr-flw-${d.day}`, "FLW (Rough Road)", d.spectrum_rr['FLW']);
                    renderChart(`chart-rr-fli-${d.day}`, "FLI (Rough Road)", d.spectrum_rr['FLI']);
                    renderChart(`chart-rr-rcc-${d.day}`, "RCC (Rough Road)", d.spectrum_rr['RCC']);
                    renderChart(`chart-rr-rrw-${d.day}`, "RRW (Rough Road)", d.spectrum_rr['RRW']);
                }
                if (d.spectrum_sr) {
                    renderChart(`chart-sr-flw-${d.day}`, "FLW (Smooth Road)", d.spectrum_sr['FLW']);
                    renderChart(`chart-sr-fli-${d.day}`, "FLI (Smooth Road)", d.spectrum_sr['FLI']);
                }
            });
        }, 100);

        section.scrollIntoView({ behavior: 'smooth' });
    }

    function qualitySelect() {
        return `<select style="padding: 2px; border: none; outline: none; background: transparent;"><option value="1">1</option><option value="2">2</option><option value="3">3</option></select>`;
    }

    function targetSelect(id) {
        return `<select style="border:none; color: blue; font-weight:bold; background: transparent; outline: none;" id="${id}">
            <option value="-1">-1</option><option value="-0.5">-0.5</option>
            <option value="동등" selected>동등</option><option value="+0.5">+0.5</option><option value="+1">+1</option>
        </select>`;
    }

    function renderReportTable(data) {
        if (!data || data.length < 5) return '';

        // Header Structure (Rows 0-3 in data are headers)
        // Row 0: ['FLI', null, 'Rough Road(RR)', null, ...]
        // Row 1: [null, null, '60kph', null, ...]
        // Row 2: [null, null, '20~500Hz', null, '20~178Hz', ...]
        // Row 3: [null, null, 'OA', null, 'Booming', ...]

        let html = '<table class="report-data-table" style="width: 100%; border-collapse: collapse; margin-top: 10px; border: 1px solid #000; font-size: 8.5pt; table-layout: fixed;">';
        
        // Find header info
        const category = data[0][0]; // FLI or RCC
        
        // Render Header with fixed widths to ensure equality
        // Category(70px) + 6 categories (each 15% of the rest)
        html += `<tr style="background:#eee;"><th rowspan="3" style="border:1px solid #000; width: 70px;">${category}</th><th colspan="10" style="border:1px solid #000;">Road Noise (Rough Road)</th><th colspan="2" style="border:1px solid #000; width: 15%;">Pattern Noise (Smooth Road)</th></tr>`;
        html += `<tr style="background:#eee;"><th colspan="2" style="border:1px solid #000; width: 15.5%;">20~500Hz</th><th colspan="2" style="border:1px solid #000; width: 15.5%;">20~178Hz</th><th colspan="2" style="border:1px solid #000; width: 15.5%;">178~224Hz</th><th colspan="2" style="border:1px solid #000; width: 15.5%;">178~224Hz</th><th colspan="2" style="border:1px solid #000; width: 15.5%;">224~500Hz</th><th colspan="2" style="border:1px solid #000; width: 15.5%;">500~4000Hz</th></tr>`;
        html += `<tr style="background:#eee;"><th colspan="2" style="border:1px solid #000;">OA</th><th colspan="2" style="border:1px solid #000;">Booming</th><th colspan="2" style="border:1px solid #000;">Cavity Peak</th><th colspan="2" style="border:1px solid #000;">Cavity RMS</th><th colspan="2" style="border:1px solid #000;">Rumble</th><th colspan="2" style="border:1px solid #000;">OA</th></tr>`;

        // Identify REF row for delta calculation
        const refRowIdx = data.findIndex(r => r[0] === 'REF');
        if (refRowIdx === -1) return 'REF 데이터를 찾을 수 없습니다.';
        const refRow = data[refRowIdx];

        // Process Data Rows (starting from row index 4 typically, but we use REF identification)
        data.forEach((row, idx) => {
            if (idx < 4) return; // Skip headers
            if (!row[0]) return; // Skip empty rows

            html += '<tr>';
            html += `<td style="border:1px solid #000; font-weight:bold;">${row[0]}</td>`; // Spec Name

            // Column indices in 'Report' sheet for OA, Booming, etc.
            // Based on inspection: B=0, D=2(OA), F=4(Booming), H=6(Cav Peak), J=8(Cav RMS), L=10(Rumble), N=12(Pattern OA)
            const dataCols = [2, 4, 6, 8, 10, 12];
            
            dataCols.forEach(colIdx => {
                const val = row[colIdx];
                const formattedVal = (typeof val === 'number') ? val.toFixed(1) : (val || '-');
                html += `<td style="border:1px solid #000;">${formattedVal}</td>`;

                // Delta Calculation
                if (row[0] === 'REF') {
                    html += `<td style="border:1px solid #000; color:#888; font-size:8pt;">-</td>`;
                } else {
                    const refVal = refRow[colIdx];
                    if (typeof val === 'number' && typeof refVal === 'number') {
                        const delta = val - refVal;
                        let style = '';
                        if (delta >= 0.3) style = 'background-color: #ffcccc; color: #cc0000; font-weight:bold;';
                        else if (delta <= -0.3) style = 'background-color: #cce5ff; color: #0000cc; font-weight:bold;';
                        
                        const deltaStr = delta > 0 ? `+${delta.toFixed(1)}` : delta.toFixed(1);
                        html += `<td style="border:1px solid #000; ${style}">${deltaStr}</td>`;
                    } else {
                        html += `<td style="border:1px solid #000;">-</td>`;
                    }
                }
            });
            html += '</tr>';
        });

        html += '</table>';
        return html;
    }

    function renderChart(canvasId, title, specData) {
        const ctx = document.getElementById(canvasId);
        if (!ctx || !specData) return;
        const datasets = Object.keys(specData.samples).map((sname, idx) => ({
            label: sname,
            data: specData.freqs.map((f, i) => ({ x: f, y: 10 * Math.log10(specData.samples[sname][i] + 1e-12) })),
            borderColor: `hsl(${idx * 70}, 70%, 50%)`,
            borderWidth: 1.5, pointRadius: 0, fill: false
        }));
        const chart = new Chart(ctx, {
            type: 'line', data: { datasets },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: {
                    title: { display: true, text: title, font: { size: 10 } },
                    legend: {
                        labels: { boxWidth: 10, font: { size: 9 } },
                        onClick: (e, item, legend) => {
                            const index = item.datasetIndex;
                            legend.chart.data.datasets.forEach((ds, i) => ds.borderWidth = (i === index) ? 3 : 1);
                            legend.chart.update();
                        }
                    },
                    zoom: { zoom: { wheel: { enabled: true }, mode: 'xy' }, pan: { enabled: true, mode: 'xy' } }
                },
                scales: { x: { type: 'linear', ticks: { font: { size: 9 } } }, y: { type: 'linear', ticks: { font: { size: 9 } } } }
            }
        });
        charts.push(chart);
    }

    window.addExtraPlace = function() {
        const container = document.getElementById('extra-places-container');
        const div = document.createElement('div');
        div.className = 'dynamic-row-item';
        div.innerHTML = `
            <p>- <span class="editable-text" contenteditable="true" style="min-width: 100px;">평가장소</span> : 
               <span class="editable-text" contenteditable="true" style="min-width: 200px;">노면</span>
               <button class="btn-remove-small" onclick="this.parentElement.parentElement.remove()" style="background: #e74c3c; color: white; border: none; cursor: pointer; padding: 2px 5px; font-size: 10px;">×</button>
            </p>
        `;
        container.appendChild(div);
    };

    window.saveFinal = async function() {
        if (!currentExtractedData) return;
        const formClone = document.getElementById('manual-form').cloneNode(true);
        
        // Remove floating controls and internal buttons
        formClone.querySelector('.edit-controls')?.remove();
        formClone.querySelectorAll('button').forEach(b => b.remove());

        // Replace contenteditables and selects with plain spans/text
        formClone.querySelectorAll('[contenteditable="true"]').forEach(el => {
            el.removeAttribute('contenteditable');
            el.style.border = 'none';
            el.style.outline = 'none';
        });

        formClone.querySelectorAll('select').forEach(el => {
            const span = document.createElement('span');
            span.textContent = el.options[el.selectedIndex].text;
            span.style.color = el.style.color;
            span.style.fontWeight = el.style.fontWeight;
            el.parentNode.replaceChild(span, el);
        });

        // Charts to images
        const canvases = document.querySelectorAll('canvas');
        const canvasImages = Array.from(canvases).map(c => c.toDataURL());
        formClone.querySelectorAll('canvas').forEach((c, i) => {
            const img = document.createElement('img');
            img.src = canvasImages[i];
            img.style.width = '100%';
            c.parentNode.replaceChild(img, c);
        });

        const finalHtml = `
            <html><head><style>p { font-size : 9pt; font-family : '맑은 고딕', sans-serif; margin:0px; line-height:180%; }</style></head>
            <body style="font-size: 9pt; font-family: '맑은 고딕', sans-serif; background: #fff; padding: 40px;">
                ${formClone.innerHTML}
            </body></html>
        `;

        const payload = {
            req_no: currentExtractedData[0].req_no,
            project: currentExtractedData[0].project,
            date: currentExtractedData[0].date,
            html_content: finalHtml,
            raw_data: { extracted: currentExtractedData, files: currentFileNames }
        };

        try {
            const response = await fetch('/save', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            if ((await response.json()).status === 'success') {
                alert('리포트가 성공적으로 저장되었습니다.');
                location.reload();
            }
        } catch (e) { alert('저장 중 오류 발생'); }
    };

    window.toggleReportList = async function() {
        const section = document.getElementById('report-list-section');
        if (section.style.display === 'none') {
            const response = await fetch('/reports');
            const data = await response.json();
            document.getElementById('report-table-body').innerHTML = data.map(r => `
                <tr>
                    <td style="padding: 10px; border: 1px solid #ccc;">${r.req_no}</td>
                    <td style="padding: 10px; border: 1px solid #ccc;">${r.project}</td>
                    <td style="padding: 10px; border: 1px solid #ccc;">${r.created_at}</td>
                    <td style="padding: 10px; border: 1px solid #ccc;"><button onclick="viewReport(${r.id})">보기</button></td>
                </tr>
            `).join('');
            section.style.display = 'block';
        } else { section.style.display = 'none'; }
    };

    window.viewReport = async function(id) {
        const response = await fetch(`/report/${id}`);
        const data = await response.json();
        document.getElementById('upload-section-main').style.display = 'none';
        document.getElementById('report-list-section').style.display = 'none';
        const ms = document.getElementById('manual-input-section');
        ms.style.display = 'block';
        document.getElementById('manual-form').innerHTML = data.content;
        document.querySelector('.save-btn').style.display = 'none';
        const b = document.createElement('button');
        b.textContent = '돌아가기'; b.className = 'save-btn'; b.style.background = '#2c3e50';
        b.onclick = () => location.reload();
        ms.appendChild(b);
    };
});
