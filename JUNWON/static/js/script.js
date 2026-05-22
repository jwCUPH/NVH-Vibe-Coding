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

    // --- Auto Dash Logic ---
    window.handleAutoDash = function(e) {
        if (e.key === 'Enter') {
            e.preventDefault();
            const start = e.target.selectionStart;
            const end = e.target.selectionEnd;
            const value = e.target.value;
            e.target.value = value.substring(0, start) + "\n- " + value.substring(end);
            e.target.selectionStart = e.target.selectionEnd = start + 3;
        }
    };

    // --- Dynamic Row Addition (Outside) ---
    window.addDynamicRowOutside = function(containerId, prefix) {
        const container = document.getElementById(containerId);
        const div = document.createElement('div');
        div.className = 'dynamic-row-item';
        div.style.display = 'flex';
        div.style.alignItems = 'center';
        
        const index = container.children.length + 1;
        div.innerHTML = `
            <span class="row-num" style="min-width: 35px;">${prefix}${index})</span>
            <textarea class="input-area-dynamic" style="flex: 1; margin: 2px 0;" rows="1"></textarea>
            <button class="btn-remove-small" onclick="this.parentElement.remove(); updateNumbersOutside('${containerId}', '${prefix}')" style="margin-left: 5px; background: #e74c3c; color: white; border: none; cursor: pointer;">×</button>
        `;
        container.appendChild(div);
    };

    window.updateNumbersOutside = function(containerId, prefix) {
        const container = document.getElementById(containerId);
        Array.from(container.children).forEach((child, i) => {
            child.querySelector('.row-num').textContent = `${prefix}${i + 1})`;
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
        
        // Setup initial structure with floating controls
        let html = `
<div class="edit-controls">
    <button class="floating-btn" onclick="addExtraPlace()">2-4) 평가장소 추가</button>
    <button class="floating-btn" onclick="addDynamicRowOutside('m-analysis-container', '-')">3-5) 구조민감도 추가</button>
    <button class="floating-btn" onclick="addDynamicRowOutside('m-summary-container', '4-')">4. 요약 추가</button>
</div>

<div style="white-space: pre-wrap; line-height: 1.6; font-size: 14px;">
■ 의뢰자 정보 : ${common.client_info}

1. Project : ${common.project}
2. Tire & Test Information
  2-1) 의뢰 번호 : ${common.req_no}
  2-2) 평가 일자 : ${common.date}
  2-3) 평가자 : ${common.testers}
  2-4) 평가 장소/노면 :
    - Road Noise : <input type="text" class="input-field" id="m-place-road" value="@모형로 Rough Asphalt 10초(60kph, D drive)" style="width: 400px;">
    - Pattern Noise : <input type="text" class="input-field" id="m-place-pattern" value="@범용로 Smooth Asphalt 10초(80kph, D drive)" style="width: 400px;">
    <div id="extra-places-container"></div>

  2-5) 평가 타이어
    2-5-1) Ref. : 사이즈(${common.tire_ref.size}), 패턴(${common.tire_ref.pattern}), 브랜드(${common.tire_ref.brand}), 마킹(${common.tire_ref.marking})
    2-5-2) Sample : 사이즈(${common.tire_sample.size}), 패턴(${common.tire_sample.pattern}), 브랜드(${common.tire_sample.brand}), 마킹(${common.tire_sample.marking})

  2-6) 평가 공기압/휠 : ${common.pressure}
  2-7) 평가 차량 : ${common.vehicle}
  2-8) 평가 온도 :
${data.map(d => `   2-8-${d.day}) Day${d.day} : 대기 ${d.temp_air}, 노면 ${d.temp_road}`).join('\n')}

3. Test Results
  3-1) 제조품질 성능검증결과
    <table class="result-table">
        <tr>
            <th style="width: 25%;">Spec.</th>
            <th style="width: 15%;">공진음 수준</th>
            <th style="width: 15%;">진동 수준</th>
            <th style="width: 25%;">비고 (발생속도 및 기타)</th>
            <th rowspan="${(common.tire_sample.markings_list.length || 0) + 2}" class="note-cell">
                ※ 제조품질 성능수준<br><br>
                수준1 : 양호~경미 (거의 신경 쓰이지 않음)<br>
                수준2 : 중간 (인지 수준이나, 불쾌감 적음)<br>
                수준3 : 강함 (명확히 인지되며, 불쾌감 유발)
            </th>
        </tr>
        <tr>
            <td>Ref.</td>
            <td>${qualitySelect()}</td>
            <td>${qualitySelect()}</td>
            <td><input type="text" value="-" style="width: 90%;"></td>
        </tr>
        ${(common.tire_sample.markings_list || []).map(m => `
            <tr>
                <td><input type="text" value="${m}" style="width: 90%;"></td>
                <td>${qualitySelect()}</td>
                <td>${qualitySelect()}</td>
                <td><input type="text" value="-" style="width: 90%;"></td>
            </tr>
        `).join('')}
    </table>

  3-2) 로드노이즈 <span style="color: blue;">[Target : 부밍 ${targetSelect('m-target-booming')}dB(A), 캐비티 ${targetSelect('m-target-cavity')}dB(A), 럼블 ${targetSelect('m-target-rumble')}dB(A)]</span>
    - <textarea class="input-area" id="m-roadnoise-text" rows="2" onkeydown="handleAutoDash(event)">- </textarea>
${data.map(d => `
    ${multiDay ? `    3-2-${d.day}) Day${d.day}` : ''}
    <div class="chart-grid">
        <div class="chart-wrapper"><canvas id="chart-rr-flw-${d.day}"></canvas></div>
        <div class="chart-wrapper"><canvas id="chart-rr-fli-${d.day}"></canvas></div>
        <div class="chart-wrapper"><canvas id="chart-rr-rcc-${d.day}"></canvas></div>
        <div class="chart-wrapper"><canvas id="chart-rr-rrw-${d.day}"></canvas></div>
    </div>
`).join('\n')}

  3-3) 패턴노이즈 <span style="color: blue;">[Target : 패턴 ${targetSelect('m-target-pattern')}dB(A)]</span>
    - <textarea class="input-area" id="m-patternnoise-text" rows="2" onkeydown="handleAutoDash(event)">- </textarea>
${data.map(d => `
    ${multiDay ? `    3-3-${d.day}) Day${d.day}` : ''}
    <div class="chart-grid">
        <div class="chart-wrapper"><canvas id="chart-sr-flw-${d.day}"></canvas></div>
        <div class="chart-wrapper"><canvas id="chart-sr-fli-${d.day}"></canvas></div>
    </div>
`).join('\n')}

  3-4) 실내소음 정량치 결과
${data.map(d => `
    ${multiDay ? `    3-4-${d.day}) Day${d.day}` : ''}
    <div class="table-container">${renderReportTable(d.table_3_4_1)}</div>
    <div class="table-container">${renderReportTable(d.table_3_4_2)}</div>
`).join('\n')}

  3-5) 추가 분석 : 구조 민감도 정리
    <div id="m-analysis-container"></div>

  4. Test Summary
    <div id="m-summary-container">
        <div class="dynamic-row-item" style="display: flex; align-items: center;">
            <span class="row-num" style="min-width: 35px;">4-1)</span>
            <textarea class="input-area-dynamic" style="flex: 1; margin: 2px 0;" rows="1"></textarea>
        </div>
    </div>

  5. 첨부
    5-1) NVH 실차평가 보고서 ---- ${data.length}개 파일
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
        return `<select style="padding: 2px;"><option value="1">1</option><option value="2">2</option><option value="3">3</option></select>`;
    }

    function targetSelect(id) {
        return `<select style="border:none; color: blue; font-weight:bold;" id="${id}">
            <option value="-1">-1</option><option value="-0.5">-0.5</option>
            <option value="동등" selected>동등</option><option value="+0.5">+0.5</option><option value="+1">+1</option>
        </select>`;
    }

    function renderReportTable(data) {
        if (!data || data.length === 0) return '';
        let html = '<table class="report-data-table">';
        data.forEach(row => {
            html += '<tr>';
            row.forEach(cell => html += `<td>${cell !== null ? cell : ''}</td>`);
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
                    title: { display: true, text: title },
                    legend: {
                        onClick: (e, item, legend) => {
                            const index = item.datasetIndex;
                            legend.chart.data.datasets.forEach((ds, i) => ds.borderWidth = (i === index) ? 3 : 1);
                            legend.chart.update();
                        }
                    },
                    zoom: { zoom: { wheel: { enabled: true }, mode: 'xy' }, pan: { enabled: true, mode: 'xy' } }
                },
                scales: { x: { type: 'linear' }, y: { type: 'linear' } }
            }
        });
        charts.push(chart);
    }

    window.addExtraPlace = function() {
        const container = document.getElementById('extra-places-container');
        const div = document.createElement('div');
        div.className = 'dynamic-row-item';
        div.style.display = 'flex'; div.style.marginBottom = '2px';
        div.innerHTML = `
            <span style="min-width: 120px;">- 평가장소 : 노면 :</span>
            <input type="text" class="input-field" style="flex: 1; border:none; border-bottom:1px solid #ccc;">
            <button class="btn-remove-small" onclick="this.parentElement.remove()" style="background: #e74c3c; color: white; border: none; cursor: pointer;">×</button>
        `;
        container.appendChild(div);
    };

    window.saveFinal = async function() {
        if (!currentExtractedData) return;
        const formClone = document.getElementById('manual-form').cloneNode(true);
        
        // Remove floating controls and internal buttons
        formClone.querySelector('.edit-controls')?.remove();
        formClone.querySelectorAll('button').forEach(b => b.remove());

        // Replace inputs/selects with spans
        formClone.querySelectorAll('input, textarea, select').forEach(el => {
            const span = document.createElement('span');
            span.textContent = el.value || (el.options ? el.options[el.selectedIndex].text : '-');
            if (el.tagName === 'TEXTAREA') {
                span.style.whiteSpace = 'pre-wrap';
                span.style.display = 'block';
            }
            span.style.fontWeight = 'bold';
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
            <div style="background: #fff; padding: 40px; border: 1px solid #000; font-family: 'Malgun Gothic'; width: 210mm; margin: 0 auto;">
                <div style="white-space: pre-wrap; font-size: 14px; line-height: 1.6;">${formClone.innerHTML}</div>
            </div>
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
