document.addEventListener('DOMContentLoaded', () => {
    // Initial row count
    let rowCount = 1;

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
            const currentDay = index + 1;
            row.querySelector('.day-label').textContent = `Day ${currentDay}`;
            // We don't necessarily need to change the ID or names, 
            // but it's cleaner for display.
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
            const response = await fetch('/validate', {
                method: 'POST',
                body: formData
            });
            const data = await response.json();
            
            if (data.status === 'OK') {
                statusSpan.textContent = 'OK';
                statusSpan.className = 'status ok';
            } else {
                statusSpan.textContent = 'NG';
                statusSpan.className = 'status ng';
            }
        } catch (error) {
            statusSpan.textContent = 'NG';
            statusSpan.className = 'status ng';
        }
    };

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

        if (validFiles === 0) {
            alert('최소 하나의 정상적인(OK) 엑셀 파일을 업로드해주세요.');
            return;
        }

        const btn = document.querySelector('.btn-extract');
        const overlay = document.getElementById('loading-overlay');
        const loadingText = document.getElementById('loading-text');
        const progressBar = document.getElementById('progress-bar');
        const progressPercent = document.getElementById('progress-percent');
        
        btn.disabled = true;
        overlay.style.display = 'flex';
        loadingText.textContent = '분석 준비 중...';
        progressBar.style.width = '0%';
        progressPercent.textContent = '0%';

        try {
            const response = await fetch('/extract', {
                method: 'POST',
                body: formData
            });

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
                            if (data.percent !== undefined) {
                                progressBar.style.width = data.percent + '%';
                                progressPercent.textContent = data.percent + '%';
                            }
                        } else if (data.status === 'success') {
                            displayManualInput(data.data, data.file_names);
                        } else if (data.status === 'error') {
                            alert('에러: ' + data.message);
                        }
                    }
                }
            }
        } catch (error) {
            alert('통신 에러: ' + error.message);
        } finally {
            btn.disabled = false;
            setTimeout(() => {
                overlay.style.display = 'none';
            }, 500);
        }
    };

    let currentExtractedData = null;
    let currentFileNames = [];

    function displayManualInput(data, fileNames) {
        currentExtractedData = data;
        currentFileNames = fileNames;
        const section = document.getElementById('manual-input-section');
        const form = document.getElementById('manual-form');
        const uploadSection = document.getElementById('upload-section-main');
        
        section.style.display = 'block';
        uploadSection.style.display = 'none'; // Hide upload while editing
        
        const common = data[0]; 
        
        let html = `
■ 의뢰자 정보 : ${common.client_info}

1. Project : ${common.project}
2. Tire & Test Information
  2-1) 의뢰 번호 : ${common.req_no}
  2-2) 평가 일자 : ${common.date}
  2-3) 평가자 : ${common.testers}
  2-4) 평가 장소/노면 : <input type="text" class="input-field" id="m-location" style="width: 300px;">
  2-5) 평가 타이어
    2-5-1) Ref. : 사이즈(${common.tire_ref.size}), 패턴(${common.tire_ref.pattern}), 브랜드(${common.tire_ref.brand}), 마킹(${common.tire_ref.marking})
    2-5-2) Sample : 사이즈(${common.tire_sample.size}), 패턴(${common.tire_sample.pattern}), 브랜드(${common.tire_sample.brand}), 마킹(${common.tire_sample.marking})

  2-6) 평가 공기압/휠 : <input type="text" class="input-field" id="m-tire-spec" style="width: 300px;">
  2-7) 평가 차량 : ${common.vehicle}
  2-8) 평가 온도 :
${data.map(d => `   2-8-${d.day}) Day${d.day} : 대기 ${d.temp_air}, 노면 ${d.temp_road}`).join('\n')}

3. Test Results
  3-1) 제조품질 성능검증결과

3-2) 로드노이즈
 - <textarea class="input-area" id="m-roadnoise" rows="2"></textarea>

3-3) 패턴노이즈
 - <textarea class="input-area" id="m-patternnoise" rows="2"></textarea>

3-4) 실내소음 정량치 결과
${data.map(d => `  3-4-${d.day}) Day${d.day}
`).join('\n')}

3-5) 추가 분석 : 구조 민감도 정리
  - <textarea class="input-area" id="m-analysis" rows="2"></textarea>

4. Test Summary
  4-1) <textarea class="input-area" id="m-summary-1" rows="2"></textarea>
  4-2) <textarea class="input-area" id="m-summary-2" rows="2"></textarea>

5. 첨부
  5-1) NVH 실차평가 보고서 ---- ${data.length}개 파일
${fileNames.map((name, i) => `  5-2) ${name}`).join('\n')}
        `;
        form.innerHTML = html;
        section.scrollIntoView({ behavior: 'smooth' });
    }

    window.saveFinal = async function() {
        if (!currentExtractedData) return;

        const common = currentExtractedData[0];
        
        // Collect all raw inputs
        const raw_inputs = {
            location: document.getElementById('m-location').value,
            tire_spec: document.getElementById('m-tire-spec').value,
            roadnoise: document.getElementById('m-roadnoise').value,
            patternnoise: document.getElementById('m-patternnoise').value,
            analysis: document.getElementById('m-analysis').value,
            summary_1: document.getElementById('m-summary-1').value,
            summary_2: document.getElementById('m-summary-2').value
        };

        // Generate clean HTML for viewing (replacing inputs with text)
        const formClone = document.getElementById('manual-form').cloneNode(true);
        formClone.querySelectorAll('.input-field, .input-area').forEach(input => {
            const span = document.createElement('span');
            span.textContent = input.value || '-';
            span.style.fontWeight = 'bold';
            span.style.textDecoration = 'underline';
            input.parentNode.replaceChild(span, input);
        });

        const finalHtml = `
            <div class="paper-container" style="background: #fff; padding: 40px; box-shadow: none; border: 1px solid #000; font-family: 'Malgun Gothic';">
                <div style="white-space: pre-wrap; font-size: 15px; line-height: 1.8;">${formClone.innerHTML}</div>
            </div>
        `;

        const payload = {
            req_no: common.req_no,
            project: common.project,
            date: common.date,
            html_content: finalHtml,
            raw_data: {
                extracted: currentExtractedData,
                files: currentFileNames,
                inputs: raw_inputs
            }
        };

        try {
            const response = await fetch('/save', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            const result = await response.json();
            if (result.status === 'success') {
                alert('리포트가 성공적으로 저장되었습니다.');
                location.reload(); 
            } else {
                alert('저장 실패: ' + result.message);
            }
        } catch (e) {
            alert('저장 중 오류 발생');
        }
    };

    window.viewReport = async function(id) {
        const response = await fetch(`/report/${id}`);
        const data = await response.json();
        
        const uploadSection = document.getElementById('upload-section-main');
        const manualSection = document.getElementById('manual-input-section');
        const reportList = document.getElementById('report-list-section');
        const form = document.getElementById('manual-form');
        const saveBtn = document.querySelector('.save-btn');
        
        uploadSection.style.display = 'none';
        reportList.style.display = 'none';
        manualSection.style.display = 'block';
        manualSection.querySelector('h2').textContent = '저장된 보고서 보기';
        
        form.innerHTML = data.content;
        saveBtn.style.display = 'none';
        
        // Buttons container
        const btnContainer = document.createElement('div');
        btnContainer.style.textAlign = 'center';
        btnContainer.style.marginTop = '20px';

        // Edit Button (if raw data exists)
        if (data.raw_data) {
            const editBtn = document.createElement('button');
            editBtn.textContent = '수정하기';
            editBtn.className = 'save-btn';
            editBtn.style.display = 'inline-block';
            editBtn.style.width = '150px';
            editBtn.style.margin = '0 10px';
            editBtn.onclick = () => {
                displayManualInput(data.raw_data.extracted, data.raw_data.files);
                // Fill in the inputs
                document.getElementById('m-location').value = data.raw_data.inputs.location;
                document.getElementById('m-tire-spec').value = data.raw_data.inputs.tire_spec;
                document.getElementById('m-roadnoise').value = data.raw_data.inputs.roadnoise;
                document.getElementById('m-patternnoise').value = data.raw_data.inputs.patternnoise;
                document.getElementById('m-analysis').value = data.raw_data.inputs.analysis;
                document.getElementById('m-summary-1').value = data.raw_data.inputs.summary_1;
                document.getElementById('m-summary-2').value = data.raw_data.inputs.summary_2;
                saveBtn.style.display = 'block';
                manualSection.querySelector('h2').textContent = '보고서 수정';
                btnContainer.remove();
            };
            btnContainer.appendChild(editBtn);
        }

        const backBtn = document.createElement('button');
        backBtn.textContent = '메인으로 돌아가기';
        backBtn.className = 'save-btn';
        backBtn.style.display = 'inline-block';
        backBtn.style.width = '200px';
        backBtn.style.margin = '0 10px';
        backBtn.style.background = '#2c3e50';
        backBtn.onclick = () => location.reload();
        btnContainer.appendChild(backBtn);
        
        manualSection.appendChild(btnContainer);
    };
});
