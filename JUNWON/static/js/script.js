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
        
        btn.disabled = true;
        overlay.style.display = 'flex';
        loadingText.textContent = '분석 준비 중...';

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
                        } else if (data.status === 'success') {
                            displayManualInput(data.data);
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
            overlay.style.display = 'none';
        }
    };

    function displayManualInput(data, isMultiDay) {
        const section = document.getElementById('manual-input-section');
        const form = document.getElementById('manual-form');
        section.style.display = 'block';
        
        const common = data[0]; // Day 1 data for common fields
        
        let html = `
            <div class="input-group">
                <h3>1. 기본 정보 (자동 추출)</h3>
                <p><strong>의뢰 번호:</strong> ${common.req_no}</p>
                <p><strong>평가 일자:</strong> ${common.date}</p>
                <p><strong>평가자:</strong> ${common.testers}</p>
                <p><strong>평가 차량:</strong> ${common.vehicle}</p>
                <p><strong>Ref. 타이어:</strong> ${common.tire_ref.brand} / ${common.tire_ref.pattern} / ${common.tire_ref.size} / ${common.tire_ref.marking}</p>
            </div>
            <hr>
            <h3>2. 수동 기입 및 환경 정보</h3>
            <div class="input-group">
                <label>2-4) 평가 장소/노면 (사용자 입력):</label>
                <input type="text" id="m-location" placeholder="예: 남양연구소 고속주회로">
            </div>
            <div class="input-group">
                <label>2-5-2) Sample 타이어 정보 (사용자 입력):</label>
                <textarea id="m-sample-tire" rows="2"></textarea>
            </div>
            <div class="input-group">
                <label>2-8) 평가 온도 (자동 추출):</label>
                <div id="temp-display" style="background: #f9f9f9; padding: 10px; border-radius: 4px;">
                    ${data.map(d => `
                        <p><strong>Day ${d.day}:</strong> 대기 ${d.temp_air}, 노면 ${d.temp_road}</p>
                    `).join('')}
                </div>
            </div>
            <hr>
            <h3>3. 평가 결과 (수동 입력)</h3>
            <div class="input-group">
                <label>3-2) 로드노이즈 Comment:</label>
                <textarea id="m-roadnoise" rows="3"></textarea>
            </div>
            <div class="input-group">
                <label>3-3) 패턴노이즈 Comment:</label>
                <textarea id="m-patternnoise" rows="3"></textarea>
            </div>
            <div class="input-group">
                <label>3-4) 실내소음 정량치 결과 (설명):</label>
                <textarea id="m-quant-data" rows="2" placeholder="정량치 Table에 대한 설명 기입"></textarea>
            </div>
            <div class="input-group">
                <label>4. Test Summary:</label>
                <textarea id="m-summary" rows="4"></textarea>
            </div>
            <button class="save-btn" onclick="saveFinal()">최종 HTML 파일로 변환 (저장)</button>
        `;
        form.innerHTML = html;
        section.scrollIntoView({ behavior: 'smooth' });
    }

    window.saveFinal = function() {
        alert('3단계: HTML 생성 및 DB 저장 기능이 곧 구현됩니다.');
    };
});
