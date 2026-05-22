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
        // Collect all OK files
        const inputs = document.querySelectorAll('input[type="file"]');
        const formData = new FormData();
        let validFiles = 0;

        inputs.forEach(input => {
            if (input.files.length > 0) {
                // Check if status is OK
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

        try {
            const response = await fetch('/extract', {
                method: 'POST',
                body: formData
            });
            const result = await response.json();
            
            if (result.status === 'success') {
                // Show manual input section (Step 2 logic)
                document.getElementById('manual-input-section').style.display = 'block';
                alert('추출 완료 (Step 2에서 로직 연동 예정)');
            } else {
                alert('에러: ' + result.message);
            }
        } catch (error) {
            alert('통신 에러');
        }
    };
});
