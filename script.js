let appData = {
    headers: [],
    rows: [],
    fileName: 'data_export'
};

const els = {
    dropZone: document.getElementById('drop-zone'),
    fileInput: document.getElementById('file-input'),
    loader: document.getElementById('loader'),
    uploadSection: document.getElementById('upload-section'),
    dashboard: document.getElementById('dashboard'),
    cardsContainer: document.getElementById('cards-container'),
    // Stats
    totalRows: document.getElementById('val-total-rows'),
    barTotal: document.getElementById('bar-total'),
    missingRows: document.getElementById('val-missing-rows'),
    barMissing: document.getElementById('bar-missing'),
    txtMissingPerc: document.getElementById('txt-missing-perc'),
    // Charts
    colChart: document.getElementById('column-chart'),
    heatmap: document.getElementById('heatmap-table')
};

document.addEventListener('DOMContentLoaded', () => {
    els.dropZone.addEventListener('click', () => els.fileInput.click());
    els.dropZone.addEventListener('dragover', (e) => { e.preventDefault(); els.dropZone.style.borderColor = '#6366f1'; });
    els.dropZone.addEventListener('dragleave', () => { els.dropZone.style.borderColor = ''; });
    els.dropZone.addEventListener('drop', handleDrop);
    els.fileInput.addEventListener('change', (e) => processFile(e.target.files[0]));
    document.getElementById('search-input').addEventListener('input', (e) => renderCards(e.target.value));
});

function handleDrop(e) {
    e.preventDefault();
    if (e.dataTransfer.files.length) processFile(e.dataTransfer.files[0]);
}

function processFile(file) {
    if (!file) return;
    els.loader.classList.remove('hidden');
    appData.fileName = file.name.split('.')[0];

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        
        // 1. Convert to Array of Arrays first to get accurate headers
        const jsonSheet = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
        
        if (jsonSheet.length > 0) {
            // Assume Row 0 is Header. Filter out empty header cells if any.
            // This fixes "Unknown 1, 2" issue
            appData.headers = jsonSheet[0].filter(h => h !== null && h !== undefined && String(h).trim() !== "");
            
            // Map remaining rows to objects based on these headers
            appData.rows = [];
            for(let i = 1; i < jsonSheet.length; i++) {
                let rowArray = jsonSheet[i];
                let rowObj = { _id: i - 1 }; // Internal ID (0-based)
                
                // Map values to headers by index
                appData.headers.forEach((header, index) => {
                    // Find the index of this header in the original row 0
                    const originalIndex = jsonSheet[0].indexOf(header);
                    rowObj[header] = rowArray[originalIndex];
                });
                appData.rows.push(rowObj);
            }

            initDashboard();
        } else {
            alert("File appears empty!");
            els.loader.classList.add('hidden');
        }
    };
    reader.readAsArrayBuffer(file);
}

function countMissing(row) {
    let count = 0;
    appData.headers.forEach(h => {
        const val = row[h];
        if(val === null || val === undefined || String(val).trim() === "") count++;
    });
    return count;
}

function initDashboard() {
    els.uploadSection.classList.add('hidden');
    els.dashboard.classList.remove('hidden');
    updateGlobalStats();
    renderCharts();
    renderCards();
}

function updateGlobalStats() {
    const total = appData.rows.length;
    const missingCount = appData.rows.filter(r => countMissing(r) > 0).length;
    const missingPerc = total > 0 ? Math.round((missingCount / total) * 100) : 0;
    const completePerc = 100 - missingPerc;

    els.totalRows.innerText = total;
    els.missingRows.innerText = missingCount;
    els.barMissing.style.width = `${missingPerc}%`;
    els.txtMissingPerc.innerText = `${missingPerc}% Incomplete`;

    // Print Stats Update
    document.getElementById('print-date').innerText = new Date().toLocaleString();
    document.getElementById('p-total').innerText = total;
    document.getElementById('p-missing').innerText = missingCount;
    document.getElementById('p-complete').innerText = `${completePerc}%`;
}

function renderCharts() {
    // 1. Column Chart (Using actual headers)
    els.colChart.innerHTML = '';
    let maxMissing = 0;
    const colData = appData.headers.map(h => {
        const count = appData.rows.filter(r => r[h] === null || r[h] === "" || r[h] === undefined).length;
        if(count > maxMissing) maxMissing = count;
        return { header: h, count: count };
    });

    colData.forEach(d => {
        // Show all columns, even if 0 missing, for completeness
        const hPct = maxMissing > 0 ? (d.count / maxMissing) * 100 : 0;
        const bar = document.createElement('div');
        bar.className = 'c-bar-group';
        bar.innerHTML = `
            <div class="c-val-label">${d.count}</div>
            <div class="c-bar" style="height:${Math.max(hPct, 2)}%"></div>
            <div class="c-label" title="${d.header}">${d.header}</div>
        `;
        els.colChart.appendChild(bar);
    });

    // 2. Heatmap (Row / Name vs Actual Headers)
    const table = els.heatmap;
    let html = `<thead><tr><th class="hm-row-header">Row / Name</th>`;
    appData.headers.forEach(h => html += `<th>${h}</th>`);
    html += '</tr></thead><tbody>';

    appData.rows.slice(0, 50).forEach((row, i) => {
        const nameKey = appData.headers.find(h => h.toLowerCase().includes('name')) || appData.headers[0];
        const nameVal = row[nameKey] || `Row ${i+1}`;
        
        html += `<tr><td class="hm-row-header">${i+1}. ${nameVal}</td>`;
        appData.headers.forEach(h => {
            const isMissing = row[h] === null || row[h] === "" || row[h] === undefined;
            html += `<td class="${isMissing ? 'cell-missing' : 'cell-ok'}">
                ${isMissing ? '<i class="ph ph-x"></i>' : '<i class="ph ph-check"></i>'}
            </td>`;
        });
        html += '</tr>';
    });
    html += '</tbody>';
    table.innerHTML = html;
}

function renderCards(filter = '') {
    els.cardsContainer.innerHTML = '';
    const term = filter.toLowerCase();

    appData.rows.forEach(row => {
        // Dynamic Name Detection
        const nameKey = appData.headers.find(h => h.toLowerCase().includes('name')) || appData.headers[0];
        const nameVal = row[nameKey] || 'Unknown';
        
        // Dynamic Title Format: "01 - NAME"
        // Pad ID with leading zero if single digit
        const paddedId = String(row._id + 1).padStart(2, '0');
        const cardTitle = `${paddedId} - ${nameVal}`;

        // Search Filter
        if(term && !cardTitle.toLowerCase().includes(term) && !JSON.stringify(row).toLowerCase().includes(term)) return;

        const missingCount = countMissing(row);
        const statusClass = missingCount === 0 ? 'status-ok' : 'status-err';
        const statusText = missingCount === 0 ? 'Complete' : `${missingCount} Missing`;

        // Wrapper
        const scene = document.createElement('div');
        scene.className = 'card-scene';
        scene.id = `scene-${row._id}`;
        
        // --- FRONT SIDE (View Details) ---
        let frontHtml = '';
        let displayCount = 0;
        appData.headers.forEach(h => {
            if(displayCount < 6 && row[h] && h !== nameKey) {
                frontHtml += `
                    <div class="data-row">
                        <span class="d-label">${h}</span>
                        <span class="d-val">${row[h]}</span>
                    </div>`;
                displayCount++;
            }
        });

        // --- BACK SIDE (Edit Inputs - ONLY MISSING) ---
        let backHtml = '';
        let inputCount = 0;
        
        if(missingCount === 0) {
            backHtml = `<div style="text-align:center; padding:40px; color:#10b981;">
                <i class="ph ph-check-circle" style="font-size:48px;"></i><br>
                <p style="margin-top:10px">No Data Missing!</p>
            </div>`;
        } else {
            appData.headers.forEach(h => {
                // Show Input ONLY if missing
                if(row[h] === null || row[h] === "" || row[h] === undefined) {
                    backHtml += `
                        <div class="edit-group">
                            <label class="edit-label">${h} (Missing)</label>
                            <input type="text" class="edit-input" id="input-${row._id}-${inputCount}" data-header="${h}" placeholder="Enter value...">
                        </div>`;
                    inputCount++;
                }
            });
        }

        // HTML Structure
        scene.innerHTML = `
            <div class="card-face">
                <div class="card-header">
                    <span class="status-badge ${statusClass}">${statusText}</span>
                    <div class="card-title" title="${cardTitle}">${cardTitle}</div>
                </div>
                <div class="card-body">${frontHtml}</div>
                <div class="card-footer">
                    <button class="flip-btn" onclick="flipCard(${row._id}, true)">
                        ${missingCount > 0 ? 'Fix Missing' : 'View Data'} <i class="ph ph-arrow-right"></i>
                    </button>
                </div>
            </div>
            <div class="card-face card-back">
                <div class="card-header">
                    <span class="status-badge ${statusClass}">${statusText}</span>
                    <div class="card-title">${cardTitle}</div>
                </div>
                <div class="card-body">
                    ${backHtml}
                </div>
                <div class="card-footer">
                    <button class="flip-btn" onclick="flipCard(${row._id}, false)">
                         Cancel
                    </button>
                    ${missingCount > 0 ? `<button class="save-btn" onclick="saveCardData(${row._id})">Save Changes</button>` : ''}
                </div>
            </div>
        `;
        els.cardsContainer.appendChild(scene);
    });
}

function flipCard(id, isFlipping) {
    const el = document.getElementById(`scene-${id}`);
    if(isFlipping) el.classList.add('flipped');
    else el.classList.remove('flipped');
}

function saveCardData(rowId) {
    const row = appData.rows.find(r => r._id === rowId);
    const scene = document.getElementById(`scene-${rowId}`);
    const inputs = scene.querySelectorAll('.edit-input');
    
    // Update Data Model
    inputs.forEach(inp => {
        if(inp.value.trim() !== "") {
            const header = inp.getAttribute('data-header');
            row[header] = inp.value.trim();
        }
    });

    // Re-Calculate and Re-Render
    updateGlobalStats();
    renderCharts();
    
    // Re-render just this card is tricky without refreshing list, so refresh list for simplicity/accuracy
    // Or flip back manually and update DOM. Let's refresh full list to ensure sort/filter logic holds.
    renderCards(document.getElementById('search-input').value);
}

// --- Export & Print ---

function exportData(type) {
    // Clean Data (Remove _id)
    let finalData = appData.rows.map(({_id, ...rest}) => rest);

    if(type === 'missing') {
        finalData = finalData.filter(r => {
            return appData.headers.some(h => r[h] === null || r[h] === "" || r[h] === undefined);
        });
    }

    if(finalData.length === 0) {
        alert("No data found for this criteria.");
        return;
    }

    const ws = XLSX.utils.json_to_sheet(finalData, { header: appData.headers });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Data");
    XLSX.writeFile(wb, `${appData.fileName}_${type}.xlsx`);
}

function printReport() {
    const thead = document.getElementById('print-thead');
    const tbody = document.getElementById('print-tbody');

    // Build Header (Dynamic from file, skipping index)
    let headHtml = '<tr><th style="width:50px;">Serial</th>';
    appData.headers.forEach(h => headHtml += `<th>${h}</th>`);
    headHtml += '</tr>';
    thead.innerHTML = headHtml;

    // Build Body
    let bodyHtml = '';
    appData.rows.forEach((row, i) => {
        bodyHtml += `<tr><td>${i+1}</td>`;
        appData.headers.forEach(h => {
            const val = row[h];
            const isMiss = val === null || val === "" || val === undefined;
            // PDF Requirement: Show "N/A" if missing
            bodyHtml += `<td class="${isMiss ? 'print-val-missing' : ''}">${val || 'N/A'}</td>`;
        });
        bodyHtml += '</tr>';
    });
    tbody.innerHTML = bodyHtml;

    window.print();
}