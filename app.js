// ===== DPV Flash Update Generator - Application Logic =====

let workbookData = {};  // Parsed Excel data by sheet name
let lastGeneratedHtml = '';

// --- Step Navigation ---

function goToStep(stepNum) {
    document.querySelectorAll('.step-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.step-content').forEach(content => content.classList.remove('active'));
    
    document.querySelector(`.step-btn[data-step="${stepNum}"]`).classList.add('active');
    document.getElementById(`step-${stepNum}`).classList.add('active');

    if (stepNum === 3) {
        generateFlash();
    }
}

// Initialize step button clicks
document.addEventListener('DOMContentLoaded', () => {
    document.querySelectorAll('.step-btn').forEach(btn => {
        btn.addEventListener('click', () => goToStep(parseInt(btn.dataset.step)));
    });

    // Set today's date
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('update-date').value = today;

    // Calculate current week number
    const now = new Date();
    const start = new Date(now.getFullYear(), 0, 1);
    const weekNum = Math.ceil(((now - start) / 86400000 + start.getDay() + 1) / 7);
    document.getElementById('update-week').value = `WK${weekNum}`;

    // Show placeholder in preview
    const iframe = document.getElementById('preview-iframe');
    const doc = iframe.contentDocument || iframe.contentWindow.document;
    doc.open();
    doc.write('<html><body style="display:flex;align-items:center;justify-content:center;height:100vh;font-family:Segoe UI,sans-serif;color:#a0aec0;text-align:center;padding:40px;"><div><p style="font-size:48px;margin-bottom:16px;">⚡</p><p style="font-size:16px;">Upload your data and edit the narrative to generate the flash update</p></div></body></html>');
    doc.close();

    initFileUpload();
});

// --- File Upload & Excel Parsing ---

function initFileUpload() {
    const uploadZone = document.getElementById('upload-zone');
    const fileInput = document.getElementById('file-input');

    uploadZone.addEventListener('click', () => fileInput.click());
    
    uploadZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadZone.classList.add('dragover');
    });

    uploadZone.addEventListener('dragleave', () => {
        uploadZone.classList.remove('dragover');
    });

    uploadZone.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadZone.classList.remove('dragover');
        if (e.dataTransfer.files.length) {
            handleFile(e.dataTransfer.files[0]);
        }
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) {
            handleFile(e.target.files[0]);
        }
    });
}

function handleFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            workbookData = {};
            workbook.SheetNames.forEach(name => {
                workbookData[name] = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1 });
            });

            // Update UI
            document.getElementById('upload-zone').classList.add('hidden');
            const fileStatus = document.getElementById('file-status');
            fileStatus.classList.remove('hidden');
            fileStatus.querySelector('.file-name').textContent = `📄 ${file.name}`;
            fileStatus.querySelector('.file-sheets').textContent = `${workbook.SheetNames.length} sheets detected`;

            // Show sheet mapping
            const mappingSection = document.getElementById('sheet-mapping');
            mappingSection.classList.remove('hidden');
            populateSheetDropdowns(workbook.SheetNames);

            showToast(`✅ Loaded: ${file.name}`);
        } catch (err) {
            showToast('❌ Error reading file. Please check the format.');
            console.error(err);
        }
    };
    reader.readAsArrayBuffer(file);
}

function populateSheetDropdowns(sheetNames) {
    const tableLabels = ['table1', 'table2', 'table3', 'table4', 'table5', 'table6'];
    const autoMatchKeywords = [
        ['derisking', 'de-risking', 'status', 'table 1', 'table1'],
        ['enforcement', 'upcoming', 'table 2', 'table2'],
        ['vos', 'voc', 'sentiment', 'anecdote', 'table 3', 'table3'],
        ['3p', 'engagement 3p', 'table 4', 'table4'],
        ['1p', 'engagement 1p', 'retail', 'table 5', 'table5'],
        ['roadmap', 'feature', 'table 6', 'table6']
    ];

    tableLabels.forEach((tableId, idx) => {
        const select = document.getElementById(`map-${tableId}`);
        // Clear existing options except first
        select.innerHTML = '<option value="">— Select sheet —</option>';
        
        sheetNames.forEach(name => {
            const option = document.createElement('option');
            option.value = name;
            option.textContent = name;
            select.appendChild(option);
        });

        // Auto-match based on keywords
        const keywords = autoMatchKeywords[idx];
        const match = sheetNames.find(name => 
            keywords.some(kw => name.toLowerCase().includes(kw))
        );
        if (match) {
            select.value = match;
        }
    });
}

function clearFile() {
    workbookData = {};
    document.getElementById('upload-zone').classList.remove('hidden');
    document.getElementById('file-status').classList.add('hidden');
    document.getElementById('sheet-mapping').classList.add('hidden');
    document.getElementById('file-input').value = '';
}

// --- Get Table Data ---

function getTableData(tableNum) {
    const sheetName = document.getElementById(`map-table${tableNum}`).value;
    if (!sheetName || !workbookData[sheetName]) return null;
    return workbookData[sheetName];
}

// --- Generate Flash Update HTML ---

function generateFlash() {
    const week = document.getElementById('update-week').value || 'WK__';
    const date = document.getElementById('update-date').value;
    const bcc = document.getElementById('bcc-list').value;

    // Narrative sections
    const narrativeProgram = document.getElementById('narrative-program').value;
    const narrativeStakeholder = document.getElementById('narrative-stakeholder').value;
    const narrativeToys = document.getElementById('narrative-toys').value;
    const narrativePowerbanks = document.getElementById('narrative-powerbanks').value;
    const narrativeSentiment = document.getElementById('narrative-sentiment').value;
    const narrativeEngToys = document.getElementById('narrative-engagement-toys').value;
    const narrativeEngPB = document.getElementById('narrative-engagement-pb').value;

    // Table data
    const table1Data = getTableData(1);
    const table2Data = getTableData(2);
    const table3Data = getTableData(3);
    const table4Data = getTableData(4);
    const table5Data = getTableData(5);
    const table6Data = getTableData(6);

    const formattedDate = date ? new Date(date).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }) : '';

    lastGeneratedHtml = buildEmailHtml({
        week, date: formattedDate, bcc,
        narrativeProgram, narrativeStakeholder,
        narrativeToys, narrativePowerbanks,
        narrativeSentiment,
        narrativeEngToys, narrativeEngPB,
        table1Data, table2Data, table3Data,
        table4Data, table5Data, table6Data
    });

    // Render in iframe
    const iframe = document.getElementById('preview-iframe');
    const doc = iframe.contentDocument || iframe.contentWindow.document;
    doc.open();
    doc.write(lastGeneratedHtml);
    doc.close();

    goToStep(3);
    showToast('✅ Flash update generated');
}

function buildEmailHtml(data) {
    return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>DPV Flash Update - ${escHtml(data.week)}</title>
    <style>
        body {
            font-family: 'Segoe UI', Calibri, Arial, sans-serif;
            font-size: 11pt;
            color: #1a1a1a;
            margin: 0;
            padding: 32px 48px;
            line-height: 1.6;
            background: #ffffff;
        }
        .bcc-note {
            font-size: 9pt;
            color: #718096;
            font-style: italic;
            margin-bottom: 20px;
            padding-bottom: 12px;
            border-bottom: 1px solid #e2e8f0;
        }
        .intro {
            font-size: 10.5pt;
            color: #4a5568;
            margin-bottom: 24px;
            font-style: italic;
        }
        h2 {
            font-size: 13pt;
            color: #232f3e;
            margin-top: 28px;
            margin-bottom: 8px;
            padding-bottom: 4px;
            border-bottom: 2px solid #232f3e;
        }
        h3 {
            font-size: 11.5pt;
            color: #37475a;
            margin-top: 18px;
            margin-bottom: 6px;
        }
        p {
            margin: 8px 0;
            text-align: justify;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 12px 0 20px 0;
            font-size: 10pt;
        }
        th {
            background: #232f3e;
            color: #ffffff;
            text-align: left;
            padding: 8px 10px;
            font-size: 9pt;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.3px;
        }
        td {
            padding: 7px 10px;
            border-bottom: 1px solid #e2e8f0;
            vertical-align: top;
        }
        tr:nth-child(even) td {
            background: #f7fafc;
        }
        .table-caption {
            font-size: 9.5pt;
            font-weight: 600;
            color: #4a5568;
            margin-top: 16px;
            margin-bottom: 4px;
            font-style: italic;
        }
        .status-green {
            color: #276749;
            font-weight: 700;
        }
        .status-yellow {
            color: #b7791f;
            font-weight: 700;
        }
        .status-red {
            color: #c53030;
            font-weight: 700;
        }
        .highlight {
            background: #fffbea;
            border-left: 4px solid #f6ad55;
            padding: 10px 14px;
            margin: 12px 0;
            font-size: 10.5pt;
        }
        .footer {
            margin-top: 36px;
            padding-top: 12px;
            border-top: 2px solid #e2e8f0;
            font-size: 9pt;
            color: #718096;
        }
    </style>
</head>
<body>

<div class="bcc-note">In BCC: ${escHtml(data.bcc)} - subscribe to receive DPV updates</div>

<div class="intro">This is a bi-weekly update on the Program and enforcement status for DPV Safety Controls (Children's Toys and Powerbanks).</div>

<h2>1. Program & Business Update</h2>
<p>${formatNarrative(data.narrativeProgram)}</p>

${data.narrativeStakeholder ? `<h3>1.1. Industry Stakeholder Feedback</h3>
<p>${formatNarrative(data.narrativeStakeholder)}</p>` : ''}

<h2>2. Current Status and Forecast</h2>

<h3>2.1. Children's Toys</h3>
<p>${formatNarrative(data.narrativeToys)}</p>

${data.table1Data ? `<p class="table-caption">Table 1: De-risking Status</p>
${renderTable(data.table1Data)}` : ''}

${data.table2Data ? `<p class="table-caption">Table 2: Children's Toys Upcoming Enforcement</p>
${renderTable(data.table2Data)}` : ''}

<h3>2.2. Power Banks</h3>
<p>${formatNarrative(data.narrativePowerbanks)}</p>

<h2>3. Seller Sentiment (SSM)</h2>
<p>${formatNarrative(data.narrativeSentiment)}</p>

${data.table3Data ? `<p class="table-caption">Table 3: SPs VoS Summary Table</p>
${renderTable(data.table3Data)}` : ''}

<h2>4. Selling Partners Engagement Plan</h2>

<h3>4.1. Children's Toys</h3>
<p>${formatNarrative(data.narrativeEngToys)}</p>

${data.table4Data ? `<p class="table-caption">Table 4: 3P Engagement Status</p>
${renderTable(data.table4Data)}` : ''}

${data.table5Data ? `<p class="table-caption">Table 5: 1P Engagement Status</p>
${renderTable(data.table5Data)}` : ''}

<h3>4.2. Power Banks</h3>
<p>${formatNarrative(data.narrativeEngPB)}</p>

${data.table6Data ? `<p class="table-caption">Table 6: 2026 DPV SP Feature Roadmap</p>
${renderTable(data.table6Data)}` : ''}

<div class="footer">
    <p>DPV Flash Update — ${escHtml(data.week)} | ${escHtml(data.date)}</p>
</div>

</body>
</html>`;
}

// --- Utility Functions ---

function escHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function formatNarrative(text) {
    if (!text) return '<span style="color:#a0aec0;font-style:italic;">[Enter narrative in Step 2]</span>';
    // Convert line breaks to <br>, preserve paragraph structure
    return escHtml(text).replace(/\n\n/g, '</p><p>').replace(/\n/g, '<br>');
}

function renderTable(data) {
    if (!data || data.length === 0) return '<p style="color:#a0aec0;font-style:italic;">[No data]</p>';

    const headers = data[0];
    const rows = data.slice(1).filter(row => row.some(cell => cell !== null && cell !== undefined && cell !== ''));

    if (headers.length === 0) return '<p style="color:#a0aec0;font-style:italic;">[Empty table]</p>';

    let html = '<table>\n<thead><tr>';
    headers.forEach(h => {
        html += `<th>${escHtml(String(h || ''))}</th>`;
    });
    html += '</tr></thead>\n<tbody>';

    rows.forEach(row => {
        html += '<tr>';
        headers.forEach((_, idx) => {
            const cell = row[idx];
            let cellStr = cell !== null && cell !== undefined ? String(cell) : '';
            
            // Auto-format: detect percentages, currency, status colors
            let style = '';
            if (cellStr.toLowerCase() === 'green' || cellStr.toLowerCase() === '🟢') {
                style = ' class="status-green"';
                cellStr = '🟢 ' + cellStr;
            } else if (cellStr.toLowerCase() === 'yellow' || cellStr.toLowerCase() === '🟡') {
                style = ' class="status-yellow"';
                cellStr = '🟡 ' + cellStr;
            } else if (cellStr.toLowerCase() === 'red' || cellStr.toLowerCase() === '🔴') {
                style = ' class="status-red"';
                cellStr = '🔴 ' + cellStr;
            }

            html += `<td${style}>${escHtml(cellStr)}</td>`;
        });
        html += '</tr>\n';
    });

    html += '</tbody></table>';
    return html;
}

// --- Export Functions ---

function copyToClipboard() {
    if (!lastGeneratedHtml) {
        generateFlash();
    }

    navigator.clipboard.writeText(lastGeneratedHtml).then(() => {
        showToast('✅ HTML copied to clipboard — paste into your email client');
    }).catch(() => {
        const textarea = document.createElement('textarea');
        textarea.value = lastGeneratedHtml;
        document.body.appendChild(textarea);
        textarea.select();
        document.execCommand('copy');
        document.body.removeChild(textarea);
        showToast('✅ HTML copied to clipboard');
    });
}

function downloadHtml() {
    if (!lastGeneratedHtml) {
        generateFlash();
    }

    const week = document.getElementById('update-week').value || 'WK__';
    const filename = `DPV_Flash_Update_${week}.html`;

    const blob = new Blob([lastGeneratedHtml], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);

    showToast(`✅ Downloaded: ${filename}`);
}

// --- Toast Notifications ---

function showToast(message) {
    let toast = document.querySelector('.toast');
    if (!toast) {
        toast = document.createElement('div');
        toast.className = 'toast';
        document.body.appendChild(toast);
    }
    toast.textContent = message;
    toast.classList.add('show');
    setTimeout(() => toast.classList.remove('show'), 3000);
}
