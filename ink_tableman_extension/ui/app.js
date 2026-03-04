document.addEventListener('DOMContentLoaded', () => {
    // ─── DOM References ──────────────────────────────────────
    const screenLanding = document.getElementById('screen-landing');
    const screenEditor = document.getElementById('screen-editor');
    const tableList = document.getElementById('table-list');
    const editorTitle = document.getElementById('editor-title');
    const editorForm = document.getElementById('editor-form');
    const tableIdInput = document.getElementById('table-id');
    const tableLabelInput = document.getElementById('table-label');
    const tableRowsInput = document.getElementById('table-rows');
    const tableColsInput = document.getElementById('table-cols');
    const contentGrid = document.getElementById('content-grid').querySelector('tbody');
    const fontSelect = document.getElementById('font-family');
    const fontSizeInput = document.getElementById('font-size');
    const textColorInput = document.getElementById('text-color');
    const headerTextColorInput = document.getElementById('header-text-color');
    const headerFillInput = document.getElementById('header-fill');
    const bodyFillInput = document.getElementById('body-fill');
    const statusContainer = document.getElementById('status-container');
    const progressBar = document.getElementById('progress-bar');
    const statusMessage = document.getElementById('status-message');

    let pollInterval = null;

    // ─── STATE ───────────────────────────────────────────────
    let merges = [];
    let cellStyles = []; // 2D array of {bold, italic, underline, strikethrough}
    let selectionStart = null;
    let selectionEnd = null;

    // ─── NAVIGATION ──────────────────────────────────────────

    function showLanding() {
        screenEditor.classList.add('hidden');
        screenLanding.classList.remove('hidden');
        loadTableList();
    }

    function showEditor(title = 'New Table') {
        screenLanding.classList.add('hidden');
        screenEditor.classList.remove('hidden');
        editorTitle.textContent = title;
        statusContainer.classList.add('hidden');
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
        document.querySelector('[data-tab="tab-layout"]').classList.add('active');
        document.getElementById('tab-layout').classList.add('active');
    }

    // ─── TAB SWITCHING ───────────────────────────────────────

    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
            btn.classList.add('active');
            document.getElementById(btn.dataset.tab).classList.add('active');
            if (btn.dataset.tab === 'tab-content') buildGrid();
        });
    });

    // ─── LOAD SYSTEM FONTS ───────────────────────────────────

    async function loadFonts(selectedFont) {
        try {
            const res = await fetch('/fonts');
            const data = await res.json();
            fontSelect.innerHTML = '';

            const generics = ['sans-serif', 'serif', 'monospace'];
            generics.forEach(f => {
                const opt = document.createElement('option');
                opt.value = f;
                opt.textContent = f;
                fontSelect.appendChild(opt);
            });

            const sep = document.createElement('option');
            sep.disabled = true;
            sep.textContent = '───────────────';
            fontSelect.appendChild(sep);

            if (data.fonts && data.fonts.length > 0) {
                data.fonts.forEach(f => {
                    if (!generics.includes(f)) {
                        const opt = document.createElement('option');
                        opt.value = f;
                        opt.textContent = f;
                        fontSelect.appendChild(opt);
                    }
                });
            }

            if (selectedFont) fontSelect.value = selectedFont;
        } catch (err) {
            fontSelect.innerHTML = '<option value="sans-serif">sans-serif</option>' +
                '<option value="serif">serif</option>' +
                '<option value="monospace">monospace</option>';
        }
    }

    // ─── LIVE PREVIEW ────────────────────────────────────────

    function applyLivePreview() {
        const font = fontSelect.value;
        const size = fontSizeInput.value + 'px';
        const color = textColorInput.value;
        const hColor = headerTextColorInput.value;
        const hFill = headerFillInput.value;
        const bFill = bodyFillInput.value;

        contentGrid.querySelectorAll('td').forEach(td => {
            const r = parseInt(td.dataset.row);
            td.style.fontFamily = font;
            td.style.fontSize = size;
            if (r === 0) {
                td.style.color = hColor;
                td.style.backgroundColor = hFill;
            } else {
                td.style.color = color;
                td.style.backgroundColor = bFill;
            }
        });
    }

    // Bind live preview listeners
    [fontSelect, fontSizeInput, textColorInput, headerTextColorInput,
        headerFillInput, bodyFillInput].forEach(el => {
            el.addEventListener('input', applyLivePreview);
            el.addEventListener('change', applyLivePreview);
        });

    // ─── LANDING: LOAD TABLE LIST ────────────────────────────

    async function loadTableList() {
        tableList.innerHTML = '<div class="table-list-loading">Scanning SVG...</div>';
        try {
            const res = await fetch('/tables');
            const data = await res.json();

            if (data.tables.length === 0) {
                tableList.innerHTML = '<div class="table-list-empty">No tables found.<br>Create one to get started!</div>';
            } else {
                tableList.innerHTML = '';
                data.tables.forEach(t => {
                    const card = document.createElement('div');
                    card.className = 'table-card';
                    card.innerHTML = `
                        <div class="table-card-info">
                            <div class="table-card-name">${escapeHtml(t.label)}</div>
                            <div class="table-card-meta">${t.rows}×${t.cols}</div>
                        </div>
                        <div class="table-card-actions">
                            <button class="btn-small btn-edit" data-id="${escapeHtml(t.id)}">Edit</button>
                            <button class="btn-small btn-danger" data-id="${escapeHtml(t.id)}">Del</button>
                        </div>
                    `;
                    tableList.appendChild(card);
                });

                tableList.querySelectorAll('.btn-edit').forEach(btn => {
                    btn.addEventListener('click', () => loadTableForEdit(btn.dataset.id));
                });
                tableList.querySelectorAll('.btn-danger').forEach(btn => {
                    btn.addEventListener('click', () => deleteTable(btn.dataset.id));
                });
            }
        } catch (err) {
            tableList.innerHTML = '<div class="table-list-empty">Failed to scan tables.</div>';
        }
    }

    // ─── LOAD TABLE FOR EDITING ──────────────────────────────

    async function loadTableForEdit(id) {
        try {
            const res = await fetch(`/table/${encodeURIComponent(id)}`);
            const data = await res.json();
            if (data.error) { alert('Table not found'); return; }

            tableIdInput.value = data.id;
            tableLabelInput.value = data.label.replace(/^Tableman:\s*/, '');
            tableRowsInput.value = data.rows;
            tableColsInput.value = data.cols;
            document.getElementById('cell-width').value = data.cell_width;
            document.getElementById('cell-height').value = data.cell_height;
            document.getElementById('border-width').value = data.border_width;
            document.getElementById('border-color').value = data.border_color;
            headerFillInput.value = data.header_fill;
            bodyFillInput.value = data.body_fill;
            fontSizeInput.value = data.font_size;
            textColorInput.value = data.text_color;
            headerTextColorInput.value = data.header_text_color;

            await loadFonts(data.font_family);

            window._tableData = data.data || [];
            merges = data.merges || [];
            cellStyles = data.cell_styles || [];

            showEditor(`Edit: ${data.label.replace(/^Tableman:\s*/, '')}`);
            buildGrid();
        } catch (err) {
            alert('Failed to load table: ' + err.message);
        }
    }

    // ─── DELETE TABLE ────────────────────────────────────────

    async function deleteTable(id) {
        if (!confirm('Delete this table from SVG?')) return;
        try {
            const res = await fetch('/delete', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ id })
            });
            const result = await res.json();
            if (result.status === 'deleted') loadTableList();
            else alert('Failed: ' + (result.message || 'Unknown'));
        } catch (err) { alert('Error: ' + err.message); }
    }

    // ─── EDITABLE GRID ───────────────────────────────────────

    function getMergeAt(r, c) {
        return merges.find(m => m.row === r && m.col === c) || null;
    }

    function isHiddenByMerge(r, c) {
        for (const m of merges) {
            if (r >= m.row && r < m.row + m.rowspan &&
                c >= m.col && c < m.col + m.colspan) {
                if (r === m.row && c === m.col) return false;
                return true;
            }
        }
        return false;
    }

    function ensureCellStyles(rows, cols) {
        while (cellStyles.length < rows) cellStyles.push([]);
        for (let r = 0; r < rows; r++) {
            while (cellStyles[r].length < cols) {
                cellStyles[r].push({ bold: false, italic: false, underline: false, strikethrough: false });
            }
        }
    }

    function getCellStyle(r, c) {
        if (r < cellStyles.length && c < cellStyles[r].length) {
            return cellStyles[r][c] || {};
        }
        return {};
    }

    function applyCellStyleToTd(td, r, c) {
        const s = getCellStyle(r, c);
        td.style.fontWeight = s.bold ? 'bold' : 'normal';
        td.style.fontStyle = s.italic ? 'italic' : 'normal';
        td.style.textDecoration = [
            s.underline ? 'underline' : '',
            s.strikethrough ? 'line-through' : ''
        ].filter(Boolean).join(' ') || 'none';
    }

    function buildGrid() {
        const rows = parseInt(tableRowsInput.value) || 3;
        const cols = parseInt(tableColsInput.value) || 3;
        const existingData = window._tableData || [];

        ensureCellStyles(rows, cols);
        contentGrid.innerHTML = '';
        clearSelection();

        for (let r = 0; r < rows; r++) {
            const tr = document.createElement('tr');
            for (let c = 0; c < cols; c++) {
                if (isHiddenByMerge(r, c)) continue;

                const td = document.createElement('td');
                td.contentEditable = true;
                td.dataset.row = r;
                td.dataset.col = c;

                const merge = getMergeAt(r, c);
                if (merge) {
                    td.rowSpan = merge.rowspan;
                    td.colSpan = merge.colspan;
                    td.classList.add('merged-cell');
                }

                if (r < existingData.length && c < existingData[r].length) {
                    td.textContent = existingData[r][c];
                } else if (r === 0) {
                    td.textContent = `Col ${c + 1}`;
                }

                td.classList.add(r === 0 ? 'header-cell' : 'body-cell');

                // Apply per-cell formatting
                applyCellStyleToTd(td, r, c);

                // Selection
                td.addEventListener('mousedown', (e) => {
                    if (e.shiftKey && selectionStart) {
                        selectionEnd = { row: r, col: c };
                    } else {
                        selectionStart = { row: r, col: c };
                        selectionEnd = { row: r, col: c };
                    }
                    highlightSelection();
                    updateFormatButtons();
                });

                tr.appendChild(td);
            }
            contentGrid.appendChild(tr);
        }

        applyLivePreview();
    }

    function collectGridData() {
        const rows = parseInt(tableRowsInput.value) || 3;
        const cols = parseInt(tableColsInput.value) || 3;
        const data = [];
        for (let r = 0; r < rows; r++) data.push(new Array(cols).fill(''));
        contentGrid.querySelectorAll('td').forEach(td => {
            const r = parseInt(td.dataset.row);
            const c = parseInt(td.dataset.col);
            if (r < rows && c < cols) data[r][c] = td.textContent;
        });
        return data;
    }

    function storeGridData() { window._tableData = collectGridData(); }

    // ─── SELECTION ───────────────────────────────────────────

    function clearSelection() {
        selectionStart = null;
        selectionEnd = null;
        contentGrid.querySelectorAll('.selected').forEach(td => td.classList.remove('selected'));
    }

    function getSelectionRange() {
        if (!selectionStart || !selectionEnd) return null;
        return {
            minR: Math.min(selectionStart.row, selectionEnd.row),
            maxR: Math.max(selectionStart.row, selectionEnd.row),
            minC: Math.min(selectionStart.col, selectionEnd.col),
            maxC: Math.max(selectionStart.col, selectionEnd.col),
        };
    }

    function highlightSelection() {
        contentGrid.querySelectorAll('.selected').forEach(td => td.classList.remove('selected'));
        const range = getSelectionRange();
        if (!range) return;

        contentGrid.querySelectorAll('td').forEach(td => {
            const r = parseInt(td.dataset.row);
            const c = parseInt(td.dataset.col);
            const merge = getMergeAt(r, c);
            const rs = merge ? merge.rowspan : 1;
            const cs = merge ? merge.colspan : 1;
            if (r + rs > range.minR && r <= range.maxR &&
                c + cs > range.minC && c <= range.maxC) {
                td.classList.add('selected');
            }
        });
    }

    // ─── FORMAT BUTTONS (B / I / U / S) ──────────────────────

    function updateFormatButtons() {
        if (!selectionStart) return;
        const s = getCellStyle(selectionStart.row, selectionStart.col);
        document.getElementById('fmt-bold').classList.toggle('active', !!s.bold);
        document.getElementById('fmt-italic').classList.toggle('active', !!s.italic);
        document.getElementById('fmt-underline').classList.toggle('active', !!s.underline);
        document.getElementById('fmt-strike').classList.toggle('active', !!s.strikethrough);
    }

    function toggleFormat(prop) {
        const range = getSelectionRange();
        if (!range) { alert('Select cells first.'); return; }

        const rows = parseInt(tableRowsInput.value);
        const cols = parseInt(tableColsInput.value);
        ensureCellStyles(rows, cols);

        // Determine toggle value from first selected cell
        const first = getCellStyle(range.minR, range.minC);
        const newVal = !first[prop];

        for (let r = range.minR; r <= range.maxR; r++) {
            for (let c = range.minC; c <= range.maxC; c++) {
                if (!isHiddenByMerge(r, c) || (r === range.minR && c === range.minC)) {
                    cellStyles[r][c][prop] = newVal;
                }
            }
        }

        // Apply to DOM
        contentGrid.querySelectorAll('td.selected').forEach(td => {
            const r = parseInt(td.dataset.row);
            const c = parseInt(td.dataset.col);
            applyCellStyleToTd(td, r, c);
        });

        updateFormatButtons();
    }

    document.getElementById('fmt-bold').addEventListener('click', () => toggleFormat('bold'));
    document.getElementById('fmt-italic').addEventListener('click', () => toggleFormat('italic'));
    document.getElementById('fmt-underline').addEventListener('click', () => toggleFormat('underline'));
    document.getElementById('fmt-strike').addEventListener('click', () => toggleFormat('strikethrough'));

    // ─── MERGE / BREAK ───────────────────────────────────────

    document.getElementById('btn-merge').addEventListener('click', () => {
        const range = getSelectionRange();
        if (!range) { alert('Select cells first.'); return; }
        const rowspan = range.maxR - range.minR + 1;
        const colspan = range.maxC - range.minC + 1;
        if (rowspan === 1 && colspan === 1) { alert('Select 2+ cells.'); return; }

        merges = merges.filter(m => {
            const mEndR = m.row + m.rowspan - 1;
            const mEndC = m.col + m.colspan - 1;
            return (m.row > range.maxR || mEndR < range.minR || m.col > range.maxC || mEndC < range.minC);
        });

        storeGridData();
        merges.push({ row: range.minR, col: range.minC, rowspan, colspan });
        buildGrid();
    });

    document.getElementById('btn-break').addEventListener('click', () => {
        if (!selectionStart) { alert('Select a merged cell.'); return; }
        const idx = merges.findIndex(m => m.row === selectionStart.row && m.col === selectionStart.col);
        if (idx === -1) { alert('Not merged.'); return; }
        storeGridData();
        merges.splice(idx, 1);
        buildGrid();
    });

    // ─── GRID TOOLBAR ────────────────────────────────────────

    document.getElementById('btn-add-row').addEventListener('click', () => {
        storeGridData(); tableRowsInput.value = parseInt(tableRowsInput.value) + 1; buildGrid();
    });
    document.getElementById('btn-remove-row').addEventListener('click', () => {
        if (parseInt(tableRowsInput.value) > 1) {
            storeGridData();
            const n = parseInt(tableRowsInput.value) - 1;
            tableRowsInput.value = n;
            merges = merges.filter(m => m.row + m.rowspan <= n + 1);
            buildGrid();
        }
    });
    document.getElementById('btn-add-col').addEventListener('click', () => {
        storeGridData(); tableColsInput.value = parseInt(tableColsInput.value) + 1; buildGrid();
    });
    document.getElementById('btn-remove-col').addEventListener('click', () => {
        if (parseInt(tableColsInput.value) > 1) {
            storeGridData();
            const n = parseInt(tableColsInput.value) - 1;
            tableColsInput.value = n;
            merges = merges.filter(m => m.col + m.colspan <= n + 1);
            buildGrid();
        }
    });

    // ─── SUBMIT ──────────────────────────────────────────────

    editorForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        storeGridData();

        const payload = {
            id: tableIdInput.value,
            label: tableLabelInput.value,
            rows: parseInt(tableRowsInput.value),
            cols: parseInt(tableColsInput.value),
            cell_width: parseFloat(document.getElementById('cell-width').value),
            cell_height: parseFloat(document.getElementById('cell-height').value),
            border_width: parseFloat(document.getElementById('border-width').value),
            border_color: document.getElementById('border-color').value,
            header_fill: headerFillInput.value,
            body_fill: bodyFillInput.value,
            font_family: fontSelect.value,
            font_size: parseFloat(fontSizeInput.value),
            text_color: textColorInput.value,
            header_text_color: headerTextColorInput.value,
            data: window._tableData || collectGridData(),
            merges: merges,
            cell_styles: cellStyles
        };

        const submitBtn = document.getElementById('btn-submit');
        submitBtn.disabled = true;
        statusContainer.classList.remove('hidden');
        updateProgress(5, 'Submitting...', 'normal');

        try {
            const res = await fetch('/submit', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            if (res.ok) pollStatus();
            else { updateProgress(0, 'Failed.', 'error'); submitBtn.disabled = false; }
        } catch (err) {
            updateProgress(0, `Error: ${err.message}`, 'error');
            submitBtn.disabled = false;
        }
    });

    // ─── BUTTONS ─────────────────────────────────────────────

    document.getElementById('btn-new-table').addEventListener('click', () => {
        editorForm.reset();
        tableIdInput.value = '';
        tableRowsInput.value = 4;
        tableColsInput.value = 4;
        window._tableData = [];
        merges = [];
        cellStyles = [];
        showEditor('New Table');
        buildGrid();
    });

    document.getElementById('btn-back').addEventListener('click', () => {
        window._tableData = []; merges = []; cellStyles = []; showLanding();
    });
    document.getElementById('btn-cancel-editor').addEventListener('click', () => {
        window._tableData = []; merges = []; cellStyles = []; showLanding();
    });
    document.getElementById('btn-close-landing').addEventListener('click', async () => {
        try { await fetch('/close', { method: 'POST' }); } catch (e) { }
    });

    // ─── PROGRESS ────────────────────────────────────────────

    function updateProgress(p, msg, state = 'normal') {
        progressBar.style.width = `${p}%`;
        statusMessage.textContent = msg;
        progressBar.classList.toggle('error', state === 'error');
    }

    async function pollStatus() {
        if (pollInterval) clearInterval(pollInterval);
        pollInterval = setInterval(async () => {
            try {
                const res = await fetch('/status');
                const s = await res.json();
                updateProgress(s.progress, s.message, s.status === 'error' ? 'error' : 'normal');
                if (s.status === 'completed' || s.status === 'error') {
                    clearInterval(pollInterval);
                    document.getElementById('btn-submit').disabled = false;
                }
            } catch (err) {
                clearInterval(pollInterval);
                document.getElementById('btn-submit').disabled = false;
            }
        }, 500);
    }

    function escapeHtml(str) {
        const d = document.createElement('div');
        d.textContent = str;
        return d.innerHTML;
    }

    // ─── INIT ────────────────────────────────────────────────
    loadFonts();
    loadTableList();
});
