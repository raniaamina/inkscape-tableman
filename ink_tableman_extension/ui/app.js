document.addEventListener('DOMContentLoaded', () => {
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
    const gridTable = document.getElementById('content-grid');
    const statusContainer = document.getElementById('status-container');
    const progressBar = document.getElementById('progress-bar');
    const statusMessage = document.getElementById('status-message');

    const gFont = document.getElementById('font-family');
    const gFontSize = document.getElementById('font-size');
    const gTextColor = document.getElementById('text-color');
    const gHdrTextColor = document.getElementById('header-text-color');
    const gHeaderFill = document.getElementById('header-fill');
    const gBodyFill = document.getElementById('body-fill');

    const tbFont = document.getElementById('tb-font');
    const tbFontSize = document.getElementById('tb-font-size');
    const tbTextColor = document.getElementById('tb-text-color');
    const tbFillColor = document.getElementById('tb-fill-color');

    let pollInterval = null;
    let merges = [];
    let cellStyles = [];
    let colWidths = [];
    let rowHeights = [];
    let selStart = null, selEnd = null;

    const defCS = () => ({
        bold: false, italic: false, underline: false, strikethrough: false,
        textColor: null, fillColor: null, fontFamily: null, fontSize: null,
        hAlign: 'center', vAlign: 'middle', wrap: false, rotation: 0
    });

    function ensureStyles(rows, cols) {
        while (cellStyles.length < rows) cellStyles.push([]);
        for (let r = 0; r < rows; r++) {
            while (cellStyles[r].length < cols) cellStyles[r].push(defCS());
            for (let c = 0; c < cols; c++) {
                if (!cellStyles[r][c] || typeof cellStyles[r][c] !== 'object') cellStyles[r][c] = defCS();
                else { const d = defCS(); for (const k in d) if (!(k in cellStyles[r][c])) cellStyles[r][c][k] = d[k]; }
            }
        }
    }
    function getCS(r, c) { return (r < cellStyles.length && c < cellStyles[r].length) ? cellStyles[r][c] : defCS(); }
    function esc(s) { const d = document.createElement('div'); d.textContent = s; return d.innerHTML; }

    function ensureDimArrays() {
        const rows = parseInt(tableRowsInput.value) || 3;
        const cols = parseInt(tableColsInput.value) || 3;
        const defW = parseFloat(document.getElementById('cell-width').value) || 120;
        const defH = parseFloat(document.getElementById('cell-height').value) || 40;
        while (colWidths.length < cols) colWidths.push(defW);
        while (rowHeights.length < rows) rowHeights.push(defH);
        colWidths.length = cols;
        rowHeights.length = rows;
    }

    // ─── NAV ─────────────────────────────────────────────────
    function showLanding() { screenEditor.classList.add('hidden'); screenLanding.classList.remove('hidden'); loadTableList(); }
    function showEditor(t) {
        screenLanding.classList.add('hidden'); screenEditor.classList.remove('hidden');
        editorTitle.textContent = t || 'New Table'; statusContainer.classList.add('hidden');
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
        document.querySelector('[data-tab="tab-layout"]').classList.add('active');
        document.getElementById('tab-layout').classList.add('active');
    }

    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
            btn.classList.add('active'); document.getElementById(btn.dataset.tab).classList.add('active');
            if (btn.dataset.tab === 'tab-content') { syncToolbarFonts(); buildGrid(); }
        });
    });

    // ─── FONTS ───────────────────────────────────────────────
    async function loadFonts(sel) {
        try {
            const d = await (await fetch('/fonts')).json();
            const g = ['sans-serif', 'serif', 'monospace'];
            [gFont, tbFont].forEach(s => {
                s.innerHTML = '';
                g.forEach(f => s.appendChild(new Option(f, f)));
                const sep = new Option('─────────'); sep.disabled = true; s.appendChild(sep);
                if (d.fonts) d.fonts.forEach(f => { if (!g.includes(f)) s.appendChild(new Option(f, f)); });
            });
            if (sel) { gFont.value = sel; tbFont.value = sel; }
        } catch (e) {
            [gFont, tbFont].forEach(s => { s.innerHTML = '';['sans-serif', 'serif', 'monospace'].forEach(f => s.appendChild(new Option(f, f))); });
        }
    }
    function syncToolbarFonts() { tbFont.value = gFont.value; tbFontSize.value = gFontSize.value; }

    // ─── TABLE LIST ──────────────────────────────────────────
    async function loadTableList() {
        tableList.innerHTML = '<div class="table-list-loading">Scanning SVG...</div>';
        try {
            const d = await (await fetch('/tables')).json();
            if (!d.tables.length) { tableList.innerHTML = '<div class="table-list-empty">No tables found.<br>Create one!</div>'; return; }
            tableList.innerHTML = '';
            d.tables.forEach(t => {
                const c = document.createElement('div'); c.className = 'table-card';
                c.innerHTML = `<div class="table-card-info"><div class="table-card-name">${esc(t.label)}</div><div class="table-card-meta">${t.rows}×${t.cols}</div></div><div class="table-card-actions"><button class="btn-small btn-edit" data-id="${esc(t.id)}">Edit</button><button class="btn-small btn-danger" data-id="${esc(t.id)}">Del</button></div>`;
                tableList.appendChild(c);
            });
            tableList.querySelectorAll('.btn-edit').forEach(b => b.addEventListener('click', () => loadEdit(b.dataset.id)));
            tableList.querySelectorAll('.btn-danger').forEach(b => b.addEventListener('click', () => delTable(b.dataset.id)));
        } catch (e) { tableList.innerHTML = '<div class="table-list-empty">Failed.</div>'; }
    }

    async function loadEdit(id) {
        try {
            const d = await (await fetch(`/table/${encodeURIComponent(id)}`)).json();
            if (d.error) { alert('Not found'); return; }
            tableIdInput.value = d.id; tableLabelInput.value = d.label.replace(/^Tableman:\s*/, '');
            tableRowsInput.value = d.rows; tableColsInput.value = d.cols;
            document.getElementById('cell-width').value = d.cell_width;
            document.getElementById('cell-height').value = d.cell_height;
            document.getElementById('border-width').value = d.border_width;
            document.getElementById('border-color').value = d.border_color;
            gHeaderFill.value = d.header_fill; gBodyFill.value = d.body_fill;
            gFontSize.value = d.font_size; gTextColor.value = d.text_color; gHdrTextColor.value = d.header_text_color;
            await loadFonts(d.font_family);
            window._tableData = d.data || []; merges = d.merges || []; cellStyles = d.cell_styles || [];
            colWidths = d.col_widths || []; rowHeights = d.row_heights || [];
            showEditor(`Edit: ${d.label.replace(/^Tableman:\s*/, '')}`); buildGrid();
        } catch (e) { alert(e.message); }
    }

    async function delTable(id) {
        if (!confirm('Delete this table?')) return;
        try {
            const d = await (await fetch('/delete', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ id }) })).json();
            if (d.status === 'deleted') loadTableList(); else alert(d.message || 'Failed');
        } catch (e) { alert(e.message); }
    }

    // ─── MERGE ───────────────────────────────────────────────
    function getMerge(r, c) { return merges.find(m => m.row === r && m.col === c) || null; }
    function isHidden(r, c) {
        for (const m of merges) { if (r >= m.row && r < m.row + m.rowspan && c >= m.col && c < m.col + m.colspan && !(r === m.row && c === m.col)) return true; }
        return false;
    }

    // ─── BUILD GRID ──────────────────────────────────────────
    function buildGrid() {
        const rows = parseInt(tableRowsInput.value) || 3;
        const cols = parseInt(tableColsInput.value) || 3;
        const data = window._tableData || [];
        ensureStyles(rows, cols);
        ensureDimArrays();

        // Build colgroup for widths
        let cg = gridTable.querySelector('colgroup');
        if (cg) cg.remove();
        cg = document.createElement('colgroup');
        for (let c = 0; c < cols; c++) {
            const col = document.createElement('col');
            col.style.width = colWidths[c] + 'px';
            cg.appendChild(col);
        }
        gridTable.prepend(cg);

        contentGrid.innerHTML = '';
        clearSel();

        for (let r = 0; r < rows; r++) {
            const tr = document.createElement('tr');
            tr.style.height = rowHeights[r] + 'px';
            for (let c = 0; c < cols; c++) {
                if (isHidden(r, c)) continue;
                const td = document.createElement('td');
                td.contentEditable = true; td.dataset.row = r; td.dataset.col = c;

                const mg = getMerge(r, c);
                if (mg) { td.rowSpan = mg.rowspan; td.colSpan = mg.colspan; td.classList.add('merged-cell'); }
                if (r < data.length && c < data[r].length) td.textContent = data[r][c];
                else if (r === 0) td.textContent = `Col ${c + 1}`;

                applyCellVisual(td, r, c);

                td.addEventListener('mousedown', e => {
                    if (e.shiftKey && selStart) selEnd = { row: r, col: c };
                    else { selStart = { row: r, col: c }; selEnd = { row: r, col: c }; }
                    highlightSel(); syncToolbarFromCell(r, c);
                });

                tr.appendChild(td);
            }
            contentGrid.appendChild(tr);
        }

        addResizeHandles();
    }

    function applyCellVisual(td, r, c) {
        const s = getCS(r, c); const isH = r === 0;
        td.style.fontFamily = s.fontFamily || gFont.value;
        td.style.fontSize = (s.fontSize || gFontSize.value) + 'px';
        td.style.color = s.textColor || (isH ? gHdrTextColor.value : gTextColor.value);
        td.style.fontWeight = s.bold ? 'bold' : 'normal';
        td.style.fontStyle = s.italic ? 'italic' : 'normal';
        td.style.textDecoration = [s.underline ? 'underline' : '', s.strikethrough ? 'line-through' : ''].filter(Boolean).join(' ') || 'none';
        td.style.textAlign = s.hAlign || 'center';
        td.style.verticalAlign = s.vAlign || 'middle';
        td.style.whiteSpace = s.wrap ? 'normal' : 'nowrap';
        td.style.wordBreak = s.wrap ? 'break-word' : 'normal';

        // Rotation — safe preview using writing-mode only
        const rot = parseFloat(s.rotation || 0);
        td.style.writingMode = ''; td.style.textOrientation = ''; td.style.transform = '';
        if (rot === 90 || rot === -90) {
            td.style.writingMode = 'vertical-rl';
            td.style.textOrientation = 'mixed';
            if (rot === -90) td.style.transform = 'rotate(180deg)';
        } else if (rot === 270) {
            td.style.writingMode = 'vertical-rl';
            td.style.textOrientation = 'upright';
        } else if (rot === 45 || rot === -45) {
            // Mild tilt — no transform on td (would break layout), show via data attr
            td.dataset.rotHint = rot + '°';
        }

        td.style.backgroundColor = s.fillColor || (isH ? gHeaderFill.value : gBodyFill.value);
    }

    function collectData() {
        const rows = parseInt(tableRowsInput.value) || 3, cols = parseInt(tableColsInput.value) || 3;
        const data = []; for (let r = 0; r < rows; r++)data.push(new Array(cols).fill(''));
        contentGrid.querySelectorAll('td').forEach(td => {
            const r = parseInt(td.dataset.row), c = parseInt(td.dataset.col);
            if (r < rows && c < cols) data[r][c] = td.textContent;
        }); return data;
    }
    function storeData() { window._tableData = collectData(); }

    // ─── RESIZE HANDLES ──────────────────────────────────────

    function addResizeHandles() {
        // Remove old handles
        document.querySelectorAll('.col-resize-handle,.row-resize-handle').forEach(h => h.remove());

        const wrapper = document.querySelector('.grid-wrapper');
        const tableRect = gridTable.getBoundingClientRect();
        const wrapRect = wrapper.getBoundingClientRect();
        const cols = parseInt(tableColsInput.value) || 3;
        const rows = parseInt(tableRowsInput.value) || 3;

        // Column resize handles (at right edge of each col except last)
        let xOff = 0;
        for (let c = 0; c < cols - 1; c++) {
            xOff += colWidths[c];
            const handle = document.createElement('div');
            handle.className = 'col-resize-handle';
            handle.style.left = (xOff - 2) + 'px';
            handle.style.top = '0';
            handle.style.height = gridTable.offsetHeight + 'px';
            handle.dataset.col = c;
            wrapper.appendChild(handle);

            handle.addEventListener('mousedown', e => {
                e.preventDefault();
                const startX = e.clientX;
                const startW = colWidths[c];
                const onMove = ev => {
                    const delta = ev.clientX - startX;
                    colWidths[c] = Math.max(30, startW + delta);
                    buildGrid();
                };
                const onUp = () => { document.removeEventListener('mousemove', onMove); document.removeEventListener('mouseup', onUp); };
                document.addEventListener('mousemove', onMove);
                document.addEventListener('mouseup', onUp);
            });
        }

        // Row resize handles (at bottom edge of each row)
        let yOff = 0;
        for (let r = 0; r < rows - 1; r++) {
            yOff += rowHeights[r];
            const handle = document.createElement('div');
            handle.className = 'row-resize-handle';
            handle.style.top = (yOff - 2) + 'px';
            handle.style.left = '0';
            handle.style.width = gridTable.offsetWidth + 'px';
            handle.dataset.row = r;
            wrapper.appendChild(handle);

            handle.addEventListener('mousedown', e => {
                e.preventDefault();
                const startY = e.clientY;
                const startH = rowHeights[r];
                const onMove = ev => {
                    const delta = ev.clientY - startY;
                    rowHeights[r] = Math.max(20, startH + delta);
                    buildGrid();
                };
                const onUp = () => { document.removeEventListener('mousemove', onMove); document.removeEventListener('mouseup', onUp); };
                document.addEventListener('mousemove', onMove);
                document.addEventListener('mouseup', onUp);
            });
        }
    }

    // ─── SELECTION ───────────────────────────────────────────
    function clearSel() { selStart = selEnd = null; contentGrid.querySelectorAll('.selected').forEach(td => td.classList.remove('selected')); }
    function getRange() {
        if (!selStart || !selEnd) return null;
        return {
            minR: Math.min(selStart.row, selEnd.row), maxR: Math.max(selStart.row, selEnd.row),
            minC: Math.min(selStart.col, selEnd.col), maxC: Math.max(selStart.col, selEnd.col)
        };
    }
    function highlightSel() {
        contentGrid.querySelectorAll('.selected').forEach(td => td.classList.remove('selected'));
        const rng = getRange(); if (!rng) return;
        contentGrid.querySelectorAll('td').forEach(td => {
            const r = parseInt(td.dataset.row), c = parseInt(td.dataset.col);
            const m = getMerge(r, c); const rs = m ? m.rowspan : 1, cs = m ? m.colspan : 1;
            if (r + rs > rng.minR && r <= rng.maxR && c + cs > rng.minC && c <= rng.maxC) td.classList.add('selected');
        });
    }

    // ─── TOOLBAR SYNC ────────────────────────────────────────
    function syncToolbarFromCell(r, c) {
        const s = getCS(r, c);
        tbFont.value = s.fontFamily || gFont.value;
        tbFontSize.value = s.fontSize || gFontSize.value;
        tbTextColor.value = s.textColor || (r === 0 ? gHdrTextColor.value : gTextColor.value);
        tbFillColor.value = s.fillColor || (r === 0 ? gHeaderFill.value : gBodyFill.value);
        document.getElementById('tb-text-color-bar').style.background = tbTextColor.value;
        document.getElementById('tb-fill-color-bar').style.background = tbFillColor.value;
        ['bold', 'italic', 'underline', 'strike'].forEach(k => {
            document.getElementById(`tb-${k}`).classList.toggle('active', !!s[k === 'strike' ? 'strikethrough' : k]);
        });
        updateDDActive('dd-halign', s.hAlign || 'center');
        updateDDActive('dd-valign', s.vAlign || 'middle');
        updateDDActive('dd-wrap', s.wrap ? 'true' : 'false');
        updateDDActive('dd-rotation', String(s.rotation || 0));
    }

    function updateDDActive(id, val) {
        document.getElementById(id).querySelectorAll('.tb-popup-item').forEach(i => i.classList.toggle('active', i.dataset.val === val));
    }

    function applyToSel(prop, val) {
        const rng = getRange(); if (!rng) return;
        const rows = parseInt(tableRowsInput.value), cols = parseInt(tableColsInput.value);
        ensureStyles(rows, cols);
        for (let r = rng.minR; r <= rng.maxR; r++)
            for (let c = rng.minC; c <= rng.maxC; c++)
                if (!isHidden(r, c)) cellStyles[r][c][prop] = val;
        contentGrid.querySelectorAll('td.selected').forEach(td => applyCellVisual(td, parseInt(td.dataset.row), parseInt(td.dataset.col)));
    }

    function toggleSel(prop) {
        if (!selStart) return;
        applyToSel(prop, !getCS(selStart.row, selStart.col)[prop]);
        const m = { bold: 'tb-bold', italic: 'tb-italic', underline: 'tb-underline', strikethrough: 'tb-strike', wrap: 'tb-wrap' };
        if (m[prop]) document.getElementById(m[prop]).classList.toggle('active');
    }

    // ─── TOOLBAR EVENTS ──────────────────────────────────────
    document.getElementById('tb-bold').addEventListener('click', () => toggleSel('bold'));
    document.getElementById('tb-italic').addEventListener('click', () => toggleSel('italic'));
    document.getElementById('tb-underline').addEventListener('click', () => toggleSel('underline'));
    document.getElementById('tb-strike').addEventListener('click', () => toggleSel('strikethrough'));

    tbFont.addEventListener('change', () => applyToSel('fontFamily', tbFont.value));
    tbFontSize.addEventListener('change', () => applyToSel('fontSize', parseFloat(tbFontSize.value)));
    document.querySelector('.tb-size-dec').addEventListener('click', () => { tbFontSize.value = Math.max(6, parseInt(tbFontSize.value) - 1); applyToSel('fontSize', parseFloat(tbFontSize.value)); });
    document.querySelector('.tb-size-inc').addEventListener('click', () => { tbFontSize.value = Math.min(72, parseInt(tbFontSize.value) + 1); applyToSel('fontSize', parseFloat(tbFontSize.value)); });

    tbTextColor.addEventListener('input', () => { document.getElementById('tb-text-color-bar').style.background = tbTextColor.value; applyToSel('textColor', tbTextColor.value); });
    tbFillColor.addEventListener('input', () => { document.getElementById('tb-fill-color-bar').style.background = tbFillColor.value; applyToSel('fillColor', tbFillColor.value); });

    // Dropdowns
    function setupDD(id, prop, xform) {
        const dd = document.getElementById(id);
        dd.querySelector('.tb-dd-trigger').addEventListener('click', e => {
            e.stopPropagation();
            document.querySelectorAll('.tb-dropdown.open').forEach(d => { if (d !== dd) d.classList.remove('open'); });
            dd.classList.toggle('open');
        });
        dd.querySelectorAll('.tb-popup-item').forEach(item => {
            item.addEventListener('click', e => {
                e.stopPropagation();
                applyToSel(prop, xform ? xform(item.dataset.val) : item.dataset.val);
                dd.querySelectorAll('.tb-popup-item').forEach(i => i.classList.remove('active'));
                item.classList.add('active'); dd.classList.remove('open');
            });
        });
    }
    setupDD('dd-halign', 'hAlign');
    setupDD('dd-valign', 'vAlign');
    setupDD('dd-wrap', 'wrap', v => v === 'true');
    setupDD('dd-rotation', 'rotation', v => parseFloat(v));
    document.addEventListener('click', () => document.querySelectorAll('.tb-dropdown.open').forEach(d => d.classList.remove('open')));

    [gFont, gFontSize, gTextColor, gHdrTextColor, gHeaderFill, gBodyFill].forEach(el => {
        el.addEventListener('input', refreshAll); el.addEventListener('change', refreshAll);
    });
    function refreshAll() { contentGrid.querySelectorAll('td').forEach(td => applyCellVisual(td, parseInt(td.dataset.row), parseInt(td.dataset.col))); }

    // ─── MERGE/BREAK ─────────────────────────────────────────
    document.getElementById('btn-merge').addEventListener('click', () => {
        const rng = getRange(); if (!rng) return;
        const rs = rng.maxR - rng.minR + 1, cs = rng.maxC - rng.minC + 1; if (rs === 1 && cs === 1) return;
        merges = merges.filter(m => m.row > rng.maxR || m.row + m.rowspan - 1 < rng.minR || m.col > rng.maxC || m.col + m.colspan - 1 < rng.minC);
        storeData(); merges.push({ row: rng.minR, col: rng.minC, rowspan: rs, colspan: cs }); buildGrid();
    });
    document.getElementById('btn-break').addEventListener('click', () => {
        if (!selStart) return; const idx = merges.findIndex(m => m.row === selStart.row && m.col === selStart.col);
        if (idx === -1) return; storeData(); merges.splice(idx, 1); buildGrid();
    });

    // ─── GRID TOOLBAR ────────────────────────────────────────
    document.getElementById('btn-add-row').addEventListener('click', () => { storeData(); tableRowsInput.value = parseInt(tableRowsInput.value) + 1; ensureDimArrays(); buildGrid(); });
    document.getElementById('btn-remove-row').addEventListener('click', () => { if (parseInt(tableRowsInput.value) > 1) { storeData(); tableRowsInput.value = parseInt(tableRowsInput.value) - 1; merges = merges.filter(m => m.row + m.rowspan <= parseInt(tableRowsInput.value) + 1); ensureDimArrays(); buildGrid(); } });
    document.getElementById('btn-add-col').addEventListener('click', () => { storeData(); tableColsInput.value = parseInt(tableColsInput.value) + 1; ensureDimArrays(); buildGrid(); });
    document.getElementById('btn-remove-col').addEventListener('click', () => { if (parseInt(tableColsInput.value) > 1) { storeData(); tableColsInput.value = parseInt(tableColsInput.value) - 1; merges = merges.filter(m => m.col + m.colspan <= parseInt(tableColsInput.value) + 1); ensureDimArrays(); buildGrid(); } });

    // ─── SUBMIT ──────────────────────────────────────────────
    document.getElementById('btn-submit').addEventListener('click', async () => {
        storeData(); ensureDimArrays();
        const payload = {
            id: tableIdInput.value, label: tableLabelInput.value,
            rows: parseInt(tableRowsInput.value), cols: parseInt(tableColsInput.value),
            cell_width: parseFloat(document.getElementById('cell-width').value),
            cell_height: parseFloat(document.getElementById('cell-height').value),
            col_widths: colWidths, row_heights: rowHeights,
            border_width: parseFloat(document.getElementById('border-width').value),
            border_color: document.getElementById('border-color').value,
            header_fill: gHeaderFill.value, body_fill: gBodyFill.value,
            font_family: gFont.value, font_size: parseFloat(gFontSize.value),
            text_color: gTextColor.value, header_text_color: gHdrTextColor.value,
            data: window._tableData, merges, cell_styles: cellStyles
        };
        document.getElementById('btn-submit').disabled = true;
        statusContainer.classList.remove('hidden'); updateProg(5, 'Submitting...');
        try {
            const res = await fetch('/submit', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
            if (res.ok) pollStatus(); else { updateProg(0, 'Failed.', 'error'); document.getElementById('btn-submit').disabled = false; }
        } catch (err) { updateProg(0, err.message, 'error'); document.getElementById('btn-submit').disabled = false; }
    });
    editorForm.addEventListener('submit', e => e.preventDefault());

    // ─── NAV ─────────────────────────────────────────────────
    document.getElementById('btn-new-table').addEventListener('click', () => {
        editorForm.reset(); tableIdInput.value = ''; tableRowsInput.value = 4; tableColsInput.value = 4;
        window._tableData = []; merges = []; cellStyles = []; colWidths = []; rowHeights = [];
        showEditor('New Table'); buildGrid();
    });
    document.getElementById('btn-back').addEventListener('click', () => { window._tableData = []; merges = []; cellStyles = []; colWidths = []; rowHeights = []; showLanding(); });
    document.getElementById('btn-cancel-editor').addEventListener('click', () => { window._tableData = []; merges = []; cellStyles = []; colWidths = []; rowHeights = []; showLanding(); });
    document.getElementById('btn-close-landing').addEventListener('click', async () => { try { await fetch('/close', { method: 'POST' }); } catch (e) { } });

    // ─── PROGRESS ────────────────────────────────────────────
    function updateProg(p, m, s = 'normal') { progressBar.style.width = p + '%'; statusMessage.textContent = m; progressBar.classList.toggle('error', s === 'error'); }
    async function pollStatus() {
        if (pollInterval) clearInterval(pollInterval);
        pollInterval = setInterval(async () => {
            try {
                const s = await (await fetch('/status')).json();
                updateProg(s.progress, s.message, s.status === 'error' ? 'error' : 'normal');
                if (s.status === 'completed' || s.status === 'error') { clearInterval(pollInterval); document.getElementById('btn-submit').disabled = false; }
            }
            catch (e) { clearInterval(pollInterval); document.getElementById('btn-submit').disabled = false; }
        }, 500);
    }

    loadFonts(); loadTableList();
});
