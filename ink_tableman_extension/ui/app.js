document.addEventListener('DOMContentLoaded', () => {
    const screenLanding = document.getElementById('screen-landing');
    const screenEditor = document.getElementById('screen-editor');
    const tableList = document.getElementById('table-list');
    const editorTitle = document.getElementById('editor-title');
    const editorForm = document.getElementById('editor-form');
    const tableIdInput = document.getElementById('table-id');
    const themeSelect = document.getElementById('theme-select');
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
        hAlign: 'center', vAlign: 'middle', wrap: false, rotation: 0,
        format: 'auto', decimals: 2, currencySymbol: null
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

    function formatValue(val, format, decs = 2, symbol = '$') {
        if (!val || format === 'text' || format === 'auto' && isNaN(val)) return val;
        const num = parseFloat(val.toString().replace(/[^0-9.-]/g, ''));
        if (isNaN(num)) return val;

        try {
            const locale = 'id-ID';
            switch (format) {
                case 'number': return new Intl.NumberFormat(locale, { minimumFractionDigits: decs, maximumFractionDigits: decs }).format(num);
                case 'percent': return new Intl.NumberFormat(locale, { style: 'percent', minimumFractionDigits: decs, maximumFractionDigits: decs }).format(num / 100);
                case 'currency': return new Intl.NumberFormat(locale, { style: 'currency', currency: 'IDR', minimumFractionDigits: decs }).format(num);
                case 'currency_usd': return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: decs }).format(num);
                case 'currency_eur': return new Intl.NumberFormat('de-DE', { style: 'currency', currency: 'EUR', minimumFractionDigits: decs }).format(num);
                case 'currency_custom':
                    const n = new Intl.NumberFormat(locale, { minimumFractionDigits: decs, maximumFractionDigits: decs }).format(num);
                    return `${symbol || '$'}${n}`;
                case 'currency_round': return new Intl.NumberFormat(locale, { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(num);
                case 'scientific': return num.toExponential(decs);
                case 'date': return isNaN(Date.parse(val)) ? val : new Intl.DateTimeFormat(locale).format(new Date(val));
                case 'date_iso': return isNaN(Date.parse(val)) ? val : new Date(val).toISOString().split('T')[0];
                case 'date_us': return isNaN(Date.parse(val)) ? val : new Intl.DateTimeFormat('en-US').format(new Date(val));
                case 'date_long': return isNaN(Date.parse(val)) ? val : new Intl.DateTimeFormat(locale, { dateStyle: 'long' }).format(new Date(val));
                case 'time': return isNaN(Date.parse(val)) ? val : new Intl.DateTimeFormat(locale, { hour: 'numeric', minute: 'numeric', second: 'numeric' }).format(new Date(val));
                case 'datetime': return isNaN(Date.parse(val)) ? val : new Intl.DateTimeFormat(locale, { dateStyle: 'short', timeStyle: 'short' }).format(new Date(val));
                default: return val;
            }
        } catch (e) { return val; }
    }

    function ensureDimArrays() {
        const rows = parseInt(tableRowsInput.value) || 3, cols = parseInt(tableColsInput.value) || 3;
        const defW = parseFloat(document.getElementById('cell-width').value) || 120;
        const defH = parseFloat(document.getElementById('cell-height').value) || 40;
        while (colWidths.length < cols) colWidths.push(defW);
        while (rowHeights.length < rows) rowHeights.push(defH);
        colWidths.length = cols; rowHeights.length = rows;
    }

    // ─── NAV ─────────────────────────────────────────────────
    function showLanding() { screenEditor.classList.add('hidden'); screenLanding.classList.remove('hidden'); loadTableList(); }

    // About Modal Logic
    document.getElementById('btn-show-about').addEventListener('click', () => document.getElementById('about-modal').classList.remove('hidden'));
    document.getElementById('btn-close-about').addEventListener('click', () => document.getElementById('about-modal').classList.add('hidden'));
    window.addEventListener('click', e => { if (e.target.id === 'about-modal') document.getElementById('about-modal').classList.add('hidden'); });
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

            // Fills & Transparency
            const setFill = (val, input, btnId) => {
                const btn = document.getElementById(btnId);
                if (val === 'none') {
                    btn.classList.add('active');
                } else {
                    btn.classList.remove('active');
                    input.value = val;
                }
            };
            setFill(d.header_fill, gHeaderFill, 'btn-none-header');
            setFill(d.body_fill, gBodyFill, 'btn-none-body');

            // Banded Rows
            const brCheck = document.getElementById('banded-rows');
            const brColor = document.getElementById('banded-color');
            brCheck.checked = !!d.banded_rows;
            setFill(d.banded_color || '#2a2a2a', brColor, 'btn-none-banded');

            gFontSize.value = d.font_size; gTextColor.value = d.text_color; gHdrTextColor.value = d.header_text_color;
            await loadFonts(d.font_family);
            window._tableData = d.data || []; merges = d.merges || []; cellStyles = d.cell_styles || [];
            colWidths = d.col_widths || []; rowHeights = d.row_heights || [];
            showEditor(`Edit: ${d.label}`); buildGrid();
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
        const rows = parseInt(tableRowsInput.value) || 3, cols = parseInt(tableColsInput.value) || 3;
        const data = window._tableData || [];
        ensureStyles(rows, cols); ensureDimArrays();

        // Build colgroup
        let cg = gridTable.querySelector('colgroup'); if (cg) cg.remove();
        cg = document.createElement('colgroup');
        for (let c = 0; c < cols; c++) {
            const col = document.createElement('col');
            col.style.width = colWidths[c] + 'px';
            cg.appendChild(col);
        }
        gridTable.prepend(cg);

        contentGrid.innerHTML = ''; clearSel();

        for (let r = 0; r < rows; r++) {
            const tr = document.createElement('tr');
            tr.style.height = rowHeights[r] + 'px';
            for (let c = 0; c < cols; c++) {
                if (isHidden(r, c)) continue;
                const td = document.createElement('td');
                td.dataset.row = r; td.dataset.col = c;

                const mg = getMerge(r, c);
                if (mg) { td.rowSpan = mg.rowspan; td.colSpan = mg.colspan; td.classList.add('merged-cell'); }

                // Inner div for rotation support
                const inner = document.createElement('div');
                inner.className = 'cell-inner';
                inner.contentEditable = true;

                if (r < data.length && c < data[r].length) {
                    const s = getCS(r, c);
                    inner.textContent = formatValue(data[r][c], s.format, s.decimals, s.currencySymbol);
                } else if (r === 0) inner.textContent = `Col ${c + 1}`;

                td.appendChild(inner);
                applyCellVisual(td, inner, r, c);

                inner.addEventListener('focus', () => {
                    const raw = (r < data.length && c < data[r].length) ? data[r][c] : '';
                    inner.textContent = raw;
                });

                inner.addEventListener('blur', () => {
                    const s = getCS(r, c);
                    const val = inner.textContent;
                    if (r >= data.length) while (data.length <= r) data.push([]);
                    data[r][c] = val;
                    inner.textContent = formatValue(val, s.format, s.decimals, s.currencySymbol);
                });

                td.addEventListener('mousedown', e => {
                    if (e.target.classList.contains('cell-inner')) return; // let inner handle focus
                    if (e.shiftKey && selStart) selEnd = { row: r, col: c };
                    else { selStart = { row: r, col: c }; selEnd = { row: r, col: c }; }
                    highlightSel(); syncToolbarFromCell(r, c);
                });

                inner.addEventListener('mousedown', e => {
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

    function applyCellVisual(td, inner, r, c) {
        const s = getCS(r, c); const isH = r === 0;

        // Text styles on inner
        inner.style.fontFamily = s.fontFamily || gFont.value;
        inner.style.fontSize = (s.fontSize || gFontSize.value) + 'px';
        inner.style.color = s.textColor || (isH ? gHdrTextColor.value : gTextColor.value);
        inner.style.fontWeight = s.bold ? 'bold' : 'normal';
        inner.style.fontStyle = s.italic ? 'italic' : 'normal';
        inner.style.textDecoration = [s.underline ? 'underline' : '', s.strikethrough ? 'line-through' : ''].filter(Boolean).join(' ') || 'none';

        // Alignment on td and flex inner
        const hMap = { 'left': 'flex-start', 'center': 'center', 'right': 'flex-end' };
        const vMap = { 'top': 'flex-start', 'middle': 'center', 'bottom': 'flex-end' };
        const ha = s.hAlign || 'center';
        const va = s.vAlign || 'middle';

        td.style.textAlign = ha;
        td.style.verticalAlign = va;
        inner.style.justifyContent = hMap[ha] || 'center';
        inner.style.alignItems = vMap[va] || 'center';

        // Wrap
        inner.style.whiteSpace = s.wrap ? 'normal' : 'nowrap';
        inner.style.wordBreak = s.wrap ? 'break-word' : 'normal';
        inner.style.textAlign = ha;

        // Fill on td
        let fill = s.fillColor;
        if (!fill) {
            const isBanded = document.getElementById('banded-rows').checked;
            const bColor = document.getElementById('banded-color').value;
            const bNone = document.getElementById('btn-none-banded').classList.contains('active');

            if (isH) {
                fill = document.getElementById('btn-none-header').classList.contains('active') ? 'none' : gHeaderFill.value;
            } else {
                const bodyNone = document.getElementById('btn-none-body').classList.contains('active');
                if (isBanded && r % 2 === 0) {
                    fill = bNone ? 'none' : bColor;
                } else {
                    fill = bodyNone ? 'none' : gBodyFill.value;
                }
            }
        }
        td.style.backgroundColor = (fill === 'none') ? 'transparent' : fill;
        if (fill === 'none') td.dataset.fillNone = 'true'; else delete td.dataset.fillNone;

        // Rotation on inner div — like Google Sheets
        const rot = String(s.rotation || 0);
        inner.style.transform = '';
        inner.style.writingMode = '';
        inner.style.textOrientation = '';
        inner.style.display = '';
        inner.style.width = '';
        inner.style.lineHeight = '';
        inner.style.letterSpacing = '';
        inner.classList.remove('rot-stack');

        if (rot === 'stack') {
            // Stack vertical: each character on its own line
            inner.classList.add('rot-stack');
        } else if (rot === '90') {
            // Rotate down — text reads top to bottom (like Google Sheets)
            inner.style.writingMode = 'vertical-rl';
            inner.style.textOrientation = 'mixed';
        } else if (rot === '-90') {
            // Rotate up — text reads bottom to top (like Google Sheets)
            inner.style.writingMode = 'vertical-rl';
            inner.style.textOrientation = 'mixed';
            inner.style.transform = 'rotate(180deg)';
        } else if (rot === '45') {
            inner.style.transform = 'rotate(-45deg)';
            inner.style.display = 'inline-block';
        } else if (rot === '-45') {
            inner.style.transform = 'rotate(45deg)';
            inner.style.display = 'inline-block';
        }
    }

    function collectData() {
        const rows = parseInt(tableRowsInput.value) || 3, cols = parseInt(tableColsInput.value) || 3;
        const data = []; for (let r = 0; r < rows; r++) data.push(new Array(cols).fill(''));
        contentGrid.querySelectorAll('td').forEach(td => {
            const r = parseInt(td.dataset.row), c = parseInt(td.dataset.col);
            const inner = td.querySelector('.cell-inner');
            if (r < rows && c < cols) data[r][c] = inner ? inner.textContent : td.textContent;
        });
        return data;
    }
    function storeData() { window._tableData = collectData(); }

    // ─── RESIZE HANDLES ──────────────────────────────────────
    function addResizeHandles() {
        document.querySelectorAll('.col-resize-handle,.row-resize-handle').forEach(h => h.remove());
        const wrapper = document.getElementById('grid-inner-rel');
        const cols = parseInt(tableColsInput.value) || 3, rows = parseInt(tableRowsInput.value) || 3;

        let xOff = 0;
        for (let c = 0; c < cols - 1; c++) {
            xOff += colWidths[c];
            const h = document.createElement('div'); h.className = 'col-resize-handle';
            h.style.left = (xOff - 2) + 'px'; h.style.top = '0'; h.style.height = gridTable.offsetHeight + 'px';
            h.dataset.col = c; wrapper.appendChild(h);
            h.addEventListener('mousedown', e => {
                e.preventDefault(); const sx = e.clientX, sw = colWidths[c];
                const mv = ev => { colWidths[c] = Math.max(30, sw + (ev.clientX - sx)); buildGrid(); };
                const up = () => { document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up); };
                document.addEventListener('mousemove', mv); document.addEventListener('mouseup', up);
            });
        }

        let yOff = 0;
        for (let r = 0; r < rows - 1; r++) {
            yOff += rowHeights[r];
            const h = document.createElement('div'); h.className = 'row-resize-handle';
            h.style.top = (yOff - 2) + 'px'; h.style.left = '0'; h.style.width = gridTable.offsetWidth + 'px';
            h.dataset.row = r; wrapper.appendChild(h);
            h.addEventListener('mousedown', e => {
                e.preventDefault(); const sy = e.clientY, sh = rowHeights[r];
                const mv = ev => { rowHeights[r] = Math.max(20, sh + (ev.clientY - sy)); buildGrid(); };
                const up = () => { document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up); };
                document.addEventListener('mousemove', mv); document.addEventListener('mouseup', up);
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
        updateDDActive('dd-wrap', s.wrap ? 'true' : 'false');
        updateDDActive('dd-rotation', String(s.rotation || 0));
        updateDDActive('dd-number-format', s.format || 'auto');
    }

    function updateDDActive(id, val) {
        document.getElementById(id).querySelectorAll('.tb-popup-item').forEach(i => i.classList.toggle('active', i.dataset.val === val));
    }

    function applyToSel(prop, val) {
        const rng = getRange(); if (!rng) return;
        ensureStyles(parseInt(tableRowsInput.value), parseInt(tableColsInput.value));
        for (let r = rng.minR; r <= rng.maxR; r++)
            for (let c = rng.minC; c <= rng.maxC; c++)
                if (!isHidden(r, c)) cellStyles[r][c][prop] = val;

        // Re-apply visuals
        contentGrid.querySelectorAll('td.selected').forEach(td => {
            const inner = td.querySelector('.cell-inner');
            const r = parseInt(td.dataset.row), c = parseInt(td.dataset.col);
            if (inner) {
                applyCellVisual(td, inner, r, c);
                // Also update text if not focused to show new format
                if (document.activeElement !== inner) {
                    const data = window._tableData || [];
                    const val = (r < data.length && c < data[r].length) ? data[r][c] : '';
                    const s = getCS(r, c);
                    inner.textContent = formatValue(val, s.format, s.decimals, s.currencySymbol);
                }
            }
        });
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

    // Formatting buttons
    document.getElementById('btn-fmt-currency').addEventListener('click', () => applyToSel('format', 'currency'));
    document.getElementById('btn-fmt-percent').addEventListener('click', () => applyToSel('format', 'percent'));
    document.getElementById('btn-fmt-dec-less').addEventListener('click', () => {
        if (!selStart) return;
        const s = getCS(selStart.row, selStart.col);
        applyToSel('decimals', Math.max(0, (s.decimals || 0) - 1));
    });
    document.getElementById('btn-fmt-dec-more').addEventListener('click', () => {
        if (!selStart) return;
        const s = getCS(selStart.row, selStart.col);
        applyToSel('decimals', Math.min(10, (s.decimals || 0) + 1));
    });

    // Dropdowns
    setupDD('dd-halign', 'hAlign');
    setupDD('dd-valign', 'vAlign');
    setupDD('dd-wrap', 'wrap', v => v === 'true');
    setupDD('dd-rotation', 'rotation', v => v === 'stack' ? 'stack' : parseInt(v));
    setupDD('dd-number-format', 'format');

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
                if (item.dataset.val === 'currency_custom') {
                    const old = (selStart && getCS(selStart.row, selStart.col).currencySymbol) || '$';
                    const sym = prompt('Enter custom currency symbol:', old);
                    if (sym !== null) applyToSel('currencySymbol', sym);
                    else return;
                }
                applyToSel(prop, xform ? xform(item.dataset.val) : item.dataset.val);
                dd.querySelectorAll('.tb-popup-item').forEach(i => i.classList.remove('active'));
                item.classList.add('active'); dd.classList.remove('open');
            });
        });
    }
    setupDD('dd-rotation', 'rotation', v => v === 'stack' ? 'stack' : parseFloat(v));
    document.addEventListener('click', () => document.querySelectorAll('.tb-dropdown.open').forEach(d => d.classList.remove('open')));

    [gFont, gFontSize, gTextColor, gHdrTextColor, gHeaderFill, gBodyFill,
        document.getElementById('banded-rows'), document.getElementById('banded-color'),
        document.getElementById('header-fill'), document.getElementById('body-fill')].forEach(el => {
            if (!el) return;
            el.addEventListener('input', refreshAll); el.addEventListener('change', refreshAll);
        });

    // None (Transparent) Toggles
    function setupNone(btnId, colorId) {
        const btn = document.getElementById(btnId);
        const color = document.getElementById(colorId);
        btn.addEventListener('click', () => {
            btn.classList.toggle('active');
            refreshAll();
        });
    }
    setupNone('btn-none-header', 'header-fill');
    setupNone('btn-none-body', 'body-fill');
    setupNone('btn-none-banded', 'banded-color');
    function refreshAll() {
        contentGrid.querySelectorAll('td').forEach(td => {
            const inner = td.querySelector('.cell-inner');
            if (inner) applyCellVisual(td, inner, parseInt(td.dataset.row), parseInt(td.dataset.col));
        });
    }

    // ─── MERGE/BREAK ─────────────────────────────────────────
    document.getElementById('btn-merge').addEventListener('click', () => {
        const rng = getRange(); if (!rng) return;
        const rs = rng.maxR - rng.minR + 1, cs = rng.maxC - rng.minC + 1; if (rs === 1 && cs === 1) return;
        merges = merges.filter(m => m.row > rng.maxR || m.row + m.rowspan - 1 < rng.minR || m.col > rng.maxC || m.col + m.colspan - 1 < rng.minC);
        storeData(); merges.push({ row: rng.minR, col: rng.minC, rowspan: rs, colspan: cs }); buildGrid();
    });
    document.getElementById('btn-break').addEventListener('click', () => {
        if (!selStart) return;
        const idx = merges.findIndex(m => m.row === selStart.row && m.col === selStart.col);
        if (idx === -1) return; storeData(); merges.splice(idx, 1); buildGrid();
    });

    // ─── GRID MUTATION ───────────────────────────────────────
    function insertRow(at) {
        storeData();
        const rows = parseInt(tableRowsInput.value), cols = parseInt(tableColsInput.value);
        if (!window._tableData || !window._tableData.length) window._tableData = collectData();
        window._tableData.splice(at, 0, new Array(cols).fill(''));
        cellStyles.splice(at, 0, Array.from({ length: cols }, () => defCS()));
        rowHeights.splice(at, 0, rowHeights[at] || rowHeights[rowHeights.length - 1] || 40);
        merges.forEach(m => {
            if (m.row >= at) m.row++;
            else if (m.row < at && m.row + m.rowspan > at) m.rowspan++;
        });
        tableRowsInput.value = window._tableData.length;
        buildGrid();
    }

    function deleteRow(at) {
        if (parseInt(tableRowsInput.value) <= 1) return;
        storeData();
        window._tableData.splice(at, 1);
        cellStyles.splice(at, 1);
        rowHeights.splice(at, 1);
        merges = merges.filter(m => !(m.row === at && m.rowspan === 1)).map(m => {
            if (m.row > at) m.row--;
            else if (m.row <= at && m.row + m.rowspan > at) m.rowspan--;
            return m;
        }).filter(m => m.rowspan > 0);
        tableRowsInput.value = window._tableData.length;
        buildGrid();
    }

    function insertCol(at) {
        storeData();
        const rows = parseInt(tableRowsInput.value), cols = parseInt(tableColsInput.value);
        if (!window._tableData || !window._tableData.length) window._tableData = collectData();
        window._tableData.forEach(r => r.splice(at, 0, ''));
        cellStyles.forEach(r => r.splice(at, 0, defCS()));
        colWidths.splice(at, 0, colWidths[at] || colWidths[colWidths.length - 1] || 120);
        merges.forEach(m => {
            if (m.col >= at) m.col++;
            else if (m.col < at && m.col + m.colspan > at) m.colspan++;
        });
        tableColsInput.value = window._tableData[0].length;
        buildGrid();
    }

    function deleteCol(at) {
        if (parseInt(tableColsInput.value) <= 1) return;
        storeData();
        window._tableData.forEach(r => r.splice(at, 1));
        cellStyles.forEach(r => r.splice(at, 1));
        colWidths.splice(at, 1);
        merges = merges.filter(m => !(m.col === at && m.colspan === 1)).map(m => {
            if (m.col > at) m.col--;
            else if (m.col <= at && m.col + m.colspan > at) m.colspan--;
            return m;
        }).filter(m => m.colspan > 0);
        tableColsInput.value = window._tableData[0].length;
        buildGrid();
    }

    function setupGridDD(id, type) {
        const dd = document.getElementById(id);
        if (!dd) return;
        dd.querySelector('.tb-dd-trigger').addEventListener('click', e => {
            e.stopPropagation();
            document.querySelectorAll('.tb-dropdown.open').forEach(d => { if (d !== dd) d.classList.remove('open'); });
            dd.classList.toggle('open');
        });
        dd.querySelectorAll('.tb-popup-item').forEach(item => {
            item.addEventListener('click', () => {
                const val = item.dataset.val;
                const r = selStart ? selStart.row : 0;
                const c = selStart ? selStart.col : 0;
                if (type === 'row') {
                    if (val === 'above') insertRow(r);
                    else if (val === 'below') insertRow(r + 1);
                    else if (val === 'start') insertRow(0);
                    else if (val === 'end') insertRow(parseInt(tableRowsInput.value));
                    else if (val === 'delete') deleteRow(r);
                } else {
                    if (val === 'left') insertCol(c);
                    else if (val === 'right') insertCol(c + 1);
                    else if (val === 'start') insertCol(0);
                    else if (val === 'end') insertCol(parseInt(tableColsInput.value));
                    else if (val === 'delete') deleteCol(c);
                }
                dd.classList.remove('open');
            });
        });
    }
    setupGridDD('dd-row-actions', 'row');
    setupGridDD('dd-col-actions', 'col');

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
            header_fill: document.getElementById('btn-none-header').classList.contains('active') ? 'none' : gHeaderFill.value,
            body_fill: document.getElementById('btn-none-body').classList.contains('active') ? 'none' : gBodyFill.value,
            banded_rows: document.getElementById('banded-rows').checked,
            banded_color: document.getElementById('btn-none-banded').classList.contains('active') ? 'none' : document.getElementById('banded-color').value,
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
            } catch (e) { clearInterval(pollInterval); document.getElementById('btn-submit').disabled = false; }
        }, 500);
    }

    // ─── THEME ────────────────────────────────────────────────
    function applyTheme(theme) {
        document.documentElement.setAttribute('data-theme', theme);
    }

    async function initTheme() {
        try {
            const res = await fetch('/settings');
            const s = await res.json();
            const theme = s.theme || 'system';
            themeSelect.value = theme;
            applyTheme(theme);
        } catch (e) { applyTheme('system'); }
    }

    themeSelect.addEventListener('change', async () => {
        const val = themeSelect.value;
        applyTheme(val);
        try {
            await fetch('/settings', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ theme: val })
            });
        } catch (e) { }
    });

    initTheme();

    loadFonts(); loadTableList();
});
