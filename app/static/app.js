/* SmartDoc & Code Review v7.0 — only .py and .docx, no AI */
(() => {
    'use strict';
    console.log('[SmartDoc v7.4] JS loaded');
    const $ = id => document.getElementById(id);
    const dropzone = $('dropzone'), fileInput = $('file-input'), browseButton = $('browse-button');
    const uploadIdle = $('upload-idle'), uploadLoading = $('upload-loading'), loadingFilename = $('loading-filename');
    const reportSection = $('report-section'), reportType = $('report-type');
    const reportFilename = $('report-filename'), reportSubtitle = $('report-subtitle');
    const reportBody = $('report-body'), resetButton = $('reset-button');
    const heroSection = $('upload');
    function setReportVisible(v) { reportSection.hidden = !v; if (heroSection) heroSection.hidden = v; }
    const autofixButton = $('autofix-button'), pdfButton = $('pdf-button');
    const statTotal = $('stat-total'), statHigh = $('stat-high');
    const statMedium = $('stat-medium'), statLow = $('stat-low');
    const verdictCard = $('verdict-card'), verdictIcon = $('verdict-icon');
    const verdictTitle = $('verdict-title'), verdictSub = $('verdict-sub');
    const sevBarHigh = $('sev-bar-high'), sevBarMedium = $('sev-bar-medium'), sevBarLow = $('sev-bar-low');
    const sourceViewer = $('source-viewer'), sourceCode = $('source-code'), sourceMinimap = $('source-minimap');
    const themeToggle = $('theme-toggle');
    const historyList = $('history-list'), historyEmpty = $('history-empty'), clearHistoryBtn = $('clear-history-button');

    const ALLOWED = ['.py', '.docx'];
    const HISTORY_KEY = 'smartdoc-history';

    const VERDICTS = {
        good: { icon: '\u2713', title: 'Всё в порядке',            sub: 'Замечаний по оформлению нет', cls: 'stat--verdict-good' },
        ok:   { icon: '!',      title: 'Нужны небольшие правки',  sub: 'Критичных замечаний нет или одно', cls: 'stat--verdict-ok' },
        bad:  { icon: '\u2716', title: 'Нужно доработать',        sub: 'Есть критичные замечания', cls: 'stat--verdict-bad' },
    };
    function setVerdict(key) {
        const v = VERDICTS[key] || VERDICTS.bad;
        verdictCard.classList.remove('stat--verdict-good','stat--verdict-ok','stat--verdict-bad');
        verdictCard.classList.add(v.cls);
        verdictIcon.textContent = v.icon;
        verdictTitle.textContent = v.title;
        verdictSub.textContent = v.sub;
    }
    function resetVerdict() {
        verdictCard.classList.remove('stat--verdict-good','stat--verdict-ok','stat--verdict-bad');
        verdictIcon.textContent = '?';
        verdictTitle.textContent = '\u2014';
        verdictSub.textContent = '';
    }
    const RULES_KEY = 'smartdoc-docx-rules';
    let currentFile = null, currentReport = null;

    // ─── DOCX Rules ───
    const rulesToggle = $('rules-toggle'), rulesModal = $('rules-modal');
    const rulesModalClose = $('rules-modal-close'), rulesSave = $('rules-save');
    const presetChips = $('preset-chips'), rulesFields = $('rules-fields');
    const rulesPresetLabel = $('rules-preset-label');
    const ruleInputs = {
        font: $('rule-font'), fontSize: $('rule-font-size'),
        spacing: $('rule-spacing'), indent: $('rule-indent'),
        alignment: $('rule-alignment'), bibName: $('rule-bib-name'),
        mLeft: $('rule-m-left'), mRight: $('rule-m-right'),
        mTop: $('rule-m-top'), mBottom: $('rule-m-bottom'),
    };
    const ruleChecks = {
        headings: $('rule-check-headings'),
        pageNumbers: $('rule-check-page-numbers'),
        toc: $('rule-check-toc'),
        bibliography: $('rule-check-bibliography'),
        hyperlinks: $('rule-check-hyperlinks'),
        textColor: $('rule-check-text-color'),
        tables: $('rule-check-tables'),
        spaces: $('rule-check-spaces'),
        blankLines: $('rule-check-blank-lines'),
    };

    const PRESETS = {
        agu: {
            font_name:'Times New Roman', font_size_pt:14, line_spacing:1.5,
            first_line_indent_cm:1.25, alignment:'justify', bib_name:'any',
            margins_cm:{left:3.5,right:1,top:2.5,bottom:2.5},
            checks:{headings:true,pageNumbers:true,toc:true,bibliography:true,
                    hyperlinks:true,textColor:true,tables:true,spaces:true,blankLines:true},
            label:'АГУ (кафедра ИТ)'
        },
    };

    let currentRules = loadRules();

    function loadRules() {
        try { const s = localStorage.getItem(RULES_KEY); return s ? JSON.parse(s) : { preset: 'agu' }; }
        catch { return { preset: 'agu' }; }
    }
    function saveRules(r) { currentRules = r; localStorage.setItem(RULES_KEY, JSON.stringify(r)); updatePresetLabel(); }

    function updatePresetLabel() {
        const p = currentRules.preset;
        if (p && PRESETS[p]) {
            rulesPresetLabel.textContent = PRESETS[p].label;
        } else {
            rulesPresetLabel.textContent = 'Свои';
        }
    }

    function fillRulesForm(r) {
        const data = r.preset && PRESETS[r.preset] ? PRESETS[r.preset] : r;
        ruleInputs.font.value = data.font_name || 'Times New Roman';
        ruleInputs.fontSize.value = data.font_size_pt || 14;
        ruleInputs.spacing.value = data.line_spacing || 1.5;
        ruleInputs.indent.value = data.first_line_indent_cm || 1.25;
        ruleInputs.alignment.value = data.alignment || 'justify';
        ruleInputs.bibName.value = data.bib_name || 'any';
        const m = data.margins_cm || {left:3.5,right:1,top:2.5,bottom:2.5};
        ruleInputs.mLeft.value = m.left; ruleInputs.mRight.value = m.right;
        ruleInputs.mTop.value = m.top; ruleInputs.mBottom.value = m.bottom;
        const ch = data.checks || {};
        Object.keys(ruleChecks).forEach(k => {
            ruleChecks[k].checked = ch[k] !== false;
        });

        const isCustom = !r.preset || !PRESETS[r.preset];
        setFieldsDisabled(!isCustom);

        presetChips.querySelectorAll('.preset-chip').forEach(ch => {
            ch.classList.toggle('is-active', isCustom ? ch.dataset.preset === 'custom' : ch.dataset.preset === r.preset);
        });
    }

    function setFieldsDisabled(disabled) {
        Object.values(ruleInputs).forEach(inp => inp.disabled = disabled);
        // Checkboxes always enabled
    }

    function readRulesForm() {
        const checks = {};
        Object.keys(ruleChecks).forEach(k => { checks[k] = ruleChecks[k].checked; });
        return {
            font_name: ruleInputs.font.value.trim(),
            font_size_pt: parseFloat(ruleInputs.fontSize.value) || 14,
            line_spacing: parseFloat(ruleInputs.spacing.value) || 1.5,
            first_line_indent_cm: parseFloat(ruleInputs.indent.value) || 1.25,
            alignment: ruleInputs.alignment.value,
            bib_name: ruleInputs.bibName.value,
            margins_cm: {
                left: parseFloat(ruleInputs.mLeft.value) || 3,
                right: parseFloat(ruleInputs.mRight.value) || 1.5,
                top: parseFloat(ruleInputs.mTop.value) || 2,
                bottom: parseFloat(ruleInputs.mBottom.value) || 2,
            },
            checks: checks,
        };
    }

    function getRulesForAPI() {
        if (currentRules.preset && PRESETS[currentRules.preset]) {
            const checks = currentRules.checks || {};
            return { preset: currentRules.preset, checks: checks };
        }
        return currentRules;
    }

    rulesToggle.addEventListener('click', (e) => {
        e.stopPropagation();
        fillRulesForm(currentRules);
        rulesModal.hidden = false;
    });
    rulesModalClose.addEventListener('click', () => { rulesModal.hidden = true; });
    rulesModal.addEventListener('click', e => { if (e.target === rulesModal) rulesModal.hidden = true; });

    presetChips.addEventListener('click', e => {
        const chip = e.target.closest('.preset-chip');
        if (!chip) return;
        const preset = chip.dataset.preset;
        presetChips.querySelectorAll('.preset-chip').forEach(c => c.classList.remove('is-active'));
        chip.classList.add('is-active');
        if (preset === 'custom') {
            setFieldsDisabled(false);
        } else {
            fillRulesForm({ preset });
        }
    });

    rulesSave.addEventListener('click', () => {
        const activeChip = presetChips.querySelector('.preset-chip.is-active');
        const presetName = activeChip ? activeChip.dataset.preset : 'custom';
        const checks = {};
        Object.keys(ruleChecks).forEach(k => { checks[k] = ruleChecks[k].checked; });
        if (presetName !== 'custom' && PRESETS[presetName]) {
            saveRules({ preset: presetName, checks: checks });
        } else {
            saveRules(readRulesForm());
        }
        rulesModal.hidden = true;
    });

    updatePresetLabel();

    // Export rules as JSON
    const rulesExport = $('rules-export');
    rulesExport.addEventListener('click', () => {
        const data = readRulesForm();
        const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = 'smartdoc_rules.json'; a.click();
        URL.revokeObjectURL(url);
    });

    // Import rules from JSON
    const rulesImport = $('rules-import');
    rulesImport.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = () => {
            try {
                const data = JSON.parse(reader.result);
                fillRulesForm(data);
                // Switch to custom
                presetChips.querySelectorAll('.preset-chip').forEach(c => c.classList.remove('is-active'));
                const customChip = presetChips.querySelector('[data-preset="custom"]');
                if (customChip) customChip.classList.add('is-active');
                setFieldsDisabled(false);
            } catch { alert('Некорректный JSON-файл'); }
        };
        reader.readAsText(file);
        e.target.value = '';
    });

    // Generate template
    const templateButton = $('template-button');
    templateButton.addEventListener('click', async () => {
        templateButton.disabled = true;
        const orig = templateButton.textContent;
        templateButton.textContent = 'Генерируем…';
        try {
            const fd = new FormData();
            fd.append('docx_rules', JSON.stringify(getRulesForAPI()));
            const r = await fetch('/api/generate-template', { method: 'POST', body: fd });
            if (!r.ok) { const d = await r.json().catch(() => ({})); throw new Error(d.detail || `Ошибка ${r.status}`); }
            const blob = await r.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url; a.download = 'Шаблон_АГУ.docx'; a.click();
            URL.revokeObjectURL(url);
        } catch (err) { alert('Не удалось: ' + err.message); }
        finally { templateButton.disabled = false; templateButton.textContent = orig; }
    });

    // ─── Scroll-reveal observer ───

    // ─── Paste Area ───
    const pasteToggle = $('paste-toggle'), pasteEditor = $('paste-editor');
    const pasteTextarea = $('paste-textarea'), pasteCheck = $('paste-check'), pasteClear = $('paste-clear');

    pasteToggle.addEventListener('click', (e) => {
        e.stopPropagation();
        const opening = pasteEditor.hidden;
        pasteEditor.hidden = !opening;
        pasteToggle.classList.toggle('is-open', opening);
        if (opening) pasteTextarea.focus();
    });

    pasteClear.addEventListener('click', () => {
        pasteTextarea.value = '';
        pasteTextarea.focus();
    });

    pasteCheck.addEventListener('click', () => {
        const code = pasteTextarea.value;
        if (!code.trim()) { alert('Вставьте код для проверки'); return; }
        // Create a virtual .py file from the pasted code
        const blob = new Blob([code], { type: 'text/x-python' });
        const file = new File([blob], 'pasted_code.py', { type: 'text/x-python' });
        handleFile(file);
    });

    // Global Ctrl+V / Cmd+V handler — if paste area is not focused, open it and paste
    document.addEventListener('paste', (e) => {
        // Don't intercept if user is in an input/textarea already
        const active = document.activeElement;
        if (active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA')) return;
        // Don't intercept if modal is open
        if (!rulesModal.hidden) return;

        const text = e.clipboardData.getData('text/plain');
        if (!text || !text.trim()) return;

        // Check if it looks like code (has newlines or common code patterns)
        const looksLikeCode = text.includes('\n') || /^\s*(import |from |def |class |if |for |while |#)/.test(text);
        if (!looksLikeCode) return;

        e.preventDefault();
        pasteEditor.hidden = false;
        pasteToggle.classList.add('is-open');
        pasteTextarea.value = text;
        pasteTextarea.focus();
        pasteTextarea.scrollIntoView({ behavior: 'smooth', block: 'center' });
    });

    // ─── Scroll-reveal observer ───
    const revealObserver = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.classList.add('is-visible');
                revealObserver.unobserve(entry.target);
            }
        });
    }, { threshold: 0.15, rootMargin: '0px 0px -40px 0px' });

    // Observe info cards
    document.querySelectorAll('.info-card').forEach(card => revealObserver.observe(card));

    // ─── Theme ───
    if (themeToggle) themeToggle.addEventListener('click', () => {
        const t = document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
        document.documentElement.setAttribute('data-theme', t);
        localStorage.setItem('smartdoc-theme', t);
    });

    // ─── File selection ───
    browseButton.addEventListener('click', e => { e.stopPropagation(); fileInput.click(); });
    dropzone.addEventListener('click', e => { if (!e.target.closest('.link-button') && !e.target.closest('.paste-area') && !e.target.closest('.paste-area__toggle') && !e.target.closest('.paste-area__editor')) fileInput.click(); });
    fileInput.addEventListener('change', e => {
        const files = Array.from(e.target.files);
        if (files.length === 0) return;
        if (files.length === 1) handleFile(files[0]);
        else handleBatch(files);
    });
    ['dragenter','dragover'].forEach(ev => dropzone.addEventListener(ev, e => { e.preventDefault(); e.stopPropagation(); dropzone.classList.add('is-dragover'); }));
    ['dragleave','drop'].forEach(ev => dropzone.addEventListener(ev, e => { e.preventDefault(); e.stopPropagation(); dropzone.classList.remove('is-dragover'); }));
    dropzone.addEventListener('drop', e => {
        const files = Array.from(e.dataTransfer.files);
        if (files.length === 0) return;
        if (files.length === 1) handleFile(files[0]);
        else handleBatch(files);
    });

    resetButton.addEventListener('click', () => {
        setReportVisible(false); sourceViewer.hidden = true;
        uploadIdle.hidden = false; uploadLoading.hidden = true;
        fileInput.value = '';
        currentFile = null; currentReport = null;
        closeTooltip();
        window.scrollTo({ top: 0, behavior: 'smooth' });
    });

    // ─── Autofix ───
    autofixButton.addEventListener('click', async () => {
        if (!currentFile) return;
        autofixButton.disabled = true;
        const orig = autofixButton.innerHTML;
        autofixButton.textContent = 'Исправляем…';
        const beforeIssues = currentReport ? currentReport.total_issues : 0;
        try {
            const fd = new FormData(); fd.append('file', currentFile);
            const ext = '.' + currentFile.name.split('.').pop().toLowerCase();
            if (ext === '.docx') {
                fd.append('docx_rules', JSON.stringify(getRulesForAPI()));
            }
            const r = await fetch('/api/autofix', { method: 'POST', body: fd });
            if (!r.ok) { const d = await r.json().catch(() => ({})); throw new Error(d.detail || `Ошибка ${r.status}`); }
            const blob = await r.blob();
            const disp = r.headers.get('Content-Disposition') || '';
            let fn = currentFile.name.replace(/\.([^.]+)$/, '_fixed.$1');
            // Сначала filename*= (UTF-8, кириллица), потом ASCII-запасной вариант
            const mu = disp.match(/filename\*=UTF-8''([^;]+)/i);
            if (mu) { try { fn = decodeURIComponent(mu[1]); } catch (_) {} }
            else { const m = disp.match(/filename="([^"]+)"/); if (m) fn = m[1]; }
            downloadBlob(blob, fn);

            // Re-check the fixed file
            autofixButton.textContent = 'Проверяем результат…';
            const fixedFile = new File([blob], fn, { type: blob.type });
            const fd2 = new FormData(); fd2.append('file', fixedFile);
            if (ext === '.docx') fd2.append('docx_rules', JSON.stringify(getRulesForAPI()));
            const r2 = await fetch('/api/check', { method: 'POST', body: fd2 });
            if (r2.ok) {
                const report = await r2.json();
                currentReport = report;
                currentFile = fixedFile;
                renderReport(report);
                // Show comparison banner
                const fixed = beforeIssues - report.total_issues;
                if (fixed > 0) {
                    const banner = document.createElement('div');
                    banner.className = 'success-banner';
                    banner.innerHTML = `✓ Исправлено: было <strong>${beforeIssues}</strong> → стало <strong>${report.total_issues}</strong> (−${fixed})`;
                    reportBody.insertBefore(banner, reportBody.firstChild);
                }
            }
        } catch (err) { alert('Не удалось: ' + err.message); }
        finally { autofixButton.disabled = false; autofixButton.innerHTML = orig; }
    });

    // ─── PDF ───
    pdfButton.addEventListener('click', async () => {
        if (!currentReport) return;
        pdfButton.disabled = true;
        const orig = pdfButton.innerHTML;
        pdfButton.textContent = 'Генерируем…';
        try {
            const r = await fetch('/api/export-pdf', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(currentReport) });
            if (!r.ok) { const d = await r.json().catch(() => ({})); throw new Error(d.detail || `Ошибка сервера (${r.status})`); }
            downloadBlob(await r.blob(), currentReport.filename.replace(/\.[^.]+$/, '') + '_report.pdf');
        } catch (err) {
            if (err instanceof TypeError) {
                alert('Сервер недоступен. Проверьте соединение.');
            } else {
                alert('Не удалось сгенерировать PDF: ' + err.message);
            }
        }
        finally { pdfButton.disabled = false; pdfButton.innerHTML = orig; }
    });

    // ─── Main handler ───
    async function handleFile(file) {
        const ext = '.' + file.name.split('.').pop().toLowerCase();
        if (!ALLOWED.includes(ext)) { showError(`Формат ${ext} не поддерживается. Допустимы: .py, .docx`); return; }
        showLoading(file.name);
        currentFile = file;
        const fd = new FormData(); fd.append('file', file);
        // Attach docx rules if checking a .docx file
        if (ext === '.docx') {
            fd.append('docx_rules', JSON.stringify(getRulesForAPI()));
        }
        try {
            const r = await fetch('/api/check', { method: 'POST', body: fd });
            if (!r.ok) { const d = await r.json().catch(() => ({})); throw new Error(d.detail || `Ошибка сервера (${r.status})`); }
            const report = await r.json();
            currentReport = report;
            renderReport(report);
            saveToHistory(report);
        } catch (err) {
            if (err instanceof TypeError && err.message.includes('fetch')) {
                showError('Сервер недоступен. Проверьте, что uvicorn запущен, и обновите страницу.');
            } else if (err.name === 'TypeError') {
                showError('Не удалось подключиться к серверу. Проверьте соединение.');
            } else {
                showError(err.message || 'Неизвестная ошибка');
            }
        }
        finally { hideLoading(); }
    }

    // ─── Batch check ───
    async function handleBatch(files) {
        showLoading(`${files.length} файлов…`);
        currentFile = null;
        const fd = new FormData();
        files.forEach(f => fd.append('files', f));
        const hasDocx = files.some(f => f.name.toLowerCase().endsWith('.docx'));
        if (hasDocx) fd.append('docx_rules', JSON.stringify(getRulesForAPI()));

        try {
            const r = await fetch('/api/check-batch', { method: 'POST', body: fd });
            if (!r.ok) { const d = await r.json().catch(() => ({})); throw new Error(d.detail || `Ошибка сервера (${r.status})`); }
            const data = await r.json();
            renderBatchReport(data);
        } catch (err) {
            showError(err.message || 'Ошибка пакетной проверки');
        }
        finally { hideLoading(); }
    }

    function renderBatchReport(data) {
        setReportVisible(true);
        sourceViewer.hidden = true;
        reportType.textContent = '📦';
        reportFilename.textContent = `Пакетная проверка (${data.file_count} файлов)`;
        reportSubtitle.textContent = '';

        animV(statTotal, data.total_issues); animV(statHigh, data.summary.high);
        animV(statMedium, data.summary.medium); animV(statLow, data.summary.low);
        if (data.total_issues === 0) setVerdict('good');
        else if (data.summary.high === 0) setVerdict('ok');
        else setVerdict('bad');
        const tot = data.total_issues || 1;
        sevBarHigh.style.width = (data.summary.high/tot*100)+'%';
        sevBarMedium.style.width = (data.summary.medium/tot*100)+'%';
        sevBarLow.style.width = (data.summary.low/tot*100)+'%';

        autofixButton.disabled = true;
        pdfButton.disabled = true;
        reportBody.innerHTML = '';

        if (data.total_issues === 0) {
            reportBody.innerHTML = '<div class="empty-state"><div class="empty-state__icon">✓</div><div class="empty-state__title">Замечаний не найдено</div><div class="empty-state__text">Все файлы соответствуют требованиям.</div></div>';
            return;
        }

        // Summary table
        let tableHTML = '<div class="batch-summary-table"><table class="batch-table"><thead><tr>'
            + '<th>Файл</th><th>Тип</th><th>Статус</th><th>Ошибок</th><th>Крит.</th><th>Средн.</th><th>Незнач.</th>'
            + '</tr></thead><tbody>';
        data.reports.forEach(r => {
            const badge = r.file_type === 'python' ? 'PY' : r.file_type === 'docx' ? 'DOCX' : '?';
            const vKey = r.total_issues === 0 ? 'good' : (r.verdict || (r.summary.high > 0 ? 'bad' : 'ok'));
            const vLabel = VERDICTS[vKey] ? VERDICTS[vKey].title : '—';
            const rowClass = r.total_issues === 0 ? 'batch-table__row--ok' : r.summary.high > 0 ? 'batch-table__row--high' : '';
            tableHTML += `<tr class="${rowClass}"><td>${esc(r.filename)}</td><td>${badge}</td>`
                + `<td><strong>${vLabel}</strong></td><td>${r.total_issues}</td>`
                + `<td>${r.summary.high}</td><td>${r.summary.medium}</td><td>${r.summary.low}</td></tr>`;
        });
        tableHTML += '</tbody></table></div>';
        reportBody.innerHTML = tableHTML;

        // Per-file accordion
        data.reports.forEach((report, idx) => {
            const badge = report.file_type === 'python' ? 'PY' : report.file_type === 'docx' ? 'DOCX' : '?';
            const statusIcon = report.error ? '⚠' : report.total_issues === 0 ? '✓' : report.total_issues;
            const statusClass = report.error ? 'batch-file--error' : report.total_issues === 0 ? 'batch-file--ok' : 'batch-file--issues';

            const section = document.createElement('div');
            section.className = `batch-file ${statusClass}`;

            let headerHTML = `<div class="batch-file__header" data-batch-idx="${idx}">
                <span class="batch-file__badge">${badge}</span>
                <span class="batch-file__name">${esc(report.filename)}</span>
                <span class="batch-file__count">${statusIcon}</span>
                <span class="batch-file__arrow">▸</span>
            </div>`;

            let bodyHTML = `<div class="batch-file__body" id="batch-body-${idx}" hidden>`;
            if (report.error) {
                bodyHTML += `<div class="error-banner">${esc(report.error)}</div>`;
            } else if (report.issues.length === 0) {
                bodyHTML += '<div style="padding:12px;color:var(--c-green)">✓ Замечаний нет</div>';
            } else {
                report.issues.forEach(iss => {
                    const el = renderIssue(iss, report);
                    bodyHTML += el.outerHTML;
                });
            }
            bodyHTML += '</div>';

            section.innerHTML = headerHTML + bodyHTML;
            section.querySelector('.batch-file__header').addEventListener('click', () => {
                const body = section.querySelector('.batch-file__body');
                const arrow = section.querySelector('.batch-file__arrow');
                body.hidden = !body.hidden;
                arrow.textContent = body.hidden ? '▸' : '▾';
            });

            reportBody.appendChild(section);
        });

        setTimeout(() => reportSection.scrollIntoView({ behavior:'smooth', block:'start' }), 100);
    }

    function showLoading(fn) { uploadIdle.hidden = true; uploadLoading.hidden = false; loadingFilename.textContent = fn; setReportVisible(false); }
    function hideLoading() { uploadIdle.hidden = false; uploadLoading.hidden = true; }

    // ─── Render report ───
    function renderReport(report) {
        setReportVisible(true);
        reportType.textContent = report.file_type === 'python' ? 'PY' : 'DOCX';
        reportFilename.textContent = report.filename;
        const parts = [];
        if (report.file_type === 'python') {
            parts.push('PEP 8 · flake8');
            if (report.source_lines && report.source_lines.length) parts.push(`${report.source_lines.length} строк`);
        } else {
            parts.push('Нормоконтроль · ГОСТ/АГУ');
            if (typeof report.paragraphs_checked === 'number') parts.push(`${report.paragraphs_checked} абзацев`);
        }
        reportSubtitle.textContent = parts.join(' · ');

        animV(statTotal, report.total_issues); animV(statHigh, report.summary.high);
        animV(statMedium, report.summary.medium); animV(statLow, report.summary.low);
        if (report.verdict) {
            setVerdict(report.verdict);
        } else {
            // Python files: derive verdict from summary
            if (report.total_issues === 0) setVerdict('good');
            else if (report.summary.high === 0) setVerdict('ok');
            else setVerdict('bad');
        }
        const tot = report.total_issues || 1;
        sevBarHigh.style.width = (report.summary.high/tot*100)+'%';
        sevBarMedium.style.width = (report.summary.medium/tot*100)+'%';
        sevBarLow.style.width = (report.summary.low/tot*100)+'%';

        reportBody.innerHTML = '';
        autofixButton.disabled = !!report.error || !currentFile;
        pdfButton.disabled = !!report.error;
        if (report.error) { reportBody.innerHTML = `<div class="error-banner">${esc(report.error)}</div>`; sourceViewer.hidden = true; return; }
        if (!report.issues.length) {
            reportBody.innerHTML = '<div class="empty-state"><div class="empty-state__icon">✓</div><div class="empty-state__title">Замечаний не найдено</div><div class="empty-state__text">Файл соответствует всем проверяемым требованиям.</div></div>';
            autofixButton.disabled = true;
            // Still show source code without errors if available
            if (report.file_type === 'python' && report.source_lines && report.source_lines.length) {
                renderSourceViewer(report);
            } else {
                sourceViewer.hidden = true;
            }
            return;
        }

        // Build error lookup by line number
        const errorsByLine = {};
        report.issues.forEach(iss => {
            const ln = iss.line;
            if (ln) {
                if (!errorsByLine[ln]) errorsByLine[ln] = [];
                errorsByLine[ln].push(iss);
            }
        });

        // Render source viewer for Python files
        if (report.file_type === 'python' && report.source_lines && report.source_lines.length) {
            renderSourceViewer(report, errorsByLine);
        } else {
            sourceViewer.hidden = true;
        }

        // Group issues by code
        const groups = {};
        const GROUP_NAMES = {
            FONT_MISMATCH: 'Шрифт', FONT_NOT_SET: 'Шрифт', FONT_SIZE_MISMATCH: 'Размер шрифта', FONT_SIZE_NOT_SET: 'Размер шрифта',
            LINE_SPACING_MISMATCH: 'Межстрочный интервал', LINE_SPACING_NOT_SET: 'Межстрочный интервал',
            INDENT_MISMATCH: 'Красная строка', INDENT_NOT_SET: 'Красная строка',
            PARA_SPACING_BEFORE: 'Интервалы между абзацами', PARA_SPACING_AFTER: 'Интервалы между абзацами',
            MARGIN_MISMATCH: 'Поля страницы', MARGIN_MISSING: 'Поля страницы', ALIGN_NOT_JUSTIFY: 'Выравнивание',
            HEADING_ENDS_WITH_DOT: 'Заголовки', HEADING_NOT_BOLD: 'Заголовки', HEADING_NUM_GAP: 'Нумерация заголовков', HEADING_LEVEL_SKIP: 'Нумерация заголовков',
            FIGURE_CAPTION_FORMAT: 'Подписи к рисункам', FIGURE_DASH_WRONG: 'Подписи к рисункам', FIGURE_NO_REFERENCE: 'Ссылки на рисунки',
            TABLE_CAPTION_FORMAT: 'Подписи к таблицам', TABLE_DASH_WRONG: 'Подписи к таблицам', TABLE_NO_CAPTION: 'Подписи к таблицам', TABLE_NO_REFERENCE: 'Ссылки на таблицы',
            TABLE_FONT_MISMATCH: 'Таблицы',
            BIBLIOGRAPHY_MISSING: 'Список литературы', BIBLIOGRAPHY_EMPTY: 'Список литературы', BIBLIOGRAPHY_ENTRY_FORMAT: 'Список литературы', BIB_NO_YEAR: 'Список литературы',
            APPENDIX_FORMAT: 'Приложения', APPENDIX_ORDER: 'Приложения',
            NO_PAGE_NUMBERS: 'Нумерация страниц', NO_TABLE_OF_CONTENTS: 'Оглавление',
            HYPERLINKS_FOUND: 'Гиперссылки', COLORED_TEXT: 'Цвет текста',
            DOUBLE_SPACES: 'Пробелы', SPACE_BEFORE_PUNCT: 'Пробелы', EXTRA_BLANK_LINES: 'Пустые строки',
        };
        report.issues.forEach(iss => {
            const group = GROUP_NAMES[iss.code] || iss.code;
            if (!groups[group]) groups[group] = [];
            groups[group].push(iss);
        });

        const issueObserver = new IntersectionObserver((entries) => {
            entries.forEach(entry => {
                if (entry.isIntersecting) {
                    entry.target.classList.add('is-visible');
                    issueObserver.unobserve(entry.target);
                }
            });
        }, { threshold: 0.05, rootMargin: '0px 0px -20px 0px' });

        // For Python with few issues or single group — flat list
        const groupKeys = Object.keys(groups);
        if (report.file_type === 'python' || (groupKeys.length <= 1 && report.issues.length <= 10)) {
            report.issues.forEach((iss, i) => {
                const el = renderIssue(iss, report);
                el.style.transitionDelay = Math.min(i * 40, 400) + 'ms';
                if (iss.line && report.file_type === 'python') {
                    el.style.cursor = 'pointer';
                    el.addEventListener('click', () => scrollToSourceLine(iss.line));
                }
                reportBody.appendChild(el);
                issueObserver.observe(el);
            });
        } else {
            // Grouped view
            let globalIdx = 0;
            groupKeys.forEach(groupName => {
                const issues = groups[groupName];
                const severity = issues.some(i=>i.severity==='high') ? 'high' : issues.some(i=>i.severity==='medium') ? 'medium' : 'low';

                const section = document.createElement('div');
                section.className = 'issue-group';

                const header = document.createElement('div');
                header.className = `issue-group__header issue-group__header--${severity}`;
                header.innerHTML = `<span class="issue-group__name">${esc(groupName)}</span>`
                    + `<span class="issue-group__count">${issues.length}</span>`
                    + `<span class="issue-group__arrow">▾</span>`;

                const body = document.createElement('div');
                body.className = 'issue-group__body';

                issues.forEach(iss => {
                    const el = renderIssue(iss, report);
                    el.style.transitionDelay = Math.min(globalIdx * 20, 200) + 'ms';
                    if (iss.line && report.file_type === 'python') {
                        el.style.cursor = 'pointer';
                        el.addEventListener('click', () => scrollToSourceLine(iss.line));
                    }
                    body.appendChild(el);
                    issueObserver.observe(el);
                    globalIdx++;
                });

                header.addEventListener('click', () => {
                    body.hidden = !body.hidden;
                    header.querySelector('.issue-group__arrow').textContent = body.hidden ? '▸' : '▾';
                });

                section.appendChild(header);
                section.appendChild(body);
                reportBody.appendChild(section);
            });
        }

        setTimeout(() => reportSection.scrollIntoView({ behavior:'smooth', block:'start' }), 100);
    }

    // ─── Source Viewer ───
    let activeTooltip = null;

    function renderSourceViewer(report, errorsByLine) {
        errorsByLine = errorsByLine || {};
        sourceViewer.hidden = false;
        sourceCode.innerHTML = '';
        sourceMinimap.innerHTML = '';

        const totalLines = report.source_lines.length;

        // Build minimap marks
        report.issues.forEach(iss => {
            if (!iss.line) return;
            const mark = document.createElement('div');
            mark.className = `minimap__mark minimap__mark--${iss.severity}`;
            mark.style.left = ((iss.line - 1) / totalLines * 100) + '%';
            mark.title = `стр. ${iss.line}: ${iss.code}`;
            mark.addEventListener('click', () => scrollToSourceLine(iss.line));
            sourceMinimap.appendChild(mark);
        });

        // Build lines
        report.source_lines.forEach((text, idx) => {
            const lineNum = idx + 1;
            const lineEl = document.createElement('div');
            lineEl.className = 'src-line';
            lineEl.setAttribute('data-line', lineNum);

            const errors = errorsByLine[lineNum];
            if (errors) {
                // Use highest severity on the line
                const sev = errors.some(e => e.severity === 'high') ? 'high'
                    : errors.some(e => e.severity === 'medium') ? 'medium' : 'low';
                lineEl.classList.add('src-line--' + sev);
            }

            const noEl = document.createElement('span');
            noEl.className = 'src-line__no';
            noEl.textContent = lineNum;

            const textEl = document.createElement('span');
            textEl.className = 'src-line__text';
            textEl.textContent = text;

            lineEl.appendChild(noEl);
            lineEl.appendChild(textEl);

            // Error badges on the right
            if (errors) {
                const badgesWrap = document.createElement('span');
                badgesWrap.className = 'src-line__errors';
                errors.forEach(err => {
                    const badge = document.createElement('span');
                    badge.className = `src-err-badge src-err-badge--${err.severity}`;
                    badge.textContent = err.code;
                    badge.title = err.description || err.message;
                    badge.addEventListener('click', (e) => {
                        e.stopPropagation();
                        showSourceTooltip(badge, err);
                    });
                    badgesWrap.appendChild(badge);
                });
                lineEl.appendChild(badgesWrap);
            }

            // Click line to show errors
            if (errors) {
                lineEl.style.cursor = 'pointer';
                lineEl.addEventListener('click', () => {
                    if (errors.length === 1) {
                        const badge = lineEl.querySelector('.src-err-badge');
                        showSourceTooltip(badge, errors[0]);
                    } else {
                        showMultiTooltip(lineEl, errors);
                    }
                });
            }

            sourceCode.appendChild(lineEl);
        });
    }

    function scrollToSourceLine(lineNum) {
        const lineEl = sourceCode.querySelector(`[data-line="${lineNum}"]`);
        if (!lineEl) return;

        // Scroll the source viewer into viewport first
        sourceViewer.scrollIntoView({ behavior: 'smooth', block: 'start' });

        // Then scroll within the code container
        setTimeout(() => {
            const containerRect = sourceCode.getBoundingClientRect();
            const lineRect = lineEl.getBoundingClientRect();
            const offset = lineRect.top - containerRect.top + sourceCode.scrollTop - sourceCode.clientHeight / 3;
            sourceCode.scrollTo({ top: offset, behavior: 'smooth' });

            // Flash animation
            lineEl.classList.remove('src-line--active');
            void lineEl.offsetWidth; // force reflow
            lineEl.classList.add('src-line--active');
        }, 400);
    }

    function showSourceTooltip(anchor, err) {
        closeTooltip();
        const tip = document.createElement('div');
        tip.className = 'src-tooltip';
        tip.innerHTML = `<button class="src-tooltip__close" aria-label="Закрыть">&times;</button>`
            + `<span class="src-tooltip__code src-tooltip__code--${err.severity}">${esc(err.code)}</span>`
            + `<span class="src-tooltip__desc">${esc(err.description || err.message)}</span>`;
        tip.querySelector('.src-tooltip__close').addEventListener('click', closeTooltip);

        // Position relative to source code container
        sourceCode.style.position = 'relative';
        const anchorRect = anchor.getBoundingClientRect();
        const codeRect = sourceCode.getBoundingClientRect();
        tip.style.top = (anchorRect.bottom - codeRect.top + sourceCode.scrollTop + 4) + 'px';
        tip.style.right = '12px';

        sourceCode.appendChild(tip);
        activeTooltip = tip;

        // Close on outside click
        setTimeout(() => document.addEventListener('click', onTooltipOutside), 0);
    }

    function showMultiTooltip(lineEl, errors) {
        closeTooltip();
        const tip = document.createElement('div');
        tip.className = 'src-tooltip';
        let html = `<button class="src-tooltip__close" aria-label="Закрыть">&times;</button>`;
        errors.forEach((err, i) => {
            if (i > 0) html += '<div style="border-top:1px solid var(--c-line-soft);margin:6px 0"></div>';
            html += `<div><span class="src-tooltip__code src-tooltip__code--${err.severity}">${esc(err.code)}</span>`
                + `<span class="src-tooltip__desc">${esc(err.description || err.message)}</span></div>`;
        });
        tip.innerHTML = html;
        tip.querySelector('.src-tooltip__close').addEventListener('click', closeTooltip);

        sourceCode.style.position = 'relative';
        const lineRect = lineEl.getBoundingClientRect();
        const codeRect = sourceCode.getBoundingClientRect();
        tip.style.top = (lineRect.bottom - codeRect.top + sourceCode.scrollTop + 4) + 'px';
        tip.style.right = '12px';

        sourceCode.appendChild(tip);
        activeTooltip = tip;
        setTimeout(() => document.addEventListener('click', onTooltipOutside), 0);
    }

    function closeTooltip() {
        if (activeTooltip) { activeTooltip.remove(); activeTooltip = null; }
        document.removeEventListener('click', onTooltipOutside);
    }

    function onTooltipOutside(e) {
        if (activeTooltip && !activeTooltip.contains(e.target)) closeTooltip();
    }

    function animV(el, target) {
        const dur=600, start=parseInt(el.textContent)||0;
        if (start===target) { el.textContent=target; return; }
        const t0=performance.now();
        (function tick(now){
            const p=Math.min((now-t0)/dur,1);
            // ease-out cubic
            el.textContent=Math.round(start+(target-start)*(1-Math.pow(1-p,3)));
            if (p<1) requestAnimationFrame(tick);
        })(t0);
    }

    function renderIssue(iss, report) {
        const el = document.createElement('article');
        el.className = `issue issue--${iss.severity}`;
        const loc = document.createElement('div'); loc.className = 'issue__location';
        loc.textContent = report.file_type === 'python'
            ? `стр. ${iss.line}${iss.column?':'+iss.column:''}`
            : (iss.location||'');
        const body = document.createElement('div'); body.className = 'issue__body';
        // Description first (human-readable), code secondary
        body.innerHTML = `<div class="issue__description">${esc(iss.description||iss.message)}</div>`;
        body.innerHTML += `<div class="issue__code-row"><span class="issue__code">${esc(iss.code)}</span></div>`;
        if (iss.expected || iss.actual) {
            let h = '<div class="issue__expected">';
            if (iss.expected) h += `<span>✓ ${esc(iss.expected)}</span>`;
            if (iss.actual) h += `<span>✗ ${esc(iss.actual)}</span>`;
            body.innerHTML += h + '</div>';
        }
        if (report.file_type === 'python' && report.source_lines && report.source_lines.length && iss.line) {
            const s = Math.max(1,iss.line-1), e = Math.min(report.source_lines.length,iss.line+1);
            const pre = document.createElement('div'); pre.className = 'issue__source';
            for (let i=s; i<=e; i++) {
                const ln = document.createElement('span');
                ln.className = 'issue__source-line'+(i===iss.line?' is-target':'');
                ln.innerHTML = `<span class="lineno">${i}</span>`;
                ln.appendChild(document.createTextNode(report.source_lines[i-1]||''));
                pre.appendChild(ln);
            }
            body.appendChild(pre);
        }
        el.appendChild(loc); el.appendChild(body);
        return el;
    }

    function showError(msg) {
        setReportVisible(true);
        reportType.textContent = '!'; reportFilename.textContent = 'Ошибка'; reportSubtitle.textContent = '';
        ['stat-total','stat-high','stat-medium','stat-low'].forEach(id => $(id).textContent = '—');
        resetVerdict();
        sevBarHigh.style.width='0%'; sevBarMedium.style.width='0%'; sevBarLow.style.width='0%';
        reportBody.innerHTML = `<div class="error-banner">${esc(msg)}</div>`;
        autofixButton.disabled = true; pdfButton.disabled = true;
    }

    // ─── History ───
    function getHistory() { try { return JSON.parse(localStorage.getItem(HISTORY_KEY)||'[]'); } catch { return []; } }

    function saveToHistory(r) {
        const h = getHistory();
        // Сохраняем полный отчёт + timestamp
        const entry = Object.assign({}, r, { timestamp: Date.now() });
        h.unshift(entry);
        // Ограничиваем 30 записями; для экономии localStorage — обрезаем source_lines у старых
        const trimmed = h.slice(0, 30).map((e, idx) => {
            if (idx > 10 && e.source_lines && e.source_lines.length > 50) {
                e.source_lines = e.source_lines.slice(0, 50);
            }
            return e;
        });
        try {
            localStorage.setItem(HISTORY_KEY, JSON.stringify(trimmed));
        } catch (e) {
            // localStorage переполнен — удаляем старые записи
            trimmed.pop();
            try { localStorage.setItem(HISTORY_KEY, JSON.stringify(trimmed)); } catch {}
        }
        renderHistory();
    }

    function renderHistory() {
        const h = getHistory();
        if (!h.length) { historyList.innerHTML = ''; historyEmpty.hidden = false; historyList.appendChild(historyEmpty); return; }
        historyEmpty.hidden = true; historyList.innerHTML = '';
        h.forEach((e, idx) => {
            const el = document.createElement('div'); el.className = 'history-item';
            el.style.cursor = 'pointer';
            const badge = e.file_type === 'python' ? 'PY' : 'DOCX';
            const d = new Date(e.timestamp);
            const ds = d.toLocaleDateString('ru-RU') + ' ' + d.toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' });
            const s = e.summary || {};
            const issuesLabel = e.total_issues === 0 ? '<span style="color:var(--c-green)">✓ без замечаний</span>' : `${e.total_issues} замечаний`;
            el.innerHTML = `<span class="history-item__badge">${badge}</span>`
                + `<div class="history-item__info"><div class="history-item__name">${esc(e.filename)}</div>`
                + `<div class="history-item__meta">${ds} · ${issuesLabel}</div></div>`
                + `<div class="history-item__stats">`
                + `<span class="history-item__stat history-item__stat--high"><span class="history-item__dot history-item__dot--high"></span>${s.high || 0}</span>`
                + `<span class="history-item__stat history-item__stat--medium"><span class="history-item__dot history-item__dot--medium"></span>${s.medium || 0}</span>`
                + `<span class="history-item__stat history-item__stat--low"><span class="history-item__dot history-item__dot--low"></span>${s.low || 0}</span></div>`
                + `<span class="history-item__open" title="Открыть отчёт">→</span>`;

            el.addEventListener('click', () => openHistoryEntry(idx));
            historyList.appendChild(el);
        });
    }

    function openHistoryEntry(idx) {
        const h = getHistory();
        if (idx >= h.length) return;
        const report = h[idx];
        if (!report.issues) { alert('Данные этой проверки неполные.'); return; }
        currentReport = report;
        currentFile = null; // нет файла — нельзя повторно исправлять
        renderReport(report);
        // Прокрутка к отчёту
        setTimeout(() => reportSection.scrollIntoView({ behavior: 'smooth', block: 'start' }), 100);
    }

    clearHistoryBtn.addEventListener('click', () => { if (confirm('Очистить историю проверок?')) { localStorage.removeItem(HISTORY_KEY); renderHistory(); } });
    renderHistory();

    function downloadBlob(blob, fn) {
        const u=URL.createObjectURL(blob), a=document.createElement('a');
        a.href=u; a.download=fn; document.body.appendChild(a); a.click();
        document.body.removeChild(a); URL.revokeObjectURL(u);
    }
    function esc(t) { const d=document.createElement('div'); d.textContent=String(t); return d.innerHTML; }

    // ─── Onboarding (first visit) ───
    const ONBOARD_KEY = 'smartdoc-onboarded';
    if (!localStorage.getItem(ONBOARD_KEY)) {
        const onboardModal = $('onboarding-modal');
        if (onboardModal) {
            onboardModal.hidden = false;
            const dismiss = () => { onboardModal.hidden = true; localStorage.setItem(ONBOARD_KEY, '1'); };
            $('onboarding-close').addEventListener('click', dismiss);
            $('onboarding-start').addEventListener('click', dismiss);
            onboardModal.addEventListener('click', e => { if (e.target === onboardModal) dismiss(); });
        }
    }
})();
