(function () {
    'use strict';

    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const fileInfo = document.getElementById('fileInfo');
    const fileNameEl = document.getElementById('fileName');
    const fileSizeEl = document.getElementById('fileSize');
    const clearBtn = document.getElementById('clearBtn');
    const convertRow = document.getElementById('convertRow');
    const convertBtn = document.getElementById('convertBtn');
    const statusEl = document.getElementById('status');
    const segBtns = Array.from(document.querySelectorAll('.seg-btn'));

    const MAX_BYTES = 50 * 1024 * 1024;
    let currentFile = null;
    let targetFormat = 'xlsx';

    function setStatus(message, kind) {
        if (!message) {
            statusEl.hidden = true;
            statusEl.textContent = '';
            statusEl.className = 'status';
            return;
        }
        statusEl.hidden = false;
        statusEl.textContent = message;
        statusEl.className = 'status is-' + (kind || 'info');
    }

    function formatBytes(bytes) {
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
        return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
    }

    function extensionOf(name) {
        const idx = name.lastIndexOf('.');
        return idx === -1 ? '' : name.slice(idx + 1).toLowerCase();
    }

    function isCsvFile(file) {
        const ext = extensionOf(file.name);
        if (ext === 'csv' || ext === 'tsv' || ext === 'txt') return true;
        if (ext === 'xlsx' || ext === 'xls') return false;
        return (file.type || '').toLowerCase().includes('csv');
    }

    function setTarget(target) {
        targetFormat = target;
        segBtns.forEach(function (btn) {
            const active = btn.dataset.target === target;
            btn.setAttribute('aria-checked', active ? 'true' : 'false');
        });
    }

    function pickFile(file) {
        if (!file) return;
        if (file.size > MAX_BYTES) {
            setStatus('File is larger than 50 MB. Try splitting it first.', 'error');
            return;
        }
        const ext = extensionOf(file.name);
        const allowed = ['csv', 'tsv', 'txt', 'xlsx', 'xls'];
        if (!allowed.includes(ext)) {
            setStatus('Unsupported file type. Please upload a .csv, .xlsx, or .xls file.', 'error');
            return;
        }

        currentFile = file;
        fileNameEl.textContent = file.name;
        fileSizeEl.textContent = formatBytes(file.size);
        fileInfo.hidden = false;
        convertRow.hidden = false;
        setStatus('', null);

        setTarget(isCsvFile(file) ? 'xlsx' : 'csv');
    }

    function resetFile() {
        currentFile = null;
        fileInput.value = '';
        fileInfo.hidden = true;
        convertRow.hidden = true;
        setStatus('', null);
    }

    function triggerDownload(blob, filename) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
    }

    function baseName(name) {
        const idx = name.lastIndexOf('.');
        return idx === -1 ? name : name.slice(0, idx);
    }

    async function convert() {
        if (!currentFile) return;
        if (typeof XLSX === 'undefined') {
            setStatus('Converter is still loading — please try again in a moment.', 'error');
            return;
        }

        convertBtn.disabled = true;
        setStatus('Converting…', 'info');

        try {
            const buffer = await currentFile.arrayBuffer();
            const inputIsCsv = isCsvFile(currentFile);
            const readOpts = inputIsCsv
                ? { type: 'array', raw: false, codepage: 65001 }
                : { type: 'array' };
            const workbook = XLSX.read(buffer, readOpts);

            if (!workbook.SheetNames.length) {
                throw new Error('No sheets found in the file.');
            }

            const outName = baseName(currentFile.name);

            if (targetFormat === 'xlsx') {
                const out = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
                triggerDownload(
                    new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
                    outName + '.xlsx'
                );
            } else {
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const csv = XLSX.utils.sheet_to_csv(firstSheet);
                const bom = '\uFEFF';
                triggerDownload(
                    new Blob([bom + csv], { type: 'text/csv;charset=utf-8' }),
                    outName + '.csv'
                );
            }

            const sheetNote = targetFormat === 'csv' && workbook.SheetNames.length > 1
                ? ' Only the first sheet (“' + workbook.SheetNames[0] + '”) was exported.'
                : '';
            setStatus('Done. Your file has been downloaded.' + sheetNote, 'success');
        } catch (err) {
            console.error(err);
            setStatus('Something went wrong while converting the file. Please check the file and try again.', 'error');
        } finally {
            convertBtn.disabled = false;
        }
    }

    fileInput.addEventListener('change', function (e) {
        const file = e.target.files && e.target.files[0];
        if (file) pickFile(file);
    });

    dropZone.addEventListener('keydown', function (e) {
        if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            fileInput.click();
        }
    });

    ['dragenter', 'dragover'].forEach(function (evt) {
        dropZone.addEventListener(evt, function (e) {
            e.preventDefault();
            e.stopPropagation();
            dropZone.classList.add('is-dragover');
        });
    });

    ['dragleave', 'drop'].forEach(function (evt) {
        dropZone.addEventListener(evt, function (e) {
            e.preventDefault();
            e.stopPropagation();
            if (evt === 'dragleave' && dropZone.contains(e.relatedTarget)) return;
            dropZone.classList.remove('is-dragover');
        });
    });

    dropZone.addEventListener('drop', function (e) {
        const file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
        if (file) pickFile(file);
    });

    clearBtn.addEventListener('click', resetFile);
    convertBtn.addEventListener('click', convert);

    segBtns.forEach(function (btn) {
        btn.addEventListener('click', function () {
            setTarget(btn.dataset.target);
        });
    });

    setTarget('xlsx');
})();
