/* global XLSX, LuckyExcel */
// Keep the original XLSX formatting when possible; use SheetJS for legacy XLS
// and workbooks that the older LuckyExcel converter cannot import.
function excelNumberFormat(format) {
    // Older SSF versions interpret unquoted dots in dates as fractional seconds.
    // Quote date separators, leaving quoted literals and escaped characters intact.
    return format.replace(/"[^"]*"|\\.|([dmy]+)\.(?=[dmy])/gi,
        (match, token) => token ? token + '"."' : match);
}

function spreadsheetFallback(workbook) {
    return workbook.SheetNames.map((name, index) => {
        const sheet = workbook.Sheets[name];
        const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
        const celldata = [];
        const cells = new Map();
        for (const address of Object.keys(sheet)) {
            if (address.startsWith('!')) continue;
            const cell = sheet[address];
            const position = XLSX.utils.decode_cell(address);
            let formatted = cell.w;
            if (formatted == null && cell.z && cell.t === 'n') {
                try { formatted = XLSX.SSF.format(excelNumberFormat(cell.z), cell.v, { date1904: !!workbook.Workbook?.WBProps?.date1904 }); } catch (_) {}
            }
            formatted ??= XLSX.utils.format_cell(cell);
            const customFormat = cell.z && cell.z !== 'General';
            const value = {
                v: customFormat ? formatted : cell.v ?? '',
                m: formatted,
                // Luckysheet's older formatter rejects some valid Excel formats
                // (including dd.mm.yyyy). Keep their already formatted text.
                ct: { fa: customFormat ? '@' : 'General', t: customFormat ? 's' : cell.t === 'n' ? 'n' : cell.t === 'b' ? 'b' : 's' }
            };
            // Display cached formula results; do not recalculate imported formulas.
            if (cell.f && cell.v == null) value.v = value.m = '=' + cell.f;
            const entry = { r: position.r, c: position.c, v: value };
            celldata.push(entry);
            cells.set(address, entry);
        }
        const config = { merge: {}, rowlen: {}, columnlen: {} };
        for (const merge of sheet['!merges'] || []) {
            const key = merge.s.r + '_' + merge.s.c;
            const mc = { r: merge.s.r, c: merge.s.c, rs: merge.e.r - merge.s.r + 1, cs: merge.e.c - merge.s.c + 1 };
            config.merge[key] = mc;
            const address = XLSX.utils.encode_cell(merge.s);
            let entry = cells.get(address);
            if (!entry) {
                entry = { r: merge.s.r, c: merge.s.c, v: { v: '', m: '', ct: { t: 's', fa: 'General' } } };
                celldata.push(entry);
            }
            entry.v.mc = mc;
        }
        (sheet['!rows'] || []).forEach((row, i) => {
            if (row.hpx || row.hpt) config.rowlen[i] = row.hpx || row.hpt * 4 / 3;
        });
        (sheet['!cols'] || []).forEach((column, i) => {
            if (column.wpx || column.wch) config.columnlen[i] = column.wpx || column.wch * 7 + 5;
        });
        return { name, index: String(index), order: index, status: index === 0 ? 1 : 0,
            row: Math.max(36, range.e.r + 1), column: Math.max(18, range.e.c + 1), celldata, config };
    });
}

function importXlsx(blob, fileName) {
    return new Promise((resolve, reject) => {
        const timer = setTimeout(() => reject(new Error('Конвертер Excel не ответил')), 5000);
        const done = (callback, value) => { clearTimeout(timer); callback(value); };
        try {
            LuckyExcel.transformExcelToLucky(new File([blob], fileName), result => {
                if (!result?.sheets?.length) done(reject, new Error('В книге нет листов'));
                else done(resolve, result.sheets);
            }, error => done(reject, error));
        } catch (error) { done(reject, error); }
    });
}

function displaySpreadsheet(sheets, title) {
    const iframe = document.getElementById('luckysheet-iframe');
    iframe.style.display = 'block';
    const requestId = crypto.randomUUID();
    let active = false;
    sheets.forEach(sheet => {
        for (const entry of sheet.celldata || []) {
            if (typeof entry.v?.ct?.fa === 'string') entry.v.ct.fa = excelNumberFormat(entry.v.ct.fa);
        }
        sheet.status = !active && !sheet.hide && Number(sheet.status) === 1 ? 1 : 0;
        if (sheet.status) active = true;
    });
    if (!active) sheets[0].status = 1;
    return new Promise((resolve, reject) => {
        let sent = false;
        const cleanup = () => {
            clearTimeout(timer);
            window.removeEventListener('message', receive);
            iframe.removeEventListener('load', ping);
        };
        const timer = setTimeout(() => { cleanup(); reject(new Error('Табличный просмотрщик не ответил. Перезагрузите расширение и вкладку MAX.')); }, 20000);
        const ping = () => iframe.contentWindow.postMessage({ action: 'pingExcel' }, window.location.origin);
        const receive = event => {
            if (event.source !== iframe.contentWindow || event.origin !== window.location.origin || !event.data) return;
            if (event.data.action === 'iframeReady' && !sent) {
                sent = true;
                iframe.contentWindow.postMessage({ action: 'loadExcel', sheets, title, requestId }, window.location.origin);
            } else if (event.data.requestId === requestId && ['excelLoaded', 'excelError'].includes(event.data.action)) {
                cleanup();
                if (event.data.action === 'excelError') reject(new Error(event.data.error));
                else resolve();
            }
        };
        window.addEventListener('message', receive);
        iframe.addEventListener('load', ping);
        ping();
    });
}

async function renderExcel(blob, fileName) {
    const data = new Uint8Array(await blob.arrayBuffer());
    const zip = data[0] === 0x50 && data[1] === 0x4b;
    const ole = data[0] === 0xd0 && data[1] === 0xcf && data[2] === 0x11 && data[3] === 0xe0;
    const biff = data[0] === 0x09 && [0, 2, 4, 8].includes(data[1]);
    if (!zip && !ole && !biff) throw new Error('Файл не является книгой Excel или повреждён.');
    const workbook = XLSX.read(data, { type: 'array', cellNF: true, cellStyles: true });
    if (!workbook.SheetNames.length) throw new Error('В книге Excel нет листов.');
    let loaded = false;
    if (zip) {
        try {
            const sheets = await importXlsx(blob, fileName);
            await displaySpreadsheet(sheets, fileName);
            loaded = true;
        } catch (error) { console.warn('Excel import fallback:', error); }
    }
    if (!loaded) await displaySpreadsheet(spreadsheetFallback(workbook), fileName);
    document.getElementById('loading').style.display = 'none';
}
