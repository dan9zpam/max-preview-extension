const { chromium } = require('playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const vm = require('node:vm');
const root = path.resolve(__dirname, '..');
const XLSX = require('../viewer/libs/xlsx.full.min.js');
const zipContext = { setTimeout, clearTimeout, setImmediate, Uint8Array, ArrayBuffer, Promise, TextEncoder, TextDecoder };
vm.createContext(zipContext);
vm.runInContext(fs.readFileSync(path.join(root, 'viewer/libs/jszip.min.js'), 'utf8'), zipContext);

function makePdf(pages = 1) {
    const fontId = 3 + pages * 2;
    const kids = Array.from({ length: pages }, (_, i) => `${3 + i * 2} 0 R`).join(' ');
    const objects = ['<< /Type /Catalog /Pages 2 0 R >>', `<< /Type /Pages /Kids [${kids}] /Count ${pages} >>`];
    for (let i = 0; i < pages; i++) {
        objects.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Resources << /Font << /F1 ${fontId} 0 R >> >> /Contents ${4 + i * 2} 0 R >>`);
        const stream = `BT /F1 18 Tf 20 100 Td (Page ${i + 1}) Tj ET`;
        objects.push(`<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`);
    }
    objects.push('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>');
    let result = '%PDF-1.4\n';
    const offsets = [0];
    objects.forEach((object, i) => { offsets.push(result.length); result += `${i + 1} 0 obj\n${object}\nendobj\n`; });
    const xref = result.length;
    result += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n` + offsets.slice(1).map(n => `${String(n).padStart(10, '0')} 00000 n \n`).join('');
    return Buffer.from(result + `trailer\n<< /Root 1 0 R /Size ${objects.length + 1} >>\nstartxref\n${xref}\n%%EOF`);
}
async function makeDocx() {
    const zip = new zipContext.JSZip();
    zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
    zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
    zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Проверка DOCX</w:t></w:r></w:p><w:sectPr/></w:body></w:document>');
    zip.file('padding.txt', 'x'.repeat(600000));
    return Buffer.from(await zip.generateAsync({ type: 'uint8array' }));
}
function makeExcel(bookType) {
    const book = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet([['Проверка Excel', null, 'Дата'], ['Значение', 42, 46000], ['Итого', 100]]);
    sheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
    sheet.B3 = { t: 'n', v: 84, f: 'B2*2' };
    sheet.C2.z = 'dd.mm.yyyy';
    XLSX.utils.book_append_sheet(book, sheet, 'Данные');
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['Второй лист', 7]]), 'Итоги');
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([]), 'Пустой');
    return Buffer.from(XLSX.write(book, { type: 'buffer', bookType }));
}
(async () => {
    const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'max-preview-test-'));
    const downloadDir = path.join(profile, 'downloads');
    fs.mkdirSync(downloadDir);
    fs.mkdirSync(path.join(profile, 'Default'));
    fs.writeFileSync(path.join(profile, 'Default/Preferences'), JSON.stringify({ download: { default_directory: downloadDir, prompt_for_download: false } }));
    const context = await chromium.launchPersistentContext(profile, {
        headless: true, channel: 'chromium', downloadsPath: path.join(profile, 'downloads'),
        args: [`--disable-extensions-except=${root}`, `--load-extension=${root}`]
    });
    try {
        let worker = context.serviceWorkers()[0];
        if (!worker) worker = await context.waitForEvent('serviceworker');
        const id = new URL(worker.url()).host;
        const pdf = makePdf();
        const docx = await makeDocx();
        const archive = new zipContext.JSZip();
        const xlsx = makeExcel('xlsx');
        const xls = makeExcel('xls');
        archive.file('inner.pdf', new Uint8Array(pdf));
        archive.file('inner.xlsx', new Uint8Array(xlsx));
        const archiveBytes = Buffer.from(await archive.generateAsync({ type: 'uint8array' }));
        await context.route('https://web.max.ru/**', route => route.fulfill({ contentType: 'text/html', body: '<a id="file" href="https://web.max.ru/report.pdf#page=1">Документ\n42 КБ</a>' }));
        const files = new Map([
            ['/opaque-name', { body: pdf, headers: { 'content-disposition': 'attachment; filename=2dfc5491c802494793456b359d897986' } }],
            ['/mojibake-name', { body: docx, headers: { 'content-disposition': 'attachment; filename="' + Buffer.from('Практическая работа.docx').toString('latin1') + '"' } }],
            ['/utf8-name', { body: pdf, headers: { 'content-disposition': "attachment; filename*=UTF-8'ru'%D0%9F%D0%BB%D0%B0%D0%BD" } }],
            ['/xlsx', { body: xlsx }], ['/xls', { body: xls }],
            ['/multipage', { body: makePdf(3) }], ['/archive', { body: archiveBytes }], ['/pdf', { body: pdf }], ['/docx', { body: docx }],
            ['/named', { body: docx, headers: { 'content-disposition': "attachment; filename*=UTF-8''%D0%A2%D0%B5%D1%81%D1%82.docx" } }],
            ['/broken', { body: Buffer.from('broken data') }],
            ['/login', { body: '<html>Sign in</html>', contentType: 'text/html' }],
            ['/empty', { body: '' }], ['/denied', { status: 403, body: 'denied' }]
        ]);
        let cookieHeader;
        await context.addCookies([{ name: 'session', value: 'test', domain: 'files.max.ru', path: '/', secure: true, sameSite: 'None' }]);
        await context.route('https://files.max.ru/**', async route => {
            cookieHeader = route.request().headers().cookie;
            await route.fulfill({ contentType: 'application/octet-stream', ...files.get(new URL(route.request().url()).pathname) });
        });
        const page = await context.newPage();
        // Playwright's allowAndName mode skips extension filename determination.
        // Exercise Chrome's native download path and inspect its final filenames.
        await (await context.newCDPSession(page)).send('Browser.setDownloadBehavior', { behavior: 'default' });
        const errors = [];
        page.on('pageerror', e => errors.push(e.message));
        async function open(endpoint, name) {
            await page.goto(`chrome-extension://${id}/viewer/viewer.html?url=${encodeURIComponent('https://files.max.ru/' + endpoint)}&name=${encodeURIComponent(name)}`);
        }
        async function saveAndVerify(targetPage, button, expectedName, bytes) {
            const before = await targetPage.evaluate(async () => (await chrome.downloads.search({})).map(item => item.id));
            await targetPage.locator(button).click();
            let saved;
            const deadline = Date.now() + 30000;
            while (!saved && Date.now() < deadline) {
                saved = await targetPage.evaluate(async previous => {
                    const items = await chrome.downloads.search({});
                    return items.find(item => !previous.includes(item.id) && item.state === 'complete');
                }, before);
                if (!saved) await targetPage.waitForTimeout(100);
            }
            assert.ok(saved, `Download did not complete: ${expectedName}`);
            assert.equal(path.basename(saved.filename), expectedName);
            assert.equal(path.dirname(saved.filename), downloadDir);
            assert.deepEqual(fs.readFileSync(saved.filename), bytes);
        }
        async function expectExcel(targetPage = page) {
            try { await targetPage.waitForFunction(() => document.querySelector('#loading').style.display === 'none'); }
            catch (error) { console.log('LOADING', await targetPage.locator('#loading').textContent(), errors); throw error; }
            const frame = targetPage.frameLocator('#luckysheet-iframe');
            await frame.locator('#luckysheet-cell-main').waitFor();
            const result = await targetPage.frames().find(f => f.url().includes('luckysheet_iframe')).evaluate(() => {
                const sheets = luckysheet.getAllSheets();
                return { names: sheets.map(s => s.name), value: sheets[0].data[1][1].v,
                    formulaResult: sheets[0].data[2][1].v, merge: sheets[0].config.merge['0_0'],
                    date: sheets[0].data[1][2].m };
            });
            assert.deepEqual(result.names, ['Данные', 'Итоги', 'Пустой']);
            assert.equal(result.value, 42);
            assert.equal(result.formulaResult, 84);
            assert.equal(result.merge.cs, 2);
            assert.match(result.date, /\d{2}\.\d{2}\.\d{4}/);
            return frame;
        }
        await open('xlsx', 'Таблица.xlsx');
        let frame = await expectExcel();
        await saveAndVerify(page, '#downloadBtn', 'Таблица.xlsx', xlsx);
        await frame.locator('.luckysheet-sheets-item').filter({ hasText: 'Итоги' }).click();
        await page.waitForFunction(() => document.querySelector('#luckysheet-iframe').contentWindow.luckysheet.getSheet().name === 'Итоги');
        console.log('PASS XLSX data, sheets, empty sheet, merged cells, dates and cached formulas');
        await open('xls', 'Таблица.xls');
        await expectExcel();
        console.log('PASS legacy XLS');
        await open('xlsx', 'Документ');
        await expectExcel();
        assert.equal(await page.locator('#fileName').textContent(), 'Документ.xlsx');
        await page.evaluate(() => { LuckyExcel.transformExcelToLucky = (_file, _ok, fail) => fail(new Error('Unsupported workbook')); });
        await page.evaluate(() => loadFile());
        await expectExcel();
        await page.evaluate(() => { LuckyExcel.transformExcelToLucky = () => {}; });
        await page.evaluate(() => loadFile());
        await expectExcel();
        console.log('PASS opaque Excel URL, converter errors and nonresponding converter fallback');
        await open('pdf', 'Отчёт');
        await page.waitForSelector('.pdf-page');
        await page.waitForFunction(() => document.querySelector('#loading').style.display === 'none');
        assert.equal(await page.locator('#fileName').textContent(), 'Отчёт.pdf');
        assert.match(cookieHeader, /session=test/);
        await saveAndVerify(page, '#downloadBtn', 'Отчёт.pdf', pdf);
        console.log('PASS PDF signature detection, real PDF.js worker, session cookies');
        await open('multipage', 'Pages.pdf');
        await page.waitForFunction(() => document.querySelectorAll('.pdf-page').length === 3);
        assert.equal(await page.locator('#loading').isVisible(), false);
        console.log('PASS multipage PDF with standard font');
        for (const [endpoint, name, expected] of [['docx', 'Документ', 'Документ.docx'], ['named', 'Document', 'Тест.docx']]) {
            await open(endpoint, name);
            await page.waitForFunction(() => document.querySelector('#docx-container').textContent.includes('Проверка DOCX'));
            assert.equal(await page.locator('#fileName').textContent(), expected);
        }
        await saveAndVerify(page, '#downloadBtn', 'Тест.docx', docx);
        for (const [endpoint, provided, expected, bytes] of [
            ['opaque-name', 'Отчёт для группы', 'Отчёт для группы.pdf', pdf],
            ['mojibake-name', 'Document', 'Практическая работа.docx', docx],
            ['utf8-name', 'Document', 'План.pdf', pdf]
        ]) {
            await open(endpoint, provided);
            await page.waitForFunction(() => document.querySelector('#loading').style.display === 'none');
            assert.equal(await page.locator('#fileName').textContent(), expected);
            await saveAndVerify(page, '#downloadBtn', expected, bytes);
        }
        console.log('PASS native downloads: Cyrillic names, extensions, storage IDs, UTF-8 and misencoded headers');
        console.log('PASS real DOCX renderer, opaque URLs, UTF-8 Content-Disposition');
        for (const [endpoint, name, message] of [['broken', 'test.xlsx', 'Файл не является книгой Excel'], ['broken', 'test.xls', 'Файл не является книгой Excel'], ['broken', 'test.pdf', 'Ошибка:'], ['broken', 'test.docx', 'Ошибка:'], ['login', 'test.docx', 'Вместо документа'], ['empty', 'test.pdf', 'пустой файл'], ['denied', 'test.docx', 'HTTP 403']]) {
            await open(endpoint, name);
            await page.waitForFunction(text => document.querySelector('#loading').textContent.includes(text), message);
        }
        console.log('PASS corrupt PDF/DOCX, HTML login response, empty file, HTTP 403');
        await open('archive', 'archive.zip');
        await page.locator('#zip-list li').filter({ hasText: 'inner.xlsx' }).locator('button').click();
        await expectExcel();
        await open('archive', 'archive.zip');
        await page.locator('#zip-list li').filter({ hasText: 'inner.pdf' }).locator('button').click();
        await page.waitForSelector('.pdf-page');
        assert.equal(await page.locator('.pdf-page').count(), 1);
        assert.equal(await page.locator('#downloadOriginalBtn').isVisible(), true);
        await saveAndVerify(page, '#downloadBtn', 'inner.pdf', pdf);
        await saveAndVerify(page, '#downloadOriginalBtn', 'archive.zip', archiveBytes);
        console.log('PASS inner PDF rendering and archive download control');
        const source = await context.newPage();
        await source.goto('https://web.max.ru/');
        const popupPromise = context.waitForEvent('page');
        await source.locator('#file').click();
        const popup = await popupPromise;
        await popup.waitForLoadState();
        assert.equal(new URL(popup.url()).searchParams.get('name'), 'report.pdf');
        await popup.close();
        const blobUrl = await source.evaluate(bytes => {
            const url = URL.createObjectURL(new Blob([new Uint8Array(bytes)]));
            const a = document.querySelector('#file');
            a.href = url; a.innerText = 'Файл.docx\n600 КБ';
            return url;
        }, [...docx]);
        const blobPopupPromise = context.waitForEvent('page');
        await source.locator('#file').click();
        const blobPopup = await blobPopupPromise;
        await blobPopup.waitForFunction(() => document.querySelector('#docx-container').textContent.includes('Проверка DOCX'));
        assert.equal(new URL(blobPopup.url()).searchParams.get('url'), blobUrl);
        await saveAndVerify(blobPopup, '#downloadBtn', 'Файл.docx', docx);
        await source.evaluate(bytes => {
            const a = document.querySelector('#file');
            a.href = URL.createObjectURL(new Blob([new Uint8Array(bytes)]));
            a.innerText = 'Таблица.xlsx\n20 КБ';
        }, [...xlsx]);
        const excelPopupPromise = context.waitForEvent('page');
        await source.locator('#file').click();
        await expectExcel(await excelPopupPromise);
        console.log('PASS page-owned XLSX blob');
        console.log('PASS content-script filename detection and page-owned blob transfer across multiple chunks');
        assert.deepEqual(errors, []);
    } finally {
        await context.close();
        fs.rmSync(profile, { recursive: true, force: true });
    }
})().catch(error => { console.error(error); process.exitCode = 1; });
