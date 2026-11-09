
const urlParams = new URLSearchParams(window.location.search);
const fileUrl = urlParams.get('url');
const suppliedFileName = urlParams.get('name') || '';
let originalFileName = MAXFileNames.resolve({ provided: suppliedFileName, url: fileUrl });
const sourceTab = Number(urlParams.get('sourceTab'));

window.onerror = function(msg, url, lineNo, columnNo, error) {
    const loadingEl = document.getElementById('loading');
    if(loadingEl) {
        loadingEl.style.display = 'block';
        loadingEl.textContent = `Ошибка: ${msg} (строка ${lineNo})`;
    }
    return false;
};

document.getElementById('fileName').textContent = originalFileName;


let currentBlobUrl = null;
let currentFileName = originalFileName;
let originalBlobUrl = null;

async function downloadFile(url, name) {
    if (!url) return;
    const filename = MAXFileNames.clean(name) || 'Document';
    if (chrome?.downloads) {
        // onDeterminingFilename must explicitly return this name for blob URLs.
        const key = 'max_download_name:' + url;
        await chrome.storage.session.set({ [key]: { filename, created: Date.now() } });
        try {
            await chrome.downloads.download({ url, filename, conflictAction: 'uniquify' });
        } catch (error) {
            await chrome.storage.session.remove(key);
            throw error;
        }
    } else {
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
    }
}

document.getElementById('downloadBtn').addEventListener('click', () => {
    downloadFile(currentBlobUrl || fileUrl, currentFileName).catch(showError);
});

document.getElementById('downloadOriginalBtn')?.addEventListener('click', () => {
    downloadFile(originalBlobUrl || fileUrl, originalFileName).catch(showError);
});

function hideAllContainers() {
    ['loading', 'docx-container', 'luckysheet-iframe', 'pdf-container', 'rtf-container', 'img-container', 'zip-container', 'text-container', 'media-container', 'unsupported-container', 'odt-container', 'pptx-container', 'code-container'].forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.style.display = 'none';
            if (id === 'zip-container') document.getElementById('zip-list').innerHTML = '';
            if (id === 'text-container') el.textContent = '';
            if (id === 'media-container') el.innerHTML = '';
            if (id === 'rtf-container' || id === 'odt-container' || id === 'pptx-container') el.innerHTML = '';
            if (id === 'docx-container') el.innerHTML = '';
        }
    });
}

async function renderBlob(blob, fileName, isInnerFile = false) {
    // ZIP extraction and some servers label binary documents as text/plain.
    // Chrome may otherwise replace the requested extension at download time.
    blob = new Blob([blob], { type: MAXFileNames.mimeFor(fileName) });
    if (currentBlobUrl && currentBlobUrl !== originalBlobUrl) {
        URL.revokeObjectURL(currentBlobUrl);
    }
    currentBlobUrl = URL.createObjectURL(blob);
    currentFileName = MAXFileNames.clean(fileName) || 'Document';
    if (!isInnerFile) {
        if (originalBlobUrl) URL.revokeObjectURL(originalBlobUrl);
        originalBlobUrl = currentBlobUrl;
    }
    
    const originalBtn = document.getElementById('downloadOriginalBtn');
    if (originalBtn) {
        originalBtn.style.display = isInnerFile ? 'inline-block' : 'none';
    }

    hideAllContainers();
    const loadingEl = document.getElementById('loading');
    loadingEl.style.display = 'block';
    loadingEl.innerHTML = 'Обработка файла...';
    
    document.getElementById('fileName').textContent = fileName;
    
    const extMatch = fileName.match(/\.([a-zA-Z0-9]+)$/);
    let ext = extMatch ? extMatch[1].toLowerCase() : '';

    try {
        if (ext === 'rtf') {
            loadingEl.style.display = 'none';
            const container = document.getElementById('rtf-container');
            container.style.display = 'block';
            
            const buffer = await blob.arrayBuffer();
            RTFJS.loggingEnabled(false);
            const doc = new RTFJS.Document(new Uint8Array(buffer));
            const elements = await doc.render();
            elements.forEach(el => container.appendChild(el));
        } else if (ext === 'png' || ext === 'jpeg' || ext === 'jpg') {
            loadingEl.style.display = 'none';
            const container = document.getElementById('img-container');
            document.getElementById('img-preview').src = URL.createObjectURL(blob);
            container.style.display = 'block';
        } else if (ext === 'pdf') {
            await renderPdf(blob);
        } else if (ext === 'docx') {
            loadingEl.style.display = 'none';
            const container = document.getElementById('docx-container');
            container.style.display = 'block';
            await docx.renderAsync(await blob.arrayBuffer(), container);
        } else if (ext === 'xlsx' || ext === 'xls') {
            await renderExcel(blob, fileName);
        } else if (ext === 'pptx') {
            loadingEl.style.display = 'none';
            const container = document.getElementById('pptx-container');
            container.style.display = 'block';
            
            try {
                const jszip = new JSZip();
                const zip = await jszip.loadAsync(blob);
                
                // Find all slide xml files
                const slideFiles = [];
                zip.folder("ppt/slides").forEach((relativePath, file) => {
                    if (relativePath.match(/^slide\d+\.xml$/)) {
                        slideFiles.push({
                            name: relativePath,
                            file: file,
                            num: parseInt(relativePath.replace('slide', '').replace('.xml', ''))
                        });
                    }
                });
                
                slideFiles.sort((a, b) => a.num - b.num);
                
                if (slideFiles.length === 0) {
                    throw new Error('Не найдено слайдов в презентации');
                }
                
                for (let slide of slideFiles) {
                    const slideDiv = document.createElement('div');
                    slideDiv.style.border = '1px solid #ccc';
                    slideDiv.style.marginBottom = '20px';
                    slideDiv.style.padding = '20px';
                    slideDiv.style.boxShadow = '0 4px 12px rgba(0,0,0,0.1)';
                    slideDiv.style.background = 'white';
                    slideDiv.style.minHeight = '300px';
                    slideDiv.style.borderRadius = '8px';
                    slideDiv.style.position = 'relative';
                    
                    const slideTitle = document.createElement('div');
                    slideTitle.style.position = 'absolute';
                    slideTitle.style.top = '10px';
                    slideTitle.style.left = '10px';
                    slideTitle.style.color = '#999';
                    slideTitle.style.fontSize = '12px';
                    slideTitle.innerText = 'Слайд ' + slide.num;
                    slideDiv.appendChild(slideTitle);
                    
                    const contentDiv = document.createElement('div');
                    contentDiv.style.marginTop = '20px';
                    contentDiv.style.fontSize = '18px';
                    contentDiv.style.lineHeight = '1.5';
                    
                    const xmlString = await slide.file.async('string');
                    const parser = new DOMParser();
                    const doc = parser.parseFromString(xmlString, "text/xml");
                    
                    
                    // Extract images
                    const relsFile = zip.file(`ppt/slides/_rels/slide${slide.num}.xml.rels`);
                    const relMap = {};
                    if (relsFile) {
                        const relsXml = await relsFile.async('string');
                        const relsDoc = parser.parseFromString(relsXml, 'text/xml');
                        const rels = relsDoc.getElementsByTagName('Relationship');
                        for (let rel of rels) {
                            relMap[rel.getAttribute('Id')] = rel.getAttribute('Target');
                        }
                    }

                    // Elements inside spTree can be text or pictures
                    const spTree = doc.getElementsByTagName('p:spTree')[0];
                    if (spTree) {
                        for (let child of spTree.children) {
                            if (child.tagName === 'p:sp') {
                                // Text shape
                                const paragraphs = child.getElementsByTagName('a:p');
                                for (let p of paragraphs) {
                                    const pEl = document.createElement('p');
                                    pEl.style.margin = '0 0 10px 0';
                                    const texts = p.getElementsByTagName('a:t');
                                    let pText = '';
                                    for (let t of texts) {
                                        pText += t.textContent;
                                    }
                                    if (pText.trim() === '') pEl.style.minHeight = '1em';
                                    else pEl.textContent = pText;
                                    contentDiv.appendChild(pEl);
                                }
                            } else if (child.tagName === 'p:pic') {
                                // Picture
                                const blip = child.getElementsByTagName('a:blip')[0];
                                if (blip) {
                                    const embedId = blip.getAttribute('r:embed');
                                    const target = relMap[embedId];
                                    if (target) {
                                        // target is usually "../media/image1.png"
                                        const imgPath = target.replace('../', 'ppt/');
                                        const imgFile = zip.file(imgPath);
                                        if (imgFile) {
                                            const imgBlob = await imgFile.async('blob');
                                            const imgEl = document.createElement('img');
                                            imgEl.src = URL.createObjectURL(imgBlob);
                                            imgEl.style.maxWidth = '100%';
                                            imgEl.style.maxHeight = '300px';
                                            imgEl.style.display = 'block';
                                            imgEl.style.margin = '10px 0';
                                            contentDiv.appendChild(imgEl);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    slideDiv.appendChild(contentDiv);

                    container.appendChild(slideDiv);
                }
            } catch (err) {
                container.innerHTML = '<div style="color:red; font-family:sans-serif;">Ошибка чтения PPTX: ' + err.message + '</div>';
            }
        
        } else if (ext === 'odt') {
            loadingEl.style.display = 'none';
            const container = document.getElementById('odt-container');
            container.style.display = 'block';
            
            try {
                const jszip = new JSZip();
                const zip = await jszip.loadAsync(blob);
                if (!zip.file('content.xml')) throw new Error('Некорректный ODT файл: отсутствует content.xml');
                
                const xmlString = await zip.file('content.xml').async('string');
                const parser = new DOMParser();
                const doc = parser.parseFromString(xmlString, "text/xml");
                
                // Простая конвертация ODT XML -> HTML
                function convertNode(node) {
                    if (node.nodeType === 3) return document.createTextNode(node.nodeValue); // Text node
                    if (node.nodeType !== 1) return null; // Only elements
                    
                    let el = null;
                    const tag = node.tagName.toLowerCase();
                    
                    if (tag === 'text:p') el = document.createElement('p');
                    else if (tag === 'text:h') {
                        const level = node.getAttribute('text:outline-level') || '1';
                        el = document.createElement('h' + Math.min(level, 6));
                    }
                    else if (tag === 'text:span') el = document.createElement('span');
                    else if (tag === 'text:a') {
                        el = document.createElement('a');
                        el.href = node.getAttribute('xlink:href') || '#';
                        el.style.color = 'blue';
                        el.style.textDecoration = 'underline';
                    }
                    else if (tag === 'text:list') el = document.createElement('ul');
                    else if (tag === 'text:list-item') el = document.createElement('li');
                    else if (tag === 'table:table') {
                        el = document.createElement('table');
                        el.style.borderCollapse = 'collapse';
                        el.style.width = '100%';
                        el.style.marginBottom = '1em';
                    }
                    else if (tag === 'table:table-row') el = document.createElement('tr');
                    else if (tag === 'table:table-cell') {
                        el = document.createElement('td');
                        el.style.border = '1px solid #ccc';
                        el.style.padding = '4px 8px';
                    }
                    else if (tag === 'draw:image') {
                        el = document.createElement('img');
                        el.style.maxWidth = '100%';
                        const href = node.getAttribute('xlink:href');
                        if (href && zip.file(href)) {
                            zip.file(href).async("blob").then(imgBlob => {
                                el.src = URL.createObjectURL(imgBlob);
                            });
                        }
                    }
                    else if (tag === 'text:s') {
                        el = document.createElement('span');
                        const spaces = parseInt(node.getAttribute('text:c') || '1', 10);
                        el.innerHTML = '&nbsp;'.repeat(spaces);
                    }
                    else if (tag === 'text:tab') {
                        el = document.createElement('span');
                        el.innerHTML = '&nbsp;&nbsp;&nbsp;&nbsp;';
                    }
                    
                    if (!el) {
                        el = document.createElement('span');
                    }
                    
                    if (el.tagName === 'P') {
                        el.style.margin = '0 0 1em 0';
                        el.style.minHeight = '1em';
                    }
                    
                    for (let child of node.childNodes) {
                        const converted = convertNode(child);
                        if (converted) el.appendChild(converted);
                    }
                    
                    return el;
                }
                
                const textElement = doc.getElementsByTagName('office:text')[0];
                if (textElement) {
                    for (let child of textElement.childNodes) {
                        const converted = convertNode(child);
                        if (converted) container.appendChild(converted);
                    }
                }
                
            } catch (err) {
                container.innerHTML = '<div style="color:red; font-family:sans-serif;">Ошибка чтения ODT: ' + err.message + '</div>';
            }
        } else if (ext === 'zip') {
            loadingEl.style.display = 'none';
            const container = document.getElementById('zip-container');
            const list = document.getElementById('zip-list');
            container.style.display = 'block';
            
            try {
                const jszip = new JSZip();
                const zip = await jszip.loadAsync(blob);
                
                Object.keys(zip.files).forEach(filename => {
                    const file = zip.files[filename];
                    const li = document.createElement('li');
                    li.style.padding = '8px 0';
                    li.style.borderBottom = '1px solid #eee';
                    li.style.display = 'flex';
                    li.style.alignItems = 'center';
                    li.style.justifyContent = 'space-between';
                    
                    const left = document.createElement('div');
                    
                    if (file.dir) {
                        left.innerHTML = '📁 <strong>' + filename + '</strong>';
                        li.appendChild(left);
                    } else {
                        const size = file._data && file._data.uncompressedSize ? (file._data.uncompressedSize / 1024).toFixed(1) + ' KB' : '';
                        left.innerHTML = '📄 ' + filename + (size ? ' <span style="color:#999;font-size:0.9em;">(' + size + ')</span>' : '');
                        
                        const btn = document.createElement('button');
                        btn.textContent = 'Открыть';
                        btn.style.padding = '4px 12px';
                        btn.style.cursor = 'pointer';
                        btn.style.border = '1px solid #ccc';
                        btn.style.background = '#fff';
                        btn.style.borderRadius = '4px';
                        
                        btn.onclick = async () => {
                            try {
                                btn.textContent = 'Загрузка...';
                                btn.disabled = true;
                                const fileBlob = await file.async("blob");
                                // We don't have mime type perfectly here, but renderBlob uses extension!
                                await renderBlob(fileBlob, filename, true);
                            } catch (e) {
                                alert("Ошибка открытия файла: " + e.message);
                                btn.textContent = 'Открыть';
                                btn.disabled = false;
                            }
                        };
                        
                        li.appendChild(left);
                        li.appendChild(btn);
                    }
                    list.appendChild(li);
                });
            } catch (err) {
                list.innerHTML = '<li style="color:red;">Ошибка чтения архива: ' + err.message + '</li>';
            }
        } else if (['py','js','c','cpp','java','cs','go','php','rb','swift','ts','sh','html','css','xml','json','md'].includes(ext) || ext === 'txt' || ext === 'csv') {

            if (['py','js','c','cpp','java','cs','go','php','rb','swift','ts','sh','html','css','xml','json','md'].includes(ext)) {
                loadingEl.style.display = 'none';
                const container = document.getElementById('code-container');
                container.style.display = 'block';
                
                const codeContent = await blob.text();
                const codeEl = document.getElementById('code-content');
                codeEl.textContent = codeContent;
                codeEl.className = 'language-' + ext;
                
                document.getElementById('code-title').textContent = 'Исходный код (' + ext + ')';
                
                // Highlight syntax
                if (window.hljs) {
                    hljs.highlightElement(codeEl);
                }
                
                // Setup Run Button
                const runBtn = document.getElementById('runCodeBtn');
                const outputWrapper = document.getElementById('code-output-wrapper');
                const outputEl = document.getElementById('code-output');
                const iframe = document.getElementById('code-runner-iframe');
                
                outputWrapper.style.display = 'none';
                outputEl.textContent = '';
                
                if (ext === 'js' || ext === 'py') {
                    runBtn.style.display = 'block';
                    runBtn.onclick = () => {
                        outputWrapper.style.display = 'block';
                        outputEl.textContent = 'Выполнение...\n';
                        outputEl.style.color = '#d4d4d4';
                        
                        // Send message to sandboxed iframe
                        iframe.contentWindow.postMessage({
                            action: 'runCode',
                            code: codeContent,
                            lang: ext,
                            messageId: Date.now()
                        }, '*');
                    };
                    
                    // Listen for results
                    window.addEventListener('message', (event) => {
                        if (event.source !== iframe.contentWindow) return;
                        if (event.data.action === 'runCodeResult') {
                            outputEl.textContent = event.data.output || '(нет вывода)\n';
                            if (event.data.isError) {
                                outputEl.style.color = '#ff5555';
                            } else {
                                outputEl.style.color = '#d4d4d4';
                            }
                            // Auto scroll to bottom
                            outputEl.scrollTop = outputEl.scrollHeight;
                        }
                    });
                    
                } else {
                    runBtn.style.display = 'none';
                }
                
            } else if (['txt','csv'].includes(ext)) {
                loadingEl.style.display = 'none';
                const container = document.getElementById('text-container');
                container.style.display = 'block';
                container.textContent = await blob.text();
            }
        } else if (['mp4','webm','ogg','mp3','wav'].includes(ext)) {
            loadingEl.style.display = 'none';
            const container = document.getElementById('media-container');
            container.style.display = 'flex';
            
            const isVideo = ['mp4','webm','ogg'].includes(ext);
            const media = document.createElement(isVideo ? 'video' : 'audio');
            media.controls = true;
            media.style.maxWidth = '100%';
            media.style.maxHeight = '100%';
            media.src = URL.createObjectURL(blob);
            container.appendChild(media);
        } else {
            loadingEl.style.display = 'none';
            document.getElementById('unsupported-container').style.display = 'flex';
        }
    } catch (e) {
        console.error(e);
        loadingEl.style.display = 'block';
        showError(e);
    }
}

function showError(error) {
    const loadingEl = document.getElementById('loading');
    loadingEl.style.display = 'block';
    loadingEl.textContent = 'Ошибка: ' + error.message;
}

async function readSourceBlob() {
    if (!urlParams.has('sourceTab') || !Number.isInteger(sourceTab)) {
        throw new Error('Откройте файл заново из вкладки MAX.');
    }
    const chunks = [];
    let offset = 0;
    let result;
    try {
        do {
            result = await chrome.tabs.sendMessage(sourceTab, {
                action: 'readPreviewBlob', url: fileUrl, offset
            });
            if (!result || result.error) throw new Error(result?.error || 'Вкладка MAX недоступна');
            const bytes = Uint8Array.from(atob(result.data), char => char.charCodeAt(0));
            if (!bytes.length && offset < result.size) throw new Error('Неполные данные файла');
            chunks.push(bytes);
            offset += bytes.length;
        } while (offset < result.size);
        return new Blob(chunks, { type: result.type });
    } finally {
        chrome.tabs.sendMessage(sourceTab, { action: 'releasePreviewBlob', url: fileUrl }).catch(() => {});
    }
}

async function loadFile() {
    try {
        if (!fileUrl) throw new Error('Не указан адрес файла. Откройте документ из MAX.');
        document.getElementById('downloadBtn').disabled = true;
        document.getElementById('loading').textContent = 'Загрузка файла...';
        let blob;
        let contentDisposition = '';
        if (fileUrl.startsWith('blob:')) {
            blob = await readSourceBlob();
        } else {
            const response = await fetch(fileUrl, { credentials: 'include' });
            if (!response.ok) throw new Error('HTTP ' + response.status + '. Откройте файл заново из MAX.');
            contentDisposition = response.headers.get('content-disposition') || '';
            blob = await response.blob();
        }
        if (!blob.size) throw new Error('Получен пустой файл.');
        const type = blob.type.split(';')[0];
        if (type === 'text/html' && !/\.html$/i.test(originalFileName)) {
            throw new Error('Вместо документа сервер вернул страницу. Возможно, ссылка устарела или требуется вход в MAX.');
        }
        originalFileName = MAXFileNames.resolve({ provided: suppliedFileName, header: contentDisposition, url: fileUrl, mime: blob.type });
        currentFileName = originalFileName;
        const signature = new Uint8Array(await blob.slice(0, 1024).arrayBuffer());
        const isZip = signature[0] === 0x50 && signature[1] === 0x4b;
        let detected = '';
        if (!isZip && new TextDecoder().decode(signature).includes('%PDF-')) detected = 'pdf';
        else if (isZip) {
            const zip = await JSZip.loadAsync(await blob.arrayBuffer());
            if (zip.file('word/document.xml')) detected = 'docx';
            else if (zip.file('xl/workbook.xml')) detected = 'xlsx';
            else if (zip.file('ppt/presentation.xml')) detected = 'pptx';
            else if (zip.file('mimetype') && (await zip.file('mimetype').async('string')).trim() === 'application/vnd.oasis.opendocument.text') detected = 'odt';
            else detected = 'zip';
        } else if (signature[0] === 0xd0 && signature[1] === 0xcf && !MAXFileNames.extension(suppliedFileName)) detected = 'xls';
        originalFileName = MAXFileNames.resolve({ provided: suppliedFileName, header: contentDisposition, url: fileUrl, mime: blob.type, detected });
        currentFileName = originalFileName;
        await renderBlob(blob, originalFileName);
    } catch (error) {
        console.error(error);
        showError(error);
        // CDN hosts vary. Ask only for the actual origin, on a user gesture.
        if (error instanceof TypeError && /^https:/.test(fileUrl || '') && chrome.permissions) {
            const button = document.createElement('button');
            button.textContent = 'Разрешить загрузку с сервера файла и повторить';
            button.addEventListener('click', async () => {
                const origin = new URL(fileUrl).origin + '/*';
                if (await chrome.permissions.request({ origins: [origin] })) {
                    await loadFile();
                }
            });
            document.getElementById('loading').appendChild(button);
        }
    } finally {
        document.getElementById('downloadBtn').disabled = false;
    }
}

async function renderPdf(blob) {
    const container = document.getElementById('pdf-container');
    const loadingEl = document.getElementById('loading');
    container.replaceChildren();
    container.style.display = 'block';
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'libs/pdfjs/pdf.worker.min.js';
    const task = pdfjsLib.getDocument({
        data: new Uint8Array(await blob.arrayBuffer()),
        cMapUrl: 'libs/pdfjs/cmaps/',
        cMapPacked: true,
        standardFontDataUrl: 'libs/pdfjs/standard_fonts/',
        isEvalSupported: false
    });
    task.onPassword = (updatePassword, reason) => {
        const password = window.prompt(reason === 2 ? 'Неверный пароль. Введите пароль PDF:' : 'Введите пароль PDF:');
        if (password === null) {
            task.destroy();
            showError(new Error('Ввод пароля отменён. Откройте файл заново для повторной попытки.'));
        } else updatePassword(password);
    };
    const pdf = await task.promise;
    loadingEl.style.display = 'none';
    try {
        for (let number = 1; number <= pdf.numPages; number++) {
            const page = await pdf.getPage(number);
            const base = page.getViewport({ scale: 1 });
            const scale = Math.min(1.5, Math.max(1, container.clientWidth - 40) / base.width);
            const viewport = page.getViewport({ scale });
            const canvas = document.createElement('canvas');
            canvas.className = 'pdf-page';
            canvas.width = Math.ceil(viewport.width);
            canvas.height = Math.ceil(viewport.height);
            container.appendChild(canvas);
            await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
            page.cleanup();
        }
    } finally {
        await pdf.destroy();
    }
}

loadFile();
