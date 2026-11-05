document.addEventListener('click', (e) => {
    // Пытаемся найти ссылку, по которой кликнули (или родительский элемент ссылки)
    const target = e.target.closest('a');
    
    if (target && target.href) {
        const url = target.href;
        let urlName;
        try { urlName = decodeURIComponent(new URL(url).pathname.split('/').pop()); }
        catch (_) { urlName = ''; }
        const supported = /\.(pdf|docx|xlsx|xls|png|jpeg|jpg|rtf|zip|odt|pptx|txt|csv|json|xml|md|js|css|html|py|c|cpp|java|cs|go|php|rb|swift|ts|sh)$/i;
        const label = (target.getAttribute('download') || target.innerText || '').trim();
        const labelName = label.split(/\r?\n/).map(line => line.trim()).find(line => supported.test(line));
        const fileName = labelName || urlName;
        if (supported.test(fileName)) {
            e.preventDefault();
            e.stopImmediatePropagation();

            // Отправляем сообщение в background script для открытия вкладки
            chrome.runtime.sendMessage({
                action: 'openViewer',
                fileUrl: url,
                fileName: fileName
            });
        }
    }
}, true); // Используем capture phase, чтобы перехватить клик до того, как его обработает React/Vue





// Вспомогательная функция для получения точного CSS пути элемента
function getCssPath(el) {
    if (!(el instanceof Element)) return '';
    let path = [];
    while (el.nodeType === Node.ELEMENT_NODE) {
        let selector = el.nodeName.toLowerCase();
        if (el.id) {
            selector += '#' + el.id;
            path.unshift(selector);
            break;
        } else {
            let sib = el, nth = 1;
            while (sib = sib.previousElementSibling) {
                if (sib.nodeName.toLowerCase() === selector) nth++;
            }
            if (nth !== 1 || el.nextElementSibling) selector += ":nth-of-type("+nth+")";
        }
        path.unshift(selector);
        el = el.parentNode;
    }
    return path.join(" > ");
}



// Запоминаем последний клик в левой части экрана (список чатов)
document.addEventListener('click', (e) => {
    try {
        if (chrome && chrome.storage && chrome.storage.local) {
            chrome.storage.local.set({ max_last_click: Date.now() });
        }
    } catch (err) {}

    try {
        if (e.clientX < window.innerWidth * 0.5) {
            let target = e.target.closest('a') || e.target.closest('li') || e.target.closest('[class*="item"]') || e.target;
            
            // 1. Пробуем найти явный идентификатор (href или data-id)
            let identifier = target.getAttribute('href') || target.getAttribute('data-id');
            if (!identifier) {
                let parent = target.closest('[data-id], a');
                if (parent) identifier = parent.getAttribute('href') || parent.getAttribute('data-id');
            }

            if (identifier && identifier !== '#' && identifier !== '/' && !identifier.startsWith('javascript:')) {
                localStorage.setItem('max_last_chat_id', identifier);
                localStorage.removeItem('max_last_chat_text');
            } else {
                // 2. Если нет ID, берем самую длинную строку текста (обычно это имя контакта)
                let textLines = target.innerText ? target.innerText.split('\n').map(s=>s.trim()).filter(s=>s) : [];
                let longestText = '';
                for (let t of textLines) {
                    if (t.length > longestText.length && t.length < 50) {
                        longestText = t;
                    }
                }
                if (longestText.length >= 2) {
                    localStorage.setItem('max_last_chat_text', longestText);
                    localStorage.removeItem('max_last_chat_id');
                }
            }
            
            // Сохраняем URL через секунду, если он изменился (для SPA)
            setTimeout(() => {
                if (location.pathname.length > 2 || location.hash.length > 2) {
                    localStorage.setItem('max_last_url', location.href);
                }
            }, 1000);
        }
    } catch (err) { console.error(err); }
}, true);

// Восстанавливаем при загрузке
window.addEventListener('load', () => {
    // 1. Пытаемся восстановить по URL
    let lastUrl = localStorage.getItem('max_last_url');
    if (lastUrl && lastUrl !== location.href && (location.pathname === '/' || location.pathname === '' || location.pathname === '/messages')) {
        // Защита от бесконечного цикла редиректов
        if (!sessionStorage.getItem('max_redirected')) {
            sessionStorage.setItem('max_redirected', 'true');
            location.href = lastUrl;
            return;
        }
    }
    sessionStorage.removeItem('max_redirected');

    // 2. Пытаемся восстановить кликом по DOM
    let id = localStorage.getItem('max_last_chat_id');
    let text = localStorage.getItem('max_last_chat_text');
    if (!id && !text) return;

    let attempts = 0;
    let interval = setInterval(() => {
        attempts++;
        if (attempts > 15) { // 7.5 секунд
            clearInterval(interval);
            return;
        }
        
        let target = null;
        
        if (id) {
            try { target = document.querySelector(`[href="${id}"], [data-id="${id}"]`); } catch(e) {}
        }
        
        if (!target && text) {
            const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
            let node;
            while (node = walker.nextNode()) {
                if (node.nodeValue.trim() === text) {
                    let rect = node.parentElement.getBoundingClientRect();
                    if (rect.width > 0 && rect.left < window.innerWidth * 0.5) {
                        target = node.parentElement;
                        break;
                    }
                }
            }
        }
        
        if (target) {
            let clickable = target.closest('a') || target.closest('button') || target.closest('li') || target.closest('[class*="item"]') || target;
            clickable.click();
            clearInterval(interval);
        }
    }, 500);
});

// Blob URLs belong to the MAX tab. Read them there, in bounded chunks,
// rather than trying to fetch a page-owned URL from the extension origin.
const previewBlobs = new Map();
chrome.runtime.onMessage.addListener((message, sender, respond) => {
    if (sender.id !== chrome.runtime.id) return;
    if (message.action === 'readPreviewBlob') {
        (async () => {
            const url = new URL(message.url);
            if (url.protocol !== 'blob:' || url.origin !== location.origin) {
                throw new Error('Недопустимый адрес файла');
            }
            let blob = previewBlobs.get(message.url);
            if (!blob) {
                const response = await fetch(message.url);
                if (!response.ok) throw new Error('HTTP ' + response.status);
                blob = await response.blob();
                previewBlobs.set(message.url, blob);
                setTimeout(() => previewBlobs.delete(message.url), 60000);
            }
            const offset = Math.max(0, Number(message.offset) || 0);
            const bytes = new Uint8Array(await blob.slice(offset, offset + 262144).arrayBuffer());
            let binary = '';
            for (const byte of bytes) binary += String.fromCharCode(byte);
            respond({ data: btoa(binary), size: blob.size, type: blob.type });
        })().catch(error => respond({ error: error.message }));
        return true;
    }
    if (message.action === 'releasePreviewBlob') previewBlobs.delete(message.url);
});
