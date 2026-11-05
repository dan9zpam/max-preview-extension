
chrome.downloads.onDeterminingFilename.addListener((downloadItem, suggest) => {
    if (downloadItem.byExtensionId === chrome.runtime.id) {
        suggest();
        return;
    }

    const filename = downloadItem.filename.toLowerCase();
    
    // Получаем время последнего клика из storage
    chrome.storage.local.get(['max_last_click'], (res) => {
        let isFromMax = false;
        
        // 1. Если клик на странице был меньше 5 секунд назад - это 100% наш файл!
        const lastClick = res.max_last_click || 0;
        if (Date.now() - lastClick < 5000) {
            isFromMax = true;
        }
        
        // 2. Проверяем URL и Referrer (как страховка)
        if ((downloadItem.url && downloadItem.url.includes('max.ru')) || 
            (downloadItem.referrer && downloadItem.referrer.includes('max.ru')) ||
            (downloadItem.finalUrl && downloadItem.finalUrl.includes('max.ru'))) {
            isFromMax = true;
        }
        
        // 3. Дополнительно проверяем вкладки
        chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
            if (tabs && tabs.length > 0 && tabs[0].url && tabs[0].url.includes('max.ru')) {
                isFromMax = true;
            }
            
            if (!isFromMax) {
                suggest();
                return;
            }

            if (filename.match(/\.(pdf|docx|xlsx|xls|png|jpeg|jpg|rtf|zip|pptx|odt|txt|csv|py|c|cpp|java|cs|go|php|rb|swift|ts|sh|json|xml|md|js|css|html|mp3|wav|ogg|mp4|webm|avi|mov|mkv)$/i)) {
                chrome.downloads.cancel(downloadItem.id, () => {
                    const viewerUrl = chrome.runtime.getURL(`viewer/viewer.html?url=${encodeURIComponent(downloadItem.finalUrl || downloadItem.url)}&name=${encodeURIComponent(downloadItem.filename)}&sourceTab=${tabs[0] ? tabs[0].id : ''}`);
                    chrome.tabs.create({ url: viewerUrl });
                    suggest(); // REQUIRED BY CHROME API
                });
            } else {
                suggest();
            }
        });
    });

    return true; // Асинхронный вызов
});

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (message.action === 'openViewer') {
        if (!sender.tab || !sender.url || new URL(sender.url).hostname !== 'web.max.ru') return;
        const viewerUrl = chrome.runtime.getURL(`viewer/viewer.html?url=${encodeURIComponent(message.fileUrl)}&name=${encodeURIComponent(message.fileName)}&sourceTab=${sender.tab.id}`);
        chrome.tabs.create({ url: viewerUrl });
    }
});
