
        window.cleanLuckysheetCopy = function(html) {
            try {
                const div = document.createElement('div');
                div.innerHTML = html;
                let text = '';
                const rows = div.querySelectorAll('tr');
                for (let r = 0; r < rows.length; r++) {
                    const cols = rows[r].querySelectorAll('td, th');
                    let rowText = [];
                    for (let c = 0; c < cols.length; c++) {
                        let cellText = cols[c].innerText || cols[c].textContent || '';
                        // Remove newlines inside cells or replace them with spaces
                        cellText = cellText.replace(/\r\n|\n|\r/gm, ' ');
                        rowText.push(cellText);
                    }
                    text += rowText.join('\t') + (r < rows.length - 1 ? '\n' : '');
                }
                return text;
            } catch(e) {
                return html; // Fallback
            }
        };


        window.addEventListener('message', (event) => {
            if (event.source !== window.parent || event.origin !== window.location.origin || !event.data) return;
            if (event.data.action === 'pingExcel') {
                window.parent.postMessage({ action: 'iframeReady' }, window.location.origin);
                return;
            }
            if (event.data.action !== 'loadExcel') return;
            const requestId = event.data.requestId;
            const reply = message => window.parent.postMessage({ ...message, requestId }, window.location.origin);
            try {
                if (!window.luckysheet) throw new Error('Не удалось загрузить табличный просмотрщик.');
                try { window.luckysheet.destroy(); } catch (_) {}
                window.luckysheet.create({
                    container: 'luckysheet', data: event.data.sheets,
                    title: event.data.title, showinfobar: false, lang: 'en',
                    hook: { workbookCreateAfter: () => reply({ action: 'excelLoaded' }) }
                });
            } catch (error) {
                console.error('Excel create failed', error.stack || error);
                reply({ action: 'excelError', error: error.message || String(error) });
            }
        });
        window.parent.postMessage({ action: 'iframeReady' }, window.location.origin);
