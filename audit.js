let auditCardsHTML = [];
let auditRenderedCount = 0;
let auditObserver = null;

// ==========================================
// ЛОГІКА АУДИТУ КІЛЕЦЬ ТА ЛІКУВАННЯ
// ==========================================

let oldAuditRingsMap = {}; 
window.currentAuditResults = null; // Глобально зберігаємо результати перевірки
let currentAuditFilter = 'all'; // Поточний фільтр ('all', 'perfect', 'modified', 'healed', 'broken')
let healingContext = null; // Пам'ять для відкритого модального вікна лікування

const auditDropZone = document.getElementById('audit_drop_zone');
const auditFileInput = document.getElementById('audit_file_input');

// --- 1. ЗАВАНТАЖЕННЯ ФАЙЛУ ---
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    if(auditDropZone) auditDropZone.addEventListener(eventName, e => { e.preventDefault(); e.stopPropagation(); }, false);
});

let auditDragCounter = 0;

if(auditDropZone) {
    auditDropZone.addEventListener('dragenter', e => {
        auditDragCounter++;
        auditDropZone.classList.add('dragover');
    });

    auditDropZone.addEventListener('dragover', e => {
        if (!auditDropZone.classList.contains('dragover')) {
            auditDropZone.classList.add('dragover');
        }
    });

    auditDropZone.addEventListener('dragleave', e => {
        auditDragCounter--;
        if (auditDragCounter === 0) {
            auditDropZone.classList.remove('dragover');
        }
    });

    auditDropZone.onclick = () => auditFileInput.click();
    
    auditDropZone.addEventListener('drop', e => {
        // Обязательно сбрасываем счетчик при бросании файла
        auditDragCounter = 0;
        auditDropZone.classList.remove('dragover');
        handleAuditFile(e.dataTransfer.files[0]);
    });
}
if(auditFileInput) auditFileInput.onchange = e => handleAuditFile(e.target.files[0]);

function handleAuditFile(file) {
    if (!file) return;
    document.getElementById('audit-content').innerHTML = '<div class="empty-msg">⏳ Читаємо файл...</div>';
    
    const reader = new FileReader();
    reader.onload = e => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
        
        oldAuditRingsMap = {};
        let loadedTripsCount = 0;

        // Починаємо читати з 4-го рядка (де починаються самі дані)
        for (let i = 4; i < rows.length; i++) {
            const grf = rows[i][2]; // Стовпець C (індекс 2) — GRF
            const rawColumnE = String(rows[i][4] || ""); // Стовпець E (індекс 4) — Текст із кодом
            
            // Шукаємо маску NNNN_NN (4 цифри, підкреслення, 2 цифри)
            const match = rawColumnE.match(/\d{4}_\d{2}/);
            const kmRingName = match ? match[0] : null;

            if (grf && kmRingName) {
                let oldTrip = new Trip(rows[i]);
                
                // 🛠 ПАТЧ: Відновлюємо правильний астрономічний час для порожніх перегонів
                if (grf.startsWith('EMPTY_') && oldTrip.depStr) {
                    const [h, m] = oldTrip.depStr.split(':').map(Number);
                    let dM = (h * 60) + m;
                    let correctAstroDay = oldTrip.logisticDay;
                    // Якщо виїзд до 12:00, значить астрономічно це вже наступний день
                    if (dM < 720) correctAstroDay = (oldTrip.logisticDay + 1) % 7;
                    
                    oldTrip.trueStart = correctAstroDay * 1440 + dM;
                    
                    // Перевірка на перехід через тиждень (неділя -> понеділок)
                    if (oldTrip.logisticDay === 6 && oldTrip.trueStart < 1440) {
                        oldTrip.trueStart += 10080;
                    }
                }
                
                if (!oldAuditRingsMap[kmRingName]) oldAuditRingsMap[kmRingName] = [];
                oldAuditRingsMap[kmRingName].push(oldTrip);
                loadedTripsCount++;
            }
        }
        
        // 🧹 Спрощене і надійне сортування за абсолютним часом
        for (let ringName in oldAuditRingsMap) {
            oldAuditRingsMap[ringName].sort((a, b) => a.trueStart - b.trueStart);
        }
        
        document.getElementById('audit-content').innerHTML = `
            <div class="empty-msg" style="color: #1e8e3e;">
                ✅ Прочитано ${Object.keys(oldAuditRingsMap).length} кілець (графіків: ${loadedTripsCount}).<br><br>
                Натисніть "Запустити перевірку".
            </div>`;
        document.getElementById('btn_run_audit').style.display = 'inline-flex';
    };
    reader.readAsArrayBuffer(file);
}

// --- 2. ЯДРО ПЕРЕВІРКИ ---
// --- 2. ЯДРО ПЕРЕВІРКИ ---
function runAudit() {
    const btn = document.getElementById('btn_run_audit');
    btn.innerText = "⏳ Перевіряємо...";
    btn.disabled = true;

    const reparkMins = parseInt(document.getElementById('repark_time')?.value) || 30;
    const toleranceMins = parseInt(document.getElementById('audit_tolerance')?.value) || 120;
    const mode = document.getElementById('mode_select')?.value || 'node';

    let results = { perfect: [], modified: [], healed: [], broken: [] };
    let availableFreshTrips = window.allTrips.filter(t => t.ringId === null);

    for (let [ringName, oldTrips] of Object.entries(oldAuditRingsMap)) {
        let mappedTrips = [];
        let isMissing = false;
        let hasModifications = false;

        for (let i = 0; i < oldTrips.length; i++) {
            let oldT = oldTrips[i];
            let freshT = window.allTrips.find(t => t.grf === oldT.grf);

            if (!freshT) {
                isMissing = true;
                mappedTrips.push({ status: 'missing', oldTrip: oldT });
            } else {
                // Перевіряємо, чи були зміни. trueStart охоплює і час, і зміну дня тижня!
                let modDetails = { 
                    auto: oldT.auto !== freshT.auto, 
                    start: oldT.trueStart !== freshT.trueStart, 
                    end: oldT.trueEnd !== freshT.trueEnd,
                    day: oldT.logisticDay !== freshT.logisticDay // Фіксуємо зміну логістичного дня
                };
                
                let modified = modDetails.auto || modDetails.start || modDetails.end || modDetails.day;
                if (modified) hasModifications = true;
                
                mappedTrips.push({ status: 'found', freshTrip: freshT, oldTrip: oldT, modified: modified, modDetails: modDetails });
            }
        }

        // МАКСИМАЛЬНО ПРОСТА ЛОГІКА
        if (isMissing) {
            // Якщо хоч один графік випав - кільце поламане. Пробуємо знайти йому заміну.
            let healedRing = tryAutoHeal(mappedTrips, availableFreshTrips, reparkMins, toleranceMins, mode);
            if (healedRing) {
                results.healed.push({ ringName, trips: healedRing.trips });
                healedRing.usedTrips.forEach(ut => availableFreshTrips = availableFreshTrips.filter(t => t.id !== ut.id));
            } else {
                results.broken.push({ ringName, mappedTrips, reason: 'Випав графік (немає автозаміни)' });
            }
        } else {
            // Всі графіки на місці! Жодних перевірок на локації чи перепарковки між ними.
            if (hasModifications) {
                results.modified.push({ ringName, mappedTrips });
            } else {
                results.perfect.push({ ringName, mappedTrips });
            }
        }
    }

    window.currentAuditResults = results;
    currentAuditFilter = 'all';
    renderAuditView();
    
    btn.innerText = "🔍 Запустити перевірку";
    btn.disabled = false;
}

function checkSequenceValid(trips, reparkMins, mode) {
    for(let i = 0; i < trips.length - 1; i++) {
        let current = trips[i], next = trips[i+1];
        let effectiveEnd = current.trueEnd + (current.trueEnd < current.trueStart ? 10080 : 0);
        if (current.getPointName('dest', mode) !== next.getPointName('origin', mode)) return false;
        if (current.auto !== next.auto) return false;
        if (next.trueStart < effectiveEnd + reparkMins) return false;
    }
    return true;
}

function tryAutoHeal(mappedTrips, availableFreshTrips, reparkMins, toleranceMins, mode) {
    let healedTrips = [], usedTrips = [];
    for (let i = 0; i < mappedTrips.length; i++) {
        let item = mappedTrips[i];
        if (item.status === 'found') healedTrips.push(item);
        else if (item.status === 'missing') {
            let prev = i > 0 ? mappedTrips[i-1].freshTrip : null;
            let next = i < mappedTrips.length - 1 ? mappedTrips[i+1].freshTrip : null;
            let oldMissing = item.oldTrip;

            let candidate = availableFreshTrips.find(c => {
                if (c.auto !== oldMissing.auto || c.getPointName('origin', mode) !== oldMissing.getPointName('origin', mode) || c.getPointName('dest', mode) !== oldMissing.getPointName('dest', mode)) return false;
                if (Math.abs(c.trueStart - oldMissing.trueStart) > toleranceMins) return false;
                if (prev && c.trueStart < prev.trueEnd + (prev.trueEnd < prev.trueStart ? 10080 : 0) + reparkMins) return false;
                if (next && next.trueStart < c.trueEnd + (c.trueEnd < c.trueStart ? 10080 : 0) + reparkMins) return false;
                return true;
            });

            if (candidate) {
                healedTrips.push({ status: 'healed', freshTrip: candidate, oldTrip: oldMissing });
                usedTrips.push(candidate);
            } else return null;
        }
    }
    return { trips: healedTrips, usedTrips: usedTrips };
}

// --- 3. ІНТЕРФЕЙС ТА ФІЛЬТРИ ---
function setAuditFilter(filterType) {
    currentAuditFilter = (currentAuditFilter === filterType) ? 'all' : filterType;
    renderAuditView();
}

// --- 3. ІНТЕРФЕЙС ТА ФІЛЬТРИ ---
function setAuditFilter(filterType) {
    currentAuditFilter = (currentAuditFilter === filterType) ? 'all' : filterType;
    renderAuditView();
}

// --- 3. ІНТЕРФЕЙС ТА ФІЛЬТРИ ---
function setAuditFilter(filterType) {
    currentAuditFilter = (currentAuditFilter === filterType) ? 'all' : filterType;
    renderAuditView();
}

function renderAuditView() {
    const headerPanel = document.getElementById('audit-header-panel');
    if (!window.currentAuditResults) {
        headerPanel.innerHTML = '';
        return;
    }
    
    const res = window.currentAuditResults;
    const getOpacity = (type) => (currentAuditFilter === 'all' || currentAuditFilter === type) ? '1' : '0.4';
    const getScale = (type) => (currentAuditFilter === type) ? 'transform: scale(1.02); box-shadow: 0 4px 8px rgba(0,0,0,0.1); z-index:2;' : 'transform: scale(1);';

    // Логіка показу кнопки: показуємо, якщо фільтр НЕ "Поламані" і є що затверджувати
    let showApproveAll = (currentAuditFilter !== 'broken') && 
                         (res.perfect.length > 0 || res.modified.length > 0 || res.healed.length > 0);

    let html = `
        <div style="display: flex; gap: 15px; margin-bottom: ${showApproveAll ? '10px' : '15px'}; min-width: 920px;">
            <div onclick="setAuditFilter('perfect')" style="cursor: pointer; flex: 1; background: #e6f4ea; border: 2px solid #1e8e3e; padding: 12px; border-radius: 8px; text-align: center; opacity: ${getOpacity('perfect')}; ${getScale('perfect')} transition: all 0.2s ease;">
                <h3 style="margin: 0; color: #1e8e3e; user-select: none; font-size: 14px;">✅ Ідеальні: ${res.perfect.length}</h3>
            </div>
            <div onclick="setAuditFilter('modified')" style="cursor: pointer; flex: 1; background: #fff8e1; border: 2px solid #fbc02d; padding: 12px; border-radius: 8px; text-align: center; opacity: ${getOpacity('modified')}; ${getScale('modified')} transition: all 0.2s ease;">
                <h3 style="margin: 0; color: #f57f17; user-select: none; font-size: 14px;">⚠️ Зі змінами: ${res.modified.length}</h3>
            </div>
            <div onclick="setAuditFilter('healed')" style="cursor: pointer; flex: 1; background: #e8f0fe; border: 2px solid #1a73e8; padding: 12px; border-radius: 8px; text-align: center; opacity: ${getOpacity('healed')}; ${getScale('healed')} transition: all 0.2s ease;">
                <h3 style="margin: 0; color: #1a73e8; user-select: none; font-size: 14px;">💊 Зцілені: ${res.healed.length}</h3>
            </div>
            <div onclick="setAuditFilter('broken')" style="cursor: pointer; flex: 1; background: #fce8e6; border: 2px solid #d93025; padding: 12px; border-radius: 8px; text-align: center; opacity: ${getOpacity('broken')}; ${getScale('broken')} transition: all 0.2s ease;">
                <h3 style="margin: 0; color: #d93025; user-select: none; font-size: 14px;">❌ Поламані: ${res.broken.length}</h3>
            </div>
        </div>
    `;

    let showDeleteAllBroken = (currentAuditFilter === 'broken' || currentAuditFilter === 'all') && res.broken.length > 0;

    let actionButtonsHtml = `<div style="display: flex; justify-content: flex-end; gap: 10px; margin-bottom: 10px;">`;
    if (showApproveAll) {
        actionButtonsHtml += `<button class="success-btn" onclick="approveAllAuditRings()">✔ Затвердити видимі (Ідеальні, Змінені, Зцілені)</button>`;
    }
    if (showDeleteAllBroken) {
        actionButtonsHtml += `<button class="danger-btn" onclick="deleteAllBrokenAuditRings()">🗑️ Видалити всі поламані</button>`;
    }
    actionButtonsHtml += `</div>`;

    if (showApproveAll || showDeleteAllBroken) {
        html += actionButtonsHtml;
    }

    headerPanel.innerHTML = html;
    renderAuditCards(); // Викликаємо генерацію карток
}

function renderAuditCards() {
    const res = window.currentAuditResults;
    auditCardsHTML = []; 

    const renderCard = (ringName, items, type, reason = '') => {
        let headerColor = type === 'perfect' ? '#c9fed8' : type === 'modified' ? '#fff59d' : type === 'healed' ? '#c1d7ff' : '#fad2cf';
        let actions = '';

        if (type === 'broken') {
            actions = `<button class="action-btn" onclick="openHealModal('${ringName}', '${type}')">💊 Лікувати вручну</button>
                       <button class="danger-btn" onclick="deleteAuditRing('${ringName}', '${type}')">🗑️ Видалити</button>`;
        } else if (type === 'modified' || type === 'healed') {
            actions = `<button class="success-btn" onclick="approveAuditRing('${ringName}', '${type}')">✔ Затвердити</button>
                       <button class="action-btn" onclick="openHealModal('${ringName}', '${type}')">✏️ Редагувати</button>
                       <button class="danger-btn" onclick="deleteAuditRing('${ringName}', '${type}')">🗑️ Видалити</button>`;
        } else {
            actions = `<button class="success-btn" onclick="approveAuditRing('${ringName}', '${type}')">✔ Затвердити</button>
                       <button class="danger-btn" onclick="deleteAuditRing('${ringName}', '${type}')">🗑️ Видалити</button>`;
        }
        let sortedItems = [...items].sort((a, b) => {
            let tripA = a.freshTrip || a.oldTrip;
            let tripB = b.freshTrip || b.oldTrip;
            // Більше не порівнюємо logisticDay, тільки абсолютний час
            return tripA.trueStart - tripB.trueStart;
        });
        return `
        <div class="ring-card" style="margin-bottom: 20px;">
            <div class="ring-header" style="background: ${headerColor}; display: flex; justify-content: space-between; padding: 6px 12px; align-items: center;">
                <div><strong>Кільце: ${ringName}</strong> ${reason ? `<span style="color: #d93025; font-size: 11px; margin-left: 10px;">(${reason})</span>` : ''}</div>
                <div style="gap: 10px; display: flex;">${actions}</div>
            </div>
            <div class="table-container mini-table">
                <table>
                    <thead>
                        <tr>
                            <th class="col-short">Статус</th><th class="col-short">GRF</th><th class="col-med">Наряд</th>
                            <th class="col-short">Авто</th><th class="col-long">Маршрут</th>
                            <th class="col-day">Пн</th><th class="col-day">Вт</th><th class="col-day">Ср</th>
                            <th class="col-day">Чт</th><th class="col-day">Пт</th><th class="col-day">Сб</th><th class="col-day">Нд</th>
                            <th class="col-short">Виїзд</th><th class="col-short">Приїзд</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${sortedItems.map(item => {  /* ТУТ МІНЯЄМО items.map НА sortedItems.map */
                            if (item.status === 'missing') {
                                let daysHtml = item.oldTrip.logDays.map(d => `<td class="col-day">${d === '+' ? '+' : ''}</td>`).join('');
                                return `<tr style="background: #ffebee; color: #d93025;"><td>❌ Зник</td><td>${item.oldTrip.grf}</td><td>${item.oldTrip.naryad || ''}</td><td>${item.oldTrip.auto}</td><td title="${item.oldTrip.route}">${item.oldTrip.route}</td>${daysHtml}<td class="time-cell">${item.oldTrip.depStr}</td><td class="time-cell">${item.oldTrip.arrStr}</td></tr>`;
                            } 
                            // ... решта умов без змін ...
                            if (item.status === 'healed') {
                                let daysHtml = item.freshTrip.logDays.map(d => `<td class="col-day">${d === '+' ? '+' : ''}</td>`).join('');
                                return `<tr style="background: #e8f0fe; color: #1a73e8; font-weight: bold;"><td title="Замість ${item.oldTrip.grf}">💊 Заміна</td><td>${item.freshTrip.grf}</td><td>${item.freshTrip.naryad || ''}</td><td>${item.freshTrip.auto}</td><td title="${item.freshTrip.route}">${item.freshTrip.route}</td>${daysHtml}<td class="time-cell">${item.freshTrip.depStr}</td><td class="time-cell">${item.freshTrip.arrStr}</td></tr>`;
                            }
                            let bg = item.modified ? "background: #fffde7;" : "";
                            let autoBg = item.modDetails?.auto ? 'background: #fff59d;' : '';
                            let startBg = item.modDetails?.start ? 'background: #fff59d;' : '';
                            let endBg = item.modDetails?.end ? 'background: #fff59d;' : '';
                            let daysHtml = item.freshTrip.logDays.map(d => `<td class="col-day">${d === '+' ? '+' : ''}</td>`).join('');
                            return `<tr style="${bg}"><td>${item.modified ? "⚠️ Змінено" : "✅ Ок"}</td><td>${item.freshTrip.grf}</td><td>${item.freshTrip.naryad || ''}</td><td style="${autoBg}">${item.freshTrip.auto}</td><td title="${item.freshTrip.route}">${item.freshTrip.route}</td>${daysHtml}<td class="time-cell" style="${startBg}">${item.freshTrip.depStr}</td><td class="time-cell" style="${endBg}">${item.freshTrip.arrStr}</td></tr>`;
                        }).join('')}
                    </tbody>
                </table>
            </div>
        </div>`;
    };

    if ((currentAuditFilter === 'all' || currentAuditFilter === 'broken') && res.broken.length > 0) auditCardsHTML.push(...res.broken.map(r => renderCard(r.ringName, r.mappedTrips, 'broken', r.reason)));
    if ((currentAuditFilter === 'all' || currentAuditFilter === 'healed') && res.healed.length > 0) auditCardsHTML.push(...res.healed.map(r => renderCard(r.ringName, r.trips, 'healed')));
    if ((currentAuditFilter === 'all' || currentAuditFilter === 'modified') && res.modified.length > 0) auditCardsHTML.push(...res.modified.map(r => renderCard(r.ringName, r.mappedTrips, 'modified')));
    if ((currentAuditFilter === 'all' || currentAuditFilter === 'perfect') && res.perfect.length > 0) auditCardsHTML.push(...res.perfect.map(r => renderCard(r.ringName, r.mappedTrips, 'perfect')));

    auditRenderedCount = 0;
    document.getElementById('audit-content').innerHTML = '';
    
    if (auditCardsHTML.length === 0) {
        document.getElementById('audit-content').innerHTML = '<div class="empty-msg" style="margin-top: 20px;">Немає кілець у цій категорії.</div>';
    } else {
        loadMoreAudit();
        setupAuditObserver();
    }
}

// --- ФУНКЦІЇ НЕСКІНЧЕННОГО СКРОЛУ ---
function setupAuditObserver() {
    if (auditObserver) auditObserver.disconnect();
    const sentinel = document.getElementById('audit-sentinel');
    if (!sentinel) return;
    
    auditObserver = new IntersectionObserver((entries) => {
        if (entries[0].isIntersecting) loadMoreAudit();
    }, { root: document.getElementById('audit-scroll'), rootMargin: '200px' });
    
    auditObserver.observe(sentinel);
}

function loadMoreAudit() {
    if (auditRenderedCount >= auditCardsHTML.length) return;
    
    const content = document.getElementById('audit-content');
    const CARDS_PER_PAGE = 20; // Скільки карток завантажувати за раз
    const nextBatch = auditCardsHTML.slice(auditRenderedCount, auditRenderedCount + CARDS_PER_PAGE);
    
    content.insertAdjacentHTML('beforeend', nextBatch.join(''));
    auditRenderedCount += CARDS_PER_PAGE;
}


// --- 4. РУЧНА ХІРУРГІЯ (ЛІКУВАННЯ) ---
function openHealModal(ringName, type) {
    let ringData = window.currentAuditResults[type].find(r => r.ringName === ringName);
    if (!ringData) return;
    
    // Визначаємо, з якого масиву брати графіки
    let tripsArray = (type === 'healed') ? ringData.trips : ringData.mappedTrips;
    
    healingContext = {
        ringName: ringName,
        type: type, // Запам'ятовуємо категорію, щоб знати, звідки видаляти після збереження
        trips: [...tripsArray],
        selectedIndex: -1
    };
    
    document.getElementById('heal-ring-name').innerText = ringName;
    document.getElementById('heal-modal').style.display = 'flex';
    document.getElementById('heal-bottom-container').style.display = 'none';
    renderHealModalTop();
}

function closeHealModal() {
    document.getElementById('heal-modal').style.display = 'none';
    healingContext = null;
}

function renderHealModalTop() {
    healingContext.trips.sort((a, b) => {
        let tripA = a.freshTrip || a.oldTrip;
        let tripB = b.freshTrip || b.oldTrip;
        return tripA.trueStart - tripB.trueStart;
    });

    let allHealed = true;
    let html = `<table>
        <thead><tr>
            <th style="width: 55px;">Дія</th>
            <th class="col-short">GRF</th><th class="col-med">Наряд</th><th class="col-short">Авто</th>
            <th class="col-long">Маршрут</th><th class="col-med">Відпр</th><th class="col-med">Отр</th>
            <th class="col-short">Виїзд</th><th class="col-short">Приїзд</th>
        </tr></thead>
        <tbody>`;
        
    healingContext.trips.forEach((item, index) => {
        // Збираємо дві кнопки в один блок
        let actionBtns = `
            <div style="display: flex; gap: 2px; justify-content: center;">
                <button class="action-btn" style="height: 20px; width: 20px; padding: 0;" onclick="findHealCandidates(${index})" title="Знайти альтернативу">🔍</button>
                <button class="danger-btn" style="height: 20px; width: 20px; padding: 0;" onclick="removeTripFromHeal(${index})" title="Видалити графік з кільця">❌</button>
            </div>`;

        if (item.status === 'missing') {
            allHealed = false;
            html += `<tr style="background: #ffebee; color: #d93025;">
                <td>${actionBtns}</td>
                <td>❌ ${item.oldTrip.grf}</td><td title="${item.oldTrip.naryad || ''}">${item.oldTrip.naryad || ''}</td><td>${item.oldTrip.auto}</td><td title="${item.oldTrip.route}">${item.oldTrip.route}</td>
                <td>${item.oldTrip.origin}</td><td>${item.oldTrip.destination}</td><td class="time-cell">${item.oldTrip.depStr}</td><td class="time-cell">${item.oldTrip.arrStr}</td>
            </tr>`;
        } else if (item.status === 'healed') {
            html += `<tr style="background: #e8f0fe;">
                <td>${actionBtns}</td>
                <td>✅ ${item.freshTrip.grf}</td><td title="${item.freshTrip.naryad || ''}">${item.freshTrip.naryad || ''}</td><td>${item.freshTrip.auto}</td><td title="${item.freshTrip.route}">${item.freshTrip.route}</td>
                <td>${item.freshTrip.origin}</td><td>${item.freshTrip.destination}</td><td class="time-cell">${item.freshTrip.depStr}</td><td class="time-cell">${item.freshTrip.arrStr}</td>
            </tr>`;
        } else {
            let bg = item.modified ? "background: #fffde7;" : "";
            let icon = item.modified ? "⚠️" : "✅";
            html += `<tr style="${bg}">
                <td>${actionBtns}</td>
                <td>${icon} ${item.freshTrip.grf}</td><td title="${item.freshTrip.naryad || ''}">${item.freshTrip.naryad || ''}</td><td>${item.freshTrip.auto}</td><td title="${item.freshTrip.route}">${item.freshTrip.route}</td>
                <td>${item.freshTrip.origin}</td><td>${item.freshTrip.destination}</td><td class="time-cell">${item.freshTrip.depStr}</td><td class="time-cell">${item.freshTrip.arrStr}</td>
            </tr>`;
        }
    });
    
    html += `</tbody></table>`;
    document.getElementById('heal-top-panel').innerHTML = html;
    
    document.getElementById('btn-save-heal').style.display = allHealed ? 'inline-flex' : 'none';
}

function removeTripFromHeal(index) {
    if (confirm("Видалити цей графік з кільця?")) {
        // Видаляємо елемент з масиву
        healingContext.trips.splice(index, 1);
        
        // Якщо користувач видалив усі графіки з кільця
        if (healingContext.trips.length === 0) {
            alert("Кільце порожнє. Воно буде автоматично видалене з перевірки.");
            window.currentAuditResults[healingContext.type] = window.currentAuditResults[healingContext.type].filter(r => r.ringName !== healingContext.ringName);
            closeHealModal();
            renderAuditView();
            return;
        }
        
        // Перемальовуємо вікно
        renderHealModalTop();
        
        // Якщо ми випадково відкрили панель пошуку для графіка, що був нижче - ховаємо її
        document.getElementById('heal-bottom-container').style.display = 'none';
    }
}

function findHealCandidates(index) {
    healingContext.selectedIndex = index;
    const mode = document.getElementById('mode_select')?.value || 'node';
    const reparkMins = parseInt(document.getElementById('repark_time')?.value) || 30;
    
    let oldMissing = healingContext.trips[index].oldTrip;
    
    // Знаходимо найближчі живі графіки ЗВЕРХУ і ЗНИЗУ від дірки
    let prev = null, next = null;
    for(let i = index - 1; i >= 0; i--) { if(healingContext.trips[i].freshTrip) { prev = healingContext.trips[i].freshTrip; break; } }
    for(let i = index + 1; i < healingContext.trips.length; i++) { if(healingContext.trips[i].freshTrip) { next = healingContext.trips[i].freshTrip; break; } }
    
    let candidates = window.allTrips.filter(c => {
        if(c.ringId !== null) return false;
        if(c.auto !== oldMissing.auto) return false;
        if(c.getPointName('origin', mode) !== oldMissing.getPointName('origin', mode)) return false;
        if(c.getPointName('dest', mode) !== oldMissing.getPointName('dest', mode)) return false;
        
        // Жорстка перевірка: чи влізе цей кандидат фізично між сусідами
        if(prev) {
            let prevEnd = prev.trueEnd + (prev.trueEnd < prev.trueStart ? 10080 : 0);
            if(c.trueStart < prevEnd + reparkMins) return false;
        }
        if(next) {
            let cEnd = c.trueEnd + (c.trueEnd < c.trueStart ? 10080 : 0);
            if(next.trueStart < cEnd + reparkMins) return false;
        }
        return true;
    });
    
    // Сортуємо кандидатів так, щоб найбільш схожий за часом був першим
    candidates.sort((a, b) => Math.abs(a.trueStart - oldMissing.trueStart) - Math.abs(b.trueStart - oldMissing.trueStart));

    let html = `<table><thead><tr><th class="col-btn">Дія</th><th class="col-short">GRF</th><th class="col-long">Маршрут</th><th class="col-short">Виїзд</th><th class="col-short">Приїзд</th></tr></thead><tbody>`;
    if(candidates.length === 0) {
        html += `<tr><td colspan="5" style="text-align:center;">Немає вільних графіків, які влізуть у це "вікно" 🤷‍♂️</td></tr>`;
    } else {
        html += candidates.map(c => `<tr>
            <td><button class="success-btn" style="height: 20px; padding: 0 5px;" onclick="applyHealCandidate('${c.id}')">Підставити</button></td>
            <td>${c.grf}</td><td>${c.route}</td><td class="time-cell">${c.depStr}</td><td class="time-cell">${c.arrStr}</td>
        </tr>`).join('');
    }
    html += `</tbody></table>`;
    
    document.getElementById('heal-bottom-panel').innerHTML = html;
    document.getElementById('heal-bottom-container').style.display = 'block';
}

function applyHealCandidate(tripId) {
    let candidate = window.allTrips.find(t => t.id === tripId);
    if(candidate && healingContext.selectedIndex !== -1) {
        let idx = healingContext.selectedIndex;
        healingContext.trips[idx] = { status: 'healed', freshTrip: candidate, oldTrip: healingContext.trips[idx].oldTrip };
        healingContext.selectedIndex = -1;
        document.getElementById('heal-bottom-container').style.display = 'none';
        renderHealModalTop();
    }
}

// --- 5. ЗАТВЕРДЖЕННЯ / ВИДАЛЕННЯ ---
function saveHealedRing() {
    let newId = `approved_${Date.now()}_healed`;
    
    // Проставляємо нові ID
    healingContext.trips.forEach(item => { 
        if (item.freshTrip) item.freshTrip.ringId = newId; 
    });
    
    window.ringNamesMap[newId] = healingContext.ringName;
    
    // Видаляємо кільце з тієї категорії, в якій воно лежало до редагування
    window.currentAuditResults[healingContext.type] = window.currentAuditResults[healingContext.type].filter(r => r.ringName !== healingContext.ringName);
    
    closeHealModal();
    renderAuditView(); 
    renderArchive(); 
    render(window.allTrips);
}

function approveAuditRing(ringName, type) {
    let ringData = window.currentAuditResults[type].find(r => r.ringName === ringName);
    let newId = `approved_${Date.now()}_audit`;
    let tripsArray = (type === 'healed') ? ringData.trips : ringData.mappedTrips;
    
    tripsArray.forEach(item => { item.freshTrip.ringId = newId; });
    window.ringNamesMap[newId] = ringName;
    window.currentAuditResults[type] = window.currentAuditResults[type].filter(r => r.ringName !== ringName);
    
    renderAuditView(); renderArchive(); render(window.allTrips);
}

function deleteAuditRing(ringName, type) {
    if(confirm("Видалити кільце з перевірки? Вцілілі графіки залишаться вільними у реєстрі.")) {
        window.currentAuditResults[type] = window.currentAuditResults[type].filter(r => r.ringName !== ringName);
        renderAuditView();
    }
}

// Додаємо можливість тягати вікно Хірургії мишкою
const healHeader = document.getElementById("heal-header");
const healModal = document.getElementById("heal-modal");
if (healHeader && healModal) {
    let pos1 = 0, pos2 = 0, pos3 = 0, pos4 = 0;
    healHeader.onmousedown = function(e) {
        e.preventDefault();
        pos3 = e.clientX; pos4 = e.clientY;
        document.onmouseup = () => { document.onmouseup = null; document.onmousemove = null; };
        document.onmousemove = function(e) {
            e.preventDefault();
            pos1 = pos3 - e.clientX; pos2 = pos4 - e.clientY;
            pos3 = e.clientX; pos4 = e.clientY;
            healModal.style.top = (healModal.offsetTop - pos2) + "px";
            healModal.style.left = (healModal.offsetLeft - pos1) + "px";
        };
    };
}

function approveAllAuditRings() {
    if (!window.currentAuditResults) return;

    // Визначаємо, які категорії затверджувати залежно від обраного фільтра
    let typesToApprove = [];
    if (currentAuditFilter === 'all') {
        typesToApprove = ['perfect', 'modified', 'healed']; // Всі, крім поламаних
    } else if (currentAuditFilter !== 'broken') {
        typesToApprove = [currentAuditFilter]; // Тільки обрана категорія
    }

    if (typesToApprove.length === 0) return;

    let approvedCount = 0;

    typesToApprove.forEach(type => {
        if (window.currentAuditResults[type] && window.currentAuditResults[type].length > 0) {
            
            window.currentAuditResults[type].forEach(ringData => {
                let newId = `approved_${Date.now()}_${Math.random().toString(36).substr(2, 5)}`;
                let tripsArray = (type === 'healed') ? ringData.trips : ringData.mappedTrips;
                
                tripsArray.forEach(item => { 
                    if (item.freshTrip) item.freshTrip.ringId = newId; 
                });
                
                window.ringNamesMap[newId] = ringData.ringName;
                approvedCount++;
            });
            
            // Очищаємо оброблену категорію
            window.currentAuditResults[type] = [];
        }
    });

    if (approvedCount > 0) {
        alert(`Успішно затверджено кілець: ${approvedCount}\nВони додані у вкладку "Затверджені кільця".`);
        renderAuditView(); 
        renderArchive();   
        render(window.allTrips); 
    } else {
        alert("Немає кілець для затвердження у поточній видимій категорії.");
    }
}

function deleteAllBrokenAuditRings() {
    if (confirm("Ви впевнені, що хочете видалити всі поламані кільця з перевірки?\nВцілілі графіки з них залишаться доступними у загальному реєстрі.")) {
        // Просто обнуляємо масив поламаних
        window.currentAuditResults.broken = [];
        renderAuditView(); // Перемальовуємо інтерфейс
    }
}