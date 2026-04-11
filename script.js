// 1. Константы и настройки
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbwtRtzMPV6OykBEJ5OoPNYDhW0Tp1FXaAsMzfMthqL7nj8J3Jke2jdj7dzRL46kTKcO/exec';
let nodeDictionary = new Map();
window.allTrips = []; // Основное хранилище данных
window.ringNamesMap = {};

const MATRIX_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbwtRtzMPV6OykBEJ5OoPNYDhW0Tp1FXaAsMzfMthqL7nj8J3Jke2jdj7dzRL46kTKcO/exec?action=matrix';

// ==========================================
// СКРОЛЛ ДЛЯ ЧИСЛОВИХ ПОЛІВ (Зміна значень коліщатком)
// ==========================================
['repark_time', 'min_trips'].forEach(id => {
    const input = document.getElementById(id);
    if (input) {
        input.addEventListener('wheel', function(e) {
            // Забороняємо прокрутку всієї сторінки, поки миша над полем
            e.preventDefault(); 
            
            let val = parseInt(this.value) || 0;
            
            // e.deltaY < 0 означає скролл коліщатком вгору
            if (e.deltaY < 0) {
                val += 1;
            } else {
                val -= 1;
            }
            
            // Захист від від'ємних значень (щоб не скрутили перепарковку в мінус)
            if (val < 0) val = 0;
            
            this.value = val;
            
            // Штучно викликаємо подію 'change'.
            // Оскільки в тебе вже висить слухач на зміну repark_time, 
            // таблиця буде миттєво перемальовуватись прямо під час скролу!
            this.dispatchEvent(new Event('change'));
        });
    }
});


async function assignRingNames() {
    const btn = document.getElementById('btn_name_rings');
    btn.innerText = "⏳ Отримуємо матрицю...";
    btn.disabled = true;

    try {
        // 1. Завантажуємо матрицю
        const response = await fetch(MATRIX_WEB_APP_URL);
        const matrixArray = await response.json();
        
        // 2. Перетворюємо масив у зручний словник.
        // ЩОБ ШУКАТИ ТІЛЬКИ ПО МІСТУ: проганяємо отримані дані через довідник, 
        // щоб відсікти будь-які локації та залишити чисті міста.
        const matrixLookup = {};
        matrixArray.forEach(item => {
            const originRaw = String(item.origin).trim();
            const destRaw = String(item.dest).trim();
            
            const originData = nodeDictionary.get(originRaw) || { city: originRaw };
            const destData = nodeDictionary.get(destRaw) || { city: destRaw };

            const originCity = String(originData.city).trim().toLowerCase();
            const destCity = String(destData.city).trim().toLowerCase();

            const key = `${originCity}_${destCity}`;
            matrixLookup[key] = item.code;
        });

        // 3. Збираємо всі затверджені кільця
        const archiveMap = {};
        window.allTrips.forEach(t => {
            if (t.ringId && t.ringId.startsWith('approved_')) {
                if (!archiveMap[t.ringId]) archiveMap[t.ringId] = [];
                archiveMap[t.ringId].push(t);
            }
        });

        const rings = Object.values(archiveMap);
        if (rings.length === 0) {
            alert("Немає затверджених кілець для іменування.");
            return;
        }

        // 4. Лічильник для однакових кодів (щоб робити 01, 02, 03)
        const codeCounters = {};

        // 5. Проходимося по кожному кільцю і даємо йому ім'я
        rings.forEach(ring => {
            const rId = ring[0].ringId;
            
            // Сортуємо графіки хронологічно
            ring.sort((a, b) => {
                if (a.logisticDay !== b.logisticDay) return a.logisticDay - b.logisticDay;
                return a.trueStart - b.trueStart;
            });

            // ШУКАЄМО ТІЛЬКИ ПО МІСТУ: Знаходимо перший МІЖМІСЬКИЙ рейс. 
            // Якщо першим стоїть локальний переїзд (Київ-Київ), ми його пропускаємо
            let mainTrip = ring.find(t => {
                let oCity = String(t.getPointName('origin', 'city')).trim().toLowerCase();
                let dCity = String(t.getPointName('dest', 'city')).trim().toLowerCase();
                return oCity !== dCity; 
            });
            
            // Якщо всі рейси в кільці локальні (в межах одного міста) — беремо перший
            if (!mainTrip) mainTrip = ring[0];
            
            // Беремо базове місто (суворо функцією getPointName)
            const originCity = String(mainTrip.getPointName('origin', 'city')).trim().toLowerCase();
            const destCity = String(mainTrip.getPointName('dest', 'city')).trim().toLowerCase();
            
            const lookupKey = `${originCity}_${destCity}`;

            // Шукаємо код
            let baseCode = matrixLookup[lookupKey] || "XXXX";

            // Рахуємо, яке це по рахунку кільце з таким кодом
            if (!codeCounters[baseCode]) codeCounters[baseCode] = 0;
            codeCounters[baseCode]++;

            // Форматуємо порядковий номер (додаємо нуль спереду, якщо число < 10)
            const sequenceNum = String(codeCounters[baseCode]).padStart(2, '0');

            // Записуємо готове ім'я
            window.ringNamesMap[rId] = `${baseCode}_${sequenceNum}`;
        });

        // Перемальовуємо архів, щоб показати нові імена
        renderArchive();
        
    } catch (e) {
        console.error("Помилка іменування:", e);
        alert("Помилка завантаження матриці. Перевірте URL та консоль.");
    } finally {
        btn.innerText = "🏷️ Дати імена кільцям";
        btn.disabled = false;
    }
}


// 2. Элементы интерфейса (всегда объявляем в самом начале!)
const modeSelect = document.getElementById('mode_select');
const reparkInput = document.getElementById('repark_time');
//const filterFrom = document.getElementById('filter_from');
//const filterTo = document.getElementById('filter_to');
//const filterAuto = document.getElementById('filter_auto'); // НОВЕ
//const filterType = document.getElementById('filter_type'); // НОВЕ
const activeFilters = { origin: new Set(), dest: new Set(), auto: new Set(), type: new Set() };
let currentFilterColumn = null; // Зберігає інформацію, який фільтр зараз відкритий

const fileInput = document.getElementById('file_input');
const dropZone = document.getElementById('drop_zone');
const tableBody = document.getElementById('table_body');
const status = document.getElementById('status');
const minTripsInput = document.getElementById('min_trips');
let clusterizeInstance = null;
//let clusterizeDraft = null;   // Добавить это
//let clusterizeArchive = null; // Добавить это

// Змінні для нескінченного скроллу
let draftCardsHTML = [];
let archiveCardsHTML = [];
let draftRenderedCount = 0;
let archiveRenderedCount = 0;
const CARDS_PER_PAGE = 20; // Скільки карток підвантажувати за раз
let draftObserver = null;
let archiveObserver = null;


let isAlgoRunning = false; // Флаг роботи алгоритму

function stopAlgo() {
    isAlgoRunning = false;
    //status.innerText = `Доступно графіків: ${filtered.length} (всього завантажено: ${trips.length})`;
    //document.getElementById('status').innerText = "Пошук перервано користувачем...";
}

// 3. Слушатели событий
// Если меняется любой фильтр или режим — перерисовываем
[reparkInput].forEach(el => {
    el.addEventListener('change', () => {
        render(window.allTrips);
    });
});


// Загрузка справочника
async function loadDictionary() {
    status.innerText = "Завантаження довідника вузлів...";
    try {
        const response = await fetch(WEB_APP_URL);
        const data = await response.json();
        nodeDictionary.clear();
        data.forEach(item => {
            nodeDictionary.set(String(item.node).trim(), {
                city: item.city,
                city2: item.city2
            });
        });
        status.innerText = `Довідник завантажено (${nodeDictionary.size} вузлів)`;
    } catch (e) {
        console.error("Помилка API:", e);
        status.innerText = "Помилка довідника (див. консоль)";
    }
}

// Класс данных рейса
class Trip {
    constructor(r) {
        this.rawRow = [...r];
        this.id = 'trip_' + Math.random().toString(36).substr(2, 9); // НОВОЕ: Уникальный ID
        this.grf = r[2]; this.digit = r[3]; this.code = r[4]; this.group = r[5];
        this.naryad = r[6]; this.type = r[7]; this.auto = r[8]; this.load = r[9];
        this.route = r[10]; this.origin = String(r[11] || "").trim();
        this.destination = String(r[15] || "").trim();

        this.originData = nodeDictionary.get(this.origin) || { city: this.origin, city2: this.origin };
        this.destData = nodeDictionary.get(this.destination) || { city: this.destination, city2: this.destination };

        const astroDays = [r[16], r[17], r[18], r[19], r[20], r[21], r[22]];
        this.dayIndex = astroDays.findIndex(d => d === '+');
        if (this.dayIndex === -1) this.dayIndex = 0;

        this.logDays = [r[23], r[24], r[25], r[26], r[27], r[28], r[29]];
        this.logisticDay = this.logDays.findIndex(d=>d==="+");
        if (this.logisticDay === -1) this.logisticDay = 0;

        this.drivers = r[30]; this.deadline = r[31];
        
        this.podachaStr = formatTime(r[32]);
        this.depStr = formatTime(r[33]);
        this.arrStr = formatTime(r[40]);
        this.freeStr = formatTime(r[41]);

        this.calculateTimeline();
        this.calculateTrueTimes();
        this.comment = r[54];
        this.ringId = null;
        this.originalRingId = null; // НОВИЙ РЯДОК: Пам'ять для автостеплера
    }

    calculateTrueTimes() {
        const isBDF = String(this.auto || "").toUpperCase().includes("БДФ");
        this.trueStart = isBDF ? this.depInt : this.podachaInt;
        this.trueEnd = isBDF ? this.arrInt : this.freeInt;

        // Если логистический день - Воскресенье (индекс 6), 
        // а астрономическое время выезда упало на Понедельник (trueStart < 1440 минут)
        if (this.logisticDay === 6 && this.trueStart < 1440) {
            this.trueStart += 10080; // Переносим на условный "8-й день"
        }

        // Обязательная страховка: если финиш оказался раньше старта, 
        // перекидываем его вслед за стартом на следующую неделю
        if (this.trueEnd < this.trueStart) {
            this.trueEnd += 10080;
        }
    }

    calculateTimeline() {
        const minInDay = 1440; const minInWeek = 10080;
        const dayStart = this.dayIndex * minInDay;
        const toMin = (str) => {
            if (!str) return 0;
            const [h, m] = str.split(':').map(Number);
            return (h * 60) + m;
        };
        const dM = toMin(this.depStr);
        this.depInt = dayStart + dM;
        let pM = toMin(this.podachaStr);
        let pInt = dayStart + pM;
        if (pM > dM) pInt -= minInDay;
        this.podachaInt = pInt < 0 ? pInt + minInWeek : pInt;
        let aM = toMin(this.arrStr);
        let aInt = dayStart + aM;
        if (aM < dM) aInt += minInDay;
        this.arrInt = aInt >= minInWeek ? aInt - minInWeek : aInt;
        let fM = toMin(this.freeStr);
        let fInt = dayStart + fM;
        if (fM < aM) fInt = aInt + (fM + minInDay - aM);
        else if (aInt > (dayStart + minInDay)) fInt += minInDay;
        this.freeInt = fInt >= minInWeek ? fInt - minInWeek : fInt;
    }

    getPointName(point, mode) {
        const data = point === 'origin' ? this.originData : this.destData;
        if (mode === 'city') return data.city;
        if (mode === 'city2') return data.city2;
        return point === 'origin' ? this.origin : this.destination;
    }
}

// Вспомогательные функции
function formatTime(val) {
    if (val == null || val === "") return "";
    if (typeof val === 'number') {
        const totalMins = Math.round(val * 1440);
        return `${Math.floor(totalMins/60).toString().padStart(2,'0')}:${(totalMins%60).toString().padStart(2,'0')}`;
    }
    return val.toString().trim().substring(0, 5);
}


function clearFilters() {
    // Очищаємо всі масиви вибраних значень
    activeFilters.origin.clear();
    activeFilters.dest.clear();
    activeFilters.auto.clear();
    activeFilters.type.clear();
    
    // Скидаємо колір іконок і перемальовуємо таблицю
    updateFilterIcons();
    render(window.allTrips);
}

// Загрузка файла
function handleFile(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = e => {
        // НОВЕ: Зберігаємо сирі дані файлу, щоб брати чистий оригінал при кожному експорті
        window.uploadedFileData = e.target.result; 
        
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
        
        window.originalHeaders = rows.slice(0, 4);

        // НОВЕ: Записуємо графіки, запам'ятовуючи їх оригінальний рядок в Excel
        window.allTrips = [];
        for (let i = 4; i < rows.length; i++) {
            if (rows[i][2]) { // Перевірка наявності GRF
                let t = new Trip(rows[i]);
                t.originalRowIndex = i; // Зберігаємо оригінальний індекс рядка (0-based)
                window.allTrips.push(t);
            }
        }

        window.allTrips.sort((a, b) => {
            if (a.trueStart !== b.trueStart) {
                return a.trueStart - b.trueStart;
            }
            return a.logisticDay - b.logisticDay;
        });
        
        render(window.allTrips);
        checkMissingNodes();
        updateSidebarAutoButtons();
    };
    reader.readAsArrayBuffer(file);
}

// Рендер таблицы с виртуальным скроллом
function render(trips) {
    const mode = modeSelect.value;
    const reparkMinutes = parseInt(reparkInput.value) || 0;

    const fromVals = Array.from(activeFilters.origin);
    const toVals = Array.from(activeFilters.dest);
    const autoVals = Array.from(activeFilters.auto);
    const typeVals = Array.from(activeFilters.type);

    // Фільтруємо: ТІЛЬКИ вільні графіки
    const filtered = trips.filter(t => {
        if (t.ringId !== null) return false; 
        
        const originName = t.origin;
        const destName = t.destination;

        // Логіка для напрямків (А->Б або Б->А) з масивами
        let matchDirection = true;
        if (fromVals.length > 0 && toVals.length > 0) {
            const straight = fromVals.includes(originName) && toVals.includes(destName);
            const reverse = toVals.includes(originName) && fromVals.includes(destName);
            matchDirection = straight || reverse;
        } else if (fromVals.length > 0) {
            matchDirection = fromVals.includes(originName);
        } else if (toVals.length > 0) {
            matchDirection = toVals.includes(destName);
        }

        if (!matchDirection) return false;

        // Логіка для нових фільтрів
        if (autoVals.length > 0 && !autoVals.includes(t.auto)) return false;
        if (typeVals.length > 0 && !typeVals.includes(t.type)) return false;

        return true; 
    });

    const dayNames = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Нд'];
    
    // Формуємо масив HTML-рядків (кожен <tr> — окремий елемент масиву)
    const rowsData = filtered.map(t => {
        const finalTrueEnd = t.trueEnd + reparkMinutes;
        
        return `<tr data-id="${t.id}" title="Логістичний день: ${dayNames[t.logisticDay]}">
            <td>${t.grf || ''}</td><td>${t.digit || ''}</td><td>${t.code || ''}</td>
            <td>${t.group || ''}</td><td>${t.naryad || ''}</td><td>${t.type || ''}</td>
            <td>${t.auto || ''}</td><td>${t.load || ''}</td><td title="${t.route || ''}">${t.route || ''}</td>
            <td class="highlight-node">${t.origin}</td>
            <td class="highlight-node">${t.destination}</td>
            ${t.logDays.map(d => `<td class="col-day">${d === '+' ? '+' : ''}</td>`).join('')}
            <td>${t.drivers || ''}</td><td>${t.deadline || ''}</td>
            <td class="time-cell" title="Подача: ${t.podachaInt} хв.&#10;TrueStart: ${t.trueStart} хв.">${t.podachaStr}</td>
            <td class="time-cell" style="color: #1a73e8;" title="Виїзд: ${t.depInt} хв.&#10;TrueStart: ${t.trueStart} хв.">${t.depStr}</td>
            <td class="time-cell" title="Приїзд: ${t.arrInt} хв.&#10;TrueEnd: ${t.trueEnd} хв.">${t.arrStr}</td>
            <td class="time-cell" title="Вільний: ${t.freeInt} хв.&#10;TrueEnd: ${t.trueEnd} хв.">${t.freeStr}</td>
            <td title="${t.comment || ''}">${t.comment || ''}</td>
        </tr>`;
    });

    // Якщо скролл ще не створений — ініціалізуємо його
    if (!clusterizeInstance) {
        clusterizeInstance = new Clusterize({
            rows: rowsData,
            scrollId: 'scrollArea',
            contentId: 'table_body',
            no_data_text: 'Немає даних для відображення'
        });
    } else {
        // Якщо вже створений — просто оновлюємо дані (працює миттєво)
        clusterizeInstance.update(rowsData);
    }
    
    status.innerText = `Доступно графіків: ${filtered.length} (всього завантажено: ${trips.length})`;
    updateDraftImbalanceStats();
}

// Инициализация Drag and Drop
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, e => { e.preventDefault(); e.stopPropagation(); }, false);
});

let dragCounter = 0;

dropZone.addEventListener('dragenter', e => {
    dragCounter++;
    dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragover', e => {
    if (!dropZone.classList.contains('dragover')) {
        dropZone.classList.add('dragover');
    }
});

dropZone.addEventListener('dragleave', e => {
    dragCounter--;
    if (dragCounter === 0) {
        dropZone.classList.remove('dragover');
    }
});




dropZone.onclick = () => fileInput.click();
fileInput.onchange = e => handleFile(e.target.files[0]);
dropZone.addEventListener('drop', e => {
    // Гарантированно убираем класс и сбрасываем счетчик
    dragCounter = 0;
    dropZone.classList.remove('dragover');
    
    handleFile(e.dataTransfer.files[0]);
});


// Переключение вкладок
function switchTab(tabId) {
    document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    
    document.getElementById(tabId).classList.add('active');
    
    // Знаходимо кнопку, на яку натиснули, за текстом або атрибутом
    const targetBtn = Array.from(document.querySelectorAll('.tab-btn')).find(b => b.getAttribute('onclick').includes(tabId));
    if (targetBtn) targetBtn.classList.add('active');
    
    document.getElementById('draft-actions').style.display = (tabId === 'draft-tab') ? 'flex' : 'none';
    
    // Перемальовуємо дані при переході
    if (tabId === 'register-tab') render(window.allTrips);
    if (tabId === 'draft-tab') renderDraft();
    if (tabId === 'archive-tab') renderArchive();
    if (tabId === 'stapler-draft-tab') renderStaplerDraft(); // НОВИЙ РЯДОК
}

// --- НОВА ЛОГІКА ІНТЕРФЕЙСУ ПОШУКУ ---
function toggleAlgoSettings() {
    const algo = document.getElementById('master_algo_select').value;
    const patternBlock = document.getElementById('dynamic_pattern_settings');
    const ppBlock = document.getElementById('dynamic_pp_settings');

    patternBlock.style.display = (algo === 'pattern') ? 'flex' : 'none';
    ppBlock.style.display = (algo === 'pp') ? 'flex' : 'none';
}

async function runMasterAlgorithm() {
    const algo = document.getElementById('master_algo_select').value;
    
    isAlgoRunning = true;
    const btnRun = document.getElementById('btn_run_algo');
    const btnStop = document.getElementById('btn_stop_algo');
    
    if (btnRun && btnStop) {
        btnRun.style.display = 'none';
        btnStop.style.display = 'inline-flex';
    }

    try {
        // Мікро-пауза, щоб браузер встиг перемалювати кнопки до зависання розрахунками
        await new Promise(resolve => setTimeout(resolve, 10));

        if (['fifo', 'filo', 'tree', 'tree_opt'].includes(algo)) {
            await runShuttleAlgo(algo); // Передаємо стратегію явно
        } else if (algo === 'pattern') {
            await runPatternAlgo();
        } else if (algo === 'pp') {
            await balanceDraftTrips();
        }
    } finally {
        isAlgoRunning = false;
        if (btnRun && btnStop) {
            btnRun.style.display = 'inline-flex';
            btnStop.style.display = 'none';
        }
        renderDraft();
        render(window.allTrips);
    }
}

async function runShuttleAlgo(strategy) {
    const mode = modeSelect.value;
    const repark = parseInt(reparkInput.value) || 0;
    const minTrips = parseInt(document.getElementById('min_trips')?.value || 4);
    
    // Беремо галочку з нових глобальних налаштувань
    const requireReturn = document.getElementById('global_return')?.checked || false;

    window.allTrips.forEach(t => { 
        if(t.ringId && t.ringId.startsWith('draft_')) t.ringId = null; 
    });

    const fromVals = Array.from(activeFilters.origin);
    const toVals = Array.from(activeFilters.dest);
    const autoVals = Array.from(activeFilters.auto);
    const typeVals = Array.from(activeFilters.type);

    const workingTrips = window.allTrips.filter(t => {
        const originName = t.getPointName('origin', mode);
        const destName = t.getPointName('dest', mode);
        
        let matchDirection = true;
        if (fromVals.length > 0 && toVals.length > 0) {
            const straight = fromVals.includes(originName) && toVals.includes(destName);
            const reverse = toVals.includes(originName) && fromVals.includes(destName);
            matchDirection = straight || reverse;
        } else if (fromVals.length > 0) {
            matchDirection = fromVals.includes(originName);
        } else if (toVals.length > 0) {
            matchDirection = toVals.includes(destName);
        }

        if (!matchDirection) return false;
        if (autoVals.length > 0 && !autoVals.includes(t.auto)) return false;
        if (typeVals.length > 0 && !typeVals.includes(t.type)) return false;

        return true;
    });

    if (strategy === 'fifo') {
        await algoFIFO(workingTrips, mode, repark, minTrips, requireReturn);
    } else if (strategy === 'filo') {
        await algoFILO(workingTrips, mode, repark, minTrips, requireReturn);
    } else if (strategy === 'tree') {
        await algoTree(workingTrips, mode, repark, minTrips, requireReturn);
    } else if (strategy === 'tree_opt') { 
        await algoTreeOptimized(workingTrips, mode, repark, minTrips, requireReturn);
    }
}

// --- АЛГОРИТМ 1: FIFO (Швидкий човник А-Б-А) ---
async function algoFIFO(workingTrips, mode, repark, minTrips, requireReturn) {
    let ringCounter = 0;
    for (let i = 0; i < workingTrips.length; i++) {
        if (!isAlgoRunning) break; 
        // Даємо браузеру "дихнути" кожні 50 ітерацій, щоб кнопка Стоп натискалася
        if (i % 50 === 0) await new Promise(resolve => setTimeout(resolve, 0));

        let anchor = workingTrips[i];
        if (anchor.ringId !== null || anchor.logisticDay > 5) continue;

        const pointA = anchor.getPointName('origin', mode);
        const pointB = anchor.getPointName('dest', mode);
        let currentChain = [anchor];
        anchor.ringId = 'temp'; 
        let currentDest = pointB;
        let lastTrip = anchor;
        let searching = true;

        while (searching && isAlgoRunning) {
            const targetOrigin = currentDest;
            const targetDest = (currentDest === pointA) ? pointB : pointA;
            let effectiveLastEnd = lastTrip.trueEnd + (lastTrip.trueEnd < lastTrip.trueStart ? 10080 : 0);

            const isLastEmpty = String(lastTrip.type || '').trim().toLowerCase() === "порожній";

            let nextTrip = workingTrips.find(candidate => {
                const isCandidateEmpty = String(candidate.type || '').trim().toLowerCase() === "порожній";
                if (isLastEmpty && isCandidateEmpty) return false; 
                return candidate.ringId === null && 
                       candidate.getPointName('origin', mode) === targetOrigin &&
                       candidate.getPointName('dest', mode) === targetDest &&
                       candidate.auto === anchor.auto &&
                       candidate.trueStart > lastTrip.trueStart && 
                       candidate.trueStart >= (effectiveLastEnd + repark);
            });

            if (nextTrip) {
                nextTrip.ringId = 'temp';
                currentChain.push(nextTrip);
                currentDest = nextTrip.getPointName('dest', mode);
                lastTrip = nextTrip;
            } else {
                searching = false; 
            }
        }

        if (requireReturn) {
            while (currentChain.length > 0) {
                let last = currentChain[currentChain.length - 1];
                if (last.getPointName('dest', mode) === pointA) {
                    break;
                }
                let removed = currentChain.pop();
                removed.ringId = null; 
            }
        }

        if (currentChain.length >= minTrips) {
            const draftId = `draft_${Date.now()}_${ringCounter++}`;
            currentChain.forEach(t => t.ringId = draftId);
        } else {
            currentChain.forEach(t => t.ringId = null); 
        }
    }
}

// --- АЛГОРИТМ 2: FILO (Відкладений човник А-Б-А) ---
async function algoFILO(workingTrips, mode, repark, minTrips, requireReturn) {
    let ringCounter = 0;
    for (let i = 0; i < workingTrips.length; i++) {
        if (!isAlgoRunning) break;
        if (i % 50 === 0) await new Promise(resolve => setTimeout(resolve, 0));

        let anchor = workingTrips[i];
        if (anchor.ringId !== null || anchor.logisticDay > 4) continue;

        const pointA = anchor.getPointName('origin', mode);
        const pointB = anchor.getPointName('dest', mode);
        let currentChain = [anchor];
        anchor.ringId = 'temp';
        let currentDest = pointB;
        let lastTrip = anchor;
        let searching = true;

        while (searching && isAlgoRunning) {
            const targetOrigin = currentDest;
            const targetDest = (currentDest === pointA) ? pointB : pointA;
            let effectiveLastEnd = lastTrip.trueEnd + (lastTrip.trueEnd < lastTrip.trueStart ? 10080 : 0);
            const targetLogDay = (lastTrip.logisticDay + 1) % 7;
            const isLastEmpty = String(lastTrip.type || '').trim().toLowerCase() === "порожній";
            let nextTrip = workingTrips.findLast(candidate => {
                const isCandidateEmpty = String(candidate.type || '').trim().toLowerCase() === "порожній";
                if (isLastEmpty && isCandidateEmpty) return false; 
                return candidate.ringId === null && 
                       candidate.getPointName('origin', mode) === targetOrigin &&
                       candidate.getPointName('dest', mode) === targetDest &&
                       candidate.auto === anchor.auto &&
                       candidate.logisticDay === targetLogDay &&
                       candidate.trueStart > lastTrip.trueStart && 
                       candidate.trueStart >= (effectiveLastEnd + repark);
            });

            if (nextTrip) {
                nextTrip.ringId = 'temp';
                currentChain.push(nextTrip);
                currentDest = nextTrip.getPointName('dest', mode);
                lastTrip = nextTrip;
            } else {
                searching = false; 
            }
        }

        if (requireReturn) {
            while (currentChain.length > 0) {
                let last = currentChain[currentChain.length - 1];
                if (last.getPointName('dest', mode) === pointA) {
                    break;
                }
                let removed = currentChain.pop();
                removed.ringId = null; 
            }
        }

        if (currentChain.length >= minTrips) {
            const draftId = `draft_${Date.now()}_${ringCounter++}`;
            currentChain.forEach(t => t.ringId = draftId);
        } else {
            currentChain.forEach(t => t.ringId = null); 
        }
    }
}

// --- АЛГОРИТМ 3: ДЕРЕВО (Транзит, пошук найдовшого ланцюга) ---
async function algoTree(workingTrips, mode, repark, minTrips, requireReturn) {
    let ringCounter = 0;

    for (let i = 0; i < workingTrips.length; i++) {
        // Якщо натиснули "Зупинити" — виходимо з головного циклу
        if (!isAlgoRunning) break; 
        if (i % 10 === 0) await new Promise(resolve => setTimeout(resolve, 0));

        let anchor = workingTrips[i];
        if (anchor.ringId !== null || anchor.logisticDay > 5) continue;

        let bestChain = [];
        let exploreCounter = 0; // Додаємо лічильник ітерацій рекурсії

        // РОБИМО ФУНКЦІЮ explore АСИНХРОННОЮ
        async function explore(currentTrip, currentChain) {
            if (!isAlgoRunning) return; 

            // Даємо браузеру час "почути" клік кожні 200 гілок
            exploreCounter++;
            if (exploreCounter % 200 === 0) {
                await new Promise(resolve => setTimeout(resolve, 0));
            }
            
            if (!isAlgoRunning) return; // Перевіряємо ще раз після мікро-паузи

            if (currentChain.length > bestChain.length) {
                bestChain = [...currentChain];
            }

            let effectiveLastEnd = currentTrip.trueEnd + (currentTrip.trueEnd < currentTrip.trueStart ? 10080 : 0);

            const currentDest = currentTrip.getPointName('dest', mode);
            const isCurrentEmpty = String(currentTrip.type || '').trim().toLowerCase() === "порожній";

            let candidates = workingTrips.filter(candidate => {
                const isCandidateEmpty = String(candidate.type || '').trim().toLowerCase() === "порожній";
                if (isCurrentEmpty && isCandidateEmpty) return false; // Забороняємо 2 порожніх підряд
                return candidate.ringId === null && 
                       !currentChain.includes(candidate) && 
                       candidate.getPointName('origin', mode) === currentDest &&
                       candidate.auto === anchor.auto && 
                       candidate.trueStart > currentTrip.trueStart && 
                       candidate.trueStart >= (effectiveLastEnd + repark);
            });

            for (let candidate of candidates) {
                if (!isAlgoRunning) break; // Обриваємо перебір сусідів
                // ОБОВ'ЯЗКОВО додаємо await перед рекурсивним викликом
                await explore(candidate, [...currentChain, candidate]);
            }
        }

        // Запускаємо асинхронну рекурсію (з await)
        await explore(anchor, [anchor]);

        if (requireReturn) {
            const startPoint = anchor.getPointName('origin', mode);
            while (bestChain.length > 0) {
                let last = bestChain[bestChain.length - 1];
                if (last.getPointName('dest', mode) === startPoint) {
                    break;
                }
                bestChain.pop();
            }
        }
        // Після виходу з рекурсії перевіряємо, чи не зупинили ми алгоритм
        if (bestChain.length >= minTrips && isAlgoRunning) {
            const draftId = `draft_${Date.now()}_${ringCounter++}`;
            bestChain.forEach(t => t.ringId = draftId);
            
            renderDraft(); 
            // Пауза дає браузеру можливість "почути" клік по кнопці "Зупинити"
            await new Promise(resolve => setTimeout(resolve, 10)); 
        }
    }
}


// --- АЛГОРИТМ 4: ОПТИМІЗОВАНЕ ДЕРЕВО (Beam Search + Max Wait) ---
async function algoTreeOptimized(workingTrips, mode, repark, minTrips, requireReturn) {
    let ringCounter = 0;
    
    // НАЛАШТУВАННЯ АЛГОРИТМУ
    const BEAM_WIDTH = 3; // Залишаємо тільки 3 найкращі варіанти продовження
    const MAX_WAIT_MINS = 24 * 60; // Відсікаємо все, де машина чекає більше 48 годин

    for (let i = 0; i < workingTrips.length; i++) {
        if (!isAlgoRunning) break; 
        if (i % 10 === 0) await new Promise(resolve => setTimeout(resolve, 0));

        let anchor = workingTrips[i];
        if (anchor.ringId !== null || anchor.logisticDay > 5) continue;

        let bestChain = [];
        let exploreCounter = 0; 

        async function explore(currentTrip, currentChain) {
            if (!isAlgoRunning) return; 

            exploreCounter++;
            if (exploreCounter % 200 === 0) {
                await new Promise(resolve => setTimeout(resolve, 0));
            }
            
            if (!isAlgoRunning) return; 

            if (currentChain.length > bestChain.length) {
                bestChain = [...currentChain];
            }

            let effectiveLastEnd = currentTrip.trueEnd + (currentTrip.trueEnd < currentTrip.trueStart ? 10080 : 0);
            const currentDest = currentTrip.getPointName('dest', mode);
            const isCurrentEmpty = String(currentTrip.type || '').trim().toLowerCase() === "порожній";
            let candidates = workingTrips.filter(candidate => {
                const isCandidateEmpty = String(candidate.type || '').trim().toLowerCase() === "порожній";
                if (isCurrentEmpty && isCandidateEmpty) return false; // Забороняємо 2 порожніх підряд
                if (candidate.ringId !== null) return false;
                if (currentChain.includes(candidate)) return false;
                if (candidate.getPointName('origin', mode) !== currentDest) return false;
                if (candidate.auto !== anchor.auto) return false;
                
                let waitTime = candidate.trueStart - effectiveLastEnd;
                
                // 1. Time-Window Pruning: Перевіряємо, чи вписується простій у вікно
                return waitTime >= repark && waitTime <= MAX_WAIT_MINS;
            });

            // 2. Beam Search: Сортуємо кандидатів за часом простою (від найменшого до найбільшого)
            candidates.sort((a, b) => {
                let waitA = a.trueStart - effectiveLastEnd;
                let waitB = b.trueStart - effectiveLastEnd;
                return waitA - waitB;
            });

            // Залишаємо лише ТОП-N кандидатів
            let topCandidates = candidates.slice(0, BEAM_WIDTH);

            // Рекурсивно запускаємо тільки для найкращих
            for (let candidate of topCandidates) {
                if (!isAlgoRunning) break; 
                await explore(candidate, [...currentChain, candidate]);
            }
        }

        await explore(anchor, [anchor]);

        if (requireReturn) {
            const startPoint = anchor.getPointName('origin', mode);
            while (bestChain.length > 0) {
                let last = bestChain[bestChain.length - 1];
                if (last.getPointName('dest', mode) === startPoint) {
                    break;
                }
                bestChain.pop();
            }
        }

        if (bestChain.length >= minTrips && isAlgoRunning) {
            const draftId = `draft_opt_${Date.now()}_${ringCounter++}`;
            bestChain.forEach(t => t.ringId = draftId);
            
            renderDraft(); 
            await new Promise(resolve => setTimeout(resolve, 10)); 
        }
    }
}


// Функція видалення конкретного кільця (додайте її)
/*function deleteRingFromDraft(draftId) {
    window.allTrips.forEach(t => {
        if (t.ringId === draftId) t.ringId = null;
    });
    renderDraft();
    render(window.allTrips);
}*/

function deleteRingFromDraft(draftRingId) {
    window.allTrips.forEach(t => {
        if (t.ringId === draftRingId) {
            // Якщо це було розширене кільце з архіву - повертаємо старий ID
            if (t.originalRingId) {
                t.ringId = t.originalRingId;
                t.originalRingId = null;
            } else {
                t.ringId = null; // Якщо це звичайний новий графік - викидаємо в реєстр
            }
        }
    });
    cleanOrphanedEmpties(); // ПРИБИРАЄМО СМІТТЯ
    renderDraft();
    renderArchive(); // Оновлюємо архів, бо туди могли повернутися кільця
    
    if (document.getElementById('register-tab').classList.contains('active')) {
        render(window.allTrips);
    }
}

// ФУНКЦІЯ ОЧИЩЕННЯ ЧЕРНЕТКИ
function clearDraft() {
    window.allTrips.forEach(t => {
        if (t.ringId && t.ringId.startsWith('draft_')) {
            t.ringId = null;
        }
    });
    cleanOrphanedEmpties(); // ПРИБИРАЄМО СМІТТЯ
    render(window.allTrips);
    renderDraft(); // Перемалювати порожню чернетку
}

// ФУНКЦІЯ ЗАТВЕРДЖЕННЯ КІЛЬЦЯ
function approveRing(draftRingId) {
    const approvedId = draftRingId.replace('draft_', 'approved_');
    window.allTrips.forEach(t => {
        if (t.ringId === draftRingId) {
            t.ringId = approvedId;
            t.originalRingId = null; // Очищаємо пам'ять про відкат, кільце затверджено
        }
    });
    renderDraft();
    renderArchive();
}

// --- ЛОГІКА НЕСКІНЧЕННОГО СКРОЛЛУ ---

function setupDraftObserver() {
    if (draftObserver) draftObserver.disconnect();
    const sentinel = document.getElementById('draft-sentinel');
    draftObserver = new IntersectionObserver((entries) => {
        if (entries[0].isIntersecting) loadMoreDrafts();
    }, { root: document.getElementById('draft-scroll'), rootMargin: '200px' });
    draftObserver.observe(sentinel);
}

function loadMoreDrafts() {
    if (draftRenderedCount >= draftCardsHTML.length) return;
    const content = document.getElementById('draft-content');
    const nextBatch = draftCardsHTML.slice(draftRenderedCount, draftRenderedCount + CARDS_PER_PAGE);
    
    if (draftRenderedCount === 0) content.innerHTML = ''; // Очищаємо перед першою партією
    content.insertAdjacentHTML('beforeend', nextBatch.join(''));
    draftRenderedCount += CARDS_PER_PAGE;
}

function setupArchiveObserver() {
    if (archiveObserver) archiveObserver.disconnect();
    const sentinel = document.getElementById('archive-sentinel');
    archiveObserver = new IntersectionObserver((entries) => {
        if (entries[0].isIntersecting) loadMoreArchives();
    }, { root: document.getElementById('archive-scroll'), rootMargin: '200px' });
    archiveObserver.observe(sentinel);
}

function loadMoreArchives() {
    if (archiveRenderedCount >= archiveCardsHTML.length) return;
    const content = document.getElementById('archive-content');
    const nextBatch = archiveCardsHTML.slice(archiveRenderedCount, archiveRenderedCount + CARDS_PER_PAGE);
    
    if (archiveRenderedCount === 0) content.innerHTML = '';
    content.insertAdjacentHTML('beforeend', nextBatch.join(''));
    archiveRenderedCount += CARDS_PER_PAGE;
}

// --- ОНОВЛЕНІ ФУНКЦІЇ РЕНДЕРУ ---

function renderDraft() {
    const mode = modeSelect.value;
    const draftRingsMap = {};
    
    window.allTrips.forEach(t => {
        if (t.ringId && t.ringId.startsWith('draft_')) {
            if (!draftRingsMap[t.ringId]) draftRingsMap[t.ringId] = [];
            draftRingsMap[t.ringId].push(t);
        }
    });

    const rings = Object.values(draftRingsMap);

    // ==========================================
    // НОВІ РЯДКИ: Оновлюємо цифру на екрані
    const countBadge = document.getElementById('draft-count-badge');
    if (countBadge) countBadge.innerText = rings.length;
    // ==========================================

    if (rings.length === 0) {
        document.getElementById('draft-content').innerHTML = '<div class="empty-msg" style="width:100%;">Чернетка порожня. Запустіть алгоритм або закольцуйте вручну.</div>';
        draftCardsHTML = [];
        return;
    }

    // Генеруємо масив усіх карток, але поки не вставляємо в DOM
    draftCardsHTML = rings.map((ring, idx) => {
        const rId = ring[0].ringId;
        ring.sort((a, b) => {
            if (a.logisticDay !== b.logisticDay) return a.logisticDay - b.logisticDay;
            return a.trueStart - b.trueStart;
        });

        // НОВОЕ: Определяем цвет шапки
        const autoType = String(ring[0].auto || "").toUpperCase();
        const headerBg = autoType.includes("БДФ") ? "#c9fed8" : "#c1d7ff";

        const isExtended = rId.includes('_ext_');
        const titleText = isExtended ? `🔄 Докільцьований наряд #${idx + 1} (${ring[0].auto})` : `Наряд #${idx + 1} (${ring[0].auto})`;
        const deleteBtnText = isExtended ? `❌ Відмінити докільцювання` : `🗑️ Видалити`;

        return `
        <div class="ring-card" id="${rId}">
            <div class="ring-header" style="background: ${headerBg};">
                <strong>${titleText}</strong>
                <div style="gap: 10px; display: flex;">
                    <button class="action-btn" onclick="editRing('${rId}')">✏️ Редагувати</button>
                    <button class="success-btn" onclick="approveRing('${rId}')">✔ Затвердити</button>
                    <button class="danger-btn" onclick="deleteRingFromDraft('${rId}')">${deleteBtnText}</button>
                </div>
            </div>
            <div class="table-container mini-table">
                <table>
                    <thead>
                        <tr>
                            <th class="col-short">GRF</th>
                            <th class="col-med">Тип</th>
                            <th class="col-short">ФЗ</th>
                            <th class="col-long">Маршрут</th>
                            <th class="col-med">Відправник</th><th class="col-med">Отримувач</th>
                            <th class="col-day">Пн</th><th class="col-day">Вт</th><th class="col-day">Ср</th>
                            <th class="col-day">Чт</th><th class="col-day">Пт</th><th class="col-day">Сб</th><th class="col-day">Нд</th>
                            <th class="col-short">Подача</th>
                            <th class="col-short">Виїзд</th>
                            <th class="col-short">Приїзд</th>
                            <th class="col-short">Вільний</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${ring.map(t => `
                            <tr>
                                <td>${t.grf}</td>
                                <td>${t.type || ''}</td>
                                <td>${t.load || ''}</td>
                                <td title="${t.route}">${t.route}</td>
                                <td>${t.getPointName('origin', mode)}</td><td>${t.getPointName('dest', mode)}</td>
                                ${t.logDays.map(d => `<td class="col-day">${d === '+' ? '+' : ''}</td>`).join('')}
                                <td class="time-cell" title="Подача: ${t.podachaInt} хв.&#10;TrueStart: ${t.trueStart} хв.">${t.podachaStr}</td>
                                <td class="time-cell" title="Виїзд: ${t.depInt} хв.&#10;TrueStart: ${t.trueStart} хв.">${t.depStr}</td>
                                <td class="time-cell" title="Приїзд: ${t.arrInt} хв.&#10;TrueEnd: ${t.trueEnd} хв.">${t.arrStr}</td>
                                <td class="time-cell" title="Вільний: ${t.freeInt} хв.&#10;TrueEnd: ${t.trueEnd} хв.">${t.freeStr}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
        </div>`;
    });

    // Скидаємо лічильник та ініціюємо перше завантаження
    draftRenderedCount = 0;
    document.getElementById('draft-content').innerHTML = ''; 
    loadMoreDrafts();
    setupDraftObserver();
    updateDraftImbalanceStats();
}

function updateArchiveStats() {
    if (!window.allTrips) return;

    let totalTrips = window.allTrips.length;
    let ringedTrips = 0;
    let approvedRingsMap = {};

    // Рахуємо графіки та збираємо унікальні кільця
    window.allTrips.forEach(t => {
        if (t.ringId && t.ringId.startsWith('approved_')) {
            ringedTrips++;
            if (!approvedRingsMap[t.ringId]) {
                approvedRingsMap[t.ringId] = { auto: t.auto };
            }
        }
    });

    let totalRings = Object.keys(approvedRingsMap).length;
    let bdfRings = 0;
    let otherRings = 0;

    // Рахуємо типи авто в кільцях
    Object.values(approvedRingsMap).forEach(ring => {
        if (String(ring.auto || "").toUpperCase().includes("БДФ")) {
            bdfRings++;
        } else {
            otherRings++;
        }
    });

    // Рахуємо відсоток кільцювання
    let percent = totalTrips > 0 ? ((ringedTrips / totalTrips) * 100).toFixed(1) : 0;

    // Оновлюємо DOM
    document.getElementById('stat-total-rings').innerText = totalRings;
    document.getElementById('stat-bdf-rings').innerText = bdfRings;
    document.getElementById('stat-other-rings').innerText = otherRings;
    document.getElementById('stat-total-trips').innerText = totalTrips;
    document.getElementById('stat-ringed-trips').innerText = ringedTrips;
    document.getElementById('stat-percent').innerText = `${percent}%`;
}

function renderArchive() {
    const mode = modeSelect.value;
    const archiveMap = {};
    
    window.allTrips.forEach(t => {
        if (t.ringId && t.ringId.startsWith('approved_')) {
            if (!archiveMap[t.ringId]) archiveMap[t.ringId] = [];
            archiveMap[t.ringId].push(t);
        }
    });

    const rings = Object.values(archiveMap);
    const searchTerm = document.getElementById('archive_search')?.value.toLowerCase().trim() || "";

    // Підготовлюємо масив з іменами, щоб не втратити оригінальну нумерацію при фільтрації
    let processedRings = rings.map((ring, idx) => {
        const rId = ring[0].ringId;
        const displayName = window.ringNamesMap[rId] ? window.ringNamesMap[rId] : `Затверджений наряд #${idx + 1}`;
        return { ring, rId, displayName };
    });

    // Фільтруємо, якщо щось введено в поле пошуку
    if (searchTerm) {
        processedRings = processedRings.filter(item => item.displayName.toLowerCase().includes(searchTerm));
    }

    if (processedRings.length === 0) {
        document.getElementById('archive-content').innerHTML = '<div class="empty-msg" style="width:100%;">Поки що немає затверджених кілець або за вашим запитом нічого не знайдено.</div>';
        archiveCardsHTML = [];
        return;
    }

    archiveCardsHTML = processedRings.map((item) => {
        const { ring, rId, displayName } = item;
        
        ring.sort((a, b) => {
            if (a.logisticDay !== b.logisticDay) return a.logisticDay - b.logisticDay;
            return a.trueStart - b.trueStart;
        });

        // НОВОЕ: Определяем цвет шапки
        const autoType = String(ring[0].auto || "").toUpperCase();
        const headerBg = autoType.includes("БДФ") ? "#c9fed8" : "#c1d7ff";

        return `
        <div class="ring-card approved" id="${rId}">
            <div class="ring-header" style="background: ${headerBg};">
                <strong>Наряд: ${displayName} (${ring[0].auto})</strong>
                <div style="gap: 10px; display: flex;">
                    <button class="action-btn" onclick="openStapler('${rId}')">📎 Знайти пару</button>
                    <button class="action-btn" onclick="editRing('${rId}')">✏️ Редагувати</button>
                    <button class="danger-btn" onclick="deleteRingFromArchive('${rId}')">🗑️ Видалити</button>
                </div>
            </div>
            <div class="table-container mini-table">
                <table>
                    <thead>
                        <tr>
                            <th class="col-short">GRF</th>
                            <th class="col-med">Тип</th>
                            <th class="col-short">ФЗ</th>
                            <th class="col-long">Маршрут</th>
                            <th class="col-med">Відправник</th><th class="col-med">Отримувач</th>
                            <th class="col-day">Пн</th><th class="col-day">Вт</th><th class="col-day">Ср</th>
                            <th class="col-day">Чт</th><th class="col-day">Пт</th><th class="col-day">Сб</th><th class="col-day">Нд</th>
                            <th class="col-short">Подача</th>
                            <th class="col-short">Виїзд</th>
                            <th class="col-short">Приїзд</th>
                            <th class="col-short">Вільний</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${ring.map(t => `
                            <tr>
                                <td>${t.grf}</td>
                                <td>${t.type || ''}</td>
                                <td>${t.load || ''}</td>
                                <td title="${t.route}">${t.route}</td>
                                <td>${t.getPointName('origin', mode)}</td><td>${t.getPointName('dest', mode)}</td>
                                ${t.logDays.map(d => `<td class="col-day">${d === '+' ? '+' : ''}</td>`).join('')}
                                <td class="time-cell" title="Подача: ${t.podachaInt} хв.&#10;TrueStart: ${t.trueStart} хв.">${t.podachaStr}</td>
                                <td class="time-cell" title="Виїзд: ${t.depInt} хв.&#10;TrueStart: ${t.trueStart} хв.">${t.depStr}</td>
                                <td class="time-cell" title="Приїзд: ${t.arrInt} хв.&#10;TrueEnd: ${t.trueEnd} хв.">${t.arrStr}</td>
                                <td class="time-cell" title="Вільний: ${t.freeInt} хв.&#10;TrueEnd: ${t.trueEnd} хв.">${t.freeStr}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
        </div>`;
    });

    archiveRenderedCount = 0;
    document.getElementById('archive-content').innerHTML = '';
    loadMoreArchives();
    updateArchiveStats();
    setupArchiveObserver();
}

function approveAllRings() {
    let hasDrafts = false;
    window.allTrips.forEach(t => {
        if (t.ringId && t.ringId.startsWith('draft_')) {
            t.ringId = t.ringId.replace('draft_', 'approved_');
            hasDrafts = true;
        }
    });
    
    if (hasDrafts) {
        renderDraft(); // Очистит вкладку черновика
        renderArchive(); // Добавит все в архив
    }
}

function unapproveRing(approvedId) {
    const draftId = approvedId.replace('approved_', 'draft_');
    window.allTrips.forEach(t => {
        if (t.ringId === approvedId) {
            t.ringId = draftId;
        }
    });
    renderArchive();
    renderDraft();
}

// Відкриття спливаючого вікна
function openFilter(event, column) {
    event.stopPropagation();
    currentFilterColumn = column;
    
    const popup = document.getElementById('filter-popup');
    const itemsContainer = document.getElementById('filter-popup-items');
    const selectAllCb = document.getElementById('filter-select-all');
    const mode = modeSelect.value;
    
    // Збираємо унікальні значення з поточних даних
    const uniqueValues = new Set();
    window.allTrips.forEach(t => {
        let val;
        // ЗМІНЕНО: Беремо сирі дані складу
        if (column === 'origin') val = t.origin;
        else if (column === 'dest') val = t.destination;
        else if (column === 'auto') val = t.auto;
        else if (column === 'type') val = t.type;
        
        if (val) uniqueValues.add(val);
    });
    
    // Рендеримо чекбокси
    const sortedVals = Array.from(uniqueValues).sort();
    itemsContainer.innerHTML = sortedVals.map(val => {
        const isChecked = activeFilters[column].has(val) ? 'checked' : '';
        return `
            <label class="filter-popup-item">
                <input type="checkbox" value="${val}" class="filter-cb" ${isChecked}>
                ${val}
            </label>
        `;
    }).join('');
    
    // Перевіряємо, чи обрані всі, для галочки "Вибрати всі"
    const allChecked = Array.from(itemsContainer.querySelectorAll('.filter-cb')).every(cb => cb.checked);
    selectAllCb.checked = sortedVals.length > 0 && allChecked && activeFilters[column].size > 0;
    
    // Позиціонування вікна під колонкою
    const rect = event.target.getBoundingClientRect();
    popup.style.display = 'block';
    popup.style.top = `${rect.bottom + window.scrollY + 5}px`;
    
    // Перевіряємо, чи не вилазить за правий край екрану
    let leftPos = rect.left + window.scrollX;
    if (leftPos + 200 > window.innerWidth) leftPos = window.innerWidth - 210;
    popup.style.left = `${leftPos}px`;
}

// Кнопка "Вибрати всі"
function toggleAllFilterItems(mainCheckbox) {
    document.querySelectorAll('.filter-cb').forEach(cb => {
        cb.checked = mainCheckbox.checked;
    });
}

// Застосування фільтра
function applyPopupFilter() {
    if (!currentFilterColumn) return;
    activeFilters[currentFilterColumn].clear();
    
    document.querySelectorAll('.filter-cb').forEach(cb => {
        if (cb.checked) activeFilters[currentFilterColumn].add(cb.value);
    });
    
    closeFilterPopup();
    updateFilterIcons();
    render(window.allTrips);
}

// Скидання поточного фільтра
function clearPopupFilter() {
    if (!currentFilterColumn) return;
    activeFilters[currentFilterColumn].clear();
    closeFilterPopup();
    updateFilterIcons();
    render(window.allTrips);
}

// Закриття вікна
function closeFilterPopup() {
    document.getElementById('filter-popup').style.display = 'none';
}

// Оновлення кольору іконок у шапці
function updateFilterIcons() {
    ['origin', 'dest', 'auto', 'type'].forEach(col => {
        const icon = document.querySelector(`.filter-icon[data-col="${col}"]`);
        if (icon) {
            if (activeFilters[col].size > 0) {
                icon.classList.add('active');
            } else {
                icon.classList.remove('active');
            }
        }
    });
}

// Закривати попап при кліку поза ним
document.addEventListener('click', (e) => {
    const popup = document.getElementById('filter-popup');
    if (popup.style.display === 'block' && !popup.contains(e.target) && !e.target.classList.contains('filter-icon')) {
        closeFilterPopup();
    }
});


// ==========================================
// ЛОГІКА РУЧНОГО КОНСТРУКТОРА КІЛЕЦЬ
// ==========================================
let constructorRing = []; // Масив графіків, які ми зараз збираємо
let constructorOriginalStatus = 'draft';

let constructorEditingRingId = null; 
let editingTripsBackup = [];

// 1. Слухач подвійного кліку по головній таблиці
document.getElementById('table_body').addEventListener('dblclick', function(e) {
    const tr = e.target.closest('tr');
    if (!tr || !tr.dataset.id) return;
    
    // Знімаємо виділення тексту, якщо браузер встиг його зробити
    window.getSelection().removeAllRanges();
    
    startConstructor(tr.dataset.id);
});

// 2. Запуск конструктора
function startConstructor(tripId) {
    const trip = window.allTrips.find(t => t.id === tripId);
    if (!trip || trip.ringId !== null) return; // Якщо вже в кільці - ігноруємо

    constructorRing = [trip]; // Починаємо нове кільце
    
    // Показуємо вкладку і перемикаємося на неї
    document.getElementById('constructor-tab-btn').style.display = 'block';
    switchTab('constructor-tab');
    
    renderConstructor();
}

// 3. Рендер панелей Конструктора
function renderConstructor() {
    const mode = modeSelect.value;
    const repark = parseInt(reparkInput.value) || 0;
    
    // Рендер ВЕРХНЬОЇ панелі (поточне кільце)
    const topBody = document.getElementById('constructor-current-body');
    topBody.innerHTML = constructorRing.map((t, idx) => {
        // Дозволяємо видаляти лише останній доданий графік (щоб не розірвати ланцюг)
        let actionBtn = '';
        if (idx === 0) {
            // Кнопка для першого (індекс 0)
            actionBtn = `<button class="btn-remove" onclick="removeFromConstructor(0)" title="Відсікти початок">X</button>`;
        } else if (idx === constructorRing.length - 1) {
            // Кнопка для останнього
            actionBtn = `<button class="btn-remove" onclick="removeFromConstructor(${idx})" title="Відсікти кінець">X</button>`;
        }
        
        return `
        <tr>
            <td>${actionBtn}</td>
            <td>${t.grf}</td>
            <td>${t.type || ''}</td>
            <td>${t.load || ''}</td>
            <td title="${t.route}">${t.route}</td>
            <td>${t.getPointName('origin', mode)}</td><td>${t.getPointName('dest', mode)}</td>
            ${t.logDays.map(d => `<td class="col-day">${d === '+' ? '+' : ''}</td>`).join('')}
            <td class="time-cell">${t.podachaStr}</td>
            <td class="time-cell">${t.depStr}</td>
            <td class="time-cell">${t.arrStr}</td>
            <td class="time-cell">${t.freeStr}</td>
        </tr>`;
    }).join('');

    // Знаходимо кандидатів для НИЖНЬОЇ панелі
    const lastTrip = constructorRing[constructorRing.length - 1];
    const targetOrigin = lastTrip.getPointName('dest', mode);
    let effectiveLastEnd = lastTrip.trueEnd + (lastTrip.trueEnd < lastTrip.trueStart ? 10080 : 0);
    const isLastEmpty = String(lastTrip.type || '').trim().toLowerCase() === "порожній";
    const candidates = window.allTrips.filter(t => {
        const isCandidateEmpty = String(t.type || '').trim().toLowerCase() === "порожній";
        if (isLastEmpty && isCandidateEmpty) return false; // Забороняємо пропонувати порожній після порожнього
        if (t.ringId !== null) return false; // Вже зайнятий
        if (constructorRing.includes(t)) return false; // Вже в цьому кільці
        if (t.getPointName('origin', mode) !== targetOrigin) return false; // Не збігається місто
        if (t.auto !== lastTrip.auto) return false; // Різні авто
        if (t.trueStart < lastTrip.trueStart) return false; // Захист від подорожі в минуле
        if (t.trueStart < (effectiveLastEnd + repark)) return false; // Не встигає з перепарковкою
        return true;
    });

    // НОВЕ: Сортуємо знайдених кандидатів
    candidates.sort((a, b) => {
        if (a.trueStart !== b.trueStart) {
            return a.trueStart - b.trueStart;
        }
        return a.logisticDay - b.logisticDay;
    });

    // Рендер НИЖНЬОЇ панелі
    const bottomBody = document.getElementById('constructor-candidates-body');
    if (candidates.length === 0) {
        bottomBody.innerHTML = `<tr><td colspan="18" style="text-align:center; padding: 20px; color:#666;">Немає доступних продовжень для цього маршруту 🤷‍♂️</td></tr>`;
    } else {
        bottomBody.innerHTML = candidates.map(t => `
        <tr>
            <td><button class="btn-add" onclick="addToConstructor('${t.id}')">+</button></td>
            <td>${t.grf}</td>
            <td>${t.type || ''}</td>
            <td>${t.load || ''}</td>
            <td title="${t.route}">${t.route}</td>
            <td>${t.getPointName('origin', mode)}</td><td>${t.getPointName('dest', mode)}</td>
            ${t.logDays.map(d => `<td class="col-day">${d === '+' ? '+' : ''}</td>`).join('')}
            <td class="time-cell" title="Подача: ${t.podachaInt} хв.&#10;TrueStart: ${t.trueStart} хв.">${t.podachaStr}</td>
            <td class="time-cell" title="Виїзд: ${t.depInt} хв.&#10;TrueStart: ${t.trueStart} хв.">${t.depStr}</td>
            <td class="time-cell" title="Приїзд: ${t.arrInt} хв.&#10;TrueEnd: ${t.trueEnd} хв.">${t.arrStr}</td>
            <td class="time-cell" title="Вільний: ${t.freeInt} хв.&#10;TrueEnd: ${t.trueEnd} хв.">${t.freeStr}</td>
        </tr>`).join('');
    }
}

// 4. Додавання до кільця
function addToConstructor(tripId) {
    const trip = window.allTrips.find(t => t.id === tripId);
    if (trip) {
        constructorRing.push(trip);
        renderConstructor();
    }
}

// 5. Видалення останнього рейсу з кільця
function removeFromConstructor(index) {
    // Якщо індекс не передано, за замовчуванням видаляємо останній
    if (index === undefined) index = constructorRing.length - 1;

    if (constructorRing.length > 1) {
        if (index === 0) {
            // Видаляємо перший елемент масиву
            constructorRing.shift(); 
        } else {
            // Видаляємо останній елемент масиву
            constructorRing.pop(); 
        }
        cleanOrphanedEmpties();
        renderConstructor();
    } else {
        cancelConstructor(); // Якщо видалили останній/єдиний графік - закриваємо конструктор
    }
}

// 6. Скасування (вихід)
function cancelConstructor() {
    if (constructorEditingRingId) {
        // Якщо ми редагували існуюче кільце - відновлюємо його з бекапу
        editingTripsBackup.forEach(t => {
            t.ringId = constructorEditingRingId;
        });

        // Повертаємося на ту вкладку, звідки прийшли
        if (constructorOriginalStatus === 'approved') {
            renderArchive();
            switchTab('archive-tab');
        } else {
            renderDraft();
            switchTab('draft-tab');
        }
    } else {
        // Якщо ми збирали кільце з нуля - просто закриваємо і йдемо в Реєстр
        switchTab('register-tab');
    }

    // Очищаємо всі змінні
    constructorRing = [];
    editingTripsBackup = [];
    constructorEditingRingId = null;
    constructorOriginalStatus = 'draft';
    
    document.getElementById('constructor-tab-btn').style.display = 'none';
    render(window.allTrips); // Оновлюємо реєстр
}

// 7. Збереження кільця (відправляємо в чернетку)
// 7. Збереження кільця (відправляємо куди треба)
function saveConstructorRing() {
    if (constructorRing.length === 0) return;
    
    const prefix = constructorOriginalStatus === 'approved' ? 'approved' : 'draft';
    const newId = `${prefix}_${Date.now()}_manual`;
    
    constructorRing.forEach(t => t.ringId = newId);
    
    constructorRing = [];
    document.getElementById('constructor-tab-btn').style.display = 'none';
    
    if (constructorOriginalStatus === 'approved') {
        renderArchive();
        switchTab('archive-tab');
    } else {
        renderDraft();
        switchTab('draft-tab');
    }

    // ==========================================
    // НОВЕ: Скидаємо всі статуси і бекапи
    constructorOriginalStatus = 'draft'; 
    constructorEditingRingId = null;
    editingTripsBackup = [];
    // ==========================================
    
    render(window.allTrips); 
}

// 8. Редагування існуючого кільця
function editRing(ringId) {
    // Знаходимо всі графіки, які належать до цього кільця
    const ringTrips = window.allTrips.filter(t => t.ringId === ringId);
    
    if (ringTrips.length === 0) return;

    // НОВЕ: Запам'ятовуємо, чи це кільце з архіву
    constructorOriginalStatus = ringId.startsWith('approved_') ? 'approved' : 'draft';

    //==========================================
    // НОВЕ: Запам'ятовуємо оригінальне кільце для відкату
    constructorEditingRingId = ringId;
    editingTripsBackup = [...ringTrips];
    // ==========================================

    // Обов'язково сортуємо їх у правильному порядку (як вони йшли в кільці)
    ringTrips.sort((a, b) => {
        if (a.logisticDay !== b.logisticDay) return a.logisticDay - b.logisticDay;
        return a.trueStart - b.trueStart;
    });

    // Завантажуємо кільце в Конструктор
    constructorRing = [...ringTrips];

    // "Розбираємо" кільце: знімаємо статус затвердженого/чернетки
    ringTrips.forEach(t => t.ringId = null);

    // Оновлюємо всі вкладки, щоб кільце зникло з Чернетки/Архіву і з'явилося в Реєстрі
    renderDraft();
    renderArchive();
    render(window.allTrips); 

    // Перемикаємося на вкладку Конструктора
    document.getElementById('constructor-tab-btn').style.display = 'block';
    switchTab('constructor-tab');
    renderConstructor();
}

// ==========================================
// ЛОГІКА СТЕПЛЕРА (Об'єднання затверджених кілець)
// ==========================================

let currentStaplerSourceId = null;

function openStapler(sourceRingId) {
    const mode = document.getElementById('stapler_mode_select').value;
    const repark = parseInt(reparkInput.value) || 0;
    currentStaplerSourceId = sourceRingId;

    // Збираємо всі кільця з архіву в об'єкт
    const archiveMap = {};
    window.allTrips.forEach(t => {
        if (t.ringId && t.ringId.startsWith('approved_')) {
            if (!archiveMap[t.ringId]) archiveMap[t.ringId] = [];
            archiveMap[t.ringId].push(t);
        }
    });

    // Отримуємо вихідне кільце та сортуємо його
    const sourceRing = archiveMap[sourceRingId];
    sourceRing.sort((a, b) => {
        if (a.logisticDay !== b.logisticDay) return a.logisticDay - b.logisticDay;
        return a.trueStart - b.trueStart;
    });

    const lastTrip = sourceRing[sourceRing.length - 1];
    const targetOrigin = lastTrip.getPointName('dest', mode);
    let effectiveLastEnd = lastTrip.trueEnd + (lastTrip.trueEnd < lastTrip.trueStart ? 10080 : 0);
    const isLastEmpty = String(lastTrip.type || '').trim().toLowerCase() === "порожній";
    const candidatesRings = [];

    // Перебираємо інші кільця в архіві
    for (const [rId, ringTrips] of Object.entries(archiveMap)) {
        if (rId === sourceRingId) continue; // Самого себе пропускаємо

        ringTrips.sort((a, b) => {
            if (a.logisticDay !== b.logisticDay) return a.logisticDay - b.logisticDay;
            return a.trueStart - b.trueStart;
        });

        const firstTrip = ringTrips[0];
        const isCandidateEmpty = String(firstTrip.type || '').trim().toLowerCase() === "порожній";
        if (isLastEmpty && isCandidateEmpty) continue; // Не показуємо як варіант для склеювання

        // Перевіряємо умови стиковки
        if (firstTrip.auto !== lastTrip.auto) continue;
        if (firstTrip.getPointName('origin', mode) !== targetOrigin) continue;
        if (firstTrip.trueStart < (effectiveLastEnd + repark)) continue;

        candidatesRings.push(ringTrips);
    }

    const contentDiv = document.getElementById('stapler-content');
    
    if (candidatesRings.length === 0) {
        contentDiv.innerHTML = `<div class="empty-msg" style="width:100%;">Підходящих кілець для продовження не знайдено 🤷‍♂️</div>`;
    } else {
        // Рендеримо картки кандидатів
        contentDiv.innerHTML = candidatesRings.map((ring, idx) => {
            const rId = ring[0].ringId;
            
            // НОВОЕ: Определяем цвет шапки
            const autoType = String(ring[0].auto || "").toUpperCase();
            const headerBg = autoType.includes("БДФ") ? "#c9fed8" : "#c1d7ff";

            return `
            <div class="ring-card approved" style="margin: 0 auto 15px auto;">
                <div class="ring-header" style="background: ${headerBg};">
                    <strong>Можливе продовження (${ring[0].auto})</strong>
                    <button class="success-btn" onclick="stitchRings('${rId}')">🔗 Причепити сюди</button>
                </div>
                <div class="table-container mini-table">
                    <table>
                        <thead>
                            <tr>
                                <th class="col-short">GRF</th><th class="col-med">Тип</th><th class="col-short">ФЗ</th><th class="col-long">Маршрут</th>
                                <th class="col-med">Відправник</th><th class="col-med">Отримувач</th>
                                <th class="col-day">Пн</th><th class="col-day">Вт</th><th class="col-day">Ср</th><th class="col-day">Чт</th><th class="col-day">Пт</th><th class="col-day">Сб</th><th class="col-day">Нд</th>
                                <th class="col-short">Подача</th><th class="col-short">Виїзд</th><th class="col-short">Приїзд</th><th class="col-short">Вільний</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${ring.map(t => `
                                <tr>
                                    <td>${t.grf}</td><td>${t.type || ''}</td><td>${t.load || ''}</td><td title="${t.route}">${t.route}</td>
                                    <td>${t.getPointName('origin', mode)}</td><td>${t.getPointName('dest', mode)}</td>
                                    ${t.logDays.map(d => `<td class="col-day">${d === '+' ? '+' : ''}</td>`).join('')}
                                    <td class="time-cell" title="Подача: ${t.podachaInt} хв.&#10;TrueStart: ${t.trueStart} хв.">${t.podachaStr}</td>
                                    <td class="time-cell" title="Виїзд: ${t.depInt} хв.&#10;TrueStart: ${t.trueStart} хв.">${t.depStr}</td>
                                    <td class="time-cell" title="Приїзд: ${t.arrInt} хв.&#10;TrueEnd: ${t.trueEnd} хв.">${t.arrStr}</td>
                                    <td class="time-cell" title="Вільний: ${t.freeInt} хв.&#10;TrueEnd: ${t.trueEnd} хв.">${t.freeStr}</td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                </div>
            </div>`;
        }).join('');
    }

    document.getElementById('stapler-modal').style.display = 'flex';
}

function closeStapler() {
    document.getElementById('stapler-modal').style.display = 'none';
    currentStaplerSourceId = null;
}

function stitchRings(targetRingId) {
    if (!currentStaplerSourceId) return;

    const mergedRingId = currentStaplerSourceId; // Запоминаем ID исходного кольца

    // Всім графікам цільового кільця присвоюємо ID вихідного кільця
    window.allTrips.forEach(t => {
        if (t.ringId === targetRingId) {
            t.ringId = mergedRingId;
        }
    });

    closeStapler();
    renderArchive(); // Перемальовуємо архів (оновлює сторінку та скидає скрол)

    // Ищем индекс нашего обновленного кольца в массиве сгенерированных карточек
    const ringIndex = archiveCardsHTML.findIndex(html => html.includes(`id="${mergedRingId}"`));

    if (ringIndex !== -1) {
        // Если кольцо за пределами первых 20 отрендеренных, заставляем программу догрузить карточки
        while (archiveRenderedCount <= ringIndex) {
            loadMoreArchives();
        }

        // Даем браузеру миллисекунду на отрисовку элементов в DOM и делаем скролл
        setTimeout(() => {
            const targetElement = document.getElementById(mergedRingId);
            if (targetElement) {
                // Плавный скролл так, чтобы карточка оказалась по центру экрана
                targetElement.scrollIntoView({ behavior: 'smooth', block: 'center' });
                
                // Делаем красивую зеленую вспышку-подсветку, чтобы сразу найти обновленное кольцо
                targetElement.style.transition = 'box-shadow 0.4s ease-in-out';
                targetElement.style.boxShadow = '0 0 20px rgba(40, 167, 69, 0.8)';
                
                // Убираем подсветку через 2 секунды
                setTimeout(() => {
                    targetElement.style.boxShadow = '';
                }, 2000);
            }
        }, 100);
    }
}

// Функція для перетягування вікна
dragElement(document.getElementById("stapler-modal"));

function dragElement(elmnt) {
    let pos1 = 0, pos2 = 0, pos3 = 0, pos4 = 0;
    const header = document.getElementById("stapler-header");
    
    if (header) {
        header.onmousedown = dragMouseDown;
    } else {
        elmnt.onmousedown = dragMouseDown;
    }

    function dragMouseDown(e) {
        e.preventDefault();
        pos3 = e.clientX;
        pos4 = e.clientY;
        document.onmouseup = closeDragElement;
        document.onmousemove = elementDrag;
    }

    function elementDrag(e) {
        e.preventDefault();
        pos1 = pos3 - e.clientX;
        pos2 = pos4 - e.clientY;
        pos3 = e.clientX;
        pos4 = e.clientY;
        elmnt.style.top = (elmnt.offsetTop - pos2) + "px";
        elmnt.style.left = (elmnt.offsetLeft - pos1) + "px";
    }

    function closeDragElement() {
        document.onmouseup = null;
        document.onmousemove = null;
    }
}

// ==========================================
// ЛОГІКА ПОШУКУ ЗА ШАБЛОНОМ
// ==========================================

// --- АЛГОРИТМ ПОШУКУ ЗА ШАБЛОНОМ ---
async function runPatternAlgo() {
    const rawPattern = document.getElementById('pattern_input').value;
    const requireReturn = document.getElementById('global_return')?.checked || false;
    const mode = document.getElementById('mode_select').value; 
    
    // Переводимо весь шаблон у нижній регістр і чистимо пробіли для надійності
    const patternNodes = rawPattern.split('-').map(s => s.trim().toLowerCase()).filter(Boolean);
    
    if (patternNodes.length < 2) {
        alert("Введіть хоча б два міста, наприклад: Київ - Полтава");
        return;
    }

    const repark = parseInt(document.getElementById('repark_time').value) || 0;
    const minTrips = parseInt(document.getElementById('min_trips')?.value || 4);

    window.allTrips.forEach(t => { 
        if(t.ringId && t.ringId.startsWith('draft_')) t.ringId = null; 
    });

    const autoVals = Array.from(activeFilters.auto);
    const typeVals = Array.from(activeFilters.type);

    const workingTrips = window.allTrips.filter(t => {
        if (t.ringId !== null) return false;
        if (autoVals.length > 0 && !autoVals.includes(t.auto)) return false;
        if (typeVals.length > 0 && !typeVals.includes(t.type)) return false;
        return true;
    });

    let ringCounter = 0;
    const patternLength = patternNodes.length;

    for (let i = 0; i < workingTrips.length; i++) {
        if (!isAlgoRunning) break; 
        if (i % 50 === 0) await new Promise(resolve => setTimeout(resolve, 0));

        let anchor = workingTrips[i];
        if (anchor.ringId !== null || anchor.logisticDay > 4) continue;

        // БЕЗПЕЧНА ПЕРЕВІРКА ЯКОРЯ (Чистимо від пробілів та регістру)
        const anchorOriginCity = String(anchor.getPointName('origin', 'city')).trim().toLowerCase();
        const anchorDestCity = String(anchor.getPointName('dest', 'city')).trim().toLowerCase();

        if (anchorOriginCity !== patternNodes[0] || anchorDestCity !== patternNodes[1]) {
            continue;
        }

        let currentChain = [anchor];
        anchor.ringId = 'temp';
        let lastTrip = anchor;
        let searching = true;
        let step = 1; 

        while (searching && isAlgoRunning) {
            let nextOriginIdx = step % patternLength;
            let nextDestIdx = (step + 1) % patternLength;
            
            let targetCityOrigin = patternNodes[nextOriginIdx];
            let targetCityDest = patternNodes[nextDestIdx];

            let effectiveLastEnd = lastTrip.trueEnd + (lastTrip.trueEnd < lastTrip.trueStart ? 10080 : 0);
            
            let requiredOriginPoint = lastTrip.getPointName('dest', mode);
            const isLastEmpty = String(lastTrip.type || '').trim().toLowerCase() === "порожній";
            
            let nextTrip = workingTrips.find(c => {
                const isCandidateEmpty = String(c.type || '').trim().toLowerCase() === "порожній";
                if (isLastEmpty && isCandidateEmpty) return false; 
                
                // Безпечне отримання міст кандидата
                const cOriginCity = String(c.getPointName('origin', 'city')).trim().toLowerCase();
                const cDestCity = String(c.getPointName('dest', 'city')).trim().toLowerCase();

                return c.ringId === null &&
                       c.auto === anchor.auto && 
                       cOriginCity === targetCityOrigin &&
                       cDestCity === targetCityDest &&
                       c.getPointName('origin', mode) === requiredOriginPoint &&
                       c.trueStart > lastTrip.trueStart &&
                       c.trueStart >= (effectiveLastEnd + repark);
            });

            if (nextTrip) {
                nextTrip.ringId = 'temp';
                currentChain.push(nextTrip);
                lastTrip = nextTrip;
                step++;
            } else {
                searching = false;
            }
        }

        // БЕЗПЕЧНА ПЕРЕВІРКА ПОВЕРНЕННЯ
        if (requireReturn) {
            while (currentChain.length > 0) {
                let last = currentChain[currentChain.length - 1];
                let lastDestCity = String(last.getPointName('dest', 'city')).trim().toLowerCase();
                if (lastDestCity === patternNodes[0]) {
                    break;
                }
                let removed = currentChain.pop();
                removed.ringId = null; 
            }
        }

        if (currentChain.length >= minTrips) {
            const draftId = `draft_pattern_${Date.now()}_${ringCounter++}`;
            currentChain.forEach(t => t.ringId = draftId);
        } else {
            currentChain.forEach(t => t.ringId = null);
        }
    }
}

function runAutoStapler() {
    const mode = modeSelect.value;
    const repark = parseInt(reparkInput.value) || 0;
    const maxWaitMins = 100 * 60; // Максимум 24 години між кільцями
    const maxDurationMins = 10 * 1440; // Максимум 6 днів на весь новий наряд
    // 1. Збираємо всі затверджені кільця
    const archiveMap = {};
    window.allTrips.forEach(t => {
        if (t.ringId && t.ringId.startsWith('approved_')) {
            if (!archiveMap[t.ringId]) archiveMap[t.ringId] = [];
            archiveMap[t.ringId].push(t);
        }
    });

    const availableRings = Object.values(archiveMap);
    if (availableRings.length < 2) {
        alert("Недостатньо кілець в Архіві для склейки. Потрібно хоча б два.");
        return;
    }

    // 2. Готуємо зручні обгортки для кілець (щоб не рахувати старт/фініш по сто разів)
    const ringProps = availableRings.map(ring => {
        ring.sort((a, b) => {
            if (a.logisticDay !== b.logisticDay) return a.logisticDay - b.logisticDay;
            return a.trueStart - b.trueStart;
        });
        
        const lastTrip = ring[ring.length - 1];
        let effectiveEnd = lastTrip.trueEnd + (lastTrip.trueEnd < lastTrip.trueStart ? 10080 : 0);
        
        return {
            id: ring[0].ringId,
            trips: ring,
            firstStart: ring[0].trueStart,
            lastEnd: effectiveEnd,
            startOrigin: ring[0].getPointName('origin', mode),
            endDest: lastTrip.getPointName('dest', mode),
            auto: ring[0].auto,
            used: false
        };
    });

    // Сортуємо кільця хронологічно за першим виїздом
    ringProps.sort((a, b) => a.firstStart - b.firstStart);

    let draftCounter = 0;
    let changesMade = false;

    // 3. Жадібний пошук (йдемо по кожному кільцю і ліпимо до нього все, що знайдемо)
    for (let i = 0; i < ringProps.length; i++) {
        let currentRing = ringProps[i];
        if (currentRing.used) continue;

        let chain = [currentRing];
        currentRing.used = true;
        let searching = true;

        while (searching) {
            let lastInChain = chain[chain.length - 1];
            let currentEndMins = lastInChain.lastEnd;
            let targetOrigin = lastInChain.endDest;
            let chainStartTime = chain[0].firstStart;

            let bestCandidate = null;
            let minWait = Infinity;

            // Шукаємо найближче ідеальне продовження
            for (let j = 0; j < ringProps.length; j++) {
                let candidate = ringProps[j];
                if (candidate.used) continue;
                if (candidate.auto !== currentRing.auto) continue;
                if (candidate.startOrigin !== targetOrigin) continue;

                // НОВЕ: Перевірка на стику двох кілець
                const lastTripOfChain = lastInChain.trips[lastInChain.trips.length - 1];
                const firstTripOfCandidate = candidate.trips[0];
                const isLastEmpty = String(lastTripOfChain.type || '').trim().toLowerCase() === "порожній";
                const isCandidateEmpty = String(firstTripOfCandidate.type || '').trim().toLowerCase() === "порожній";
                if (isLastEmpty && isCandidateEmpty) continue; // Не зшиваємо 2 порожніх графіки на стику

                let waitTime = candidate.firstStart - currentEndMins;
                
                // Перевіряємо, чи вписуємося в перепарковку і ліміт простою
                if (waitTime >= repark && waitTime <= maxWaitMins) {
                    // Перевіряємо загальну довжину майбутнього наряду
                    let totalDuration = candidate.lastEnd - chainStartTime;
                    if (totalDuration <= maxDurationMins) {
                        if (waitTime < minWait) {
                            minWait = waitTime;
                            bestCandidate = candidate;
                        }
                    }
                }
            }

            if (bestCandidate) {
                bestCandidate.used = true;
                chain.push(bestCandidate);
            } else {
                searching = false; // Більше немає що причепити
            }
        }

        // Якщо вдалося склеїти хоча б 2 кільця — кидаємо їх в Чернетку Степлера
        if (chain.length > 1) {
            changesMade = true;
            const newDraftId = `stapler_draft_${Date.now()}_${draftCounter++}`;
            
            chain.forEach(rObj => {
                rObj.trips.forEach(t => {
                    t.originalRingId = t.ringId; // Запам'ятовуємо старе ім'я
                    t.ringId = newDraftId;       // Даємо нове тимчасове ім'я
                });
            });
        }
    }

    if (changesMade) {
        renderArchive();
        document.getElementById('stapler-tab-btn').style.display = 'block'; // Показываем вкладку!
        switchTab('stapler-draft-tab');
    } else {
        alert("Автостеплер нічого не знайшов");
    }
}

// 4. Функція малювання Чернетки Степлера
function renderStaplerDraft() {
    const mode = modeSelect.value;
    const staplerMap = {};
    
    window.allTrips.forEach(t => {
        if (t.ringId && t.ringId.startsWith('stapler_draft_')) {
            if (!staplerMap[t.ringId]) staplerMap[t.ringId] = [];
            staplerMap[t.ringId].push(t);
        }
    });

    const rings = Object.values(staplerMap);
    const content = document.getElementById('stapler-draft-content');

    if (rings.length === 0) {
        document.getElementById('stapler-tab-btn').style.display = 'none'; // Прячем кнопку вкладки
        content.innerHTML = ''; // Очищаем контент
        
        // Если мы находились на этой вкладке в момент очистки - перекидываем в Архив
        if (document.getElementById('stapler-draft-tab').classList.contains('active')) {
            switchTab('archive-tab');
        }
        return;
    }

    document.getElementById('stapler-tab-btn').style.display = 'block';

    content.innerHTML = rings.map((ring, idx) => {
        const rId = ring[0].ringId;
        ring.sort((a, b) => {
            if (a.logisticDay !== b.logisticDay) return a.logisticDay - b.logisticDay;
            return a.trueStart - b.trueStart;
        });

        // НОВОЕ: Определяем цвет шапки
        const autoType = String(ring[0].auto || "").toUpperCase();
        const headerBg = autoType.includes("БДФ") ? "#c9fed8" : "#c1d7ff";

        return `
        <div class="ring-card" id="${rId}" style="border: 2px solid #6f42c1 !important;">
            <div class="ring-header" style="background: ${headerBg};">
                <strong>🤖 Мега-збірка #${idx + 1} (${ring[0].auto})</strong>
                <div style="gap: 10px; display: flex;">
                    <button class="success-btn" onclick="approveStaplerRing('${rId}')">✔ Затвердити</button>
                    <button class="danger-btn" onclick="rejectStaplerRing('${rId}')">❌ Розбити назад</button>
                </div>
            </div>
            <div class="table-container mini-table">
                <table>
                    <thead>
                        <tr>
                            <th class="col-short">GRF</th><th class="col-med">Тип</th><th class="col-short">ФЗ</th><th class="col-long">Маршрут</th>
                            <th class="col-med">Відпр</th><th class="col-med">Отр</th>
                            <th class="col-day">Пн</th><th class="col-day">Вт</th><th class="col-day">Ср</th><th class="col-day">Чт</th><th class="col-day">Пт</th><th class="col-day">Сб</th><th class="col-day">Нд</th>
                            <th class="col-short">Подача</th><th class="col-short">Виїзд</th><th class="col-short">Приїзд</th><th class="col-short">Вільний</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${ring.map(t => `
                            <tr>
                                <td>${t.grf}</td><td>${t.type || ''}</td><td>${t.load || ''}</td><td title="${t.route}">${t.route}</td>
                                <td>${t.getPointName('origin', mode)}</td><td>${t.getPointName('dest', mode)}</td>
                                ${t.logDays.map(d => `<td class="col-day">${d === '+' ? '+' : ''}</td>`).join('')}
                                <td class="time-cell" title="Подача: ${t.podachaInt} хв.&#10;TrueStart: ${t.trueStart} хв.">${t.podachaStr}</td>
                                <td class="time-cell" title="Виїзд: ${t.depInt} хв.&#10;TrueStart: ${t.trueStart} хв.">${t.depStr}</td>
                                <td class="time-cell" title="Приїзд: ${t.arrInt} хв.&#10;TrueEnd: ${t.trueEnd} хв.">${t.arrStr}</td>
                                <td class="time-cell" title="Вільний: ${t.freeInt} хв.&#10;TrueEnd: ${t.trueEnd} хв.">${t.freeStr}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
        </div>`;
    }).join('');
}

// 5. Функції кнопок у Чернетці Степлера
function approveStaplerRing(draftId) {
    const finalId = draftId.replace('stapler_draft_', 'approved_');
    window.allTrips.forEach(t => {
        if (t.ringId === draftId) {
            t.ringId = finalId;
            t.originalRingId = null; // Очищаємо пам'ять, кільце злите назавжди
        }
    });
    renderStaplerDraft();
    renderArchive();
}

function rejectStaplerRing(draftId) {
    window.allTrips.forEach(t => {
        if (t.ringId === draftId) {
            t.ringId = t.originalRingId; // Відкочуємо до старих ID
            t.originalRingId = null;
        }
    });
    renderStaplerDraft();
    renderArchive();
}

function approveAllStapler() {
    window.allTrips.forEach(t => {
        if (t.ringId && t.ringId.startsWith('stapler_draft_')) {
            t.ringId = t.ringId.replace('stapler_draft_', 'approved_');
            t.originalRingId = null;
        }
    });
    renderStaplerDraft();
    renderArchive();
    switchTab('archive-tab');
}

function rejectAllStapler() {
    window.allTrips.forEach(t => {
        if (t.ringId && t.ringId.startsWith('stapler_draft_')) {
            t.ringId = t.originalRingId;
            t.originalRingId = null;
        }
    });
    renderStaplerDraft();
    renderArchive();
    switchTab('archive-tab');
}

function exportToExcel() {
    if (!window.allTrips || window.allTrips.length === 0) {
        alert("Немає даних для експорту.");
        return;
    }

    if (!window.uploadedFileData) {
        alert("Оригінальний файл не знайдено в пам'яті. Завантажте файл знову.");
        return;
    }

    // Читаем файл с параметрами для максимального сохранения форматов чисел и дат
    const data = new Uint8Array(window.uploadedFileData);
    const wb = XLSX.read(data, { type: 'array', cellStyles: true, cellNF: true }); 
    const wsName = wb.SheetNames[0];
    const ws = wb.Sheets[wsName];
    
    // Отримуємо межі оригінальної таблиці
    let range = XLSX.utils.decode_range(ws['!ref']);
    const newColIndex = range.e.c + 1; 

    // 2. Нумерація затверджених кілець
    const approvedRings = new Set();
    window.allTrips.forEach(t => {
        if (t.ringId && t.ringId.startsWith('approved_')) {
            approvedRings.add(t.ringId);
        }
    });

    const ringNumberMap = {};
    let ringCounter = 1;
    approvedRings.forEach(id => {
        ringNumberMap[id] = ringCounter++;
    });

    // 3. Записуємо заголовок колонки у 4-й рядок
    XLSX.utils.sheet_add_aoa(ws, [["Номер кільця"]], { origin: { r: 3, c: newColIndex } });

    // ДОПОМІЖНА ФУНКЦІЯ: перетворює "HH:MM" у дробове число для Excel
    const timeToExcelFraction = (timeStr) => {
        if (!timeStr) return "";
        const [h, m] = timeStr.split(':').map(Number);
        if (isNaN(h) || isNaN(m)) return timeStr;
        return (h * 60 + m) / 1440;
    };

    // 4. Розподіляємо дані
    const emptyTripsData = []; 
    
    window.allTrips.forEach(t => {
        let ringName = "";
        if (t.ringId && t.ringId.startsWith('approved_')) {
            ringName = window.ringNamesMap[t.ringId] || ringNumberMap[t.ringId];
        }

        if (t.originalRowIndex !== undefined) {
            // Оригінальний графік - точково записуємо номер в його ж рядок
            XLSX.utils.sheet_add_aoa(ws, [[ringName]], { origin: { r: t.originalRowIndex, c: newColIndex } });
        } else {
            // Це віртуальний порожній перегон
            let newRow = [...t.rawRow];
            
            // ПЕРЕТВОРЮЄМО ЧАС З ТЕКСТУ В ЧИСЛА EXCEL
            newRow[32] = timeToExcelFraction(t.podachaStr);
            newRow[33] = timeToExcelFraction(t.depStr);
            newRow[40] = timeToExcelFraction(t.arrStr);
            newRow[41] = timeToExcelFraction(t.freeStr);

            // Добиваємо рядок порожніми комірками до нової колонки
            while (newRow.length < newColIndex) {
                newRow.push("");
            }
            newRow[newColIndex] = ringName; 
            
            emptyTripsData.push(newRow);
        }
    });

    // 5. Дописуємо віртуальні перегони в самий низ і вішаємо на них маску часу
    if (emptyTripsData.length > 0) {
        const startEmptyRow = range.e.r + 1;
        XLSX.utils.sheet_add_aoa(ws, emptyTripsData, { origin: { r: startEmptyRow, c: 0 } });
        
        // ПРИМУСОВО задаємо маску часу "ГГ:ХВ" для доданих комірок
        for (let i = 0; i < emptyTripsData.length; i++) {
            const R = startEmptyRow + i;
            [32, 33, 40, 41].forEach(C => {
                const cellAddress = XLSX.utils.encode_cell({r: R, c: C});
                if (ws[cellAddress] && typeof ws[cellAddress].v === 'number') {
                    ws[cellAddress].z = 'hh:mm';
                }
            });
        }
    }

    // 6. Оновлюємо внутрішні межі аркуша
    const maxRow = range.e.r + emptyTripsData.length;
    ws['!ref'] = XLSX.utils.encode_range({s: {r:0, c:0}, e: {r: maxRow, c: newColIndex}});

    // 7. Зберігаємо файл
    XLSX.writeFile(wb, "Кольцмейстер_Експорт.xlsx");
}

function deleteRingFromArchive(archiveRingId) {
    if (confirm("Ви впевнені, що хочете видалити цей затверджений наряд? Графіки повернуться в реєстр.")) {
        window.allTrips.forEach(t => {
            if (t.ringId === archiveRingId) t.ringId = null;
        });
        cleanOrphanedEmpties(); // ПРИБИРАЄМО СМІТТЯ
        renderArchive();
        
        // Оновлюємо Реєстр, щоб графіки з'явилися там
        if (document.getElementById('register-tab').classList.contains('active')) {
            render(window.allTrips);
        }
    }
}

// ==========================================
// ЛОГІКА ДОКІЛЬЦЮВАННЯ ЗАТВЕРДЖЕНИХ КІЛЕЦЬ
// ==========================================
function extendApprovedRings() {
    const mode = modeSelect.value;
    const repark = parseInt(reparkInput.value) || 0;
    const maxWaitMins = 24 * 60; // Максимум 24 години очікування для нового рейсу
    let extendedCount = 0;

    // 1. Збираємо всі затверджені кільця
    const archiveMap = {};
    window.allTrips.forEach(t => {
        if (t.ringId && t.ringId.startsWith('approved_')) {
            if (!archiveMap[t.ringId]) archiveMap[t.ringId] = [];
            archiveMap[t.ringId].push(t);
        }
    });

    // 2. Фільтруємо вільні графіки
    const availableTrips = window.allTrips.filter(t => t.ringId === null);

    // 3. Проходимося по кожному архівному кільцю і пробуємо причепити хвіст
    for (const [rId, ringTrips] of Object.entries(archiveMap)) {
        // Сортуємо кільце, щоб знайти останній рейс
        ringTrips.sort((a, b) => {
            if (a.logisticDay !== b.logisticDay) return a.logisticDay - b.logisticDay;
            return a.trueStart - b.trueStart;
        });

        let currentChain = [...ringTrips];
        let searching = true;
        let addedNew = false;

        while (searching) {
            let lastTrip = currentChain[currentChain.length - 1];
            let effectiveLastEnd = lastTrip.trueEnd + (lastTrip.trueEnd < lastTrip.trueStart ? 10080 : 0);
            let targetOrigin = lastTrip.getPointName('dest', mode);
            const isLastEmpty = String(lastTrip.type || '').trim().toLowerCase() === "порожній";
            // Шукаємо всі можливі продовження
            let candidates = availableTrips.filter(c => {
                const isCandidateEmpty = String(c.type || '').trim().toLowerCase() === "порожній";
                if (isLastEmpty && isCandidateEmpty) return false;
                return c.ringId === null &&
                       c.auto === lastTrip.auto &&
                       c.getPointName('origin', mode) === targetOrigin &&
                       c.trueStart >= (effectiveLastEnd + repark) &&
                       (c.trueStart - effectiveLastEnd) <= maxWaitMins; 
            });

            if (candidates.length > 0) {
                // Беремо найближчий за часом (жадібний пошук)
                candidates.sort((a, b) => a.trueStart - b.trueStart);
                let bestNext = candidates[0];
                
                bestNext.ringId = 'temp'; // Тимчасово бронюємо
                currentChain.push(bestNext);
                addedNew = true;
            } else {
                searching = false; // Більше немає продовжень
            }
        }

        // Якщо вдалося щось причепити — міняємо ID і відправляємо в Чернетку
        if (addedNew) {
            const newDraftId = `draft_ext_${Date.now()}_${extendedCount++}`;
            currentChain.forEach(t => {
                // Для старих рейсів запам'ятовуємо їхній архівний ID для можливого відкату
                if (t.ringId !== 'temp') {
                    t.originalRingId = t.ringId; 
                } else {
                    t.originalRingId = null; // Це нові графіки, вони при відкаті просто стануть null
                }
                t.ringId = newDraftId;
            });
        }
    }

    if (extendedCount > 0) {
        renderArchive();
        renderDraft();
        switchTab('draft-tab');
        render(window.allTrips);
    } else {
        alert("Не знайдено жодних підходящих вільних графіків для продовження існуючих кілець.");
    }
}

// ==========================================
// ФУНКЦІЯ ЗСУВУ ШАБЛОНУ ПО КОЛУ
// ==========================================
function shiftPattern() {
    const inputField = document.getElementById('pattern_input');
    const rawValue = inputField.value;

    // Якщо поле порожнє, нічого не робимо
    if (!rawValue) return;

    // Розбиваємо рядок по дефісу, прибираємо зайві пробіли і відкидаємо порожні елементи
    let nodes = rawValue.split('-').map(s => s.trim()).filter(Boolean);

    // Зсуваємо, тільки якщо є хоча б 2 міста
    if (nodes.length > 1) {
        // Беремо перше місто і видаляємо його з початку масиву
        let firstNode = nodes.shift();
        // Додаємо це місто в кінець масиву
        nodes.push(firstNode);
        
        // Збираємо масив назад у рядок і записуємо в поле
        inputField.value = nodes.join(' - ');
    }
}

// ==========================================
// ФУНКЦІЯ МАСОВОГО ВИДАЛЕННЯ ЗАТВЕРДЖЕНИХ КІЛЕЦЬ
// ==========================================
function deleteAllArchiveRings() {
    // 1. Отримуємо поточний пошуковий запит
    const searchTerm = document.getElementById('archive_search')?.value.toLowerCase().trim() || "";
    
    // 2. Формуємо текст попередження залежно від того, чи є фільтр
    const confirmMsg = searchTerm 
        ? `Ви впевнені, що хочете видалити всі ЗНАЙДЕНІ за запитом "${searchTerm}" наряди? Вони повернуться в реєстр.`
        : "УВАГА! Ви впевнені, що хочете видалити ВСІ затверджені наряди? Вони повернуться в реєстр.";
        
    if (!confirm(confirmMsg)) return;

    // 3. Збираємо всі затверджені кільця
    const archiveMap = {};
    window.allTrips.forEach(t => {
        if (t.ringId && t.ringId.startsWith('approved_')) {
            if (!archiveMap[t.ringId]) archiveMap[t.ringId] = [];
            archiveMap[t.ringId].push(t);
        }
    });

    const rings = Object.values(archiveMap);
    let ringsToDelete = [];
    
    // 4. Визначаємо, які саме ID треба видалити (з урахуванням пошуку)
    if (searchTerm) {
        rings.forEach((ring, idx) => {
            const rId = ring[0].ringId;
            const displayName = window.ringNamesMap[rId] ? window.ringNamesMap[rId] : `Затверджений наряд #${idx + 1}`;
            
            // Якщо ім'я кільця містить текст з пошуку — додаємо в список на видалення
            if (displayName.toLowerCase().includes(searchTerm)) {
                ringsToDelete.push(rId);
            }
        });
    } else {
        // Якщо пошуку немає, видаляємо всі
        ringsToDelete = Object.keys(archiveMap);
    }

    // 5. Очищаємо ringId для знайдених графіків
    window.allTrips.forEach(t => {
        if (ringsToDelete.includes(t.ringId)) {
            t.ringId = null;
        }
    });
    cleanOrphanedEmpties();
    // 6. Перемальовуємо інтерфейс
    renderArchive();
    
    // Якщо користувач раптом перемикається на реєстр, він теж має бути оновленим
    if (document.getElementById('register-tab').classList.contains('active')) {
        render(window.allTrips);
    }
}

// ==========================================
// ПЕРЕВІРКА ВІДСУТНІХ ВУЗЛІВ ПІСЛЯ ЗАВАНТАЖЕННЯ ФАЙЛУ
// ==========================================
function checkMissingNodes() {
    // Якщо довідник ще не завантажився або порожній — немає сенсу перевіряти
    if (!nodeDictionary || nodeDictionary.size === 0) return;

    const missingNodes = new Set();

    // Проходимо по всіх графіках і перевіряємо Відправника та Отримувача
    window.allTrips.forEach(t => {
        // Якщо поле не порожнє і його немає в довіднику — додаємо в Set
        if (t.origin && !nodeDictionary.has(t.origin)) {
            missingNodes.add(t.origin);
        }
        if (t.destination && !nodeDictionary.has(t.destination)) {
            missingNodes.add(t.destination);
        }
    });

    // Якщо знайшли хоча б один невідомий вузол — показуємо модалку
    if (missingNodes.size > 0) {
        const list = document.getElementById('missing-nodes-list');
        const modal = document.getElementById('missing-nodes-modal');
        
        // Сортуємо за алфавітом для краси і перетворюємо в HTML
        const sortedNodes = Array.from(missingNodes).sort();
        list.innerHTML = sortedNodes.map(node => `<li>${node}</li>`).join('');
        
        modal.style.display = 'flex';
    }
}

function closeMissingNodesModal() {
    document.getElementById('missing-nodes-modal').style.display = 'none';
}

// ==========================================
// ЛОГІКА ПІДРАХУНКУ БАЛАНСУ МАРШРУТІВ (САЙДБАР)
// ==========================================
let imbalanceSort = { col: 'diff', asc: false };

// Функція, яка спрацьовує при кліку на заголовок таблиці
function setImbalanceSort(col) {
    if (imbalanceSort.col === col) {
        imbalanceSort.asc = !imbalanceSort.asc; // Змінюємо напрямок, якщо клікнули на ту ж колонку
    } else {
        imbalanceSort.col = col;
        // За замовчуванням текст сортуємо від А до Я, а цифри - від більшого до меншого
        imbalanceSort.asc = (col === 'route'); 
    }
    updateDraftImbalanceStats();
}

function updateDraftImbalanceStats() {
    const content = document.getElementById('imbalance-content');
    if (!content || !window.allTrips) return;

    const autoVals = Array.from(activeFilters.auto);
    const typeVals = Array.from(activeFilters.type);
    
    // Отримуємо текст для пошуку по маршруту
    const searchTerm = (document.getElementById('sidebar_route_search')?.value || '').toLowerCase().trim();
    
    const imbalanceMap = {};
    
    window.allTrips.forEach(t => {
        if (t.ringId !== null) return; 
        
        // Глобальні фільтри
        if (autoVals.length > 0 && !autoVals.includes(t.auto)) return;
        if (typeVals.length > 0 && !typeVals.includes(t.type)) return;

        // Новий локальний фільтр по кнопках сайдбару
        if (sidebarActiveAutos.size > 0 && !sidebarActiveAutos.has(t.auto)) return;

        const origin = t.getPointName('origin', 'city') || 'Невідомо';
        const dest = t.getPointName('dest', 'city') || 'Невідомо';
        const autoType = t.auto || 'Невідомо'; // Беремо тип авто

        if (origin === dest) return; 

        let cityA = origin < dest ? origin : dest;
        let cityB = origin < dest ? dest : origin;
        
        // ГРУПУЄМО ПО МАРШРУТУ + ТИПУ АВТО
        let key = `${cityA}_${cityB}_${autoType}`;

        if (!imbalanceMap[key]) {
            imbalanceMap[key] = { cityA, cityB, auto: autoType, aToB: 0, bToA: 0 };
        }

        if (origin === cityA) {
            imbalanceMap[key].aToB++;
        } else {
            imbalanceMap[key].bToA++;
        }
    });

    let imbalances = Object.values(imbalanceMap).filter(item => item.aToB > 0 || item.bToA > 0);

    // Фільтрація по тексту маршруту
    if (searchTerm) {
        imbalances = imbalances.filter(item => {
            const route1 = `${item.cityA} - ${item.cityB}`.toLowerCase();
            const route2 = `${item.cityB} - ${item.cityA}`.toLowerCase();
            return route1.includes(searchTerm) || route2.includes(searchTerm);
        });
    }

    // НОВА ЛОГІКА СОРТУВАННЯ (Додано колонку 'auto')
    imbalances.sort((a, b) => {
        let valA, valB;
        if (imbalanceSort.col === 'route') {
            valA = a.cityA + a.cityB + a.auto;
            valB = b.cityA + b.cityB + b.auto;
            return imbalanceSort.asc ? valA.localeCompare(valB) : valB.localeCompare(valA);
        } else if (imbalanceSort.col === 'auto') {
            valA = a.auto + a.cityA + a.cityB;
            valB = b.auto + b.cityA + b.cityB;
            return imbalanceSort.asc ? valA.localeCompare(valB) : valB.localeCompare(valA);
        } else if (imbalanceSort.col === 'tu') {
            valA = a.aToB;
            valB = b.aToB;
        } else if (imbalanceSort.col === 'na') {
            valA = a.bToA;
            valB = b.bToA;
        } else { // 'diff' - сортування за перекосом
            valA = Math.abs(a.aToB - a.bToA);
            valB = Math.abs(b.aToB - b.bToA);
        }
        
        return imbalanceSort.asc ? (valA - valB) : (valB - valA);
    });

    if (imbalances.length === 0) {
        content.innerHTML = '<div class="empty-msg" style="font-size: 11px;">Вільних міжміських графіків не знайдено</div>';
        return;
    }

    const getSortIcon = (col) => imbalanceSort.col === col ? (imbalanceSort.asc ? ' ▲' : ' ▼') : '';

    let html = `
        <table class="imbalance-table">
            <thead>
                <tr>
                    <th class="sortable-th" style="text-align:left;" onclick="setImbalanceSort('route')" title="Сортувати за маршрутом">
                        Маршрут<span class="sort-icon">${getSortIcon('route')}</span>
                    </th>
                    <th class="sortable-th" style="text-align:center;" onclick="setImbalanceSort('auto')" title="Сортувати за авто">
                        Авто<span class="sort-icon">${getSortIcon('auto')}</span>
                    </th>
                    <th class="sortable-th" onclick="setImbalanceSort('tu')" title="Сортувати за кількістю туди">
                        ➡️<span class="sort-icon">${getSortIcon('tu')}</span>
                    </th>
                    <th class="sortable-th" onclick="setImbalanceSort('na')" title="Сортувати за кількістю назад">
                        ⬅️<span class="sort-icon">${getSortIcon('na')}</span>
                    </th>
                </tr>
            </thead>
            <tbody>
    `;

    imbalances.forEach(item => {
        let aToB_html = item.aToB > item.bToA ? `<span class="val-tu">${item.aToB}</span>` : item.aToB;
        let bToA_html = item.bToA > item.aToB ? `<span class="val-na">${item.bToA}</span>` : item.bToA;
        
        // Форматуємо тип авто (якщо дуже довгий - обріжеться завдяки CSS)
        let autoDisplay = item.auto.length > 8 ? item.auto.substring(0, 7) + '…' : item.auto;

        html += `
            <tr>
                <td class="city-col" title="${item.cityA} - ${item.cityB}">
                    ${item.cityA} - ${item.cityB}
                </td>
                <td style="text-align: center; font-size: 9.5px; color: #666; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${item.auto}">
                    ${autoDisplay}
                </td>
                <td>${aToB_html}</td>
                <td>${bToA_html}</td>
            </tr>
        `;
    });

    html += `</tbody></table>`;
    content.innerHTML = html;
}
// ==========================================
// ЛОГІКА БАЛАНСУВАННЯ ПОРОЖНІМИ ПЕРЕГОНАМИ
// ==========================================

window.transitMatrix = {}; // Ключ: "ВузолА_ВузолБ", Значення: хвилини
window.emptiesPriorities = {}; // Ключ: "Місто", Значення: ["Мамка1", "Мамка2"...]

// 1. Завантаження довідників (Виклич цю функцію поруч із loadDictionary() в кінці файлу)
async function loadBalancingData() {
    try {
        const [transitRes, emptiesRes] = await Promise.all([
            fetch(WEB_APP_URL + '?action=transit'),
            fetch(WEB_APP_URL + '?action=empties')
        ]);
        
        const transitData = await transitRes.json();
        transitData.forEach(item => {
            window.transitMatrix[`${item.originNode}_${item.destNode}`] = item.durationMins;
        });

        window.emptiesPriorities = await emptiesRes.json();
        console.log("Довідники балансування завантажено успішно.");
    } catch (e) {
        console.error("Помилка завантаження матриць балансування:", e);
    }
}
// ДОДАЙ ВИКЛИК: loadBalancingData(); в самий низ файлу поруч із loadDictionary();

// 2. Фабрика віртуальних перегонів
function createVirtualEmptyTrip(originNode, destNode, startMins, durationMins, autoType) {
    let dummyRow = new Array(55).fill("");
    dummyRow[2] = "EMPTY_" + Math.floor(Math.random() * 1000000); // GRF
    dummyRow[7] = "Порожній"; // Тип
    dummyRow[8] = autoType;   // Авто
    dummyRow[10] = `${originNode} - ${destNode}`; // Маршрут
    dummyRow[11] = originNode; // Відправник
    dummyRow[15] = destNode;   // Отримувач
    
    let shiftedMins = startMins - 720;
    
    // Захист: якщо час падає на понеділок до 12:00 (shiftedMins < 0), 
    // він належить до логістичної неділі попереднього тижня
    if (shiftedMins < 0) {
        shiftedMins += 10080; 
    }
    
    let dayIdx = Math.floor(shiftedMins / 1440) % 7;

    // Ставимо "+" у відповідні колонки сирого масиву (Астрономічні 16-22, Логістичні 23-29)
    dummyRow[16 + dayIdx] = "+"; 
    dummyRow[23 + dayIdx] = "+"; 

    let vTrip = new Trip(dummyRow);
    
    // Жорстко перезаписуємо час
    vTrip.trueStart = startMins;
    vTrip.trueEnd = startMins + durationMins;
    
    // Оновлюємо внутрішні хвилини, щоб спливаючі підказки при наведенні працювали коректно
    vTrip.podachaInt = startMins - 1;
    vTrip.depInt = startMins;
    vTrip.arrInt = startMins + durationMins;
    vTrip.freeInt = startMins + durationMins + 1;
    
    // Форматуємо час суворо як "ГГ:ХВ" (без днів тижня)
    const formatVirtualTime = (totalMins) => {
        // Додаємо 10080 (хвилин у тижні), щоб -1 хвилина коректно відобразилась як 23:59 попереднього дня
        let safeMins = (totalMins + 10080) % 10080; 
        let h = Math.floor((safeMins % 1440) / 60).toString().padStart(2, '0');
        let m = (Math.round(safeMins) % 60).toString().padStart(2, '0');
        return `${h}:${m}`;
    };

    vTrip.podachaStr = formatVirtualTime(vTrip.podachaInt);
    vTrip.depStr = formatVirtualTime(vTrip.depInt);
    vTrip.arrStr = formatVirtualTime(vTrip.arrInt);
    vTrip.freeStr = formatVirtualTime(vTrip.freeInt);
    
    vTrip.comment = "🤖 Авто-перегон";
    return vTrip;
}

function findLoadedConnection(lastTrip, availableTrips, reparkMins, mode) {
    let targetOrigin = lastTrip.getPointName('dest', mode);
    let effectiveLastEnd = lastTrip.trueEnd + (lastTrip.trueEnd < lastTrip.trueStart ? 10080 : 0);
    const isLastEmpty = String(lastTrip.type || '').trim().toLowerCase() === "порожній";

    let candidates = availableTrips.filter(c => {
        const isCandidateEmpty = String(c.type || '').trim().toLowerCase() === "порожній";
        if (isLastEmpty && isCandidateEmpty) return false;

        return c.ringId === null &&
               c.auto === lastTrip.auto &&
               c.getPointName('origin', mode) === targetOrigin &&
               c.trueStart >= (effectiveLastEnd + reparkMins) &&
               (c.trueStart - effectiveLastEnd) <= (24 * 60); // Максимум 24 год очікування
    });

    if (candidates.length > 0) {
        // Сортуємо хронологічно, щоб брати найближчий за часом
        candidates.sort((a, b) => a.trueStart - b.trueStart);
        return candidates[0]; 
    }
    return null;
}

// 3. Основне ядро балансування (повертає знайдений перегон + наступний графік або null)
function findEmptyConnection(lastTrip, availableTrips, reparkMins, returnToOrigin = false) {
    // Захист від двох порожніх підряд
    if (String(lastTrip.type || '').trim().toLowerCase() === "порожній") return null;

    let originCity = lastTrip.getPointName('dest', 'city');
    let originNode = lastTrip.destination; // Сирий вузол з довідника
    let effectiveLastEnd = lastTrip.trueEnd + (lastTrip.trueEnd < lastTrip.trueStart ? 10080 : 0);

    let priorities = [];
    if (returnToOrigin) {
        // РЕЖИМ "ДОДОМУ": Шукаємо наступний рейс з міста, з якого виїхав попередній
        priorities = [lastTrip.getPointName('origin', 'city')];
    } else {
        // РЕЖИМ МАТРИЦІ: Беремо пріоритети з довідника
        priorities = window.emptiesPriorities[originCity];
        if (!priorities || priorities.length === 0) return null;
    }

    // Перебираємо міста в порядку пріоритету
    for (let targetCity of priorities) {
        // Шукаємо потенційні наступні графіки
        let candidates = availableTrips.filter(c => {
            const isCandidateEmpty = String(c.type || '').trim().toLowerCase() === "порожній";
            return c.ringId === null && 
                   !isCandidateEmpty && 
                   c.auto === lastTrip.auto &&
                   c.getPointName('origin', 'city') === targetCity &&
                   c.trueStart > effectiveLastEnd; // Хоча б пізніше, ніж ми звільнились
        });

        // Сортуємо кандидатів хронологічно, щоб брати найближчі
        candidates.sort((a, b) => a.trueStart - b.trueStart);

        for (let candidate of candidates) {
            let destNode = candidate.origin;
            let transitKey = `${originNode}_${destNode}`;
            
            // Примусово робимо числом і ставимо 180 хвилин (3 год), якщо скрипт не зміг прочитати таблицю
            let transitTime = Number(window.transitMatrix[transitKey]);
            if (isNaN(transitTime) || transitTime <= 0) {
                transitTime = 180; 
            }

            if (transitTime > 0) {
                // Перевіряємо, чи встигаємо: Кінець попереднього + перепарковка + дорога + перепарковка <= Старт наступного
                let arrivalTime = effectiveLastEnd + reparkMins + transitTime;
                
                // Дозволяємо чекати наступного графіка не більше 24 годин (можеш змінити)
                let waitTime = candidate.trueStart - arrivalTime;
                
                if (arrivalTime + reparkMins <= candidate.trueStart && waitTime <= (24 * 60)) {
                    // БІНГО! Знайшли маршрут. Створюємо віртуальний перегон.
                    let virtualTrip = createVirtualEmptyTrip(originNode, destNode, effectiveLastEnd + reparkMins, transitTime, lastTrip.auto);
                    return { virtualTrip, nextTrip: candidate };
                }
            }
        }
    }
    return null;
}

// 4. ІНСТРУМЕНТ 1: Докільцювати затверджені кільця
function balanceApprovedRings() {
    const repark = parseInt(reparkInput.value) || 0;
    const mode = modeSelect.value;
    let extendedCount = 0;

    const archiveMap = {};
    window.allTrips.forEach(t => {
        if (t.ringId && t.ringId.startsWith('approved_')) {
            if (!archiveMap[t.ringId]) archiveMap[t.ringId] = [];
            archiveMap[t.ringId].push(t);
        }
    });

    let availableTrips = window.allTrips.filter(t => t.ringId === null);

    for (const [rId, ringTrips] of Object.entries(archiveMap)) {
        ringTrips.sort((a, b) => {
            if (a.logisticDay !== b.logisticDay) return a.logisticDay - b.logisticDay;
            return a.trueStart - b.trueStart;
        });

        let currentChain = [...ringTrips];
        let searching = true;
        let addedNew = false;

        while (searching) {
            let lastTrip = currentChain[currentChain.length - 1];
            
            // 1. Спочатку шукаємо вантаж
            let loadedNext = findLoadedConnection(lastTrip, availableTrips, repark, mode);

            if (loadedNext) {
                loadedNext.ringId = 'temp';
                currentChain.push(loadedNext);
                // Видаляємо з пулу доступних
                availableTrips = availableTrips.filter(t => t.id !== loadedNext.id);
                addedNew = true;
            } else {
                // 2. Потім шукаємо перегон (поки без галочки "додому", за матрицею)
                let connection = findEmptyConnection(lastTrip, availableTrips, repark, false); 
                
                if (connection) {
                    connection.virtualTrip.ringId = 'temp';
                    connection.nextTrip.ringId = 'temp';
                    currentChain.push(connection.virtualTrip, connection.nextTrip);
                    // Видаляємо наступний рейс з пулу доступних
                    availableTrips = availableTrips.filter(t => t.id !== connection.nextTrip.id);
                    addedNew = true;
                } else {
                    searching = false;
                }
            }
        }

        if (addedNew) {
            const newDraftId = `draft_bal_${Date.now()}_${extendedCount++}`;
            currentChain.forEach(t => {
                if (t.ringId !== 'temp') t.originalRingId = t.ringId; 
                t.ringId = newDraftId;
                if (t.grf.startsWith('EMPTY_') && !window.allTrips.includes(t)) {
                    window.allTrips.push(t);
                }
            });
        }
    }

    if (extendedCount > 0) {
        renderArchive();
        renderDraft();
        switchTab('draft-tab');
        render(window.allTrips);
    } else {
        alert("Не вдалося знайти продовжень (ані вантажних, ані порожніх) для існуючих кілець.");
    }
}

// 5. ІНСТРУМЕНТ 2: Склеїти залишки в нові кільця
async function balanceDraftTrips() {
    const repark = parseInt(reparkInput.value) || 0;
    const minTrips = parseInt(document.getElementById('min_trips')?.value || 4);
    const returnMode = document.getElementById('empty_return_mode')?.checked || false; 
    const mode = modeSelect.value; 
    let ringCounter = 0;

    let workingTrips = window.allTrips.filter(t => t.ringId === null);

    for (let i = 0; i < workingTrips.length; i++) {
        if (!isAlgoRunning) break; // Реакція на кнопку СТОП

        if (i % 50 === 0) await new Promise(resolve => setTimeout(resolve, 0));

        let anchor = workingTrips[i];
        if (anchor.ringId !== null) continue;

        let currentChain = [anchor];
        anchor.ringId = 'temp';
        let searching = true;

        while (searching && isAlgoRunning) {
            let lastTrip = currentChain[currentChain.length - 1];
            let freeTrips = workingTrips.filter(t => t.ringId === null);
            
            let loadedNext = findLoadedConnection(lastTrip, freeTrips, repark, mode);

            if (loadedNext) {
                loadedNext.ringId = 'temp';
                currentChain.push(loadedNext);
            } else {
                let connection = findEmptyConnection(lastTrip, freeTrips, repark, returnMode);

                if (connection) {
                    connection.virtualTrip.ringId = 'temp';
                    connection.nextTrip.ringId = 'temp';
                    currentChain.push(connection.virtualTrip, connection.nextTrip);
                } else {
                    searching = false;
                }
            }
        }

        if (currentChain.length >= minTrips) {
            const draftId = `draft_bal_new_${Date.now()}_${ringCounter++}`;
            currentChain.forEach(t => {
                t.ringId = draftId;
                if (t.grf.startsWith('EMPTY_') && !window.allTrips.includes(t)) {
                    window.allTrips.push(t);
                }
            });
        } else {
            currentChain.forEach(t => t.ringId = null);
        }
    }
}

function cleanOrphanedEmpties() {
    window.allTrips = window.allTrips.filter(t => {
        // Перевіряємо чи це наша пустишка і чи стала вона "вільною"
        if (t.grf && t.grf.startsWith('EMPTY_') && t.ringId === null) {
            // Захист: якщо ми прямо зараз редагуємо кільце в Конструкторі і пустишка лежить там - не чіпаємо!
            if (typeof constructorRing !== 'undefined' && constructorRing.includes(t)) {
                return true; 
            }
            // В усіх інших випадках - повністю видаляємо з пам'яті системи
            return false; 
        }
        return true;
    });
}

// ==========================================
// ПЕРЕГЛЯД ТА ОНОВЛЕННЯ МАТРИЦІ ТРАНЗИТІВ
// ==========================================

// 1. Примусове оновлення
async function forceUpdateBalancingData() {
    const btn = document.getElementById('btn_update_matrix');
    if (btn) btn.innerText = "⏳ Оновлення...";
    
    await loadBalancingData(); // Викликаємо нашу існуючу функцію
    
    if (btn) btn.innerText = "🔄 Оновити довідники";
    alert(`Довідники оновлено!\nТранзитних маршрутів: ${Object.keys(window.transitMatrix).length}\nМіст з пріоритетами: ${Object.keys(window.emptiesPriorities).length}`);
}

// 2. Логіка модального вікна та сортування
let matrixSort = { col: 'origin', asc: true };

function showTransitMatrix() {
    document.getElementById('matrix-modal').style.display = 'flex';
    renderTransitMatrix();
}

function closeMatrixModal() {
    document.getElementById('matrix-modal').style.display = 'none';
}

function setMatrixSort(col) {
    if (matrixSort.col === col) {
        matrixSort.asc = !matrixSort.asc;
    } else {
        matrixSort.col = col;
        matrixSort.asc = true;
    }
    renderTransitMatrix();
}

function renderTransitMatrix() {
    let data = [];
    // Перетворюємо об'єкт window.transitMatrix назад у масив для таблиці
    for (let key in window.transitMatrix) {
        let [origin, dest] = key.split('_');
        data.push({ origin, dest, time: window.transitMatrix[key] });
    }

    // Сортуємо
    data.sort((a, b) => {
        let valA = a[matrixSort.col];
        let valB = b[matrixSort.col];
        if (matrixSort.col === 'time') {
            return matrixSort.asc ? valA - valB : valB - valA;
        } else {
            return matrixSort.asc ? valA.localeCompare(valB) : valB.localeCompare(valA);
        }
    });

    const getIcon = (col) => matrixSort.col === col ? (matrixSort.asc ? ' ▲' : ' ▼') : '';

    let html = `
        <thead>
            <tr>
                <th class="sortable-th" style="text-align: left;" onclick="setMatrixSort('origin')">Відправник (Вузол)<span class="sort-icon">${getIcon('origin')}</span></th>
                <th class="sortable-th" style="text-align: left;" onclick="setMatrixSort('dest')">Отримувач (Вузол)<span class="sort-icon">${getIcon('dest')}</span></th>
                <th class="sortable-th" onclick="setMatrixSort('time')">Хвилини<span class="sort-icon">${getIcon('time')}</span></th>
                <th>ГГ:ХВ</th>
            </tr>
        </thead>
        <tbody>
    `;

    if (data.length === 0) {
        html += `<tr><td colspan="4" style="text-align:center; padding:20px;">Матриця порожня. Натисніть "Оновити довідники".</td></tr>`;
    } else {
        data.forEach(item => {
            let h = Math.floor(item.time / 60).toString().padStart(2, '0');
            let m = (item.time % 60).toString().padStart(2, '0');
            
            html += `<tr>
                <td style="text-align: left;">${item.origin}</td>
                <td style="text-align: left;">${item.dest}</td>
                <td>${item.time}</td>
                <td style="font-family: monospace; font-weight: bold; color: #1a73e8;">${h}:${m}</td>
            </tr>`;
        });
    }
    
    html += `</tbody>`;
    document.getElementById('transit-table').innerHTML = html;
}

// ==========================================
// ЛОКАЛЬНІ ФІЛЬТРИ ДЛЯ САЙДБАРУ БАЛАНСУ
// ==========================================
let sidebarActiveAutos = new Set();

function updateSidebarAutoButtons() {
    const container = document.getElementById('sidebar_auto_buttons');
    if (!container || !window.allTrips) return;

    // Збираємо всі унікальні типи авто з бази
    const uniqueAutos = new Set();
    window.allTrips.forEach(t => {
        if (t.auto) uniqueAutos.add(t.auto);
    });

    const sortedAutos = Array.from(uniqueAutos).sort();
    
    // Малюємо кнопки
    container.innerHTML = sortedAutos.map(auto => {
        const isActive = sidebarActiveAutos.has(auto) ? 'active' : '';
        return `<button class="sidebar-auto-btn ${isActive}" onclick="toggleSidebarAuto('${auto}')">${auto}</button>`;
    }).join('');
}

function toggleSidebarAuto(auto) {
    if (sidebarActiveAutos.has(auto)) {
        sidebarActiveAutos.delete(auto);
    } else {
        sidebarActiveAutos.add(auto);
    }
    updateSidebarAutoButtons();
    updateDraftImbalanceStats(); // Одразу перемальовуємо таблицю
}

window.addEventListener('DOMContentLoaded', () => {
    const notifyUrl = WEB_APP_URL + '?action=notifyVisit';
    
    // Добавили { mode: 'no-cors' }
    fetch(notifyUrl, { mode: 'no-cors' })
        .catch(e => console.log("Уведомление не ушло, ну и ладно"));
});

// ==========================================
// ЗВОРОТНІЙ ЗВ'ЯЗОК (FEEDBACK)
// ==========================================
function openFeedbackModal() {
    document.getElementById('feedback-text').value = ''; // Очищаем поле
    document.getElementById('feedback-modal').style.display = 'flex';
}

function closeFeedbackModal() {
    document.getElementById('feedback-modal').style.display = 'none';
}

function sendFeedback() {
    const textEl = document.getElementById('feedback-text');
    const btn = document.getElementById('btn-send-feedback');
    const text = textEl.value.trim();

    if (!text) {
        alert("Напишіть хоча б пару слів 😉");
        return;
    }

    // Блокируем кнопку, чтобы не нажали дважды
    btn.innerText = "⏳ Надсилаємо...";
    btn.disabled = true;

    // Формируем URL с текстом. encodeURIComponent нужен, чтобы пробелы и спецсимволы не сломали ссылку
    const url = WEB_APP_URL + '?action=feedback&text=' + encodeURIComponent(text);

    // Отправляем с флагом no-cors, чтобы избежать ошибок браузера
    fetch(url, { mode: 'no-cors' })
        .then(() => {
            alert("Дякуємо! Ваше побажання відправлено розробнику 🚀");
            closeFeedbackModal();
        })
        .catch(e => {
            alert("Ой, щось пішло не так. Перевірте інтернет та спробуйте пізніше.");
            console.error("Ошибка отправки фидбека:", e);
        })
        .finally(() => {
            // Возвращаем кнопку в исходное состояние
            btn.innerText = "🚀 Надіслати";
            btn.disabled = false;
        });
}

loadDictionary();
loadBalancingData();