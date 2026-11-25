/**
 * BACKEND - Logika aplikacji (Apps Script)
 */

function getProjectFolder_(folderName) {
 const normalize = name => String(name).trim().toLowerCase();
 const onlyFolder = (iterator, description) => {
   const folders = [];
   while (iterator.hasNext()) folders.push(iterator.next());
   if (folders.length !== 1) throw new Error(description + ': oczekiwano dokładnie jednego folderu, znaleziono ' + folders.length + '.');
   return folders[0];
 };
 const matchingFolders = (parent, name) => {
   const matches = [];
   const folders = parent.getFolders();
   while (folders.hasNext()) {
     const folder = folders.next();
     if (normalize(folder.getName()) === normalize(name)) matches.push(folder);
   }
   return matches;
 };
 const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
 if (!spreadsheet) throw new Error('Skrypt nie jest powiązany z arkuszem źródłowym.');
 const tableFolder = onlyFolder(DriveApp.getFileById(spreadsheet.getId()).getParents(), 'Folder arkusza źródłowego');
 if (normalize(tableFolder.getName()) !== 'tabela') throw new Error('Arkusz źródłowy musi znajdować się bezpośrednio w folderze Tabela.');
 const project = onlyFolder(tableFolder.getParents(), 'Folder nadrzędny Tabela');
 const tables = matchingFolders(project, 'Tabela');
 if (tables.length !== 1) throw new Error('Oczekiwano dokładnie jednego folderu Tabela, znaleziono ' + tables.length + '.');
 const matches = matchingFolders(project, folderName);
 if (matches.length !== 1) throw new Error('Oczekiwano dokładnie jednego folderu ' + folderName + ', znaleziono ' + matches.length + '.');
 return matches[0];
}

const BRAND_CONFIG = {
 projectName: "Aplikacja-FOREMAN",
 country: "PL",
 seaviaLogoUrl: "https://drive.google.com/thumbnail?id=1L4VgENQplibUvoIXim4tpkwvIj_UgTLO&sz=w400",
 applicationLogoUrl: "https://drive.google.com/thumbnail?id=1lxSTmQcvgVlUts-T_ARGJM63CFPRk7SX&sz=w400",
 applicationLogoAlt: "Logo aplikacji",
 clientLogoUrl: "https://drive.google.com/thumbnail?id=1lxSTmQcvgVlUts-T_ARGJM63CFPRk7SX&sz=w400",
 clientLogoAlt: "Candido Hermida",
 clientLogoFallback: "https://via.placeholder.com/150x40?text=CANDIDO+HERMIDA"
};

// Funkcja wymuszająca uprawnienia do Google Drive podczas autoryzacji
function forceDrivePermission() {
 try { DriveApp.getFileById("12345_test_id"); } catch (e) { /* Ignoruj */ }
}

function doGet() {
 const template = HtmlService.createTemplateFromFile('Index');
 const brandConfig = getBrandConfig();
 template.brandConfig = brandConfig;

 return template
 .evaluate()
 .setTitle('Aplikacja-FOREMAN-dev')
 .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
 .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getBrandConfig() {
 const savedConfig = PropertiesService.getScriptProperties().getProperty('brandConfig');
 if (!savedConfig) return BRAND_CONFIG;

 try {
   const config = Object.assign({}, BRAND_CONFIG, JSON.parse(savedConfig));
   config.seaviaLogoUrl = normalizeDriveImageUrl(config.seaviaLogoUrl);
  config.applicationLogoUrl = normalizeDriveImageUrl(config.applicationLogoUrl);
   config.clientLogoUrl = normalizeDriveImageUrl(config.clientLogoUrl);
   return config;
 } catch (error) {
   return BRAND_CONFIG;
 }
}

function saveBrandConfig(config) {
 if (!config || typeof config !== 'object') throw new Error('Nieprawidłowe dane konfiguracji.');

 const cleanConfig = {
   projectName: validateBrandText(config.projectName, 'Nazwa projektu', 80),
   country: validateBrandText(config.country, 'Kraj', 2).toUpperCase(),
  seaviaLogoUrl: normalizeDriveImageUrl(validateBrandUrl(config.seaviaLogoUrl, 'Adres logo Seavia')),
  applicationLogoUrl: normalizeDriveImageUrl(validateBrandUrl(config.applicationLogoUrl, 'Adres logo aplikacji')),
  applicationLogoAlt: validateBrandText(config.applicationLogoAlt, 'Opis logo aplikacji', 80),
  clientLogoUrl: normalizeDriveImageUrl(validateBrandUrl(config.clientLogoUrl, 'Adres logo klienta')),
   clientLogoAlt: validateBrandText(config.clientLogoAlt, 'Nazwa klienta', 80),
   clientLogoFallback: BRAND_CONFIG.clientLogoFallback
 };

 PropertiesService.getScriptProperties().setProperty('brandConfig', JSON.stringify(cleanConfig));
 return { success: true };
}

function validateBrandText(value, label, maxLength) {
 const text = String(value || '').trim();
 if (!text || text.length > maxLength) throw new Error(label + ' jest wymagane i może mieć maksymalnie ' + maxLength + ' znaków.');
 return text;
}

function validateBrandUrl(value, label) {
 const url = String(value || '').trim();
 if (!/^(https?:\/\/\S+|[\w-]{20,})$/i.test(url)) throw new Error(label + ' musi być linkiem lub identyfikatorem pliku Google Drive.');
 return url;
}

function normalizeDriveImageUrl(value) {
 const url = String(value || '').trim();
 const fileIdMatch = url.match(/(?:\/d\/|[?&]id=)([-\w]{20,})/i);
 const fileId = fileIdMatch ? fileIdMatch[1] : (/^[-\w]{20,}$/.test(url) ? url : '');
 return fileId ? 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w400' : url;
}

function include(filename) {
 return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function normalizeTechnicianStatus(value) {
 const status = String(value || "").trim().toLowerCase();
 if (!status) return "inactive";

 const normalized = status.replace(/[^a-z]/g, '');
 if (["aktywny", "active", "aktywna", "working", "online", "wpracy", "pracuje"].includes(status)) return "active";
 if (["nieaktywny", "inactive", "nieaktywna", "offline", "off", "disabled", "nieaktywna"].includes(status)) return "inactive";
 if (["nieaktywny", "inactive", "offline", "disabled"].some(v => normalized.includes(v.replace(/[^a-z]/g, '')))) return "inactive";
 if (["aktywny", "active", "aktywna", "working", "online", "wpracy", "pracuje"].some(v => normalized.includes(v.replace(/[^a-z]/g, '')))) return "active";
 return "inactive";
}

function getTechnicianStatsFromSheet() {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const dataSheet = ss.getSheetByName('Data sheet');
 const result = { technicians: [], techniciansDetails: [], techStats: { total: 0, active: 0 } };

 if (!dataSheet) return result;

 const lastRow = dataSheet.getLastRow();
 if (lastRow < 9) return result;

 const maxCols = Math.max(3, dataSheet.getLastColumn());
 const techData = dataSheet.getRange(9, 1, lastRow - 8, maxCols).getValues();
 const seenNames = new Set();

 techData.forEach(row => {
   const team = String(row[0] || "").trim();
   const name = String(row[1] || "").trim();
   const status = normalizeTechnicianStatus(row[2]);

   if (!name) return;
   const normalizedName = name.toLowerCase();
   if (seenNames.has(normalizedName)) return;

   seenNames.add(normalizedName);
   result.technicians.push(name);
   result.techniciansDetails.push({
     name: name,
     team: team,
     status: status,
     hoursByDeck: {},
     totalHours: 0
   });

   result.techStats.total++;
   if (status === "active") result.techStats.active++;
 });

 return result;
}

/**
 * Pobiera początkowe dane do dashboardu (KPI + Wykresy)
 * Scalona wersja obsługująca zarówno spisHours, jak i extraJobs!
 */
function getInitialData() {
 const ss = SpreadsheetApp.getActiveSpreadsheet();

 const techData = getTechnicianStatsFromSheet();
 let technicians = techData.technicians;
 let techniciansDetails = techData.techniciansDetails;
 let techStats = techData.techStats;

 let totalWorkedHours = 0;
 let workedHoursByTask = {};
 let spisHours = []; 

let extraJobs = {
 totalHours: 0,
 byDeck: {},
 byArea: {}
 };

 const sheetSpis = ss.getSheetByName('Spis wykonanych prac');
 if (sheetSpis && sheetSpis.getLastRow() >= 2) {
 const spisData = sheetSpis.getRange(2, 2, sheetSpis.getLastRow() - 1, 7).getDisplayValues(); 
spisData.forEach(row => {
 let timeStart = String(row[0]).trim(); // B
 let timeEnd = String(row[1]).trim(); // C

 let deck = normalizeDeckName(row[2]); // D
 let area = String(row[3]).trim(); // E
 let taskName = String(row[4]).trim(); // F
 let taskId = String(row[5]).trim(); // G 
let isCancelled = String(row[4]).includes("[ANULOWANE]") || taskId.includes("[ANULOWANE]");
 let techNameFromSpis = String(row[6]).trim(); // H

 if (timeStart && timeEnd && !isCancelled) {
 let duration = parseWorkTimeBackend(timeEnd) - parseWorkTimeBackend(timeStart);
 if (duration < 0) duration += 24; 

duration = Math.round(duration * 100) / 100;
 if (duration > 0) {
 if (taskId.toLowerCase().includes("- extra job")) {
 extraJobs.totalHours += duration;
 if (!extraJobs.byDeck[deck]) extraJobs.byDeck[deck] = 0;
 extraJobs.byDeck[deck] += duration;
 if (!extraJobs.byArea[area]) extraJobs.byArea[area] = 0;
 extraJobs.byArea[area] += duration;
 }

 totalWorkedHours += duration;
 let key = deck + "_" + taskName;
 if(!workedHoursByTask[key]) workedHoursByTask[key] = 0;
 workedHoursByTask[key] += duration;

 spisHours.push({
 hours: duration,
 deck: deck,
 area: area,
 idPrac: taskId, 
taskName: taskName
 });

 let techObj = techniciansDetails.find(t => t.name.toLowerCase() === techNameFromSpis.toLowerCase());
 if (techObj) {
 techObj.totalHours += duration;
 if (!techObj.hoursByDeck[deck]) techObj.hoursByDeck[deck] = 0;
 techObj.hoursByDeck[deck] += duration;
 }
 }
 }
 });
 }

 totalWorkedHours = Math.round(totalWorkedHours * 10) / 10;
 extraJobs.totalHours = Math.round(extraJobs.totalHours * 10) / 10;
 Object.keys(extraJobs.byDeck).forEach(k => {
 extraJobs.byDeck[k] = Math.round(extraJobs.byDeck[k] * 10) / 10;
 });
 Object.keys(extraJobs.byArea).forEach(k => {
 extraJobs.byArea[k] = Math.round(extraJobs.byArea[k] * 10) / 10;
 });
 techniciansDetails.forEach(tech => {
 tech.totalHours = Math.round(tech.totalHours * 10) / 10;
 Object.keys(tech.hoursByDeck).forEach(d => {
 tech.hoursByDeck[d] = Math.round(tech.hoursByDeck[d] * 10) / 10;
 });
 });

 let tasksTree = {}; 
const sheetCandido = ss.getSheetByName('Lista zadań');
 if (sheetCandido) {
 const lastRow = sheetCandido.getLastRow();

 if (lastRow >= 3) {
 // ZMIANA: Rozszerzamy zakres do 20 kolumn, aby objąć nowe kolumny (R, S, T)
 const data = sheetCandido.getRange(1, 1, lastRow, 20).getValues();
 // Wartości wyświetlane używane dla kolumn ID/typ/nazwa zadania, żeby komórka
 // błędnie rozpoznana przez Sheets jako data pokazywała się jak w arkuszu,
 // a nie jako surowy obiekt Date (np. "Tue Aug 25 2026 09:00:00 GMT+0200...").
 const displayData = sheetCandido.getRange(1, 1, lastRow, 20).getDisplayValues();
 let currentDeck = "";
 let currentArea = "";

 for (let i = 2; i < data.length; i++) {
 let deckVal = normalizeDeckName(data[i][2]);
 if (deckVal === "" && String(data[i][1] || "").toLowerCase().includes("dk")) {
 deckVal = normalizeDeckName(data[i][1]);
 }
 let areaVal = String(data[i][3] || "").trim();
 let roomVal = String(data[i][4] || "").trim(); 
 let taskType = String(displayData[i][7] || "").trim(); // H - typ zadania

 let taskName = String(displayData[i][9] || "").trim(); // J
 if (taskName === "") {
 let taskH = String(displayData[i][8] || "").trim();
 let taskJ = String(displayData[i][10] || "").trim();
 if (taskH !== "" && !taskH.toLowerCase().includes("seavia")) taskName = taskH;
 else if (taskJ !== "") taskName = taskJ;
 }
 if (!taskName && taskType) taskName = taskType;
 let plannedStart = data[i][12];
 let plannedEnd = data[i][13];
 let plannedDays = 1;
 if (plannedStart instanceof Date && plannedEnd instanceof Date) {
 let diffTime = Math.abs(plannedEnd - plannedStart);
 plannedDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
 }

 let rawProgress = data[i][16]; // Q
 let progressNum = 0;
 if (typeof rawProgress === 'number') {
 progressNum = (rawProgress > 0 && rawProgress <= 1) ? Math.round(rawProgress * 100) : Math.round(rawProgress);
 } else if (typeof rawProgress === 'string') {
 progressNum = parseInt(rawProgress.replace('%', '').trim()) || 0;
 }

 // NOWE: Odczyt Jednostki i Ilości/Powierzchni
 let taskUnit = String(displayData[i][5] || "").trim().toLowerCase(); // F
 let taskAmountRaw = String(displayData[i][17] || "").replace(',', '.').replace(/[^0-9.-]+/g,""); // R
 let taskAmount = parseFloat(taskAmountRaw) || 0;

 if (deckVal !== "") {
 currentDeck = deckVal; currentArea = ""; 
if (!tasksTree[currentDeck]) tasksTree[currentDeck] = {};
 }
 if (areaVal !== "") currentArea = areaVal;

 if (currentDeck === "" || taskName === "") continue;
 let area = currentArea === "" ? "Brak strefy" : currentArea;

 if (!tasksTree[currentDeck]) tasksTree[currentDeck] = {}; 
if (!tasksTree[currentDeck][area]) tasksTree[currentDeck][area] = [];

 let taskId = String(displayData[i][6] || displayData[i][1] || "").trim();
let key = currentDeck + "_" + taskName;
 let taskHours = workedHoursByTask[key] || 0;
 taskHours = Math.round(taskHours * 10) / 10;

 tasksTree[currentDeck][area].push({
 row: i + 1,
 roomId: roomVal, 
 taskId: taskId,
 taskName: taskName,
 taskType: taskType,
 progress: progressNum,
 plannedDays: plannedDays,
 workedHours: taskHours,
 unit: taskUnit, // Przekazujemy do frontendu
 amount: taskAmount, // Przekazujemy do frontendu (WAGA)
 });
 }
 }
 }

 const rooms = getRoomCatalog();

 return {
 technicians: technicians,
 techniciansDetails: techniciansDetails,
 techStats: techStats, 
totalWorkedHours: totalWorkedHours, 
tasksTree: tasksTree,
 spisHours: spisHours, 
extraJobs: extraJobs,
 rooms: rooms 
};
}

function normalizeDeckName(value) {
 const raw = String(value ?? '').trim();
 if (!raw) return '';
 const upper = raw.toUpperCase();
 const compact = upper.replace(/[_\-\s]+/g, '');
 const match = compact.match(/^DK(\d+)$/i);
 if (match) return 'DK' + match[1];
 return upper;
}

function normalizeAreaCategory(value) {
  const raw = String(value ?? '').trim();
  if (!raw) return '';
  const key = raw.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z]/g, '');

  if (key.includes('sciany') || key.includes('sciana')) return 'ŚCIANY';
  if (key.includes('podloga') || key.includes('podlogi')) return 'PODŁOGI';
  if (key.includes('sufit') || key.includes('sufity')) return 'SUFITY';
  if (key.includes('meble')) return 'MEBLE';
  return raw;
}

function normalizeHeaderName(value) {
  const cleaned = String(value ?? '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[łŁ]/g, 'l')
    .replace(/[ąĄ]/g, 'a')
    .replace(/[ćĆ]/g, 'c')
    .replace(/[ęĘ]/g, 'e')
    .replace(/[óÓ]/g, 'o')
    .replace(/[śŚ]/g, 's')
    .replace(/[żŻ]/g, 'z')
    .replace(/[źŹ]/g, 'z')
    .replace(/\b(m|m2|sqm|m\^2|sqm2)\b/g, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  return cleaned;
}

function getHeaderIndex(headers, ...aliases) {
  const wanted = aliases
    .flatMap(alias => [normalizeHeaderName(alias)])
    .filter(Boolean);

  if (!wanted.length) return -1;

  for (const alias of wanted) {
    const exact = headers.indexOf(alias);
    if (exact >= 0) return exact;
  }

  for (let i = 0; i < headers.length; i++) {
    const header = normalizeHeaderName(headers[i]);
    if (wanted.some(alias => header === alias || header.includes(alias) || alias.includes(header))) {
      return i;
    }
  }

  return -1;
}

function parseWorkTimeBackend(val) {
 if (!val) return 0;
 if (val instanceof Date) return val.getHours() + (val.getMinutes() / 60);
 let str = String(val).replace(',', '.').trim();
 if (str.includes(':')) {
 let parts = str.split(':');
 return parseInt(parts[0]||0) + (parseInt(parts[1]||0)/60);
 }
 return parseFloat(str) || 0;
}

function normalizeRoomMatch(value) {
  return String(value ?? '').trim().toLowerCase().replace(/\s+/g, '');
}

function getRoomSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = ['Pomieszczenia', 'Pokoje', 'Rooms', 'Room Data'];

  for (const name of sheetNames) {
    const candidate = ss.getSheetByName(name);
    if (candidate) return candidate;
  }

  if (ss.getSheetByName) {
    const fallback = ss.getSheetByName('Pomieszczenia');
    if (fallback) return fallback;
  }

  if (ss.insertSheet) return ss.insertSheet('Pomieszczenia');
  return null;
}

function getRoomCatalog() {
  try {
    const sheet = getRoomSheet();
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const values = sheet.getRange(1, 1, lastRow, 20).getValues();
    const headers = values[0].map(h => normalizeHeaderName(h));

    return values.slice(1).map((row, index) => {
      const deckIdx = getHeaderIndex(headers, 'poklad', 'deck');
      const areaIdx = getHeaderIndex(headers, 'strefa', 'area');
      const nameIdx = getHeaderIndex(headers, 'nazwa pokoju', 'nazwa', 'room name');
      const roomIdx = getHeaderIndex(headers, 'numer pokoju', 'numer', 'room number');
      const widthIdx = getHeaderIndex(headers, 'wymiar a', 'wymiar a m', 'width');
      const lengthIdx = getHeaderIndex(headers, 'wymiar b', 'wymiar b m', 'length');
      const heightIdx = getHeaderIndex(headers, 'wysokosc h', 'wysokosc', 'height');
      const floorIdx = getHeaderIndex(headers, 'powierzchnia podlogi', 'podloga', 'floor area');
      const ceilingIdx = getHeaderIndex(headers, 'powierzchnia sufitu', 'sufit', 'ceiling area');
      const wallIdx = getHeaderIndex(headers, 'powierzchnia scian', 'sciany', 'wall area');

      const width = parseFloat(String(row[widthIdx] ?? '').replace(',', '.')) || 0;
      const length = parseFloat(String(row[lengthIdx] ?? '').replace(',', '.')) || 0;
      const height = parseFloat(String(row[heightIdx] ?? '').replace(',', '.')) || 0;
      const floorSurface = parseFloat(String(row[floorIdx] ?? '').replace(',', '.')) || (width && length ? width * length : 0);
      const ceilingSurface = parseFloat(String(row[ceilingIdx] ?? '').replace(',', '.')) || floorSurface;
      const wallSurface = parseFloat(String(row[wallIdx] ?? '').replace(',', '.')) || ((width && length && height) ? (2 * (width + length)) * height : 0);

      const record = {
        id: index + 2,
        deck: deckIdx >= 0 ? String(row[deckIdx] || '') : (row[0] ? String(row[0]) : ''),
        area: areaIdx >= 0 ? String(row[areaIdx] || '') : '',
        name: nameIdx >= 0 ? String(row[nameIdx] || '') : '',
        room: roomIdx >= 0 ? String(row[roomIdx] || '') : (row[1] ? String(row[1]) : ''),
        width: width,
        length: length,
        height: height,
        surface: floorSurface,
        ceilingSurface: ceilingSurface,
        wallSurface: wallSurface
      };

      if (!record.name && !record.room) return null;
      return record;
    }).filter(Boolean).filter(row => String(row.deck || '').trim() || String(row.name || '').trim() || String(row.room || '').trim());
  } catch (error) {
    console.error('Błąd pobierania katalogu pokoi:', error);
    return [];
  }
}

function getRoomSurface(deck, area, room) {
  try {
    const roomSheet = getRoomSheet();
    if (!roomSheet) return 0;

    const lastRow = roomSheet.getLastRow();
    if (lastRow < 2) return 0;

    const data = roomSheet.getRange(1, 1, lastRow, 11).getValues();
    const headers = data[0].map(h => normalizeHeaderName(h));

    const deckIdx = getHeaderIndex(headers, 'poklad', 'deck');
    const roomIdx = getHeaderIndex(headers, 'numer pokoju', 'numer', 'room number');
    const nameIdx = getHeaderIndex(headers, 'nazwa pokoju', 'nazwa', 'room name');
    const widthIdx = getHeaderIndex(headers, 'wymiar a', 'wymiar a m', 'width');
    const lengthIdx = getHeaderIndex(headers, 'wymiar b', 'wymiar b m', 'length');
    const heightIdx = getHeaderIndex(headers, 'wysokosc h', 'wysokosc', 'height');
    const floorIdx = getHeaderIndex(headers, 'powierzchnia podlogi', 'podloga', 'floor area');
    const ceilingIdx = getHeaderIndex(headers, 'powierzchnia sufitu', 'sufit', 'ceiling area');
    const wallIdx = getHeaderIndex(headers, 'powierzchnia scian', 'sciany', 'wall area');

    const targetDeck = normalizeDeckName(deck);
    const targetRoom = String(room || '').trim();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowDeck = normalizeDeckName(row[deckIdx] ?? '');
      const rowName = String(row[nameIdx] ?? '').trim();
      const rowRoom = String(row[roomIdx] ?? '').trim();

      if (!rowDeck && !rowName && !rowRoom) continue;
      if (targetDeck && rowDeck && rowDeck !== targetDeck) continue;
      if (targetRoom) {
        const roomCandidates = [rowRoom, rowName].filter(Boolean).map(v => normalizeRoomMatch(v));
        if (!roomCandidates.includes(normalizeRoomMatch(targetRoom))) continue;
      }

      const width = parseFloat(String(row[widthIdx] ?? '').replace(',', '.')) || 0;
      const length = parseFloat(String(row[lengthIdx] ?? '').replace(',', '.')) || 0;
      const height = parseFloat(String(row[heightIdx] ?? '').replace(',', '.')) || 0;
      const floor = parseFloat(String(row[floorIdx] ?? '').replace(',', '.')) || (width && length ? width * length : 0);
      const ceiling = parseFloat(String(row[ceilingIdx] ?? '').replace(',', '.')) || floor;
      const wall = parseFloat(String(row[wallIdx] ?? '').replace(',', '.')) || ((width && length && height) ? (2 * (width + length)) * height : 0);

      if (floor) return Math.round(floor * 100) / 100;
      if (ceiling) return Math.round(ceiling * 100) / 100;
      if (wall) return Math.round(wall * 100) / 100;
    }

    return 0;
  } catch (error) {
    console.error('Błąd odczytu powierzchni pokoju:', error);
    return 0;
  }
}

function saveRoomRecord(payload) {
  try {
    if (!payload) return { success: false, message: 'Brak danych pokoju.' };

    const deck = String(payload.deck || '').trim();
    const roomArea = String(payload.area || '').trim();
    const roomName = String(payload.name || '').trim();
    const roomNumber = String(payload.room || '').trim();
    if (!deck || (!roomName && !roomNumber)) {
      return { success: false, message: 'Pokład oraz nazwa lub numer pokoju są wymagane.' };
    }

    const sheet = getRoomSheet();
    const headerRow = sheet.getRange(1, 1, 1, 20).getValues()[0];
    const headers = headerRow.map(h => normalizeHeaderName(h));
    const hasAreaColumn = getHeaderIndex(headers, 'strefa', 'area') >= 0;
    const standardHeaders = ['Pokład', 'Strefa (Area)', 'Nazwa pokoju', 'Numer pokoju', 'Wymiar A (m)', 'Wymiar B (m)', 'Wysokość H (m)', 'Powierzchnia podłogi', 'Powierzchnia sufitu', 'Powierzchnia ścian'];

    if (sheet.getLastRow() < 1) {
      sheet.appendRow(standardHeaders);
    }

    const recordId = Number(payload.id || 0);
    const existingRow = findRoomInSheet(sheet, deck, roomName, roomNumber, recordId);
    if (existingRow && !recordId) {
      return { success: false, message: 'Na tym pokładzie istnieje już pomieszczenie o tej nazwie. Zmień nazwę, aby zapisać nowe wymiary.' };
    }
    const rowIndex = existingRow || (sheet.getLastRow() + 1);

    const width = Number(payload.width || 0);
    const length = Number(payload.length || 0);
    const height = Number(payload.height || 0);
    const floor = Number(payload.surface || 0) || (width && length ? width * length : 0);
    const ceiling = Number(payload.ceilingSurface || 0) || floor;
    const wall = Number(payload.wallSurface || 0) || ((width && length && height) ? (2 * (width + length)) * height : 0);

    const values = hasAreaColumn ? [
      normalizeDeckName(deck),
      roomArea || '',
      roomName || '',
      roomNumber || '',
      width || '',
      length || '',
      height || '',
      floor || '',
      ceiling || '',
      wall || ''
    ] : [
      normalizeDeckName(deck),
      roomName || '',
      roomNumber || '',
      width || '',
      length || '',
      height || '',
      floor || '',
      ceiling || '',
      wall || ''
    ];

    sheet.getRange(rowIndex, 1, 1, values.length).setValues([values]);
    const syncedTasks = syncRoomTasks_(deck, roomName, roomNumber);

    return { success: true, message: 'Pokój zapisany poprawnie. Zaktualizowano ilości w ' + syncedTasks + ' zadaniach.' };
  } catch (error) {
    console.error('Błąd zapisu pokoju:', error);
    return { success: false, message: 'Błąd zapisu pokoju: ' + error.toString() };
  }
}

function findRoomInSheet(sheet, deck, roomName, roomNumber, idOverride) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const values = sheet.getRange(1, 1, lastRow, 20).getValues();
  const headers = values[0].map(h => normalizeHeaderName(h));
  const deckIdx = getHeaderIndex(headers, 'poklad', 'deck');
  const nameIdx = getHeaderIndex(headers, 'nazwa pokoju', 'nazwa', 'room name');
  const roomIdx = getHeaderIndex(headers, 'numer pokoju', 'numer', 'room number');

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowId = i + 1;
    if (idOverride && rowId === idOverride) return rowId;

    const rowDeck = normalizeDeckName(row[deckIdx] ?? '');
    const rowName = String(row[nameIdx] ?? '').trim();
    const rowRoom = String(row[roomIdx] ?? '').trim();
    const targetName = normalizeRoomMatch(roomName);
    const targetRoom = normalizeRoomMatch(roomNumber);

    if (rowDeck === normalizeDeckName(deck) && ((targetName && normalizeRoomMatch(rowName) === targetName) || (targetRoom && normalizeRoomMatch(rowRoom) === targetRoom) || (!targetName && !targetRoom))) {
      return rowId;
    }
  }

  return 0;
}

function deleteRoomRecord(id) {
  try {
    const sheet = getRoomSheet();
    if (!sheet || !id) return { success: false, message: 'Brak identyfikatora pokoju.' };

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'Brak danych pokoju.' };

    const values = sheet.getRange(1, 1, lastRow, 7).getValues();
    for (let i = 1; i < values.length; i++) {
      if (i + 1 === Number(id)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Pokój usunięty.' };
      }
    }

    return { success: false, message: 'Nie znaleziono pokoju do usunięcia.' };
  } catch (error) {
    console.error('Błąd usuwania pokoju:', error);
    return { success: false, message: 'Błąd usuwania pokoju: ' + error.toString() };
  }
}

function getRoomsData() {
  return { rooms: getRoomCatalog() };
}

function addRoomToTaskForm(deck, area, room) {
  return getTaskSurfaceForTask(deck, area, room, '');
}

function getTaskSurfaceForTask(deck, area, room, taskName) {
  const targetDeck = normalizeDeckName(deck);
  const targetArea = String(area || '').trim();
  const targetRoom = String(room || '').trim();
  const candidates = getRoomCatalog().filter(row => {
    if (targetDeck && normalizeDeckName(row.deck) && normalizeDeckName(row.deck) !== targetDeck) return false;
    if (targetArea && String(row.area || '').trim() && String(row.area || '').trim() !== targetArea) return false;
    if (targetRoom) {
      const roomMatches = [String(row.room || '').trim(), String(row.name || '').trim()].filter(Boolean).map(v => normalizeRoomMatch(v));
      return roomMatches.includes(normalizeRoomMatch(targetRoom));
    }
    return true;
  });

  if (candidates.length === 0) {
    return 0;
  }

  const bestMatch = candidates.sort((a, b) => {
    const scoreA = (
      (String(a.area || '').trim() === targetArea ? 4 : 0) +
      (String(a.room || '').trim() === targetRoom ? 3 : 0) +
      (String(a.name || '').trim() === String(taskName || '').trim() ? 2 : 0)
    );
    const scoreB = (
      (String(b.area || '').trim() === targetArea ? 4 : 0) +
      (String(b.room || '').trim() === targetRoom ? 3 : 0) +
      (String(b.name || '').trim() === String(taskName || '').trim() ? 2 : 0)
    );
    return scoreB - scoreA;
  })[0];

  const surface = Number(bestMatch.surface || 0) || (Number(bestMatch.width || 0) * Number(bestMatch.length || 0));
  return Math.round((surface || 0) * 100) / 100;
}

function resolveTaskSurfaceForRow(deck, area, room, taskName, taskType) {
  const text = [area || '', taskType || '', taskName || ''].join(' ').toLowerCase();
  const targetDeck = normalizeDeckName(deck);
  const targetArea = String(area || '').trim();
  const targetRoom = String(room || '').trim();
  const roomMatches = getRoomCatalog().filter(row => {
    if (targetDeck && normalizeDeckName(row.deck) && normalizeDeckName(row.deck) !== targetDeck) return false;
    if (targetArea && String(row.area || '').trim() && String(row.area || '').trim() !== targetArea) return false;
    if (!targetRoom) return true;
    return [String(row.room || '').trim(), String(row.name || '').trim()].filter(Boolean).map(v => normalizeRoomMatch(v)).includes(normalizeRoomMatch(targetRoom));
  });

  const match = roomMatches[0] || null;
  if (!match) return getTaskSurfaceForTask(targetDeck, targetArea, targetRoom, taskName || '');

  if (/podłog|wykładzin|wylew|szlif|posadz|płyt|laminat|frez|terak/.test(text)) {
    return Math.round((Number(match.surface || match.floorSurface || 0) || 0) * 100) / 100;
  }
  if (/ścian|scian|malowani|tapet|grunt|gładzi|farb|pobielen|tynk|siding/.test(text)) {
    return Math.round((Number(match.wallSurface || 0) || 0) * 100) / 100;
  }
  if (/sufit|strop|sufitu|gładzi/.test(text)) {
    return Math.round((Number(match.ceilingSurface || 0) || 0) * 100) / 100;
  }

  const fallback = Number(match.surface || 0) || (Number(match.width || 0) * Number(match.length || 0));
  return Math.round((fallback || getTaskSurfaceForTask(targetDeck, targetArea, targetRoom, taskName || '')) * 100) / 100;
}

function syncTaskSurfaceForRow(rowIndex, deck, area, room, taskName, taskType) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Lista zadań');
    const rowNum = Number(rowIndex);
    if (!sheet || !Number.isFinite(rowNum) || rowNum < 3) return null;

    const resolved = resolveTaskSurfaceForRow(deck, area, room, taskName, taskType);
    sheet.getRange(rowNum, 18).setValue(resolved);
    return resolved;
  } catch (err) {
    console.error('Błąd aktualizacji powierzchni zadania:', err);
    return null;
  }
}

function parseTimeToMinutes(value) {
 if (!value) return null;
 const text = String(value).trim();
 if (!text) return null;
 const match = text.match(/(\d{1,2}):(\d{2})/);
 if (!match) return null;
 const hours = parseInt(match[1], 10);
 const minutes = parseInt(match[2], 10);
 if (isNaN(hours) || isNaN(minutes)) return null;
 return hours * 60 + minutes;
}

function normalizeTaskTypeForSheet(value) {
  const raw = String(value ?? '').trim();
  if (!raw) return '';
  const upper = raw.toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  const compact = upper.replace(/\s+/g, ' ').trim();
  const map = {
    'ROZBIORKA': 'ROZBIÓRKA',
    'MONTAZ': 'MONTAŻ',
    'ROZBIORKA Z ODZYSKIEM': 'ROZBIÓRKA Z ODZYSKIEM',
    'MONTAZ Z ODZYSKU': 'MONTAŻ Z ODZYSKU',
    'MONTAZ Z ODZYSKIEM': 'MONTAŻ Z ODZYSKU'
  };
  return map[compact] || compact;
}

function getCurrentUserLogin() {
  try {
    const user = Session.getActiveUser();
    if (!user) return 'system';
    if (user.getEmail && user.getEmail()) return user.getEmail();
    if (user.getUserLoginId && user.getUserLoginId()) return user.getUserLoginId();
    return 'system';
  } catch (err) {
    return 'system';
  }
}

function getTasksForTechnicians(techNamesArray, dateStr, startTime, endTime) {
 let tasks = [];
 if (!techNamesArray || techNamesArray.length === 0 || !dateStr) return tasks;

 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const sheetSpis = ss.getSheetByName('Spis wykonanych prac');
 if (!sheetSpis) return tasks;

 const lastRow = sheetSpis.getLastRow();
 if (lastRow < 2) return tasks;

 const data = sheetSpis.getRange(2, 1, lastRow - 1, 10).getDisplayValues();
 const rawDates = sheetSpis.getRange(2, 1, lastRow - 1, 1).getValues();
 const searchTechs = techNamesArray.map(t => String(t).trim().toLowerCase());
 const selectedStartMin = parseTimeToMinutes(startTime);
 const selectedEndMin = parseTimeToMinutes(endTime);

 for (let i = 0; i < data.length; i++) {
   let rowDate = parseDateToYMD(rawDates[i][0]);
   let rowStart = data[i][1];
   let rowEnd = data[i][2];
   let rowDeck = String(data[i][3]).trim().toUpperCase();
   let rowArea = String(data[i][4] || '').trim();
   let rowTask = data[i][5];
  let rowTaskId = String(data[i][6] || '').trim();
   let rowRoomId = String(data[i][8] || '').trim();
   let rowTech = String(data[i][7] || "").trim().toLowerCase();
   let rowTechOriginal = String(data[i][7] || "").trim();

   if (!rowStart || !rowEnd) continue;
   if (!searchTechs.includes(rowTech) || rowDate !== dateStr) continue;

   const taskStartMin = parseTimeToMinutes(rowStart);
   const taskEndMin = parseTimeToMinutes(rowEnd);
   let isInSelectedWindow = true;

   if (selectedStartMin !== null && selectedEndMin !== null && selectedEndMin > selectedStartMin) {
     isInSelectedWindow = (taskStartMin !== null && taskEndMin !== null) &&
       taskStartMin < selectedEndMin && taskEndMin > selectedStartMin;
   }

   if (!isInSelectedWindow) continue;

   tasks.push({
     techName: rowTechOriginal,
     start: rowStart,
     end: rowEnd,
     deck: rowDeck,
     area: rowArea,
     roomId: rowRoomId,
    taskId: rowTaskId,
     taskName: rowTask,
     dbRowIndex: i + 2
   });
 }

 tasks.sort((a, b) => {
   let nameCmp = a.techName.localeCompare(b.techName);
   if (nameCmp !== 0) return nameCmp;
   return a.start.localeCompare(b.start);
 });

 return tasks;
}

function updateWorkReportTask(rowIndex, payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSpis = ss.getSheetByName('Spis wykonanych prac');
    if (!sheetSpis) return { success: false, message: 'Brak zakładki Spis wykonanych prac.' };
    if (!rowIndex || rowIndex < 2 || rowIndex > sheetSpis.getLastRow()) {
      return { success: false, message: 'Nie znaleziono wpisu do edycji.' };
    }

    if (payload && payload.start) sheetSpis.getRange(rowIndex, 2).setValue(payload.start);
    if (payload && payload.end) sheetSpis.getRange(rowIndex, 3).setValue(payload.end);
    sheetSpis.getRange(rowIndex, 10).setValue(getCurrentUserLogin());

    if (payload && payload.progress !== undefined) {
      const sheetTasks = ss.getSheetByName('Lista zadań');
      if (sheetTasks && sheetTasks.getLastRow() >= 3) {
        const taskRows = sheetTasks.getRange(3, 1, sheetTasks.getLastRow() - 2, 10).getDisplayValues();
        const taskId = String(payload.taskId || '').trim().toLowerCase();
        const taskName = String(payload.taskName || '').trim().toLowerCase();
        const deck = normalizeDeckName(payload.deck || '');
        const area = String(payload.area || '').trim().toLowerCase();
        const roomId = normalizeRoomMatch(payload.roomId || '');
        let taskIndex = -1;
        if (taskName) {
          taskIndex = taskRows.findIndex(row =>
            normalizeDeckName(row[2]) === deck &&
            String(row[3] || '').trim().toLowerCase() === area &&
            normalizeRoomMatch(row[4]) === roomId &&
            String(row[9] || '').trim().toLowerCase() === taskName
          );
        }
        if (taskIndex < 0 && taskId) {
          taskIndex = taskRows.findIndex(row =>
            [row[6], row[1]].some(value => String(value || '').trim().toLowerCase() === taskId)
          );
        }
        if (taskIndex >= 0) {
          const taskRow = taskIndex + 3;
          const progress = Number(payload.progress) || 0;
          sheetTasks.getRange(taskRow, 17).setValue(progress + '%');
          sheetTasks.getRange(taskRow, 14).setValue(progress >= 100 ? sheetSpis.getRange(rowIndex, 1).getValue() : '');
        }
      }
    }

    return { success: true, message: 'Zadanie technika zostało zaktualizowane.' };
  } catch (err) {
    return { success: false, message: 'Błąd edycji zadania: ' + err.toString() };
  }
}

function cancelWorkReport(rowIndex) {
 try {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const sheetSpis = ss.getSheetByName('Spis wykonanych prac');
 if (sheetSpis) {
 if (rowIndex > 1 && rowIndex <= sheetSpis.getLastRow()) {
 let userEmail = Session.getActiveUser().getEmail() || "Nieznany użytkownik";
 let timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd.MM.yyyy HH:mm");
 let cancelMsg = "[ANULOWANE] przez: " + userEmail + " (" + timestamp + ")";
 sheetSpis.getRange(rowIndex, 2, 1, 2).clearContent();
 sheetSpis.getRange(rowIndex, 6).setValue(cancelMsg);
 return { success: true, message: "Wpis anulowano pomyślnie (ślad pozostał w bazie)." };
 }
 }
 return { success: false, message: "Nie udało się zlokalizować wpisu w bazie." };
 } catch (error) {
 return { success: false, message: "Błąd podczas usuwania: " + error.toString() };
 }
}

function addTechnician(name) {
 if (!name || name.trim() === "") return { success: false, message: "Nazwisko nie może być puste!" };
 try {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const techName = name.trim();
 const dataSheet = ss.getSheetByName('Data sheet');
 if (!dataSheet) return { success: false, message: "Błąd: Brak zakładki 'Data sheet' w arkuszu." };

 const lastRow = dataSheet.getLastRow();
 let added = false;

 if (lastRow >= 9) {
 const values = dataSheet.getRange("B9:B").getValues();
 for (let i = 0; i < values.length; i++) {
 if (String(values[i][0]).trim() === "") {
 dataSheet.getRange(9 + i, 2).setValue(techName);
 added = true;
 break;
 }
 }
 }
 if (!added) dataSheet.getRange(Math.max(lastRow + 1, 9), 2).setValue(techName);
 return { success: true, message: "Pomyślnie dodano technika: " + techName, newTech: techName };
 } catch (error) {
 return { success: false, message: "Błąd podczas dodawania technika: " + error.toString() };
 }
}

function toggleTechnicianStatus(name, newStatus) {
 try {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const dataSheet = ss.getSheetByName('Data sheet');
 if (!dataSheet) return { success: false, message: "Brak zakładki 'Data sheet'." };

 const lastRow = dataSheet.getLastRow();
 if (lastRow >= 9) {
 const techData = dataSheet.getRange(9, 2, lastRow - 8, 1).getValues();
 for (let i = 0; i < techData.length; i++) {
 if (String(techData[i][0]).trim().toLowerCase() === name.trim().toLowerCase()) {
 dataSheet.getRange(9 + i, 3).setValue(newStatus);
 return { success: true, message: "Status technika zaktualizowany!" };
 }
 }
 }
 return { success: false, message: "Nie znaleziono technika w bazie." };
 } catch (error) {
 return { success: false, message: "Błąd zmiany statusu: " + error.toString() };
 }
}

function assignTechnicianToTeam(techName, teamName) {
 try {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const dataSheet = ss.getSheetByName('Data sheet');
 if (!dataSheet) return { success: false, message: "Brak zakładki 'Data sheet'." };

 const lastRow = dataSheet.getLastRow();
 if (lastRow >= 9) {
 const techData = dataSheet.getRange(9, 2, lastRow - 8, 1).getValues(); 
for (let i = 0; i < techData.length; i++) {
 if (String(techData[i][0]).trim().toLowerCase() === techName.trim().toLowerCase()) {
 dataSheet.getRange(9 + i, 1).setValue(teamName);
 return { success: true, message: "Drużyna zaktualizowana!" };
 }
 }
 }
 return { success: false, message: "Nie znaleziono technika w bazie." };
 } catch (error) {
 return { success: false, message: "Błąd zmiany drużyny: " + error.toString() };
 }
}

function parseDateToYMD(val) {
 if (!val) return null;
 if (val instanceof Date) {
 let d = val.getDate().toString().padStart(2, '0');
 let m = (val.getMonth() + 1).toString().padStart(2, '0');
 let y = val.getFullYear();
 return y + "-" + m + "-" + d;
 }
 if (typeof val === 'string') {
 let str = val.trim();
 if (str.match(/^\d{4}-\d{2}-\d{2}$/)) return str;
 let m1 = str.match(/(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{2,4})/);
 if (m1) {
 let d = m1[1].padStart(2, '0');
 let m = m1[2].padStart(2, '0');
 let y = m1[3];
 if (y.length === 2) y = "20" + y;
 return y + "-" + m + "-" + d;
 }
 }
 return null;
}

 function submitWorkReport(formData) {
 try {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 let idPrac = "";
 let savedTaskName = String(formData.taskName || '').trim();
 let savedTaskAmount = '';
 let savedWorkName = '';
 let savedWorkId = '';
 const sheetCandido = ss.getSheetByName('Lista zadań');
 let techsArray = formData.techniciansArray;
 if (!techsArray || !Array.isArray(techsArray)) {
   return { success: false, message: "Brak techników do przypisania." };
 }
 if (formData.workDate && formData.timeStart && formData.timeEnd) {
   const conflicts = getTasksForTechnicians(techsArray, formData.workDate, formData.timeStart, formData.timeEnd);
   if (conflicts.length > 0) {
     const names = [...new Set(conflicts.map(t => t.techName))].join(', ');
     return { success: false, message: 'Konflikt godzinowy: zadania kolidują z technikami: ' + names };
   }
 }
 if (formData.isCustomTask) {
 if (sheetCandido) {
 const deck = normalizeDeckName(formData.deck);
 const taskName = String(formData.taskName || '').trim();
 const taskType = normalizeTaskTypeForSheet(formData.taskType);
 const isBreakOrTravelTask = String(deck || '').toUpperCase() === 'TRAVEL' || String(deck || '').toUpperCase() === 'BREAK';
 const savedTaskType = taskType || (isBreakOrTravelTask ? '-' : '');
 
 if (!deck) return { success: false, message: "Błąd: Pokład jest wymagany." };
 if (!taskName) return { success: false, message: "Błąd: Nazwa zadania jest wymagana." };
 if (!taskType && !isBreakOrTravelTask) return { success: false, message: "Błąd: Typ zadania jest wymagany. Dopuszczone wartości: ROZBIÓRKA, MONTAŻ, ROZBIÓRKA Z ODZYSKIEM, MONTAŻ Z ODZYSKU." };
 
 const normalizeTaskKey = value => String(value || '').replace(/\s*-\s*EXTRA JOB$/i, '').trim().toLowerCase();
 const areaPrefix = value => {
   const normalized = String(value || '').toLowerCase();
   if (normalized.includes('podłog') || normalized.includes('podlog')) return 'P';
   if (normalized.includes('ścian') || normalized.includes('scian')) return 'W';
   if (normalized.includes('sufit')) return 'C';
   if (normalized.includes('instal')) return 'A';
   return 'A';
 };
 const taskRows = sheetCandido.getLastRow() >= 3
   ? sheetCandido.getRange(3, 1, sheetCandido.getLastRow() - 2, 10).getDisplayValues()
   : [];
 const existingTask = taskRows.find(row => normalizeTaskKey(row[9]) === normalizeTaskKey(taskName));
 const existingId = existingTask ? String(existingTask[6] || existingTask[1] || '').trim() : '';
 if (existingId) {
   idPrac = existingId.replace(/-EJ$/i, '');
 } else if (String(formData.customTaskId || '').trim()) {
   idPrac = String(formData.customTaskId).trim();
 } else {
   const prefix = areaPrefix(formData.area);
   const usedNumbers = taskRows.flatMap(row => [row[1], row[6]])
     .map(value => String(value || '').match(new RegExp('^' + prefix + '(\\d{4})(?:-EJ)?$', 'i')))
     .filter(Boolean)
     .map(match => Number(match[1]));
   const nextNumber = usedNumbers.length ? Math.max(...usedNumbers) + 1 : 1;
   idPrac = prefix + String(nextNumber).padStart(4, '0');
 }
 let taskNameForSaving = taskName;
 if (formData.isExtraJob) {
 if (!idPrac.toUpperCase().endsWith("-EJ")) idPrac += "-EJ";
 if (!taskNameForSaving.toUpperCase().includes("- EXTRA JOB")) {
 taskNameForSaving += " - EXTRA JOB";
 }
 }
 const formattedDate = String(formData.workDate || '').trim();
 const taskUnit = String(formData.taskUnit || '').trim().toLowerCase();
 const taskAmount = taskUnit === 'pis'
   ? Number(formData.taskQuantity) || 0
   : taskUnit === 'm'
     ? Number(formData.taskLength) || 0
     : taskUnit === 'm2'
       ? resolveTaskSurfaceForRow(formData.deck, formData.area, formData.roomId, taskNameForSaving, taskType)
       : 0;
 const completedDate = String(formData.progress || '') === '100' ? formattedDate : '';
 const newTaskRow = [
 "", // 0 - A
 idPrac, // 1 - B - ID zadania
 deck, // 2 - C - Pokład
 formData.area || "", // 3 - D - Strefa
 formData.roomId || "", // 4 - E - Pokój
 formData.taskUnit || "", // 5 - F - Jednostka miary
 idPrac, // 6 - G - ID prac
 savedTaskType, // 7 - H - Typ zadania
 "Seavia", // 8 - I
 String(taskNameForSaving || '').trim(), // 9 - J - Nazwa zadania
 formattedDate, // 10 - K - Data zadania
 "", // 11 - L
 "", // 12 - M
 completedDate, // 13 - N - Data osiągnięcia 100%
 "", // 14 - O
 "", // 15 - P
 formData.progress ? formData.progress + "%" : "", // 16 - Q - Progress
 taskAmount, // 17 - R - Powierzchnia/Ilość
 ];
 const firstDataRow = 3;
 const lastRow = Math.max(sheetCandido.getLastRow(), firstDataRow - 1);
 const sheetRowsForInsert = lastRow >= firstDataRow
   ? sheetCandido.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, newTaskRow.length).getDisplayValues()
   : [];
 const emptyRowIndex = sheetRowsForInsert.findIndex(row => row.every(value => String(value).trim() === ""));
 const targetRow = emptyRowIndex === -1 ? lastRow + 1 : firstDataRow + emptyRowIndex;
 const canonicalTaskRow = [...newTaskRow];
 sheetCandido.getRange(targetRow, 8).setDataValidation(
   SpreadsheetApp.newDataValidation()
     .requireValueInList(['ROZBIÓRKA', 'MONTAŻ', 'ROZBIÓRKA Z ODZYSKIEM', 'MONTAŻ Z ODZYSKU', '-'], true)
     .setAllowInvalid(false)
     .build()
 );
 sheetCandido.getRange(targetRow, 1, 1, canonicalTaskRow.length).setValues([canonicalTaskRow]);
 savedTaskName = sheetCandido.getRange(targetRow, 10).getDisplayValue();
 savedTaskAmount = sheetCandido.getRange(targetRow, 18).getDisplayValue();
 Logger.log(`Lista zadań row ${targetRow}: H=${sheetCandido.getRange(targetRow, 8).getDisplayValue()}, J=${savedTaskName}, R=${savedTaskAmount}, B=${sheetCandido.getRange(targetRow, 2).getDisplayValue()}, F=${sheetCandido.getRange(targetRow, 6).getDisplayValue()}, G=${sheetCandido.getRange(targetRow, 7).getDisplayValue()}`);
 }
 } else {
 if (formData.taskRow) {
 const taskRowNumber = Number(formData.taskRow);
 let isNewLengthTask = false;
 if (sheetCandido && Number.isFinite(taskRowNumber) && taskRowNumber >= 3 && taskRowNumber <= sheetCandido.getLastRow()) {
 idPrac = sheetCandido.getRange(taskRowNumber, 7).getValue();
 if (!idPrac || idPrac === "") idPrac = "Wiersz " + formData.taskRow;
 const taskUnit = String(formData.taskUnit || '').trim().toLowerCase();
 savedTaskName = String((formData.taskName || '').trim() || sheetCandido.getRange(taskRowNumber, 10).getDisplayValue() || '');
 const currentAmount = Number(String(sheetCandido.getRange(taskRowNumber, 18).getDisplayValue() || '0').replace(',', '.')) || 0;
 const selectedAmount = Number(taskUnit === 'pis' ? formData.taskQuantity : formData.taskLength) || 0;
 savedTaskAmount = String(currentAmount);
 if (taskUnit === 'pis' && selectedAmount > 0) {
   sheetCandido.getRange(taskRowNumber, 18).setValue(selectedAmount);
   savedTaskAmount = String(selectedAmount);
 } else if (taskUnit === 'm' && formData.taskLengthConfirmed && selectedAmount > 0) {
   if (Math.abs(selectedAmount - currentAmount) > 0.0001) {
     const taskValues = sheetCandido.getRange(taskRowNumber, 1, 1, 18).getValues()[0];
     const newTaskId = generateTaskIdForArea_(sheetCandido, taskValues[3], String(taskValues[6] || '').toUpperCase().endsWith('-EJ'));
     taskValues[1] = newTaskId;
     taskValues[6] = newTaskId;
     taskValues[10] = formData.workDate || taskValues[10];
     taskValues[13] = '';
     taskValues[16] = '0%';
     taskValues[17] = selectedAmount;
     const lastTaskRow = sheetCandido.getLastRow();
     const taskSheetRows = lastTaskRow >= 3 ? sheetCandido.getRange(3, 1, lastTaskRow - 2, 19).getDisplayValues() : [];
     const blankIndex = taskSheetRows.findIndex(row => row.every(value => String(value || '').trim() === ''));
     const newRowIndex = blankIndex < 0 ? lastTaskRow + 1 : blankIndex + 3;
     sheetCandido.getRange(newRowIndex, 8).setDataValidation(
       SpreadsheetApp.newDataValidation()
         .requireValueInList(['ROZBIÓRKA', 'MONTAŻ', 'ROZBIÓRKA Z ODZYSKIEM', 'MONTAŻ Z ODZYSKU', '-'], true)
         .setAllowInvalid(false)
         .build()
     );
     sheetCandido.getRange(newRowIndex, 1, 1, taskValues.length).setValues([taskValues]);
     idPrac = newTaskId;
     savedTaskAmount = String(selectedAmount);
     isNewLengthTask = true;
 }
 } else if (taskUnit === 'm2') {
   const taskType = sheetCandido.getRange(taskRowNumber, 8).getDisplayValue();
   const measuredSurface = resolveTaskSurfaceForRow(formData.deck, formData.area, formData.roomId, savedTaskName, taskType);
   if (measuredSurface > 0) {
     sheetCandido.getRange(taskRowNumber, 18).setValue(measuredSurface);
     savedTaskAmount = String(measuredSurface);
   }
 }
 if (formData.progress && !isNewLengthTask) {
 sheetCandido.getRange(taskRowNumber, 17).setValue(formData.progress + "%");
 if (String(formData.progress) === "100" && formData.workDate) {
   sheetCandido.getRange(taskRowNumber, 14).setValue(formData.workDate);
 }
 }
 // Dla istniejacego zadania nie przeliczamy ponownie kolumny R.
 // Zapisana ilosc/powierzchnia pozostaje taka sama przy kontynuacji pracy.
 }
 }
 }
 let sheetSpis = ss.getSheetByName('Spis wykonanych prac');
 if (!sheetSpis) sheetSpis = ss.insertSheet('Spis wykonanych prac');
 if (sheetSpis.getLastRow() === 0) {
 sheetSpis.appendRow(["Data", "Godzina rozpoczęcia pracy", "Godzina zakończenia pracy", "Pokład", "Typ prac", "Nazwa prac", "ID prac", "Technik", "Pokój", "Edytował"]);
 sheetSpis.getRange(1, 1, 1, 10).setFontWeight("bold");
 }
 let roomSurface = Number(formData.roomSurface) || 0;
 if (!techsArray || !Array.isArray(techsArray)) {
   return { success: false, message: "Brak techników do przypisania." };
 }
 const editorLogin = getCurrentUserLogin();
 let taskNameForReport = String(formData.taskName || '').trim();
 if (formData.isExtraJob && !taskNameForReport.toUpperCase().includes("- EXTRA JOB")) {
   taskNameForReport += " - EXTRA JOB";
 }
 techsArray.forEach(tech => {
 sheetSpis.appendRow([
 formData.workDate, 
formData.timeStart, 
formData.timeEnd, 
formData.deck ? String(formData.deck).trim().toUpperCase() : "",
 formData.area,
 taskNameForReport, 
idPrac, 
tech,
 formData.roomId || "",
 editorLogin
 ]);
 });
 const savedWorkRow = sheetSpis.getLastRow();
 savedWorkName = sheetSpis.getRange(savedWorkRow, 6).getDisplayValue();
 savedWorkId = sheetSpis.getRange(savedWorkRow, 7).getDisplayValue();
 Logger.log(`Spis wykonanych prac row ${savedWorkRow}: F=${savedWorkName}, G=${savedWorkId}`);
 let msgSuffix = techsArray.length > 1 ? " (Dla " + techsArray.length + " techników)" : '';
 return { success: true, message: `Pomyślnie zapisano godziny i postęp zadania!${msgSuffix}\nLista zadań: J=${savedTaskName}, R=${savedTaskAmount}\nSpis wykonanych prac: F=${savedWorkName}, G=${savedWorkId}` };
 } catch (error) {
 return { success: false, message: "Błąd podczas zapisu: " + error.toString() };
 }
}

function getCandidoWeeklyReport(dateStr, includeBreaks) {
 if (!dateStr) return null;
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const sheetSpis = ss.getSheetByName('Spis wykonanych prac');
 if (!sheetSpis || sheetSpis.getLastRow() < 2) return { decks: {}, weekInfo: {} };

 let selectedDate = new Date(dateStr);
 let dayOfWeek = selectedDate.getDay();
 let diff = selectedDate.getDate() - dayOfWeek + (dayOfWeek === 0 ? -6 : 1);
 let monday = new Date(selectedDate.setDate(diff));
 let weekDates = [];
 for (let i = 0; i < 7; i++) {
 let d = new Date(monday);
 d.setDate(monday.getDate() + i);
 weekDates.push(parseDateToYMD(d));
 }
 let target = new Date(monday.valueOf());
 let dayNr = (monday.getDay() + 6) % 7;
 target.setDate(target.getDate() - dayNr + 3);
 let firstThursday = target.valueOf();
 target.setMonth(0, 1);
 if (target.getDay() !== 4) {
 target.setMonth(0, 1 + ((4 - target.getDay()) + 7) % 7);
 }
 let weekNum = 1 + Math.ceil((firstThursday - target) / 604800000);

 let reportData = {
 settings: { includeBreaks: includeBreaks },
 weekInfo: {
 weekNumber: "W" + weekNum,
 start: weekDates[0],
 end: weekDates[6]
 },
 decks: {},
 allTasks: [],
 extraTasks: [],
 extraDeck: { tasks: {} }
 };

 let uniqueTasks = new Set();
 const sheetCandido = ss.getSheetByName('Lista zadań');
 if (sheetCandido && sheetCandido.getLastRow() >= 3) {
 // ZMIANA 3: Indeks kolumny zadania 9 -> 10
 const candidoTasks = sheetCandido.getRange(3, 10, sheetCandido.getLastRow() - 2, 1).getValues();
 candidoTasks.forEach(row => {
 let t = String(row[0]).trim();
 if (t && !t.includes("[ANULOWANE]") && !t.toLowerCase().includes("extra job")) {
 uniqueTasks.add(t);
 }
 });
 }

 const spisData = sheetSpis.getRange(2, 1, sheetSpis.getLastRow() - 1, 8).getDisplayValues();
 const rawDates = sheetSpis.getRange(2, 1, sheetSpis.getLastRow() - 1, 1).getValues();
 for (let i = 0; i < spisData.length; i++) {
 let rowDate = parseDateToYMD(rawDates[i][0]);
 let dayIndex = weekDates.indexOf(rowDate);
 if (dayIndex === -1) continue;

 let rowStart = spisData[i][1];
 let rowEnd = spisData[i][2];
 let rowDeck = String(spisData[i][3]).trim().toUpperCase();
 let rowTask = String(spisData[i][5]).trim();
 let taskId = String(spisData[i][6]).trim();
 let isCancelled = rowTask.includes("[ANULOWANE]") || taskId.includes("[ANULOWANE]");
 let isBreak = rowTask.toLowerCase().includes("przerwa") || rowTask.toLowerCase().includes("break");

 if (isBreak && !includeBreaks) continue;
 if (!rowStart || !rowEnd || isCancelled || !rowDeck || !rowTask) continue;

 if (rowTask.toLowerCase().includes("site manager") || rowTask.toLowerCase().includes("meetings") || rowTask.toLowerCase() === "supervision") {
 rowTask = "Supervision";
 }

 let duration = parseWorkTimeBackend(rowEnd) - parseWorkTimeBackend(rowStart);
 if (duration < 0) duration += 24; 
duration = Math.round(duration * 100) / 100;

 if (duration > 0) {
 let isExtraJob = taskId.toLowerCase().includes("- extra job");
 if (isExtraJob) {
 let extraTaskKey = rowDeck + ", " + rowTask;
 if (!reportData.extraDeck.tasks[extraTaskKey]) {
 reportData.extraDeck.tasks[extraTaskKey] = [0, 0, 0, 0, 0, 0, 0];
 }
 reportData.extraDeck.tasks[extraTaskKey][dayIndex] += duration;
 } else {
 if (rowDeck === "TRAVEL") {
 rowTask = "Travel";
 }
 if (!reportData.decks[rowDeck]) {
 reportData.decks[rowDeck] = { tasks: {} };
 }
 if (!reportData.decks[rowDeck].tasks[rowTask]) {
 reportData.decks[rowDeck].tasks[rowTask] = [0, 0, 0, 0, 0, 0, 0];
 }
 reportData.decks[rowDeck].tasks[rowTask][dayIndex] += duration;
 if(rowTask === "Supervision") {
 uniqueTasks.add(rowTask);
 }
 }
 }
 }

 reportData.allTasks = Array.from(uniqueTasks).sort((a, b) => {
 return a.localeCompare(b);
 });

 let extraTasksSet = new Set();
 Object.keys(reportData.extraDeck.tasks).forEach(k => {
 let total = reportData.extraDeck.tasks[k].reduce((a, b) => a + b, 0);
 if (total > 0) {
 extraTasksSet.add(k);
 }
 });
 reportData.extraTasks = Array.from(extraTasksSet).sort((a, b) => a.localeCompare(b));

 Object.keys(reportData.decks).forEach(d => {
 Object.keys(reportData.decks[d].tasks).forEach(t => {
 reportData.decks[d].tasks[t] = reportData.decks[d].tasks[t].map(h => Math.round(h * 10) / 10);
 });
 });
 Object.keys(reportData.extraDeck.tasks).forEach(t => {
 reportData.extraDeck.tasks[t] = reportData.extraDeck.tasks[t].map(h => Math.round(h * 10) / 10);
 });
 return reportData;
}

function getTimeByWeekReport(includeBreaks) {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const sheetSpis = ss.getSheetByName('Spis wykonanych prac');
 if (!sheetSpis || sheetSpis.getLastRow() < 2) return { decks: {}, allTasks: [], allWeeks: [] };

 let reportData = {
 settings: { includeBreaks: includeBreaks },
 decks: {},
 allTasks: [],
 allWeeks: new Set(),
 extraTasks: [],
 extraDeck: { tasks: {} }
 };

 const spisData = sheetSpis.getRange(2, 1, sheetSpis.getLastRow() - 1, 8).getDisplayValues();
 const rawDates = sheetSpis.getRange(2, 1, sheetSpis.getLastRow() - 1, 1).getValues();

 const getWeekNumber = (d) => {
 d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
 d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
 var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
 var weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
 return weekNo;
 };

 for (let i = 0; i < spisData.length; i++) {
 let rowDate = rawDates[i][0];
 if (!(rowDate instanceof Date)) continue;
 let weekNum = getWeekNumber(rowDate);
 let rowStart = spisData[i][1];
 let rowEnd = spisData[i][2];
 let rowDeck = String(spisData[i][3]).trim().toUpperCase();
 let rowTask = String(spisData[i][5]).trim();
 let taskId = String(spisData[i][6]).trim();
 let isCancelled = rowTask.toUpperCase().includes("ANULOWANE") || taskId.toUpperCase().includes("ANULOWANE");
 let isBreak = /przerwa|break|brake/i.test(rowTask) || /przerwa|break|brake/i.test(rowDeck);

 if (isBreak && !includeBreaks) continue;
 if (!rowStart || !rowEnd || isCancelled || !rowDeck || !rowTask) continue;

 if (rowTask.toLowerCase().includes("site manager") || rowTask.toLowerCase().includes("meetings") || rowTask.toLowerCase() === "supervision") {
 rowTask = "Supervision";
 }

 let duration = parseWorkTimeBackend(rowEnd) - parseWorkTimeBackend(rowStart);
 if (duration < 0) duration += 24;
 duration = Math.round(duration * 100) / 100;

 if (duration > 0) {
 let isExtraJob = /-\s*extra job/i.test(rowTask) || /-\s*extra job/i.test(taskId);
 if (isExtraJob) {
 let extraTaskKey = rowDeck + ", " + rowTask;
 if (!reportData.extraDeck.tasks[extraTaskKey]) {
 reportData.extraDeck.tasks[extraTaskKey] = { total: 0, weeklyHours: {} };
 }
 if (!reportData.extraDeck.tasks[extraTaskKey].weeklyHours[weekNum]) {
 reportData.extraDeck.tasks[extraTaskKey].weeklyHours[weekNum] = 0;
 }
 reportData.extraDeck.tasks[extraTaskKey].weeklyHours[weekNum] += duration;
 reportData.extraDeck.tasks[extraTaskKey].total += duration;
 } else {
 if (rowDeck === "TRAVEL") {
 rowTask = "Travel";
 }
 if (!reportData.decks[rowDeck]) {
 reportData.decks[rowDeck] = { tasks: {} };
 }
 if (!reportData.decks[rowDeck].tasks[rowTask]) {
 reportData.decks[rowDeck].tasks[rowTask] = { total: 0, weeklyHours: {} };
 }
 if (!reportData.decks[rowDeck].tasks[rowTask].weeklyHours[weekNum]) {
 reportData.decks[rowDeck].tasks[rowTask].weeklyHours[weekNum] = 0;
 }
 reportData.decks[rowDeck].tasks[rowTask].weeklyHours[weekNum] += duration;
 reportData.decks[rowDeck].tasks[rowTask].total += duration;
 if (reportData.allTasks.indexOf(rowTask) === -1) {
 reportData.allTasks.push(rowTask);
 }
 }
 reportData.allWeeks.add(weekNum);
 }
 }

 reportData.allTasks = reportData.allTasks.sort((a, b) => a.localeCompare(b));
 reportData.allWeeks = Array.from(reportData.allWeeks).sort((a, b) => a - b);

 let extraTasksSet = new Set();
 Object.keys(reportData.extraDeck.tasks).forEach(k => {
 if (reportData.extraDeck.tasks[k].total > 0) {
 extraTasksSet.add(k);
 }
 });
 reportData.extraTasks = Array.from(extraTasksSet).sort((a, b) => a.localeCompare(b));

 Object.keys(reportData.decks).forEach(d => {
 Object.keys(reportData.decks[d].tasks).forEach(t => {
 let task = reportData.decks[d].tasks[t];
 task.total = Math.round(task.total * 10) / 10;
 Object.keys(task.weeklyHours).forEach(w => {
 task.weeklyHours[w] = Math.round(task.weeklyHours[w] * 10) / 10;
 });
 });
 });
 Object.keys(reportData.extraDeck.tasks).forEach(t => {
 let task = reportData.extraDeck.tasks[t];
 task.total = Math.round(task.total * 10) / 10;
 Object.keys(task.weeklyHours).forEach(w => {
 task.weeklyHours[w] = Math.round(task.weeklyHours[w] * 10) / 10;
 });
 });
 return reportData;
}

 function getTechnicianWeeklyReports(dateStr, dateToStr, includeBreaks) {
 // Compatibility with the previous two-argument client call.
 if (typeof dateToStr === 'boolean') {
   includeBreaks = dateToStr;
   dateToStr = dateStr;
 }
 if (!dateStr || !dateToStr) return null;

 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const sheetSpis = ss.getSheetByName('Spis wykonanych prac');
 const dataSheet = ss.getSheetByName('Data sheet');
 if (!sheetSpis || !dataSheet || sheetSpis.getLastRow() < 2) return null;

 const startDate = new Date(dateStr + 'T00:00:00');
 const endDate = new Date(dateToStr + 'T00:00:00');
 if (isNaN(startDate.getTime()) || isNaN(endDate.getTime()) || endDate < startDate) return null;

 const weekDates = [];
 for (let date = new Date(startDate); date <= endDate; date.setDate(date.getDate() + 1)) {
   weekDates.push(parseDateToYMD(date));
 }

 // Google Sheets NUM.TYG without an explicit return type starts weeks on Sunday.
 const yearStart = new Date(startDate.getFullYear(), 0, 1);
 const daysFromYearStart = Math.floor((startDate - yearStart) / 86400000);
 const weekNum = Math.floor((daysFromYearStart + yearStart.getDay()) / 7) + 1;
 const headerValues = dataSheet.getRangeList(['C1', 'C2', 'C4']).getRanges().map(range => range.getDisplayValue());
 const reportData = {
   header: {
     title: headerValues[0],
     project: headerValues[1],
     company: headerValues[2]
   },
   weekInfo: {
     weekNumber: weekNum,
     year: startDate.getFullYear(),
     dates: weekDates
   },
   technicians: {}
 }; 
 
const spisData = sheetSpis.getRange(2, 1, sheetSpis.getLastRow() - 1, 9).getDisplayValues();
 const rawDates = sheetSpis.getRange(2, 1, sheetSpis.getLastRow() - 1, 1).getValues();
 
for (let i = 0; i < spisData.length; i++) {
 let rowDate = parseDateToYMD(rawDates[i][0]);
 let dayIndex = weekDates.indexOf(rowDate);
 if (dayIndex === -1) continue;
 
let rowStart = String(spisData[i][1]).trim();
 let rowEnd = String(spisData[i][2]).trim();
 let rowDeck = String(spisData[i][3]).trim().toUpperCase();
 let rowArea = String(spisData[i][4]).trim();
 let rowTask = String(spisData[i][5]).trim();
 let taskId = String(spisData[i][6]).trim(); // ID prac
 let techName = String(spisData[i][7]).trim(); // Technik
 let rowRoomId = String(spisData[i][8] || "").trim(); // NOWE: Kolumna I (Pokój)
 let isCancelled = rowTask.includes("[ANULOWANE]") || rowTask.includes("ANULOWANE");
 let isBreak = rowTask.toLowerCase().includes("przerwa") || rowTask.toLowerCase().includes("break") || rowTask.toLowerCase().includes("brake");
 
if (isBreak && !includeBreaks) continue;
 if (!rowStart || !rowEnd || isCancelled || !techName || !rowTask) continue;
 
let duration = parseWorkTimeBackend(rowEnd) - parseWorkTimeBackend(rowStart);
 if (duration < 0) duration += 24;
 duration = Math.round(duration * 100) / 100;
 
if (duration > 0) {
 if (!reportData.technicians[techName]) {
 reportData.technicians[techName] = {
 totalHours: 0,
 days: weekDates.map(() => [])
 }; 
 }
 reportData.technicians[techName].days[dayIndex].push({
 start: rowStart,
 end: rowEnd,
 hours: duration,
 taskId: taskId,
 task: rowTask,
 deck: rowDeck,
 area: rowArea,
 roomId: rowRoomId
 });
 reportData.technicians[techName].totalHours += duration;
 }
 }
 
Object.keys(reportData.technicians).forEach(techName => {
 reportData.technicians[techName].totalHours = Math.round(reportData.technicians[techName].totalHours * 10) / 10;
 });
 
// URUCHOMIENIE W TLE: Synchronizacja z REV OCEAN - timesheet
 // Nie przerywamy generowania kart pracy, jeśli synchronizacja do arkusza timesheet
 // nie powinna blokować eksportu danych do PDF.
 try {
   syncWeeklyTimesheet(reportData);
 } catch (error) {
   Logger.log('Błąd synchronizacji REV OCEAN - timesheet podczas generowania kart pracy: ' + error);
 }

 return reportData;
}

// =======================================================================
// NOWY MODUŁ: ZLECENIA I PROTOKOŁY ZDAWCZO-ODBIORCZE (wg zakresu dat)
// UWAGA: funkcja tylko ODCZYTUJE dane z getTechnicianWeeklyReports (bez zmian
// w logice liczenia czasu na kartach pracy) i buduje na ich podstawie nowe,
// niezależne zestawienie do generowania Zlecenia oraz Protokołu.
// =======================================================================
function getContractorsInfoMap_() {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const dataSheet = ss.getSheetByName('Data sheet');
 const map = {};
 if (!dataSheet) return map;

 const lastRow = dataSheet.getLastRow();
 if (lastRow < 9) return map;

 // Kolumny: B=Imię i nazwisko, F=Stawka, H=Numer umowy
 const rows = dataSheet.getRange(9, 1, lastRow - 8, 8).getValues();
 rows.forEach(row => {
 const name = String(row[1] || "").trim();
 if (!name) return;
 map[name.toLowerCase()] = {
 name: name,
   nip: String(row[6] || "").trim(),
 rate: parseFloat(String(row[5] || "0").replace(',', '.')) || 0,
 contractNumber: String(row[7] || "").trim()
 };
 });
 return map;
}

function parseDocumentDate_(value) {
 if (value === null || value === undefined || value === '') return null;
 if (value instanceof Date) {
   const normalized = new Date(value.getTime() - value.getTimezoneOffset() * 60000);
   return Utilities.formatDate(normalized, Session.getScriptTimeZone(), 'yyyy-MM-dd');
 }
 const text = String(value).trim();
 if (!text) return null;
 const match = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
 if (match) {
   const year = Number(match[1]);
   const month = Number(match[2]);
   const day = Number(match[3]);
   const date = new Date(year, month - 1, day);
   if (date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day) {
     return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
   }
 }
 const parsed = new Date(String(value) + 'T00:00:00');
 if (Number.isNaN(parsed.getTime())) return null;
 return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getDocumentRegistrySheet_(sheetName) {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 let sheet = ss.getSheetByName(sheetName);
 if (!sheet) {
   sheet = ss.insertSheet(sheetName);
 }
 if (sheet.getLastRow() === 0) {
   const header = sheetName === 'Zlecenia'
     ? ['Wykonawca', 'Data od', 'Data do', 'Numer zlecenia']
     : ['Wykonawca', 'Data od', 'Data do', 'Numer protokołu', 'Numer zlecenia'];
   sheet.appendRow(header);
 }
 return sheet;
}

function getDocumentStatusForRange_(sheetName, techName, dateFromStr, dateToStr) {
 const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
 if (!sheet) return { status: 'new' };
 const lastRow = sheet.getLastRow();
 if (lastRow < 2) return { status: 'new' };

 const rows = sheet.getRange(1, 1, lastRow, Math.max(5, sheet.getLastColumn())).getValues();
 if (!rows.length) return { status: 'new' };

 const header = rows[0].map(value => String(value || '').trim().toLowerCase());
 const techIndex = header.indexOf('wykonawca');
 const fromIndex = header.indexOf('data od');
 const toIndex = header.indexOf('data do');
 const numberIndex = header.indexOf('numer zlecenia');
 const protocolNumberIndex = header.indexOf('numer protokołu');
 const documentNumberIndex = sheetName === 'Protokoły' ? protocolNumberIndex : numberIndex;
 if (techIndex === -1 || fromIndex === -1 || toIndex === -1 || documentNumberIndex === -1) {
   return { status: 'new' };
 }

 const normalizedTech = String(techName || '').trim().toLowerCase();
 const fromDate = parseDocumentDate_(dateFromStr);
 const toDate = parseDocumentDate_(dateToStr);
 if (!fromDate || !toDate) return { status: 'new' };

 const toMillis = new Date(toDate + 'T00:00:00').getTime();
 const fromMillis = new Date(fromDate + 'T00:00:00').getTime();

 for (let i = 1; i < rows.length; i++) {
   const row = rows[i];
   const rowTech = String(row[techIndex] || '').trim().toLowerCase();
   if (!rowTech || rowTech !== normalizedTech) continue;

   const rowFrom = parseDocumentDate_(row[fromIndex]);
   const rowTo = parseDocumentDate_(row[toIndex]);
   if (!rowFrom || !rowTo) continue;

   const rowFromMillis = new Date(rowFrom + 'T00:00:00').getTime();
   const rowToMillis = new Date(rowTo + 'T00:00:00').getTime();
   const sameRange = rowFrom === fromDate && rowTo === toDate;
   if (sameRange) {
     return {
       status: 'existing',
       documentNumber: String(row[documentNumberIndex] || '').trim(),
       range: [rowFrom, rowTo]
     };
   }

   const overlaps = !(toMillis < rowFromMillis || fromMillis > rowToMillis);
   if (overlaps) {
     return {
       status: 'conflict',
       existing: [rowFrom, rowTo],
       documentNumber: String(row[documentNumberIndex] || '').trim()
     };
   }
 }

 return { status: 'new' };
}

function getNextRegistryDocumentNumber_(sheetName, techName, prefix, year) {
 const sheet = getDocumentRegistrySheet_(sheetName);
 const lastRow = sheet.getLastRow();
 if (lastRow < 2) return prefix + '001-' + year;
 const values = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
 const techColumn = values[0].findIndex(cell => String(cell || '').trim().toLowerCase() === 'wykonawca');
 const numberColumn = values[0].findIndex(cell => String(cell || '').trim().toLowerCase().includes('numer'));
 if (techColumn === -1 || numberColumn === -1) return prefix + '001-' + year;

 let next = 1;
 const normalizedTech = String(techName || '').trim().toLowerCase();
 const pattern = new RegExp('^' + prefix + '(\\d{3})-' + year + '$', 'i');
 for (let i = 1; i < values.length; i++) {
   const row = values[i];
   const rowTech = String(row[techColumn] || '').trim().toLowerCase();
   if (rowTech !== normalizedTech) continue;
   const value = String(row[numberColumn] || '').trim();
   const match = value.match(pattern);
   if (match) next = Math.max(next, Number(match[1]) + 1);
 }
 return prefix + String(next).padStart(3, '0') + '-' + year;
}

function registerDocumentRange_(sheetName, techName, dateFromStr, dateToStr, documentNumber, relatedOrderNumber) {
 const sheet = getDocumentRegistrySheet_(sheetName);
 const values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
 const header = values[0] || [];
 const techIdx = header.findIndex(cell => String(cell || '').trim().toLowerCase() === 'wykonawca');
 const fromIdx = header.findIndex(cell => String(cell || '').trim().toLowerCase() === 'data od');
 const toIdx = header.findIndex(cell => String(cell || '').trim().toLowerCase() === 'data do');
 const numIdx = sheetName === 'Protokoły'
   ? header.findIndex(cell => String(cell || '').trim().toLowerCase() === 'numer protokołu')
   : header.findIndex(cell => String(cell || '').trim().toLowerCase() === 'numer zlecenia');
 const relatedIdx = header.findIndex(cell => String(cell || '').trim().toLowerCase() === 'numer zlecenia');
 const row = [String(techName || '').trim(), parseDocumentDate_(dateFromStr), parseDocumentDate_(dateToStr), documentNumber];
 if (sheetName === 'Protokoły') row.push(relatedOrderNumber || '');
 const targetRow = sheet.getLastRow() + 1;
 sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
 return { success: true, documentNumber: documentNumber };
}

function generateZleceniaIProtokolyData(dateFromStr, dateToStr, includeBreaks) {
 // Ponowne wykorzystanie istniejącej, sprawdzonej logiki liczenia godzin - bez modyfikacji.
 const weeklyData = getTechnicianWeeklyReports(dateFromStr, dateToStr, includeBreaks);
 if (!weeklyData) return { success: false, message: "Brak danych dla podanego zakresu dat." };

 const contractorsInfo = getContractorsInfoMap_();
 const taskCatalog = {};
 const taskCatalogById = {};
 const taskCatalogByName = {};
 const parseReportNumber = value => {
   if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
   const text = String(value ?? '').trim().replace(/\s/g, '');
   if (!text) return 0;
   const normalized = text.includes(',')
     ? text.replace(/\./g, '').replace(',', '.')
     : text;
   return Number(normalized.replace(/[^0-9.-]/g, '')) || 0;
 };
 const sheetTasks = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Lista zadań');
 if (sheetTasks && sheetTasks.getLastRow() >= 3) {
 const lastTaskRow = sheetTasks.getLastRow();
 const taskValues = sheetTasks.getRange(1, 1, lastTaskRow, 20).getValues();
 const taskDisplayValues = sheetTasks.getRange(1, 1, lastTaskRow, 20).getDisplayValues();
 for (let i = 2; i < taskValues.length; i++) {
 const deck = normalizeDeckName(taskValues[i][2]);
 const area = String(taskValues[i][3] || '').trim();
 const room = String(taskValues[i][4] || '').trim();
 const taskName = String(taskDisplayValues[i][9] || '').trim();
 if (!deck || !taskName) continue;

 const rawProgress = taskValues[i][16];
 const progress = typeof rawProgress === 'number'
   ? (rawProgress <= 1 ? rawProgress * 100 : rawProgress)
   : parseInt(String(taskDisplayValues[i][16] || '').replace('%', '').trim(), 10) || 0;
 const unit = String(taskDisplayValues[i][5] || '').trim().toLowerCase();
 const amount = parseReportNumber(taskDisplayValues[i][17]);
 const catalogEntry = { unit: unit, amount: amount, progress: progress };
 const taskId = String(taskDisplayValues[i][6] || taskDisplayValues[i][1] || '').trim();
 if (taskId) taskCatalogById[taskId.toLowerCase()] = catalogEntry;
 const nameKey = taskName.toLowerCase();
 if (!taskCatalogByName[nameKey]) taskCatalogByName[nameKey] = catalogEntry;
 taskCatalog[deck + '|' + area + '|' + room + '|' + taskName] = catalogEntry;
 if (!room) taskCatalog[deck + '|' + area + '||' + taskName] = catalogEntry;
 }
 }
 const AGGREGATED_LABEL = "prace pozostałe niesklasyfikowane";
 const AGGREGATED_UNIT = "roboczogodzina";

 const isAggregatedRow = (deck, task) => {
 const upperDeck = String(deck || "").trim().toUpperCase();
 const normalizedTask = String(task || "").trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
 return upperDeck === "TRAVEL" || upperDeck === "BREAK" ||
 normalizedTask.includes("przerwa") || normalizedTask.includes("break") || normalizedTask.includes("brake") ||
 normalizedTask.startsWith("transport");
 };

 const technicians = {};
 const allTaskHours = {};

 Object.keys(weeklyData.technicians).forEach(techName => {
 const techData = weeklyData.technicians[techName];
 techData.days.forEach(dayEntries => {
 dayEntries.forEach(entry => {
 if (isAggregatedRow(entry.deck, entry.task)) return;
 const key = entry.deck + "||" + entry.area + "||" + entry.roomId + "||" + entry.task;
 allTaskHours[key] = (allTaskHours[key] || 0) + entry.hours;
 });
 });
 });

 Object.keys(weeklyData.technicians).forEach(techName => {
 const techData = weeklyData.technicians[techName];
 const info = contractorsInfo[techName.toLowerCase()] || { name: techName, nip: "", rate: 0, contractNumber: "" };

 const taskLines = []; // Kolejność wg pierwszego wystąpienia unikatowego zadania
 const taskLineIndex = {};
 let aggregatedHours = 0;

 techData.days.forEach(dayEntries => {
 dayEntries.forEach(entry => {
 if (isAggregatedRow(entry.deck, entry.task)) {
 aggregatedHours += entry.hours;
 return;
 }
 const key = entry.deck + "||" + entry.area + "||" + entry.roomId + "||" + entry.task;
 const catalogKey = entry.deck + "|" + entry.area + "|" + entry.roomId + "|" + entry.task;
 const catalogEntry = taskCatalogById[String(entry.taskId || '').trim().toLowerCase()] ||
   taskCatalog[catalogKey] ||
   taskCatalog[entry.deck + "|" + entry.area + "||" + entry.task] ||
   taskCatalogByName[String(entry.task || '').trim().toLowerCase()] ||
   { unit: "h", amount: 0, progress: 100 };
 const normalizedTaskName = String(entry.task || '').trim().toLowerCase();
 const normalizedUnit = String(catalogEntry.unit || 'h').trim().toLowerCase();
 const uniqueTaskKey = normalizedTaskName + "||" + normalizedUnit;
 const allocatedAmount = Number(catalogEntry.amount || 0) * (Number(catalogEntry.progress || 0) / 100) *
   (allTaskHours[key] > 0 ? entry.hours / allTaskHours[key] : 0);
 if (!taskLineIndex.hasOwnProperty(uniqueTaskKey)) {
 taskLineIndex[uniqueTaskKey] = taskLines.length;
 taskLines.push({ key: uniqueTaskKey, task: entry.task, hours: 0, unit: catalogEntry.unit, amount: 0 });
 }
 taskLines[taskLineIndex[uniqueTaskKey]].hours += entry.hours;
 taskLines[taskLineIndex[uniqueTaskKey]].amount += allocatedAmount;
 });
 });

 const lines = taskLines.map((t, idx) => {
 const hours = Math.round(t.hours * 100) / 100;
 const allocatedAmount = t.amount;
 const reportedAmount = ['m', 'm2', 'pis'].includes(String(t.unit || '').toLowerCase())
   ? Math.floor(allocatedAmount)
   : Math.round(allocatedAmount * 100) / 100;
 const protocolAmount = String(t.unit || '').toLowerCase() === 'h'
   ? Math.ceil(allocatedAmount * 2) / 2
   : String(t.unit || '').toLowerCase() === 'pis'
     ? Math.floor(allocatedAmount)
     : Math.floor(allocatedAmount * 100) / 100;
 const laborValue = hours * info.rate;
 const calculationUnitPrice = allocatedAmount > 0 ? laborValue / allocatedAmount : 0;
 const unitPrice = Math.floor(calculationUnitPrice * 10) / 10;
 const value = Math.round((Number(protocolAmount) * Number(unitPrice || 0)) * 100) / 100;
 return {
 lp: idx + 1,
 label: t.task,
 unit: t.unit || "h",
 hours: hours,
 amount: reportedAmount,
 protocolAmount: protocolAmount,
 unitPrice: unitPrice,
 value: value,
 residualValue: Math.round((laborValue - value) * 100) / 100
 };
 });

 aggregatedHours = Math.round(aggregatedHours * 100) / 100;
 const totalWorkedValue = Math.round(lines.reduce((sum, line) => sum + (Number(line.hours || 0) * Number(info.rate || 0)), 0) * 100) / 100;
 const displayedValueTotal = Math.round(lines.reduce((sum, line) => sum + line.value, 0) * 100) / 100;
 const residualValue = Math.max(0, Math.round((totalWorkedValue - displayedValueTotal) * 100) / 100);
 const residualHoursRaw = info.rate > 0 ? residualValue / info.rate : 0;
 const residualHours = residualHoursRaw > 0 ? Math.round(residualHoursRaw / 0.25) * 0.25 : 0;
 const otherWorkHours = Math.round((aggregatedHours + residualHours) / 0.25) * 0.25;
 if (otherWorkHours > 0) {
 lines.push({
 lp: lines.length + 1,
 label: AGGREGATED_LABEL,
 unit: AGGREGATED_UNIT,
 hours: otherWorkHours,
 amount: otherWorkHours,
 protocolAmount: otherWorkHours,
 unitPrice: Math.floor(info.rate * 10) / 10,
 value: Math.round(otherWorkHours * Number((Math.floor(info.rate * 10) / 10).toFixed(1)) * 100) / 100,
 residualValue: 0
 });
 }

 const totalHours = Math.round(lines.reduce((sum, l) => sum + l.hours, 0) * 100) / 100;
 const totalValue = Math.round(lines.reduce((sum, l) => sum + l.value, 0) * 100) / 100;
 const orderLines = lines.map(line => Object.assign({}, line));
 const protocolLines = lines.map(line => {
   const protocolLine = Object.assign({}, line);
   delete protocolLine.residualValue;
   return protocolLine;
 });

 technicians[techName] = {
 contractNumber: info.contractNumber,
 nip: info.nip,
 rate: info.rate,
 lines: lines,
 orderLines: orderLines,
 protocolLines: protocolLines,
 totalHours: totalHours,
 totalValue: totalValue
 };
 });

 return {
 success: true,
 dateFrom: dateFromStr,
 dateTo: dateToStr,
 includeBreaks: !!includeBreaks,
 technicians: technicians
 };
}

function getApplicationReportFolders_() {
 return { zlec: getProjectFolder_('ZLEC'), pzo: getProjectFolder_('PZO') };
}

function getReportProjectData_() {
 const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data sheet');
 if (!sheet) throw new Error("Brak zakładki 'Data sheet'.");
 return {
   project: sheet.getRange('C2').getDisplayValue().trim(),
   location: sheet.getRange('C3').getDisplayValue().trim()
 };
}

function nextReportDocumentNumber_(folder, prefix, nip, contractNumber, year) {
 const base = String(contractNumber || '').trim().replace(/\//g, '-');
 const files = folder.getFiles();
 let next = 1;
 const pattern = new RegExp('^' + String(nip || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '-' + String(base).replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '-' + prefix + '(\\d{3})-' + year + '\\.(pdf)$', 'i');
 while (files.hasNext()) {
   const match = files.next().getName().match(pattern);
   if (match) next = Math.max(next, Number(match[1]) + 1);
 }
 return prefix + String(next).padStart(3, '0') + '-' + year;
}

function getDocumentFileByNumber_(folder, number) {
 if (!folder || !number) return null;
 const files = folder.getFiles();
 while (files.hasNext()) {
   const file = files.next();
   if (file.getName().includes(number)) return file;
 }
 return null;
}

function shiftReportDate_(dateString, days) {
 const date = new Date(String(dateString || '') + 'T00:00:00');
 if (Number.isNaN(date.getTime())) return '';
 date.setDate(date.getDate() + days);
 return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd.MM.yyyy');
}

function buildReportDocumentHtml_(kind, techName, tech, data, projectData, documentNumber, relatedOrderNumber) {
 const isOrder = kind === 'ZLEC';
 const lines = isOrder ? tech.orderLines : tech.protocolLines;
 const contractNumber = String(tech.contractNumber || '—');
 const fullDocumentNumber = contractNumber + '/' + documentNumber;
 const rows = lines.map((line, index) => {
   const label = escapeHtml_(line.label);
   const unit = escapeHtml_(line.unit);
  const price = Number(line.unitPrice || 0).toFixed(2);
   const rawAmount = isOrder ? (line.amount ?? line.hours) : (line.protocolAmount ?? line.amount ?? line.hours);
   const amount = String(line.unit || '').toLowerCase() === 'h'
     ? escapeHtml_(Number(rawAmount || 0).toFixed(2) + ' h')
     : Number(rawAmount || 0).toFixed(String(line.unit || '').toLowerCase() === 'pis' ? 0 : 2);
   const value = Number(line.value || 0).toFixed(2);
   return `<tr><td>${index + 1}.</td><td>${label}</td><td>${price}</td><td>€</td><td>${unit}</td>${isOrder ? '' : `<td>${amount}</td><td>${value} €</td>`}</tr>`;
 }).join('');
 const total = isOrder ? '' : `<tr class="total"><td colspan="6">RAZEM</td><td>${Number(tech.totalValue || 0).toFixed(2)} €</td></tr>`;
 const title = isOrder
   ? `ZAŁĄCZNIK NR 1 do umowy nr ${escapeHtml_(contractNumber)}`
   : `ZAŁĄCZNIK NR 2 Protokół zdawczo-odbiorczy Usług do umowy nr ${escapeHtml_(contractNumber)}`;
 const documentTitle = isOrder
   ? `ZLECENIE ${escapeHtml_(fullDocumentNumber)}`
   : `PROTOKÓŁ ODBIORU ${escapeHtml_(fullDocumentNumber)} do ZLECENIA ${escapeHtml_(contractNumber + '/' + (relatedOrderNumber || '—'))}`;
 const intro = isOrder
   ? 'Niniejsze Zlecenie stanowi integralny załącznik do Umowy o świadczenie usług zawartej pomiędzy Seavia Marine sp. z o.o. a Wykonawcą wskazanym poniżej i podlega jej postanowieniom.'
   : 'Niniejszy protokół stanowi podstawę do rozliczenia zadań określonych w Umowie o świadczenie usług zawartej pomiędzy Seavia Marine sp. z o.o. a Wykonawcą wskazanym poniżej i podlega jej postanowieniom.';
 const orderScope = 'Zamawiający zleca ogólnopojęte prace związane z wykończeniem wnętrz części mieszkalnej, do głównych zakresów realizowanych prac należeć będą: demontaż oraz montaż izolacji; demontaż / montaż paneli ściennych; montaż wykładziny winylowej; demontaż / montaż sufitów; montaż drzwi B15; transport materiałów oraz inne prace niesklasyfikowane poniżej:';
 const footer = `<div class="footer"><div class="date-place"><b>Data i miejsce:</b> ${isOrder ? shiftReportDate_(data.dateFrom, -1) : shiftReportDate_(data.dateTo, 1)}, Lubieszewo</div><span class="sign">Odbierający (Koordynator projektu)</span><span class="sign">Przekazujący (Wykonawca)</span></div>`;
 const orderExtra = '<p class="intro"><b>Koszty logistyczne (zgodnie z § 10 Umowy):</b><br>Zamawiający dobrowolnie pokrywa koszty mobilizacji/demobilizacji w zakresie określonym w Umowie o Świadczeniu Usług</p><p class="intro"><b>Sprzęt specjalistyczny przekazywany Wykonawcy (jeżeli dotyczy):</b><br>Nie dotyczy</p><p class="intro">Wykonawca potwierdza, że przyjęcie niniejszego Zlecenia jest dobrowolne i że samodzielnie organizuje sposób oraz czas realizacji Usług w ramach wskazanego okna czasowego, zgodnie z § 4 Umowy.</p>';
 const protocolExtra = '<div class="quality"><div class="quality-copy"><b>Ocena jakości i zgodności z wymogami klasyfikatora/armatora:</b><p>□ Prace wykonane prawidłowo, bez zastrzeżeń.</p><p>□ Prace wykonane z zastrzeżeniami (opis):</p></div><div class="remarks"><div></div><div></div><div></div><div></div><div></div></div></div><div class="settlement"><b>Podstawa rozliczenia:</b><br>Niniejszy Protokół stanowi podstawę do wystawienia faktury/rachunku za odebrany zakres Usług, zgodnie z § 6 Umowy. Strony potwierdzają, że wynagrodzenie należne jest za wykonany i odebrany zakres prac.<br><span class="value"><b>Wartość odebranego zakresu prac (EUR netto):</b> ' + Number(tech.totalValue || 0).toFixed(2) + ' €</span></div>';
 return `<!doctype html><html><head><meta charset="utf-8"><style>
@page{size:A4 landscape;margin:7mm}*{box-sizing:border-box}body{font-family:Arial,sans-serif;color:#111;font-size:9pt;margin:0}h1{text-align:center;font-size:14pt;margin:0 0 2.5mm}h2{text-align:center;font-size:11pt;margin:0 0 3mm}.intro{margin:2mm 0;line-height:1.15}.meta{border:.35px solid #555;margin:2mm 0 3mm}.meta div{display:grid;grid-template-columns:62mm minmax(0,1fr);min-height:5.2mm;border-bottom:.35px solid #777;line-height:1.05}.meta div:last-child{border-bottom:0}.meta b,.meta span{min-width:0;padding:.8mm 1.2mm;overflow-wrap:anywhere}.meta b{white-space:normal}.section{background:#c9c9c9;border:.35px solid #555;font-weight:bold;padding:1.2mm}.table{width:100%;border-collapse:collapse}.table th,.table td{border:.35px solid #777;padding:.85mm;text-align:center;line-height:1.05}.table th:nth-child(2),.table td:nth-child(2){text-align:left}.table th{background:#e6e6e6}.total td{font-weight:bold;background:#eee}.quality{display:grid;grid-template-columns:44% 56%;gap:2mm;margin-top:3mm;line-height:1.2;min-height:25mm;align-items:stretch}.quality-copy,.remarks{min-height:25mm}.quality p{margin:1mm 0 0 5mm}.remarks{display:grid;grid-template-rows:repeat(5,1fr);width:100%;padding:0 0 0 2mm;align-items:end}.remarks div{width:100%;height:.35px;border-bottom:.35px dotted #555}.settlement{border:.35px solid #555;margin-top:3mm;padding:1.5mm;line-height:1.15}.value{display:block;margin-top:1mm}.footer{margin-top:3mm;border-top:.35px solid #555;padding-top:1.5mm;text-align:center}.date-place{text-align:left;margin-bottom:0}.sign{display:inline-block;width:38%;margin:14mm 5% 0;padding-top:0;border-top:.35px solid #777;line-height:1.2}
 </style></head><body><h1>${title}</h1><h2>${documentTitle}</h2><p class="intro">${intro}</p>
 <div class="meta"><div><b>Wykonawca</b><span>${escapeHtml_(techName)}</span></div><div><b>NIP</b><span>${escapeHtml_(tech.nip || '—')}</span></div><div><b>Oznaczenie Projektu/Nazwa Jednostki</b><span>${escapeHtml_(projectData.project || '—')}</span></div><div><b>Miejsce realizacji</b><span>${escapeHtml_(projectData.location || '—')}</span></div><div><b>Okno czasowe realizacji (od – do)</b><span>${escapeHtml_(data.dateFrom)} do ${escapeHtml_(data.dateTo)}</span></div></div>
 ${isOrder ? `<p class="intro">${orderScope}</p>` : ''}<div class="section">${isOrder ? 'Zasady wynagrodzenia dla niniejszego zakresu według stawki za jednostkę zakresu:' : 'Zakres Usług zgłoszony do odbioru:'}</div><table class="table"><thead><tr><th>LP.</th><th>NAZWA ZAKRESU</th><th>CENA JEDNOSTKOWA</th><th>WALUTA</th><th>JEDNOSTKA MIARY</th>${isOrder ? '' : '<th>ILOŚĆ</th><th>WARTOŚĆ</th>'}</tr></thead><tbody>${rows}${total}</tbody></table>
 ${isOrder ? orderExtra : protocolExtra}${footer}</body></html>`;
}

function escapeHtml_(value) {
 return String(value || '').replace(/[&<>"']/g, character => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[character]));
}

function getFileByName_(folder, name) {
 const files = folder.getFilesByName(name);
 return files.hasNext() ? files.next() : null;
}

function inspectReportTemplates() {
 const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
 const templateNames = ['Zlecenie-szablon', 'Protokół-szablon'];
 const result = {};

 templateNames.forEach(name => {
   const sheet = spreadsheet.getSheetByName(name);
   if (!sheet) {
     result[name] = { found: false };
     return;
   }

   const dataRange = sheet.getDataRange();
   const values = dataRange.getDisplayValues();
   const nonEmptyCells = [];
   values.forEach((row, rowIndex) => row.forEach((value, columnIndex) => {
     if (String(value || '').trim()) {
       const cell = sheet.getRange(rowIndex + 1, columnIndex + 1);
       nonEmptyCells.push({
         a1: cell.getA1Notation(),
         value: value,
         numberFormat: cell.getNumberFormat(),
         horizontalAlignment: cell.getHorizontalAlignment(),
         verticalAlignment: cell.getVerticalAlignment(),
         wrapStrategy: cell.getWrapStrategy()
       });
     }
   }));

   result[name] = {
     found: true,
     sheetId: sheet.getSheetId(),
     maxRows: sheet.getMaxRows(),
     maxColumns: sheet.getMaxColumns(),
     dataRange: dataRange.getA1Notation(),
     frozenRows: sheet.getFrozenRows(),
     frozenColumns: sheet.getFrozenColumns(),
     mergedRanges: sheet.getDataRange().getMergedRanges().map(range => range.getA1Notation()),
     nonEmptyCells: nonEmptyCells
   };
 });

 Logger.log(JSON.stringify(result, null, 2));
 return result;
}

function writeReportSpreadsheetSheet_(sheet, kind, techName, tech, data, projectData, documentNumber, relatedOrderNumber) {
 const isOrder = kind === 'ZLEC';
 const contractNumber = String(tech.contractNumber || '—');
 const documentTitle = isOrder
   ? 'ZLECENIE ' + contractNumber + '/' + documentNumber
   : 'PROTOKÓŁ ODBIORU ' + contractNumber + '/' + documentNumber + ' do ZLECENIA ' + contractNumber + '/' + (relatedOrderNumber || '—');
 const intro = isOrder
   ? 'Niniejszy Zlecenie stanowi integralny załącznik do Umowy o świadczenie usług zawartej pomiędzy Seavia Marine sp. z o.o. a Wykonawcą wskazanym poniżej i podlega jej postanowieniom.'
   : 'Niniejszy Protokół stanowi podstawę do rozliczenia zadań określonych w Umowie o świadczenie usług zawartej pomiędzy Seavia Marine sp. z o.o. a Wykonawcą wskazanym poniżej i podlega jej postanowieniom.';
 const orderScope = 'Zamawiający zleca ogólnopojęte prace związane z wykończeniem wnętrz części mieszkalnej, do głównych zakresów realizowanych prac należeć będą: demontaż oraz montaż izolacji; demontaż / montaż paneli ściennych; montaż wykładziny winylowej; demontaż / montaż sufitów; montaż drzwi B15; transport materiałów oraz inne prace niesklasyfikowane poniżej:';
 const lines = isOrder ? tech.orderLines : tech.protocolLines;
 const headers = isOrder
   ? ['LP.', 'NAZWA ZAKRESU', 'CENA JEDNOSTKOWA', 'WALUTA', 'JEDNOSTKA MIARY']
   : ['LP.', 'NAZWA ZAKRESU', 'CENA JEDNOSTKOWA', 'WALUTA', 'JEDNOSTKA MIARY', 'ILOŚĆ', 'WARTOŚĆ'];
 const rows = lines.map(line => isOrder
   ? [line.lp, line.label, Number(line.unitPrice || 0), 'EUR', line.unit]
   : [line.lp, line.label, Number(line.unitPrice || 0), 'EUR', line.unit, Number(line.protocolAmount ?? line.amount ?? line.hours ?? 0), Number(line.value || 0)]
 );
 const columnCount = headers.length;
 const title = isOrder
   ? 'ZAŁĄCZNIK NR 1 do umowy nr ' + contractNumber
   : 'ZAŁĄCZNIK NR 2 Protokół zdawczo-odbiorczy Usług do umowy nr ' + contractNumber;
 const metaRows = [
   ['Wykonawca', techName],
   ['NIP', tech.nip || '—'],
   ['Oznaczenie Projektu/Nazwa Jednostki', projectData.project || '—'],
  ['Miejsce realizacji', projectData.location || '—'],
   ['Okno czasowe realizacji (od – do)', data.dateFrom + ' do ' + data.dateTo]
 ];

 sheet.getDataRange().breakApart();
 sheet.clear();
 sheet.setHiddenGridlines(true);
 sheet.getRange(1, 1, 1, columnCount).merge().setValue(title).setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
 sheet.getRange(2, 1, 1, columnCount).merge().setValue(documentTitle).setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
 sheet.getRange(4, 1, 1, columnCount).merge().setValue(intro).setWrap(true);
 sheet.getRange(4, 1, 1, columnCount).setVerticalAlignment('middle');
 sheet.setRowHeight(4, 42);
 sheet.getRange(6, 1, metaRows.length, 1).setValues(metaRows.map(row => [row[0]])).setFontWeight('bold');
 metaRows.forEach((row, index) => sheet.getRange(6 + index, 2, 1, columnCount - 1).merge().setValue(row[1]));
 sheet.getRange(6, 1, metaRows.length, columnCount).setBorder(true, true, true, true, true, true).setWrap(true);
 sheet.getRange(6, 1, metaRows.length, 1).setBackground('#f1f5f9');

 let headerRow = 12;
 if (isOrder) {
   sheet.getRange(headerRow, 1, 1, columnCount).merge().setValue(orderScope).setWrap(true);
   sheet.setRowHeight(headerRow, 55);
   headerRow += 2;
 }
 sheet.getRange(headerRow, 1, 1, columnCount).merge()
   .setValue(isOrder ? 'Zasady wynagrodzenia dla niniejszego zakresu według stawki za jednostkę zakresu:' : 'Zakres Usług zgłoszony do odbioru:')
   .setFontWeight('bold').setBackground('#c9c9c9');
 headerRow += 1;
 sheet.getRange(headerRow, 1, 1, columnCount).setValues([headers]).setFontWeight('bold').setBackground('#e6e6e6').setHorizontalAlignment('center');
 if (rows.length) {
   sheet.getRange(headerRow + 1, 1, rows.length, columnCount).setValues(rows).setBorder(true, true, true, true, true, true).setWrap(true);
   sheet.getRange(headerRow + 1, 3, rows.length, 1).setNumberFormat('0.00');
   if (!isOrder) {
     sheet.getRange(headerRow + 1, columnCount - 1, rows.length, 2).setNumberFormat('0.00');
   }
 }
 const afterTableRow = headerRow + rows.length + 2;
 if (!isOrder) {
   const totalRow = afterTableRow - 1;
   sheet.getRange(totalRow, 1, 1, columnCount - 1).merge().setValue('RAZEM').setFontWeight('bold').setHorizontalAlignment('right').setBackground('#e2e8f0');
   sheet.getRange(totalRow, columnCount).setValue(Number(tech.totalValue || 0)).setNumberFormat('0.00').setFontWeight('bold').setBackground('#e2e8f0');
   sheet.getRange(afterTableRow, 1, 1, columnCount).merge().setValue('Ocena jakości i zgodności z wymogami klasyfikatora/armatora:').setFontWeight('bold');
   sheet.getRange(afterTableRow + 1, 1, 1, columnCount).merge().setValue('□ Prace wykonane prawidłowo, bez zastrzeżeń.');
   sheet.getRange(afterTableRow + 2, 1, 1, columnCount).merge().setValue('□ Prace wykonane z zastrzeżeniami (opis):');
   sheet.getRange(afterTableRow + 3, 1, 5, columnCount).merge().setBorder(false, false, true, false, false, false);
   sheet.getRange(afterTableRow + 9, 1, 1, columnCount).merge().setValue('Podstawa rozliczenia:').setFontWeight('bold');
   sheet.getRange(afterTableRow + 10, 1, 1, columnCount).merge().setValue('Niniejszy Protokół stanowi podstawę do wystawienia faktury/rachunku za odebrany zakres Usług, zgodnie z § 6 Umowy. Strony potwierdzają, że wynagrodzenie należne jest za wykonany i odebrany zakres prac.').setWrap(true);
   sheet.getRange(afterTableRow + 11, 1, 1, columnCount).merge().setValue('Wartość odebranego zakresu prac (EUR netto): ' + Number(tech.totalValue || 0).toFixed(2) + ' €').setFontWeight('bold');
 } else {
   sheet.getRange(afterTableRow, 1, 1, columnCount).merge().setValue('Koszty logistyczne (zgodnie z § 10 Umowy):').setFontWeight('bold');
   sheet.getRange(afterTableRow + 1, 1, 1, columnCount).merge().setValue('Zamawiający dobrowolnie pokrywa koszty mobilizacji/demobilizacji w zakresie określonym w Umowie o Świadczeniu Usług').setWrap(true);
   sheet.getRange(afterTableRow + 3, 1, 1, columnCount).merge().setValue('Sprzęt specjalistyczny przekazywany Wykonawcy (jeżeli dotyczy):').setFontWeight('bold');
   sheet.getRange(afterTableRow + 4, 1, 1, columnCount).merge().setValue('Nie dotyczy');
   sheet.getRange(afterTableRow + 6, 1, 1, columnCount).merge().setValue('Wykonawca potwierdza, że przyjęcie niniejszego Zlecenia jest dobrowolne i że samodzielnie organizuje sposób oraz czas realizacji Usług w ramach wskazanego okna czasowego, zgodnie z § 4 Umowy.').setWrap(true);
 }
 const footerRow = isOrder ? afterTableRow + 9 : afterTableRow + 14;
 const documentDate = isOrder ? shiftReportDate_(data.dateFrom, -1) : shiftReportDate_(data.dateTo, 1);
 sheet.getRange(footerRow, 1, 1, columnCount).merge().setValue('Data i miejsce: ' + documentDate + ', Lubieszewo').setFontWeight('bold');
 sheet.getRange(footerRow + 2, 1, 1, Math.floor(columnCount / 2)).merge().setValue('Odbierający (Koordynator projektu)').setHorizontalAlignment('center').setBorder(true, false, false, false, false, false);
 sheet.getRange(footerRow + 2, Math.floor(columnCount / 2) + 1, 1, columnCount - Math.floor(columnCount / 2)).merge().setValue('Przekazujący (Wykonawca)').setHorizontalAlignment('center').setBorder(true, false, false, false, false, false);
 sheet.setColumnWidth(1, 55);
 sheet.setColumnWidth(2, 420);
 for (let column = 3; column <= headers.length; column++) sheet.setColumnWidth(column, 130);
}

function createReportSpreadsheet_(folder, fileName, techName, tech, data, projectData, zlecNumber, pzoNumber) {
 const existingFile = getFileByName_(folder, fileName);
 const spreadsheet = existingFile
   ? SpreadsheetApp.openById(existingFile.getId())
   : SpreadsheetApp.create(fileName);
 if (!existingFile) DriveApp.getFileById(spreadsheet.getId()).moveTo(folder);
 const orderSheet = spreadsheet.getSheets()[0];
 orderSheet.setName('Zlecenie');
 const protocolSheet = spreadsheet.getSheetByName('Protokół') || spreadsheet.insertSheet('Protokół');
 writeReportSpreadsheetSheet_(orderSheet, 'ZLEC', techName, tech, data, projectData, zlecNumber, pzoNumber);
 writeReportSpreadsheetSheet_(protocolSheet, 'PZO', techName, tech, data, projectData, pzoNumber, zlecNumber);
 SpreadsheetApp.flush();
 return spreadsheet.getUrl();
}

function generateZleceniaIProtokolyFiles(dateFromStr, dateToStr, includeBreaks) {
 const data = generateZleceniaIProtokolyData(dateFromStr, dateToStr, includeBreaks);
 if (!data || !data.success) return data;
 const folders = getApplicationReportFolders_();
 const projectData = getReportProjectData_();
 const year = String(new Date(dateToStr + 'T00:00:00').getFullYear());
 const links = [];
 Object.keys(data.technicians).forEach(techName => {
   const tech = data.technicians[techName];
   const contractNumber = String(tech.contractNumber || '').trim();
   const nip = String(tech.nip || '').trim();
   if (!contractNumber || !nip) throw new Error('Brak numeru umowy lub NIP dla technika: ' + techName);

   const zlecStatus = getDocumentStatusForRange_('Zlecenia', techName, dateFromStr, dateToStr);
   if (zlecStatus.status === 'conflict') {
     throw new Error('Zazębianie zakresów dla zlecenia: ' + techName + ' (' + zlecStatus.existing[0] + ' - ' + zlecStatus.existing[1] + '). Wybierz inny okres dat.');
   }

   const pzoStatus = getDocumentStatusForRange_('Protokoły', techName, dateFromStr, dateToStr);
   if (pzoStatus.status === 'conflict') {
     throw new Error('Zazębianie zakresów dla protokołu: ' + techName + ' (' + pzoStatus.existing[0] + ' - ' + pzoStatus.existing[1] + '). Wybierz inny okres dat.');
   }

   let zlecNumber = zlecStatus.status === 'existing' ? zlecStatus.documentNumber : getNextRegistryDocumentNumber_('Zlecenia', techName, 'ZL', year);
   let pzoNumber = pzoStatus.status === 'existing' ? pzoStatus.documentNumber : getNextRegistryDocumentNumber_('Protokoły', techName, 'P', year);

   if (zlecStatus.status !== 'existing') {
     registerDocumentRange_('Zlecenia', techName, dateFromStr, dateToStr, zlecNumber);
   }
   if (pzoStatus.status !== 'existing') {
     registerDocumentRange_('Protokoły', techName, dateFromStr, dateToStr, pzoNumber, zlecNumber);
   }

   const existingZlecFile = getDocumentFileByNumber_(folders.zlec, zlecNumber);
   const existingPzoFile = getDocumentFileByNumber_(folders.pzo, pzoNumber);
   const escapedNip = String(nip).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
   const contractFilePart = contractNumber.replace(new RegExp('^' + escapedNip + '/', 'i'), '').replace(/\//g, '-');
   const sheetFileName = nip + '-' + contractFilePart + '-' + zlecNumber + '-' + pzoNumber;
   if (zlecStatus.status === 'existing' && pzoStatus.status === 'existing' && existingZlecFile && existingPzoFile) {
     const arkusz = createReportSpreadsheet_(folders.zlec, sheetFileName, techName, tech, data, projectData, zlecNumber, pzoNumber);
     links.push({ technician: techName, zlecenie: existingZlecFile.getUrl(), protokol: existingPzoFile.getUrl(), arkusz: arkusz, reused: true });
     return;
   }

   const zlecFileName = nip + '-' + contractFilePart + '-' + zlecNumber + '.pdf';
   const pzoFileName = nip + '-' + contractFilePart + '-' + pzoNumber + '.pdf';
   const zlecFile = folders.zlec.createFile(Utilities.newBlob(buildReportDocumentHtml_('ZLEC', techName, tech, data, projectData, zlecNumber), 'text/html', zlecNumber + '.html').getAs(MimeType.PDF).setName(zlecFileName));
   const pzoFile = folders.pzo.createFile(Utilities.newBlob(buildReportDocumentHtml_('PZO', techName, tech, data, projectData, pzoNumber, zlecNumber), 'text/html', pzoNumber + '.html').getAs(MimeType.PDF).setName(pzoFileName));
  const arkusz = createReportSpreadsheet_(folders.zlec, sheetFileName, techName, tech, data, projectData, zlecNumber, pzoNumber);
  links.push({ technician: techName, zlecenie: zlecFile.getUrl(), protokol: pzoFile.getUrl(), arkusz: arkusz, reused: false });
 });
 return { success: true, files: links };
}

function generateStyledExcelReport(reportData, reportType) {
 try {
 const isWeekly = reportType === 'weekly';
 const baseName = "Candido_" + (isWeekly ? reportData.weekInfo.weekNumber : 'Annual') + "_Report";
 const tempFileName = baseName + "_TEMP_" + Date.now();
 const targetFolder = getProjectFolder_('Raporty');

 const spreadsheet = SpreadsheetApp.create(tempFileName);
 const tempFile = DriveApp.getFileById(spreadsheet.getId());
 tempFile.moveTo(targetFolder);

 const sheet = spreadsheet.getSheets()[0];
 sheet.setName(isWeekly ? "Weekly Report" : "Annual Report");
 const headerColor = "#f1f5f9";
 const totalHeaderColor = "#e2e8f0";
 const seaviaBlue = "#0f3460";
 const lightGreen = "#dcfce7";
 const green = "#bbf7d0";
 const borderColor = "#cbd5e1";
 let currentRow = 1;

 sheet.getRange(currentRow, 1, 1, 2).setValues([["Project name:", "REV OCEAN 19577"]]).setFontWeight("bold");
 sheet.getRange(currentRow, 2).setFontColor(seaviaBlue).setHorizontalAlignment("right");
 currentRow++;

 if (isWeekly) {
 sheet.getRange(currentRow, 1, 1, 2).setValues([["Week no.:", reportData.weekInfo.weekNumber]]).setFontWeight("bold");
 sheet.getRange(currentRow, 2).setFontColor(seaviaBlue).setHorizontalAlignment("right");
 currentRow++;
 sheet.getRange(currentRow, 1, 1, 2).setValues([["Date range:", reportData.weekInfo.start + " / " + reportData.weekInfo.end]]).setFontWeight("bold");
 sheet.getRange(currentRow, 2).setHorizontalAlignment("right");
 currentRow++;
 } else {
 sheet.getRange(currentRow, 1, 1, 2).setValues([["Data generacji:", new Date().toLocaleDateString('pl-PL')]]).setFontWeight("bold");
 sheet.getRange(currentRow, 2).setHorizontalAlignment("right");
 currentRow++;
 }

 sheet.getRange(currentRow, 1, 1, 2).setValues([["Paid Breaks:", reportData.settings.includeBreaks ? "YES" : "NO"]]).setFontWeight("bold");
 sheet.getRange(currentRow, 2).setFontColor(reportData.settings.includeBreaks ? "#166534" : "#94a3b8").setHorizontalAlignment("right");
 currentRow += 2;

 Object.keys(reportData.decks).sort().forEach(deckName => {
 let deckData = reportData.decks[deckName];
 if (!deckData.tasks || Object.keys(deckData.tasks).length === 0) return;

 sheet.getRange(currentRow, 1).setValue(isWeekly ? "WEEKLY SUM " + deckName : "ANNUAL TIME REPORT - " + deckName).setFontWeight("bold").setFontColor(seaviaBlue);
 currentRow++;

 let headers = ["No.", "Tasks", "Total"];
 let columnCount;
 if (isWeekly) {
 headers.push('Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun');
 columnCount = 10;
 } else {
 let sortedWeeks = reportData.allWeeks.sort((a, b) => a - b);
 sortedWeeks.forEach(w => headers.push("W" + w));
 columnCount = 3 + sortedWeeks.length;
 }

 if (columnCount > 0) {
 let headerRange = sheet.getRange(currentRow, 1, 1, columnCount);
 headerRange.setValues([headers]).setBackground(headerColor).setFontWeight("bold");
 sheet.getRange(currentRow, 3).setBackground(totalHeaderColor);
 }
 currentRow++;

 const renderRow = (taskName, index) => {
 let rowData = [index, taskName];
 if (isWeekly) {
 let hours = (deckData.tasks && deckData.tasks[taskName]) ? deckData.tasks[taskName] : Array(7).fill(0);
 let total = hours.reduce((a, b) => a + b, 0);
 rowData.push(total > 0 ? total : 0, ...hours.map(h => h > 0 ? h : 0));
 } else {
 let sortedWeeks = reportData.allWeeks.sort((a, b) => a - b);
 let task = (deckData.tasks && deckData.tasks[taskName]) ? deckData.tasks[taskName] : { total: 0, weeklyHours: {} };
 rowData.push(task.total > 0 ? task.total : 0);
 sortedWeeks.forEach(w => {
 let h = task.weeklyHours[w] || 0;
 rowData.push(h > 0 ? h : 0);
 });
 }

 if (rowData.length > 2) {
 let rowRange = sheet.getRange(currentRow, 1, 1, rowData.length);
 rowRange.setValues([rowData]).setNumberFormat("0.0");
 sheet.getRange(currentRow, 1, 1, 2).setNumberFormat("0");
 sheet.getRange(currentRow, 2).setHorizontalAlignment("left");

 let totalCell = sheet.getRange(currentRow, 3);
 if (rowData[2] > 0) totalCell.setBackground(green).setFontWeight("bold"); else totalCell.setBackground(totalHeaderColor);

 for (let i = 4; i <= rowData.length; i++) {
 let cell = sheet.getRange(currentRow, i);
 if (cell.getValue() > 0) cell.setBackground(lightGreen).setFontWeight("bold");
 }
 }
 currentRow++;
 };

 if (deckName !== "TRAVEL" && (isWeekly || reportData.allTasks.includes("Supervision"))) {
 renderRow("Supervision", 0);
 }

 let taskIndex = 1;
 let filteredTasks = reportData.allTasks;
 if (!reportData.settings.includeBreaks) {
 filteredTasks = filteredTasks.filter(t => !t.toLowerCase().includes('break') && !t.toLowerCase().includes('przerwa'));
 }

 if (deckName === "TRAVEL") {
 filteredTasks = filteredTasks.filter(t => t.toLowerCase() === "travel");
 } else {
 filteredTasks = filteredTasks.filter(t => t.toLowerCase() !== "travel");
 }

 filteredTasks.forEach(taskName => {
 if (taskName !== "Supervision") {
 renderRow(taskName, taskIndex++);
 }
 });

 let renderedRowCount = taskIndex;
 if (deckName !== "TRAVEL" && (isWeekly || reportData.allTasks.includes("Supervision"))) {
 renderedRowCount += 1;
 } else {
 renderedRowCount = filteredTasks.length;
 }
 if (renderedRowCount > 0) {
 let tableStartRow = currentRow - renderedRowCount - 1; 
let numRows = renderedRowCount + 1; 
if (tableStartRow > 0 && numRows > 0 && columnCount > 0) {
 let tableRange = sheet.getRange(tableStartRow, 1, numRows, columnCount);
 tableRange.setBorder(true, true, true, true, true, true, borderColor, SpreadsheetApp.BorderStyle.SOLID);
 }
 }
 sheet.setColumnWidth(2, 250);
 currentRow += 2;
 });

 if (reportData.extraTasks && reportData.extraTasks.length > 0) {
 sheet.getRange(currentRow, 1).setValue("PRACE DODATKOWE (EXTRA JOBS)").setFontWeight("bold").setFontColor("#dc2626");
 currentRow++;

 let headers = ["No.", "Tasks", "Total"];
 let columnCount;
 if (isWeekly) {
 headers.push('Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun');
 columnCount = 10;
 } else {
 let sortedWeeks = reportData.allWeeks.sort((a, b) => a - b);
 sortedWeeks.forEach(w => headers.push("W" + w));
 columnCount = 3 + sortedWeeks.length;
 }

 if (columnCount > 0) {
 let headerRange = sheet.getRange(currentRow, 1, 1, columnCount);
 headerRange.setValues([headers]).setBackground(headerColor).setFontWeight("bold");
 sheet.getRange(currentRow, 3).setBackground(totalHeaderColor);
 }
 currentRow++;

 let extraTaskIndex = 1;
 reportData.extraTasks.forEach(taskName => {
 let rowData = [extraTaskIndex, taskName];
 if (isWeekly) {
 let hours = reportData.extraDeck.tasks[taskName] || Array(7).fill(0);
 let total = hours.reduce((a, b) => a + b, 0);
 rowData.push(total > 0 ? total : 0, ...hours.map(h => h > 0 ? h : 0));
 } else {
 let sortedWeeks = reportData.allWeeks.sort((a, b) => a - b);
 let task = reportData.extraDeck.tasks[taskName] || { total: 0, weeklyHours: {} };
 rowData.push(task.total > 0 ? task.total : 0);
 sortedWeeks.forEach(w => {
 let h = task.weeklyHours[w] || 0;
 rowData.push(h > 0 ? h : 0);
 });
 }

 if (rowData.length > 2) {
 let rowRange = sheet.getRange(currentRow, 1, 1, rowData.length);
 rowRange.setValues([rowData]).setNumberFormat("0.0");
 sheet.getRange(currentRow, 1, 1, 2).setNumberFormat("0");
 sheet.getRange(currentRow, 2).setHorizontalAlignment("left");
 rowRange.setFontColor("#dc2626"); 

let totalCell = sheet.getRange(currentRow, 3);
 if (rowData[2] > 0) totalCell.setBackground(green).setFontWeight("bold"); else totalCell.setBackground(totalHeaderColor);

 for (let i = 4; i <= rowData.length; i++) {
 let cell = sheet.getRange(currentRow, i);
 if (cell.getValue() > 0) cell.setBackground(lightGreen).setFontWeight("bold");
 }
 }
 currentRow++;
 extraTaskIndex++;
 });

 let renderedRowCount = reportData.extraTasks.length;
 if (renderedRowCount > 0) {
 let tableStartRow = currentRow - renderedRowCount - 1;
 let numRows = renderedRowCount + 1;
 if (tableStartRow > 0 && numRows > 0 && columnCount > 0) {
 let tableRange = sheet.getRange(tableStartRow, 1, numRows, columnCount);
 tableRange.setBorder(true, true, true, true, true, true, borderColor, SpreadsheetApp.BorderStyle.SOLID);
 }
 }
 sheet.setColumnWidth(2, 250);
 currentRow += 2;
 }

 SpreadsheetApp.flush();

 const url = "https://docs.google.com/spreadsheets/d/" + spreadsheet.getId() + "/export?format=xlsx";
 const token = ScriptApp.getOAuthToken();
 const response = UrlFetchApp.fetch(url, {
 headers: { 'Authorization': 'Bearer ' + token },
 muteHttpExceptions: true
 });

 if (response.getResponseCode() !== 200) {
 throw new Error("Błąd podczas konwersji arkusza do formatu Excel (.xlsx): " + response.getContentText());
 }

 const excelBlob = response.getBlob().setName(`${baseName}.xlsx`);
 const excelFile = targetFolder.createFile(excelBlob);
 tempFile.setTrashed(true);

 const downloadUrl = `https://drive.google.com/uc?export=download&id=${excelFile.getId()}`;
 return {
 fileUrl: downloadUrl,
 fileName: excelFile.getName()
 };
 } catch (e) {
 Logger.log("Krytyczny błąd w generateStyledExcelReport: " + e.message + " Stack: " + e.stack);
 return { error: "Wystąpił błąd po stronie serwera: " + e.message };
 }
}

function deleteTeam(teamName) {
 if (!teamName || teamName.trim() === "") {
 return { success: false, message: "Nazwa drużyny nie może być pusta!" };
 }
 try {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const dataSheet = ss.getSheetByName('Data sheet');
 if (!dataSheet) return { success: false, message: "Brak zakładki 'Data sheet'." };
 const lastRow = dataSheet.getLastRow();
 let updatedCount = 0;
 if (lastRow >= 9) {
 const range = dataSheet.getRange(9, 1, lastRow - 8, 1);
 const values = range.getValues();
 for (let i = 0; i < values.length; i++) {
 if (String(values[i][0]).trim().toLowerCase() === teamName.trim().toLowerCase()) {
 dataSheet.getRange(9 + i, 1).setValue(""); 
updatedCount++;
 }
 }
 }
 return { success: true, message: `Pomyślnie usunięto drużynę ${teamName} (odpięto ${updatedCount} techników).`, updatedCount: updatedCount };
 } catch (error) {
 return { success: false, message: "Błąd podczas usuwania drużyny: " + error.toString() };
 }
}

function renameTeam(oldName, newName) {
 if (!oldName || !newName || newName.trim() === "") {
 return { success: false, message: "Nazwy drużyn nie mogą być puste!" };
 }
 try {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const dataSheet = ss.getSheetByName('Data sheet');
 if (!dataSheet) return { success: false, message: "Brak zakładki 'Data sheet'." };
 const lastRow = dataSheet.getLastRow();
 let updatedCount = 0;
 if (lastRow >= 9) {
 const range = dataSheet.getRange(9, 1, lastRow - 8, 1);
 const values = range.getValues();
 for (let i = 0; i < values.length; i++) {
 if (String(values[i][0]).trim().toLowerCase() === oldName.trim().toLowerCase()) {
 dataSheet.getRange(9 + i, 1).setValue(newName.trim());
 updatedCount++;
 }
 }
 }
 return { success: true, message: `Pomyślnie zmieniono nazwę drużyny z ${oldName} na ${newName} (zaktualizowano ${updatedCount} techników).`, updatedCount: updatedCount };
 } catch (error) {
 return { success: false, message: "Błąd podczas zmiany nazwy drużyny: " + error.toString() };
 }
}

function getCostFolderUrl() {
 return getProjectFolder_('Rachunki').getUrl();
}

function saveReportPdf(base64Data, fileName, reportType) {
 const folderName = reportType === 'cards' ? 'protokoły' : reportType === 'report' ? 'Raporty' : null;
 if (!folderName) throw new Error('Nieznany rodzaj raportu PDF.');
 const match = /^data:application\/pdf(?:;[^,]*)?;base64,(.+)$/i.exec(base64Data);
 if (!match) throw new Error('Nieprawidłowe dane pliku PDF.');
 const folder = getProjectFolder_(folderName);
 const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(match[1]), 'application/pdf', fileName));
 if (!file || !file.getId()) throw new Error('Nie potwierdzono zapisu pliku PDF na Dysku.');
 return { success: true, fileId: file.getId() };
}

function uploadCostFile(base64Data, fileName, costName, costDate, fileIndex = 0, totalFiles = 1) {
 try {
 const folder = getProjectFolder_('Rachunki');
 
const parts = base64Data.split(',');
 const contentType = parts[0].match(/:(.*?);/)[1];
 const rawBase64 = parts[1];
 const decoded = Utilities.base64Decode(rawBase64);
 
// 1. Wyciąganie oryginalnego rozszerzenia (np. .jpg, .pdf)
 let extension = "";
 const lastDotIndex = fileName.lastIndexOf('.');
 if (lastDotIndex !== -1) {
 extension = fileName.substring(lastDotIndex);
 } else {
 if (contentType.includes('pdf')) extension = '.pdf';
 else if (contentType.includes('jpeg') || contentType.includes('jpg')) extension = '.jpg';
 else if (contentType.includes('png')) extension = '.png';
 }
 
// 2. Czyszczenie nazwy (usuwamy tylko niedozwolone znaki w nazwach plików, pozwalamy na spacje)
 const cleanCostName = costName.replace(/[\/\\?%*:|"<>]/g, '').trim();
 
// 3. Dodanie numeracji, jeśli w jednej paczce jest więcej niż 1 plik 
let suffix = "";
 if (totalFiles > 1) {
 suffix = ` - plik ${fileIndex + 1}`;
 }
 
// 4. Budowa finalnej nowej nazwy
 const newFileName = `${costDate} ${cleanCostName}${suffix}${extension}`;
 
const blob = Utilities.newBlob(decoded, contentType, newFileName);
 const file = folder.createFile(blob);
 
const ss = SpreadsheetApp.getActiveSpreadsheet();
 let sheetKoszty = ss.getSheetByName('Koszty');
 if (!sheetKoszty) {
 sheetKoszty = ss.insertSheet('Koszty');
 sheetKoszty.appendRow(["Data dodania", "Data kosztu", "Nazwa kosztu", "Link do pliku", "Użytkownik"]);
 sheetKoszty.getRange(1, 1, 1, 5).setFontWeight("bold");
 }
 const userEmail = Session.getActiveUser().getEmail() || "Nieznany użytkownik";
 // Zapis do arkusza Google z zachowaniem oryginalnie wpisanej nazwy
 sheetKoszty.appendRow([new Date(), costDate, costName, file.getUrl(), userEmail]);
 
return { success: true, message: "Koszt został pomyślnie zapisany w bazie i folderze!", fileUrl: file.getUrl() };
 } catch (e) {
 Logger.log("Błąd w uploadCostFile: " + e.message);
 return { success: false, message: "Błąd serwera podczas zapisywania pliku: " + e.toString() };
 }
}

// =======================================================================
// NOWE: SYNCHRONIZACJA Z REV OCEAN - TIMESHEET (W tle)
// =======================================================================
function syncWeeklyTimesheet(reportData) {
 if (!reportData || !reportData.weekInfo || !reportData.technicians) return { success: false, message: "Brak danych." };

 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const sheetTime = ss.getSheetByName('REV OCEAN - timesheet');
 if (!sheetTime) return { success: false, message: "Brak zakładki 'REV OCEAN - timesheet'." };

 // Wymuszamy pobranie pełnego zakresu
 const timeData = sheetTime.getDataRange().getValues();
 if (timeData.length < 3) return { success: false, message: "Pusty arkusz timesheet." };

 const timezone = ss.getSpreadsheetTimeZone();

 // 1. Zabezpieczone mapowanie dat (nagłówki) z Wiersza 3
 let dateToColMap = {};
 if (timeData[2]) {
 for (let c = 3; c < timeData[2].length; c++) {
 let rawDate = timeData[2][c];
 let dStr = "";
 
if (rawDate instanceof Date) {
 dStr = Utilities.formatDate(rawDate, timezone, "yyyy-MM-dd");
 } else if (typeof rawDate === 'string') {
 let parsed = parseDateToYMD(rawDate);
 if (parsed) dStr = parsed;
 }

 if (dStr) {
 dateToColMap[dStr] = c + 1; // +1 dla metody getRange()
 }
 }
 }

 // 2. Bezpieczne mapowanie wszystkich istniejących techników z kolumny B!
 let techToRowMap = {};
 
// Skanujemy CAŁY plik, patrząc TYLKO NA KOLUMNĘ B (timeData[r][1])
 for (let r = 3; r < timeData.length; r++) {
 if (!timeData[r]) continue;

 let tName = String(timeData[r][1] || "").trim(); // TYLKO KOLUMNA B
 
if (tName) {
 let normalized = tName.replace(/\s+/g, '').toLowerCase();
 techToRowMap[normalized] = r + 1;
 }
 }

 let updatesCount = 0;
 let addedTechs = 0;

 // Najpierw: czyścimy komórki z analizowanego zakresu dat dla WSZYSTKICH techników już
 // obecnych w timesheet (nie tylko tych z reportData) - inaczej godziny technika, który
 // po korekcie zniknął z karty pracy (np. usunięto mu zadanie), zostałyby w arkuszu na stałe.
 Object.keys(techToRowMap).forEach(normalizedTechKey => {
 let rowIdx = techToRowMap[normalizedTechKey];
 reportData.weekInfo.dates.forEach(dDate => {
 let colIdx = dateToColMap[dDate];
 if (colIdx) {
 let cell = sheetTime.getRange(rowIdx, colIdx);
 cell.setValue("");
 cell.setBackground(null);
 }
 });
 });


 // 3. Główna pętla wpisująca godziny i ewentualnie nowe osoby na czysto
 Object.keys(reportData.technicians).forEach(techName => {
 
let normalizedTechKey = String(techName).replace(/\s+/g, '').toLowerCase();
 let rowIdx = techToRowMap[normalizedTechKey];

 // Jeśli technika nie ma na liście w kolumnie B - dopisujemy!
 if (!rowIdx) {
 let targetSheetLastRow = sheetTime.getLastRow();
 
// Szukamy pierwszego pustego miejsca od dołu patrząc WYŁĄCZNIE w kolumnę B.
 let insertRow = targetSheetLastRow + 1;
 const columnBValues = sheetTime.getRange("B1:B" + targetSheetLastRow).getValues();
 
for(let i = columnBValues.length - 1; i >= 0; i--) {
 if (String(columnBValues[i][0]).trim() !== "") {
 insertRow = i + 2; 
break;
 }
 }

 // Zapis do nowej linii - ZAWSZE w Kolumnie B
 sheetTime.getRange(insertRow, 1).setValue(insertRow - 3); 
sheetTime.getRange(insertRow, 2).setValue(techName.toUpperCase()); 
sheetTime.getRange(insertRow, 2).setFontWeight("bold");

 techToRowMap[normalizedTechKey] = insertRow; 
rowIdx = insertRow;
 addedTechs++;
 }

 let techObj = reportData.technicians[techName];

 // Iteracja po 7 dniach analizowanego tygodnia
 techObj.days.forEach((tasksArray, dayIndex) => {
 if (tasksArray.length === 0) return;

 let dDate = reportData.weekInfo.dates[dayIndex];
 let colIdx = dateToColMap[dDate];

 if (!colIdx) return; 

let dailyTotalHours = 0;
 let hasTravel = false;

 tasksArray.forEach(t => {
 dailyTotalHours += t.hours;
 if (t.task.toLowerCase().includes("travel")) {
 hasTravel = true;
 }
 });

 if (dailyTotalHours > 0) {
 let cell = sheetTime.getRange(rowIdx, colIdx);
 cell.setValue(dailyTotalHours).setHorizontalAlignment("center");
 
// Oznaczanie Travel
 if (hasTravel) {
 cell.setBackground("#FFF2CC");
 }
 updatesCount++;
 }
 });
 });

 // Wymuszamy, żeby Google od razu to narysował w bazie
 SpreadsheetApp.flush(); 

return { success: true, updates: updatesCount, newTechs: addedTechs };
}

// =======================================================================
// EKSPORT ZAKŁADKI REV OCEAN DO EXCELA
// =======================================================================
function exportRevOceanTimesheet() {
 try {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const sheet = ss.getSheetByName('REV OCEAN - timesheet');
 if (!sheet) return { error: "Brak zakładki 'REV OCEAN - timesheet' w bazie." };
 const targetFolder = getProjectFolder_('TIME SHEET');

 const tempFileName = "Eksport_Timesheet_" + new Date().getTime();
 const tempSs = SpreadsheetApp.create(tempFileName);
 
sheet.copyTo(tempSs).setName("REV OCEAN - timesheet");
 
const defaultSheet = tempSs.getSheets()[0];
 if(tempSs.getSheets().length > 1) {
 tempSs.deleteSheet(defaultSheet);
 }

 const tempFile = DriveApp.getFileById(tempSs.getId());
 tempFile.moveTo(targetFolder);

 SpreadsheetApp.flush();

 const url = "https://docs.google.com/spreadsheets/d/" + tempSs.getId() + "/export?format=xlsx";
 const token = ScriptApp.getOAuthToken();
 const response = UrlFetchApp.fetch(url, {
 headers: { 'Authorization': 'Bearer ' + token },
 muteHttpExceptions: true
 });

 if (response.getResponseCode() !== 200) {
 throw new Error("Błąd podczas konwersji arkusza.");
 }

 const dateFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
 const excelBlob = response.getBlob().setName(`Candido_Timesheet_${dateFormatted}.xlsx`);
 
const finalFile = targetFolder.createFile(excelBlob);
 tempFile.setTrashed(true);

 return { success: true, fileUrl: `https://drive.google.com/uc?export=download&id=${finalFile.getId()}` };
 } catch (e) {
 return { error: e.message };
 }
}

// =======================================================================
// POBIERANIE DANYCH TECHNIKÓW Z PLIKU EXCEL
// =======================================================================
function syncTechniciansFromExcel() {
 try {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const fileId = ss.getId();
 const sheetFile = DriveApp.getFileById(fileId);
 const parents = sheetFile.getParents();
 if (!parents.hasNext()) {
 return { success: false, message: "Arkusz główny nie znajduje się w żadnym folderze." };
 }
 const folder = parents.next();
 
let excelFile = null;
 let files = folder.getFilesByName("Dane wykonawców.xlsx");
 if (files.hasNext()) excelFile = files.next();
 if (!excelFile) {
 files = folder.getFilesByName("Dane wykonawców.xlsm"); // Poprawiono rozszerzenie
 if (files.hasNext()) excelFile = files.next();
 }
 
if (!excelFile) {
 return { success: false, message: "Nie znaleziono pliku 'Dane wykonawców.xlsx' (ani .xlsm) w tym samym folderze." };
 }
 
// Tymczasowa konwersja do Google Sheets (Wymaga włączonego Drive API w usługach "Advanced Services" Apps Script)
 let tempSpreadsheetId;
 try {
 const blob = excelFile.getBlob();
 const config = {
 name: "Temp_Dane_Wykonawcow", // W Drive API v3 używamy 'name' zamiast 'title'
 mimeType: MimeType.GOOGLE_SHEETS
 };
 // Zmiana z v2 (insert) na v3 (create)
 const convertedFile = Drive.Files.create(config, blob);
 tempSpreadsheetId = convertedFile.id;
 } catch (apiError) {
 return { success: false, message: "Błąd otwierania pliku. Upewnij się, że usługa 'Drive API' jest włączona w sekcji 'Usługi' (Advanced Services) w lewym panelu Apps Script. Szczegóły: " + apiError.toString() };
 }
 
const tempSs = SpreadsheetApp.openById(tempSpreadsheetId);
 const sheet = tempSs.getSheets()[0];
 const data = sheet.getDataRange().getValues();
 
// Usunięcie tymczasowego pliku
 DriveApp.getFileById(tempSpreadsheetId).setTrashed(true);
 
if (data.length < 3) {
 return { success: false, message: "Plik danych jest pusty lub posiada za mało wierszy (wymagane minimum 3)." };
 }
 
// Kolumny mogą się przesuwać w pliku źródłowym (np. dodanie kolumny "Numer umowy"),
 // więc szukamy ich po nagłówkach zamiast polegać na sztywnych indeksach.
 const findHeaderCol = (headerRow, keywords) => {
 for (let c = 0; c < headerRow.length; c++) {
 const cell = String(headerRow[c] || "").trim().toLowerCase();
 if (!cell) continue;
 if (keywords.some(k => cell.includes(k))) return c;
 }
 return -1;
 };
 
let headerRowIdx = -1;
 let nameIdx = -1, nipIdx = -1, rateIdx = -1, phoneIdx = -1, emailIdx = -1, contractNumIdx = -1;
 
// Nagłówki mogą znajdować się w jednym z pierwszych kilku wierszy pliku.
 for (let r = 0; r < Math.min(5, data.length); r++) {
 const row = data[r];
 const tryNip = findHeaderCol(row, ['nip']);
 const tryRate = findHeaderCol(row, ['stawka']);
 const tryPhone = findHeaderCol(row, ['tel']);
 const tryEmail = findHeaderCol(row, ['mail']);
 const tryName = findHeaderCol(row, ['nazwisko']);
 const tryContractNum = findHeaderCol(row, ['numer umowy']);
 if (tryNip >= 0 && tryRate >= 0 && tryPhone >= 0 && tryEmail >= 0) {
 headerRowIdx = r;
 nameIdx = tryName;
 nipIdx = tryNip;
 rateIdx = tryRate;
 phoneIdx = tryPhone;
 emailIdx = tryEmail;
 contractNumIdx = tryContractNum;
 break;
 }
 }
 
if (headerRowIdx === -1) {
 return { success: false, message: "Nie udało się rozpoznać nagłówków kolumn (NIP, Stawka, Telefon, E-mail) w pliku źródłowym. Sprawdź, czy struktura pliku nie została zmieniona." };
 }
 
if (nameIdx < 0) nameIdx = findHeaderCol(data[headerRowIdx], ['imię', 'imie']);
 if (nameIdx < 0) nameIdx = 1; // domyślnie kolumna B, zgodnie z dotychczasową strukturą
 
const startRowIdx = headerRowIdx + 1;
 
const dataSheet = ss.getSheetByName('Data sheet');
 if (!dataSheet) return { success: false, message: "Brak zakładki 'Data sheet' w głównym arkuszu." };
 
// Pobranie bieżącej listy z zakłądki 'Data sheet' żeby wiedzieć czy dodawać czy edytować
 const lastRow = Math.max(dataSheet.getLastRow(), 8);
 let currentData = [];
 if (lastRow >= 9) {
 // Pobieramy całą tabelę od kolumny A (1) do H (8), aby uwzględnić Stawkę (F), NIP (G) i Numer umowy (H)
 currentData = dataSheet.getRange(9, 1, lastRow - 8, 8).getValues();
 }
 
const nameToIndex = {};
 currentData.forEach((row, idx) => {
 const n = String(row[1]).trim().toLowerCase(); // Index 1 to Kolumna B (Imię i nazwisko)
 if (n) nameToIndex[n] = idx;
 });
 
let addedCount = 0;
 let updatedCount = 0;
 
// Rozpoczynamy import danych od wiersza następującego po wykrytym nagłówku
 for (let i = startRowIdx; i < data.length; i++) {
 const row = data[i];
 const name = String(row[nameIdx] || "").trim();
 if (!name) continue; // Pomija wiersz jeśli nie ma w nim imienia i nazwiska
 
const email = String(row[emailIdx] || "").trim();
 const phone = String(row[phoneIdx] || "").trim();
 const nip = String(row[nipIdx] || "").trim();
 const rate = String(row[rateIdx] || "").trim();
 const contractNum = contractNumIdx >= 0 ? String(row[contractNumIdx] || "").trim() : "";
 
const lowerName = name.toLowerCase();
 
if (nameToIndex.hasOwnProperty(lowerName)) {
 // Aktualizacja istniejącego wpisu (w pamięci)
 const idx = nameToIndex[lowerName];
 let changed = false;
 
if (String(currentData[idx][3]).trim() !== email) { currentData[idx][3] = email; changed = true; }
 if (String(currentData[idx][4]).trim() !== phone) { currentData[idx][4] = phone; changed = true; }
 if (String(currentData[idx][5]).trim() !== rate) { currentData[idx][5] = rate; changed = true; } 
if (String(currentData[idx][6]).trim() !== nip) { currentData[idx][6] = nip; changed = true; }   
if (String(currentData[idx][7] || "").trim() !== contractNum) { currentData[idx][7] = contractNum; changed = true; }
 
if (changed) {
 updatedCount++;
 }
 } else {
 // Dodanie nowego technika na koniec tabeli w pamięci
 currentData.push(["", name, "aktywny", email, phone, rate, nip, contractNum]); 
nameToIndex[lowerName] = currentData.length - 1;
 addedCount++;
 }
 }
 
// Masowy zapis zaktualizowanych i nowych danych do arkusza
 if (currentData.length > 0) {
 dataSheet.getRange(9, 1, currentData.length, 8).setValues(currentData);
 }
 
return { success: true, message: `Pobrano dane pomyślnie.\nNowi technicy: ${addedCount}\nZaktualizowano: ${updatedCount}.` };
 } catch (e) {
 return { success: false, message: "Wystąpił nieoczekiwany błąd podczas pobierania danych z pliku: " + e.toString() };
 }
}

// =======================================================================
// NOWE: OBLICZANIE PROPORCJONALNYCH PROTOKOŁÓW (WAGA CZASU x %)
// =======================================================================
function generateProtocolsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSpis = ss.getSheetByName('Spis wykonanych prac');
  const sheetZadan = ss.getSheetByName('Lista zadań');

  // 1. Zbudowanie mapy zadań (Pobranie Powierzchni, Jednostki i aktualnego Procentu z bazy)
  let taskMap = {};
  if (sheetZadan) {
    const lastRow = sheetZadan.getLastRow();
    if (lastRow >= 3) {
      // Czytamy pierwsze 20 kolumn, żeby objąć R (Jednostka) i S (Powierzchnia/Ilość)
      const zadanData = sheetZadan.getRange(1, 1, lastRow, 20).getValues();
      for (let i = 2; i < zadanData.length; i++) {
        let deck = String(zadanData[i][2] || "").trim().toUpperCase();
        let area = String(zadanData[i][3] || "").trim();
        let room = String(zadanData[i][4] || "").trim(); // Pokój - Kolumna E
        let taskName = String(zadanData[i][9] || "").trim(); // Nazwa zadania - Kolumna J
        
        let progressRaw = zadanData[i][16]; // Kolumna Q (%)
        let progress = 0;
        if (typeof progressRaw === 'number') {
            progress = progressRaw <= 1 ? progressRaw * 100 : progressRaw;
        } else {
            progress = parseInt(String(progressRaw).replace('%', '').trim()) || 0;
        }
        
        let unit = String(zadanData[i][5] || "").trim().toLowerCase(); // Kolumna F (Jednostka)
        let explicitAmount = parseFloat(String(zadanData[i][18] || "").replace(',', '.')) || 0;
        let amount = explicitAmount || getTaskSurfaceForTask(deck, area, room, taskName) || 0;
        
        // Unikalny klucz łączący Pokład, Strefę i Zadanie (oraz pokój jeśli istnieje)
        let key = deck + "|" + area + "|" + room + "|" + taskName;
        if (deck && taskName) {
          taskMap[key] = { 
            deck: deck, area: area, room: room, taskName: taskName, 
            progress: progress, unit: unit, amount: amount, 
            totalHours: 0, techs: {} 
          };
        }
      }
    }
  }

  // 2. Sumowanie godzin z poszczególnych zadań (Dla całości i per Technik)
  if (sheetSpis && sheetSpis.getLastRow() >= 2) {
    // Odczyt od kolumny B(1) do I(8)
    const spisData = sheetSpis.getRange(2, 2, sheetSpis.getLastRow() - 1, 8).getDisplayValues();
    for (let i = 0; i < spisData.length; i++) {
      let start = spisData[i][0]; // Kolumna B
      let end = spisData[i][1];   // Kolumna C
      let deck = String(spisData[i][2]).trim().toUpperCase(); // Kolumna D
      let area = String(spisData[i][3]).trim(); // Kolumna E
      let taskName = String(spisData[i][4]).trim(); // Kolumna F
      let tech = String(spisData[i][6]).trim(); // Kolumna H (Technik)
      let room = String(spisData[i][7] || "").trim(); // Kolumna I (Pokój)
      
      if (start && end && taskName && !taskName.includes("[ANULOWANE]")) {
        let dur = parseWorkTimeBackend(end) - parseWorkTimeBackend(start);
        if (dur < 0) dur += 24;
        dur = Math.round(dur * 100) / 100;
        
        if (dur > 0) {
          let key = deck + "|" + area + "|" + room + "|" + taskName;
          // Zabezpieczenie: jeśli nie wbiło pokoju, szuka zadania ogólnego
          if (!taskMap[key]) {
             key = deck + "|" + area + "||" + taskName; 
          }
          
          if (taskMap[key]) {
            taskMap[key].totalHours += dur;
            if (!taskMap[key].techs[tech]) taskMap[key].techs[tech] = 0;
            taskMap[key].techs[tech] += dur; // Dopisanie godzin konkretnemu technikowi
          }
        }
      }
    }
  }

  // 3. Obliczenie Proporcji (Udziału Metrażowego) dla każdego technika
  let protocols = [];
  Object.keys(taskMap).forEach(key => {
    let t = taskMap[key];
    // Liczymy tylko zadania, które mają zdefiniowany metraż (amount > 0) i jakikolwiek postęp > 0%
    if (t.totalHours > 0 && t.amount > 0 && t.progress > 0) {
      
      // Ilość zrobiona np. ze 100 m2 (70%) => 70 m2
      let completedAmount = t.amount * (t.progress / 100);
      
      // Dzielimy zrobiony metraż przez udziały czasowe
      Object.keys(t.techs).forEach(techName => {
        let techHours = t.techs[techName];
        let proportion = techHours / t.totalHours;
        let techShareAmount = completedAmount * proportion;
        techShareAmount = unit.toLowerCase() === 'pis'
          ? Math.round(techShareAmount)
          : Math.round(techShareAmount * 100) / 100;
        
        protocols.push({
          techName: techName,
          deck: t.deck,
          area: t.area,
          room: t.room,
          taskName: t.taskName,
          totalTaskHours: t.totalHours,
          techHours: techHours,
          progress: t.progress,
          totalAmount: t.amount,
          unit: t.unit || 'm2', // Domyślna jednostka jeśli brak
          techShareAmount: techShareAmount
        });
      });
    }
  });

  // Grupujemy dane na potrzeby interfejsu (Sortowanie alfabetyczne po nazwisku)
  let groupedProtocols = {};
  protocols.forEach(p => {
    if (!groupedProtocols[p.techName]) groupedProtocols[p.techName] = [];
    groupedProtocols[p.techName].push(p);
  });

  return groupedProtocols;
}

function resolveRoomTaskAmount_(deck, area, room, taskName, taskType, unit) {
  const normalizedUnit = String(unit || '').trim().toLowerCase();
  const roomRecord = getRoomCatalog().find(record =>
    normalizeDeckName(record.deck) === normalizeDeckName(deck) &&
    [record.room, record.name].some(value => normalizeRoomMatch(value) === normalizeRoomMatch(room))
  );
  if (!roomRecord) return 0;

  if (normalizedUnit === 'm2') return resolveTaskSurfaceForRow(deck, area, room, taskName, taskType);
  return 0;
}

function syncRoomTasks_(deck, roomName, roomNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Lista zadań');
  if (!sheet || sheet.getLastRow() < 3) return 0;

  const rows = sheet.getRange(3, 1, sheet.getLastRow() - 2, 18).getDisplayValues();
  const roomKeys = [roomName, roomNumber].map(normalizeRoomMatch).filter(Boolean);
  let updated = 0;
  rows.forEach((row, index) => {
    if (normalizeDeckName(row[2]) !== normalizeDeckName(deck)) return;
    if (!roomKeys.includes(normalizeRoomMatch(row[4]))) return;
    const unit = String(row[5] || '').trim().toLowerCase();
    if (unit !== 'm2') return;
    const amount = resolveRoomTaskAmount_(deck, row[3], row[4], row[9], row[7], unit);
    if (amount <= 0) return;
    sheet.getRange(index + 3, 18).setValue(Math.round(amount * 100) / 100);
    updated++;
  });
  return updated;
}

function generateTaskIdForArea_(sheet, area, isExtraJob) {
  const normalized = String(area || '').toLowerCase();
  const prefix = normalized.includes('podłog') || normalized.includes('podlog') ? 'P'
    : normalized.includes('ścian') || normalized.includes('scian') ? 'W'
      : normalized.includes('sufit') ? 'C'
        : normalized.includes('instal') ? 'A' : 'A';
  const lastRow = sheet.getLastRow();
  const existingIds = lastRow >= 3 ? sheet.getRange(3, 2, lastRow - 2, 6).getDisplayValues().flatMap(row => [row[0], row[5]]) : [];
  const numbers = existingIds
    .map(value => String(value || '').match(new RegExp('^' + prefix + '(\\d{4})(?:-EJ)?$', 'i')))
    .filter(Boolean)
    .map(match => Number(match[1]));
  const nextId = prefix + String((numbers.length ? Math.max(...numbers) : 0) + 1).padStart(4, '0');
  return isExtraJob ? nextId + '-EJ' : nextId;
}