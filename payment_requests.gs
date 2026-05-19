const PAY_CFG = {
  QUEUE_SHEET: 'Payment_Link_Map',
  SUMMARY_SHEET: 'Сводная',
  HEADER_ROW: 2,
  DATA_START_ROW: 3,
  REGISTRY_ID_DEFAULT: '1BUFF4-tlLg-H8faxgKUcnlcaL8aJS84x9bZGza9_K_M',
  REGISTRY_SHEET_DEFAULT: 'Реестр',
  ROOT_FOLDER_DEFAULT: '1wMYUKNsNixmNcI1HdsmScn4V-G5Vy6wL'
};

function addPaymentRegistryMenu_(ui) {
  ui.createMenu('💳 Оплаты поставщику')
    .addItem('🔐 Авторизация для менеджера', 'payAuthorizeUser')
    .addSeparator()
    .addItem('1) Подать заявку на оплату', 'payOpenManagerRequestDialog')
    .addItem('2) Проверка/отправка в реестр', 'payOpenApprovalDialog')
    .addSeparator()
    .addItem('3) Синхронизировать оплаченные', 'paySyncPaidStatuses')
    .addItem('4) Обновить шапки листов менеджера (3 типа заявок)', 'payMigrateManagerHeadersForPaymentTypes')
    .addToUi();
}

function payAuthorizeUser() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    // Вызываем и UI, и Drive, чтобы запросились все нужные разрешения сразу.
    const cell = sheet.getRange(1, 1);
    const oldValue = cell.getValue();
    cell.setValue(oldValue);
    const rootId = payGetProp_('PAYMENT_REGISTRY_ROOT_FOLDER_ID', PAY_CFG.ROOT_FOLDER_DEFAULT);
    DriveApp.getFolderById(rootId).getName();
    DriveApp.getRootFolder().getName();
    // Явно запрашиваем право создания папок/файлов на Drive.
    const testFolder = DriveApp.getRootFolder().createFolder('AUTH_TEST_' + new Date().getTime());
    testFolder.setTrashed(true);
    SpreadsheetApp.getUi().alert(
      '✅ Авторизация выполнена',
      'Разрешения (таблица + Drive: чтение/запись/создание папок) выданы. Теперь можно подавать заявки через меню "💳 Оплаты поставщику".',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      '⚠️ Требуется подтверждение доступа',
      'Нажмите "Продолжить" и подтвердите доступ Google, затем повторите действие.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

function payGetProp_(key, fallback) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  return v == null || String(v).trim() === '' ? fallback : String(v).trim();
}

function payNorm_(v) {
  return String(v || '').trim();
}

function payQueueId_() {
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss-SSS');
  const rnd = Math.floor(Math.random() * 1000);
  return 'PRQ-' + ts + '-' + String(rnd).padStart(3, '0');
}

function payQueueSignatureFromRow_(r) {
  // Устойчивая сигнатура дубликата: без суммы/срока, чтобы не зависеть от форматов.
  // manager_sheet|manager_name|spec|payment_type|percent|counterparty
  return [
    payCanon_(r[3]),
    payCanon_(r[4]),
    payCanon_(r[5]),
    payCanon_(r[6]),
    String(r[7] == null ? '' : r[7]),
    payCanon_(r[8])
  ].join('|');
}

function payCanon_(v) {
  return payNorm_(v).toLowerCase().replace(/\s+/g, '');
}

/**
 * Канонизация именно ЗАГОЛОВКОВ листа (а не произвольных значений в строках).
 * Согласована с `syncManagerCanonHeader_` (main.gs): «№», «#», точки, скобки,
 * слэши, кавычки, дефисы и подчёркивания → пробел; NBSP → space; ё → е.
 * Без этого «№ спецификации», «Номер_спецификации», «Спецификация/инвойс»
 * не матчатся с «Номер спецификации», и `payFindCol_` молча падает в фолбэк
 * (что давало баг «вместо номера спецификации тянется сумма»).
 */
function payHeaderCanon_(s) {
  return String(s == null ? '' : s)
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/\u00A0/g, ' ')
    .replace(/[№#]/g, ' ')
    .replace(/[.,:;()/\\\[\]{}'"“”«»]/g, ' ')
    .replace(/[-_]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function payHeaderMap_(headers) {
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    const k = payHeaderCanon_(headers[i]);
    if (k && map[k] == null) map[k] = i;
  }
  return map;
}

/**
 * Поиск колонки по списку алиасов. Возвращает 0-based индекс или -1.
 * В отличие от старой версии НЕ имеет фолбэка на цифровой номер — если
 * колонку не нашли, caller обязан сообщить пользователю осмысленную ошибку.
 */
function payFindHeaderCol_(map, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const hit = map[payHeaderCanon_(aliases[i])];
    if (hit != null) return hit;
  }
  return -1;
}

const PAY_SPEC_ALIASES = [
  'Номер спецификации', 'Номер Спецификации', '№ спецификации',
  'Номер_спецификации', 'Номер спец', '№ спец',
  'Спецификация', 'Спецификации', 'Спец', 'Спец №',
  'Спецификация/инвойс', 'Спецификация инвойса', 'Спецификация инвойс',
  'Номер спецификации/инвойса', 'Номер спецификации инвойса',
  'Spec', 'Spec number', 'Spec No', 'Specification'
];

const PAY_AMOUNT_ALIASES = [
  'Сумма', 'Сумма итого', 'Итого сумма', 'Итого', 'Итого, сумма',
  'Сумма строки', 'Сумма позиции',
  'Amount', 'Total', 'Sum'
];

/**
 * Бросает читаемую ошибку с перечнем фактических заголовков листа.
 * Зовём её, когда ни один из алиасов не нашёлся.
 */
function payThrowHeaderNotFound_(kindLabel, aliases, headers) {
  const visible = headers
    .map(function (h) { return String(h == null ? '' : h).trim(); })
    .filter(function (s) { return s; })
    .slice(0, 30);
  throw new Error(
    'Не нашёл колонку «' + kindLabel + '» на активном листе.\n' +
    'Искал по вариантам: ' + aliases.slice(0, 6).join(', ') + ' …\n' +
    'Фактические заголовки строки ' + PAY_CFG.HEADER_ROW + ': ' + visible.join(' | ') + '\n' +
    'Переименуй колонку обратно в «' + aliases[0] + '» или сообщи, какое название использовать.'
  );
}

function payManagerNameFromSheet_(sheetName) {
  const m = String(sheetName || '').match(/^(.+?)\s+\d{2}\/\d{2}$/);
  return m ? m[1].trim() : payNorm_(sheetName);
}

function payGetCounterparties() {
  const raw = PropertiesService.getDocumentProperties().getProperty('PAY_COUNTERPARTIES');
  if (!raw) return [];
  try {
    const arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    return [];
  }
}

function paySaveCounterparty_(name) {
  const v = payNorm_(name);
  if (!v) return;
  const list = payGetCounterparties();
  const key = payCanon_(v);
  const exists = list.some(function (x) { return payCanon_(x) === key; });
  if (!exists) {
    list.push(v);
    PropertiesService.getDocumentProperties().setProperty('PAY_COUNTERPARTIES', JSON.stringify(list));
  }
}

function payCollectSpecsFromActiveSheet_() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastCol = Math.max(sheet.getLastColumn(), 30);
  const lastRow = sheet.getLastRow();
  const headers = sheet.getRange(PAY_CFG.HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const map = payHeaderMap_(headers);
  const specCol = payFindHeaderCol_(map, PAY_SPEC_ALIASES);
  const amountCol = payFindHeaderCol_(map, PAY_AMOUNT_ALIASES);
  if (specCol < 0) payThrowHeaderNotFound_('Номер спецификации', PAY_SPEC_ALIASES, headers);
  if (amountCol < 0) payThrowHeaderNotFound_('Сумма', PAY_AMOUNT_ALIASES, headers);
  if (specCol === amountCol) {
    throw new Error(
      'Колонки «Номер спецификации» и «Сумма» сошлись в одну — это исключено.\n' +
      'Найденный индекс: ' + (specCol + 1) + '. Проверь заголовки строки ' + PAY_CFG.HEADER_ROW + ' активного листа.'
    );
  }
  const rows = lastRow >= PAY_CFG.DATA_START_ROW
    ? sheet.getRange(PAY_CFG.DATA_START_ROW, 1, lastRow - PAY_CFG.DATA_START_ROW + 1, lastCol).getValues()
    : [];
  const out = {};
  rows.forEach(function (r) {
    const spec = payNorm_(r[specCol]);
    if (!spec) return;
    const n = parseFloat(String(r[amountCol]).replace(/\s/g, '').replace(',', '.'));
    out[spec] = (out[spec] || 0) + (isFinite(n) ? n : 0);
  });
  return Object.keys(out).sort().map(function (spec) {
    return { spec: spec, amount: out[spec] };
  });
}

function payOpenManagerRequestDialog() {
  const specs = payCollectSpecsFromActiveSheet_();
  if (!specs.length) {
    SpreadsheetApp.getUi().alert('Не найдено спецификаций на активном листе менеджера.');
    return;
  }
  const counterparties = payGetCounterparties();
  const html = HtmlService.createHtmlOutput(payManagerDialogHtml_(specs, counterparties)).setWidth(640).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, 'Заявка на оплату');
}

function payManagerDialogHtml_(specs, counterparties) {
  return (
    '<!doctype html><html><head><base target="_top"><meta charset="utf-8">' +
    '<style>body{font-family:Arial;padding:14px}.row{margin:10px 0}label{display:block;font-weight:600;margin-bottom:4px}input,select{width:100%;padding:8px;box-sizing:border-box}</style>' +
    '</head><body>' +
    '<div class="row"><label>Номер спецификации</label><select id="spec">' +
    specs.map(function (x) { return '<option value="' + x.spec + '">' + x.spec + ' (сумма: ' + x.amount.toFixed(2) + ')</option>'; }).join('') +
    '</select></div>' +
    '<div class="row"><label>Тип оплаты</label><select id="ptype"><option>Аванс</option><option>Баланс</option><option>Отсрочка</option></select></div>' +
    '<div class="row"><label>Размер оплаты, %</label><input id="percent" type="number" min="0" max="100" value="20"></div>' +
    '<div class="row"><label>Валюта заявки</label><select id="currency"><option value="CNY" selected>CNY</option><option value="USD">USD</option><option value="EUR">EUR</option><option value="RUB">RUB</option></select></div>' +
    '<div class="row"><label>Оплатить до</label><input id="due" type="date"></div>' +
    '<div class="row"><label>Контрагент (обязательно)</label><input id="counterparty" list="cpList" placeholder="Введите или выберите"></div>' +
    '<datalist id="cpList">' +
    counterparties.map(function (c) { return '<option value="' + c + '"></option>'; }).join('') +
    '</datalist>' +
    '<div class="row"><label>Подписанный документ (обязательно)</label><input id="files" type="file" multiple></div>' +
    '<div class="row"><button id="sendBtn" onclick="send()">Отправить на проверку</button></div>' +
    '<p style="font-size:12px;color:#555">Крупные файлы отправляются частями (без одного большого запроса).</p>' +
    '<script>var sending=false;var CHUNK=450000;' +
    'function toB64(f){return new Promise(function(res,rej){var r=new FileReader();r.onload=function(){res(String(r.result).split(\",\")[1]);};r.onerror=rej;r.readAsDataURL(f);});}' +
    'function runPrepare(payload){return new Promise(function(ok,er){google.script.run.withSuccessHandler(ok).withFailureHandler(er).payManagerDialogPrepare(payload);});}' +
    'function runChunk(token,folderId,fi,ci,total,name,mime,part){return new Promise(function(ok,er){google.script.run.withSuccessHandler(ok).withFailureHandler(er).payManagerDialogUploadChunk(token,folderId,fi,ci,total,name,mime,part);});}' +
    'function runFinalize(token,folderId){return new Promise(function(ok,er){google.script.run.withSuccessHandler(ok).withFailureHandler(er).payManagerDialogFinalize(token,folderId);});}' +
    'async function send(){var spec=document.getElementById(\"spec\").value;var ptype=document.getElementById(\"ptype\").value;var pct=document.getElementById(\"percent\").value;var currency=document.getElementById(\"currency\").value;var due=document.getElementById(\"due\").value;var cp=document.getElementById(\"counterparty\").value;' +
    'var fs=[].slice.call(document.getElementById(\"files\").files);if(!cp){alert(\"Контрагент обязателен\");return;}if(!fs.length){alert(\"Нужен подписанный документ\");return;}if(sending)return;sending=true;document.getElementById(\"sendBtn\").disabled=true;' +
    'try{var payload={spec:spec,paymentType:ptype,percent:pct,currency:currency,dueDate:due,counterparty:cp};var prep=await runPrepare(payload);' +
    'for(var fi=0;fi<fs.length;fi++){var b64=await toB64(fs[fi]);var total=Math.max(1,Math.ceil(b64.length/CHUNK));' +
    'for(var ci=0;ci<total;ci++){var part=b64.substring(ci*CHUNK,(ci+1)*CHUNK);await runChunk(prep.uploadToken,prep.folderId,fi,ci,total,fs[fi].name,fs[fi].type||\'application/octet-stream\',part);}}' +
    'var msg=await runFinalize(prep.uploadToken,prep.folderId);alert(msg);google.script.host.close();}catch(e){alert(e.message||e);}finally{sending=false;document.getElementById(\"sendBtn\").disabled=false;}}</script>' +
    '</body></html>'
  );
}

/**
 * Точки входа для google.script.run из формы менеджера.
 * Функции с суффиксом `_*` в ряде окружений не отдаются клиенту — вызов даёт "is not a function".
 */
function payManagerDialogPrepare(payload) {
  return payPrepareManagerRequest_(payload);
}

function payManagerDialogUploadChunk(uploadToken, folderId, fileIndex, chunkIndex, totalChunks, fileName, mimeType, base64Chunk) {
  payUploadManagerFileChunk_(
    uploadToken,
    folderId,
    fileIndex,
    chunkIndex,
    totalChunks,
    fileName,
    mimeType,
    base64Chunk
  );
}

function payManagerDialogFinalize(uploadToken, folderId) {
  return payFinalizeManagerRequest_(uploadToken, folderId);
}

function payEnsureQueueSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(PAY_CFG.QUEUE_SHEET);
  if (!sh) sh = ss.insertSheet(PAY_CFG.QUEUE_SHEET);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, 19).setValues([[
      'queue_id', 'created_at', 'created_by', 'manager_sheet', 'manager_name', 'spec', 'payment_type',
      'percent', 'counterparty', 'amount', 'currency', 'due_date', 'manager_rows_json',
      'folder_url', 'file_links_json', 'status', 'registry_request_no', 'registry_row', 'reject_reason'
    ]]);
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < 19) {
    sh.getRange(1, 19).setValue('reject_reason');
  }
  return sh;
}

/**
 * Подготовка черновика заявки: папка на Drive + meta JSON. Файлы грузятся отдельными вызовами (чанки), чтобы не упираться в HTTP 413.
 */
function payPrepareManagerRequest_(payload) {
  const bundle = payManagerRequestBuildDraft_(payload);
  const metaBlob = Utilities.newBlob(
    JSON.stringify(bundle.meta),
    'application/json',
    '__pay_upload_meta.json'
  );
  bundle.folder.createFile(metaBlob);
  return { uploadToken: bundle.queueId, folderId: bundle.folder.getId(), queueId: bundle.queueId };
}

function payManagerRequestBuildDraft_(payload) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const managerSheetName = sheet.getName();
  const managerName = payManagerNameFromSheet_(managerSheetName);
  const spec = payNorm_(payload.spec);
  const pType = payNorm_(payload.paymentType);
  const pct = parseFloat(payload.percent);
  const currency = payNorm_(payload.currency || 'CNY').toUpperCase();
  const counterparty = payNorm_(payload.counterparty);
  if (!spec || !pType || !isFinite(pct) || pct <= 0 || pct > 100 || !counterparty) {
    throw new Error('Проверьте обязательные поля заявки.');
  }
  if (!/^[A-Z]{3}$/.test(currency)) {
    throw new Error('Некорректная валюта заявки.');
  }

  const lastCol = Math.max(sheet.getLastColumn(), 30);
  const lastRow = sheet.getLastRow();
  const headers = sheet.getRange(PAY_CFG.HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const map = payHeaderMap_(headers);
  const specCol = payFindHeaderCol_(map, PAY_SPEC_ALIASES);
  const amountCol = payFindHeaderCol_(map, PAY_AMOUNT_ALIASES);
  if (specCol < 0) payThrowHeaderNotFound_('Номер спецификации', PAY_SPEC_ALIASES, headers);
  if (amountCol < 0) payThrowHeaderNotFound_('Сумма', PAY_AMOUNT_ALIASES, headers);
  if (specCol === amountCol) {
    throw new Error(
      'Колонки «Номер спецификации» и «Сумма» сошлись в одну — это исключено.\n' +
      'Найденный индекс: ' + (specCol + 1) + '. Проверь заголовки строки ' + PAY_CFG.HEADER_ROW + ' активного листа.'
    );
  }

  const rows = sheet.getRange(PAY_CFG.DATA_START_ROW, 1, Math.max(0, lastRow - PAY_CFG.DATA_START_ROW + 1), lastCol).getValues();
  const matchedRows = [];
  let baseAmount = 0;
  for (let i = 0; i < rows.length; i++) {
    if (payNorm_(rows[i][specCol]) !== spec) continue;
    const absRow = PAY_CFG.DATA_START_ROW + i;
    matchedRows.push(absRow);
    const n = parseFloat(String(rows[i][amountCol]).replace(/\s/g, '').replace(',', '.'));
    if (isFinite(n)) baseAmount += n;
  }
  if (!matchedRows.length) throw new Error('По выбранной спецификации не найдены строки менеджера.');

  const amount = +(baseAmount * pct / 100).toFixed(2);
  const queueId = payQueueId_();
  const root = payGetDraftRootFolder_();
  const folder = root.createFolder(queueId + ' - ' + counterparty);
  const createdBy = Session.getActiveUser().getEmail() || '';
  const meta = {
    v: 1,
    queueId: queueId,
    createdBy: createdBy,
    managerSheetName: managerSheetName,
    managerName: managerName,
    spec: spec,
    paymentType: pType,
    percent: pct,
    counterparty: counterparty,
    amount: amount,
    currency: currency,
    dueDate: payload.dueDate ? String(payload.dueDate) : '',
    matchedRows: matchedRows
  };
  return { queueId: queueId, folder: folder, meta: meta };
}

function payUploadManagerFileChunk_(uploadToken, folderId, fileIndex, chunkIndex, totalChunks, fileName, mimeType, base64Chunk) {
  const folder = DriveApp.getFolderById(folderId);
  payAssertUploadMeta_(folder, uploadToken);
  const fi = Number(fileIndex);
  const ci = Number(chunkIndex);
  const tc = Number(totalChunks);
  if (!isFinite(fi) || fi < 0 || !isFinite(ci) || ci < 0 || !isFinite(tc) || tc < 1) {
    throw new Error('Некорректные параметры загрузки файла.');
  }
  if (ci === 0) {
    payTrashFileByName_(folder, '__fileinfo_f' + fi + '.json');
    const infoBlob = Utilities.newBlob(
      JSON.stringify({ name: fileName || 'file', mimeType: mimeType || 'application/octet-stream', totalChunks: tc }),
      'application/json',
      '__fileinfo_f' + fi + '.json'
    );
    folder.createFile(infoBlob);
  }
  const partName = '__part_f' + fi + '_c' + ci + '.b64';
  payTrashFileByName_(folder, partName);
  folder.createFile(Utilities.newBlob(String(base64Chunk || ''), 'text/plain', partName));
}

function payTrashFileByName_(folder, name) {
  const it = folder.getFilesByName(name);
  while (it.hasNext()) {
    it.next().setTrashed(true);
  }
}

function payFinalizeManagerRequest_(uploadToken, folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const meta = payReadUploadMeta_(folder);
  if (!meta || payNorm_(meta.queueId) !== payNorm_(uploadToken)) {
    throw new Error('Сессия загрузки недействительна или устарела.');
  }

  const files = folder.getFiles();
  const fileInfoByIdx = {};
  const partsByFile = {};
  while (files.hasNext()) {
    const f = files.next();
    const n = f.getName();
    const mInfo = /^__fileinfo_f(\d+)\.json$/i.exec(n);
    if (mInfo) {
      const idx = parseInt(mInfo[1], 10);
      const txt = f.getBlob().getDataAsString();
      try {
        fileInfoByIdx[idx] = JSON.parse(txt);
      } catch (e) {
        throw new Error('Повреждены служебные данные загрузки.');
      }
      continue;
    }
    const mPart = /^__part_f(\d+)_c(\d+)\.b64$/i.exec(n);
    if (mPart) {
      const fi = parseInt(mPart[1], 10);
      const ci = parseInt(mPart[2], 10);
      if (!partsByFile[fi]) partsByFile[fi] = {};
      partsByFile[fi][ci] = f.getBlob().getDataAsString();
    }
  }

  const idxs = Object.keys(fileInfoByIdx)
    .map(function (x) {
      return parseInt(x, 10);
    })
    .filter(function (n) {
      return isFinite(n);
    })
    .sort(function (a, b) {
      return a - b;
    });
  if (!idxs.length) throw new Error('Нет загруженных файлов (чанки не найдены).');

  const fileLinks = [];
  for (let k = 0; k < idxs.length; k++) {
    const fi = idxs[k];
    const info = fileInfoByIdx[fi];
    const tc = info && info.totalChunks ? parseInt(info.totalChunks, 10) : 0;
    const name = info && info.name ? info.name : 'file';
    const mime = info && info.mimeType ? info.mimeType : 'application/octet-stream';
    if (!tc || tc < 1) throw new Error('Некорректная мета файла: ' + name);

    const chunks = partsByFile[fi];
    if (!chunks) throw new Error('Нет данных для файла: ' + name);

    let joined = '';
    for (let c = 0; c < tc; c++) {
      const piece = chunks[c];
      if (piece == null) throw new Error('Неполная загрузка файла (нет части ' + c + '): ' + name);
      joined += piece;
    }

    const blob = Utilities.newBlob(Utilities.base64Decode(joined), mime, name);
    const gf = folder.createFile(blob);
    fileLinks.push(gf.getUrl());
  }

  payTrashUploadTempFiles_(folder);

  const sh = payEnsureQueueSheet_();
  const newRow = sh.getLastRow() + 1;
  sh.getRange(newRow, 1, 1, 19).setValues([[
    meta.queueId,
    new Date(),
    meta.createdBy,
    meta.managerSheetName,
    meta.managerName,
    meta.spec,
    meta.paymentType,
    meta.percent,
    meta.counterparty,
    meta.amount,
    meta.currency,
    meta.dueDate ? new Date(meta.dueDate) : '',
    JSON.stringify(meta.matchedRows),
    folder.getUrl(),
    JSON.stringify(fileLinks),
    'На проверке',
    '',
    '',
    ''
  ]]);

  paySaveCounterparty_(meta.counterparty);
  return 'Заявка создана: ' + meta.queueId;
}

function payAssertUploadMeta_(folder, uploadToken) {
  const meta = payReadUploadMeta_(folder);
  if (!meta || payNorm_(meta.queueId) !== payNorm_(uploadToken)) {
    throw new Error('Неверный токен загрузки или папка.');
  }
}

function payReadUploadMeta_(folder) {
  const it = folder.getFilesByName('__pay_upload_meta.json');
  if (!it.hasNext()) return null;
  const txt = it.next().getBlob().getDataAsString();
  try {
    return JSON.parse(txt);
  } catch (e) {
    return null;
  }
}

function payTrashUploadTempFiles_(folder) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    const n = f.getName();
    if (
      n === '__pay_upload_meta.json' ||
      /^__fileinfo_f\d+\.json$/i.test(n) ||
      /^__part_f\d+_c\d+\.b64$/i.test(n)
    ) {
      f.setTrashed(true);
    }
  }
}

/**
 * Совместимость: одна заявка одним вызовом (малые файлы). Крупные — через prepare / chunks / finalize.
 */
function paySubmitManagerRequest(payload, files) {
  if (!files || !files.length) throw new Error('Нужно приложить подписанный документ.');
  const prep = payPrepareManagerRequest_(payload);
  const CHUNK = 450000;
  for (let i = 0; i < files.length; i++) {
    const f = files[i];
    const data = f.data || '';
    const total = Math.max(1, Math.ceil(data.length / CHUNK));
    for (let c = 0; c < total; c++) {
      payUploadManagerFileChunk_(
        prep.uploadToken,
        prep.folderId,
        i,
        c,
        total,
        f.name || 'file',
        f.type || 'application/octet-stream',
        data.substring(c * CHUNK, (c + 1) * CHUNK)
      );
    }
  }
  return payFinalizeManagerRequest_(prep.uploadToken, prep.folderId);
}

function payOpenApprovalDialog() {
  payAssertApprover_();
  const list = payGetPendingQueue_();
  if (!list.length) {
    SpreadsheetApp.getUi().alert('Нет заявок на проверке.');
    return;
  }
  const html = HtmlService.createHtmlOutput(payApprovalDialogHtml_(list)).setWidth(760).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Проверка заявок');
}

function payAssertApprover_() {
  const allowed = payGetProp_('PAYMENT_APPROVER_EMAIL', 'banych83@gmail.com');
  const me = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!me || me !== allowed.toLowerCase()) {
    throw new Error('Только согласующий (' + allowed + ') может выполнять это действие.');
  }
}

function payGetPendingQueue_() {
  const sh = payEnsureQueueSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const data = sh.getRange(2, 1, last - 1, 19).getValues();
  const out = [];
  const seenById = {};
  const seenBySig = {};
  data.forEach(function (r, i) {
    if (payNorm_(r[15]) !== 'На проверке') return;
    const qid = payNorm_(r[0]);
    const sig = payQueueSignatureFromRow_(r);
    if (qid && seenById[qid]) return;
    if (sig && seenBySig[sig]) return;
    if (qid) seenById[qid] = true;
    if (sig) seenBySig[sig] = true;
    out.push({
      row: i + 2,
      queueId: qid || ('ROW-' + (i + 2)),
      manager: r[4],
      spec: r[5],
      paymentType: r[6],
      percent: r[7],
      counterparty: r[8],
      amount: r[9],
      dueDate: r[11] ? Utilities.formatDate(new Date(r[11]), Session.getScriptTimeZone(), 'dd.MM.yyyy') : '',
      folderUrl: r[13],
      fileLinks: (function () {
        try { return JSON.parse(r[14] || '[]'); } catch (e) { return []; }
      })()
    });
  });
  return out;
}

function payApprovalDialogHtml_(list) {
  return (
    '<!doctype html><html><head><base target="_top"><meta charset="utf-8"><style>body{font-family:Arial;padding:12px}table{width:100%;border-collapse:collapse}td,th{border:1px solid #ddd;padding:6px;font-size:12px}button{padding:4px 8px}</style></head><body>' +
    '<table><tr><th>ID</th><th>Менеджер</th><th>Спец</th><th>Тип</th><th>%</th><th>Сумма</th><th>Контрагент</th><th>Срок</th><th>Документы</th><th>Действие</th></tr>' +
    list.map(function (x) {
      var docs = [];
      if (x.folderUrl) docs.push('<a href="' + x.folderUrl + '" target="_blank">Папка</a>');
      if (x.fileLinks && x.fileLinks.length) {
        for (var i = 0; i < x.fileLinks.length; i++) {
          docs.push('<a href="' + x.fileLinks[i] + '" target="_blank">Файл ' + (i + 1) + '</a>');
        }
      }
      return '<tr><td>' + x.queueId + '</td><td>' + x.manager + '</td><td>' + x.spec + '</td><td>' + x.paymentType + '</td><td>' + x.percent + '</td><td>' + x.amount + '</td><td>' + x.counterparty + '</td><td>' + x.dueDate + '</td><td>' + (docs.join('<br>') || '-') + '</td><td><button onclick="ap(' + x.row + ')">Одобрить</button> <button onclick="rej(' + x.row + ')">Отклонить</button></td></tr>';
    }).join('') +
    '</table><script>var busy=false;function lockUi(){if(busy)return false;busy=true;var bs=document.querySelectorAll(\"button\");for(var i=0;i<bs.length;i++)bs[i].disabled=true;return true;}function unlockUi(){busy=false;var bs=document.querySelectorAll(\"button\");for(var i=0;i<bs.length;i++)bs[i].disabled=false;}function ap(row){if(!lockUi())return;google.script.run.withSuccessHandler(function(m){alert(m);google.script.host.close();}).withFailureHandler(function(e){unlockUi();alert(e.message||e);}).payApproveQueueRow(row);}function rej(row){if(!lockUi())return;var reason=prompt(\"Причина отклонения:\",\"\");if(reason===null){unlockUi();return;}if(!String(reason).trim()){unlockUi();alert(\"Укажите причину\");return;}google.script.run.withSuccessHandler(function(m){alert(m);google.script.host.close();}).withFailureHandler(function(e){unlockUi();alert(e.message||e);}).payRejectQueueRow(row, reason);}</script></body></html>'
  );
}

function payGetDraftRootFolder_() {
  const rootId = payGetProp_('PAYMENT_REGISTRY_ROOT_FOLDER_ID', PAY_CFG.ROOT_FOLDER_DEFAULT);
  const root = DriveApp.getFolderById(rootId);
  const draftName = payGetProp_('PAYMENT_DRAFT_FOLDER_NAME', '_DRAFT_PAYMENT_REQUESTS');
  const it = root.getFoldersByName(draftName);
  if (it.hasNext()) return it.next();
  return root.createFolder(draftName);
}

function payApproveQueueRow(queueRow) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
  payAssertApprover_();
  const sh = payEnsureQueueSheet_();
  const r = sh.getRange(queueRow, 1, 1, 19).getValues()[0];
  const status = payNorm_(r[15]);
  if (status !== 'На проверке') {
    const existingReq = payNorm_(r[16]);
    if (existingReq) return 'Заявка уже обработана ранее: ' + existingReq;
    throw new Error('Строка уже обработана.');
  }

  const queueId = r[0];
  const createdBy = r[2];
  const managerSheet = r[3];
  const managerName = r[4];
  const spec = r[5];
  const paymentType = r[6];
  const percent = r[7];
  const counterparty = r[8];
  const amount = r[9];
  const currency = payNorm_(r[10]) || 'CNY';
  const dueDate = r[11];
  const managerRows = JSON.parse(r[12] || '[]');
  const draftFolderUrl = payNorm_(r[13]);
  const fileLinks = JSON.parse(r[14] || '[]');

  const registryId = payGetProp_('PAYMENT_REGISTRY_SPREADSHEET_ID', PAY_CFG.REGISTRY_ID_DEFAULT);
  const registrySheetName = payGetProp_('PAYMENT_REGISTRY_SHEET_NAME', PAY_CFG.REGISTRY_SHEET_DEFAULT);
  const regSs = SpreadsheetApp.openById(registryId);
  const regSh = regSs.getSheetByName(registrySheetName);
  if (!regSh) throw new Error('Лист реестра не найден: ' + registrySheetName);

  const last = regSh.getLastRow();
  const reqNo = 'ЗАЯВКА-' + new Date().getFullYear() + '-' + String(last).padStart(4, '0');
  const purpose = paymentType + ' по инвойсу (спецификации) № ' + spec + ' в размере ' + percent + '%.';
  const approver = Session.getActiveUser().getEmail() || '';
  const finalFolderUrl = payCreateApprovedFolder_(reqNo, counterparty, fileLinks, draftFolderUrl);
  const newRow = last + 1;
  regSh.insertRowAfter(last);
  regSh.getRange(newRow, 1).setValue(new Date());
  regSh.getRange(newRow, 2).setValue(dueDate || '');
  regSh.getRange(newRow, 3).setValue('Печёнкин А.А.');
  regSh.getRange(newRow, 4).setValue('Снабжение');
  regSh.getRange(newRow, 5).setValue(counterparty);
  regSh.getRange(newRow, 6).setValue(purpose);
  regSh.getRange(newRow, 7).setValue(amount);
  regSh.getRange(newRow, 8).setValue(currency);
  regSh.getRange(newRow, 9).setValue(fileLinks.join('\n'));
  regSh.getRange(newRow, 10).setValue('На согласовании');
  regSh.getRange(newRow, 11).setValue('');
  regSh.getRange(newRow, 12).setValue('');
  regSh.getRange(newRow, 13).setValue('Передано из менеджерской заявки ' + queueId);
  regSh.getRange(newRow, 14).setValue('Не оплачено');
  if (currency.toUpperCase() === 'RUB') {
    regSh.getRange(newRow, 16).setValue(amount);
  } else {
    regSh.getRange(newRow, 16).setFormula('=G' + newRow + '*GOOGLEFINANCE("CURRENCY:' + currency.toUpperCase() + 'RUB")');
  }
  regSh.getRange(newRow, 17).setValue(approver);
  regSh.getRange(newRow, 18).setValue(reqNo);
  regSh.getRange(newRow, 19).setValue(finalFolderUrl);

  sh.getRange(queueRow, 16).setValue('Отправлено в реестр');
  sh.getRange(queueRow, 17).setValue(reqNo);
  sh.getRange(queueRow, 18).setValue(newRow);
  sh.getRange(queueRow, 19).setValue('');

  sh.getRange(queueRow, 14).setValue(finalFolderUrl);
  payWriteRequestToManagerAndSummary_(managerSheet, managerName, spec, managerRows, reqNo, finalFolderUrl, paymentType);

  // Авто-пометка дублей этой же заявки в очереди, чтобы они не висели "На проверке".
  const allLast = sh.getLastRow();
  if (allLast >= 2) {
    const all = sh.getRange(2, 1, allLast - 1, 19).getValues();
    const currentSig = payQueueSignatureFromRow_(r);
    for (let i = 0; i < all.length; i++) {
      const absRow = i + 2;
      if (absRow === queueRow) continue;
      if (payNorm_(all[i][15]) !== 'На проверке') continue;
      if (payQueueSignatureFromRow_(all[i]) !== currentSig) continue;
      sh.getRange(absRow, 16).setValue('Отклонено');
      sh.getRange(absRow, 19).setValue('Авто: дубликат заявки, одобрена ' + reqNo);
    }
  }

  return 'Заявка отправлена в реестр: ' + reqNo;
  } finally {
    lock.releaseLock();
  }
}

function payRejectQueueRow(queueRow, reason) {
  payAssertApprover_();
  const sh = payEnsureQueueSheet_();
  const status = payNorm_(sh.getRange(queueRow, 16).getValue());
  if (status !== 'На проверке') throw new Error('Можно отклонить только заявки со статусом "На проверке".');
  const txt = payNorm_(reason);
  if (!txt) throw new Error('Укажите причину отклонения.');
  sh.getRange(queueRow, 16).setValue('Отклонено');
  sh.getRange(queueRow, 19).setValue(txt);
  return 'Заявка отклонена.';
}

function payCreateApprovedFolder_(requestNo, counterparty, sourceFileLinks, draftFolderUrl) {
  const rootId = payGetProp_('PAYMENT_REGISTRY_ROOT_FOLDER_ID', PAY_CFG.ROOT_FOLDER_DEFAULT);
  const root = DriveApp.getFolderById(rootId);
  const safeCounterparty = String(counterparty || 'Контрагент').replace(/[^\w\sа-яА-ЯёЁ.-]/g, '').trim();
  const finalFolder = root.createFolder(requestNo + ' - ' + (safeCounterparty || 'Контрагент'));

  // Пытаемся перенести (копировать) файлы из черновой папки/ссылок менеджера.
  (sourceFileLinks || []).forEach(function (url) {
    const fileId = payExtractDriveFileId_(url);
    if (!fileId) return;
    try {
      const f = DriveApp.getFileById(fileId);
      f.makeCopy(f.getName(), finalFolder);
    } catch (e) {
      Logger.log('[PAY][WARN] Не удалось скопировать файл в финальную папку: ' + url + ' | ' + e);
    }
  });

  // Добавим короткую памятку в финальной папке, если перенос не удался.
  try {
    finalFolder.createFile(
      'README.txt',
      'Заявка: ' + requestNo + '\nЧерновая папка менеджера: ' + (draftFolderUrl || '-') + '\nЕсли части файлов нет, откройте черновую папку.'
    );
  } catch (e) {}

  return finalFolder.getUrl();
}

function payExtractDriveFileId_(url) {
  const s = String(url || '');
  const m1 = s.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m1) return m1[1];
  const m2 = s.match(/[?&]id=([a-zA-Z0-9_-]{20,})/);
  if (m2) return m2[1];
  return null;
}

/**
 * Пишет результат одобренной заявки на лист менеджера.
 *
 * Что делает:
 *  • Колонки AB/AC («Номер заявки» / «Ссылка на заявку») — обновляются как «последняя
 *    поданная заявка по строке» (поведение совместимо со старыми листами).
 *  • Дополнительно: в пару колонок «<Тип> № заявки» / «<Тип> ссылка» — зависит от
 *    paymentType (Аванс / Баланс / Отсрочка). Если этих колонок нет — создаст справа.
 *    Так на одной строке видно все три заявки и три ссылки на документы.
 *
 * В `Сводная` НЕ пишем сознательно: лист пересобирается автоматически из листов
 * менеджеров (см. `main.gs`). Запись в неё номера заявки/ссылки в AG/AH ранее
 * затирала «Период (MM/YY)» и «Плановая дата поступления» и ломала планирование.
 */
function payWriteRequestToManagerAndSummary_(managerSheetName, managerName, spec, managerRows, requestNo, folderUrl, paymentType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mgr = ss.getSheetByName(managerSheetName);
  if (!mgr || !managerRows || !managerRows.length) return;

  const ptCols = payEnsurePaymentTypeColumns_(mgr, paymentType);
  managerRows.forEach(function (r) {
    mgr.getRange(r, 28).setValue(requestNo); // AB — последняя заявка
    mgr.getRange(r, 29).setValue(folderUrl); // AC — последняя ссылка
    if (ptCols) {
      mgr.getRange(r, ptCols.noCol).setValue(requestNo);
      mgr.getRange(r, ptCols.linkCol).setValue(folderUrl);
    }
  });
}

/**
 * Гарантирует наличие пары колонок «<Тип> № заявки» / «<Тип> ссылка» в шапке листа
 * менеджера (строка PAY_CFG.HEADER_ROW). Возвращает {noCol, linkCol} 1-based.
 *
 * Поиск устойчив к разным написаниям через `payHeaderCanon_` (см. PAY_PT_ALIASES).
 * Если колонок нет — добавляет их справа от текущей последней колонки. Создание
 * происходит атомарно для типа: сначала «№», следом «ссылка», чтобы соседи всегда
 * стояли парой и порядок типов не зависел от порядка одобрений.
 */
function payEnsurePaymentTypeColumns_(sheet, paymentType) {
  const key = payCanon_(paymentType);
  const aliases = PAY_PT_ALIASES[key];
  if (!aliases) return null;
  const headerRow = PAY_CFG.HEADER_ROW;
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const map = payHeaderMap_(headers);
  let noCol = payFindHeaderCol_(map, aliases.no);
  let linkCol = payFindHeaderCol_(map, aliases.link);

  // Колонок нет — допишем справа.
  let writeAt = lastCol;
  if (noCol < 0) {
    writeAt += 1;
    sheet.getRange(headerRow, writeAt).setValue(aliases.no[0]).setFontWeight('bold');
    noCol = writeAt - 1; // переведём обратно в 0-based для единообразия
  }
  if (linkCol < 0) {
    writeAt += 1;
    sheet.getRange(headerRow, writeAt).setValue(aliases.link[0]).setFontWeight('bold');
    linkCol = writeAt - 1;
  }
  return { noCol: noCol + 1, linkCol: linkCol + 1 };
}

const PAY_PT_ALIASES = {
  'аванс': {
    no: ['Аванс № заявки', 'Аванс №', 'Аванс номер заявки', 'Аванс заявка'],
    link: ['Аванс ссылка', 'Аванс папка', 'Аванс ссылка на заявку']
  },
  'баланс': {
    no: ['Баланс № заявки', 'Баланс №', 'Баланс номер заявки', 'Баланс заявка'],
    link: ['Баланс ссылка', 'Баланс папка', 'Баланс ссылка на заявку']
  },
  'отсрочка': {
    no: ['Отсрочка № заявки', 'Отсрочка №', 'Отсрочка номер заявки', 'Отсрочка заявка'],
    link: ['Отсрочка ссылка', 'Отсрочка папка', 'Отсрочка ссылка на заявку']
  }
};

function paySyncPaidStatuses() {
  const r = paySyncPaidStatusesImpl_({});
  if (r && r.message && !r.silent) {
    SpreadsheetApp.getUi().alert(r.message);
  }
}

/**
 * @param {{ silent?: boolean }} opt
 * @returns {{ updated: number, message: string, silent: boolean }}
 */
function paySyncPaidStatusesImpl_(opt) {
  const silent = !!(opt && opt.silent);
  const sh = payEnsureQueueSheet_();
  const last = sh.getLastRow();
  if (last < 2) {
    const msg = 'Нет заявок для синхронизации.';
    if (!silent) return { updated: 0, message: msg, silent: false };
    return { updated: 0, message: msg, silent: true };
  }
  const rows = sh.getRange(2, 1, last - 1, 19).getValues();
  const registryId = payGetProp_('PAYMENT_REGISTRY_SPREADSHEET_ID', PAY_CFG.REGISTRY_ID_DEFAULT);
  const registrySheetName = payGetProp_('PAYMENT_REGISTRY_SHEET_NAME', PAY_CFG.REGISTRY_SHEET_DEFAULT);
  const regSh = SpreadsheetApp.openById(registryId).getSheetByName(registrySheetName);
  if (!regSh) throw new Error('Лист реестра не найден.');
  const regLast = regSh.getLastRow();
  const regData = regLast >= 2 ? regSh.getRange(2, 1, regLast - 1, 19).getValues() : [];
  const regMap = {};
  regData.forEach(function (r) {
    const no = payNorm_(r[17]); // 18
    if (!no) return;
    regMap[no] = { status: payNorm_(r[13]), paidDate: r[14], folder: payNorm_(r[18]) };
  });

  let updated = 0;
  for (let i = 0; i < rows.length; i++) {
    const q = rows[i];
    const status = payNorm_(q[15]);
    if (status !== 'Отправлено в реестр') continue;
    const reqNo = payNorm_(q[16]);
    if (!reqNo || !regMap[reqNo]) continue;
    const info = regMap[reqNo];
    if (info.status !== 'Оплачено' || !info.paidDate) continue;

    const managerSheet = q[3];
    const managerName = q[4];
    const spec = q[5];
    const pType = payNorm_(q[6]);
    const mgrRows = JSON.parse(q[12] || '[]');
    payApplyFactDate_(managerSheet, managerName, spec, mgrRows, pType, info.paidDate, info.folder || payNorm_(q[13]));
    sh.getRange(i + 2, 16).setValue('Оплачено (синхр.)');
    updated++;
  }
  const msg = 'Синхронизация завершена. Обновлено заявок: ' + updated;
  return { updated: updated, message: msg, silent: silent };
}

/**
 * Проставляет дату факта оплаты на лист менеджера.
 *
 * Колонки на листе менеджера (книга «01_Операционка», шапка строки 2):
 *   S(19) — «Дата факт Аванс», V(22) — «Дата Факт Баланс», Y(25) — «Дата Факт Отсрочка».
 * Эти позиции выверены по реальной разметке, не менять без обновления здесь.
 *
 * В `Сводная` сознательно не пишем: лист пересобирается автоматически и тянет
 * даты факта с листов менеджеров (см. `main.gs`). Ранее запись `folderUrl` в
 * Сводная!AH затирала «Плановая дата поступления» и ломала планирование.
 *
 * `folderUrl` дополнительно прокатываем в пару колонок по типу платежа, чтобы
 * на одной строке менеджера были видны три комплекта ссылок одновременно.
 */
function payApplyFactDate_(managerSheetName, managerName, spec, managerRows, paymentType, paidDate, folderUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mgr = ss.getSheetByName(managerSheetName);
  if (!mgr || !managerRows || !managerRows.length) return;

  let mgrCol = 19; // S — Дата факт Аванс
  if (payCanon_(paymentType) === payCanon_('Баланс')) mgrCol = 22; // V
  else if (payCanon_(paymentType) === payCanon_('Отсрочка')) mgrCol = 25; // Y

  const ptCols = payEnsurePaymentTypeColumns_(mgr, paymentType);
  managerRows.forEach(function (r) {
    mgr.getRange(r, mgrCol).setValue(paidDate);
    if (folderUrl) {
      mgr.getRange(r, 29).setValue(folderUrl); // AC — последняя ссылка
      if (ptCols) mgr.getRange(r, ptCols.linkCol).setValue(folderUrl);
    }
  });
}

/**
 * Разовая миграция: на всех менеджерских листах (имя вида «Имя MM/YY») гарантирует
 * наличие шести колонок справа от текущего конца листа:
 *   Аванс № заявки | Аванс ссылка | Баланс № заявки | Баланс ссылка | Отсрочка № заявки | Отсрочка ссылка.
 *
 * Безопасно прогонять повторно: колонки находятся по канонизированным заголовкам,
 * пары добавляются только если их нет. Уже заполненные данные не трогаются.
 */
function payMigrateManagerHeadersForPaymentTypes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const summary = { processed: 0, added: 0, sheets: [] };

  sheets.forEach(function (sh) {
    const name = sh.getName();
    if (payManagerNameFromSheet_(name) === payNorm_(name)) return;
    const beforeCols = sh.getLastColumn();
    payEnsurePaymentTypeColumns_(sh, 'Аванс');
    payEnsurePaymentTypeColumns_(sh, 'Баланс');
    payEnsurePaymentTypeColumns_(sh, 'Отсрочка');
    const afterCols = sh.getLastColumn();
    summary.processed += 1;
    summary.added += Math.max(0, afterCols - beforeCols);
    summary.sheets.push(name + ' (+' + Math.max(0, afterCols - beforeCols) + ')');
  });

  if (!summary.processed) {
    SpreadsheetApp.getUi().alert('Не нашёл листов менеджеров (формат имени «Имя MM/YY»).');
    return;
  }
  SpreadsheetApp.getUi().alert(
    'Шапки обновлены.\n' +
    'Листов обработано: ' + summary.processed + '\n' +
    'Колонок добавлено суммарно: ' + summary.added + '\n\n' +
    summary.sheets.join('\n')
  );
}
