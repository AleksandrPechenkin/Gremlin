/**
 * Книга 05 — приёмка на складе РФ: импорт файла 3PL, сверка, расхождения.
 */

const RECEIPT_CFG = {
  CHUNK_SIZE: 450000,
  DRAFT_FOLDER: '_DRAFT_RECEIPT_3PL',
  PROPS: {
    CUSTOMS_QTY_TOLERANCE: 'RECEIPT_CUSTOMS_QTY_TOLERANCE'
  },
  SHEETS: {
    JOURNAL: 'Приемка_журнал',
    LINES: 'Приемка_строки',
    DISCREPANCIES: 'Расхождения_приемки',
    MAPPING: 'Справочник_маппинг_3PL'
  },
  JOURNAL_HEADER: [
    'Загрузка_ID', 'Загрузка_время', 'SHIPMENT_ID', 'Имя_файла', 'Пресет_3PL',
    'Строк_импорта', 'Дата_приемки', 'Статус', 'Ошибки'
  ],
  LINES_HEADER: [
    'Загрузка_ID', 'Загрузка_время', 'SHIPMENT_ID', 'Строка_ключ', 'Артикул_ВБ', 'ШК',
    'Номер_спецификации', 'Заявка_3PL_шт', 'Принято_шт', 'Годных_шт', 'Разница_3PL_шт',
    'Брак_шт', 'Комментарий_3PL', 'Статус_матча'
  ],
  DISC_HEADER: [
    'Загрузка_ID', 'SHIPMENT_ID', 'Артикул_ВБ', 'ШК', 'Номер_спецификации', 'Тип_расхождения',
    'Ожид_партии', 'Ожид_ДТ', 'Факт_принято', 'Факт_годных', 'Δ_шт', 'Статус', 'Нужна_претензия',
    'Сумма_претензии_RUB', 'Комментарий', 'Обновлено'
  ],
  MAPPING_HEADER: [
    'Пресет', 'Активен', 'Кол_артикул', 'Кол_штрихкод', 'Кол_количество', 'Кол_годных',
    'Кол_брак', 'Кол_спецификация', 'Кол_рейс', 'Кол_комментарий', 'Первая_строка_данных'
  ],
  DISC_TYPES: {
    SHORTAGE: 'SHORTAGE',
    SURPLUS: 'SURPLUS',
    DEFECT: 'DEFECT',
    ASSORTMENT: 'ASSORTMENT',
    CUSTOMS: 'CUSTOMS_VS_PHYSICAL'
  },
  STATUS: {
    OPEN: 'открыто',
    CONFIRMED: 'подтверждено',
    CLAIM: 'претензия',
    CLOSED: 'закрыто'
  }
};

function addCostingReceiptMenu_(ui) {
  ui.createMenu('📥 Приёмка (3PL)')
    .addItem('Создать листы приёмки', 'costingReceiptEnsureSheetsMenu_')
    .addSeparator()
    .addItem('Загрузить файл 3PL (CSV / XLSX)', 'costingReceiptOpenUploadDialog_')
    .addSeparator()
    .addItem('Сверка по рейсу (dry-run)', 'costingReceiptReconcileDryRunMenu_')
    .addItem('Сверка по рейсу (записать расхождения)', 'costingReceiptReconcileLiveMenu_')
    .addSeparator()
    .addItem('Применить подтверждённые к партиям', 'costingReceiptApplyConfirmedMenu_')
    .addItem('Записать событие RECEIPT_ACCEPTED', 'costingReceiptLogisticsEventMenu_')
    .addToUi();
}

function costingReceiptEnsureSheetsMenu_() {
  const r = costingReceiptEnsureSheets_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Приёмка', r.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {string[]}
 */
function costingReceiptEnsureSheets_(ss) {
  const lines = [];
  costingReceiptEnsureSheetWithHeader_(ss, RECEIPT_CFG.SHEETS.JOURNAL, RECEIPT_CFG.JOURNAL_HEADER);
  lines.push('«' + RECEIPT_CFG.SHEETS.JOURNAL + '» — OK');
  const linesSh = costingReceiptEnsureSheetWithHeader_(ss, RECEIPT_CFG.SHEETS.LINES, RECEIPT_CFG.LINES_HEADER);
  costingReceiptAlignLinesHeader_(linesSh);
  lines.push('«' + RECEIPT_CFG.SHEETS.LINES + '» — OK');
  costingReceiptEnsureSheetWithHeader_(ss, RECEIPT_CFG.SHEETS.DISCREPANCIES, RECEIPT_CFG.DISC_HEADER);
  lines.push('«' + RECEIPT_CFG.SHEETS.DISCREPANCIES + '» — OK');
  const mapSh = costingReceiptEnsureSheetWithHeader_(ss, RECEIPT_CFG.SHEETS.MAPPING, RECEIPT_CFG.MAPPING_HEADER);
  if (mapSh.getLastRow() < 2) {
    mapSh.getRange(2, 1, 3, RECEIPT_CFG.MAPPING_HEADER.length).setValues([
      [
        'default', 'да',
        'Артикул ВБ;Артикул WB;SKU;Артикул',
        'ШК;Barcode;Штрихкод',
        'Количество;Принято;Qty;Кол-во;По факту (стало);По факту',
        'Годных;Годное;Accepted good',
        'Брак;Бой;Defect',
        'Спецификация;Номер спецификации',
        'Рейс;SHIPMENT_ID',
        'Комментарий;Примечание',
        2
      ],
      [
        'расхождение_3pl', 'да',
        'Артикул;Артикул ВБ;Артикул WB',
        'ШК;Barcode',
        'По факту (стало);По факту;стало;Количество;Принято',
        'По факту (стало);По факту',
        'Брак;Бой',
        'Спецификация',
        'Рейс;SHIPMENT_ID',
        'Наименование;Комментарий',
        2
      ],
      [
        'auto', 'да',
        'Артикул ВБ;Артикул WB;SKU;Артикул',
        'ШК;Barcode;Штрихкод',
        'Количество;Принято;По факту (стало);По факту;Qty',
        'Годных;По факту (стало)',
        'Брак;Бой',
        'Спецификация;Номер спецификации',
        'Рейс;SHIPMENT_ID',
        'Наименование;Комментарий',
        2
      ]
    ]);
    lines.push('Шаблоны пресетов default / расхождение_3pl / auto в «' + RECEIPT_CFG.SHEETS.MAPPING + '».');
  }
  const batchMsg = costingReceiptEnsureBatchesAllocQtyColumn_(ss);
  if (batchMsg) lines.push(batchMsg);
  const addedPresets = costingReceiptEnsureMappingPresets_(mapSh);
  if (addedPresets) lines.push(addedPresets);
  return lines;
}

/** Добавляет пресеты расхождение_3pl / auto, если лист создан раньше без них. */
function costingReceiptEnsureMappingPresets_(mapSh) {
  if (!mapSh || mapSh.getLastRow() < 2) return '';
  const data = mapSh.getDataRange().getValues();
  const have = {};
  for (let r = 1; r < data.length; r++) {
    have[String(data[r][0] || '').trim().toLowerCase()] = true;
  }
  const toAdd = [];
  if (!have['расхождение_3pl']) {
    toAdd.push([
      'расхождение_3pl', 'да',
      'Артикул;Артикул ВБ;Артикул WB',
      'ШК;Barcode',
      'По факту (стало);По факту;Количество;Принято',
      'По факту (стало);По факту',
      'Брак;Бой',
      'Спецификация',
      'Рейс;SHIPMENT_ID',
      'Наименование;Комментарий',
      2
    ]);
  }
  if (!have['auto']) {
    toAdd.push([
      'auto', 'да',
      'Артикул ВБ;Артикул WB;SKU;Артикул',
      'ШК;Barcode;Штрихкод',
      'Количество;Принято;По факту (стало);По факту;Qty',
      'По факту (стало);Годных',
      'Брак;Бой',
      'Спецификация;Номер спецификации',
      'Рейс;SHIPMENT_ID',
      'Наименование;Комментарий',
      2
    ]);
  }
  if (!toAdd.length) return '';
  const start = mapSh.getLastRow() + 1;
  mapSh.getRange(start, 1, toAdd.length, RECEIPT_CFG.MAPPING_HEADER.length).setValues(toAdd);
  return 'Добавлены пресеты: ' + toAdd.map(function (row) { return row[0]; }).join(', ');
}

/** Подгоняет шапку «Приемка_строки» к актуальному RECEIPT_CFG.LINES_HEADER (миграция). */
function costingReceiptAlignLinesHeader_(sh) {
  if (!sh) return;
  const target = RECEIPT_CFG.LINES_HEADER;
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const row1 = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (costingHeaderCanonForLookup_(row1[0]) !== costingHeaderCanonForLookup_(target[0])) return;
  const idxOld = {};
  let c;
  for (c = 0; c < row1.length; c++) {
    const k = costingHeaderCanonForLookup_(row1[c]);
    if (k) idxOld[k] = c;
  }
  if (idxOld[costingHeaderCanonForLookup_('Заявка_3PL_шт')] != null &&
      idxOld[costingHeaderCanonForLookup_('Разница_3PL_шт')] != null) {
    return;
  }
  const lr = sh.getLastRow();
  const newData = [];
  let r;
  for (r = 2; r <= lr; r++) {
    const old = sh.getRange(r, 1, 1, lastCol).getValues()[0];
    const nr = new Array(target.length).fill('');
    let t;
    for (t = 0; t < target.length; t++) {
      const k = costingHeaderCanonForLookup_(target[t]);
      if (idxOld[k] != null) nr[t] = old[idxOld[k]];
    }
    newData.push(nr);
  }
  sh.getRange(1, 1, 1, target.length).setValues([target]).setFontWeight('bold');
  sh.setFrozenRows(1);
  if (newData.length) {
    sh.getRange(2, 1, newData.length, target.length).setValues(newData);
  }
  if (lr > newData.length + 1) {
    sh.deleteRows(newData.length + 2, lr - newData.length - 1);
  }
}

function costingReceiptEnsureSheetWithHeader_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const cur = sh.getLastRow() >= 1 ? sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), header.length)).getValues()[0] : [];
  const needRewrite = !cur[0] || costingHeaderCanonForLookup_(cur[0]) !== costingHeaderCanonForLookup_(header[0]);
  if (needRewrite || sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Колонка Qty_для_аллокации в «Партии_в_рейсе» — для себестоимости после приёмки. */
function costingReceiptEnsureBatchesAllocQtyColumn_(ss) {
  const sh = typeof costingFindSheetByRole_ === 'function'
    ? costingFindSheetByRole_(ss, 'BATCHES')
    : ss.getSheetByName(COST_CFG.SHEETS.BATCHES);
  if (!sh || sh.getLastRow() < 1) return '«Партии_в_рейсе» не найден — колонка Qty_для_аллокации не добавлена.';
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const aliases = ['Qty_для_аллокации', 'Qty для аллокации', 'Количество_для_аллокации'];
  if (costingFindColOptional_(hdr, aliases) != null) return '';
  const ncol = hdr.length + 1;
  sh.getRange(1, ncol).setValue('Qty_для_аллокации').setFontWeight('bold');
  return 'Добавлена колонка «Qty_для_аллокации» в «Партии_в_рейсе».';
}

function costingReceiptLineKey_(shipmentId, article, spec, barcode) {
  return [
    String(shipmentId || '').trim(),
    String(article || '').trim(),
    String(spec || '').trim(),
    String(barcode || '').trim()
  ].join('\t');
}

function costingReceiptImportRowKey_(uploadId, lineKey) {
  return String(uploadId || '').trim() + '\t' + String(lineKey || '').trim();
}

/* ---------- Upload dialog ---------- */

function costingReceiptOpenUploadDialog_() {
  costingReceiptEnsureSheets_(SpreadsheetApp.getActiveSpreadsheet());
  const html = HtmlService.createHtmlOutput(costingReceiptUploadDialogHtml_())
    .setWidth(580)
    .setHeight(460);
  SpreadsheetApp.getUi().showModalDialog(html, 'Загрузка файла 3PL');
}

function costingReceiptUploadDialogHtml_() {
  return (
    '<!doctype html><html><head><base target="_top"><meta charset="utf-8">' +
    '<style>body{font-family:Arial;padding:14px;font-size:13px}.row{margin:10px 0}label{display:block;font-weight:600;margin-bottom:4px}' +
    'input,select{width:100%;padding:7px;box-sizing:border-box}.note{font-size:11px;color:#666;margin-top:8px}button{padding:8px 14px}</style></head><body>' +
    '<div class="row"><label>SHIPMENT_ID (рейс)</label><input id="shipment" placeholder="TR-2026-0010"></div>' +
    '<div class="row"><label>Пресет маппинга</label><input id="preset" value="auto" title="auto — автошапка и колонки; расхождение_3pl — как файл склада с «По факту (стало)»"></div>' +
    '<div class="row"><label>Дата приёмки (ГГГГ-ММ-ДД)</label><input id="rdate" placeholder="сегодня — оставить пустым"></div>' +
    '<div class="row"><label>Файл CSV или XLSX</label><input id="file" type="file" accept=".csv,.xlsx,.xls,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"></div>' +
    '<div class="row"><button id="btn" onclick="upload()">Загрузить</button></div>' +
    '<p class="note">После загрузки: «Сверка по рейсу» (dry-run или запись).</p>' +
    '<script>var sending=false;var CHUNK=' + RECEIPT_CFG.CHUNK_SIZE + ';' +
    'function toB64(f){return new Promise(function(res,rej){var r=new FileReader();r.onload=function(){res(String(r.result).split(",")[1]);};r.onerror=rej;r.readAsDataURL(f);});}' +
    'function runPrep(p){return new Promise(function(ok,er){google.script.run.withSuccessHandler(ok).withFailureHandler(er).costingReceipt3plPrepare(p);});}' +
    'function runChunk(t,fid,ci,tot,name,mime,part){return new Promise(function(ok,er){google.script.run.withSuccessHandler(ok).withFailureHandler(er).costingReceipt3plUploadChunk(t,fid,ci,tot,name,mime,part);});}' +
    'function runFin(t,fid,p){return new Promise(function(ok,er){google.script.run.withSuccessHandler(ok).withFailureHandler(er).costingReceipt3plFinalize(t,fid,p);});}' +
    'async function upload(){var sid=document.getElementById("shipment").value.trim();var preset=document.getElementById("preset").value.trim()||"default";' +
    'var rdate=document.getElementById("rdate").value.trim();var fs=document.getElementById("file").files;' +
    'if(!sid){alert("Укажите SHIPMENT_ID");return;}if(!fs.length){alert("Выберите файл");return;}if(sending)return;sending=true;document.getElementById("btn").disabled=true;' +
    'try{var f=fs[0];var prep=await runPrep({shipmentId:sid,preset:preset,receiptDate:rdate});var b64=await toB64(f);var total=Math.max(1,Math.ceil(b64.length/CHUNK));' +
    'for(var ci=0;ci<total;ci++){await runChunk(prep.uploadToken,prep.folderId,ci,total,f.name,f.type||"application/octet-stream",b64.substring(ci*CHUNK,(ci+1)*CHUNK));}' +
    'alert(await runFin(prep.uploadToken,prep.folderId,{shipmentId:sid,fileName:f.name,preset:preset,receiptDate:rdate}));google.script.host.close();}' +
    'catch(e){alert(e.message||e);}finally{sending=false;document.getElementById("btn").disabled=false;}}</script></body></html>'
  );
}

function costingReceipt3plPrepare(payload) {
  return costingReceipt3plPrepare_(payload);
}

function costingReceipt3plUploadChunk(uploadToken, folderId, chunkIndex, totalChunks, fileName, mimeType, base64Chunk) {
  return costingReceipt3plUploadChunk_(uploadToken, folderId, chunkIndex, totalChunks, fileName, mimeType, base64Chunk);
}

function costingReceipt3plFinalize(uploadToken, folderId, payload) {
  return costingReceipt3plFinalize_(uploadToken, folderId, payload);
}

function costingReceipt3plPrepare_(payload) {
  const shipmentId = String(payload && payload.shipmentId || '').trim();
  if (!shipmentId) throw new Error('SHIPMENT_ID не указан.');
  costingReceiptValidateShipment_(shipmentId);
  const root = costingReceiptGetDraftRootFolder_();
  const uploadToken = Utilities.getUuid();
  const folder = root.createFolder(uploadToken + '_' + shipmentId.replace(/[^\w\-]+/g, '_'));
  folder.createFile(Utilities.newBlob(JSON.stringify({
    uploadToken: uploadToken,
    folderId: folder.getId(),
    shipmentId: shipmentId,
    createdAt: new Date().toISOString()
  }), 'application/json', '__meta.json'));
  return { uploadToken: uploadToken, folderId: folder.getId() };
}

function costingReceipt3plUploadChunk_(uploadToken, folderId, chunkIndex, totalChunks, fileName, mimeType, base64Chunk) {
  if (!folderId) throw new Error('Сессия загрузки не инициализирована.');
  const folder = DriveApp.getFolderById(folderId);
  folder.createFile(Utilities.newBlob(String(base64Chunk || ''), 'text/plain', '__chunk_' + chunkIndex + '_of_' + totalChunks + '.txt'));
  if (chunkIndex === 0) {
    folder.createFile(Utilities.newBlob(
      JSON.stringify({ fileName: fileName, mimeType: mimeType || '', totalChunks: totalChunks }),
      'application/json',
      '__file_info.json'
    ));
  }
  return { ok: true };
}

function costingReceipt3plFinalize_(uploadToken, folderId, payload) {
  const shipmentId = String(payload && payload.shipmentId || '').trim();
  const fileName = String(payload && payload.fileName || 'receipt.csv').trim();
  const preset = String(payload && payload.preset || 'default').trim() || 'default';
  let receiptDate = costingReceiptParseDate_(payload && payload.receiptDate);
  if (!receiptDate) receiptDate = new Date();

  const folder = DriveApp.getFolderById(folderId);
  const blob = costingReceiptAssembleBlobFromFolder_(folder);
  const table = costingReceiptParseTableFromBlob_(blob, fileName);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let mapping = costingReceiptLoadMappingPreset_(ss, preset);
  mapping = costingReceiptApplyAutoHeaderDetection_(table, mapping, preset);
  const parsed = costingReceiptMapTableToLines_(table, mapping, shipmentId);
  if (!parsed.length) {
    throw new Error(costingReceiptExplainEmptyImport_(table, mapping));
  }

  costingReceiptEnsureSheets_(ss);
  const uploadId = uploadToken || Utilities.getUuid();
  const now = new Date();
  const linesSh = ss.getSheetByName(RECEIPT_CFG.SHEETS.LINES);
  costingReceiptAlignLinesHeader_(linesSh);
  const rows = [];
  let unmatched = 0;
  let actDisc = 0;
  for (let i = 0; i < parsed.length; i++) {
    const L = parsed[i];
    const lineKey = costingReceiptLineKey_(shipmentId, L.article, L.spec, L.barcode);
    if (!L.article && !L.barcode) {
      L.matchStatus = 'unmatched';
      unmatched++;
    } else {
      L.matchStatus = 'imported';
    }
    if (L.qtyPlan > 0 && Math.abs(L.qtyReceived - L.qtyPlan) > 0.0001) actDisc++;
    rows.push([
      uploadId,
      now,
      shipmentId,
      lineKey,
      L.article,
      L.barcode,
      L.spec,
      L.qtyPlan,
      L.qtyReceived,
      L.qtyGood,
      L.qtyDiff3pl,
      L.qtyDefect,
      L.comment,
      L.matchStatus
    ]);
  }
  const upsert = costingReceiptUpsertLines_(linesSh, rows, uploadId);
  const journalSh = ss.getSheetByName(RECEIPT_CFG.SHEETS.JOURNAL);
  costingReceiptUpsertJournal_(journalSh, [
    uploadId,
    now,
    shipmentId,
    fileName,
    preset,
    rows.length,
    receiptDate,
    'импортировано',
    unmatched ? ('строк без артикула/ШК: ' + unmatched) : ''
  ], uploadId, fileName);

  return (
    'Импорт завершён.\n' +
    'Загрузка_ID: ' + uploadId + '\n' +
    'Строк: ' + rows.length + '\n' +
    'С расхождением в акте (было≠стало): ' + actDisc + '\n' +
    (upsert.updated ? ('Обновлено: ' + upsert.updated + '\n') : '') +
    (upsert.inserted ? ('Новых: ' + upsert.inserted + '\n') : '') +
    (unmatched ? ('⚠️ Без артикула/ШК: ' + unmatched + '\n') : '') +
    '\nКолонки «Заявка_3PL_шт» и «Разница_3PL_шт» — из файла склада.\n' +
    'Далее: «Сверка по рейсу» (акт + партии).'
  );
}

function costingReceiptAssembleBlobFromFolder_(folder) {
  const chunks = folder.getFilesByName('__file_info.json');
  if (!chunks.hasNext()) throw new Error('Нет метаданных файла в сессии загрузки.');
  const info = JSON.parse(chunks.next().getBlob().getDataAsString());
  const total = Number(info.totalChunks) || 1;
  let b64 = '';
  for (let i = 0; i < total; i++) {
    const it = folder.getFilesByName('__chunk_' + i + '_of_' + total + '.txt');
    if (!it.hasNext()) throw new Error('Не найден chunk ' + i);
    b64 += it.next().getBlob().getDataAsString();
  }
  const bytes = Utilities.base64Decode(b64);
  const fname = info.fileName || 'file.csv';
  const mime = costingReceiptMimeFromFileName_(fname) || info.mimeType || 'text/csv';
  return Utilities.newBlob(bytes, mime, fname);
}

function costingReceiptMimeFromFileName_(fileName) {
  const n = String(fileName || '').toLowerCase();
  if (n.indexOf('.xlsx') >= 0) {
    return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  }
  if (n.indexOf('.xls') >= 0) {
    return 'application/vnd.ms-excel';
  }
  if (n.indexOf('.csv') >= 0) {
    return 'text/csv';
  }
  return '';
}

function costingReceiptParseTableFromBlob_(blob, fileName) {
  const name = String(fileName || blob.getName() || '').toLowerCase();
  if (name.indexOf('.xlsx') >= 0 || name.indexOf('.xls') >= 0 ||
      String(blob.getContentType() || '').indexOf('spreadsheet') >= 0) {
    return costingReceiptReadXlsxTable_(blob, fileName);
  }
  const text = blob.getDataAsString('UTF-8');
  if (!text) throw new Error('Пустой файл.');
  return Utilities.parseCsv(text);
}

function costingReceiptReadXlsxTable_(blob, fileName) {
  try {
    const convertedId = costingReceiptDriveConvertToSheetId_(blob, fileName);
    const tmp = SpreadsheetApp.openById(convertedId);
    const sh = tmp.getSheets()[0];
    const lr = sh.getLastRow();
    const lc = sh.getLastColumn();
    if (lr < 1 || lc < 1) throw new Error('Пустой лист после конвертации XLSX.');
    const data = sh.getRange(1, 1, lr, lc).getValues();
    try {
      DriveApp.getFileById(convertedId).setTrashed(true);
    } catch (e) { /* ignore */ }
    return data;
  } catch (e) {
    throw new Error(
      'Не удалось прочитать XLSX: ' + (e.message || e) +
      '. Сохраните файл как CSV UTF-8 и загрузите снова.'
    );
  }
}

/** Конвертация Office → Google Sheet через Drive API v3 (OAuth токен скрипта). */
function costingReceiptDriveConvertToSheetId_(blob, fileName) {
  const url = 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id';
  const boundary = '-------receipt3pl' + Utilities.getUuid().replace(/-/g, '');
  const metadata = JSON.stringify({
    name: 'receipt_tmp_' + Date.now(),
    mimeType: MimeType.GOOGLE_SHEETS
  });
  const payload =
    '--' + boundary + '\r\n' +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    metadata + '\r\n' +
    '--' + boundary + '\r\n' +
    'Content-Type: ' + (blob.getContentType() || 'application/octet-stream') + '\r\n\r\n';
  const ending = '\r\n--' + boundary + '--';
  const body = Utilities.newBlob(payload).getBytes()
    .concat(blob.getBytes())
    .concat(Utilities.newBlob(ending).getBytes());
  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'multipart/related; boundary=' + boundary,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: body,
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() >= 300) {
    throw new Error('Drive API: ' + resp.getResponseCode() + ' ' + resp.getContentText().slice(0, 200));
  }
  return JSON.parse(resp.getContentText()).id;
}

function costingReceiptLoadMappingPreset_(ss, presetName) {
  const sh = ss.getSheetByName(RECEIPT_CFG.SHEETS.MAPPING);
  if (!sh || sh.getLastRow() < 2) {
    return costingReceiptDefaultMapping_();
  }
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const idxPreset = costingFindColOptional_(h, ['Пресет', 'Preset']) != null
    ? costingFindColOptional_(h, ['Пресет', 'Preset']) : 0;
  const idxActive = costingFindColOptional_(h, ['Активен', 'Active']);
  const cols = {
    article: costingFindColOptional_(h, ['Кол_артикул']) != null ? costingFindColOptional_(h, ['Кол_артикул']) : 2,
    barcode: costingFindColOptional_(h, ['Кол_штрихкод']) != null ? costingFindColOptional_(h, ['Кол_штрихкод']) : 3,
    qty: costingFindColOptional_(h, ['Кол_количество']) != null ? costingFindColOptional_(h, ['Кол_количество']) : 4,
    good: costingFindColOptional_(h, ['Кол_годных']) != null ? costingFindColOptional_(h, ['Кол_годных']) : 5,
    defect: costingFindColOptional_(h, ['Кол_брак']) != null ? costingFindColOptional_(h, ['Кол_брак']) : 6,
    spec: costingFindColOptional_(h, ['Кол_спецификация']) != null ? costingFindColOptional_(h, ['Кол_спецификация']) : 7,
    ship: costingFindColOptional_(h, ['Кол_рейс']) != null ? costingFindColOptional_(h, ['Кол_рейс']) : 8,
    comment: costingFindColOptional_(h, ['Кол_комментарий']) != null ? costingFindColOptional_(h, ['Кол_комментарий']) : 9,
    dataStart: costingFindColOptional_(h, ['Первая_строка_данных']) != null ? costingFindColOptional_(h, ['Первая_строка_данных']) : 10
  };
  const want = String(presetName || 'default').trim().toLowerCase();
  for (let r = 1; r < data.length; r++) {
    const p = String(data[r][idxPreset] || '').trim().toLowerCase();
    if (p !== want) continue;
    if (idxActive != null) {
      const act = String(data[r][idxActive] || '').trim().toLowerCase();
      if (act === 'нет' || act === 'no' || act === '0' || act === 'false') continue;
    }
    return costingReceiptRowToMapping_(data[r], cols);
  }
  return costingReceiptDefaultMapping_();
}

function costingReceiptDefaultMapping_() {
  return {
    articleAliases: ['Артикул ВБ', 'Артикул WB', 'SKU', 'Артикул'],
    barcodeAliases: ['ШК', 'Barcode', 'Штрихкод'],
    qtyAliases: [
      'Количество', 'Принято', 'Qty', 'Кол-во', 'Принято шт',
      'По факту (стало)', 'По факту', 'по факту стало', 'Факт'
    ],
    qtyPlanAliases: ['По заявке (было)', 'По заявке', 'было', 'Заявка'],
    diffAliases: ['Разница', 'Δ', 'разница'],
    goodAliases: ['Годных', 'Годное', 'Accepted', 'По факту (стало)', 'По факту'],
    defectAliases: ['Брак', 'Бой', 'Defect'],
    specAliases: ['Спецификация', 'Номер спецификации'],
    shipAliases: ['Рейс', 'SHIPMENT_ID'],
    nameAliases: ['Наименование', 'Номенклатура', 'Товар'],
    commentAliases: ['Комментарий', 'Примечание'],
    dataStartRow: 2,
    headerRow: 1
  };
}

/** Сканирует первые строки файла и подставляет номер строки шапки. */
function costingReceiptDetectHeaderRow_(table) {
  const limit = Math.min(table.length, 15);
  let bestIdx = 0;
  let bestScore = 0;
  for (let r = 0; r < limit; r++) {
    const row = table[r] || [];
    let score = 0;
    for (let c = 0; c < row.length; c++) {
      const key = costingHeaderCanonForLookup_(row[c]);
      if (key.indexOf('артикул') >= 0) score += 4;
      if (key.indexOf('наименование') >= 0) score += 1;
      if (key.indexOf('факт') >= 0 || key.indexOf('стало') >= 0) score += 3;
      if (key.indexOf('заявк') >= 0 || key.indexOf('было') >= 0) score += 2;
      if (key.indexOf('разниц') >= 0) score += 1;
      if (key === 'шк' || key.indexOf('штрих') >= 0) score += 2;
      if (key.indexOf('принято') >= 0 || key.indexOf('количество') >= 0) score += 2;
    }
    if (score > bestScore) {
      bestScore = score;
      bestIdx = r;
    }
  }
  return {
    headerRow: bestIdx + 1,
    dataStartRow: bestIdx + 2,
    score: bestScore
  };
}

function costingReceiptApplyAutoHeaderDetection_(table, mapping, presetName) {
  const out = mapping || costingReceiptDefaultMapping_();
  const preset = String(presetName || '').trim().toLowerCase();
  const detected = costingReceiptDetectHeaderRow_(table);
  if (preset === 'auto' || detected.score >= 4) {
    out.headerRow = detected.headerRow;
    out.dataStartRow = detected.dataStartRow;
  }
  if (preset === 'auto' || preset === 'расхождение_3pl') {
    costingReceiptEnrichMappingDiscrepancy_(out);
  }
  return out;
}

function costingReceiptEnrichMappingDiscrepancy_(mapping) {
  const merge = function (base, extra) {
    const seen = {};
    const out = [];
    let i;
    for (i = 0; i < base.length; i++) {
      const k = costingHeaderCanonForLookup_(base[i]);
      if (!k || seen[k]) continue;
      seen[k] = true;
      out.push(base[i]);
    }
    for (i = 0; i < extra.length; i++) {
      const k = costingHeaderCanonForLookup_(extra[i]);
      if (!k || seen[k]) continue;
      seen[k] = true;
      out.push(extra[i]);
    }
    return out;
  };
  mapping.articleAliases = merge(mapping.articleAliases || [], ['Артикул', 'Артикул ВБ']);
  mapping.qtyAliases = merge(mapping.qtyAliases || [], ['По факту (стало)', 'По факту', 'Количество']);
  mapping.qtyPlanAliases = merge(mapping.qtyPlanAliases || [], ['По заявке (было)', 'По заявке', 'было']);
  mapping.diffAliases = merge(mapping.diffAliases || [], ['Разница', 'Δ', 'разница']);
  mapping.goodAliases = merge(mapping.goodAliases || [], ['По факту (стало)', 'По факту']);
  mapping.nameAliases = merge(mapping.nameAliases || [], ['Наименование']);
  return mapping;
}

/** ШК из текста «… ШК:2628» / «ШК:0006» в колонке наименования. */
function costingReceiptExtractBarcodeFromText_(text) {
  const s = String(text || '');
  let m = s.match(/шк\s*[:#]?\s*([0-9]{3,})/i);
  if (m) return String(m[1]).trim();
  m = s.match(/шк\s*[:#]?\s*(\d+)/i);
  if (m) {
    const digits = String(m[1]).trim();
    return digits.length < 4 ? ('0000' + digits).slice(-4) : digits;
  }
  return '';
}

function costingReceiptIsTotalRow_(article, nameText) {
  const a = costingHeaderCanonForLookup_(article);
  const n = costingHeaderCanonForLookup_(nameText);
  return a.indexOf('итого') >= 0 || n.indexOf('итого') >= 0;
}

function costingReceiptExplainEmptyImport_(table, mapping) {
  const det = costingReceiptDetectHeaderRow_(table);
  const hdr = (table[det.headerRow - 1] || []).map(function (c) {
    return String(c || '').trim();
  }).filter(Boolean);
  return (
    'Не распознано ни одной строки приёмки.\n' +
    'Шапка (строка ' + det.headerRow + '): ' + (hdr.join(' | ') || '—') + '\n' +
    'Проверьте пресет (auto / расхождение_3pl) и лист «Справочник_маппинг_3PL».\n' +
    'Для файла «Расхождение по приемке» нужны колонки «Артикул» и «По факту (стало)».'
  );
}

function costingReceiptRowToMapping_(row, cols) {
  const splitAliases = function (cell) {
    return String(cell || '').split(/[;|]/).map(function (s) {
      return String(s).trim();
    }).filter(Boolean);
  };
  const dr = costingToNumber_(row[cols.dataStart]);
  return {
    articleAliases: splitAliases(row[cols.article]),
    barcodeAliases: splitAliases(row[cols.barcode]),
    qtyAliases: splitAliases(row[cols.qty]),
    goodAliases: splitAliases(row[cols.good]),
    defectAliases: splitAliases(row[cols.defect]),
    specAliases: splitAliases(row[cols.spec]),
    shipAliases: splitAliases(row[cols.ship]),
    commentAliases: splitAliases(row[cols.comment]),
    dataStartRow: dr >= 2 ? dr : 2,
    headerRow: Math.max(1, (dr >= 2 ? dr : 2) - 1)
  };
}

function costingReceiptMapTableToLines_(table, mapping, defaultShipmentId) {
  if (!table || !table.length) return [];
  const headerRow = Math.max(0, (mapping.headerRow || 1) - 1);
  const headers = table[headerRow] || [];
  const idxArticle = costingReceiptFindColInTable_(headers, mapping.articleAliases);
  const idxBarcode = costingReceiptFindColInTable_(headers, mapping.barcodeAliases);
  const idxQty = costingReceiptFindColInTable_(headers, mapping.qtyAliases);
  const idxQtyPlan = mapping.qtyPlanAliases
    ? costingReceiptFindColInTable_(headers, mapping.qtyPlanAliases)
    : null;
  let idxGood = costingReceiptFindColInTable_(headers, ['Годных', 'Годное', 'Accepted good']);
  if (idxGood == null && idxQty != null) idxGood = idxQty;
  const idxDefect = costingReceiptFindColInTable_(headers, mapping.defectAliases);
  const idxSpec = costingReceiptFindColInTable_(headers, mapping.specAliases);
  const idxShip = costingReceiptFindColInTable_(headers, mapping.shipAliases);
  const idxName = mapping.nameAliases
    ? costingReceiptFindColInTable_(headers, mapping.nameAliases)
    : null;
  const idxComment = costingReceiptFindColInTable_(headers, mapping.commentAliases);
  const idxDiff = mapping.diffAliases
    ? costingReceiptFindColInTable_(headers, mapping.diffAliases)
    : null;
  const start = Math.max(mapping.dataStartRow - 1, headerRow + 1);
  const out = [];
  const agg = {};
  for (let r = start; r < table.length; r++) {
    const row = table[r];
    if (!row || !row.length) continue;
    const nameText = idxName != null ? String(row[idxName] || '').trim() : '';
    let article = idxArticle != null ? String(row[idxArticle] || '').trim() : '';
    let barcode = idxBarcode != null ? String(row[idxBarcode] || '').trim() : '';
    if (!barcode && nameText) barcode = costingReceiptExtractBarcodeFromText_(nameText);
    if (costingReceiptIsTotalRow_(article, nameText)) continue;
    if (!article && !barcode) continue;
    const ship = idxShip != null ? String(row[idxShip] || '').trim() : defaultShipmentId;
    const spec = idxSpec != null ? String(row[idxSpec] || '').trim() : '';
    let qtyReceived = idxQty != null ? costingToNumber_(row[idxQty]) : 0;
    let qtyGood = idxGood != null ? costingToNumber_(row[idxGood]) : 0;
    const qtyDefect = idxDefect != null ? costingToNumber_(row[idxDefect]) : 0;
    if (!qtyGood && qtyReceived) qtyGood = Math.max(0, qtyReceived - qtyDefect);
    if (!qtyReceived && qtyGood) qtyReceived = qtyGood + qtyDefect;
    let comment = idxComment != null ? String(row[idxComment] || '').trim() : '';
    if (nameText && comment.indexOf(nameText) < 0) {
      comment = (comment ? comment + '; ' : '') + nameText;
    }
    const qtyPlan = idxQtyPlan != null ? costingToNumber_(row[idxQtyPlan]) : 0;
    let qtyDiff3pl = idxDiff != null ? costingToNumber_(row[idxDiff]) : NaN;
    if (!isFinite(qtyDiff3pl) && qtyPlan > 0 && qtyReceived > 0) {
      qtyDiff3pl = qtyReceived - qtyPlan;
    }
    if (!isFinite(qtyDiff3pl)) qtyDiff3pl = 0;
    const key = costingReceiptLineKey_(ship, article, spec, barcode);
    if (!agg[key]) {
      agg[key] = {
        article: article,
        barcode: barcode,
        spec: spec,
        qtyPlan: 0,
        qtyReceived: 0,
        qtyGood: 0,
        qtyDiff3pl: 0,
        qtyDefect: 0,
        comment: comment
      };
    }
    agg[key].qtyPlan += qtyPlan;
    agg[key].qtyReceived += qtyReceived;
    agg[key].qtyGood += qtyGood;
    agg[key].qtyDiff3pl += isFinite(qtyDiff3pl) ? qtyDiff3pl : 0;
    agg[key].qtyDefect += qtyDefect;
    if (comment) agg[key].comment = comment;
  }
  const keys = Object.keys(agg);
  for (let k = 0; k < keys.length; k++) out.push(agg[keys[k]]);
  return out;
}

function costingReceiptFindColInTable_(headers, aliases) {
  const norm = headers.map(costingHeaderCanonForLookup_);
  for (let a = 0; a < aliases.length; a++) {
    const key = costingHeaderCanonForLookup_(aliases[a]);
    const idx = norm.indexOf(key);
    if (idx !== -1) return idx;
  }
  return null;
}

function costingReceiptUpsertLines_(linesSh, rows, uploadId) {
  const ncol = RECEIPT_CFG.LINES_HEADER.length;
  let inserted = 0;
  let updated = 0;
  const lr = linesSh.getLastRow();
  const existing = {};
  if (lr >= 2) {
    const data = linesSh.getRange(2, 1, lr - 1, ncol).getValues();
    const idxUpload = RECEIPT_CFG.LINES_HEADER.indexOf('Загрузка_ID');
    const idxKey = RECEIPT_CFG.LINES_HEADER.indexOf('Строка_ключ');
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][idxUpload] || '').trim() === uploadId) {
        existing[String(data[i][idxKey] || '').trim()] = i + 2;
      }
    }
  }
  const toAppend = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const key = String(row[RECEIPT_CFG.LINES_HEADER.indexOf('Строка_ключ')] || '').trim();
    if (existing[key]) {
      linesSh.getRange(existing[key], 1, 1, ncol).setValues([row]);
      updated++;
    } else {
      toAppend.push(row);
      inserted++;
    }
  }
  if (toAppend.length) {
    const start = linesSh.getLastRow() + 1;
    linesSh.getRange(start, 1, toAppend.length, ncol).setValues(toAppend);
  }
  return { inserted: inserted, updated: updated };
}

function costingReceiptUpsertJournal_(journalSh, rowValues, uploadId, fileName) {
  const lr = journalSh.getLastRow();
  if (lr < 2) {
    journalSh.appendRow(rowValues);
    return;
  }
  const data = journalSh.getDataRange().getValues();
  const idxUpload = costingFindColOptional_(data[0], ['Загрузка_ID']) != null
    ? costingFindColOptional_(data[0], ['Загрузка_ID']) : 0;
  const idxFile = costingFindColOptional_(data[0], ['Имя_файла']);
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][idxUpload] || '').trim() === uploadId &&
        (!idxFile || String(data[r][idxFile] || '').trim() === fileName)) {
      journalSh.getRange(r + 1, 1, 1, rowValues.length).setValues([rowValues]);
      return;
    }
  }
  journalSh.appendRow(rowValues);
}

/* ---------- Reconcile ---------- */

function costingReceiptReconcileDryRunMenu_() {
  costingReceiptReconcileMenu_(true);
}

function costingReceiptReconcileLiveMenu_() {
  costingReceiptReconcileMenu_(false);
}

function costingReceiptReconcileMenu_(dryRun) {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    dryRun ? 'Сверка (dry-run)' : 'Сверка (запись)',
    'SHIPMENT_ID и опционально Загрузка_ID (пусто = последняя загрузка по рейсу):',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const parts = String(res.getResponseText() || '').trim().split(/[,;\s]+/);
  const shipmentId = parts[0] || '';
  const uploadId = parts[1] || '';
  if (!shipmentId) {
    ui.alert('Укажите SHIPMENT_ID.');
    return;
  }
  try {
    const report = costingReceiptReconcile_(SpreadsheetApp.getActiveSpreadsheet(), {
      shipmentId: shipmentId,
      uploadId: uploadId,
      dryRun: dryRun,
      writeLogisticsEvent: false
    });
    ui.alert(dryRun ? 'Dry-run сверки' : 'Сверка записана', report.summary, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Ошибка', e.message || String(e), ui.ButtonSet.OK);
  }
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {{ shipmentId: string, uploadId?: string, dryRun?: boolean, writeLogisticsEvent?: boolean, receiptDate?: Date }} opt
 */
function costingReceiptReconcile_(ss, opt) {
  opt = opt || {};
  const shipmentId = String(opt.shipmentId || '').trim();
  if (!shipmentId) throw new Error('SHIPMENT_ID не указан.');
  costingReceiptEnsureSheets_(ss);
  let uploadId = String(opt.uploadId || '').trim();
  if (!uploadId) uploadId = costingReceiptFindLatestUploadForShipment_(ss, shipmentId);
  if (!uploadId) throw new Error('Нет загрузок приёмки для рейса ' + shipmentId + '.');

  const receiptLines = costingReceiptLoadReceiptLines_(ss, shipmentId, uploadId);
  const batches = costingReceiptLoadBatchesIndex_(ss, shipmentId);
  const declQty = costingReceiptLoadDeclQtyBySku_(ss, shipmentId);
  const tolerance = costingReceiptGetCustomsTolerance_();

  const discRows = [];
  const seenBatch = {};
  let shortage = 0;
  let surplus = 0;
  let defect = 0;
  let customsDiff = 0;
  const now = new Date();

  for (let i = 0; i < receiptLines.length; i++) {
    const L = receiptLines[i];
    if (L.matchStatus === 'unmatched' || (!L.article && !L.barcode)) {
      discRows.push(costingReceiptBuildDiscRow_({
        uploadId: uploadId,
        shipmentId: shipmentId,
        article: L.article,
        barcode: L.barcode,
        spec: L.spec,
        type: RECEIPT_CFG.DISC_TYPES.ASSORTMENT,
        expectedBatch: 0,
        expectedDecl: 0,
        factReceived: L.qtyReceived,
        factGood: L.qtyGood,
        delta: L.qtyReceived,
        status: RECEIPT_CFG.STATUS.OPEN,
        comment: 'Не сматчилось с партией рейса',
        now: now
      }));
      continue;
    }
    const batch = costingReceiptFindBatch_(batches, L.article, L.barcode, L.spec);
    const batchKey = batch ? batch.key : '';
    if (batch) seenBatch[batchKey] = true;
    const expected = batch ? batch.qty : 0;
    const declQ = L.article ? (declQty[L.article] || 0) : 0;
    const good = L.qtyGood > 0 ? L.qtyGood : L.qtyReceived;

    if (L.qtyPlan > 0) {
      const deltaAct = L.qtyDiff3pl !== 0 && isFinite(L.qtyDiff3pl)
        ? L.qtyDiff3pl
        : (good - L.qtyPlan);
      if (Math.abs(deltaAct) > 0.0001) {
        if (deltaAct < 0) shortage++;
        else surplus++;
        discRows.push(costingReceiptBuildDiscRow_({
          uploadId: uploadId,
          shipmentId: shipmentId,
          article: L.article,
          barcode: L.barcode,
          spec: L.spec,
          type: deltaAct < 0 ? RECEIPT_CFG.DISC_TYPES.SHORTAGE : RECEIPT_CFG.DISC_TYPES.SURPLUS,
          expectedBatch: L.qtyPlan,
          expectedDecl: declQ,
          factReceived: L.qtyReceived,
          factGood: good,
          delta: deltaAct,
          status: RECEIPT_CFG.STATUS.OPEN,
          comment: 'акт 3PL: было ' + L.qtyPlan + ' → стало ' + good,
          now: now
        }));
      }
    }

    if (!batch) {
      discRows.push(costingReceiptBuildDiscRow_({
        uploadId: uploadId,
        shipmentId: shipmentId,
        article: L.article,
        barcode: L.barcode,
        spec: L.spec,
        type: RECEIPT_CFG.DISC_TYPES.ASSORTMENT,
        expectedBatch: 0,
        expectedDecl: declQ,
        factReceived: L.qtyReceived,
        factGood: good,
        delta: L.qtyReceived,
        status: RECEIPT_CFG.STATUS.OPEN,
        comment: 'Нет строки в Партии_в_рейсе',
        now: now
      }));
      continue;
    }

    const delta = good - expected;
    if (delta < -0.0001) {
      shortage++;
      const claimRub = batch.unitRub * Math.abs(delta);
      discRows.push(costingReceiptBuildDiscRow_({
        uploadId: uploadId,
        shipmentId: shipmentId,
        article: L.article,
        barcode: L.barcode,
        spec: L.spec,
        type: RECEIPT_CFG.DISC_TYPES.SHORTAGE,
        expectedBatch: expected,
        expectedDecl: declQ,
        factReceived: L.qtyReceived,
        factGood: good,
        delta: delta,
        status: RECEIPT_CFG.STATUS.OPEN,
        needClaim: claimRub > 0,
        claimRub: claimRub,
        comment: '',
        now: now
      }));
    } else if (delta > 0.0001) {
      surplus++;
      discRows.push(costingReceiptBuildDiscRow_({
        uploadId: uploadId,
        shipmentId: shipmentId,
        article: L.article,
        barcode: L.barcode,
        spec: L.spec,
        type: RECEIPT_CFG.DISC_TYPES.SURPLUS,
        expectedBatch: expected,
        expectedDecl: declQ,
        factReceived: L.qtyReceived,
        factGood: good,
        delta: delta,
        status: RECEIPT_CFG.STATUS.OPEN,
        comment: '',
        now: now
      }));
    }

    if (L.qtyDefect > 0.0001 || (L.qtyReceived > good + 0.0001)) {
      defect++;
      discRows.push(costingReceiptBuildDiscRow_({
        uploadId: uploadId,
        shipmentId: shipmentId,
        article: L.article,
        barcode: L.barcode,
        spec: L.spec,
        type: RECEIPT_CFG.DISC_TYPES.DEFECT,
        expectedBatch: expected,
        expectedDecl: declQ,
        factReceived: L.qtyReceived,
        factGood: good,
        delta: L.qtyDefect || (L.qtyReceived - good),
        status: RECEIPT_CFG.STATUS.OPEN,
        comment: 'Брак/бой при приёмке',
        now: now
      }));
    }

    if (declQ > 0 && Math.abs(declQ - good) > tolerance) {
      customsDiff++;
      discRows.push(costingReceiptBuildDiscRow_({
        uploadId: uploadId,
        shipmentId: shipmentId,
        article: L.article,
        barcode: L.barcode,
        spec: L.spec,
        type: RECEIPT_CFG.DISC_TYPES.CUSTOMS,
        expectedBatch: expected,
        expectedDecl: declQ,
        factReceived: L.qtyReceived,
        factGood: good,
        delta: good - declQ,
        status: RECEIPT_CFG.STATUS.OPEN,
        comment: 'ДТ vs физика (годные шт)',
        now: now
      }));
    }
  }

  const batchKeys = Object.keys(batches.byKey);
  for (let b = 0; b < batchKeys.length; b++) {
    const bk = batchKeys[b];
    if (seenBatch[bk]) continue;
    const batch = batches.byKey[bk];
    shortage++;
    discRows.push(costingReceiptBuildDiscRow_({
      uploadId: uploadId,
      shipmentId: shipmentId,
      article: batch.article,
      barcode: batch.barcode,
      spec: batch.spec,
      type: RECEIPT_CFG.DISC_TYPES.SHORTAGE,
      expectedBatch: batch.qty,
      expectedDecl: declQty[batch.article] || 0,
      factReceived: 0,
      factGood: 0,
      delta: -batch.qty,
      status: RECEIPT_CFG.STATUS.OPEN,
      needClaim: batch.unitRub * batch.qty > 0,
      claimRub: batch.unitRub * batch.qty,
      comment: 'Не принято на складе (нет в файле 3PL)',
      now: now
    }));
  }

  if (!opt.dryRun) {
    costingReceiptWriteDiscrepancies_(ss, shipmentId, uploadId, discRows);
    if (opt.writeLogisticsEvent) {
      costingReceiptWriteLogisticsEvent_(ss, shipmentId, opt.receiptDate || new Date());
    }
  }

  return {
    summary:
      'Рейс: ' + shipmentId + '\n' +
      'Загрузка_ID: ' + uploadId + '\n' +
      'Строк приёмки: ' + receiptLines.length + '\n' +
      'Записей расхождений: ' + discRows.length + '\n' +
      '  недостача: ' + shortage + '\n' +
      '  излишек: ' + surplus + '\n' +
      '  брак: ' + defect + '\n' +
      '  ДТ vs физика: ' + customsDiff + '\n' +
      (opt.dryRun ? '\n(dry-run — лист не изменён)' : '\nЗаписано на «' + RECEIPT_CFG.SHEETS.DISCREPANCIES + '».'),
    discCount: discRows.length,
    uploadId: uploadId
  };
}

function costingReceiptBuildDiscRow_(p) {
  return [
    p.uploadId,
    p.shipmentId,
    p.article || '',
    p.barcode || '',
    p.spec || '',
    p.type,
    p.expectedBatch,
    p.expectedDecl,
    p.factReceived,
    p.factGood,
    p.delta,
    p.status || RECEIPT_CFG.STATUS.OPEN,
    p.needClaim ? 'да' : '',
    p.claimRub ? Math.round(p.claimRub * 100) / 100 : '',
    p.comment || '',
    p.now
  ];
}

function costingReceiptWriteDiscrepancies_(ss, shipmentId, uploadId, newRows) {
  const sh = ss.getSheetByName(RECEIPT_CFG.SHEETS.DISCREPANCIES);
  const ncol = RECEIPT_CFG.DISC_HEADER.length;
  const kept = [RECEIPT_CFG.DISC_HEADER.slice()];
  if (sh.getLastRow() >= 2) {
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, ncol).getValues();
    const idxShip = RECEIPT_CFG.DISC_HEADER.indexOf('SHIPMENT_ID');
    const idxUpload = RECEIPT_CFG.DISC_HEADER.indexOf('Загрузка_ID');
    for (let i = 0; i < data.length; i++) {
      const sameShip = String(data[i][idxShip] || '').trim() === shipmentId;
      const sameUpload = String(data[i][idxUpload] || '').trim() === uploadId;
      if (sameShip && sameUpload) continue;
      kept.push(data[i]);
    }
  }
  for (let j = 0; j < newRows.length; j++) kept.push(newRows[j]);
  sh.clearContents();
  sh.getRange(1, 1, kept.length, ncol).setValues(kept);
  sh.getRange(1, 1, 1, ncol).setFontWeight('bold');
  sh.setFrozenRows(1);
}

function costingReceiptLoadReceiptLines_(ss, shipmentId, uploadId) {
  const sh = ss.getSheetByName(RECEIPT_CFG.SHEETS.LINES);
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const idxUpload = costingFindColOptional_(h, ['Загрузка_ID']);
  const idxShip = costingFindColOptional_(h, ['SHIPMENT_ID']);
  const idxArt = costingFindColOptional_(h, ['Артикул_ВБ', 'Артикул ВБ']);
  const idxBc = costingFindColOptional_(h, ['ШК']);
  const idxSpec = costingFindColOptional_(h, ['Номер_спецификации', 'Номер спецификации']);
  const idxRec = costingFindColOptional_(h, ['Принято_шт']);
  const idxGood = costingFindColOptional_(h, ['Годных_шт']);
  const idxDef = costingFindColOptional_(h, ['Брак_шт']);
  const idxMatch = costingFindColOptional_(h, ['Статус_матча']);
  const out = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (String(row[idxUpload] || '').trim() !== uploadId) continue;
    if (idxShip != null && String(row[idxShip] || '').trim() !== shipmentId) continue;
    out.push({
      article: idxArt != null ? String(row[idxArt] || '').trim() : '',
      barcode: idxBc != null ? String(row[idxBc] || '').trim() : '',
      spec: idxSpec != null ? String(row[idxSpec] || '').trim() : '',
      qtyPlan: costingFindColOptional_(h, ['Заявка_3PL_шт']) != null
        ? costingToNumber_(row[costingFindColOptional_(h, ['Заявка_3PL_шт'])]) : 0,
      qtyReceived: idxRec != null ? costingToNumber_(row[idxRec]) : 0,
      qtyGood: idxGood != null ? costingToNumber_(row[idxGood]) : 0,
      qtyDiff3pl: costingFindColOptional_(h, ['Разница_3PL_шт']) != null
        ? costingToNumber_(row[costingFindColOptional_(h, ['Разница_3PL_шт'])]) : 0,
      qtyDefect: idxDef != null ? costingToNumber_(row[idxDef]) : 0,
      matchStatus: idxMatch != null ? String(row[idxMatch] || '').trim() : ''
    });
  }
  return out;
}

function costingReceiptLoadBatchesIndex_(ss, shipmentId) {
  const sh = typeof costingFindSheetByRole_ === 'function'
    ? costingFindSheetByRole_(ss, 'BATCHES')
    : ss.getSheetByName(COST_CFG.SHEETS.BATCHES);
  const byKey = {};
  const bySupArt = {};
  if (!sh || sh.getLastRow() < 2) return { byKey: byKey, bySupArt: bySupArt, list: [] };
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const idxShip = costingFindCol_(h, ['SHIPMENT_ID', 'ID_рейса']);
  const idxSku = costingFindCol_(h, ['Артикул ВБ', 'Артикул_ВБ']);
  const idxSupArt = costingFindColOptional_(h, ['Артикул поставщика', 'Артикул_поставщика']);
  const idxBc = costingFindColOptional_(h, ['ШК', 'Barcode']);
  const idxSpec = costingFindColOptional_(h, ['Номер спецификации', 'Номер_спецификации']);
  const idxQty = costingFindCol_(h, ['Количество', 'Qty']);
  const idxPrice = costingFindColOptional_(h, ['Цена', 'Price']);
  const idxAmount = costingFindColOptional_(h, ['Сумма итого', 'Стоимость', 'Сумма']);
  const list = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (String(row[idxShip] || '').trim() !== shipmentId) continue;
    const article = String(row[idxSku] || '').trim();
    const barcode = idxBc != null ? String(row[idxBc] || '').trim() : '';
    const spec = idxSpec != null ? String(row[idxSpec] || '').trim() : '';
    const qty = costingToNumber_(row[idxQty]);
    let unitRub = 0;
    if (idxAmount != null && qty > 0) unitRub = costingToNumber_(row[idxAmount]) / qty;
    else if (idxPrice != null) unitRub = costingToNumber_(row[idxPrice]);
    const key = costingReceiptBatchKey_(article, barcode, spec);
    const entry = {
      key: key,
      row: r + 1,
      article: article,
      barcode: barcode,
      spec: spec,
      qty: qty,
      unitRub: unitRub
    };
    byKey[key] = entry;
    list.push(entry);
    if (idxSupArt != null) {
      const sup = String(row[idxSupArt] || '').trim();
      if (sup) {
        const sk = costingHeaderCanonForLookup_(sup);
        if (!bySupArt[sk]) bySupArt[sk] = entry;
      }
    }
    if (article && /[^0-9]/.test(article)) {
      const ak = costingHeaderCanonForLookup_(article);
      if (!bySupArt[ak]) bySupArt[ak] = entry;
    }
  }
  return { byKey: byKey, bySupArt: bySupArt, list: list };
}

function costingReceiptBatchKey_(article, barcode, spec) {
  return [
    String(article || '').trim().toLowerCase(),
    String(barcode || '').trim(),
    String(spec || '').trim().toLowerCase()
  ].join('\t');
}

function costingReceiptFindBatch_(batches, article, barcode, spec) {
  const keys = [
    costingReceiptBatchKey_(article, barcode, spec),
    costingReceiptBatchKey_(article, barcode, ''),
    costingReceiptBatchKey_(article, '', spec),
    costingReceiptBatchKey_(article, '', '')
  ];
  for (let i = 0; i < keys.length; i++) {
    if (batches.byKey[keys[i]]) return batches.byKey[keys[i]];
  }
  if (barcode) {
    for (let j = 0; j < batches.list.length; j++) {
      if (batches.list[j].barcode && batches.list[j].barcode === barcode) return batches.list[j];
    }
  }
  if (article && batches.bySupArt) {
    const ak = costingHeaderCanonForLookup_(article);
    if (batches.bySupArt[ak]) return batches.bySupArt[ak];
  }
  return null;
}

function costingReceiptLoadDeclQtyBySku_(ss, shipmentId) {
  const out = {};
  const sh = typeof costingFindSheetByRole_ === 'function'
    ? costingFindSheetByRole_(ss, 'DECL_LINES')
    : ss.getSheetByName(COST_CFG.SHEETS.DECL_LINES);
  if (!sh || sh.getLastRow() < 2) return out;
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const idxShip = costingFindColOptional_(h, ['SHIPMENT_ID', 'ID_рейса']);
  const idxSku = costingFindColOptional_(h, ['Артикул_ВБ', 'Артикул ВБ']);
  const idxQty = costingFindColOptional_(h, ['Qty', 'Количество']);
  const idxStatus = costingFindColOptional_(h, ['Статус_сопоставления', 'Статус сопоставления']);
  if (idxShip == null || idxSku == null || idxQty == null) return out;
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (String(row[idxShip] || '').trim() !== shipmentId) continue;
    const sku = String(row[idxSku] || '').trim();
    if (!sku) continue;
    if (idxStatus != null) {
      const st = String(row[idxStatus] || '').trim().toLowerCase();
      if (st === 'pending' || st.indexOf('needs') >= 0) continue;
    }
    const q = costingToNumber_(row[idxQty]);
    out[sku] = (out[sku] || 0) + q;
  }
  return out;
}

function costingReceiptFindLatestUploadForShipment_(ss, shipmentId) {
  const j = ss.getSheetByName(RECEIPT_CFG.SHEETS.JOURNAL);
  if (!j || j.getLastRow() < 2) return '';
  const data = j.getDataRange().getValues();
  const idxUpload = 0;
  const idxShip = RECEIPT_CFG.JOURNAL_HEADER.indexOf('SHIPMENT_ID');
  const idxTime = RECEIPT_CFG.JOURNAL_HEADER.indexOf('Загрузка_время');
  let best = '';
  let bestTime = 0;
  for (let r = data.length - 1; r >= 1; r--) {
    if (String(data[r][idxShip] || '').trim() !== shipmentId) continue;
    const t = data[r][idxTime] instanceof Date ? data[r][idxTime].getTime() : 0;
    if (t >= bestTime) {
      bestTime = t;
      best = String(data[r][idxUpload] || '').trim();
    }
  }
  return best;
}

/* ---------- Apply confirmed ---------- */

function costingReceiptApplyConfirmedMenu_() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    'Применить к партиям',
    'SHIPMENT_ID (только строки «подтверждено» в Расхождения_приемки):',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const shipmentId = String(res.getResponseText() || '').trim();
  if (!shipmentId) return;
  try {
    const r = costingReceiptApplyConfirmed_(SpreadsheetApp.getActiveSpreadsheet(), shipmentId);
    ui.alert('Применено', r.summary, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Ошибка', e.message || String(e), ui.ButtonSet.OK);
  }
}

function costingReceiptApplyConfirmed_(ss, shipmentId) {
  costingReceiptEnsureSheets_(ss);
  const discSh = ss.getSheetByName(RECEIPT_CFG.SHEETS.DISCREPANCIES);
  if (!discSh || discSh.getLastRow() < 2) throw new Error('Нет расхождений для применения.');
  const data = discSh.getDataRange().getValues();
  const h = data[0];
  const idxShip = costingFindCol_(h, ['SHIPMENT_ID']);
  const idxArt = costingFindCol_(h, ['Артикул_ВБ', 'Артикул ВБ']);
  const idxBc = costingFindColOptional_(h, ['ШК']);
  const idxSpec = costingFindColOptional_(h, ['Номер_спецификации']);
  const idxGood = costingFindCol_(h, ['Факт_годных']);
  const idxStatus = costingFindCol_(h, ['Статус']);
  const idxType = costingFindCol_(h, ['Тип_расхождения']);

  const goodByKey = {};
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (String(row[idxShip] || '').trim() !== shipmentId) continue;
    const st = String(row[idxStatus] || '').trim().toLowerCase();
    if (st !== RECEIPT_CFG.STATUS.CONFIRMED && st !== 'confirmed') continue;
    const type = String(row[idxType] || '').trim();
    if (type === RECEIPT_CFG.DISC_TYPES.ASSORTMENT) continue;
    const article = String(row[idxArt] || '').trim();
    const barcode = idxBc != null ? String(row[idxBc] || '').trim() : '';
    const spec = idxSpec != null ? String(row[idxSpec] || '').trim() : '';
    const good = costingToNumber_(row[idxGood]);
    if (!(good > 0) && type !== RECEIPT_CFG.DISC_TYPES.SHORTAGE) continue;
    const key = costingReceiptBatchKey_(article, barcode, spec);
    goodByKey[key] = good;
  }
  if (!Object.keys(goodByKey).length) {
    throw new Error('Нет строк со статусом «подтверждено». Сначала подтвердите расхождения вручную.');
  }

  const backupName = costingSnapshotBatches_(ss);
  const batchesSh = costingGetSheetByRole_(ss, 'BATCHES');
  const bh = batchesSh.getRange(1, 1, 1, batchesSh.getLastColumn()).getValues()[0];
  const idxBShip = costingFindCol_(bh, ['SHIPMENT_ID', 'ID_рейса']);
  const idxBSku = costingFindCol_(bh, ['Артикул ВБ', 'Артикул_ВБ']);
  const idxBBc = costingFindColOptional_(bh, ['ШК']);
  const idxBSpec = costingFindColOptional_(bh, ['Номер спецификации', 'Номер_спецификации']);
  const idxBQty = costingFindCol_(bh, ['Количество', 'Qty']);
  let idxAlloc = costingFindColOptional_(bh, ['Qty_для_аллокации', 'Qty для аллокации']);
  if (idxAlloc == null) {
    costingReceiptEnsureBatchesAllocQtyColumn_(ss);
    const bh2 = batchesSh.getRange(1, 1, 1, batchesSh.getLastColumn()).getValues()[0];
    idxAlloc = costingFindCol_(bh2, ['Qty_для_аллокации']);
  }
  const bData = batchesSh.getDataRange().getValues();
  let updated = 0;
  for (let i = 1; i < bData.length; i++) {
    if (String(bData[i][idxBShip] || '').trim() !== shipmentId) continue;
    const article = String(bData[i][idxBSku] || '').trim();
    const barcode = idxBBc != null ? String(bData[i][idxBBc] || '').trim() : '';
    const spec = idxBSpec != null ? String(bData[i][idxBSpec] || '').trim() : '';
    const key = costingReceiptBatchKey_(article, barcode, spec);
    if (goodByKey[key] == null) continue;
    const good = goodByKey[key];
    bData[i][idxBQty] = good;
    bData[i][idxAlloc] = good;
    batchesSh.getRange(i + 1, idxBQty + 1).setValue(good);
    batchesSh.getRange(i + 1, idxAlloc + 1).setValue(good);
    updated++;
  }
  return {
    summary:
      'Рейс: ' + shipmentId + '\n' +
      'Снимок: ' + backupName + '\n' +
      'Обновлено строк партий: ' + updated + '\n' +
      'Запустите пересчёт себестоимости для рейса.'
  };
}

/* ---------- Logistics event ---------- */

function costingReceiptLogisticsEventMenu_() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Событие приёмки', 'SHIPMENT_ID:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const shipmentId = String(res.getResponseText() || '').trim();
  if (!shipmentId) return;
  const dateRes = ui.prompt('Дата приёмки', 'ГГГГ-ММ-ДД (пусто = сегодня):', ui.ButtonSet.OK_CANCEL);
  let dt = new Date();
  if (dateRes.getSelectedButton() === ui.Button.OK && dateRes.getResponseText()) {
    const parsed = costingReceiptParseDate_(dateRes.getResponseText());
    if (parsed) dt = parsed;
  }
  try {
    costingReceiptWriteLogisticsEvent_(SpreadsheetApp.getActiveSpreadsheet(), shipmentId, dt);
    ui.alert('Записано', 'RECEIPT_ACCEPTED для ' + shipmentId, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Ошибка', e.message || String(e), ui.ButtonSet.OK);
  }
}

function costingReceiptWriteLogisticsEvent_(ss, shipmentId, eventDate) {
  if (typeof logisticsAppendEventByCode_ === 'function') {
    logisticsAppendEventByCode_(ss, shipmentId, 'RECEIPT_ACCEPTED', eventDate, 'Приёмка 3PL');
    return;
  }
  if (typeof logisticsAppendEvent_ === 'function') {
    logisticsAppendEvent_(ss, shipmentId, 'Приёмка завершена (RECEIPT_ACCEPTED)', eventDate, 'Приёмка 3PL');
    return;
  }
  throw new Error('logistics.gs не подключён в проекте 05.');
}

function costingReceiptGetDraftRootFolder_() {
  const iter = DriveApp.getFoldersByName(RECEIPT_CFG.DRAFT_FOLDER);
  if (iter.hasNext()) return iter.next();
  return DriveApp.createFolder(RECEIPT_CFG.DRAFT_FOLDER);
}

function costingReceiptValidateShipment_(shipmentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (typeof costingLoadTripsRegistry_ === 'function') {
    const reg = costingLoadTripsRegistry_(ss);
    if (Object.keys(reg).length && !reg[shipmentId]) {
      throw new Error('Рейс ' + shipmentId + ' не найден в листе «Рейсы».');
    }
  }
}

function costingReceiptGetCustomsTolerance_() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(RECEIPT_CFG.PROPS.CUSTOMS_QTY_TOLERANCE);
    const n = parseFloat(String(raw || '0').replace(',', '.'));
    return isFinite(n) && n >= 0 ? n : 0;
  } catch (e) {
    return 0;
  }
}

function costingReceiptParseDate_(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  const s = String(val).trim();
  if (!s) return null;
  const iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (iso) return new Date(parseInt(iso[1], 10), parseInt(iso[2], 10) - 1, parseInt(iso[3], 10));
  const ru = s.match(/^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})/);
  if (ru) {
    let y = parseInt(ru[3], 10);
    if (y < 100) y += 2000;
    return new Date(y, parseInt(ru[2], 10) - 1, parseInt(ru[1], 10));
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}
