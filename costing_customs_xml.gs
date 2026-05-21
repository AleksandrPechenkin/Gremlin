/**
 * Книга 05 — импорт таможенных деклараций (XML) и сопоставление с партиями рейса.
 */

const COST_XML_CHUNK_SIZE_ = 450000;
const COST_XML_DRAFT_FOLDER_ = '_DRAFT_CUSTOMS_XML';

function costingCustomsXmlOpenUploadDialog_() {
  const html = HtmlService.createHtmlOutput(costingCustomsXmlUploadDialogHtml_())
    .setWidth(560)
    .setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, 'Загрузка декларации (XML)');
}

/** HtmlService: без суффикса _, иначе google.script.run не видит функцию. */
function costingCustomsXmlPrepare(payload) {
  return costingCustomsXmlPrepare_(payload);
}

function costingCustomsXmlUploadChunk(uploadToken, folderId, chunkIndex, totalChunks, fileName, mimeType, base64Chunk) {
  return costingCustomsXmlUploadChunk_(uploadToken, folderId, chunkIndex, totalChunks, fileName, mimeType, base64Chunk);
}

function costingCustomsXmlFinalize(uploadToken, folderId, payload) {
  return costingCustomsXmlFinalize_(uploadToken, folderId, payload);
}

/**
 * Дописать строки в лист. В GAS getRange(row, col, numRows, numCols) — 3-й аргумент это
 * число строк, а не индекс последней строки (getRange(2,1,2,19) = 2 строки, не одна).
 */
function costingSheetAppendRows_(sheet, values, ncol) {
  if (!values || !values.length) return;
  const colCount = ncol || (values[0] ? values[0].length : 1);
  const normalized = values.map(function (r) {
    const row = (r || []).slice();
    while (row.length < colCount) row.push('');
    if (row.length > colCount) row.length = colCount;
    return row;
  });
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, normalized.length, colCount).setValues(normalized);
  const idxTnved = COST_CFG.DECL_LINES_HEADER.indexOf('Код_ТНВЭД');
  if (idxTnved >= 0) {
    sheet.getRange(startRow, idxTnved + 1, normalized.length, 1).setNumberFormat('@');
  }
}

/** Ключ строки декларации: рейс + ГТД + № строки + ТН ВЭД. */
function costingDeclLineImportKey_(shipmentId, gtd, lineNo, tnved) {
  return [
    String(shipmentId || '').trim(),
    String(gtd || '').trim(),
    String(lineNo != null ? lineNo : '').trim(),
    costingNormalizeTnved_(tnved || '')
  ].join('\t');
}

function costingNormalizeDeclImportRow_(row, ncol) {
  const out = (row || []).slice();
  while (out.length < ncol) out.push('');
  if (out.length > ncol) out.length = ncol;
  return out;
}

function costingDeclLineColIndex_(headerRow, stdName, aliases) {
  const idx = costingFindColOptional_(headerRow, aliases);
  if (idx != null) return idx;
  const std = COST_CFG.DECL_LINES_HEADER.indexOf(stdName);
  return std >= 0 ? std : null;
}

function costingDeclLinesColMap_(headerRow) {
  return {
    upload: costingDeclLineColIndex_(headerRow, 'Загрузка_ID', ['Загрузка_ID', 'ID']),
    ship: costingDeclLineColIndex_(headerRow, 'SHIPMENT_ID', ['SHIPMENT_ID', 'ID_рейса']),
    gtd: costingDeclLineColIndex_(headerRow, 'Номер_ГТД', ['Номер_ГТД', 'Номер ГТД', 'ГТД']),
    lineNo: costingDeclLineColIndex_(headerRow, '№_строки', [
      '№_строки', '№ строки', '№ строки декларации', 'N строки', '№_товара', '№ товара'
    ]),
    tnved: costingDeclLineColIndex_(headerRow, 'Код_ТНВЭД', ['Код_ТНВЭД', 'Код ТН ВЭД', 'ТН ВЭД', 'ТНВЭД'])
  };
}

function costingDeclLineCell_(row, colIdx) {
  if (colIdx == null || colIdx < 0) return '';
  return row.length > colIdx ? row[colIdx] : '';
}

function costingDeclLineKeyFromDataRow_(row, cols) {
  return costingDeclLineImportKey_(
    costingDeclLineCell_(row, cols.ship),
    costingDeclLineCell_(row, cols.gtd),
    costingDeclLineCell_(row, cols.lineNo),
    costingDeclLineCell_(row, cols.tnved)
  );
}

/** Убрать повторяющиеся строки на листе (оставить первую). Возвращает число удалённых. */
function costingDedupeDeclLinesSheet_(linesSh) {
  const ncol = COST_CFG.DECL_LINES_HEADER.length;
  const lastRow = linesSh.getLastRow();
  if (lastRow < 2) return 0;
  const data = linesSh.getRange(1, 1, lastRow, Math.max(linesSh.getLastColumn(), ncol)).getValues();
  const cols = costingDeclLinesColMap_(data[0]);
  const seen = {};
  const out = [data[0]];
  let removed = 0;
  for (let r = 1; r < data.length; r++) {
    const key = costingDeclLineKeyFromDataRow_(data[r], cols);
    if (!key || key === '\t\t\t') {
      out.push(costingNormalizeDeclImportRow_(data[r], ncol));
      continue;
    }
    if (seen[key]) {
      removed++;
      continue;
    }
    seen[key] = true;
    out.push(costingNormalizeDeclImportRow_(data[r], ncol));
  }
  if (!removed) return 0;
  linesSh.getRange(2, 1, out.length - 1, ncol).setValues(out.slice(1));
  const extra = linesSh.getLastRow() - out.length;
  if (extra > 0) linesSh.deleteRows(out.length + 1, extra);
  return removed;
}

/**
 * Импорт: обновить существующую строку (рейс+ГТД+№+ТН ВЭД) или добавить новую; дубли на листе схлопнуть.
 */
function costingUpsertDeclLinesFromImport_(linesSh, rows, shipmentId, gtdNumber) {
  const ncol = COST_CFG.DECL_LINES_HEADER.length;
  const shipNorm = String(shipmentId || '').trim();
  const gtdNorm = String(gtdNumber || '').trim();
  if (!rows || !rows.length) return { inserted: 0, updated: 0, removedDupes: 0 };

  const ss = linesSh.getParent();
  if (ss) costingEnsureDeclLinesSheet_(ss);

  let data;
  if (linesSh.getLastRow() < 1) {
    data = [COST_CFG.DECL_LINES_HEADER.slice()];
  } else {
    const lastRow = linesSh.getLastRow();
    data = linesSh.getRange(1, 1, lastRow, Math.max(linesSh.getLastColumn(), ncol)).getValues();
  }
  const cols = costingDeclLinesColMap_(data[0]);
  const seen = {};
  const deduped = [data[0]];
  let removedDupes = 0;

  for (let r = 1; r < data.length; r++) {
    const key = costingDeclLineKeyFromDataRow_(data[r], cols);
    if (!key || key === '\t\t\t') {
      deduped.push(costingNormalizeDeclImportRow_(data[r], ncol));
      continue;
    }
    if (seen[key] != null) {
      removedDupes++;
      continue;
    }
    seen[key] = deduped.length;
    deduped.push(costingNormalizeDeclImportRow_(data[r], ncol));
  }
  data = deduped;

  let inserted = 0;
  let updated = 0;
  const idxLine = COST_CFG.DECL_LINES_HEADER.indexOf('№_строки');
  const idxTnved = COST_CFG.DECL_LINES_HEADER.indexOf('Код_ТНВЭД');

  for (let i = 0; i < rows.length; i++) {
    const row = costingNormalizeDeclImportRow_(rows[i], ncol);
    const key = costingDeclLineImportKey_(
      shipNorm,
      gtdNorm,
      idxLine >= 0 ? row[idxLine] : '',
      idxTnved >= 0 ? row[idxTnved] : ''
    );
    if (!key || key === '\t\t\t') {
      seen[key || '__empty__' + i] = data.length;
      data.push(row);
      inserted++;
      continue;
    }
    if (seen[key] != null) {
      data[seen[key]] = row;
      updated++;
    } else {
      seen[key] = data.length;
      data.push(row);
      inserted++;
    }
  }

  const body = [];
  for (let b = 1; b < data.length; b++) body.push(costingNormalizeDeclImportRow_(data[b], ncol));
  if (body.length) {
    linesSh.getRange(2, 1, body.length, ncol).setValues(body);
    const extra = linesSh.getLastRow() - (body.length + 1);
    if (extra > 0) linesSh.deleteRows(body.length + 2, extra);
    if (idxTnved >= 0) {
      linesSh.getRange(2, idxTnved + 1, body.length, 1).setNumberFormat('@');
    }
  }
  return { inserted: inserted, updated: updated, removedDupes: removedDupes };
}

/** Запись в журнал: тот же рейс + имя файла — обновить строку, не дублировать. */
function costingUpsertDeclJournalFromImport_(journalSh, rowValues) {
  const ncol = rowValues.length;
  const lastRow = journalSh.getLastRow();
  if (lastRow < 1) {
    journalSh.appendRow(rowValues);
    return { updated: false };
  }
  const data = journalSh.getDataRange().getValues();
  const h = data[0];
  const colShip = costingFindColOptional_(h, ['SHIPMENT_ID', 'ID_рейса']);
  const colFile = costingFindColOptional_(h, ['Имя_файла', 'Имя файла', 'Файл']);
  const stdShip = COST_CFG.DECL_JOURNAL_HEADER.indexOf('SHIPMENT_ID');
  const stdFile = COST_CFG.DECL_JOURNAL_HEADER.indexOf('Имя_файла');
  const ship = String(rowValues[colShip != null ? colShip : stdShip] || '').trim();
  const fileName = String(rowValues[colFile != null ? colFile : stdFile] || '').trim();

  for (let r = 1; r < data.length; r++) {
    if (
      String(data[r][colShip] || '').trim() === ship &&
      String(data[r][colFile] != null ? data[r][colFile] : '').trim() === fileName
    ) {
      const row = costingNormalizeDeclImportRow_(rowValues, Math.max(ncol, data[0].length));
      journalSh.getRange(r + 1, 1, 1, row.length).setValues([row]);
      return { updated: true };
    }
  }
  journalSh.appendRow(rowValues);
  return { updated: false };
}

/** Удалить из журнала загрузки, на которые больше нет строк декларации. */
function costingPruneOrphanDeclJournal_(ss) {
  const linesSh = costingFindSheetByRole_(ss, 'DECL_LINES');
  const journalSh = costingFindSheetByRole_(ss, 'DECL_JOURNAL');
  if (!linesSh || !journalSh || journalSh.getLastRow() < 2) return 0;

  const inUse = {};
  const lData = linesSh.getDataRange().getValues();
  const colUpload = costingFindColOptional_(lData[0], ['Загрузка_ID', 'ID']);
  if (colUpload == null) return 0;
  for (let r = 1; r < lData.length; r++) {
    const u = String(lData[r][colUpload] || '').trim();
    if (u) inUse[u] = true;
  }

  const jData = journalSh.getDataRange().getValues();
  const jUpload = costingFindColOptional_(jData[0], ['Загрузка_ID', 'ID']);
  if (jUpload == null) return 0;
  let removed = 0;
  for (let j = jData.length - 1; j >= 1; j--) {
    const u = String(jData[j][jUpload] || '').trim();
    if (u && !inUse[u]) {
      journalSh.deleteRow(j + 1);
      removed++;
    }
  }
  return removed;
}

/** Подсчёт дублей без удаления (для отчёта сопоставления). */
function costingCountDeclLineDuplicates_(linesData, declCols) {
  const cols = {
    ship: declCols.lShip,
    gtd: declCols.lGtd,
    lineNo: costingFindColOptional_(linesData[0], ['№_строки', '№ строки']),
    tnved: declCols.lTnved
  };
  const seen = {};
  let dupes = 0;
  for (let r = 1; r < linesData.length; r++) {
    const key = costingDeclLineKeyFromDataRow_(linesData[r], cols);
    if (!key || key === '\t\t\t') continue;
    if (seen[key]) dupes++;
    else seen[key] = true;
  }
  return dupes;
}

function costingCustomsXmlUploadDialogHtml_() {
  return (
    '<!doctype html><html><head><base target="_top"><meta charset="utf-8">' +
    '<style>body{font-family:Arial;padding:14px;font-size:13px}' +
    '.row{margin:10px 0}label{display:block;font-weight:600;margin-bottom:4px}' +
    'input{width:100%;padding:7px;box-sizing:border-box}' +
    '.note{font-size:11px;color:#666;margin-top:8px}' +
    'button{padding:8px 14px;font-size:13px}</style></head><body>' +
    '<div class="row"><label>SHIPMENT_ID (рейс)</label><input id="shipment" value="TR-2026-0008"></div>' +
    '<div class="row"><label>Файлы декларации (.xml)</label>' +
    '<input id="file" type="file" accept=".xml,text/xml,application/xml" multiple></div>' +
    '<div class="row"><button id="btn" onclick="upload()">Загрузить</button></div>' +
    '<p class="note">После загрузки выполните «Сопоставить строки декларации с SKU».</p>' +
    '<script>var sending=false;var CHUNK=' + COST_XML_CHUNK_SIZE_ + ';' +
    'function toB64(f){return new Promise(function(res,rej){var r=new FileReader();r.onload=function(){res(String(r.result).split(",")[1]);};r.onerror=rej;r.readAsDataURL(f);});}' +
    'function runPrep(p){return new Promise(function(ok,er){google.script.run.withSuccessHandler(ok).withFailureHandler(er).costingCustomsXmlPrepare(p);});}' +
    'function runChunk(t,fid,ci,tot,name,mime,part){return new Promise(function(ok,er){google.script.run.withSuccessHandler(ok).withFailureHandler(er).costingCustomsXmlUploadChunk(t,fid,ci,tot,name,mime,part);});}' +
    'function runFin(t,fid,p){return new Promise(function(ok,er){google.script.run.withSuccessHandler(ok).withFailureHandler(er).costingCustomsXmlFinalize(t,fid,p);});}' +
    'async function upload(){var sid=document.getElementById("shipment").value.trim();var fs=document.getElementById("file").files;' +
    'if(!sid){alert("Укажите SHIPMENT_ID");return;}if(!fs.length){alert("Выберите XML");return;}if(sending)return;sending=true;document.getElementById("btn").disabled=true;' +
    'try{var out=[];for(var fi=0;fi<fs.length;fi++){var prep=await runPrep({shipmentId:sid});var b64=await toB64(fs[fi]);var total=Math.max(1,Math.ceil(b64.length/CHUNK));' +
    'for(var ci=0;ci<total;ci++){await runChunk(prep.uploadToken,prep.folderId,ci,total,fs[fi].name,fs[fi].type||"application/xml",b64.substring(ci*CHUNK,(ci+1)*CHUNK));}' +
    'out.push(await runFin(prep.uploadToken,prep.folderId,{shipmentId:sid,fileName:fs[fi].name}));}' +
    'alert(out.join("\\n\\n---\\n\\n"));google.script.host.close();}catch(e){alert(e.message||e);}finally{sending=false;document.getElementById("btn").disabled=false;}}</script>' +
    '</body></html>'
  );
}

function costingCustomsXmlPrepare_(payload) {
  const shipmentId = String(payload && payload.shipmentId || '').trim();
  if (!shipmentId) throw new Error('SHIPMENT_ID не указан.');
  costingValidateShipmentForCustoms_(shipmentId);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const root = costingCustomsXmlGetDraftRootFolder_();
  const uploadToken = Utilities.getUuid();
  const folderName = uploadToken + '_' + shipmentId.replace(/[^\w\-]+/g, '_');
  const folder = root.createFolder(folderName);
  const meta = {
    uploadToken: uploadToken,
    folderId: folder.getId(),
    shipmentId: shipmentId,
    createdAt: new Date().toISOString()
  };
  folder.createFile(Utilities.newBlob(JSON.stringify(meta), 'application/json', '__meta.json'));
  return { uploadToken: uploadToken, folderId: folder.getId() };
}

function costingCustomsXmlUploadChunk_(uploadToken, folderId, chunkIndex, totalChunks, fileName, mimeType, base64Chunk) {
  if (!uploadToken || !folderId) throw new Error('Сессия загрузки не инициализирована.');
  const folder = DriveApp.getFolderById(folderId);
  const partName = '__chunk_' + chunkIndex + '_of_' + totalChunks + '.txt';
  folder.createFile(Utilities.newBlob(String(base64Chunk || ''), 'text/plain', partName));
  if (chunkIndex === 0) {
    folder.createFile(Utilities.newBlob(
      JSON.stringify({ fileName: fileName, mimeType: mimeType || 'application/xml', totalChunks: totalChunks }),
      'application/json',
      '__file_info.json'
    ));
  }
  return { ok: true, chunkIndex: chunkIndex };
}

function costingCustomsXmlFinalize_(uploadToken, folderId, payload) {
  const shipmentId = String(payload && payload.shipmentId || '').trim();
  const fileName = String(payload && payload.fileName || 'declaration.xml').trim();
  if (!shipmentId) throw new Error('SHIPMENT_ID не указан.');

  const folder = DriveApp.getFolderById(folderId);
  const xmlText = costingCustomsXmlAssembleFromFolder_(folder);
  const parsed = costingCustomsXmlParse_(xmlText);
  const uploadId = uploadToken || Utilities.getUuid();
  const now = new Date();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  costingEnsureDeclLinesSheet_(ss);
  costingEnsureDeclJournalSheet_(ss);

  const linesSh = costingGetSheetByRole_(ss, 'DECL_LINES');

  const rows = [];
  for (let i = 0; i < parsed.lines.length; i++) {
    const L = parsed.lines[i];
    rows.push([
      uploadId,
      now,
      shipmentId,
      parsed.gtdNumber,
      parsed.gtdDate,
      L.lineNo,
      L.tnved,
      L.description,
      L.qty,
      L.netWeight,
      L.currency,
      L.customsValue,
      L.duty,
      L.vat,
      L.fee,
      L.duty + L.vat + L.fee,
      '',
      'pending',
      'import:' + parsed.preset
    ]);
  }
  let upsertStats = { inserted: 0, updated: 0, removedDupes: 0 };
  if (rows.length) {
    upsertStats = costingUpsertDeclLinesFromImport_(linesSh, rows, shipmentId, parsed.gtdNumber);
    costingPruneOrphanDeclJournal_(ss);
  }

  const journalSh = costingFindSheetByRole_(ss, 'DECL_JOURNAL') || ss.insertSheet(COST_CFG.SHEETS.DECL_JOURNAL);
  costingEnsureDeclJournalSheet_(ss);
  const journalRow = [
    uploadId,
    now,
    shipmentId,
    fileName,
    parsed.preset,
    rows.length,
    parsed.totalDuty,
    parsed.totalVat,
    parsed.totalFee != null ? parsed.totalFee : 0,
    parsed.warnings.join('; ').slice(0, 500)
  ];
  const journalUpsert = costingUpsertDeclJournalFromImport_(journalSh, journalRow);
  costingPruneOrphanDeclJournal_(ss);

  try {
    folder.createFile(Utilities.newBlob(xmlText, 'application/xml', fileName));
  } catch (e) { /* ignore */ }

  const feeTotal = parsed.totalFee != null ? parsed.totalFee : 0;
  let upsertNote = '';
  if (upsertStats.updated > 0) {
    upsertNote += '\nОбновлено существующих строк: ' + upsertStats.updated + ' (повторная загрузка той же ГТД).';
  }
  if (upsertStats.inserted > 0 && upsertStats.updated === 0) {
    upsertNote += '\nНовых строк: ' + upsertStats.inserted + '.';
  } else if (upsertStats.inserted > 0) {
    upsertNote += '\nДобавлено новых строк: ' + upsertStats.inserted + '.';
  }
  if (upsertStats.removedDupes > 0) {
    upsertNote += '\nУдалено дублей на листе: ' + upsertStats.removedDupes + '.';
  }
  if (journalUpsert.updated) {
    upsertNote += '\nЗапись в журнале загрузок обновлена (тот же файл).';
  }
  return (
    'Импорт завершён.\n' +
    'Пресет: ' + parsed.preset + '\n' +
    'Строк товаров: ' + rows.length + '\n' +
    'ГТД: ' + (parsed.gtdNumber || '—') + '\n' +
    'Пошлина: ' + parsed.totalDuty.toFixed(2) + ' ₽\n' +
    'НДС: ' + parsed.totalVat.toFixed(2) + ' ₽\n' +
    'Сбор: ' + feeTotal.toFixed(2) + ' ₽' +
    upsertNote +
    (parsed.warnings.length ? ('\nЗамечания: ' + parsed.warnings.join('; ')) : '') +
    '\n\nДалее: меню → «Сопоставить строки декларации с SKU» (можно один раз на TR-2026-0008 после всех 4 файлов).'
  );
}

function costingCustomsXmlAssembleFromFolder_(folder) {
  const files = folder.getFiles();
  const partsByIdx = {};
  let fileInfo = null;
  while (files.hasNext()) {
    const f = files.next();
    const n = f.getName();
    if (n === '__file_info.json') {
      try { fileInfo = JSON.parse(f.getBlob().getDataAsString()); } catch (e) { fileInfo = null; }
      continue;
    }
    const m = n.match(/^__chunk_(\d+)_of_(\d+)\.txt$/);
    if (m) partsByIdx[Number(m[1])] = f.getBlob().getDataAsString();
  }
  const total = fileInfo && fileInfo.totalChunks ? fileInfo.totalChunks : Object.keys(partsByIdx).length;
  let joined = '';
  for (let i = 0; i < total; i++) {
    if (partsByIdx[i] == null) throw new Error('Не найден chunk ' + i + ' из ' + total);
    joined += partsByIdx[i];
  }
  return Utilities.newBlob(Utilities.base64Decode(joined)).getDataAsString('UTF-8');
}

function costingCustomsXmlGetDraftRootFolder_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parent = DriveApp.getFileById(ss.getId()).getParents().hasNext()
    ? DriveApp.getFileById(ss.getId()).getParents().next()
    : DriveApp.getRootFolder();
  const name = COST_XML_DRAFT_FOLDER_;
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

function costingValidateShipmentForCustoms_(shipmentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const batches = costingFindSheetByRole_(ss, 'BATCHES');
  if (!batches || batches.getLastRow() < 2) {
    throw new Error('Лист «Партии_в_рейсе» пуст — сначала наполните партии.');
  }
  const data = batches.getDataRange().getValues();
  const idx = costingFindCol_(data[0], ['SHIPMENT_ID', 'ID_рейса']);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx] || '').trim() === shipmentId) return;
  }
  throw new Error('В «Партии_в_рейсе» нет строк с SHIPMENT_ID = ' + shipmentId);
}

function costingCustomsXmlParse_(xmlText) {
  const text = String(xmlText || '').trim();
  if (!text) throw new Error('XML пуст.');
  const preset = costingCustomsXmlDetectPreset_(text);
  if (preset === 'fts_esadout_cu') {
    return costingCustomsXmlParseFtsEsadoutCu_(text);
  }
  let doc;
  try {
    doc = XmlService.parse(text);
  } catch (e) {
    throw new Error('Не удалось разобрать XML: ' + (e.message || String(e)));
  }
  return costingCustomsXmlParseGenericEec_(doc.getRootElement(), text);
}

/** ФТС / ED_Container: ESADout_CU + GTDoutCustomsMark (как ESAD0001–0004). */
function costingCustomsXmlDetectPreset_(text) {
  const raw = String(text || '').toLowerCase();
  if (raw.indexOf('esadout_cugoods') !== -1 || raw.indexOf('ed_container') !== -1) {
    return 'fts_esadout_cu';
  }
  if (raw.indexOf('urn:customs.ru:information:customsdocuments:esadout_cu') !== -1) {
    return 'fts_esadout_cu';
  }
  if (raw.indexOf('alta') !== -1) return 'alta';
  return 'generic_eec';
}

function costingCustomsXmlParseFtsEsadoutCu_(text) {
  const warnings = [];
  const gtd = costingCustomsXmlExtractGtdFts_(text);
  const blocks = costingCustomsXmlSplitGoodsBlocks_(text);
  if (!blocks.length) {
    warnings.push('Не найдены блоки ESADout_CUGoods — проверьте файл');
  }

  const docPayments = costingCustomsXmlParseFtsPayments_(text);
  const lines = [];
  let totalDuty = 0;
  let totalVat = 0;
  let totalFee = 0;
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    const lineNo = costingCustomsXmlRegexInBlock_(block, /<(?:[\w]+:)?GoodsNumeric>(\d+)</i) || String(i + 1);
    const tnved = costingNormalizeTnved_(
      costingCustomsXmlRegexInBlock_(block, /<(?:[\w]+:)?GoodsTNVEDCode>(\d+)</i)
    );
    const description = costingCustomsXmlJoinTagTexts_(block, 'GoodsDescription');
    const qty = costingToNumber_(
      costingCustomsXmlRegexInBlock_(block, /<(?:[\w]+:)?GoodsGroupQuantity>[\s\S]*?<(?:[\w]+:)?GoodsQuantity>([\d.]+)</i) ||
      costingCustomsXmlRegexInBlock_(block, /<(?:[\w]+:)?SupplementaryGoodsQuantity>[\s\S]*?<(?:[\w]+:)?GoodsQuantity>([\d.]+)</i) ||
      costingCustomsXmlRegexInBlock_(block, /<(?:[\w]+:)?GoodsQuantity>([\d.]+)</i)
    );
    const netWeight = costingToNumber_(
      costingCustomsXmlRegexInBlock_(block, /<(?:[\w]+:)?NetWeightQuantity>([\d.]+)</i)
    );
    const customsValue = costingToNumber_(
      costingCustomsXmlRegexInBlock_(block, /<(?:[\w]+:)?CustomsCost>([\d.]+)</i)
    );
    const payments = costingCustomsXmlParseFtsPayments_(block);
    totalDuty += payments.duty;
    totalVat += payments.vat;
    totalFee += payments.fee;
    if (!tnved) warnings.push('Строка ' + lineNo + ': нет кода ТН ВЭД');
    lines.push({
      lineNo: lineNo,
      tnved: tnved,
      description: description,
      qty: qty,
      netWeight: netWeight,
      currency: 'RUB',
      customsValue: customsValue,
      duty: payments.duty,
      vat: payments.vat,
      fee: payments.fee
    });
  }

  let linesFeeSum = 0;
  for (let li = 0; li < lines.length; li++) linesFeeSum += costingToNumber_(lines[li].fee);
  const docFee = costingToNumber_(docPayments.fee);
  if (docFee > linesFeeSum + 0.01) {
    costingCustomsXmlDistributeFeeToLines_(lines, docFee - linesFeeSum);
    totalFee = docFee;
    if (linesFeeSum <= 0) {
      warnings.push('Таможенный сбор (1010) взят с уровня документа и распределён по строкам товаров');
    }
  } else if (linesFeeSum > totalFee) {
    totalFee = linesFeeSum;
  }

  return {
    preset: 'fts_esadout_cu',
    gtdNumber: gtd.number,
    gtdDate: gtd.date,
    lines: lines,
    totalDuty: totalDuty,
    totalVat: totalVat,
    totalFee: totalFee,
    warnings: warnings
  };
}

/** Коды платежей ФТС: 1010 — сбор, 2010 — пошлина, 5010 — НДС. */
function costingCustomsXmlParseFtsPayments_(block) {
  const out = { duty: 0, vat: 0, fee: 0 };
  const re = /<(?:[\w]+:)?ESADout_CUCustomsPaymentCalculation>([\s\S]*?)<\/(?:[\w]+:)?ESADout_CUCustomsPaymentCalculation>/gi;
  let m;
  while ((m = re.exec(block)) !== null) {
    const chunk = m[1];
    const code = costingCustomsXmlRegexInBlock_(chunk, /<(?:[\w]+:)?PaymentModeCode>(\d+)</i);
    const amount = costingToNumber_(
      costingCustomsXmlRegexInBlock_(chunk, /<(?:[\w]+:)?PaymentAmount>([\d.]+)</i)
    );
    if (!code || !amount) continue;
    if (code === '1010') out.fee += amount;
    else if (code === '2010') out.duty += amount;
    else if (code === '5010') out.vat += amount;
  }
  return out;
}

/** Разнести target по строкам декларации (вес: пошлина+НДС). */
function costingSpreadFeesToDeclRowIdxs_(linesData, cols, rowIdxs, target) {
  if (cols.lFee == null || !rowIdxs.length || costingToNumber_(target) <= 0) return false;
  const pseudoLines = rowIdxs.map(function (ri) {
    const row = linesData[ri];
    return {
      duty: cols.lDuty != null ? costingToNumber_(row[cols.lDuty]) : 0,
      vat: cols.lVat != null ? costingToNumber_(row[cols.lVat]) : 0,
      customsValue: 0,
      qty: cols.lQty != null ? costingToNumber_(row[cols.lQty]) : 0,
      fee: 0
    };
  });
  costingCustomsXmlDistributeFeeToLines_(pseudoLines, target);
  for (let i = 0; i < rowIdxs.length; i++) {
    linesData[rowIdxs[i]][cols.lFee] = pseudoLines[i].fee;
  }
  return true;
}

/**
 * Сбор на уровне ГТД (один на весь XML) — разнести по всем товарным строкам этой ГТД.
 * Иначе сумма в сверке верная, а в колонке «Сбор» — только у 1-й позиции.
 */
function costingDistributeDeclFeesByGtdOnLines_(linesData, cols, shipmentFilter) {
  if (cols.lFee == null || cols.lGtd == null) return 0;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feeByUpload = costingLoadJournalFeeByUpload_(ss);
  const byGtd = {};

  for (let r = 1; r < linesData.length; r++) {
    const row = linesData[r];
    const sid = cols.lShip != null ? String(row[cols.lShip] || '').trim() : '';
    if (!sid) continue;
    if (shipmentFilter && sid !== shipmentFilter) continue;
    const gtd = String(row[cols.lGtd] || '').trim();
    if (!gtd) continue;
    const key = sid + '\t' + gtd;
    if (!byGtd[key]) byGtd[key] = { rowIdxs: [], lineSum: 0, uploads: {} };
    const g = byGtd[key];
    g.rowIdxs.push(r);
    g.lineSum += costingToNumber_(row[cols.lFee]);
    if (cols.lUpload != null) {
      const uid = String(row[cols.lUpload] || '').trim();
      if (uid) g.uploads[uid] = true;
    }
  }

  let fixed = 0;
  const keys = Object.keys(byGtd);
  for (let ki = 0; ki < keys.length; ki++) {
    const g = byGtd[keys[ki]];
    if (g.rowIdxs.length <= 1) continue;
    let journalSum = 0;
    const uids = Object.keys(g.uploads);
    for (let ui = 0; ui < uids.length; ui++) {
      const jf = feeByUpload[uids[ui]];
      if (jf > 0) journalSum += jf;
    }
    const target = Math.max(g.lineSum, journalSum);
    if (target <= 0.01) continue;
    if (costingSpreadFeesToDeclRowIdxs_(linesData, cols, g.rowIdxs, target)) {
      fixed += g.rowIdxs.length;
    }
  }
  return fixed;
}

/** Таможенный сбор часто один на весь XML, а не в каждом ESADout_CUGoods — дополняем строки. */
function costingCustomsXmlDistributeFeeToLines_(lines, feeToAllocate) {
  const add = costingToNumber_(feeToAllocate);
  if (!lines.length || add <= 0) return;
  let totalWeight = 0;
  const weights = [];
  for (let i = 0; i < lines.length; i++) {
    const L = lines[i];
    const w = costingToNumber_(L.duty) + costingToNumber_(L.vat);
    const fallback = costingToNumber_(L.customsValue) || costingToNumber_(L.qty) || 1;
    const weight = w > 0 ? w : fallback;
    weights.push(weight);
    totalWeight += weight;
  }
  if (!totalWeight) totalWeight = lines.length;
  for (let j = 0; j < lines.length; j++) {
    lines[j].fee = costingToNumber_(lines[j].fee) + add * (weights[j] / totalWeight);
  }
}

/** После сопоставления: сбор по ГТД → все SKU FACT этой ГТД (по доле пошлины). Цель — только сопоставленные группы (не все строки листа). */
function costingRedistributeFactFeeByGtd_(factAgg, declGroups, shipmentFilter) {
  const feeTargetByGtd = {};
  const groupKeys = Object.keys(declGroups);
  for (let gi = 0; gi < groupKeys.length; gi++) {
    const g = declGroups[groupKeys[gi]];
    if (shipmentFilter && g.sid !== shipmentFilter) continue;
    const gtd = g.meta && g.meta.gtd ? String(g.meta.gtd).trim() : '';
    const key = g.sid + '\t' + gtd;
    if (!feeTargetByGtd[key]) feeTargetByGtd[key] = 0;
    feeTargetByGtd[key] += costingToNumber_(g.fee);
  }

  const factKeysByGtd = {};
  const aggKeys = Object.keys(factAgg);
  for (let i = 0; i < aggKeys.length; i++) {
    const k = aggKeys[i];
    const f = factAgg[k];
    const gtdKey = f.shipmentId + '\t' + (f.gtd || '');
    if (!factKeysByGtd[gtdKey]) factKeysByGtd[gtdKey] = [];
    factKeysByGtd[gtdKey].push(k);
  }

  const gtdKeys = Object.keys(feeTargetByGtd);
  for (let gi = 0; gi < gtdKeys.length; gi++) {
    const gtdKey = gtdKeys[gi];
    const target = feeTargetByGtd[gtdKey];
    if (target <= 0) continue;
    const fKeys = factKeysByGtd[gtdKey];
    if (!fKeys || !fKeys.length) continue;

    let dutySum = 0;
    for (let fi = 0; fi < fKeys.length; fi++) {
      dutySum += factAgg[fKeys[fi]].duty;
    }

    if (dutySum > 0) {
      for (let fi = 0; fi < fKeys.length; fi++) {
        const f = factAgg[fKeys[fi]];
        f.fee = target * (f.duty / dutySum);
      }
    } else {
      const each = target / fKeys.length;
      for (let fi = 0; fi < fKeys.length; fi++) factAgg[fKeys[fi]].fee = each;
    }
  }
}

function costingCustomsXmlExtractGtdFts_(text) {
  const nd = text.match(/<!--\s*ND=([^>\s]+)\s*-->/i);
  if (nd && nd[1]) {
    return { number: String(nd[1]).trim(), date: '' };
  }
  const m = text.match(
    /<GTDID>[\s\S]*?<(?:[\w]+:)?CustomsCode>(\d+)<\/(?:[\w]+:)?CustomsCode>[\s\S]*?<(?:[\w]+:)?RegistrationDate>(\d{4}-\d{2}-\d{2})<\/(?:[\w]+:)?RegistrationDate>[\s\S]*?<(?:[\w]+:)?GTDNumber>(\d+)<\/(?:[\w]+:)?GTDNumber>/i
  );
  if (!m) return { number: '', date: '' };
  const iso = m[2];
  const ddmmyy = iso.substring(8, 10) + iso.substring(5, 7) + iso.substring(2, 4);
  return {
    number: m[1] + '/' + ddmmyy + '/' + m[3],
    date: iso
  };
}

function costingCustomsXmlSplitGoodsBlocks_(text) {
  const blocks = [];
  const re = /<ESADout_CUGoods>([\s\S]*?)<\/ESADout_CUGoods>/gi;
  let m;
  while ((m = re.exec(text)) !== null) {
    blocks.push(m[1]);
  }
  return blocks;
}

function costingCustomsXmlRegexInBlock_(block, re) {
  const m = String(block || '').match(re);
  return m && m[1] ? String(m[1]).trim() : '';
}

function costingCustomsXmlJoinTagTexts_(block, localTag) {
  const parts = [];
  const re = new RegExp('<(?:[\\w]+:)?' + localTag + '>([\\s\\S]*?)<\\/(?:[\\w]+:)?' + localTag + '>', 'gi');
  let m;
  while ((m = re.exec(block)) !== null) {
    const t = String(m[1] || '').replace(/\s+/g, ' ').trim();
    if (t) parts.push(t);
  }
  return parts.join(' ').slice(0, 500);
}

function costingCustomsXmlParseGenericEec_(root, text) {
  const warnings = [];
  const gtd = costingCustomsXmlExtractGtdFts_(text);
  const gtdNumber = gtd.number || costingCustomsXmlFirstText_(root, [
    'GTDNumber', 'RegistrationNumber', 'PRDocumentNumber', 'EDocumentNumber'
  ]);
  const gtdDate = gtd.date || costingCustomsXmlFirstText_(root, ['GTDDate', 'RegistrationDate', 'ExecutionDate']);

  const goods = costingCustomsXmlFindGoodsElementsXml_(root);
  if (!goods.length) {
    warnings.push('Товарные позиции не найдены — проверьте формат XML');
  }

  const lines = [];
  let totalDuty = 0;
  let totalVat = 0;
  for (let i = 0; i < goods.length; i++) {
    const g = goods[i];
    const lineNo = costingCustomsXmlFirstText_(g, ['GoodsNumeric', 'ConsignmentItemOrdinal', 'Number']) || String(i + 1);
    const tnved = costingNormalizeTnved_(
      costingCustomsXmlFirstText_(g, ['GoodsTNVEDCode', 'TNVEDCode', 'CommodityCode', 'HSCode'])
    );
    const description = costingCustomsXmlFirstText_(g, [
      'GoodsDescription', 'CommercialDescription', 'GoodsDescriptionText'
    ]);
    const qty = costingToNumber_(costingCustomsXmlFirstText_(g, ['GoodsQuantity', 'QuantityFact', 'Quantity']));
    const netWeight = costingToNumber_(costingCustomsXmlFirstText_(g, ['NetWeightQuantity', 'NetWeight']));
    const currency = costingCustomsXmlFirstText_(g, ['CurrencyCode', 'CurrencyA3Code']) || 'RUB';
    const customsValue = costingToNumber_(costingCustomsXmlFirstText_(g, ['CustomsCost']));
    const duty = costingCustomsXmlSumPayments_(g, ['duty', 'пошлин']);
    const vat = costingCustomsXmlSumPayments_(g, ['vat', 'ндс', 'nds']);
    const fee = costingCustomsXmlSumPayments_(g, ['fee', 'сбор']);
    totalDuty += duty;
    totalVat += vat;
    lines.push({
      lineNo: lineNo,
      tnved: tnved,
      description: description,
      qty: qty,
      netWeight: netWeight,
      currency: currency,
      customsValue: customsValue,
      duty: duty,
      vat: vat,
      fee: fee
    });
  }

  return {
    preset: 'generic_eec',
    gtdNumber: gtdNumber,
    gtdDate: gtdDate,
    lines: lines,
    totalDuty: totalDuty,
    totalVat: totalVat,
    warnings: warnings
  };
}

function costingCustomsXmlFindGoodsElementsXml_(root) {
  const out = [];
  if (!root) return out;
  const itemNames = {
    esadout_cugoods: true,
    goodsitem: true,
    consignmentitem: true
  };
  const descendants = root.getDescendants();
  for (let j = 0; j < descendants.length; j++) {
    const node = descendants[j];
    if (node.getType && node.getType() !== XmlService.ContentTypes.ELEMENT) continue;
    const el = node.asElement();
    const ln = String(el.getName() || '').toLowerCase();
    const loc = ln.indexOf(':') !== -1 ? ln.split(':').pop() : ln;
    if (itemNames[loc]) out.push(el);
  }
  return out;
}

function costingCustomsXmlFirstText_(el, localNames) {
  if (!el) return '';
  const wanted = {};
  for (let i = 0; i < localNames.length; i++) {
    wanted[String(localNames[i]).toLowerCase()] = true;
  }
  const stack = [el];
  while (stack.length) {
    const cur = stack.pop();
    const nm = String(cur.getName() || '').toLowerCase();
    const loc = nm.indexOf(':') !== -1 ? nm.split(':').pop() : nm;
    if (wanted[loc]) {
      const t = String(cur.getText() || '').trim();
      if (t) return t;
    }
    const kids = cur.getChildren();
    for (let k = 0; k < kids.length; k++) stack.push(kids[k]);
  }
  return '';
}

function costingCustomsXmlSumPayments_(goodsEl, keywords) {
  let sum = 0;
  const stack = [goodsEl];
  while (stack.length) {
    const cur = stack.pop();
    const nm = String(cur.getName() || '').toLowerCase();
    const loc = nm.indexOf(':') !== -1 ? nm.split(':').pop() : nm;
    if (loc.indexOf('payment') !== -1 || loc.indexOf('duty') !== -1 || loc.indexOf('tax') !== -1) {
      const kwHit = keywords.some(function (kw) {
        return nm.indexOf(kw) !== -1 || String(cur.getText() || '').toLowerCase().indexOf(kw) !== -1;
      });
      const amount = costingCustomsXmlChildAmount_(cur);
      if (amount && kwHit) sum += amount;
    }
    const kids = cur.getChildren();
    for (let k = 0; k < kids.length; k++) stack.push(kids[k]);
  }
  if (!sum) {
    sum = costingCustomsXmlSumByTag_(goodsEl, ['PaymentAmount', 'CustPaymentAmount', 'DutyTaxFee']);
  }
  return sum;
}

function costingCustomsXmlChildAmount_(el) {
  const stack = [el];
  while (stack.length) {
    const cur = stack.pop();
    const nm = String(cur.getName() || '').toLowerCase();
    if (nm.indexOf('amount') !== -1 || nm.indexOf('sum') !== -1) {
      const v = costingToNumber_(cur.getText());
      if (v) return v;
    }
    const kids = cur.getChildren();
    for (let k = 0; k < kids.length; k++) stack.push(kids[k]);
  }
  return 0;
}

function costingCustomsXmlSumByTag_(el, tags) {
  const wanted = {};
  for (let i = 0; i < tags.length; i++) wanted[String(tags[i]).toLowerCase()] = true;
  let sum = 0;
  const stack = [el];
  while (stack.length) {
    const cur = stack.pop();
    const nm = String(cur.getName() || '').toLowerCase();
    const loc = nm.indexOf(':') !== -1 ? nm.split(':').pop() : nm;
    if (wanted[loc]) {
      sum += costingToNumber_(cur.getText());
    }
    const kids = cur.getChildren();
    for (let k = 0; k < kids.length; k++) stack.push(kids[k]);
  }
  return sum;
}

function costingCustomsXmlRegexFirst_(text, re) {
  const m = text.match(re);
  return m && m[1] ? String(m[1]).trim() : '';
}


/** Декларация часто объединяет несколько SKU в одну строку по ТН ВЭД. */
function costingQtyDeclMatch_(declQty, batchQty) {
  if (declQty <= 0) return true;
  if (batchQty <= 0) return false;
  return Math.abs(batchQty - declQty) <= Math.max(0.01, declQty * 0.002);
}

/** Коды ТН ВЭД партии: справочник + колонка партии (оба участвуют в сопоставлении). */
function costingBatchTnvedCodes_(batch) {
  const out = [];
  if (batch.tnvedRef) out.push(batch.tnvedRef);
  if (batch.tnvedBatch && out.indexOf(batch.tnvedBatch) === -1) out.push(batch.tnvedBatch);
  if (batch.tnved && out.indexOf(batch.tnved) === -1) out.push(batch.tnved);
  return out;
}

function costingBatchMatchesDeclTnved_(batch, declTnved) {
  if (!declTnved) return false;
  const codes = costingBatchTnvedCodes_(batch);
  for (let i = 0; i < codes.length; i++) {
    if (costingTnvedCodesMatch_(codes[i], declTnved)) return true;
  }
  return false;
}

/** Кандидаты: точное совпадение ТН ВЭД; при отсутствии — префикс 8/6 цифр (субкоды в одной строке декларации). */
function costingFilterBatchesForDeclTnved_(pool, declTnved) {
  const decl = costingNormalizeTnved_(declTnved);
  if (!decl) return [];
  const bySku = {};
  function addBatch(b) {
    if (!b || !b.sku || bySku[b.sku]) return;
    bySku[b.sku] = b;
  }
  function addByCodePrefix(prefixLen) {
    if (prefixLen < 4) return;
    const head = decl.slice(0, prefixLen);
    pool.forEach(function (b) {
      const codes = costingBatchTnvedCodes_(b);
      for (let i = 0; i < codes.length; i++) {
        const c = codes[i];
        if (c.length >= prefixLen && c.slice(0, prefixLen) === head) {
          addBatch(b);
          break;
        }
      }
    });
  }

  pool.filter(function (b) { return costingBatchMatchesDeclTnved_(b, decl); }).forEach(addBatch);
  if (!Object.keys(bySku).length) {
    if (decl.length >= 8) addByCodePrefix(8);
    else if (decl.length >= 6) addByCodePrefix(6);
  } else if (decl.length >= 8) {
    addByCodePrefix(8);
  }
  return Object.keys(bySku).map(function (k) { return bySku[k]; });
}

function costingMatchDeclToBatchCandidates_(candidates, declQty, desc) {
  if (!candidates.length) return { status: 'unmatched', skus: [], shares: [], hint: '', paymentScale: 1 };
  if (candidates.length === 1) {
    const b = candidates[0];
    if (costingQtyDeclMatch_(declQty, b.qty)) {
      return { status: 'matched', skus: [b.sku], shares: [1], hint: '', paymentScale: 1 };
    }
    if (declQty > 0 && b.qty > 0 && b.qty < declQty) {
      return {
        status: 'matched_group_partial',
        skus: [b.sku],
        shares: [1],
        hint: 'частично по qty: в партии ' + b.qty + ' из ' + declQty + ' в декларации; платежи целиком на этот SKU',
        paymentScale: 1
      };
    }
    if (declQty > 0 && b.qty > declQty && b.qty / declQty <= 80) {
      return {
        status: 'matched',
        skus: [b.sku],
        shares: [1],
        hint: 'qty партии ' + b.qty + ' > qty декл. ' + declQty + ' (ед. учёта могут отличаться)',
        paymentScale: 1
      };
    }
    return { status: 'unmatched', skus: [], shares: [], hint: 'qty декл. ' + declQty + ' ≠ qty партии ' + b.qty, paymentScale: 1 };
  }
  const sumQty = candidates.reduce(function (s, b) { return s + (b.qty > 0 ? b.qty : 0); }, 0);
  if (costingQtyDeclMatch_(declQty, sumQty)) {
    return {
      status: 'matched_group',
      skus: candidates.map(function (b) { return b.sku; }),
      shares: candidates.map(function (b) { return sumQty > 0 ? b.qty / sumQty : 1 / candidates.length; }),
      hint: '',
      paymentScale: 1
    };
  }
  if (sumQty > 0 && declQty > 0 && sumQty < declQty && !costingQtyDeclMatch_(declQty, sumQty)) {
    const ratio = sumQty / declQty;
    if (ratio >= 0.05) {
      return {
        status: 'matched_group_partial',
        skus: candidates.map(function (b) { return b.sku; }),
        shares: candidates.map(function (b) { return b.qty / sumQty; }),
        hint: 'частично по qty: в партиях ' + sumQty + ' из ' + declQty +
          ' в декларации; пошлина/НДС распределены на найденные SKU по их qty',
        paymentScale: 1
      };
    }
  }
  if (sumQty > 0 && declQty > 0 && sumQty > declQty) {
    const ratioHigh = sumQty / declQty;
    if (ratioHigh <= 80) {
      return {
        status: 'matched_group',
        skus: candidates.map(function (b) { return b.sku; }),
        shares: candidates.map(function (b) { return b.qty / sumQty; }),
        hint: 'qty партий ' + sumQty + ' > qty декл. ' + declQty + ' (×' + ratioHigh.toFixed(1) +
          '); платежи по долям qty партий',
        paymentScale: 1
      };
    }
    return {
      status: 'unmatched',
      skus: [],
      shares: [],
      hint: candidates.length + ' SKU, qty декл. ' + declQty + ', сумма партий ' + sumQty +
        ' (разрыв ×' + ratioHigh.toFixed(0) + ' — проверьте единицы qty или лишние SKU в группе ТН ВЭД)',
      paymentScale: 1
    };
  }
  let best = candidates[0];
  let bestScore = Infinity;
  for (let c = 0; c < candidates.length; c++) {
    const cand = candidates[c];
    let score = declQty > 0 ? Math.abs(cand.qty - declQty) : 0;
    if (desc && cand.supplierArticle && desc.indexOf(cand.supplierArticle) !== -1) score -= 0.5;
    if (score < bestScore) { bestScore = score; best = cand; }
  }
  const tied = candidates.filter(function (cand) {
    let score = declQty > 0 ? Math.abs(cand.qty - declQty) : 0;
    if (desc && cand.supplierArticle && desc.indexOf(cand.supplierArticle) !== -1) score -= 0.5;
    return Math.abs(score - bestScore) < 0.001;
  });
  if (tied.length === 1 && costingQtyDeclMatch_(declQty, best.qty)) {
    return { status: 'matched', skus: [best.sku], shares: [1], hint: '', paymentScale: 1 };
  }
  if (tied.length > 1) return { status: 'ambiguous', skus: [], shares: [], hint: '', paymentScale: 1 };
  const skuList = candidates.map(function (b) { return b.sku; }).join(', ');
  return {
    status: 'unmatched',
    skus: [],
    shares: [],
    hint: candidates.length + ' SKU [' + skuList + '], qty декл. ' + declQty + ', сумма партий ' + sumQty,
    paymentScale: 1
  };
}

function costingFactAggAddSkuShares_(factAgg, sid, skus, shares, declTnved, duty, vat, fee, meta, paymentScale) {
  const scale = paymentScale != null && paymentScale > 0 && paymentScale <= 1 ? paymentScale : 1;
  for (let i = 0; i < skus.length; i++) {
    const sku = skus[i];
    const share = shares[i] != null ? shares[i] : 0;
    if (!sku || share <= 0) continue;
    const key = sid + '||' + sku;
    if (!factAgg[key]) {
      factAgg[key] = {
        shipmentId: sid,
        sku: sku,
        tnved: declTnved,
        duty: 0,
        vat: 0,
        fee: 0,
        gtd: meta.gtd || '',
        uploadId: meta.uploadId || ''
      };
    }
    factAgg[key].duty += duty * share * scale;
    factAgg[key].vat += vat * share * scale;
    factAgg[key].fee += fee * share * scale;
  }
}

/* ===================== Сопоставление с SKU ===================== */

function costingMatchDeclarationPrompt_() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    'Сопоставление декларации',
    'Введите SHIPMENT_ID (пусто = все рейсы с pending):',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const shipmentId = String(res.getResponseText() || '').trim();
  const report = costingMatchDeclarationLinesToBatches_(shipmentId || null);
  ui.alert('Сопоставление', report.summary, ui.ButtonSet.OK);
}

function costingIsDeclStatusMatched_(st) {
  const s = String(st || '').trim().toLowerCase();
  return s === 'matched' || s === 'matched_by_qty' || s === 'matched_group' || s === 'matched_group_partial';
}

/** Собрать группы строк декларации (агрегат по SHIPMENT_ID + ТН ВЭД). */
function costingCollectDeclGroups_(linesData, linesDisplay, cols, shipmentFilter, onlyPending) {
  const declGroups = {};
  let withoutTnved = 0;
  for (let r = 1; r < linesData.length; r++) {
    const row = linesData[r];
    const sid = String(row[cols.lShip] || '').trim();
    if (!sid) continue;
    if (shipmentFilter && sid !== shipmentFilter) continue;
    if (onlyPending && cols.lStatus != null && costingIsDeclStatusMatched_(row[cols.lStatus])) continue;

    const declTnved = cols.lTnved != null
      ? costingTnvedFromCells_(row[cols.lTnved], linesDisplay[r][cols.lTnved])
      : '';
    if (!declTnved) {
      withoutTnved++;
      continue;
    }
    const gKey = sid + '\t' + declTnved;
    if (!declGroups[gKey]) {
      declGroups[gKey] = {
        sid: sid,
        declTnved: declTnved,
        qty: 0,
        duty: 0,
        vat: 0,
        fee: 0,
        desc: '',
        meta: { gtd: '', uploadId: '' },
        rowIdxs: []
      };
    }
    const g = declGroups[gKey];
    g.qty += cols.lQty != null ? costingToNumber_(row[cols.lQty]) : 0;
    g.duty += cols.lDuty != null ? costingToNumber_(row[cols.lDuty]) : 0;
    g.vat += cols.lVat != null ? costingToNumber_(row[cols.lVat]) : 0;
    g.fee += cols.lFee != null ? costingToNumber_(row[cols.lFee]) : 0;
    const descPart = cols.lDesc != null ? String(row[cols.lDesc] || '').trim().toLowerCase() : '';
    if (descPart.length > g.desc.length) g.desc = descPart;
    if (cols.lGtd != null && !g.meta.gtd) g.meta.gtd = String(row[cols.lGtd] || '').trim();
    if (cols.lUpload != null && !g.meta.uploadId) g.meta.uploadId = String(row[cols.lUpload] || '').trim();
    g.rowIdxs.push(r);
  }
  return { groups: declGroups, withoutTnved: withoutTnved };
}

/** Сбор по upload из журнала (колонка Сбор_итого или legacy-число в Предупреждения). */
function costingLoadJournalFeeByUpload_(ss) {
  const feeByUpload = {};
  const journalSh = costingFindSheetByRole_(ss, 'DECL_JOURNAL');
  if (!journalSh || journalSh.getLastRow() < 2) return feeByUpload;
  const jData = journalSh.getDataRange().getValues();
  const jh = jData[0];
  const jUpload = costingFindColOptional_(jh, ['Загрузка_ID', 'ID']);
  if (jUpload == null) return feeByUpload;
  const jFee = costingFindColOptional_(jh, ['Сбор_итого', 'Сбор итого', 'Сбор_итог', 'Сбор итог']);
  const jWarn = costingFindColOptional_(jh, ['Предупреждения', 'Ошибки', 'Замечания']);
  for (let j = 1; j < jData.length; j++) {
    const uid = String(jData[j][jUpload] || '').trim();
    if (!uid) continue;
    let fee = jFee != null ? costingToNumber_(jData[j][jFee]) : 0;
    if (fee <= 0 && jWarn != null) {
      const v = jData[j][jWarn];
      const n = costingToNumber_(v);
      const s = String(v || '');
      if (n > 50 && s.indexOf(';') < 0 && s.indexOf('Тамож') < 0) fee = n;
    }
    if (fee > 0) feeByUpload[uid] = fee;
  }
  return feeByUpload;
}

/** Сверить сбор на строках декларации с журналом загрузки; при расхождении — перераспределить по всем строкам upload. */
function costingBackfillDeclFeesFromJournal_(linesData, cols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (cols.lUpload == null || cols.lFee == null) return 0;
  const feeByUpload = costingLoadJournalFeeByUpload_(ss);
  if (!Object.keys(feeByUpload).length) return 0;

  const rowsByUpload = {};
  for (let r = 1; r < linesData.length; r++) {
    const uid = String(linesData[r][cols.lUpload] || '').trim();
    if (!uid || !feeByUpload[uid]) continue;
    if (!rowsByUpload[uid]) rowsByUpload[uid] = [];
    rowsByUpload[uid].push(r);
  }

  let fixed = 0;
  const uploadIds = Object.keys(rowsByUpload);
  for (let u = 0; u < uploadIds.length; u++) {
    const uid = uploadIds[u];
    const rowIdxs = rowsByUpload[uid];
    const target = feeByUpload[uid];
    let existing = 0;
    for (let i = 0; i < rowIdxs.length; i++) {
      existing += costingToNumber_(linesData[rowIdxs[i]][cols.lFee]);
    }
    if (Math.abs(existing - target) < 0.02 && rowIdxs.length <= 1) continue;

    if (costingSpreadFeesToDeclRowIdxs_(linesData, cols, rowIdxs, target)) {
      fixed += rowIdxs.length;
    }
  }
  return fixed;
}

function costingMatchDeclarationLinesToBatches_(shipmentFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  costingEnsureDeclLinesSheet_(ss);
  costingEnsureCustomsFactColumns_(ss);
  costingEnsureDeclJournalSheet_(ss);

  const linesSh = costingGetSheetByRole_(ss, 'DECL_LINES');
  const dupesRemovedOnSheet = costingDedupeDeclLinesSheet_(linesSh);
  const linesData = linesSh.getDataRange().getValues();
  const linesDisplay = linesSh.getDataRange().getDisplayValues();
  if (linesData.length < 2) {
    return { summary: 'Нет строк на листе «Декларации_строки».', matched: 0, ambiguous: 0, unmatched: 0 };
  }
  const lh = linesData[0];
  const declCols = {
    lShip: costingFindCol_(lh, ['SHIPMENT_ID']),
    lSku: costingFindColOptional_(lh, ['Артикул_ВБ', 'Артикул ВБ']),
    lStatus: costingFindColOptional_(lh, ['Статус_сопоставления']),
    lTnved: costingFindColOptional_(lh, ['Код_ТНВЭД', 'Код ТН ВЭД', 'ТН ВЭД', 'ТНВЭД', 'TNVED']),
    lDesc: costingFindColOptional_(lh, ['Описание']),
    lQty: costingFindColOptional_(lh, ['Qty', 'Количество']),
    lDuty: costingFindColOptional_(lh, ['Пошлина']),
    lVat: costingFindColOptional_(lh, ['НДС']),
    lFee: costingFindColOptional_(lh, ['Сбор']),
    lGtd: costingFindColOptional_(lh, ['Номер_ГТД']),
    lUpload: costingFindColOptional_(lh, ['Загрузка_ID']),
    lComment: costingFindColOptional_(lh, ['Комментарий'])
  };
  const lShip = declCols.lShip;
  const lSku = declCols.lSku;
  const lStatus = declCols.lStatus;
  const lTnved = declCols.lTnved;
  const lDesc = declCols.lDesc;
  const lQty = declCols.lQty;
  const lDuty = declCols.lDuty;
  const lVat = declCols.lVat;
  const lFee = declCols.lFee;
  const lGtd = declCols.lGtd;
  const lUpload = declCols.lUpload;
  const lComment = declCols.lComment;

  const dupesRemaining = costingCountDeclLineDuplicates_(linesData, declCols);
  const feeSpreadByGtd = costingDistributeDeclFeesByGtdOnLines_(linesData, declCols, shipmentFilter);
  const feeBackfilled = costingBackfillDeclFeesFromJournal_(linesData, declCols);

  const batchesSh = costingGetSheetByRole_(ss, 'BATCHES');
  const batchesData = batchesSh.getDataRange().getValues();
  const bh = batchesData[0];
  const bShip = costingFindCol_(bh, ['SHIPMENT_ID', 'ID_рейса']);
  const bSku = costingFindCol_(bh, ['Артикул ВБ', 'Артикул_ВБ']);
  const bTnved = costingFindColOptional_(bh, ['Код ТН ВЭД', 'Код_ТНВЭД', 'ТН ВЭД', 'ТНВЭД', 'TNVED', 'HS CODE']);
  const bQty = costingFindCol_(bh, ['Количество', 'Qty']);
  const bSupArt = costingFindColOptional_(bh, ['Артикул поставщика', 'Артикул_поставщика']);
  const skuToTnved = costingLoadSkuToTnvedMap_(ss);
  const batchesDisplay = batchesSh.getDataRange().getDisplayValues();

  const batchesByShipment = {};
  let batchRowsInFilter = 0;
  let batchRowsWithTnved = 0;
  for (let i = 1; i < batchesData.length; i++) {
    const sid = String(batchesData[i][bShip] || '').trim();
    if (!sid) continue;
    if (shipmentFilter && sid !== shipmentFilter) continue;
    batchRowsInFilter++;
    const sku = String(batchesData[i][bSku] || '').trim();
    const fromRef = sku && skuToTnved[sku] ? costingNormalizeTnved_(skuToTnved[sku]) : '';
    const fromBatch = bTnved != null
      ? costingTnvedFromCells_(batchesData[i][bTnved], batchesDisplay[i] ? batchesDisplay[i][bTnved] : '')
      : '';
    const tnved = fromRef || fromBatch;
    if (tnved) batchRowsWithTnved++;
    if (!batchesByShipment[sid]) batchesByShipment[sid] = [];
    batchesByShipment[sid].push({
      sku: sku,
      tnved: tnved,
      tnvedRef: fromRef,
      tnvedBatch: fromBatch,
      qty: costingToNumber_(batchesData[i][bQty]),
      supplierArticle: bSupArt != null ? String(batchesData[i][bSupArt] || '').trim().toLowerCase() : ''
    });
  }

  let matched = 0;
  let matchedGroups = 0;
  let matchedPartial = 0;
  let ambiguous = 0;
  let unmatched = 0;
  let declWithoutTnved = 0;
  const factAgg = {};
  const diagHints = [];
  let declDutyTotal = 0;
  let declVatTotal = 0;
  let declFeeTotal = 0;

  let collected = costingCollectDeclGroups_(linesData, linesDisplay, declCols, shipmentFilter, true);
  let declGroups = collected.groups;
  declWithoutTnved = collected.withoutTnved;
  let rebuiltFactFromAllDecl = false;

  if (!Object.keys(declGroups).length && shipmentFilter) {
    collected = costingCollectDeclGroups_(linesData, linesDisplay, declCols, shipmentFilter, false);
    declGroups = collected.groups;
    declWithoutTnved += collected.withoutTnved;
    rebuiltFactFromAllDecl = Object.keys(declGroups).length > 0;
  }

  const groupKeys = Object.keys(declGroups);
  for (let gi = 0; gi < groupKeys.length; gi++) {
    const g = declGroups[groupKeys[gi]];
    const pool = batchesByShipment[g.sid] || [];
    let candidates = costingFilterBatchesForDeclTnved_(pool, g.declTnved);

    let matchResult = costingMatchDeclToBatchCandidates_(candidates, g.qty, g.desc);
    if (matchResult.status === 'unmatched' && !candidates.length && g.qty > 0) {
      const byQty = pool.filter(function (b) {
        return b.qty > 0 && Math.abs(b.qty - g.qty) < 0.01;
      });
      if (byQty.length === 1) {
        matchResult = { status: 'matched_by_qty', skus: [byQty[0].sku], shares: [1], hint: '', paymentScale: 1 };
      }
    }

    const status = matchResult.status;
    const skuStr = matchResult.skus ? matchResult.skus.join('; ') : '';
    let comment = '';
    if (status === 'matched_group') {
      comment = matchResult.hint || ('группа ' + matchResult.skus.length + ' SKU по ТН ВЭД (qty=' + g.qty + ')');
    } else if (status === 'matched_group_partial') {
      comment = matchResult.hint || ('частичная группа ' + matchResult.skus.length + ' SKU');
    } else if (g.rowIdxs.length > 1) {
      comment = 'агрегат ' + g.rowIdxs.length + ' строк декларации';
    }

    for (let ri = 0; ri < g.rowIdxs.length; ri++) {
      const r = g.rowIdxs[ri];
      const row = linesData[r];
      if (lSku != null) row[lSku] = skuStr;
      if (lStatus != null) row[lStatus] = status;
      if (lComment != null && comment) row[lComment] = comment;
      linesData[r] = row;
    }

    if (status === 'matched' || status === 'matched_by_qty' || status === 'matched_group' || status === 'matched_group_partial') {
      matched++;
      if (status === 'matched_group') matchedGroups++;
      if (status === 'matched_group_partial') matchedPartial++;
      declDutyTotal += g.duty;
      declVatTotal += g.vat;
      declFeeTotal += g.fee;
      costingFactAggAddSkuShares_(
        factAgg, g.sid, matchResult.skus, matchResult.shares, g.declTnved, g.duty, g.vat, g.fee, g.meta,
        matchResult.paymentScale
      );
    } else if (status === 'ambiguous') {
      ambiguous++;
    } else {
      unmatched++;
      if (diagHints.length < 3) {
        const hint = matchResult.hint || ('ТН ВЭД ' + g.declTnved + ', qty ' + g.qty + ' — нет партии в рейсе ' + g.sid);
        diagHints.push(hint);
      }
    }
  }

  if (Object.keys(factAgg).length) {
    costingRedistributeFactFeeByGtd_(factAgg, declGroups, shipmentFilter);
  }

  if (linesData.length > 1) {
    const ncol = linesData[0].length;
    linesSh.getRange(1, 1, linesData.length, ncol).setValues(linesData);
  }

  const customsSh = costingGetSheetByRole_(ss, 'CUSTOMS');
  if (Object.keys(factAgg).length) {
    costingUpsertCustomsFactRows_(customsSh, factAgg, shipmentFilter);
  }

  let factDutyTotal = 0;
  let factVatTotal = 0;
  let factFeeTotal = 0;
  const factKeys = Object.keys(factAgg);
  for (let fi = 0; fi < factKeys.length; fi++) {
    const f = factAgg[factKeys[fi]];
    factDutyTotal += f.duty;
    factVatTotal += f.vat;
    factFeeTotal += f.fee;
  }

  let summary =
    'Сопоставлено: ' + matched +
    (matchedGroups ? (' (полных групп: ' + matchedGroups + ')') : '') +
    (matchedPartial ? (' (частичных: ' + matchedPartial + ')') : '') + '\n' +
    'Неоднозначно: ' + ambiguous + '\n' +
    'Без пары: ' + unmatched + '\n' +
    'Записей FACT на «Таможенные платежи»: ' + Object.keys(factAgg).length;
  if (shipmentFilter) {
    summary += '\n\nРейс ' + shipmentFilter + ': партий ' + batchRowsInFilter +
      ', с ТН ВЭД (партия или справочник): ' + batchRowsWithTnved;
  } else {
    summary += '\n\nПартий в фильтре: ' + batchRowsInFilter +
      ', с ТН ВЭД: ' + batchRowsWithTnved;
  }
  if (declWithoutTnved) {
    summary += '\nСтрок декларации без ТН ВЭД: ' + declWithoutTnved;
  }
  if (unmatched && batchRowsInFilter && !batchRowsWithTnved) {
    summary += '\n\nПодсказка: в «Партии_в_рейсе» нет ТН ВЭД — заполните колонку или синхронизируйте справочник товаров (меню себестоимости).';
  } else if (unmatched && batchRowsWithTnved) {
    summary += '\n\nПодсказка: проверьте, что SHIPMENT_ID в декларации совпадает с рейсом в партиях и коды ТН ВЭД совпадают со справочником.';
  }
  if (matched && (declDutyTotal > 0 || declVatTotal > 0)) {
    summary += '\n\nСверка сумм (сопоставленные строки декларации → FACT):';
    summary += '\n• Пошлина: декл. ' + declDutyTotal.toFixed(2) + ' → FACT ' + factDutyTotal.toFixed(2) +
      (Math.abs(declDutyTotal - factDutyTotal) > 1 ? ' ⚠' : ' ✓');
    summary += '\n• НДС: декл. ' + declVatTotal.toFixed(2) + ' → FACT ' + factVatTotal.toFixed(2) +
      (Math.abs(declVatTotal - factVatTotal) > 1 ? ' ⚠' : ' ✓');
    if (declFeeTotal > 0 || factFeeTotal > 0) {
      summary += '\n• Сбор: декл. ' + declFeeTotal.toFixed(2) + ' → FACT ' + factFeeTotal.toFixed(2) +
        (Math.abs(declFeeTotal - factFeeTotal) > 1 ? ' ⚠' : ' ✓');
    }
  }
  if (dupesRemovedOnSheet > 0) {
    summary += '\n\nℹ️ Удалены дубликаты строк декларации на листе: ' + dupesRemovedOnSheet + '.';
  }
  if (dupesRemaining > 0) {
    summary += '\n\n⚠️ Остались дубликаты строк (' + dupesRemaining + ') — проверьте № строки и ТН ВЭД.';
  }
  if (feeSpreadByGtd > 0) {
    summary += '\n\nℹ️ Сбор разнесён по товарным строкам внутри ГТД (' + feeSpreadByGtd + ' строк).';
  }
  if (feeBackfilled > 0) {
    summary += '\n\nℹ️ Сбор по строкам декларации восстановлен из журнала загрузки (' + feeBackfilled + ' строк).';
  }
  if (rebuiltFactFromAllDecl) {
    summary += '\n\nℹ️ Новых pending-строк не было — FACT пересобран по всем строкам декларации рейса.';
  } else if (!groupKeys.length && !Object.keys(factAgg).length) {
    summary += '\n\nℹ️ Нет строк для обработки (все уже сопоставлены или нет pending). Лист «Таможенные платежи» не изменён.';
  }
  if (diagHints.length) {
    summary += '\n\nПримеры без пары:\n• ' + diagHints.join('\n• ');
  }
  return { summary: summary, matched: matched, ambiguous: ambiguous, unmatched: unmatched };
}

function costingUpsertCustomsFactRows_(customsSh, factAgg, shipmentFilter) {
  const aggKeys = Object.keys(factAgg);
  if (!aggKeys.length) return;
  let data = customsSh.getDataRange().getValues();
  let header = data.length ? data[0].slice() : ['SHIPMENT_ID'];
  let hmap = costingHeaderMap_(header);
  const idxShip = costingEnsureHeaderColumn_(customsSh, hmap, 'SHIPMENT_ID');
  const idxSku = costingEnsureHeaderColumn_(customsSh, hmap, 'Артикул_ВБ');
  const idxTnved = costingEnsureHeaderColumn_(customsSh, hmap, 'Код_ТНВЭД');
  const idxScenario = costingEnsureHeaderColumn_(customsSh, hmap, 'Сценарий');
  const idxSource = costingEnsureHeaderColumn_(customsSh, hmap, 'Источник');
  const idxGtd = costingEnsureHeaderColumn_(customsSh, hmap, 'Номер_ГТД');
  const idxUpload = costingEnsureHeaderColumn_(customsSh, hmap, 'Загрузка_ID');
  const idxDuty = costingEnsureHeaderColumn_(customsSh, hmap, 'Пошлина_RUB');
  const idxVat = costingEnsureHeaderColumn_(customsSh, hmap, 'НДС_RUB');
  const idxFee = costingEnsureHeaderColumn_(customsSh, hmap, 'Таможенный_сбор_RUB');

  data = customsSh.getDataRange().getValues();
  header = data[0].slice();
  const colCount = header.length;

  const replaceKeys = {};
  for (let i = 0; i < aggKeys.length; i++) replaceKeys[aggKeys[i]] = true;
  const filterShip = shipmentFilter ? String(shipmentFilter).trim() : '';
  const refreshShipment = filterShip && aggKeys.length > 0;

  const kept = [header];
  for (let r = 1; r < data.length; r++) {
    const row = data[r].slice();
    while (row.length < colCount) row.push('');
    const scen = String(row[idxScenario] || '').trim().toUpperCase();
    const sid = String(row[idxShip] || '').trim();
    const sku = String(row[idxSku] || '').trim();
    if (scen === 'FACT' && sku) {
      const key = sid + '||' + sku;
      if (replaceKeys[key]) continue;
      if (refreshShipment && sid === filterShip) continue;
    }
    kept.push(row);
  }

  const keys = aggKeys;
  for (let i = 0; i < keys.length; i++) {
    const f = factAgg[keys[i]];
    const newRow = new Array(colCount).fill('');
    newRow[idxShip] = f.shipmentId;
    newRow[idxSku] = f.sku;
    newRow[idxTnved] = f.tnved;
    newRow[idxScenario] = 'FACT';
    newRow[idxSource] = 'XML';
    newRow[idxGtd] = f.gtd;
    newRow[idxUpload] = f.uploadId;
    newRow[idxDuty] = f.duty;
    newRow[idxVat] = f.vat;
    newRow[idxFee] = f.fee;
    kept.push(newRow);
  }

  customsSh.clearContents();
  customsSh.getRange(1, 1, kept.length, colCount).setValues(kept);
  customsSh.getRange(1, 1, 1, colCount).setFontWeight('bold');
}

/** Отдельное меню (книга 05): видно даже если в GAS устарел costing.gs */
function addCostingCustomsMenu_(ui) {
  ui.createMenu('📄 Таможня (XML + факт)')
    .addItem('Загрузить декларацию (XML)', 'costingCustomsXmlOpenUploadDialog_')
    .addItem('Сопоставить строки декларации с SKU', 'costingMatchDeclarationPrompt_')
    .addSeparator()
    .addItem('Dry-run факт (по рейсу)', 'costingDryRunFactPrompt_')
    .addItem('Пересчитать фактическую себестоимость', 'rebuildCostingFact_')
    .addItem('📊 Обновить лист «Сверка план vs факт»', 'costingComparePlanFactPrompt_')
    .addToUi();
}