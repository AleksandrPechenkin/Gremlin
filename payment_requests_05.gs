/**
 * Книга 05 — заявки на оплату услуг логистики.
 *
 * Источник заявки — строки листа «Затраты рейса». Менеджер логистики выделяет в этом
 * листе одну или несколько строк (статьи расходов одного рейса/нескольких рейсов
 * у одного контрагента и в одной валюте), вызывает «Подать заявку», прикладывает
 * подписанный документ. Очередь живёт на отдельном листе книги 05
 * (`Payment_Link_Map_Logistics`), согласующий — тот же, что в книге 01
 * (Script Property `PAYMENT_APPROVER_EMAIL`). Одобренная заявка отправляется в общий
 * Реестр платежей (та же таблица и лист, что используются для оплат менеджеров).
 *
 * Связующее звено товара ↔ машины ↔ расхода ↔ оплаты — `SHIPMENT_ID`. В строке
 * Реестра он попадает и в «Назначение», и в комментарий (колонка M).
 *
 * Возврат факта оплаты — `payLogSyncPaidStatuses`: статус «Оплачено» + дата факта
 * подтягиваются по `№ заявки` и проставляются на тех же строках «Затраты рейса»,
 * по которым подавалась заявка.
 *
 * Состав обязательных скрипт-свойств (общих с книгой 01):
 *   PAYMENT_REGISTRY_SPREADSHEET_ID, PAYMENT_REGISTRY_SHEET_NAME,
 *   PAYMENT_REGISTRY_ROOT_FOLDER_ID, PAYMENT_APPROVER_EMAIL.
 * Дополнительные (опциональные, со значениями по умолчанию):
 *   PAYMENT_DRAFT_FOLDER_NAME_LOG (имя подпапки черновиков; по умолчанию
 *   `_DRAFT_PAYMENT_REQUESTS_LOG`), PAYMENT_INITIATOR_NAME (по умолчанию
 *   «Печёнкин А.А.»), PAYMENT_LOG_DEPARTMENT (по умолчанию «Снабжение»).
 */

const PAY_LOG_CFG = {
  EXPENSES_SHEET: 'Затраты рейса',
  EXPENSES_ALIASES: ['Затраты рейса', 'Затраты_рейса'],
  EXPENSES_HEADER_ROW: 1,
  EXPENSES_DATA_START_ROW: 2,
  QUEUE_SHEET: 'Payment_Link_Map_Logistics',
  REGISTRY_ID_DEFAULT: '1BUFF4-tlLg-H8faxgKUcnlcaL8aJS84x9bZGza9_K_M',
  REGISTRY_SHEET_DEFAULT: 'Реестр',
  ROOT_FOLDER_DEFAULT: '1wMYUKNsNixmNcI1HdsmScn4V-G5Vy6wL',
  DRAFT_FOLDER_NAME_DEFAULT: '_DRAFT_PAYMENT_REQUESTS_LOG',
  APPROVER_EMAIL_DEFAULT: 'banych83@gmail.com',
  INITIATOR_DEFAULT: 'Печёнкин А.А.',
  DEPARTMENT_DEFAULT: 'Снабжение'
};

const PAY_LOG_COL_ALIASES = {
  shipment: ['SHIPMENT_ID', 'ID_рейса', 'ID рейса', 'Рейс'],
  article: ['Статья_затрат', 'Статья затрат', 'Статья'],
  counterparty: ['Контрагент'],
  currency: ['Валюта'],
  amount: ['Сумма_в_валюте', 'Сумма в валюте', 'Сумма'],
  amountRub: ['Сумма_RUB', 'Сумма RUB'],
  payRequestNo: ['№ заявки', '№ Заявки', 'Номер заявки', 'Номер_заявки'],
  payStatus: ['Статус оплаты', 'Статус_оплаты'],
  payDate: ['Дата оплаты', 'Дата_оплаты'],
  payFolder: ['Ссылка на папку', 'Ссылка_на_папку', 'Папка оплат']
};

const PAY_LOG_NEW_COLS = ['№ заявки', 'Статус оплаты', 'Дата оплаты', 'Ссылка на папку'];

/* ===================== Меню ===================== */

function addPaymentLogisticsMenu_(ui) {
  ui.createMenu('💳 Оплаты логистики')
    .addItem('🔐 Авторизация для логиста', 'payLogAuthorizeUser')
    .addSeparator()
    .addItem('1) Подать заявку (по выделенным строкам «Затраты рейса»)', 'payLogOpenRequestDialog')
    .addItem('2) Проверка/отправка в реестр', 'payLogOpenApprovalDialog')
    .addSeparator()
    .addItem('3) Синхронизировать оплаченные', 'payLogSyncPaidStatuses')
    .addItem('🛠 Добавить недостающие колонки оплат в «Затраты рейса»', 'payLogEnsurePayColsMenu')
    .addToUi();
}

function payLogAuthorizeUser() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const cell = sheet.getRange(1, 1);
    cell.setValue(cell.getValue());
    const rootId = payLogProp_('PAYMENT_REGISTRY_ROOT_FOLDER_ID', PAY_LOG_CFG.ROOT_FOLDER_DEFAULT);
    DriveApp.getFolderById(rootId).getName();
    DriveApp.getRootFolder().getName();
    const test = DriveApp.getRootFolder().createFolder('AUTH_TEST_05_' + new Date().getTime());
    test.setTrashed(true);
    SpreadsheetApp.getUi().alert(
      '✅ Авторизация выполнена',
      'Разрешения (таблица + Drive: чтение/запись/создание папок) выданы. Можно подавать заявки через «💳 Оплаты логистики».',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      '⚠️ Требуется подтверждение доступа',
      'Нажмите «Продолжить» и подтвердите доступ Google, затем повторите действие.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw e;
  }
}

function payLogEnsurePayColsMenu() {
  const added = payLogEnsurePayCols_();
  SpreadsheetApp.getUi().alert(
    added.length
      ? 'Добавлены колонки в «Затраты рейса»: ' + added.join(', ')
      : 'Все необходимые колонки уже есть в «Затраты рейса».'
  );
}

/* ===================== Helpers ===================== */

function payLogProp_(key, fallback) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  return v == null || String(v).trim() === '' ? fallback : String(v).trim();
}

function payLogNorm_(v) { return String(v == null ? '' : v).trim(); }
function payLogCanon_(v) { return payLogNorm_(v).toLowerCase().replace(/\s+/g, ''); }

function payLogHeaderCanon_(s) {
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

function payLogHeaderMap_(headers) {
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    const k = payLogHeaderCanon_(headers[i]);
    if (k && map[k] == null) map[k] = i;
  }
  return map;
}

function payLogFindCol_(map, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const hit = map[payLogHeaderCanon_(aliases[i])];
    if (hit != null) return hit;
  }
  return -1;
}

function payLogParseAmount_(raw) {
  if (raw == null || raw === '') return 0;
  if (typeof raw === 'number' && isFinite(raw)) return raw;
  const cleaned = String(raw)
    .replace(/\u00A0/g, ' ')
    .replace(/\s/g, '')
    .replace(/,/g, '.')
    .replace(/[^\d.\-]/g, '');
  const n = parseFloat(cleaned);
  return isFinite(n) ? n : 0;
}

/** Сумма к оплате в валюте заявки: «Сумма_в_валюте»/«Сумма»; для RUB — запасной «Сумма_RUB». */
function payLogAmountFromExpenseRow_(row, cols, currency) {
  const cur = payLogNorm_(currency).toUpperCase();
  if (cols.amount >= 0) {
    const a = payLogParseAmount_(row[cols.amount]);
    if (a > 0) return a;
  }
  if (cols.amountRub >= 0 && cur === 'RUB') return payLogParseAmount_(row[cols.amountRub]);
  if (cols.amount >= 0) return payLogParseAmount_(row[cols.amount]);
  if (cols.amountRub >= 0) return payLogParseAmount_(row[cols.amountRub]);
  return 0;
}

function payLogQueueId_() {
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss-SSS');
  const rnd = Math.floor(Math.random() * 1000);
  return 'LRQ-' + ts + '-' + String(rnd).padStart(3, '0');
}

function payLogGetExpensesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (let i = 0; i < PAY_LOG_CFG.EXPENSES_ALIASES.length; i++) {
    const sh = ss.getSheetByName(PAY_LOG_CFG.EXPENSES_ALIASES[i]);
    if (sh) return sh;
  }
  throw new Error('Не найден лист «Затраты рейса» (искал по: ' + PAY_LOG_CFG.EXPENSES_ALIASES.join(', ') + ').');
}

function payLogReadExpensesHeader_(sh) {
  const lastCol = Math.max(sh.getLastColumn(), 30);
  const headers = sh.getRange(PAY_LOG_CFG.EXPENSES_HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const map = payLogHeaderMap_(headers);
  return { headers: headers, map: map, lastCol: lastCol };
}

function payLogResolveCols_(map, headers, required) {
  const out = {};
  for (const k in PAY_LOG_COL_ALIASES) {
    out[k] = payLogFindCol_(map, PAY_LOG_COL_ALIASES[k]);
  }
  if (required) {
    const missing = [];
    for (let i = 0; i < required.length; i++) {
      const key = required[i];
      if (key === 'amount') {
        if (out.amount < 0 && out.amountRub < 0) missing.push('Сумма_в_валюте / Сумма / Сумма_RUB');
        continue;
      }
      if (out[key] < 0) missing.push(PAY_LOG_COL_ALIASES[key][0]);
    }
    if (missing.length) {
      const visible = headers
        .map(function (h) { return String(h == null ? '' : h).trim(); })
        .filter(function (s) { return s; })
        .slice(0, 30)
        .join(' | ');
      throw new Error(
        'В «Затраты рейса» не нашёл обязательные колонки: ' + missing.join(', ') +
          '.\nФактические заголовки строки ' + PAY_LOG_CFG.EXPENSES_HEADER_ROW + ': ' + visible
      );
    }
  }
  return out;
}

/* ===================== Ensure pay-cols ===================== */

/**
 * Добавляет в «Затраты рейса» недостающие колонки оплат (4 шт.) в конец шапки.
 * Возвращает список реально добавленных названий (пустой массив, если всё уже было).
 */
function payLogEnsurePayCols_() {
  const sh = payLogGetExpensesSheet_();
  const hdr = payLogReadExpensesHeader_(sh);
  const added = [];
  let lastCol = sh.getLastColumn();
  for (let i = 0; i < PAY_LOG_NEW_COLS.length; i++) {
    const name = PAY_LOG_NEW_COLS[i];
    const aliasGroup =
      name === '№ заявки' ? PAY_LOG_COL_ALIASES.payRequestNo
      : name === 'Статус оплаты' ? PAY_LOG_COL_ALIASES.payStatus
      : name === 'Дата оплаты' ? PAY_LOG_COL_ALIASES.payDate
      : PAY_LOG_COL_ALIASES.payFolder;
    if (payLogFindCol_(hdr.map, aliasGroup) >= 0) continue;
    lastCol += 1;
    sh.getRange(PAY_LOG_CFG.EXPENSES_HEADER_ROW, lastCol).setValue(name).setFontWeight('bold');
    hdr.map[payLogHeaderCanon_(name)] = lastCol - 1;
    added.push(name);
  }
  return added;
}

/* ===================== Подача заявки ===================== */

function payLogCollectSelectedExpenseRows_() {
  const sh = payLogGetExpensesSheet_();
  if (sh.getSheetId() !== SpreadsheetApp.getActiveSheet().getSheetId()) {
    throw new Error('Откройте лист «Затраты рейса» и выделите там строки расходов, по которым подаём заявку.');
  }
  const ranges = SpreadsheetApp.getActiveRangeList()
    ? SpreadsheetApp.getActiveRangeList().getRanges()
    : [SpreadsheetApp.getActiveRange()];
  const rowsSet = {};
  for (let i = 0; i < ranges.length; i++) {
    const r = ranges[i];
    const start = r.getRow();
    const num = r.getNumRows();
    for (let k = 0; k < num; k++) {
      const abs = start + k;
      if (abs >= PAY_LOG_CFG.EXPENSES_DATA_START_ROW) rowsSet[abs] = true;
    }
  }
  const rowIdxs = Object.keys(rowsSet).map(function (x) { return parseInt(x, 10); }).sort(function (a, b) { return a - b; });
  if (!rowIdxs.length) throw new Error('Выделите 1+ строк с расходами на листе «Затраты рейса».');
  return { sheet: sh, rowIdxs: rowIdxs };
}

function payLogReadSelectedRows_(sh, rowIdxs) {
  payLogEnsurePayCols_();
  const hdr = payLogReadExpensesHeader_(sh);
  const cols = payLogResolveCols_(hdr.map, hdr.headers, ['shipment', 'counterparty', 'currency', 'amount']);
  const lastCol = Math.max(sh.getLastColumn(), hdr.lastCol);
  const out = [];
  for (let i = 0; i < rowIdxs.length; i++) {
    const r = rowIdxs[i];
    const row = sh.getRange(r, 1, 1, lastCol).getValues()[0];
    out.push({
      absRow: r,
      shipmentId: payLogNorm_(row[cols.shipment]),
      article: cols.article >= 0 ? payLogNorm_(row[cols.article]) : '',
      counterparty: payLogNorm_(row[cols.counterparty]),
      currency: payLogNorm_(row[cols.currency]).toUpperCase(),
      amount: payLogAmountFromExpenseRow_(row, cols, row[cols.currency]),
      payRequestNo: cols.payRequestNo >= 0 ? payLogNorm_(row[cols.payRequestNo]) : '',
      payStatus: cols.payStatus >= 0 ? payLogNorm_(row[cols.payStatus]) : ''
    });
  }
  return { rows: out, cols: cols };
}

function payLogValidateSelection_(rows) {
  const problems = [];
  let cp = '', cur = '';
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!r.shipmentId) problems.push('строка ' + r.absRow + ': пустой SHIPMENT_ID');
    if (!r.counterparty) problems.push('строка ' + r.absRow + ': пустой Контрагент');
    if (!r.currency || !/^[A-Z]{3}$/.test(r.currency)) problems.push('строка ' + r.absRow + ': пустая или некорректная Валюта (нужен код ISO, напр. CNY)');
    if (!(r.amount > 0)) problems.push('строка ' + r.absRow + ': Сумма_в_валюте (или Сумма_RUB для RUB) ≤ 0 или не число');
    const active = r.payStatus && /^(на проверке|на согласовании|оплачено)$/i.test(r.payStatus);
    if (active) problems.push('строка ' + r.absRow + ': уже есть активная заявка (' + r.payRequestNo + ' / ' + r.payStatus + ')');
    if (i === 0) { cp = r.counterparty; cur = r.currency; }
    else {
      if (payLogCanon_(r.counterparty) !== payLogCanon_(cp)) problems.push('строка ' + r.absRow + ': Контрагент отличается от первой строки выделения');
      if (r.currency !== cur) problems.push('строка ' + r.absRow + ': Валюта отличается от первой строки выделения');
    }
  }
  if (problems.length) {
    throw new Error('Заявку подать нельзя — исправьте:\n  • ' + problems.join('\n  • '));
  }
  return { counterparty: cp, currency: cur };
}

function payLogOpenRequestDialog() {
  const sel = payLogCollectSelectedExpenseRows_();
  const read = payLogReadSelectedRows_(sel.sheet, sel.rowIdxs);
  const meta = payLogValidateSelection_(read.rows);
  const total = read.rows.reduce(function (a, x) { return a + (isFinite(x.amount) ? x.amount : 0); }, 0);
  const shipments = payLogUniq_(read.rows.map(function (x) { return x.shipmentId; }));
  const articles = payLogUniq_(read.rows.map(function (x) { return x.article; }).filter(function (s) { return s; }));
  const defaultPurpose =
    'Услуги логистики' +
    (articles.length ? ' (' + articles.join('; ') + ')' : '') +
    ' по рейсу ' + shipments.join(', ');

  const payload = {
    counterparty: meta.counterparty,
    currency: meta.currency,
    amount: +total.toFixed(2),
    shipments: shipments,
    articles: articles,
    rowsPreview: read.rows.map(function (x) {
      return {
        absRow: x.absRow,
        shipmentId: x.shipmentId,
        article: x.article,
        amount: +Number(x.amount).toFixed(2)
      };
    }),
    purpose: defaultPurpose
  };
  const html = HtmlService.createHtmlOutput(payLogRequestDialogHtml_(payload)).setWidth(700).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, 'Заявка на оплату услуг логистики');
}

function payLogUniq_(arr) {
  const seen = {};
  const out = [];
  for (let i = 0; i < arr.length; i++) {
    const v = String(arr[i] == null ? '' : arr[i]);
    if (!v) continue;
    if (seen[v]) continue;
    seen[v] = true;
    out.push(v);
  }
  return out;
}

function payLogEscapeHtml_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function payLogRequestDialogHtml_(p) {
  const rowsHtml = p.rowsPreview.map(function (r) {
    return '<tr><td>' + r.absRow + '</td><td>' + payLogEscapeHtml_(r.shipmentId) + '</td><td>' +
      payLogEscapeHtml_(r.article) + '</td><td style="text-align:right">' + r.amount.toFixed(2) + '</td></tr>';
  }).join('');
  return (
    '<!doctype html><html><head><base target="_top"><meta charset="utf-8">' +
    '<style>body{font-family:Arial;padding:14px;font-size:13px}' +
    '.row{margin:8px 0}label{display:block;font-weight:600;margin-bottom:4px}' +
    'input,select,textarea{width:100%;padding:7px;box-sizing:border-box;font-family:inherit;font-size:13px}' +
    'textarea{height:60px}' +
    'table{width:100%;border-collapse:collapse;margin-top:4px}' +
    'th,td{border:1px solid #ddd;padding:4px 6px;font-size:12px}' +
    'th{background:#f6f6f6}' +
    '.summary{background:#f6faff;border:1px solid #d6e6ff;padding:8px;border-radius:4px;font-size:13px}' +
    '.note{font-size:11px;color:#666;margin-top:8px}' +
    'button{padding:8px 14px;font-size:13px}' +
    '</style></head><body>' +
    '<div class="summary"><b>Контрагент:</b> ' + payLogEscapeHtml_(p.counterparty) + '<br>' +
      '<b>Валюта:</b> ' + payLogEscapeHtml_(p.currency) + '<br>' +
      '<b>Рейсы:</b> ' + p.shipments.map(payLogEscapeHtml_).join(', ') + '<br>' +
      '<b>Сумма по выделению:</b> ' + p.amount.toFixed(2) + ' ' + payLogEscapeHtml_(p.currency) + '</div>' +
    '<div class="row"><table><thead><tr><th>Строка</th><th>SHIPMENT_ID</th><th>Статья</th><th>Сумма</th></tr></thead>' +
      '<tbody>' + rowsHtml + '</tbody></table></div>' +
    '<div class="row"><label>Назначение платежа (можно отредактировать)</label>' +
      '<textarea id="purpose">' + payLogEscapeHtml_(p.purpose) + '</textarea></div>' +
    '<div class="row"><label>Оплатить до</label><input id="due" type="date"></div>' +
    '<div class="row"><label>Подписанный документ (обязательно)</label><input id="files" type="file" multiple></div>' +
    '<div class="row"><button id="sendBtn" onclick="send()">Отправить на проверку</button></div>' +
    '<p class="note">Крупные файлы отправляются частями (без одного большого запроса).</p>' +
    '<script>var sending=false;var CHUNK=450000;' +
    'function toB64(f){return new Promise(function(res,rej){var r=new FileReader();r.onload=function(){res(String(r.result).split(\",\")[1]);};r.onerror=rej;r.readAsDataURL(f);});}' +
    'function runPrepare(payload){return new Promise(function(ok,er){google.script.run.withSuccessHandler(ok).withFailureHandler(er).payLogDialogPrepare(payload);});}' +
    'function runChunk(token,folderId,fi,ci,total,name,mime,part){return new Promise(function(ok,er){google.script.run.withSuccessHandler(ok).withFailureHandler(er).payLogDialogUploadChunk(token,folderId,fi,ci,total,name,mime,part);});}' +
    'function runFinalize(token,folderId){return new Promise(function(ok,er){google.script.run.withSuccessHandler(ok).withFailureHandler(er).payLogDialogFinalize(token,folderId);});}' +
    'async function send(){var purpose=document.getElementById(\"purpose\").value;var due=document.getElementById(\"due\").value;var fs=[].slice.call(document.getElementById(\"files\").files);' +
    'if(!fs.length){alert(\"Нужен подписанный документ\");return;}if(sending)return;sending=true;document.getElementById(\"sendBtn\").disabled=true;' +
    'try{var payload={purpose:purpose,dueDate:due};var prep=await runPrepare(payload);' +
    'for(var fi=0;fi<fs.length;fi++){var b64=await toB64(fs[fi]);var total=Math.max(1,Math.ceil(b64.length/CHUNK));' +
    'for(var ci=0;ci<total;ci++){var part=b64.substring(ci*CHUNK,(ci+1)*CHUNK);await runChunk(prep.uploadToken,prep.folderId,fi,ci,total,fs[fi].name,fs[fi].type||\'application/octet-stream\',part);}}' +
    'var msg=await runFinalize(prep.uploadToken,prep.folderId);alert(msg);google.script.host.close();}catch(e){alert(e.message||e);}finally{sending=false;document.getElementById(\"sendBtn\").disabled=false;}}</script>' +
    '</body></html>'
  );
}

/* ===================== Prepare / Upload chunks / Finalize ===================== */

function payLogDialogPrepare(payload) { return payLogPrepareRequest_(payload); }
function payLogDialogUploadChunk(uploadToken, folderId, fileIndex, chunkIndex, totalChunks, fileName, mimeType, base64Chunk) {
  payLogUploadFileChunk_(uploadToken, folderId, fileIndex, chunkIndex, totalChunks, fileName, mimeType, base64Chunk);
}
function payLogDialogFinalize(uploadToken, folderId) { return payLogFinalizeRequest_(uploadToken, folderId); }

function payLogGetDraftRootFolder_() {
  const rootId = payLogProp_('PAYMENT_REGISTRY_ROOT_FOLDER_ID', PAY_LOG_CFG.ROOT_FOLDER_DEFAULT);
  const root = DriveApp.getFolderById(rootId);
  const draftName = payLogProp_('PAYMENT_DRAFT_FOLDER_NAME_LOG', PAY_LOG_CFG.DRAFT_FOLDER_NAME_DEFAULT);
  const it = root.getFoldersByName(draftName);
  if (it.hasNext()) return it.next();
  return root.createFolder(draftName);
}

function payLogPrepareRequest_(payload) {
  const sel = payLogCollectSelectedExpenseRows_();
  const read = payLogReadSelectedRows_(sel.sheet, sel.rowIdxs);
  const meta = payLogValidateSelection_(read.rows);
  const total = read.rows.reduce(function (a, x) { return a + (isFinite(x.amount) ? x.amount : 0); }, 0);
  const shipments = payLogUniq_(read.rows.map(function (x) { return x.shipmentId; }));
  const articles = payLogUniq_(read.rows.map(function (x) { return x.article; }).filter(function (s) { return s; }));

  const purpose = payLogNorm_(payload && payload.purpose);
  if (!purpose) throw new Error('Назначение платежа обязательно.');

  const queueId = payLogQueueId_();
  const root = payLogGetDraftRootFolder_();
  const folder = root.createFolder(queueId + ' - ' + meta.counterparty);

  const draft = {
    v: 1,
    queueId: queueId,
    createdBy: Session.getActiveUser().getEmail() || '',
    expenseRows: read.rows.map(function (x) { return x.absRow; }),
    shipmentIds: shipments,
    articles: articles,
    counterparty: meta.counterparty,
    currency: meta.currency,
    amount: +Number(total).toFixed(2),
    dueDate: payload && payload.dueDate ? String(payload.dueDate) : '',
    purpose: purpose
  };
  const metaBlob = Utilities.newBlob(JSON.stringify(draft), 'application/json', '__pay_upload_meta.json');
  folder.createFile(metaBlob);
  return { uploadToken: queueId, folderId: folder.getId(), queueId: queueId };
}

function payLogUploadFileChunk_(uploadToken, folderId, fileIndex, chunkIndex, totalChunks, fileName, mimeType, base64Chunk) {
  const folder = DriveApp.getFolderById(folderId);
  payLogAssertUploadMeta_(folder, uploadToken);
  const fi = Number(fileIndex);
  const ci = Number(chunkIndex);
  const tc = Number(totalChunks);
  if (!isFinite(fi) || fi < 0 || !isFinite(ci) || ci < 0 || !isFinite(tc) || tc < 1) {
    throw new Error('Некорректные параметры загрузки файла.');
  }
  if (ci === 0) {
    payLogTrashFileByName_(folder, '__fileinfo_f' + fi + '.json');
    const infoBlob = Utilities.newBlob(
      JSON.stringify({ name: fileName || 'file', mimeType: mimeType || 'application/octet-stream', totalChunks: tc }),
      'application/json',
      '__fileinfo_f' + fi + '.json'
    );
    folder.createFile(infoBlob);
  }
  const partName = '__part_f' + fi + '_c' + ci + '.b64';
  payLogTrashFileByName_(folder, partName);
  folder.createFile(Utilities.newBlob(String(base64Chunk || ''), 'text/plain', partName));
}

function payLogTrashFileByName_(folder, name) {
  const it = folder.getFilesByName(name);
  while (it.hasNext()) it.next().setTrashed(true);
}

function payLogFinalizeRequest_(uploadToken, folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const meta = payLogReadUploadMeta_(folder);
  if (!meta || payLogNorm_(meta.queueId) !== payLogNorm_(uploadToken)) {
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
      try { fileInfoByIdx[idx] = JSON.parse(f.getBlob().getDataAsString()); }
      catch (e) { throw new Error('Повреждены служебные данные загрузки.'); }
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
    .map(function (x) { return parseInt(x, 10); })
    .filter(function (n) { return isFinite(n); })
    .sort(function (a, b) { return a - b; });
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
  payLogTrashUploadTempFiles_(folder);

  // Запись в очередь книги 05.
  const sh = payLogEnsureQueueSheet_();
  const newRow = sh.getLastRow() + 1;
  sh.getRange(newRow, 1, 1, 17).setValues([[
    meta.queueId,
    new Date(),
    meta.createdBy,
    JSON.stringify(meta.expenseRows || []),
    (meta.shipmentIds || []).join(', '),
    (meta.articles || []).join('; '),
    meta.counterparty,
    meta.amount,
    meta.currency,
    meta.dueDate ? new Date(meta.dueDate) : '',
    meta.purpose,
    folder.getUrl(),
    JSON.stringify(fileLinks),
    'На проверке',
    '',
    '',
    ''
  ]]);

  // Метим строки «Затраты рейса».
  payLogWriteStatusToExpenseRows_(meta.expenseRows || [], {
    payRequestNo: meta.queueId,
    payStatus: 'На проверке',
    payDate: '',
    payFolder: folder.getUrl()
  });

  return 'Заявка создана: ' + meta.queueId;
}

function payLogAssertUploadMeta_(folder, uploadToken) {
  const meta = payLogReadUploadMeta_(folder);
  if (!meta || payLogNorm_(meta.queueId) !== payLogNorm_(uploadToken)) {
    throw new Error('Неверный токен загрузки или папка.');
  }
}

function payLogReadUploadMeta_(folder) {
  const it = folder.getFilesByName('__pay_upload_meta.json');
  if (!it.hasNext()) return null;
  try { return JSON.parse(it.next().getBlob().getDataAsString()); } catch (e) { return null; }
}

function payLogTrashUploadTempFiles_(folder) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    const n = f.getName();
    if (
      n === '__pay_upload_meta.json' ||
      /^__fileinfo_f\d+\.json$/i.test(n) ||
      /^__part_f\d+_c\d+\.b64$/i.test(n)
    ) f.setTrashed(true);
  }
}

/* ===================== Очередь + запись в строки «Затраты рейса» ===================== */

function payLogEnsureQueueSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(PAY_LOG_CFG.QUEUE_SHEET);
  if (!sh) sh = ss.insertSheet(PAY_LOG_CFG.QUEUE_SHEET);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, 17).setValues([[
      'queue_id', 'created_at', 'created_by', 'expense_rows_json', 'shipment_ids',
      'articles', 'counterparty', 'amount', 'currency', 'due_date', 'payment_purpose',
      'folder_url', 'file_links_json', 'status', 'registry_request_no', 'registry_row', 'reject_reason'
    ]]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function payLogWriteStatusToExpenseRows_(absRows, patch) {
  if (!absRows || !absRows.length) return;
  const sh = payLogGetExpensesSheet_();
  const hdr = payLogReadExpensesHeader_(sh);
  const cols = payLogResolveCols_(hdr.map, hdr.headers, []);
  for (let i = 0; i < absRows.length; i++) {
    const r = absRows[i];
    if (!r) continue;
    if (cols.payRequestNo >= 0 && patch.payRequestNo !== undefined) sh.getRange(r, cols.payRequestNo + 1).setValue(patch.payRequestNo);
    if (cols.payStatus >= 0 && patch.payStatus !== undefined) sh.getRange(r, cols.payStatus + 1).setValue(patch.payStatus);
    if (cols.payDate >= 0 && patch.payDate !== undefined) sh.getRange(r, cols.payDate + 1).setValue(patch.payDate);
    if (cols.payFolder >= 0 && patch.payFolder !== undefined) sh.getRange(r, cols.payFolder + 1).setValue(patch.payFolder);
  }
}

/* ===================== Согласование ===================== */

function payLogAssertApprover_() {
  const allowed = payLogProp_('PAYMENT_APPROVER_EMAIL', PAY_LOG_CFG.APPROVER_EMAIL_DEFAULT);
  const me = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!me || me !== allowed.toLowerCase()) {
    throw new Error('Только согласующий (' + allowed + ') может выполнять это действие.');
  }
}

function payLogOpenApprovalDialog() {
  payLogAssertApprover_();
  const list = payLogGetPendingQueue_();
  if (!list.length) {
    SpreadsheetApp.getUi().alert('Нет логистических заявок на проверке.');
    return;
  }
  const html = HtmlService.createHtmlOutput(payLogApprovalDialogHtml_(list)).setWidth(820).setHeight(580);
  SpreadsheetApp.getUi().showModalDialog(html, 'Проверка логистических заявок');
}

function payLogGetPendingQueue_() {
  const sh = payLogEnsureQueueSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const data = sh.getRange(2, 1, last - 1, 17).getValues();
  const out = [];
  const seenById = {};
  const seenBySig = {};
  data.forEach(function (r, i) {
    if (payLogNorm_(r[13]) !== 'На проверке') return;
    const qid = payLogNorm_(r[0]);
    const sig = payLogQueueSig_(r);
    if (qid && seenById[qid]) return;
    if (sig && seenBySig[sig]) return;
    if (qid) seenById[qid] = true;
    if (sig) seenBySig[sig] = true;
    out.push({
      row: i + 2,
      queueId: qid || ('ROW-' + (i + 2)),
      counterparty: r[6],
      amount: r[7],
      currency: r[8],
      dueDate: r[9] ? Utilities.formatDate(new Date(r[9]), Session.getScriptTimeZone(), 'dd.MM.yyyy') : '',
      shipments: r[4],
      articles: r[5],
      purpose: r[10],
      folderUrl: r[11],
      fileLinks: (function () { try { return JSON.parse(r[12] || '[]'); } catch (e) { return []; } })()
    });
  });
  return out;
}

function payLogQueueSig_(r) {
  let expenseRows = '';
  try { expenseRows = (JSON.parse(r[3] || '[]') || []).slice().sort(function (a, b) { return a - b; }).join(','); } catch (e) {}
  return [
    payLogCanon_(r[6]),
    payLogCanon_(r[8]),
    payLogCanon_(r[4]),
    expenseRows,
    payLogCanon_(r[10])
  ].join('|');
}

function payLogApprovalDialogHtml_(list) {
  return (
    '<!doctype html><html><head><base target="_top"><meta charset="utf-8"><style>body{font-family:Arial;padding:12px}' +
    'table{width:100%;border-collapse:collapse}td,th{border:1px solid #ddd;padding:6px;font-size:12px;vertical-align:top}' +
    'button{padding:4px 8px}</style></head><body>' +
    '<table><tr><th>ID</th><th>Контрагент</th><th>Рейсы</th><th>Статьи</th><th>Назначение</th><th>Сумма</th><th>Срок</th><th>Документы</th><th>Действие</th></tr>' +
    list.map(function (x) {
      var docs = [];
      if (x.folderUrl) docs.push('<a href="' + x.folderUrl + '" target="_blank">Папка</a>');
      if (x.fileLinks && x.fileLinks.length) {
        for (var i = 0; i < x.fileLinks.length; i++) docs.push('<a href="' + x.fileLinks[i] + '" target="_blank">Файл ' + (i + 1) + '</a>');
      }
      return '<tr><td>' + payLogEscapeHtml_(x.queueId) + '</td><td>' + payLogEscapeHtml_(x.counterparty) +
        '</td><td>' + payLogEscapeHtml_(x.shipments) + '</td><td>' + payLogEscapeHtml_(x.articles) +
        '</td><td>' + payLogEscapeHtml_(x.purpose) + '</td><td>' + Number(x.amount).toFixed(2) + ' ' + payLogEscapeHtml_(x.currency) +
        '</td><td>' + payLogEscapeHtml_(x.dueDate) + '</td><td>' + (docs.join('<br>') || '-') +
        '</td><td><button onclick="ap(' + x.row + ')">Одобрить</button> <button onclick="rej(' + x.row + ')">Отклонить</button></td></tr>';
    }).join('') +
    '</table><script>var busy=false;function lockUi(){if(busy)return false;busy=true;var bs=document.querySelectorAll(\"button\");for(var i=0;i<bs.length;i++)bs[i].disabled=true;return true;}function unlockUi(){busy=false;var bs=document.querySelectorAll(\"button\");for(var i=0;i<bs.length;i++)bs[i].disabled=false;}' +
    'function ap(row){if(!lockUi())return;google.script.run.withSuccessHandler(function(m){alert(m);google.script.host.close();}).withFailureHandler(function(e){unlockUi();alert(e.message||e);}).payLogApproveQueueRow(row);}' +
    'function rej(row){if(!lockUi())return;var reason=prompt(\"Причина отклонения:\",\"\");if(reason===null){unlockUi();return;}if(!String(reason).trim()){unlockUi();alert(\"Укажите причину\");return;}google.script.run.withSuccessHandler(function(m){alert(m);google.script.host.close();}).withFailureHandler(function(e){unlockUi();alert(e.message||e);}).payLogRejectQueueRow(row, reason);}' +
    '</script></body></html>'
  );
}

function payLogApproveQueueRow(queueRow) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    payLogAssertApprover_();
    const sh = payLogEnsureQueueSheet_();
    const r = sh.getRange(queueRow, 1, 1, 17).getValues()[0];
    const status = payLogNorm_(r[13]);
    if (status !== 'На проверке') {
      const existingReq = payLogNorm_(r[14]);
      if (existingReq) return 'Заявка уже обработана ранее: ' + existingReq;
      throw new Error('Строка уже обработана.');
    }

    const queueId = r[0];
    const expenseRows = (function () { try { return JSON.parse(r[3] || '[]') || []; } catch (e) { return []; } })();
    const shipmentsCsv = payLogNorm_(r[4]);
    const articlesCsv = payLogNorm_(r[5]);
    const counterparty = r[6];
    const amount = r[7];
    const currency = payLogNorm_(r[8]).toUpperCase() || 'CNY';
    const dueDate = r[9];
    const purpose = payLogNorm_(r[10]);
    const draftFolderUrl = payLogNorm_(r[11]);
    const fileLinks = (function () { try { return JSON.parse(r[12] || '[]') || []; } catch (e) { return []; } })();

    const registryId = payLogProp_('PAYMENT_REGISTRY_SPREADSHEET_ID', PAY_LOG_CFG.REGISTRY_ID_DEFAULT);
    const registrySheetName = payLogProp_('PAYMENT_REGISTRY_SHEET_NAME', PAY_LOG_CFG.REGISTRY_SHEET_DEFAULT);
    const regSs = SpreadsheetApp.openById(registryId);
    const regSh = regSs.getSheetByName(registrySheetName);
    if (!regSh) throw new Error('Лист реестра не найден: ' + registrySheetName);

    const last = regSh.getLastRow();
    const reqNo = 'ЗАЯВКА-' + new Date().getFullYear() + '-' + String(last).padStart(4, '0');
    const initiator = payLogProp_('PAYMENT_INITIATOR_NAME', PAY_LOG_CFG.INITIATOR_DEFAULT);
    const department = payLogProp_('PAYMENT_LOG_DEPARTMENT', PAY_LOG_CFG.DEPARTMENT_DEFAULT);
    const approver = Session.getActiveUser().getEmail() || '';
    const finalFolderUrl = payLogCreateApprovedFolder_(reqNo, counterparty, fileLinks, draftFolderUrl);
    const comment = 'Передано из логистической заявки ' + queueId + '. Рейс: ' + shipmentsCsv;

    const newRow = last + 1;
    regSh.insertRowAfter(last);
    regSh.getRange(newRow, 1).setValue(new Date());
    regSh.getRange(newRow, 2).setValue(dueDate || '');
    regSh.getRange(newRow, 3).setValue(initiator);
    regSh.getRange(newRow, 4).setValue(department);
    regSh.getRange(newRow, 5).setValue(counterparty);
    regSh.getRange(newRow, 6).setValue(purpose);
    regSh.getRange(newRow, 7).setValue(amount);
    regSh.getRange(newRow, 8).setValue(currency);
    regSh.getRange(newRow, 9).setValue(fileLinks.join('\n'));
    regSh.getRange(newRow, 10).setValue('На согласовании');
    regSh.getRange(newRow, 11).setValue('');
    regSh.getRange(newRow, 12).setValue('');
    regSh.getRange(newRow, 13).setValue(comment);
    regSh.getRange(newRow, 14).setValue('Не оплачено');
    if (currency === 'RUB') {
      regSh.getRange(newRow, 16).setValue(amount);
    } else {
      regSh.getRange(newRow, 16).setFormula('=G' + newRow + '*GOOGLEFINANCE("CURRENCY:' + currency + 'RUB")');
    }
    regSh.getRange(newRow, 17).setValue(approver);
    regSh.getRange(newRow, 18).setValue(reqNo);
    regSh.getRange(newRow, 19).setValue(finalFolderUrl);

    sh.getRange(queueRow, 12).setValue(finalFolderUrl);
    sh.getRange(queueRow, 14).setValue('Отправлено в реестр');
    sh.getRange(queueRow, 15).setValue(reqNo);
    sh.getRange(queueRow, 16).setValue(newRow);
    sh.getRange(queueRow, 17).setValue('');

    payLogWriteStatusToExpenseRows_(expenseRows, {
      payRequestNo: reqNo,
      payStatus: 'На согласовании',
      payDate: '',
      payFolder: finalFolderUrl
    });

    // Авто-отклонение дублей этой же заявки в очереди.
    const allLast = sh.getLastRow();
    if (allLast >= 2) {
      const all = sh.getRange(2, 1, allLast - 1, 17).getValues();
      const sig = payLogQueueSig_(r);
      for (let i = 0; i < all.length; i++) {
        const abs = i + 2;
        if (abs === queueRow) continue;
        if (payLogNorm_(all[i][13]) !== 'На проверке') continue;
        if (payLogQueueSig_(all[i]) !== sig) continue;
        sh.getRange(abs, 14).setValue('Отклонено');
        sh.getRange(abs, 17).setValue('Авто: дубликат заявки, одобрена ' + reqNo);
      }
    }

    return 'Заявка отправлена в реестр: ' + reqNo;
  } finally {
    lock.releaseLock();
  }
}

function payLogRejectQueueRow(queueRow, reason) {
  payLogAssertApprover_();
  const sh = payLogEnsureQueueSheet_();
  const r = sh.getRange(queueRow, 1, 1, 17).getValues()[0];
  if (payLogNorm_(r[13]) !== 'На проверке') {
    throw new Error('Можно отклонить только заявки со статусом «На проверке».');
  }
  const txt = payLogNorm_(reason);
  if (!txt) throw new Error('Укажите причину отклонения.');
  sh.getRange(queueRow, 14).setValue('Отклонено');
  sh.getRange(queueRow, 17).setValue(txt);
  const expenseRows = (function () { try { return JSON.parse(r[3] || '[]') || []; } catch (e) { return []; } })();
  payLogWriteStatusToExpenseRows_(expenseRows, {
    payRequestNo: '',
    payStatus: 'Отклонено',
    payDate: '',
    payFolder: ''
  });
  return 'Заявка отклонена.';
}

function payLogCreateApprovedFolder_(requestNo, counterparty, sourceFileLinks, draftFolderUrl) {
  const rootId = payLogProp_('PAYMENT_REGISTRY_ROOT_FOLDER_ID', PAY_LOG_CFG.ROOT_FOLDER_DEFAULT);
  const root = DriveApp.getFolderById(rootId);
  const safeCounterparty = String(counterparty || 'Контрагент').replace(/[^\w\sа-яА-ЯёЁ.-]/g, '').trim();
  const finalFolder = root.createFolder(requestNo + ' - ' + (safeCounterparty || 'Контрагент'));

  (sourceFileLinks || []).forEach(function (url) {
    const fileId = payLogExtractDriveFileId_(url);
    if (!fileId) return;
    try {
      const f = DriveApp.getFileById(fileId);
      f.makeCopy(f.getName(), finalFolder);
    } catch (e) {
      Logger.log('[PAY-LOG][WARN] Не удалось скопировать файл в финальную папку: ' + url + ' | ' + e);
    }
  });

  try {
    finalFolder.createFile(
      'README.txt',
      'Заявка: ' + requestNo + '\nЧерновая папка логиста: ' + (draftFolderUrl || '-') + '\nЕсли части файлов нет, откройте черновую папку.'
    );
  } catch (e) {}

  return finalFolder.getUrl();
}

function payLogExtractDriveFileId_(url) {
  const s = String(url || '');
  const m1 = s.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m1) return m1[1];
  const m2 = s.match(/[?&]id=([a-zA-Z0-9_-]{20,})/);
  if (m2) return m2[1];
  return null;
}

/* ===================== Синхронизация факта оплаты ===================== */

function payLogSyncPaidStatuses() {
  const r = payLogSyncPaidStatusesImpl_({});
  if (r && r.message && !r.silent) {
    SpreadsheetApp.getUi().alert(r.message);
  }
}

/**
 * @param {{ silent?: boolean }} opt
 * @returns {{ updated: number, message: string, silent: boolean }}
 */
function payLogSyncPaidStatusesImpl_(opt) {
  const silent = !!(opt && opt.silent);
  const sh = payLogEnsureQueueSheet_();
  const last = sh.getLastRow();
  if (last < 2) {
    const msg = 'Нет логистических заявок для синхронизации.';
    return { updated: 0, message: msg, silent: silent };
  }
  const rows = sh.getRange(2, 1, last - 1, 17).getValues();

  const registryId = payLogProp_('PAYMENT_REGISTRY_SPREADSHEET_ID', PAY_LOG_CFG.REGISTRY_ID_DEFAULT);
  const registrySheetName = payLogProp_('PAYMENT_REGISTRY_SHEET_NAME', PAY_LOG_CFG.REGISTRY_SHEET_DEFAULT);
  const regSh = SpreadsheetApp.openById(registryId).getSheetByName(registrySheetName);
  if (!regSh) throw new Error('Лист реестра не найден.');
  const regLast = regSh.getLastRow();
  const regData = regLast >= 2 ? regSh.getRange(2, 1, regLast - 1, 19).getValues() : [];
  const regMap = {};
  regData.forEach(function (r) {
    const no = payLogNorm_(r[17]);
    if (!no) return;
    regMap[no] = { status: payLogNorm_(r[13]), paidDate: r[14], folder: payLogNorm_(r[18]) };
  });

  let updated = 0;
  for (let i = 0; i < rows.length; i++) {
    const q = rows[i];
    if (payLogNorm_(q[13]) !== 'Отправлено в реестр') continue;
    const reqNo = payLogNorm_(q[14]);
    if (!reqNo || !regMap[reqNo]) continue;
    const info = regMap[reqNo];
    if (info.status !== 'Оплачено' || !info.paidDate) continue;

    const expenseRows = (function () { try { return JSON.parse(q[3] || '[]') || []; } catch (e) { return []; } })();
    payLogWriteStatusToExpenseRows_(expenseRows, {
      payRequestNo: reqNo,
      payStatus: 'Оплачено',
      payDate: info.paidDate,
      payFolder: info.folder || payLogNorm_(q[11])
    });
    sh.getRange(i + 2, 14).setValue('Оплачено (синхр.)');
    updated++;
  }
  const msg = 'Синхронизация завершена. Обновлено заявок: ' + updated;
  return { updated: updated, message: msg, silent: silent };
}
