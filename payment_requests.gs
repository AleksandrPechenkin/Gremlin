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

function payCanon_(v) {
  return payNorm_(v).toLowerCase().replace(/\s+/g, '');
}

function payHeaderMap_(headers) {
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    const k = payCanon_(headers[i]);
    if (k) map[k] = i;
  }
  return map;
}

function payFindCol_(map, keys, fallback1Based) {
  for (let i = 0; i < keys.length; i++) {
    const hit = map[payCanon_(keys[i])];
    if (hit != null) return hit;
  }
  return fallback1Based - 1;
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
  const specCol = payFindCol_(map, ['Номер спецификации', 'Номер Спецификации'], 12);
  const amountCol = payFindCol_(map, ['Сумма'], 15);
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
    '<div class="row"><label>Оплатить до</label><input id="due" type="date"></div>' +
    '<div class="row"><label>Контрагент (обязательно)</label><input id="counterparty" list="cpList" placeholder="Введите или выберите"></div>' +
    '<datalist id="cpList">' +
    counterparties.map(function (c) { return '<option value="' + c + '"></option>'; }).join('') +
    '</datalist>' +
    '<div class="row"><label>Подписанный документ (обязательно)</label><input id="files" type="file" multiple></div>' +
    '<div class="row"><button onclick="send()">Отправить на проверку</button></div>' +
    '<script>function toB64(f){return new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(String(r.result).split(\",\")[1]);r.onerror=rej;r.readAsDataURL(f);});}' +
    'async function send(){const spec=document.getElementById(\"spec\").value;const ptype=document.getElementById(\"ptype\").value;const pct=document.getElementById(\"percent\").value;const due=document.getElementById(\"due\").value;const cp=document.getElementById(\"counterparty\").value;' +
    'const fs=[...document.getElementById(\"files\").files];if(!cp){alert(\"Контрагент обязателен\");return;}if(!fs.length){alert(\"Нужен подписанный документ\");return;}' +
    'const files=[];for(const f of fs){files.push({name:f.name,type:f.type||\"application/octet-stream\",data:await toB64(f)});}' +
    'google.script.run.withSuccessHandler(function(msg){alert(msg);google.script.host.close();}).withFailureHandler(function(e){alert(e.message||e);}).paySubmitManagerRequest({spec:spec,paymentType:ptype,percent:pct,dueDate:due,counterparty:cp},files);}</script>' +
    '</body></html>'
  );
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

function paySubmitManagerRequest(payload, files) {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const managerSheetName = sheet.getName();
  const managerName = payManagerNameFromSheet_(managerSheetName);
  const spec = payNorm_(payload.spec);
  const pType = payNorm_(payload.paymentType);
  const pct = parseFloat(payload.percent);
  const counterparty = payNorm_(payload.counterparty);
  if (!spec || !pType || !isFinite(pct) || pct <= 0 || pct > 100 || !counterparty) {
    throw new Error('Проверьте обязательные поля заявки.');
  }
  if (!files || !files.length) throw new Error('Нужно приложить подписанный документ.');

  const lastCol = Math.max(sheet.getLastColumn(), 30);
  const lastRow = sheet.getLastRow();
  const headers = sheet.getRange(PAY_CFG.HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const map = payHeaderMap_(headers);
  const specCol = payFindCol_(map, ['Номер спецификации', 'Номер Спецификации'], 12);
  const amountCol = payFindCol_(map, ['Сумма'], 15);

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
  const queueId = 'PRQ-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');

  const root = payGetDraftRootFolder_();
  const folder = root.createFolder(queueId + ' - ' + counterparty);
  const fileLinks = [];
  files.forEach(function (f) {
    const blob = Utilities.newBlob(Utilities.base64Decode(f.data), f.type || 'application/octet-stream', f.name || 'file');
    const gf = folder.createFile(blob);
    fileLinks.push(gf.getUrl());
  });

  const sh = payEnsureQueueSheet_();
  const newRow = sh.getLastRow() + 1;
  const createdBy = Session.getActiveUser().getEmail() || '';
  sh.getRange(newRow, 1, 1, 19).setValues([[
    queueId, new Date(), createdBy, managerSheetName, managerName, spec, pType,
    pct, counterparty, amount, 'CNY', payload.dueDate ? new Date(payload.dueDate) : '',
    JSON.stringify(matchedRows), folder.getUrl(), JSON.stringify(fileLinks), 'На проверке', '', '', ''
  ]]);

  paySaveCounterparty_(counterparty);
  ui.alert('Заявка отправлена на проверку: ' + queueId);
  return 'Заявка создана: ' + queueId;
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
  data.forEach(function (r, i) {
    if (payNorm_(r[15]) !== 'На проверке') return;
    out.push({
      row: i + 2,
      queueId: r[0],
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
    '</table><script>function ap(row){google.script.run.withSuccessHandler(function(m){alert(m);google.script.host.close();}).withFailureHandler(function(e){alert(e.message||e);}).payApproveQueueRow(row);}function rej(row){var reason=prompt("Причина отклонения:","");if(reason===null)return;if(!String(reason).trim()){alert("Укажите причину");return;}google.script.run.withSuccessHandler(function(m){alert(m);google.script.host.close();}).withFailureHandler(function(e){alert(e.message||e);}).payRejectQueueRow(row, reason);}</script></body></html>'
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
  payAssertApprover_();
  const sh = payEnsureQueueSheet_();
  const r = sh.getRange(queueRow, 1, 1, 19).getValues()[0];
  if (payNorm_(r[15]) !== 'На проверке') throw new Error('Строка уже обработана.');

  const queueId = r[0];
  const createdBy = r[2];
  const managerSheet = r[3];
  const managerName = r[4];
  const spec = r[5];
  const paymentType = r[6];
  const percent = r[7];
  const counterparty = r[8];
  const amount = r[9];
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
  regSh.getRange(newRow, 8).setValue('CNY');
  regSh.getRange(newRow, 9).setValue(fileLinks.join('\n'));
  regSh.getRange(newRow, 10).setValue('На согласовании');
  regSh.getRange(newRow, 11).setValue('');
  regSh.getRange(newRow, 12).setValue('');
  regSh.getRange(newRow, 13).setValue('Передано из менеджерской заявки ' + queueId);
  regSh.getRange(newRow, 14).setValue('Не оплачено');
  regSh.getRange(newRow, 17).setValue(approver);
  regSh.getRange(newRow, 18).setValue(reqNo);
  regSh.getRange(newRow, 19).setValue(finalFolderUrl);

  sh.getRange(queueRow, 16).setValue('Отправлено в реестр');
  sh.getRange(queueRow, 17).setValue(reqNo);
  sh.getRange(queueRow, 18).setValue(newRow);
  sh.getRange(queueRow, 19).setValue('');

  sh.getRange(queueRow, 14).setValue(finalFolderUrl);
  payWriteRequestToManagerAndSummary_(managerSheet, managerName, spec, managerRows, reqNo, finalFolderUrl);
  return 'Заявка отправлена в реестр: ' + reqNo;
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

function payWriteRequestToManagerAndSummary_(managerSheetName, managerName, spec, managerRows, requestNo, folderUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mgr = ss.getSheetByName(managerSheetName);
  if (mgr && managerRows && managerRows.length) {
    managerRows.forEach(function (r) {
      mgr.getRange(r, 28).setValue(requestNo); // AB
      mgr.getRange(r, 29).setValue(folderUrl); // AC
    });
  }

  const sumSh = ss.getSheetByName(PAY_CFG.SUMMARY_SHEET);
  if (!sumSh) return;
  const last = sumSh.getLastRow();
  if (last < PAY_CFG.DATA_START_ROW) return;
  const data = sumSh.getRange(PAY_CFG.DATA_START_ROW, 1, last - PAY_CFG.DATA_START_ROW + 1, 34).getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const sumSpec = payNorm_(row[11]); // L
    const sumManager = payNorm_(row[7]); // H
    if (sumSpec !== payNorm_(spec)) continue;
    if (payCanon_(sumManager) !== payCanon_(managerName)) continue;
    const abs = PAY_CFG.DATA_START_ROW + i;
    sumSh.getRange(abs, 33).setValue(requestNo); // AG
    sumSh.getRange(abs, 34).setValue(folderUrl); // AH
  }
}

function paySyncPaidStatuses() {
  const sh = payEnsureQueueSheet_();
  const last = sh.getLastRow();
  if (last < 2) return SpreadsheetApp.getUi().alert('Нет заявок для синхронизации.');
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
  SpreadsheetApp.getUi().alert('Синхронизация завершена. Обновлено заявок: ' + updated);
}

function payApplyFactDate_(managerSheetName, managerName, spec, managerRows, paymentType, paidDate, folderUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mgr = ss.getSheetByName(managerSheetName);
  let mgrCol = 19; // S
  let sumCol = 22; // V
  if (payCanon_(paymentType) === payCanon_('Баланс')) {
    mgrCol = 22; // V
    sumCol = 25; // Y
  } else if (payCanon_(paymentType) === payCanon_('Отсрочка')) {
    mgrCol = 25; // Y
    sumCol = 28; // AB
  }
  if (mgr && managerRows && managerRows.length) {
    managerRows.forEach(function (r) {
      mgr.getRange(r, mgrCol).setValue(paidDate);
      if (folderUrl) mgr.getRange(r, 29).setValue(folderUrl); // AC
    });
  }

  const sumSh = ss.getSheetByName(PAY_CFG.SUMMARY_SHEET);
  if (!sumSh) return;
  const last = sumSh.getLastRow();
  if (last < PAY_CFG.DATA_START_ROW) return;
  const data = sumSh.getRange(PAY_CFG.DATA_START_ROW, 1, last - PAY_CFG.DATA_START_ROW + 1, 34).getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (payNorm_(row[11]) !== payNorm_(spec)) continue; // L spec
    if (payCanon_(row[7]) !== payCanon_(managerName)) continue; // H manager
    const abs = PAY_CFG.DATA_START_ROW + i;
    sumSh.getRange(abs, sumCol).setValue(paidDate);
    if (folderUrl) sumSh.getRange(abs, 34).setValue(folderUrl); // AH
  }
}
