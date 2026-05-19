/**
 * Генерация файла Google Sheets: Invoice + Packing List по строкам вкладки менеджера.
 *
 * Script Properties (Файл → Настройки проекта → Свойства скрипта):
 *   INVOICE_DRIVE_FOLDER_ID — ID папки на Диске (из URL папки).
 *   MAIN_SPREADSHEET_ID — ID книги со сводной; если пусто, используется активная книга.
 *
 * Необязательно:
 *   INVOICE_REQUIRED_STATUS — текст статуса для запуска (по умолчанию: 04. Сформировать инвойс).
 *   INVOICE_MANAGER_LINK_COL_1BASED — колонка ссылки на вкладке менеджера (по умолчанию 27 = AA).
 *   INVOICE_SUMMARY_LINK_COL_1BASED — колонка ссылки на листе «Сводная» (по умолчанию 4 = D).
 *   INVOICE_STATUS_COL_1BASED — если заголовки нестандартные: номер колонки статуса (по умолчанию 8 = H).
 *   INVOICE_SPEC_COL_1BASED — номер колонки спецификации, если нет заголовка (по умолчанию 12 = L).
 *
 * Справочники продавца и условий оплаты хранятся в свойствах документа (Document Properties),
 * как в исходном скрипте: SAVED_SELLERS, SAVED_PAYMENT_TERMS (JSON-массивы).
 */

const INV_HEADER_ROW = 2;
const INV_DATA_START_ROW = 3;

const INV_DEFAULT_STATUS = '04. Сформировать инвойс';

const DEFAULT_SELLER = {
  name: 'YIWU DONGCHI IMPORT AND EXPORT CO.,LTD.',
  address:
    'First floor, No. 2, 6, 8, Changchun 8th Street, Futian Street, Yiwu City, Jinhua City, Zhejiang Province, China',
  contacts: 'E-mail: bavier_logistic@163.com    WhatsApp/WeChat: +86 158 1818 0019',
  bank: 'SWIFT: VTBRCNSH    Bank: Foreign Trade Bank of Russia Shanghai Branch    CNY A/C: 40807156700610047967',
  contractNo: '',
  contractDate: ''
};

const DEFAULT_PAYMENT_TERMS =
  'A deposit of 20% of the invoice amount is paid to the supplier 45 days after the order is placed in production. The final balance of 80% is paid to the supplier within 45 days from the date of shipment of the goods from the factory.';

const SUMMARY_SHEET_NAME = 'Сводная';
const SUMMARY_LINK_HEADER = 'Ссылка на инвойс';

const SELLER_EDITOR_EMAIL = 'banych83@gmail.com';

const INVOICE_FORM_PROFORMA = 'proforma';
const INVOICE_FORM_COMMERCIAL = 'commercial';

const INVOICE_FORMS = {
  [INVOICE_FORM_PROFORMA]: {
    key: INVOICE_FORM_PROFORMA,
    label: 'Proforma Invoice',
    heading: 'PROFORMA INVOICE',
    plLabel: 'Proforma Invoice No'
  },
  [INVOICE_FORM_COMMERCIAL]: {
    key: INVOICE_FORM_COMMERCIAL,
    label: 'Commercial Invoice',
    heading: 'COMMERCIAL INVOICE',
    plLabel: 'Commercial Invoice No'
  }
};

function invResolveInvoiceForm_(formKey) {
  const raw = String(formKey == null ? '' : formKey).trim().toLowerCase();
  if (INVOICE_FORMS[raw]) return INVOICE_FORMS[raw];
  return INVOICE_FORMS[INVOICE_FORM_PROFORMA];
}

function invLog_(level, message, context) {
  const tail = context != null ? ' | ' + JSON.stringify(context) : '';
  Logger.log(`[INVOICE][${level}] ${message}${tail}`);
}

function invThrowUser_(msg) {
  invLog_('ERROR', msg, {});
  throw new Error(msg);
}

function invGetProp_(key, defaultValue) {
  if (typeof getScriptProp === 'function') return getScriptProp(key, defaultValue);
  const v = PropertiesService.getScriptProperties().getProperty(key);
  if (v == null || String(v).trim() === '') return defaultValue === undefined ? '' : defaultValue;
  return String(v).trim();
}

function invGetPropOrThrow_(key) {
  const v = invGetProp_(key, '');
  if (!v) invThrowUser_(`Не заполнен Script Property: ${key}`);
  return v;
}

function invCanon_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/ё/g, 'е');
}

function invCompanyKey_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/["'`«»„“”]/g, '')
    .replace(/[.,;:!?()[\]{}\/\\_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/\s/g, '');
}

function invNorm_(s) {
  return String(s || '')
    .replace(/\u00a0/g, ' ')
    .trim();
}

function invNormStatus_(s) {
  return invNorm_(s).replace(/\s+/g, ' ');
}

function invBuildHeaderIndex_(headers) {
  const map = {};
  for (let c = 0; c < headers.length; c++) {
    const key = invCanon_(headers[c]);
    if (key && map[key] == null) map[key] = c;
  }
  return map;
}

function invColByHeader_(map, variants) {
  for (let i = 0; i < variants.length; i++) {
    const k = invCanon_(variants[i]);
    if (map[k] != null) return map[k];
  }
  return null;
}

/**
 * Все известные написания заголовка колонки с номером спецификации
 * (в т.ч. с дополнениями «/Инвойса», «/Инвойс»). Используем во всех местах,
 * где нужно найти эту колонку — менеджерская вкладка и «Сводная».
 */
function invSpecHeaderVariants_() {
  return [
    'Номер спецификации',
    'Номер Спецификации',
    'Спецификация',
    'Номер спецификации/инвойса',
    'Номер Спецификации/Инвойса',
    'Номер спецификации/Инвойса',
    'Номер Спецификации/инвойса',
    'Спецификация/Инвойс',
    'Спецификация/инвойс',
    'Спецификация/Инвойса',
    'Спецификация/инвойса',
    'Спецификация / Инвойс',
    'Спецификация / инвойс'
  ];
}

/**
 * Доп. защита: иногда заголовок отличается мелочами (точка, «№», лишний
 * слэш и т.п.). Если явных вариантов не хватило — ищем колонку, у которой
 * канон-имя начинается на «номерспецификации» или «спецификация», но не
 * содержит «сумма»/«цена»/«статус» и т.п. (чтобы не утащить случайную
 * соседнюю колонку).
 */
function invFindSpecColLoose_(headers) {
  const blacklist = ['сумма', 'цена', 'статус', 'количество', 'артикул', 'дата', 'поставщик', 'объем', 'вес', 'ссылка', 'отгрузка'];
  for (let c = 0; c < headers.length; c++) {
    const k = invCanon_(headers[c]);
    if (!k) continue;
    const isSpec = k.indexOf('номерспецификации') === 0 || k.indexOf('спецификации') === 0 || k.indexOf('спецификация') === 0;
    if (!isSpec) continue;
    let bad = false;
    for (let b = 0; b < blacklist.length; b++) {
      if (k.indexOf(blacklist[b]) >= 0) { bad = true; break; }
    }
    if (!bad) return c;
  }
  return null;
}

function invRequiredStatus_() {
  return invNormStatus_(invGetProp_('INVOICE_REQUIRED_STATUS', INV_DEFAULT_STATUS));
}

function invManagerLinkCol_() {
  const n = parseInt(invGetProp_('INVOICE_MANAGER_LINK_COL_1BASED', '27'), 10);
  return n > 0 ? n : 27;
}

function invSummaryLinkCol_() {
  const n = parseInt(invGetProp_('INVOICE_SUMMARY_LINK_COL_1BASED', '4'), 10);
  return n > 0 ? n : 4;
}

function invFallbackStatusCol_() {
  return parseInt(invGetProp_('INVOICE_STATUS_COL_1BASED', '8'), 10) || 8;
}

function invFallbackSpecCol_() {
  return parseInt(invGetProp_('INVOICE_SPEC_COL_1BASED', '12'), 10) || 12;
}

function invIsLikelyUrl_(v) {
  const s = invNorm_(v);
  return /^https?:\/\//i.test(s) || /^https?:\/\//i.test(String(v));
}

function invCurrentUserEmail_() {
  try {
    const active = String((Session.getActiveUser && Session.getActiveUser().getEmail()) || '').trim();
    if (active) return active.toLowerCase();
  } catch (e) {}
  try {
    const eff = String((Session.getEffectiveUser && Session.getEffectiveUser().getEmail()) || '').trim();
    if (eff) return eff.toLowerCase();
  } catch (e) {}
  return '';
}

/** Доступна google.script.run — используется HTML-формой для скрытия/показа полей. */
function invCanEditSellers() {
  return invCurrentUserEmail_() === String(SELLER_EDITOR_EMAIL).toLowerCase();
}

function invRequireSellerEditor_() {
  if (!invCanEditSellers()) {
    throw new Error(
      'Редактирование справочника поставщиков разрешено только пользователю ' +
      SELLER_EDITOR_EMAIL + '. Текущий пользователь: ' + (invCurrentUserEmail_() || 'неизвестен')
    );
  }
}

/**
 * Приводит дату контракта к формату dd.MM.yyyy для отображения в инвойсе.
 * Поддерживает входные форматы: yyyy-MM-dd (HTML input type=date), dd.MM.yyyy,
 * dd/MM/yyyy, объект Date. Если формат не распознан — возвращает строку как есть.
 */
function invFormatContractDate_(v) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || 'GMT+8', 'dd.MM.yyyy');
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return m[3] + '.' + m[2] + '.' + m[1];
  m = s.match(/^(\d{2})[.\/-](\d{2})[.\/-](\d{4})$/);
  if (m) return m[1] + '.' + m[2] + '.' + m[3];
  return s;
}

/** Второе меню: вызывайте из onOpen — addSupplierInvoiceMenu_(ui); */
function addSupplierInvoiceMenu_(ui) {
  ui.createMenu('📄 Инвойсы поставщику')
    .addItem('Сформировать инвойс по спецификации…', 'invStartWizard')
    .addItem('Сформировать общий инвойс (несколько спецификаций)…', 'invStartCombinedWizard')
    .addSeparator()
    .addItem('Справочник: компании (продавцы)', 'invShowSellersSettings')
    .addItem('Справочник: условия оплаты', 'invShowPaymentTermsSettings')
    .addToUi();
}

function invStartWizard() {
  const ui = SpreadsheetApp.getUi();
  try {
    invGetPropOrThrow_('INVOICE_DRIVE_FOLDER_ID');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    if (!sheet) {
      ui.alert('Нет активного листа.');
      return;
    }
    if (invNorm_(sheet.getName()) === invNorm_(SUMMARY_SHEET_NAME)) {
      ui.alert('Откройте вкладку менеджера (не «Сводная») и повторите.');
      return;
    }

    const lastCol = Math.max(sheet.getLastColumn(), invFallbackSpecCol_(), invManagerLinkCol_());
    const lastRow = sheet.getLastRow();
    if (lastRow < INV_DATA_START_ROW) {
      ui.alert('Нет данных: заполните строки начиная с ' + INV_DATA_START_ROW + '.');
      return;
    }

    const headerRange = sheet.getRange(INV_HEADER_ROW, 1, INV_HEADER_ROW, lastCol);
    const headers = headerRange.getValues()[0];
    const hmap = invBuildHeaderIndex_(headers);

    const statusCol0 =
      invColByHeader_(hmap, ['Статус заказа', 'Статус', 'STATUS']) ??
      invFallbackStatusCol_() - 1;
    const specCol0 =
      invColByHeader_(hmap, invSpecHeaderVariants_()) ??
      invFindSpecColLoose_(headers) ??
      invFallbackSpecCol_() - 1;

    const need = invRequiredStatus_();
    const specsSet = {};
    const values = sheet.getRange(INV_DATA_START_ROW, 1, lastRow, lastCol).getValues();
    for (let r = 0; r < values.length; r++) {
      const row = values[r];
      const st = invNormStatus_(row[statusCol0]);
      if (st !== need) continue;
      const sp = invNorm_(row[specCol0]);
      if (sp) specsSet[sp] = true;
    }

    const specs = Object.keys(specsSet).sort();
    if (!specs.length) {
      ui.alert(
        `Не найдено строк со статусом «${need}» и заполненным номером спецификации на листе «${sheet.getName()}».`
      );
      return;
    }

    const paymentTerms = getSavedTerms();
    const html = HtmlService.createHtmlOutput(invBuildSpecDialogHtml_(specs, paymentTerms))
      .setWidth(440)
      .setHeight(390);
    SpreadsheetApp.getUi().showModalDialog(html, 'Выбор спецификации');
  } catch (e) {
    invLog_('ERROR', 'invStartWizard_', { error: String(e && e.message ? e.message : e) });
    ui.alert('Ошибка', String(e && e.message ? e.message : e), ui.ButtonSet.OK);
  }
}

// Совместимость со старым именем пункта меню.
function invStartWizard_() {
  return invStartWizard();
}

/**
 * Запуск мастера «общий инвойс»: один файл на несколько спецификаций
 * одной и той же компании из колонки «Отгрузка через».
 */
function invStartCombinedWizard() {
  const ui = SpreadsheetApp.getUi();
  try {
    invGetPropOrThrow_('INVOICE_DRIVE_FOLDER_ID');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    if (!sheet) { ui.alert('Нет активного листа.'); return; }
    if (invNorm_(sheet.getName()) === invNorm_(SUMMARY_SHEET_NAME)) {
      ui.alert('Откройте вкладку менеджера (не «Сводная») и повторите.');
      return;
    }

    const data = invCollectShipViaSpecsForSheet_(sheet);
    if (!data.companies.length) {
      ui.alert(
        'Не найдено строк со статусом «' + invRequiredStatus_() +
        '» и одновременно заполненными колонками «Отгрузка через» и «Номер спецификации» на листе «' +
        sheet.getName() + '».'
      );
      return;
    }

    const paymentTerms = getSavedTerms();
    const html = HtmlService.createHtmlOutput(invBuildCombinedDialogHtml_(data, paymentTerms))
      .setWidth(560)
      .setHeight(580);
    ui.showModalDialog(html, 'Общий инвойс на несколько спецификаций');
  } catch (e) {
    invLog_('ERROR', 'invStartCombinedWizard', { error: String(e && e.message ? e.message : e) });
    ui.alert('Ошибка', String(e && e.message ? e.message : e), ui.ButtonSet.OK);
  }
}

/**
 * Возвращает компании из колонки «Отгрузка через» и их спецификации со статусом
 * «Сформировать инвойс» для активной вкладки. Считает строки и сумму по каждой
 * спецификации — нужно для пометки в диалоге.
 */
function invCollectShipViaSpecsForSheet_(sheet) {
  const result = { companies: [] };
  const lastCol = Math.max(sheet.getLastColumn(), invFallbackSpecCol_(), invManagerLinkCol_());
  const lastRow = sheet.getLastRow();
  if (lastRow < INV_DATA_START_ROW) return result;

  const headers = sheet.getRange(INV_HEADER_ROW, 1, INV_HEADER_ROW, lastCol).getValues()[0];
  const hmap = invBuildHeaderIndex_(headers);

  const statusCol0 =
    invColByHeader_(hmap, ['Статус заказа', 'Статус', 'STATUS']) ?? invFallbackStatusCol_() - 1;
  const specCol0 =
    invColByHeader_(hmap, invSpecHeaderVariants_()) ??
    invFindSpecColLoose_(headers) ??
    invFallbackSpecCol_() - 1;
  const shipViaCol0 = invColByHeader_(hmap, ['Отгрузка через']);
  const qtyCol0 = invColByHeader_(hmap, ['Итоговое количество', 'Количество']);
  const priceCol0 = invColByHeader_(hmap, ['Цена']);
  const amountCol0 = invColByHeader_(hmap, ['Сумма']);

  if (shipViaCol0 == null) {
    throw new Error('На активной вкладке не найдена колонка «Отгрузка через».');
  }

  const need = invRequiredStatus_();
  const data = sheet.getRange(INV_DATA_START_ROW, 1, lastRow, lastCol).getValues();
  const map = {};
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (invNormStatus_(row[statusCol0]) !== need) continue;
    const ship = invNorm_(row[shipViaCol0]);
    const spec = invNorm_(row[specCol0]);
    if (!ship || !spec) continue;
    const qty = qtyCol0 != null ? invParseNumber_(row[qtyCol0]) : NaN;
    const price = priceCol0 != null ? invParseNumber_(row[priceCol0]) : NaN;
    let amount = amountCol0 != null ? invParseNumber_(row[amountCol0]) : NaN;
    if (!isFinite(amount) && isFinite(qty) && isFinite(price)) amount = qty * price;
    if (!isFinite(amount)) amount = 0;

    if (!map[ship]) map[ship] = {};
    if (!map[ship][spec]) map[ship][spec] = { linesCount: 0, amount: 0 };
    map[ship][spec].linesCount++;
    map[ship][spec].amount += amount;
  }

  Object.keys(map).sort().forEach(function (cn) {
    const specs = Object.keys(map[cn]).sort().map(function (sp) {
      return {
        number: sp,
        linesCount: map[cn][sp].linesCount,
        amount: Math.round(map[cn][sp].amount * 100) / 100
      };
    });
    result.companies.push({ company: cn, specs: specs });
  });
  return result;
}

function invBuildCombinedDialogHtml_(data, paymentTerms) {
  const terms = Array.isArray(paymentTerms) && paymentTerms.length
    ? paymentTerms
    : [DEFAULT_PAYMENT_TERMS];
  const safeJson = JSON.stringify(data).replace(/</g, '\\u003c').replace(/>/g, '\\u003e');
  const termsOptions = terms.map(function (t, i) {
    const shortLabel = String(t || '').length > 90 ? String(t).slice(0, 90) + '...' : String(t || '');
    return '<option value="' + i + '">' + invEscapeHtml_(shortLabel) + '</option>';
  }).join('');

  return [
    '<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8">',
    '<style>',
    'body{font-family:Arial,sans-serif;padding:12px;margin:0}',
    'label{font-size:12px;color:#444}',
    'select,input[type=text]{width:100%;box-sizing:border-box;margin:4px 0 8px}',
    '.btn{margin-top:4px;padding:7px 14px}',
    '.specsBox{margin:6px 0 8px;border:1px solid #ddd;border-radius:4px;padding:8px;max-height:210px;overflow:auto;background:#fafafa}',
    '.specRow{display:block;padding:4px 0;border-bottom:1px dashed #e0e0e0;font-size:13px}',
    '.specRow:last-child{border-bottom:none}',
    '.specRow input{margin-right:8px;vertical-align:middle}',
    '.meta{color:#666;font-size:11px;margin-left:4px}',
    '.bar{display:flex;justify-content:space-between;align-items:center;margin:4px 0}',
    '.barInfo{font-size:12px;color:#444}',
    '.empty{color:#888;font-size:12px;padding:8px}',
    '.hint{font-size:11px;color:#666;margin:-4px 0 6px}',
    '</style></head><body>',

    '<label>Форма документа</label>',
    '<select id="form">',
    '<option value="', INVOICE_FORM_PROFORMA, '" selected>Proforma Invoice (предварительный)</option>',
    '<option value="', INVOICE_FORM_COMMERCIAL, '">Commercial Invoice (итоговый)</option>',
    '</select>',

    '<label>Отгрузка через</label>',
    '<select id="company" onchange="renderSpecs()"></select>',
    '<div class="hint">Список спецификаций обновится автоматически при смене компании.</div>',

    '<div class="bar">',
    '<div class="barInfo" id="selSummary">Выбрано: 0 из 0</div>',
    '<div>',
    '<button type="button" class="btn" onclick="selectAll(true)">Выбрать все</button> ',
    '<button type="button" class="btn" onclick="selectAll(false)">Снять все</button>',
    '</div></div>',
    '<div id="specsBox" class="specsBox"></div>',

    '<label>Вариант оплаты</label>',
    '<select id="term">', termsOptions, '</select>',

    '<div style="margin-top:10px;font-size:12px;color:#555">Будет создан один файл, объединяющий выбранные спецификации. Номера будут перечислены через «/».</div>',

    '<button class="btn" type="button" id="goBtn" onclick="go()" disabled>Сформировать</button> ',
    '<button class="btn" type="button" onclick="google.script.host.close()">Отмена</button>',

    '<script>',
    'var DATA=', safeJson, ';',
    'function esc(s){return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");}',
    'function fmtNum(n){if(!isFinite(n))return "0.00";return Number(n).toLocaleString("ru-RU",{minimumFractionDigits:2,maximumFractionDigits:2});}',
    'function getCompany(){var sel=document.getElementById("company");return DATA.companies.filter(function(c){return c.company===sel.value;})[0];}',
    'function renderSpecs(){var box=document.getElementById("specsBox");var c=getCompany();',
    'if(!c){box.innerHTML="<div class=\\"empty\\">(нет данных)</div>";updateSummary();return;}',
    'box.innerHTML=c.specs.map(function(s){return "<label class=\\"specRow\\"><input type=\\"checkbox\\" class=\\"spcChk\\" value=\\""+esc(s.number)+"\\" onchange=\\"updateSummary()\\">"+esc(s.number)+" <span class=\\"meta\\">— "+s.linesCount+" поз., "+fmtNum(s.amount)+" CNY</span></label>";}).join("");',
    'updateSummary();}',
    'function getChecked(){return Array.prototype.slice.call(document.querySelectorAll(".spcChk")).filter(function(x){return x.checked;}).map(function(x){return x.value;});}',
    'function updateSummary(){var total=document.querySelectorAll(".spcChk").length;var checked=getChecked().length;document.getElementById("selSummary").textContent="Выбрано: "+checked+" из "+total;document.getElementById("goBtn").disabled=(checked===0);}',
    'function selectAll(on){Array.prototype.slice.call(document.querySelectorAll(".spcChk")).forEach(function(x){x.checked=!!on;});updateSummary();}',
    'function go(){var f=document.getElementById("form").value;var t=Number(document.getElementById("term").value);var c=document.getElementById("company").value;var specs=getChecked();if(!specs.length){alert("Выберите хотя бы одну спецификацию.");return;}',
    'var btns=document.querySelectorAll("button");btns.forEach(function(b){b.disabled=true;});',
    'google.script.run.withSuccessHandler(function(){google.script.host.close();})',
    '.withFailureHandler(function(e){btns.forEach(function(b){b.disabled=false;});updateSummary();alert(e.message||e);})',
    '.invGenerateForSpecsChosen(specs, t, f, c);}',
    'function init(){var sel=document.getElementById("company");sel.innerHTML=DATA.companies.map(function(c){return "<option value=\\""+esc(c.company)+"\\">"+esc(c.company)+" — "+c.specs.length+" спец.</option>";}).join("");renderSpecs();}',
    'init();',
    '</script></body></html>'
  ].join('');
}

function invBuildSpecDialogHtml_(specs, paymentTerms) {
  const terms = Array.isArray(paymentTerms) && paymentTerms.length
    ? paymentTerms
    : [DEFAULT_PAYMENT_TERMS];
  return (
    '<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8">' +
    '<style>body{font-family:Arial,sans-serif;padding:12px}select,input{width:100%;box-sizing:border-box;margin:6px 0}' +
    'label{font-size:12px;color:#444}.btn{margin-top:10px;padding:8px 14px}' +
    '.hint{margin-top:4px;font-size:11px;color:#666}</style></head><body>' +
    '<label>Форма документа</label>' +
    '<select id="form">' +
    '<option value="' + INVOICE_FORM_PROFORMA + '" selected>Proforma Invoice (предварительный)</option>' +
    '<option value="' + INVOICE_FORM_COMMERCIAL + '">Commercial Invoice (итоговый)</option>' +
    '</select>' +
    '<div class="hint">Влияет только на название документа в шапке и имени файла.</div>' +
    '<label style="margin-top:10px;display:block">Номер спецификации</label>' +
    '<select id="spec">' +
    specs.map(function (s) {
      return '<option value="' + invEscapeHtml_(s) + '">' + invEscapeHtml_(s) + '</option>';
    }).join('') +
    '</select>' +
    '<label style="margin-top:10px;display:block">Вариант оплаты</label>' +
    '<select id="term">' +
    terms.map(function (t, i) {
      const shortLabel = String(t || '').length > 90 ? String(t).slice(0, 90) + '...' : String(t || '');
      return '<option value="' + i + '">' + invEscapeHtml_(shortLabel) + '</option>';
    }).join('') +
    '</select>' +
    '<div style="margin-top:12px;font-size:12px;color:#555">После создания файла ссылки запишутся в колонку ' +
    invColLetter_(invManagerLinkCol_()) +
    ' этой вкладки и в колонку ' +
    invColLetter_(invSummaryLinkCol_()) +
    ' на «Сводной» (если найдены совпадающие строки).</div>' +
    '<button class="btn" onclick="go()">Сформировать</button> ' +
    '<button class="btn" onclick="google.script.host.close()">Отмена</button>' +
    '<script>function go(){var s=document.getElementById("spec").value;var t=document.getElementById("term").value;' +
    'var f=document.getElementById("form").value;' +
    'var btns=document.querySelectorAll(\"button\");btns.forEach(function(b){b.disabled=true;});' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close();})' +
    '.withFailureHandler(function(e){btns.forEach(function(b){b.disabled=false;});alert(e.message||e);})' +
    '.invGenerateForSpecChosen(s, Number(t), f);}</script></body></html>'
  );
}

function invEscapeHtml_(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function invColLetter_(n) {
  let s = '';
  let c = n;
  while (c > 0) {
    const m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s || '?';
}

/**
 * Старая точка вызова из диалога одной спецификации. Делегирует в новый поток.
 * @param {string} specNumber
 * @param {number=} paymentTermIndex
 * @param {string=} invoiceFormKey  'proforma' (по умолчанию) или 'commercial'
 */
function invGenerateForSpecChosen(specNumber, paymentTermIndex, invoiceFormKey) {
  return invGenerateForSpecsChosen([specNumber], paymentTermIndex, invoiceFormKey, '');
}

/**
 * Создаёт один файл (Proforma/Commercial Invoice + Packing List) на одну
 * или несколько спецификаций, относящихся к одной компании из колонки
 * «Отгрузка через». Проставляет ссылки во все строки этих спецификаций
 * на активной вкладке и на «Сводной».
 *
 * @param {string[]} specNumbers   Список номеров спецификаций.
 * @param {number=}  paymentTermIndex
 * @param {string=}  invoiceFormKey  'proforma' (по умолчанию) или 'commercial'
 * @param {string=}  shipViaCompanyHint  Имя компании из диалога «общий инвойс»
 *                                       (для подсказки/совместимости — не критично).
 */
function invGenerateForSpecsChosen(specNumbers, paymentTermIndex, invoiceFormKey, shipViaCompanyHint) {
  const ui = SpreadsheetApp.getUi();

  const rawList = Array.isArray(specNumbers) ? specNumbers : [specNumbers];
  const normSpecs = [];
  const seen = {};
  for (let i = 0; i < rawList.length; i++) {
    const s = invNorm_(rawList[i]);
    if (!s || seen[s]) continue;
    seen[s] = true;
    normSpecs.push(s);
  }
  if (!normSpecs.length) {
    throw new Error('Не выбраны спецификации.');
  }

  const invoiceForm = invResolveInvoiceForm_(invoiceFormKey);

  const folderId = invGetPropOrThrow_('INVOICE_DRIVE_FOLDER_ID');
  const folder = DriveApp.getFolderById(folderId);

  const activeSs = SpreadsheetApp.getActiveSpreadsheet();
  const managerSheet = activeSs.getActiveSheet();
  if (!managerSheet || invNorm_(managerSheet.getName()) === invNorm_(SUMMARY_SHEET_NAME)) {
    throw new Error('Активным должен быть лист менеджера.');
  }

  const lastCol = Math.max(managerSheet.getLastColumn(), invFallbackSpecCol_(), invManagerLinkCol_());
  const lastRow = managerSheet.getLastRow();
  const headers = managerSheet.getRange(INV_HEADER_ROW, 1, INV_HEADER_ROW, lastCol).getValues()[0];
  const hmap = invBuildHeaderIndex_(headers);

  const statusCol0 =
    invColByHeader_(hmap, ['Статус заказа', 'Статус', 'STATUS']) ?? invFallbackStatusCol_() - 1;
  const specCol0 =
    invColByHeader_(hmap, invSpecHeaderVariants_()) ??
    invFindSpecColLoose_(headers) ??
    invFallbackSpecCol_() - 1;
  const wbCol0 = invColByHeader_(hmap, ['Артикул ВБ', 'Артикул WB', 'WB']);
  const supplierArticleCol0 = invColByHeader_(hmap, ['Артикул поставщика', 'Supplier article']);
  const barcodeCol0 = invColByHeader_(hmap, ['ШК', 'Barcode']);
  const qtyCol0 = invColByHeader_(hmap, ['Итоговое количество', 'Количество']);
  const qtyPerBoxCol0 = invColByHeader_(hmap, ['Количество в коробке', 'Pcs/Carton', 'Коробок в упаковке']);
  const priceCol0 = invColByHeader_(hmap, ['Цена']);
  const amountCol0 = invColByHeader_(hmap, ['Сумма']);
  const supCol0 = invColByHeader_(hmap, ['Поставщик']);
  const shipViaCol0 = invColByHeader_(hmap, ['Отгрузка через']);
  const volCol0 = invColByHeader_(hmap, ['Объем']);
  const weightCol0 = invColByHeader_(hmap, ['Вес']);

  const need = invRequiredStatus_();
  const specSet = {};
  normSpecs.forEach(function (s) { specSet[s] = true; });

  const data = managerSheet.getRange(INV_DATA_START_ROW, 1, lastRow, lastCol).getValues();

  const hitRows = [];
  const wbPairs = [];
  const shipViasSeen = {};
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (invNormStatus_(row[statusCol0]) !== need) continue;
    const sp = invNorm_(row[specCol0]);
    if (!sp || !specSet[sp]) continue;
    const absRow = INV_DATA_START_ROW + i;
    hitRows.push({ absRow: absRow, row: row, spec: sp });
    const wb = wbCol0 != null ? invNorm_(row[wbCol0]) : invNorm_(row[0]);
    if (wb) wbPairs.push({ spec: sp, wb: wb });
    const sv = shipViaCol0 != null ? invNorm_(row[shipViaCol0]) : '';
    if (sv) shipViasSeen[sv] = (shipViasSeen[sv] || 0) + 1;
  }

  if (!hitRows.length) {
    throw new Error('Не найдено строк по выбранным спецификациям и статусу.');
  }

  const foundSpecs = {};
  hitRows.forEach(function (h) { foundSpecs[h.spec] = true; });
  const missing = normSpecs.filter(function (s) { return !foundSpecs[s]; });
  if (missing.length) {
    throw new Error(
      'Для спецификаций ' + missing.join(', ') +
      ' не найдены строки со статусом «' + need + '».'
    );
  }

  const shipViaNames = Object.keys(shipViasSeen);
  if (shipViaNames.length > 1) {
    throw new Error(
      'В выбранных спецификациях разные компании в колонке «Отгрузка через»: ' +
      shipViaNames.join(', ') +
      '. Для общего инвойса все строки должны иметь одну и ту же компанию.'
    );
  }

  const shipViaValue = shipViaNames.length
    ? shipViaNames[0]
    : invNorm_(shipViaCompanyHint);
  const seller = invResolveSellerByShipViaStrict_(shipViaValue);
  if (!invNorm_(seller.contractNo) || !invNorm_(seller.contractDate)) {
    throw new Error(
      'У компании «' + seller.name + '» не заполнен номер и/или дата контракта. ' +
      'Откройте «📄 Инвойсы поставщику → Справочник: компании (продавцы)» и заполните реквизиты контракта.'
    );
  }

  const linkCol = invManagerLinkCol_();
  let hadLink = false;
  for (let j = 0; j < hitRows.length; j++) {
    if (invIsLikelyUrl_(hitRows[j].row[linkCol - 1])) {
      hadLink = true;
      break;
    }
  }
  if (hadLink) {
    const res = ui.alert(
      'Повторная генерация',
      'У части строк уже есть ссылка в колонке ' +
        invColLetter_(linkCol) +
        '. Создать новый файл и перезаписать ссылки?',
      ui.ButtonSet.YES_NO
    );
    if (res !== ui.Button.YES) return;
  }

  const terms = invGetPaymentTermByIndex_(paymentTermIndex);
  const invoiceNo = invNextInvoiceNumber_();

  const directoryMap = invGetProductDirectoryMap_();
  const lines = hitRows.map(function (h) {
    const row = h.row;
    const wb = wbCol0 != null ? invNorm_(row[wbCol0]) : '';
    const barcode = barcodeCol0 != null ? invNorm_(row[barcodeCol0]) : '';
    const name = supplierArticleCol0 != null ? invNorm_(row[supplierArticleCol0]) : wb;
    const dir = directoryMap[invCanon_(wb)] || directoryMap[invCanon_(name)] || null;
    const description = dir && dir.description
      ? dir.description
      : [name, 'Spec: ' + h.spec].filter(Boolean).join(' | ');
    const packaging = dir && dir.packaging ? dir.packaging : '';
    const qty = qtyCol0 != null ? invParseNumber_(row[qtyCol0]) : invParseNumber_(row[6]);
    const pcsPerCarton = qtyPerBoxCol0 != null ? invParseNumber_(row[qtyPerBoxCol0]) : NaN;
    const price = priceCol0 != null ? invParseNumber_(row[priceCol0]) : invParseNumber_(row[13]);
    let amount = amountCol0 != null ? invParseNumber_(row[amountCol0]) : NaN;
    if (!isFinite(amount) && isFinite(qty) && isFinite(price)) amount = qty * price;
    const cartons = isFinite(qty) && isFinite(pcsPerCarton) && pcsPerCarton > 0
      ? Math.ceil(qty / pcsPerCarton)
      : 0;
    const vol = volCol0 != null ? invParseNumber_(row[volCol0]) : NaN;
    const weight = weightCol0 != null ? invParseNumber_(row[weightCol0]) : NaN;
    const perCartonVol = cartons > 0 && isFinite(vol) ? vol / cartons : NaN;
    const perCartonWeight = cartons > 0 && isFinite(weight) ? weight / cartons : NaN;
    return {
      article: wb,
      name: name,
      description: description,
      photoUrl: dir && dir.photoUrl ? dir.photoUrl : '',
      barcode: barcode,
      qty: isFinite(qty) ? qty : '',
      cartons: cartons,
      pcsPerCarton: isFinite(pcsPerCarton) ? pcsPerCarton : '',
      price: isFinite(price) ? price : '',
      amount: isFinite(amount) ? amount : '',
      volume: isFinite(vol) ? vol : '',
      weight: isFinite(weight) ? weight : '',
      nwPerCarton: isFinite(perCartonWeight) ? perCartonWeight : '',
      gwPerCarton: isFinite(perCartonWeight) ? perCartonWeight : '',
      volumePerCarton: isFinite(perCartonVol) ? perCartonVol : '',
      packaging: packaging,
      spec: h.spec
    };
  });

  const supplierName = supCol0 != null ? invNorm_(hitRows[0].row[supCol0]) : '';

  const specsDisplay = normSpecs.join('/');
  const specsLabel = normSpecs.length > 1 ? ('Specs ' + specsDisplay) : ('Spec ' + specsDisplay);
  const titleBase =
    invoiceForm.label + ' ' + invoiceNo + ' — ' + specsLabel +
    (supplierName ? ' — ' + supplierName : '');
  invLog_('INFO', 'Создание файла', {
    title: titleBase,
    lines: lines.length,
    form: invoiceForm.key,
    specs: normSpecs.length
  });

  const newSs = invCreateInvoiceWorkbook_(
    titleBase, invoiceNo, specsDisplay, supplierName, seller, terms, lines, invoiceForm
  );
  const file = DriveApp.getFileById(newSs.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  const url = newSs.getUrl();

  for (let k = 0; k < hitRows.length; k++) {
    managerSheet.getRange(hitRows[k].absRow, linkCol).setValue(url);
  }

  invUpdateSummaryLinksMulti_(wbPairs, url);

  invLog_('INFO', 'Готово', { url: url });
  ui.alert('Готово', 'Файл создан:\n' + url, ui.ButtonSet.OK);
}

function invGetPaymentTermByIndex_(index) {
  const list = getSavedTerms();
  if (!list.length) return DEFAULT_PAYMENT_TERMS;
  const idx = Number(index);
  if (!isFinite(idx) || idx < 0 || idx >= list.length) return String(list[0]);
  return String(list[idx]);
}

function invParseNumber_(v) {
  if (v === '' || v == null) return NaN;
  if (typeof v === 'number' && isFinite(v)) return v;
  const s = String(v).replace(/\s/g, '').replace(',', '.');
  const n = parseFloat(s);
  return n;
}

function invGetProductDirectoryMap_() {
  const out = {};
  const refId = invGetProp_('PRODUCT_REFERENCE_SPREADSHEET_ID', '1PXWd05ENcZGvPYYbAVvf-1EPevdwkxr4IvjRbbojOlg');
  const sheetName = invGetProp_('PRODUCT_REFERENCE_SHEET_NAME', 'Справочник с названием товаров');

  try {
    const ss = SpreadsheetApp.openById(refId);
    const sh = ss.getSheetByName(sheetName);
    if (!sh) {
      invLog_('WARN', 'Справочник товаров не найден', { refId: refId, sheetName: sheetName });
      return out;
    }
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return out;
    // E: артикул, U: фото, Z: Description, AA: Packaging
    const values = sh.getRange(2, 1, lastRow - 1, 27).getValues();
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const article = invNorm_(row[4]);
      if (!article) continue;
      const key = invCanon_(article);
      out[key] = {
        photoUrl: invNorm_(row[20]),
        description: invNorm_(row[25]),
        packaging: invNorm_(row[26])
      };
    }
  } catch (e) {
    invLog_('WARN', 'Ошибка чтения справочника товаров', { error: String(e && e.message ? e.message : e) });
  }
  return out;
}

function invResolveSellerByShipViaStrict_(shipViaValue) {
  const sellers = getSavedSellers();
  if (!sellers.length) {
    throw new Error('Справочник компаний пуст. Добавьте хотя бы одну компанию.');
  }
  const raw = invNorm_(shipViaValue);
  if (!raw) {
    throw new Error('В колонке "Отгрузка через" пусто. Укажите компанию и повторите.');
  }

  const key = invCompanyKey_(raw);
  if (!key) {
    throw new Error('Не удалось распознать название в колонке "Отгрузка через".');
  }

  let exact = null;
  const fuzzyCandidates = [];
  for (let i = 0; i < sellers.length; i++) {
    const s = sellers[i];
    const sName = invNorm_(s && s.name);
    if (!sName) continue;
    const sKey = invCompanyKey_(sName);
    if (sKey === key) {
      exact = s;
      break;
    }
    if (sKey && (sKey.indexOf(key) >= 0 || key.indexOf(sKey) >= 0)) {
      fuzzyCandidates.push(s);
    }
  }

  if (!exact && fuzzyCandidates.length === 1) {
    exact = fuzzyCandidates[0];
  }

  if (!exact) {
    const sample = sellers
      .slice(0, 5)
      .map(function (s) { return s && s.name ? s.name : ''; })
      .filter(Boolean)
      .join(' | ');
    throw new Error(
      'Компания из колонки "Отгрузка через" не найдена в справочнике: "' +
        raw +
        '". Проверьте написание. Примеры из справочника: ' +
        sample
    );
  }
  invLog_('INFO', 'Выбор компании для инвойса', { shipVia: raw, seller: exact.name });
  return exact;
}

function invCreateInvoiceWorkbook_(title, invoiceNo, specsDisplay, buyerHint, seller, terms, lines, invoiceForm) {
  const form = invoiceForm && invoiceForm.heading ? invoiceForm : invResolveInvoiceForm_('');
  const ss = SpreadsheetApp.create(title);
  const invSh = ss.getSheets()[0];
  invSh.setName('Invoice');
  const plSh = ss.insertSheet('Packing List');
  const invoiceDate = Utilities.formatDate(new Date(), 'GMT+8', 'dd.MM.yyyy');

  const specsText = String(specsDisplay || '').trim();
  const specsCount = specsText ? specsText.split('/').filter(Boolean).length : 0;
  const specsHeaderLabel = specsCount > 1
    ? ('Specs: ' + specsText)
    : (specsText ? ('Spec: ' + specsText) : '');

  // INVOICE HEADER (Proforma / Commercial — отличие только в названии)
  invSh.getRange('A1:K1').mergeAcross().setValue(form.heading)
    .setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center');
  const headerLine2 = '№ ' + invoiceNo + '    Date: ' + invoiceDate +
    (specsHeaderLabel ? '    ' + specsHeaderLabel : '');
  invSh.getRange('A2:K2').mergeAcross().setValue(headerLine2)
    .setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center');

  const contractNo = invNorm_(seller && seller.contractNo);
  const contractDate = invFormatContractDate_(seller && seller.contractDate);
  const contractLine = contractNo
    ? 'Contract No: ' + contractNo + (contractDate ? '    dated ' + contractDate : '')
    : '';
  invSh.getRange('A3:K3').mergeAcross().setValue(contractLine)
    .setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center');

  invSh.getRange('A4').setValue('Seller / Shipper:').setFontWeight('bold');
  invSh.getRange('A5').setValue(seller.name).setFontWeight('bold');
  invSh.getRange('A6').setValue(seller.address).setWrap(false);
  invSh.getRange('A7').setValue(seller.contacts).setWrap(false);

  invSh.getRange('A9').setValue('Buyer / Consignee:').setFontWeight('bold');
  invSh.getRange('A10').setValue('PRIVATE ENTREPRENEUR EROKLINTSEV VLADIMIR VLADIMIROVICH').setFontWeight('bold');
  invSh.getRange('A11').setValue('Legal address: 670014, Russia, Republic of Buriatia, Ulan-Ude city, Peterbuzhskaya street, 62 building').setWrap(false);
  invSh.getRange('A12').setValue('TIN: 645325166024    PSRN: 321645100012882').setWrap(false);

  const invHeaders = ['№', 'Article', 'Photo', 'Name / Наименование', 'Description', 'Packaging', 'Barcode', 'Qty (pcs)', 'Cartons', 'Price (CNY)', 'Amount (CNY)'];
  invSh.getRange(14, 1, 1, 11).setValues([invHeaders])
    .setBackground('#0d47a1').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);

  let row = 15;
  lines.forEach(function (ln, idx) {
    invSh.getRange(row, 1, 1, 11).setValues([[
      idx + 1,
      ln.article || '',
      '',
      ln.name || '',
      ln.description || '',
      ln.packaging || '',
      ln.barcode || '',
      ln.qty || '',
      ln.cartons || '',
      ln.price || '',
      ln.amount || ''
    ]]);
    if (ln.photoUrl) {
      invSh.getRange(row, 12).setValue(ln.photoUrl);
      invSh.getRange(row, 3).setFormula('=IFERROR(IMAGE(L' + row + '); "")');
    }
    invSh.getRange(row, 4, 1, 3).setWrap(true);
    invSh.getRange(row, 7).setNumberFormat('@');
    invSh.getRange(row, 10, 1, 2).setNumberFormat('#,##0.00');
    invSh.setRowHeight(row, 90);
    row++;
  });

  const totalRow = row + 1;
  invSh.getRange(totalRow, 10).setValue('TOTAL:').setFontWeight('bold').setHorizontalAlignment('right');
  invSh.getRange(totalRow, 11).setFormula('=SUM(K15:K' + (row - 1) + ')').setFontWeight('bold').setNumberFormat('#,##0.00').setBackground('#c8e6c9');

  const termsRow = totalRow + 3;
  invSh.getRange(termsRow, 1, 1, 11).mergeAcross().setValue('Terms and Conditions')
    .setBackground('#0d47a1').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');

  const termsData = [
    ['Payment Terms:', terms],
    ['Delivery Terms:', 'FOB YIWU'],
    ['Country of Origin:', 'China'],
    ['Bank Details:', seller.bank || '']
  ];

  let currentTermRow = termsRow + 1;
  termsData.forEach(function (x) {
    invSh.getRange(currentTermRow, 1, 1, 3).mergeAcross().setValue(x[0])
      .setFontWeight('bold').setBackground('#f0f0f0').setVerticalAlignment('top');
    invSh.getRange(currentTermRow, 4, 1, 8).mergeAcross().setValue(x[1])
      .setWrap(true).setVerticalAlignment('top');
    currentTermRow++;
  });

  invSh.setColumnWidth(1, 40);
  invSh.setColumnWidth(2, 100);
  invSh.setColumnWidth(3, 110);
  invSh.setColumnWidth(4, 150);
  invSh.setColumnWidth(5, 150);
  invSh.setColumnWidth(6, 150);
  invSh.setColumnWidth(7, 110);
  invSh.setColumnWidths(8, 2, 70);
  invSh.setColumnWidths(10, 2, 85);
  invSh.getRange(14, 1, Math.max(1, row - 14), 11)
    .setBorder(true, true, true, true, true, true)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  invSh.hideColumns(12, 1);

  // PACKING LIST
  plSh.getRange('A1:J1').mergeAcross().setValue('PACKING LIST')
    .setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center');
  const plLine2 = form.plLabel + ': ' + invoiceNo + '     Date: ' + invoiceDate +
    (specsHeaderLabel ? '     ' + specsHeaderLabel : '');
  plSh.getRange('A2:J2').mergeAcross().setValue(plLine2)
    .setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center');
  plSh.getRange('A3:J3').mergeAcross().setValue(contractLine)
    .setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center');

  const plHeaders = ['№', 'Article', 'Name / Наименование', 'Description', 'Barcode', 'Qty (pcs)', 'Cartons', 'N.W.(kg)', 'G.W.(kg)', 'Volume(m³)'];
  plSh.getRange(4, 1, 1, 10).setValues([plHeaders])
    .setBackground('#1565c0').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);

  let plRow = 5;
  lines.forEach(function (ln, idx) {
    const cartons = ln.cartons || 0;
    const nw = cartons && ln.nwPerCarton ? cartons * ln.nwPerCarton : (ln.weight || 0);
    const gw = cartons && ln.gwPerCarton ? cartons * ln.gwPerCarton : (ln.weight || 0);
    const vol = cartons && ln.volumePerCarton ? cartons * ln.volumePerCarton : (ln.volume || 0);
    plSh.getRange(plRow, 1, 1, 10).setValues([[
      idx + 1,
      ln.article || '',
      ln.name || '',
      ln.description || '',
      ln.barcode || '',
      ln.qty || '',
      cartons,
      nw,
      gw,
      vol
    ]]);
    plSh.getRange(plRow, 3, 1, 2).setWrap(true);
    plSh.getRange(plRow, 5).setNumberFormat('@');
    plSh.getRange(plRow, 8, 1, 2).setNumberFormat('0.00');
    plSh.getRange(plRow, 10).setNumberFormat('0.000');
    plRow++;
  });

  const plTotalRow = plRow + 1;
  plSh.getRange(plTotalRow, 1).setValue('TOTAL').setFontWeight('bold').setBackground('#e8f5e9');
  plSh.getRange(plTotalRow, 6).setFormula('=SUM(F5:F' + (plRow - 1) + ')').setFontWeight('bold');
  plSh.getRange(plTotalRow, 7).setFormula('=SUM(G5:G' + (plRow - 1) + ')').setFontWeight('bold');
  plSh.getRange(plTotalRow, 8).setFormula('=SUM(H5:H' + (plRow - 1) + ')').setFontWeight('bold').setNumberFormat('0.00');
  plSh.getRange(plTotalRow, 9).setFormula('=SUM(I5:I' + (plRow - 1) + ')').setFontWeight('bold').setNumberFormat('0.00');
  plSh.getRange(plTotalRow, 10).setFormula('=SUM(J5:J' + (plRow - 1) + ')').setFontWeight('bold').setNumberFormat('0.000');

  plSh.setColumnWidth(1, 40);
  plSh.setColumnWidth(2, 100);
  plSh.setColumnWidth(3, 180);
  plSh.setColumnWidth(4, 210);
  plSh.setColumnWidth(5, 110);
  plSh.setColumnWidths(6, 2, 70);
  plSh.setColumnWidths(8, 3, 75);
  plSh.getRange(4, 1, Math.max(1, plRow - 4), 10)
    .setBorder(true, true, true, true, true, true)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');

  return ss;
}

/**
 * Обновляет колонку «Ссылка на инвойс» на «Сводной» для строк, попавших
 * в инвойс. wbPairs — массив { spec, wb }. Для каждой спецификации
 * обновляются только строки с подходящим артикулом ВБ; если у строки на
 * «Сводной» артикул пустой — она тоже обновляется, чтобы не оставлять
 * её без ссылки.
 */
function invUpdateSummaryLinksMulti_(wbPairs, url) {
  const mainId = invGetProp_('MAIN_SPREADSHEET_ID', '');
  let ss;
  try {
    ss = mainId ? SpreadsheetApp.openById(mainId) : SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    invLog_('WARN', 'Не удалось открыть MAIN_SPREADSHEET_ID', { mainId: mainId, err: String(e && e.message) });
    return;
  }

  const sh = ss.getSheetByName(SUMMARY_SHEET_NAME);
  if (!sh) {
    invLog_('WARN', 'Лист Сводная не найден', {});
    return;
  }

  const lastRow = sh.getLastRow();
  const lastCol = Math.max(sh.getLastColumn(), invSummaryLinkCol_(), 20);
  if (lastRow < INV_DATA_START_ROW) return;

  const headers = sh.getRange(INV_HEADER_ROW, 1, INV_HEADER_ROW, lastCol).getValues()[0];
  const hmap = invBuildHeaderIndex_(headers);
  const linkCol0 = invColByHeader_(hmap, [
    SUMMARY_LINK_HEADER,
    'Ссылка на инвойс',
    'Инвойс',
    'Invoice link'
  ]);
  const linkCol1 = linkCol0 != null ? linkCol0 + 1 : invSummaryLinkCol_();

  const specCol0 =
    invColByHeader_(hmap, invSpecHeaderVariants_()) ??
    invFindSpecColLoose_(headers) ??
    11;
  const wbCol0 = invColByHeader_(hmap, ['Артикул ВБ', 'Артикул WB', 'WB']) ?? 0;

  sh.getRange(INV_HEADER_ROW, linkCol1).setValue(SUMMARY_LINK_HEADER);

  // spec -> { wb -> true }
  const specMap = {};
  const pairs = Array.isArray(wbPairs) ? wbPairs : [];
  for (let i = 0; i < pairs.length; i++) {
    const sp = invNorm_(pairs[i] && pairs[i].spec);
    const wb = invNorm_(pairs[i] && pairs[i].wb);
    if (!sp) continue;
    if (!specMap[sp]) specMap[sp] = {};
    if (wb) specMap[sp][wb] = true;
  }
  if (!Object.keys(specMap).length) {
    invLog_('INFO', 'Сводная: нет пар spec/wb для обновления', {});
    return;
  }

  const data = sh.getRange(INV_DATA_START_ROW, 1, lastRow, lastCol).getValues();
  let updated = 0;
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowSpec = invNorm_(row[specCol0]);
    const allowedWbs = specMap[rowSpec];
    if (!allowedWbs) continue;
    const rowWb = invNorm_(row[wbCol0]);
    if (rowWb && Object.keys(allowedWbs).length > 0 && allowedWbs[rowWb] !== true) continue;
    sh.getRange(INV_DATA_START_ROW + i, linkCol1).setValue(url);
    updated++;
  }
  invLog_('INFO', 'Сводная обновлена', { updated: updated, linkCol: linkCol1, specs: Object.keys(specMap).length });
}

/**
 * Тонкая обёртка над `invUpdateSummaryLinksMulti_` для совместимости со
 * старыми вызовами по одной спецификации.
 */
function invUpdateSummaryLinks_(spec, wbArticles, url) {
  const pairs = (wbArticles || []).map(function (wb) {
    return { spec: spec, wb: wb };
  });
  return invUpdateSummaryLinksMulti_(pairs, url);
}

function invNextInvoiceNumber_() {
  const key = 'INVOICE_LAST_SERIAL';
  const sp = PropertiesService.getScriptProperties();
  const cur = parseInt(sp.getProperty(key) || '0', 10) || 0;
  const next = cur + 1;
  sp.setProperty(key, String(next));
  const prefix = invGetProp_('INVOICE_NUMBER_PREFIX', 'INV');
  return prefix + '-' + Utilities.formatString('%05d', next);
}

function invGetPrimarySeller_() {
  const list = getSavedSellers();
  return list.length ? list[0] : DEFAULT_SELLER;
}

function invGetPrimaryTerms_() {
  const list = getSavedTerms();
  return list.length ? String(list[0]) : DEFAULT_PAYMENT_TERMS;
}

function getSavedSellers() {
  const raw = PropertiesService.getDocumentProperties().getProperty('SAVED_SELLERS');
  if (!raw) return [DEFAULT_SELLER];
  try {
    const arr = JSON.parse(raw);
    const src = Array.isArray(arr) ? arr : [];
    const merged = [DEFAULT_SELLER];
    src.forEach(function (s) {
      if (!s || !invNorm_(s.name)) return;
      if (invCanon_(s.name) === invCanon_(DEFAULT_SELLER.name)) return;
      merged.push({
        name: invNorm_(s.name),
        address: invNorm_(s.address || ''),
        contacts: invNorm_(s.contacts || ''),
        bank: invNorm_(s.bank || ''),
        contractNo: invNorm_(s.contractNo || ''),
        contractDate: invNorm_(s.contractDate || '')
      });
    });
    return merged;
  } catch (e) {
    return [DEFAULT_SELLER];
  }
}

function addSavedSeller(name, address, contacts, bank, contractNo, contractDate) {
  invRequireSellerEditor_();
  const safeName = invNorm_(name);
  if (!safeName) throw new Error('Название компании обязательно.');

  const safeAddress = invNorm_(address);
  const safeContacts = invNorm_(contacts);
  const safeBank = invNorm_(bank);
  const safeContractNo = invNorm_(contractNo);
  const safeContractDate = invNorm_(contractDate);

  if (!safeContractNo) throw new Error('Номер контракта обязателен.');
  if (!safeContractDate) throw new Error('Дата контракта обязательна.');

  const list = getSavedSellers().filter(function (s) {
    return s && invNorm_(s.name);
  });

  const existingIdx = list.findIndex(function (s) {
    return invCanon_(s.name) === invCanon_(safeName);
  });

  if (existingIdx >= 0) {
    list[existingIdx] = {
      name: safeName,
      address: safeAddress,
      contacts: safeContacts,
      bank: safeBank,
      contractNo: safeContractNo,
      contractDate: safeContractDate
    };
  } else {
    list.push({
      name: safeName,
      address: safeAddress,
      contacts: safeContacts,
      bank: safeBank,
      contractNo: safeContractNo,
      contractDate: safeContractDate
    });
  }

  PropertiesService.getDocumentProperties().setProperty('SAVED_SELLERS', JSON.stringify(list));
  return list;
}

/**
 * Обновляет существующую запись по индексу (в порядке getSavedSellers()).
 * Допускает переименование. Запрещает редактирование базовой (DEFAULT_SELLER) —
 * она задана в коде.
 */
function updateSavedSeller(index, name, address, contacts, bank, contractNo, contractDate) {
  invRequireSellerEditor_();
  const all = getSavedSellers();
  const idx = Number(index);
  if (!isFinite(idx) || idx < 0 || idx >= all.length) {
    throw new Error('Неверный индекс компании.');
  }
  if (idx === 0 && invCanon_(all[0].name) === invCanon_(DEFAULT_SELLER.name)) {
    throw new Error('Базовую компанию редактировать нельзя — она задана в коде.');
  }

  const safeName = invNorm_(name);
  if (!safeName) throw new Error('Название компании обязательно.');
  const safeContractNo = invNorm_(contractNo);
  if (!safeContractNo) throw new Error('Номер контракта обязателен.');
  const safeContractDate = invNorm_(contractDate);
  if (!safeContractDate) throw new Error('Дата контракта обязательна.');

  for (let i = 0; i < all.length; i++) {
    if (i === idx) continue;
    if (invCanon_(all[i].name) === invCanon_(safeName)) {
      throw new Error('Компания с таким названием уже есть в справочнике.');
    }
  }

  const updated = {
    name: safeName,
    address: invNorm_(address || ''),
    contacts: invNorm_(contacts || ''),
    bank: invNorm_(bank || ''),
    contractNo: safeContractNo,
    contractDate: safeContractDate
  };

  const list = all.filter(function (s) {
    return s && invNorm_(s.name);
  });
  list[idx] = updated;

  PropertiesService.getDocumentProperties().setProperty('SAVED_SELLERS', JSON.stringify(list));
  return list;
}

function deleteSavedSeller(index) {
  invRequireSellerEditor_();
  const list = getSavedSellers().filter(function (s) {
    return s && invNorm_(s.name);
  });
  if (index < 0 || index >= list.length) return;
  if (invCanon_(list[index].name) === invCanon_(DEFAULT_SELLER.name)) {
    throw new Error('Базовую компанию удалить нельзя.');
  }
  list.splice(index, 1);
  PropertiesService.getDocumentProperties().setProperty('SAVED_SELLERS', JSON.stringify(list));
}

function invShowSellersSettings() {
  const html = HtmlService.createHtmlOutput(invSellersEditorHtml_()).setWidth(620).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Компании (продавцы)');
}

function invShowSellersSettings_() {
  return invShowSellersSettings();
}

function invSellersEditorHtml_() {
  const styles = [
    'body{font-family:Arial;padding:10px;margin:0}',
    'textarea,input,select{width:100%;box-sizing:border-box}',
    'label{font-size:12px;color:#444}',
    '.req{color:#c62828}',
    '.row2{display:flex;gap:8px}.row2>div{flex:1}',
    '.list{margin-top:10px;font-size:12px;color:#333;max-height:160px;overflow:auto;border:1px solid #ddd;padding:8px}',
    '.warn{color:#c62828}',
    '.editingBadge{margin-top:6px;padding:6px 8px;background:#fff3cd;border:1px solid #ffeeba;color:#856404;font-size:12px;border-radius:4px}',
    '.readonlyHint{margin:6px 0 10px;padding:8px;background:#e8f0fe;border:1px solid #c2d4f9;color:#1a3e8c;font-size:12px;border-radius:4px}',
    '.actions{margin-top:8px;display:flex;flex-wrap:wrap;gap:6px}',
    '.actions button{padding:6px 12px}',
    'h3{margin:14px 0 6px;font-size:13px;color:#333}'
  ].join('');

  const html = [
    '<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8">',
    '<style>', styles, '</style></head><body>',

    '<div id="readonlyHint" class="readonlyHint" style="display:none">',
      'Режим просмотра. Редактирование справочника доступно только пользователю ',
      '<b>', SELLER_EDITOR_EMAIL, '</b>. Вы вошли как <span id="curEmail">—</span>.',
    '</div>',

    '<div id="formBox" style="display:none">',
      '<p style="margin:4px 0 8px">Карточка продавца. Поля со <span class="req">*</span> обязательны.</p>',
      '<div id="editingBadge" class="editingBadge" style="display:none">Редактируется: <b><span id="editingName"></span></b></div>',
      '<label>Название <span class="req">*</span></label><input id="n">',
      '<label>Адрес</label><textarea id="a" rows="2"></textarea>',
      '<label>Контакты</label><textarea id="c" rows="2"></textarea>',
      '<label>Банк</label><textarea id="b" rows="2"></textarea>',
      '<div class="row2" style="margin-top:6px">',
        '<div><label>Номер контракта <span class="req">*</span></label><input id="cno" placeholder="например, BD-2025-001"></div>',
        '<div><label>Дата контракта <span class="req">*</span></label><input id="cdt" type="date"></div>',
      '</div>',
      '<div class="actions">',
        '<button id="saveBtn" onclick="save()">Добавить</button>',
        '<button id="cancelEditBtn" onclick="cancelEdit()" style="display:none">Отмена редактирования</button>',
        '<button onclick="resetForm()">Очистить форму</button>',
      '</div>',
    '</div>',

    '<h3>Список компаний</h3>',
    '<div id="lst" class="list">Загрузка...</div>',

    '<div id="manageBox" style="display:none">',
      '<div class="row2" style="margin-top:8px">',
        '<div><label>Редактировать</label><select id="editSel" onchange="onPickEdit()"></select></div>',
        '<div><label>Удалить</label><select id="delSeller"></select></div>',
      '</div>',
      '<div class="actions">',
        '<button onclick="delSellerClick()">Удалить выбранную</button>',
      '</div>',
    '</div>',

    '<div class="actions" style="margin-top:12px">',
      '<button onclick="google.script.host.close()">Закрыть</button>',
    '</div>',

    '<script>',
      'var _sellers=[];var _canEdit=false;var _editingIdx=null;',

      'function esc(s){return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");}',
      'function fmtDate(s){s=String(s||"").trim();var m=s.match(/^(\\d{4})-(\\d{2})-(\\d{2})$/);if(m)return m[3]+"."+m[2]+"."+m[1];return s;}',
      'function isoDate(s){s=String(s||"").trim();var m=s.match(/^(\\d{2})[.\\/-](\\d{2})[.\\/-](\\d{4})$/);if(m)return m[3]+"-"+m[2]+"-"+m[1];if(/^\\d{4}-\\d{2}-\\d{2}$/.test(s))return s;return "";}',

      'function showReadonly(email){document.getElementById("curEmail").textContent=email||"неизвестен";document.getElementById("readonlyHint").style.display="block";document.getElementById("formBox").style.display="none";document.getElementById("manageBox").style.display="none";}',
      'function showEditor(){document.getElementById("readonlyHint").style.display="none";document.getElementById("formBox").style.display="block";document.getElementById("manageBox").style.display="block";}',

      'function renderList(){',
        'var lst=document.getElementById("lst");',
        'if(!_sellers.length){lst.innerHTML="(пусто)";}',
        'else{lst.innerHTML=_sellers.map(function(x,i){',
          'var meta=[];if(x.contractNo)meta.push("контракт "+esc(x.contractNo));if(x.contractDate)meta.push("от "+esc(fmtDate(x.contractDate)));',
          'var miss=(!x.contractNo||!x.contractDate)?" <span class=\\"warn\\">(не заполнен контракт)</span>":"";',
          'return (i+1)+". "+esc(x.name||"")+(meta.length?" — "+meta.join(", "):"")+miss;',
        '}).join("<br>");}',
        'if(!_canEdit)return;',
        'var del=document.getElementById("delSeller");del.innerHTML="";',
        'var edt=document.getElementById("editSel");edt.innerHTML="<option value=\\"\\\">— выберите для редактирования —</option>";',
        '_sellers.forEach(function(x,i){',
          'var label=(i+1)+". "+(x.name||"");',
          'var od=document.createElement("option");od.value=String(i);od.text=label;del.appendChild(od);',
          'if(i>0){var oe=document.createElement("option");oe.value=String(i);oe.text=label;edt.appendChild(oe);}',
        '});',
        'if(_editingIdx!=null)document.getElementById("editSel").value=String(_editingIdx);',
      '}',

      'function reload(){google.script.run.withSuccessHandler(function(list){_sellers=list||[];renderList();}).withFailureHandler(function(e){document.getElementById("lst").innerText=(e.message||e);}).getSavedSellers();}',

      'function resetForm(){_editingIdx=null;["n","a","c","b","cno","cdt"].forEach(function(id){document.getElementById(id).value="";});document.getElementById("editingBadge").style.display="none";document.getElementById("saveBtn").textContent="Добавить";document.getElementById("cancelEditBtn").style.display="none";var es=document.getElementById("editSel");if(es)es.value="";}',
      'function cancelEdit(){resetForm();}',

      'function onPickEdit(){var v=document.getElementById("editSel").value;if(v===""){resetForm();return;}var idx=Number(v);var s=_sellers[idx];if(!s){resetForm();return;}_editingIdx=idx;document.getElementById("n").value=s.name||"";document.getElementById("a").value=s.address||"";document.getElementById("c").value=s.contacts||"";document.getElementById("b").value=s.bank||"";document.getElementById("cno").value=s.contractNo||"";document.getElementById("cdt").value=isoDate(s.contractDate);document.getElementById("editingName").textContent=s.name||"";document.getElementById("editingBadge").style.display="block";document.getElementById("saveBtn").textContent="Сохранить изменения";document.getElementById("cancelEditBtn").style.display="inline-block";}',

      'function readPayload(){return{name:document.getElementById("n").value,address:document.getElementById("a").value,contacts:document.getElementById("c").value,bank:document.getElementById("b").value,contractNo:document.getElementById("cno").value,contractDate:document.getElementById("cdt").value};}',

      'function save(){var p=readPayload();',
        'if(!p.name||!p.name.trim()){alert("Укажите название.");return;}',
        'if(!p.contractNo||!p.contractNo.trim()){alert("Укажите номер контракта.");return;}',
        'if(!p.contractDate||!p.contractDate.trim()){alert("Укажите дату контракта.");return;}',
        'var ok=function(msg){alert(msg||"Сохранено");resetForm();reload();};',
        'var fail=function(e){alert(e.message||e);};',
        'if(_editingIdx!=null){google.script.run.withSuccessHandler(ok).withFailureHandler(fail).invUpdateSellerUi(_editingIdx,p);}',
        'else{google.script.run.withSuccessHandler(ok).withFailureHandler(fail).invAddSellerUi(p);}',
      '}',

      'function delSellerClick(){var v=document.getElementById("delSeller").value;if(v===""){alert("Нечего удалять");return;}if(!confirm("Удалить выбранную компанию?")){return;}google.script.run.withSuccessHandler(function(msg){alert(msg||"Удалено");resetForm();reload();}).withFailureHandler(function(e){alert(e.message||e);}).invDeleteSellerUi(Number(v));}',

      'function init(){google.script.run.withSuccessHandler(function(can){_canEdit=!!can;if(_canEdit){showEditor();}else{google.script.run.withSuccessHandler(function(email){showReadonly(email);}).invCurrentUserEmailForUi();}reload();}).withFailureHandler(function(e){document.getElementById("lst").innerText=(e.message||e);}).invCanEditSellers();}',

      'init();',
    '</script></body></html>'
  ].join('');

  return html;
}

/** Возвращает текущий email активного пользователя — для отображения в UI. */
function invCurrentUserEmailForUi() {
  return invCurrentUserEmail_() || '';
}

function invAddSellerUi(payload) {
  invRequireSellerEditor_();
  if (!payload || !invNorm_(payload.name)) throw new Error('Укажите название.');
  if (!invNorm_(payload.contractNo)) throw new Error('Укажите номер контракта.');
  if (!invNorm_(payload.contractDate)) throw new Error('Укажите дату контракта.');
  const list = addSavedSeller(
    payload.name,
    payload.address || '',
    payload.contacts || '',
    payload.bank || '',
    payload.contractNo || '',
    payload.contractDate || ''
  );
  return 'Компания сохранена. Всего в справочнике: ' + list.length;
}

function invUpdateSellerUi(index, payload) {
  invRequireSellerEditor_();
  if (!payload || !invNorm_(payload.name)) throw new Error('Укажите название.');
  if (!invNorm_(payload.contractNo)) throw new Error('Укажите номер контракта.');
  if (!invNorm_(payload.contractDate)) throw new Error('Укажите дату контракта.');
  const list = updateSavedSeller(
    index,
    payload.name,
    payload.address || '',
    payload.contacts || '',
    payload.bank || '',
    payload.contractNo || '',
    payload.contractDate || ''
  );
  return 'Компания обновлена. Всего в справочнике: ' + list.length;
}

function invDeleteSellerUi(index) {
  invRequireSellerEditor_();
  const listBefore = getSavedSellers();
  if (!listBefore.length) return 'Список пуст.';
  deleteSavedSeller(index);
  const listAfter = getSavedSellers();
  return 'Компания удалена. Осталось: ' + listAfter.length;
}

function getSavedTerms() {
  const raw = PropertiesService.getDocumentProperties().getProperty('SAVED_PAYMENT_TERMS');
  if (!raw) return [DEFAULT_PAYMENT_TERMS];
  try {
    const arr = JSON.parse(raw);
    const src = Array.isArray(arr) ? arr : [];
    const unique = {};
    const merged = [DEFAULT_PAYMENT_TERMS];
    unique[invCanon_(DEFAULT_PAYMENT_TERMS)] = true;
    src.forEach(function (t) {
      const safe = invNorm_(t);
      if (!safe) return;
      const key = invCanon_(safe);
      if (unique[key]) return;
      unique[key] = true;
      merged.push(safe);
    });
    return merged;
  } catch (e) {
    return [DEFAULT_PAYMENT_TERMS];
  }
}

function addSavedTerm(term) {
  const safeTerm = invNorm_(term);
  if (!safeTerm) throw new Error('Пустой текст условий оплаты.');

  const list = getSavedTerms().map(function (t) {
    return invNorm_(t);
  });

  const existingIdx = list.findIndex(function (t) {
    return invCanon_(t) === invCanon_(safeTerm);
  });

  if (existingIdx >= 0) {
    list[existingIdx] = safeTerm;
  } else {
    list.push(safeTerm);
  }

  PropertiesService.getDocumentProperties().setProperty('SAVED_PAYMENT_TERMS', JSON.stringify(list));
  return list;
}

function deleteSavedTerm(index) {
  const list = getSavedTerms().map(function (t) {
    return invNorm_(t);
  });
  if (index < 0 || index >= list.length) return;
  if (invCanon_(list[index]) === invCanon_(DEFAULT_PAYMENT_TERMS)) {
    throw new Error('Базовые условия удалить нельзя.');
  }
  list.splice(index, 1);
  PropertiesService.getDocumentProperties().setProperty('SAVED_PAYMENT_TERMS', JSON.stringify(list));
}

function invShowPaymentTermsSettings() {
  const html = HtmlService.createHtmlOutput(invTermsEditorHtml_()).setWidth(520).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, 'Условия оплаты');
}

function invShowPaymentTermsSettings_() {
  return invShowPaymentTermsSettings();
}

function invTermsEditorHtml_() {
  return (
    '<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8">' +
    '<style>body{font-family:Arial;padding:10px}textarea,select{width:100%;box-sizing:border-box}textarea{height:140px}.list{margin-top:10px;font-size:12px;color:#333;max-height:120px;overflow:auto;border:1px solid #ddd;padding:8px}</style></head><body>' +
    '<label>Новый текст условий</label><textarea id="t"></textarea>' +
    '<div style="margin-top:8px;font-size:12px;color:#666">Текущий список условий:</div>' +
    '<div id="lst" class="list">Загрузка...</div>' +
    '<label style="margin-top:8px;display:block;">Удалить условие</label>' +
    '<select id="delTerm"></select>' +
    '<button onclick="save()">Добавить в справочник</button> <button onclick="delTerm()">Удалить выбранное</button> <button onclick="google.script.host.close()">Закрыть</button>' +
    '<script>' +
    'function esc(s){return String(s||\"\").replace(/&/g,\"&amp;\").replace(/</g,\"&lt;\").replace(/>/g,\"&gt;\");}' +
    'function shorten(s){s=String(s||\"\");return s.length>120?s.slice(0,120)+\"...\":s;}' +
    'function render(list){var sel=document.getElementById(\"delTerm\");sel.innerHTML=\"\";' +
    'if(!list||!list.length){document.getElementById(\"lst\").innerHTML=\"(пусто)\";sel.innerHTML=\"<option value=\\\"\\\">(пусто)</option>\";return;}' +
    'document.getElementById(\"lst\").innerHTML=list.map(function(x,i){return (i+1)+\". \"+esc(shorten(x));}).join(\"<br>\");' +
    'list.forEach(function(x,i){var o=document.createElement(\"option\");o.value=String(i);o.text=(i+1)+\". \"+shorten(x);sel.appendChild(o);});}' +
    'function reload(){google.script.run.withSuccessHandler(render).withFailureHandler(function(e){document.getElementById(\"lst\").innerText=(e.message||e);}).getSavedTerms();}' +
    'function save(){var v=document.getElementById("t").value;' +
    'google.script.run.withSuccessHandler(function(msg){alert(msg||\"Сохранено\");reload();document.getElementById(\"t\").value=\"\";})' +
    '.withFailureHandler(function(e){alert(e.message||e);}).invAddTermUi(v);}' +
    'function delTerm(){var v=document.getElementById(\"delTerm\").value;if(v===\"\"){alert(\"Нечего удалять\");return;}' +
    'if(!confirm(\"Удалить выбранное условие?\")){return;}' +
    'google.script.run.withSuccessHandler(function(msg){alert(msg||\"Удалено\");reload();})' +
    '.withFailureHandler(function(e){alert(e.message||e);}).invDeleteTermUi(Number(v));}' +
    'reload();' +
    '</script></body></html>'
  );
}

function invAddTermUi(term) {
  if (!invNorm_(term)) throw new Error('Пустой текст.');
  const list = addSavedTerm(invNorm_(term));
  return 'Условия сохранены. Всего в справочнике: ' + list.length;
}

function invDeleteTermUi(index) {
  const listBefore = getSavedTerms();
  if (!listBefore.length) return 'Список пуст.';
  deleteSavedTerm(index);
  const listAfter = getSavedTerms();
  return 'Условие удалено. Осталось: ' + listAfter.length;
}
