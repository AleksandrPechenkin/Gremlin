function updateExternalPurchases() {
  const ui = SpreadsheetApp.getUi();
  const externalSpreadsheetId = CONFIG.EXTERNAL_SPREADSHEET_ID;
  const externalSheetName = CONFIG.EXTERNAL_SHEET_NAME;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const srcSheet = ss.getActiveSheet();
    if (!srcSheet) {
      ui.alert('❌ Ошибка', 'Не найден активный лист-источник.', ui.ButtonSet.OK);
      return;
    }
    const srcLastRow = srcSheet.getLastRow();
    const srcLastCol = srcSheet.getLastColumn();
    if (srcLastRow < 3) {
      ui.alert('ℹ️ Инфо', 'На листе "Сводная" нет строк для синхронизации оплат.', ui.ButtonSet.OK);
      return;
    }
    const srcData = srcSheet.getRange(1, 1, srcLastRow, srcLastCol).getValues();
    const srcDisplay = srcSheet.getRange(1, 1, srcLastRow, srcLastCol).getDisplayValues();
    const hdr = srcDisplay[Math.min(1, srcDisplay.length - 1)] || []; // обычно строка 2

    // Канон, согласованный с syncManagerCanonHeader_ (main.gs): обрабатывает №, #, точки,
    // скобки, слэши, дефисы, подчёркивания, NBSP, ё→е. Без этого «Аванс_сумма», «№ спецификации»,
    // «Дата.факт.Аванс» молча уходили в цифровой фолбэк, а валидатор это пропускал.
    const canon = (s) => String(s == null ? '' : s)
      .toLowerCase()
      .replace(/ё/g, 'е')
      .replace(/\u00A0/g, ' ')
      .replace(/[№#]/g, ' ')
      .replace(/[.,:;()/\\\[\]{}'"“”«»]/g, ' ')
      .replace(/[-_]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    // Никакого цифрового фолбэка: не нашли — вернули -1, валидатор payValidateSourceSyncColumns_
    // покажет понятное сообщение со списком отсутствующих обязательных колонок.
    const findCol = (variants) => {
      const cv = variants.map(canon);
      for (let c = 0; c < hdr.length; c++) {
        const h = canon(hdr[c]);
        if (cv.indexOf(h) >= 0) return c;
      }
      return -1;
    };
    const findColOptional = findCol;
    const srcArticleCol = findCol(['Артикул ВБ', 'Артикул WB']);
    const srcSupplierCol = findColOptional(['Поставщик']);
    const srcSpecCol = findCol(['Номер спецификации', 'Номер Спецификации', '№ спецификации', 'Спецификация']);
    const srcAdvSumCol = findCol(['Аванс сумма', 'Аванс, сумма']);
    const srcAdvPlanCol = findCol(['Аванс план', 'Аванс план дата', 'Дата план Аванс']);
    const srcAdvFactCol = findCol(['Аванс факт', 'Аванс факт дата', 'Дата факт Аванс']);
    const srcBalSumCol = findCol(['Баланс сумма', 'Баланс, сумма', 'Остаток сумма', 'Остаток, сумма']);
    const srcBalPlanCol = findCol(['Баланс план', 'Баланс план дата', 'Дата план Баланс', 'Остаток план']);
    const srcBalFactCol = findCol(['Баланс факт', 'Баланс факт дата', 'Дата факт Баланс', 'Остаток факт']);
    const srcDefSumCol = findCol(['Отсрочка сумма', 'Отсрочка, сумма']);
    const srcDefPlanCol = findCol(['Отсрочка план', 'Отсрочка план дата', 'Дата План Отсрочка']);
    const srcDefFactCol = findCol(['Отсрочка факт', 'Отсрочка факт дата', 'Дата Факт Отсрочка']);
    const srcValidationError = payValidateSourceSyncColumns_({
      headers: hdr,
      srcArticleCol: srcArticleCol,
      srcSpecCol: srcSpecCol,
      srcAdvSumCol: srcAdvSumCol,
      srcAdvPlanCol: srcAdvPlanCol,
      srcAdvFactCol: srcAdvFactCol,
      srcBalSumCol: srcBalSumCol,
      srcBalPlanCol: srcBalPlanCol,
      srcBalFactCol: srcBalFactCol,
      srcDefSumCol: srcDefSumCol,
      srcDefPlanCol: srcDefPlanCol,
      srcDefFactCol: srcDefFactCol
    });
    if (srcValidationError) {
      ui.alert('❌ Ошибка структуры листа', srcValidationError, ui.ButtonSet.OK);
      return;
    }

    const destSpreadsheet = payOpenSpreadsheetByIdWithRetry_(externalSpreadsheetId, 8);
    const destSheet = destSpreadsheet.getSheetByName(externalSheetName);
    if (!destSheet) {
      ui.alert('❌ Ошибка', `Лист "${externalSheetName}" не найден.`, ui.ButtonSet.OK);
      return;
    }

    const numRows = destSheet.getLastRow();
    if (numRows < 2) {
      ui.alert('ℹ️ Инфо', `На листе "${externalSheetName}" нет данных для обновления.`, ui.ButtonSet.OK);
      return;
    }
    const destHeader = destSheet.getRange(1, 1, 1, destSheet.getLastColumn()).getDisplayValues()[0] || [];
    // Тот же мощный канон, что и для источника — чтобы «Номер_спецификации» / «№ спецификации»
    // в «Закуплено» тоже находились.
    const canonDest = canon;
    const findDestCol1Based = (variants) => {
      const cv = variants.map(canonDest);
      for (let c = 0; c < destHeader.length; c++) {
        if (cv.indexOf(canonDest(destHeader[c])) >= 0) return c + 1;
      }
      return -1;
    };
    const destSpecCol1Based = findDestCol1Based(
      ['Номер закупки/спецификации', 'Номер закупки', 'Номер спецификации', '№ спецификации', 'Спецификация', 'Спецификации']
    );
    const destValidationError = payValidateDestSyncColumns_(destHeader, destSpecCol1Based);
    if (destValidationError) {
      ui.alert('❌ Ошибка структуры листа "Закуплено"', destValidationError, ui.ButtonSet.OK);
      return;
    }
    const keyColsCount = Math.max(9, destSpecCol1Based);
    const keyRange = destSheet.getRange(1, 1, numRows, keyColsCount).getValues();
    const targetRange = destSheet.getRange(1, 64, numRows, 9); // BL:BT
    const targetValues = targetRange.getValues();

    const normalize = (val) => {
      if (val == null || val === '') return '';
      return String(val).replace(/\.0$/, '').toLowerCase().trim().replace(/\s+/g, '');
    };
    const normalizeSpec = (val) => {
      if (val == null || val === '') return '';
      if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
        const d = val.getDate();
        const m = val.getMonth() + 1;
        return d + '/' + m;
      }
      let s = String(val).toLowerCase().trim().replace(/\s+/g, '').replace(/\.0$/, '');
      let m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:t.*)?$/);
      if (m) return parseInt(m[3], 10) + '/' + parseInt(m[2], 10);
      m = s.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
      if (m) return parseInt(m[1], 10) + '/' + parseInt(m[2], 10);
      m = s.match(/^(\d{1,2})\/0(\d)$/);
      if (m) return parseInt(m[1], 10) + '/' + parseInt(m[2], 10);
      return s;
    };
    const cleanLog = (val) => val ? String(val).replace(/\n/g, ' ').trim() : '-';
    const hasFactValue = (rawVal, displayVal) => {
      if (rawVal === true) return true; // чекбокс
      if (rawVal === false) return false;
      if (rawVal instanceof Date && !isNaN(rawVal.getTime())) return true;
      const rawText = rawVal == null ? '' : String(rawVal).trim();
      const dispText = displayVal == null ? '' : String(displayVal).trim();
      if (/^https?:\/\//i.test(rawText) || /^https?:\/\//i.test(dispText)) return false;
      // Текст считаем фактом только если он похож на дату.
      if (/^\d{1,2}[./-]\d{1,2}([./-]\d{2,4})?$/.test(rawText)) return true;
      if (/^\d{1,2}[./-]\d{1,2}([./-]\d{2,4})?$/.test(dispText)) return true;
      return false;
    };
    const parseAmount = (rawVal, displayVal) => {
      const rawText = rawVal == null ? '' : String(rawVal).trim();
      const disp = displayVal == null ? '' : String(displayVal).trim();
      // Защита от гиперссылок/URL: это не сумма.
      if (/^https?:\/\//i.test(rawText) || /^https?:\/\//i.test(disp)) return '';
      if (typeof rawVal === 'number' && isFinite(rawVal)) return rawVal;
      if (/^-?\d+(?:[.,]\d+)?$/.test(rawText)) return parseFloat(rawText.replace(',', '.'));
      const cleaned = disp
        .replace(/\u00A0/g, ' ')
        .replace(/\s+/g, '')
        .replace(/,/g, '.')
        .replace(/[^\d.-]/g, '');
      if (/^-?\d+(?:\.\d+)?$/.test(cleaned)) return parseFloat(cleaned);
      return '';
    };

    const destMap = new Map(); // art|sup|spec -> row
    const destMapNoArticle = new Map(); // sup|spec -> row
    const destMapArtSpec = new Map(); // art|spec -> row (только если уникально)
    const destMapSpecUnique = new Map(); // spec -> row (только если уникально)
    const specCounts = {};
    const artSpecCounts = {};
    for (let r = 1; r < numRows; r++) {
      const destArt = normalize(keyRange[r][3]);  // D
      const destSup = normalize(keyRange[r][6]);  // G
      const destSpec = normalizeSpec(keyRange[r][destSpecCol1Based - 1]);
      if (destArt || destSup || destSpec) {
        destMap.set(`${destArt}|${destSup}|${destSpec}`, r);
        if (destSup || destSpec) destMapNoArticle.set(`${destSup}|${destSpec}`, r);
        if (destArt || destSpec) {
          const kArtSpec = `${destArt}|${destSpec}`;
          artSpecCounts[kArtSpec] = (artSpecCounts[kArtSpec] || 0) + 1;
          destMapArtSpec.set(kArtSpec, r);
        }
        if (destSpec) specCounts[destSpec] = (specCounts[destSpec] || 0) + 1;
        if (destSpec) destMapSpecUnique.set(destSpec, r);
      }
    }

    const successLog = [];
    const errorLog = [];
    let hasChanges = false;
    const changedRows = [];

    for (let i = 2; i < srcData.length; i++) {
      const row = srcData[i];
      const origArt = row[srcArticleCol];
      const origSup = srcSupplierCol >= 0 ? row[srcSupplierCol] : '';
      const origSpec = row[srcSpecCol];
      const srcArt = normalize(origArt);
      const srcSup = normalize(origSup);
      const srcSpec = normalizeSpec(origSpec);
      // Для связки должна быть спецификация и хотя бы поставщик или артикул.
      if (!srcSpec || (!srcArt && !srcSup)) continue;

      const key = `${srcArt}|${srcSup}|${srcSpec}`;
      const keyNoArticle = `${srcSup}|${srcSpec}`;
      const keyArtSpec = `${srcArt}|${srcSpec}`;
      let destRowIdx = -1;
      if (destMap.has(key)) destRowIdx = destMap.get(key);
      else if (destMapNoArticle.has(keyNoArticle)) destRowIdx = destMapNoArticle.get(keyNoArticle);
      else if (srcArt && srcSpec && artSpecCounts[keyArtSpec] === 1 && destMapArtSpec.has(keyArtSpec)) destRowIdx = destMapArtSpec.get(keyArtSpec);
      else if (srcSpec && specCounts[srcSpec] === 1 && destMapSpecUnique.has(srcSpec)) destRowIdx = destMapSpecUnique.get(srcSpec);
      if (destRowIdx < 0) {
        errorLog.push(`Стр. ${i + 1} | Арт: ${cleanLog(origArt)} | Пост: ${cleanLog(origSup)} | Спец: ${cleanLog(origSpec)}`);
        continue;
      }

      const avansSumPlan = parseAmount(row[srcAdvSumCol], srcDisplay[i][srcAdvSumCol]);
      const avansDatePlan = row[srcAdvPlanCol];
      const avansDateFact = row[srcAdvFactCol];
      const avansFactPresent = hasFactValue(avansDateFact, srcDisplay[i][srcAdvFactCol]);

      const balSumPlan = parseAmount(row[srcBalSumCol], srcDisplay[i][srcBalSumCol]);
      const balDatePlan = row[srcBalPlanCol];
      const balDateFact = row[srcBalFactCol];
      const balFactPresent = hasFactValue(balDateFact, srcDisplay[i][srcBalFactCol]);

      const defSumPlan = parseAmount(row[srcDefSumCol], srcDisplay[i][srcDefSumCol]);
      const defDatePlan = row[srcDefPlanCol];
      const defDateFact = row[srcDefFactCol];
      const defFactPresent = hasFactValue(defDateFact, srcDisplay[i][srcDefFactCol]);

      targetValues[destRowIdx][0] = avansDatePlan; // BL
      targetValues[destRowIdx][1] = avansSumPlan;  // BM
      targetValues[destRowIdx][2] = avansFactPresent ? avansSumPlan : ''; // BN

      targetValues[destRowIdx][3] = balDatePlan; // BO
      targetValues[destRowIdx][4] = balSumPlan;  // BP
      targetValues[destRowIdx][5] = balFactPresent ? balSumPlan : ''; // BQ

      targetValues[destRowIdx][6] = defDatePlan; // BR
      targetValues[destRowIdx][7] = defSumPlan;  // BS
      targetValues[destRowIdx][8] = defFactPresent ? defSumPlan : ''; // BT

      hasChanges = true;
      changedRows.push(destRowIdx);
      successLog.push(`Стр. ${i + 1} | Арт: ${cleanLog(origArt)} | Пост: ${cleanLog(origSup)} | Спец: ${cleanLog(origSpec)}`);
    }

    if (hasChanges) {
      payWriteChangedBlocksWithRetry_(destSheet, targetValues, changedRows, 64);
    }
    showSyncReport(successLog, errorLog, srcSheet.getName(), externalSheetName);
  } catch (error) {
    logError('Ошибка обновления таблицы "Закуплено"', error);
    ui.alert('❌ Ошибка', `Произошла ошибка при обновлении:\n${error.message}`, ui.ButtonSet.OK);
  }
}

function payWriteChangedBlocksWithRetry_(sheet, allTargetValues, changedRowsZeroBased, targetStartCol1Based) {
  const uniq = {};
  for (let i = 0; i < changedRowsZeroBased.length; i++) {
    uniq[changedRowsZeroBased[i]] = true;
  }
  const rows = Object.keys(uniq)
    .map(function (x) {
      return parseInt(x, 10);
    })
    .filter(function (x) {
      return isFinite(x) && x >= 0;
    })
    .sort(function (a, b) {
      return a - b;
    });

  if (!rows.length) return;

  let start = rows[0];
  let prev = rows[0];
  const blocks = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (r === prev + 1) {
      prev = r;
      continue;
    }
    blocks.push({ from: start, to: prev });
    start = r;
    prev = r;
  }
  blocks.push({ from: start, to: prev });

  const maxBlockRows = 120;
  for (let bi = 0; bi < blocks.length; bi++) {
    const b = blocks[bi];
    let from = b.from;
    while (from <= b.to) {
      const to = Math.min(b.to, from + maxBlockRows - 1);
      const rowsCount = to - from + 1;
      const payload = allTargetValues.slice(from, to + 1);
      let ok = false;
      let lastErr = null;
      for (let a = 1; a <= 5; a++) {
        try {
          sheet.getRange(from + 1, targetStartCol1Based, rowsCount, 9).setValues(payload);
          ok = true;
          break;
        } catch (e) {
          lastErr = e;
          Utilities.sleep(900 * a);
        }
      }
      if (!ok) {
        throw new Error(
          'Не удалось записать обновления в «Закуплено» после нескольких попыток. ' +
            (lastErr && lastErr.message ? lastErr.message : String(lastErr))
        );
      }
      from = to + 1;
    }
  }
}

function payOpenSpreadsheetByIdWithRetry_(spreadsheetId, maxAttempts) {
  const attempts = Math.max(1, maxAttempts || 3);
  let lastErr = null;
  for (let i = 1; i <= attempts; i++) {
    try {
      return SpreadsheetApp.openById(spreadsheetId);
    } catch (e) {
      lastErr = e;
      const msg = String(e && e.message ? e.message : e);
      const transient =
        msg.indexOf('слишком долго не может получить доступ') >= 0 ||
        msg.indexOf('Service Spreadsheets timed out') >= 0 ||
        msg.indexOf('Service unavailable') >= 0 ||
        msg.indexOf('Exception: Service Spreadsheets') >= 0;
      if (!transient) break;
      if (i < attempts) Utilities.sleep(800 * i);
    }
  }
  throw new Error(
    'Не удалось открыть таблицу по ID после ' +
      attempts +
      ' попыток: ' +
      spreadsheetId +
      '. Детали: ' +
      (lastErr && lastErr.message ? lastErr.message : String(lastErr))
  );
}

function payValidateSourceSyncColumns_(cfg) {
  const h = cfg && cfg.headers ? cfg.headers : [];
  if (!h.length) return 'Не удалось прочитать заголовки источника (строка 2).';
  const required = [
    ['Артикул ВБ', cfg.srcArticleCol],
    ['Номер спецификации', cfg.srcSpecCol],
    ['Аванс сумма', cfg.srcAdvSumCol],
    ['Дата план Аванс', cfg.srcAdvPlanCol],
    ['Дата факт Аванс', cfg.srcAdvFactCol],
    ['Баланс сумма', cfg.srcBalSumCol],
    ['Дата план Баланс', cfg.srcBalPlanCol],
    ['Дата факт Баланс', cfg.srcBalFactCol],
    ['Отсрочка сумма', cfg.srcDefSumCol],
    ['Дата План Отсрочка', cfg.srcDefPlanCol],
    ['Дата Факт Отсрочка', cfg.srcDefFactCol]
  ];
  const maxCols = h.length;
  // findCol теперь возвращает -1 при отсутствии колонки (раньше — цифровой фолбэк,
  // и условие `idx >= 0` всегда было истинным — валидатор пропускал даже отсутствующие шапки).
  const missing = required
    .filter(function (x) {
      const idx = x[1];
      return !(isFinite(idx) && idx >= 0 && idx < maxCols);
    })
    .map(function (x) {
      return x[0];
    });
  if (!missing.length) return '';
  const visible = h
    .map(function (s) { return String(s == null ? '' : s).trim(); })
    .filter(function (s) { return s; })
    .slice(0, 30)
    .join(' | ');
  return (
    'В активном листе не найдены обязательные колонки: ' +
    missing.join(', ') +
    '.\nФактические заголовки строки 2: ' + visible +
    '\nПереименуйте колонки в эти варианты или сообщите, какое название использовать.'
  );
}

function payValidateDestSyncColumns_(destHeader, destSpecCol1Based) {
  const h = destHeader || [];
  if (!h.length) return 'Не удалось прочитать заголовки в листе "Закуплено".';
  const maxCols = h.length;
  const requiredCols = [
    ['Артикул ВБ (D)', 4],
    ['Поставщик (G)', 7],
    ['Номер закупки/спецификации', destSpecCol1Based],
    ['Дата оплаты (Аванс) BL', 64],
    ['Сумма аванса (план) BM', 65],
    ['Сумма аванса (факт) BN', 66],
    ['Дата оплаты (Остаток) BO', 67],
    ['Сумма остатка (план) BP', 68],
    ['Сумма остатка (факт) BQ', 69],
    ['Дата оплаты отсрочки BR', 70],
    ['Сумма отсрочки (план) BS', 71],
    ['Сумма отсрочки (факт) BT', 72]
  ];
  const missing = requiredCols
    .filter(function (x) {
      const c = x[1];
      return !(isFinite(c) && c >= 1 && c <= maxCols);
    })
    .map(function (x) {
      return x[0];
    });
  if (!missing.length) return '';
  const visible = h
    .map(function (s) { return String(s == null ? '' : s).trim(); })
    .filter(function (s) { return s; })
    .slice(0, 30)
    .join(' | ');
  return (
    'В листе "Закуплено" не хватает обязательных колонок: ' +
    missing.join(', ') +
    '.\nФактические заголовки строки 1: ' + visible
  );
}

function showSyncReport(successLog, errorLog, sourceSheetName, targetSheetName) {
  const src = sourceSheetName || 'источник';
  const dst = targetSheetName || 'Закуплено';
  const htmlContent = `
    <div style="font-family: sans-serif; padding: 10px;">
      <h3 style="color: #1a73e8; text-align: center; margin-top: 0;">Отчет об обновлении оплат</h3>
      <p style="font-size:12px;color:#666;margin:0 0 10px 0;">Источник: <b>${src}</b> → Цель: <b>${dst}</b></p>
      
      <h4 style="color: #1e8e3e;">✅ Успешно найдено и обновлено (${successLog.length} шт.):</h4>
      <textarea style="width: 100%; height: 120px; box-sizing: border-box; background: #f1f8f1; border: 1px solid #c6e0c6; padding: 5px;" readonly>${successLog.length > 0 ? successLog.join('\\n') : 'Пусто'}</textarea>
      
      <h4 style="color: #d93025; margin-top: 20px;">❌ Не найдено в таблице "${dst}" (${errorLog.length} шт.):</h4>
      <textarea style="width: 100%; height: 150px; box-sizing: border-box; background: #fce8e6; border: 1px solid #f2cfcf; padding: 5px;" readonly>${errorLog.length > 0 ? errorLog.join('\\n') : 'Пусто'}</textarea>
      
      <p style="font-size: 12px; color: #666; margin-top: 15px;"><i>*Если строка не найдена, проверьте совпадение Артикула, Поставщика и Спецификации в обеих таблицах.</i></p>
    </div>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(600)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📊 Результаты синхронизации');
}
