function updateExternalPurchases() {
  const ui = SpreadsheetApp.getUi();
  const externalSpreadsheetId = CONFIG.EXTERNAL_SPREADSHEET_ID;
  const externalSheetName = CONFIG.EXTERNAL_SHEET_NAME;

  try {
    const srcSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!srcSheet) {
      ui.alert('❌ Ошибка', `Лист "${SHEET_NAME}" не найден.`, ui.ButtonSet.OK);
      return;
    }
    const srcData = srcSheet.getDataRange().getValues();

    const destSpreadsheet = SpreadsheetApp.openById(externalSpreadsheetId);
    const destSheet = destSpreadsheet.getSheetByName(externalSheetName);
    if (!destSheet) {
      ui.alert('❌ Ошибка', `Лист "${externalSheetName}" не найден.`, ui.ButtonSet.OK);
      return;
    }

    const destData = destSheet.getDataRange().getValues();
    const numRows = destData.length;
    const targetRange = destSheet.getRange(1, 64, numRows, 9); // BL:BT
    const targetValues = targetRange.getValues();

    const normalize = (val) => {
      if (val == null || val === '') return '';
      return String(val).replace(/\.0$/, '').toLowerCase().trim().replace(/\s+/g, '');
    };
    const cleanLog = (val) => val ? String(val).replace(/\n/g, ' ').trim() : '-';

    const destMap = new Map();
    for (let r = 1; r < numRows; r++) {
      const destArt = normalize(destData[r][3]);  // D
      const destSup = normalize(destData[r][6]);  // G
      const destSpec = normalize(destData[r][8]); // I
      if (destArt || destSup || destSpec) {
        destMap.set(`${destArt}|${destSup}|${destSpec}`, r);
      }
    }

    const successLog = [];
    const errorLog = [];
    let hasChanges = false;

    for (let i = 2; i < srcData.length; i++) {
      const row = srcData[i];
      const origArt = row[0];
      const origSup = row[8];
      const origSpec = row[11];
      const srcArt = normalize(origArt);
      const srcSup = normalize(origSup);
      const srcSpec = normalize(origSpec);
      if (!srcArt && !srcSup && !srcSpec) continue;

      const key = `${srcArt}|${srcSup}|${srcSpec}`;
      if (!destMap.has(key)) {
        errorLog.push(`Стр. ${i + 1} | Арт: ${cleanLog(origArt)} | Пост: ${cleanLog(origSup)} | Спец: ${cleanLog(origSpec)}`);
        continue;
      }

      const destRowIdx = destMap.get(key);
      const avansSumPlan = row[19];
      const avansDatePlan = row[20];
      const avansDateFact = row[21];

      const balSumPlan = row[22];
      const balDatePlan = row[23];
      const balDateFact = row[24];

      const defSumPlan = row[25];
      const defDatePlan = row[26];
      const defDateFact = row[27];

      targetValues[destRowIdx][0] = avansDatePlan; // BL
      targetValues[destRowIdx][1] = avansSumPlan;  // BM
      targetValues[destRowIdx][2] = (avansDateFact && String(avansDateFact).trim() !== '') ? avansSumPlan : ''; // BN

      targetValues[destRowIdx][3] = balDatePlan; // BO
      targetValues[destRowIdx][4] = balSumPlan;  // BP
      targetValues[destRowIdx][5] = (balDateFact && String(balDateFact).trim() !== '') ? balSumPlan : ''; // BQ

      targetValues[destRowIdx][6] = defDatePlan; // BR
      targetValues[destRowIdx][7] = defSumPlan;  // BS
      targetValues[destRowIdx][8] = (defDateFact && String(defDateFact).trim() !== '') ? defSumPlan : ''; // BT

      hasChanges = true;
      successLog.push(`Стр. ${i + 1} | Арт: ${cleanLog(origArt)} | Пост: ${cleanLog(origSup)} | Спец: ${cleanLog(origSpec)}`);
    }

    if (hasChanges) targetRange.setValues(targetValues);
    showSyncReport(successLog, errorLog);
  } catch (error) {
    logError('Ошибка обновления таблицы "Закуплено"', error);
    ui.alert('❌ Ошибка', `Произошла ошибка при обновлении:\n${error.message}`, ui.ButtonSet.OK);
  }
}

function showSyncReport(successLog, errorLog) {
  const htmlContent = `
    <div style="font-family: sans-serif; padding: 10px;">
      <h3 style="color: #1a73e8; text-align: center; margin-top: 0;">Отчет об обновлении оплат</h3>
      
      <h4 style="color: #1e8e3e;">✅ Успешно найдено и обновлено (${successLog.length} шт.):</h4>
      <textarea style="width: 100%; height: 120px; box-sizing: border-box; background: #f1f8f1; border: 1px solid #c6e0c6; padding: 5px;" readonly>${successLog.length > 0 ? successLog.join('\\n') : 'Пусто'}</textarea>
      
      <h4 style="color: #d93025; margin-top: 20px;">❌ Не найдено в таблице "Закуплено" (${errorLog.length} шт.):</h4>
      <textarea style="width: 100%; height: 150px; box-sizing: border-box; background: #fce8e6; border: 1px solid #f2cfcf; padding: 5px;" readonly>${errorLog.length > 0 ? errorLog.join('\\n') : 'Пусто'}</textarea>
      
      <p style="font-size: 12px; color: #666; margin-top: 15px;"><i>*Если строка не найдена, проверьте совпадение Артикула, Поставщика и Спецификации в обеих таблицах.</i></p>
    </div>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(600)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📊 Результаты синхронизации');
}
