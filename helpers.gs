function pushStringAttr(attr, id, val) {
  if (!id) return;
  const v = safeString(val);
  if (v !== '') attr.push({ meta: buildAttrMeta(id).meta, value: v });
}

function pushNumberAttr(attr, id, val) {
  if (!id) return;
  const v = parseNumber(val);
  if (v !== null) attr.push({ meta: buildAttrMeta(id).meta, value: v });
}

function pushTimeAttr(attr, id, val) {
  if (!id) return;
  const v = parseDateToMS(val);
  if (v) attr.push({ meta: buildAttrMeta(id).meta, value: v });
}

function parseNumber(val) {
  if (val === null || val === undefined || val === '') return null;
  if (typeof val === 'number') return isNaN(val) ? null : val;
  const num = parseFloat(val.toString().replace(/\s/g, '').replace(',', '.'));
  return isNaN(num) ? null : num;
}

function parseDateToMS(val) {
  if (!val || val === '' || val === '-') return null;
  let d;
  if (Object.prototype.toString.call(val) === '[object Date]') {
    d = val;
  } else {
    const str = val.toString().trim();
    const m = str.match(/^(\d{1,2})[\.\-\/](\d{1,2})[\.\-\/](\d{2,4})/);
    if (m) {
      let year = parseInt(m[3], 10);
      if (year < 100) year += 2000;
      d = new Date(year, parseInt(m[2], 10) - 1, parseInt(m[1], 10), 12, 0, 0);
    } else {
      d = new Date(str);
    }
  }
  if (!d || isNaN(d.getTime())) return null;
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function safeString(val) {
  return val == null ? '' : val.toString().trim();
}
