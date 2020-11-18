const { workerData, parentPort } = require('worker_threads');
const JSDOM = require('jsdom').JSDOM;
const XLSX = require('xlsx');

const calculateLimit = (sheet) => {
  const { s, e } = XLSX.utils.decode_range(sheet['!ref']);

  let el = { c: 0, r: 0 };

  for (let row = s.r; row <= e.r; row++) {
    for (let col = s.c; col <= e.c; col++) {
      const range = { r: row, c: col };
      const address = XLSX.utils.encode_cell(range);
      const cell = sheet[address];
      if (cell && cell.w !== null && cell.v !== null) {
        el = range;
      }
    }
  }

  return XLSX.utils.encode_range({
    s, e: {
      c: el.c + 10,
      r: el.r + 10
    }
  });
}

const sheet = workerData;

if (typeof sheet === 'undefined') {
  parentPort.postMessage({ error: true, dom: new JSDOM('<table><tr color="red"><td style="width: 100%; max-width: 100%;">このバージョンではまだ存在しないor削除されたシートです</td></tr></table>').serialize() });
}
else if (typeof sheet['!ref'] === 'undefined') {
  parentPort.postMessage({ error: true, dom: new JSDOM(`<table><tr color="red"><td style="width: 100%; max-width: 100%;">設定が存在しないシートです。</td></tr></table>`).serialize() });
}
else {
  const ref = calculateLimit(sheet);

  sheet['!ref'] = ref;

  const shhetR = sheet[`!rows`];
  const sheetC = sheet[`!cols`];

  const dom = new JSDOM(XLSX.utils.sheet_to_html(sheet));
  const document = dom.window.document;

  const targetRows = document.getElementsByTagName('tr');
  if (shhetR) {
    for (let i = 0; i < shhetR.length; i++) {
      const e = shhetR[i];
      if (e != null) {
        const targetRow = targetRows[i];
        if (typeof targetRow !== 'undefined') {
          const height = `${e.hpx}px`;
          targetRow.style.minHeight = height;
          targetRow.style.maxHeight = height;
          targetRow.style.height = height;
        }
      }
    };
  }

  const sumTdWidth = (s, e) => {
    let result = 0;
    while (s <= e) {
      const c = sheetC[s];
      if (c != null) result += c.wpx;
      else result += 72;
      s++;
    }
    return result;
  }

  if (sheetC) {
    const targetTds = document.getElementsByTagName('td');
    let i = 0;
    const length = targetTds.length;
    while (i < length) {
      const targetTab = targetTds[i];
      const colSpan = targetTab.getAttribute('colspan');
      const [_, c] = targetTab.id.split('-');
      const s = XLSX.utils.decode_cell(c).c;
      const e = colSpan != null ? s + Number(colSpan) - 1 : s;
      const width = `${sumTdWidth(s, e)}px`;
      const style = targetTab.style;
      style.minWidth = width;
      style.maxWidth = width;
      style.width = width;
      i++;
    }
  }

  parentPort.postMessage({
    error: false,
    dom: dom.serialize()
  });
}

