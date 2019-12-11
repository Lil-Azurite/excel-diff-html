const { workerData, parentPort} = require('worker_threads');
const JSDOM = require('jsdom').JSDOM;
const XLSX = require('xlsx');

const sheet = workerData;

if (typeof sheet === 'undefined') {
  parentPort.postMessage({ error: true, dom: new JSDOM('<table><tr color="red"><td style="width: 100%; max-width: 100%;">このバージョンではまだ存在しないor削除されたシートです</td></tr></table>').serialize()});
}
else if (typeof sheet['!ref'] === 'undefined') {
  parentPort.postMessage({ error: true, dom: new JSDOM(`<table><tr color="red"><td style="width: 100%; max-width: 100%;">設定が存在しないシートです。</td></tr></table>`).serialize()});
}
else {
  const [_s, _e]= sheet['!ref'].split(':');
  const d_s = XLSX.utils.decode_cell(_s);
  const d_e = XLSX.utils.decode_cell(_e);
  if (5000 < Object.keys(sheet).length || 1000 < d_e.c - d_s.c || 1000 < d_e.c - d_s.c) {
    parentPort.postMessage({ error: true, dom: new JSDOM(`<table><tr color="red"><td style="width: 100%; max-width: 100%;">行数・列数が多すぎですexcelの作り方を見直してください。このシートのセル範囲「${sheet['!ref']}」</td></tr></table>`).serialize()});
  }
  else {
    const dom = new JSDOM(XLSX.utils.sheet_to_html(sheet));
    const document = dom.window.document;
    const targetRows = document.getElementsByTagName('tr');

    if (sheet[`!rows`]) {
      for(let i = 0; i < sheet[`!rows`].length; i++) {
        const e = sheet[`!rows`][i];
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
        const c = sheet[`!cols`][s];
        if (c != null) result += c.wpx;
        else result += 72;
        s++;
      }
      return result;
    }

    if (sheet[`!cols`]) {
      const targetTds = document.getElementsByTagName('td');
      let i = 0;
      const length = targetTds.length;
      while (i < length) {
        const targetTab = targetTds[i];
        const colSpan = targetTab.getAttribute('colspan');
        const [_, c] = targetTab.id.split('-');
        const s = XLSX.utils.decode_cell(c).c;
        const e = colSpan!= null ? s + Number(colSpan) - 1 : s;
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
}
