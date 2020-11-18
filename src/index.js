const { Worker } = require('worker_threads');
const fs = require('fs');
const path = require('path');
const JSDOM = require('jsdom').JSDOM;
const XLSX = require('xlsx');
const diff = require("deep-object-diff"); //alt xlsx-populate
const templateHtml = fs.readFileSync(path.join(__dirname, './template.html')).toString();

const writeDiffExcelhtml = async (oldExcelPath, newExcelPath, outputPath) => {
  console.log('Start Diff');
  console.log('Old File ->', oldExcelPath);
  console.log('New File ->', newExcelPath);
  console.log('Output Path ->', outputPath);

  const options = {
    cellStyles: true,
  }

  const w1 = XLSX.readFile(oldExcelPath);
  const w2 = XLSX.readFile(newExcelPath);

  const addList = diff.addedDiff(w1, w2).Sheets;
  const deleteList = diff.deletedDiff(w1, w2).Sheets;
  const updateList = diff.updatedDiff(w1, w2).Sheets;

  const workbook1 = XLSX.readFile(oldExcelPath, options);
  const workbook2 = XLSX.readFile(newExcelPath, options);

  const changeSheets = {
    add: addList,
    delete: deleteList,
    update: updateList
  }

  const changeResult = {
    add: {},
    delete: {},
    update: {}
  };

  const changeSheetNameList = {};

  Object.keys(changeSheets).forEach((changeType) => {
    const sheets = changeSheets[changeType];
    if (typeof sheets !== 'undefined') {
      Object.keys(sheets).forEach((sheet) => {
        const sheetResult = [];

        if (!workbook1.Sheets[sheet] && workbook2.Sheets[sheet]) sheetResult.push({ c: 'add' });
        if (workbook1.Sheets[sheet] && !workbook2.Sheets[sheet]) sheetResult.push({ c: 'delete' });
        if (typeof sheets[sheet] !== 'undefined') {
          Object.keys(sheets[sheet]).forEach((cechages) => {
            switch (changeType) {
              case 'add':
                if (sheets[sheet][cechages].v) sheetResult.push({
                  c: cechages,
                  detail: `${cechages}に値≪${sheets[sheet][cechages].v}≫が追加されました。`
                });
                break;
              case 'delete':
                sheetResult.push({
                  c: cechages,
                  detail: `${cechages}の値は削除されました。`
                });
                break;
              case 'update':
                // if(cechages === '!merges')
                if (sheets[sheet][cechages].v) sheetResult.push({
                  c: cechages,
                  detail: `${cechages}の値は<<${workbook1.Sheets[sheet][cechages].v}>>から≪${sheets[sheet][cechages].v}≫に変更されました。`
                });
                break;
            }
          });
        }

        if (sheetResult.length !== 0) changeSheetNameList[sheet] = true;
        changeResult[changeType][sheet] = sheetResult;
      });
    }
  });

  const templateDOM = new JSDOM(templateHtml);
  const templateDocument = templateDOM.window.document;

  const sheetNameArray = Object.keys(changeSheetNameList);
  const promiseArray = [];

  const createWorker = (sheet) => {
    return new Promise((resolve, reject) => {
      const worker = new Worker(path.join(__dirname, './createTable.js'), {
        workerData: sheet
      });
      worker.on('message', (result) => {
        resolve({ error: result.error, dom: new JSDOM(result.dom) });
      });
      worker.on('exit', (code) => {
        if (code !== 0) reject(new Error());
      });
    });
  }

  for (let i = 0; i < sheetNameArray.length; i++) {
    const sheet = sheetNameArray[i];
    promiseArray.push(createWorker(workbook1.Sheets[sheet]));
    promiseArray.push(createWorker(workbook2.Sheets[sheet]));
  }
  const createTableResult = await Promise.all(promiseArray);

  for (let index = 0; index < sheetNameArray.length; index++) {
    const sheet = sheetNameArray[index];
    console.log('warite sheet ', sheet);
    const oldDomInfo = createTableResult[index * 2];
    const newDomInfo = createTableResult[index * 2 + 1];
    const oldDom = oldDomInfo.dom;
    const newDom = newDomInfo.dom;

    if (!oldDomInfo.error && !newDomInfo.error) {
      Object.keys(changeResult).forEach((changeType) => {
        if (changeResult[changeType][sheet]) {
          changeResult[changeType][sheet].forEach(e => {
            if (workbook1.Sheets[sheet] && workbook2.Sheets[sheet]) {
              const document1 = oldDom.window.document;
              const element1 = document1.getElementById(`sjs-${e.c}`);
              if (element1) {
                element1.title = e.detail;
                element1.style.backgroundColor = '#ff9494';
                element1.style.cursor = 'pointer';
              }
              const document2 = newDom.window.document;
              const element2 = document2.getElementById(`sjs-${e.c}`);
              if (element2) {
                element2.title = e.detail;
                element2.style.backgroundColor = '#a0ea9e';
                element2.style.cursor = 'pointer';
              }
            }
          });
        }
      });
    }

    const oldTable = oldDom.window.document.getElementsByTagName('table')[0];
    const newTable = newDom.window.document.getElementsByTagName('table')[0];

    oldTable.id = `old-table-${index}`;
    newTable.id = `new-table-${index}`;

    if (index !== 0) {
      oldTable.style.display = 'none';
      newTable.style.display = 'none';
    }

    templateDocument.getElementById('old-table').appendChild(oldTable);
    templateDocument.getElementById('new-table').appendChild(newTable);

    const span = templateDocument.createElement('span');
    const newContent = templateDocument.createTextNode(sheet);
    span.appendChild(newContent);
    span.id = `sheet-${index}`;
    span.className = 'sheet-name'
    if (index === 0) {
      span.style.color = 'whitesmoke';
      span.style.backgroundColor = 'gray';
    }

    templateDocument.getElementById('sheet-list').appendChild(span);
  };

  fs.writeFileSync(outputPath, templateDOM.serialize());
  console.log('End Diff');

  return true;
}

module.exports = writeDiffExcelhtml;