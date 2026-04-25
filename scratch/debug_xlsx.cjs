const XLSX = require('xlsx');
const fs = require('fs');

const processData = (data) => {
  if (data.length === 0) return data;

  const rawHeader = data[0];
  const header = {
    A: rawHeader.A,
    D: rawHeader.D,
    F: rawHeader.F,
    H: rawHeader.H,
    L: rawHeader.L,
    N: rawHeader.N,
    P: rawHeader.P,
    T: rawHeader.T,
    W: rawHeader.W
  };

  const rows = data.slice(1).filter(row => {
    const isSTTUGDK_O = String(row.D || '').trim().toUpperCase() === 'O';
    const isSIFOJ_6 = String(row.P || '').trim().startsWith('6');
    return isSTTUGDK_O && !isSIFOJ_6;
  });

  const grouped = rows.reduce((acc, row) => {
    const key = String(row.A || '').trim();
    if (!key) return acc;

    if (!acc[key]) {
      acc[key] = { 
        A: row.A,
        D: row.D,
        F: row.F,
        H: row.H,
        L: row.L,
        N: row.N,
        P: row.P,
        T: parseFloat(row.T) || 0,
        W: parseFloat(row.W) || 0
      };
    } else {
      acc[key].T += parseFloat(row.T) || 0;
      acc[key].W += parseFloat(row.W) || 0;
    }
    return acc;
  }, {});

  const processedRows = Object.values(grouped)
    .map(row => ({
      ...row,
      T: Number(row.T.toFixed(2)),
      W: Number(row.W.toFixed(2))
    }))
    .filter(row => row.T >= 1000);

  return [header, ...processedRows];
};

const wb = XLSX.readFile('src/assets/test.xlsx');
const sheet = wb.Sheets[wb.SheetNames[0]];
const jsonData = XLSX.utils.sheet_to_json(sheet, { header: "A", range: 9 });

console.log('JSON Data Length:', jsonData.length);
const processedData = processData(jsonData);
console.log('Processed Data Length:', processedData.length);

const headerRow = processedData[0];
const dataRows = processedData.slice(1);

const rows13 = dataRows.filter(row => String(row.N || '').trim().startsWith('13'));
const rows14 = dataRows.filter(row => String(row.N || '').trim().startsWith('14'));
const rows16 = dataRows.filter(row => String(row.N || '').trim().startsWith('16'));

console.log('13-LTL count:', rows13.length);
console.log('14-FTL count:', rows14.length);
console.log('16-MINIFTL count:', rows16.length);
