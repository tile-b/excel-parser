import { useState, useCallback } from 'react';
import XLSX from 'xlsx-js-style';

/**
 * Isolated function for processing the parsed Excel data.
 * This is where .map() and .filter() logic will be added later.
 * 
 * @param {Array} data - The array of objects parsed from the Excel file.
 * @returns {Array} - The processed array of objects.
 */
const processData = (data) => {
  if (data.length === 0) return data;

  const rawHeader = data[0];

  // Dynamic header mapping
  const findKey = (name) => Object.keys(rawHeader).find(k =>
    String(rawHeader[k] || '').trim().toUpperCase() === name.toUpperCase()
  );

  const keys = {
    BRUGDK: findKey('BRUGDK') || 'A',
    STTUGDK: findKey('STTUGDK') || 'D',
    NAZIV_MES_ISP: findKey('NAZIV_MES_ISP') || 'F',
    MESTO_MES_ISP: findKey('MESTO_MES_ISP') || 'H',
    DC: findKey('DC') || 'L',
    OTPREMA: findKey('OTPREMA') || 'N',
    SIFOJ: findKey('SIFOJ') || 'P',
    BRUTOMASA: findKey('BRUTOMASA') || 'T',
    PALETE: findKey('PALETE') || 'W',
  };

  const header = {
    F: rawHeader[keys.NAZIV_MES_ISP],
    H: rawHeader[keys.MESTO_MES_ISP],
    L: rawHeader[keys.DC],
    P: rawHeader[keys.SIFOJ],
    T: rawHeader[keys.BRUTOMASA],
    W: rawHeader[keys.PALETE]
  };

  const validRows = data.slice(1).filter(row => {
    // Column B (STTUGDK) should be 'O'
    return String(row[keys.STTUGDK] || '').trim().toUpperCase() === 'O';
  });

  const groupRows = (rowsToGroup) => {
    const grouped = rowsToGroup.reduce((acc, row) => {
      const key = String(row[keys.BRUGDK] || '').trim();
      if (!key) return acc;

      if (!acc[key]) {
        acc[key] = {
          F: row[keys.NAZIV_MES_ISP],
          H: row[keys.MESTO_MES_ISP],
          L: row[keys.DC],
          P: row[keys.SIFOJ],
          T: parseFloat(row[keys.BRUTOMASA]) || 0,
          W: parseFloat(row[keys.PALETE]) || 0
        };
      } else {
        acc[key].T += parseFloat(row[keys.BRUTOMASA]) || 0;
        acc[key].W += parseFloat(row[keys.PALETE]) || 0;
      }
      return acc;
    }, {});

    return Object.values(grouped).map(row => ({
      ...row,
      T: Number(row.T.toFixed(2)),
      W: Number(row.W.toFixed(2))
    }));
  };

  const rows13 = validRows.filter(row => String(row[keys.OTPREMA] || '').trim().startsWith('13'));
  const all13 = groupRows(rows13);
  const over1000_13 = all13.filter(row => row.T >= 1000);
  const under1000_13 = all13.filter(row => row.T < 1000);

  const rows14 = validRows.filter(row => String(row[keys.OTPREMA] || '').trim().startsWith('14'));
  const all14 = groupRows(rows14);

  const rows16 = validRows.filter(row => String(row[keys.OTPREMA] || '').trim().startsWith('16'));
  const all16 = groupRows(rows16);

  return { header, over1000_13, under1000_13, all13, all14, all16 };
};

export default function LtlParser() {
  const [file, setFile] = useState(null);
  const [error, setError] = useState('');
  const [processing, setProcessing] = useState(false);

  const handleFileChange = (e) => {
    const selectedFile = e.target.files[0];
    validateAndSetFile(selectedFile);
  };

  const validateAndSetFile = (selectedFile) => {
    setError('');
    if (!selectedFile) return;

    const isExcel = selectedFile.name.endsWith('.xlsx') ||
      selectedFile.name.endsWith('.xls') ||
      selectedFile.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
      selectedFile.type === 'application/vnd.ms-excel';

    if (!isExcel) {
      setError('Molimo otpremite važeću Excel datoteku (.xlsx ili .xls)');
      setFile(null);
      return;
    }

    setFile(selectedFile);
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    e.stopPropagation();
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    const droppedFile = e.dataTransfer.files[0];
    validateAndSetFile(droppedFile);
  };

  const handleProcess = async () => {
    if (!file) return;

    setProcessing(true);
    setError('');

    try {
      const reader = new FileReader();

      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });

          // Get the first sheet
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];

          // Convert sheet to JSON
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: "A", range: 9 });

          // Process data
          const { header: headerRow, over1000_13: dataRows, all13, all14, all16 } = processData(jsonData);

          // Create new workbook
          const newWorkbook = XLSX.utils.book_new();

          // Helper to create sheet with header, totals, and advanced styling
          const createSheet = (rows) => {
            const totalT = rows.reduce((sum, row) => sum + (Number(row.T) || 0), 0);
            const totalW = rows.reduce((sum, row) => sum + (Number(row.W) || 0), 0);

            const totalRow = {
              F: 'TOTAL',
              T: Number(totalT.toFixed(2)),
              W: Number(totalW.toFixed(2))
            };

            const sheetData = [headerRow, ...rows, totalRow];
            const ws = XLSX.utils.json_to_sheet(sheetData, { skipHeader: true });

            // Apply Advanced Styling
            const range = XLSX.utils.decode_range(ws['!ref']);
            for (let R = range.s.r; R <= range.e.r; ++R) {
              for (let C = range.s.c; C <= range.e.c; ++C) {
                const cell_address = { c: C, r: R };
                const cell_ref = XLSX.utils.encode_cell(cell_address);
                if (!ws[cell_ref]) continue;

                // Base style for all cells
                ws[cell_ref].s = {
                  font: { name: "Calibri", sz: 11 },
                  alignment: { vertical: "center", horizontal: "left" },
                  border: {
                    top: { style: "thin", color: { rgb: "000000" } },
                    bottom: { style: "thin", color: { rgb: "000000" } },
                    left: { style: "thin", color: { rgb: "000000" } },
                    right: { style: "thin", color: { rgb: "000000" } }
                  }
                };

                // 1. Header Styling (Blue background, White bold text)
                if (R === 0) {
                  ws[cell_ref].s.fill = { fgColor: { rgb: "2563EB" } }; // Blue-600
                  ws[cell_ref].s.font = { color: { rgb: "FFFFFF" }, bold: true, sz: 11 };
                  ws[cell_ref].s.alignment.horizontal = "center";
                }



                // 4. Number Formatting (Thousand separators for T and W columns)
                if (R > 0 && (C === 4 || C === 5)) {
                  ws[cell_ref].z = '#,##0.00';
                  ws[cell_ref].s.alignment.horizontal = "right";
                }
              }
            }



            // Add Auto-Filters
            ws['!autofilter'] = { ref: ws['!ref'] };

            // Auto-fit Column Widths
            const colWidths = [
              { wch: Math.max(String(headerRow.F || '').length, ...rows.map(r => String(r.F || '').length), 5) + 2 },
              { wch: Math.max(String(headerRow.H || '').length, ...rows.map(r => String(r.H || '').length), 5) + 2 },
              { wch: Math.max(String(headerRow.L || '').length, ...rows.map(r => String(r.L || '').length), 5) + 2 },
              { wch: Math.max(String(headerRow.P || '').length, ...rows.map(r => String(r.P || '').length), 5) + 2 },
              { wch: Math.max(String(headerRow.T || '').length, ...rows.map(r => String(r.T || '').length), String(totalRow.T).length, 5) + 2 },
              { wch: Math.max(String(headerRow.W || '').length, ...rows.map(r => String(r.W || '').length), String(totalRow.W).length, 5) + 2 },
            ];
            ws['!cols'] = colWidths;

            return ws;
          };

          // Helper to create the Summary sheet with 3 tables
          const buildSummarySheet = (over13, all13) => {
            const dcs = ['BG', 'BP', 'PZ', 'NI'];
            const ws = XLSX.utils.aoa_to_sheet([[]]);

            // 1. Calculate Stats
            const overStats = dcs.map(dc => {
              const rows = over13.filter(r => String(r.L || '').trim().toUpperCase() === dc);
              return [dc, rows.reduce((s, r) => s + r.T, 0), rows.reduce((s, r) => s + r.W, 0)];
            });
            const tOverM = overStats.reduce((s, r) => s + r[1], 0);
            const tOverP = overStats.reduce((s, r) => s + r[2], 0);

            const allStats = dcs.map(dc => {
              const rows = all13.filter(r => String(r.L || '').trim().toUpperCase() === dc);
              return [dc, rows.reduce((s, r) => s + r.T, 0), rows.reduce((s, r) => s + r.W, 0)];
            });
            const tAllM = allStats.reduce((s, r) => s + r[1], 0);
            const tAllP = allStats.reduce((s, r) => s + r[2], 0);

            // 2. Build Tables
            const table1 = [['FINAL LTL', '', ''], ['DC', 'Bruto masa', 'Palete'], ...allStats, ['TOTAL', tAllM, tAllP]];
            const table2 = [['LTL>1000kg', '', ''], ['DC', 'Bruto masa', 'Palete'], ...overStats, ['TOTAL', tOverM, tOverP]];

            // For Table 3, calculate % for each DC = over_mass / all_mass for that DC
            const percRows = overStats.map((r, i) => {
              const allMass = allStats[i][1];
              return [r[0], allMass > 0 ? r[1] / allMass : 0];
            });
            const totalPerc = tAllM > 0 ? tOverM / tAllM : 0;
            const table3 = [['LTL>1000kg(%)', ''], ['DC', '%'], ...percRows, ['TOTAL', totalPerc]];

            XLSX.utils.sheet_add_aoa(ws, table1, { origin: 'A10' });
            XLSX.utils.sheet_add_aoa(ws, table2, { origin: 'E10' });
            XLSX.utils.sheet_add_aoa(ws, table3, { origin: 'I10' });

            // 3. Styling
            const range = XLSX.utils.decode_range(ws['!ref']);
            for (let R = range.s.r; R <= range.e.r; ++R) {
              for (let C = range.s.c; C <= range.e.c; ++C) {
                const addr = XLSX.utils.encode_cell({ c: C, r: R });
                if (!ws[addr]) continue;

                const isTitleRow = R === 9;
                const isHeader = R === 10;
                const isTotal = R === 15;
                const isDC = (C === 0 || C === 4 || C === 8) && ['BG', 'BP', 'PZ', 'NI'].includes(String(ws[addr].v).toUpperCase());
                const isNumber = typeof ws[addr].v === 'number';
                const isTotalLabel = isTotal && (C === 0 || C === 4 || C === 8);

                ws[addr].s = {
                  font: { name: "Calibri", sz: 11, bold: isHeader || isTotal || isDC || isNumber },
                  alignment: { vertical: "center", horizontal: (C === 0 || C === 4 || C === 8) ? "left" : "right" }
                };

                if (!isTitleRow) {
                  ws[addr].s.border = {
                    top: { style: "thin", color: { rgb: "000000" } },
                    bottom: { style: "thin", color: { rgb: "000000" } },
                    left: { style: "thin", color: { rgb: "000000" } },
                    right: { style: "thin", color: { rgb: "000000" } }
                  };
                }

                if (isHeader || isDC || isTotalLabel) {
                  ws[addr].s.fill = { fgColor: { rgb: "2563EB" } };
                  ws[addr].s.font.color = { rgb: "FFFFFF" };
                  if (isHeader) ws[addr].s.alignment.horizontal = "center";
                } else if (isTotal) {
                  ws[addr].s.fill = { fgColor: { rgb: "F1F5F9" } };
                }

                if (R > 10) {
                  if ([1, 2, 5, 6].includes(C)) ws[addr].z = '#,##0.00';
                  if (C === 9) ws[addr].z = '0%';
                }
              }
            }

            ws['!cols'] = [{ wch: 8 }, { wch: 15 }, { wch: 10 }, { wch: 5 }, { wch: 8 }, { wch: 15 }, { wch: 10 }, { wch: 5 }, { wch: 8 }, { wch: 10 }];
            return ws;
          };

          // 0. Summary Sheet (First)
          XLSX.utils.book_append_sheet(newWorkbook, buildSummarySheet(dataRows, all13), 'TOTAL-LTL');

          // 1. Main Sheet (13-LTL)
          XLSX.utils.book_append_sheet(newWorkbook, createSheet(dataRows), '13-LTL');

          // 2. DC Specific Sheets (Filtered from 13-LTL)
          const dcList = ['BG', 'BP', 'PZ', 'NI'];
          dcList.forEach(dc => {
            const dcRows = dataRows.filter(row => String(row.L || '').trim().toUpperCase() === dc);
            XLSX.utils.book_append_sheet(newWorkbook, createSheet(dcRows), dc);
          });

          // 3. 14 FTL and 16 MINI FTL Sheets
          XLSX.utils.book_append_sheet(newWorkbook, createSheet(all14), '14 FTL');
          XLSX.utils.book_append_sheet(newWorkbook, createSheet(all16), '16 MINI FTL');

          // Export and download
          const today = new Date().toISOString().split('T')[0];
          XLSX.writeFile(newWorkbook, `LTL-${today}.xlsx`);

          setProcessing(false);
        } catch (err) {
          setError('Greška pri obradi datoteke. Uverite se da je to ispravna Excel datoteka.');
          setProcessing(false);
        }
      };

      reader.onerror = () => {
        setError('Greška pri čitanju datoteke.');
        setProcessing(false);
      };

      reader.readAsArrayBuffer(file);
    } catch (err) {
      setError('Došlo je do neočekivane greške.');
      setProcessing(false);
    }
  };

  return (
    <div className="w-full max-w-xl bg-white rounded-3xl shadow-2xl p-10 border-4 border-emerald-500 ring-8 ring-emerald-50">
      <header className="text-center mb-10">
        <h1 className="text-4xl font-black text-slate-900 tracking-tighter mb-2 uppercase">
          LTL izvestaj
        </h1>
        <p className="text-slate-500 text-lg">
          Otpremite, obradite i preuzmite svoje podatke odmah.
        </p>
      </header>

      <main className="space-y-8">
        {/* Drag & Drop Zone */}
        <div
          onDragOver={handleDragOver}
          onDrop={handleDrop}
          className={`
            relative group cursor-pointer
            border-2 border-dashed rounded-2xl p-12
            flex flex-col items-center justify-center
            transition-all duration-300 ease-in-out
            ${file
              ? 'border-emerald-200 bg-emerald-50'
              : 'border-slate-200 hover:border-indigo-400 hover:bg-indigo-50/30'}
          `}
        >
          <input
            type="file"
            accept=".xlsx, .xls"
            onChange={handleFileChange}
            className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
          />

          <div className={`
            w-16 h-16 rounded-full flex items-center justify-center mb-4 transition-transform duration-300 group-hover:scale-110
            ${file ? 'bg-emerald-100 text-emerald-600' : 'bg-indigo-100 text-indigo-600'}
          `}>
            {file ? (
              <svg xmlns="http://www.w3.org/2000/svg" className="h-8 w-8" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
              </svg>
            ) : (
              <svg xmlns="http://www.w3.org/2000/svg" className="h-8 w-8" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
              </svg>
            )}
          </div>

          <div className="text-center">
            {file ? (
              <div className="animate-in fade-in slide-in-from-bottom-2 duration-300">
                <p className="text-emerald-700 font-semibold text-lg truncate max-w-xs mx-auto">
                  {file.name}
                </p>
                <p className="text-emerald-600/70 text-sm">
                  Datoteka je spremna za obradu
                </p>
              </div>
            ) : (
              <>
                <p className="text-slate-700 font-medium text-lg">
                  Prevucite i otpustite vašu Excel datoteku ovde
                </p>
                <p className="text-slate-400 text-sm mt-1">
                  ili kliknite da izaberete sa računara
                </p>
              </>
            )}
          </div>
        </div>

        {/* Error Message */}
        {error && (
          <div className="bg-rose-50 border border-rose-100 text-rose-600 px-4 py-3 rounded-xl text-sm flex items-start gap-3 animate-in fade-in zoom-in-95 duration-200">
            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 flex-shrink-0" viewBox="0 0 20 20" fill="currentColor">
              <path fillRule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7 4a1 1 0 11-2 0 1 1 0 012 0zm-1-9a1 1 0 00-1 1v4a1 1 0 102 0V6a1 1 0 00-1-1z" clipRule="evenodd" />
            </svg>
            <span>{error}</span>
          </div>
        )}

        {/* Action Button */}
        <button
          onClick={handleProcess}
          disabled={!file || processing}
          className={`
            w-full py-4 px-6 rounded-2xl font-bold text-lg shadow-lg transition-all duration-300 cursor-pointer
            flex items-center justify-center gap-3
            ${!file || processing
              ? 'bg-slate-100 text-slate-400 cursor-not-allowed'
              : 'bg-indigo-600 text-white hover:bg-indigo-700 hover:shadow-indigo-200 active:scale-[0.98]'}
          `}
        >
          {processing ? (
            <>
              <svg className="animate-spin h-5 w-5 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
              </svg>
              Obrađuje se...
            </>
          ) : (
            <>
              <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
              </svg>
              Obradi i preuzmi
            </>
          )}
        </button>
      </main>
    </div>
  );
}
