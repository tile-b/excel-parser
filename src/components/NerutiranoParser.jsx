import { useState } from 'react';
import XLSX from 'xlsx-js-style';

/**
 * Placeholder logic for NERUTIRANO parser.
 * To be updated based on new requirements.
 */
const processData = (data) => {
  if (data.length === 0) return { header: {}, dataRows: [] };

  const rawHeader = data[0];

  // Dynamic header mapping to be robust
  const findKey = (name) => Object.keys(rawHeader).find(k =>
    String(rawHeader[k] || '').trim().toUpperCase() === name.toUpperCase()
  );

  const keys = {
    DC: findKey('DC') || 'C',
    DATUM_NASTANKA: findKey('DATUM NASTANKA DOKUMENTA') || 'E',
    NAZIV_MESTA: findKey('NAZIV MESTA ISPORUKE') || 'L',
    MESTO: findKey('MESTO') || 'N',
    MASA: findKey('Sum of BRUTO MASA') || 'P',
  };

  const header = {
    A: rawHeader[keys.DC],
    B: 'DATUM',
    C: rawHeader[keys.NAZIV_MESTA],
    D: 'BROJ LOKALA',
    E: rawHeader[keys.MESTO],
    F: rawHeader[keys.MASA],
  };

  const groupedData = data.slice(1)
    .filter(row => String(row[keys.DC] || '').trim() !== '(blank)')
    .reduce((acc, row) => {
      const mesto = String(row[keys.MESTO] || '').trim();
      if (!mesto) return acc;

      const nazivMesta = String(row[keys.NAZIV_MESTA] || '').trim();
      const masa = parseFloat(row[keys.MASA]) || 0;

      if (!acc[mesto]) {
        acc[mesto] = {
          A: row[keys.DC],
          B: row[keys.DATUM_NASTANKA],
          C: nazivMesta ? [nazivMesta] : [],
          D: 0,
          E: mesto,
          F: masa
        };
      } else {
        if (nazivMesta && !acc[mesto].C.includes(nazivMesta)) {
          acc[mesto].C.push(nazivMesta);
        }
        acc[mesto].F += masa;
      }
      return acc;
    }, {});

  const dataRows = Object.values(groupedData).map(row => ({
    A: String(row.A || '').replace('NECTAR DOO', '').trim(),
    B: row.B,
    C: row.C.join(' - '),
    D: row.C.length,
    E: `${row.E}, Srbija`,
    F: `${row.E} ${Math.round(row.F)} kg`
  }));

  return { header, dataRows };
};

export default function NerutiranoParser() {
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

          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];

          // Data starts from row 10 (range: 9)
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: "A", range: 9 });

          const { header: headerRow, dataRows } = processData(jsonData);

          const newWorkbook = XLSX.utils.book_new();
          
          // Helper for styling and sheet creation
          const createStyledSheet = (rows) => {
            const sheetData = [headerRow, ...rows];
            const ws = XLSX.utils.json_to_sheet(sheetData, { skipHeader: true });

            const range = XLSX.utils.decode_range(ws['!ref']);
            for (let R = range.s.r; R <= range.e.r; ++R) {
              for (let C = range.s.c; C <= range.e.c; ++C) {
                const cell = XLSX.utils.encode_cell({ c: C, r: R });
                if (!ws[cell]) continue;

                // Base style
                ws[cell].s = {
                  font: { name: "Calibri", sz: 11 },
                  alignment: { vertical: "center", horizontal: "left" },
                  border: {
                    top: { style: "thin", color: { rgb: "000000" } },
                    bottom: { style: "thin", color: { rgb: "000000" } },
                    left: { style: "thin", color: { rgb: "000000" } },
                    right: { style: "thin", color: { rgb: "000000" } }
                  }
                };

                // Header styling (Rose/Red theme)
                if (R === 0) {
                  ws[cell].s.fill = { fgColor: { rgb: "E11D48" } }; // rose-600
                  ws[cell].s.font = { color: { rgb: "FFFFFF" }, bold: true, sz: 11 };
                  ws[cell].s.alignment.horizontal = "center";
                } 
              }
            }

            // Add Auto-Filters
            ws['!autofilter'] = { ref: ws['!ref'] };

            // Auto-fit Column Widths
            const colWidths = [
              { wch: Math.max(String(headerRow.A || '').length, ...rows.map(r => String(r.A || '').length), 5) + 2 },
              { wch: Math.max(String(headerRow.B || '').length, ...rows.map(r => String(r.B || '').length), 5) + 2 },
              { wch: Math.max(String(headerRow.C || '').length, ...rows.map(r => String(r.C || '').length), 5) + 2 },
              { wch: Math.max(String(headerRow.D || '').length, ...rows.map(r => String(r.D || '').length), 5) + 2 },
              { wch: Math.max(String(headerRow.E || '').length, ...rows.map(r => String(r.E || '').length), 5) + 2 },
              { wch: Math.max(String(headerRow.F || '').length, ...rows.map(r => String(r.F || '').length), 5) + 2 },
            ];
            ws['!cols'] = colWidths;

            return ws;
          };

          XLSX.utils.book_append_sheet(newWorkbook, createStyledSheet(dataRows), 'NERUTIRANO');

          const today = new Date().toISOString().split('T')[0];
          XLSX.writeFile(newWorkbook, `NERUTIRANO-${today}.xlsx`);

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
    <div className="w-full max-w-xl bg-white rounded-3xl shadow-2xl p-10 border-4 border-rose-500 ring-8 ring-rose-50">
      <header className="text-center mb-10">
        <h1 className="text-4xl font-black text-slate-900 tracking-tighter mb-2 uppercase">
          NERUTIRANO
        </h1>
        <p className="text-slate-500 text-lg">
          Novi parser sa novim pravilima i kolonama.
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
              ? 'border-rose-200 bg-rose-50'
              : 'border-slate-200 hover:border-rose-400 hover:bg-rose-50/30'}
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
            ${file ? 'bg-rose-100 text-rose-600' : 'bg-slate-100 text-slate-600 group-hover:text-rose-600 group-hover:bg-rose-100'}
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
                <p className="text-rose-700 font-semibold text-lg truncate max-w-xs mx-auto">
                  {file.name}
                </p>
                <p className="text-rose-600/70 text-sm">
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
              : 'bg-rose-600 text-white hover:bg-rose-700 hover:shadow-rose-200 active:scale-[0.98]'}
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
