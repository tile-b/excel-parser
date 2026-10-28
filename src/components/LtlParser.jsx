import { useState } from 'react';
import XLSX from 'xlsx-js-style';
import { parseLtlFile, reportFileName } from '../lib/ltlReport';

const formatNumber = (value) => {
  const rounded = Math.round(Number(value) || 0);
  return rounded.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
};

const splitSifoj = (sifoj) => {
  const str = String(sifoj || '');
  const dashIndex = str.indexOf('-');
  if (dashIndex === -1) {
    return { codePart: str, namePart: '' };
  }
  return {
    codePart: str.slice(0, dashIndex + 1),
    namePart: str.slice(dashIndex + 1),
  };
};

const formatSifojHtml = (sifoj) => {
  const { codePart, namePart } = splitSifoj(sifoj);
  if (!namePart) return codePart;
  return `<span style="color:#000000;">${codePart}</span><span style="color:#5F6368;">${namePart}</span>`;
};

const renderSifoj = (sifoj) => {
  if (!sifoj) return null;
  const { codePart, namePart } = splitSifoj(sifoj);
  if (!namePart) {
    return <span>{codePart}</span>;
  }
  return (
    <>
      <span className="text-black">{codePart}</span>
      <span className="text-[#5F6368]">{namePart}</span>
    </>
  );
};

const buildClipboardPayload = (table) => {
  const tsvLines = [
    table.header.join('\t'),
    ...table.rows.map(([dc, sifoj, masa, palete]) =>
      [dc, sifoj, formatNumber(masa), formatNumber(palete)].join('\t')
    ),
    [table.total[0], table.total[1], formatNumber(table.total[2]), formatNumber(table.total[3])].join('\t'),
  ];
  const tsv = tsvLines.join('\n');

  const border = 'border:1px solid #000000;padding:6px 10px;font-family:Calibri,sans-serif;font-size:11pt;';
  const headerCells = table.header
    .map(
      (h) =>
        `<th style="${border}background:#2563EB;color:#FFFFFF;font-weight:bold;text-align:center;">${h}</th>`
    )
    .join('');

  const bodyRows = table.rows
    .map(([dc, sifoj, masa, palete]) => {
      const isSubtotal = !dc && !sifoj;
      if (isSubtotal) {
        return `
        <tr style="height:20px;line-height:20px;">
          <td style="${border}background:#D9D9D9;height:20px;padding:0 10px;font-weight:bold;text-align:left;"></td>
          <td style="${border}background:#D9D9D9;height:20px;padding:0 10px;font-weight:bold;text-align:center;"></td>
          <td style="${border}background:#D9D9D9;height:20px;padding:0 10px;font-weight:bold;text-align:right;">${formatNumber(masa)}</td>
          <td style="${border}background:#D9D9D9;height:20px;padding:0 10px;font-weight:bold;text-align:right;">${formatNumber(palete)}</td>
        </tr>`;
      }
      return `
        <tr>
          <td style="${border}background:#2563EB;color:#FFFFFF;font-weight:bold;text-align:left;">${dc}</td>
          <td style="${border}font-weight:bold;text-align:center;">${formatSifojHtml(sifoj)}</td>
          <td style="${border}font-weight:bold;text-align:right;">${formatNumber(masa)}</td>
          <td style="${border}font-weight:bold;text-align:right;">${formatNumber(palete)}</td>
        </tr>`;
    })
    .join('');

  const totalRow = `
    <tr>
      <td style="${border}background:#2563EB;color:#FFFFFF;font-weight:bold;text-align:left;">${table.total[0]}</td>
      <td style="${border}background:#F1F5F9;font-weight:bold;text-align:center;">${table.total[1]}</td>
      <td style="${border}background:#F1F5F9;font-weight:bold;text-align:right;">${formatNumber(table.total[2])}</td>
      <td style="${border}background:#F1F5F9;font-weight:bold;text-align:right;">${formatNumber(table.total[3])}</td>
    </tr>`;

  const html = `<table style="border-collapse:collapse;"><thead><tr>${headerCells}</tr></thead><tbody>${bodyRows}${totalRow}</tbody></table>`;

  return { tsv, html };
};

export default function LtlParser() {
  const [file, setFile] = useState(null);
  const [error, setError] = useState('');
  const [processing, setProcessing] = useState(false);
  const [report, setReport] = useState(null);
  const [copied, setCopied] = useState(false);

  const handleFileChange = (e) => {
    const selectedFile = e.target.files[0];
    validateAndSetFile(selectedFile);
  };

  const validateAndSetFile = async (selectedFile) => {
    setError('');
    setCopied(false);
    if (!selectedFile) return;

    const isExcel =
      selectedFile.name.endsWith('.xlsx') ||
      selectedFile.name.endsWith('.xls') ||
      selectedFile.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
      selectedFile.type === 'application/vnd.ms-excel';

    if (!isExcel) {
      setError('Molimo otpremite važeću Excel datoteku (.xlsx ili .xls)');
      setFile(null);
      setReport(null);
      return;
    }

    setFile(selectedFile);
    setProcessing(true);

    try {
      const data = new Uint8Array(await selectedFile.arrayBuffer());
      setReport(parseLtlFile(data));
    } catch (err) {
      console.error(err);
      setReport(null);
      setError(err.userMessage || 'Greška pri obradi datoteke. Uverite se da je to ispravna Excel datoteka.');
    } finally {
      setProcessing(false);
    }
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
      let currentReport = report;
      if (!currentReport) {
        const data = new Uint8Array(await file.arrayBuffer());
        currentReport = parseLtlFile(data);
        setReport(currentReport);
      }
      XLSX.writeFile(currentReport.workbook, reportFileName());
    } catch (err) {
      console.error(err);
      setError(err.userMessage || 'Greška pri obradi datoteke. Uverite se da je to ispravna Excel datoteka.');
    } finally {
      setProcessing(false);
    }
  };

  const handleCopy = async () => {
    if (!report?.sifojTable) return;
    const { tsv, html } = buildClipboardPayload(report.sifojTable);

    try {
      if (typeof ClipboardItem !== 'undefined' && navigator.clipboard?.write) {
        await navigator.clipboard.write([
          new ClipboardItem({
            'text/plain': new Blob([tsv], { type: 'text/plain' }),
            'text/html': new Blob([html], { type: 'text/html' }),
          }),
        ]);
      } else {
        await navigator.clipboard.writeText(tsv);
      }
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch {
      await navigator.clipboard.writeText(tsv);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    }
  };

  const sifojTable = report?.sifojTable;

  return (
    <div className="w-full max-w-2xl bg-white rounded-3xl shadow-2xl p-10 border-4 border-emerald-500 ring-8 ring-emerald-50">
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
          disabled={!file || processing || !report}
          className={`
            w-full py-4 px-6 rounded-2xl font-bold text-lg shadow-lg transition-all duration-300 cursor-pointer
            flex items-center justify-center gap-3
            ${!file || processing || !report
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

        {/* FINAL LTL PO SIFOJ Preview */}
        {sifojTable && (
          <section className="pt-4 border-t border-slate-200 space-y-3">
            <div className="flex items-center justify-between">
              <h2 className="text-sm font-extrabold tracking-wide text-slate-800 uppercase">
                {sifojTable.title}
              </h2>
              <button
                type="button"
                onClick={handleCopy}
                className={`
                  inline-flex items-center gap-2 px-3.5 py-1.5 rounded-xl text-xs font-bold transition-all cursor-pointer
                  ${copied
                    ? 'bg-emerald-600 text-white shadow-sm'
                    : 'bg-slate-100 text-slate-700 hover:bg-blue-600 hover:text-white'}
                `}
              >
                {copied ? (
                  <>
                    <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2.5} d="M5 13l4 4L19 7" />
                    </svg>
                    Kopirano!
                  </>
                ) : (
                  <>
                    <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M8 16H6a2 2 0 01-2-2V6a2 2 0 012-2h8a2 2 0 012 2v2m-6 12h8a2 2 0 002-2v-8a2 2 0 00-2-2h-8a2 2 0 00-2 2v8a2 2 0 002 2z" />
                    </svg>
                    Kopiraj tabelu
                  </>
                )}
              </button>
            </div>

            <div className="overflow-x-auto rounded-xl border border-slate-300 shadow-sm">
              <table className="w-full border-collapse text-sm font-bold">
                <thead>
                  <tr className="bg-blue-600 text-white">
                    {sifojTable.header.map((col) => (
                      <th
                        key={col}
                        className="border border-slate-900/20 px-3 py-2 text-center"
                      >
                        {col}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {sifojTable.rows.map(([dc, sifoj, masa, palete], idx) => {
                    const isSubtotal = !dc && !sifoj;
                    return (
                      <tr
                        key={`${dc}-${sifoj}-${idx}`}
                        className={isSubtotal ? 'bg-[#D9D9D9] h-[20px]' : 'bg-white'}
                        style={isSubtotal ? { height: '20px' } : undefined}
                      >
                        <td
                          className={`border border-slate-300 ${
                            isSubtotal
                              ? 'bg-[#D9D9D9] text-slate-900 px-3 py-0 h-[20px] leading-[20px]'
                              : 'bg-blue-600 text-white px-3 py-1.5'
                          } text-left font-bold`}
                        >
                          {dc}
                        </td>
                        <td
                          className={`border border-slate-300 text-center font-bold ${
                            isSubtotal ? 'px-3 py-0 h-[20px] leading-[20px]' : 'px-3 py-1.5'
                          }`}
                        >
                          {renderSifoj(sifoj)}
                        </td>
                        <td
                          className={`border border-slate-300 text-right text-slate-900 tabular-nums font-bold ${
                            isSubtotal ? 'px-3 py-0 h-[20px] leading-[20px]' : 'px-3 py-1.5'
                          }`}
                        >
                          {formatNumber(masa)}
                        </td>
                        <td
                          className={`border border-slate-300 text-right text-slate-900 tabular-nums font-bold ${
                            isSubtotal ? 'px-3 py-0 h-[20px] leading-[20px]' : 'px-3 py-1.5'
                          }`}
                        >
                          {formatNumber(palete)}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
                <tfoot>
                  <tr className="bg-slate-100">
                    <td className="bg-blue-600 text-white border border-slate-300 px-3 py-2 text-left">
                      {sifojTable.total[0]}
                    </td>
                    <td className="border border-slate-300 px-3 py-2 text-center text-slate-900">
                      {sifojTable.total[1]}
                    </td>
                    <td className="border border-slate-300 px-3 py-2 text-right text-slate-900 tabular-nums">
                      {formatNumber(sifojTable.total[2])}
                    </td>
                    <td className="border border-slate-300 px-3 py-2 text-right text-slate-900 tabular-nums">
                      {formatNumber(sifojTable.total[3])}
                    </td>
                  </tr>
                </tfoot>
              </table>
            </div>
          </section>
        )}
      </main>
    </div>
  );
}
