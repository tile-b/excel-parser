import XLSX from 'xlsx-js-style';

/* ------------------------------------------------------------------ */
/* Input format (fixed: same columns, DCs and SIFOJs on every export)  */
/* ------------------------------------------------------------------ */

/** 0-based index of the header row in the source sheet (Excel row 10). */
export const HEADER_ROW_INDEX = 9;

/** Source columns used by the report, looked up by header name. */
const COLUMNS = {
  brugdk: 'BRUGDK',
  status: 'STTUGDK',
  naziv: 'NAZIV_MES_ISP',
  mesto: 'MESTO_MES_ISP',
  dc: 'DC',
  otprema: 'OTPREMA',
  sifoj: 'SIFOJ',
  masa: 'BRUTOMASA',
  palete: 'PALETE',
};

/** Distribution centers, in the order they appear in every report table. */
const DCS = ['BG', 'BP', 'PZ', 'NI'];

/** Only shipments with this STTUGDK status are included. */
const VALID_STATUS = 'O';

/** LTL shipments at or above this gross mass (kg) go into 13-LTL / "LTL>1000kg". */
const LTL_MIN_MASS = 1000;

/** Columns (and their order) shown on the detail sheets. */
const DETAIL_FIELDS = ['naziv', 'mesto', 'dc', 'sifoj', 'masa', 'palete'];
const DETAIL_NUMERIC = ['masa', 'palete'];

/* ------------------------------------------------------------------ */
/* Styling                                                             */
/* ------------------------------------------------------------------ */

const BLUE = '2563EB';
const WHITE = 'FFFFFF';
const GREY = 'F1F5F9';
const GREY_SUBTOTAL = 'D9D9D9';
const NUM_FMT = '#,##0.00';
const INT_NUM_FMT = '#,##0';
const PCT_FMT = '0%';

const thinBorder = () => {
  const side = () => ({ style: 'thin', color: { rgb: '000000' } });
  return { top: side(), bottom: side(), left: side(), right: side() };
};

/* ------------------------------------------------------------------ */
/* Helpers                                                             */
/* ------------------------------------------------------------------ */

const text = (v) => String(v ?? '').trim();
const num = (v) => parseFloat(v) || 0;
const round2 = (v) => Number(v.toFixed(2));
const sum = (rows, field) => rows.reduce((s, r) => s + r[field], 0);
const dcOf = (row) => text(row.dc).toUpperCase();
const byOtprema = (prefix) => (row) => text(row.otprema).startsWith(prefix);

/** SIFOJ numeric codes mapped to their descriptive names. */
export const SIFOJ_NAMES = {
  '3800': 'Skladište BP',
  '4840': 'DC Beograd',
  '6800': 'Skladište Bujanovac',
};

/** Formats SIFOJ code with dash divider and meaning, e.g. 3800-Skladište BP */
export const formatSifoj = (sifoj) => {
  let code = text(sifoj);
  if (!code) return '';
  if (code.endsWith('.0')) code = code.slice(0, -2);
  const name = SIFOJ_NAMES[code];
  return name ? `${code}-${name}` : code;
};

/** User-facing error (message is shown in the UI). */
const userError = (message) => Object.assign(new Error(message), { userMessage: message });

/* ------------------------------------------------------------------ */
/* Parsing                                                             */
/* ------------------------------------------------------------------ */

/**
 * Maps the raw `sheet_to_json({ header: 'A' })` output to named fields.
 * Columns are resolved by header name, so column order in the export doesn't matter.
 */
const parseRows = (jsonData) => {
  const [rawHeader = {}, ...data] = jsonData;

  const letters = {};
  const missing = [];
  Object.entries(COLUMNS).forEach(([field, name]) => {
    const letter = Object.keys(rawHeader).find((k) => text(rawHeader[k]).toUpperCase() === name);
    if (letter) letters[field] = letter;
    else missing.push(name);
  });
  if (missing.length) {
    throw userError(`Datoteka nema očekivane kolone: ${missing.join(', ')} (zaglavlje se očekuje u redu ${HEADER_ROW_INDEX + 1}).`);
  }

  const headerNames = Object.fromEntries(Object.keys(COLUMNS).map((f) => [f, rawHeader[letters[f]]]));
  const rows = data.map((raw) => Object.fromEntries(Object.keys(COLUMNS).map((f) => [f, raw[letters[f]]])));

  return { headerNames, rows };
};

/** Merges item lines into one row per shipment (BRUGDK), summing mass and pallets. */
const groupByShipment = (rows) => {
  const shipments = new Map();
  rows.forEach((row) => {
    const key = text(row.brugdk);
    if (!key) return;

    const existing = shipments.get(key);
    if (existing) {
      existing.masa += num(row.masa);
      existing.palete += num(row.palete);
    } else {
      shipments.set(key, {
        naziv: row.naziv,
        mesto: row.mesto,
        dc: row.dc,
        sifoj: row.sifoj,
        masa: num(row.masa),
        palete: num(row.palete),
      });
    }
  });

  return [...shipments.values()].map((s) => ({ ...s, masa: round2(s.masa), palete: round2(s.palete) }));
};

/* ------------------------------------------------------------------ */
/* Detail sheets (13-LTL, per DC, 14 FTL, 16 MINI FTL)                 */
/* ------------------------------------------------------------------ */

const buildDetailSheet = (headerNames, rows) => {
  const total = { naziv: 'TOTAL', masa: round2(sum(rows, 'masa')), palete: round2(sum(rows, 'palete')) };
  const aoa = [headerNames, ...rows, total].map((r) => DETAIL_FIELDS.map((f) => r[f]));
  const ws = XLSX.utils.aoa_to_sheet(aoa);

  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let R = range.s.r; R <= range.e.r; ++R) {
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
      if (!cell) continue;

      cell.s = {
        font: { name: 'Calibri', sz: 11 },
        alignment: { vertical: 'center', horizontal: 'left' },
        border: thinBorder(),
      };

      if (R === 0) {
        cell.s.fill = { fgColor: { rgb: BLUE } };
        cell.s.font = { color: { rgb: WHITE }, bold: true, sz: 11 };
        cell.s.alignment.horizontal = 'center';
      } else if (DETAIL_NUMERIC.includes(DETAIL_FIELDS[C])) {
        cell.z = NUM_FMT;
        cell.s.alignment.horizontal = 'right';
      }
    }
  }

  ws['!autofilter'] = { ref: ws['!ref'] };

  // Auto-fit column widths
  ws['!cols'] = DETAIL_FIELDS.map((f) => {
    const lengths = rows.map((r) => String(r[f] || '').length);
    lengths.push(String(headerNames[f] || '').length, 5);
    if (DETAIL_NUMERIC.includes(f)) lengths.push(String(total[f]).length);
    return { wch: Math.max(...lengths) + 2 };
  });

  return ws;
};

/* ------------------------------------------------------------------ */
/* Summary sheet (TOTAL-LTL)                                           */
/* ------------------------------------------------------------------ */

/**
 * Writes a styled table: title row, header row, data rows and a TOTAL row.
 * The first column is the label column (blue). Layout/styling is relative to `origin`,
 * so tables can be placed anywhere without hardcoded row/column indexes.
 *
 * @param {object} ws - worksheet
 * @param {{r:number,c:number}} origin - top-left cell (0-based)
 * @param {object} table
 * @param {string} table.title
 * @param {string[]} table.header
 * @param {Array[]} table.rows
 * @param {Array} table.total
 * @param {Object<number,string>} [table.formats] - number format per column index
 * @param {Object<number,string>} [table.align] - horizontal alignment per column index (default: right)
 * @returns {number} number of rows written
 */
const addTable = (ws, origin, { title, header, rows, total, formats = {}, align = {} }) => {
  const aoa = [[title, ...Array(header.length - 1).fill('')], header, ...rows, total];
  XLSX.utils.sheet_add_aoa(ws, aoa, { origin });

  aoa.forEach((line, i) => {
    const isTitle = i === 0;
    const isHeader = i === 1;
    const isTotal = i === aoa.length - 1;
    const isSubtotal = !isTitle && !isHeader && !isTotal && line[0] === '' && line[1] === '';

    if (isSubtotal) {
      if (!ws['!rows']) ws['!rows'] = [];
      ws['!rows'][origin.r + i] = { hpx: 20 };
    }

    line.forEach((_, j) => {
      const cell = ws[XLSX.utils.encode_cell({ r: origin.r + i, c: origin.c + j })];
      if (!cell) return;

      const isLabel = j === 0 && !isTitle && !isHeader && !isSubtotal;

      cell.s = {
        font: { name: 'Calibri', sz: 11, bold: !isTitle },
        alignment: { vertical: 'center', horizontal: j === 0 ? 'left' : align[j] || 'right' },
      };
      if (isTitle) return;

      cell.s.border = thinBorder();
      if (isHeader || isLabel) {
        cell.s.fill = { fgColor: { rgb: BLUE } };
        cell.s.font.color = { rgb: WHITE };
        if (isHeader) cell.s.alignment.horizontal = 'center';
      } else if (isSubtotal) {
        cell.s.fill = { fgColor: { rgb: GREY_SUBTOTAL } };
      } else if (isTotal) {
        cell.s.fill = { fgColor: { rgb: GREY } };
      }

      if (!isHeader && formats[j]) cell.z = formats[j];
    });
  });

  return aoa.length;
};

/** Mass/pallet totals per DC: [[dc, masa, palete], ...] */
const statsPerDc = (rows) =>
  DCS.map((dc) => {
    const dcRows = rows.filter((r) => dcOf(r) === dc);
    return [dc, sum(dcRows, 'masa'), sum(dcRows, 'palete')];
  });

/** Mass/pallet totals per DC + SIFOJ, plus DC sum total rows: [[dc, sifoj, masa, palete], ...] */
const statsPerDcAndSifoj = (rows) =>
  DCS.flatMap((dc) => {
    const dcRows = rows.filter((r) => dcOf(r) === dc);
    const bySifoj = new Map();
    dcRows.forEach((r) => {
      const s = text(r.sifoj);
      const acc = bySifoj.get(s) || { masa: 0, palete: 0 };
      acc.masa += r.masa;
      acc.palete += r.palete;
      bySifoj.set(s, acc);
    });

    const sifojEntries = [...bySifoj.keys()]
      .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
      .map((s) => [dc, formatSifoj(s), round2(bySifoj.get(s).masa), round2(bySifoj.get(s).palete)]);

    const dcMasa = round2(sum(dcRows, 'masa'));
    const dcPalete = round2(sum(dcRows, 'palete'));
    const totalRow = ['', '', dcMasa, dcPalete];

    if (sifojEntries.length === 0) {
      return [
        [dc, '', 0, 0],
        totalRow,
      ];
    }

    return [...sifojEntries, totalRow];
  });

const buildSummarySheet = (all13, over13) => {
  const ws = XLSX.utils.aoa_to_sheet([[]]);

  const allStats = statsPerDc(all13);
  const overStats = statsPerDc(over13);
  const allMasa = allStats.reduce((s, r) => s + r[1], 0);
  const allPalete = allStats.reduce((s, r) => s + r[2], 0);
  const overMasa = overStats.reduce((s, r) => s + r[1], 0);
  const overPalete = overStats.reduce((s, r) => s + r[2], 0);

  const dcHeader = ['DC', 'Bruto masa', 'Palete'];
  const massFormats = { 1: NUM_FMT, 2: NUM_FMT };

  // Row 1: three tables side by side, starting at Excel row 10
  const top = HEADER_ROW_INDEX;
  const height = addTable(ws, { r: top, c: 0 }, {
    title: 'FINAL LTL',
    header: dcHeader,
    rows: allStats,
    total: ['TOTAL', allMasa, allPalete],
    formats: massFormats,
  });

  addTable(ws, { r: top, c: 4 }, {
    title: `LTL>${LTL_MIN_MASS}kg`,
    header: dcHeader,
    rows: overStats,
    total: ['TOTAL', overMasa, overPalete],
    formats: massFormats,
  });

  addTable(ws, { r: top, c: 8 }, {
    title: `LTL>${LTL_MIN_MASS}kg(%)`,
    header: ['DC', '%'],
    rows: overStats.map(([dc, masa], i) => [dc, allStats[i][1] > 0 ? masa / allStats[i][1] : 0]),
    total: ['TOTAL', allMasa > 0 ? overMasa / allMasa : 0],
    formats: { 1: PCT_FMT },
  });

  // Row 2: FINAL LTL split per SIFOJ, two empty rows below the tables above
  addTable(ws, { r: top + height + 2, c: 0 }, {
    title: 'FINAL LTL PO SIFOJ',
    header: ['DC', 'SIFOJ', 'Bruto masa', 'Palete'],
    rows: statsPerDcAndSifoj(all13),
    total: ['TOTAL', '', allMasa, allPalete],
    formats: { 2: INT_NUM_FMT, 3: INT_NUM_FMT },
    align: { 1: 'center' },
  });

  ws['!cols'] = [12, 28, 14, 12, 8, 15, 10, 5, 8, 10].map((wch) => ({ wch }));
  return ws;
};

/* ------------------------------------------------------------------ */
/* Public API                                                          */
/* ------------------------------------------------------------------ */

/**
 * Builds the complete LTL report workbook and the FINAL LTL PO SIFOJ preview data.
 *
 * @param {Array<object>} jsonData - `XLSX.utils.sheet_to_json(ws, { header: 'A', range: HEADER_ROW_INDEX })`
 * @returns {{ workbook: object, sifojTable: { title: string, header: string[], rows: Array[], total: Array } }}
 */
export const buildLtlReportData = (jsonData) => {
  const { headerNames, rows } = parseRows(jsonData);
  const valid = rows.filter((r) => text(r.status).toUpperCase() === VALID_STATUS);

  const all13 = groupByShipment(valid.filter(byOtprema('13')));
  const over13 = all13.filter((r) => r.masa >= LTL_MIN_MASS);
  const all14 = groupByShipment(valid.filter(byOtprema('14')));
  const all16 = groupByShipment(valid.filter(byOtprema('16')));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, buildSummarySheet(all13, over13), 'TOTAL-LTL');
  XLSX.utils.book_append_sheet(wb, buildDetailSheet(headerNames, over13), '13-LTL');
  DCS.forEach((dc) => {
    XLSX.utils.book_append_sheet(wb, buildDetailSheet(headerNames, over13.filter((r) => dcOf(r) === dc)), dc);
  });
  XLSX.utils.book_append_sheet(wb, buildDetailSheet(headerNames, all14), '14 FTL');
  XLSX.utils.book_append_sheet(wb, buildDetailSheet(headerNames, all16), '16 MINI FTL');

  const sifojRows = statsPerDcAndSifoj(all13).map(([dc, sifoj, masa, palete]) => [
    dc,
    sifoj,
    round2(masa),
    round2(palete),
  ]);
  const totalMasa = round2(sum(all13, 'masa'));
  const totalPalete = round2(sum(all13, 'palete'));

  const sifojTable = {
    title: 'FINAL LTL PO SIFOJ',
    header: ['DC', 'SIFOJ', 'Bruto masa', 'Palete'],
    rows: sifojRows,
    total: ['TOTAL', '', totalMasa, totalPalete],
  };

  return { workbook: wb, sifojTable };
};

export const buildLtlWorkbook = (jsonData) => buildLtlReportData(jsonData).workbook;

/** Reads the source workbook (ArrayBuffer/Uint8Array) and returns `{ workbook, sifojTable }`. */
export const parseLtlFile = (data) => {
  const source = XLSX.read(data, { type: 'array' });
  const sheet = source.Sheets[source.SheetNames[0]];
  const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 'A', range: HEADER_ROW_INDEX });
  return buildLtlReportData(jsonData);
};

/** Reads the source workbook (ArrayBuffer/Uint8Array) and returns the report workbook. */
export const createLtlReport = (data) => parseLtlFile(data).workbook;

export const reportFileName = () => `LTL-${new Date().toISOString().split('T')[0]}.xlsx`;

