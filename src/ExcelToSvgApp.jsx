/* ExcelToSvgApp.jsx
   -------------------------------------------------------
   Web‑app React che converte fogli Excel in una o più
   tabelle grafiche SVG (2048×912 px) con due layout
   “Light” e “Strong”.
   -------------------------------------------------------
*/
import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

// Dimensioni fisse delle immagini SVG
const WIDTH = 2048;
const HEIGHT = 912;

const COLORS = {
  yellow: '#ffd600',
  blue:   '#1d48b4',
  gray:   '#4d4c4c',
  grayRow: '#c9cbd1'
};

// Family con fallback eleganti se il font proprietario non fosse disponibile
const FONT_FAMILY = "'Avenir Next LT Pro', 'Avenir Next', 'Helvetica Neue', Helvetica, Arial, sans-serif";
const FONT_SIZE   = 11; // px

export default function ExcelToSvgApp() {
  const [layout, setLayout]     = useState('light');
  const [svgPages, setSvgPages] = useState([]);
  const [status, setStatus]     = useState('Seleziona un file Excel …');

  /* --------------- Gestione upload ---------------- */
  const onFileChange = useCallback((e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setStatus('Caricamento in corso…');
    const reader = new FileReader();

    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const sheetName = wb.SheetNames[0];
        const sheet = wb.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });

        setStatus('Generazione SVG…');
        const pages = buildSvgPages(json, layout);
        setSvgPages(pages);
        setStatus(`${pages.length} SVG pronti!`);
      } catch (err) {
        console.error(err);
        setStatus('Errore nella lettura del file');
      }
    };
    reader.readAsArrayBuffer(file);
  }, [layout]);

  /* --------------- Download ----------------------- */
  const downloadSvg = useCallback((svgContent, idx) => {
    const blob = new Blob([svgContent], { type: 'image/svg+xml;charset=utf-8' });
    saveAs(blob, `table_${idx + 1}.svg`);
  }, []);

  /* --------------- UI ----------------------------- */
  return (
    <div className="min-h-screen bg-[#ffd600] flex flex-col items-center p-6">
      <div className="w-full max-w-4xl bg-white rounded-2xl shadow-lg p-8 space-y-6">
        {/* Titolo */}
        <h1 className="text-3xl font-bold text-[#1d48b4] font-[Avenir Next LT Pro]">
          Excel ➜ SVG Table Generator
        </h1>

        {/* File input & layout selector */}
        <div className="grid gap-4 md:grid-cols-2">
          <label className="block">
            <span className="text-[#1d48b4] font-semibold">Carica Excel</span>
            <input
              type="file"
              accept=".xlsx,.xls"
              onChange={onFileChange}
              className="block w-full mt-2 cursor-pointer file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-[#1d48b4] file:text-white hover:file:bg-[#16388f]"
            />
          </label>

          <label className="block">
            <span className="text-[#1d48b4] font-semibold">Layout</span>
            <select
              value={layout}
              onChange={(e) => setLayout(e.target.value)}
              className="mt-2 w-full border border-[#1d48b4] rounded-xl p-2 focus:outline-none focus:ring-2 focus:ring-[#1d48b4]"
            >
              <option value="light">Light</option>
              <option value="strong">Strong</option>
            </select>
          </label>
        </div>

        {/* Stato */}
        <p className="text-sm text-[#4d4c4c] italic">{status}</p>

        {/* Bottoni di download */}
        {svgPages.length > 0 && (
          <div className="space-y-4">
            <h2 className="text-xl font-bold text-[#1d48b4]">Download SVG</h2>
            <div className="grid gap-3 md:grid-cols-3">
              {svgPages.map((svg, idx) => (
                <button
                  key={idx}
                  onClick={() => downloadSvg(svg, idx)}
                  className="bg-[#1d48b4] text-white rounded-xl py-2 px-4 hover:bg-[#16388f] active:scale-95 transition-transform"
                >
                  Scarica pagina {idx + 1}
                </button>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

/* -------------------------------------------------------------------------- */
/* ⚙️  Logica di generazione delle pagine SVG                                */
/* -------------------------------------------------------------------------- */
function buildSvgPages(sheetData, layout) {
  if (!sheetData.length) return [];

  const headerHeight = 40;
  const rowHeight = 28;
  const pages = [];
  let currentRows = [];
  let currentHeight = headerHeight;
  const header = sheetData[0];

  const pushPage = () => {
    pages.push(renderSvg(header, currentRows, layout));
    currentRows = [];
    currentHeight = headerHeight;
  };

  for (let i = 1; i < sheetData.length; i++) {
    if (currentHeight + rowHeight > HEIGHT) {
      pushPage();
    }
    currentRows.push(sheetData[i]);
    currentHeight += rowHeight;
  }
  pushPage();
  return pages;
}

function renderSvg(header, rows, layout) {
  const cellPadding = 12;
  const colCount = header.length;
  const colWidth = (WIDTH - cellPadding * 2) / colCount;
  let y = 0;
  const lines = [];

  // CSS interno SVG
  lines.push(`
    <style>
      text { font-family: ${FONT_FAMILY}; font-size: ${FONT_SIZE}px; dominant-baseline: middle; }
      .header { font-weight: bold; }
      .light-line  { stroke: ${COLORS.gray}; stroke-width: 1; }
      .header-line { stroke: ${COLORS.blue}; stroke-width: 1; }
      .strong-line { stroke: ${COLORS.blue}; stroke-width: 1; }
    </style>
  `);

  /* ---------- Header ---------- */
  const headerHeight = 40;
  header.forEach((cell, idx) => {
    const x = cellPadding + idx * colWidth + colWidth / 2;
    lines.push(`<text x="${x}" y="${y + headerHeight / 2}" class="header" text-anchor="middle">${escapeXml(cell)}</text>`);
  });
  lines.push(`<line x1="0" y1="${y + headerHeight}" x2="${WIDTH}" y2="${y + headerHeight}" class="${layout === 'light' ? 'header-line' : 'strong-line'}"/>`);
  y += headerHeight;

  /* ---------- Body rows -------- */
  rows.forEach((row, rowIdx) => {
    if (layout === 'strong' && rowIdx % 2 === 1) {
      lines.push(`<rect x="0" y="${y}" width="${WIDTH}" height="28" fill="${COLORS.grayRow}" />`);
    }

    row.forEach((cell, idx) => {
      const x = cellPadding + idx * colWidth + colWidth / 2;
      const isBold = /^\*\*/.test(cell);
      lines.push(`<text x="${x}" y="${y + 14}" ${isBold ? 'font-weight="bold"' : ''} text-anchor="middle" fill="${COLORS.gray}">${escapeXml(cell)}</text>`);
    });

    lines.push(`<line x1="0" y1="${y + 28}" x2="${WIDTH}" y2="${y + 28}" class="${layout === 'light' ? 'light-line' : 'strong-line'}"/>`);
    y += 28;
  });

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${WIDTH}" height="${HEIGHT}" viewBox="0 0 ${WIDTH} ${HEIGHT}">${lines.join('\n')}</svg>`;
}

function escapeXml(unsafe) {
  if (unsafe == null) return '';
  return String(unsafe).replace(/[&<>"']/g, c => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&apos;'
  }[c]));
}