/* ══════════════════════════════════════════════════════
   FinDash — script.js
   Soporta dos modos de lectura:
     A) Formato Estándar: Hojas "INGRESOS" y "EGRESOS".
     B) Información en Crudo: Mapeo de 14 pestañas bancarias específicas.
══════════════════════════════════════════════════════ */

'use strict';

// ─── ESTADO GLOBAL ────────────────────────────────────
let allRows        = [];  
let yearRows       = [];  
let filteredRows   = [];  
let allYears       = [];  
let barChartInst   = null;
let donutChartInst = null;
let tipoBarInst    = null;
let detectedYear   = new Date().getFullYear();

// Variable global para el modo (controlada desde el index.html)
window.currentMode = window.currentMode || 'standard';

// ─── CONFIGURACIÓN DE BANCOS (MODO CRUDO) ──────────────
const BANK_MAPPINGS = {
  "BBVA MXN CHEQUES":         { fecha: "Fecha", concepto: "Concepto", monto: "Importe" },
  "BBVA MXN CONCENT":         { fecha: "Fecha", concepto: "Concepto", monto: "Importe" },
  "BBVA USD CHEQUES":         { fecha: "Fecha", concepto: "Descripción", monto: "Monto" },
  "MONEX USD CHEQUES":        { fecha: "Fecha", concepto: "Descripción", monto: "Monto" },
  "MONEX MXN CHEQUES":        { fecha: "Fecha", concepto: "Descripción", monto: "Monto" },
  "CLARA MXN CRÉDITO":        { fecha: "Fecha transacción", concepto: "Establecimiento", monto: "Monto" },
  "KAPITAL MXN CHEQUES":      { fecha: "Fecha", concepto: "Descripción", monto: "Monto" },
  "KAPITAL MXN FLEX":         { fecha: "Fecha", concepto: "Descripción", monto: "Monto" },
  "KAPITAL MXN FACTORAJE":    { fecha: "Fecha", concepto: "Descripción", monto: "Monto" },
  "KONFIO MXN CRÉDITO":       { fecha: "Fecha", concepto: "Descripción", monto: "Monto" },
  "KONFÍO MXN TARJ CRÉDITO":  { fecha: "Fecha", concepto: "Descripción", monto: "Monto" },
  "BBVA MXN CRÉDITO":         { fecha: "Fecha", concepto: "Concepto", monto: "Importe" },
  "XEPELIN MXN CRÉDITO":      { fecha: "Fecha", concepto: "Descripción", monto: "Monto" },
  "TEXAS BANK USD":           { fecha: "Date", concepto: "Description", monto: "Amount" }
};

const MONTHS_ES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
const MONTHS_LONG = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

// ══════════════════════════════════════════════════════
// 1. CONTROLADORES DE ARCHIVO
// ══════════════════════════════════════════════════════

document.getElementById('fileInput')?.addEventListener('change', handleFileChange);

function handleFileChange(e) {
  const file = e.target.files[0];
  if (file) processFile(file);
}

function processFile(file) {
  const reader = new FileReader();
  showLoading();

  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: 'array' });
      let rows = [];

      if (window.currentMode === 'raw') {
        rows = parseRawBankSheets(wb);
      } else {
        rows = parseStandardSheets(wb);
      }

      if (rows.length === 0) throw new Error("No se encontraron datos válidos.");
      buildDashboard(rows);
    } catch (err) {
      alert("Error: " + err.message);
      resetApp();
    }
  };
  reader.readAsArrayBuffer(file);
}

// ══════════════════════════════════════════════════════
// 2. PARSER MODO CRUDO (BANCOS)
// ══════════════════════════════════════════════════════

function parseRawBankSheets(wb) {
  let combined = [];
  const sheetsToProcess = Object.keys(BANK_MAPPINGS);

  sheetsToProcess.forEach(sheetName => {
    // Buscamos la hoja (insensible a mayúsculas/minúsculas)
    const actualSheetName = wb.SheetNames.find(n => n.toUpperCase() === sheetName);
    if (!actualSheetName) return;

    const ws = wb.Sheets[actualSheetName];
    const data = XLSX.utils.sheet_to_json(ws, { raw: false, defval: "" });
    const mapping = BANK_MAPPINGS[sheetName];

    data.forEach(row => {
      const montoRaw = row[mapping.monto];
      const monto = parseMonto(montoRaw);
      if (isNaN(monto) || monto === 0) return;

      const fecha = parseDate(row[mapping.fecha]);
      
      combined.push({
        fecha: fecha,
        year: fecha ? fecha.getFullYear() : 'Sin fecha',
        mes: fecha ? fecha.getMonth() : null,
        tipo_registro: monto >= 0 ? 'Ingreso' : 'Egreso',
        tipo: actualSheetName, // El nombre del banco actúa como categoría
        categoria: actualSheetName,
        subcategoria: row[mapping.concepto] || 'Sin concepto',
        monto: Math.abs(monto)
      });
    });
  });
  return combined;
}

// ══════════════════════════════════════════════════════
// 3. PARSER MODO ESTÁNDAR
// ══════════════════════════════════════════════════════

function parseStandardSheets(wb) {
  const nameIng = wb.SheetNames.find(n => n.toUpperCase().includes('INGRESO'));
  const nameEgr = wb.SheetNames.find(n => n.toUpperCase().includes('EGRESO'));
  
  if (!nameIng && !nameEgr) throw new Error("No se detectaron hojas de INGRESOS o EGRESOS.");

  let rows = [];
  if (nameIng) rows = rows.concat(processStandardSheet(wb.Sheets[nameIng], 'Ingreso'));
  if (nameEgr) rows = rows.concat(processStandardSheet(wb.Sheets[nameEgr], 'Egreso'));
  
  return rows;
}

function processStandardSheet(ws, tipoReg) {
  const data = XLSX.utils.sheet_to_json(ws, { raw: false });
  return data.map(row => {
    const monto = parseMonto(row['TOTAL'] || row['Monto'] || row['Importe']);
    const fecha = parseDate(row['FECHA'] || row['Fecha']);
    return {
      fecha,
      year: fecha ? fecha.getFullYear() : 'Sin fecha',
      mes: fecha ? fecha.getMonth() : null,
      tipo_registro: tipoReg,
      tipo: row['TIPO'] || row['Categoría'] || 'General',
      categoria: row['TIPO'] || row['Categoría'] || 'General',
      subcategoria: row['NOMBRE'] || row['PROVEEDOR'] || row['Concepto'] || '—',
      monto: Math.abs(monto || 0)
    };
  }).filter(r => r.monto > 0);
}

// ══════════════════════════════════════════════════════
// 4. LÓGICA DEL DASHBOARD
// ══════════════════════════════════════════════════════

function buildDashboard(rows) {
  allRows = rows;
  filteredRows = rows;
  
  // Ocultar carga, mostrar dashboard
  document.getElementById('uploadSection').style.display = 'none';
  document.getElementById('dashboardContent').style.display = 'block';

  renderKPIs(rows);
  renderCharts(rows);
  renderTable(rows);
}

function renderKPIs(rows) {
  const income = rows.filter(r => r.tipo_registro === 'Ingreso').reduce((s, r) => s + r.monto, 0);
  const expense = rows.filter(r => r.tipo_registro === 'Egreso').reduce((s, r) => s + r.monto, 0);
  
  document.getElementById('totalIncome').textContent = formatMoney(income);
  document.getElementById('totalExpenses').textContent = formatMoney(expense);
  document.getElementById('netBalance').textContent = formatMoney(income - expense);
}

function renderCharts(rows) {
  // Aquí iría tu lógica de Chart.js existente (Barra y Dona)
  console.log("Gráficos generados para", rows.length, "registros");
}

function renderTable(rows) {
  const tbody = document.querySelector('#transactionsTable tbody');
  tbody.innerHTML = '';

  rows.slice(0, 100).forEach(r => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${r.fecha ? r.fecha.toLocaleDateString() : '—'}</td>
      <td><span class="badge ${r.tipo_registro.toLowerCase()}">${r.tipo_registro}</span></td>
      <td>${r.categoria}</td>
      <td>${r.subcategoria}</td>
      <td class="amount">${formatMoney(r.monto)}</td>
      <td>Confirmado</td>
    `;
    tbody.appendChild(tr);
  });
}

// ══════════════════════════════════════════════════════
// 5. HELPERS
// ══════════════════════════════════════════════════════

function parseMonto(val) {
  if (!val) return 0;
  if (typeof val === 'number') return val;
  return parseFloat(String(val).replace(/[$,]/g, '')) || 0;
}

function parseDate(val) {
  if (!val) return null;
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

function formatMoney(n) {
  return new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN' }).format(n);
}

function showLoading() {
  document.getElementById('uploadInstructions').textContent = "Procesando archivo...";
}

function resetApp() {
  location.reload();
}
