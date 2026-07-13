(function () {
  'use strict';

  /* =========================================================================
   * Constants
   * ======================================================================= */

  var MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio',
    'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  var MESES_ABBR = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];

  var VALES_CATEGORIAS = ['Insumos de Fábrica', 'Aceite y Lubricante', 'Embalajes',
    'Elementos de Protección Personal', 'Elementos Limpieza Industrial',
    'Herramientas Pequeñas', 'Gases Industriales', 'Combustible', 'Gastos Comunes'];
  var VALES_CATEGORIAS_NORM = VALES_CATEGORIAS.map(normKey);

  var META_CHATARRA_PLANTA = 0.046;
  // Split de la meta planta en sus dos componentes (máquinas + desorillado Braner),
  // igual al desglose "Meta Maqu. / Meta Bran." del Excel de referencia (2,70%+1,90%=4,60%).
  var META_CHATARRA_MAQUINAS = 0.027;
  var META_CHATARRA_BRANER = 0.019;

  // Reference targets seen identically in both sample workbooks' "no borrar"/"Metas"
  // sheets. Some monthly exports (e.g. BD_REPORTE_OPERACIONES.xlsx) don't include that
  // table at all — when a workbook has no meta table of its own, fall back to these
  // instead of leaving every machine blank. The source workbook's own table, when
  // present, always wins; this only fills in for a workbook that has none.
  var FALLBACK_CHATARRA_META = {
    'Perfiladora 1': 0.045, 'Perfiladora 2': 0.027, 'Perfiladora 4': 0.012, 'Perfiladora 5': 0.024,
    'Perfiladora 7': 0.012, 'Tubera 2': 0.025, 'Tubera 4': 0.045, 'Tubera 5': 0.02
  };
  var FALLBACK_OEE_META = {
    'Perfiladora 1': [0.21, 0.2, 0.17, 0.22, 0.15, 0.15, 0.18, 0.15, 0.17, 0.19, 0.17, 0.21],
    'Perfiladora 2': [0.36, 0.42, 0.39, 0.4, 0.42, 0.38, 0.36, 0.35, 0.36, 0.39, 0.38, 0.37],
    'Perfiladora 4': [0.26, 0.26, 0.26, 0.26, 0.26, 0.26, 0.26, 0.26, 0.26, 0.26, 0.26, 0.26],
    'Perfiladora 5': [0.52, 0.51, 0.52, 0.52, 0.52, 0.51, 0.52, 0.51, 0.51, 0.52, 0.52, 0.52],
    'Perfiladora 7': [0.53, 0.53, 0.53, 0.53, 0.53, 0.53, 0.53, 0.53, 0.53, 0.53, 0.53, 0.53],
    'Tubera 2': [0.56, 0.55, 0.59, 0.59, 0.57, 0.58, 0.57, 0.53, 0.55, 0.57, 0.59, 0.55],
    'Tubera 4': [0.26, 0.19, 0.19, 0.25, 0.21, 0.17, 0.26, 0.14, 0.18, 0.2, 0.18, 0.23],
    'Tubera 5': [0.56, 0.55, 0.56, 0.57, 0.56, 0.56, 0.54, 0.52, 0.55, 0.53, 0.57, 0.52],
    'OEE Planta': [0.514, 0.517, 0.526, 0.529, 0.523, 0.531, 0.513, 0.501, 0.514, 0.514, 0.531, 0.504]
  };
  var MACHINE_CODE_FALLBACK_PREFIX = { 'Perfiladora': 'PE', 'Tubera': 'TB' };

  var BOBINAS_LABELS = {
    desorillado: '__Kg Desorillado - Bobinas Cortadas + Flejes 2ª Pasada',
    bobinasCortadas: '__Cantidad Bobinas Cortadas',
    diasHabiles: '__Días Trabajados por Corte Bobinas',
    promedioDiario: '__Rollos por Día',
    desorilladoPct: '__% Kg Desorillado por Corte Bobinas'
  };

  /* =========================================================================
   * String / number utilities
   * ======================================================================= */

  function normKey(s) {
    if (s == null) return '';
    return String(s)
      .normalize('NFD').replace(/[̀-ͯ]/g, '')
      .toLowerCase()
      .replace(/[^a-z0-9]/g, '');
  }
  var EXCEL_ERROR_RE = /^#(VALUE|NAME|REF|DIV\/0|NULL|NUM|N\/A)[!?]?$/i;
  function trimStr(s) {
    if (s == null) return '';
    var t = String(s).trim();
    // a live SAP export can leave a failed lookup formula as a literal "#VALUE!"
    // (or similar) in a cell — treat it as blank, never as real data, since it's
    // especially dangerous as a join/grouping key (silently merges unrelated rows)
    return EXCEL_ERROR_RE.test(t) ? '' : t;
  }
  function isNum(v) { return typeof v === 'number' && isFinite(v); }
  function toNum(v) {
    if (v == null || v === '' || v === '-') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    var n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
    return isFinite(n) ? n : null;
  }
  function zeroToNull(v) { return (v == null || !isFinite(v) || v === 0) ? null : v; }
  function toClp(v) {
    // Chilean-formatted currency strings like "$3.445" (thousands sep ".") or
    // "$0,17866" (decimal sep ","). plain toNum() would misread "$3.445" as 3.445.
    if (typeof v === 'number') return isFinite(v) ? v : null;
    if (v == null) return null;
    var s = String(v).trim();
    if (!s || s === '-') return null;
    s = s.replace(/[^0-9,.\-]/g, '').replace(/\./g, '').replace(',', '.');
    var n = parseFloat(s);
    return isFinite(n) ? n : null;
  }
  function ratio(num, den) {
    if (num == null || den == null || den === 0) return null;
    return num / den;
  }
  function sum(arr) { var s = 0, has = false; for (var i = 0; i < arr.length; i++) { var v = arr[i]; if (isNum(v)) { s += v; has = true; } } return has ? s : 0; }
  function isMachineLike(name) { return /^(perfiladora|tubera)[0-9]+$/.test(normKey(name)); }
  // ordinary least-squares fit y = a + b*x over {x,y} points — the trend line behind
  // every "forecast" in the Proyecciones tab. Returns null with fewer than 2 points.
  function linreg(points) {
    var n = points.length;
    if (n < 2) return null;
    var sx = 0, sy = 0, sxx = 0, sxy = 0;
    points.forEach(function (p) { sx += p.x; sy += p.y; sxx += p.x * p.x; sxy += p.x * p.y; });
    var denom = n * sxx - sx * sx;
    if (!denom) return { a: sy / n, b: 0 };
    var b = (n * sxy - sx * sy) / denom;
    var a = (sy - b * sx) / n;
    return { a: a, b: b };
  }

  var NF_INT = new Intl.NumberFormat('es-CL', { maximumFractionDigits: 0 });
  var NF_1D = new Intl.NumberFormat('es-CL', { maximumFractionDigits: 1, minimumFractionDigits: 1 });
  function fmtInt(n) { return n == null ? '-' : NF_INT.format(Math.round(n)); }
  function fmt1(n) { return n == null ? '-' : NF_1D.format(n); }
  function fmtPct(n, dec) {
    if (n == null) return '-';
    return (n * 100).toLocaleString('es-CL', { maximumFractionDigits: dec == null ? 1 : dec, minimumFractionDigits: dec == null ? 1 : dec }) + '%';
  }
  function fmtMoney(n) { return n == null ? '-' : '$' + NF_INT.format(Math.round(n)); }
  function fmtSigned(n, formatter) {
    if (n == null) return '-';
    var s = formatter(Math.abs(n));
    return (n < 0 ? '-' : '+') + s;
  }
  function deltaClass(n, goodIsPositive) {
    if (n == null) return 'neutral';
    var good = goodIsPositive ? n >= 0 : n <= 0;
    return good ? 'good' : 'bad';
  }
  function td(text, cls) { return '<td' + (cls ? ' class="' + cls + '"' : '') + '>' + text + '</td>'; }
  function tdv(n, formatter, cls) {
    var classes = [];
    if (n == null) classes.push('dash');
    else if (n < 0) classes.push('neg');
    if (cls) classes.push(cls);
    return '<td class="' + classes.join(' ') + '">' + (n == null ? '-' : formatter(n)) + '</td>';
  }
  function escapeHtml(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  /* =========================================================================
   * Workbook / sheet access helpers
   * ======================================================================= */

  function findSheet(wb, wantKey, opts) {
    opts = opts || {};
    var names = wb.SheetNames;
    // exact normalized match first
    for (var i = 0; i < names.length; i++) {
      if (normKey(names[i]) === wantKey) return wb.Sheets[names[i]];
    }
    if (opts.startsWith) {
      for (var j = 0; j < names.length; j++) {
        if (normKey(names[j]).indexOf(wantKey) === 0) return wb.Sheets[names[j]];
      }
    }
    return null;
  }

  function sheetToAOA(ws) {
    if (!ws) return null;
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  }

  function headerIndex(headerRow) {
    var idx = {};
    (headerRow || []).forEach(function (h, i) {
      if (h == null || h === '') return;
      idx[normKey(h)] = i;
    });
    return idx;
  }

  function monthColIndexes(idx) {
    // returns array of 12 column indexes (or null) matching MESES_ABBR keys
    return MESES_ABBR.map(function (abbr) {
      var k = normKey(abbr);
      return idx.hasOwnProperty(k) ? idx[k] : null;
    });
  }

  /* =========================================================================
   * Sheet parsers
   * ======================================================================= */

  function parseProduccion(aoa) {
    var out = [];
    if (!aoa || aoa.length < 2) return out;
    var idx = headerIndex(aoa[0]);
    var cAno = idx['ano'], cMes = idx['mes'], cMaq = idx['nombremaquina'],
      cFecha = idx['fecha'], cLargo = idx['largo'], cCant = idx['cantproducida'],
      cKgReal = idx['cantkgreal'], cTotalUE = idx['totalunest'], cEsp = idx['espesor'],
      cClas = idx['clasificacionprod1'], cCodMaq = idx['codmaq'], cOt = idx['ot'],
      cDesc = idx['descripcionarticulo'], cFamilia = idx['familia'], cSku = idx['codigoarticulo'];
    if (cAno == null || cMes == null || cMaq == null) return out;
    for (var r = 1; r < aoa.length; r++) {
      var row = aoa[r];
      var maq = trimStr(row[cMaq]);
      if (!maq) continue;
      // some exports embed a per-machine "TOTAL MAQUINA" subtotal row alongside
      // the real transaction rows, with the machine populated but no Fecha —
      // skip it or every sum would be inflated (subtotal + its own detail rows)
      if (cFecha != null && !row[cFecha]) continue;
      out.push({
        anio: toNum(row[cAno]),
        mes: toNum(row[cMes]),
        maquina: maq,
        codigo: cCodMaq != null ? trimStr(row[cCodMaq]) : '',
        fecha: cFecha != null ? row[cFecha] : null,
        ot: cOt != null ? toNum(row[cOt]) : null,
        largo: cLargo != null ? toNum(row[cLargo]) : null,
        cantProducida: cCant != null ? toNum(row[cCant]) : null,
        kgReal: cKgReal != null ? toNum(row[cKgReal]) : null,
        totalUnEst: cTotalUE != null ? toNum(row[cTotalUE]) : null,
        espesor: cEsp != null ? toNum(row[cEsp]) : null,
        clasif: cClas != null ? trimStr(row[cClas]) : '',
        descripcion: cDesc != null ? trimStr(row[cDesc]) : '',
        familia: cFamilia != null ? trimStr(row[cFamilia]) : '',
        sku: cSku != null ? trimStr(row[cSku]) : ''
      });
    }
    return out;
  }

  function parseChatarra(aoa) {
    var out = [];
    if (!aoa || aoa.length < 2) return out;
    var idx = headerIndex(aoa[0]);
    var cAno = idx['ano'], cMes = idx['mes'], cMaq = idx['nommaquina'], cTotalUE = idx['totalunest'],
      cFecha = idx['fechacontprod'], cFamilia = idx['familiacodsolic'], cEcot = idx['ecot'],
      cCodSolic = idx['codsolic'], cDescSolic = idx['descripcodsolic'];
    if (cAno == null || cMes == null || cMaq == null) return out;
    for (var r = 1; r < aoa.length; r++) {
      var row = aoa[r];
      var maq = trimStr(row[cMaq]);
      if (!maq) continue;
      // same defensive skip as Producción: exclude embedded per-machine subtotal rows
      if (cFecha != null && !row[cFecha]) continue;
      out.push({
        anio: toNum(row[cAno]),
        mes: toNum(row[cMes]),
        maquina: maq,
        totalUnEst: cTotalUE != null ? toNum(row[cTotalUE]) : null,
        familia: cFamilia != null ? trimStr(row[cFamilia]) : '',
        fecha: cFecha != null ? row[cFecha] : null,
        ot: cEcot != null ? toNum(row[cEcot]) : null,
        codSolic: cCodSolic != null ? trimStr(row[cCodSolic]) : '',
        descripcion: cDescSolic != null ? trimStr(row[cDescSolic]) : ''
      });
    }
    return out;
  }

  function parseValesConsumo(aoa) {
    var out = [];
    if (!aoa || aoa.length < 2) return out;
    var idx = headerIndex(aoa[0]);
    var cFecha = idx['fechacontab'], cValor = idx['valortotal'],
      cCC = idx['descripccosto'], cCuenta = idx['nombredecuenta'],
      cArt = idx['descripcionarticulo'], cCant = idx['cantidad'], cCU = idx['costounit'];
    if (cFecha == null || cValor == null) return out;
    for (var r = 1; r < aoa.length; r++) {
      var row = aoa[r];
      var fecha = row[cFecha];
      if (!(fecha instanceof Date)) continue;
      var valor = toNum(row[cValor]);
      if (valor == null) continue;
      out.push({
        anio: fecha.getUTCFullYear(),
        mes: fecha.getUTCMonth() + 1,
        t: fecha.getTime(),
        valor: valor,
        ccosto: cCC != null ? trimStr(row[cCC]) : '',
        cuenta: cCuenta != null ? trimStr(row[cCuenta]) : '',
        articulo: cArt != null ? trimStr(row[cArt]) : '',
        cantidad: cCant != null ? toNum(row[cCant]) : null,
        costoUnit: cCU != null ? toNum(row[cCU]) : null
      });
    }
    return out;
  }

  function parseMonthlyMatrixByLabel(aoa, labelKey) {
    // sheets shaped: [#, Año, Descripción/Nom.Maq., Ene..Dic, Total]
    var out = {}; // out[anio][label] = [12 values]
    if (!aoa || aoa.length < 2) return out;
    var idx = headerIndex(aoa[0]);
    var cAno = idx['ano'], cLabel = idx[labelKey];
    var months = monthColIndexes(idx);
    if (cAno == null || cLabel == null) return out;
    for (var r = 1; r < aoa.length; r++) {
      var row = aoa[r];
      var anio = toNum(row[cAno]);
      var label = trimStr(row[cLabel]);
      if (anio == null || !label) continue;
      if (!out[anio]) out[anio] = {};
      out[anio][label] = months.map(function (ci) { return ci == null ? null : toNum(row[ci]); });
    }
    return out;
  }

  function parseOEEPl(aoa) {
    // returns { machines: [...], data: {mes: {metric: {machine: value}}} }
    var res = { machines: [], data: {} };
    if (!aoa || aoa.length < 2) return res;
    var header = aoa[0];
    var machines = [];
    for (var c = 2; c < header.length; c++) {
      var h = trimStr(header[c]);
      if (h) machines.push({ name: h, col: c });
    }
    res.machines = machines.map(function (m) { return m.name; });
    for (var r = 1; r < aoa.length; r++) {
      var row = aoa[r];
      var mes = trimStr(row[0]);
      var metric = trimStr(row[1]);
      if (!mes || !metric) continue;
      if (!res.data[mes]) res.data[mes] = {};
      if (!res.data[mes][metric]) res.data[mes][metric] = {};
      machines.forEach(function (m) {
        res.data[mes][metric][m.name] = toNum(row[m.col]);
      });
    }
    return res;
  }

  function parseHorasAtraso(aoa) {
    var out = {}; // out[machine][mes] = value
    if (!aoa || aoa.length < 2) return out;
    var header = aoa[0];
    var monthCols = [];
    for (var c = 1; c < header.length; c++) {
      var h = trimStr(header[c]);
      var mi = MESES.findIndex(function (m) { return normKey(m) === normKey(h); });
      if (mi >= 0) monthCols.push({ mes: MESES[mi], col: c });
    }
    for (var r = 1; r < aoa.length; r++) {
      var row = aoa[r];
      var maq = trimStr(row[0]);
      if (!maq) continue;
      out[maq] = {};
      monthCols.forEach(function (mc) { out[maq][mc.mes] = toNum(row[mc.col]); });
    }
    return out;
  }

  function parsePptoProduccion(aoa) {
    var out = {}; // out[machine][mes] = value
    if (!aoa || aoa.length < 2) return out;
    var header = aoa[0];
    var monthCols = [];
    for (var c = 1; c < header.length; c++) {
      var h = trimStr(header[c]);
      var mi = MESES.findIndex(function (m) { return normKey(m) === normKey(h); });
      if (mi >= 0) monthCols.push({ mes: MESES[mi], col: c });
    }
    for (var r = 1; r < aoa.length; r++) {
      var row = aoa[r];
      var maq = trimStr(row[0]);
      if (!maq) continue;
      out[maq] = {};
      monthCols.forEach(function (mc) { out[maq][mc.mes] = toNum(row[mc.col]); });
    }
    return out;
  }

  function parsePptoValesConsumo(aoa) {
    // "Ppto Vales de consumo" sheet: one row per (Centro de Costo, Máquina), 12 month
    // budget ($) columns. Returns both the machine-total sum (existing behavior, used
    // for the plant-wide budget comparison) and the breakdown kept by category, since
    // individual cost-center rows (Aceite y Lubricante, Insumos de Fábrica, Embalajes,
    // etc.) are needed for the per-category budget-vs-real analysis.
    var total = {}; // total[machine][mes] = summed budget $ across cost-centers
    var porCategoria = {}; // porCategoria[machine][categoria][mes] = budget $
    if (!aoa || aoa.length < 2) return { total: total, porCategoria: porCategoria };
    var header = aoa[0];
    var monthCols = [];
    for (var c = 2; c < header.length; c++) {
      var h = trimStr(header[c]);
      var mi = MESES.findIndex(function (m) { return normKey(m) === normKey(h); });
      if (mi >= 0) monthCols.push({ mes: MESES[mi], col: c });
    }
    for (var r = 1; r < aoa.length; r++) {
      var row = aoa[r];
      var categoria = trimStr(row[0]);
      var maq = trimStr(row[1]);
      if (!maq) continue;
      if (!total[maq]) {
        total[maq] = {};
        monthCols.forEach(function (mc) { total[maq][mc.mes] = 0; });
      }
      if (categoria) {
        if (!porCategoria[maq]) porCategoria[maq] = {};
        if (!porCategoria[maq][categoria]) {
          porCategoria[maq][categoria] = {};
          monthCols.forEach(function (mc) { porCategoria[maq][categoria][mc.mes] = 0; });
        }
      }
      monthCols.forEach(function (mc) {
        var v = toNum(row[mc.col]);
        if (v == null) return;
        total[maq][mc.mes] += v;
        if (categoria) porCategoria[maq][categoria][mc.mes] += v;
      });
    }
    return { total: total, porCategoria: porCategoria };
  }

  function parseMetas(sheets) {
    // Precise locators (this helper sheet is cluttered with several unrelated
    // mini-tables, so whole-sheet fuzzy scanning is not safe):
    //  (a) OEE-meta matrix: a row that reads ["Máquina","Enero",...,"Diciembre"]
    //      followed by rows of label + 12 numeric values in the SAME columns.
    //  (b) chatarra-meta-estándar: a single cell whose text is exactly "Metas"
    //      followed downward, in that column and the next, by machine + 1 value.
    var oeeMeta = {}; // oeeMeta[label][mes] = value
    var chatarraMeta = {}; // chatarraMeta[machine] = value
    var mesKeys = MESES.map(normKey);

    sheets.forEach(function (aoa) {
      if (!aoa) return;

      for (var r = 0; r < aoa.length; r++) {
        var row = aoa[r] || [];
        if (normKey(row[0]) !== 'maquina') continue;
        var monthCols = mesKeys.map(function (mk) {
          for (var c = 1; c < row.length; c++) { if (normKey(row[c]) === mk) return c; }
          return null;
        });
        if (monthCols.some(function (c) { return c == null; })) continue;
        for (var rr = r + 1; rr < aoa.length; rr++) {
          var label = trimStr((aoa[rr] || [])[0]);
          if (!label) break;
          if (!oeeMeta[label]) oeeMeta[label] = {};
          MESES.forEach(function (m, mi) { oeeMeta[label][m] = toNum((aoa[rr] || [])[monthCols[mi]]); });
        }
      }

      for (var r2 = 0; r2 < aoa.length; r2++) {
        var row2 = aoa[r2] || [];
        for (var c2 = 0; c2 < row2.length; c2++) {
          if (normKey(row2[c2]) !== 'metas') continue;
          for (var rr2 = r2 + 1; rr2 < aoa.length; rr2++) {
            var maq = trimStr((aoa[rr2] || [])[c2]);
            if (!isMachineLike(maq)) break;
            var val = toNum((aoa[rr2] || [])[c2 + 1]);
            if (val != null && chatarraMeta[maq] == null) chatarraMeta[maq] = val;
          }
        }
      }
    });
    return { oeeMeta: oeeMeta, chatarraMeta: chatarraMeta };
  }

  function rechazoRowsFrom(aoa, headerRow) {
    // Parses the espesor/rechazo/material table starting right after headerRow.
    // The sheet isn't exclusive to this table — below it (past a blank-row gap)
    // often sits unrelated content (dropdown helper lists, a "Metas OEE" block,
    // etc.) whose columns can coincidentally parse as numeric espesor/rechazo
    // values. Stop at the first fully-blank row once we've collected real rows,
    // so that trailing content never gets folded in as bogus rechazo entries.
    var out = [];
    var idx = headerIndex(aoa[headerRow]);
    var cEsp = idx['espesor'], cRech = idx['rechazototal'], cMat = idx['material'];
    if (cEsp == null || cRech == null || cMat == null) return out;
    for (var r = headerRow + 1; r < aoa.length; r++) {
      var row = aoa[r];
      if (out.length && (!row || row.every(function (v) { return v == null || v === ''; }))) break;
      if (!row) continue;
      var esp = toNum(row[cEsp]), rech = toNum(row[cRech]), mat = trimStr(row[cMat]);
      if (esp == null || rech == null || !mat) continue;
      out.push({ espesor: esp, rechazo: rech, material: mat });
    }
    return out;
  }
  function parseRechazoTabla(aoa) {
    // 'no borrar'/'Estandar' style table: espesor, rechazo total (%), material
    // (short code). Some workbooks keep it at the top of its own sheet (row 0);
    // others bury it further down inside the "Metas" sheet, below other tables
    // (e.g. starting around row 29) — so the header is searched for row by row
    // instead of assumed to be aoa[0].
    if (!aoa || aoa.length < 2) return [];
    for (var h = 0; h < aoa.length; h++) {
      if (!aoa[h]) continue;
      var idx = headerIndex(aoa[h]);
      if (idx['espesor'] == null || idx['rechazototal'] == null || idx['material'] == null) continue;
      var rows = rechazoRowsFrom(aoa, h);
      if (rows.length) return rows;
    }
    return [];
  }

  function parsePptoInsumos(aoa) {
    // period-aggregate reference table (no date column): item, cost-center/machine,
    // consumption count, and Chilean-formatted price columns ("$3.445" / "$0,17866")
    var out = [];
    if (!aoa || aoa.length < 2) return out;
    var idx = headerIndex(aoa[0]);
    var cArt = idx['descripcionarticulo'], cCC = idx['descripccosto'], cConsumo = idx['consumo'],
      cProm = idx['precioprom'], cUlt = idx['ultpreciocompra'], cMax = idx['preciomax'],
      cIpc = idx['precioipc'], cSKg = idx['skg'];
    if (cArt == null) return out;
    for (var r = 1; r < aoa.length; r++) {
      var row = aoa[r];
      var articulo = trimStr(row[cArt]);
      if (!articulo) continue;
      var consumo = cConsumo != null ? toNum(row[cConsumo]) : null;
      var precioProm = cProm != null ? toClp(row[cProm]) : null;
      out.push({
        articulo: articulo,
        ccosto: cCC != null ? trimStr(row[cCC]) : '',
        consumo: consumo,
        precioProm: precioProm,
        ultPrecioCompra: cUlt != null ? toClp(row[cUlt]) : null,
        precioMax: cMax != null ? toClp(row[cMax]) : null,
        precioIPC: cIpc != null ? toClp(row[cIpc]) : null,
        sKg: cSKg != null ? toClp(row[cSKg]) : null,
        valorEstimado: (consumo != null && precioProm != null) ? consumo * precioProm : null
      });
    }
    return out;
  }

  function insumosRanking(pptoInsumos) {
    var map = {};
    pptoInsumos.forEach(function (r) {
      if (!map[r.articulo]) map[r.articulo] = { articulo: r.articulo, consumo: 0, valorEstimado: 0, ultPrecios: [], ipcs: [] };
      var g = map[r.articulo];
      if (isNum(r.consumo)) g.consumo += r.consumo;
      if (isNum(r.valorEstimado)) g.valorEstimado += r.valorEstimado;
      if (isNum(r.ultPrecioCompra)) g.ultPrecios.push(r.ultPrecioCompra);
      if (isNum(r.precioIPC)) g.ipcs.push(r.precioIPC);
    });
    var list = Object.keys(map).map(function (k) {
      var g = map[k];
      var ultPrecio = g.ultPrecios.length ? sum(g.ultPrecios) / g.ultPrecios.length : null;
      var ipc = g.ipcs.length ? sum(g.ipcs) / g.ipcs.length : null;
      var overrun = (ultPrecio != null && ipc != null && ipc !== 0) ? (ultPrecio - ipc) / ipc : null;
      return {
        articulo: g.articulo, consumo: g.consumo || null, valorEstimado: g.valorEstimado || null,
        ultPrecio: ultPrecio, precioIPC: ipc, overrun: overrun
      };
    });
    var topGasto = list.filter(function (x) { return x.valorEstimado; })
      .sort(function (a, b) { return b.valorEstimado - a.valorEstimado; }).slice(0, 12);
    var topSobrecosto = list.filter(function (x) { return x.overrun != null && x.valorEstimado; })
      .sort(function (a, b) { return b.overrun - a.overrun; }).slice(0, 10);
    return { topGasto: topGasto, topSobrecosto: topSobrecosto };
  }

  /* =========================================================================
   * Workbook loader
   * ======================================================================= */

  function mergeConsolidadoConMensual(consolidado, mensual) {
    // Both sheets are one logical table: the main sheet holds closed months, the "M"
    // sheet holds the in-progress month being filled daily ("es la continuación del
    // otro"). Dedup at (año, mes, máquina) granularity, and when the same month+machine
    // exists in both, keep whichever side reaches the LATER day — the main sheet can
    // hold a stale partial copy of the current month while the M sheet has newer days,
    // and preferring main blindly would silently drop the freshest data.
    function maxDia(rows) {
      var d = 0;
      rows.forEach(function (r) {
        if (r.fecha instanceof Date) { var x = r.fecha.getUTCDate(); if (x > d) d = x; }
      });
      return d || rows.length; // no dates at all → fall back to row count as freshness
    }
    var porClaveMain = {}, porClaveM = {};
    consolidado.forEach(function (r) {
      var k = r.anio + '-' + r.mes + '-' + r.maquina;
      (porClaveMain[k] = porClaveMain[k] || []).push(r);
    });
    mensual.forEach(function (r) {
      var k = r.anio + '-' + r.mes + '-' + r.maquina;
      (porClaveM[k] = porClaveM[k] || []).push(r);
    });
    var out = [];
    Object.keys(porClaveMain).forEach(function (k) {
      var m = porClaveM[k];
      if (m && maxDia(m) > maxDia(porClaveMain[k])) out = out.concat(m);
      else out = out.concat(porClaveMain[k]);
    });
    Object.keys(porClaveM).forEach(function (k) {
      if (!porClaveMain[k]) out = out.concat(porClaveM[k]);
    });
    return out;
  }

  function loadWorkbook(arrayBuffer) {
    var wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });

    var wsProduccion = findSheet(wb, 'produccion');
    var wsChatarra = findSheet(wb, 'chatarra');
    var wsProduccionM = findSheet(wb, 'produccionm');
    var wsChatarraM = findSheet(wb, 'chatarram');
    var wsVales = findSheet(wb, 'valesdeconsumo', { startsWith: true });
    var wsBobinas = findSheet(wb, 'bobinasm');
    var wsProdResumen = findSheet(wb, 'prodresumenanu');
    var wsOEE = findSheet(wb, 'oeepl');
    var wsHoras = findSheet(wb, 'horasdeatraso');
    var wsPptoProd = findSheet(wb, 'pptoproduccion');
    var wsPptoVales = findSheet(wb, 'pptovalesdeconsumo');
    var wsPptoInsumos = findSheet(wb, 'pptoinsumos');
    var wsMetas = findSheet(wb, 'metas');
    var wsNoBorrar = findSheet(wb, 'noborrar');

    var missing = [];
    if (!wsProduccion) missing.push('Produccion');
    if (!wsChatarra) missing.push('Chatarra');
    if (missing.length) {
      throw new Error('No se encontraron las hojas requeridas: ' + missing.join(', ') +
        '. Verifica que el Excel tenga la misma estructura del reporte de operaciones.');
    }

    // "Produccion"/"Chatarra" hold closed, completed months; "Produccion M"/"Chatarra M"
    // hold the current month being filled in day by day from SAP and get folded into the
    // main sheet once the month closes. Use the M sheet only for months not yet present
    // in the main sheet, so an in-progress month still shows up without risking double counts.
    var produccion = mergeConsolidadoConMensual(
      parseProduccion(sheetToAOA(wsProduccion)),
      wsProduccionM ? parseProduccion(sheetToAOA(wsProduccionM)) : []
    );
    var chatarra = mergeConsolidadoConMensual(
      parseChatarra(sheetToAOA(wsChatarra)),
      wsChatarraM ? parseChatarra(sheetToAOA(wsChatarraM)) : []
    );

    // Recover the article behind scrap rows whose "Cód.Solic." formula broke
    // (live SAP exports leave literal "#VALUE!", which trimStr blanks out): the
    // scrap row's EC/OT points at the production order, so the production rows
    // tell us which article that scrap belongs to. Without this, most of the
    // in-progress month's scrap loses its article identity and vanishes from
    // the SKU search even though the kilos exist.
    var otMap = {};
    produccion.forEach(function (p) {
      if (p.ot != null && p.sku && !otMap[p.ot]) otMap[p.ot] = p;
    });
    chatarra.forEach(function (c) {
      if (!c.codSolic && c.ot != null && otMap[c.ot]) {
        var p = otMap[c.ot];
        c.codSolic = p.sku;
        if (!c.descripcion) c.descripcion = p.descripcion;
        if (!c.familia) c.familia = p.familia;
      }
    });

    var vales = wsVales ? parseValesConsumo(sheetToAOA(wsVales)) : [];
    var bobinasM = wsBobinas ? parseMonthlyMatrixByLabel(sheetToAOA(wsBobinas), 'descripcion') : {};
    var prodResumenAnu = wsProdResumen ? parseMonthlyMatrixByLabel(sheetToAOA(wsProdResumen), 'nommaq') : {};
    var oee = wsOEE ? parseOEEPl(sheetToAOA(wsOEE)) : { machines: [], data: {} };
    var horasAtraso = wsHoras ? parseHorasAtraso(sheetToAOA(wsHoras)) : {};
    var pptoProduccion = wsPptoProd ? parsePptoProduccion(sheetToAOA(wsPptoProd)) : {};
    var pptoValesParsed = wsPptoVales ? parsePptoValesConsumo(sheetToAOA(wsPptoVales)) : { total: {}, porCategoria: {} };
    var pptoVales = pptoValesParsed.total;
    var pptoValesPorCategoria = pptoValesParsed.porCategoria;
    var pptoInsumos = wsPptoInsumos ? parsePptoInsumos(sheetToAOA(wsPptoInsumos)) : [];
    var metasCandidates = [wsMetas, wsNoBorrar].filter(Boolean).map(sheetToAOA);
    var metas = parseMetas(metasCandidates);
    if (!Object.keys(metas.chatarraMeta).length) metas.chatarraMeta = FALLBACK_CHATARRA_META;
    if (!Object.keys(metas.oeeMeta).length) {
      var fbOee = {};
      Object.keys(FALLBACK_OEE_META).forEach(function (k) {
        fbOee[k] = {};
        MESES.forEach(function (m, i) { fbOee[k][m] = FALLBACK_OEE_META[k][i]; });
      });
      metas.oeeMeta = fbOee;
    }
    // La tabla vive en su propia hoja ("no borrar"/"Estandar") en algunos Excel;
    // en otros está enterrada más abajo dentro de la hoja "Metas", así que se
    // busca en ambas y se usa la primera que dé resultado.
    var rechazoTabla = [];
    [wsNoBorrar, wsMetas].some(function (ws) {
      if (!ws) return false;
      rechazoTabla = parseRechazoTabla(sheetToAOA(ws));
      return rechazoTabla.length > 0;
    });

    var machines = oee.machines.length ? oee.machines : Array.from(new Set(
      produccion.filter(function (p) { return isMachineLike(p.maquina); }).map(function (p) { return p.maquina; })
    ));

    var machineCodes = {};
    produccion.forEach(function (p) {
      if (p.codigo && !machineCodes[p.maquina]) machineCodes[p.maquina] = p.codigo;
    });

    var groupSet = {};
    produccion.concat(chatarra).forEach(function (p) {
      if (isMachineLike(p.maquina)) groupSet[p.maquina] = true;
    });
    machines.forEach(function (m) { groupSet[m] = true; });
    var groupMachines = Object.keys(groupSet);

    var years = Array.from(new Set(produccion.map(function (p) { return p.anio; }).filter(isNum))).sort();

    return {
      produccion: produccion, chatarra: chatarra, vales: vales, bobinasM: bobinasM,
      prodResumenAnu: prodResumenAnu, oee: oee, horasAtraso: horasAtraso,
      pptoProduccion: pptoProduccion, pptoVales: pptoVales, pptoValesPorCategoria: pptoValesPorCategoria, metas: metas,
      rechazoTabla: rechazoTabla, machines: machines, groupMachines: groupMachines, years: years,
      machineCodes: machineCodes, pptoInsumos: pptoInsumos
    };
  }

  function machineShortLabel(name, codesMap) {
    if (codesMap && codesMap[name]) return codesMap[name];
    var m = /^(Perfiladora|Tubera)\s*([0-9]+)/i.exec(name);
    if (m && MACHINE_CODE_FALLBACK_PREFIX[m[1]]) return MACHINE_CODE_FALLBACK_PREFIX[m[1]] + m[2];
    return name;
  }

  /* =========================================================================
   * Aggregation engine (mirrors the INDAMA workbook formulas)
   * ======================================================================= */

  function makeEngine(D) {
    var inSet = function (set) { var m = {}; set.forEach(function (s) { m[s] = true; }); return function (name) { return !!m[name]; }; };

    function prodRows(anio, mes, maquinaFilter) {
      return D.produccion.filter(function (p) {
        return p.anio === anio && p.mes === mes && maquinaFilter(p.maquina);
      });
    }
    function chatarraRows(anio, mes, maquinaFilter) {
      return D.chatarra.filter(function (p) {
        return p.anio === anio && p.mes === mes && maquinaFilter(p.maquina);
      });
    }

    function prodEstandar(anio, mes, maquinaFilter) {
      return zeroToNull(sum(prodRows(anio, mes, maquinaFilter).map(function (r) { return r.totalUnEst; })));
    }
    function prodReal(anio, mes, maquinaFilter) {
      return zeroToNull(sum(prodRows(anio, mes, maquinaFilter).map(function (r) { return r.kgReal; })));
    }
    function metrosLineales(anio, mes, maquinaFilter, requireFecha) {
      var rows = prodRows(anio, mes, maquinaFilter);
      if (requireFecha) rows = rows.filter(function (r) { return r.fecha != null; });
      var v = sum(rows.map(function (r) { return (r.largo || 0) * (r.cantProducida || 0); })) / 1000;
      return zeroToNull(v);
    }
    function unidadesFabricadas(anio, mes, maquinaFilter) {
      return zeroToNull(sum(prodRows(anio, mes, maquinaFilter).map(function (r) { return r.cantProducida; })));
    }
    function chatarraKg(anio, mes, maquinaFilter) {
      return zeroToNull(sum(chatarraRows(anio, mes, maquinaFilter).map(function (r) { return r.totalUnEst; })));
    }

    function bobinaValor(anio, label, mesIdx) {
      var m = D.bobinasM[anio];
      if (!m || !m[label]) return null;
      // future months are pre-filled with 0 in this sheet, not left blank —
      // treat 0 as "no data yet" so it renders as "-" instead of a fake 0,00%
      return zeroToNull(m[label][mesIdx]);
    }
    function flejeVenta(anio, nombre, mesIdx) {
      var m = D.prodResumenAnu[anio];
      if (!m || !m[nombre]) return null;
      var v = m[nombre][mesIdx];
      return zeroToNull(v);
    }

    function valesValor(maquina, anio, mes) {
      var v = 0, has = false;
      D.vales.forEach(function (r) {
        if (r.anio === anio && r.mes === mes && r.ccosto === maquina && VALES_CATEGORIAS_NORM.indexOf(normKey(r.cuenta)) >= 0) {
          v += r.valor; has = true;
        }
      });
      return has && v !== 0 ? v : null;
    }

    function oeeVal(mesNombre, maquina, metric) {
      var d = D.oee.data[mesNombre];
      if (!d || !d[metric]) return null;
      var v = d[metric][maquina];
      return zeroToNull(v);
    }
    function aporteOEEMes(mesNombre) {
      var d = D.oee.data[mesNombre];
      if (!d || !d['Aporte OEE']) return null;
      var vals = D.machines.map(function (m) { return d['Aporte OEE'][m]; }).filter(isNum);
      return vals.length ? zeroToNull(sum(vals)) : null;
    }

    function horasAtrasoVal(maquina, mesNombre) {
      // mirrors INDEX/MATCH on the source sheet: a found row with a blank
      // month cell reads as 0 (not an error), only a missing row is "-".
      var m = D.horasAtraso[maquina];
      if (!m) return null;
      var v = m[mesNombre];
      return isNum(v) ? v : 0;
    }

    function pptoProdVal(maquina, mesNombre) {
      var m = D.pptoProduccion[maquina];
      if (!m) return null;
      var v = m[mesNombre];
      return isNum(v) ? v : null;
    }
    function pptoTotalVal(mesNombre) {
      // total planta budget row is keyed "Braner" in the source workbook
      return pptoProdVal('Braner', mesNombre) != null ? pptoProdVal('Braner', mesNombre) : pptoProdVal('PLANTA', mesNombre);
    }
    function pptoValesVal(maquina, mesNombre) {
      var m = D.pptoVales[maquina];
      if (!m) return null;
      var v = m[mesNombre];
      return isNum(v) ? v : null;
    }

    /* ---- Costos e Insumos: real vs. presupuesto por categoría (mensual) ----
     * Real = "Vales de Consumo" transactions (Nombre de cuenta = categoría) filtered
     * to the selected machine(s). Presupuesto = "Ppto Vales de consumo" sheet, same
     * categoría/máquina breakdown. Both keyed by the exact category label (e.g.
     * "Aceite y Lubricante", "Insumos de Fábrica", "Embalajes"). */
    function costosCategoria(anio, categoria, opts) {
      opts = opts || {};
      var catNorm = normKey(categoria);
      var maquinas = opts.maquina ? [opts.maquina] : D.machines;
      function pptoForMonth(mesNombre) {
        var total = 0, has = false;
        maquinas.forEach(function (maq) {
          var byCat = D.pptoValesPorCategoria[maq];
          if (!byCat) return;
          Object.keys(byCat).forEach(function (catKey) {
            if (normKey(catKey) !== catNorm) return;
            var v = byCat[catKey][mesNombre];
            if (isNum(v)) { total += v; has = true; }
          });
        });
        return has ? total : null;
      }
      var months = MESES.map(function (mesNombre, i) {
        var mesN = i + 1;
        var real = 0, hasReal = false;
        D.vales.forEach(function (r) {
          if (r.anio === anio && r.mes === mesN && normKey(r.cuenta) === catNorm && maquinas.indexOf(r.ccosto) >= 0) {
            real += r.valor; hasReal = true;
          }
        });
        var realVal = hasReal ? real : null;
        var pptoVal = pptoForMonth(mesNombre);
        var desviacion = (realVal != null && pptoVal != null) ? realVal - pptoVal : null;
        var desviacionPct = (desviacion != null && pptoVal) ? desviacion / pptoVal : null;
        return { mes: mesNombre, mesAbbr: MESES_ABBR[i], real: realVal, ppto: pptoVal, desviacion: desviacion, desviacionPct: desviacionPct };
      });
      var realMonths = months.map(function (m) { return m.real; }).filter(isNum);
      var promedioAnual = realMonths.length ? sum(realMonths) / realMonths.length : null;
      return { months: months, promedioAnual: promedioAnual };
    }

    function oeeMetaVal(label, mesNombre) {
      var m = D.metas.oeeMeta[label];
      if (!m) return null;
      var v = m[mesNombre];
      return zeroToNull(v);
    }
    function chatarraMetaEstandarVal(maquina) {
      var v = D.metas.chatarraMeta[maquina];
      return isNum(v) ? v : null;
    }

    var groupFilter = inSet(D.groupMachines);

    /* ---- Resumen mensual PLANTA (rows 18-51 of INDAMA 2026) ---- */
    function resumenPlanta(anio) {
      var meses = MESES.map(function (mesNombre, i) {
        var mesN = i + 1;
        var prodEst = prodEstandar(anio, mesN, groupFilter);
        var prodR = prodReal(anio, mesN, groupFilter);
        var ml = metrosLineales(anio, mesN, groupFilter, true);
        var unidades = unidadesFabricadas(anio, mesN, groupFilter);
        var chatMaq = chatarraKg(anio, mesN, groupFilter);
        var flejeVenta1 = flejeVenta(anio, 'Cortador 2', i);
        var flejeVenta2 = flejeVenta(anio, 'Flejes Venta', i);
        var desorillado = bobinaValor(anio, BOBINAS_LABELS.desorillado, i);
        var bobinasCortadas = bobinaValor(anio, BOBINAS_LABELS.bobinasCortadas, i);
        var diasHabiles = bobinaValor(anio, BOBINAS_LABELS.diasHabiles, i);
        var promDiario = bobinaValor(anio, BOBINAS_LABELS.promedioDiario, i);
        var desorilladoPctRaw = bobinaValor(anio, BOBINAS_LABELS.desorilladoPct, i);
        var desorilladoPct = desorilladoPctRaw == null ? null : desorilladoPctRaw / 100;

        var totalProduccion = zeroToNull(sum([prodEst, flejeVenta1, flejeVenta2].filter(isNum)));
        var totalChatarraReal = zeroToNull(sum([chatMaq, desorillado].filter(isNum)));
        var pctTuberasPerf = ratio(chatMaq, sum([prodEst, chatMaq].filter(isNum)));
        var pctChatarraPlanta = (pctTuberasPerf == null && desorilladoPct == null) ? null :
          (pctTuberasPerf || 0) + (desorilladoPct || 0);

        var oeeMensual = aporteOEEMes(mesNombre);

        return {
          mes: mesNombre, mesAbbr: MESES_ABBR[i], mesN: mesN,
          prodEstandar: prodEst, prodReal: prodR, diferencia: (prodEst != null && prodR != null) ? prodEst - prodR : null,
          variacion: (prodEst && prodR != null) ? 1 - prodR / prodEst : null,
          metrosLineales: ml, unidades: unidades, factorKm: ratio(prodEst, ml),
          chatarraMaquinas: chatMaq, chatarraMaquinasPct: pctTuberasPerf,
          desorilladoBranerPct: desorilladoPct, desorilladoKg: desorillado,
          bobinasCortadas: bobinasCortadas, diasHabiles: diasHabiles, promDiario: promDiario,
          totalProduccion: totalProduccion, totalChatarraReal: totalChatarraReal,
          pctChatarraPlanta: pctChatarraPlanta, oeeMensual: oeeMensual
        };
      });
      return meses;
    }

    /* ---- KPI header (rows 1-13) ---- */
    function kpiHeader(anio, resumen) {
      var kilosFabricados = zeroToNull(sum(resumen.map(function (m) { return m.totalProduccion; })));
      var chatarraKgTotal = zeroToNull(sum(resumen.map(function (m) { return m.totalChatarraReal; })));

      var sumChatMaq = sum(resumen.map(function (m) { return m.chatarraMaquinas; }));
      var sumProdEst = sum(resumen.map(function (m) { return m.prodEstandar; }));
      var chatarraMaquinasPct = ratio(sumChatMaq, sumProdEst + sumChatMaq);

      var branerVals = resumen.map(function (m) { return m.desorilladoBranerPct; }).filter(function (v) { return v != null && v !== 0; });
      var chatarraBranerPct = branerVals.length ? sum(branerVals) / branerVals.length : null;

      var chatarraPlantaPct = (chatarraMaquinasPct == null && chatarraBranerPct == null) ? null :
        (chatarraMaquinasPct || 0) + (chatarraBranerPct || 0);
      var desviacionChatarra = chatarraPlantaPct == null ? null : META_CHATARRA_PLANTA - chatarraPlantaPct;

      // meta calculada anual: promedio de la meta por espesor de cada mes, ponderado
      // por kg producidos ese mes (mismo criterio que el OEE ponderado más abajo)
      var numCalc = 0, denCalc = 0;
      resumen.forEach(function (m) {
        var calc = chatarraPorEspesorTarget(anio, m.mes);
        if (isNum(calc)) {
          var w = (m.totalProduccion || 0) + (m.totalChatarraReal || 0);
          numCalc += calc * w; denCalc += w;
        }
      });
      var metaChatarraCalculada = denCalc ? numCalc / denCalc : null;
      var desviacionChatarraCalculada = (metaChatarraCalculada != null && chatarraPlantaPct != null) ?
        metaChatarraCalculada - chatarraPlantaPct : null;
      // meta calculada total = componente de máquinas (arriba) + el estándar de
      // Braner/desorillado (no tiene tabla por espesor propia) — comparable directo
      // contra "Chatarra Planta", que también es la suma de ambos componentes.
      var metaChatarraCalculadaTotal = metaChatarraCalculada != null ? metaChatarraCalculada + META_CHATARRA_BRANER : null;
      var desviacionChatarraCalculadaTotal = (metaChatarraCalculadaTotal != null && chatarraPlantaPct != null) ?
        metaChatarraCalculadaTotal - chatarraPlantaPct : null;

      var numOEE = 0, denOEE = 0;
      resumen.forEach(function (m) {
        if (isNum(m.oeeMensual)) {
          var w = (m.totalProduccion || 0) + (m.totalChatarraReal || 0);
          numOEE += m.oeeMensual * w; denOEE += w;
        }
      });
      var oeePlantaGlobal = denOEE ? numOEE / denOEE : null;

      return {
        oeePlantaGlobal: oeePlantaGlobal, kilosFabricados: kilosFabricados, chatarraKg: chatarraKgTotal,
        chatarraMaquinasPct: chatarraMaquinasPct, chatarraBranerPct: chatarraBranerPct,
        chatarraPlantaPct: chatarraPlantaPct, metaChatarra: META_CHATARRA_PLANTA, desviacionChatarra: desviacionChatarra,
        metaChatarraCalculada: metaChatarraCalculada, desviacionChatarraCalculada: desviacionChatarraCalculada,
        metaChatarraCalculadaTotal: metaChatarraCalculadaTotal, desviacionChatarraCalculadaTotal: desviacionChatarraCalculadaTotal
      };
    }

    /* ---- Target de chatarra dinámico por espesor: promedia el rechazo teórico
     * estándar (D.rechazoTabla, tabla material×espesor de la hoja "Estandar"/"no
     * borrar") ponderado por los kg realmente producidos de cada espesor ese mes —
     * un mes que fabrica más material delgado tiene naturalmente más desperdicio
     * teórico, así que el target se ajusta solo en vez de usar un número fijo todo
     * el año. Si el Excel no trae esa tabla (no todos la traen), devuelve null y
     * quien llama cae de vuelta a la meta fija de planta. Cuando el mismo espesor
     * aparece con más de un material, se promedia su rechazo (no distinguimos
     * material por artículo producido, solo por espesor). ---- */
    function chatarraPorEspesorTarget(anio, mesNombre, maquina) {
      if (!D.rechazoTabla.length) return null;
      var mesN = MESES.indexOf(mesNombre) + 1;
      var rows = D.produccion.filter(function (p) {
        return p.anio === anio && p.mes === mesN && isNum(p.espesor) && p.espesor > 0 && isNum(p.totalUnEst) &&
          (!maquina || p.maquina === maquina);
      });
      if (!rows.length) return null;
      var rechazoPorEsp = {};
      D.rechazoTabla.forEach(function (r) {
        if (!rechazoPorEsp[r.espesor]) rechazoPorEsp[r.espesor] = [];
        rechazoPorEsp[r.espesor].push(r.rechazo);
      });
      var kgTotal = 0, kgPonderado = 0;
      rows.forEach(function (p) {
        var arr = rechazoPorEsp[p.espesor];
        if (!arr) return; // espesor producido sin estándar de rechazo definido — se excluye del promedio
        var rechazoProm = sum(arr) / arr.length;
        kgPonderado += p.totalUnEst * rechazoProm;
        kgTotal += p.totalUnEst;
      });
      return kgTotal ? kgPonderado / kgTotal : null;
    }

    /* ---- Presupuesto vs real (rows 34-51) ---- */
    function presupuesto(anio, resumen) {
      var produccion = resumen.map(function (m) {
        var real = m.totalProduccion;
        var ppto = real == null ? null : pptoTotalVal(m.mes);
        var desv = (real != null && ppto != null) ? real - ppto : null;
        var desvPct = ratio(desv, ppto);
        return { mes: m.mes, mesAbbr: m.mesAbbr, real: real, ppto: ppto, desv: desv, desvPct: desvPct };
      });
      var chatarra = resumen.map(function (m) {
        var real = m.totalChatarraReal;
        var pct = m.pctChatarraPlanta;
        // "pct"/"real" son planta completa (máquinas + desorillado Braner), así que
        // la meta comparada también tiene que sumar sus dos componentes — no solo
        // la meta calculada de máquinas (que por sí sola solo cubre el rechazo por
        // espesor, nunca el desorillado). Mismo criterio que calidadPlantaDetalle.
        // Se exponen las dos metas (administrativa fija y calculada) con su propia
        // desviación cada una; "meta"/"desv" quedan apuntando a la calculada porque
        // es la que se grafica (la administrativa se muestra solo en la tabla).
        var metaDinamica = chatarraPorEspesorTarget(anio, m.mes);
        var metaMaq = metaDinamica != null ? metaDinamica : META_CHATARRA_MAQUINAS;
        var metaCalculada = pct == null ? null : (metaMaq + META_CHATARRA_BRANER);
        var metaAdministrativa = pct == null ? null : META_CHATARRA_PLANTA;
        var desvCalculada = (metaCalculada != null && pct != null) ? metaCalculada - pct : null;
        var desvAdministrativa = (metaAdministrativa != null && pct != null) ? metaAdministrativa - pct : null;
        return {
          mes: m.mes, mesAbbr: m.mesAbbr, real: real, pct: pct,
          meta: metaCalculada, desv: desvCalculada,
          metaCalculada: metaCalculada, desvCalculada: desvCalculada,
          metaAdministrativa: metaAdministrativa, desvAdministrativa: desvAdministrativa,
          metaEsDinamica: metaDinamica != null
        };
      });
      return { produccion: produccion, chatarra: chatarra };
    }

    /* ---- Calidad Planta: chatarra máquinas + desorillado Braner, cada uno vs.
     * su propia meta (rows "CALIDAD PLANTA" del Excel de referencia) ---- */
    function calidadPlantaDetalle(anio) {
      var resumen = resumenPlanta(anio);
      return resumen.map(function (m) {
        var realMaq = m.chatarraMaquinasPct;
        var realBran = m.desorilladoBranerPct;
        var real = m.pctChatarraPlanta;
        var metaDinamica = chatarraPorEspesorTarget(anio, m.mes);
        var metaMaq = metaDinamica != null ? metaDinamica : META_CHATARRA_MAQUINAS;
        var metaBran = META_CHATARRA_BRANER;
        var metaTotal = metaMaq + metaBran;
        var desvMaq = realMaq != null ? metaMaq - realMaq : null;
        var desvBran = realBran != null ? metaBran - realBran : null;
        var desvTotal = (real != null) ? metaTotal - real : null;
        return {
          mes: m.mes, mesAbbr: m.mesAbbr,
          real: real, esperado: metaTotal,
          realMaquinas: realMaq, metaMaquinas: metaMaq, desvMaquinas: desvMaq,
          realBraner: realBran, metaBraner: metaBran, desvBraner: desvBran,
          desvTotal: desvTotal, metaEsDinamica: metaDinamica != null
        };
      });
    }

    /* ---- Detalle por máquina - mes (rows 108-170) ---- */
    function detalleMes(anio, mesNombre) {
      var mesIdx = MESES.indexOf(mesNombre);
      var mesN = mesIdx + 1;
      var maquinas = D.machines.concat(['Braner']);

      var cols = maquinas.map(function (maq) {
        var isBraner = maq === 'Braner';
        var prodEst, ml, factor, unidades, chat;
        if (isBraner) {
          var m = D.prodResumenAnu[anio];
          prodEst = zeroToNull((m && m['Cortador 2']) ? m['Cortador 2'][mesIdx] : null);
          // fall back: Braner cortado total is tracked via Bobinas M "3_Kg..." style rows; use fleje total if present
          if (prodEst == null) {
            var flejeTotal = D.bobinasM[anio] && D.bobinasM[anio]['1_Kg Bobinas Cortadas'];
            prodEst = flejeTotal ? zeroToNull(flejeTotal[mesIdx]) : null;
          }
          ml = null; factor = null;
          unidades = bobinaValor(anio, BOBINAS_LABELS.bobinasCortadas, mesIdx);
          chat = bobinaValor(anio, BOBINAS_LABELS.desorillado, mesIdx);
        } else {
          var filt = function (name) { return name === maq; };
          prodEst = prodEstandarSingle(anio, mesN, maq);
          ml = metrosLinealesSingle(anio, mesN, maq);
          factor = ratio(prodEst, ml);
          unidades = unidadesFabricadasSingle(anio, mesN, maq);
          chat = chatarraKg(anio, mesN, function (n) { return n === maq; });
        }
        var chatPct = ratio(chat, sum([prodEst, chat].filter(isNum)));
        var metaAdmin = (isBraner || prodEst == null) ? null : chatarraMetaEstandarVal(maq);
        var metaCalc = (isBraner || prodEst == null) ? null : chatarraPorEspesorTarget(anio, mesNombre, maq);
        var metaChatEst = metaCalc != null ? metaCalc : metaAdmin;
        var metaChatEsCalc = metaCalc != null;
        var desvChatEst = (metaChatEst != null && chatPct != null) ? metaChatEst - chatPct : null;

        var valorVales = valesValor(maq, anio, mesN);
        var costoPorM = ratio(valorVales, ml);
        var costoPorKg = (prodEst != null && prodEst !== 0) ? ratio(valorVales, prodEst) : null;
        var metaPresupKg = null;
        var pptoVal = pptoValesVal(maq, mesNombre);
        var pptoProdKg = pptoProdVal(maq === 'Braner' ? 'Braner' : maq, mesNombre);
        if (pptoVal != null && pptoProdKg) metaPresupKg = pptoVal / pptoProdKg;
        var desvCosto = (metaPresupKg != null && costoPorKg != null) ? zeroToNull(metaPresupKg - costoPorKg) : null;

        var disp = isBraner ? null : oeeVal(mesNombre, maq, 'Disponibilidad');
        var cal = isBraner ? null : oeeVal(mesNombre, maq, 'Calidad');
        var ritmo = isBraner ? null : oeeVal(mesNombre, maq, 'Ritmo');
        var oeeM = isBraner ? null : oeeVal(mesNombre, maq, 'OEE');
        var oeeMeta = isBraner ? null : oeeMetaVal(maq, mesNombre);
        var desvOEE = (oeeM != null && oeeMeta != null) ? zeroToNull(oeeM - oeeMeta) : null;
        var horas = horasAtrasoVal(maq, mesNombre);

        return {
          maquina: maq, isBraner: isBraner,
          prodEstandar: prodEst, metrosLineales: ml, factor: factor, unidades: unidades,
          chatarra: chat, chatarraPct: chatPct, metaChatarraEstandar: metaChatEst, metaChatarraEsCalculada: metaChatEsCalc, desviacionChatarraEstandar: desvChatEst,
          metaAdministrativa: metaAdmin, metaCalculada: metaCalc,
          valorVales: valorVales, costoPorMetro: costoPorM, costoPorKilo: costoPorKg, metaPresupuesto: metaPresupKg, desviacionCosto: desvCosto,
          disponibilidad: disp, calidad: cal, ritmo: ritmo, oee: oeeM, oeeMeta: oeeMeta, desviacionOEE: desvOEE,
          horasAtraso: horas
        };
      });

      // planta-weighted cost/kg (F144/F145 pattern), excludes Braner
      var numC = 0, denC = 0, numM = 0;
      cols.forEach(function (c) {
        if (c.isBraner) return;
        if (isNum(c.valorVales) && isNum(c.costoPorKilo)) { numC += c.valorVales * c.costoPorKilo; denC += c.valorVales; }
        if (isNum(c.valorVales) && isNum(c.metaPresupuesto)) numM += c.valorVales * c.metaPresupuesto;
      });
      var costoPlanta = denC ? numC / denC : null;
      var metaPlanta = denC ? numM / denC : null;

      return { mes: mesNombre, columnas: cols, costoPlanta: costoPlanta, metaPlanta: metaPlanta, desviacionCostoPlanta: (costoPlanta != null && metaPlanta != null) ? metaPlanta - costoPlanta : null };
    }

    function prodEstandarSingle(anio, mesN, maq) { return prodEstandar(anio, mesN, function (n) { return n === maq; }); }
    function metrosLinealesSingle(anio, mesN, maq) { return metrosLineales(anio, mesN, function (n) { return n === maq; }, false); }
    function unidadesFabricadasSingle(anio, mesN, maq) { return unidadesFabricadas(anio, mesN, function (n) { return n === maq; }); }

    /* ---- Detalle por máquina - año (rows 176-238) ---- */
    function detalleAnio(anio, maquina) {
      var meses = MESES.map(function (mesNombre, i) {
        var mesN = i + 1;
        var prodEst = prodEstandarSingle(anio, mesN, maquina);
        var ml = metrosLinealesSingle(anio, mesN, maquina);
        var factor = ratio(prodEst, ml);
        var unidades = unidadesFabricadasSingle(anio, mesN, maquina);
        var chat = chatarraKg(anio, mesN, function (n) { return n === maquina; });
        var chatPct = ratio(chat, sum([prodEst, chat].filter(isNum)));
        var metaAdmin = prodEst == null ? null : chatarraMetaEstandarVal(maquina);
        var metaCalc = prodEst == null ? null : chatarraPorEspesorTarget(anio, mesNombre, maquina);
        var metaChatEst = metaCalc != null ? metaCalc : metaAdmin;
        var metaChatEsCalc = metaCalc != null;
        var desvChatEst = (metaChatEst != null && chatPct != null) ? metaChatEst - chatPct : null;

        var valorVales = valesValor(maquina, anio, mesN);
        var costoPorM = ratio(valorVales, ml);
        var costoPorKg = (prodEst != null && prodEst !== 0) ? ratio(valorVales, prodEst) : null;
        var metaPresupKg = null;
        if (prodEst != null) {
          var pptoVal = pptoValesVal(maquina, mesNombre);
          var pptoProdKg = pptoProdVal(maquina, mesNombre);
          if (pptoVal != null && pptoProdKg) metaPresupKg = pptoVal / pptoProdKg;
        }
        var desvCosto = (metaPresupKg != null && costoPorKg != null) ? zeroToNull(metaPresupKg - costoPorKg) : null;

        var disp = oeeVal(mesNombre, maquina, 'Disponibilidad');
        var cal = oeeVal(mesNombre, maquina, 'Calidad');
        var ritmo = oeeVal(mesNombre, maquina, 'Ritmo');
        var oeeM = oeeVal(mesNombre, maquina, 'OEE');
        var oeeMeta = prodEst == null ? null : oeeMetaVal(maquina, mesNombre);
        var desvOEE = (oeeM != null && oeeMeta != null) ? zeroToNull(oeeM - oeeMeta) : null;
        var horas = zeroToNull(horasAtrasoVal(maquina, mesNombre));

        return {
          mes: mesNombre, mesAbbr: MESES_ABBR[i],
          prodEstandar: prodEst, pptoProduccion: pptoProdVal(maquina, mesNombre),
          metrosLineales: ml, factor: factor, unidades: unidades,
          chatarra: chat, chatarraPct: chatPct, metaChatarraEstandar: metaChatEst, metaChatarraEsCalculada: metaChatEsCalc, desviacionChatarraEstandar: desvChatEst,
          metaAdministrativa: metaAdmin, metaCalculada: metaCalc,
          valorVales: valorVales, costoPorMetro: costoPorM, costoPorKilo: costoPorKg, metaPresupuesto: metaPresupKg, desviacionCosto: desvCosto,
          disponibilidad: disp, calidad: cal, ritmo: ritmo, oee: oeeM, oeeMeta: oeeMeta, desviacionOEE: desvOEE,
          horasAtraso: horas
        };
      });

      var oeeVals = meses.map(function (m) { return m.oee; }).filter(isNum);
      var oeePonderado = oeeVals.length ? sum(oeeVals) / oeeVals.length : null;
      return { maquina: maquina, meses: meses, oeePonderado: oeePonderado };
    }

    /* ---- Análisis por espesor (rows 259-283) ---- */
    function espesorAnalysis(anio) {
      var espesores = Array.from(new Set(
        D.produccion.filter(function (p) { return p.anio === anio && isNum(p.espesor) && p.espesor > 0; })
          .map(function (p) { return p.espesor; })
      )).sort(function (a, b) { return a - b; });

      var kgTable = espesores.map(function (esp) {
        var row = { espesor: esp, meses: [] };
        for (var i = 0; i < 12; i++) {
          var mesN = i + 1;
          var kg = sum(D.produccion.filter(function (p) { return p.anio === anio && p.mes === mesN && p.espesor === esp; })
            .map(function (p) { return p.totalUnEst; }));
          row.meses.push(kg);
        }
        return row;
      });
      var totalPorMes = [];
      for (var i = 0; i < 12; i++) {
        var mesN = i + 1;
        totalPorMes.push(sum(D.produccion.filter(function (p) { return p.anio === anio && p.mes === mesN; }).map(function (p) { return p.totalUnEst; })));
      }
      var pctTable = kgTable.map(function (row) {
        return { espesor: row.espesor, meses: row.meses.map(function (v, i) { return totalPorMes[i] ? v / totalPorMes[i] : null; }) };
      });

      return { espesores: espesores, kgTable: kgTable, pctTable: pctTable, totalPorMes: totalPorMes };
    }

    /* ---- Torre de control: estado por máquina en el último mes con datos ---- */
    function estado(actual, meta, higherIsBetter) {
      if (actual == null || meta == null) return 'sindato';
      var good = higherIsBetter ? actual >= meta : actual <= meta;
      return good ? 'ok' : 'alerta';
    }
    function torreControl(anio, mesNombre) {
      var det = detalleMes(anio, mesNombre);
      return det.columnas.map(function (c) {
        return {
          maquina: c.maquina, isBraner: c.isBraner,
          chatarraPct: c.chatarraPct, metaChatarraEstandar: c.metaChatarraEstandar,
          metaAdministrativa: c.metaAdministrativa, metaCalculada: c.metaCalculada,
          chatarraEstado: estado(c.chatarraPct, c.metaChatarraEstandar, false),
          oee: c.oee, oeeMeta: c.oeeMeta,
          oeeEstado: estado(c.oee, c.oeeMeta, true),
          costoPorKilo: c.costoPorKilo, metaPresupuesto: c.metaPresupuesto,
          costoEstado: estado(c.costoPorKilo, c.metaPresupuesto, false)
        };
      });
    }

    /* ---- Aporte OEE por máquina (contribución al OEE de planta) ---- */
    function aporteOEEPorMaquina(mesNombre) {
      return D.machines.map(function (maq) {
        return { maquina: maq, aporte: oeeVal(mesNombre, maq, 'Aporte OEE') };
      });
    }

    /* ---- Mix de productos: todos los gráficos comparten el mismo filtro ---- */
    function prodFilter(anio, opts) {
      var mesN = opts.mes ? MESES.indexOf(opts.mes) + 1 : null;
      return function (p) {
        return p.anio === anio &&
          (mesN == null || p.mes === mesN) &&
          (!opts.maquina || p.maquina === opts.maquina) &&
          (opts.espesor == null || p.espesor === opts.espesor);
      };
    }
    function productMix(anio, opts) {
      opts = opts || {};
      function groupSum(rows, keyFn) {
        var map = {};
        rows.forEach(function (r) {
          var k = keyFn(r);
          if (!k) return;
          if (!map[k]) map[k] = 0;
          if (isNum(r.totalUnEst)) map[k] += r.totalUnEst;
        });
        return Object.keys(map).map(function (k) { return { label: k, kg: map[k] }; })
          .sort(function (a, b) { return b.kg - a.kg; });
      }
      var prodAnio = D.produccion.filter(prodFilter(anio, opts));
      return {
        topProductos: groupSum(prodAnio, function (r) { return r.descripcion; }).slice(0, 20),
        porFamilia: groupSum(prodAnio, function (r) { return r.familia; })
      };
    }

    /* ---- Producción mensual total (kg) por máquina ---- */
    function produccionMensualPorMaquina(anio, opts) {
      opts = opts || {};
      var mesN = opts.mes ? MESES.indexOf(opts.mes) + 1 : null;
      var maquinas = opts.maquina ? [opts.maquina] : D.machines;
      return maquinas.map(function (maq) {
        var values = MESES.map(function (mesNombre, i) {
          if (mesN != null && i + 1 !== mesN) return null;
          return sum(D.produccion.filter(function (p) {
            return p.anio === anio && p.mes === i + 1 && p.maquina === maq &&
              (opts.espesor == null || p.espesor === opts.espesor);
          }).map(function (p) { return p.totalUnEst; }));
        });
        return { maquina: maq, values: values };
      });
    }

    /* ---- Producción por espesor (kg), filtrable por máquina/mes/espesor ---- */
    function espesorKgFiltrado(anio, opts) {
      opts = opts || {};
      var filt = prodFilter(anio, opts);
      var rows = D.produccion.filter(function (p) { return filt(p) && isNum(p.espesor) && p.espesor > 0; });
      var espesores = Array.from(new Set(rows.map(function (p) { return p.espesor; }))).sort(function (a, b) { return a - b; });
      var mesN = opts.mes ? MESES.indexOf(opts.mes) + 1 : null;
      var porMes = espesores.map(function (esp) {
        var values = MESES.map(function (_, i) {
          if (mesN != null && i + 1 !== mesN) return null;
          return sum(rows.filter(function (p) { return p.mes === i + 1 && p.espesor === esp; })
            .map(function (p) { return p.totalUnEst; }));
        });
        return { espesor: esp, values: values, total: sum(values.filter(isNum)) };
      });
      porMes.sort(function (a, b) { return b.total - a.total; });
      return porMes;
    }
    function espesoresDisponibles(anio) {
      return Array.from(new Set(
        D.produccion.filter(function (p) { return p.anio === anio && isNum(p.espesor) && p.espesor > 0; })
          .map(function (p) { return p.espesor; })
      )).sort(function (a, b) { return a - b; });
    }

    /* ---- Proyección de cierre del mes en curso (run-rate) ----
     * The in-progress month arrives day by day via "Producción M". At the pace
     * produced so far (kg / last day with data), project the full-month close and
     * compare against that month's production budget — per machine and plant-wide. */
    function proyeccionCierre(anio) {
      var mesN = null;
      D.produccion.forEach(function (p) {
        if (p.anio === anio && isNum(p.totalUnEst) && p.totalUnEst > 0 && (mesN == null || p.mes > mesN)) mesN = p.mes;
      });
      if (mesN == null) return null;
      var rows = D.produccion.filter(function (p) { return p.anio === anio && p.mes === mesN; });
      var maxDay = 0;
      rows.forEach(function (p) {
        if (p.fecha instanceof Date) {
          var d = p.fecha.getUTCDate();
          if (d > maxDay) maxDay = d;
        }
      });
      var diasMes = new Date(Date.UTC(anio, mesN, 0)).getUTCDate();
      if (!maxDay) maxDay = diasMes;
      var mesNombre = MESES[mesN - 1];
      var factor = diasMes / maxDay;
      var maquinas = D.machines.filter(function (m) { return m !== 'Braner'; }).map(function (maq) {
        var kg = sum(rows.filter(function (p) { return p.maquina === maq; }).map(function (p) { return p.totalUnEst; }));
        var proy = kg * factor;
        var ppto = pptoProdVal(maq, mesNombre);
        var desvPct = (ppto != null && ppto !== 0) ? (proy - ppto) / ppto : null;
        return { maquina: maq, kg: kg, proy: proy, ppto: ppto, desvPct: desvPct,
                 estado: desvPct == null ? 'sindato' : (desvPct >= 0 ? 'ok' : 'alerta') };
      });
      var totalKg = sum(rows.map(function (p) { return p.totalUnEst; }));
      var totalProy = totalKg * factor;
      var totalPpto = pptoTotalVal(mesNombre);
      var totalDesv = (totalPpto != null && totalPpto !== 0) ? (totalProy - totalPpto) / totalPpto : null;
      return {
        mes: mesNombre, mesN: mesN, diaActual: maxDay, diasMes: diasMes, cerrado: maxDay >= diasMes,
        maquinas: maquinas,
        total: { kg: totalKg, proy: totalProy, ppto: totalPpto, desvPct: totalDesv,
                 estado: totalDesv == null ? 'sindato' : (totalDesv >= 0 ? 'ok' : 'alerta') }
      };
    }

    /* ---- Proyecciones y estadística: cierre de mes y de año por tendencia lineal ----
     * Método: los meses cerrados (y el mes en curso, prorrateado al ritmo actual —
     * mismo cálculo que proyeccionCierre) son los puntos "reales"; se ajusta una recta
     * de mínimos cuadrados sobre esos puntos y se usa para proyectar los meses que aún
     * no tienen ningún dato. Kilos y $ (magnitudes que se pueden sumar) se proyectan
     * por separado y el % o el $/kg del cierre anual se derivan de esos totales — nunca
     * promediando razones directamente, que sesga el resultado. Es una proyección
     * simple y transparente (no un modelo de series de tiempo), correcta como primera
     * lectura de "vamos bien o mal para cerrar el mes/año", no como pronóstico fino. */
    function proyeccionAnalitica(anio, opts) {
      opts = opts || {};
      var maquina = opts.maquina || null;

      var mesesReales;
      if (maquina) {
        var det = detalleAnio(anio, maquina);
        mesesReales = det.meses.map(function (m) {
          return {
            mes: m.mes, mesAbbr: m.mesAbbr,
            prodKg: m.prodEstandar, pptoProdKg: m.pptoProduccion,
            chatKg: m.chatarra, chatPct: m.chatarraPct, chatMeta: m.metaChatarraEstandar,
            valorVales: m.valorVales, costoKg: m.costoPorKilo, costoMeta: m.metaPresupuesto,
            oee: m.oee, oeeMeta: m.oeeMeta
          };
        });
      } else {
        var res = resumenPlanta(anio);
        mesesReales = MESES.map(function (mesNombre, i) {
          var dm = detalleMes(anio, mesNombre);
          var valorValesPlanta = zeroToNull(sum(dm.columnas.filter(function (c) { return !c.isBraner; })
            .map(function (c) { return c.valorVales; })));
          var r = res[i];
          var chatMetaCalc = chatarraPorEspesorTarget(anio, mesNombre);
          return {
            mes: mesNombre, mesAbbr: MESES_ABBR[i],
            prodKg: r.prodEstandar, pptoProdKg: pptoTotalVal(mesNombre),
            chatKg: r.totalChatarraReal, chatPct: r.pctChatarraPlanta,
            chatMeta: chatMetaCalc != null ? chatMetaCalc : META_CHATARRA_PLANTA,
            valorVales: valorValesPlanta, costoKg: dm.costoPlanta, costoMeta: dm.metaPlanta,
            oee: r.oeeMensual, oeeMeta: oeeMetaVal('OEE Planta', mesNombre)
          };
        });
      }

      var cierre = proyeccionCierre(anio);
      var mesActualN = cierre ? cierre.mesN : null;
      var enCurso = cierre ? !cierre.cerrado : false;
      var factorDias = enCurso ? cierre.diasMes / cierre.diaActual : 1;

      // kg / $ (extensivas): el mes en curso se prorratea al ritmo actual antes de
      // entrar a la regresión; los meses futuros se extrapolan con la recta ajustada
      function forecastExtensiva(field) {
        var pts = [];
        for (var i = 1; i <= (mesActualN || 0); i++) {
          var v = mesesReales[i - 1][field];
          if (!isNum(v)) continue;
          pts.push({ x: i, y: (i === mesActualN && enCurso) ? v * factorDias : v });
        }
        var reg = linreg(pts);
        var serie = mesesReales.map(function (m, i0) {
          var idx = i0 + 1;
          if (idx <= (mesActualN || 0)) {
            var proy = (idx === mesActualN && enCurso) ? m[field] * factorDias : null;
            return { mes: m.mes, mesAbbr: m.mesAbbr, actual: m[field], proyectado: proy, esFuturo: false };
          }
          var yhat = reg ? Math.max(0, reg.a + reg.b * idx) : null;
          return { mes: m.mes, mesAbbr: m.mesAbbr, actual: null, proyectado: yhat, esFuturo: true };
        });
        var total = 0, has = false;
        serie.forEach(function (s) {
          var v = s.proyectado != null ? s.proyectado : s.actual;
          if (isNum(v)) { total += v; has = true; }
        });
        return { serie: serie, total: has ? total : null, pendiente: reg ? reg.b : null };
      }

      // %/tasas: se regresiona directo sobre el valor observado (más simple; una tasa
      // no se "acumula" mes a mes como sí lo hacen los kg o los $)
      function forecastTasa(field) {
        var pts = [];
        for (var i = 1; i <= (mesActualN || 0); i++) {
          var v = mesesReales[i - 1][field];
          if (isNum(v)) pts.push({ x: i, y: v });
        }
        var reg = linreg(pts);
        return mesesReales.map(function (m, i0) {
          var idx = i0 + 1;
          if (idx <= (mesActualN || 0)) return { mes: m.mes, mesAbbr: m.mesAbbr, actual: m[field], proyectado: null, esFuturo: false };
          var yhat = reg ? reg.a + reg.b * idx : null;
          return { mes: m.mes, mesAbbr: m.mesAbbr, actual: null, proyectado: yhat, esFuturo: true };
        });
      }

      var prod = forecastExtensiva('prodKg');
      var chat = forecastExtensiva('chatKg');
      var vales = forecastExtensiva('valorVales');
      var oeeSerie = forecastTasa('oee');
      var chatPctSerie = forecastTasa('chatPct');
      var costoKgSerie = forecastTasa('costoKg');

      function promedioMeta(field) {
        var vals = mesesReales.map(function (m) { return m[field]; }).filter(isNum);
        return vals.length ? sum(vals) / vals.length : null;
      }
      var pptoAnual = zeroToNull(sum(mesesReales.map(function (m) { return m.pptoProdKg; }).filter(isNum)));
      var chatMetaRef = promedioMeta('chatMeta');
      var costoMetaProm = promedioMeta('costoMeta');
      var oeeMetaProm = promedioMeta('oeeMeta');

      var chatPctAnual = (chat.total != null && prod.total != null && (chat.total + prod.total) > 0) ?
        chat.total / (chat.total + prod.total) : null;
      var costoKgAnual = (vales.total != null && prod.total) ? vales.total / prod.total : null;
      var oeeAnual = (function () {
        var vals = oeeSerie.map(function (s) { return s.proyectado != null ? s.proyectado : s.actual; }).filter(isNum);
        return vals.length ? sum(vals) / vals.length : null;
      })();

      return {
        maquina: maquina, mes: cierre ? cierre.mes : null,
        diaActual: cierre ? cierre.diaActual : null, diasMes: cierre ? cierre.diasMes : null, enCurso: enCurso,
        produccion: prod, chatarra: chat, valorVales: vales,
        oee: { serie: oeeSerie }, chatarraPct: { serie: chatPctSerie }, costoKg: { serie: costoKgSerie },
        anual: {
          prodKg: prod.total, pptoKg: pptoAnual,
          desvProdPct: (prod.total != null && pptoAnual) ? (prod.total - pptoAnual) / pptoAnual : null,
          chatKg: chat.total, chatPct: chatPctAnual, chatMeta: chatMetaRef,
          desvChatPct: (chatPctAnual != null && chatMetaRef != null) ? chatMetaRef - chatPctAnual : null,
          costoKg: costoKgAnual, costoMeta: costoMetaProm,
          desvCostoPct: (costoKgAnual != null && costoMetaProm) ? (costoMetaProm - costoKgAnual) / costoMetaProm : null,
          oee: oeeAnual, oeeMeta: oeeMetaProm,
          desvOeePct: (oeeAnual != null && oeeMetaProm != null) ? oeeAnual - oeeMetaProm : null
        }
      };
    }

    /* ---- Tendencias de planta: costo/kg vs. meta y calidad promedio, por mes ---- */
    function tendenciasPlanta(anio) {
      return MESES.map(function (mesNombre, i) {
        var det = detalleMes(anio, mesNombre);
        var calidades = det.columnas.filter(function (c) { return !c.isBraner && isNum(c.calidad); })
          .map(function (c) { return c.calidad; });
        return {
          mes: mesNombre, mesAbbr: MESES_ABBR[i],
          costoKg: det.costoPlanta, costoMeta: det.metaPlanta,
          costoDesv: (det.costoPlanta != null && det.metaPlanta != null) ? det.metaPlanta - det.costoPlanta : null,
          calidad: calidades.length ? sum(calidades) / calidades.length : null
        };
      });
    }

    /* ---- Puntos críticos: TODAS las desviaciones vs. meta en un solo lugar ----
     * Reúne, para un mes, cada indicador con meta definida — OEE, producción vs.
     * presupuesto, chatarra vs. meta estándar, costo/kg vs. presupuesto y las tres
     * categorías de costo a nivel planta — normalizando la "gravedad" como desviación
     * relativa en la dirección mala, para poder rankear peor-primero entre unidades
     * distintas (kg, %, $/kg). */
    function puntosCriticos(anio, mesNombre, opts) {
      opts = opts || {};
      var det = detalleMes(anio, mesNombre);
      // if the selected month is still in progress, prorate ABSOLUTE budgets (kg, $)
      // to the days elapsed — comparing 9 days of production against a 31-day budget
      // would flag every machine as critical. Ratios (%, $/kg) need no proration.
      var proy = proyeccionCierre(anio);
      var enCurso = proy && proy.mes === mesNombre && !proy.cerrado;
      var factorDias = enCurso ? proy.diaActual / proy.diasMes : 1;
      var items = [];
      function push(grupo, tipo, maquina, real, meta, badWhenHigher, fmtKind) {
        if (real == null || meta == null || meta === 0) return;
        if (opts.maquina && maquina !== opts.maquina) return;
        var rel = (real - meta) / Math.abs(meta);
        var malo = badWhenHigher ? rel : -rel; // >0 = fuera de meta, magnitud = gravedad
        items.push({
          grupo: grupo, tipo: tipo, maquina: maquina, real: real, meta: meta, fmtKind: fmtKind,
          desvRel: rel, gravedad: malo,
          estado: malo <= 0 ? 'ok' : (malo >= 0.2 ? 'critico' : (malo >= 0.08 ? 'alerta' : 'atencion'))
        });
      }
      var sufijo = enCurso ? ' (al día ' + proy.diaActual + ')' : '';
      det.columnas.forEach(function (c) {
        if (!c.isBraner) {
          var ppto = pptoProdVal(c.maquina, mesNombre);
          push('Producción', 'Producción vs Ppto' + sufijo, c.maquina, c.prodEstandar, ppto != null ? ppto * factorDias : null, false, 'kg');
          push('OEE', 'OEE vs Meta', c.maquina, c.oee, c.oeeMeta, false, 'pct1');
        }
        push('Chatarra', 'Chatarra vs Meta', c.maquina, c.chatarraPct, c.metaChatarraEstandar, true, 'pct2');
        push('Costos', 'Costo/kg vs Ppto', c.maquina, c.costoPorKilo, c.metaPresupuesto, true, 'moneypkg');
      });
      ['Aceite y Lubricante', 'Insumos de Fábrica', 'Embalajes'].forEach(function (cat) {
        var cc = costosCategoria(anio, cat, opts.maquina ? { maquina: opts.maquina } : {});
        var m = cc.months.filter(function (x) { return x.mes === mesNombre; })[0];
        if (m) push('Costos', 'Costo ' + cat + sufijo, opts.maquina || 'Planta', m.real, m.ppto != null ? m.ppto * factorDias : null, true, 'money');
      });
      items.sort(function (a, b) { return b.gravedad - a.gravedad; });
      var resumen = { critico: 0, alerta: 0, atencion: 0, ok: 0 };
      items.forEach(function (it) { resumen[it.estado]++; });
      return { items: items, resumen: resumen, mes: mesNombre, enCurso: enCurso, diaActual: enCurso ? proy.diaActual : null, diasMes: enCurso ? proy.diasMes : null };
    }

    /* ---- Paretos 80/20: qué artículos concentran la chatarra / la producción ---- */
    function paretoDe(rows, anio, opts) {
      opts = opts || {};
      var mesN = opts.mes ? MESES.indexOf(opts.mes) + 1 : null;
      var map = {};
      rows.forEach(function (c) {
        if (c.anio !== anio || !c.descripcion) return;
        if (mesN != null && c.mes !== mesN) return;
        if (opts.maquina && c.maquina !== opts.maquina) return;
        if (!isNum(c.totalUnEst) || c.totalUnEst <= 0) return;
        map[c.descripcion] = (map[c.descripcion] || 0) + c.totalUnEst;
      });
      var items = Object.keys(map).map(function (k) { return { label: k, kg: map[k] }; })
        .sort(function (a, b) { return b.kg - a.kg; });
      var total = sum(items.map(function (it) { return it.kg; }));
      if (!total) return { items: [], total: 0, n80: 0 };
      var acum = 0, n80 = 0;
      items.forEach(function (it, i) {
        acum += it.kg;
        it.share = it.kg / total;
        it.cum = acum / total;
        if (n80 === 0 && it.cum >= 0.8) n80 = i + 1;
      });
      return { items: items, total: total, n80: n80 || items.length };
    }
    function paretoChatarra(anio, opts) { return paretoDe(D.chatarra, anio, opts); }
    function paretoProduccion(anio, opts) { return paretoDe(D.produccion, anio, opts); }

    /* ---- Análisis PxQ de insumos: primer precio vs. último, efecto precio × cantidad.
     * Fuente: transacciones de Vales de Consumo (fecha, artículo, cantidad, costo unit.),
     * cruzadas con el precio IPC de la hoja Ppto Insumos cuando existe. ---- */
    function insumosPxQ(anio, opts) {
      opts = opts || {};
      var mesN = opts.mes ? MESES.indexOf(opts.mes) + 1 : null;
      var categoriaKey = opts.categoria ? normKey(opts.categoria) : null;
      var map = {};
      D.vales.forEach(function (r) {
        if (r.anio !== anio || !r.articulo) return;
        if (mesN != null && r.mes !== mesN) return;
        if (opts.maquina && r.ccosto !== opts.maquina) return;
        if (categoriaKey && normKey(r.cuenta) !== categoriaKey) return;
        var g = map[r.articulo];
        if (!g) g = map[r.articulo] = { articulo: r.articulo, cantidad: 0, gasto: 0, tMin: null, tMax: null, pPrimero: null, pUltimo: null };
        if (isNum(r.cantidad)) g.cantidad += r.cantidad;
        if (isNum(r.valor)) g.gasto += r.valor;
        if (isNum(r.costoUnit) && r.t != null) {
          if (g.tMin == null || r.t < g.tMin) { g.tMin = r.t; g.pPrimero = r.costoUnit; }
          if (g.tMax == null || r.t > g.tMax) { g.tMax = r.t; g.pUltimo = r.costoUnit; }
        }
      });
      var ipcPor = {};
      (D.pptoInsumos || []).forEach(function (r) {
        var k = normKey(r.articulo);
        if (ipcPor[k] == null && isNum(r.precioIPC)) ipcPor[k] = r.precioIPC;
      });
      var items = Object.keys(map).map(function (k) {
        var g = map[k];
        var deltaPrecio = (isNum(g.pPrimero) && g.pPrimero !== 0 && isNum(g.pUltimo)) ? (g.pUltimo - g.pPrimero) / g.pPrimero : null;
        // efecto precio: cuánto del gasto se explica por el cambio de precio (ΔP × Q)
        var efectoPrecio = (deltaPrecio != null) ? (g.pUltimo - g.pPrimero) * g.cantidad : null;
        var ipc = ipcPor[normKey(g.articulo)];
        var vsIpc = (isNum(ipc) && ipc !== 0 && isNum(g.pUltimo)) ? (g.pUltimo - ipc) / ipc : null;
        return {
          articulo: g.articulo, cantidad: g.cantidad || null, gasto: g.gasto || null,
          pPrimero: g.pPrimero, pUltimo: g.pUltimo, deltaPrecio: deltaPrecio,
          efectoPrecio: efectoPrecio, precioIPC: isNum(ipc) ? ipc : null, vsIpc: vsIpc
        };
      });
      items.sort(function (a, b) { return (b.gasto || 0) - (a.gasto || 0); });
      return items;
    }

    /* ---- Movimiento de precio y cantidad por ítem, mes objetivo vs. mes base -----
     * mes base = el primer mes del año con vales de ese ítem (ej. si compró por
     * primera vez en enero, ese es el base, sin importar el filtro de mes actual).
     * mes objetivo = el mes filtrado, o el último mes con datos si el filtro es
     * "Todos" (año completo). El precio de cada mes es un promedio ponderado
     * (gasto ÷ cantidad de ese mes), no el precio de una sola transacción — más
     * robusto si un ítem se compró varias veces en el mismo mes a precios distintos.
     * precioIPC es una referencia de inflación, no un presupuesto real por ítem
     * (el Excel solo trae presupuesto a nivel de categoría × máquina × mes). ---- */
    function insumosMovimiento(anio, opts) {
      opts = opts || {};
      var mesObjetivoFiltro = opts.mes ? MESES.indexOf(opts.mes) + 1 : null;
      var categoriaKey = opts.categoria ? normKey(opts.categoria) : null;
      var porArticulo = {};
      D.vales.forEach(function (r) {
        if (r.anio !== anio || !r.articulo) return;
        if (opts.maquina && r.ccosto !== opts.maquina) return;
        if (categoriaKey && normKey(r.cuenta) !== categoriaKey) return;
        var g = porArticulo[r.articulo];
        if (!g) g = porArticulo[r.articulo] = { articulo: r.articulo, meses: {} };
        var gm = g.meses[r.mes];
        if (!gm) gm = g.meses[r.mes] = { cantidad: 0, gasto: 0 };
        if (isNum(r.cantidad)) gm.cantidad += r.cantidad;
        if (isNum(r.valor)) gm.gasto += r.valor;
      });

      var ipcPor = {};
      (D.pptoInsumos || []).forEach(function (r) {
        var k = normKey(r.articulo);
        if (ipcPor[k] == null && isNum(r.precioIPC)) ipcPor[k] = r.precioIPC;
      });

      function precioMes(gm) { return (gm && isNum(gm.cantidad) && gm.cantidad !== 0) ? gm.gasto / gm.cantidad : null; }

      var items = Object.keys(porArticulo).map(function (k) {
        var g = porArticulo[k];
        var mesesConDatos = Object.keys(g.meses).map(Number).sort(function (a, b) { return a - b; });
        if (!mesesConDatos.length) return null;
        var mesBaseN = mesesConDatos[0];
        var mesObjetivoN = mesObjetivoFiltro != null ? mesObjetivoFiltro : mesesConDatos[mesesConDatos.length - 1];
        var base = g.meses[mesBaseN];
        var obj = g.meses[mesObjetivoN];
        var precioBase = precioMes(base), precioObj = precioMes(obj);
        var esPrimeraCompra = mesObjetivoN === mesBaseN;
        var deltaPrecio = (!esPrimeraCompra && isNum(precioBase) && precioBase !== 0 && isNum(precioObj)) ? (precioObj - precioBase) / precioBase : null;
        var deltaCantidad = (!esPrimeraCompra && base && obj && base.cantidad) ? (obj.cantidad - base.cantidad) / base.cantidad : null;
        var ipc = ipcPor[normKey(g.articulo)];
        var vsIpc = (isNum(ipc) && ipc !== 0 && isNum(precioObj)) ? (precioObj - ipc) / ipc : null;
        // gasto/cantidad del período visible: el mes filtrado, o el año completo si no hay filtro
        var cantidadPeriodo = 0, gastoPeriodo = 0;
        (mesObjetivoFiltro != null ? [mesObjetivoFiltro] : mesesConDatos).forEach(function (m) {
          var gm = g.meses[m];
          if (gm) { cantidadPeriodo += gm.cantidad; gastoPeriodo += gm.gasto; }
        });
        return {
          articulo: g.articulo, mesBase: MESES[mesBaseN - 1], mesObjetivo: MESES[mesObjetivoN - 1],
          esPrimeraCompra: esPrimeraCompra, sinComprasMesObjetivo: mesObjetivoFiltro != null && !obj,
          cantidadBase: base ? base.cantidad : null, cantidadObjetivo: obj ? obj.cantidad : null,
          precioBase: precioBase, precioObjetivo: precioObj,
          deltaPrecio: deltaPrecio, deltaCantidad: deltaCantidad,
          precioIPC: isNum(ipc) ? ipc : null, vsIpc: vsIpc,
          cantidad: cantidadPeriodo || null, gasto: gastoPeriodo || null
        };
      }).filter(Boolean);

      items.sort(function (a, b) { return (b.gasto || 0) - (a.gasto || 0); });
      return items;
    }

    /* ---- Buscador de SKU: producción vs. chatarra por artículo × máquina ---- */
    // Identity key: the article DESCRIPTION (normalized), falling back to the code
    // when a row has no description. Producción keys articles by "Código Artículo"
    // and Chatarra by "Cód.Solic.", and the same physical article can carry a
    // different code on each side (verified in the source data) — joining by code
    // split such articles into a fake "0% chatarra" production row plus a fake
    // "100% chatarra" scrap row. Descriptions match on both sides.
    function artKey(descripcion, codigo) {
      var d = normKey(descripcion);
      return d || String(codigo);
    }
    function skuBuscador(anio, opts) {
      opts = opts || {};
      var mesN = opts.mes ? MESES.indexOf(opts.mes) + 1 : null;
      var prodRows = D.produccion.filter(function (p) {
        return p.anio === anio && p.sku && (mesN == null || p.mes === mesN) && (!opts.maquina || p.maquina === opts.maquina);
      });
      var chatRows = D.chatarra.filter(function (c) {
        return c.anio === anio && c.codSolic && (mesN == null || c.mes === mesN) && (!opts.maquina || c.maquina === opts.maquina);
      });
      // one row per artículo × máquina — production and its corresponding chatarra don't
      // always land in the exact same "Mes" cell (the scrap requisition can be logged a
      // few days after the production run, sometimes crossing a month boundary), so
      // grouping by month as well used to split one article's totals into two incomplete
      // rows. The "Mes" filter above already scopes both sides to a single month.
      var map = {};
      prodRows.forEach(function (p) {
        var key = artKey(p.descripcion, p.sku) + '|' + p.maquina;
        if (!map[key]) map[key] = { sku: p.sku, descripcion: p.descripcion, familia: p.familia, maquina: p.maquina, prodKg: 0, chatKg: 0 };
        if (isNum(p.totalUnEst)) map[key].prodKg += p.totalUnEst;
      });
      chatRows.forEach(function (c) {
        var key = artKey(c.descripcion, c.codSolic) + '|' + c.maquina;
        if (!map[key]) map[key] = { sku: c.codSolic, descripcion: c.descripcion, familia: c.familia, maquina: c.maquina, prodKg: 0, chatKg: 0 };
        if (isNum(c.totalUnEst)) map[key].chatKg += c.totalUnEst;
      });
      var rows = Object.keys(map).map(function (k) {
        var r = map[k];
        // keep prodKg/chatKg as real 0s (not "-") — a SKU with production but no
        // recorded scrap is a genuine 0% chatarra, not missing data, and vice versa
        var total = r.prodKg + r.chatKg;
        r.chatPct = total > 0 ? r.chatKg / total : null;
        r.mesLabel = opts.mes || 'Año completo';
        return r;
      });
      if (opts.query) {
        var q = normKey(opts.query);
        rows = rows.filter(function (r) {
          return normKey(String(r.sku)).indexOf(q) >= 0 || normKey(r.descripcion).indexOf(q) >= 0;
        });
      }
      rows.sort(function (a, b) { return (b.prodKg || 0) - (a.prodKg || 0); });
      return rows;
    }

    /* ---- Tendencia mensual de chatarra para el filtro actual del buscador SKU ---- */
    function skuMonthlyTrend(anio, opts) {
      opts = opts || {};
      var q = opts.query ? normKey(opts.query) : null;
      return MESES.map(function (mesNombre, i) {
        var mesN = i + 1;
        var prodRows = D.produccion.filter(function (p) {
          if (p.anio !== anio || p.mes !== mesN || !p.sku) return false;
          if (opts.maquina && p.maquina !== opts.maquina) return false;
          if (q && normKey(String(p.sku)).indexOf(q) < 0 && normKey(p.descripcion).indexOf(q) < 0) return false;
          return true;
        });
        var chatRows = D.chatarra.filter(function (c) {
          if (c.anio !== anio || c.mes !== mesN || !c.codSolic) return false;
          if (opts.maquina && c.maquina !== opts.maquina) return false;
          if (q && normKey(String(c.codSolic)).indexOf(q) < 0 && normKey(c.descripcion).indexOf(q) < 0) return false;
          return true;
        });
        var prodKg = sum(prodRows.map(function (p) { return p.totalUnEst; }));
        var chatKg = sum(chatRows.map(function (c) { return c.totalUnEst; }));
        var total = prodKg + chatKg;
        return {
          mes: mesNombre, mesAbbr: MESES_ABBR[i],
          prodKg: prodKg, chatKg: chatKg,
          chatPct: total > 0 ? chatKg / total : null
        };
      });
    }

    /* ---- Una serie por artículo (no agregado) para el gráfico de burbujas ---- */
    function skuArticleMonthlyTrend(anio, opts) {
      opts = opts || {};
      var q = opts.query ? normKey(opts.query) : null;
      var mesN = opts.mes ? MESES.indexOf(opts.mes) + 1 : null;
      function matchProd(p) {
        if (p.anio !== anio || !p.sku) return false;
        if (opts.maquina && p.maquina !== opts.maquina) return false;
        if (mesN != null && p.mes !== mesN) return false;
        if (q && normKey(String(p.sku)).indexOf(q) < 0 && normKey(p.descripcion).indexOf(q) < 0) return false;
        return true;
      }
      function matchChat(c) {
        if (c.anio !== anio || !c.codSolic) return false;
        if (opts.maquina && c.maquina !== opts.maquina) return false;
        if (mesN != null && c.mes !== mesN) return false;
        if (q && normKey(String(c.codSolic)).indexOf(q) < 0 && normKey(c.descripcion).indexOf(q) < 0) return false;
        return true;
      }
      var prodAll = D.produccion.filter(matchProd);
      var chatAll = D.chatarra.filter(matchChat);

      // group both sides by the article's description key (same identity rule as
      // skuBuscador) so production and scrap combine even when their codes differ
      var labelOf = {};
      prodAll.forEach(function (p) { var k = artKey(p.descripcion, p.sku); if (!labelOf[k]) labelOf[k] = p.descripcion || String(p.sku); });
      chatAll.forEach(function (c) { var k = artKey(c.descripcion, c.codSolic); if (!labelOf[k]) labelOf[k] = c.descripcion || String(c.codSolic); });

      var chatTotals = {};
      chatAll.forEach(function (c) {
        var k = artKey(c.descripcion, c.codSolic);
        chatTotals[k] = (chatTotals[k] || 0) + (isNum(c.totalUnEst) ? c.totalUnEst : 0);
      });
      // every article with recorded chatarra for the current filter — no top-N cap
      var allKeys = Object.keys(labelOf).filter(function (k) { return (chatTotals[k] || 0) > 0; });
      var orderedKeys = allKeys.slice().sort(function (a, b) { return (chatTotals[b] || 0) - (chatTotals[a] || 0); });

      var series = orderedKeys.map(function (k) {
        var months = MESES.map(function (mesNombre, i) {
          var mesN = i + 1;
          var prodKg = sum(prodAll.filter(function (p) { return artKey(p.descripcion, p.sku) === k && p.mes === mesN; }).map(function (p) { return p.totalUnEst; }));
          var chatKg = sum(chatAll.filter(function (c) { return artKey(c.descripcion, c.codSolic) === k && c.mes === mesN; }).map(function (c) { return c.totalUnEst; }));
          var total = prodKg + chatKg;
          return { mes: mesNombre, mesAbbr: MESES_ABBR[i], chatKg: chatKg, chatPct: total > 0 ? chatKg / total : null };
        });
        return { sku: k, label: labelOf[k] || k, months: months, totalChat: chatTotals[k] || 0 };
      });
      return { series: series, totalArticles: allKeys.length };
    }

    return {
      resumenPlanta: resumenPlanta, kpiHeader: kpiHeader, presupuesto: presupuesto,
      detalleMes: detalleMes, detalleAnio: detalleAnio, espesorAnalysis: espesorAnalysis,
      pptoTotalVal: pptoTotalVal, oeeMetaVal: oeeMetaVal,
      torreControl: torreControl, aporteOEEPorMaquina: aporteOEEPorMaquina, productMix: productMix,
      produccionMensualPorMaquina: produccionMensualPorMaquina, costosCategoria: costosCategoria,
      proyeccionCierre: proyeccionCierre, proyeccionAnalitica: proyeccionAnalitica,
      paretoChatarra: paretoChatarra, paretoProduccion: paretoProduccion,
      tendenciasPlanta: tendenciasPlanta, calidadPlantaDetalle: calidadPlantaDetalle,
      puntosCriticos: puntosCriticos, insumosPxQ: insumosPxQ, insumosMovimiento: insumosMovimiento,
      espesorKgFiltrado: espesorKgFiltrado, espesoresDisponibles: espesoresDisponibles,
      skuBuscador: skuBuscador, skuMonthlyTrend: skuMonthlyTrend, skuArticleMonthlyTrend: skuArticleMonthlyTrend
    };
  }

  /* =========================================================================
   * Chart rendering (bullet-style bar + target tick, per dataviz mark specs)
   * ======================================================================= */

  var chartUid = 0;
  // Charts render at viewBox W but stretch to 100% of their container's real pixel
  // width, so a fixed small W badly oversizes text/strokes inside a full-width card
  // (roughly 2x wider than a half-width .grid-2 card). Picking W from the container's
  // actual layout slot keeps every chart's fonts at their true, designed size.
  // Below the 980px breakpoint .grid-2 collapses to one column (see CSS), so every
  // card — half or full — becomes viewport-width; clamp W to that instead of
  // inflating it, or full-width charts would render with tiny, illegible text on phones.
  var CHART_W_FULL = 1280;
  function chartWidth(container, halfW) {
    var vw = window.innerWidth || CHART_W_FULL;
    if (vw < 980) return Math.min(CHART_W_FULL, Math.max(halfW, vw - 40));
    return (container.closest && container.closest('.grid-2')) ? halfW : CHART_W_FULL;
  }
  function renderBulletChart(container, opts) {
    // opts: {categories, values, targets, formatValue, formatTarget, barColorVar, seriesLabel, targetLabel}
    chartUid++;
    var W = chartWidth(container, 640), H = 244, padL = 8, padR = 8, padT = 22, padB = 26;
    var plotW = W - padL - padR, plotH = H - padT - padB;
    var n = opts.categories.length;
    var bandW = plotW / n;
    var barW = Math.max(6, Math.min(28, bandW * 0.5));

    var allVals = opts.values.concat(opts.targets).filter(isNum);
    var maxV = allVals.length ? Math.max.apply(null, allVals) : 1;
    var minV = allVals.length ? Math.min(0, Math.min.apply(null, allVals)) : 0;
    maxV = maxV * 1.22 || 1;

    function y(v) { return padT + plotH - ((v - minV) / (maxV - minV)) * plotH; }
    var baseline = y(0);

    var gridLines = 4, gridHtml = '', labelsHtml = '';
    for (var g = 0; g <= gridLines; g++) {
      var v = minV + (maxV - minV) * g / gridLines;
      var yy = y(v);
      gridHtml += '<line class="grid-line" x1="' + padL + '" x2="' + (W - padR) + '" y1="' + yy + '" y2="' + yy + '"></line>';
      labelsHtml += '<text class="axis-label" x="' + (padL) + '" y="' + (yy - 3) + '">' + opts.formatValue(v, true) + '</text>';
    }

    var barsHtml = '', ticksHtml = '', xLabelsHtml = '', valueLabelsHtml = '';
    var tipId = 'tip' + chartUid;
    var bars = [];
    opts.categories.forEach(function (cat, i) {
      var cx = padL + bandW * i + bandW / 2;
      var val = opts.values[i];
      var tgt = opts.targets[i];
      xLabelsHtml += '<text class="axis-label" x="' + cx + '" y="' + (H - 8) + '" text-anchor="middle">' + cat + '</text>';
      if (isNum(val)) {
        var y0 = baseline, y1 = y(val);
        var top = Math.min(y0, y1), h = Math.max(1, Math.abs(y1 - y0));
        var r = Math.min(4, barW / 2, h);
        var x0 = cx - barW / 2;
        var d = 'M' + x0 + ',' + (top + h) +
          'L' + x0 + ',' + (top + r) +
          'Q' + x0 + ',' + top + ' ' + (x0 + r) + ',' + top +
          'L' + (x0 + barW - r) + ',' + top +
          'Q' + (x0 + barW) + ',' + top + ' ' + (x0 + barW) + ',' + (top + r) +
          'L' + (x0 + barW) + ',' + (top + h) + 'Z';
        barsHtml += '<path class="bar" data-i="' + i + '" d="' + d + '" fill="' + opts.barColorVar + '"></path>';
        var labelY = val >= 0 ? Math.min(top, baseline) - 6 : Math.max(top + h, baseline) + 12;
        valueLabelsHtml += '<text class="value-label" data-i="' + i + '" x="' + cx + '" y="' + labelY + '" text-anchor="middle">' + opts.formatValue(val) + '</text>';
        bars.push({ i: i, cx: cx, val: val, tgt: tgt, cat: cat });
      }
      if (isNum(tgt)) {
        var yt = y(tgt);
        ticksHtml += '<line x1="' + (cx - barW / 2 - 3) + '" x2="' + (cx + barW / 2 + 3) + '" y1="' + yt + '" y2="' + yt + '" stroke="var(--text-primary)" stroke-width="2"></line>';
      }
    });

    var legendHtml = '<div class="chart-legend">' +
      '<span class="sw"><span class="dot" style="background:' + opts.barColorVar + '"></span>' + opts.seriesLabel + '</span>' +
      (opts.targetLabel ? '<span class="sw"><span class="dot" style="background:var(--text-primary)"></span>' + opts.targetLabel + '</span>' : '') +
      '</div>';

    var svg = '<div class="chart-wrap">' + legendHtml +
      '<svg class="chart" viewBox="0 0 ' + W + ' ' + H + '" id="svg' + chartUid + '">' +
      '<line class="baseline" x1="' + padL + '" x2="' + (W - padR) + '" y1="' + baseline + '" y2="' + baseline + '"></line>' +
      gridHtml + labelsHtml + barsHtml + ticksHtml + valueLabelsHtml + xLabelsHtml +
      '</svg>' +
      '<div class="chart-tooltip" id="' + tipId + '"></div>' +
      '</div>';
    container.innerHTML = svg;

    var tip = container.querySelector('#' + tipId);
    var wrapEl = container.querySelector('.chart-wrap');
    container.querySelectorAll('.bar').forEach(function (barEl) {
      var i = parseInt(barEl.getAttribute('data-i'), 10);
      var b = bars.filter(function (x) { return x.i === i; })[0];
      if (!b) return;
      var label = container.querySelector('.value-label[data-i="' + i + '"]');
      function activate(ev) {
        var rect = wrapEl.getBoundingClientRect();
        var scale = rect.width / W;
        tip.style.left = (b.cx * scale) + 'px';
        tip.style.top = (y(b.val) * scale) + 'px';
        tip.style.opacity = 1;
        tip.innerHTML = '<strong>' + b.cat + '</strong><br>' + opts.seriesLabel + ': ' + opts.formatValue(b.val) +
          (isNum(b.tgt) ? '<br>' + opts.targetLabel + ': ' + opts.formatValue(b.tgt) : '');
        barEl.classList.add('bar-active');
        if (label) label.classList.add('value-label-active');
      }
      function deactivate() {
        tip.style.opacity = 0;
        barEl.classList.remove('bar-active');
        if (label) label.classList.remove('value-label-active');
      }
      barEl.addEventListener('mousemove', activate);
      barEl.addEventListener('mouseleave', deactivate);
      if (label) {
        label.addEventListener('mousemove', activate);
        label.addEventListener('mouseleave', deactivate);
      }
    });
  }

  function renderLineChart(container, opts) {
    // opts: {categories, series:[{name,color,values}], formatValue}
    chartUid++;
    var multi = opts.series.length > 1;
    var W = 640, H = multi ? 260 : 250, padL = 8, padR = 8, padT = 26, padB = 26;
    var plotW = W - padL - padR, plotH = H - padT - padB;
    var n = opts.categories.length;

    var allVals = [];
    opts.series.forEach(function (s) { s.values.forEach(function (v) { if (isNum(v)) allVals.push(v); }); });
    var maxV = allVals.length ? Math.max.apply(null, allVals) : 1;
    var minV = allVals.length ? Math.min(0, Math.min.apply(null, allVals)) : 0;
    maxV = maxV * 1.15 || 1;
    minV = minV < 0 ? minV * 1.15 : minV;

    function x(i) { return n <= 1 ? padL + plotW / 2 : padL + plotW * i / (n - 1); }
    function y(v) { return padT + plotH - ((v - minV) / ((maxV - minV) || 1)) * plotH; }
    var baseline = y(0);

    var gridLines = 4, gridHtml = '', labelsHtml = '';
    for (var g = 0; g <= gridLines; g++) {
      var v = minV + (maxV - minV) * g / gridLines;
      var yy = y(v);
      gridHtml += '<line class="grid-line" x1="' + padL + '" x2="' + (W - padR) + '" y1="' + yy + '" y2="' + yy + '"></line>';
      labelsHtml += '<text class="axis-label" x="' + padL + '" y="' + (yy - 3) + '">' + opts.formatValue(v, true) + '</text>';
    }
    var xLabelsHtml = '';
    opts.categories.forEach(function (cat, i) {
      xLabelsHtml += '<text class="axis-label" x="' + x(i) + '" y="' + (H - 8) + '" text-anchor="middle">' + cat + '</text>';
    });

    var seriesHtml = '', markersHtml = '', valueLabelsHtml = '', legendHtml = '<div class="chart-legend">';
    opts.series.forEach(function (s, si) {
      var d = '', open = false;
      s.values.forEach(function (v, i) {
        if (!isNum(v)) { open = false; return; }
        d += (open ? 'L' : 'M') + x(i) + ',' + y(v) + ' ';
        open = true;
      });
      seriesHtml += '<path fill="none" stroke="' + s.color + '" stroke-width="2.25" stroke-linecap="round" stroke-linejoin="round" d="' + d.trim() + '"></path>';
      s.values.forEach(function (v, i) {
        if (!isNum(v)) return;
        markersHtml += '<circle class="pt" data-s="' + si + '" data-i="' + i + '" cx="' + x(i) + '" cy="' + y(v) + '" r="4" fill="' + s.color + '" stroke="var(--surface-1)" stroke-width="2"></circle>';
        var above = si % 2 === 0;
        var ly = above ? y(v) - 10 : y(v) + 17;
        valueLabelsHtml += '<text class="value-label" data-s="' + si + '" data-i="' + i + '" x="' + x(i) + '" y="' + ly + '" text-anchor="middle" fill="' + s.color + '">' + opts.formatValue(v) + '</text>';
      });
      legendHtml += '<span class="sw"><span class="dot" style="background:' + s.color + '"></span>' + s.name + '</span>';
    });
    legendHtml += '</div>';

    var tipId = 'tip' + chartUid;
    var crossId = 'cross' + chartUid;
    var bandW = n > 1 ? plotW / (n - 1) : plotW;
    var hitHtml = '';
    opts.categories.forEach(function (cat, i) {
      var hx = x(i);
      hitHtml += '<rect class="hit-col" data-i="' + i + '" x="' + (hx - bandW / 2) + '" y="' + padT + '" width="' + bandW + '" height="' + plotH + '" fill="transparent"></rect>';
    });

    container.innerHTML = '<div class="chart-wrap">' + legendHtml +
      '<svg class="chart" viewBox="0 0 ' + W + ' ' + H + '" id="svg' + chartUid + '">' +
      '<line class="baseline" x1="' + padL + '" x2="' + (W - padR) + '" y1="' + baseline + '" y2="' + baseline + '"></line>' +
      gridHtml + labelsHtml +
      '<line id="' + crossId + '" class="crosshair" x1="0" x2="0" y1="' + padT + '" y2="' + (padT + plotH) + '" opacity="0"></line>' +
      seriesHtml + markersHtml + valueLabelsHtml + xLabelsHtml + hitHtml +
      '</svg><div class="chart-tooltip" id="' + tipId + '"></div></div>';

    var tip = container.querySelector('#' + tipId);
    var cross = container.querySelector('#' + crossId);
    var wrapEl = container.querySelector('.chart-wrap');
    container.querySelectorAll('.hit-col').forEach(function (hit) {
      var i = parseInt(hit.getAttribute('data-i'), 10);
      hit.addEventListener('mousemove', function () {
        var rect = wrapEl.getBoundingClientRect();
        var scale = rect.width / W;
        var firstVal = opts.series.map(function (s) { return s.values[i]; }).filter(isNum)[0];
        var top = (isNum(firstVal) ? y(firstVal) : padT) * scale;
        tip.style.left = (x(i) * scale) + 'px';
        tip.style.top = top + 'px';
        var html = '<strong>' + opts.categories[i] + '</strong>';
        opts.series.forEach(function (s) {
          var v = s.values[i];
          html += '<br><span style="color:' + s.color + '">●</span> ' + s.name + ': ' + (isNum(v) ? opts.formatValue(v) : '-');
        });
        tip.innerHTML = html;
        tip.style.opacity = 1;
        cross.setAttribute('x1', x(i)); cross.setAttribute('x2', x(i)); cross.setAttribute('opacity', 1);
        container.querySelectorAll('.pt').forEach(function (p) { p.classList.toggle('pt-active', p.getAttribute('data-i') === String(i)); });
      });
      hit.addEventListener('mouseleave', function () {
        tip.style.opacity = 0;
        cross.setAttribute('opacity', 0);
        container.querySelectorAll('.pt').forEach(function (p) { p.classList.remove('pt-active'); });
      });
    });
  }

  function renderStackedBarChart(container, opts) {
    // opts: {categories, series:[{name,color,values(fractions 0..1)}]}
    chartUid++;
    var W = chartWidth(container, 680), H = 260, padL = 8, padR = 8, padT = 10, padB = 26;
    var plotW = W - padL - padR, plotH = H - padT - padB;
    var n = opts.categories.length;
    var bandW = plotW / n;
    var barW = Math.max(10, Math.min(34, bandW * 0.6));
    var gap = 2;

    var gridLines = 4, gridHtml = '', labelsHtml = '';
    for (var g = 0; g <= gridLines; g++) {
      var frac = g / gridLines;
      var yy = padT + plotH - frac * plotH;
      gridHtml += '<line class="grid-line" x1="' + padL + '" x2="' + (W - padR) + '" y1="' + yy + '" y2="' + yy + '"></line>';
      labelsHtml += '<text class="axis-label" x="' + padL + '" y="' + (yy - 3) + '">' + Math.round(frac * 100) + '%</text>';
    }

    var barsHtml = '', xLabelsHtml = '', stackLabelsHtml = '';
    var segs = [];
    opts.categories.forEach(function (cat, i) {
      var cx = padL + bandW * i + bandW / 2;
      xLabelsHtml += '<text class="axis-label" x="' + cx + '" y="' + (H - 8) + '" text-anchor="middle">' + cat + '</text>';
      var total = 0;
      opts.series.forEach(function (s) { total += (s.values[i] || 0); });
      if (!total) return;
      var yCursor = padT + plotH;
      opts.series.forEach(function (s, si) {
        var v = s.values[i] || 0;
        if (v <= 0) return;
        var segH = (v / total) * plotH;
        var top = yCursor - segH;
        var x0 = cx - barW / 2;
        var hDraw = Math.max(0, segH - gap);
        barsHtml += '<rect class="bar" x="' + x0 + '" y="' + top + '" width="' + barW + '" height="' + hDraw + '" rx="2" fill="' + s.color + '" data-i="' + i + '" data-s="' + si + '"></rect>';
        // name the series INSIDE the segment — matching colors across 8 thicknesses
        // is hard, so tall segments carry "3mm 45%" and medium ones just "3mm"
        if (hDraw >= 26) {
          stackLabelsHtml += '<text class="mix-stack-label" x="' + cx + '" y="' + (top + hDraw / 2 - 3) + '" text-anchor="middle">' + escapeHtml(s.name.replace(' mm', 'mm')) + '</text>' +
            '<text class="mix-stack-label" x="' + cx + '" y="' + (top + hDraw / 2 + 9) + '" text-anchor="middle">' + fmtPct(v / total, 0) + '</text>';
        } else if (hDraw >= 13) {
          stackLabelsHtml += '<text class="mix-stack-label" x="' + cx + '" y="' + (top + hDraw / 2 + 3.5) + '" text-anchor="middle">' + escapeHtml(s.name.replace(' mm', 'mm')) + '</text>';
        }
        segs.push({ i: i, si: si, cat: cat, name: s.name, v: v / total, cx: cx, top: top });
        yCursor -= segH;
      });
    });

    var legendHtml = '<div class="chart-legend">' + opts.series.map(function (s) {
      return '<span class="sw"><span class="dot" style="background:' + s.color + '"></span>' + s.name + '</span>';
    }).join('') + '</div>';

    var tipId = 'tip' + chartUid;
    container.innerHTML = '<div class="chart-wrap">' + legendHtml +
      '<svg class="chart" viewBox="0 0 ' + W + ' ' + H + '" id="svg' + chartUid + '">' +
      '<line class="baseline" x1="' + padL + '" x2="' + (W - padR) + '" y1="' + (padT + plotH) + '" y2="' + (padT + plotH) + '"></line>' +
      gridHtml + labelsHtml + barsHtml + stackLabelsHtml + xLabelsHtml +
      '</svg><div class="chart-tooltip" id="' + tipId + '"></div></div>';

    var tip = container.querySelector('#' + tipId);
    var wrapEl = container.querySelector('.chart-wrap');
    container.querySelectorAll('.bar').forEach(function (barEl) {
      var i = parseInt(barEl.getAttribute('data-i'), 10), si = parseInt(barEl.getAttribute('data-s'), 10);
      var seg = segs.filter(function (s) { return s.i === i && s.si === si; })[0];
      if (!seg) return;
      barEl.addEventListener('mousemove', function () {
        var rect = wrapEl.getBoundingClientRect();
        var scale = rect.width / W;
        tip.style.left = (seg.cx * scale) + 'px';
        tip.style.top = (seg.top * scale) + 'px';
        tip.style.opacity = 1;
        tip.innerHTML = '<strong>' + seg.cat + '</strong><br>' + seg.name + ': ' + fmtPct(seg.v, 1);
        barEl.classList.add('bar-active');
      });
      barEl.addEventListener('mouseleave', function () { tip.style.opacity = 0; barEl.classList.remove('bar-active'); });
    });
  }

  function renderGroupedBarChart(container, opts) {
    // opts: {categories, series:[{name,color,values}], formatValue}
    chartUid++;
    var W = chartWidth(container, 680), H = 250, padL = 8, padR = 8, padT = 12, padB = 26;
    var plotW = W - padL - padR, plotH = H - padT - padB;
    var n = opts.categories.length;
    var nS = opts.series.length;
    var bandW = plotW / n;
    var gap = 2;
    var barW = Math.max(4, (bandW * 0.72) / nS);

    var allVals = [];
    opts.series.forEach(function (s) { s.values.forEach(function (v) { if (isNum(v)) allVals.push(v); }); });
    var maxV = allVals.length ? Math.max.apply(null, allVals) : 1;
    maxV = maxV * 1.18 || 1;

    function y(v) { return padT + plotH - (v / maxV) * plotH; }
    var baseline = y(0);

    var gridLines = 4, gridHtml = '', labelsHtml = '';
    for (var g = 0; g <= gridLines; g++) {
      var v = maxV * g / gridLines;
      var yy = y(v);
      gridHtml += '<line class="grid-line" x1="' + padL + '" x2="' + (W - padR) + '" y1="' + yy + '" y2="' + yy + '"></line>';
      labelsHtml += '<text class="axis-label" x="' + padL + '" y="' + (yy - 3) + '">' + opts.formatValue(v, true) + '</text>';
    }

    var barsHtml = '', xLabelsHtml = '';
    var tipId = 'tip' + chartUid;
    var bars = [];
    opts.categories.forEach(function (cat, i) {
      var groupX0 = padL + bandW * i + (bandW - barW * nS) / 2;
      xLabelsHtml += '<text class="axis-label" x="' + (padL + bandW * i + bandW / 2) + '" y="' + (H - 8) + '" text-anchor="middle">' + cat + '</text>';
      opts.series.forEach(function (s, si) {
        var val = s.values[i];
        if (!isNum(val)) return;
        var x0 = groupX0 + si * barW;
        var top = Math.min(baseline, y(val)), h = Math.max(1, Math.abs(y(val) - baseline));
        var r = Math.min(3, (barW - gap) / 2, h);
        var w = Math.max(1, barW - gap);
        barsHtml += '<rect class="bar" data-i="' + i + '" data-s="' + si + '" x="' + x0 + '" y="' + top + '" width="' + w + '" height="' + h + '" rx="' + r + '" fill="' + s.color + '"></rect>';
        bars.push({ i: i, si: si, cat: cat, name: s.name, val: val, cx: x0 + w / 2, top: top });
      });
    });

    var legendHtml = '<div class="chart-legend">' + opts.series.map(function (s) {
      return '<span class="sw"><span class="dot" style="background:' + s.color + '"></span>' + s.name + '</span>';
    }).join('') + '</div>';

    container.innerHTML = '<div class="chart-wrap">' + legendHtml +
      '<svg class="chart" viewBox="0 0 ' + W + ' ' + H + '" id="svg' + chartUid + '">' +
      '<line class="baseline" x1="' + padL + '" x2="' + (W - padR) + '" y1="' + baseline + '" y2="' + baseline + '"></line>' +
      gridHtml + labelsHtml + barsHtml + xLabelsHtml +
      '</svg><div class="chart-tooltip" id="' + tipId + '"></div></div>';

    var tip = container.querySelector('#' + tipId);
    var wrapEl = container.querySelector('.chart-wrap');
    container.querySelectorAll('.bar').forEach(function (barEl) {
      var i = parseInt(barEl.getAttribute('data-i'), 10), si = parseInt(barEl.getAttribute('data-s'), 10);
      var b = bars.filter(function (x) { return x.i === i && x.si === si; })[0];
      if (!b) return;
      barEl.addEventListener('mousemove', function () {
        var rect = wrapEl.getBoundingClientRect();
        var scale = rect.width / W;
        tip.style.left = (b.cx * scale) + 'px';
        tip.style.top = (b.top * scale) + 'px';
        tip.style.opacity = 1;
        tip.innerHTML = '<strong>' + b.cat + '</strong><br>' + b.name + ': ' + opts.formatValue(b.val);
        barEl.classList.add('bar-active');
      });
      barEl.addEventListener('mouseleave', function () { tip.style.opacity = 0; barEl.classList.remove('bar-active'); });
    });
  }

  function renderRankedBarChart(container, opts) {
    // opts: {items:[{label,value}], formatValue, colorVar}
    // horizontal bars — the right layout for long text category names (item/product
    // names), where a vertical bar chart's x-axis labels would collide or need
    // unreadable truncation.
    chartUid++;
    var items = opts.items.filter(function (it) { return isNum(it.value); });
    var n = items.length;
    var rowH = 30, padR = 70, padTop = 4, padBottom = 4;
    // Long item names (purchase-order style descriptions run 50-65+ chars) used to size
    // the label column with no cap, so a single very long name could force the column
    // to eat most of a half-width card, leaving barely any room for the bars themselves.
    // Capping what's drawn keeps proportions consistent regardless of card width — the
    // full name is still always available on hover.
    var LABEL_MAX = 42;
    function shortLabel(s) { return s.length > LABEL_MAX ? s.slice(0, LABEL_MAX - 1) + '…' : s; }
    // 6.6px/char approximates the 11px semibold font's real advance width.
    var maxLabelLen = items.length ? Math.max.apply(null, items.map(function (it) { return Math.min((it.label || '').length, LABEL_MAX); })) : 10;
    var labelW = Math.max(140, Math.min(440, maxLabelLen * 6.6 + 20));
    var targetW = chartWidth(container, 640);
    var plotW = Math.max(220, targetW - labelW - padR);
    var W = labelW + plotW + padR;
    var H = n * rowH + padTop + padBottom || rowH;

    var maxV = items.length ? Math.max.apply(null, items.map(function (it) { return Math.abs(it.value); })) : 1;
    maxV = maxV * 1.12 || 1;

    var barsHtml = '', tipId = 'tip' + chartUid;
    var bars = [];
    items.forEach(function (it, i) {
      var y0 = padTop + i * rowH;
      var barW = Math.max(1, (Math.abs(it.value) / maxV) * plotW);
      var barH = rowH - 10;
      var by = y0 + 5;
      barsHtml += '<text class="axis-label rank-label" x="' + (labelW - 10) + '" y="' + (by + barH / 2 + 3.5) + '" text-anchor="end">' + escapeHtml(shortLabel(it.label)) + '</text>';
      barsHtml += '<rect class="bar" data-i="' + i + '" x="' + labelW + '" y="' + by + '" width="' + barW + '" height="' + barH + '" rx="3" fill="' + opts.colorVar + '"></rect>';
      barsHtml += '<text class="value-label" x="' + (labelW + barW + 6) + '" y="' + (by + barH / 2 + 3.5) + '" text-anchor="start">' + opts.formatValue(it.value) + '</text>';
      bars.push({ i: i, label: it.label, value: it.value, cx: labelW + barW / 2, top: by });
    });

    container.innerHTML = '<div class="chart-wrap">' +
      '<svg class="chart" viewBox="0 0 ' + W + ' ' + H + '" id="svg' + chartUid + '">' +
      barsHtml +
      '</svg><div class="chart-tooltip" id="' + tipId + '"></div></div>';

    var tip = container.querySelector('#' + tipId);
    var wrapEl = container.querySelector('.chart-wrap');
    container.querySelectorAll('.bar').forEach(function (barEl) {
      var i = parseInt(barEl.getAttribute('data-i'), 10);
      var b = bars.filter(function (x) { return x.i === i; })[0];
      if (!b) return;
      barEl.addEventListener('mousemove', function () {
        var rect = wrapEl.getBoundingClientRect();
        var scale = rect.width / W;
        tip.style.left = (b.cx * scale) + 'px';
        tip.style.top = (b.top * scale) + 'px';
        tip.style.opacity = 1;
        tip.innerHTML = '<strong>' + escapeHtml(b.label) + '</strong><br>' + opts.formatValue(b.value);
        barEl.classList.add('bar-active');
      });
      barEl.addEventListener('mouseleave', function () { tip.style.opacity = 0; barEl.classList.remove('bar-active'); });
    });
  }

  function renderDonutChart(container, opts) {
    // opts: {items:[{label, value}], formatValue} — top 8 + "Otros", % share labels,
    // legend carries identity (colors follow the fixed categorical order).
    chartUid++;
    var MAXS = 8;
    var items = opts.items.filter(function (it) { return isNum(it.value) && it.value > 0; });
    var rest = items.slice(MAXS);
    items = items.slice(0, MAXS);
    if (rest.length) items.push({ label: 'Otros (' + rest.length + ')', value: sum(rest.map(function (it) { return it.value; })), otros: true });
    var total = sum(items.map(function (it) { return it.value; }));
    if (!total) { container.innerHTML = '<div class="cap">Sin datos para este filtro.</div>'; return; }

    var W = 720, H = 260, cx = 190, cy = H / 2, R = 100, r0 = 58;
    var segs = '', labels = '';
    var a0 = -Math.PI / 2;
    var pts = [];
    items.forEach(function (it, i) {
      var frac = it.value / total;
      var a1 = a0 + frac * Math.PI * 2;
      // 2px surface gap between segments via stroke on the path
      var large = (a1 - a0) > Math.PI ? 1 : 0;
      var x0 = cx + R * Math.cos(a0), y0 = cy + R * Math.sin(a0);
      var x1 = cx + R * Math.cos(a1), y1 = cy + R * Math.sin(a1);
      var xi1 = cx + r0 * Math.cos(a1), yi1 = cy + r0 * Math.sin(a1);
      var xi0 = cx + r0 * Math.cos(a0), yi0 = cy + r0 * Math.sin(a0);
      var color = it.otros ? 'var(--text-muted)' : 'var(--series-' + ((i % 8) + 1) + ')';
      segs += '<path class="bar" data-i="' + i + '" d="M' + x0 + ',' + y0 +
        ' A' + R + ',' + R + ' 0 ' + large + ' 1 ' + x1 + ',' + y1 +
        ' L' + xi1 + ',' + yi1 +
        ' A' + r0 + ',' + r0 + ' 0 ' + large + ' 0 ' + xi0 + ',' + yi0 + ' Z" fill="' + color + '" stroke="var(--surface-1)" stroke-width="2"></path>';
      var am = (a0 + a1) / 2;
      if (frac >= 0.04) {
        var lx = cx + (R + r0) / 2 * Math.cos(am), ly = cy + (R + r0) / 2 * Math.sin(am);
        labels += '<text class="stack-label" x="' + lx + '" y="' + (ly + 3.5) + '" text-anchor="middle">' + fmtPct(frac, 0) + '</text>';
      }
      pts.push({ i: i, label: it.label, value: it.value, frac: frac, mx: cx + R * Math.cos(am), my: cy + R * Math.sin(am) });
      a0 = a1;
    });
    labels += '<text x="' + cx + '" y="' + (cy - 4) + '" text-anchor="middle" style="font-size:19px;font-weight:750;fill:var(--text-primary);">' + fmtKgTick(total) + '</text>' +
      '<text class="axis-label" x="' + cx + '" y="' + (cy + 14) + '" text-anchor="middle">kg totales</text>';

    var legend = '<div style="display:flex;flex-direction:column;gap:7px;justify-content:center;">' + pts.map(function (p, i) {
      var color = items[i].otros ? 'var(--text-muted)' : 'var(--series-' + ((i % 8) + 1) + ')';
      return '<span class="sw" style="display:inline-flex;align-items:center;gap:7px;font-size:12px;color:var(--text-secondary);">' +
        '<span class="dot" style="width:9px;height:9px;border-radius:2px;background:' + color + ';flex:0 0 9px;"></span>' +
        '<span>' + escapeHtml(p.label) + ' <span class="subtle">' + fmtPct(p.frac, 1) + '</span></span></span>';
    }).join('') + '</div>';

    var tipId = 'tip' + chartUid;
    container.innerHTML = '<div class="chart-wrap" style="display:flex;gap:18px;align-items:center;flex-wrap:wrap;">' +
      '<svg class="chart" viewBox="0 0 380 ' + H + '" style="max-width:380px;flex:1 1 300px;" id="svg' + chartUid + '">' + segs + labels + '</svg>' +
      legend + '<div class="chart-tooltip" id="' + tipId + '"></div></div>';

    var tip = container.querySelector('#' + tipId);
    var wrapEl = container.querySelector('.chart-wrap');
    container.querySelectorAll('path.bar').forEach(function (elm) {
      var p = pts[parseInt(elm.getAttribute('data-i'), 10)];
      if (!p) return;
      elm.addEventListener('mousemove', function (ev) {
        var rect = wrapEl.getBoundingClientRect();
        tip.style.left = (ev.clientX - rect.left) + 'px';
        tip.style.top = (ev.clientY - rect.top - 8) + 'px';
        tip.style.opacity = 1;
        tip.innerHTML = '<strong>' + escapeHtml(p.label) + '</strong><br>' + opts.formatValue(p.value) + ' — ' + fmtPct(p.frac, 1);
      });
      elm.addEventListener('mouseleave', function () { tip.style.opacity = 0; });
    });
  }

  function renderParetoChart(container, opts) {
    // opts: {items:[{label, kg, share, cum}], n80}
    // Classic Pareto with ONE axis: bars show each article's % share of total scrap,
    // the line shows the cumulative %, both on the same 0-100% scale (kg go in the
    // tooltip). Top items individually + the rest folded into "Resto".
    chartUid++;
    var MAXB = 15;
    var items = opts.items.slice(0, MAXB);
    var resto = opts.items.slice(MAXB);
    if (resto.length) {
      var restoKg = sum(resto.map(function (it) { return it.kg; }));
      var restoShare = sum(resto.map(function (it) { return it.share; }));
      items.push({ label: 'Resto (' + resto.length + ' artículos)', kg: restoKg, share: restoShare, cum: 1 });
    }
    var n = items.length;
    var W = chartWidth(container, 720), H = 300, padL = 40, padR = 20, padT = 16, padB = 24;
    var plotW = W - padL - padR, plotH = H - padT - padB;
    var bandW = plotW / Math.max(n, 1);
    var barW = Math.max(8, Math.min(34, bandW * 0.6));

    function y(f) { return padT + plotH - f * plotH; }
    var gridHtml = '', labelsHtml = '';
    [0, 0.25, 0.5, 0.75, 1].forEach(function (f) {
      var yy = y(f);
      gridHtml += '<line class="grid-line" x1="' + padL + '" x2="' + (W - padR) + '" y1="' + yy + '" y2="' + yy + '"></line>';
      labelsHtml += '<text class="axis-label" x="' + (padL - 8) + '" y="' + (yy + 3) + '" text-anchor="end">' + Math.round(f * 100) + '%</text>';
    });
    // 80% reference
    var y80 = y(0.8);
    gridHtml += '<line x1="' + padL + '" x2="' + (W - padR) + '" y1="' + y80 + '" y2="' + y80 + '" stroke="var(--critical)" stroke-width="1" stroke-dasharray="4 3" opacity=".55"></line>';
    labelsHtml += '<text class="axis-label" x="' + (W - padR) + '" y="' + (y80 - 4) + '" text-anchor="end" fill="var(--critical)">80%</text>';

    var barsHtml = '', lineHtml = '', ptsHtml = '', xLabelsHtml = '';
    var pts = [];
    items.forEach(function (it, i) {
      var cx = padL + bandW * i + bandW / 2;
      var isResto = i === items.length - 1 && resto.length;
      var bh = Math.max(1, it.share * plotH);
      barsHtml += '<rect class="bar" data-i="' + i + '" x="' + (cx - barW / 2) + '" y="' + (padT + plotH - bh) + '" width="' + barW + '" height="' + bh + '" rx="3" fill="' + (isResto ? 'var(--text-muted)' : 'var(--series-1)') + '"></rect>';
      pts.push({ i: i, cx: cx, cy: y(it.cum), it: it });
      xLabelsHtml += '<text class="axis-label" x="' + cx + '" y="' + (H - 8) + '" text-anchor="middle">' + (i + 1) + '</text>';
    });
    var d = pts.map(function (p, i) { return (i ? 'L' : 'M') + p.cx + ',' + p.cy; }).join('');
    lineHtml = '<path d="' + d + '" fill="none" stroke="var(--series-8)" stroke-width="2"></path>';
    pts.forEach(function (p) {
      ptsHtml += '<circle class="bar" data-i="' + p.i + '" cx="' + p.cx + '" cy="' + p.cy + '" r="3.5" fill="var(--series-8)" stroke="var(--surface-1)" stroke-width="1.5"></circle>';
    });

    var legendHtml = '<div class="chart-legend">' +
      '<span class="sw"><span class="dot" style="background:var(--series-1)"></span>% del total</span>' +
      '<span class="sw"><span class="dot" style="background:var(--series-8)"></span>% acumulado</span></div>';
    var tipId = 'tip' + chartUid;
    container.innerHTML = '<div class="chart-wrap">' + legendHtml +
      '<svg class="chart" viewBox="0 0 ' + W + ' ' + H + '" id="svg' + chartUid + '">' +
      gridHtml + labelsHtml + barsHtml + lineHtml + ptsHtml + xLabelsHtml +
      '</svg><div class="chart-tooltip" id="' + tipId + '"></div></div>';

    var tip = container.querySelector('#' + tipId);
    var wrapEl = container.querySelector('.chart-wrap');
    container.querySelectorAll('.bar').forEach(function (elm) {
      var i = parseInt(elm.getAttribute('data-i'), 10);
      var p = pts[i];
      if (!p) return;
      elm.addEventListener('mousemove', function () {
        var rect = wrapEl.getBoundingClientRect();
        var scale = rect.width / W;
        tip.style.left = (p.cx * scale) + 'px';
        tip.style.top = ((padT + 8) * scale) + 'px';
        tip.style.opacity = 1;
        tip.innerHTML = '<strong>#' + (i + 1) + ' ' + escapeHtml(p.it.label) + '</strong><br>' +
          fmtInt(p.it.kg) + ' kg — ' + fmtPct(p.it.share, 1) + ' del total<br>acumulado: ' + fmtPct(p.it.cum, 1);
      });
      elm.addEventListener('mouseleave', function () { tip.style.opacity = 0; });
    });
  }

  function renderBubbleChart(container, opts) {
    // opts: {categories, series:[{label,color,values(0..1 y-axis),sizes(kg, bubble area)}], formatValue, formatSize}
    // one bubble per article × month: center label = kg (formatSize), outside label = full article name.
    // Every article is shown (no top-N cap), but the chart stays a fixed width that fits the
    // page (no horizontal scrollbar): crowded months push their labels up onto extra vertical
    // tiers instead of pushing the chart wider, and bubbles shrink a little in busy months.
    chartUid++;
    var series = opts.series;
    var nS = series.length;
    var n = opts.categories.length;
    var W = chartWidth(container, 720), padL = 46, padR = 20, padB = 34;
    var plotW = W - padL - padR, plotH = 230;

    var allVals = [];
    series.forEach(function (s) { s.values.forEach(function (v) { if (isNum(v)) allVals.push(v); }); });
    var maxV = allVals.length ? Math.max.apply(null, allVals) : 1;
    maxV = maxV * 1.3 || 1;

    var allSizes = [];
    series.forEach(function (s) { s.sizes.forEach(function (v) { if (isNum(v) && v > 0) allSizes.push(v); }); });
    var maxSize = allSizes.length ? Math.max.apply(null, allSizes) : 1;
    var minR = 6, maxRBase = nS > 1 ? 20 : 28;
    function radius(s, capR) {
      if (!isNum(s) || s <= 0 || !maxSize) return 0;
      return minR + Math.sqrt(s / maxSize) * (capR - minR);
    }
    // rough width (px) a horizontal label needs at the 9px bold font used for name labels
    function labelW(text) { return Math.max(40, Math.min(220, text.length * 5.4 + 14)); }

    var bandW = plotW / n;
    function xBase(i) { return padL + bandW * i + bandW / 2; }

    // bucket every (article, month) point into its category band, spread evenly across the
    // band's fixed width, and shrink the bubble radius a bit when many items share a band.
    var slots = [];
    for (var i = 0; i < n; i++) slots.push([]);
    series.forEach(function (s, si) {
      s.values.forEach(function (val, i) {
        if (!isNum(val)) return;
        slots[i].push({ val: val, size: s.sizes[i], label: s.label, color: s.color });
      });
    });
    var items = [];
    slots.forEach(function (list, i) {
      var bandLeft = padL + bandW * i;
      var count = list.length;
      var capR = count > 1 ? Math.min(maxRBase, Math.max(minR + 2, bandW / (count * 2.1))) : maxRBase;
      list.forEach(function (it, idx) {
        var sub = bandW / count;
        items.push({
          cx: bandLeft + sub * (idx + 0.5), val: it.val, size: it.size, label: it.label, color: it.color,
          cat: opts.categories[i], r: Math.max(3, radius(it.size, capR))
        });
      });
    });

    // vertical name labels: each article's name stands upright over its own bubble,
    // so labels can't collide horizontally unless two bubbles nearly overlap — in
    // that rare case the later label is dropped (hover tooltip still names it).
    var lastLabelX = null;
    items.slice().sort(function (a, b) { return a.cx - b.cx; }).forEach(function (it) {
      it.hideLabel = lastLabelX != null && (it.cx - lastLabelX) < 11;
      if (!it.hideLabel) lastLabelX = it.cx;
    });
    // Long article names (30+ chars is common) rotated vertically need room above
    // their own bubble — but reserving a top margin sized for the single longest
    // name in the whole chart (regardless of which bubble it belongs to) leaves a
    // huge empty gap whenever that long-named article isn't the highest-value one.
    // Instead: start from a small margin, see how far each bubble's own label
    // actually reaches above it, and only grow the margin by whatever the worst
    // individual overflow turns out to be — self-adjusting instead of worst-case.
    var NAME_LABEL_MAX = 24;
    function shortLabel(s) { return s.length > NAME_LABEL_MAX ? s.slice(0, NAME_LABEL_MAX - 1) + '…' : s; }
    var maxR = items.length ? Math.max.apply(null, items.map(function (it) { return it.r; })) : minR;
    var padT0 = Math.max(30, maxR + 10);
    var H0 = padT0 + plotH + padB;

    // A square-root y-scale, not linear: chatarra % is usually a long right-tailed
    // distribution (most articles near 0%, an occasional article near 100%), so a
    // linear axis squashes every normal value into a thin sliver at the bottom to
    // make room for one outlier. Square-root compresses the top of the range and
    // gives the bottom (where almost every point lives) much more room, without
    // hiding or relabeling the outlier's true value.
    var maxVSqrt = Math.sqrt(maxV) || 1;
    function yWithPadT(padTv, v) { return padTv + plotH - (Math.sqrt(Math.max(0, v)) / maxVSqrt) * plotH; }
    items.forEach(function (it) { it.cy = yWithPadT(padT0, it.val); });

    var topSafety = 8;
    var worstOverflow = 0;
    items.forEach(function (it) {
      if (it.hideLabel) return;
      var nameLen = Math.min(it.label.length, NAME_LABEL_MAX);
      var labelTop = it.cy - it.r - 5 - nameLen * 5.4;
      if (labelTop < topSafety) worstOverflow = Math.max(worstOverflow, topSafety - labelTop);
    });
    var padT = padT0 + worstOverflow;
    var H = H0 + worstOverflow;
    function y(v) { return yWithPadT(padT, v); }
    var baseline = y(0);
    items.forEach(function (it) { it.cy += worstOverflow; });

    var gridLines = 4, gridHtml = '', labelsHtml = '';
    for (var g = 0; g <= gridLines; g++) {
      var v = maxV * (g / gridLines) * (g / gridLines);
      var yy = y(v);
      gridHtml += '<line class="grid-line" x1="' + padL + '" x2="' + (W - padR) + '" y1="' + yy + '" y2="' + yy + '"></line>';
      labelsHtml += '<text class="axis-label" x="' + (padL - 8) + '" y="' + (yy + 3) + '" text-anchor="end">' + opts.formatValue(v, true) + '</text>';
    }
    var xLabelsHtml = '';
    opts.categories.forEach(function (cat, i) {
      xLabelsHtml += '<text class="axis-label" x="' + xBase(i) + '" y="' + (H - 10) + '" text-anchor="middle">' + cat + '</text>';
    });
    // dashed vertical rule between each month's band so it's clear where one month's
    // articles end and the next month's begin
    var monthSepHtml = '';
    for (var sepI = 1; sepI < n; sepI++) {
      var sepX = padL + bandW * sepI;
      monthSepHtml += '<line class="month-sep" x1="' + sepX + '" x2="' + sepX + '" y1="8" y2="' + (H - padB + 6) + '"></line>';
    }

    var bubblesHtml = '', centerLabelsHtml = '', nameLabelsHtml = '', leadersHtml = '';
    var tipId = 'tip' + chartUid;
    var bubbles = [];
    var uid = 0;
    items.forEach(function (it) {
      uid++;
      var cx = it.cx, cy = it.cy, r = it.r;
      bubblesHtml += '<circle class="bubble" data-u="' + uid + '" cx="' + cx + '" cy="' + cy + '" r="' + r + '" fill="' + it.color + '" fill-opacity="0.35" stroke="' + it.color + '" stroke-width="2"></circle>';
      if (isNum(it.size) && it.size > 0 && r >= 13) {
        centerLabelsHtml += '<text class="bubble-center-label" data-u="' + uid + '" x="' + cx + '" y="' + (cy + 3.5) + '" text-anchor="middle">' + opts.formatSize(it.size) + '</text>';
      }
      if (!it.hideLabel) {
        // vertical label: anchored at the bubble's top edge, reading bottom-to-top
        var nameY = cy - r - 5;
        nameLabelsHtml += '<text class="bubble-name-label" data-u="' + uid + '" x="' + cx + '" y="' + nameY + '" text-anchor="start" transform="rotate(-90 ' + cx + ' ' + nameY + ')" fill="' + it.color + '">' + escapeHtml(shortLabel(it.label)) + '</text>';
      }
      bubbles.push({ uid: uid, label: it.label, cat: it.cat, val: it.val, size: it.size, cx: cx, cy: cy, r: r });
    });

    container.innerHTML = '<div class="chart-wrap">' +
      '<svg class="chart" viewBox="0 0 ' + W + ' ' + H + '" id="svg' + chartUid + '">' +
      '<line class="baseline" x1="' + padL + '" x2="' + (W - padR) + '" y1="' + baseline + '" y2="' + baseline + '"></line>' +
      gridHtml + monthSepHtml + labelsHtml + leadersHtml + bubblesHtml + centerLabelsHtml + nameLabelsHtml + xLabelsHtml +
      '</svg><div class="chart-tooltip" id="' + tipId + '"></div></div>';

    var tip = container.querySelector('#' + tipId);
    var wrapEl = container.querySelector('.chart-wrap');
    container.querySelectorAll('.bubble').forEach(function (bEl) {
      var u = parseInt(bEl.getAttribute('data-u'), 10);
      var b = bubbles.filter(function (x) { return x.uid === u; })[0];
      if (!b) return;
      var centerLabel = container.querySelector('.bubble-center-label[data-u="' + u + '"]');
      var nameLabel = container.querySelector('.bubble-name-label[data-u="' + u + '"]');
      function activate() {
        var rect = wrapEl.getBoundingClientRect();
        var scale = rect.width / W;
        tip.style.left = (b.cx * scale) + 'px';
        tip.style.top = ((b.cy - b.r) * scale) + 'px';
        tip.style.opacity = 1;
        tip.innerHTML = '<strong>' + escapeHtml(b.label) + '</strong><br>' + b.cat + ' — % Chatarra: ' + opts.formatValue(b.val) +
          (isNum(b.size) ? '<br>Chatarra: ' + opts.formatSize(b.size) : '');
        bEl.classList.add('bar-active');
        if (nameLabel) nameLabel.classList.add('value-label-active');
      }
      function deactivate() {
        tip.style.opacity = 0; bEl.classList.remove('bar-active');
        if (nameLabel) nameLabel.classList.remove('value-label-active');
      }
      bEl.addEventListener('mousemove', activate);
      bEl.addEventListener('mouseleave', deactivate);
      if (nameLabel) { nameLabel.addEventListener('mousemove', activate); nameLabel.addEventListener('mouseleave', deactivate); }
      if (centerLabel) { centerLabel.addEventListener('mousemove', activate); centerLabel.addEventListener('mouseleave', deactivate); }
    });
    return { hiddenLabels: items.filter(function (it) { return it.hideLabel; }).length };
  }

  function renderForecastChart(container, opts) {
    // opts: {categories, actual:[], proyectado:[], meta:[]?, formatValue, barColor}
    // A "ghost bar" chart: for the current (in-progress) month and every future month,
    // a wide, dashed, low-opacity bar shows the linear-trend projection; closed months
    // and the current month's actual-to-date show as a solid, narrower bar drawn on
    // top of it — so real and projected are always visually distinguishable at a glance.
    chartUid++;
    var W = chartWidth(container, 720), H = 270, padL = 10, padR = 12, padT = 18, padB = 26;
    var plotW = W - padL - padR, plotH = H - padT - padB;
    var n = opts.categories.length;
    var bandW = plotW / n;
    var wGhost = Math.max(10, Math.min(34, bandW * 0.62));
    var wSolid = wGhost * 0.5;

    var all = [].concat(opts.actual, opts.proyectado, opts.meta || []).filter(isNum);
    var maxV = all.length ? Math.max.apply(null, all) * 1.18 : 1;
    var minV = all.length ? Math.min(0, Math.min.apply(null, all)) : 0;
    if (minV < 0) minV = minV * 1.15;

    function y(v) { return padT + plotH - ((v - minV) / (maxV - minV)) * plotH; }
    var baseline = y(0);

    var gridHtml = '', labelsHtml = '';
    for (var g = 0; g <= 4; g++) {
      var v = minV + (maxV - minV) * g / 4;
      var yy = y(v);
      gridHtml += '<line class="grid-line" x1="' + padL + '" x2="' + (W - padR) + '" y1="' + yy + '" y2="' + yy + '"></line>';
      labelsHtml += '<text class="axis-label" x="' + padL + '" y="' + (yy - 3) + '">' + opts.formatValue(v, true) + '</text>';
    }

    function cxOf(i) { return padL + bandW * i + bandW / 2; }
    var ghostHtml = '', solidHtml = '', metaHtml = '', xLabelsHtml = '', valueLabelsHtml = '';
    var pts = [];
    opts.categories.forEach(function (cat, i) {
      var cx = cxOf(i);
      xLabelsHtml += '<text class="axis-label" x="' + cx + '" y="' + (H - 8) + '" text-anchor="middle">' + cat + '</text>';
      var proy = opts.proyectado[i], act = opts.actual[i];
      if (isNum(proy)) {
        var topG = Math.min(baseline, y(proy)), hG = Math.max(1, Math.abs(y(proy) - baseline));
        ghostHtml += '<rect x="' + (cx - wGhost / 2) + '" y="' + topG + '" width="' + wGhost + '" height="' + hG +
          '" rx="3" fill="' + opts.barColor + '" fill-opacity="0.22" stroke="' + opts.barColor + '" stroke-width="1.3" stroke-dasharray="3 2"></rect>';
        if (!isNum(act)) valueLabelsHtml += '<text class="value-label" x="' + cx + '" y="' + (topG - 5) + '" text-anchor="middle" opacity=".75">≈' + opts.formatValue(proy) + '</text>';
      }
      if (isNum(act)) {
        var top = Math.min(baseline, y(act)), h = Math.max(1, Math.abs(y(act) - baseline));
        solidHtml += '<rect class="bar" data-i="' + i + '" x="' + (cx - wSolid / 2) + '" y="' + top + '" width="' + wSolid + '" height="' + h + '" rx="2.5" fill="' + opts.barColor + '"></rect>';
        valueLabelsHtml += '<text class="value-label" x="' + cx + '" y="' + (top - 5) + '" text-anchor="middle">' + opts.formatValue(act) + '</text>';
      }
      if (opts.meta && isNum(opts.meta[i])) {
        var ym = y(opts.meta[i]);
        metaHtml += '<line x1="' + (cx - wGhost / 2 - 3) + '" x2="' + (cx + wGhost / 2 + 3) + '" y1="' + ym + '" y2="' + ym + '" stroke="var(--text-primary)" stroke-width="2"></line>';
      }
      pts.push({ i: i, cx: cx, cat: cat, act: act, proy: proy, meta: opts.meta ? opts.meta[i] : null });
    });

    var legendHtml = '<div class="chart-legend">' +
      '<span class="sw"><span class="dot" style="background:' + opts.barColor + '"></span>Real</span>' +
      '<span class="sw"><span class="dot" style="background:' + opts.barColor + ';opacity:.4;border:1px dashed ' + opts.barColor + ';"></span>Proyectado</span>' +
      (opts.meta ? '<span class="sw"><span class="dot" style="background:var(--text-primary)"></span>Meta / Ppto</span>' : '') + '</div>';

    var tipId = 'tip' + chartUid;
    container.innerHTML = '<div class="chart-wrap">' + legendHtml +
      '<svg class="chart" viewBox="0 0 ' + W + ' ' + H + '" id="svg' + chartUid + '">' +
      '<line class="baseline" x1="' + padL + '" x2="' + (W - padR) + '" y1="' + baseline + '" y2="' + baseline + '"></line>' +
      gridHtml + labelsHtml + ghostHtml + solidHtml + metaHtml + valueLabelsHtml + xLabelsHtml +
      '</svg><div class="chart-tooltip" id="' + tipId + '"></div></div>';

    var tip = container.querySelector('#' + tipId);
    var wrapEl = container.querySelector('.chart-wrap');
    container.querySelectorAll('.bar').forEach(function (elm) {
      var p = pts[parseInt(elm.getAttribute('data-i'), 10)];
      if (!p) return;
      elm.addEventListener('mousemove', function () {
        var rect = wrapEl.getBoundingClientRect();
        var scale = rect.width / W;
        tip.style.left = (p.cx * scale) + 'px';
        tip.style.top = ((padT + 4) * scale) + 'px';
        tip.style.opacity = 1;
        tip.innerHTML = '<strong>' + p.cat + '</strong>' +
          (isNum(p.act) ? '<br>Real: ' + opts.formatValue(p.act) : '') +
          (isNum(p.proy) ? '<br>Proyectado: ' + opts.formatValue(p.proy) : '') +
          (isNum(p.meta) ? '<br>Meta/Ppto: ' + opts.formatValue(p.meta) : '');
      });
      elm.addEventListener('mouseleave', function () { tip.style.opacity = 0; });
    });
  }

  function renderComboChart(container, opts) {
    // opts: {categories, barName, barColor, barValues, metaName, metaValues,
    //        desvName, desvValues, formatValue}
    // The requested "todo en uno": the real value as bars, its meta/budget as a
    // line, and the deviation as a line WITH points — all three share one axis
    // (same unit), which is what makes the overlay legitimate. The axis extends
    // below zero when deviations go negative.
    chartUid++;
    var W = chartWidth(container, 720), H = 280, padL = 10, padR = 12, padT = 20, padB = 26;
    var plotW = W - padL - padR, plotH = H - padT - padB;
    var n = opts.categories.length;
    var bandW = plotW / n;
    var barW = Math.max(8, Math.min(30, bandW * 0.52));

    var all = [].concat(opts.barValues, opts.metaValues, opts.desvValues).filter(isNum);
    var maxV = all.length ? Math.max.apply(null, all) : 1;
    var minV = all.length ? Math.min(0, Math.min.apply(null, all)) : 0;
    maxV = maxV * 1.18 || 1;
    if (minV < 0) minV = minV * 1.15;

    function y(v) { return padT + plotH - ((v - minV) / (maxV - minV)) * plotH; }
    var baseline = y(0);

    var gridHtml = '', labelsHtml = '';
    for (var g = 0; g <= 4; g++) {
      var v = minV + (maxV - minV) * g / 4;
      var yy = y(v);
      gridHtml += '<line class="grid-line" x1="' + padL + '" x2="' + (W - padR) + '" y1="' + yy + '" y2="' + yy + '"></line>';
      labelsHtml += '<text class="axis-label" x="' + padL + '" y="' + (yy - 3) + '">' + opts.formatValue(v, true) + '</text>';
    }

    function cxOf(i) { return padL + bandW * i + bandW / 2; }
    var barsHtml = '', xLabelsHtml = '', valueLabelsHtml = '';
    var pts = [];
    opts.categories.forEach(function (cat, i) {
      var cx = cxOf(i);
      xLabelsHtml += '<text class="axis-label" x="' + cx + '" y="' + (H - 8) + '" text-anchor="middle">' + cat + '</text>';
      var val = opts.barValues[i];
      if (isNum(val)) {
        var top = Math.min(baseline, y(val)), h = Math.max(1, Math.abs(y(val) - baseline));
        barsHtml += '<rect class="bar" data-i="' + i + '" x="' + (cx - barW / 2) + '" y="' + top + '" width="' + barW + '" height="' + h + '" rx="3" fill="' + opts.barColor + '"></rect>';
        valueLabelsHtml += '<text class="value-label" x="' + cx + '" y="' + (top - 5) + '" text-anchor="middle">' + opts.formatValue(val) + '</text>';
      }
      pts.push({ i: i, cx: cx, cat: cat, val: val, meta: opts.metaValues[i], desv: opts.desvValues[i] });
    });

    function linePath(values) {
      var d = '', started = false;
      values.forEach(function (v, i) {
        if (!isNum(v)) { started = false; return; }
        d += (started ? 'L' : 'M') + cxOf(i) + ',' + y(v);
        started = true;
      });
      return d;
    }
    var metaHtml = '<path d="' + linePath(opts.metaValues) + '" fill="none" stroke="var(--text-primary)" stroke-width="1.8" stroke-dasharray="5 3" opacity=".8"></path>';
    var lineValueLabelsHtml = '';
    opts.metaValues.forEach(function (v, i) {
      if (!isNum(v)) return;
      metaHtml += '<circle cx="' + cxOf(i) + '" cy="' + y(v) + '" r="2.6" fill="var(--text-primary)"></circle>';
      lineValueLabelsHtml += '<text class="value-label" x="' + cxOf(i) + '" y="' + (y(v) - 8) + '" text-anchor="middle" font-size="9.5" fill="var(--text-primary)">' + opts.formatValue(v) + '</text>';
    });
    var desvHtml = '<path d="' + linePath(opts.desvValues) + '" fill="none" stroke="var(--series-8)" stroke-width="2"></path>';
    opts.desvValues.forEach(function (v, i) {
      if (!isNum(v)) return;
      desvHtml += '<circle class="bar" data-i="' + i + '" cx="' + cxOf(i) + '" cy="' + y(v) + '" r="3.6" fill="var(--series-8)" stroke="var(--surface-1)" stroke-width="1.5"></circle>';
      lineValueLabelsHtml += '<text class="value-label" x="' + cxOf(i) + '" y="' + (y(v) + 14) + '" text-anchor="middle" font-size="9.5" fill="var(--series-8)">' + opts.formatValue(v) + '</text>';
    });

    var legendHtml = '<div class="chart-legend">' +
      '<span class="sw"><span class="dot" style="background:' + opts.barColor + '"></span>' + escapeHtml(opts.barName) + '</span>' +
      '<span class="sw"><span class="dot" style="background:var(--text-primary)"></span>' + escapeHtml(opts.metaName) + '</span>' +
      '<span class="sw"><span class="dot" style="background:var(--series-8)"></span>' + escapeHtml(opts.desvName) + '</span></div>';

    var tipId = 'tip' + chartUid;
    container.innerHTML = '<div class="chart-wrap">' + legendHtml +
      '<svg class="chart" viewBox="0 0 ' + W + ' ' + H + '" id="svg' + chartUid + '">' +
      '<line class="baseline" x1="' + padL + '" x2="' + (W - padR) + '" y1="' + baseline + '" y2="' + baseline + '"></line>' +
      gridHtml + labelsHtml + barsHtml + metaHtml + desvHtml + valueLabelsHtml + lineValueLabelsHtml + xLabelsHtml +
      '</svg><div class="chart-tooltip" id="' + tipId + '"></div></div>';

    var tip = container.querySelector('#' + tipId);
    var wrapEl = container.querySelector('.chart-wrap');
    container.querySelectorAll('.bar').forEach(function (elm) {
      var p = pts[parseInt(elm.getAttribute('data-i'), 10)];
      if (!p) return;
      elm.addEventListener('mousemove', function () {
        var rect = wrapEl.getBoundingClientRect();
        var scale = rect.width / W;
        tip.style.left = (p.cx * scale) + 'px';
        tip.style.top = ((padT + 6) * scale) + 'px';
        tip.style.opacity = 1;
        tip.innerHTML = '<strong>' + p.cat + '</strong>' +
          (isNum(p.val) ? '<br>' + escapeHtml(opts.barName) + ': ' + opts.formatValue(p.val) : '') +
          (isNum(p.meta) ? '<br>' + escapeHtml(opts.metaName) + ': ' + opts.formatValue(p.meta) : '') +
          (isNum(p.desv) ? '<br>' + escapeHtml(opts.desvName) + ': ' + opts.formatValue(p.desv) : '');
      });
      elm.addEventListener('mouseleave', function () { tip.style.opacity = 0; });
    });
  }

  function renderTopLineBottomBarChart(container, opts) {
    // opts: {categories, lines: [{name,color,values}, ...] (1-2, dashed, plotted as-is —
    //        typically real/esperado, well above zero), bars: [{name,color,values,signed}, ...]
    //        (1-2, grouped, extend up/down from the zero baseline — typically desviaciones),
    //        totalLine: {name,color,values} (optional solid line+dots overlaid on the bars,
    //        e.g. desviación total), formatValue}
    // Same one-axis-for-everything approach as renderComboChart: legitimate because every
    // series here shares one unit (%), just at very different magnitudes (lines ~5%, bar
    // deviations ~±1%) — matches the reference "Calidad Planta" chart's layout exactly.
    chartUid++;
    var W = chartWidth(container, 720), H = 300, padL = 10, padR = 12, padT = 20, padB = 26;
    var plotW = W - padL - padR, plotH = H - padT - padB;
    var n = opts.categories.length;
    var bandW = plotW / n;
    var barCount = opts.bars.length;
    var groupW = Math.min(bandW * 0.6, 46);
    var barW = Math.max(6, groupW / barCount - 3);

    var allVals = [];
    opts.lines.forEach(function (l) { allVals = allVals.concat(l.values.filter(isNum)); });
    opts.bars.forEach(function (b) { allVals = allVals.concat(b.values.filter(isNum)); });
    if (opts.totalLine) allVals = allVals.concat(opts.totalLine.values.filter(isNum));
    var maxV = allVals.length ? Math.max.apply(null, allVals) : 1;
    var minV = allVals.length ? Math.min(0, Math.min.apply(null, allVals)) : 0;
    maxV = maxV * 1.15 || 1;
    minV = minV < 0 ? minV * 1.4 : -maxV * 0.3;

    function y(v) { return padT + plotH - ((v - minV) / (maxV - minV)) * plotH; }
    var baseline = y(0);

    var gridHtml = '', labelsHtml = '';
    for (var g = 0; g <= 4; g++) {
      var v = minV + (maxV - minV) * g / 4;
      var yy = y(v);
      gridHtml += '<line class="grid-line" x1="' + padL + '" x2="' + (W - padR) + '" y1="' + yy + '" y2="' + yy + '"></line>';
      labelsHtml += '<text class="axis-label" x="' + padL + '" y="' + (yy - 3) + '">' + opts.formatValue(v, true) + '</text>';
    }

    function cxOf(i) { return padL + bandW * i + bandW / 2; }
    var xLabelsHtml = '';
    var pts = [];
    opts.categories.forEach(function (cat, i) {
      xLabelsHtml += '<text class="axis-label" x="' + cxOf(i) + '" y="' + (H - 8) + '" text-anchor="middle">' + cat + '</text>';
      pts.push({ i: i, cx: cxOf(i), cat: cat });
    });

    // barras agrupadas alrededor de cada categoría
    var barsHtml = '', barValueLabelsHtml = '';
    opts.bars.forEach(function (series, si) {
      var groupLeft = function (cx) { return cx - groupW / 2 + si * (barW + 3); };
      series.values.forEach(function (val, i) {
        if (!isNum(val)) return;
        var cx = cxOf(i), x0 = groupLeft(cx);
        var top = Math.min(baseline, y(val)), h = Math.max(1, Math.abs(y(val) - baseline));
        var color = series.signed ? (val >= 0 ? 'var(--good)' : 'var(--critical)') : series.color;
        barsHtml += '<rect class="bar" data-series="b' + si + '" data-i="' + i + '" x="' + x0 + '" y="' + top + '" width="' + barW + '" height="' + h + '" rx="2.5" fill="' + color + '"></rect>';
        var labelY = val >= 0 ? top - 4 : top + h + 11;
        barValueLabelsHtml += '<text class="value-label" x="' + (x0 + barW / 2) + '" y="' + labelY + '" text-anchor="middle" font-size="9.5">' + opts.formatValue(val) + '</text>';
        pts[i]['bar' + si] = val;
      });
    });

    function linePath(values) {
      var d = '', started = false;
      values.forEach(function (v, i) {
        if (!isNum(v)) { started = false; return; }
        d += (started ? 'L' : 'M') + cxOf(i) + ',' + y(v);
        started = true;
      });
      return d;
    }
    var linesHtml = '', lineValueLabelsHtml = '';
    opts.lines.forEach(function (l, li) {
      linesHtml += '<path d="' + linePath(l.values) + '" fill="none" stroke="' + l.color + '" stroke-width="1.8" stroke-dasharray="5 3" opacity=".9"></path>';
      var dy = li === 0 ? -8 : 14;
      l.values.forEach(function (v, i) {
        if (!isNum(v)) return;
        linesHtml += '<circle cx="' + cxOf(i) + '" cy="' + y(v) + '" r="2.6" fill="' + l.color + '"></circle>';
        lineValueLabelsHtml += '<text class="value-label" x="' + cxOf(i) + '" y="' + (y(v) + dy) + '" text-anchor="middle" font-size="9.5" fill="' + l.color + '">' + opts.formatValue(v) + '</text>';
        pts[i]['line' + li] = v;
      });
    });

    var totalHtml = '';
    if (opts.totalLine) {
      totalHtml += '<path d="' + linePath(opts.totalLine.values) + '" fill="none" stroke="' + opts.totalLine.color + '" stroke-width="2"></path>';
      opts.totalLine.values.forEach(function (v, i) {
        if (!isNum(v)) return;
        totalHtml += '<circle class="bar" data-series="total" data-i="' + i + '" cx="' + cxOf(i) + '" cy="' + y(v) + '" r="3.6" fill="' + opts.totalLine.color + '" stroke="var(--surface-1)" stroke-width="1.5"></circle>';
        pts[i].total = v;
      });
    }

    var legendHtml = '<div class="chart-legend">' +
      opts.bars.map(function (b) { return '<span class="sw"><span class="dot" style="background:' + (b.signed ? 'var(--good)' : b.color) + '"></span>' + escapeHtml(b.name) + '</span>'; }).join('') +
      opts.lines.map(function (l) { return '<span class="sw"><span class="dot" style="background:' + l.color + '"></span>' + escapeHtml(l.name) + '</span>'; }).join('') +
      (opts.totalLine ? '<span class="sw"><span class="dot" style="background:' + opts.totalLine.color + '"></span>' + escapeHtml(opts.totalLine.name) + '</span>' : '') +
      '</div>';

    var tipId = 'tip' + chartUid;
    container.innerHTML = '<div class="chart-wrap">' + legendHtml +
      '<svg class="chart" viewBox="0 0 ' + W + ' ' + H + '" id="svg' + chartUid + '">' +
      '<line class="baseline" x1="' + padL + '" x2="' + (W - padR) + '" y1="' + baseline + '" y2="' + baseline + '"></line>' +
      gridHtml + labelsHtml + barsHtml + linesHtml + totalHtml + barValueLabelsHtml + lineValueLabelsHtml + xLabelsHtml +
      '</svg><div class="chart-tooltip" id="' + tipId + '"></div></div>';

    var tip = container.querySelector('#' + tipId);
    var wrapEl = container.querySelector('.chart-wrap');
    container.querySelectorAll('[data-i]').forEach(function (elm) {
      var p = pts[parseInt(elm.getAttribute('data-i'), 10)];
      if (!p) return;
      elm.addEventListener('mousemove', function () {
        var rect = wrapEl.getBoundingClientRect();
        var scale = rect.width / W;
        tip.style.left = (p.cx * scale) + 'px';
        tip.style.top = ((padT + 6) * scale) + 'px';
        tip.style.opacity = 1;
        var html = '<strong>' + p.cat + '</strong>';
        opts.lines.forEach(function (l, li) { if (isNum(p['line' + li])) html += '<br>' + escapeHtml(l.name) + ': ' + opts.formatValue(p['line' + li]); });
        opts.bars.forEach(function (b, bi) { if (isNum(p['bar' + bi])) html += '<br>' + escapeHtml(b.name) + ': ' + opts.formatValue(p['bar' + bi]); });
        if (opts.totalLine && isNum(p.total)) html += '<br>' + escapeHtml(opts.totalLine.name) + ': ' + opts.formatValue(p.total);
        tip.innerHTML = html;
      });
      elm.addEventListener('mouseleave', function () { tip.style.opacity = 0; });
    });
  }

  function renderDeviationChart(container, opts) {
    // opts: {categories, values, formatValue, seriesLabel, goodIsPositive}
    chartUid++;
    var W = chartWidth(container, 640), H = 220, padL = 8, padR = 8, padT = 20, padB = 26;
    var plotW = W - padL - padR, plotH = H - padT - padB;
    var n = opts.categories.length;
    var bandW = plotW / n;
    var barW = Math.max(8, Math.min(30, bandW * 0.55));
    var goodIsPositive = opts.goodIsPositive !== false;

    var allVals = opts.values.filter(isNum);
    var maxAbs = allVals.length ? Math.max.apply(null, allVals.map(Math.abs)) : 1;
    maxAbs = maxAbs * 1.25 || 1;
    var minV = -maxAbs, maxV = maxAbs;

    function y(v) { return padT + plotH - ((v - minV) / (maxV - minV)) * plotH; }
    var baseline = y(0);

    var gridLines = 4, gridHtml = '', labelsHtml = '';
    for (var g = 0; g <= gridLines; g++) {
      var v = minV + (maxV - minV) * g / gridLines;
      var yy = y(v);
      gridHtml += '<line class="grid-line" x1="' + padL + '" x2="' + (W - padR) + '" y1="' + yy + '" y2="' + yy + '"></line>';
      labelsHtml += '<text class="axis-label" x="' + padL + '" y="' + (yy - 3) + '">' + opts.formatValue(v, true) + '</text>';
    }

    var barsHtml = '', xLabelsHtml = '', valueLabelsHtml = '';
    var tipId = 'tip' + chartUid;
    var bars = [];
    opts.categories.forEach(function (cat, i) {
      var cx = padL + bandW * i + bandW / 2;
      var val = opts.values[i];
      xLabelsHtml += '<text class="axis-label" x="' + cx + '" y="' + (H - 8) + '" text-anchor="middle">' + cat + '</text>';
      if (isNum(val)) {
        var good = goodIsPositive ? val >= 0 : val <= 0;
        var color = good ? 'var(--good)' : 'var(--critical)';
        var y0 = baseline, y1 = y(val);
        var top = Math.min(y0, y1), h = Math.max(1, Math.abs(y1 - y0));
        var r = Math.min(4, barW / 2, h);
        var x0 = cx - barW / 2;
        var d = 'M' + x0 + ',' + (top + h) +
          'L' + x0 + ',' + (top + r) +
          'Q' + x0 + ',' + top + ' ' + (x0 + r) + ',' + top +
          'L' + (x0 + barW - r) + ',' + top +
          'Q' + (x0 + barW) + ',' + top + ' ' + (x0 + barW) + ',' + (top + r) +
          'L' + (x0 + barW) + ',' + (top + h) + 'Z';
        barsHtml += '<path class="bar" data-i="' + i + '" d="' + d + '" fill="' + color + '"></path>';
        var labelY = val >= 0 ? top - 6 : top + h + 12;
        valueLabelsHtml += '<text class="value-label" data-i="' + i + '" x="' + cx + '" y="' + labelY + '" text-anchor="middle" fill="' + color + '">' + opts.formatValue(val) + '</text>';
        bars.push({ i: i, cx: cx, val: val, cat: cat, good: good });
      }
    });

    var legendHtml = '<div class="chart-legend">' +
      '<span class="sw"><span class="dot" style="background:var(--good)"></span>Favorable</span>' +
      '<span class="sw"><span class="dot" style="background:var(--critical)"></span>Desfavorable</span>' +
      '</div>';

    container.innerHTML = '<div class="chart-wrap">' + legendHtml +
      '<svg class="chart" viewBox="0 0 ' + W + ' ' + H + '" id="svg' + chartUid + '">' +
      '<line class="baseline" x1="' + padL + '" x2="' + (W - padR) + '" y1="' + baseline + '" y2="' + baseline + '"></line>' +
      gridHtml + labelsHtml + barsHtml + valueLabelsHtml + xLabelsHtml +
      '</svg><div class="chart-tooltip" id="' + tipId + '"></div></div>';

    var tip = container.querySelector('#' + tipId);
    var wrapEl = container.querySelector('.chart-wrap');
    container.querySelectorAll('.bar').forEach(function (barEl) {
      var i = parseInt(barEl.getAttribute('data-i'), 10);
      var b = bars.filter(function (x) { return x.i === i; })[0];
      if (!b) return;
      var label = container.querySelector('.value-label[data-i="' + i + '"]');
      function activate() {
        var rect = wrapEl.getBoundingClientRect();
        var scale = rect.width / W;
        tip.style.left = (b.cx * scale) + 'px';
        tip.style.top = (y(b.val) * scale) + 'px';
        tip.style.opacity = 1;
        tip.innerHTML = '<strong>' + b.cat + '</strong><br>' + opts.seriesLabel + ': ' + opts.formatValue(b.val) +
          '<br>' + (b.good ? 'Favorable' : 'Desfavorable');
        barEl.classList.add('bar-active');
        if (label) label.classList.add('value-label-active');
      }
      function deactivate() {
        tip.style.opacity = 0;
        barEl.classList.remove('bar-active');
        if (label) label.classList.remove('value-label-active');
      }
      barEl.addEventListener('mousemove', activate);
      barEl.addEventListener('mouseleave', deactivate);
      if (label) { label.addEventListener('mousemove', activate); label.addEventListener('mouseleave', deactivate); }
    });
  }

  var SERIES_COLORS = ['var(--series-1)', 'var(--series-2)', 'var(--series-3)', 'var(--series-4)',
    'var(--series-5)', 'var(--series-6)', 'var(--series-7)', 'var(--series-8)'];

  /* =========================================================================
   * Table rendering helpers
   * ======================================================================= */

  function renderMonthlyTable(container, title, rows, cols) {
    // rows: [{label, unit, values:[12], fmt}]
    var html = '<table class="wide"><thead><tr><th>' + escapeHtml(title) + '</th><th></th>' +
      MESES_ABBR.map(function (m) { return '<th>' + m + '</th>'; }).join('') + '</tr></thead><tbody>';
    rows.forEach(function (r) {
      html += '<tr' + (r.total ? ' class="total"' : '') + '>' +
        '<td>' + escapeHtml(r.label) + '</td><td class="unit">' + (r.unit || '') + '</td>' +
        r.values.map(function (v) { return tdv(v, r.fmt || fmtInt); }).join('') + '</tr>';
    });
    html += '</tbody></table>';
    container.innerHTML = html;
  }

  function renderMachineTable(container, title, machines, rows) {
    var html = '<table class="wide"><thead><tr><th>' + escapeHtml(title) + '</th><th></th>' +
      machines.map(function (m) { return '<th>' + escapeHtml(m) + '</th>'; }).join('') + '</tr></thead><tbody>';
    rows.forEach(function (r) {
      html += '<tr' + (r.total ? ' class="total"' : '') + '>' +
        '<td>' + escapeHtml(r.label) + '</td><td class="unit">' + (r.unit || '') + '</td>' +
        r.values.map(function (v) { return tdv(v, r.fmt || fmtInt); }).join('') + '</tr>';
    });
    html += '</tbody></table>';
    container.innerHTML = html;
  }

  /* =========================================================================
   * App wiring
   * ======================================================================= */

  var STATE = {
    data: null, engine: null, year: null, month: null, machine: null, fileMeta: null,
    restored: false,
    sku: { maquina: null, mes: null, query: '' },
    productos: { maquina: null, mes: null, espesor: null },
    insumos: { maquina: null, mes: null },
    criticos: { mes: null, maquina: null, estado: null },
    proyecciones: { maquina: null },
    oeeCalidad: { maquina: null, mes: null }
  };

  function el(id) { return document.getElementById(id); }

  function fmtPctTick(v) { return (v * 100).toLocaleString('es-CL', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + '%'; }
  function fmtKgTick(v) {
    if (Math.abs(v) >= 1e6) return (v / 1e6).toFixed(1) + 'M';
    if (Math.abs(v) >= 1e3) return Math.round(v / 1e3) + 'k';
    return Math.round(v);
  }
  function fmtMoneyTick(v, isAxisTick) {
    if (!isAxisTick) return fmtMoney(v);
    if (Math.abs(v) >= 1e6) return '$' + (v / 1e6).toFixed(1) + 'M';
    if (Math.abs(v) >= 1e3) return '$' + Math.round(v / 1e3) + 'k';
    return fmtMoney(v);
  }
  // costo/kg (y su meta/presupuesto y desviación) siempre en 2 decimales — a
  // diferencia de los montos totales en pesos, es una tasa chica donde los
  // centavos importan para comparar máquinas o meses.
  function fmtMoney2(n) { return n == null ? '-' : '$' + n.toLocaleString('es-CL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  function fmtMoneyTick2(v, isAxisTick) {
    if (!isAxisTick) return fmtMoney2(v);
    if (Math.abs(v) >= 1e6) return '$' + (v / 1e6).toFixed(1) + 'M';
    if (Math.abs(v) >= 1e3) return '$' + Math.round(v / 1e3) + 'k';
    return fmtMoney2(v);
  }

  var ICON_GAUGE = '<path d="M12 21a9 9 0 1 1 9-9"></path><line x1="12" y1="12" x2="16" y2="8"></line><circle cx="12" cy="12" r="1"></circle>';
  var ICON_BOX = '<rect x="4" y="7" width="16" height="13" rx="1.5"></rect><path d="M4 7l8-4 8 4"></path><path d="M12 12v8"></path>';
  var ICON_TRASH = '<path d="M4 7h16"></path><path d="M10 11v6"></path><path d="M14 11v6"></path><path d="M6 7l1 12a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2l1-12"></path><path d="M9 7V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v3"></path>';
  var ICON_GEAR = '<circle cx="12" cy="12" r="3"></circle><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"></path>';
  var ICON_LAYERS = '<path d="M12 2 2 7l10 5 10-5-10-5z"></path><path d="M2 17l10 5 10-5"></path><path d="M2 12l10 5 10-5"></path>';
  var ICON_TARGET = '<circle cx="12" cy="12" r="9"></circle><circle cx="12" cy="12" r="5"></circle><circle cx="12" cy="12" r="1"></circle>';
  var ICON_TREND = '<polyline points="3 17 9 11 13 15 21 7"></polyline><polyline points="14 7 21 7 21 14"></polyline>';

  function renderKPIs(kpi) {
    var tiles = [
      { label: 'OEE Planta Global', value: fmtPct(kpi.oeePlantaGlobal, 1), icon: ICON_GAUGE, accent: 'var(--series-5)' },
      { label: 'Kilos Fabricados', value: fmtInt(kpi.kilosFabricados) + ' kg', icon: ICON_BOX, accent: 'var(--series-1)' },
      { label: 'Chatarra', value: fmtInt(kpi.chatarraKg) + ' kg', icon: ICON_TRASH, accent: 'var(--series-6)' },
      { label: 'Chatarra Máquinas', value: fmtPct(kpi.chatarraMaquinasPct, 2), icon: ICON_GEAR, accent: 'var(--series-8)' },
      { label: 'Chatarra Braner', value: fmtPct(kpi.chatarraBranerPct, 2), icon: ICON_LAYERS, accent: 'var(--series-7)' },
      { label: 'Chatarra Planta', value: fmtPct(kpi.chatarraPlantaPct, 2), icon: ICON_TARGET, accent: 'var(--series-2)' },
      { label: 'Meta Administrativa', value: fmtPct(kpi.metaChatarra, 1), icon: ICON_TARGET, accent: 'var(--series-4)' },
      {
        label: 'Desviación Administrativa', value: fmtSigned(kpi.desviacionChatarra, function (v) { return fmtPct(v, 2); }),
        icon: ICON_TREND, accent: 'var(--series-3)',
        deltaClass: deltaClass(kpi.desviacionChatarra, true),
        deltaText: kpi.desviacionChatarra == null ? '' : (kpi.desviacionChatarra >= 0 ? 'sobre la meta' : 'bajo la meta')
      },
      { label: 'Meta Calculada Máquinas', value: fmtPct(kpi.metaChatarraCalculada, 2), icon: ICON_TARGET, accent: 'var(--series-6)' },
      { label: 'Meta Calculada', value: fmtPct(kpi.metaChatarraCalculadaTotal, 2), icon: ICON_TARGET, accent: 'var(--series-6)' },
      {
        label: 'Desviación Calculada', value: fmtSigned(kpi.desviacionChatarraCalculadaTotal, function (v) { return fmtPct(v, 2); }),
        icon: ICON_TREND, accent: 'var(--series-3)',
        deltaClass: deltaClass(kpi.desviacionChatarraCalculadaTotal, true),
        deltaText: kpi.desviacionChatarraCalculadaTotal == null ? '' : (kpi.desviacionChatarraCalculadaTotal >= 0 ? 'sobre la meta' : 'bajo la meta')
      }
    ];
    el('kpiGrid').innerHTML = tiles.map(function (t) {
      return '<div class="tile" style="--accent:' + t.accent + '">' +
        '<div class="tile-top"><span class="icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' + t.icon + '</svg></span></div>' +
        '<div class="label">' + t.label + '</div><div class="value">' + t.value + '</div>' +
        (t.deltaClass ? '<div class="delta ' + t.deltaClass + '">' + (t.deltaText || '') + '</div>' : '') + '</div>';
    }).join('');
  }

  function renderResumen(anio) {
    var resumen = STATE.engine.resumenPlanta(anio);
    var rows = [
      { label: 'Total Producción', unit: 'kg', values: resumen.map(function (m) { return m.totalProduccion; }), total: true },
      { label: 'Total Chatarra Real', unit: 'kg', values: resumen.map(function (m) { return m.totalChatarraReal; }), total: true },
      { label: 'Chatarra Tuberas & Perfiladoras', unit: '%', values: resumen.map(function (m) { return m.chatarraMaquinasPct; }), fmt: function (v) { return fmtPct(v, 2); } },
      { label: 'Desorillado Braner', unit: '%', values: resumen.map(function (m) { return m.desorilladoBranerPct; }), fmt: function (v) { return fmtPct(v, 2); } },
      { label: 'Chatarra Planta', unit: '%', values: resumen.map(function (m) { return m.pctChatarraPlanta; }), fmt: function (v) { return fmtPct(v, 2); } },
      { label: 'OEE Planta Mensual', unit: '%', values: resumen.map(function (m) { return m.oeeMensual; }), fmt: function (v) { return fmtPct(v, 1); } }
    ];
    renderMonthlyTable(el('resumenTable'), 'Planta ' + anio, rows);
    return resumen;
  }

  function renderTendencias(anio, resumen) {
    var tend = STATE.engine.tendenciasPlanta(anio);
    renderComboChart(el('chartProdTrend'), {
      categories: tend.map(function (m) { return m.mesAbbr; }),
      barName: 'Costo/kg real', barColor: 'var(--series-1)',
      barValues: tend.map(function (m) { return m.costoKg; }),
      metaName: 'Meta presupuesto', metaValues: tend.map(function (m) { return m.costoMeta; }),
      desvName: 'Desviación (meta − real)', desvValues: tend.map(function (m) { return m.costoDesv; }),
      formatValue: fmtMoneyTick2
    });

    var metaSerie = resumen.map(function (m) { return STATE.engine.oeeMetaVal('OEE Planta', m.mes); });
    renderComboChart(el('chartOEETrend'), {
      categories: resumen.map(function (m) { return m.mesAbbr; }),
      barName: 'OEE real', barColor: 'var(--series-2)',
      barValues: resumen.map(function (m) { return m.oeeMensual; }),
      metaName: 'Meta', metaValues: metaSerie,
      desvName: 'Desviación (real − meta)',
      desvValues: resumen.map(function (m, i) {
        return (isNum(m.oeeMensual) && isNum(metaSerie[i])) ? m.oeeMensual - metaSerie[i] : null;
      }),
      formatValue: fmtPctTick
    });
  }

  function renderPresupuesto(anio, resumen) {
    var ppto = STATE.engine.presupuesto(anio, resumen);

    renderComboChart(el('chartProdPpto'), {
      categories: ppto.produccion.map(function (m) { return m.mesAbbr; }),
      barName: 'Producción real', barColor: 'var(--series-1)',
      barValues: ppto.produccion.map(function (m) { return m.real; }),
      metaName: 'Presupuesto', metaValues: ppto.produccion.map(function (m) { return m.ppto; }),
      desvName: 'Desviación', desvValues: ppto.produccion.map(function (m) { return m.desv; }),
      formatValue: fmtKgTick
    });

    // el target dinámico por espesor (si el Excel trae la tabla de rechazo estándar)
    // reemplaza la meta fija mes a mes — se avisa en el subtítulo cuál está activo,
    // ya que la mayoría de los Excel de este reporte no traen esa tabla todavía.
    var hayMetaDinamica = ppto.chatarra.some(function (m) { return m.metaEsDinamica; });
    el('chatPptoCap').textContent = hayMetaDinamica ?
      '% chatarra planta vs. meta dinámica por espesor (se ajusta según el mix de espesores producido cada mes; meses sin tabla de rechazo estándar usan 4,6% fijo)' :
      '% chatarra planta vs. meta (4,6% fijo — este Excel no trae la tabla de rechazo estándar por espesor para calcular una meta dinámica)';
    renderComboChart(el('chartChatPpto'), {
      categories: ppto.chatarra.map(function (m) { return m.mesAbbr; }),
      barName: 'Chatarra planta', barColor: 'var(--series-6)',
      barValues: ppto.chatarra.map(function (m) { return m.pct; }),
      metaName: hayMetaDinamica ? 'Meta por espesor' : 'Meta (4,6%)', metaValues: ppto.chatarra.map(function (m) { return m.meta; }),
      desvName: 'Desviación (meta − real)', desvValues: ppto.chatarra.map(function (m) { return m.desv; }),
      formatValue: fmtPctTick
    });

    renderDeviationChart(el('chartDesvProd'), {
      categories: ppto.produccion.map(function (m) { return m.mesAbbr; }),
      values: ppto.produccion.map(function (m) { return m.desv; }),
      formatValue: fmtKgTick, seriesLabel: 'Desviación', goodIsPositive: true
    });
    renderDeviationChart(el('chartDesvChat'), {
      categories: ppto.chatarra.map(function (m) { return m.mesAbbr; }),
      values: ppto.chatarra.map(function (m) { return m.desv; }),
      formatValue: function (v) { return fmtPct(v, 2); }, seriesLabel: 'Desviación', goodIsPositive: true
    });

    var rows = [
      { label: 'Total Producción', unit: 'kg', values: ppto.produccion.map(function (m) { return m.real; }), total: true },
      { label: 'Presupuesto Producción', unit: 'kg', values: ppto.produccion.map(function (m) { return m.ppto; }) },
      { label: 'Desviación Ppto vs Real', unit: 'kg', values: ppto.produccion.map(function (m) { return m.desv; }), fmt: fmtInt },
      { label: 'Desviación Ppto vs Real', unit: '%', values: ppto.produccion.map(function (m) { return m.desvPct; }), fmt: function (v) { return fmtPct(v, 1); } },
      { label: 'Total Chatarra Real', unit: 'kg', values: ppto.chatarra.map(function (m) { return m.real; }), total: true },
      { label: 'Chatarra Planta', unit: '%', values: ppto.chatarra.map(function (m) { return m.pct; }), fmt: function (v) { return fmtPct(v, 2); } },
      { label: 'Meta Administrativa', unit: '%', values: ppto.chatarra.map(function (m) { return m.metaAdministrativa; }), fmt: function (v) { return fmtPct(v, 1); } },
      { label: 'Desviación Administrativa', unit: '%', values: ppto.chatarra.map(function (m) { return m.desvAdministrativa; }), fmt: function (v) { return fmtPct(v, 2); } },
      { label: 'Meta Calculada', unit: '%', values: ppto.chatarra.map(function (m) { return m.metaCalculada; }), fmt: function (v) { return fmtPct(v, 1); } },
      { label: 'Desviación Calculada', unit: '%', values: ppto.chatarra.map(function (m) { return m.desvCalculada; }), fmt: function (v) { return fmtPct(v, 2); } }
    ];
    renderMonthlyTable(el('pptoTable'), 'Producción vs. Presupuesto — ' + anio, rows);
  }

  function renderOeeCalidad(anio) {
    var maquina = STATE.oeeCalidad.maquina;
    var mesFiltro = STATE.oeeCalidad.mes;
    var maquinaLabel = maquina ? machineShortLabel(maquina, STATE.data.machineCodes) : 'Planta';
    var categories, oeeReal, oeeMeta, calReal, calEsperado, calDesvMaq, calDesvBran, calDesvTotal, hayDesglose, hayMetaDinamica;

    if (!maquina) {
      var resumen = STATE.engine.resumenPlanta(anio);
      categories = resumen.map(function (m) { return m.mesAbbr; });
      oeeReal = resumen.map(function (m) { return m.oeeMensual; });
      oeeMeta = resumen.map(function (m) { return STATE.engine.oeeMetaVal('OEE Planta', m.mes); });
      var cal = STATE.engine.calidadPlantaDetalle(anio);
      calReal = cal.map(function (m) { return m.real; });
      calEsperado = cal.map(function (m) { return m.esperado; });
      calDesvMaq = cal.map(function (m) { return m.desvMaquinas; });
      calDesvBran = cal.map(function (m) { return m.desvBraner; });
      calDesvTotal = cal.map(function (m) { return m.desvTotal; });
      hayMetaDinamica = cal.some(function (m) { return m.metaEsDinamica; });
      hayDesglose = true;
    } else {
      var det = STATE.engine.detalleAnio(anio, maquina);
      categories = det.meses.map(function (m) { return m.mesAbbr; });
      oeeReal = det.meses.map(function (m) { return m.oee; });
      oeeMeta = det.meses.map(function (m) { return m.oeeMeta; });
      calReal = det.meses.map(function (m) { return m.chatarraPct; });
      calEsperado = det.meses.map(function (m) { return m.metaChatarraEstandar; });
      calDesvTotal = det.meses.map(function (m) { return m.desviacionChatarraEstandar; });
      hayMetaDinamica = det.meses.some(function (m) { return m.metaChatarraEsCalculada; });
      hayDesglose = false;
    }
    var oeeDesv = oeeReal.map(function (v, i) { return (isNum(v) && isNum(oeeMeta[i])) ? v - oeeMeta[i] : null; });

    el('oeeCalidadOeeCap').textContent = 'OEE real vs. meta — ' + maquinaLabel;
    renderTopLineBottomBarChart(el('chartOeeCalidadOEE'), {
      categories: categories,
      lines: [
        { name: 'OEE Real', color: 'var(--series-2)', values: oeeReal },
        { name: 'OEE Meta', color: 'var(--series-5)', values: oeeMeta }
      ],
      bars: [{ name: 'Desviación OEE', signed: true, values: oeeDesv }],
      formatValue: fmtPctTick
    });

    if (hayDesglose) {
      el('oeeCalidadCalCap').textContent = hayMetaDinamica ?
        'Calidad Planta: chatarra máquinas + desorillado Braner vs. meta dinámica por espesor' :
        'Calidad Planta: chatarra máquinas + desorillado Braner vs. meta (4,6% fijo: 2,70% máquinas + 1,90% Braner)';
      renderTopLineBottomBarChart(el('chartOeeCalidadCalidad'), {
        categories: categories,
        lines: [
          { name: 'Calidad Planta Real', color: 'var(--series-1)', values: calReal },
          { name: 'Calidad Planta Esperado', color: 'var(--series-6)', values: calEsperado }
        ],
        bars: [
          { name: 'Desviación Máquinas', color: 'var(--series-3)', values: calDesvMaq },
          { name: 'Desviación Braner', color: 'var(--series-7)', values: calDesvBran }
        ],
        totalLine: { name: 'Desviación Total', color: 'var(--series-8)', values: calDesvTotal },
        formatValue: fmtPctTick
      });
    } else {
      el('oeeCalidadCalCap').textContent = 'Chatarra ' + maquinaLabel + ' vs. ' +
        (hayMetaDinamica ? 'meta calculada por mix de espesor producido' : 'meta estándar (sin tabla de rechazo para calcularla)') +
        ' — el desorillado es un concepto de planta, no se puede atribuir a una máquina individual';
      renderTopLineBottomBarChart(el('chartOeeCalidadCalidad'), {
        categories: categories,
        lines: [
          { name: 'Chatarra Real', color: 'var(--series-1)', values: calReal },
          { name: 'Meta', color: 'var(--series-6)', values: calEsperado }
        ],
        bars: [{ name: 'Desviación', signed: true, values: calDesvTotal }],
        formatValue: fmtPctTick
      });
    }

    var mesCapEl = el('oeeCalidadMesDetalle');
    if (mesFiltro) {
      var idx = MESES.indexOf(mesFiltro);
      mesCapEl.style.display = 'block';
      mesCapEl.innerHTML = '<strong>' + mesFiltro + '</strong> — OEE real ' + fmtPct(oeeReal[idx], 1) + ' vs. meta ' + fmtPct(oeeMeta[idx], 1) +
        ' (desv. ' + fmtSigned(oeeDesv[idx], function (v) { return fmtPct(v, 1); }) + ')' +
        ' &nbsp;·&nbsp; Calidad Planta real ' + fmtPct(calReal[idx], 2) + ' vs. esperado ' + fmtPct(calEsperado[idx], 2) +
        ' (desv. ' + fmtSigned(calDesvTotal[idx], function (v) { return fmtPct(v, 2); }) + ')';
    } else {
      mesCapEl.style.display = 'none';
      mesCapEl.innerHTML = '';
    }
  }

  function renderDetalleMes(anio, mesNombre) {
    var det = STATE.engine.detalleMes(anio, mesNombre);
    var maquinas = det.columnas.map(function (c) { return c.maquina; });
    ['mesLabel', 'mesLabel2', 'mesLabel3', 'mesLabel4', 'mesLabel5', 'mesLabel6'].forEach(function (id) { el(id).textContent = mesNombre; });

    var maquinasCortas = maquinas.map(function (m) { return m === 'Braner' ? 'Braner' : machineShortLabel(m, STATE.data.machineCodes); });
    renderBulletChart(el('chartMesProd'), {
      categories: maquinasCortas,
      values: det.columnas.map(function (c) { return c.prodEstandar; }),
      targets: det.columnas.map(function () { return null; }),
      formatValue: fmtKgTick, barColorVar: 'var(--series-1)',
      seriesLabel: 'Producción estándar', targetLabel: ''
    });
    renderBulletChart(el('chartMesChatarra'), {
      categories: maquinasCortas,
      values: det.columnas.map(function (c) { return c.chatarraPct; }),
      targets: det.columnas.map(function (c) { return c.metaChatarraEstandar; }),
      formatValue: fmtPctTick, barColorVar: 'var(--series-6)',
      seriesLabel: 'Chatarra', targetLabel: 'Meta'
    });
    renderBulletChart(el('chartMesOEE'), {
      categories: maquinasCortas.filter(function (_, i) { return !det.columnas[i].isBraner; }),
      values: det.columnas.filter(function (c) { return !c.isBraner; }).map(function (c) { return c.oee; }),
      targets: det.columnas.filter(function (c) { return !c.isBraner; }).map(function (c) { return c.oeeMeta; }),
      formatValue: fmtPctTick, barColorVar: 'var(--series-2)',
      seriesLabel: 'OEE', targetLabel: 'Meta'
    });

    var noBraner = det.columnas.filter(function (c) { return !c.isBraner; });
    var noBranerCortas = maquinasCortas.filter(function (_, i) { return !det.columnas[i].isBraner; });
    renderGroupedBarChart(el('chartMesOEEDesglose'), {
      categories: noBranerCortas,
      series: [
        { name: 'Disponibilidad', color: 'var(--series-1)', values: noBraner.map(function (c) { return c.disponibilidad; }) },
        { name: 'Calidad', color: 'var(--series-2)', values: noBraner.map(function (c) { return c.calidad; }) },
        { name: 'Ritmo', color: 'var(--series-3)', values: noBraner.map(function (c) { return c.ritmo; }) }
      ],
      formatValue: function (v) { return fmtPct(v, 0); }
    });

    var aporte = STATE.engine.aporteOEEPorMaquina(mesNombre);
    renderBulletChart(el('chartMesAporteOEE'), {
      categories: aporte.map(function (a) { return machineShortLabel(a.maquina, STATE.data.machineCodes); }),
      values: aporte.map(function (a) { return a.aporte; }),
      targets: aporte.map(function () { return null; }),
      formatValue: function (v) { return fmtPct(v, 1); }, barColorVar: 'var(--series-5)',
      seriesLabel: 'Aporte OEE', targetLabel: ''
    });

    renderMachineTable(el('detalleMesProd'), 'Producción & Chatarra — ' + mesNombre, maquinas, [
      { label: 'Producción Estándar', unit: 'kg', values: det.columnas.map(function (c) { return c.prodEstandar; }) },
      { label: 'Metros Lineales', unit: 'm', values: det.columnas.map(function (c) { return c.metrosLineales; }) },
      { label: 'Factor Kilo/Metro', unit: 'F', values: det.columnas.map(function (c) { return c.factor; }), fmt: function (v) { return fmt1(v); } },
      { label: 'Unidades Fabricadas', unit: 'un', values: det.columnas.map(function (c) { return c.unidades; }) },
      { label: 'Chatarra', unit: 'kg', values: det.columnas.map(function (c) { return c.chatarra; }) },
      { label: 'Chatarra', unit: '%', values: det.columnas.map(function (c) { return c.chatarraPct; }), fmt: function (v) { return fmtPct(v, 2); } },
      { label: 'Meta Chatarra', unit: '%', values: det.columnas.map(function (c) { return c.metaChatarraEstandar; }), fmt: function (v) { return fmtPct(v, 2); } },
      { label: 'Desviación', unit: '%', values: det.columnas.map(function (c) { return c.desviacionChatarraEstandar; }), fmt: function (v) { return fmtPct(v, 2); } }
    ]);

    var valesRows = [
      { label: 'Valor Vales de Consumo', unit: '$', values: det.columnas.map(function (c) { return c.valorVales; }), fmt: fmtMoney },
      { label: 'Costo por Metro Lineal', unit: '$', values: det.columnas.map(function (c) { return c.costoPorMetro; }), fmt: fmtMoney },
      { label: 'Costo Por Kilo', unit: '$', values: det.columnas.map(function (c) { return c.costoPorKilo; }), fmt: fmtMoney2 },
      { label: 'Meta Presupuesto', unit: '$', values: det.columnas.map(function (c) { return c.metaPresupuesto; }), fmt: fmtMoney2 },
      { label: 'Desviación', unit: '$', values: det.columnas.map(function (c) { return c.desviacionCosto; }), fmt: fmtMoney2 }
    ];
    renderMachineTable(el('detalleMesVales'), 'Vales de Consumo — ' + mesNombre, maquinas, valesRows);
    if (det.costoPlanta != null) {
      el('detalleMesVales').innerHTML += '<p class="cap">Costo Por Kilo Planta: ' + fmtMoney2(det.costoPlanta) +
        (det.metaPlanta != null ? ' · Meta: ' + fmtMoney2(det.metaPlanta) + ' · Desviación: ' + fmtMoney2(det.desviacionCostoPlanta) : '') + '</p>';
    }

    renderMachineTable(el('detalleMesOEE'), 'Indicadores OEE — ' + mesNombre, D.machines ? D.machines : maquinas.slice(0, -1), [
      { label: 'Disponibilidad', unit: '%', values: det.columnas.filter(function(c){return !c.isBraner;}).map(function (c) { return c.disponibilidad; }), fmt: function (v) { return fmtPct(v, 1); } },
      { label: 'Calidad', unit: '%', values: det.columnas.filter(function(c){return !c.isBraner;}).map(function (c) { return c.calidad; }), fmt: function (v) { return fmtPct(v, 1); } },
      { label: 'Ritmo', unit: '%', values: det.columnas.filter(function(c){return !c.isBraner;}).map(function (c) { return c.ritmo; }), fmt: function (v) { return fmtPct(v, 1); } },
      { label: 'OEE Máquina', unit: '%', values: det.columnas.filter(function(c){return !c.isBraner;}).map(function (c) { return c.oee; }), fmt: function (v) { return fmtPct(v, 1); }, total: true },
      { label: 'OEE Meta', unit: '%', values: det.columnas.filter(function(c){return !c.isBraner;}).map(function (c) { return c.oeeMeta; }), fmt: function (v) { return fmtPct(v, 1); } },
      { label: 'Desviación OEE', unit: '%', values: det.columnas.filter(function(c){return !c.isBraner;}).map(function (c) { return c.desviacionOEE; }), fmt: function (v) { return fmtPct(v, 1); } },
      { label: 'Horas de Atraso / Adelanto', unit: 'h', values: det.columnas.filter(function(c){return !c.isBraner;}).map(function (c) { return c.horasAtraso; }) }
    ]);
  }

  var D; // convenience alias set on load

  function renderDetalleAnio(anio, maquina) {
    var det = STATE.engine.detalleAnio(anio, maquina);
    ['maqLabel', 'maqLabel2', 'maqLabel3'].forEach(function (id) { el(id).textContent = maquina; });

    renderComboChart(el('chartAnioProd'), {
      categories: det.meses.map(function (m) { return m.mesAbbr; }),
      barName: 'Producción estándar', barColor: 'var(--series-1)',
      barValues: det.meses.map(function (m) { return m.prodEstandar; }),
      metaName: 'Presupuesto', metaValues: det.meses.map(function (m) { return m.pptoProduccion; }),
      desvName: 'Desviación',
      desvValues: det.meses.map(function (m) {
        return (isNum(m.prodEstandar) && isNum(m.pptoProduccion)) ? m.prodEstandar - m.pptoProduccion : null;
      }),
      formatValue: fmtKgTick
    });
    renderComboChart(el('chartAnioOEE'), {
      categories: det.meses.map(function (m) { return m.mesAbbr; }),
      barName: 'OEE real', barColor: 'var(--series-2)',
      barValues: det.meses.map(function (m) { return m.oee; }),
      metaName: 'Meta', metaValues: det.meses.map(function (m) { return m.oeeMeta; }),
      desvName: 'Desviación', desvValues: det.meses.map(function (m) { return m.desviacionOEE; }),
      formatValue: fmtPctTick
    });

    renderMonthlyTable(el('detalleAnioProd'), 'Producción & Chatarra — ' + maquina, [
      { label: 'Producción Estándar', unit: 'kg', values: det.meses.map(function (m) { return m.prodEstandar; }) },
      { label: 'Metros Lineales', unit: 'm', values: det.meses.map(function (m) { return m.metrosLineales; }) },
      { label: 'Factor Kilo/Metro', unit: 'F', values: det.meses.map(function (m) { return m.factor; }), fmt: fmt1 },
      { label: 'Unidades Fabricadas', unit: 'un', values: det.meses.map(function (m) { return m.unidades; }) },
      { label: 'Chatarra', unit: 'kg', values: det.meses.map(function (m) { return m.chatarra; }) },
      { label: 'Chatarra', unit: '%', values: det.meses.map(function (m) { return m.chatarraPct; }), fmt: function (v) { return fmtPct(v, 2); } },
      { label: 'Meta Chatarra', unit: '%', values: det.meses.map(function (m) { return m.metaChatarraEstandar; }), fmt: function (v) { return fmtPct(v, 2); } },
      { label: 'Desviación', unit: '%', values: det.meses.map(function (m) { return m.desviacionChatarraEstandar; }), fmt: function (v) { return fmtPct(v, 2); } }
    ]);
    renderMonthlyTable(el('detalleAnioVales'), 'Vales de Consumo — ' + maquina, [
      { label: 'Valor Vales de Consumo', unit: '$', values: det.meses.map(function (m) { return m.valorVales; }), fmt: fmtMoney },
      { label: 'Costo por Metro Lineal', unit: '$', values: det.meses.map(function (m) { return m.costoPorMetro; }), fmt: fmtMoney },
      { label: 'Costo Por Kilo', unit: '$', values: det.meses.map(function (m) { return m.costoPorKilo; }), fmt: fmtMoney2 },
      { label: 'Meta Presupuesto', unit: '$', values: det.meses.map(function (m) { return m.metaPresupuesto; }), fmt: fmtMoney2 },
      { label: 'Desviación', unit: '$', values: det.meses.map(function (m) { return m.desviacionCosto; }), fmt: fmtMoney2 }
    ]);
    var rows = [
      { label: 'Disponibilidad', unit: '%', values: det.meses.map(function (m) { return m.disponibilidad; }), fmt: function (v) { return fmtPct(v, 1); } },
      { label: 'Calidad', unit: '%', values: det.meses.map(function (m) { return m.calidad; }), fmt: function (v) { return fmtPct(v, 1); } },
      { label: 'Ritmo', unit: '%', values: det.meses.map(function (m) { return m.ritmo; }), fmt: function (v) { return fmtPct(v, 1); } },
      { label: 'OEE Máquina', unit: '%', values: det.meses.map(function (m) { return m.oee; }), fmt: function (v) { return fmtPct(v, 1); }, total: true },
      { label: 'OEE Meta', unit: '%', values: det.meses.map(function (m) { return m.oeeMeta; }), fmt: function (v) { return fmtPct(v, 1); } },
      { label: 'Desviación OEE', unit: '%', values: det.meses.map(function (m) { return m.desviacionOEE; }), fmt: function (v) { return fmtPct(v, 1); } },
      { label: 'Horas de Atraso / Adelanto', unit: 'h', values: det.meses.map(function (m) { return m.horasAtraso; }) }
    ];
    renderMonthlyTable(el('detalleAnioOEE'), 'OEE — ' + maquina + (det.oeePonderado != null ? ' (ponderado ' + fmtPct(det.oeePonderado, 1) + ')' : ''), rows);
  }

  function renderEspesor(anio) {
    var esp = STATE.engine.espesorAnalysis(anio);

    // cap the stacked chart at 8 categorical slots; fold the smallest into "Otros"
    var totals = esp.pctTable.map(function (r) { return { espesor: r.espesor, total: sum(r.meses.filter(isNum)) }; });
    totals.sort(function (a, b) { return b.total - a.total; });
    var keep = totals.slice(0, 8).map(function (t) { return t.espesor; });
    var chartSeries = [];
    esp.pctTable.forEach(function (r, idx) {
      if (keep.indexOf(r.espesor) === -1) return;
      chartSeries.push({ name: r.espesor + ' mm', color: SERIES_COLORS[chartSeries.length % SERIES_COLORS.length], values: r.meses.map(function (v) { return v || 0; }) });
    });
    if (totals.length > 8) {
      var otrosEsp = totals.slice(8).map(function (t) { return t.espesor; });
      var otrosVals = MESES.map(function (_, i) {
        var v = 0;
        esp.pctTable.forEach(function (r) { if (otrosEsp.indexOf(r.espesor) >= 0 && isNum(r.meses[i])) v += r.meses[i]; });
        return v;
      });
      chartSeries.push({ name: 'Otros', color: 'var(--text-muted)', values: otrosVals });
    }
    renderStackedBarChart(el('chartEspesorStack'), { categories: MESES_ABBR, series: chartSeries });

    // a thickness not produced in a month is "-", not a noisy 0
    var pctRows = esp.pctTable.map(function (r) { return { label: r.espesor + ' mm', values: r.meses.map(zeroToNull), fmt: function (v) { return fmtPct(v, 1); } }; });
    pctRows.push({ label: 'Total', values: esp.totalPorMes.map(function (v) { return v ? 1 : null; }), fmt: function (v) { return fmtPct(v, 0); }, total: true });
    renderMonthlyTable(el('espesorPctTable'), 'Producción por Espesor (%) — ' + anio, pctRows);

    var kgRows = esp.kgTable.map(function (r) { return { label: r.espesor + ' mm', values: r.meses.map(zeroToNull) }; });
    kgRows.push({ label: 'Total', values: esp.totalPorMes.map(zeroToNull), total: true });
    renderMonthlyTable(el('espesorKgTable'), 'Producción por Espesor (kg) — ' + anio, kgRows);
  }

  function renderSegmented(containerId, options, activeValue, onPick) {
    var container = el(containerId);
    container.innerHTML = options.map(function (opt) {
      var val = typeof opt === 'object' ? opt.value : opt;
      var label = typeof opt === 'object' ? opt.label : opt;
      var active = String(val) === String(activeValue);
      return '<button type="button" class="seg-btn' + (active ? ' active' : '') + '" data-value="' + escapeHtml(val) + '">' + escapeHtml(label) + '</button>';
    }).join('');
    container.querySelectorAll('.seg-btn').forEach(function (btn) {
      btn.addEventListener('click', function () {
        container.querySelectorAll('.seg-btn').forEach(function (b) { b.classList.remove('active'); });
        btn.classList.add('active');
        onPick(btn.getAttribute('data-value'));
        saveFilters();
      });
    });
  }

  function populateSelectors() {
    // with a single year there's nothing to pick — show it as plain text, not a button
    if (STATE.data.years.length <= 1) {
      el('yearSeg').innerHTML = '<span style="font-size:15px;font-weight:750;letter-spacing:-0.01em;padding:4px 2px;display:inline-block;">' + STATE.year + '</span>';
      el('yearSeg').style.border = 'none';
      el('yearSeg').style.background = 'none';
    } else {
      el('yearSeg').style.border = '';
      el('yearSeg').style.background = '';
      renderSegmented('yearSeg', STATE.data.years, STATE.year, function (val) {
        STATE.year = parseInt(val, 10);
        STATE.month = pickDefaultMonth(STATE.year);
        populateSelectors();
        renderAll();
      });
    }

    renderSegmented('monthSeg', MESES.map(function (m, i) { return { value: m, label: MESES_ABBR[i] }; }), STATE.month, function (val) {
      STATE.month = val;
      renderDetalleMes(STATE.year, STATE.month);
    });

    var machines = STATE.data.machines;
    renderSegmented('machineSeg', machines.map(function (m) { return { value: m, label: machineShortLabel(m, STATE.data.machineCodes) }; }), STATE.machine, function (val) {
      STATE.machine = val;
      renderDetalleAnio(STATE.year, STATE.machine);
    });

    var skuMachineOptions = [{ value: '', label: 'Todas' }].concat(
      machines.map(function (m) { return { value: m, label: machineShortLabel(m, STATE.data.machineCodes) }; })
    );
    // every filter defaults to "todas las máquinas" / "año completo" — the user picks
    // a narrower scope explicitly rather than starting pre-filtered
    renderSegmented('skuMachineSeg', skuMachineOptions, STATE.sku.maquina || '', function (val) {
      STATE.sku.maquina = val || null;
      renderSkuBuscador();
    });
    var skuMonthOptions = [{ value: '', label: 'Todos' }].concat(MESES.map(function (m, i) { return { value: m, label: MESES_ABBR[i] }; }));
    renderSegmented('skuMonthSeg', skuMonthOptions, STATE.sku.mes || '', function (val) {
      STATE.sku.mes = val || null;
      renderSkuBuscador();
    });

    renderSegmented('productosMachineSeg', skuMachineOptions, STATE.productos.maquina || '', function (val) {
      STATE.productos.maquina = val || null;
      renderProductos(STATE.year);
    });
    var productosMonthOptions = [{ value: '', label: 'Anual' }].concat(MESES.map(function (m, i) { return { value: m, label: MESES_ABBR[i] }; }));
    renderSegmented('productosMonthSeg', productosMonthOptions, STATE.productos.mes || '', function (val) {
      STATE.productos.mes = val || null;
      renderProductos(STATE.year);
    });
    var espesorOptions = [{ value: '', label: 'Todos' }].concat(
      STATE.engine.espesoresDisponibles(STATE.year).map(function (e) { return { value: String(e), label: e + ' mm' }; })
    );
    renderSegmented('productosEspesorSeg', espesorOptions, STATE.productos.espesor != null ? String(STATE.productos.espesor) : '', function (val) {
      STATE.productos.espesor = val === '' ? null : parseFloat(val);
      renderProductos(STATE.year);
    });

    renderSegmented('insumosMachineSeg', skuMachineOptions, STATE.insumos.maquina || '', function (val) {
      STATE.insumos.maquina = val || null;
      renderInsumos();
    });
    var insumosMonthOptions = [{ value: '', label: 'Todos' }].concat(MESES.map(function (m, i) { return { value: m, label: MESES_ABBR[i] }; }));
    renderSegmented('insumosMonthSeg', insumosMonthOptions, STATE.insumos.mes || '', function (val) {
      STATE.insumos.mes = val || null;
      renderInsumos();
    });

    renderSegmented('criticosMonthSeg', MESES.map(function (m, i) { return { value: m, label: MESES_ABBR[i] }; }),
      STATE.criticos.mes || STATE.month, function (val) {
        STATE.criticos.mes = val;
        renderCriticos(STATE.year);
      });
    renderSegmented('criticosMachineSeg', skuMachineOptions, STATE.criticos.maquina || '', function (val) {
      STATE.criticos.maquina = val || null;
      renderCriticos(STATE.year);
    });
    var estadoOptions = [
      { value: '', label: 'Todos' }, { value: 'critico', label: 'Críticos' },
      { value: 'alerta', label: 'Alertas' }, { value: 'atencion', label: 'Atención' },
      { value: 'ok', label: 'En meta' }
    ];
    renderSegmented('criticosEstadoSeg', estadoOptions, STATE.criticos.estado || '', function (val) {
      STATE.criticos.estado = val || null;
      renderCriticos(STATE.year);
    });

    renderSegmented('proyeccionesMachineSeg', skuMachineOptions, STATE.proyecciones.maquina || '', function (val) {
      STATE.proyecciones.maquina = val || null;
      renderProyecciones(STATE.year);
    });

    renderSegmented('oeeCalidadMachineSeg', skuMachineOptions, STATE.oeeCalidad.maquina || '', function (val) {
      STATE.oeeCalidad.maquina = val || null;
      renderOeeCalidad(STATE.year);
    });
    var oeeCalidadMonthOptions = [{ value: '', label: 'Todos' }].concat(MESES.map(function (m, i) { return { value: m, label: MESES_ABBR[i] }; }));
    renderSegmented('oeeCalidadMonthSeg', oeeCalidadMonthOptions, STATE.oeeCalidad.mes || '', function (val) {
      STATE.oeeCalidad.mes = val || null;
      renderOeeCalidad(STATE.year);
    });
  }

  function pickDefaultMonth(anio) {
    var lastWithData = null;
    STATE.data.produccion.forEach(function (p) {
      if (p.anio === anio && isNum(p.kgReal) && p.kgReal !== 0) {
        if (lastWithData == null || p.mes > lastWithData) lastWithData = p.mes;
      }
    });
    return MESES[(lastWithData || 1) - 1];
  }

  var PILL_LABEL = { ok: 'En meta', alerta: 'Fuera de meta', sindato: 'Sin dato' };
  function renderTorreControl(anio, mesNombre) {
    el('torreMesLabel').textContent = mesNombre;
    var rows = STATE.engine.torreControl(anio, mesNombre);
    function pill(estado) {
      var cls = estado === 'ok' ? 'good' : (estado === 'alerta' ? 'bad' : '');
      return '<span class="pill' + (cls ? ' ' + cls : ' pill-muted') + '">' + PILL_LABEL[estado] + '</span>';
    }
    // desviaciones firmadas de modo que POSITIVO = a favor de la meta en todos los
    // indicadores (chatarra/costo: meta − real; OEE: real − meta), y rojo cuando negativo
    function desvTd(real, meta, badWhenHigher, fmt) {
      if (real == null || meta == null) return '<td class="dash">-</td>';
      var d = badWhenHigher ? meta - real : real - meta;
      return '<td class="' + (d < 0 ? 'neg' : 'pos') + '">' + fmtSigned(d, fmt) + '</td>';
    }
    var pct2 = function (v) { return fmtPct(v, 2); };
    var pct1 = function (v) { return fmtPct(v, 1); };
    // soft dashed rule between the Chatarra / OEE / Costo-por-kg column groups
    var html = '<table class="wide"><thead><tr><th>Máquina</th>' +
      '<th>Chatarra</th><th>Meta Admin.</th><th>Desv. Admin.</th><th>Meta Calc.</th><th>Desv. Calc.</th><th>Estado</th>' +
      '<th class="group-sep">OEE</th><th>Meta</th><th>Desv.</th><th>Estado</th>' +
      '<th class="group-sep">Costo/kg</th><th>Meta</th><th>Desv.</th><th>Estado</th></tr></thead><tbody>';
    rows.forEach(function (r) {
      html += '<tr>' +
        '<td>' + escapeHtml(machineShortLabel(r.maquina, STATE.data.machineCodes)) + ' <span class="subtle">' + escapeHtml(r.maquina) + '</span></td>' +
        tdv(r.chatarraPct, pct2) + tdv(r.metaAdministrativa, pct2) +
        desvTd(r.chatarraPct, r.metaAdministrativa, true, pct2) +
        tdv(r.metaCalculada, pct2) +
        desvTd(r.chatarraPct, r.metaCalculada, true, pct2) +
        '<td>' + pill(r.chatarraEstado) + '</td>' +
        tdv(r.oee, pct1, 'group-sep') + tdv(r.oeeMeta, pct1) +
        desvTd(r.oee, r.oeeMeta, false, pct1) +
        '<td>' + pill(r.oeeEstado) + '</td>' +
        tdv(r.costoPorKilo, fmtMoney2, 'group-sep') + tdv(r.metaPresupuesto, fmtMoney2) +
        desvTd(r.costoPorKilo, r.metaPresupuesto, true, fmtMoney2) +
        '<td>' + pill(r.costoEstado) + '</td>' +
        '</tr>';
    });
    html += '</tbody></table>';
    el('torreControl').innerHTML = html;
  }

  var CRITICOS_FMT = {
    kg: fmtInt,
    pct1: function (v) { return fmtPct(v, 1); },
    pct2: function (v) { return fmtPct(v, 2); },
    money: fmtMoney,
    moneypkg: fmtMoney2
  };
  var CRITICOS_PILL = {
    critico: '<span class="pill bad">Crítico</span>',
    alerta: '<span class="pill bad" style="opacity:.75;">Alerta</span>',
    atencion: '<span class="pill" style="background:color-mix(in srgb, var(--warning) 22%, transparent);color:var(--serious);">Atención</span>',
    ok: '<span class="pill good">En meta</span>'
  };
  var CRITICOS_GRUPOS = ['OEE', 'Producción', 'Chatarra', 'Costos'];
  function renderCriticos(anio) {
    var mes = STATE.criticos.mes || STATE.month;
    var pc = STATE.engine.puntosCriticos(anio, mes, { maquina: STATE.criticos.maquina });
    el('criticosMesLabel').textContent = mes + (pc.enCurso ? ' (en curso, día ' + pc.diaActual + ' de ' + pc.diasMes + ' — presupuestos de kg y $ prorrateados)' : '');

    // chips count the machine-filtered scope; the estado filter narrows the tables below
    var chips = [
      { label: 'Críticos (≥20%)', value: pc.resumen.critico, accent: 'var(--critical)' },
      { label: 'Alertas (8-20%)', value: pc.resumen.alerta, accent: 'var(--serious)' },
      { label: 'Atención (<8%)', value: pc.resumen.atencion, accent: 'var(--warning)' },
      { label: 'En meta', value: pc.resumen.ok, accent: 'var(--good)' }
    ];
    el('criticosChips').innerHTML = chips.map(function (c) {
      return '<div class="tile" style="--accent:' + c.accent + ';">' +
        '<div class="label">' + c.label + '</div>' +
        '<div class="value">' + c.value + '</div></div>';
    }).join('');

    var estadoFiltro = STATE.criticos.estado || null;
    var visibles = pc.items.filter(function (it) { return !estadoFiltro || it.estado === estadoFiltro; });
    // one card per exact indicator (its "tipo" becomes the card title, e.g. "OEE vs
    // Meta", "Costo Aceite y Lubricante"), grouped in the OEE/Producción/Chatarra/
    // Costos order and no longer needing a redundant "Indicador" column
    var html = '';
    CRITICOS_GRUPOS.forEach(function (grupo) {
      var tipos = [];
      visibles.forEach(function (it) { if (it.grupo === grupo && tipos.indexOf(it.tipo) < 0) tipos.push(it.tipo); });
      tipos.forEach(function (tipo) {
        var items = visibles.filter(function (it) { return it.tipo === tipo; });
        if (!items.length) return;
        html += '<div class="card"><h3>' + escapeHtml(tipo) + '</h3>' +
          '<table class="wide"><thead><tr><th>Máquina</th><th>Real</th><th>Meta / Ppto</th><th>Desviación</th><th>Estado</th></tr></thead><tbody>';
        items.forEach(function (it) {
          var fmt = CRITICOS_FMT[it.fmtKind] || fmtInt;
          html += '<tr>' +
            '<td>' + (it.maquina === 'Planta' ? 'Planta' : escapeHtml(machineShortLabel(it.maquina, STATE.data.machineCodes)) + ' <span class="subtle">' + escapeHtml(it.maquina) + '</span>') + '</td>' +
            tdv(it.real, fmt) + tdv(it.meta, fmt) +
            '<td class="' + (it.gravedad > 0 ? 'neg' : 'pos') + '">' + fmtSigned(it.desvRel, function (v) { return fmtPct(v, 1); }) + '</td>' +
            '<td>' + CRITICOS_PILL[it.estado] + '</td></tr>';
        });
        html += '</tbody></table></div>';
      });
    });
    el('criticosSecciones').innerHTML = html ||
      '<div class="card"><div class="cap">Sin indicadores para este filtro.</div></div>';
  }

  function renderProyeccion(anio) {
    var p = STATE.engine.proyeccionCierre(anio);
    var section = el('proyeccionSection');
    if (!p) { section.style.display = 'none'; return; }
    section.style.display = 'block';
    el('proyeccionMesLabel').textContent = p.mes;
    el('proyeccionCap').textContent = p.cerrado ?
      'Mes completo (' + p.diasMes + ' días) — la proyección coincide con el cierre real.' :
      'Avance al día ' + p.diaActual + ' de ' + p.diasMes + ' — proyección al ritmo actual (kg a la fecha ÷ días transcurridos × días del mes) vs. presupuesto.';
    function pill(estado) {
      var cls = estado === 'ok' ? 'good' : (estado === 'alerta' ? 'bad' : '');
      var label = estado === 'ok' ? 'Cumple' : (estado === 'alerta' ? 'En riesgo' : 'Sin ppto');
      return '<span class="pill' + (cls ? ' ' + cls : ' pill-muted') + '">' + label + '</span>';
    }
    var html = '<table class="wide"><thead><tr><th>Máquina</th><th>Real al día ' + p.diaActual + ' (kg)</th><th>Proyección cierre (kg)</th><th>Presupuesto (kg)</th><th>Desv. proyectada</th><th>Estado</th></tr></thead><tbody>';
    p.maquinas.forEach(function (m) {
      html += '<tr>' +
        '<td>' + escapeHtml(machineShortLabel(m.maquina, STATE.data.machineCodes)) + ' <span class="subtle">' + escapeHtml(m.maquina) + '</span></td>' +
        tdv(m.kg, fmtInt) + tdv(m.proy, fmtInt) + tdv(m.ppto, fmtInt) +
        tdv(m.desvPct, function (v) { return fmtSigned(v, function (x) { return fmtPct(x, 1); }); }) +
        '<td>' + pill(m.estado) + '</td></tr>';
    });
    html += '<tr class="total"><td>Total Planta</td>' +
      tdv(p.total.kg, fmtInt) + tdv(p.total.proy, fmtInt) + tdv(p.total.ppto, fmtInt) +
      tdv(p.total.desvPct, function (v) { return fmtSigned(v, function (x) { return fmtPct(x, 1); }); }) +
      '<td>' + pill(p.total.estado) + '</td></tr>';
    html += '</tbody></table>';
    el('proyeccionTable').innerHTML = html;
  }

  function proyChip(label, value, sub, accent) {
    return '<div class="tile" style="--accent:' + accent + ';">' +
      '<div class="label">' + label + '</div>' +
      '<div class="value">' + value + '</div>' +
      (sub ? '<div class="delta ' + sub.cls + '">' + sub.text + '</div>' : '') + '</div>';
  }
  function proyDeltaSub(desvPct, favorableLabel, contraLabel) {
    if (desvPct == null) return null;
    var good = desvPct >= 0;
    return { cls: good ? 'good' : 'bad', text: fmtSigned(desvPct, function (v) { return fmtPct(v, 1); }) + ' ' + (good ? favorableLabel : contraLabel) };
  }

  function renderProyecciones(anio) {
    var maquina = STATE.proyecciones.maquina;
    var maquinaTxt = maquina ? machineShortLabel(maquina, STATE.data.machineCodes) : 'toda la planta';
    var pa = STATE.engine.proyeccionAnalitica(anio, { maquina: maquina });
    var empty = el('proyeccionesEmpty'), content = el('proyeccionesContent');
    if (!pa.mes) { empty.style.display = 'block'; content.style.display = 'none'; return; }
    empty.style.display = 'none'; content.style.display = 'block';

    // meta/presupuesto mensual (12 valores), calculado una sola vez y reutilizado en
    // los 4 gráficos y en las tarjetas de cierre de mes
    var detMaquina = maquina ? STATE.engine.detalleAnio(anio, maquina) : null;
    var pptoProdSerie = MESES.map(function (m, i) {
      return detMaquina ? detMaquina.meses[i].pptoProduccion : STATE.engine.pptoTotalVal(m);
    });
    var costoMetaSerie = MESES.map(function (m, i) {
      return detMaquina ? detMaquina.meses[i].metaPresupuesto : null;
    });
    var oeeMetaSerie = MESES.map(function (m, i) {
      return detMaquina ? detMaquina.meses[i].oeeMeta : STATE.engine.oeeMetaVal('OEE Planta', m);
    });

    el('proyMesLabel').textContent = pa.mes + ' — ' + maquinaTxt;
    el('proyMesCap').textContent = pa.enCurso ?
      ('Avance al día ' + pa.diaActual + ' de ' + pa.diasMes + ': se proyecta el cierre del mes prorrateando el ritmo actual (valor a la fecha ÷ días transcurridos × días del mes).') :
      (pa.mes + ' ya cerró — el cierre de mes es el dato real, sin proyección.');

    var mesIdx = MESES.indexOf(pa.mes);
    var mesProd = pa.produccion.serie[mesIdx], mesChat = pa.chatarraPct.serie[mesIdx];
    var mesCosto = pa.costoKg.serie[mesIdx], mesOee = pa.oee.serie[mesIdx];
    var prodCierreVal = mesProd ? (mesProd.proyectado != null ? mesProd.proyectado : mesProd.actual) : null;
    var pptoMesKg = pptoProdSerie[mesIdx];
    var desvMesProd = (prodCierreVal != null && pptoMesKg) ? (prodCierreVal - pptoMesKg) / pptoMesKg : null;
    var chatCierreVal = mesChat ? mesChat.actual : null; // % del mes: se usa el valor real observado tal cual
    var desvMesChat = (chatCierreVal != null && pa.anual.chatMeta != null) ? pa.anual.chatMeta - chatCierreVal : null;
    var costoCierreVal = mesCosto ? mesCosto.actual : null;
    var costoMetaMes = costoMetaSerie[mesIdx];
    var desvMesCosto = (costoCierreVal != null && costoMetaMes) ? (costoMetaMes - costoCierreVal) / costoMetaMes : null;
    var oeeCierreVal = mesOee ? mesOee.actual : null;
    var oeeMetaMes = oeeMetaSerie[mesIdx];
    var desvMesOee = (oeeCierreVal != null && oeeMetaMes != null) ? oeeCierreVal - oeeMetaMes : null;

    el('proyMesChips').innerHTML = [
      proyChip('Producción proyectada', fmtKgTick(prodCierreVal) + ' kg', proyDeltaSub(desvMesProd, 'sobre ppto', 'bajo ppto'), 'var(--series-1)'),
      proyChip('Chatarra', chatCierreVal != null ? fmtPct(chatCierreVal, 2) : '-', proyDeltaSub(desvMesChat, 'bajo meta', 'sobre meta'), 'var(--series-6)'),
      proyChip('Costo/kg', costoCierreVal != null ? fmtMoney2(costoCierreVal) : '-', proyDeltaSub(desvMesCosto, 'bajo ppto', 'sobre ppto'), 'var(--series-3)'),
      proyChip('OEE', oeeCierreVal != null ? fmtPct(oeeCierreVal, 1) : '-', proyDeltaSub(desvMesOee, 'sobre meta', 'bajo meta'), 'var(--series-2)')
    ].join('');

    var a = pa.anual;
    el('proyAnioChips').innerHTML = [
      proyChip('Producción proyectada al año', fmtKgTick(a.prodKg) + ' kg', proyDeltaSub(a.desvProdPct, 'sobre ppto', 'bajo ppto'), 'var(--series-1)'),
      proyChip('Chatarra proyectada', a.chatPct != null ? fmtPct(a.chatPct, 2) : '-', proyDeltaSub(a.desvChatPct, 'bajo meta', 'sobre meta'), 'var(--series-6)'),
      proyChip('Costo/kg proyectado', a.costoKg != null ? fmtMoney2(a.costoKg) : '-', proyDeltaSub(a.desvCostoPct, 'bajo ppto', 'sobre ppto'), 'var(--series-3)'),
      proyChip('OEE proyectado', a.oee != null ? fmtPct(a.oee, 1) : '-', proyDeltaSub(a.desvOeePct, 'sobre meta', 'bajo meta'), 'var(--series-2)')
    ].join('');

    renderForecastChart(el('chartProyProduccion'), {
      categories: MESES_ABBR,
      actual: pa.produccion.serie.map(function (s) { return s.actual; }),
      proyectado: pa.produccion.serie.map(function (s) { return s.proyectado; }),
      meta: pptoProdSerie, formatValue: fmtKgTick, barColor: 'var(--series-1)'
    });
    renderForecastChart(el('chartProyChatarra'), {
      categories: MESES_ABBR,
      actual: pa.chatarraPct.serie.map(function (s) { return s.actual; }),
      proyectado: pa.chatarraPct.serie.map(function (s) { return s.proyectado; }),
      meta: MESES.map(function () { return a.chatMeta; }),
      formatValue: fmtPctTick, barColor: 'var(--series-6)'
    });
    renderForecastChart(el('chartProyCosto'), {
      categories: MESES_ABBR,
      actual: pa.costoKg.serie.map(function (s) { return s.actual; }),
      proyectado: pa.costoKg.serie.map(function (s) { return s.proyectado; }),
      meta: costoMetaSerie, formatValue: fmtMoneyTick2, barColor: 'var(--series-3)'
    });
    renderForecastChart(el('chartProyOEE'), {
      categories: MESES_ABBR,
      actual: pa.oee.serie.map(function (s) { return s.actual; }),
      proyectado: pa.oee.serie.map(function (s) { return s.proyectado; }),
      meta: oeeMetaSerie, formatValue: fmtPctTick, barColor: 'var(--series-2)'
    });
  }

  function renderInsumos() {
    var data = STATE.data.pptoInsumos || [];
    if (!data.length) {
      el('insumosEmpty').style.display = 'block';
      el('insumosContent').style.display = 'none';
      return;
    }
    el('insumosEmpty').style.display = 'none';
    el('insumosContent').style.display = 'block';

    var maquina = STATE.insumos.maquina;
    var maquinaTxt = maquina ? machineShortLabel(maquina, STATE.data.machineCodes) : 'todas las máquinas';
    var mesTxt = STATE.insumos.mes ? STATE.insumos.mes : 'año completo';
    var filtroTxt = maquinaTxt + ', ' + mesTxt;

    // "Mayor gasto por insumo" and the table below it are sourced from the vale
    // transactions (STATE.engine.insumosPxQ), which carry a real date — that's what
    // lets both the Máquina AND Mes filters apply here. "Ppto Insumos" (used only for
    // the sobrecosto/IPC chart) is a period-aggregate reference with no date column at
    // all, so that one chart stays annual regardless of the Mes filter — noted in its caption.
    var pxq = STATE.engine.insumosPxQ(STATE.year, { maquina: maquina, mes: STATE.insumos.mes });

    el('insumosGastoCap').textContent = 'gasto real (vales de consumo) — ' + filtroTxt;
    var rk = insumosRanking(maquina ? data.filter(function (r) { return r.ccosto === maquina; }) : data);
    el('insumosSobrecostoCap').textContent = 'último precio de compra vs. precio ajustado por IPC — ' + maquinaTxt + ' (referencia anual, no varía por mes)';

    renderRankedBarChart(el('chartInsumosGasto'), {
      items: pxq.slice(0, 12).map(function (r) { return { label: r.articulo, value: r.gasto }; }),
      formatValue: fmtMoney, colorVar: 'var(--series-1)'
    });
    renderRankedBarChart(el('chartInsumosSobrecosto'), {
      items: rk.topSobrecosto.map(function (r) { return { label: r.articulo, value: r.overrun }; }),
      formatValue: function (v) { return fmtPct(v, 0); }, colorVar: 'var(--series-6)'
    });

    // Cada categoría agrupa: su tendencia anual real vs. presupuesto (comparación real
    // contra presupuesto, a nivel de categoría — el único nivel al que el Excel trae
    // presupuesto de verdad) y el movimiento de precio/cantidad por ítem individual
    // (el Excel no trae presupuesto por ítem, así que ahí se usa el precio IPC como
    // referencia de inflación, dejado explícito en el texto).
    renderCostosCategoria('Aceite y Lubricante', 'chartCostosAceite', 'costosAceiteCap', maquina, maquinaTxt);
    renderInsumoMovimiento('Aceite y Lubricante', 'pxqTableAceite', 'pxqAceiteCap', maquina, STATE.insumos.mes, filtroTxt);

    renderCostosCategoria('Insumos de Fábrica', 'chartCostosInsumosFab', 'costosInsumosFabCap', maquina, maquinaTxt);
    renderInsumoMovimiento('Insumos de Fábrica', 'pxqTableInsumosFab', 'pxqInsumosFabCap', maquina, STATE.insumos.mes, filtroTxt);

    renderCostosCategoria('Embalajes', 'chartCostosEmbalajes', 'costosEmbalajesCap', maquina, maquinaTxt);
    renderInsumoMovimiento('Embalajes', 'pxqTableEmbalajes', 'pxqEmbalajesCap', maquina, STATE.insumos.mes, filtroTxt);
  }

  // Un ítem por fila: compara su "mes base" (primera compra del año) contra el mes
  // filtrado (o el más reciente con datos, si el filtro es "Todos") — precio y
  // cantidad, cada uno con su propia flecha de variación. El precio también se
  // compara contra el precio IPC como referencia (no hay presupuesto real por ítem).
  function renderInsumoMovimiento(categoria, tableId, capId, maquina, mes, filtroTxt) {
    var items = STATE.engine.insumosMovimiento(STATE.year, { maquina: maquina, mes: mes, categoria: categoria }).slice(0, 15);
    el(capId).textContent = items.length ?
      ('cada ítem compara su primer mes de compra del año contra ' + (mes || 'el mes más reciente con datos') +
       ' — precio IPC es una referencia de inflación, no presupuesto real del ítem — ' + filtroTxt) :
      ('sin vales de consumo de esta categoría para este filtro (' + filtroTxt + ')');

    function deltaCell(pct, esPrimeraCompra, sinDatos, judged) {
      if (esPrimeraCompra) return '<td class="dash">primera compra</td>';
      if (sinDatos) return '<td class="dash">sin compras</td>';
      if (pct == null) return '<td class="dash">-</td>';
      var cls = judged ? (pct < 0 ? 'good' : (pct > 0 ? 'bad' : 'neutral')) : 'neutral';
      var arrow = pct < 0 ? '▼' : (pct > 0 ? '▲' : '—');
      return '<td class="delta ' + cls + '">' + arrow + ' ' + fmtPct(Math.abs(pct), 1) + '</td>';
    }

    var html = '<table class="wide"><thead><tr>' +
      '<th>Insumo</th><th>Mes base</th><th>Cantidad</th><th>Δ Cantidad</th>' +
      '<th>Precio</th><th>Δ Precio</th><th>Precio IPC (ref.)</th><th>Desv. vs IPC</th><th>Gasto período</th>' +
      '</tr></thead><tbody>';
    items.forEach(function (r) {
      html += '<tr><td>' + escapeHtml(r.articulo) + '</td>' +
        '<td class="dash">' + r.mesBase + '</td>' +
        tdv(r.cantidadObjetivo, fmtInt) +
        deltaCell(r.deltaCantidad, r.esPrimeraCompra, r.sinComprasMesObjetivo, false) +
        tdv(r.precioObjetivo, fmtMoney) +
        deltaCell(r.deltaPrecio, r.esPrimeraCompra, r.sinComprasMesObjetivo, true) +
        tdv(r.precioIPC, fmtMoney) +
        // negativo = pagando MENOS que la referencia IPC (bueno) — reutiliza deltaCell
        // en vez de tdv(), que colorea cualquier negativo en rojo sin importar el sentido.
        deltaCell(r.vsIpc, false, false, true) +
        tdv(r.gasto, fmtMoney) +
        '</tr>';
    });
    html += '</tbody></table>';
    el(tableId).innerHTML = items.length ? html : '';
  }

  // "Vale de consumo" real vs "Ppto Vale de Consumo" budget, monthly, for one cost
  // category (Aceite y Lubricante / Insumos de Fábrica / Embalajes), plus a caption
  // comparing the annual monthly average to the current month.
  function renderCostosCategoria(categoria, chartId, capId, maquina, maquinaTxt) {
    var data = STATE.engine.costosCategoria(STATE.year, categoria, { maquina: maquina });
    renderComboChart(el(chartId), {
      categories: MESES_ABBR,
      barName: 'Real', barColor: 'var(--series-1)',
      barValues: data.months.map(function (m) { return m.real; }),
      metaName: 'Presupuesto', metaValues: data.months.map(function (m) { return m.ppto; }),
      desvName: 'Desviación (ppto − real)',
      desvValues: data.months.map(function (m) { return (m.real != null && m.ppto != null) ? m.ppto - m.real : null; }),
      formatValue: fmtMoneyTick
    });
    var mesRef = STATE.insumos.mes || STATE.month;
    var mesActualRow = data.months.filter(function (m) { return m.mes === mesRef; })[0];
    var parts = [maquinaTxt];
    if (data.promedioAnual != null) parts.push('promedio anual: ' + fmtMoney(data.promedioAnual));
    if (mesActualRow && mesActualRow.real != null) parts.push(mesRef + ': ' + fmtMoney(mesActualRow.real));
    if (mesActualRow && mesActualRow.desviacionPct != null) {
      parts.push('desv. vs ppto: ' + fmtSigned(mesActualRow.desviacionPct, function (v) { return fmtPct(v, 0); }));
    }
    el(capId).textContent = parts.join(' — ');
  }

  function renderProductos(anio) {
    el('productosAnioLabel').textContent = anio;
    var opts = { maquina: STATE.productos.maquina, mes: STATE.productos.mes, espesor: STATE.productos.espesor };
    var mix = STATE.engine.productMix(anio, opts);
    var periodo = opts.mes ? opts.mes : 'año completo';
    var maquinaTxt = opts.maquina ? machineShortLabel(opts.maquina, STATE.data.machineCodes) : 'todas las máquinas';
    var espesorTxt = opts.espesor != null ? opts.espesor + ' mm' : 'todos los espesores';
    var filtroTxt = maquinaTxt + ', ' + periodo + ', ' + espesorTxt;
    el('productosTopCap').textContent = 'kg producidos — ' + filtroTxt;
    el('productosDonutCap').textContent = 'participación del kg producido — ' + filtroTxt;
    el('productosMaqMesCap').textContent = 'kg por mes — ' + filtroTxt;
    el('productosEspMesCap').textContent = 'kg por espesor y mes — ' + filtroTxt;
    el('productosEspAnualCap').textContent = 'kg totales por espesor — ' + filtroTxt;

    renderRankedBarChart(el('chartTopProductos'), {
      items: mix.topProductos.map(function (r) { return { label: r.label, value: r.kg }; }),
      formatValue: fmtKgTick, colorVar: 'var(--series-1)'
    });

    renderDonutChart(el('chartFamiliaDonut'), {
      items: mix.porFamilia.map(function (r) { return { label: r.label, value: r.kg }; }),
      formatValue: function (v) { return fmtInt(v) + ' kg'; }
    });

    var porMaquina = STATE.engine.produccionMensualPorMaquina(anio, opts);
    renderGroupedBarChart(el('chartProdMaquinaMes'), {
      categories: MESES_ABBR,
      series: porMaquina.map(function (m, i) {
        return { name: machineShortLabel(m.maquina, STATE.data.machineCodes), color: SERIES_COLORS[i % SERIES_COLORS.length], values: m.values };
      }),
      formatValue: fmtKgTick
    });

    // producción por espesor — cap to the 8 biggest thicknesses + "Otros" so the monthly
    // grouped chart doesn't get overcrowded with dozens of thin categorical slots
    var espKg = STATE.engine.espesorKgFiltrado(anio, opts);
    var espMesSeries = espKg.slice(0, 8).map(function (r, i) {
      return { name: r.espesor + ' mm', color: SERIES_COLORS[i % SERIES_COLORS.length], values: r.values };
    });
    if (espKg.length > 8) {
      var otrosVals = MESES.map(function (_, i) {
        var v = 0;
        espKg.slice(8).forEach(function (r) { if (isNum(r.values[i])) v += r.values[i]; });
        return v;
      });
      espMesSeries.push({ name: 'Otros', color: 'var(--text-muted)', values: otrosVals });
    }
    renderGroupedBarChart(el('chartEspesorMesKg'), { categories: MESES_ABBR, series: espMesSeries, formatValue: fmtKgTick });

    renderRankedBarChart(el('chartEspesorAnualKg'), {
      items: espKg.map(function (r) { return { label: r.espesor + ' mm', value: r.total }; }),
      formatValue: fmtKgTick, colorVar: 'var(--series-3)'
    });

    el('productosEspDonutCap').textContent = 'participación del kg producido por espesor — ' + filtroTxt;
    renderDonutChart(el('chartEspesorDonut'), {
      items: espKg.map(function (r) { return { label: r.espesor + ' mm', value: r.total }; }),
      formatValue: function (v) { return fmtInt(v) + ' kg'; }
    });
  }

  var SKU_ROW_LIMIT = 300;
  function renderSkuBuscador() {
    var art = STATE.engine.skuArticleMonthlyTrend(STATE.year, { maquina: STATE.sku.maquina, mes: STATE.sku.mes, query: STATE.sku.query });
    // when a single month is selected there is only one populated x-position — collapse the
    // chart to that one category so bubbles spread across the full plot width instead of
    // cramming into one narrow 1/12th-wide band (which caused labels/bubbles to overlap).
    var mesIdx = STATE.sku.mes ? MESES.indexOf(STATE.sku.mes) : -1;
    var bubbleInfo = renderBubbleChart(el('chartSkuBubble'), {
      categories: mesIdx >= 0 ? [MESES_ABBR[mesIdx]] : MESES_ABBR,
      series: art.series.map(function (s, i) {
        var months = mesIdx >= 0 ? [s.months[mesIdx]] : s.months;
        return {
          label: s.label, color: SERIES_COLORS[i % SERIES_COLORS.length],
          values: months.map(function (m) { return m.chatPct; }),
          sizes: months.map(function (m) { return m.chatKg; })
        };
      }),
      formatValue: function (v) { return fmtPct(v, 1); },
      formatSize: function (v) { return fmtInt(v) + ' kg'; }
    });
    var skuFilterTxt = (STATE.sku.mes ? STATE.sku.mes : 'año completo') + ', ' + (STATE.sku.maquina ? machineShortLabel(STATE.sku.maquina, STATE.data.machineCodes) : 'todas las máquinas');
    var hiddenTxt = bubbleInfo && bubbleInfo.hiddenLabels ? ' (' + fmtInt(bubbleInfo.hiddenLabels) + ' sin etiqueta visible por espacio — pasa el mouse sobre la burbuja)' : '';
    el('chartSkuBubble').parentNode.querySelector('.cap').textContent = art.series.length === 0 ? 'Sin artículos con chatarra para este filtro (' + skuFilterTxt + ')' :
      (fmtInt(art.totalArticles) + ' artículo' + (art.totalArticles === 1 ? '' : 's') + ' con chatarra — ' + skuFilterTxt + hiddenTxt);

    var pareto = STATE.engine.paretoChatarra(STATE.year, { maquina: STATE.sku.maquina, mes: STATE.sku.mes });
    if (pareto.items.length) {
      renderParetoChart(el('chartPareto'), pareto);
      el('paretoCap').textContent = 'Los primeros ' + pareto.n80 + ' de ' + pareto.items.length +
        ' artículos concentran el 80% de la chatarra (' + fmtInt(pareto.total) + ' kg totales) — ' + skuFilterTxt;
    } else {
      el('chartPareto').innerHTML = '';
      el('paretoCap').textContent = 'Sin chatarra atribuible a artículos para este filtro (' + skuFilterTxt + ')';
    }

    var paretoP = STATE.engine.paretoProduccion(STATE.year, { maquina: STATE.sku.maquina, mes: STATE.sku.mes });
    if (paretoP.items.length) {
      renderParetoChart(el('chartParetoProd'), paretoP);
      el('paretoProdCap').textContent = 'Los primeros ' + paretoP.n80 + ' de ' + paretoP.items.length +
        ' artículos concentran el 80% de la producción (' + fmtInt(paretoP.total) + ' kg totales) — ' + skuFilterTxt;
    } else {
      el('chartParetoProd').innerHTML = '';
      el('paretoProdCap').textContent = 'Sin producción para este filtro (' + skuFilterTxt + ')';
    }

    var rows = STATE.engine.skuBuscador(STATE.year, STATE.sku);
    var total = rows.length;
    var shown = rows.slice(0, SKU_ROW_LIMIT);
    el('skuResultCount').textContent = total === 0 ? 'Sin resultados' :
      (total > SKU_ROW_LIMIT ? 'Mostrando ' + SKU_ROW_LIMIT + ' de ' + fmtInt(total) + ' resultados — refina la búsqueda para ver más' : fmtInt(total) + ' resultado' + (total === 1 ? '' : 's'));

    var html = '<table class="wide"><thead><tr><th>Descripción</th><th>Máquina</th><th>Mes</th><th>Producción (kg)</th><th>Chatarra (kg)</th><th>Chatarra %</th></tr></thead><tbody>';
    shown.forEach(function (r) {
      html += '<tr>' +
        '<td>' + escapeHtml(r.descripcion || '-') + '</td>' +
        '<td>' + escapeHtml(machineShortLabel(r.maquina, STATE.data.machineCodes)) + '</td>' +
        '<td>' + escapeHtml(r.mesLabel) + '</td>' +
        tdv(r.prodKg, fmtInt) + tdv(r.chatKg, fmtInt) +
        tdv(r.chatPct, function (v) { return fmtPct(v, 2); }) +
        '</tr>';
    });
    html += '</tbody></table>';
    el('skuTable').innerHTML = html;
  }

  function renderAll() {
    var anio = STATE.year;
    var resumen = renderResumen(anio);
    var kpi = STATE.engine.kpiHeader(anio, resumen);
    renderKPIs(kpi);
    renderTorreControl(anio, STATE.month);
    renderProyeccion(anio);
    renderCriticos(anio);
    renderTendencias(anio, resumen);
    renderPresupuesto(anio, resumen);
    renderDetalleMes(anio, STATE.month);
    renderDetalleAnio(anio, STATE.machine);
    renderEspesor(anio);
    renderInsumos();
    renderProductos(anio);
    renderSkuBuscador();
    renderProyecciones(anio);
    renderOeeCalidad(anio);
    el('updatedLabel').textContent = 'Developed by Gino Espinosa Morales';
    saveFilters();
  }

  var TAB_META = {
    resumen: { title: 'Resumen Ejecutivo', crumb: 'Planta · Año completo' },
    criticos: { title: 'Puntos Críticos', crumb: 'Todas las desviaciones vs. meta: OEE, producción, chatarra y costos' },
    maquina: { title: 'Detalle por Máquina', crumb: 'Producción, chatarra, costos y OEE por máquina' },
    espesor: { title: 'Análisis por Espesor', crumb: 'Mix de producción por espesor' },
    insumos: { title: 'Costos e Insumos', crumb: 'Gasto y variación de precio por insumo' },
    productos: { title: 'Mix de Productos', crumb: 'Producción y chatarra por producto y familia' },
    sku: { title: 'Buscador SKU', crumb: 'Producción y chatarra por SKU, máquina y mes' },
    proyecciones: { title: 'Proyecciones', crumb: 'Cierre de mes y de año proyectado — producción, chatarra, costos y OEE' },
    oeeCalidad: { title: 'OEE y Calidad Planta', crumb: 'Tendencia mensual de OEE y Calidad Planta vs. meta, por máquina y mes' }
  };

  function switchTab(tab) {
    document.querySelectorAll('.nav-item').forEach(function (btn) {
      btn.classList.toggle('active', btn.getAttribute('data-tab') === tab);
    });
    document.querySelectorAll('.tab-panel').forEach(function (panel) {
      panel.classList.toggle('active', panel.id === 'tab-' + tab);
    });
    el('pageTitle').textContent = TAB_META[tab].title;
    el('pageCrumb').textContent = TAB_META[tab].crumb;
    el('sidebar').classList.remove('open');
    if (STATE.data) saveFilters();
  }

  /* ---- Persistencia local: último Excel (IndexedDB) + filtros (localStorage).
   * Todo queda en el navegador del usuario — nada sale del equipo. ---- */
  var IDB_NAME = 'erp-analytics', IDB_STORE = 'files', LS_FILTERS = 'erpAnalyticsFiltros';
  function idbOpen() {
    return new Promise(function (resolve, reject) {
      var req = indexedDB.open(IDB_NAME, 1);
      req.onupgradeneeded = function () { req.result.createObjectStore(IDB_STORE); };
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { reject(req.error); };
    });
  }
  function idbPut(key, value) {
    return idbOpen().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(IDB_STORE, 'readwrite');
        tx.objectStore(IDB_STORE).put(value, key);
        tx.oncomplete = function () { resolve(); };
        tx.onerror = function () { reject(tx.error); };
      });
    });
  }
  function idbGet(key) {
    return idbOpen().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(IDB_STORE, 'readonly');
        var req = tx.objectStore(IDB_STORE).get(key);
        req.onsuccess = function () { resolve(req.result); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function saveFilters() {
    try {
      var activeTab = '';
      document.querySelectorAll('.nav-item.active').forEach(function (b) { activeTab = b.getAttribute('data-tab'); });
      localStorage.setItem(LS_FILTERS, JSON.stringify({
        year: STATE.year, month: STATE.month, machine: STATE.machine,
        sku: STATE.sku, productos: STATE.productos, insumos: STATE.insumos,
        criticos: STATE.criticos, proyecciones: STATE.proyecciones, oeeCalidad: STATE.oeeCalidad, tab: activeTab
      }));
    } catch (e) { /* almacenamiento no disponible — seguir sin persistir */ }
  }
  function restoreFilters(data) {
    try {
      var f = JSON.parse(localStorage.getItem(LS_FILTERS) || 'null');
      if (!f) return null;
      if (f.year != null && data.years.indexOf(f.year) >= 0) STATE.year = f.year;
      if (f.month && MESES.indexOf(f.month) >= 0) STATE.month = f.month;
      if (f.machine && data.machines.indexOf(f.machine) >= 0) STATE.machine = f.machine;
      function maqOk(v) { return v == null || data.machines.indexOf(v) >= 0; }
      if (f.sku) STATE.sku = { maquina: maqOk(f.sku.maquina) ? f.sku.maquina : null, mes: MESES.indexOf(f.sku.mes) >= 0 ? f.sku.mes : null, query: f.sku.query || '' };
      if (f.productos) STATE.productos = {
        maquina: maqOk(f.productos.maquina) ? f.productos.maquina : null,
        mes: MESES.indexOf(f.productos.mes) >= 0 ? f.productos.mes : null,
        espesor: isNum(f.productos.espesor) ? f.productos.espesor : null
      };
      if (f.insumos) STATE.insumos = {
        maquina: maqOk(f.insumos.maquina) ? f.insumos.maquina : null,
        mes: MESES.indexOf(f.insumos.mes) >= 0 ? f.insumos.mes : null
      };
      if (f.criticos) STATE.criticos = {
        mes: MESES.indexOf(f.criticos.mes) >= 0 ? f.criticos.mes : null,
        maquina: maqOk(f.criticos.maquina) ? f.criticos.maquina : null,
        estado: ['critico', 'alerta', 'atencion', 'ok'].indexOf(f.criticos.estado) >= 0 ? f.criticos.estado : null
      };
      if (f.proyecciones) STATE.proyecciones = { maquina: maqOk(f.proyecciones.maquina) ? f.proyecciones.maquina : null };
      if (f.oeeCalidad) STATE.oeeCalidad = {
        maquina: maqOk(f.oeeCalidad.maquina) ? f.oeeCalidad.maquina : null,
        mes: MESES.indexOf(f.oeeCalidad.mes) >= 0 ? f.oeeCalidad.mes : null
      };
      STATE.restored = true; // a saved null machine means the user chose "Todas" — don't re-default it
      return f.tab || null;
    } catch (e) { return null; }
  }

  function initFromBuffer(arrayBuffer, meta) {
    var data = loadWorkbook(arrayBuffer);
    STATE.data = data;
    D = data;
    STATE.engine = makeEngine(data);
    STATE.fileMeta = meta && meta.name ? { name: meta.name, ts: meta.ts } : null;
    STATE.year = data.years.length ? data.years[data.years.length - 1] : new Date().getFullYear();
    STATE.month = pickDefaultMonth(STATE.year);
    STATE.machine = data.machines[0];
    restoreFilters(data); // restores machine/mes/etc. filters — the active tab is not restored: every login lands on Resumen
    if (STATE.sku.query) el('skuSearch').value = STATE.sku.query;
    populateSelectors();
    el('dzScreen').style.display = 'none';
    el('appShell').style.display = 'flex';
    renderAll();
    switchTab('resumen');
    window.scrollTo(0, 0);
  }

  function onFile(file) {
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var buf = e.target.result;
        initFromBuffer(buf, { name: file.name, ts: Date.now() });
        // persist AFTER a successful parse so a bad file never replaces good data
        idbPut('last', { buffer: buf, name: file.name, ts: Date.now() }).catch(function () { });
      } catch (err) {
        showError(err.message || String(err));
      }
    };
    reader.onerror = function () { showError('No se pudo leer el archivo.'); };
    reader.readAsArrayBuffer(file);
  }

  function tryRestoreSaved() {
    if (typeof indexedDB === 'undefined') return;
    idbGet('last').then(function (saved) {
      if (!saved || !saved.buffer) return;
      // only auto-load if the user hasn't already dropped a file
      if (el('appShell').style.display === 'flex') return;
      try { initFromBuffer(saved.buffer, saved); } catch (e) { /* datos guardados corruptos — ignorar */ }
    }).catch(function () { });
  }

  function showError(msg) {
    var box = el('dzError');
    box.textContent = msg;
    box.style.display = 'block';
  }

  function wireEvents() {
    var dz = el('dropzone'), input = el('fileInput');
    dz.addEventListener('click', function () { input.click(); });
    dz.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') input.click(); });
    input.addEventListener('change', function () { if (input.files[0]) onFile(input.files[0]); });
    ['dragenter', 'dragover'].forEach(function (evt) {
      dz.addEventListener(evt, function (e) { e.preventDefault(); dz.classList.add('drag'); });
    });
    ['dragleave', 'drop'].forEach(function (evt) {
      dz.addEventListener(evt, function (e) { e.preventDefault(); dz.classList.remove('drag'); });
    });
    dz.addEventListener('drop', function (e) {
      var f = e.dataTransfer.files[0];
      if (f) onFile(f);
    });
    el('reloadBtn').addEventListener('click', function () {
      el('appShell').style.display = 'none';
      el('dzScreen').style.display = 'flex';
      el('dzError').style.display = 'none';
      input.value = '';
    });
    document.querySelectorAll('.nav-item').forEach(function (btn) {
      btn.addEventListener('click', function () { switchTab(btn.getAttribute('data-tab')); });
    });
    el('menuBtn').addEventListener('click', function (ev) {
      ev.stopPropagation();
      el('sidebar').classList.toggle('open');
    });
    // tap outside the drawer closes it (the ::after backdrop is part of the sidebar,
    // so clicks on the dimmed area land on the sidebar itself — check the target)
    document.addEventListener('click', function (ev) {
      var sb = el('sidebar');
      if (!sb.classList.contains('open')) return;
      if (ev.target === sb || !sb.contains(ev.target)) sb.classList.remove('open');
    });

    var skuSearchTimer = null;
    el('skuSearch').addEventListener('input', function () {
      var val = this.value;
      clearTimeout(skuSearchTimer);
      skuSearchTimer = setTimeout(function () {
        STATE.sku.query = val;
        renderSkuBuscador();
        saveFilters();
      }, 150);
    });
  }

  /* ---- Integración opcional con Google Drive: login del usuario (OAuth) + Google Picker
   * para elegir el Excel/Sheet, y Drive API para leer sus bytes. Requiere que la página se
   * sirva por https:// (no funciona abriendo el archivo local) y que el Client ID esté
   * autorizado para ese origen en Google Cloud Console. El token de acceso vive solo en
   * memoria (nunca se guarda) — cada carga de página vuelve a pedir permiso, silenciosamente
   * si la sesión de Google sigue activa. Solo se persiste el ID del archivo elegido, para que
   * "Actualizar desde Drive" no tenga que reabrir el selector cada vez. ---- */
  var DRIVE_CLIENT_ID = '996470318055-bratn6iehc54js1rlpn9bnvp3037t37k.apps.googleusercontent.com';
  var DRIVE_API_KEY = 'AIzaSyBt1q8xioWVrFMRUedP45jEpMDX5gWJXdw';
  var DRIVE_SCOPE = 'https://www.googleapis.com/auth/drive.readonly';
  var LS_DRIVE_FILE = 'erpAnalyticsDriveFile';
  var driveTokenClient = null, driveAccessToken = null, drivePickerLoaded = false;

  function driveSavedFile() {
    try { return JSON.parse(localStorage.getItem(LS_DRIVE_FILE) || 'null'); } catch (e) { return null; }
  }
  function driveSaveFile(rec) {
    try { localStorage.setItem(LS_DRIVE_FILE, JSON.stringify(rec)); } catch (e) { }
  }
  function driveHint(msg, isErr, ctx) {
    var box = el(ctx === 'side' ? 'driveRefreshHint' : 'dzDriveHint');
    if (!box) return;
    box.textContent = msg || '';
    box.classList.toggle('err', !!isErr);
  }
  function driveSetBusy(busy, ctx) {
    var btn = el('driveConnectBtn'); if (btn) { btn.disabled = busy; btn.classList.toggle('loading', busy); }
    var rbtn = el('driveRefreshBtn'); if (rbtn) { rbtn.disabled = busy; rbtn.classList.toggle('loading', busy); }
    // el ícono chico que se convierte en spinner es fácil de pasar por alto en la
    // barra lateral — el texto del botón también cambia para que quede claro que
    // está trabajando y no solo "no hizo nada"
    var rlabel = el('driveRefreshLabel');
    if (rlabel) rlabel.textContent = busy ? 'Actualizando…' : 'Actualizar desde Drive';
    // el spinner grande solo vive en la pantalla inicial (dzScreen) — el pequeño
    // dentro del botón ya es suficiente para "Actualizar desde Drive" en el sidebar
    if (ctx !== 'side') { var big = el('dzBigLoading'); if (big) big.classList.toggle('show', busy); }
  }

  function driveEnsureToken(promptMode, onReady, onError) {
    if (typeof google === 'undefined' || !google.accounts || !google.accounts.oauth2) {
      onError('No se pudo cargar el inicio de sesión de Google. Revisa tu conexión e intenta de nuevo.');
      return;
    }
    if (!driveTokenClient) {
      driveTokenClient = google.accounts.oauth2.initTokenClient({
        client_id: DRIVE_CLIENT_ID,
        scope: DRIVE_SCOPE,
        callback: function () { }
      });
    }
    driveTokenClient.callback = function (resp) {
      if (resp && resp.access_token) { driveAccessToken = resp.access_token; onReady(); }
      else onError('No se concedió acceso a Google Drive.');
    };
    driveTokenClient.error_callback = function (err) {
      onError((err && err.message) || 'No se pudo iniciar sesión con Google.');
    };
    driveTokenClient.requestAccessToken({ prompt: promptMode || '' });
  }

  function driveEnsurePicker(cb, onError) {
    if (drivePickerLoaded) { cb(); return; }
    if (typeof gapi === 'undefined') {
      onError('No se pudo cargar el selector de archivos de Google. Revisa tu conexión e intenta de nuevo.');
      return;
    }
    gapi.load('picker', function () { drivePickerLoaded = true; cb(); });
  }

  function driveOpenPicker() {
    driveHint('Conectando con base de datos…', false, 'dz');
    driveSetBusy(true, 'dz');
    driveEnsureToken('', function () {
      driveEnsurePicker(function () {
        driveHint('', false, 'dz');
        driveSetBusy(false, 'dz');
        var view = new google.picker.DocsView(google.picker.ViewId.DOCS)
          .setMimeTypes('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,application/vnd.google-apps.spreadsheet')
          .setSelectFolderEnabled(false);
        var picker = new google.picker.PickerBuilder()
          .addView(view)
          .setOAuthToken(driveAccessToken)
          .setDeveloperKey(DRIVE_API_KEY)
          .setCallback(driveOnPicked)
          .build();
        picker.setVisible(true);
      }, function (msg) { driveSetBusy(false, 'dz'); driveHint(msg, true, 'dz'); });
    }, function (msg) { driveSetBusy(false, 'dz'); driveHint(msg, true, 'dz'); });
  }

  function driveOnPicked(data) {
    if (data.action !== google.picker.Action.PICKED) return;
    var doc = data.docs[0];
    driveDownloadFile(doc.id, doc.name, doc.mimeType, 'dz');
  }

  function driveDownloadFile(fileId, name, mimeType, ctx) {
    driveHint('Descargando "' + name + '" desde Drive…', false, ctx);
    driveSetBusy(true, ctx);
    var isGoogleSheet = mimeType === 'application/vnd.google-apps.spreadsheet';
    var url = isGoogleSheet
      ? 'https://www.googleapis.com/drive/v3/files/' + fileId + '/export?mimeType=application%2Fvnd.openxmlformats-officedocument.spreadsheetml.sheet'
      : 'https://www.googleapis.com/drive/v3/files/' + fileId + '?alt=media';
    fetch(url, { headers: { Authorization: 'Bearer ' + driveAccessToken } })
      .then(function (res) {
        if (!res.ok) throw new Error('Google Drive respondió ' + res.status + '.');
        return res.arrayBuffer();
      })
      .then(function (buf) {
        initFromBuffer(buf, { name: name, ts: Date.now() });
        idbPut('last', { buffer: buf, name: name, ts: Date.now() }).catch(function () { });
        driveSaveFile({ id: fileId, name: name, mimeType: mimeType });
        driveUpdateRefreshUi();
        driveHint('', false, ctx);
        driveSetBusy(false, ctx);
      })
      .catch(function (err) {
        driveSetBusy(false, ctx);
        driveHint('No se pudo leer el archivo desde Drive: ' + (err.message || err), true, ctx);
        if (ctx === 'dz') showError('No se pudo leer el archivo desde Google Drive.');
      });
  }

  function driveUpdateRefreshUi() {
    var saved = driveSavedFile();
    var btn = el('driveRefreshBtn');
    if (!btn) return;
    if (saved) {
      btn.style.display = '';
      btn.title = saved.name;
    } else {
      btn.style.display = 'none';
    }
  }

  function driveRefresh() {
    var saved = driveSavedFile();
    if (!saved) return;
    driveHint('Conectando con Google…', false, 'side');
    driveSetBusy(true, 'side');
    driveEnsureToken('', function () {
      driveDownloadFile(saved.id, saved.name, saved.mimeType, 'side');
    }, function (msg) { driveSetBusy(false, 'side'); driveHint(msg, true, 'side'); });
  }

  function wireDriveUI() {
    var connectBtn = el('driveConnectBtn');
    if (connectBtn) connectBtn.addEventListener('click', driveOpenPicker);
    var refreshBtn = el('driveRefreshBtn');
    if (refreshBtn) refreshBtn.addEventListener('click', driveRefresh);
    driveUpdateRefreshUi();
  }

  /* ---- Pantalla de acceso (usuario/clave). Es una barrera simple, no seguridad real:
   * como todo el panel es un solo archivo HTML sin servidor, la clave se valida en el
   * propio navegador — alguien con conocimientos técnicos podría saltársela revisando
   * el código. Se compara un hash SHA-256 en vez de la clave en texto plano solo para
   * no dejarla a la vista en el código fuente a simple lectura. La sesión no se guarda:
   * cada vez que se abre o recarga la página, vuelve a pedir usuario y clave. ---- */
  var LOGIN_USERS = {
    'gino espinosa': '629f4cf9337b0d0c76f305d860f98894cfa8c279516b425747514ca8710deb97', // 007
    'admin': 'a665a45920422f9d417e4867efdc4fb8a04a1f3fff1fa07e998e86f7f7a27ae3' // 123
  };
  function sha256Hex(str) {
    return crypto.subtle.digest('SHA-256', new TextEncoder().encode(str)).then(function (buf) {
      return Array.prototype.map.call(new Uint8Array(buf), function (b) { return b.toString(16).padStart(2, '0'); }).join('');
    });
  }
  function wireLogin(onAuthed) {
    el('logoutBtn').addEventListener('click', function () { location.reload(); });
    var form = el('loginForm');
    form.addEventListener('submit', function (ev) {
      ev.preventDefault();
      var userKey = el('loginUser').value.trim().toLowerCase();
      var expectedHash = LOGIN_USERS[userKey];
      sha256Hex(el('loginPass').value).then(function (hash) {
        if (expectedHash && hash === expectedHash) {
          document.documentElement.setAttribute('data-authed', '1');
          el('loginScreen').style.display = 'none';
          el('dzScreen').style.display = 'flex';
          onAuthed();
        } else {
          el('loginError').style.display = 'block';
          form.classList.remove('shake');
          void form.offsetWidth; // reinicia la animación aunque se repita el error
          form.classList.add('shake');
          el('loginPass').value = '';
          el('loginPass').focus();
        }
      });
    });
  }

  /* ---- Modo claro/oscuro: alterna sobre el tema efectivo y persiste la elección ---- */
  function currentTheme() {
    var t = document.documentElement.getAttribute('data-theme');
    if (t === 'dark' || t === 'light') return t;
    return (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) ? 'dark' : 'light';
  }
  function updateThemeIcon() {
    var dark = currentTheme() === 'dark';
    el('themeIconMoon').style.display = dark ? 'none' : '';
    el('themeIconSun').style.display = dark ? '' : 'none';
  }
  function wireTheme() {
    updateThemeIcon();
    el('themeBtn').addEventListener('click', function () {
      var next = currentTheme() === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', next);
      try { localStorage.setItem('erpAnalyticsTheme', next); } catch (e) { }
      updateThemeIcon();
    });
  }

  document.addEventListener('DOMContentLoaded', function () {
    // wireLogin() goes FIRST and unguarded: if anything else below throws, the login
    // form's submit handler must already be attached, or the "Ingresar" button does
    // nothing (default form submit just reloads the page — looks exactly like a
    // silent failure). Everything else is wrapped so one broken piece (e.g. a Drive
    // API script that failed to load) can't take the rest of the app down with it.
    wireLogin(function () {
      try { wireEvents(); } catch (e) { console.error('wireEvents failed', e); }
      try { wireDriveUI(); } catch (e) { console.error('wireDriveUI failed', e); }
      try { el('printBtn').addEventListener('click', function () { window.print(); }); } catch (e) { console.error('printBtn wiring failed', e); }
      // tryRestoreSaved() can reveal the previously loaded Excel straight away — only
      // runs after a successful login, so a not-yet-authenticated visitor never sees
      // real data appear behind the password form.
      try { tryRestoreSaved(); } catch (e) { console.error('tryRestoreSaved failed', e); }
    });
    try { wireTheme(); } catch (e) { console.error('wireTheme failed', e); }
  });
})();
