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
  function trimStr(s) { return s == null ? '' : String(s).trim(); }
  function isNum(v) { return typeof v === 'number' && isFinite(v); }
  function toNum(v) {
    if (v == null || v === '' || v === '-') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    var n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
    return isFinite(n) ? n : null;
  }
  function zeroToNull(v) { return (v == null || !isFinite(v) || v === 0) ? null : v; }
  function ratio(num, den) {
    if (num == null || den == null || den === 0) return null;
    return num / den;
  }
  function sum(arr) { var s = 0, has = false; for (var i = 0; i < arr.length; i++) { var v = arr[i]; if (isNum(v)) { s += v; has = true; } } return has ? s : 0; }
  function isMachineLike(name) { return /^(perfiladora|tubera)[0-9]+$/.test(normKey(name)); }

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
    var extra = n == null ? ' dash' : (cls || '');
    return '<td class="' + extra.trim() + '">' + (n == null ? '-' : formatter(n)) + '</td>';
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
      cClas = idx['clasificacionprod1'];
    if (cAno == null || cMes == null || cMaq == null) return out;
    for (var r = 1; r < aoa.length; r++) {
      var row = aoa[r];
      var maq = trimStr(row[cMaq]);
      if (!maq) continue;
      out.push({
        anio: toNum(row[cAno]),
        mes: toNum(row[cMes]),
        maquina: maq,
        fecha: cFecha != null ? row[cFecha] : null,
        largo: cLargo != null ? toNum(row[cLargo]) : null,
        cantProducida: cCant != null ? toNum(row[cCant]) : null,
        kgReal: cKgReal != null ? toNum(row[cKgReal]) : null,
        totalUnEst: cTotalUE != null ? toNum(row[cTotalUE]) : null,
        espesor: cEsp != null ? toNum(row[cEsp]) : null,
        clasif: cClas != null ? trimStr(row[cClas]) : ''
      });
    }
    return out;
  }

  function parseChatarra(aoa) {
    var out = [];
    if (!aoa || aoa.length < 2) return out;
    var idx = headerIndex(aoa[0]);
    var cAno = idx['ano'], cMes = idx['mes'], cMaq = idx['nommaquina'], cTotalUE = idx['totalunest'];
    if (cAno == null || cMes == null || cMaq == null) return out;
    for (var r = 1; r < aoa.length; r++) {
      var row = aoa[r];
      var maq = trimStr(row[cMaq]);
      if (!maq) continue;
      out.push({
        anio: toNum(row[cAno]),
        mes: toNum(row[cMes]),
        maquina: maq,
        totalUnEst: cTotalUE != null ? toNum(row[cTotalUE]) : null
      });
    }
    return out;
  }

  function parseValesConsumo(aoa) {
    var out = [];
    if (!aoa || aoa.length < 2) return out;
    var idx = headerIndex(aoa[0]);
    var cFecha = idx['fechacontab'], cValor = idx['valortotal'],
      cCC = idx['descripccosto'], cCuenta = idx['nombredecuenta'];
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
        valor: valor,
        ccosto: cCC != null ? trimStr(row[cCC]) : '',
        cuenta: cCuenta != null ? trimStr(row[cCuenta]) : ''
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
    var out = {}; // out[machine][mes] = summed budget $ across cost-centers
    if (!aoa || aoa.length < 2) return out;
    var header = aoa[0];
    var monthCols = [];
    for (var c = 2; c < header.length; c++) {
      var h = trimStr(header[c]);
      var mi = MESES.findIndex(function (m) { return normKey(m) === normKey(h); });
      if (mi >= 0) monthCols.push({ mes: MESES[mi], col: c });
    }
    for (var r = 1; r < aoa.length; r++) {
      var row = aoa[r];
      var maq = trimStr(row[1]);
      if (!maq) continue;
      if (!out[maq]) {
        out[maq] = {};
        monthCols.forEach(function (mc) { out[maq][mc.mes] = 0; });
      }
      monthCols.forEach(function (mc) {
        var v = toNum(row[mc.col]);
        if (v != null) out[maq][mc.mes] += v;
      });
    }
    return out;
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

  function parseRechazoTabla(aoa) {
    // 'no borrar' style table: espesor, rechazo total (%), material (short code)
    var out = [];
    if (!aoa || aoa.length < 2) return out;
    var idx = headerIndex(aoa[0]);
    var cEsp = idx['espesor'], cRech = idx['rechazototal'], cMat = idx['material'];
    if (cEsp == null || cRech == null || cMat == null) return out;
    for (var r = 1; r < aoa.length; r++) {
      var row = aoa[r];
      var esp = toNum(row[cEsp]), rech = toNum(row[cRech]), mat = trimStr(row[cMat]);
      if (esp == null || rech == null || !mat) continue;
      out.push({ espesor: esp, rechazo: rech, material: mat });
    }
    return out;
  }

  /* =========================================================================
   * Workbook loader
   * ======================================================================= */

  function loadWorkbook(arrayBuffer) {
    var wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });

    var wsProduccion = findSheet(wb, 'produccion');
    var wsChatarra = findSheet(wb, 'chatarra');
    var wsVales = findSheet(wb, 'valesdeconsumo', { startsWith: true });
    var wsBobinas = findSheet(wb, 'bobinasm');
    var wsProdResumen = findSheet(wb, 'prodresumenanu');
    var wsOEE = findSheet(wb, 'oeepl');
    var wsHoras = findSheet(wb, 'horasdeatraso');
    var wsPptoProd = findSheet(wb, 'pptoproduccion');
    var wsPptoVales = findSheet(wb, 'pptovalesdeconsumo');
    var wsMetas = findSheet(wb, 'metas');
    var wsNoBorrar = findSheet(wb, 'noborrar');

    var missing = [];
    if (!wsProduccion) missing.push('Produccion');
    if (!wsChatarra) missing.push('Chatarra');
    if (missing.length) {
      throw new Error('No se encontraron las hojas requeridas: ' + missing.join(', ') +
        '. Verifica que el Excel tenga la misma estructura del panel INDAMA.');
    }

    var produccion = parseProduccion(sheetToAOA(wsProduccion));
    var chatarra = parseChatarra(sheetToAOA(wsChatarra));
    var vales = wsVales ? parseValesConsumo(sheetToAOA(wsVales)) : [];
    var bobinasM = wsBobinas ? parseMonthlyMatrixByLabel(sheetToAOA(wsBobinas), 'descripcion') : {};
    var prodResumenAnu = wsProdResumen ? parseMonthlyMatrixByLabel(sheetToAOA(wsProdResumen), 'nommaq') : {};
    var oee = wsOEE ? parseOEEPl(sheetToAOA(wsOEE)) : { machines: [], data: {} };
    var horasAtraso = wsHoras ? parseHorasAtraso(sheetToAOA(wsHoras)) : {};
    var pptoProduccion = wsPptoProd ? parsePptoProduccion(sheetToAOA(wsPptoProd)) : {};
    var pptoVales = wsPptoVales ? parsePptoValesConsumo(sheetToAOA(wsPptoVales)) : {};
    var metasCandidates = [wsMetas, wsNoBorrar].filter(Boolean).map(sheetToAOA);
    var metas = parseMetas(metasCandidates);
    var rechazoTabla = wsNoBorrar ? parseRechazoTabla(sheetToAOA(wsNoBorrar)) : [];

    var machines = oee.machines.length ? oee.machines : Array.from(new Set(
      produccion.filter(function (p) { return isMachineLike(p.maquina); }).map(function (p) { return p.maquina; })
    ));

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
      pptoProduccion: pptoProduccion, pptoVales: pptoVales, metas: metas,
      rechazoTabla: rechazoTabla, machines: machines, groupMachines: groupMachines, years: years
    };
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
      var v = m[label][mesIdx];
      return v;
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
        chatarraPlantaPct: chatarraPlantaPct, metaChatarra: META_CHATARRA_PLANTA, desviacionChatarra: desviacionChatarra
      };
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
        var meta = pct == null ? null : META_CHATARRA_PLANTA;
        var desv = (meta != null && pct != null) ? meta - pct : null;
        return { mes: m.mes, mesAbbr: m.mesAbbr, real: real, pct: pct, meta: meta, desv: desv };
      });
      return { produccion: produccion, chatarra: chatarra };
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
        var metaChatEst = (isBraner || prodEst == null) ? null : chatarraMetaEstandarVal(maq);
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
          chatarra: chat, chatarraPct: chatPct, metaChatarraEstandar: metaChatEst, desviacionChatarraEstandar: desvChatEst,
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
        var metaChatEst = prodEst == null ? null : chatarraMetaEstandarVal(maquina);
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
          prodEstandar: prodEst, metrosLineales: ml, factor: factor, unidades: unidades,
          chatarra: chat, chatarraPct: chatPct, metaChatarraEstandar: metaChatEst, desviacionChatarraEstandar: desvChatEst,
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

    return {
      resumenPlanta: resumenPlanta, kpiHeader: kpiHeader, presupuesto: presupuesto,
      detalleMes: detalleMes, detalleAnio: detalleAnio, espesorAnalysis: espesorAnalysis,
      pptoTotalVal: pptoTotalVal
    };
  }

  /* =========================================================================
   * Chart rendering (bullet-style bar + target tick, per dataviz mark specs)
   * ======================================================================= */

  var chartUid = 0;
  function renderBulletChart(container, opts) {
    // opts: {categories, values, targets, formatValue, formatTarget, barColorVar, seriesLabel, targetLabel}
    chartUid++;
    var W = 640, H = 220, padL = 8, padR = 8, padT = 10, padB = 26;
    var plotW = W - padL - padR, plotH = H - padT - padB;
    var n = opts.categories.length;
    var bandW = plotW / n;
    var barW = Math.max(6, Math.min(28, bandW * 0.5));

    var allVals = opts.values.concat(opts.targets).filter(isNum);
    var maxV = allVals.length ? Math.max.apply(null, allVals) : 1;
    var minV = allVals.length ? Math.min(0, Math.min.apply(null, allVals)) : 0;
    maxV = maxV * 1.15 || 1;

    function y(v) { return padT + plotH - ((v - minV) / (maxV - minV)) * plotH; }
    var baseline = y(0);

    var gridLines = 4, gridHtml = '', labelsHtml = '';
    for (var g = 0; g <= gridLines; g++) {
      var v = minV + (maxV - minV) * g / gridLines;
      var yy = y(v);
      gridHtml += '<line class="grid-line" x1="' + padL + '" x2="' + (W - padR) + '" y1="' + yy + '" y2="' + yy + '"></line>';
      labelsHtml += '<text class="axis-label" x="' + (padL) + '" y="' + (yy - 3) + '">' + opts.formatValue(v, true) + '</text>';
    }

    var barsHtml = '', ticksHtml = '', xLabelsHtml = '';
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
        bars.push({ i: i, cx: cx, val: val, tgt: tgt, cat: cat });
      }
      if (isNum(tgt)) {
        var yt = y(tgt);
        ticksHtml += '<line x1="' + (cx - barW / 2 - 3) + '" x2="' + (cx + barW / 2 + 3) + '" y1="' + yt + '" y2="' + yt + '" stroke="var(--text-primary)" stroke-width="2"></line>';
      }
    });

    var legendHtml = '<div class="chart-legend">' +
      '<span class="sw"><span class="dot" style="background:' + opts.barColorVar + '"></span>' + opts.seriesLabel + '</span>' +
      '<span class="sw"><span class="dot" style="background:var(--text-primary)"></span>' + opts.targetLabel + '</span>' +
      '</div>';

    var svg = '<div class="chart-wrap">' + legendHtml +
      '<svg class="chart" viewBox="0 0 ' + W + ' ' + H + '" id="svg' + chartUid + '">' +
      '<line class="baseline" x1="' + padL + '" x2="' + (W - padR) + '" y1="' + baseline + '" y2="' + baseline + '"></line>' +
      gridHtml + labelsHtml + barsHtml + ticksHtml + xLabelsHtml +
      '</svg>' +
      '<div class="chart-tooltip" id="' + tipId + '"></div>' +
      '</div>';
    container.innerHTML = svg;

    var tip = container.querySelector('#' + tipId);
    var svgEl = container.querySelector('#svg' + chartUid);
    var wrapEl = container.querySelector('.chart-wrap');
    container.querySelectorAll('.bar').forEach(function (barEl) {
      var i = parseInt(barEl.getAttribute('data-i'), 10);
      var b = bars.filter(function (x) { return x.i === i; })[0];
      if (!b) return;
      barEl.addEventListener('mousemove', function (ev) {
        var rect = wrapEl.getBoundingClientRect();
        var scale = rect.width / W;
        var left = (b.cx * scale);
        var top = (y(b.val) * scale);
        tip.style.left = left + 'px';
        tip.style.top = top + 'px';
        tip.style.opacity = 1;
        tip.innerHTML = '<strong>' + b.cat + '</strong><br>' + opts.seriesLabel + ': ' + opts.formatValue(b.val) +
          (isNum(b.tgt) ? '<br>' + opts.targetLabel + ': ' + opts.formatValue(b.tgt) : '');
      });
      barEl.addEventListener('mouseleave', function () { tip.style.opacity = 0; });
    });
  }

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

  var STATE = { data: null, engine: null, year: null, month: null, machine: null };

  function el(id) { return document.getElementById(id); }

  function fmtPctTick(v) { return Math.round(v * 100) + '%'; }
  function fmtKgTick(v) {
    if (Math.abs(v) >= 1e6) return (v / 1e6).toFixed(1) + 'M';
    if (Math.abs(v) >= 1e3) return Math.round(v / 1e3) + 'k';
    return Math.round(v);
  }

  function renderKPIs(kpi) {
    var tiles = [
      { label: 'OEE Planta Global', value: fmtPct(kpi.oeePlantaGlobal, 1) },
      { label: 'Kilos Fabricados', value: fmtInt(kpi.kilosFabricados) + ' kg' },
      { label: 'Chatarra', value: fmtInt(kpi.chatarraKg) + ' kg' },
      { label: 'Chatarra Máquinas', value: fmtPct(kpi.chatarraMaquinasPct, 2) },
      { label: 'Chatarra Braner', value: fmtPct(kpi.chatarraBranerPct, 2) },
      { label: 'Chatarra Planta', value: fmtPct(kpi.chatarraPlantaPct, 2) },
      { label: 'Meta Chatarra', value: fmtPct(kpi.metaChatarra, 1) },
      {
        label: 'Desviación', value: fmtSigned(kpi.desviacionChatarra, function (v) { return fmtPct(v, 2); }),
        deltaClass: deltaClass(kpi.desviacionChatarra, true)
      }
    ];
    el('kpiGrid').innerHTML = tiles.map(function (t) {
      return '<div class="tile"><div class="label">' + t.label + '</div><div class="value">' + t.value + '</div>' +
        (t.deltaClass ? '<div class="delta ' + t.deltaClass + '"></div>' : '') + '</div>';
    }).join('');
  }

  function renderResumen(anio) {
    var resumen = STATE.engine.resumenPlanta(anio);
    var rows = [
      { label: 'Total Producción', unit: 'kg', values: resumen.map(function (m) { return m.totalProduccion; }), total: true },
      { label: 'Total Chatarra Real', unit: 'kg', values: resumen.map(function (m) { return m.totalChatarraReal; }) },
      { label: 'Chatarra Tuberas & Perfiladoras', unit: '%', values: resumen.map(function (m) { return m.chatarraMaquinasPct; }), fmt: function (v) { return fmtPct(v, 2); } },
      { label: 'Desorillado Braner', unit: '%', values: resumen.map(function (m) { return m.desorilladoBranerPct; }), fmt: function (v) { return fmtPct(v, 2); } },
      { label: 'Chatarra Planta', unit: '%', values: resumen.map(function (m) { return m.pctChatarraPlanta; }), fmt: function (v) { return fmtPct(v, 2); } },
      { label: 'OEE Planta Mensual', unit: '%', values: resumen.map(function (m) { return m.oeeMensual; }), fmt: function (v) { return fmtPct(v, 1); } }
    ];
    renderMonthlyTable(el('resumenTable'), 'Planta ' + anio, rows);
    return resumen;
  }

  function renderPresupuesto(anio, resumen) {
    var ppto = STATE.engine.presupuesto(anio, resumen);

    renderBulletChart(el('chartProdPpto'), {
      categories: ppto.produccion.map(function (m) { return m.mesAbbr; }),
      values: ppto.produccion.map(function (m) { return m.real; }),
      targets: ppto.produccion.map(function (m) { return m.ppto; }),
      formatValue: fmtKgTick, barColorVar: 'var(--series-1)',
      seriesLabel: 'Producción real', targetLabel: 'Presupuesto'
    });

    renderBulletChart(el('chartChatPpto'), {
      categories: ppto.chatarra.map(function (m) { return m.mesAbbr; }),
      values: ppto.chatarra.map(function (m) { return m.pct; }),
      targets: ppto.chatarra.map(function (m) { return m.meta; }),
      formatValue: fmtPctTick, barColorVar: 'var(--series-6)',
      seriesLabel: 'Chatarra planta', targetLabel: 'Meta (4,6%)'
    });

    var rows = [
      { label: 'Total Producción', unit: 'kg', values: ppto.produccion.map(function (m) { return m.real; }) },
      { label: 'Presupuesto Producción', unit: 'kg', values: ppto.produccion.map(function (m) { return m.ppto; }) },
      { label: 'Desviación Ppto vs Real', unit: 'kg', values: ppto.produccion.map(function (m) { return m.desv; }), fmt: fmtInt },
      { label: 'Desviación Ppto vs Real', unit: '%', values: ppto.produccion.map(function (m) { return m.desvPct; }), fmt: function (v) { return fmtPct(v, 1); } },
      { label: 'Total Chatarra Real', unit: 'kg', values: ppto.chatarra.map(function (m) { return m.real; }) },
      { label: 'Chatarra Planta', unit: '%', values: ppto.chatarra.map(function (m) { return m.pct; }), fmt: function (v) { return fmtPct(v, 2); } },
      { label: 'Meta', unit: '%', values: ppto.chatarra.map(function (m) { return m.meta; }), fmt: function (v) { return fmtPct(v, 1); } },
      { label: 'Desviación', unit: '%', values: ppto.chatarra.map(function (m) { return m.desv; }), fmt: function (v) { return fmtPct(v, 2); } }
    ];
    renderMonthlyTable(el('pptoTable'), 'Producción vs. Presupuesto — ' + anio, rows);
  }

  function renderDetalleMes(anio, mesNombre) {
    var det = STATE.engine.detalleMes(anio, mesNombre);
    var maquinas = det.columnas.map(function (c) { return c.maquina; });

    renderMachineTable(el('detalleMesProd'), 'Producción & Chatarra — ' + mesNombre, maquinas, [
      { label: 'Producción Estándar', unit: 'kg', values: det.columnas.map(function (c) { return c.prodEstandar; }) },
      { label: 'Metros Lineales', unit: 'm', values: det.columnas.map(function (c) { return c.metrosLineales; }) },
      { label: 'Factor Kilo/Metro', unit: 'F', values: det.columnas.map(function (c) { return c.factor; }), fmt: function (v) { return fmt1(v); } },
      { label: 'Unidades Fabricadas', unit: 'un', values: det.columnas.map(function (c) { return c.unidades; }) },
      { label: 'Chatarra', unit: 'kg', values: det.columnas.map(function (c) { return c.chatarra; }) },
      { label: 'Chatarra', unit: '%', values: det.columnas.map(function (c) { return c.chatarraPct; }), fmt: function (v) { return fmtPct(v, 2); } },
      { label: 'Meta Chatarra Estándar', unit: '%', values: det.columnas.map(function (c) { return c.metaChatarraEstandar; }), fmt: function (v) { return fmtPct(v, 2); } },
      { label: 'Desviación', unit: '%', values: det.columnas.map(function (c) { return c.desviacionChatarraEstandar; }), fmt: function (v) { return fmtPct(v, 2); } }
    ]);

    var valesRows = [
      { label: 'Valor Vales de Consumo', unit: '$', values: det.columnas.map(function (c) { return c.valorVales; }), fmt: fmtMoney },
      { label: 'Costo por Metro Lineal', unit: '$', values: det.columnas.map(function (c) { return c.costoPorMetro; }), fmt: fmtMoney },
      { label: 'Costo Por Kilo', unit: '$', values: det.columnas.map(function (c) { return c.costoPorKilo; }), fmt: fmtMoney },
      { label: 'Meta Presupuesto', unit: '$', values: det.columnas.map(function (c) { return c.metaPresupuesto; }), fmt: fmtMoney },
      { label: 'Desviación', unit: '$', values: det.columnas.map(function (c) { return c.desviacionCosto; }), fmt: fmtMoney }
    ];
    renderMachineTable(el('detalleMesVales'), 'Vales de Consumo — ' + mesNombre, maquinas, valesRows);
    if (det.costoPlanta != null) {
      el('detalleMesVales').innerHTML += '<p class="cap">Costo Por Kilo Planta: ' + fmtMoney(det.costoPlanta) +
        (det.metaPlanta != null ? ' · Meta: ' + fmtMoney(det.metaPlanta) + ' · Desviación: ' + fmtMoney(det.desviacionCostoPlanta) : '') + '</p>';
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
    renderMonthlyTable(el('detalleAnioProd'), 'Producción & Chatarra — ' + maquina, [
      { label: 'Producción Estándar', unit: 'kg', values: det.meses.map(function (m) { return m.prodEstandar; }) },
      { label: 'Metros Lineales', unit: 'm', values: det.meses.map(function (m) { return m.metrosLineales; }) },
      { label: 'Factor Kilo/Metro', unit: 'F', values: det.meses.map(function (m) { return m.factor; }), fmt: fmt1 },
      { label: 'Unidades Fabricadas', unit: 'un', values: det.meses.map(function (m) { return m.unidades; }) },
      { label: 'Chatarra', unit: 'kg', values: det.meses.map(function (m) { return m.chatarra; }) },
      { label: 'Chatarra', unit: '%', values: det.meses.map(function (m) { return m.chatarraPct; }), fmt: function (v) { return fmtPct(v, 2); } },
      { label: 'Meta Chatarra Estándar', unit: '%', values: det.meses.map(function (m) { return m.metaChatarraEstandar; }), fmt: function (v) { return fmtPct(v, 2); } },
      { label: 'Desviación', unit: '%', values: det.meses.map(function (m) { return m.desviacionChatarraEstandar; }), fmt: function (v) { return fmtPct(v, 2); } }
    ]);
    renderMonthlyTable(el('detalleAnioVales'), 'Vales de Consumo — ' + maquina, [
      { label: 'Valor Vales de Consumo', unit: '$', values: det.meses.map(function (m) { return m.valorVales; }), fmt: fmtMoney },
      { label: 'Costo por Metro Lineal', unit: '$', values: det.meses.map(function (m) { return m.costoPorMetro; }), fmt: fmtMoney },
      { label: 'Costo Por Kilo', unit: '$', values: det.meses.map(function (m) { return m.costoPorKilo; }), fmt: fmtMoney },
      { label: 'Meta Presupuesto', unit: '$', values: det.meses.map(function (m) { return m.metaPresupuesto; }), fmt: fmtMoney },
      { label: 'Desviación', unit: '$', values: det.meses.map(function (m) { return m.desviacionCosto; }), fmt: fmtMoney }
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
    var pctRows = esp.pctTable.map(function (r) { return { label: r.espesor + ' mm', values: r.meses, fmt: function (v) { return fmtPct(v, 1); } }; });
    pctRows.push({ label: 'Total', values: esp.totalPorMes.map(function (v) { return v ? 1 : null; }), fmt: function (v) { return fmtPct(v, 0); }, total: true });
    renderMonthlyTable(el('espesorPctTable'), 'Producción por Espesor (%) — ' + anio, pctRows);

    var kgRows = esp.kgTable.map(function (r) { return { label: r.espesor + ' mm', values: r.meses }; });
    kgRows.push({ label: 'Total', values: esp.totalPorMes, total: true });
    renderMonthlyTable(el('espesorKgTable'), 'Producción por Espesor (kg) — ' + anio, kgRows);
  }

  function populateSelectors() {
    var years = STATE.data.years;
    el('yearSelect').innerHTML = years.map(function (y) { return '<option value="' + y + '">' + y + '</option>'; }).join('');
    el('yearSelect').value = STATE.year;

    el('monthSelect').innerHTML = MESES.map(function (m) { return '<option value="' + m + '">' + m + '</option>'; }).join('');
    el('monthSelect').value = STATE.month;

    var machines = STATE.data.machines;
    el('machineSelect').innerHTML = machines.map(function (m) { return '<option value="' + escapeHtml(m) + '">' + escapeHtml(m) + '</option>'; }).join('');
    el('machineSelect').value = STATE.machine;
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

  function renderAll() {
    var anio = STATE.year;
    var resumen = renderResumen(anio);
    var kpi = STATE.engine.kpiHeader(anio, resumen);
    renderKPIs(kpi);
    renderPresupuesto(anio, resumen);
    renderDetalleMes(anio, STATE.month);
    renderDetalleAnio(anio, STATE.machine);
    renderEspesor(anio);
    el('updatedLabel').textContent = 'Año ' + anio + ' · generado ' + new Date().toLocaleString('es-CL');
  }

  function onFile(file) {
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var data = loadWorkbook(e.target.result);
        STATE.data = data;
        D = data;
        STATE.engine = makeEngine(data);
        STATE.year = data.years.length ? data.years[data.years.length - 1] : new Date().getFullYear();
        STATE.month = pickDefaultMonth(STATE.year);
        STATE.machine = data.machines[0];
        populateSelectors();
        el('dzScreen').style.display = 'none';
        el('dashboard').style.display = 'block';
        el('topControls').style.display = 'flex';
        renderAll();
      } catch (err) {
        showError(err.message || String(err));
      }
    };
    reader.onerror = function () { showError('No se pudo leer el archivo.'); };
    reader.readAsArrayBuffer(file);
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
      el('dashboard').style.display = 'none';
      el('topControls').style.display = 'none';
      el('dzScreen').style.display = 'block';
      el('dzError').style.display = 'none';
      input.value = '';
    });
    el('yearSelect').addEventListener('change', function () {
      STATE.year = parseInt(this.value, 10);
      STATE.month = pickDefaultMonth(STATE.year);
      populateSelectors();
      renderAll();
    });
    el('monthSelect').addEventListener('change', function () {
      STATE.month = this.value;
      renderDetalleMes(STATE.year, STATE.month);
    });
    el('machineSelect').addEventListener('change', function () {
      STATE.machine = this.value;
      renderDetalleAnio(STATE.year, STATE.machine);
    });
  }

  document.addEventListener('DOMContentLoaded', wireEvents);
})();
