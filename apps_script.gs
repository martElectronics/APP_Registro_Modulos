/**
 * apps_script.gs — Backend de Google Apps Script para el registro de módulos.
 *
 * Recibe por POST el CSV de módulos que envía serial_to_sheet.py (capturado del
 * BMS Master por Serial con la tecla 'g') y crea UNA HOJA NUEVA por cada toma,
 * con nombre = fecha y hora del momento, formateada como mapa de calor.
 *
 * NO vive aquí en ejecución: este archivo es solo la COPIA VERSIONADA. El código
 * real corre en el editor de Apps Script ligado a la Google Sheet destino.
 *
 * ── DESPLIEGUE ───────────────────────────────────────────────────────────────
 *   1. Google Sheet destino -> Extensiones -> Apps Script.
 *   2. Pega este código (reemplaza todo) y Guarda.
 *   3. Implementar -> Nueva implementación -> Aplicación web:
 *        Ejecutar como: Yo   |   Quién tiene acceso: CUALQUIER USUARIO.
 *   4. Copia la URL (.../exec) y ponla en serial_to_sheet.py (APPS_SCRIPT_URL).
 *   Para ACTUALIZAR tras editar: Gestionar implementaciones -> editar (lápiz) ->
 *   Versión: Nueva -> Implementar (la URL no cambia).
 *
 * ⚠ El TOKEN de aquí debe COINCIDIR con el de serial_to_sheet.py (si no: 'bad token').
 *
 * ── FORMATO DE CADA HOJA ─────────────────────────────────────────────────────
 *   Cabecera: Modulo | V1..V11 | T1..T9 | Delta V | Delta T
 *   - V: degradado verde (alta=verde, baja=blanco), escalado por módulo.
 *   - T: degradado rojo (alta=rojo, baja=blanco), escalado por módulo.
 *   - Delta V/T por módulo = máx−mín de sus celdas válidas.
 *   - Fila "Total" = delta global del pack.
 *   - Valor fuera de [V_MIN,V_MAX] / [T_MIN,T_MAX] -> celda GRIS, excluido de
 *     color, delta y del rodeo de extremos (NTC/celda mala).
 *   - Borde grueso en la celda más alta y más baja de cada módulo (V y T).
 */

const TOKEN = 'MART_cambia_esto_2026';   // debe coincidir con serial_to_sheet.py

const GREEN = [0, 176, 80];   // V altas = verde, bajas = blanco
const RED   = [255, 0, 0];    // T altas = rojo,  bajas = blanco
const GRAY  = '#cccccc';      // valor fuera de rango (NTC/celda mala)

// Límites de validez (iguales que el Config.json de la app). Ajusta si quieres.
const V_MIN = 0, V_MAX = 100;
const T_MIN = 0, T_MAX = 80;

function heat(value, lo, hi, base) {
  let r = (hi === lo) ? 0.5 : (value - lo) / (hi - lo);
  r = Math.max(0, Math.min(1, r));
  const ch = i => Math.round(255 - (255 - base[i]) * r).toString(16).padStart(2, '0');
  return '#' + ch(0) + ch(1) + ch(2);
}
// Índices con valor numérico dentro de [lo,hi]
function validIdx(arr, lo, hi) {
  const idx = [];
  arr.forEach((x, i) => { if (isFinite(x) && x >= lo && x <= hi) idx.push(i); });
  return idx;
}

function doPost(e) {
  if (!e || !e.parameter || e.parameter.token !== TOKEN)
    return ContentService.createTextOutput('bad token');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const base = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH-mm-ss');
  let name = base, n = 1;
  while (ss.getSheetByName(name)) name = base + '_' + (n++);
  const sheet = ss.insertSheet(name, 0);

  const head = ['Modulo'];
  for (let i = 1; i <= 11; i++) head.push('V' + i);
  for (let i = 1; i <= 9;  i++) head.push('T' + i);
  head.push('Delta V', 'Delta T');

  const data = [head];
  const bg   = [head.map(() => '#ffffff')];
  const borders = [];
  let gVmin = Infinity, gVmax = -Infinity, gTmin = Infinity, gTmax = -Infinity;

  (e.parameter.data || '').split(/\r?\n/).forEach(line => {
    line = line.trim();
    if (!line || line.indexOf('Modulo') === 0) return;
    const c = line.split(';');
    const mod = parseInt(String(c[0]).replace(/\D/g, ''), 10);
    const vals = c.slice(1).map(Number);
    const V = vals.slice(0, 11), T = vals.slice(11, 20);

    const vIdx = validIdx(V, V_MIN, V_MAX);
    const tIdx = validIdx(T, T_MIN, T_MAX);
    const vV = vIdx.map(i => V[i]), tV = tIdx.map(i => T[i]);
    const vmin = vV.length ? Math.min(...vV) : NaN, vmax = vV.length ? Math.max(...vV) : NaN;
    const tmin = tV.length ? Math.min(...tV) : NaN, tmax = tV.length ? Math.max(...tV) : NaN;
    if (vV.length) { gVmin = Math.min(gVmin, vmin); gVmax = Math.max(gVmax, vmax); }
    if (tV.length) { gTmin = Math.min(gTmin, tmin); gTmax = Math.max(gTmax, tmax); }

    const dV = vV.length ? Math.round((vmax - vmin) * 1000) / 1000 : '';
    const dT = tV.length ? Math.round((tmax - tmin) * 10) / 10 : '';
    // Celdas fuera de rango -> palabra "ERROR" en vez del número (como el Excel).
    const Vd = V.map((x, i) => vIdx.indexOf(i) >= 0 ? x : 'ERROR');
    const Td = T.map((x, i) => tIdx.indexOf(i) >= 0 ? x : 'ERROR');
    data.push([mod].concat(Vd, Td, [dV, dT]));

    const row = new Array(head.length).fill('#ffffff');
    for (let i = 0; i < 11; i++)
      row[1 + i]  = (vIdx.indexOf(i) >= 0) ? heat(V[i], vmin, vmax, GREEN) : GRAY;
    for (let i = 0; i < 9; i++)
      row[12 + i] = (tIdx.indexOf(i) >= 0) ? heat(T[i], tmin, tmax, RED)   : GRAY;
    bg.push(row);

    // Rodear máx/mín SOLO entre los válidos
    const r = data.length;
    if (vV.length) {
      let iMax = vIdx[0], iMin = vIdx[0];
      vIdx.forEach(i => { if (V[i] > V[iMax]) iMax = i; if (V[i] < V[iMin]) iMin = i; });
      borders.push({ r, c: 2 + iMax }, { r, c: 2 + iMin });
    }
    if (tV.length) {
      let iMax = tIdx[0], iMin = tIdx[0];
      tIdx.forEach(i => { if (T[i] > T[iMax]) iMax = i; if (T[i] < T[iMin]) iMin = i; });
      borders.push({ r, c: 13 + iMax }, { r, c: 13 + iMin });
    }
  });

  const totalRow = new Array(head.length).fill('');
  totalRow[0] = 'Total';
  totalRow[head.length - 2] = isFinite(gVmax) ? Math.round((gVmax - gVmin) * 1000) / 1000 : '';
  totalRow[head.length - 1] = isFinite(gTmax) ? Math.round((gTmax - gTmin) * 10) / 10 : '';
  data.push(totalRow);
  bg.push(head.map(() => '#ffffff'));

  sheet.getRange(1, 1, data.length, head.length).setValues(data);
  sheet.getRange(1, 1, bg.length, head.length).setBackgrounds(bg);
  sheet.getRange(1, 1, 1, head.length).setFontWeight('bold');

  borders.forEach(b =>
    sheet.getRange(b.r, b.c).setBorder(
      true, true, true, true, false, false,
      '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK));

  return ContentService.createTextOutput('ok: ' + name);
}
