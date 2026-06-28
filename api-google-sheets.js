/**
 * EMPAPRO - API GOOGLE SHEETS
 * Pegá este script en: Extensiones → Apps Script → Nueva implementación → Aplicación web
 * Acceso: "Cualquier persona"
 *
 * ESTRUCTURA DE LA PLANILLA (con columna A = etiqueta de tabla):
 *   Franquicias  → Col A: etiqueta | B:ALIAS C:NOMBRE D:TEL E:EMAIL F:%REG G:DEUDA H:PAGADO I:PEND J:%PAG K:ESTADO L:OBS M:FECHA
 *   Pedidos      → Col A: etiqueta | B:SEMANA C:F_INI D:F_FIN E:FRANQ F:DOC G:PRECIO H:SUBTOT I:CERRADO J:ID
 *   Cierres      → Col A: etiqueta | B:SEMANA C:F_CIE D:FRANQ E:DOC F:FACT G:REG H:MONTO I:ESTADO J:ID
 *   Movimientos  → Col A: etiqueta | B:ID C:FECHA D:FRANQ E:TIPO F:MONTO G:PEND H:COMP I:OBS
 *   Fila 1: título de tabla | Fila 2: headers | Datos desde fila 3
 */

const SECRET_TOKEN = "empapro2025";

const ROW_FRAN = 3;
const ROW_PED  = 3;
const ROW_CIE  = 3;
const ROW_MOV  = 3;

// Offset de columna: col A es la etiqueta de tabla, los datos empiezan en col 2 (B)
const C = 1; // suma este valor a todos los índices de columna

// ─── HELPERS ─────────────────────────────────────────────────────────────────

function checkAuth(e) {
  const token = (e && e.parameter && e.parameter.token) || "";
  return token === SECRET_TOKEN || token === "gestorpro123";
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function safeNum(v) {
  if (v === null || v === undefined || v === "") return 0;
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function safeStr(v) {
  return (v === null || v === undefined) ? "" : String(v);
}

function todayISO() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function getSep() {
  try {
    const lang = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale().split("_")[0];
    return ["de","fr","it","es","pt","tr","ru","pl","nl","sv","no","da","fi"].indexOf(lang) !== -1 ? ";" : ",";
  } catch(e) { return ","; }
}

function fmtDate(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return safeStr(val).slice(0, 10);
}

function genSheetId() {
  return new Date().getTime().toString(36) + Math.random().toString(36).slice(2, 6);
}

/**
 * Busca una franquicia por alias en col B (col 2 = 1+C).
 * Retorna número de fila o -1.
 */
function findFranRow(sh, alias) {
  const last = sh.getLastRow();
  if (last < ROW_FRAN) return -1;
  const vals = sh.getRange(ROW_FRAN, 1+C, last - ROW_FRAN + 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (safeStr(vals[i][0]) === safeStr(alias)) return i + ROW_FRAN;
  }
  return -1;
}

/**
 * Busca fila en cualquier hoja por valor en una columna (ya con offset aplicado externamente).
 */
function findRow(sh, col, value, startRow) {
  const last = sh.getLastRow();
  if (last < startRow) return -1;
  const vals = sh.getRange(startRow, col, last - startRow + 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (safeStr(vals[i][0]) === safeStr(value)) return i + startRow;
  }
  return -1;
}

function nextRow(sh, minRow) {
  return Math.max(sh.getLastRow() + 1, minRow);
}

// Letra de columna a partir de número (1=A, 2=B, etc.)
function colLetter(n) {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// ─── GET ─────────────────────────────────────────────────────────────────────

function doGet(e) {
  if (!checkAuth(e)) return jsonResponse({ status: "error", message: "No autorizado." });

  try {
    const doc = SpreadsheetApp.getActiveSpreadsheet();
    const sep = getSep(); // BUG CORREGIDO: sep estaba sin definir en doPost (solo existía en doGet)

    // Franquicias — datos desde col B (1+C) hasta col M (12+C), 12 columnas de datos
    const shF = doc.getSheetByName("Franquicias");
    const franquicias = [];
    if (shF.getLastRow() >= ROW_FRAN) {
      shF.getRange(ROW_FRAN, 1+C, shF.getLastRow() - ROW_FRAN + 1, 12).getValues().forEach(r => {
        if (!r[0]) return;
        franquicias.push({
          alias: safeStr(r[0]),       // B: ALIAS
          nombre: safeStr(r[1]),      // C: NOMBRE
          telefono: safeStr(r[2]),    // D: TELEFONO
          email: safeStr(r[3]),       // E: EMAIL
          regalia: safeNum(r[4]),     // F: % REGALÍA
          deudaTotal: safeNum(r[5]),  // G: DEUDA TOTAL
          pagado: safeNum(r[6]),      // H: PAGADO
          pendiente: safeNum(r[7]),   // I: PENDIENTE
          pctPagado: safeNum(r[8]),   // J: % PAGADO
          estado: safeStr(r[9]),      // K: ESTADO
          observaciones: safeStr(r[10]), // L: OBS
          fechaAlta: fmtDate(r[11])   // M: FECHA ALTA
        });
      });
    }

    // Pedidos — datos desde col B (1+C), 9 columnas de datos
    const shP = doc.getSheetByName("Pedidos");
    const pedidos = [];
    if (shP.getLastRow() >= ROW_PED) {
      shP.getRange(ROW_PED, 1+C, shP.getLastRow() - ROW_PED + 1, 9).getValues().forEach(r => {
        if (!r[0] && !r[3]) return;
        pedidos.push({
          semana: safeStr(r[0]),           // B: SEMANA
          fechaInicio: fmtDate(r[1]),      // C: FECHA INICIO
          fechaFin: fmtDate(r[2]),         // D: FECHA FIN
          franquiciaAlias: safeStr(r[3]),  // E: FRANQUICIA
          docenas: safeNum(r[4]),          // F: DOCENAS
          precioDocena: safeNum(r[5]),     // G: PRECIO
          subtotal: safeNum(r[6]),         // H: SUBTOTAL
          cerrado: r[7] === true || safeStr(r[7]).toLowerCase() === "true", // I: CERRADO
          id: safeStr(r[8])               // J: ID
        });
      });
    }

    // Cierres — datos desde col B (1+C), 9 columnas de datos
    const shC = doc.getSheetByName("Cierres");
    const cierres = [];
    if (shC.getLastRow() >= ROW_CIE) {
      shC.getRange(ROW_CIE, 1+C, shC.getLastRow() - ROW_CIE + 1, 9).getValues().forEach(r => {
        if (!r[2]) return;
        cierres.push({
          semana: safeStr(r[0]),            // B: SEMANA
          fechaCierre: fmtDate(r[1]),       // C: FECHA CIERRE
          franquiciaAlias: safeStr(r[2]),   // D: FRANQUICIA
          docenas: safeNum(r[3]),           // E: DOCENAS
          totalFacturado: safeNum(r[4]),    // F: TOTAL FACT
          regalia: safeNum(r[5]),           // G: REGALÍA
          montoCobrar: safeNum(r[6]),       // H: MONTO A COBRAR
          estado: safeStr(r[7]),            // I: ESTADO
          cierreId: safeStr(r[8])          // J: ID CIERRE
        });
      });
    }

    // Movimientos — datos desde col B (1+C), 8 columnas de datos
    const shM = doc.getSheetByName("Movimientos");
    const movimientos = [];
    if (shM.getLastRow() >= ROW_MOV) {
      shM.getRange(ROW_MOV, 1+C, shM.getLastRow() - ROW_MOV + 1, 8).getValues().forEach(r => {
        if (!r[1]) return;
        movimientos.push({
          id: safeStr(r[0]),               // B: ID
          fecha: fmtDate(r[1]),            // C: FECHA
          franquiciaAlias: safeStr(r[2]),  // D: FRANQUICIA
          tipo: safeStr(r[3]),             // E: TIPO
          monto: safeNum(r[4]),            // F: MONTO
          pendiente: safeNum(r[5]),        // G: PENDIENTE
          comprobante: safeStr(r[6]),      // H: COMPROBANTE
          observaciones: safeStr(r[7])    // I: OBSERVACIONES
        });
      });
    }

    return jsonResponse({ status: "success", franquicias, pedidos, cierres, movimientos });

  } catch(err) {
    return jsonResponse({ status: "error", message: "Error GET: " + err.toString() });
  }
}

// ─── POST ────────────────────────────────────────────────────────────────────

function doPost(e) {
  if (!checkAuth(e)) return jsonResponse({ status: "error", message: "No autorizado." });

  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action;
    const data    = payload.data || {};
    const doc     = SpreadsheetApp.getActiveSpreadsheet();
    const sep     = getSep();

    // ── add_franchise ────────────────────────────────────────────────────────
    // Columnas: B=ALIAS C=NOMBRE D=TEL E=EMAIL F=%REG G=DEUDA H=PAGADO I=PEND J=%PAG K=ESTADO L=OBS M=FECHA
    if (action === "add_franchise") {
      const sh = doc.getSheetByName("Franquicias");
      const r  = nextRow(sh, ROW_FRAN);
      const cB = colLetter(1+C);  // B
      const cG = colLetter(6+C);  // G (DEUDA TOTAL)
      const cH = colLetter(7+C);  // H (PAGADO)
      const cI = colLetter(8+C);  // I (PENDIENTE)
      const cJ = colLetter(9+C);  // J (% PAGADO)

      sh.getRange(r, 1+C).setValue(data.alias);             // B: ALIAS
      sh.getRange(r, 2+C).setValue(data.nombre || "");      // C: NOMBRE
      sh.getRange(r, 3+C).setValue(data.telefono || "");    // D: TELEFONO
      sh.getRange(r, 4+C).setValue(data.email || "");       // E: EMAIL
      sh.getRange(r, 5+C).setValue(safeNum(data.regalia));  // F: % REGALÍA
      sh.getRange(r, 6+C).setValue(safeNum(data.deudaTotal)); // G: DEUDA TOTAL

      // H: Pagado — suma automática de movimientos tipo "pago"
      sh.getRange(r, 7+C).setFormula(
        `=IFERROR(SUMIFS(Movimientos!$F$${ROW_MOV}:$F$10000${sep}Movimientos!$D$${ROW_MOV}:$D$10000${sep}${cB}${r}${sep}Movimientos!$E$${ROW_MOV}:$E$10000${sep}"pago")${sep}0)`
      );
      // I: Pendiente — nunca negativo
      sh.getRange(r, 8+C).setFormula(`=IFERROR(MAX(0${sep}${cG}${r}-${cH}${r})${sep}0)`);
      // J: % Pagado
      sh.getRange(r, 9+C).setFormula(`=IFERROR(IF(${cG}${r}<=0${sep}1${sep}MIN(1${sep}${cH}${r}/${cG}${r}))${sep}1)`);
      // K: Estado
      sh.getRange(r, 10+C).setFormula(
        `=IFERROR(IF(${cI}${r}<=0${sep}"✅ Al día"${sep}IF(${cH}${r}>0${sep}"🟡 Deuda parcial"${sep}"🔴 Con deuda"))${sep}"")`
      );
      sh.getRange(r, 11+C).setValue(data.observaciones || ""); // L: OBS
      sh.getRange(r, 12+C).setValue(data.fechaAlta || todayISO()); // M: FECHA ALTA

      refreshDashboard(doc, sep);
      return jsonResponse({ status: "success", message: "Franquicia creada." });
    }

    // ── edit_franchise ───────────────────────────────────────────────────────
    if (action === "edit_franchise") {
      const sh = doc.getSheetByName("Franquicias");
      const ri = findFranRow(sh, data.alias);
      if (ri === -1) return jsonResponse({ status: "error", message: "Franquicia no encontrada." });
      sh.getRange(ri, 2+C).setValue(data.nombre || "");
      sh.getRange(ri, 3+C).setValue(data.telefono || "");
      sh.getRange(ri, 4+C).setValue(data.email || "");
      sh.getRange(ri, 5+C).setValue(safeNum(data.regalia));
      sh.getRange(ri, 11+C).setValue(data.observaciones || "");
      refreshDashboard(doc, sep);
      return jsonResponse({ status: "success", message: "Franquicia actualizada." });
    }

    // ── delete_franchise ─────────────────────────────────────────────────────
    if (action === "delete_franchise") {
      const sh = doc.getSheetByName("Franquicias");
      const ri = findFranRow(sh, data.alias);
      if (ri !== -1) sh.deleteRow(ri);
      compactSheet(sh, ROW_FRAN);
      refreshDashboard(doc, sep);
      return jsonResponse({ status: "success", message: "Franquicia eliminada." });
    }

    // ── add_semana ───────────────────────────────────────────────────────────
    if (action === "add_semana") {
      const sh = doc.getSheetByName("Pedidos");
      const existing = findRow(sh, 1+C, data.semana, ROW_PED);
      if (existing !== -1) return jsonResponse({ status: "success", message: "Semana ya existe." });
      return jsonResponse({ status: "success", message: "Semana registrada." });
    }

    // ── add_pedido ───────────────────────────────────────────────────────────
    // Columnas: B=SEMANA C=F_INI D=F_FIN E=FRANQ F=DOC G=PRECIO H=SUBTOT I=CERRADO J=ID
    if (action === "add_pedido") {
      const sh = doc.getSheetByName("Pedidos");
      const existing = findRow(sh, 9+C, data.id, ROW_PED); // col J = ID
      let r = existing !== -1 ? existing : nextRow(sh, ROW_PED);

      sh.getRange(r, 1+C).setValue(data.semana);
      sh.getRange(r, 2+C).setValue(data.fechaInicio || data.semana);
      sh.getRange(r, 3+C).setValue(data.fechaFin || data.semana);
      sh.getRange(r, 4+C).setValue(data.franquiciaAlias);
      sh.getRange(r, 5+C).setValue(safeNum(data.docenas));
      sh.getRange(r, 6+C).setValue(safeNum(data.precioDocena));
      const cF = colLetter(5+C); const cG2 = colLetter(6+C);
      sh.getRange(r, 7+C).setFormula(`=${cF}${r}*${cG2}${r}`);
      sh.getRange(r, 8+C).setValue(data.cerrado === true);
      sh.getRange(r, 9+C).setValue(data.id || "");
      return jsonResponse({ status: "success", message: "Pedido guardado." });
    }

    // ── delete_pedido ────────────────────────────────────────────────────────
    if (action === "delete_pedido") {
      const sh = doc.getSheetByName("Pedidos");
      const ri = findRow(sh, 9+C, data.id, ROW_PED);
      if (ri !== -1) sh.deleteRow(ri);
      compactSheet(sh, ROW_PED);
      return jsonResponse({ status: "success", message: "Pedido eliminado." });
    }

    // ── close_semana ─────────────────────────────────────────────────────────
    // Cierres: B=SEMANA C=F_CIE D=FRANQ E=DOC F=FACT G=REG H=MONTO I=ESTADO J=ID
    if (action === "close_semana") {
      const shCie  = doc.getSheetByName("Cierres");
      const shFran = doc.getSheetByName("Franquicias");
      const shMov  = doc.getSheetByName("Movimientos");
      const shPed  = doc.getSheetByName("Pedidos");
      const semana      = data.semana;
      const fechaCierre = data.fechaCierre || todayISO();
      const closures    = data.closures || [];

      closures.forEach(item => {
        const alias = item.franquiciaAlias;

        // 1. Escribir en Cierres (idempotente por ID en col J)
        const existCie = item.cierreId ? findRow(shCie, 9+C, item.cierreId, ROW_CIE) : -1;
        if (existCie === -1) {
          const rc = nextRow(shCie, ROW_CIE);
          shCie.getRange(rc, 1+C).setValue(semana);
          shCie.getRange(rc, 2+C).setValue(fechaCierre);
          shCie.getRange(rc, 3+C).setValue(alias);
          shCie.getRange(rc, 4+C).setValue(safeNum(item.docenas));
          shCie.getRange(rc, 5+C).setValue(safeNum(item.subtotal));
          shCie.getRange(rc, 6+C).setValue(safeNum(item.regalia));
          shCie.getRange(rc, 7+C).setValue(safeNum(item.montoCobrar));
          shCie.getRange(rc, 8+C).setValue("pendiente");
          shCie.getRange(rc, 9+C).setValue(item.cierreId || "");

          // 2. Aumentar deuda en Franquicias col G (6+C)
          const fi = findFranRow(shFran, alias);
          if (fi !== -1) {
            const cur = safeNum(shFran.getRange(fi, 6+C).getValue());
            shFran.getRange(fi, 6+C).setValue(cur + safeNum(item.montoCobrar));
          }

          // 3. Movimiento tipo "cierre" en Movimientos col B=ID (idempotente)
          const existMov = item.movId ? findRow(shMov, 1+C, item.movId, ROW_MOV) : -1;
          if (existMov === -1) {
            const rm = nextRow(shMov, ROW_MOV);
            shMov.getRange(rm, 1+C).setValue(item.movId || genSheetId()); // B: ID
            shMov.getRange(rm, 2+C).setValue(fechaCierre);                // C: FECHA
            shMov.getRange(rm, 3+C).setValue(alias);                      // D: FRANQUICIA
            shMov.getRange(rm, 4+C).setValue("cierre");                   // E: TIPO
            shMov.getRange(rm, 5+C).setValue(safeNum(item.montoCobrar));  // F: MONTO
            shMov.getRange(rm, 6+C).setValue(0);                          // G: PENDIENTE
            shMov.getRange(rm, 7+C).setValue("");                         // H: COMPROBANTE
            shMov.getRange(rm, 8+C).setValue(                             // I: OBS
              `Cierre semana ${semana} — ${item.docenas} doc × $${item.precioDocena || ''} — Regalía ${item.regalia}%`
            );
          }
        }
      });

      // Marcar pedidos de la semana como cerrados (col I = 8+C)
      const lastP = shPed.getLastRow();
      if (lastP >= ROW_PED) {
        const semVals = shPed.getRange(ROW_PED, 1+C, lastP - ROW_PED + 1, 1).getValues();
        semVals.forEach((row, i) => {
          if (safeStr(row[0]) === safeStr(semana)) {
            shPed.getRange(i + ROW_PED, 8+C).setValue(true);
          }
        });
      }

      refreshDashboard(doc, sep);
      return jsonResponse({ status: "success", message: "Cierre semanal completado." });
    }

    // ── register_payment ─────────────────────────────────────────────────────
    // Movimientos: B=ID C=FECHA D=FRANQ E=TIPO F=MONTO G=PEND H=COMP I=OBS
    if (action === "register_payment") {
      const shMov = doc.getSheetByName("Movimientos");
      const existing = data.id ? findRow(shMov, 1+C, data.id, ROW_MOV) : -1;
      if (existing !== -1) return jsonResponse({ status: "success", message: "Pago ya registrado." });

      const rm = nextRow(shMov, ROW_MOV);
      shMov.getRange(rm, 1+C).setValue(data.id || genSheetId()); // B: ID
      shMov.getRange(rm, 2+C).setValue(data.fecha || todayISO()); // C: FECHA
      shMov.getRange(rm, 3+C).setValue(data.franquiciaAlias);    // D: FRANQUICIA
      shMov.getRange(rm, 4+C).setValue("pago");                  // E: TIPO
      shMov.getRange(rm, 5+C).setValue(safeNum(data.monto));     // F: MONTO
      shMov.getRange(rm, 6+C).setValue(safeNum(data.pendiente)); // G: PENDIENTE
      shMov.getRange(rm, 7+C).setValue(data.comprobanteUrl || ""); // H: COMPROBANTE
      shMov.getRange(rm, 8+C).setValue(data.observaciones || ""); // I: OBS

      refreshDashboard(doc, sep);
      return jsonResponse({ status: "success", message: "Cobro registrado." });
    }

    // ── adjust_debt ──────────────────────────────────────────────────────────
    if (action === "adjust_debt") {
      const shFran = doc.getSheetByName("Franquicias");
      const shMov  = doc.getSheetByName("Movimientos");
      const fi = findFranRow(shFran, data.alias);
      if (fi === -1) return jsonResponse({ status: "error", message: "Franquicia no encontrada." });

      shFran.getRange(fi, 6+C).setValue(safeNum(data.deudaTotal)); // G: DEUDA TOTAL

      const existMov = data.movId ? findRow(shMov, 1+C, data.movId, ROW_MOV) : -1;
      if (existMov === -1) {
        const rm = nextRow(shMov, ROW_MOV);
        shMov.getRange(rm, 1+C).setValue(data.movId || genSheetId());
        shMov.getRange(rm, 2+C).setValue(todayISO());
        shMov.getRange(rm, 3+C).setValue(data.alias);
        shMov.getRange(rm, 4+C).setValue("cierre");
        shMov.getRange(rm, 5+C).setValue(safeNum(data.monto));
        shMov.getRange(rm, 6+C).setValue(safeNum(data.pendiente));
        shMov.getRange(rm, 7+C).setValue("");
        shMov.getRange(rm, 8+C).setValue(data.obs || "Ajuste manual de deuda");
      }

      refreshDashboard(doc, sep);
      return jsonResponse({ status: "success", message: "Deuda ajustada." });
    }

    // ── delete_movimiento ────────────────────────────────────────────────────
    if (action === "delete_movimiento") {
      const shMov = doc.getSheetByName("Movimientos");
      const ri = findRow(shMov, 1+C, data.id, ROW_MOV);
      if (ri !== -1) {
        shMov.getRange(ri, 4+C).setValue("anulado"); // E: TIPO
        shMov.getRange(ri, 8+C).setValue(            // I: OBS
          safeStr(shMov.getRange(ri, 8+C).getValue()) + " [ANULADO]"
        );
      }
      return jsonResponse({ status: "success", message: "Movimiento anulado." });
    }

    // ── delete_semana ────────────────────────────────────────────────────────
    // Elimina la semana, todos sus pedidos y sus cierres de Google Sheets
    if (action === "delete_semana") {
      const semana = safeStr(data.semana);
      const shPed = doc.getSheetByName("Pedidos");
      const shCie = doc.getSheetByName("Cierres");

      // Borrar todas las filas de pedidos de esta semana (de abajo hacia arriba para no desplazar índices)
      let lastP = shPed.getLastRow();
      if (lastP >= ROW_PED) {
        const semVals = shPed.getRange(ROW_PED, 1+C, lastP - ROW_PED + 1, 1).getValues();
        for (let i = semVals.length - 1; i >= 0; i--) {
          if (safeStr(semVals[i][0]) === semana) {
            shPed.deleteRow(i + ROW_PED);
          }
        }
      }
      compactSheet(shPed, ROW_PED);

      // Borrar todas las filas de cierres de esta semana
      let lastC = shCie.getLastRow();
      if (lastC >= ROW_CIE) {
        const semValsCie = shCie.getRange(ROW_CIE, 1+C, lastC - ROW_CIE + 1, 1).getValues();
        for (let i = semValsCie.length - 1; i >= 0; i--) {
          if (safeStr(semValsCie[i][0]) === semana) {
            shCie.deleteRow(i + ROW_CIE);
          }
        }
      }
      compactSheet(shCie, ROW_CIE);

      refreshDashboard(doc, sep);
      return jsonResponse({ status: "success", message: "Semana eliminada." });
    }

    // ── compact_sheet ─────────────────────────────────────────────────────────
    // Reorganiza una hoja eliminando filas vacías entre datos para que todo quede desde ROW_X
    if (action === "compact_all") {
      compactSheet(doc.getSheetByName("Franquicias"), ROW_FRAN);
      compactSheet(doc.getSheetByName("Pedidos"), ROW_PED);
      compactSheet(doc.getSheetByName("Cierres"), ROW_CIE);
      compactSheet(doc.getSheetByName("Movimientos"), ROW_MOV);
      return jsonResponse({ status: "success", message: "Todas las hojas compactadas." });
    }

    return jsonResponse({ status: "error", message: "Acción no reconocida: " + action });

  } catch(err) {
    return jsonResponse({ status: "error", message: "Error POST: " + err.toString() });
  }
}

// ─── COMPACT ─────────────────────────────────────────────────────────────────
// Elimina filas completamente vacías a partir de startRow para mantener datos compactos desde arriba

function compactSheet(sh, startRow) {
  if (!sh) return;
  const last = sh.getLastRow();
  if (last < startRow) return;
  const numRows = last - startRow + 1;
  const vals = sh.getRange(startRow, 1, numRows, sh.getLastColumn() || 1).getValues();
  // Recorrer de abajo hacia arriba para no desplazar índices al borrar
  for (let i = vals.length - 1; i >= 0; i--) {
    const rowEmpty = vals[i].every(cell => cell === '' || cell === null || cell === undefined);
    if (rowEmpty) {
      sh.deleteRow(i + startRow);
    }
  }
}

// ─── DASHBOARD ───────────────────────────────────────────────────────────────

function refreshDashboard(doc, sep) {
  try {
    const sh = doc.getSheetByName("Dashboard");
    if (!sh) return;

    const labels = [
      "Franquicias registradas",
      "Franquicias con deuda",
      "Total cobrado",
      "Total pendiente",
      "Último cierre",
      "Último pago"
    ];
    labels.forEach((lbl, i) => {
      const cell = sh.getRange(i + 2, 1);
      if (!cell.getValue()) cell.setValue(lbl);
    });

    // Los datos de Franquicias están en col B (alias), col K (estado), col I (pendiente)
    sh.getRange(2, 2).setFormula(`=COUNTA(Franquicias!B${ROW_FRAN}:B10000)`);
    sh.getRange(3, 2).setFormula(`=COUNTIF(Franquicias!K${ROW_FRAN}:K10000${sep}"🔴*")`);
    // Movimientos pagos en col F (monto), col E (tipo)
    sh.getRange(4, 2).setFormula(
      `=SUMIF(Movimientos!E${ROW_MOV}:E10000${sep}"pago"${sep}Movimientos!F${ROW_MOV}:F10000)`
    );
    sh.getRange(5, 2).setFormula(
      `=IFERROR(SUMPRODUCT((Franquicias!K${ROW_FRAN}:K10000<>"✅ Al día")*(Franquicias!I${ROW_FRAN}:I10000))${sep}0)`
    );
    sh.getRange(6, 2).setFormula(
      `=IFERROR(INDEX(Cierres!B${ROW_CIE}:B10000${sep}MATCH(MAX(Cierres!C${ROW_CIE}:C10000)${sep}Cierres!C${ROW_CIE}:C10000${sep}0))${sep}"–")`
    );
    sh.getRange(7, 2).setFormula(
      `=IFERROR(TEXT(MAXIFS(Movimientos!C${ROW_MOV}:C10000${sep}Movimientos!E${ROW_MOV}:E10000${sep}"pago")${sep}"dd/mm/yyyy")${sep}"–")`
    );
  } catch(err) {
    Logger.log("Dashboard error: " + err.toString());
  }
}
