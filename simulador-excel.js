/* =========================================================
   simulador-excel.js
   Arma una planilla .xlsx con fórmulas a partir de los mismos
   datos del simulador (cursos.csv, electivas.csv y lo que el
   estudiante marcó). No usa librerías externas: escribe el XML
   de Excel y lo empaqueta en un .zip sin compresión.
   La planilla funciona en Excel, LibreOffice y Google Sheets.
   ========================================================= */

(function (global) {
  "use strict";

  /* ── CONFIGURACIÓN ─────────────────────────────────── */

  const AREAS_XL = {
    cbg:   "Conocimientos Básicos y Generales",
    tm:    "Teórico-Metodológico",
    se:    "Socioespacial",
    sa:    "Sistemas Ambientales",
    tig:   "Tecnologías de la Información Geográfica",
    libre: "Extensión u optativas"
  };
  const MINIMOS = { cbg: 60, tm: 45, se: 45, sa: 40, tig: 30 };
  const TOTAL_TFG = 220;
  const FILAS_OTRAS_VACIAS = 15;
  const LARGO_BARRA = 20;

  const HOJA_AV = "Mi avance";
  const HOJA_UC = "Unidades curriculares";
  const REF_UC = "'" + HOJA_UC + "'!";
  const REF_AV = "'" + HOJA_AV + "'!";

  // Filas fijas de la hoja "Mi avance"
  const FILA_AREA = { cbg: 5, tm: 6, se: 7, sa: 8, tig: 9, libre: 10 };
  const FILA_TOTAL = 12;
  const FILA_GEO = 13;
  const FILA_TFG = 15;
  const CELDA_TOTAL = REF_AV + "$B$" + FILA_TOTAL;

  // Índices de estilo (ver estilosXML)
  const S = {
    normal: 0, titulo: 1, nota: 2, encabezado: 3, texto: 4, numero: 5,
    input: 6, inputTexto: 7, pct: 8, barra: 9, totalTexto: 10, totalNum: 11,
    seccion: 12, estado: 13, banner: 14, gris: 15, totalPct: 16, totalBarra: 17
  };

  /* ── UTILIDADES ────────────────────────────────────── */

  function esc(t) {
    return String(t)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;")
      .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  function letra(n) { // 1 → A
    let s = "";
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  function numCol(l) { // A → 1
    let n = 0;
    for (const ch of l) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }

  function barra(frac) {
    const llenos = Math.round(Math.min(1, Math.max(0, frac)) * LARGO_BARRA);
    return "█".repeat(llenos) + "░".repeat(LARGO_BARRA - llenos);
  }

  function formulaBarra(celdaFrac) {
    const n = `ROUND(${celdaFrac}*${LARGO_BARRA},0)`;
    return `REPT("█",${n})&REPT("░",${LARGO_BARRA}-${n})`;
  }

  /* ── MODELO DE HOJA ────────────────────────────────── */

  class Hoja {
    constructor(nombre) {
      this.nombre = nombre;
      this.filas = {};
      this.alturas = {};
      this.anchos = [];
      this.merges = [];
      this.cf = [];
      this.dv = [];
    }
    // valor: string | number | null; f: fórmula sin "="
    set(ref, valor, estilo = 0, f = null) {
      const m = /^([A-Z]+)(\d+)$/.exec(ref);
      const col = numCol(m[1]), fila = Number(m[2]);
      (this.filas[fila] = this.filas[fila] || {})[col] = { v: valor, s: estilo, f };
    }
    merge(rango) { this.merges.push(rango); }
    alto(fila, pt) { this.alturas[fila] = pt; }

    xml() {
      const filasXML = Object.keys(this.filas).map(Number).sort((a, b) => a - b).map(r => {
        const celdas = this.filas[r];
        const cs = Object.keys(celdas).map(Number).sort((a, b) => a - b).map(c => {
          const { v, s, f } = celdas[c];
          const ref = letra(c) + r;
          if (f) {
            if (typeof v === "number") return `<c r="${ref}" s="${s}"><f>${esc(f)}</f><v>${v}</v></c>`;
            return `<c r="${ref}" s="${s}" t="str"><f>${esc(f)}</f><v>${esc(v ?? "")}</v></c>`;
          }
          if (v === null || v === undefined || v === "") return `<c r="${ref}" s="${s}"/>`;
          if (typeof v === "number") return `<c r="${ref}" s="${s}"><v>${v}</v></c>`;
          return `<c r="${ref}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
        }).join("");
        const alto = this.alturas[r] ? ` ht="${this.alturas[r]}" customHeight="1"` : "";
        return `<row r="${r}"${alto}>${cs}</row>`;
      }).join("");

      const cols = this.anchos.length
        ? "<cols>" + this.anchos.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("") + "</cols>"
        : "";
      const merges = this.merges.length
        ? `<mergeCells count="${this.merges.length}">` + this.merges.map(m => `<mergeCell ref="${m}"/>`).join("") + "</mergeCells>"
        : "";
      let prioridad = 1;
      const cf = this.cf.map(g =>
        `<conditionalFormatting sqref="${g.sqref}">` +
        g.reglas.map(rg => `<cfRule type="expression" dxfId="${rg.dxf}" priority="${prioridad++}"><formula>${esc(rg.formula)}</formula></cfRule>`).join("") +
        "</conditionalFormatting>"
      ).join("");
      const dv = this.dv.length
        ? `<dataValidations count="${this.dv.length}">` + this.dv.map(d =>
            `<dataValidation type="${d.tipo}"${d.operador ? ` operator="${d.operador}"` : ""} allowBlank="1" showErrorMessage="1"` +
            ` errorTitle="${esc(d.errorTitulo || "Valor no válido")}" error="${esc(d.error || "")}" sqref="${d.sqref}">` +
            `<formula1>${esc(d.formula1)}</formula1></dataValidation>`).join("") + "</dataValidations>"
        : "";

      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>` +
        `<sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>` +
        `<sheetFormatPr defaultRowHeight="15"/>` + cols +
        `<sheetData>${filasXML}</sheetData>` + merges + cf + dv +
        `<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>` +
        `<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>` +
        `</worksheet>`;
    }
  }

  /* ── ESTILOS ───────────────────────────────────────── */

  function estilosXML() {
    const font = (sz, extra = "", color = "FF222222") =>
      `<font>${extra}<sz val="${sz}"/><color rgb="${color}"/><name val="Arial"/><family val="2"/></font>`;
    const fonts = [
      font(10),                                  // 0 normal
      font(14, "<b/>", "FF1F4A87"),              // 1 título
      font(10, "<i/>", "FF666666"),              // 2 nota
      font(10, "<b/>", "FFFFFFFF"),              // 3 encabezado
      font(10, "<b/>"),                          // 4 negrita
      font(12, "<b/>", "FF1F4A87"),              // 5 sección
      font(10, "", "FF3B6CB7"),                  // 6 barra
      font(10, "<b/>", "FF3B6CB7")               // 7 barra total
    ];
    const fill = rgb => `<fill><patternFill patternType="solid"><fgColor rgb="${rgb}"/><bgColor indexed="64"/></patternFill></fill>`;
    const fills = [
      `<fill><patternFill patternType="none"/></fill>`,
      `<fill><patternFill patternType="gray125"/></fill>`,
      fill("FF3B6CB7"),   // 2 azul
      fill("FFFFF2CC"),   // 3 amarillo (celdas para completar)
      fill("FFF2F4F8"),   // 4 gris claro
      fill("FFEEF2FB")    // 5 celeste
    ];
    const borders = [
      `<border><left/><right/><top/><bottom/><diagonal/></border>`,
      `<border><left style="thin"><color rgb="FFD0D5DD"/></left><right style="thin"><color rgb="FFD0D5DD"/></right>` +
      `<top style="thin"><color rgb="FFD0D5DD"/></top><bottom style="thin"><color rgb="FFD0D5DD"/></bottom><diagonal/></border>`
    ];
    // [numFmt, font, fill, border, alineación]
    const xf = (nf, fo, fi, bo, al = "") =>
      `<xf numFmtId="${nf}" fontId="${fo}" fillId="${fi}" borderId="${bo}" xfId="0"` +
      `${nf ? ' applyNumberFormat="1"' : ""} applyFont="1" applyFill="1" applyBorder="1"${al ? ' applyAlignment="1"' : ""}>${al}</xf>`;
    const al = (h, wrap = false, v = "center") =>
      `<alignment${h ? ` horizontal="${h}"` : ""} vertical="${v}"${wrap ? ' wrapText="1"' : ""}/>`;
    const cellXfs = [
      xf(0, 0, 0, 0),                              // 0 normal
      xf(0, 1, 0, 0, al("left")),                  // 1 título
      xf(0, 2, 0, 0, al("left", true, "top")),     // 2 nota
      xf(0, 3, 2, 1, al("center", true)),          // 3 encabezado
      xf(0, 0, 0, 1, al("left", true)),            // 4 texto
      xf(0, 0, 0, 1, al("center")),                // 5 número
      xf(0, 4, 3, 1, al("center")),                // 6 input centrado
      xf(0, 0, 3, 1, al("left", true)),            // 7 input texto
      xf(9, 0, 0, 1, al("center")),                // 8 porcentaje
      xf(0, 6, 0, 1, al("left")),                  // 9 barra
      xf(0, 4, 4, 1, al("left", true)),            // 10 total texto
      xf(0, 4, 4, 1, al("center")),                // 11 total número
      xf(0, 5, 0, 0, al("left")),                  // 12 sección
      xf(0, 0, 0, 1, al("left", true)),            // 13 estado
      xf(0, 4, 5, 1, al("left", true)),            // 14 banner TFG
      xf(0, 0, 4, 1, al("center")),                // 15 gris vacío
      xf(9, 4, 4, 1, al("center")),                // 16 total %
      xf(0, 7, 4, 1, al("left"))                   // 17 total barra
    ];
    const dxf = (fuente, fondo, negrita = false) =>
      `<dxf><font>${negrita ? "<b/>" : ""}<color rgb="${fuente}"/></font>` +
      `<fill><patternFill patternType="solid"><bgColor rgb="${fondo}"/></patternFill></fill></dxf>`;
    const dxfs = [
      dxf("FF1C7A3D", "FFD9F3DF", true),  // 0 verde (podés cursarla / cumplido)
      dxf("FF666666", "FFE4E4E4"),        // 1 gris (aprobada)
      dxf("FFC0392B", "FFFDECE9"),        // 2 TM
      dxf("FF9A7212", "FFFDF3D9"),        // 3 SE
      dxf("FF2F7A47", "FFEAF7EE"),        // 4 SA
      dxf("FF6A4FA0", "FFEFEAF8"),        // 5 TIG
      dxf("FF3B6CB7", "FFEEF2FB"),        // 6 CBG
      dxf("FF777777", "FFF0F0F0")         // 7 libre
    ];
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
      `<fonts count="${fonts.length}">${fonts.join("")}</fonts>` +
      `<fills count="${fills.length}">${fills.join("")}</fills>` +
      `<borders count="${borders.length}">${borders.join("")}</borders>` +
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      `<cellXfs count="${cellXfs.length}">${cellXfs.join("")}</cellXfs>` +
      `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
      `<dxfs count="${dxfs.length}">${dxfs.join("")}</dxfs>` +
      `</styleSheet>`;
  }

  const DXF_AREA = { tm: 2, se: 3, sa: 4, tig: 5, cbg: 6, libre: 7 };

  function reglasArea(colLetra, primeraFila) {
    return Object.keys(DXF_AREA).map(a => ({
      dxf: DXF_AREA[a], formula: `$${colLetra}${primeraFila}="${AREAS_XL[a]}"`
    }));
  }

  /* ── ESTADO DE CADA UNIDAD CURRICULAR (fórmula + valor) ── */

  function nombresDe(ids, cursos) {
    return ids.map(id => (cursos.find(c => c.id === id) || {}).nombre || id).join(", ");
  }

  // Devuelve { f, v } con la fórmula de la columna Estado y su valor calculado
  function estadoCurso(curso, fila, ctx) {
    const { cursos, filaDe, aprobado, total, porArea } = ctx;
    const ok = "Podés cursarla";
    let f, v;

    const ids = (curso.prereq_cursos || "").split(";").map(s => s.trim()).filter(Boolean);
    const condCursos = ids.length ? `AND(${ids.map(id => `$E$${filaDe[id]}="Sí"`).join(",")})` : "TRUE";
    const cumpleCursos = ids.every(id => aprobado[id]);

    switch (curso.prereq_tipo) {
      case "curso": {
        const msj = "Falta aprobar: " + nombresDe(ids, cursos);
        f = `IF(${condCursos},"${ok}","${msj}")`;
        v = cumpleCursos ? ok : msj;
        break;
      }
      case "curso_recomendada": {
        const msj = "Podés cursarla (se sugiere aprobar antes " + nombresDe(ids, cursos) + ")";
        f = `IF(${condCursos},"${ok}","${msj}")`;
        v = cumpleCursos ? ok : msj;
        break;
      }
      case "creditos": {
        const n = Number(curso.prereq_creditos);
        f = `IF(${CELDA_TOTAL}>=${n},"${ok}","Te faltan "&(${n}-${CELDA_TOTAL})&" créditos")`;
        v = total >= n ? ok : `Te faltan ${n - total} créditos`;
        break;
      }
      case "creditos_areas": {
        const n = Number(curso.prereq_creditos);
        const conds = [`${CELDA_TOTAL}>=${n}`];
        let cumple = total >= n;
        [["min_cbg", "cbg"], ["min_tm", "tm"], ["min_se", "se"], ["min_sa", "sa"], ["min_tig", "tig"]].forEach(([campo, a]) => {
          if (Number(curso[campo])) {
            const fa = FILA_AREA[a];
            conds.push(`${REF_AV}$B$${fa}>=${REF_AV}$C$${fa}`);
            if ((porArea[a] || 0) < MINIMOS[a]) cumple = false;
          }
        });
        const msj = "Todavía no: revisá los mínimos en la hoja Mi avance";
        f = `IF(AND(${conds.join(",")}),"${ok}","${msj}")`;
        v = cumple ? ok : msj;
        break;
      }
      default:
        f = `"${ok}"`;
        v = ok;
    }
    return {
      f: `IF($E${fila}="Sí","✓ Aprobada",${f})`,
      v: aprobado[curso.id] ? "✓ Aprobada" : v
    };
  }

  function textoPreviaturas(curso, cursos) {
    const ids = (curso.prereq_cursos || "").split(";").map(s => s.trim()).filter(Boolean);
    switch (curso.prereq_tipo) {
      case "curso": return "Aprobar " + nombresDe(ids, cursos);
      case "curso_recomendada": return "Sugerida (no obligatoria): " + nombresDe(ids, cursos);
      case "creditos": return `${curso.prereq_creditos} créditos aprobados`;
      case "creditos_areas": return `${curso.prereq_creditos} créditos y los mínimos de cada área`;
      default: return "Sin previaturas";
    }
  }

  /* ── ARMADO DEL LIBRO ──────────────────────────────── */

  function generar({ cursos, electivas, estado, conDatos }) {
    const est = conDatos && estado ? estado : { aprobadas: {}, electivas: {}, otras: [] };
    const aprobado = {};
    cursos.forEach(c => { aprobado[c.id] = !!(est.aprobadas || {})[c.id]; });
    const electivaAprobada = id => !!(est.electivas || {})[id];
    const otras = (est.otras || []).filter(o => Number(o.creditos) > 0);

    /* Valores calculados (se guardan junto a las fórmulas para que la
       planilla se vea bien incluso en visores que no recalculan) */
    const porArea = { cbg: 0, tm: 0, se: 0, sa: 0, tig: 0, libre: 0 };
    let geo = 0;
    cursos.forEach(c => { if (aprobado[c.id]) { porArea[c.area] += Number(c.creditos); geo += Number(c.creditos); } });
    electivas.forEach(e => { if (electivaAprobada(e.id)) porArea[e.area] = (porArea[e.area] || 0) + Number(e.creditos); });
    otras.forEach(o => { porArea.cbg += Number(o.creditos); });
    const total = Object.values(porArea).reduce((a, b) => a + b, 0);

    /* ── Hoja "Unidades curriculares" ── */
    const uc = new Hoja(HOJA_UC);
    uc.anchos = [12, 50, 34, 10, 12, 13, 46, 36];
    uc.set("A1", "Unidades curriculares de la Licenciatura en Geografía", S.titulo);
    uc.merge("A1:H1");
    uc.set("A2", "Elegí \"Sí\" en la columna ¿Aprobada? (celdas amarillas) para cada unidad curricular que ya aprobaste. " +
      "Los créditos obtenidos, el estado de cada unidad curricular y la hoja Mi avance se actualizan solos.", S.nota);
    uc.merge("A2:H2");
    uc.alto(2, 30);

    // Sección 1: trayectoria por semestre
    let r = 4;
    uc.set("A" + r, "1. Trayectoria sugerida por semestre", S.seccion);
    r++;
    ["Semestre", "Unidad curricular", "Área", "Créditos", "¿Aprobada?", "Créditos obtenidos", "Estado", "Previaturas"]
      .forEach((h, i) => uc.set(letra(i + 1) + r, h, S.encabezado));
    uc.alto(r, 30);
    r++;

    const filaDe = {};
    const ordenCursos = [...cursos].sort((a, b) => Number(a.semestre) - Number(b.semestre));
    ordenCursos.forEach((c, i) => { filaDe[c.id] = r + i; });
    const ctx = { cursos, filaDe, aprobado, total, porArea };
    const iniCursos = r;

    ordenCursos.forEach(c => {
      const cred = Number(c.creditos);
      uc.set("A" + r, Number(c.semestre), S.numero);
      uc.set("B" + r, c.nombre, S.texto);
      uc.set("C" + r, AREAS_XL[c.area] || c.area, S.texto);
      uc.set("D" + r, cred, S.numero);
      uc.set("E" + r, aprobado[c.id] ? "Sí" : null, S.input);
      uc.set("F" + r, aprobado[c.id] ? cred : 0, S.numero, `IF($E${r}="Sí",$D${r},0)`);
      const e = estadoCurso(c, r, ctx);
      uc.set("G" + r, e.v, S.estado, e.f);
      uc.set("H" + r, textoPreviaturas(c, cursos), S.texto);
      r++;
    });
    const finCursos = r - 1;

    // Sección 2: unidades curriculares sin semestre fijo
    r++;
    uc.set("A" + r, "2. Otras unidades curriculares (sin semestre fijo, se pueden intercalar)", S.seccion);
    r++;
    ["Semestre sugerido", "Unidad curricular", "Área", "Créditos", "¿Aprobada?", "Créditos obtenidos", "Dictada por", "Observaciones"]
      .forEach((h, i) => uc.set(letra(i + 1) + r, h, S.encabezado));
    uc.alto(r, 30);
    r++;
    const iniElect = r;
    const ordenElect = [...electivas].sort((a, b) => {
      const sa = Number(a.semestre_sugerido) || 99, sb = Number(b.semestre_sugerido) || 99;
      return sa - sb || a.nombre.localeCompare(b.nombre, "es");
    });
    ordenElect.forEach(e => {
      const cred = Number(e.creditos);
      const ap = electivaAprobada(e.id);
      const sug = Number(e.semestre_sugerido);
      uc.set("A" + r, sug || null, S.numero);
      uc.set("B" + r, e.nombre, S.texto);
      uc.set("C" + r, AREAS_XL[e.area] || e.area, S.texto);
      uc.set("D" + r, cred, S.numero);
      uc.set("E" + r, ap ? "Sí" : null, S.input);
      uc.set("F" + r, ap ? cred : 0, S.numero, `IF($E${r}="Sí",$D${r},0)`);
      uc.set("G" + r, e.propia === "si" ? "Departamento de Geografía" : e.centro, S.texto);
      uc.set("H" + r, e.area === "libre" ? "Suma al total, no a un área mínima" : null, S.texto);
      r++;
    });
    const finElect = r - 1;

    // Sección 3: otras que el estudiante agrega
    r++;
    uc.set("A" + r, "3. Otras unidades curriculares que cursaste (que no están en la lista)", S.seccion);
    r++;
    uc.set("A" + r, "Escribí el nombre, elegí el área y poné los créditos. Se suman solas al total y a su área.", S.nota);
    uc.merge(`A${r}:H${r}`);
    r++;
    ["", "Unidad curricular", "Área", "Créditos", "", "Créditos obtenidos", "", ""]
      .forEach((h, i) => uc.set(letra(i + 1) + r, h, S.encabezado));
    r++;
    const iniOtras = r;
    const filasOtras = otras.map(o => ({ nombre: o.nombre || "Sin nombre", creditos: Number(o.creditos) }))
      .concat(Array.from({ length: FILAS_OTRAS_VACIAS }, () => ({ nombre: null, creditos: null })));
    filasOtras.forEach(o => {
      uc.set("A" + r, null, S.gris);
      uc.set("B" + r, o.nombre, S.inputTexto);
      uc.set("C" + r, AREAS_XL.cbg, S.inputTexto);
      uc.set("D" + r, o.creditos, S.input);
      uc.set("E" + r, null, S.gris);
      uc.set("F" + r, o.creditos || 0, S.numero, `IF(ISNUMBER($D${r}),$D${r},0)`);
      uc.set("G" + r, null, S.gris);
      uc.set("H" + r, null, S.gris);
      r++;
    });
    const finOtras = r - 1;

    // Validaciones y formatos condicionales
    const rangoSiNo = `E${iniCursos}:E${finCursos} E${iniElect}:E${finElect}`;
    uc.dv.push({ tipo: "list", sqref: rangoSiNo, formula1: '"Sí,No"', error: "Elegí Sí o No." });
    uc.dv.push({
      tipo: "list", sqref: `C${iniOtras}:C${finOtras}`,
      formula1: '"' + Object.values(AREAS_XL).join(",") + '"', error: "Elegí un área de la lista."
    });
    uc.dv.push({
      tipo: "decimal", operador: "greaterThanOrEqual", sqref: `D${iniOtras}:D${finOtras}`,
      formula1: "0", error: "Los créditos tienen que ser un número."
    });
    uc.cf.push({ sqref: rangoSiNo, reglas: [{ dxf: 0, formula: `$E${iniCursos}="Sí"` }] });
    uc.cf.push({
      sqref: `G${iniCursos}:G${finCursos}`,
      reglas: [
        { dxf: 1, formula: `LEFT($G${iniCursos},1)="✓"` },
        { dxf: 0, formula: `LEFT($G${iniCursos},5)="Podés"` }
      ]
    });
    uc.cf.push({ sqref: `C${iniCursos}:C${finOtras}`, reglas: reglasArea("C", iniCursos) });

    /* ── Hoja "Mi avance" ── */
    const av = new Hoja(HOJA_AV);
    av.anchos = [44, 14, 14, 12, 11, 26, 14];
    av.set("A1", "Simulador de previaturas · Licenciatura en Geografía", S.titulo);
    av.merge("A1:G1");
    const hoy = new Date();
    const fecha = `${String(hoy.getDate()).padStart(2, "0")}/${String(hoy.getMonth() + 1).padStart(2, "0")}/${hoy.getFullYear()}`;
    av.set("A2", (conDatos
      ? `Planilla generada el ${fecha} con lo que marcaste en el simulador. `
      : `Planilla en blanco generada el ${fecha}. `) +
      "Esta hoja se calcula sola: marcá tus unidades curriculares aprobadas en la hoja Unidades curriculares y acá vas a ver tu avance.", S.nota);
    av.merge("A2:G2");
    av.alto(2, 30);

    ["Área", "Créditos acumulados", "Mínimo para el TFG", "Te faltan", "Avance", "Progreso", "Estado"]
      .forEach((h, i) => av.set(letra(i + 1) + 4, h, S.encabezado));
    av.alto(4, 30);

    const sumif = nombre => `SUMIF(${REF_UC}$C:$C,"${nombre}",${REF_UC}$F:$F)`;

    ["cbg", "tm", "se", "sa", "tig"].forEach(a => {
      const f = FILA_AREA[a];
      const tengo = porArea[a] || 0, min = MINIMOS[a];
      const frac = Math.min(1, tengo / min);
      av.set("A" + f, AREAS_XL[a], S.texto);
      av.set("B" + f, tengo, S.numero, sumif(AREAS_XL[a]));
      av.set("C" + f, min, S.numero);
      av.set("D" + f, Math.max(0, min - tengo), S.numero, `MAX(0,C${f}-B${f})`);
      av.set("E" + f, frac, S.pct, `IF(C${f}>0,MIN(1,B${f}/C${f}),0)`);
      av.set("F" + f, barra(frac), S.barra, formulaBarra("E" + f));
      av.set("G" + f, tengo >= min ? "✓ Cumplido" : "Pendiente", S.numero, `IF(B${f}>=C${f},"✓ Cumplido","Pendiente")`);
    });
    const fl = FILA_AREA.libre;
    av.set("A" + fl, AREAS_XL.libre + " (no tiene mínimo)", S.texto);
    av.set("B" + fl, porArea.libre || 0, S.numero, sumif(AREAS_XL.libre));
    ["C", "D", "E", "F", "G"].forEach(c => av.set(c + fl, null, S.gris));

    const ft = FILA_TOTAL;
    const fracT = Math.min(1, total / TOTAL_TFG);
    av.set("A" + ft, "Créditos totales acumulados", S.totalTexto);
    av.set("B" + ft, total, S.totalNum, `SUM(${REF_UC}$F:$F)`);
    av.set("C" + ft, TOTAL_TFG, S.totalNum);
    av.set("D" + ft, Math.max(0, TOTAL_TFG - total), S.totalNum, `MAX(0,C${ft}-B${ft})`);
    av.set("E" + ft, fracT, S.totalPct, `IF(C${ft}>0,MIN(1,B${ft}/C${ft}),0)`);
    av.set("F" + ft, barra(fracT), S.totalBarra, formulaBarra("E" + ft));
    av.set("G" + ft, total >= TOTAL_TFG ? "✓ Cumplido" : "Pendiente", S.totalNum, `IF(B${ft}>=C${ft},"✓ Cumplido","Pendiente")`);

    av.set("A" + FILA_GEO, "Créditos de la trayectoria por semestre", S.texto);
    av.set("B" + FILA_GEO, geo, S.numero, `SUM(${REF_UC}$F$${iniCursos}:$F$${finCursos})`);
    ["C", "D", "E", "F", "G"].forEach(c => av.set(c + FILA_GEO, null, S.gris));

    // Condición del Trabajo Final de Grado
    const filasMin = ["cbg", "tm", "se", "sa", "tig"].map(a => FILA_AREA[a]);
    const condTFG = [`B${ft}>=C${ft}`].concat(filasMin.map(f => `B${f}>=C${f}`)).join(",");
    const cumpleTFG = total >= TOTAL_TFG && ["cbg", "tm", "se", "sa", "tig"].every(a => (porArea[a] || 0) >= MINIMOS[a]);
    const siTFG = "✓ Ya cumplís los créditos para el Trabajo Final de Grado";
    const noTFG = "Todavía no cumplís los créditos para el Trabajo Final de Grado. Mirá la columna Te faltan.";
    av.set("A" + FILA_TFG, "Trabajo Final de Grado", S.totalTexto);
    av.set("B" + FILA_TFG, cumpleTFG ? siTFG : noTFG, S.banner, `IF(AND(${condTFG}),"${siTFG}","${noTFG}")`);
    ["C", "D", "E", "F", "G"].forEach(c => av.set(c + FILA_TFG, null, S.banner));
    av.merge(`B${FILA_TFG}:G${FILA_TFG}`);
    av.alto(FILA_TFG, 30);

    av.cf.push({
      sqref: `G5:G9 G${ft}`,
      reglas: [{ dxf: 0, formula: `LEFT(G5,1)="✓"` }]
    });
    av.cf.push({ sqref: `B${FILA_TFG}`, reglas: [{ dxf: 0, formula: `LEFT($B$${FILA_TFG},1)="✓"` }] });
    av.cf.push({ sqref: "A5:A10", reglas: reglasArea("A", 5) });

    // Guía de uso
    let g = FILA_TFG + 2;
    av.set("A" + g, "Cómo usar esta planilla", S.seccion);
    g++;
    [
      "1. Andá a la hoja Unidades curriculares y elegí \"Sí\" en la columna ¿Aprobada? de cada unidad curricular que aprobaste. Las celdas amarillas son las únicas que tenés que tocar.",
      "2. La sección 1 es la trayectoria sugerida por semestre. La columna Estado te dice si ya podés cursar cada unidad curricular o qué te falta.",
      "3. La sección 2 tiene las unidades curriculares sin semestre fijo, que podés ir intercalando. Algunas traen un semestre sugerido como orientación.",
      "4. Si cursaste algo que no está en la lista, agregalo en la sección 3 con su área y sus créditos.",
      "5. Esta hoja suma todo automáticamente y muestra cuánto te falta en cada área y en el total."
    ].forEach(t => { av.set("A" + g, t, S.nota); av.merge(`A${g}:G${g}`); av.alto(g, 28); g++; });
    g++;
    av.set("A" + g, "Los mínimos por área y los 220 créditos totales corresponden al plan de estudios vigente. " +
      "Esta planilla es una ayuda para planificar y no sustituye la información oficial de Bedelía.", S.nota);
    av.merge(`A${g}:G${g}`);
    av.alto(g, 28);

    /* ── Empaquetado ── */
    const enc = new TextEncoder();
    const archivos = [
      ["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
        `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
        `<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
        `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
        `</Types>`],
      ["_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
        `</Relationships>`],
      ["xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<bookViews><workbookView activeTab="0"/></bookViews>` +
        `<sheets><sheet name="${esc(HOJA_AV)}" sheetId="1" r:id="rId1"/><sheet name="${esc(HOJA_UC)}" sheetId="2" r:id="rId2"/></sheets>` +
        `<calcPr calcId="191029" fullCalcOnLoad="1"/>` +
        `</workbook>`],
      ["xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>` +
        `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>` +
        `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
        `</Relationships>`],
      ["xl/styles.xml", estilosXML()],
      ["xl/worksheets/sheet1.xml", av.xml()],
      ["xl/worksheets/sheet2.xml", uc.xml()]
    ].map(([nombre, contenido]) => ({ nombre, datos: enc.encode(contenido) }));

    return zipSinCompresion(archivos);
  }

  /* ── ZIP (método "store", sin compresión) ──────────── */

  const TABLA_CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(datos) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < datos.length; i++) c = TABLA_CRC[(c ^ datos[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  function zipSinCompresion(archivos) {
    const enc = new TextEncoder();
    const ahora = new Date();
    const hora = (ahora.getHours() << 11) | (ahora.getMinutes() << 5) | (ahora.getSeconds() >> 1);
    const dia = ((ahora.getFullYear() - 1980) << 9) | ((ahora.getMonth() + 1) << 5) | ahora.getDate();

    const locales = [], centrales = [];
    let offset = 0;
    archivos.forEach(a => {
      const nombre = enc.encode(a.nombre);
      const crc = crc32(a.datos);
      const tam = a.datos.length;

      const loc = new DataView(new ArrayBuffer(30));
      loc.setUint32(0, 0x04034b50, true);
      loc.setUint16(4, 20, true);
      loc.setUint16(6, 0x0800, true);
      loc.setUint16(8, 0, true);
      loc.setUint16(10, hora, true);
      loc.setUint16(12, dia, true);
      loc.setUint32(14, crc, true);
      loc.setUint32(18, tam, true);
      loc.setUint32(22, tam, true);
      loc.setUint16(26, nombre.length, true);
      loc.setUint16(28, 0, true);
      locales.push(new Uint8Array(loc.buffer), nombre, a.datos);

      const cen = new DataView(new ArrayBuffer(46));
      cen.setUint32(0, 0x02014b50, true);
      cen.setUint16(4, 20, true);
      cen.setUint16(6, 20, true);
      cen.setUint16(8, 0x0800, true);
      cen.setUint16(10, 0, true);
      cen.setUint16(12, hora, true);
      cen.setUint16(14, dia, true);
      cen.setUint32(16, crc, true);
      cen.setUint32(20, tam, true);
      cen.setUint32(24, tam, true);
      cen.setUint16(28, nombre.length, true);
      cen.setUint32(42, offset, true);
      centrales.push(new Uint8Array(cen.buffer), nombre);

      offset += 30 + nombre.length + tam;
    });

    const tamCentral = centrales.reduce((s, p) => s + p.length, 0);
    const fin = new DataView(new ArrayBuffer(22));
    fin.setUint32(0, 0x06054b50, true);
    fin.setUint16(8, archivos.length, true);
    fin.setUint16(10, archivos.length, true);
    fin.setUint32(12, tamCentral, true);
    fin.setUint32(16, offset, true);

    const partes = locales.concat(centrales, [new Uint8Array(fin.buffer)]);
    const salida = new Uint8Array(partes.reduce((s, p) => s + p.length, 0));
    let pos = 0;
    partes.forEach(p => { salida.set(p, pos); pos += p.length; });
    return salida;
  }

  /* ── DESCARGA EN EL NAVEGADOR ──────────────────────── */

  function descargar(opciones, nombreArchivo) {
    const bytes = generar(opciones);
    const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = nombreArchivo;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  const API = { generar, descargar };
  if (typeof module !== "undefined" && module.exports) module.exports = API;
  else global.SimuladorExcel = API;

})(typeof window !== "undefined" ? window : globalThis);
