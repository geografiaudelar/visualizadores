/* =========================================================
   simulador.js · Simulador de previaturas
   Departamento de Geografía · Facultad de Ciencias · UdelaR
   Lee cursos.csv, cruza con lo que el usuario marcó como
   aprobado (guardado en localStorage) y calcula qué se
   puede cursar en cada semestre.
   ========================================================= */

const CSV_PATH = "cursos.csv";
const ELECTIVAS_CSV_PATH = "electivas.csv";
const STORAGE_KEY = "geo-simulador-previaturas";

const NOMBRES_AREA = {
  tm:  "Teórico-Metodológico",
  se:  "Socioespacial",
  sa:  "Sistemas Ambientales",
  tig: "Tecnologías de la Información Geográfica"
};

const MINIMOS_AREA = { tm: 45, se: 45, sa: 40, tig: 30, cbg: 60 };

let CURSOS = [];
let ELECTIVAS = [];
let ESTADO = {};

/* ── CARGA DE DATOS ──────────────────────────────────── */

async function cargarCursos() {
  const resp = await fetch(CSV_PATH);
  if (!resp.ok) throw new Error("No se pudo leer cursos.csv");
  const texto = await resp.text();
  return parsearCSV(texto);
}

async function cargarElectivas() {
  const resp = await fetch(ELECTIVAS_CSV_PATH);
  if (!resp.ok) throw new Error("No se pudo leer electivas.csv");
  const texto = await resp.text();
  return parsearCSV(texto);
}

function parsearCSV(texto) {
  const filas = texto.trim().split(/\r?\n/);
  const encabezados = dividirLineaCSV(filas[0]);
  return filas.slice(1).map(linea => {
    const valores = dividirLineaCSV(linea);
    const obj = {};
    encabezados.forEach((h, i) => obj[h] = (valores[i] ?? "").trim());
    return obj;
  });
}

function dividirLineaCSV(linea) {
  const resultado = [];
  let actual = "";
  let entreComillas = false;
  for (let i = 0; i < linea.length; i++) {
    const c = linea[i];
    if (c === '"') { entreComillas = !entreComillas; continue; }
    if (c === "," && !entreComillas) { resultado.push(actual); actual = ""; continue; }
    actual += c;
  }
  resultado.push(actual);
  return resultado;
}

/* ── ESTADO (localStorage) ───────────────────────────── */

function cargarEstado() {
  try {
    const guardado = JSON.parse(localStorage.getItem(STORAGE_KEY));
    if (guardado && typeof guardado === "object") {
      // Migración: versiones anteriores guardaban un único número "cbg"
      if (!Array.isArray(guardado.otras)) {
        guardado.otras = guardado.cbg
          ? [{ id: "migrado", nombre: "Créditos varios", creditos: Number(guardado.cbg) }]
          : [];
      }
      if (!guardado.electivas || typeof guardado.electivas !== "object") {
        guardado.electivas = {};
      }
      delete guardado.cbg;
      return guardado;
    }
  } catch (e) { /* nada guardado o corrupto, seguimos con vacío */ }
  return { aprobadas: {}, otras: [], electivas: {} };
}

function guardarEstado() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(ESTADO));
}

/* ── CÁLCULOS ─────────────────────────────────────────── */

function creditosGeoAprobados() {
  return CURSOS
    .filter(c => ESTADO.aprobadas[c.id])
    .reduce((sum, c) => sum + Number(c.creditos), 0);
}

function creditosPorArea() {
  const totales = { tm: 0, se: 0, sa: 0, tig: 0 };
  CURSOS.filter(c => ESTADO.aprobadas[c.id]).forEach(c => {
    totales[c.area] = (totales[c.area] || 0) + Number(c.creditos);
  });
  ELECTIVAS.filter(e => ESTADO.electivas[e.id] && totales.hasOwnProperty(e.area)).forEach(e => {
    totales[e.area] = (totales[e.area] || 0) + Number(e.creditos);
  });
  return totales;
}

function creditosOtras() {
  return ESTADO.otras.reduce((sum, o) => sum + Number(o.creditos), 0);
}

// Créditos "Conocimientos Básicos y Generales": lo que el usuario cargó a mano + las
// electivas del catálogo marcadas que están etiquetadas como área "cbg"
function creditosCBG() {
  const deElectivas = ELECTIVAS
    .filter(e => ESTADO.electivas[e.id] && e.area === "cbg")
    .reduce((sum, e) => sum + Number(e.creditos), 0);
  return creditosOtras() + deElectivas;
}

// Créditos de electivas marcadas que no entran en ninguna de las 5 áreas oficiales
// (extensión, optativas libres, etc.) — cuentan para el total pero no para ningún mínimo
function creditosLibres() {
  return ELECTIVAS
    .filter(e => ESTADO.electivas[e.id] && e.area === "libre")
    .reduce((sum, e) => sum + Number(e.creditos), 0);
}

function creditosElectivasAprobadas() {
  return ELECTIVAS
    .filter(e => ESTADO.electivas[e.id])
    .reduce((sum, e) => sum + Number(e.creditos), 0);
}

function creditosTotales() {
  return creditosGeoAprobados() + creditosOtras() + creditosElectivasAprobadas();
}

// Devuelve { habilitada: bool, motivo: string } para un curso dado
function evaluarCurso(curso) {
  if (ESTADO.aprobadas[curso.id]) {
    return { habilitada: true, motivo: "aprobada" };
  }

  const faltantes = [];

  if (curso.prereq_tipo === "curso" || curso.prereq_tipo === "curso_recomendada") {
    const ids = curso.prereq_cursos.split(";").map(s => s.trim()).filter(Boolean);
    const sinAprobar = ids
      .filter(id => !ESTADO.aprobadas[id])
      .map(id => (CURSOS.find(c => c.id === id) || {}).nombre || id);
    if (sinAprobar.length) {
      const etiqueta = curso.prereq_tipo === "curso_recomendada" ? "Previatura sugerida" : "Falta aprobar";
      faltantes.push(`${etiqueta}: ${sinAprobar.join(", ")}`);
    }
  }

  if (curso.prereq_tipo === "creditos" || curso.prereq_tipo === "creditos_areas") {
    const req = Number(curso.prereq_creditos);
    const tengo = creditosTotales();
    if (tengo < req) {
      faltantes.push(`Te faltan ${req - tengo} créditos (tenés ${tengo} de ${req})`);
    }
  }

  if (curso.prereq_tipo === "creditos_areas") {
    const porArea = creditosPorArea();
    const chequeos = [
      ["min_tm", "tm"], ["min_se", "se"], ["min_sa", "sa"], ["min_tig", "tig"]
    ];
    chequeos.forEach(([campo, area]) => {
      const req = Number(curso[campo]);
      if (req && porArea[area] < req) {
        faltantes.push(`Te faltan ${req - porArea[area]} créditos en ${NOMBRES_AREA[area]}`);
      }
    });
    const reqCbg = Number(curso.min_cbg);
    if (reqCbg && creditosCBG() < reqCbg) {
      faltantes.push(`Te faltan ${reqCbg - creditosCBG()} créditos en Conocimientos Básicos y Generales`);
    }
  }

  // Las previaturas "recomendadas" no bloquean, solo avisan
  const bloqueantes = faltantes.filter(f => !f.startsWith("Previatura sugerida"));

  return { habilitada: bloqueantes.length === 0, motivo: faltantes.join(" · ") };
}

/* ── RENDER ───────────────────────────────────────────── */

function render() {
  renderSemestres();
  renderElectivas();
  renderResumen();
}

function renderSemestres() {
  const cont = document.getElementById("sim-main");
  cont.innerHTML = "";

  const semestres = [...new Set(CURSOS.map(c => Number(c.semestre)))].sort((a, b) => a - b);

  semestres.forEach(sem => {
    const cursosSemestre = CURSOS.filter(c => Number(c.semestre) === sem);

    const bloque = document.createElement("div");
    bloque.className = "semestre-block";

    const header = document.createElement("div");
    header.className = "semestre-header";
    const n = cursosSemestre.length;
    header.innerHTML = `<h3>Semestre ${sem}</h3><span>${n} ${n > 1 ? "unidades curriculares" : "unidad curricular"}</span>`;
    bloque.appendChild(header);

    const grid = document.createElement("div");
    grid.className = "materias-grid";

    cursosSemestre.forEach(curso => grid.appendChild(renderTarjetaMateria(curso)));

    bloque.appendChild(grid);

    const sugeridas = renderSugeridasSemestre(sem);
    if (sugeridas) bloque.appendChild(sugeridas);

    cont.appendChild(bloque);
  });
}

// Unidades curriculares sin semestre fijo que se sugieren para este semestre
function renderSugeridasSemestre(sem) {
  const items = ELECTIVAS.filter(e => Number(e.semestre_sugerido) === sem);
  if (!items.length) return null;

  const div = document.createElement("div");
  div.className = "semestre-sugeridas";
  const enlaces = items.map(e => {
    const clase = ESTADO.electivas[e.id] ? ' class="sug-ok"' : "";
    return `<a href="#electiva-${e.id}" data-id="${e.id}"${clase}>${e.nombre}</a>`;
  }).join(", ");
  div.innerHTML = `<strong>Para este semestre también te sugerimos:</strong> ${enlaces}`;

  div.querySelectorAll("a").forEach(a => a.addEventListener("click", ev => {
    ev.preventDefault();
    const card = document.getElementById("electiva-" + a.dataset.id);
    if (!card) return;
    card.scrollIntoView({ behavior: "smooth", block: "center" });
    card.classList.add("resaltada");
    setTimeout(() => card.classList.remove("resaltada"), 1600);
  }));
  return div;
}

function renderTarjetaMateria(curso) {
  const { habilitada, motivo } = evaluarCurso(curso);
  const aprobada = !!ESTADO.aprobadas[curso.id];

  const card = document.createElement("div");
  card.className = `materia-card area-${curso.area}` + (!aprobada && !habilitada ? " disabled" : "");

  const top = document.createElement("div");
  top.className = "materia-top";
  top.innerHTML = `
    <span class="materia-nombre">${curso.nombre}</span>
    <span class="materia-creditos">${curso.creditos} cr.</span>
  `;
  card.appendChild(top);

  let estadoHtml = "";
  if (aprobada) {
    estadoHtml = `<span class="materia-estado estado-aprobada">✓ Aprobada</span>`;
  } else if (habilitada) {
    estadoHtml = `<span class="materia-estado estado-disponible">Podés cursarla</span>`;
  } else {
    estadoHtml = `<span class="materia-estado estado-bloqueada">Todavía no</span>`;
  }
  const estadoDiv = document.createElement("div");
  estadoDiv.innerHTML = estadoHtml;
  card.appendChild(estadoDiv.firstChild);

  if (motivo && motivo !== "aprobada") {
    const req = document.createElement("div");
    req.className = "materia-req";
    req.textContent = motivo;
    card.appendChild(req);
  }

  const label = document.createElement("label");
  label.className = "materia-check";
  label.innerHTML = `<input type="checkbox" ${aprobada ? "checked" : ""}> Ya la aprobé`;
  label.querySelector("input").addEventListener("change", (e) => {
    ESTADO.aprobadas[curso.id] = e.target.checked;
    guardarEstado();
    render();
  });
  card.appendChild(label);

  return card;
}

function renderResumen() {
  document.getElementById("sum-geo").textContent = creditosGeoAprobados();
  document.getElementById("sum-total").textContent = creditosTotales();

  renderOtrasLista();

  const porArea = creditosPorArea();
  const cont = document.getElementById("area-breakdown");
  cont.innerHTML = "";
  Object.keys(NOMBRES_AREA).forEach(area => {
    const row = document.createElement("div");
    row.className = "area-row";
    row.innerHTML = `
      <span class="dot dot-${area}"></span>
      <span class="area-nombre">${NOMBRES_AREA[area]}</span>
      <span class="area-valor">${porArea[area] || 0} / ${MINIMOS_AREA[area]}</span>
    `;
    cont.appendChild(row);
  });
  const libres = creditosLibres();
  if (libres > 0) {
    const row = document.createElement("div");
    row.className = "area-row";
    row.innerHTML = `
      <span class="dot" style="background:#bbb"></span>
      <span class="area-nombre">Otras áreas (extensión, optativas)</span>
      <span class="area-valor">${libres}</span>
    `;
    cont.appendChild(row);
  }

  const minList = document.getElementById("min-list");
  minList.innerHTML = "";
  const orden = [
    ["Conocimientos Básicos y Generales", MINIMOS_AREA.cbg, creditosCBG()],
    ["Teórico-Metodológico", MINIMOS_AREA.tm, porArea.tm || 0],
    ["Socioespacial", MINIMOS_AREA.se, porArea.se || 0],
    ["Sistemas Ambientales", MINIMOS_AREA.sa, porArea.sa || 0],
    ["Tecnologías de la Información Geográfica", MINIMOS_AREA.tig, porArea.tig || 0],
  ];
  orden.forEach(([nombre, min, tengo]) => {
    const li = document.createElement("li");
    li.innerHTML = `<span>${nombre}</span><strong>${tengo} / ${min}</strong>`;
    minList.appendChild(li);
  });
}

function renderOtrasLista() {
  const cont = document.getElementById("otras-lista");
  cont.innerHTML = "";

  if (!ESTADO.otras.length) {
    cont.innerHTML = `<p class="otras-vacio">Todavía no agregaste ninguna.</p>`;
    return;
  }

  ESTADO.otras.forEach(o => {
    const item = document.createElement("div");
    item.className = "otra-item";
    item.innerHTML = `
      <span class="otra-nombre">${o.nombre || "Sin nombre"}</span>
      <span class="otra-creditos">${o.creditos} cr.</span>
      <button type="button" class="otra-quitar" title="Quitar">✕</button>
    `;
    item.querySelector(".otra-quitar").addEventListener("click", () => {
      ESTADO.otras = ESTADO.otras.filter(x => x.id !== o.id);
      guardarEstado();
      render();
    });
    cont.appendChild(item);
  });
}

const CHIP_AREA_LABEL = {
  tm: "Teórico-Metodológico", se: "Socioespacial", sa: "Sistemas Ambientales",
  tig: "Tec. Información Geográfica", cbg: "Con. Básicos y Generales", libre: "Extensión u optativa"
};

function renderElectivas() {
  const cont = document.getElementById("sim-electivas");
  cont.innerHTML = "";

  const grupos = [
    { key: "si", titulo: "Dictadas por Geografía", badge: "badge-propia" },
    { key: "no", titulo: "Dictadas por otros institutos", badge: "badge-otro" },
  ];

  grupos.forEach(g => {
    const items = ELECTIVAS.filter(e => e.propia === g.key).sort((a, b) => {
      const sa = Number(a.semestre_sugerido) || 99, sb = Number(b.semestre_sugerido) || 99;
      return sa - sb || a.nombre.localeCompare(b.nombre, "es");
    });
    if (!items.length) return;

    const grupo = document.createElement("div");
    grupo.className = "electivas-grupo";

    const header = document.createElement("div");
    header.className = "electivas-grupo-header";
    header.innerHTML = `<h3>${g.titulo}</h3><span class="${g.badge}">${items.length}</span>`;
    grupo.appendChild(header);

    const lista = document.createElement("div");
    lista.className = "electivas-lista";
    items.forEach(e => lista.appendChild(renderTarjetaElectiva(e)));
    grupo.appendChild(lista);

    cont.appendChild(grupo);
  });
}

function renderTarjetaElectiva(e) {
  const marcada = !!ESTADO.electivas[e.id];
  const card = document.createElement("div");
  card.className = "electiva-card" + (marcada ? " marcada" : "");
  card.id = "electiva-" + e.id;
  const sug = Number(e.semestre_sugerido);
  const chipSugerida = sug ? `<span class="electiva-sugerida">Sugerida en semestre ${sug}</span>` : "";

  card.innerHTML = `
    <div class="electiva-top">
      <span class="electiva-nombre">${e.nombre}</span>
      <span class="electiva-creditos">${e.creditos} cr.</span>
    </div>
    <div class="electiva-chips"><span class="electiva-area-chip chip-${e.area}">${CHIP_AREA_LABEL[e.area] || e.area}</span>${chipSugerida}</div>
    <span class="electiva-meta">${e.centro}</span>
    <label class="materia-check">
      <input type="checkbox" ${marcada ? "checked" : ""}> Ya la aprobé
    </label>
  `;
  card.querySelector("input").addEventListener("change", (ev) => {
    ESTADO.electivas[e.id] = ev.target.checked;
    guardarEstado();
    render();
  });
  return card;
}

/* ── INICIO ───────────────────────────────────────────── */

async function iniciar() {
  ESTADO = cargarEstado();

  try {
    CURSOS = await cargarCursos();
    ELECTIVAS = await cargarElectivas();
  } catch (e) {
    document.getElementById("sim-main").innerHTML =
      `<div class="summary-card">No se pudo leer <strong>cursos.csv</strong> o <strong>electivas.csv</strong>. Verificá que
       ambos archivos estén en la misma carpeta que simulador.html y que la página se esté abriendo desde un servidor web
       (por ejemplo, GitHub Pages), no con doble clic.</div>`;
    return;
  }

  const btnAgregar = document.getElementById("btn-agregar-otra");
  btnAgregar.addEventListener("click", agregarOtra);
  document.getElementById("otra-creditos").addEventListener("keydown", (e) => {
    if (e.key === "Enter") agregarOtra();
  });
  document.getElementById("otra-nombre").addEventListener("keydown", (e) => {
    if (e.key === "Enter") agregarOtra();
  });

  document.getElementById("btn-excel-datos").addEventListener("click", () => descargarExcel(true));
  document.getElementById("btn-excel-blanco").addEventListener("click", () => descargarExcel(false));

  document.getElementById("btn-reset").addEventListener("click", () => {
    if (!confirm("¿Reiniciar la simulación? Se van a borrar todas las unidades curriculares marcadas como aprobadas y los créditos agregados.")) return;
    ESTADO = { aprobadas: {}, otras: [], electivas: {} };
    guardarEstado();
    render();
  });

  render();
}

function descargarExcel(conDatos) {
  if (!window.SimuladorExcel) {
    alert("No se pudo cargar el generador de planillas (simulador-excel.js).");
    return;
  }
  const nombre = conDatos
    ? "simulador_previaturas_geografia_mi_avance.xlsx"
    : "simulador_previaturas_geografia_en_blanco.xlsx";
  window.SimuladorExcel.descargar(
    { cursos: CURSOS, electivas: ELECTIVAS, estado: ESTADO, conDatos },
    nombre
  );
}

function agregarOtra() {
  const nombreInput = document.getElementById("otra-nombre");
  const creditosInput = document.getElementById("otra-creditos");
  const creditos = Number(creditosInput.value);

  if (!creditos || creditos <= 0) {
    creditosInput.focus();
    return;
  }

  ESTADO.otras.push({
    id: "o" + Date.now() + Math.random().toString(36).slice(2, 6),
    nombre: nombreInput.value.trim(),
    creditos
  });
  guardarEstado();
  nombreInput.value = "";
  creditosInput.value = "";
  nombreInput.focus();
  render();
}

document.addEventListener("DOMContentLoaded", iniciar);
