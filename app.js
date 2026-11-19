/*
 * VEPAGAS Asistencia — versión web (iPhone y cualquier navegador).
 * Misma lógica que la APK (RepositorioAsistencia.kt / GraphSharePoint.kt / FormatoFecha.kt):
 * inicio de sesión Microsoft (MSAL, delegado, sin secretos) y Microsoft Graph directo contra las
 * listas SharePoint, con los mismos formatos que Power Apps. Sin servidor propio.
 */
"use strict";

const C = window.VEPAGAS_CONFIG;
const GRAPH = `https://graph.microsoft.com/v1.0/sites/${C.siteId}/lists`;
const SCOPES = ["https://graph.microsoft.com/Sites.ReadWrite.All", "https://graph.microsoft.com/User.Read"];
const MARCA_MIGRADO = "APK"; // dispositivos_db.Title: la fila ya no es de Power Apps (misma marca que la APK)
const REDIRECT = location.origin + location.pathname.replace(/index\.html$/, "");

// =====================================================================
// Formatos (idénticos a Power Apps). Ecuador: UTC-5 todo el año.
// =====================================================================
const DESFASE_EC = 5 * 3600 * 1000;
const MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
const MESES_EN = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];
const DIAS = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
const dos = (n) => String(n).padStart(2, "0");

const F = {
  partes(fecha) {
    const x = new Date(fecha.getTime() - DESFASE_EC);
    return { a: x.getUTCFullYear(), m: x.getUTCMonth() + 1, d: x.getUTCDate(), h: x.getUTCHours(), mi: x.getUTCMinutes(), s: x.getUTCSeconds(), sem: x.getUTCDay() };
  },
  desdeLocal(a, m, d, h, mi, s = 0) {
    if (m < 1 || m > 12 || d < 1 || d > 31 || h > 23 || mi > 59 || s > 59) return null;
    return new Date(Date.UTC(a, m - 1, d, h, mi, s) + DESFASE_EC);
  },
  /** FECHA_HORA_INICIO / SALIDA: "18 noviembre 2026 9:34" */
  texto(f) { const p = F.partes(f); return `${p.d} ${MESES[p.m - 1]} ${p.a} ${p.h}:${dos(p.mi)}`; },
  /** HORA: "09:34" */
  hora(f) { const p = F.partes(f); return `${dos(p.h)}:${dos(p.mi)}`; },
  dia(f) { const p = F.partes(f); return `${p.a}-${dos(p.m)}-${dos(p.d)}`; },
  /** sesion_db.FECHA: inicio del día de Ecuador en UTC. */
  diaSesion(f) { return `${F.dia(f)}T05:00:00Z`; },
  diaLegible(f) { const p = F.partes(f); return `${DIAS[p.sem]}, ${p.d} de ${MESES[p.m - 1]}`; },
  /** LAT / LONG: "-0,18070000" */
  coordenada(v) { return v.toFixed(8).replace(".", ","); },
  /** Enlace de Google Maps: con punto */
  coordenadaMapa(v) { return v.toFixed(8); },
  /** UBICACION_APP: "-0,1807 -78,4677983" */
  coordenadaUbicacionApp(v) { return v == null ? "" : String(Number(v.toFixed(8))).replace(".", ","); },

  /** Lee todos los formatos que existen en las listas (Power Apps en español, en inglés, filas viejas). */
  parsear(texto) {
    const t = (texto || "").replace(/[  ]/g, " ").replace(/\s+/g, " ").trim();
    if (!t) return null;
    let m;
    if (/^\d{4}-\d{2}-\d{2}T.*(Z|[+-]\d{2}:?\d{2})$/i.test(t)) { const v = Date.parse(t); if (!isNaN(v)) return new Date(v); }
    if ((m = t.match(/^(\d{1,2}) (\p{L}+) (\d{4}) (\d{1,2}):(\d{2})(?::(\d{2}))?$/iu))) {
      const mes = MESES.indexOf(m[2].toLowerCase()) + 1 || (m[2].toLowerCase() === "setiembre" ? 9 : 0);
      if (mes) return F.desdeLocal(+m[3], mes, +m[1], +m[4], +m[5], +(m[6] || 0));
    }
    if ((m = t.match(/^(\p{L}+) (\d{1,2}), (\d{4}) (\d{1,2}):(\d{2})(?::(\d{2}))? ?([ap])\.? ?m\.?$/iu))) {
      const mes = MESES_EN.indexOf(m[1].toLowerCase()) + 1;
      if (mes) return F.desdeLocal(+m[3], mes, +m[2], F.h24(+m[4], m[7]), +m[5], +(m[6] || 0));
    }
    if ((m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4}) (\d{1,2}):(\d{2})(?::(\d{2}))?$/))) {
      return F.desdeLocal(+m[3], +m[2], +m[1], +m[4], +m[5], +(m[6] || 0));
    }
    if ((m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4}) (\d{1,2}):(\d{2}) ?([ap])\.? ?m\.?$/i))) {
      const h = F.h24(+m[4], m[6]);
      return F.desdeLocal(+m[3], +m[2], +m[1], h, +m[5]) || F.desdeLocal(+m[3], +m[1], +m[2], h, +m[5]);
    }
    if ((m = t.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?$/))) {
      return F.desdeLocal(+m[1], +m[2], +m[3], +m[4], +m[5], +(m[6] || 0));
    }
    return null;
  },
  h24(h, ap) { const pm = ap.toLowerCase() === "p"; return (h % 12) + (pm ? 12 : 0); },
};

// =====================================================================
// Hora oficial: cabecera Date de Microsoft o del servidor de la página (no el reloj del teléfono).
// =====================================================================
let desfaseServidor = null;
function tomarHora(resp) {
  const d = resp.headers.get("Date");
  const v = d ? Date.parse(d) : NaN;
  if (!isNaN(v)) desfaseServidor = v - Date.now();
}
async function sincronizarHora() {
  try { tomarHora(await fetch(REDIRECT + "?t=" + Date.now(), { method: "HEAD", cache: "no-store" })); } catch { /* sin red */ }
}
const ahora = () => new Date(Date.now() + (desfaseServidor ?? 0));

// =====================================================================
// Almacenamiento local (identidad del dispositivo y operaciones pendientes)
// =====================================================================
const Local = {
  leer(k) { try { return localStorage.getItem(k); } catch { return null; } },
  guardar(k, v) { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, v); } catch { /* modo privado */ } },
};
function idDispositivo() {
  let id = Local.leer("vepagas_dispositivo");
  if (!id) {
    const uuid = crypto.randomUUID ? crypto.randomUUID()
      : "10000000-1000-4000-8000-100000000000".replace(/[018]/g, (c) => (c ^ (crypto.getRandomValues(new Uint8Array(1))[0] & (15 >> (c / 4)))).toString(16));
    id = "web-" + uuid;
    Local.guardar("vepagas_dispositivo", id);
  }
  return id;
}
const operaciones = () => { try { return JSON.parse(Local.leer("vepagas_operaciones") || "[]"); } catch { return []; } };
const escribirOperaciones = (l) => Local.guardar("vepagas_operaciones", l.length ? JSON.stringify(l) : null);
const guardarOperacion = (op) => escribirOperaciones(operaciones().filter((o) => !(o.idReg === op.idReg && o.accion === op.accion)).concat(op));
const quitarOperacion = (op) => escribirOperaciones(operaciones().filter((o) => !(o.idReg === op.idReg && o.accion === op.accion)));

// =====================================================================
// Inicio de sesión Microsoft (MSAL.js)
// =====================================================================
class SinSesion extends Error {}
class Rechazo extends Error { constructor(codigo, mensaje) { super(mensaje); this.codigo = codigo; } }
class ErrorGraph extends Error { constructor(mensaje, estado) { super(mensaje); this.estado = estado; } }

const msalApp = new msal.PublicClientApplication({
  auth: { clientId: C.clientId, authority: `https://login.microsoftonline.com/${C.tenantId}`, redirectUri: REDIRECT, navigateToLoginRequestUrl: false },
  cache: { cacheLocation: "localStorage" },
});

function cuentaActual() {
  const a = msalApp.getActiveAccount() || msalApp.getAllAccounts()[0];
  if (!a) return null;
  const c = a.idTokenClaims || {};
  const correo = c.email || c.preferred_username || a.username || "";
  return { cuenta: a, oid: c.oid || a.localAccountId, nombre: c.name || a.name || correo, correo };
}

async function token() {
  const c = cuentaActual();
  if (!c) throw new SinSesion();
  try {
    return (await msalApp.acquireTokenSilent({ scopes: SCOPES, account: c.cuenta })).accessToken;
  } catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError) throw new SinSesion();
    throw e;
  }
}

// =====================================================================
// Microsoft Graph -> listas SharePoint
// =====================================================================
const esperar = (ms) => new Promise((r) => setTimeout(r, ms));

/** POST no se reintenta (podría duplicar); GET/PATCH sí ante 429/503/504 y fallos de red. */
async function enviar(metodo, url, cuerpo) {
  const reintentable = metodo !== "POST";
  for (let intento = 1; ; intento++) {
    const t = await token();
    let r;
    try {
      r = await fetch(url, {
        method: metodo,
        headers: { Authorization: "Bearer " + t, Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly", ...(cuerpo ? { "Content-Type": "application/json" } : {}) },
        body: cuerpo ? JSON.stringify(cuerpo) : undefined,
      });
    } catch {
      if (reintentable && intento < 3) { await esperar(800 * intento); continue; }
      throw new ErrorGraph("Sin conexión con Microsoft", null);
    }
    tomarHora(r);
    if (r.ok) { const txt = await r.text(); return txt ? JSON.parse(txt) : {}; }
    if (reintentable && [429, 503, 504].includes(r.status) && intento < 3) {
      await esperar((+r.headers.get("Retry-After") || intento) * 1000); continue;
    }
    throw new ErrorGraph("SharePoint respondió " + r.status, r.status);
  }
}

async function listar(lista, campo, valor, select) {
  const filtro = encodeURIComponent(`fields/${campo} eq '${valor.replace(/'/g, "''")}'`);
  let url = `${GRAPH}/${lista}/items?$expand=fields($select=${select})&$filter=${filtro}&$top=200`;
  const todos = [];
  for (let p = 0; url && p < 25; p++) {
    const r = await enviar("GET", url);
    todos.push(...(r.value || []));
    url = r["@odata.nextLink"];
  }
  return todos;
}
const crear = async (lista, fields) => (await enviar("POST", `${GRAPH}/${lista}/items`, { fields })).id;
const actualizar = (lista, id, fields) => enviar("PATCH", `${GRAPH}/${lista}/items/${id}/fields`, fields);
const txt = (v) => (v == null ? "" : String(v));

const Graph = {
  async jornadasDe(oid) {
    const items = await listar(C.listas.registros, "ID_USUARIO", oid, "ID_USUARIO,FECHA_HORA_INICIO,FECHA_HORA_SALIDA,ID_REG,TIPO_TRABAJO");
    return items.map((it) => {
      const f = it.fields || {};
      if (txt(f.ID_USUARIO) !== oid) return null;
      const inicio = F.parsear(txt(f.FECHA_HORA_INICIO));
      if (!inicio) return null;
      return { itemId: it.id, idReg: txt(f.ID_REG), tipo: txt(f.TIPO_TRABAJO).trim(), inicio, salida: F.parsear(txt(f.FECHA_HORA_SALIDA)) };
    }).filter(Boolean);
  },
  crearJornada: (oid, idReg, tipo, inicio) => crear(C.listas.registros, { ID_USUARIO: oid, FECHA_HORA_INICIO: F.texto(inicio), ID_REG: idReg, TIPO_TRABAJO: tipo }),
  registrarSalida: (itemId, salida) => actualizar(C.listas.registros, itemId, { FECHA_HORA_SALIDA: F.texto(salida) }),
  crearUbicacion: (idReg, hora, lat, lon) => crear(C.listas.ubicaciones, {
    ID_FECHA: idReg, HORA: F.hora(hora), LAT: F.coordenada(lat), LONG: F.coordenada(lon),
    UBICACION: `https://www.google.com/maps/place/${F.coordenadaMapa(lat)},${F.coordenadaMapa(lon)}`,
  }),
  async dispositivoDe(oid) {
    const filas = (await listar(C.listas.dispositivos, "ID_DISP", oid, "ID_DISP,ID_DISPOSITIVO,Title"))
      .filter((it) => txt(it.fields?.ID_DISP) === oid)
      .sort((a, b) => (+a.id || Infinity) - (+b.id || Infinity));
    const f = filas[0];
    return f ? { itemId: f.id, idDispositivo: txt(f.fields.ID_DISPOSITIVO).trim(), titulo: txt(f.fields.Title) } : null;
  },
  crearDispositivo: (oid, id, nombre) => crear(C.listas.dispositivos, { ID_DISP: oid, ID_DISPOSITIVO: id, NOMBRE: nombre, SESSION: "AUTH", Title: MARCA_MIGRADO }),
  migrarDispositivo: (itemId, id) => actualizar(C.listas.dispositivos, itemId, { ID_DISPOSITIVO: id, Title: MARCA_MIGRADO }),
  registrarSesion: (oid, resultado) => crear(C.listas.sesiones, { ID_USUARIO: oid, FECHA: F.diaSesion(ahora()), SESSION: resultado }),
  registrarUbicacionApp: (lat, lon, nombre) => crear(C.listas.ubicacionApp, {
    UBICACION: `${F.coordenadaUbicacionApp(lat)} ${F.coordenadaUbicacionApp(lon)}`, NOMBRE: nombre,
  }),
};

// =====================================================================
// Reglas de negocio (igual que Power Apps / APK)
// =====================================================================
async function jornadaDeHoy(oid, tipo) {
  const jornadas = await Graph.jornadasDe(oid);
  const hoy = F.dia(ahora());
  return jornadas.filter((j) => F.dia(j.inicio) === hoy && j.tipo.toLowerCase() === tipo.toLowerCase())
    .sort((a, b) => a.inicio - b.inicio)[0] || null;
}

/** OnStart de Power Apps: control de dispositivos_db + sesion_db. */
async function perfil() {
  const c = cuentaActual(); if (!c) throw new SinSesion();
  const id = idDispositivo();
  const fila = await Graph.dispositivoDe(c.oid);
  let resultado;
  if (!fila) { await Graph.crearDispositivo(c.oid, id, c.nombre); resultado = "REGISTRADO"; }
  else if (fila.idDispositivo.toLowerCase() === id.toLowerCase()) resultado = "AUTORIZADO";
  else if (fila.titulo !== MARCA_MIGRADO) { await Graph.migrarDispositivo(fila.itemId, id); resultado = "MIGRADO"; }
  else resultado = "BAD";
  try { await Graph.registrarSesion(c.oid, resultado === "BAD" ? "BAD" : "GOOD"); } catch { /* como Power Apps */ }
  if (resultado === "BAD")
    throw new Rechazo("DISPOSITIVO_NO_AUTORIZADO", "Este teléfono no está autorizado para su usuario. Si cambió de teléfono, contacte a TI para habilitar el nuevo.");
  completarUbicacionesPendientes();
  return resultado;
}

async function estadoJornada(tipo) {
  const c = cuentaActual(); if (!c) throw new SinSesion();
  const j = await jornadaDeHoy(c.oid, tipo);
  return {
    estado: !j ? "SIN_INGRESO" : !j.salida ? "EN_JORNADA" : "JORNADA_CERRADA",
    horaIngreso: j ? F.hora(j.inicio) : null,
    horaSalida: j?.salida ? F.hora(j.salida) : null,
    ubicacionPendiente: operaciones().some((o) => o.tipo === tipo && o.registroHecho),
  };
}

let marcando = false;
async function marcar(accion, tipo, lat, lon) {
  if (marcando) throw new Rechazo("EN_CURSO", "Ya se está registrando, espere un momento.");
  marcando = true;
  try {
    const c = cuentaActual(); if (!c) throw new SinSesion();
    let op = operaciones().find((o) => o.accion === accion && o.tipo === tipo && !o.registroHecho);
    const reanudando = !!op;

    if (!op) {
      const jornada = await jornadaDeHoy(c.oid, tipo); // también actualiza la hora oficial
      if (desfaseServidor == null) await sincronizarHora();
      const t = ahora();
      if (accion === "INGRESO") {
        if (jornada) throw new Rechazo("INGRESO_YA_REGISTRADO", "Su ingreso ya fue registrado");
        op = { accion, tipo, idReg: crypto.randomUUID ? crypto.randomUUID() : idDispositivo().slice(4), horaIso: t.toISOString(), inicioIso: t.toISOString(), itemIdRegistro: null, registroHecho: false, lat, lon };
      } else {
        if (!jornada) throw new Rechazo("SIN_INGRESO", "No se ha registrado su entrada");
        if (jornada.salida) throw new Rechazo("SALIDA_YA_REGISTRADA", "Su salida ya fue registrada");
        op = { accion, tipo, idReg: jornada.idReg, horaIso: t.toISOString(), inicioIso: jornada.inicio.toISOString(), itemIdRegistro: jornada.itemId, registroHecho: false, lat, lon };
      }
      guardarOperacion(op);
    }

    // Paso 1: registros_tb
    if (accion === "INGRESO") {
      if (reanudando) {
        const existente = await jornadaDeHoy(c.oid, op.tipo);
        if (existente && existente.idReg !== op.idReg) { quitarOperacion(op); throw new Rechazo("INGRESO_YA_REGISTRADO", "Su ingreso ya fue registrado"); }
        op.itemIdRegistro = existente ? existente.itemId : null;
      }
      if (!op.itemIdRegistro) op.itemIdRegistro = await Graph.crearJornada(c.oid, op.idReg, op.tipo, new Date(op.horaIso));
    } else {
      await Graph.registrarSalida(op.itemIdRegistro, new Date(op.horaIso)); // PATCH: idempotente
    }
    op.registroHecho = true;
    guardarOperacion(op);

    // Paso 2: ubicacion_db (si falla, la marcación vale y se reintenta al abrir la app)
    let ubicacionOk = true;
    try { await Graph.crearUbicacion(op.idReg, new Date(op.horaIso), op.lat, op.lon); quitarOperacion(op); } catch { ubicacionOk = false; }

    const ingreso = accion === "INGRESO";
    return {
      mensaje: ubicacionOk
        ? (ingreso ? "Se registró el ingreso correctamente" : "Se registró la salida correctamente")
        : (ingreso ? "Ingreso registrado" : "Salida registrada") + ". La ubicación se enviará al abrir la app; no repita la marcación.",
      estado: {
        estado: ingreso ? "EN_JORNADA" : "JORNADA_CERRADA",
        horaIngreso: F.hora(new Date(op.inicioIso)),
        horaSalida: ingreso ? null : F.hora(new Date(op.horaIso)),
        ubicacionPendiente: !ubicacionOk,
      },
    };
  } finally { marcando = false; }
}

async function completarUbicacionesPendientes() {
  for (const op of operaciones().filter((o) => o.registroHecho)) {
    try { await Graph.crearUbicacion(op.idReg, new Date(op.horaIso), op.lat, op.lon); quitarOperacion(op); } catch { /* luego */ }
  }
}

/** Convierte excepciones en un resultado para la pantalla. */
async function intentar(fn) {
  try { return { ok: true, valor: await fn() }; }
  catch (e) {
    if (e instanceof Rechazo) return { rechazo: true, codigo: e.codigo, mensaje: e.message };
    if (e instanceof SinSesion) return { sesionExpirada: true };
    if (e instanceof ErrorGraph) {
      if (e.estado === 401) return { sesionExpirada: true };
      if (e.estado === 403) return { rechazo: true, codigo: "SIN_PERMISO", mensaje: "Su usuario no tiene permiso en las listas de SharePoint. Contacte a TI." };
      if (e.estado == null) return { incierto: true, mensaje: "Sin conexión. Verifique su internet e intente nuevamente." };
      return { incierto: true, mensaje: `SharePoint no respondió correctamente (${e.estado}). Intente nuevamente.` };
    }
    console.error(e);
    return { incierto: true, mensaje: "Ocurrió un error inesperado. Intente nuevamente." };
  }
}

// =====================================================================
// Ubicación (GPS del navegador)
// =====================================================================
function obtenerUbicacion() {
  return new Promise((res) => {
    if (!navigator.geolocation) return res({ error: "NO_SOPORTADO" });
    navigator.geolocation.getCurrentPosition(
      (p) => res({ lat: p.coords.latitude, lon: p.coords.longitude, precision: p.coords.accuracy, t: Date.now() }),
      (e) => res({ error: e.code === 1 ? "PERMISO" : e.code === 2 ? "NO_DISPONIBLE" : "TIEMPO" }),
      { enableHighAccuracy: true, timeout: 20000, maximumAge: 30000 },
    );
  });
}
const esIOS = /iphone|ipad|ipod/i.test(navigator.userAgent) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
const MENSAJES_GPS = {
  PERMISO: esIOS
    ? "Permita la ubicación: Ajustes → Privacidad y seguridad → Localización → Sitios web de Safari → «Al usar la app». Luego toque «Actualizar»."
    : "Permita el acceso a la ubicación para este sitio y toque «Actualizar».",
  NO_DISPONIBLE: "No se pudo obtener la ubicación. Active la localización del teléfono e intente de nuevo.",
  TIEMPO: "La ubicación tardó demasiado. Salga a un lugar abierto y toque «Actualizar».",
  NO_SOPORTADO: "Este navegador no permite obtener la ubicación.",
};

// =====================================================================
// Estado de la pantalla y dibujo
// =====================================================================
const S = {
  iniciando: true, sesion: null, pantalla: "TIPOS", tipos: [], tipo: null, accion: "INGRESO",
  jornada: null, cargando: false, procesando: false, buscandoUbicacion: false, ubicacion: null,
  mensaje: null, bloqueado: null, marcacionSinConfirmar: false,
};
let ultimaPantalla = null;
const app = document.getElementById("app");
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const set = (cambios) => { Object.assign(S, cambios); dibujar(); };

const ICONOS = {
  atras: '<svg viewBox="0 0 24 24"><path d="M15 18l-6-6 6-6"/></svg>',
  salir: '<svg viewBox="0 0 24 24"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><path d="M16 17l5-5-5-5"/><path d="M21 12H9"/></svg>',
  pin: (color) => `<svg viewBox="0 0 24 24"><path fill="${color}" d="M12 2a7 7 0 0 0-7 7c0 5.2 7 13 7 13s7-7.8 7-13a7 7 0 0 0-7-7z"/><circle cx="12" cy="9" r="2.8" fill="#fff"/></svg>`,
};

function saludo(h) { return h >= 5 && h < 12 ? "Buenos días" : h >= 12 && h < 19 ? "Buenas tardes" : "Buenas noches"; }
function primerNombre(n) { const p = (n || "").trim().split(/\s+/)[0] || ""; return p ? p[0].toUpperCase() + p.slice(1).toLowerCase() : ""; }
function iniciales(n) { return (n || "").split(/\s+/).filter(Boolean).slice(0, 2).map((x) => x[0].toUpperCase()).join("") || "?"; }
const etiquetaTipo = () => (S.tipos.find((t) => t.valor === S.tipo) || {}).etiqueta || S.tipo || "";

function mensajeHTML(m, descartable = true) {
  if (!m) return "";
  return `<div class="mensaje ${m.tipo}"><span class="p"></span><span class="t">${esc(m.texto)}</span>${descartable ? '<button data-accion="descartar">OK</button>' : ""}</div>`;
}

function marco({ volver = false, salir = false, cabecera, contenido }) {
  return `
    <header class="cabecera">
      <div class="cabecera-fila">
        ${volver ? `<button class="icono-cab" data-accion="volver" aria-label="Atrás">${ICONOS.atras}</button>` : ""}
        <span class="espacio"></span>
        ${salir ? `<button class="icono-cab" data-accion="pedir-salida" aria-label="Cerrar sesión">${ICONOS.salir}</button>` : ""}
      </div>
      ${cabecera}
      <div class="acento"></div>
    </header>
    <section class="contenido">${contenido}</section>
    <footer class="pie">by ICT · Vepagas</footer>`;
}

function vistaLogin() {
  return `
    <div class="login">
      <div class="login-logo">
        <img src="logo.png" alt="VEPAGAS">
        <p>Nos empeñamos en la integridad de nuestra gente</p>
      </div>
      <div class="login-panel">
        <div class="acento"></div>
        <h1>Control de asistencia</h1>
        <p class="desc">Ingrese con su cuenta @vepagas.com o con su Gmail autorizado.</p>
        ${mensajeHTML(S.mensaje)}
        <button class="btn" data-accion="login" ${S.cargando ? "disabled" : ""}>${S.cargando ? '<span class="girador"></span>' : "Iniciar sesión"}</button>
        <p class="legal">Solo cuentas autorizadas por Talento Humano. La ubicación se registra únicamente al timbrar.</p>
      </div>
    </div>`;
}

function vistaTipos(anim) {
  const t = ahora(); const p = F.partes(t);
  const mostrarInstalar = esIOS && !navigator.standalone && Local.leer("vepagas_sin_aviso_instalar") !== "1";
  return marco({
    salir: true,
    cabecera: `
      <div class="saludo tenue" data-saludo>${saludo(p.h)},</div>
      <div class="nombre-grande">${esc(primerNombre(S.sesion?.nombre) || "Bienvenido")}</div>
      <div class="reloj-fila"><span class="reloj" data-hhmm>${F.hora(t)}</span><span class="fecha tenue" data-dia>${F.diaLegible(t)}</span></div>`,
    contenido: `
      ${S.bloqueado ? `${mensajeHTML({ tipo: "error", texto: S.bloqueado }, false)}<button class="btn-texto" data-accion="reintentar" ${S.cargando ? "disabled" : ""} style="align-self:center">Reintentar</button>` : ""}
      ${S.cargando && !S.tipos.length ? '<div class="centro"><span class="girador rojo"></span>Verificando su dispositivo…</div>' : ""}
      ${S.tipos.map((tp, i) => `
        <div class="tarjeta tocable ${anim ? "entrar retraso-" + (i + 1) : ""}" data-accion="elegir" data-valor="${esc(tp.valor)}" role="button" tabindex="0">
          <div class="accion-principal">
            <div class="cuadro-icono">${ICONOS.pin("#fff")}</div>
            <div style="flex:1"><div class="titulo-tarjeta" style="margin:0">${esc(tp.etiqueta)}</div><div class="sub">Marque su ingreso y salida de hoy</div></div>
            <span class="flecha">›</span>
          </div>
        </div>`).join("")}
      ${mensajeHTML(S.mensaje)}
      ${mostrarInstalar ? `
        <div class="tarjeta ${anim ? "entrar retraso-2" : ""}">
          <div class="fila"><span class="etiqueta">Instale la app en su iPhone</span><span class="espacio"></span><button class="btn-texto" data-accion="ocultar-instalar">Ocultar</button></div>
          <p class="instalar" style="margin:8px 0 0">En Safari toque <b>Compartir</b> (el cuadro con la flecha ⬆) y luego <b>«Agregar a inicio»</b>. Así se abre como una app, sin barra del navegador.</p>
        </div>` : ""}
      <div class="logo-pastilla"><img src="logo.png" alt="VEPAGAS"></div>
      <div class="correo">${esc(S.sesion?.correo)}</div>`,
  });
}

function pasoHTML(titulo, hora, hecho, activo, linea) {
  return `
    <div class="paso ${hecho ? "hecho" : activo ? "activo" : ""}">
      <div class="col"><div class="punto">${hecho ? "✓" : "<i></i>"}</div>${linea ? '<div class="linea"></div>' : ""}</div>
      <div class="txt"><b>${titulo}</b><span>${hecho ? esc(hora) : activo ? "Pendiente" : "—"}</span></div>
    </div>`;
}

function vistaMenu(anim) {
  const t = ahora();
  const j = S.jornada; const e = j?.estado;
  const chip = { SIN_INGRESO: '<span class="chip gris">Pendiente</span>', EN_JORNADA: '<span class="chip rojo">En jornada</span>', JORNADA_CERRADA: '<span class="chip verde">Completa</span>' }[e]
    || (S.cargando ? '<span class="girador rojo" style="width:16px;height:16px;border-width:2px"></span>' : "");
  const dudoso = S.marcacionSinConfirmar || (!e && !S.cargando);
  let accion = "";
  if (dudoso) accion = '<button class="btn" data-accion="ingreso">Registrar ingreso</button><button class="btn oscuro" data-accion="salida">Registrar salida</button>';
  else if (e === "SIN_INGRESO") accion = '<button class="btn" data-accion="ingreso">Registrar ingreso</button>';
  else if (e === "EN_JORNADA") accion = '<button class="btn oscuro" data-accion="salida">Registrar salida</button>';
  else if (e === "JORNADA_CERRADA") accion = `
    <div class="tarjeta entrar"><div class="accion-principal">
      <div class="punto" style="width:40px;height:40px;background:rgba(46,125,50,.12);color:var(--verde);font-size:20px">✓</div>
      <div><div style="font-weight:600">Jornada de hoy completa</div><div class="sub">Gracias por su trabajo. ¡Hasta mañana!</div></div>
    </div></div>`;

  return marco({
    volver: true, salir: true,
    cabecera: `
      <div class="perfil">
        <div class="avatar"><span>${esc(iniciales(S.sesion?.nombre))}</span></div>
        <div><div class="tenue" style="font-size:14px">¡Bienvenido!</div><div class="n">${esc(S.sesion?.nombre)}</div><div class="s"><span>${esc(etiquetaTipo())}</span> · <span data-dia>${F.diaLegible(t)}</span></div></div>
      </div>`,
    contenido: `
      <div class="tarjeta ${anim ? "entrar" : ""}">
        <div class="fila"><span class="etiqueta">Jornada de hoy</span><span class="espacio"></span>${chip}</div>
        <div class="pasos">
          ${pasoHTML("Ingreso", j?.horaIngreso, e === "EN_JORNADA" || e === "JORNADA_CERRADA", e === "SIN_INGRESO", true)}
          ${pasoHTML("Salida", j?.horaSalida, e === "JORNADA_CERRADA", e === "EN_JORNADA", false)}
        </div>
        ${j?.ubicacionPendiente ? '<div class="sub" style="color:var(--ambar);margin-top:10px">Ubicación pendiente de sincronizar</div>' : ""}
      </div>
      ${mensajeHTML(S.mensaje)}
      ${S.marcacionSinConfirmar ? mensajeHTML({ tipo: "aviso", texto: "Hay una timbrada sin confirmar. Vuelva a timbrar la misma opción; no se duplicará." }, false) : ""}
      <div style="display:flex;flex-direction:column;gap:12px;margin-top:4px" class="${anim ? "entrar retraso-1" : ""}">${accion}</div>`,
  });
}

function vistaRegistro(anim) {
  const t = ahora(); const u = S.ubicacion; const ingreso = S.accion === "INGRESO";
  const precision = u ? `<span class="chip ${u.precision <= 50 ? "verde" : "ambar"}">±${Math.round(u.precision)} m</span>` : "";
  return marco({
    volver: true,
    cabecera: `
      <div class="tenue" style="font-size:15px">${ingreso ? "Registro de ingreso" : "Registro de salida"}</div>
      <div style="display:flex;align-items:flex-end;margin-top:4px"><span class="reloj-grande" data-hhmm>${F.hora(t)}</span><span class="segundos" data-ss>:${dos(F.partes(t).s)}</span></div>
      <div class="fecha tenue" data-dia style="margin-top:6px">${F.diaLegible(t)}</div>`,
    contenido: `
      <div class="tarjeta ${anim ? "entrar" : ""}">
        <span class="etiqueta">Colaborador</span>
        <div class="titulo-tarjeta">${esc(S.sesion?.nombre)}</div>
        <div class="sub">${esc(etiquetaTipo())}</div>
      </div>
      <div class="tarjeta ${u && !S.buscandoUbicacion ? "tocable" : ""} ${anim ? "entrar retraso-1" : ""}" ${u && !S.buscandoUbicacion ? 'data-accion="mapa" role="button" tabindex="0"' : ""}>
        <div class="fila"><span class="etiqueta">Su ubicación</span><span class="espacio"></span>
          <button class="btn-texto" data-accion="ubicar" ${S.buscandoUbicacion || S.procesando ? "disabled" : ""}>${S.buscandoUbicacion ? "Buscando…" : "Actualizar"}</button></div>
        <div class="mapa">
          ${u || S.buscandoUbicacion ? '<span class="onda"></span><span class="onda b"></span>' : ""}
          <span class="pin">${ICONOS.pin(u ? "#AF1740" : "#6B6B6B")}</span>
          ${u ? '<span class="ver">Ver en Google Maps ↗</span>' : ""}
        </div>
        <div style="margin-top:12px">
          ${u ? `<div class="fila"><div style="flex:1"><div class="coords">Lat ${u.lat.toFixed(6)} &nbsp; Lon ${u.lon.toFixed(6)}</div><div class="sub" style="font-size:11px">Toque el mapa para ver desde dónde está timbrando</div></div>${precision}</div>`
            : `<div class="sub">${S.buscandoUbicacion ? "Obteniendo su ubicación…" : "Ubicación no disponible. Toque «Actualizar»."}</div>`}
        </div>
      </div>
      ${mensajeHTML(S.mensaje)}
      <button class="btn ${ingreso ? "" : "oscuro"} ${anim ? "entrar retraso-2" : ""}" data-accion="timbrar" ${S.procesando ? "disabled" : ""}>
        ${S.procesando ? '<span class="girador"></span>' : ingreso ? "Timbrar ingreso" : "Timbrar salida"}</button>
      ${S.procesando ? '<div class="sub" style="text-align:center">Registrando…</div>' : ""}`,
  });
}

function dibujar() {
  const clave = S.iniciando ? "INICIO" : !S.sesion ? "LOGIN" : S.pantalla;
  const anim = clave !== ultimaPantalla;
  ultimaPantalla = clave;
  if (clave === "INICIO") app.innerHTML = '<div class="centro" style="flex:1;justify-content:center"><span class="girador rojo"></span></div>';
  else if (clave === "LOGIN") app.innerHTML = vistaLogin();
  else if (clave === "TIPOS") app.innerHTML = vistaTipos(anim);
  else if (clave === "MENU") app.innerHTML = vistaMenu(anim);
  else app.innerHTML = vistaRegistro(anim);
  if (anim) window.scrollTo(0, 0);
}

// Reloj: solo se actualizan los textos, sin redibujar (no reinicia animaciones).
setInterval(() => {
  const t = ahora(); const p = F.partes(t);
  app.querySelectorAll("[data-hhmm]").forEach((e) => (e.textContent = F.hora(t)));
  app.querySelectorAll("[data-ss]").forEach((e) => (e.textContent = ":" + dos(p.s)));
  app.querySelectorAll("[data-dia]").forEach((e) => (e.textContent = F.diaLegible(t)));
  app.querySelectorAll("[data-saludo]").forEach((e) => (e.textContent = saludo(p.h) + ","));
}, 1000);

// =====================================================================
// Acciones
// =====================================================================
async function iniciarSesion() {
  set({ cargando: true, mensaje: null });
  try { await msalApp.loginRedirect({ scopes: SCOPES, prompt: "select_account" }); }
  catch (e) { set({ cargando: false, mensaje: { tipo: "error", texto: "No se pudo iniciar sesión con Microsoft: " + (e.message || e) } }); }
}

async function cerrarSesion(mensaje = null) {
  const c = cuentaActual();
  Local.guardar("vepagas_operaciones", null);
  try { if (c) await msalApp.logoutRedirect({ account: c.cuenta, onRedirectNavigate: () => false }); } catch { /* local */ }
  Object.assign(S, { sesion: null, pantalla: "TIPOS", tipos: [], tipo: null, jornada: null, ubicacion: null, bloqueado: null, cargando: false, procesando: false, marcacionSinConfirmar: false });
  set({ mensaje });
}
const sesionExpirada = () => cerrarSesion({ tipo: "aviso", texto: "Su sesión expiró. Inicie sesión nuevamente." });

async function cargarPerfil() {
  set({ cargando: true, bloqueado: null, mensaje: null });
  const r = await intentar(perfil);
  if (r.sesionExpirada) return sesionExpirada();
  if (r.ok) {
    const texto = r.valor === "AUTORIZADO" ? null : "Dispositivo registrado y autorizado";
    return set({ cargando: false, tipos: C.tipos, mensaje: texto && { tipo: "exito", texto } });
  }
  if (r.codigo === "DISPOSITIVO_NO_AUTORIZADO") return set({ cargando: false, tipos: [], bloqueado: r.mensaje });
  set({ cargando: false, tipos: [], bloqueado: r.mensaje });
}

async function elegirTipo(valor) {
  set({ tipo: valor, jornada: null, pantalla: "MENU", mensaje: null, marcacionSinConfirmar: operaciones().some((o) => !o.registroHecho) });
  refrescar();
  // Como Power Apps: Patch(UBICACION_APP). En segundo plano; sin GPS se registra sin coordenadas.
  const u = await obtenerUbicacion();
  if (!u.error) S.ubicacion = u;
  try { await Graph.registrarUbicacionApp(u.error ? null : u.lat, u.error ? null : u.lon, S.sesion?.nombre || ""); } catch { /* no bloquea */ }
}

async function refrescar() {
  set({ cargando: true });
  const r = await intentar(() => estadoJornada(S.tipo));
  if (r.sesionExpirada) return sesionExpirada();
  set({ cargando: false, jornada: r.ok ? r.valor : S.jornada, mensaje: r.ok ? S.mensaje : { tipo: r.rechazo ? "error" : "aviso", texto: r.mensaje } });
}

async function actualizarUbicacion() {
  if (S.buscandoUbicacion || S.procesando) return;
  set({ buscandoUbicacion: true });
  const u = await obtenerUbicacion();
  if (u.error) set({ buscandoUbicacion: false, mensaje: { tipo: "error", texto: MENSAJES_GPS[u.error] } });
  else set({ buscandoUbicacion: false, ubicacion: u, mensaje: S.mensaje?.tipo === "error" ? null : S.mensaje });
}

async function timbrar() {
  if (S.procesando) return;
  set({ procesando: true, mensaje: null });
  let u = S.ubicacion;
  if (!u || Date.now() - u.t > 60000) {
    u = await obtenerUbicacion();
    if (u.error) return set({ procesando: false, mensaje: { tipo: "error", texto: MENSAJES_GPS[u.error] } });
    S.ubicacion = u;
  }
  const r = await intentar(() => marcar(S.accion, S.tipo, u.lat, u.lon));
  if (r.sesionExpirada) return sesionExpirada();
  if (r.ok) return set({ procesando: false, pantalla: "MENU", jornada: r.valor.estado, marcacionSinConfirmar: false, mensaje: { tipo: r.valor.estado.ubicacionPendiente ? "aviso" : "exito", texto: r.valor.mensaje } });
  if (r.rechazo) { set({ procesando: false, marcacionSinConfirmar: operaciones().some((o) => !o.registroHecho), mensaje: { tipo: "error", texto: r.mensaje } }); return; }
  set({ procesando: false, marcacionSinConfirmar: true, mensaje: { tipo: "aviso", texto: r.mensaje + " Vuelva a timbrar; no se duplicará." } });
}

function abrirMapa() {
  const u = S.ubicacion; if (!u) return;
  window.open(`https://www.google.com/maps/search/?api=1&query=${u.lat.toFixed(7)},${u.lon.toFixed(7)}`, "_blank", "noopener");
}

const dialogo = document.getElementById("dialogo");
document.addEventListener("click", (ev) => {
  const el = ev.target.closest("[data-accion]");
  if (!el || el.disabled) return;
  switch (el.dataset.accion) {
    case "login": return iniciarSesion();
    case "descartar": return set({ mensaje: null });
    case "pedir-salida": dialogo.hidden = false; return;
    case "cancelar-salida": dialogo.hidden = true; return;
    case "confirmar-salida": dialogo.hidden = true; return cerrarSesion();
    case "reintentar": return cargarPerfil();
    case "elegir": return elegirTipo(el.dataset.valor);
    case "ocultar-instalar": Local.guardar("vepagas_sin_aviso_instalar", "1"); return dibujar();
    case "volver": return set({ pantalla: S.pantalla === "REGISTRO" ? "MENU" : "TIPOS", mensaje: null });
    case "ingreso": set({ accion: "INGRESO", pantalla: "REGISTRO", mensaje: null }); return actualizarUbicacion();
    case "salida": set({ accion: "SALIDA", pantalla: "REGISTRO", mensaje: null }); return actualizarUbicacion();
    case "ubicar": ev.stopPropagation(); return actualizarUbicacion();
    case "mapa": return abrirMapa();
    case "timbrar": return timbrar();
  }
});
document.addEventListener("keydown", (ev) => {
  if ((ev.key === "Enter" || ev.key === " ") && ev.target.matches('[role="button"][data-accion]')) { ev.preventDefault(); ev.target.click(); }
});
// Al volver a la app (p. ej. desde Google Maps) se refresca el estado del día.
document.addEventListener("visibilitychange", () => { if (!document.hidden && S.sesion && S.pantalla === "MENU" && !S.cargando) refrescar(); });

// =====================================================================
// Arranque
// =====================================================================
(async function arrancar() {
  dibujar();
  sincronizarHora();
  try {
    await msalApp.initialize();
    const r = await msalApp.handleRedirectPromise();
    if (r?.account) msalApp.setActiveAccount(r.account);
  } catch (e) {
    S.mensaje = { tipo: "error", texto: "No se pudo completar el inicio de sesión: " + (e.errorMessage || e.message || e) };
  }
  const c = cuentaActual();
  if (c) msalApp.setActiveAccount(c.cuenta);
  S.iniciando = false;
  S.sesion = c && { nombre: c.nombre, correo: c.correo };
  dibujar();
  if (c) cargarPerfil();
})();
