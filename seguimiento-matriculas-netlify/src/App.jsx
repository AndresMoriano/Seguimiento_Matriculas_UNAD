import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ReferenceLine, ResponsiveContainer, Cell, LabelList,
} from "recharts";

/* ============================================================
   SEGUIMIENTO DE MATRÍCULAS — UDR VALLE DEL GUAMUEZ (UNAD)
   Seguimiento por período y anual. Actualización mediante la
   subida de bases de facturas pagadas y pendientes por pagar.
   ============================================================ */

const ESCUELAS = ["ECACEN", "ECISA", "ECBTI", "ECJP", "ECAPMA", "ECSAH", "ECEDU", "INVIL"];

const C = {
  tinta: "#0F2440",
  azul: "#1F4E9C",
  azulSuave: "#E8EFFA",
  ambar: "#F5B300",
  verde: "#1B8A5A",
  rojo: "#C0402E",
  gris: "#5B6B80",
  borde: "#DDE4EE",
  fondo: "#F4F6FA",
};

/* ---------- Datos iniciales (tomados del archivo metas.xlsx) ---------- */
// [metaCred, avCred, metaEst, avEst, factPend]
const SEED_RAW = {
  "2025": [
    { nombre: "16-01", promedio: 14, d: { ECACEN: [3237, 3680, 230, 263, null], ECISA: [1270, 1385, 95, 104, null], ECBTI: [3285, 3352, 232, 242, null], ECJP: [107, 170, 8, 12, null], ECAPMA: [1196, 1997, 85, 144, null], ECSAH: [2311, 2874, 156, 196, null], ECEDU: [2319, 2952, 168, 217, null], INVIL: [null, null, null, null, null] } },
    { nombre: "16-02", promedio: 15, d: { ECACEN: [834, 424, 60, 32, 5], ECISA: [355, 120, 25, 11, 0], ECBTI: [614, 304, 44, 24, 1], ECJP: [0, 32, 0, 3, 1], ECAPMA: [285, 285, 20, 21, 1], ECSAH: [472, 314, 34, 24, 0], ECEDU: [811, 478, 58, 42, 10], INVIL: [null, null, null, null, null] } },
    { nombre: "8-03", promedio: 7, d: { ECACEN: [106, 127, 15, 24, 0], ECISA: [16, 72, 2, 12, 0], ECBTI: [165, 207, 24, 30, 0], ECJP: [16, 22, 2, 4, 0], ECAPMA: [29, 78, 4, 17, 0], ECSAH: [146, 194, 21, 31, 0], ECEDU: [113, 195, 16, 30, 0], INVIL: [null, 3, null, 1, null] } },
    { nombre: "16-04", promedio: 14, d: { ECACEN: [3737, 4288, 266.9, 320, null], ECISA: [1553, 1731, 110.9, 133, null], ECBTI: [4120, 4138, 294.3, 303, null], ECJP: [218, 238, 15.6, 18, null], ECAPMA: [2094, 2550, 149.6, 186, null], ECSAH: [2901, 3352, 207.2, 241, null], ECEDU: [3772, 3904, 269.4, 293, null], INVIL: [null, 12, 0, 4, null] } },
    { nombre: "16-05", promedio: 7, d: { ECACEN: [463, 815, 33.1, 64, 0], ECISA: [191, 515, 13.6, 40, 0], ECBTI: [175, 679, 12.5, 51, 0], ECJP: [8, 33, 0.6, 3, 0], ECAPMA: [216, 755, 15.4, 57, 0], ECSAH: [247, 367, 17.6, 30, 0], ECEDU: [91, 895, 6.5, 67, 0], INVIL: [6, 12, 2, 5, null] } },
  ],
  "2026": [
    { nombre: "16-01", promedio: 14, d: { ECACEN: [3608, 4305, 257.7, 322, 18], ECISA: [1504, 2294, 107.4, 173, 13], ECBTI: [3364, 3965, 240.3, 289, 14], ECJP: [202, 299, 14.4, 24, 1], ECAPMA: [1736, 2797, 124, 200, 14], ECSAH: [2475, 3356, 176.8, 244, 7], ECEDU: [3010, 3685, 215, 271, 19], INVIL: [null, 3, null, 1, null] } },
    { nombre: "16-02", promedio: 14, d: { ECACEN: [364, 929, 26, 77, 8], ECISA: [195, 680, 13.9, 54, 6], ECBTI: [632, 828, 45.1, 63, 3], ECJP: [29, 70, 2.1, 5, 0], ECAPMA: [255, 641, 18.2, 48, 4], ECSAH: [572, 605, 40.9, 47, 6], ECEDU: [328, 986, 23.4, 76, 6], INVIL: [null, 3, 0, 5, null] } },
    { nombre: "8-03", promedio: 7, d: { ECACEN: [150, 215, 21.4, 27, 6], ECISA: [42, 110, 6, 15, 5], ECBTI: [171, 166, 24.4, 25, 6], ECJP: [0, 5, 0, 1, 3], ECAPMA: [65, 100, 9.3, 16, 4], ECSAH: [468, 137, 66.9, 20, 8], ECEDU: [106, 117, 15.1, 18, 2], INVIL: [9, 9, 4, 4, null] } },
    { nombre: "16-04", promedio: 13, d: { ECACEN: [4228, 3608, 325.2, 260, 0], ECISA: [1897, 1906, 145.9, 139, 0], ECBTI: [4251, 3346, 327, 241, 0], ECJP: [225, 170, 17.3, 13, 0], ECAPMA: [2184, 2393, 168, 167, 0], ECSAH: [2560, 2128, 196.9, 170, 0], ECEDU: [3817, 3532, 293.6, 255, 0], INVIL: [3, 3, 4, 2, null] } },
    { nombre: "16-05", promedio: 14, d: { ECACEN: [null, null, 0, null, null], ECISA: [null, null, 0, null, null], ECBTI: [null, null, 0, null, null], ECJP: [null, null, 0, null, null], ECAPMA: [null, null, 0, null, null], ECSAH: [null, null, 0, null, null], ECEDU: [null, null, 0, null, null], INVIL: [null, 12, 0, null, null] } },
  ],
};

function estadoInicial() {
  const anios = {};
  for (const [anio, periodos] of Object.entries(SEED_RAW)) {
    anios[anio] = {
      periodos: periodos.map((p, i) => ({
        id: `${anio}-${p.nombre}-${i}`,
        nombre: p.nombre,
        promedio: p.promedio,
        escuelas: Object.fromEntries(
          ESCUELAS.map((e) => {
            const v = p.d[e] || [null, null, null, null, null];
            return [e, { metaCred: v[0], avCred: v[1], metaEst: v[2], avEst: v[3], factPend: v[4], avEstNuevos: null, metaEstNuevos: null, metaEstAntiguos: null, metaCredNuevos: null, metaCredAntiguos: null, avCredNuevos: null, avCredAntiguos: null, credPendNuevos: null, credPendAntiguos: null, estPendNuevos: null, estPendAntiguos: null, valorPend: null }];
          })
        ),
        fuentePagadas: null,
        fuentePendientes: null,
      })),
    };
  }
  return { anios };
}

/* ---------------------- utilidades ---------------------- */
const KEY = "unad-seguimiento-matriculas-v1";

/* Almacenamiento local del navegador. Reemplaza al window.storage que solo
   existe dentro de Claude, conservando la misma forma asíncrona (get/set) para
   que el resto del código no cambie. Los datos se guardan en el equipo de quien
   usa la aplicación; no se comparten entre computadores ni navegadores. */
const storage = {
  async get(clave) {
    try {
      const v = window.localStorage.getItem(clave);
      return v === null ? null : { value: v };
    } catch (e) {
      return null;
    }
  },
  async set(clave, valor) {
    try {
      window.localStorage.setItem(clave, valor);
      return true;
    } catch (e) {
      return false;
    }
  },
};

const num = (v) => (typeof v === "number" && isFinite(v) ? v : null);
const fmt = (v, dec = 0) =>
  v === null || v === undefined ? "—" : Number(v).toLocaleString("es-CO", { maximumFractionDigits: dec, minimumFractionDigits: 0 });
/* Las metas de estudiantes se muestran siempre enteras: son personas. */
const fmtEst = (v) => (v === null || v === undefined ? "—" : fmt(Math.round(v)));
const pct = (av, meta) => {
  if (!meta || meta === 0) return null;
  return ((av || 0) / meta) * 100;
};
const fmtPct = (p) => (p === null ? "—" : `${p.toLocaleString("es-CO", { maximumFractionDigits: 1 })}%`);
const colorPct = (p) => (p === null ? C.gris : p >= 100 ? C.verde : p >= 70 ? "#B07E00" : C.rojo);

const sinTildes = (s) => String(s).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().trim();

function normalizarEscuela(valor) {
  if (valor === null || valor === undefined) return null;
  const v = sinTildes(valor);
  if (!v) return null;
  if (ESCUELAS.includes(v)) return v;
  if (v.includes("ECACEN") || v.includes("ADMINISTRATIVAS") || v.includes("CONTABLES")) return "ECACEN";
  if (v.includes("ECISA") || v.includes("ECSAL") || v.includes("SALUD")) return "ECISA";
  if (v.includes("ECBTI") || v.includes("BASICAS") || v.includes("TECNOLOGIA") || v.includes("INGENIERIA")) return "ECBTI";
  if (v.includes("ECJP") || v.includes("JURIDICAS") || v.includes("POLITICAS")) return "ECJP";
  if (v.includes("ECAPMA") || v.includes("AGRICOLAS") || v.includes("PECUARIAS") || v.includes("AMBIENTE")) return "ECAPMA";
  if (v.includes("ECSAH") || v.includes("SOCIALES") || v.includes("ARTES") || v.includes("HUMANIDADES")) return "ECSAH";
  if (v.includes("ECEDU") || v.includes("EDUCACION")) return "ECEDU";
  if (v.includes("INVIL") || v.includes("LENGUAS")) return "INVIL";
  return "SIN_CLASIFICAR";
}

function aNumero(v) {
  if (typeof v === "number") return isFinite(v) ? v : null;
  if (typeof v !== "string") return null;
  let s = v.replace(/[$\s]/g, "");
  if (!s) return null;
  // Formato colombiano: 10.000,50 → 10000.50
  if (/,\d{1,2}$/.test(s)) s = s.replace(/\./g, "").replace(",", ".");
  else s = s.replace(/,/g, "");
  const n = Number(s);
  return isFinite(n) ? n : null;
}

/* ------- lectura y agregación de las bases subidas ------- */
function detectarEncabezado(filas, tipo) {
  for (let i = 0; i < Math.min(filas.length, 20); i++) {
    const fila = filas[i].map((c) => sinTildes(c ?? ""));
    const tieneEscuela = fila.some((c) => c === "ESCUELA" || (c.includes("ESCUELA") && !c.includes("CODIGO")));
    const tieneId = fila.some((c) => c.includes("DOCUMENTO") || c === "ID" || c.includes("NOMBRES"));
    const tieneFactura = fila.some((c) => c.includes("FACTURA"));
    const tieneNuevosAntiguos = fila.some((c) => c.includes("NUEVO") || c.includes("ANTIGUO"));
    const tieneConcertado = fila.some((c) => c.includes("CONCERTADO") || c.includes("EJECUTADO"));
    const tieneBloque = fila.some((c) => c.includes("BLOQUE") || c.includes("COHORTE"));
    if (tipo === "metas") { if (tieneEscuela && (tieneNuevosAntiguos || tieneConcertado || tieneBloque)) return i; continue; }
    if (tieneEscuela && (tipo === "pagadas" ? tieneId : tieneId || tieneFactura)) return i;
  }
  return -1;
}

function indiceColumna(encabezado, pruebas) {
  const cols = encabezado.map((c) => sinTildes(c ?? ""));
  for (const prueba of pruebas) {
    const idx = cols.findIndex((c) => c && prueba(c));
    if (idx >= 0) return idx;
  }
  return -1;
}

/* Los reportes institucionales exportan los créditos con punto como separador
   de miles, que al convertirse a número quedan como 3.612 (=3612) o 3.36 (=3360,
   con el cero final perdido). Un entero se toma tal cual; un decimal se reescala
   a miles rellenando hasta tres decimales. */
function creditosDelReporte(v) {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") {
    if (!isFinite(v)) return null;
    if (Number.isInteger(v)) return v;
    return Math.round(v * 1000);
  }
  if (typeof v !== "string") return null;
  const s = v.trim();
  if (!s || s === "-") return null;
  if (/^\d{1,3}(\.\d{3})+$/.test(s)) return Number(s.replace(/\./g, ""));
  const n = aNumero(s);
  if (n === null) return null;
  return Number.isInteger(n) ? n : Math.round(n * 1000);
}

/* Reconoce el bloque venga como fecha real, como número serial de Excel
   (46038 = 16 de enero de 2026) o ya escrito ("16-01", "8-03"). Una fila sin
   bloque es un subtotal. */
function bloqueDelReporte(v) {
  if (v === null || v === undefined || v === "" || v === "-") return null;
  if (v instanceof Date) return `${String(v.getUTCDate()).padStart(2, "0")}-${String(v.getUTCMonth() + 1).padStart(2, "0")}`;
  /* Excel guarda las fechas como días desde el 30/12/1899. Un valor grande sin
     separadores en la columna de bloque es una fecha, no un número real. */
  if (typeof v === "number" && isFinite(v) && v > 20000 && v < 80000) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
    return `${String(d.getUTCDate()).padStart(2, "0")}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`;
  }
  const s = String(v).trim();
  if (!s || s === "-") return null;
  if (/^\d{4,5}(\.\d+)?$/.test(s)) {
    const n = Number(s);
    if (n > 20000 && n < 80000) {
      const d = new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000);
      return `${String(d.getUTCDate()).padStart(2, "0")}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`;
    }
  }
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return `${iso[3]}-${iso[2]}`;
  const dm = s.match(/^(\d{1,2})\s*[-/]\s*(\d{1,2})$/);
  if (dm) return `${dm[1].padStart(2, "0")}-${dm[2].padStart(2, "0")}`;
  return s;
}

/* Dos bloques coinciden si sus dos números son iguales, sin importar ceros
   a la izquierda: "16-1" del archivo base equivale a "16-01" del reporte. */
function mismoBloque(a, b) {
  if (!a || !b) return false;
  const partes = (x) => String(x).trim().split(/[-/]/).map((p) => parseInt(p, 10));
  const pa = partes(a), pb = partes(b);
  if (pa.length !== 2 || pb.length !== 2 || pa.some(isNaN) || pb.some(isNaN)) {
    return sinTildes(a) === sinTildes(b);
  }
  return pa[0] === pb[0] && pa[1] === pb[1];
}

function agregarBaseMetas(filas) {
  const hIdx = detectarEncabezado(filas, "metas");
  if (hIdx < 0) throw new Error("No se encontró la fila de encabezados. La base de metas debe incluir una columna «Escuela» y columnas de «Nuevos» y «Antiguos».");
  const enc = filas[hIdx];
  const colEscuela = indiceColumna(enc, [(c) => c === "ESCUELA", (c) => c.includes("ESCUELA") && !c.includes("CODIGO")]);
  const colCohorte = indiceColumna(enc, [(c) => c.includes("COHORTE")]);
  const colBloque = indiceColumna(enc, [(c) => c.includes("BLOQUE"), (c) => c.includes("PERIODO")]);

  /* La fila superior agrupa las columnas en Nuevos / Antiguos / Total y debajo
     se repiten Concertado, Ejecutado y Avance %. Se toma el «Concertado» de
     cada grupo: es la meta. */
  const grupo = hIdx > 0 ? filas[hIdx - 1] : [];
  const encabezados = enc.map((c) => sinTildes(c ?? ""));
  let etiquetaActual = "";
  const columnas = { NUEVOS: {}, ANTIGUOS: {}, TOTAL: {} };
  for (let c = 0; c < encabezados.length; c++) {
    const g = sinTildes((grupo && grupo[c]) ?? "");
    if (g.includes("NUEVO")) etiquetaActual = "NUEVOS";
    else if (g.includes("ANTIGUO")) etiquetaActual = "ANTIGUOS";
    else if (g.includes("TOTAL")) etiquetaActual = "TOTAL";
    if (!etiquetaActual) continue;
    const h = encabezados[c];
    if (h.includes("CONCERTADO") || h.includes("META")) columnas[etiquetaActual].concertado = c;
    else if (h.includes("EJECUTADO") || h.includes("AVANCE")) {
      if (!h.includes("%")) columnas[etiquetaActual].ejecutado = c;
    }
  }

  /* Sin fila de grupos, se buscan encabezados que nombren el grupo directamente. */
  if (columnas.NUEVOS.concertado === undefined) {
    const n = indiceColumna(enc, [(c) => c.includes("NUEVO")]);
    if (n >= 0) columnas.NUEVOS.concertado = n;
  }
  if (columnas.ANTIGUOS.concertado === undefined) {
    const a = indiceColumna(enc, [(c) => c.includes("ANTIGUO")]);
    if (a >= 0) columnas.ANTIGUOS.concertado = a;
  }

  const nuevosDetectados = columnas.NUEVOS.concertado !== undefined;
  const antiguosDetectados = columnas.ANTIGUOS.concertado !== undefined;
  if (!nuevosDetectados && !antiguosDetectados) {
    throw new Error("No se encontraron las columnas de metas de «Nuevos» ni de «Antiguos».");
  }

  const porBloque = {};   // bloque → escuela → metas
  const porCohorte = {};  // cohorte (Anual, 2026-1, …) → escuela → metas
  let filasLeidas = 0, sinClasificar = 0, filasSubtotal = 0;

  for (let i = hIdx + 1; i < filas.length; i++) {
    const fila = filas[i];
    if (!fila || fila.every((c) => c === null || c === undefined || c === "")) continue;
    const esc = normalizarEscuela(fila[colEscuela]);
    if (!esc) continue;
    filasLeidas++;
    if (esc === "SIN_CLASIFICAR") { sinClasificar++; continue; }

    const registro = {
      metaNuevos: nuevosDetectados ? creditosDelReporte(fila[columnas.NUEVOS.concertado]) : null,
      metaAntiguos: antiguosDetectados ? creditosDelReporte(fila[columnas.ANTIGUOS.concertado]) : null,
      avNuevos: columnas.NUEVOS.ejecutado !== undefined ? creditosDelReporte(fila[columnas.NUEVOS.ejecutado]) : null,
      avAntiguos: columnas.ANTIGUOS.ejecutado !== undefined ? creditosDelReporte(fila[columnas.ANTIGUOS.ejecutado]) : null,
      metaTotal: columnas.TOTAL.concertado !== undefined ? creditosDelReporte(fila[columnas.TOTAL.concertado]) : null,
      avTotal: columnas.TOTAL.ejecutado !== undefined ? creditosDelReporte(fila[columnas.TOTAL.ejecutado]) : null,
    };

    const bloque = colBloque >= 0 ? bloqueDelReporte(fila[colBloque]) : null;
    const cohorte = colCohorte >= 0 ? String(fila[colCohorte] ?? "").trim() : "";

    if (bloque) {
      if (!porBloque[bloque]) porBloque[bloque] = {};
      porBloque[bloque][esc] = registro;
    } else {
      // Fila de subtotal (Anual o cohorte completa): se conserva como referencia.
      filasSubtotal++;
      const clave = cohorte || "Total";
      if (!porCohorte[clave]) porCohorte[clave] = {};
      porCohorte[clave][esc] = registro;
    }
  }

  if (filasLeidas === 0) throw new Error("La base no contiene registros con escuela identificable.");
  const bloques = Object.keys(porBloque);
  if (bloques.length === 0) throw new Error("No se encontraron filas con bloque (período). Verifique la columna «Bloque».");

  return {
    porBloque, porCohorte, bloques, cohortes: Object.keys(porCohorte),
    filasLeidas, sinClasificar, filasSubtotal, esMetas: true,
    nuevosDetectados, antiguosDetectados,
    ejecutadoDetectado: columnas.NUEVOS.ejecutado !== undefined || columnas.ANTIGUOS.ejecutado !== undefined,
  };
}

/* Estima los créditos de una matrícula según el nivel del programa, usando
   promedios: pregrado 12, posgrado (especialización/maestría/doctorado) 7,
   INVIL (lenguas) 1. Es una ESTIMACIÓN para el escenario "si se pagaran": el
   reporte de pendientes no trae créditos reales por estudiante. */
const CRED_PREGRADO = 12, CRED_POSGRADO = 7, CRED_INVIL = 1;
function creditosEstimados(programa, escuela) {
  const p = sinTildes(programa || "");
  if (escuela === "INVIL") return CRED_INVIL;
  if (p.includes("ESPECIALIZAC") || p.includes("MAESTRIA") || p.includes("DOCTORADO")) return CRED_POSGRADO;
  return CRED_PREGRADO;
}

function agregarBase(filas, tipo) {
  if (tipo === "metas") return agregarBaseMetas(filas);
  const hIdx = detectarEncabezado(filas, tipo);
  if (hIdx < 0) throw new Error("No se encontró la fila de encabezados. Verifique que la base incluya una columna «Escuela» y una de identificación del estudiante (Documento / ID).");
  const enc = filas[hIdx];
  const colEscuela = indiceColumna(enc, [(c) => c === "ESCUELA", (c) => c.includes("ESCUELA") && !c.includes("CODIGO")]);
  const colDoc = indiceColumna(enc, [(c) => c.includes("DOCUMENTO"), (c) => c === "ID"]);
  const colCred = indiceColumna(enc, [(c) => c.includes("CREDITOS TOTALES"), (c) => c === "CREDITOS", (c) => c.includes("CREDITOS")]);
  const colFactura = indiceColumna(enc, [(c) => c === "FACTURA", (c) => c.includes("FACTURA") && !c.includes("FECHA")]);
  const colValor = indiceColumna(enc, [(c) => c.includes("VALOR")]);
  const colCond = indiceColumna(enc, [(c) => c.includes("CONDICION"), (c) => c.includes("TIPO ESTUDIANTE"), (c) => c === "ESTADO"]);
  const colPrograma = indiceColumna(enc, [(c) => c.includes("PROGRAMA")]);

  const porEscuela = {};
  const docsVistos = {};
  let filasLeidas = 0, sinClasificar = 0;

  for (let i = hIdx + 1; i < filas.length; i++) {
    const fila = filas[i];
    if (!fila || fila.every((c) => c === null || c === undefined || c === "")) continue;
    const esc = normalizarEscuela(fila[colEscuela]);
    if (!esc) continue;
    filasLeidas++;
    if (esc === "SIN_CLASIFICAR") { sinClasificar++; continue; }
    if (!porEscuela[esc]) porEscuela[esc] = { estudiantes: 0, creditos: 0, facturas: 0, valor: 0, nuevos: 0, credPendNuevos: 0, credPendAntiguos: 0, estPendNuevos: 0, estPendAntiguos: 0 };
    if (tipo === "pagadas") {
      const doc = colDoc >= 0 ? String(fila[colDoc] ?? "").trim() : "";
      const clave = esc + "|" + (doc || `fila${i}`);
      const esNuevo = colCond >= 0 && sinTildes(fila[colCond] ?? "").includes("NUEVO");
      if (!docsVistos[clave]) {
        docsVistos[clave] = true;
        porEscuela[esc].estudiantes++;
        if (esNuevo) porEscuela[esc].nuevos++;
      }
      const cr = colCred >= 0 ? aNumero(fila[colCred]) : null;
      if (cr !== null) porEscuela[esc].creditos += cr;
    } else {
      // Pendientes: cuenta estudiantes únicos (no facturas, un estudiante puede
      // tener varias) y estima los créditos por recuperar según el nivel.
      porEscuela[esc].facturas++;
      const val = colValor >= 0 ? aNumero(fila[colValor]) : null;
      if (val !== null) porEscuela[esc].valor += val;
      const doc = colDoc >= 0 ? String(fila[colDoc] ?? "").trim() : "";
      const clave = esc + "|" + (doc || `fila${i}`);
      if (!docsVistos[clave]) {
        docsVistos[clave] = true;
        porEscuela[esc].estudiantes++;
        const esNuevo = colCond >= 0 && sinTildes(fila[colCond] ?? "").includes("NUEVO");
        const cr = creditosEstimados(colPrograma >= 0 ? fila[colPrograma] : "", esc);
        if (esNuevo) { porEscuela[esc].nuevos++; porEscuela[esc].estPendNuevos++; porEscuela[esc].credPendNuevos += cr; }
        else { porEscuela[esc].estPendAntiguos++; porEscuela[esc].credPendAntiguos += cr; }
      }
    }
  }
  if (filasLeidas === 0) throw new Error("La base no contiene registros con escuela identificable.");
  return { porEscuela, filasLeidas, sinClasificar, creditosDetectados: colCred >= 0, condicionDetectada: colCond >= 0, programaDetectado: colPrograma >= 0 };
}

async function leerArchivo(file) {
  const buf = await file.arrayBuffer();
  // cellDates convierte los seriales de fecha de Excel en fechas reales.
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const hoja = wb.Sheets[wb.SheetNames[0]];
  return {
    filas: XLSX.utils.sheet_to_json(hoja, { header: 1, defval: null, raw: true }),
    nombreHoja: wb.SheetNames[0],
  };
}

/* ============================ APP ============================ */
export default function App() {
  const [estado, setEstado] = useState(null);
  const [cargando, setCargando] = useState(true);
  const [anio, setAnio] = useState("2026");
  const [vista, setVista] = useState("periodo");
  const [periodoIdx, setPeriodoIdx] = useState(0);
  const [aviso, setAviso] = useState(null);

  useEffect(() => {
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = "https://fonts.googleapis.com/css2?family=Archivo:wght@400;500;600;700;800&display=swap";
    document.head.appendChild(link);
  }, []);

  useEffect(() => {
    (async () => {
      let inicial = estadoInicial();
      try {
        const r = await storage.get(KEY);
        if (r && r.value) inicial = JSON.parse(r.value);
      } catch (e) { /* sin datos guardados: se usa la información del archivo base */ }
      setEstado(inicial);
      setCargando(false);
    })();
  }, []);

  const guardar = async (nuevo) => {
    setEstado(nuevo);
    try { await storage.set(KEY, JSON.stringify(nuevo)); }
    catch (e) { setAviso({ tipo: "error", texto: "No fue posible guardar los cambios. Se conservan solo en esta sesión." }); }
  };

  useEffect(() => {
    if (!aviso) return;
    const t = setTimeout(() => setAviso(null), 6000);
    return () => clearTimeout(t);
  }, [aviso]);

  if (cargando || !estado) {
    return (
      <div style={{ minHeight: "100vh", background: C.fondo, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "Archivo, system-ui, sans-serif", color: C.gris }}>
        Cargando seguimiento…
      </div>
    );
  }

  const anios = Object.keys(estado.anios).sort();
  const datosAnio = estado.anios[anio];
  const periodos = datosAnio ? datosAnio.periodos : [];
  const periodo = periodos[Math.min(periodoIdx, periodos.length - 1)] || null;

  const estilos = { fontFamily: "Archivo, system-ui, sans-serif", color: C.tinta };

  return (
    <div style={{ minHeight: "100vh", background: C.fondo, ...estilos }}>
      <style>{`
        .num { font-variant-numeric: tabular-nums; }
        table.seg th, table.seg td { padding: 7px 10px; }
        table.seg.compacta th, table.seg.compacta td { padding: 5px 4px; font-size: 12.5px; }
        table.seg.compacta td:first-child, table.seg.compacta th:first-child { padding-left: 2px; }
        @media (max-width: 900px) {
          table.seg.compacta th, table.seg.compacta td { padding: 4px 2px; font-size: 11.5px; }
        }
        input.campo { width: 76px; padding: 4px 6px; border: 1px solid ${C.borde}; border-radius: 6px; text-align: right; font: inherit; }
        input.campo:focus { outline: 2px solid ${C.azul}; outline-offset: 0; }
        button.tab { border: none; background: transparent; padding: 10px 14px; font: inherit; font-weight: 600; color: #B9C6DC; cursor: pointer; border-bottom: 3px solid transparent; }
        button.tab.activa { color: #fff; border-color: ${C.ambar}; }
        button.tab:focus-visible, .btn:focus-visible { outline: 2px solid ${C.ambar}; outline-offset: 2px; }
        .btn { border: none; border-radius: 8px; padding: 9px 16px; font: inherit; font-weight: 600; cursor: pointer; }
      `}</style>

      {/* -------- encabezado -------- */}
      <header style={{ background: C.tinta, color: "#fff" }}>
        <div style={{ maxWidth: 1120, margin: "0 auto", padding: "20px 20px 0" }}>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 12, alignItems: "baseline", justifyContent: "space-between" }}>
            <div>
              <div style={{ fontSize: 11, letterSpacing: "0.18em", fontWeight: 600, color: C.ambar }}>UNAD · ZONA SUR · UDR VALLE DEL GUAMUEZ</div>
              <h1 style={{ margin: "4px 0 2px", fontSize: 26, fontWeight: 800, letterSpacing: "-0.01em" }}>Seguimiento de matrículas</h1>
              <div style={{ fontSize: 13, color: "#B9C6DC" }}>Créditos y estudiantes por período académico y consolidado anual</div>
            </div>
            <label style={{ fontSize: 13, color: "#B9C6DC" }}>
              Año&nbsp;
              <select
                value={anio}
                onChange={(e) => { setAnio(e.target.value); setPeriodoIdx(0); }}
                style={{ font: "inherit", fontWeight: 700, padding: "6px 10px", borderRadius: 8, border: "none", background: "#1B3357", color: "#fff" }}
              >
                {anios.map((a) => <option key={a} value={a}>{a}</option>)}
              </select>
            </label>
          </div>
          <nav style={{ marginTop: 14, display: "flex", gap: 4, flexWrap: "wrap" }}>
            {[["periodo", "Seguimiento por período"], ["anual", "Consolidado anual"], ["cargar", "Cargar bases de datos"], ["metas", "Metas"]].map(([id, tx]) => (
              <button key={id} className={`tab ${vista === id ? "activa" : ""}`} onClick={() => setVista(id)}>{tx}</button>
            ))}
          </nav>
        </div>
      </header>

      {aviso && (
        <div style={{ maxWidth: 1120, margin: "12px auto 0", padding: "0 20px" }}>
          <div style={{ background: aviso.tipo === "error" ? "#FBEAE6" : "#E7F4EC", border: `1px solid ${aviso.tipo === "error" ? C.rojo : C.verde}`, color: aviso.tipo === "error" ? C.rojo : C.verde, borderRadius: 8, padding: "10px 14px", fontSize: 14, fontWeight: 600 }}>
            {aviso.texto}
          </div>
        </div>
      )}

      <main style={{ maxWidth: 1120, margin: "0 auto", padding: "20px 20px 60px" }}>
        {vista === "periodo" && periodo && (
          <VistaPeriodo periodos={periodos} periodoIdx={Math.min(periodoIdx, periodos.length - 1)} setPeriodoIdx={setPeriodoIdx} periodo={periodo} anio={anio} />
        )}
        {vista === "anual" && <VistaAnual periodos={periodos} anio={anio} />}
        {vista === "cargar" && (
          <VistaCargar estado={estado} guardar={guardar} anio={anio} setAviso={setAviso} />
        )}
        {vista === "metas" && (
          <VistaMetas estado={estado} guardar={guardar} anio={anio} setAviso={setAviso} setAnio={setAnio} />
        )}
      </main>
    </div>
  );
}

/* ---------------- componentes de tabla ---------------- */
function Tarjeta({ children, titulo, extra }) {
  return (
    <section style={{ background: "#fff", border: `1px solid ${C.borde}`, borderRadius: 12, padding: 18, marginBottom: 18 }}>
      {(titulo || extra) && (
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10, marginBottom: 12 }}>
          {titulo && <h2 style={{ margin: 0, fontSize: 16, fontWeight: 700 }}>{titulo}</h2>}
          {extra}
        </div>
      )}
      {children}
    </section>
  );
}

function filaTotales(escuelas) {
  const t = { metaCred: 0, avCred: 0, metaEst: 0, avEst: 0, factPend: 0, avEstNuevos: null, metaEstNuevos: null, metaEstAntiguos: null, metaCredNuevos: null, metaCredAntiguos: null, avCredNuevos: null, avCredAntiguos: null, credPendNuevos: null, credPendAntiguos: null, estPendNuevos: null, estPendAntiguos: null, valorPend: null };
  for (const e of ESCUELAS) {
    const d = escuelas[e] || {};
    t.metaCred += num(d.metaCred) || 0;
    t.avCred += num(d.avCred) || 0;
    t.metaEst += num(d.metaEst) || 0;
    t.avEst += num(d.avEst) || 0;
    t.factPend += num(d.factPend) || 0;
    for (const campo of ["avEstNuevos", "metaEstNuevos", "metaEstAntiguos", "metaCredNuevos", "metaCredAntiguos", "avCredNuevos", "avCredAntiguos", "credPendNuevos", "credPendAntiguos", "estPendNuevos", "estPendAntiguos", "valorPend"]) {
      const v = num(d[campo]);
      if (v !== null) t[campo] = (t[campo] || 0) + v;
    }
  }
  return t;
}

function avanceAntiguos(d) {
  const av = num(d.avEst), nv = num(d.avEstNuevos);
  return av !== null && nv !== null ? av - nv : null;
}

function TablaSeguimiento({ escuelas }) {
  const tot = filaTotales(escuelas);
  const celda = { borderBottom: `1px solid ${C.borde}` };
  const der = { ...celda, textAlign: "right" };
  const Fila = ({ nombre, d, negrita }) => {
    const pc = pct(d.avCred, d.metaCred);
    const pe = pct(d.avEst, d.metaEst);
    const pn = pct(d.avEstNuevos, d.metaEstNuevos);
    const avAnt = avanceAntiguos(d);
    const pa = pct(avAnt, d.metaEstAntiguos);
    const st = negrita ? { fontWeight: 800, background: C.azulSuave } : {};
    return (
      <tr style={st}>
        <td style={{ ...celda, fontWeight: negrita ? 800 : 600 }}>{nombre}</td>
        <td className="num" style={der}>{fmt(d.metaCred)}</td>
        <td className="num" style={der}>{fmt(d.avCred)}</td>
        <td className="num" style={{ ...der, color: colorPct(pc), fontWeight: 700 }}>{fmtPct(pc)}</td>
        <td className="num" style={{ ...der, borderLeft: `1px solid ${C.borde}` }}>{fmtEst(d.metaEst)}</td>
        <td className="num" style={der}>{fmt(d.avEst)}</td>
        <td className="num" style={{ ...der, color: colorPct(pe), fontWeight: 700 }}>{fmtPct(pe)}</td>
        <td className="num" style={{ ...der, borderLeft: `1px solid ${C.borde}` }}>{fmtEst(d.metaEstNuevos)}</td>
        <td className="num" style={{ ...der, color: C.azul, fontWeight: 700 }}>{fmt(d.avEstNuevos)}</td>
        <td className="num" style={{ ...der, color: colorPct(pn), fontWeight: 700 }}>{fmtPct(pn)}</td>
        <td className="num" style={{ ...der, borderLeft: `1px solid ${C.borde}` }}>{fmtEst(d.metaEstAntiguos)}</td>
        <td className="num" style={der}>{fmt(avAnt)}</td>
        <td className="num" style={{ ...der, color: colorPct(pa), fontWeight: 700 }}>{fmtPct(pa)}</td>
        <td className="num" style={{ ...der, borderLeft: `1px solid ${C.borde}`, color: (d.factPend || 0) > 0 ? C.rojo : C.gris, fontWeight: (d.factPend || 0) > 0 ? 700 : 400 }}>{fmt(d.factPend)}</td>
      </tr>
    );
  };
  return (
    <div style={{ overflowX: "auto" }}>
      <table className="seg compacta" style={{ borderCollapse: "collapse", width: "100%" }}>
        <thead>
          <tr style={{ fontSize: 10, letterSpacing: "0.06em", color: C.gris }}>
            <th></th>
            <th colSpan={3} style={{ textAlign: "center", borderBottom: `2px solid ${C.azul}`, paddingBottom: 3 }}>CRÉDITOS</th>
            <th colSpan={3} style={{ textAlign: "center", borderBottom: `2px solid ${C.ambar}`, paddingBottom: 3 }}>ESTUDIANTES</th>
            <th colSpan={3} style={{ textAlign: "center", borderBottom: `2px solid ${C.verde}`, paddingBottom: 3 }}>NUEVOS</th>
            <th colSpan={3} style={{ textAlign: "center", borderBottom: `2px solid ${C.gris}`, paddingBottom: 3 }}>ANTIGUOS</th>
            <th style={{ textAlign: "center", borderBottom: `2px solid ${C.rojo}`, paddingBottom: 3 }}>FACT.</th>
          </tr>
          <tr style={{ fontSize: 11, color: C.gris, textAlign: "right" }}>
            <th style={{ textAlign: "left" }}>Escuela</th>
            <th>Meta</th><th>Avance</th><th>%</th>
            <th>Meta</th><th>Avance</th><th>%</th>
            <th>Meta</th><th>Avance</th><th>%</th>
            <th>Meta</th><th>Avance</th><th>%</th>
            <th>S/pagar</th>
          </tr>
        </thead>
        <tbody>
          {ESCUELAS.map((e) => <Fila key={e} nombre={e} d={escuelas[e] || {}} />)}
          <Fila nombre="TOTAL" d={tot} negrita />
        </tbody>
      </table>
    </div>
  );
}

function GraficaCumplimiento({ escuelas, titulo }) {
  const datos = ESCUELAS.map((e) => {
    const d = escuelas[e] || {};
    const p = pct(d.avCred, d.metaCred);
    return { escuela: e, valor: p === null ? 0 : Math.round(p * 10) / 10 };
  }).filter((d) => d.valor > 0);
  if (datos.length === 0) return null;
  const max = Math.max(150, ...datos.map((d) => d.valor));
  return (
    <div>
      <div style={{ fontSize: 12, color: C.gris, marginBottom: 4 }}>{titulo}</div>
      <ResponsiveContainer width="100%" height={230}>
        <BarChart data={datos} margin={{ top: 18, right: 10, left: -14, bottom: 0 }}>
          <XAxis dataKey="escuela" tick={{ fontSize: 11, fill: C.gris }} axisLine={{ stroke: C.borde }} tickLine={false} />
          <YAxis domain={[0, Math.ceil(max / 50) * 50]} tick={{ fontSize: 11, fill: C.gris }} axisLine={false} tickLine={false} unit="%" />
          <Tooltip formatter={(v) => [`${v}%`, "% de la meta (créditos)"]} contentStyle={{ fontFamily: "inherit", fontSize: 12, borderRadius: 8 }} />
          <ReferenceLine y={100} stroke={C.tinta} strokeDasharray="4 4" />
          <Bar dataKey="valor" radius={[4, 4, 0, 0]}>
            <LabelList dataKey="valor" position="top" formatter={(v) => `${v}%`} style={{ fontSize: 10, fill: C.tinta, fontWeight: 600 }} />
            {datos.map((d, i) => <Cell key={i} fill={d.valor >= 100 ? C.verde : d.valor >= 70 ? C.ambar : C.rojo} />)}
          </Bar>
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}

/* ---------------- vista: período ---------------- */
function EscenarioRecuperacion({ escuelas, promedio }) {
  const [abierto, setAbierto] = useState(false);
  const hayDatos = ESCUELAS.some((e) => (num((escuelas[e] || {}).factPend) || 0) > 0);
  if (!hayDatos) return null;
  const tot = filaTotales(escuelas);

  // Créditos por recuperar por escuela (estimados) y % potencial de la meta.
  const fmtMoneda = (v) => (v === null || v === undefined ? "—" : "$ " + fmt(Math.round(v)));
  const celda = { borderBottom: `1px solid ${C.borde}` };
  const der = { ...celda, textAlign: "right" };

  const Fila = ({ nombre, d, negrita }) => {
    const credPend = (num(d.credPendNuevos) || 0) + (num(d.credPendAntiguos) || 0);
    const avActual = num(d.avCred);
    const meta = num(d.metaCred);
    const pActual = pct(avActual, meta);
    const pPotencial = pct((avActual || 0) + credPend, meta);
    const estPend = (num(d.estPendNuevos) || 0) + (num(d.estPendAntiguos) || 0);
    const st = negrita ? { fontWeight: 800, background: C.azulSuave } : {};
    return (
      <tr style={st}>
        <td style={{ ...celda, fontWeight: negrita ? 800 : 600 }}>{nombre}</td>
        <td className="num" style={der}>{fmt(estPend)}</td>
        <td className="num" style={{ ...der, color: C.azul }}>{fmt(d.estPendNuevos)}</td>
        <td className="num" style={der}>{fmt(d.estPendAntiguos)}</td>
        <td className="num" style={{ ...der, borderLeft: `1px solid ${C.borde}`, fontWeight: 700 }}>{credPend ? fmt(credPend) : "—"}</td>
        <td className="num" style={{ ...der, color: C.gris }}>{fmtPct(pActual)}</td>
        <td className="num" style={{ ...der, color: colorPct(pPotencial), fontWeight: 800 }}>{fmtPct(pPotencial)}</td>
        <td className="num" style={{ ...der, borderLeft: `1px solid ${C.borde}`, color: C.gris }}>{fmtMoneda(d.valorPend)}</td>
      </tr>
    );
  };

  const credTot = (tot.credPendNuevos || 0) + (tot.credPendAntiguos || 0);
  const pActualT = pct(tot.avCred, tot.metaCred);
  const pPotT = pct((tot.avCred || 0) + credTot, tot.metaCred);

  return (
    <div style={{ marginTop: 14 }}>
      <button className="btn" onClick={() => setAbierto(!abierto)}
        style={{ background: abierto ? C.tinta : "#fff", color: abierto ? "#fff" : C.tinta, border: `1px solid ${C.borde}`, fontSize: 13 }}>
        {abierto ? "▾ " : "▸ "}Escenario: ¿y si se pagaran las facturas pendientes?
      </button>
      {abierto && (
        <div style={{ marginTop: 12 }}>
          <div style={{ background: "#FFF8E6", border: `1px solid ${C.ambar}`, borderLeft: `4px solid ${C.ambar}`, borderRadius: 8, padding: "10px 14px", fontSize: 12.5, color: C.tinta, marginBottom: 12 }}>
            Escenario hipotético, no es avance real. Muestra cómo quedaría el cumplimiento de créditos si se recuperaran
            todas las matrículas pendientes de pago. Los créditos se estiman por nivel (pregrado 12, posgrado 7, INVIL 1 en
            promedio), porque el reporte de pendientes no trae créditos por estudiante.
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 12, marginBottom: 14 }}>
            <Indicador etiqueta="Estudiantes pendientes" valor={fmt(tot.factPend)} color={C.rojo} sub={`${fmt(tot.estPendNuevos)} nuevos · ${fmt(tot.estPendAntiguos)} antiguos`} />
            <Indicador etiqueta="Créditos por recuperar" valor={credTot ? fmt(credTot) : "—"} color={C.tinta} sub="Estimados por nivel de programa" />
            <Indicador etiqueta="Cumplimiento potencial" valor={fmtPct(pPotT)} color={colorPct(pPotT)} sub={pActualT !== null ? `Hoy: ${fmtPct(pActualT)}` : "Cargue la base de pagadas"} />
            <Indicador etiqueta="Valor pendiente" valor={tot.valorPend ? "$ " + fmt(tot.valorPend) : "—"} color={C.tinta} sub="Monto total de facturas sin pagar" />
          </div>
          <div style={{ overflowX: "auto" }}>
            <table className="seg compacta" style={{ borderCollapse: "collapse", width: "100%" }}>
              <thead>
                <tr style={{ fontSize: 10, letterSpacing: "0.06em", color: C.gris }}>
                  <th></th>
                  <th colSpan={3} style={{ textAlign: "center", borderBottom: `2px solid ${C.rojo}`, paddingBottom: 3 }}>ESTUDIANTES PENDIENTES</th>
                  <th colSpan={3} style={{ textAlign: "center", borderBottom: `2px solid ${C.azul}`, paddingBottom: 3 }}>CRÉDITOS (ESTIMADO)</th>
                  <th style={{ textAlign: "center", borderBottom: `2px solid ${C.gris}`, paddingBottom: 3 }}>VALOR</th>
                </tr>
                <tr style={{ fontSize: 11, color: C.gris, textAlign: "right" }}>
                  <th style={{ textAlign: "left" }}>Escuela</th>
                  <th>Total</th><th>Nuevos</th><th>Antiguos</th>
                  <th>Por recuperar</th><th>% hoy</th><th>% potencial</th>
                  <th>Pendiente $</th>
                </tr>
              </thead>
              <tbody>
                {ESCUELAS.map((e) => <Fila key={e} nombre={e} d={escuelas[e] || {}} />)}
                <Fila nombre="TOTAL" d={tot} negrita />
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

function DetalleCreditos({ escuelas }) {
  const [abierto, setAbierto] = useState(false);
  const hayDatos = ESCUELAS.some((e) => num((escuelas[e] || {}).metaCredNuevos) !== null || num((escuelas[e] || {}).metaCredAntiguos) !== null);
  if (!hayDatos) return null;
  const tot = filaTotales(escuelas);
  return (
    <div style={{ marginTop: 10 }}>
      <button
        className="btn"
        onClick={() => setAbierto(!abierto)}
        style={{ background: "transparent", color: C.azul, padding: "4px 0", fontSize: 12.5, textDecoration: "underline" }}
      >
        {abierto ? "Ocultar" : "Ver"} metas en créditos (nuevos y antiguos)
      </button>
      {abierto && (
        <div style={{ overflowX: "auto", marginTop: 6 }}>
          <table className="seg compacta" style={{ borderCollapse: "collapse", width: "100%" }}>
            <thead>
              <tr style={{ fontSize: 11, color: C.gris, textAlign: "right" }}>
                <th style={{ textAlign: "left" }}>Escuela</th>
                <th>Meta cr. nuevos</th><th>Ejecutado cr.</th>
                <th>Meta cr. antiguos</th><th>Ejecutado cr.</th>
              </tr>
            </thead>
            <tbody>
              {ESCUELAS.map((e) => {
                const d = escuelas[e] || {};
                return (
                  <tr key={e}>
                    <td style={{ borderBottom: `1px solid ${C.borde}`, fontWeight: 600 }}>{e}</td>
                    <td className="num" style={{ borderBottom: `1px solid ${C.borde}`, textAlign: "right" }}>{fmt(d.metaCredNuevos)}</td>
                    <td className="num" style={{ borderBottom: `1px solid ${C.borde}`, textAlign: "right", color: C.gris }}>{fmt(d.avCredNuevos)}</td>
                    <td className="num" style={{ borderBottom: `1px solid ${C.borde}`, textAlign: "right" }}>{fmt(d.metaCredAntiguos)}</td>
                    <td className="num" style={{ borderBottom: `1px solid ${C.borde}`, textAlign: "right", color: C.gris }}>{fmt(d.avCredAntiguos)}</td>
                  </tr>
                );
              })}
              <tr style={{ fontWeight: 800, background: C.azulSuave }}>
                <td>TOTAL</td>
                <td className="num" style={{ textAlign: "right" }}>{fmt(tot.metaCredNuevos)}</td>
                <td className="num" style={{ textAlign: "right" }}>{fmt(tot.avCredNuevos)}</td>
                <td className="num" style={{ textAlign: "right" }}>{fmt(tot.metaCredAntiguos)}</td>
                <td className="num" style={{ textAlign: "right" }}>{fmt(tot.avCredAntiguos)}</td>
              </tr>
            </tbody>
          </table>
          <div style={{ fontSize: 11.5, color: C.gris, marginTop: 6 }}>
            Cifras tal como vienen del reporte institucional. Las metas en estudiantes de la tabla principal se obtienen
            dividiendo estas cifras entre los créditos promedio del período.
          </div>
        </div>
      )}
    </div>
  );
}

function VistaPeriodo({ periodos, periodoIdx, setPeriodoIdx, periodo, anio }) {
  const tot = filaTotales(periodo.escuelas);
  const pc = pct(tot.avCred, tot.metaCred);
  return (
    <>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 16 }}>
        {periodos.map((p, i) => (
          <button
            key={p.id}
            className="btn"
            onClick={() => setPeriodoIdx(i)}
            style={{
              background: i === periodoIdx ? C.azul : "#fff",
              color: i === periodoIdx ? "#fff" : C.tinta,
              border: `1px solid ${i === periodoIdx ? C.azul : C.borde}`,
            }}
          >
            {p.nombre}
          </button>
        ))}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(170px, 1fr))", gap: 12, marginBottom: 18 }}>
        <Indicador etiqueta="Créditos matriculados" valor={fmt(tot.avCred)} sub={`Meta: ${fmt(tot.metaCred)}`} />
        <Indicador etiqueta="% meta en créditos" valor={fmtPct(pc)} color={colorPct(pc)} sub={`Período ${periodo.nombre} · ${anio}`} />
        <Indicador etiqueta="Estudiantes matriculados" valor={fmt(tot.avEst)} sub={`Meta: ${fmtEst(tot.metaEst)} (${periodo.promedio || "—"} cr. promedio)`} />
        <Indicador
          etiqueta="Estudiantes nuevos"
          valor={fmt(tot.avEstNuevos)}
          color={C.azul}
          sub={(() => {
            const partes = [];
            if (tot.metaEstNuevos !== null) partes.push(`Meta nuevos: ${fmtEst(tot.metaEstNuevos)}`);
            if (tot.avEstNuevos !== null && tot.avEst > 0) partes.push(`${fmt(tot.avEst - tot.avEstNuevos)} antiguos${tot.metaEstAntiguos !== null ? ` (meta: ${fmtEst(tot.metaEstAntiguos)})` : ""}`);
            return partes.length ? partes.join(" · ") : "Se identifica al subir la base de pagadas";
          })()}
        />
        <Indicador etiqueta="Facturas sin pagar" valor={fmt(tot.factPend)} color={tot.factPend > 0 ? C.rojo : C.verde} sub="Oportunidad de recuperación" />
      </div>

      <Tarjeta
        titulo={`Período ${anio} (${periodo.nombre})`}
        extra={<Fuentes periodo={periodo} />}
      >
        <TablaSeguimiento escuelas={periodo.escuelas} />
        <DetalleCreditos escuelas={periodo.escuelas} />
        <EscenarioRecuperacion escuelas={periodo.escuelas} promedio={periodo.promedio} />
        <div style={{ marginTop: 18 }}>
          <GraficaCumplimiento escuelas={periodo.escuelas} titulo="Cumplimiento de la meta de créditos por escuela (línea punteada = 100 %)" />
        </div>
      </Tarjeta>
    </>
  );
}

function Fuentes({ periodo }) {
  const f = [];
  if (periodo.fuentePagadas) f.push(`Pagadas: ${periodo.fuentePagadas.archivo} (${periodo.fuentePagadas.fecha})`);
  if (periodo.fuentePendientes) f.push(`Pendientes: ${periodo.fuentePendientes.archivo} (${periodo.fuentePendientes.fecha})`);
  if (periodo.fuenteMetas) f.push(`Metas: ${periodo.fuenteMetas.archivo} (${periodo.fuenteMetas.fecha})`);
  if (f.length === 0) return <span style={{ fontSize: 12, color: C.gris }}>Avance del archivo base — actualice subiendo las bases del período</span>;
  return <span style={{ fontSize: 12, color: C.gris }}>{f.join(" · ")}</span>;
}

function Indicador({ etiqueta, valor, sub, color }) {
  return (
    <div style={{ background: "#fff", border: `1px solid ${C.borde}`, borderRadius: 12, padding: "14px 16px" }}>
      <div style={{ fontSize: 11, letterSpacing: "0.06em", color: C.gris, textTransform: "uppercase", fontWeight: 600 }}>{etiqueta}</div>
      <div className="num" style={{ fontSize: 26, fontWeight: 800, color: color || C.tinta, margin: "2px 0" }}>{valor}</div>
      {sub && <div style={{ fontSize: 12, color: C.gris }}>{sub}</div>}
    </div>
  );
}

/* ---------------- vista: anual ---------------- */
function VistaAnual({ periodos, anio }) {
  const acumulado = useMemo(() => {
    const esc = Object.fromEntries(ESCUELAS.map((e) => [e, { metaCred: 0, avCred: 0, metaEst: 0, avEst: 0, factPend: 0, avEstNuevos: null, metaEstNuevos: null, metaEstAntiguos: null, metaCredNuevos: null, metaCredAntiguos: null, avCredNuevos: null, avCredAntiguos: null, credPendNuevos: null, credPendAntiguos: null, estPendNuevos: null, estPendAntiguos: null, valorPend: null }]));
    for (const p of periodos) {
      for (const e of ESCUELAS) {
        const d = p.escuelas[e] || {};
        esc[e].metaCred += num(d.metaCred) || 0;
        esc[e].avCred += num(d.avCred) || 0;
        esc[e].metaEst += num(d.metaEst) || 0;
        esc[e].avEst += num(d.avEst) || 0;
        esc[e].factPend += num(d.factPend) || 0;
        for (const campo of ["avEstNuevos", "metaEstNuevos", "metaEstAntiguos", "metaCredNuevos", "metaCredAntiguos", "avCredNuevos", "avCredAntiguos", "credPendNuevos", "credPendAntiguos", "estPendNuevos", "estPendAntiguos", "valorPend"]) {
          const v = num(d[campo]);
          if (v !== null) esc[e][campo] = (esc[e][campo] || 0) + v;
        }
      }
    }
    return esc;
  }, [periodos]);

  const tot = filaTotales(acumulado);
  const pc = pct(tot.avCred, tot.metaCred);
  const pe = pct(tot.avEst, tot.metaEst);

  return (
    <>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(170px, 1fr))", gap: 12, marginBottom: 18 }}>
        <Indicador etiqueta={`Créditos ${anio}`} valor={fmt(tot.avCred)} sub={`Meta anual: ${fmt(tot.metaCred)}`} />
        <Indicador etiqueta="% meta anual (créditos)" valor={fmtPct(pc)} color={colorPct(pc)} sub={pc !== null && pc < 100 ? `Faltan ${fmt(tot.metaCred - tot.avCred)} créditos` : "Meta anual superada"} />
        <Indicador etiqueta={`Estudiantes ${anio}`} valor={fmt(tot.avEst)} sub={`Meta anual: ${fmtEst(tot.metaEst)}`} />
        <Indicador etiqueta="% meta anual (estudiantes)" valor={fmtPct(pe)} color={colorPct(pe)} sub={pe !== null && pe < 100 ? `Faltan ${fmt(Math.max(0, Math.round(tot.metaEst - tot.avEst)))} estudiantes` : "Meta anual superada"} />
        <Indicador etiqueta={`Estudiantes nuevos ${anio}`} valor={fmt(tot.avEstNuevos)} color={C.azul} sub={tot.metaEstNuevos !== null ? `Meta anual de nuevos: ${fmtEst(tot.metaEstNuevos)}` : "Suma de los períodos con dato registrado"} />
      </div>

      <Tarjeta titulo={`Consolidado ${anio} — suma de los ${periodos.length} períodos`}>
        <TablaAnual acumulado={acumulado} />
        <DetalleCreditos escuelas={acumulado} />
        <EscenarioRecuperacion escuelas={acumulado} />
        <div style={{ marginTop: 18 }}>
          <GraficaCumplimiento escuelas={acumulado} titulo="Cumplimiento anual de la meta de créditos por escuela" />
        </div>
      </Tarjeta>
    </>
  );
}

function TablaAnual({ acumulado }) {
  const tot = filaTotales(acumulado);
  const celda = { borderBottom: `1px solid ${C.borde}` };
  const der = { ...celda, textAlign: "right" };
  const Fila = ({ nombre, d, negrita }) => {
    const p = pct(d.avCred, d.metaCred);
    const pe = pct(d.avEst, d.metaEst);
    const pn = pct(d.avEstNuevos, d.metaEstNuevos);
    const avAnt = avanceAntiguos(d);
    const pa = pct(avAnt, d.metaEstAntiguos);
    const faltC = d.metaCred ? d.metaCred - d.avCred : null;
    const st = negrita ? { fontWeight: 800, background: C.azulSuave } : {};
    return (
      <tr style={st}>
        <td style={{ ...celda, fontWeight: negrita ? 800 : 600 }}>{nombre}</td>
        <td className="num" style={der}>{fmt(d.metaCred)}</td>
        <td className="num" style={der}>{fmt(d.avCred)}</td>
        <td className="num" style={{ ...der, color: colorPct(p), fontWeight: 700 }}>{fmtPct(p)}</td>
        <td className="num" style={{ ...der, color: faltC !== null && faltC > 0 ? C.rojo : C.verde }}>{faltC === null ? "—" : fmt(faltC)}</td>
        <td className="num" style={{ ...der, borderLeft: `1px solid ${C.borde}` }}>{fmtEst(d.metaEst)}</td>
        <td className="num" style={der}>{fmt(d.avEst)}</td>
        <td className="num" style={{ ...der, color: colorPct(pe), fontWeight: 700 }}>{fmtPct(pe)}</td>
        <td className="num" style={{ ...der, borderLeft: `1px solid ${C.borde}` }}>{fmtEst(d.metaEstNuevos)}</td>
        <td className="num" style={{ ...der, color: C.azul, fontWeight: 700 }}>{fmt(d.avEstNuevos)}</td>
        <td className="num" style={{ ...der, color: colorPct(pn), fontWeight: 700 }}>{fmtPct(pn)}</td>
        <td className="num" style={{ ...der, borderLeft: `1px solid ${C.borde}` }}>{fmtEst(d.metaEstAntiguos)}</td>
        <td className="num" style={der}>{fmt(avAnt)}</td>
        <td className="num" style={{ ...der, color: colorPct(pa), fontWeight: 700 }}>{fmtPct(pa)}</td>
        <td className="num" style={{ ...der, borderLeft: `1px solid ${C.borde}`, color: (d.factPend || 0) > 0 ? C.rojo : C.gris }}>{fmt(d.factPend)}</td>
      </tr>
    );
  };
  return (
    <div style={{ overflowX: "auto" }}>
      <table className="seg compacta" style={{ borderCollapse: "collapse", width: "100%" }}>
        <thead>
          <tr style={{ fontSize: 10, letterSpacing: "0.06em", color: C.gris }}>
            <th></th>
            <th colSpan={4} style={{ textAlign: "center", borderBottom: `2px solid ${C.azul}`, paddingBottom: 3 }}>CRÉDITOS</th>
            <th colSpan={3} style={{ textAlign: "center", borderBottom: `2px solid ${C.ambar}`, paddingBottom: 3 }}>ESTUDIANTES</th>
            <th colSpan={3} style={{ textAlign: "center", borderBottom: `2px solid ${C.verde}`, paddingBottom: 3 }}>NUEVOS</th>
            <th colSpan={3} style={{ textAlign: "center", borderBottom: `2px solid ${C.gris}`, paddingBottom: 3 }}>ANTIGUOS</th>
            <th style={{ textAlign: "center", borderBottom: `2px solid ${C.rojo}`, paddingBottom: 3 }}>FACT.</th>
          </tr>
          <tr style={{ fontSize: 11, color: C.gris, textAlign: "right" }}>
            <th style={{ textAlign: "left" }}>Escuela</th>
            <th>Meta</th><th>Avance</th><th>%</th><th>Falta</th>
            <th>Meta</th><th>Avance</th><th>%</th>
            <th>Meta</th><th>Avance</th><th>%</th>
            <th>Meta</th><th>Avance</th><th>%</th>
            <th>S/pagar</th>
          </tr>
        </thead>
        <tbody>
          {ESCUELAS.map((e) => <Fila key={e} nombre={e} d={acumulado[e]} />)}
          <Fila nombre="TOTAL" d={tot} negrita />
        </tbody>
      </table>
    </div>
  );
}

/* ---------------- vista: cargar bases ---------------- */
function VistaCargar({ estado, guardar, anio, setAviso }) {
  const periodos = estado.anios[anio].periodos;
  const [pIdx, setPIdx] = useState(0);
  const [tipo, setTipo] = useState("pagadas");
  const [previa, setPrevia] = useState(null);
  const [procesando, setProcesando] = useState(false);
  const inputRef = useRef(null);

  const alSeleccionar = async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    setProcesando(true);
    setPrevia(null);
    try {
      const { filas, nombreHoja } = await leerArchivo(file);
      const res = agregarBase(filas, tipo);
      setPrevia({ ...res, archivo: file.name, nombreHoja });
    } catch (err) {
      setAviso({ tipo: "error", texto: `No se pudo procesar «${file.name}»: ${err.message}` });
    } finally {
      setProcesando(false);
      if (inputRef.current) inputRef.current.value = "";
    }
  };

  const aplicar = async () => {
    const nuevo = JSON.parse(JSON.stringify(estado));
    const fecha = new Date().toLocaleDateString("es-CO", { day: "2-digit", month: "short", year: "numeric" });

    if (tipo === "metas") {
      /* El reporte trae todos los bloques del año, así que se actualiza cada
         período del aplicativo cuyo nombre coincida con un bloque del archivo. */
      const periodosDelAnio = nuevo.anios[anio].periodos;
      const actualizados = [];
      for (const bloque of previa.bloques) {
        const destino = periodosDelAnio.find((per) => mismoBloque(per.nombre, bloque));
        if (!destino) continue;
        const prom = num(destino.promedio);
        for (const e of ESCUELAS) {
          const reg = previa.porBloque[bloque][e];
          if (!reg) continue;
          const d = destino.escuelas[e];
          // Los valores del reporte son créditos.
          if (reg.metaNuevos !== null) d.metaCredNuevos = reg.metaNuevos;
          if (reg.metaAntiguos !== null) d.metaCredAntiguos = reg.metaAntiguos;
          if (reg.avNuevos !== null) d.avCredNuevos = reg.avNuevos;
          if (reg.avAntiguos !== null) d.avCredAntiguos = reg.avAntiguos;
          if (reg.metaTotal !== null) d.metaCred = reg.metaTotal;
          else if (reg.metaNuevos !== null && reg.metaAntiguos !== null) d.metaCred = reg.metaNuevos + reg.metaAntiguos;
          // La meta en estudiantes se deriva con los créditos promedio del período.
          // Se redondea porque son personas: no existen fracciones de estudiante.
          if (prom && prom > 0) {
            if (reg.metaNuevos !== null) d.metaEstNuevos = Math.round(reg.metaNuevos / prom);
            if (reg.metaAntiguos !== null) d.metaEstAntiguos = Math.round(reg.metaAntiguos / prom);
            if (d.metaCred !== null && d.metaCred !== undefined) d.metaEst = Math.round(d.metaCred / prom);
          }
        }
        destino.fuenteMetas = { archivo: previa.archivo, fecha, bloque };
        actualizados.push(destino.nombre);
      }
      if (actualizados.length === 0) {
        const existentes = periodos.map((per) => per.nombre).join(", ");
        setAviso({ tipo: "error", texto: `Los bloques del archivo (${previa.bloques.join(", ")}) no coinciden con los períodos de ${anio} (${existentes}). Revise que los nombres correspondan o créelos en la pestaña Metas.` });
        return;
      }
      await guardar(nuevo);
      setPrevia(null);
      setAviso({ tipo: "ok", texto: `Metas de ${anio} actualizadas con «${previa.archivo}» en ${actualizados.length} período(s): ${actualizados.join(", ")}.` });
      return;
    }

    const p = nuevo.anios[anio].periodos[pIdx];
    for (const e of ESCUELAS) {
      const agg = previa.porEscuela[e];
      if (tipo === "pagadas") {
        p.escuelas[e].avCred = agg ? Math.round(agg.creditos) : 0;
        p.escuelas[e].avEst = agg ? agg.estudiantes : 0;
        p.escuelas[e].avEstNuevos = previa.condicionDetectada ? (agg ? agg.nuevos : 0) : null;
      } else {
        p.escuelas[e].factPend = agg ? agg.facturas : 0;
        p.escuelas[e].estPendNuevos = agg ? agg.estPendNuevos : 0;
        p.escuelas[e].estPendAntiguos = agg ? agg.estPendAntiguos : 0;
        p.escuelas[e].credPendNuevos = agg ? agg.credPendNuevos : 0;
        p.escuelas[e].credPendAntiguos = agg ? agg.credPendAntiguos : 0;
        p.escuelas[e].valorPend = agg ? Math.round(agg.valor) : 0;
      }
    }
    if (tipo === "pagadas") p.fuentePagadas = { archivo: previa.archivo, fecha };
    else p.fuentePendientes = { archivo: previa.archivo, fecha };
    await guardar(nuevo);
    setPrevia(null);
    setAviso({ tipo: "ok", texto: `Avance del período ${p.nombre} de ${anio} actualizado con «${previa.archivo}».` });
  };

  const totalPrevia = previa && previa.porEscuela
    ? Object.values(previa.porEscuela).reduce((a, v) => {
        for (const k of Object.keys(v)) {
          if (typeof v[k] === "number") a[k] = (a[k] || 0) + v[k];
        }
        return a;
      }, {})
    : null;

  return (
    <>
      <Tarjeta titulo="Actualizar avance con una base de datos">
        <p style={{ fontSize: 14, color: C.gris, marginTop: 0 }}>
          Suba el archivo Excel o CSV tal como lo descarga del sistema. La aplicación detecta los encabezados automáticamente,
          agrupa por escuela y reemplaza el avance del período seleccionado. Solo se guardan los totales por escuela; los datos
          personales de los estudiantes no se almacenan.
        </p>
        <div style={{ display: "flex", gap: 14, flexWrap: "wrap", alignItems: "end" }}>
          {tipo !== "metas" && (
            <label style={{ fontSize: 13, fontWeight: 600 }}>
              Período<br />
              <select value={pIdx} onChange={(e) => setPIdx(Number(e.target.value))} style={{ font: "inherit", padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.borde}`, marginTop: 4 }}>
                {periodos.map((p, i) => <option key={p.id} value={i}>{anio} ({p.nombre})</option>)}
              </select>
            </label>
          )}
          <div style={{ fontSize: 13, fontWeight: 600 }}>
            Tipo de base<br />
            <div style={{ display: "inline-flex", marginTop: 4, border: `1px solid ${C.borde}`, borderRadius: 8, overflow: "hidden" }}>
              {[["pagadas", "Facturas pagadas (actas de matrícula)"], ["pendientes", "Facturas pendientes por pagar"], ["metas", "Metas del período (nuevos y antiguos)"]].map(([id, tx]) => (
                <button key={id} className="btn" onClick={() => { setTipo(id); setPrevia(null); }}
                  style={{ borderRadius: 0, background: tipo === id ? C.azul : "#fff", color: tipo === id ? "#fff" : C.tinta, fontWeight: 600, fontSize: 13 }}>
                  {tx}
                </button>
              ))}
            </div>
          </div>
          <label className="btn" style={{ background: C.ambar, color: C.tinta, display: "inline-block" }}>
            {procesando ? "Procesando…" : "Seleccionar archivo"}
            <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" onChange={alSeleccionar} style={{ display: "none" }} />
          </label>
        </div>
        <p style={{ fontSize: 12, color: C.gris, marginBottom: 0 }}>
          {tipo === "pagadas"
            ? "Se cuentan estudiantes (documentos únicos), se suman los «Créditos Totales» y se identifican los estudiantes nuevos por la columna «Condición» (o «Tipo Estudiante» / «Estado») por escuela. Columnas requeridas: Escuela, Documento; opcionales: Créditos Totales, Condición."
            : tipo === "pendientes"
              ? "Se cuentan los estudiantes con facturas pendientes por escuela, separando nuevos y antiguos por la columna «Estado» (o «Condición»), y se estiman los créditos por recuperar según el nivel del programa (pregrado 12, posgrado 7, INVIL 1 en promedio). También se suma el valor pendiente en pesos. Columnas requeridas: Escuela y Documento; se aprovechan además Estado, Programa y Valor Factura si están."
              : "Reporte «Detallado por Escuela» con columnas Escuela, Cohorte, Bloque y los grupos Nuevos / Antiguos / Total (Concertado y Ejecutado). Las cifras son créditos académicos. El archivo trae todos los bloques del año, así que se aplican de una vez a los períodos correspondientes (16-01, 16-02, 8-03, 16-04, 16-05); las filas Anual y de cohorte se omiten porque son subtotales. No requiere elegir período."}
          {" "}Se lee la primera hoja del archivo.
        </p>
      </Tarjeta>

      {previa && previa.esMetas && (
        <Tarjeta titulo={`Vista previa — ${previa.archivo}`}>
          <div style={{ fontSize: 13, color: C.gris, marginBottom: 10 }}>
            Hoja «{previa.nombreHoja}» · {fmt(previa.filasLeidas)} filas leídas · {previa.bloques.length} bloque(s) detectado(s)
            {previa.filasSubtotal > 0 && <> · {previa.filasSubtotal} filas de subtotal (Anual / cohorte) omitidas</>}
            {previa.sinClasificar > 0 && <span style={{ color: C.rojo, fontWeight: 700 }}> · {previa.sinClasificar} sin escuela reconocida</span>}
          </div>
          <div style={{ background: C.azulSuave, border: `1px solid ${C.borde}`, borderRadius: 8, padding: "10px 12px", fontSize: 12.5, color: C.tinta, marginBottom: 12 }}>
            Las cifras del reporte son <strong>créditos académicos</strong>. Se registran como meta de créditos de nuevos y
            antiguos, y la meta equivalente en estudiantes se calcula dividiendo entre los créditos promedio de cada período
            (ajustables en la pestaña Metas).
          </div>
          {previa.bloques.map((bloque) => {
            const destino = periodos.find((per) => mismoBloque(per.nombre, bloque));
            const filas = ESCUELAS.filter((e) => previa.porBloque[bloque][e]);
            const t = filas.reduce((a, e) => {
              const r = previa.porBloque[bloque][e];
              return { n: a.n + (r.metaNuevos || 0), ant: a.ant + (r.metaAntiguos || 0) };
            }, { n: 0, ant: 0 });
            const prom = destino ? num(destino.promedio) : null;
            return (
              <div key={bloque} style={{ marginBottom: 16 }}>
                <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 6 }}>
                  Bloque {bloque}
                  {destino
                    ? <span style={{ fontWeight: 400, color: C.verde }}> → se aplicará al período {destino.nombre}{prom ? ` (${prom} cr. promedio)` : ""}</span>
                    : <span style={{ fontWeight: 400, color: C.rojo }}> → sin período equivalente en {anio}: no se aplicará</span>}
                </div>
                <div style={{ overflowX: "auto" }}>
                  <table className="seg" style={{ borderCollapse: "collapse", width: "100%", minWidth: 560 }}>
                    <thead>
                      <tr style={{ fontSize: 12, color: C.gris, textAlign: "right" }}>
                        <th style={{ textAlign: "left" }}>Escuela</th>
                        <th>Meta cr. nuevos</th><th>Meta cr. antiguos</th><th>Meta cr. total</th>
                        <th>Est. nuevos</th><th>Est. antiguos</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filas.map((e) => {
                        const r = previa.porBloque[bloque][e];
                        const tot = r.metaTotal !== null ? r.metaTotal : (r.metaNuevos || 0) + (r.metaAntiguos || 0);
                        return (
                          <tr key={e}>
                            <td style={{ borderBottom: `1px solid ${C.borde}`, fontWeight: 600 }}>{e}</td>
                            <td className="num" style={{ borderBottom: `1px solid ${C.borde}`, textAlign: "right", color: C.verde, fontWeight: 700 }}>{fmt(r.metaNuevos)}</td>
                            <td className="num" style={{ borderBottom: `1px solid ${C.borde}`, textAlign: "right" }}>{fmt(r.metaAntiguos)}</td>
                            <td className="num" style={{ borderBottom: `1px solid ${C.borde}`, textAlign: "right" }}>{fmt(tot)}</td>
                            <td className="num" style={{ borderBottom: `1px solid ${C.borde}`, textAlign: "right", color: C.gris }}>{prom ? fmt(Math.round(r.metaNuevos / prom)) : "—"}</td>
                            <td className="num" style={{ borderBottom: `1px solid ${C.borde}`, textAlign: "right", color: C.gris }}>{prom ? fmt(Math.round(r.metaAntiguos / prom)) : "—"}</td>
                          </tr>
                        );
                      })}
                      <tr style={{ fontWeight: 800, background: C.azulSuave }}>
                        <td>TOTAL</td>
                        <td className="num" style={{ textAlign: "right" }}>{fmt(t.n)}</td>
                        <td className="num" style={{ textAlign: "right" }}>{fmt(t.ant)}</td>
                        <td className="num" style={{ textAlign: "right" }}>{fmt(t.n + t.ant)}</td>
                        <td className="num" style={{ textAlign: "right" }}>{prom ? fmt(Math.round(t.n / prom)) : "—"}</td>
                        <td className="num" style={{ textAlign: "right" }}>{prom ? fmt(Math.round(t.ant / prom)) : "—"}</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>
            );
          })}
          <div style={{ marginTop: 14, display: "flex", gap: 10 }}>
            <button className="btn" style={{ background: C.verde, color: "#fff" }} onClick={aplicar}>
              Aplicar metas a {anio}
            </button>
            <button className="btn" style={{ background: "#fff", border: `1px solid ${C.borde}`, color: C.gris }} onClick={() => setPrevia(null)}>
              Descartar
            </button>
          </div>
        </Tarjeta>
      )}

      {previa && !previa.esMetas && (
        <Tarjeta titulo={`Vista previa — ${previa.archivo}`}>
          <div style={{ fontSize: 13, color: C.gris, marginBottom: 10 }}>
            Hoja «{previa.nombreHoja}» · {fmt(previa.filasLeidas)} registros leídos
            {previa.sinClasificar > 0 && <span style={{ color: C.rojo, fontWeight: 700 }}> · {previa.sinClasificar} sin escuela reconocida (no se contarán)</span>}
            {tipo === "pagadas" && !previa.creditosDetectados && <span style={{ color: C.rojo, fontWeight: 700 }}> · no se encontró la columna de créditos: solo se actualizarán estudiantes</span>}
            {tipo === "pagadas" && !previa.condicionDetectada && <span style={{ color: C.rojo, fontWeight: 700 }}> · no se encontró la columna «Condición»: no se identificarán los estudiantes nuevos</span>}
          </div>
          <div style={{ overflowX: "auto" }}>
            <table className="seg" style={{ borderCollapse: "collapse", width: "100%", minWidth: 460 }}>
              <thead>
                <tr style={{ fontSize: 12, color: C.gris, textAlign: "right" }}>
                  <th style={{ textAlign: "left" }}>Escuela</th>
                  {tipo === "pagadas" && (<><th>Estudiantes</th><th>Nuevos</th><th>Créditos</th></>)}
                  {tipo === "pendientes" && (<><th>Estudiantes</th><th>Nuevos</th><th>Antiguos</th><th>Créd. estim.</th><th>Valor</th></>)}
                </tr>
              </thead>
              <tbody>
                {ESCUELAS.filter((e) => previa.porEscuela[e]).map((e) => {
                  const v = previa.porEscuela[e];
                  return (
                    <tr key={e}>
                      <td style={{ borderBottom: `1px solid ${C.borde}`, fontWeight: 600 }}>{e}</td>
                      {tipo === "pagadas" && (
                        <>
                          <td className="num" style={{ borderBottom: `1px solid ${C.borde}`, textAlign: "right" }}>{fmt(v.estudiantes)}</td>
                          <td className="num" style={{ borderBottom: `1px solid ${C.borde}`, textAlign: "right", color: C.azul, fontWeight: 700 }}>{previa.condicionDetectada ? fmt(v.nuevos) : "—"}</td>
                          <td className="num" style={{ borderBottom: `1px solid ${C.borde}`, textAlign: "right" }}>{fmt(Math.round(v.creditos))}</td>
                        </>
                      )}
                      {tipo === "pendientes" && (
                        <>
                          <td className="num" style={{ borderBottom: `1px solid ${C.borde}`, textAlign: "right" }}>{fmt(v.estudiantes)}</td>
                          <td className="num" style={{ borderBottom: `1px solid ${C.borde}`, textAlign: "right", color: C.azul, fontWeight: 700 }}>{fmt(v.estPendNuevos)}</td>
                          <td className="num" style={{ borderBottom: `1px solid ${C.borde}`, textAlign: "right" }}>{fmt(v.estPendAntiguos)}</td>
                          <td className="num" style={{ borderBottom: `1px solid ${C.borde}`, textAlign: "right", fontWeight: 700 }}>{fmt((v.credPendNuevos || 0) + (v.credPendAntiguos || 0))}</td>
                          <td className="num" style={{ borderBottom: `1px solid ${C.borde}`, textAlign: "right", color: C.gris }}>{v.valor > 0 ? `$ ${fmt(Math.round(v.valor))}` : "—"}</td>
                        </>
                      )}
                    </tr>
                  );
                })}
                {totalPrevia && (
                  <tr style={{ fontWeight: 800, background: C.azulSuave }}>
                    <td>TOTAL</td>
                    {tipo === "pagadas" && (
                      <>
                        <td className="num" style={{ textAlign: "right" }}>{fmt(totalPrevia.estudiantes)}</td>
                        <td className="num" style={{ textAlign: "right", color: C.azul }}>{previa.condicionDetectada ? fmt(totalPrevia.nuevos) : "—"}</td>
                        <td className="num" style={{ textAlign: "right" }}>{fmt(Math.round(totalPrevia.creditos || 0))}</td>
                      </>
                    )}
                    {tipo === "pendientes" && (
                      <>
                        <td className="num" style={{ textAlign: "right" }}>{fmt(totalPrevia.estudiantes)}</td>
                        <td className="num" style={{ textAlign: "right", color: C.azul }}>{fmt(totalPrevia.estPendNuevos)}</td>
                        <td className="num" style={{ textAlign: "right" }}>{fmt(totalPrevia.estPendAntiguos)}</td>
                        <td className="num" style={{ textAlign: "right" }}>{fmt((totalPrevia.credPendNuevos || 0) + (totalPrevia.credPendAntiguos || 0))}</td>
                        <td className="num" style={{ textAlign: "right" }}>{(totalPrevia.valor || 0) > 0 ? `$ ${fmt(Math.round(totalPrevia.valor))}` : "—"}</td>
                      </>
                    )}
                  </tr>
                )}
              </tbody>
            </table>
          </div>
          <div style={{ marginTop: 14, display: "flex", gap: 10 }}>
            <button className="btn" style={{ background: C.verde, color: "#fff" }} onClick={aplicar}>
              Aplicar al período {anio} ({periodos[pIdx].nombre})
            </button>
            <button className="btn" style={{ background: "#fff", border: `1px solid ${C.borde}`, color: C.gris }} onClick={() => setPrevia(null)}>
              Descartar
            </button>
          </div>
        </Tarjeta>
      )}
    </>
  );
}

/* ---------------- vista: metas ---------------- */
function VistaMetas({ estado, guardar, anio, setAviso, setAnio }) {
  const periodos = estado.anios[anio].periodos;
  const [pIdx, setPIdx] = useState(0);
  const periodo = periodos[Math.min(pIdx, periodos.length - 1)];

  const cambiar = async (esc, campo, valor) => {
    const nuevo = JSON.parse(JSON.stringify(estado));
    const n = valor === "" ? null : Number(valor);
    nuevo.anios[anio].periodos[pIdx].escuelas[esc][campo] = n === null || isNaN(n) ? null : n;
    await guardar(nuevo);
  };

  const cambiarPromedio = async (valor) => {
    const nuevo = JSON.parse(JSON.stringify(estado));
    nuevo.anios[anio].periodos[pIdx].promedio = valor === "" ? null : Number(valor);
    await guardar(nuevo);
  };

  const agregarPeriodo = async () => {
    const nombre = prompt("Nombre del nuevo período (por ejemplo: 16-06):");
    if (!nombre) return;
    const nuevo = JSON.parse(JSON.stringify(estado));
    nuevo.anios[anio].periodos.push({
      id: `${anio}-${nombre}-${Date.now()}`,
      nombre: nombre.trim(),
      promedio: 14,
      escuelas: Object.fromEntries(ESCUELAS.map((e) => [e, { metaCred: null, avCred: null, metaEst: null, avEst: null, factPend: null, avEstNuevos: null, metaEstNuevos: null, metaEstAntiguos: null, metaCredNuevos: null, metaCredAntiguos: null, avCredNuevos: null, avCredAntiguos: null, credPendNuevos: null, credPendAntiguos: null, estPendNuevos: null, estPendAntiguos: null, valorPend: null }])),
      fuentePagadas: null,
      fuentePendientes: null,
    });
    await guardar(nuevo);
    setPIdx(nuevo.anios[anio].periodos.length - 1);
  };

  const agregarAnio = async () => {
    const a = prompt("Año a crear (por ejemplo: 2027):");
    if (!a || !/^\d{4}$/.test(a.trim())) return;
    const aa = a.trim();
    if (estado.anios[aa]) { setAviso({ tipo: "error", texto: `El año ${aa} ya existe.` }); return; }
    const nuevo = JSON.parse(JSON.stringify(estado));
    nuevo.anios[aa] = {
      periodos: ["16-01", "16-02", "8-03", "16-04", "16-05"].map((n, i) => ({
        id: `${aa}-${n}-${i}`,
        nombre: n,
        promedio: n === "8-03" ? 7 : 14,
        escuelas: Object.fromEntries(ESCUELAS.map((e) => [e, { metaCred: null, avCred: null, metaEst: null, avEst: null, factPend: null, avEstNuevos: null, metaEstNuevos: null, metaEstAntiguos: null, metaCredNuevos: null, metaCredAntiguos: null, avCredNuevos: null, avCredAntiguos: null, credPendNuevos: null, credPendAntiguos: null, estPendNuevos: null, estPendAntiguos: null, valorPend: null }])),
        fuentePagadas: null,
        fuentePendientes: null,
      })),
    };
    await guardar(nuevo);
    setAnio(aa);
    setAviso({ tipo: "ok", texto: `Año ${aa} creado con los cinco períodos habituales. Registre las metas de cada uno.` });
  };

  const restablecer = async () => {
    if (!confirm("¿Restablecer toda la información a la del archivo base? Se perderán las cargas y metas registradas en la aplicación.")) return;
    await guardar(estadoInicial());
    setAviso({ tipo: "ok", texto: "Información restablecida a la del archivo base." });
  };

  const celda = { borderBottom: `1px solid ${C.borde}`, fontSize: 13.5 };
  return (
    <>
      <Tarjeta
        titulo="Metas y ajustes manuales"
        extra={
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button className="btn" style={{ background: C.azul, color: "#fff", fontSize: 13 }} onClick={agregarPeriodo}>+ Período</button>
            <button className="btn" style={{ background: C.azul, color: "#fff", fontSize: 13 }} onClick={agregarAnio}>+ Año</button>
            <button className="btn" style={{ background: "#fff", border: `1px solid ${C.rojo}`, color: C.rojo, fontSize: 13 }} onClick={restablecer}>Restablecer todo</button>
          </div>
        }
      >
        <div style={{ display: "flex", gap: 14, flexWrap: "wrap", alignItems: "end", marginBottom: 14 }}>
          <label style={{ fontSize: 13, fontWeight: 600 }}>
            Período<br />
            <select value={pIdx} onChange={(e) => setPIdx(Number(e.target.value))} style={{ font: "inherit", padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.borde}`, marginTop: 4 }}>
              {periodos.map((p, i) => <option key={p.id} value={i}>{anio} ({p.nombre})</option>)}
            </select>
          </label>
          <label style={{ fontSize: 13, fontWeight: 600 }}>
            Créditos promedio por estudiante<br />
            <input className="campo" style={{ marginTop: 4 }} type="number" value={periodo.promedio ?? ""} onChange={(e) => cambiarPromedio(e.target.value)} />
          </label>
        </div>
        <p style={{ fontSize: 13, color: C.gris, marginTop: 0 }}>
          Registre aquí las metas de cada escuela. El avance se actualiza normalmente al subir las bases, pero también puede
          ajustarse manualmente (por ejemplo, para períodos históricos).
        </p>
        <div style={{ overflowX: "auto" }}>
          <table className="seg" style={{ borderCollapse: "collapse", width: "100%", minWidth: 700 }}>
            <thead>
              <tr style={{ fontSize: 12, color: C.gris, textAlign: "right" }}>
                <th style={{ textAlign: "left" }}>Escuela</th>
                <th>Meta créditos</th><th>Avance créditos</th>
                <th>Meta estudiantes</th><th>Avance estudiantes</th>
                <th>Meta cr. nuevos</th><th>Meta est. nuevos</th><th>Avance nuevos</th>
                <th>Meta cr. antiguos</th><th>Meta est. antiguos</th>
                <th>Facturas sin pagar</th>
              </tr>
            </thead>
            <tbody>
              {ESCUELAS.map((e) => {
                const d = periodo.escuelas[e];
                return (
                  <tr key={e}>
                    <td style={{ ...celda, fontWeight: 600 }}>{e}</td>
                    {["metaCred", "avCred", "metaEst", "avEst", "metaCredNuevos", "metaEstNuevos", "avEstNuevos", "metaCredAntiguos", "metaEstAntiguos", "factPend"].map((campo) => (
                      <td key={campo} style={{ ...celda, textAlign: "right" }}>
                        <input className="campo num" type="number" value={d[campo] ?? ""} placeholder="—"
                          onChange={(ev) => cambiar(e, campo, ev.target.value)} />
                      </td>
                    ))}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Tarjeta>
    </>
  );
}
