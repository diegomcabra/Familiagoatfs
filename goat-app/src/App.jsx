import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import { supabase, supabaseConfigured } from "./supabaseClient";
import * as XLSX from "xlsx";

/* ============ DATOS INICIALES (desde Familia_Goat.xlsx) ============ */
const INSUMOS_INICIALES = [{"id": "INS001", "nombre": "Nuez Pelada Mariposa Extra Light", "costo": 16605.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS002", "nombre": "Nuez de Pecan Partida", "costo": 25300.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS003", "nombre": "Avellanas Peladas Grandes - Origen Turquía", "costo": 35100.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS004", "nombre": "Almendra Pelada Non Pareil GRANDE 25/27 - Chile", "costo": 23985.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS005", "nombre": "Castaña de Cajú Natural W4 - Origen Brasil", "costo": 18900.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS006", "nombre": "Chips de Banana Deshidroazucaradas", "costo": 9880.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS007", "nombre": "Pistachos pelados naturales", "costo": 57710.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS008", "nombre": "Maní Tostado Sin Sal Bolsa", "costo": 2400.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS009", "nombre": "Mani repelado tostado con sal", "costo": 2375.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS010", "nombre": "Mani Japones crocante", "costo": 3000.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS011", "nombre": "Castañas de caju tostadas y saladas", "costo": 19989.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS012", "nombre": "Habas Fritas Saladas", "costo": 12272.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS013", "nombre": "Maiz crocante original", "costo": 10104.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS014", "nombre": "Castañas de Para entera Natural", "costo": 34850.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS015", "nombre": "Arándanos Rojos CHILE (Fruta desecada)", "costo": 15195.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS016", "nombre": "Pasas de uva negras Jumbo", "costo": 5690.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS017", "nombre": "Pasas de uva rubias", "costo": 7380.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS018", "nombre": "Gotas de Chocolate", "costo": 4799.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS019", "nombre": "Coco en Escamas (Sin marca)", "costo": 14200.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS020", "nombre": "Castañas de caju tostadas", "costo": 20500.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS022", "nombre": "Oregano Nacional", "costo": 4800.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS023", "nombre": "Aji molido", "costo": 4799.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS024", "nombre": "Provenzal", "costo": 4999.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS025", "nombre": "Pimenton Ahumado", "costo": 4599.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS026", "nombre": "Hongos Secos Boletus", "costo": 29375.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS027", "nombre": "Tomates Secos PREMIUM", "costo": 14714.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS028", "nombre": "Semilla de Zapallo Pelada AAA - SUPER PREMIUM (Sin marca)", "costo": 13400.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS029", "nombre": "Semilla de Girasol Pelado PREMIUM", "costo": 3600.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS030", "nombre": "Ajo Granulado Blanco Standard", "costo": 8500.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS031", "nombre": "Canela Rama", "costo": 22000.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS032", "nombre": "Sal marina finas hierbas ahumada (sin tacc) Dicomere 450gr", "costo": 1535.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS034", "nombre": "Etiquetas 4 cm", "costo": 211.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS035", "nombre": "Etiquetas 5 cm", "costo": 211.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS036", "nombre": "Etiquetas 7 cm", "costo": 211.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS037", "nombre": "Bolsa Doypack Metalizado 20x30", "costo": 314.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS038", "nombre": "Bolsa Doypack Metalizado 16X24", "costo": 345.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS039", "nombre": "Bolsa Doypack Metalizado 10X15 + 3CM", "costo": 124.8, "stockActual": 0, "stockMinimo": 0}, {"id": "INS040", "nombre": "Bolsa Doypack 13x22 Kraft", "costo": 220.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS041", "nombre": "Bolsa Doypack 13x25 Metalizado", "costo": 200.0, "stockActual": 0, "stockMinimo": 0}, {"id": "INS042", "nombre": "Bolsa Doypack 12x20 Metalizado", "costo": 152.0, "stockActual": 0, "stockMinimo": 0}];

const PRODUCTOS_INICIALES = [{"id": "PR001", "nombre": "Mix GOAT Clásico 1kg", "precio": 28455, "stock": 0}, {"id": "PR002", "nombre": "Mix GOAT Clásico 500g", "precio": 15684, "stock": 0}, {"id": "PR003", "nombre": "Mix GOAT Clásico 250g", "precio": 8259, "stock": 0}, {"id": "PR004", "nombre": "Mix GOAT Clásico 100g", "precio": 3705.28, "stock": 0}, {"id": "PR005", "nombre": "Mix GOAT Clásico Chocolate 1kg", "precio": 27342.6, "stock": 0}, {"id": "PR006", "nombre": "Mix GOAT Clásico Chocolate 500g", "precio": 14111.55, "stock": 0}, {"id": "PR007", "nombre": "Mix GOAT Clásico Chocolate 250g", "precio": 7472.77, "stock": 0}, {"id": "PR008", "nombre": "Mix GOAT Clásico Chocolate 100g", "precio": 3369.82, "stock": 0}, {"id": "PR009", "nombre": "Mix GOAT Premium Pistacho 1kg", "precio": 39401.6, "stock": 0}, {"id": "PR010", "nombre": "Mix GOAT Premium Pistacho 500g", "precio": 21548.25, "stock": 0}, {"id": "PR011", "nombre": "Mix GOAT Premium Pistacho 250g", "precio": 11191.12, "stock": 0}, {"id": "PR012", "nombre": "Mix GOAT Premium Pistacho 100g", "precio": 4956.32, "stock": 0}, {"id": "PR013", "nombre": "Mix GOAT Premium Pecan 1kg", "precio": 30326.8, "stock": 0}, {"id": "PR014", "nombre": "Mix GOAT Premium Pecan 500g", "precio": 16686.75, "stock": 0}, {"id": "PR015", "nombre": "Mix GOAT Premium Pecan 250g", "precio": 8760.38, "stock": 0}, {"id": "PR016", "nombre": "Mix GOAT Premium Pecan 100g", "precio": 3919.2, "stock": 0}, {"id": "PR017", "nombre": "Mix GOAT Premium Arándanos 1kg", "precio": 27540.1, "stock": 0}, {"id": "PR018", "nombre": "Mix GOAT Premium Arándanos 500g", "precio": 15193.88, "stock": 0}, {"id": "PR019", "nombre": "Mix GOAT Premium Arándanos 250g", "precio": 8013.94, "stock": 0}, {"id": "PR020", "nombre": "Mix GOAT Premium Arándanos 100g", "precio": 3600.72, "stock": 0}, {"id": "PR021", "nombre": "Mix GOAT Energético 1kg", "precio": 30980.6, "stock": 0}, {"id": "PR022", "nombre": "Mix GOAT Energético 500g", "precio": 17037, "stock": 0}, {"id": "PR023", "nombre": "Mix GOAT Energético 250g", "precio": 8935.5, "stock": 0}, {"id": "PR024", "nombre": "Mix GOAT Energético 100g", "precio": 3993.92, "stock": 0}, {"id": "PR025", "nombre": "Mix GOAT Almendra y Coco 1kg", "precio": 12713.4, "stock": 0}, {"id": "PR026", "nombre": "Mix GOAT Almendra y Coco 500g", "precio": 7251, "stock": 0}, {"id": "PR027", "nombre": "Mix GOAT Almendra y Coco 250g", "precio": 4042.5, "stock": 0}, {"id": "PR028", "nombre": "Mix GOAT Almendra y Coco 100g", "precio": 1906.24, "stock": 0}, {"id": "PR029", "nombre": "Mix GOAT Coco 1kg", "precio": 10891.3, "stock": 0}, {"id": "PR030", "nombre": "Mix GOAT Coco 500g", "precio": 6274.88, "stock": 0}, {"id": "PR031", "nombre": "Mix GOAT Coco 250g", "precio": 3554.44, "stock": 0}, {"id": "PR032", "nombre": "Mix GOAT Coco 100g", "precio": 1698, "stock": 0}, {"id": "PR033", "nombre": "Mix GOAT Cervecero 1kg", "precio": 16170.15, "stock": 0}, {"id": "PR034", "nombre": "Mix GOAT Cervecero 500g", "precio": 8525.33, "stock": 0}, {"id": "PR035", "nombre": "Mix GOAT Cervecero 250g", "precio": 4679.66, "stock": 0}, {"id": "PR036", "nombre": "Mix GOAT Cervecero 100g", "precio": 2178.1, "stock": 0}, {"id": "PR037", "nombre": "Mix GOAT Cervecero Japonés 1kg", "precio": 17520.15, "stock": 0}, {"id": "PR038", "nombre": "Mix GOAT Cervecero Japonés 500g", "precio": 9200.33, "stock": 0}, {"id": "PR039", "nombre": "Mix GOAT Cervecero Japonés 250g", "precio": 5017.16, "stock": 0}, {"id": "PR040", "nombre": "Mix GOAT Cervecero Japonés 100g", "precio": 2322.1, "stock": 0}, {"id": "PR041", "nombre": "Almendra Pelada Non Pareil Grande 1kg", "precio": 34314, "stock": 0}, {"id": "PR042", "nombre": "Almendra Pelada Non Pareil Grande 500g", "precio": 18822.75, "stock": 0}, {"id": "PR043", "nombre": "Almendra Pelada Non Pareil Grande 250g", "precio": 9828.38, "stock": 0}, {"id": "PR044", "nombre": "Almendra Pelada Non Pareil Grande 100g", "precio": 4374.88, "stock": 0}, {"id": "PR045", "nombre": "Nuez Pelada Mariposa Extra Light 1kg", "precio": 23982, "stock": 0}, {"id": "PR046", "nombre": "Nuez Pelada Mariposa Extra Light 500g", "precio": 13287.75, "stock": 0}, {"id": "PR047", "nombre": "Nuez Pelada Mariposa Extra Light 250g", "precio": 7060.88, "stock": 0}, {"id": "PR048", "nombre": "Nuez Pelada Mariposa Extra Light 100g", "precio": 3194.08, "stock": 0}, {"id": "PR049", "nombre": "Nuez de Pecan Partida 1kg", "precio": 36155, "stock": 0}, {"id": "PR050", "nombre": "Nuez de Pecan Partida 500g", "precio": 19809, "stock": 0}, {"id": "PR051", "nombre": "Nuez de Pecan Partida 250g", "precio": 10321.5, "stock": 0}, {"id": "PR052", "nombre": "Nuez de Pecan Partida 100g", "precio": 4585.28, "stock": 0}, {"id": "PR053", "nombre": "Pistacho pelado sin sal 1kg", "precio": 81529, "stock": 0}, {"id": "PR054", "nombre": "Pistacho pelado sin sal 500g", "precio": 44116.5, "stock": 0}, {"id": "PR055", "nombre": "Pistacho pelado sin sal 250g", "precio": 22475.25, "stock": 0}, {"id": "PR056", "nombre": "Pistacho pelado sin sal 100g", "precio": 9770.88, "stock": 0}, {"id": "PR057", "nombre": "Castaña de Cajú Natural 1kg", "precio": 27195, "stock": 0}, {"id": "PR058", "nombre": "Castaña de Cajú Natural 500g", "precio": 15009, "stock": 0}, {"id": "PR059", "nombre": "Castaña de Cajú Natural 250g", "precio": 7921.5, "stock": 0}, {"id": "PR060", "nombre": "Castaña de Cajú Natural 100g", "precio": 3561.28, "stock": 0}, {"id": "PR061", "nombre": "Castañas de caju tostadas 1kg", "precio": 29435, "stock": 0}, {"id": "PR062", "nombre": "Castañas de caju tostadas 500g", "precio": 16209, "stock": 0}, {"id": "PR063", "nombre": "Castañas de caju tostadas 250g", "precio": 8521.5, "stock": 0}, {"id": "PR064", "nombre": "Castañas de caju tostadas 100g", "precio": 3817.28, "stock": 0}, {"id": "PR065", "nombre": "Castañas de caju tostadas y saladas 1kg", "precio": 28719.6, "stock": 0}, {"id": "PR066", "nombre": "Castañas de caju tostadas y saladas 500g", "precio": 15825.75, "stock": 0}, {"id": "PR067", "nombre": "Castañas de caju tostadas y saladas 250g", "precio": 8329.88, "stock": 0}, {"id": "PR068", "nombre": "Castañas de caju tostadas y saladas 100g", "precio": 3735.52, "stock": 0}, {"id": "PR069", "nombre": "Maní Tostado Sin Sal Bolsa 1kg", "precio": 4095, "stock": 0}, {"id": "PR070", "nombre": "Maní Tostado Sin Sal Bolsa 500g", "precio": 2634, "stock": 0}, {"id": "PR071", "nombre": "Maní Tostado Sin Sal Bolsa 250g", "precio": 1734, "stock": 0}, {"id": "PR072", "nombre": "Maní Tostado Sin Sal Bolsa 100g", "precio": 921.28, "stock": 0}, {"id": "PR073", "nombre": "Maní Japones 1kg", "precio": 5287.5, "stock": 0}, {"id": "PR074", "nombre": "Maní Japones 500g", "precio": 3084, "stock": 0}, {"id": "PR075", "nombre": "Maní Japones 250g", "precio": 1959, "stock": 0}, {"id": "PR076", "nombre": "Maní Japones 100g", "precio": 953.7, "stock": 0}, {"id": "PR077", "nombre": "Mano Tostado con sal 1kg", "precio": 4350, "stock": 0}, {"id": "PR078", "nombre": "Mano Tostado con sal 500g", "precio": 2615.25, "stock": 0}, {"id": "PR079", "nombre": "Mano Tostado con sal 250g", "precio": 1724.62, "stock": 0}, {"id": "PR080", "nombre": "Mano Tostado con sal 100g", "precio": 859.95, "stock": 0}, {"id": "PR081", "nombre": "Arándanos Rojos CHILE 1kg", "precio": 22008, "stock": 0}, {"id": "PR082", "nombre": "Arándanos Rojos CHILE 500g", "precio": 12230.25, "stock": 0}, {"id": "PR083", "nombre": "Arándanos Rojos CHILE 250g", "precio": 6532.12, "stock": 0}, {"id": "PR084", "nombre": "Arándanos Rojos CHILE 100g", "precio": 2968.48, "stock": 0}, {"id": "PR085", "nombre": "Pasas de uva negras Jumbo 1kg", "precio": 8701, "stock": 0}, {"id": "PR086", "nombre": "Pasas de uva negras Jumbo 500g", "precio": 5101.5, "stock": 0}, {"id": "PR087", "nombre": "Pasas de uva negras Jumbo 250g", "precio": 2967.75, "stock": 0}, {"id": "PR088", "nombre": "Pasas de uva negras Jumbo 100g", "precio": 1447.68, "stock": 0}, {"id": "PR089", "nombre": "Pasas de uva rubias 1kg", "precio": 11067, "stock": 0}, {"id": "PR090", "nombre": "Pasas de uva rubias 500g", "precio": 6369, "stock": 0}, {"id": "PR091", "nombre": "Pasas de uva rubias 250g", "precio": 3601.5, "stock": 0}, {"id": "PR092", "nombre": "Pasas de uva rubias 100g", "precio": 1718.08, "stock": 0}, {"id": "PR093", "nombre": "Chips de Banana 1kg", "precio": 14021, "stock": 0}, {"id": "PR094", "nombre": "Chips de Banana 500g", "precio": 8244, "stock": 0}, {"id": "PR095", "nombre": "Chips de Banana 250g", "precio": 4539, "stock": 0}, {"id": "PR096", "nombre": "Chips de Banana 100g", "precio": 2118.08, "stock": 0}, {"id": "PR099", "nombre": "Provenzal GOAT", "precio": 0, "stock": 0}, {"id": "PR100", "nombre": "Aji molido GOAT", "precio": 0, "stock": 0}, {"id": "PR101", "nombre": "Oregano Nacional", "precio": 0, "stock": 0}, {"id": "PR102", "nombre": "Semilla de Zapallo Pelada  GOAT", "precio": 0, "stock": 0}, {"id": "PR103", "nombre": "Semilla de Girasol Pelado PREMIUM", "precio": 0, "stock": 0}, {"id": "PR104", "nombre": "Hongos Secos Boletus", "precio": 0, "stock": 0}, {"id": "PR105", "nombre": "Tomates Secos PREMIUM", "precio": 0, "stock": 0}];

const RECETAS_INICIALES = [{"idProducto": "PR001", "idInsumo": "INS001", "cantidad": 0.3}, {"idProducto": "PR001", "idInsumo": "INS003", "cantidad": 0.1}, {"idProducto": "PR001", "idInsumo": "INS004", "cantidad": 0.3}, {"idProducto": "PR001", "idInsumo": "INS005", "cantidad": 0.1}, {"idProducto": "PR001", "idInsumo": "INS016", "cantidad": 0.1}, {"idProducto": "PR001", "idInsumo": "INS017", "cantidad": 0.1}, {"idProducto": "PR001", "idInsumo": "INS035", "cantidad": 1.0}, {"idProducto": "PR001", "idInsumo": "INS037", "cantidad": 1.0}, {"idProducto": "PR002", "idInsumo": "INS001", "cantidad": 0.15}, {"idProducto": "PR002", "idInsumo": "INS003", "cantidad": 0.05}, {"idProducto": "PR002", "idInsumo": "INS004", "cantidad": 0.15}, {"idProducto": "PR002", "idInsumo": "INS005", "cantidad": 0.05}, {"idProducto": "PR002", "idInsumo": "INS016", "cantidad": 0.05}, {"idProducto": "PR002", "idInsumo": "INS017", "cantidad": 0.05}, {"idProducto": "PR002", "idInsumo": "INS035", "cantidad": 1.0}, {"idProducto": "PR002", "idInsumo": "INS038", "cantidad": 1.0}];

const STORAGE_KEY = "goat-datos-v1"; // fallback local si Supabase no está configurado
const ROW_ID = "goat"; // fila única compartida en la tabla app_data
const fmt = (n) => new Intl.NumberFormat("es-AR", { style: "currency", currency: "ARS", maximumFractionDigits: 0 }).format(Math.round(n || 0));
const fmtNum = (n, d = 2) => new Intl.NumberFormat("es-AR", { maximumFractionDigits: d }).format(n || 0);
const uid = (p) => p + Math.random().toString(36).slice(2, 8).toUpperCase();
const todayISO = () => new Date().toISOString().slice(0, 10);

// Exporta una o varias listas a un archivo .xlsx.
// sheets: { "Nombre de hoja": [ {col1: valor, col2: valor}, ... ], ... }
function exportarExcel(nombreArchivo, sheets) {
  const wb = XLSX.utils.book_new();
  Object.entries(sheets).forEach(([nombreHoja, filas]) => {
    const ws = XLSX.utils.json_to_sheet(filas);
    XLSX.utils.book_append_sheet(wb, ws, nombreHoja.slice(0, 31)); // Excel limita el nombre de hoja a 31 caracteres
  });
  XLSX.writeFile(wb, `${nombreArchivo}-${todayISO()}.xlsx`);
}

function BotonExportar({ onClick, label = "Exportar" }) {
  return (
    <button
      onClick={onClick}
      className="flex items-center gap-2 border border-stone-300 hover:bg-stone-100 text-stone-700 px-3 py-2 rounded-lg text-sm font-medium"
    >
      <Icon name="download" className="w-4 h-4" /> {label}
    </button>
  );
}

/* ============ ICONOS (línea simple) ============ */
const Icon = ({ name, className = "w-5 h-5" }) => {
  const paths = {
    dashboard: "M3 3h8v8H3V3zm10 0h8v5h-8V3zM3 13h8v8H3v-8zm10 3h8v5h-8v-5z",
    insumos: "M12 2c-3 3-5 6-5 9a5 5 0 0010 0c0-3-2-6-5-9zM7 21h10",
    productos: "M20 7l-8-4-8 4v10l8 4 8-4V7zM4 7l8 4 8-4M12 11v10",
    recetas: "M4 3h16v18l-8-4-8 4V3z M8 8h8M8 12h5",
    ventas: "M3 3v18h18M7 15l4-4 3 3 5-6",
    pedidos: "M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2M9 12l2 2 4-4",
    pvp: "M12 2v20M17 5H9.5a3.5 3.5 0 000 7h5a3.5 3.5 0 010 7H6",
    plus: "M12 5v14M5 12h14",
    trash: "M4 7h16M9 7V4h6v3M6 7l1 14h10l1-14",
    edit: "M12 20h9M16.5 3.5a2.1 2.1 0 013 3L7 19l-4 1 1-4L16.5 3.5z",
    warn: "M12 2L2 21h20L12 2zm0 6v6m0 3h.01",
    check: "M20 6L9 17l-5-5",
    save: "M19 21H5a2 2 0 01-2-2V5a2 2 0 012-2h11l5 5v11a2 2 0 01-2 2zM7 3v6h8V3M7 21v-8h10v8",
    download: "M12 3v12m0 0l-4-4m4 4l4-4M4 21h16",
  };
  return (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" className={className}>
      <path d={paths[name] || ""} />
    </svg>
  );
};

/* ============ BARRA DE NIVEL "FRASCO" (elemento distintivo) ============ */
const JarLevel = ({ actual, minimo }) => {
  const max = Math.max(actual, minimo * 2, 1);
  const pct = Math.min(100, (actual / max) * 100);
  const low = actual <= minimo;
  return (
    <div className="flex items-center gap-2 w-full">
      <div className="relative h-4 flex-1 rounded-full bg-stone-100 border border-stone-200 overflow-hidden">
        <div
          className={`absolute inset-y-0 left-0 rounded-full transition-all ${low ? "bg-orange-500" : "bg-amber-700"}`}
          style={{ width: pct + "%" }}
        />
        {minimo > 0 && (
          <div className="absolute inset-y-0 border-l-2 border-dashed border-stone-500/50" style={{ left: Math.min(100, (minimo / max) * 100) + "%" }} />
        )}
      </div>
      {low && <Icon name="warn" className="w-4 h-4 text-orange-600 shrink-0" />}
    </div>
  );
};

/* ============ APP ============ */
export default function GoatApp() {
  const [tab, setTab] = useState("dashboard");
  const [insumos, setInsumos] = useState(INSUMOS_INICIALES);
  const [productos, setProductos] = useState(PRODUCTOS_INICIALES);
  const [recetas, setRecetas] = useState(RECETAS_INICIALES);
  const [ventas, setVentas] = useState([]);
  const [pedidos, setPedidos] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [saveState, setSaveState] = useState("idle"); // idle | saving | saved
  const [syncStatus, setSyncStatus] = useState(supabaseConfigured ? "checking" : "local"); // checking | synced | error | local
  const [syncError, setSyncError] = useState("");
  const lastLocalWrite = useRef(0); // timestamp del último guardado propio, para no reprocesar el eco de Supabase

  // Cargar datos: desde Supabase si está configurado, si no desde localStorage (modo solo-este-navegador)
  useEffect(() => {
    (async () => {
      if (supabaseConfigured) {
        try {
          const { data, error } = await supabase.from("app_data").select("data").eq("id", ROW_ID).maybeSingle();
          if (error) throw error;
          if (data && data.data) {
            const d = data.data;
            if (d.insumos) setInsumos(d.insumos);
            if (d.productos) setProductos(d.productos);
            if (d.recetas) setRecetas(d.recetas);
            if (d.ventas) setVentas(d.ventas);
            if (d.pedidos) setPedidos(d.pedidos);
          } else {
            // primera vez: no existe la fila todavía, la creamos con los datos iniciales
            const { error: upsertError } = await supabase.from("app_data").upsert({ id: ROW_ID, data: { insumos, productos, recetas, ventas, pedidos } });
            if (upsertError) throw upsertError;
          }
          setSyncStatus("synced");
        } catch (e) {
          console.error("Error cargando desde Supabase, uso respaldo local:", e);
          setSyncStatus("error");
          setSyncError(e.message || "No se pudo conectar a Supabase");
          try {
            const raw = localStorage.getItem(STORAGE_KEY);
            if (raw) {
              const d = JSON.parse(raw);
              if (d.insumos) setInsumos(d.insumos);
              if (d.productos) setProductos(d.productos);
              if (d.recetas) setRecetas(d.recetas);
              if (d.ventas) setVentas(d.ventas);
              if (d.pedidos) setPedidos(d.pedidos);
            }
          } catch (e2) {}
        }
      } else {
        try {
          const raw = localStorage.getItem(STORAGE_KEY);
          if (raw) {
            const d = JSON.parse(raw);
            if (d.insumos) setInsumos(d.insumos);
            if (d.productos) setProductos(d.productos);
            if (d.recetas) setRecetas(d.recetas);
            if (d.ventas) setVentas(d.ventas);
            if (d.pedidos) setPedidos(d.pedidos);
          }
        } catch (e) {}
      }
      setLoaded(true);
    })();
  }, []);

  // Escuchar cambios hechos desde OTRAS pestañas/dispositivos y traerlos en vivo
  useEffect(() => {
    if (!supabaseConfigured) return;
    const channel = supabase
      .channel("app_data_changes")
      .on("postgres_changes", { event: "UPDATE", schema: "public", table: "app_data", filter: `id=eq.${ROW_ID}` }, (payload) => {
        // Si el cambio lo disparamos nosotros mismos hace menos de 2s, lo ignoramos (evita eco)
        if (Date.now() - lastLocalWrite.current < 2000) return;
        const d = payload.new?.data;
        if (!d) return;
        if (d.insumos) setInsumos(d.insumos);
        if (d.productos) setProductos(d.productos);
        if (d.recetas) setRecetas(d.recetas);
        if (d.ventas) setVentas(d.ventas);
        if (d.pedidos) setPedidos(d.pedidos);
      })
      .subscribe((status) => {
        // "CHANNEL_ERROR" o "TIMED_OUT" indican que Realtime no está habilitado o hay un problema de conexión
        if (status === "CHANNEL_ERROR" || status === "TIMED_OUT") {
          setSyncStatus("error");
          setSyncError("No se pudo conectar el canal en vivo (revisá que habilitaste Realtime para la tabla app_data)");
        }
      });
    return () => supabase.removeChannel(channel);
  }, []);

  // Guardar (con debounce simple): a Supabase si está configurado, si no a localStorage
  useEffect(() => {
    if (!loaded) return;
    setSaveState("saving");
    const t = setTimeout(async () => {
      const payload = { insumos, productos, recetas, ventas, pedidos };
      if (supabaseConfigured) {
        try {
          lastLocalWrite.current = Date.now();
          const { error } = await supabase.from("app_data").upsert({ id: ROW_ID, data: payload, updated_at: new Date().toISOString() });
          if (error) throw error;
          setSaveState("saved");
          setSyncStatus("synced");
          setSyncError("");
        } catch (e) {
          console.error("Error guardando en Supabase:", e);
          setSaveState("idle");
          setSyncStatus("error");
          setSyncError(e.message || "No se pudo guardar en Supabase");
        }
      } else {
        try {
          localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
          setSaveState("saved");
        } catch (e) {
          setSaveState("idle");
        }
      }
    }, 500);
    return () => clearTimeout(t);
  }, [insumos, productos, recetas, ventas, pedidos, loaded]);

  const insumoMap = useMemo(() => Object.fromEntries(insumos.map((i) => [i.id, i])), [insumos]);
  const productoMap = useMemo(() => Object.fromEntries(productos.map((p) => [p.id, p])), [productos]);

  const costoProducto = useCallback(
    (idProducto) => {
      const items = recetas.filter((r) => r.idProducto === idProducto);
      return items.reduce((sum, r) => sum + (insumoMap[r.idInsumo]?.costo || 0) * (r.cantidad || 0), 0);
    },
    [recetas, insumoMap]
  );

  const insumosBajoStock = insumos.filter((i) => i.stockActual <= i.stockMinimo && i.stockMinimo > 0);
  const productosBajoStock = productos.filter((p) => p.stock <= 0);
  const valorInventarioInsumos = insumos.reduce((s, i) => s + i.costo * i.stockActual, 0);
  const valorInventarioProductos = productos.reduce((s, p) => s + p.precio * p.stock, 0);
  const totalVentasIngresos = ventas.reduce((s, v) => s + v.total, 0);
  const totalVentasGanancia = ventas.reduce((s, v) => s + (v.ganancia || 0), 0);

  const NAV = [
    { id: "dashboard", label: "Panel", icon: "dashboard" },
    { id: "insumos", label: "Insumos", icon: "insumos" },
    { id: "productos", label: "Productos", icon: "productos" },
    { id: "recetas", label: "Recetas", icon: "recetas" },
    { id: "pedidos", label: "Pedidos", icon: "pedidos" },
    { id: "ventas", label: "Ventas", icon: "ventas" },
    { id: "pvp", label: "Calculadora PVP", icon: "pvp" },
  ];

  return (
    <div className="w-full min-h-screen bg-stone-50 text-stone-900" style={{ fontFamily: "'Public Sans', system-ui, sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&family=Public+Sans:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;700&display=swap');
        .font-display { font-family: 'Fraunces', serif; font-optical-sizing: auto; }
        .font-mono-num { font-family: 'JetBrains Mono', monospace; }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-thumb { background: #d6d3d1; border-radius: 4px; }
      `}</style>

      <div className="flex">
        {/* SIDEBAR */}
        <aside className="w-60 shrink-0 min-h-screen bg-stone-900 text-stone-100 flex flex-col">
          <div className="px-6 pt-8 pb-6 border-b border-stone-700/60">
            <div className="font-display text-3xl tracking-tight text-amber-500">GOAT</div>
            <div className="text-xs text-stone-400 mt-1 tracking-wide uppercase">Gestión de mixes</div>
          </div>
          <nav className="flex-1 px-3 py-4 space-y-1">
            {NAV.map((n) => (
              <button
                key={n.id}
                onClick={() => setTab(n.id)}
                className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${
                  tab === n.id ? "bg-amber-700 text-white" : "text-stone-300 hover:bg-stone-800"
                }`}
              >
                <Icon name={n.icon} className="w-4.5 h-4.5" />
                {n.label}
              </button>
            ))}
          </nav>
          <div className="px-6 py-4 border-t border-stone-700/60 text-xs">
            <div className={`flex items-center gap-2 ${syncStatus === "error" ? "text-orange-400" : "text-stone-500"}`}>
              <Icon name={saveState === "saving" ? "save" : syncStatus === "error" ? "warn" : "check"} className="w-3.5 h-3.5" />
              {saveState === "saving"
                ? "Guardando…"
                : syncStatus === "synced"
                ? "Sincronizado"
                : syncStatus === "checking"
                ? "Conectando…"
                : syncStatus === "error"
                ? "Sin conexión a la base"
                : "Guardado (solo este navegador)"}
            </div>
            {syncStatus === "error" && syncError && (
              <div className="mt-1 text-[10px] text-stone-500 leading-snug">{syncError}</div>
            )}
          </div>
        </aside>

        {/* MAIN */}
        <main className="flex-1 min-w-0 p-8">
          {syncStatus === "error" && (
            <div className="mb-4 bg-orange-50 border border-orange-200 text-orange-800 text-sm rounded-xl px-4 py-3 flex items-start gap-2">
              <Icon name="warn" className="w-4 h-4 mt-0.5 shrink-0" />
              <div>
                <div className="font-medium">No se pudo conectar con Supabase. Los cambios se están guardando solo en este navegador.</div>
                <div className="text-xs mt-1 text-orange-700">{syncError}</div>
              </div>
            </div>
          )}
          {tab === "dashboard" && (
            <Dashboard
              insumos={insumos}
              productos={productos}
              recetas={recetas}
              costoProducto={costoProducto}
              insumosBajoStock={insumosBajoStock}
              productosBajoStock={productosBajoStock}
              valorInventarioInsumos={valorInventarioInsumos}
              valorInventarioProductos={valorInventarioProductos}
              totalVentasIngresos={totalVentasIngresos}
              totalVentasGanancia={totalVentasGanancia}
              ventas={ventas}
              pedidos={pedidos}
              setTab={setTab}
            />
          )}
          {tab === "insumos" && <Insumos insumos={insumos} setInsumos={setInsumos} />}
          {tab === "productos" && (
            <Productos productos={productos} setProductos={setProductos} costoProducto={costoProducto} recetas={recetas} />
          )}
          {tab === "recetas" && (
            <Recetas productos={productos} insumos={insumos} recetas={recetas} setRecetas={setRecetas} costoProducto={costoProducto} />
          )}
          {tab === "pedidos" && (
            <Pedidos
              productos={productos}
              setProductos={setProductos}
              insumos={insumos}
              setInsumos={setInsumos}
              recetas={recetas}
              pedidos={pedidos}
              setPedidos={setPedidos}
              setVentas={setVentas}
              costoProducto={costoProducto}
            />
          )}
          {tab === "ventas" && <Ventas ventas={ventas} />}
          {tab === "pvp" && <CalculadoraPVP productos={productos} recetas={recetas} costoProducto={costoProducto} />}
        </main>
      </div>
    </div>
  );
}

/* ============ COMPONENTES: PANEL ============ */
function KpiCard({ label, value, sub, accent = "amber" }) {
  const accents = { amber: "text-amber-700", orange: "text-orange-600", emerald: "text-emerald-700", stone: "text-stone-700" };
  return (
    <div className="bg-white rounded-2xl border border-stone-200 p-5 shadow-sm">
      <div className="text-xs uppercase tracking-wide text-stone-500 font-medium">{label}</div>
      <div className={`font-mono-num text-2xl font-semibold mt-2 ${accents[accent]}`}>{value}</div>
      {sub && <div className="text-xs text-stone-400 mt-1">{sub}</div>}
    </div>
  );
}

function Dashboard({ insumos, productos, recetas, costoProducto, insumosBajoStock, productosBajoStock, valorInventarioInsumos, valorInventarioProductos, totalVentasIngresos, totalVentasGanancia, ventas, pedidos, setTab }) {
  const pendientes = pedidos.filter((p) => p.estado === "Pendiente");

  const exportarTodo = () => {
    exportarExcel("goat-datos-completos", {
      Insumos: insumos.map((i) => ({
        ID: i.id,
        Nombre: i.nombre,
        "Costo (kg o un.)": i.costo,
        "Stock actual": i.stockActual,
        "Stock mínimo": i.stockMinimo,
      })),
      Productos: productos.map((p) => {
        const tieneReceta = recetas.some((r) => r.idProducto === p.id);
        const costo = tieneReceta ? costoProducto(p.id) : null;
        const margen = costo != null && p.precio > 0 ? ((p.precio - costo) / p.precio) * 100 : null;
        return {
          ID: p.id,
          Nombre: p.nombre,
          "Costo (receta)": costo != null ? Number(costo.toFixed(2)) : "",
          "Precio de venta": p.precio,
          "Margen %": margen != null ? Number(margen.toFixed(1)) : "",
          Stock: p.stock,
        };
      }),
      Recetas: recetas.map((r) => {
        const prod = productos.find((p) => p.id === r.idProducto);
        const ins = insumos.find((i) => i.id === r.idInsumo);
        return {
          "ID Producto": r.idProducto,
          Producto: prod?.nombre || "",
          "ID Insumo": r.idInsumo,
          Insumo: ins?.nombre || "",
          Cantidad: r.cantidad,
          "Subtotal costo": Number(((ins?.costo || 0) * r.cantidad).toFixed(2)),
        };
      }),
      Pedidos: pedidos.map((p) => ({
        Fecha: p.fecha,
        Producto: p.nombreProducto,
        Cliente: p.cliente,
        "Medio de pago": p.medioPago,
        Cantidad: p.cantidad,
        Estado: p.estado,
        "Fecha entrega": p.fechaEntrega || "",
      })),
      Ventas: ventas.map((v) => ({
        Fecha: v.fecha,
        Producto: v.nombreProducto,
        Cliente: v.cliente,
        "Medio de pago": v.medioPago,
        Cantidad: v.cantidad,
        Total: v.total,
        Ganancia: Number((v.ganancia || 0).toFixed(2)),
      })),
    });
  };

  return (
    <div>
      <header className="mb-6 flex items-center justify-between">
        <div>
          <h1 className="font-display text-3xl font-semibold text-stone-900">Panel general</h1>
          <p className="text-stone-500 mt-1">Foto actual del negocio: stock, costos y ventas registradas.</p>
        </div>
        <BotonExportar label="Exportar todo" onClick={exportarTodo} />
      </header>

      <div className="grid grid-cols-2 md:grid-cols-5 gap-4 mb-8">
        <KpiCard label="Valor insumos en stock" value={fmt(valorInventarioInsumos)} sub={`${insumos.length} insumos cargados`} />
        <KpiCard label="Valor productos en stock" value={fmt(valorInventarioProductos)} sub={`${productos.length} productos cargados`} />
        <button onClick={() => setTab && setTab("pedidos")} className="text-left">
          <KpiCard label="Pedidos pendientes" value={pendientes.length} sub="Esperando confirmación de entrega" accent={pendientes.length > 0 ? "orange" : "stone"} />
        </button>
        <KpiCard label="Ingresos por ventas" value={fmt(totalVentasIngresos)} sub={`${ventas.length} ventas registradas`} accent="emerald" />
        <KpiCard label="Ganancia acumulada" value={fmt(totalVentasGanancia)} sub="Ingresos − costo de insumos" accent="emerald" />
      </div>

      <div className="grid md:grid-cols-2 gap-6">
        <div className="bg-white rounded-2xl border border-stone-200 p-5">
          <h2 className="font-semibold text-stone-800 flex items-center gap-2 mb-3">
            <Icon name="warn" className="w-4 h-4 text-orange-600" /> Insumos bajo stock mínimo
          </h2>
          {insumosBajoStock.length === 0 ? (
            <p className="text-sm text-stone-400">Ningún insumo por debajo del mínimo (o no configuraste mínimos todavía).</p>
          ) : (
            <ul className="divide-y divide-stone-100">
              {insumosBajoStock.map((i) => (
                <li key={i.id} className="py-2 flex justify-between text-sm">
                  <span>{i.nombre}</span>
                  <span className="font-mono-num text-orange-600">{fmtNum(i.stockActual)} / mín {fmtNum(i.stockMinimo)}</span>
                </li>
              ))}
            </ul>
          )}
        </div>
        <div className="bg-white rounded-2xl border border-stone-200 p-5">
          <h2 className="font-semibold text-stone-800 flex items-center gap-2 mb-3">
            <Icon name="warn" className="w-4 h-4 text-orange-600" /> Productos sin stock
          </h2>
          {productosBajoStock.length === 0 ? (
            <p className="text-sm text-stone-400">Todos los productos tienen stock cargado.</p>
          ) : (
            <ul className="divide-y divide-stone-100 max-h-56 overflow-y-auto">
              {productosBajoStock.slice(0, 20).map((p) => (
                <li key={p.id} className="py-2 flex justify-between text-sm">
                  <span>{p.nombre}</span>
                  <span className="font-mono-num text-orange-600">{p.stock} un.</span>
                </li>
              ))}
              {productosBajoStock.length > 20 && (
                <li className="py-2 text-xs text-stone-400">…y {productosBajoStock.length - 20} más</li>
              )}
            </ul>
          )}
        </div>
      </div>
    </div>
  );
}

/* ============ COMPONENTES: INSUMOS ============ */
function Insumos({ insumos, setInsumos }) {
  const [q, setQ] = useState("");
  const [editing, setEditing] = useState(null); // id en edición
  const [draft, setDraft] = useState({});
  const [adding, setAdding] = useState(false);
  const [newItem, setNewItem] = useState({ id: "", nombre: "", costo: 0, stockActual: 0, stockMinimo: 0 });

  const filtered = insumos.filter((i) => i.nombre.toLowerCase().includes(q.toLowerCase()) || i.id.toLowerCase().includes(q.toLowerCase()));

  const startEdit = (i) => { setEditing(i.id); setDraft(i); };
  const saveEdit = () => { setInsumos((prev) => prev.map((i) => (i.id === editing ? { ...draft, costo: Number(draft.costo), stockActual: Number(draft.stockActual), stockMinimo: Number(draft.stockMinimo) } : i))); setEditing(null); };
  const remove = (id) => { if (confirm("¿Eliminar este insumo?")) setInsumos((prev) => prev.filter((i) => i.id !== id)); };
  const addNew = () => {
    if (!newItem.nombre.trim()) return;
    const id = newItem.id.trim() || uid("INS");
    setInsumos((prev) => [...prev, { ...newItem, id, costo: Number(newItem.costo), stockActual: Number(newItem.stockActual), stockMinimo: Number(newItem.stockMinimo) }]);
    setNewItem({ id: "", nombre: "", costo: 0, stockActual: 0, stockMinimo: 0 });
    setAdding(false);
  };

  return (
    <div>
      <header className="mb-6 flex items-center justify-between">
        <div>
          <h1 className="font-display text-3xl font-semibold text-stone-900">Insumos</h1>
          <p className="text-stone-500 mt-1">Materia prima: frutos secos, deshidratados, envases y etiquetas.</p>
        </div>
        <div className="flex gap-2">
          <BotonExportar
            onClick={() =>
              exportarExcel("goat-insumos", {
                Insumos: insumos.map((i) => ({
                  ID: i.id,
                  Nombre: i.nombre,
                  "Costo (kg o un.)": i.costo,
                  "Stock actual": i.stockActual,
                  "Stock mínimo": i.stockMinimo,
                })),
              })
            }
          />
          <button onClick={() => setAdding(true)} className="flex items-center gap-2 bg-amber-700 hover:bg-amber-800 text-white px-4 py-2 rounded-lg text-sm font-medium">
            <Icon name="plus" className="w-4 h-4" /> Nuevo insumo
          </button>
        </div>
      </header>

      <input
        value={q}
        onChange={(e) => setQ(e.target.value)}
        placeholder="Buscar insumo por nombre o ID…"
        className="w-full max-w-md mb-4 px-4 py-2 rounded-lg border border-stone-300 text-sm focus:outline-none focus:ring-2 focus:ring-amber-600"
      />

      {adding && (
        <div className="bg-white border border-amber-300 rounded-xl p-4 mb-4 grid grid-cols-2 md:grid-cols-5 gap-3 items-end">
          <Field label="ID (opcional)"><input value={newItem.id} onChange={(e) => setNewItem({ ...newItem, id: e.target.value })} className="input" /></Field>
          <Field label="Nombre"><input value={newItem.nombre} onChange={(e) => setNewItem({ ...newItem, nombre: e.target.value })} className="input" /></Field>
          <Field label="Costo / kg o un."><input type="number" value={newItem.costo} onChange={(e) => setNewItem({ ...newItem, costo: e.target.value })} className="input" /></Field>
          <Field label="Stock actual"><input type="number" value={newItem.stockActual} onChange={(e) => setNewItem({ ...newItem, stockActual: e.target.value })} className="input" /></Field>
          <Field label="Stock mínimo"><input type="number" value={newItem.stockMinimo} onChange={(e) => setNewItem({ ...newItem, stockMinimo: e.target.value })} className="input" /></Field>
          <div className="col-span-2 md:col-span-5 flex gap-2 justify-end">
            <button onClick={() => setAdding(false)} className="px-3 py-1.5 text-sm rounded-lg text-stone-500 hover:bg-stone-100">Cancelar</button>
            <button onClick={addNew} className="px-3 py-1.5 text-sm rounded-lg bg-amber-700 text-white">Guardar</button>
          </div>
        </div>
      )}

      <div className="bg-white rounded-2xl border border-stone-200 overflow-hidden">
        <table className="w-full text-sm">
          <thead className="bg-stone-100 text-stone-600 text-xs uppercase tracking-wide">
            <tr>
              <th className="text-left px-4 py-3">Insumo</th>
              <th className="text-right px-4 py-3">Costo</th>
              <th className="text-left px-4 py-3 w-64">Stock (actual / mínimo)</th>
              <th className="px-4 py-3 w-24"></th>
            </tr>
          </thead>
          <tbody className="divide-y divide-stone-100">
            {filtered.map((i) =>
              editing === i.id ? (
                <tr key={i.id} className="bg-amber-50">
                  <td className="px-4 py-2"><input value={draft.nombre} onChange={(e) => setDraft({ ...draft, nombre: e.target.value })} className="input" /></td>
                  <td className="px-4 py-2"><input type="number" value={draft.costo} onChange={(e) => setDraft({ ...draft, costo: e.target.value })} className="input text-right" /></td>
                  <td className="px-4 py-2">
                    <div className="flex gap-2">
                      <input type="number" value={draft.stockActual} onChange={(e) => setDraft({ ...draft, stockActual: e.target.value })} className="input w-20" />
                      <input type="number" value={draft.stockMinimo} onChange={(e) => setDraft({ ...draft, stockMinimo: e.target.value })} className="input w-20" />
                    </div>
                  </td>
                  <td className="px-4 py-2 text-right">
                    <button onClick={saveEdit} className="text-emerald-700 mr-2"><Icon name="check" className="w-4 h-4" /></button>
                  </td>
                </tr>
              ) : (
                <tr key={i.id} className="hover:bg-stone-50">
                  <td className="px-4 py-3">
                    <div className="font-medium text-stone-800">{i.nombre}</div>
                    <div className="text-xs text-stone-400 font-mono-num">{i.id}</div>
                  </td>
                  <td className="px-4 py-3 text-right font-mono-num">{fmt(i.costo)}</td>
                  <td className="px-4 py-3">
                    <JarLevel actual={i.stockActual} minimo={i.stockMinimo} />
                    <div className="text-xs text-stone-400 mt-1 font-mono-num">{fmtNum(i.stockActual)} / mín {fmtNum(i.stockMinimo)}</div>
                  </td>
                  <td className="px-4 py-3 text-right whitespace-nowrap">
                    <button onClick={() => startEdit(i)} className="text-stone-500 hover:text-amber-700 mr-3"><Icon name="edit" className="w-4 h-4" /></button>
                    <button onClick={() => remove(i.id)} className="text-stone-500 hover:text-red-600"><Icon name="trash" className="w-4 h-4" /></button>
                  </td>
                </tr>
              )
            )}
          </tbody>
        </table>
      </div>
      <StyleHelper />
    </div>
  );
}

/* ============ COMPONENTES: PRODUCTOS ============ */
function Productos({ productos, setProductos, costoProducto, recetas }) {
  const [q, setQ] = useState("");
  const [editing, setEditing] = useState(null);
  const [draft, setDraft] = useState({});
  const [adding, setAdding] = useState(false);
  const [newItem, setNewItem] = useState({ id: "", nombre: "", precio: 0, stock: 0 });

  const filtered = productos.filter((p) => p.nombre.toLowerCase().includes(q.toLowerCase()) || p.id.toLowerCase().includes(q.toLowerCase()));
  const tieneReceta = (id) => recetas.some((r) => r.idProducto === id);

  const startEdit = (p) => { setEditing(p.id); setDraft(p); };
  const saveEdit = () => { setProductos((prev) => prev.map((p) => (p.id === editing ? { ...draft, precio: Number(draft.precio), stock: Number(draft.stock) } : p))); setEditing(null); };
  const remove = (id) => { if (confirm("¿Eliminar este producto?")) setProductos((prev) => prev.filter((p) => p.id !== id)); };
  const addNew = () => {
    if (!newItem.nombre.trim()) return;
    const id = newItem.id.trim() || uid("PR");
    setProductos((prev) => [...prev, { ...newItem, id, precio: Number(newItem.precio), stock: Number(newItem.stock) }]);
    setNewItem({ id: "", nombre: "", precio: 0, stock: 0 });
    setAdding(false);
  };

  return (
    <div>
      <header className="mb-6 flex items-center justify-between">
        <div>
          <h1 className="font-display text-3xl font-semibold text-stone-900">Productos terminados</h1>
          <p className="text-stone-500 mt-1">Mixes y productos listos para vender, con precio y stock de bolsas.</p>
        </div>
        <div className="flex gap-2">
          <BotonExportar
            onClick={() =>
              exportarExcel("goat-productos", {
                Productos: productos.map((p) => {
                  const costo = tieneReceta(p.id) ? costoProducto(p.id) : null;
                  const margen = costo != null && p.precio > 0 ? ((p.precio - costo) / p.precio) * 100 : null;
                  return {
                    ID: p.id,
                    Nombre: p.nombre,
                    "Costo (receta)": costo != null ? Number(costo.toFixed(2)) : "",
                    "Precio de venta": p.precio,
                    "Margen %": margen != null ? Number(margen.toFixed(1)) : "",
                    Stock: p.stock,
                  };
                }),
              })
            }
          />
          <button onClick={() => setAdding(true)} className="flex items-center gap-2 bg-amber-700 hover:bg-amber-800 text-white px-4 py-2 rounded-lg text-sm font-medium">
            <Icon name="plus" className="w-4 h-4" /> Nuevo producto
          </button>
        </div>
      </header>

      <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Buscar producto por nombre o ID…" className="w-full max-w-md mb-4 px-4 py-2 rounded-lg border border-stone-300 text-sm focus:outline-none focus:ring-2 focus:ring-amber-600" />

      {adding && (
        <div className="bg-white border border-amber-300 rounded-xl p-4 mb-4 grid grid-cols-2 md:grid-cols-4 gap-3 items-end">
          <Field label="ID (opcional)"><input value={newItem.id} onChange={(e) => setNewItem({ ...newItem, id: e.target.value })} className="input" /></Field>
          <Field label="Nombre"><input value={newItem.nombre} onChange={(e) => setNewItem({ ...newItem, nombre: e.target.value })} className="input" /></Field>
          <Field label="Precio de venta"><input type="number" value={newItem.precio} onChange={(e) => setNewItem({ ...newItem, precio: e.target.value })} className="input" /></Field>
          <Field label="Stock (bolsas)"><input type="number" value={newItem.stock} onChange={(e) => setNewItem({ ...newItem, stock: e.target.value })} className="input" /></Field>
          <div className="col-span-2 md:col-span-4 flex gap-2 justify-end">
            <button onClick={() => setAdding(false)} className="px-3 py-1.5 text-sm rounded-lg text-stone-500 hover:bg-stone-100">Cancelar</button>
            <button onClick={addNew} className="px-3 py-1.5 text-sm rounded-lg bg-amber-700 text-white">Guardar</button>
          </div>
        </div>
      )}

      <div className="bg-white rounded-2xl border border-stone-200 overflow-hidden">
        <table className="w-full text-sm">
          <thead className="bg-stone-100 text-stone-600 text-xs uppercase tracking-wide">
            <tr>
              <th className="text-left px-4 py-3">Producto</th>
              <th className="text-right px-4 py-3">Costo (receta)</th>
              <th className="text-right px-4 py-3">Precio venta</th>
              <th className="text-right px-4 py-3">Margen</th>
              <th className="text-right px-4 py-3">Stock</th>
              <th className="px-4 py-3 w-24"></th>
            </tr>
          </thead>
          <tbody className="divide-y divide-stone-100">
            {filtered.map((p) => {
              const costo = tieneReceta(p.id) ? costoProducto(p.id) : null;
              const margen = costo != null && p.precio > 0 ? ((p.precio - costo) / p.precio) * 100 : null;
              return editing === p.id ? (
                <tr key={p.id} className="bg-amber-50">
                  <td className="px-4 py-2"><input value={draft.nombre} onChange={(e) => setDraft({ ...draft, nombre: e.target.value })} className="input" /></td>
                  <td className="px-4 py-2 text-right text-stone-400">{costo != null ? fmt(costo) : "—"}</td>
                  <td className="px-4 py-2"><input type="number" value={draft.precio} onChange={(e) => setDraft({ ...draft, precio: e.target.value })} className="input text-right" /></td>
                  <td className="px-4 py-2 text-right text-stone-400">—</td>
                  <td className="px-4 py-2"><input type="number" value={draft.stock} onChange={(e) => setDraft({ ...draft, stock: e.target.value })} className="input text-right" /></td>
                  <td className="px-4 py-2 text-right"><button onClick={saveEdit} className="text-emerald-700"><Icon name="check" className="w-4 h-4" /></button></td>
                </tr>
              ) : (
                <tr key={p.id} className="hover:bg-stone-50">
                  <td className="px-4 py-3">
                    <div className="font-medium text-stone-800">{p.nombre}</div>
                    <div className="text-xs text-stone-400 font-mono-num">{p.id}</div>
                  </td>
                  <td className="px-4 py-3 text-right font-mono-num text-stone-500">{costo != null ? fmt(costo) : "—"}</td>
                  <td className="px-4 py-3 text-right font-mono-num">{fmt(p.precio)}</td>
                  <td className="px-4 py-3 text-right font-mono-num">
                    {margen != null ? <span className={margen < 15 ? "text-orange-600" : "text-emerald-700"}>{fmtNum(margen, 1)}%</span> : "—"}
                  </td>
                  <td className={`px-4 py-3 text-right font-mono-num ${p.stock <= 0 ? "text-orange-600" : ""}`}>{p.stock}</td>
                  <td className="px-4 py-3 text-right whitespace-nowrap">
                    <button onClick={() => startEdit(p)} className="text-stone-500 hover:text-amber-700 mr-3"><Icon name="edit" className="w-4 h-4" /></button>
                    <button onClick={() => remove(p.id)} className="text-stone-500 hover:text-red-600"><Icon name="trash" className="w-4 h-4" /></button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      <StyleHelper />
    </div>
  );
}

/* ============ COMPONENTES: RECETAS ============ */
function Recetas({ productos, insumos, recetas, setRecetas, costoProducto }) {
  const [idProducto, setIdProducto] = useState(productos[0]?.id || "");
  const [nuevoInsumo, setNuevoInsumo] = useState(insumos[0]?.id || "");
  const [cantidad, setCantidad] = useState(0.1);

  const itemsReceta = recetas.filter((r) => r.idProducto === idProducto);
  const producto = productos.find((p) => p.id === idProducto);
  const costo = costoProducto(idProducto);
  const margen = producto && producto.precio > 0 ? ((producto.precio - costo) / producto.precio) * 100 : null;

  const addLinea = () => {
    if (!idProducto || !nuevoInsumo) return;
    if (itemsReceta.some((r) => r.idInsumo === nuevoInsumo)) { alert("Ese insumo ya está en la receta."); return; }
    setRecetas((prev) => [...prev, { idProducto, idInsumo: nuevoInsumo, cantidad: Number(cantidad) }]);
  };
  const updateCantidad = (idInsumo, val) => {
    setRecetas((prev) => prev.map((r) => (r.idProducto === idProducto && r.idInsumo === idInsumo ? { ...r, cantidad: Number(val) } : r)));
  };
  const removeLinea = (idInsumo) => setRecetas((prev) => prev.filter((r) => !(r.idProducto === idProducto && r.idInsumo === idInsumo)));

  return (
    <div>
      <header className="mb-6 flex items-center justify-between">
        <div>
          <h1 className="font-display text-3xl font-semibold text-stone-900">Recetas</h1>
          <p className="text-stone-500 mt-1">Definí qué insumos (y cuánta cantidad) lleva cada producto para calcular su costo real.</p>
        </div>
        <BotonExportar
          label="Exportar todas"
          onClick={() =>
            exportarExcel("goat-recetas", {
              Recetas: recetas.map((r) => {
                const prod = productos.find((p) => p.id === r.idProducto);
                const ins = insumos.find((i) => i.id === r.idInsumo);
                return {
                  "ID Producto": r.idProducto,
                  Producto: prod?.nombre || "",
                  "ID Insumo": r.idInsumo,
                  Insumo: ins?.nombre || "",
                  Cantidad: r.cantidad,
                  "Subtotal costo": Number(((ins?.costo || 0) * r.cantidad).toFixed(2)),
                };
              }),
            })
          }
        />
      </header>

      <div className="grid md:grid-cols-3 gap-6">
        <div className="md:col-span-1">
          <label className="text-xs uppercase tracking-wide text-stone-500 font-medium">Producto</label>
          <select value={idProducto} onChange={(e) => setIdProducto(e.target.value)} className="input mt-1 w-full">
            {productos.map((p) => (
              <option key={p.id} value={p.id}>{p.nombre}</option>
            ))}
          </select>

          <div className="mt-6 bg-white rounded-2xl border border-stone-200 p-4">
            <div className="text-xs uppercase text-stone-500 mb-2">Resumen</div>
            <div className="flex justify-between text-sm py-1"><span>Costo total receta</span><span className="font-mono-num">{fmt(costo)}</span></div>
            <div className="flex justify-between text-sm py-1"><span>Precio de venta</span><span className="font-mono-num">{fmt(producto?.precio || 0)}</span></div>
            <div className="flex justify-between text-sm py-1 border-t border-stone-100 mt-1 pt-2">
              <span>Margen</span>
              <span className={`font-mono-num font-semibold ${margen != null && margen < 15 ? "text-orange-600" : "text-emerald-700"}`}>
                {margen != null ? fmtNum(margen, 1) + "%" : "—"}
              </span>
            </div>
          </div>
        </div>

        <div className="md:col-span-2">
          <div className="bg-white rounded-2xl border border-stone-200 overflow-hidden">
            <table className="w-full text-sm">
              <thead className="bg-stone-100 text-stone-600 text-xs uppercase tracking-wide">
                <tr>
                  <th className="text-left px-4 py-3">Insumo</th>
                  <th className="text-right px-4 py-3 w-32">Cantidad</th>
                  <th className="text-right px-4 py-3 w-32">Subtotal</th>
                  <th className="px-4 py-3 w-16"></th>
                </tr>
              </thead>
              <tbody className="divide-y divide-stone-100">
                {itemsReceta.length === 0 && (
                  <tr><td colSpan={4} className="px-4 py-6 text-center text-stone-400">Sin insumos cargados para este producto todavía.</td></tr>
                )}
                {itemsReceta.map((r) => {
                  const ins = insumos.find((i) => i.id === r.idInsumo);
                  return (
                    <tr key={r.idInsumo}>
                      <td className="px-4 py-2">{ins?.nombre || r.idInsumo}</td>
                      <td className="px-4 py-2">
                        <input type="number" step="0.01" value={r.cantidad} onChange={(e) => updateCantidad(r.idInsumo, e.target.value)} className="input text-right" />
                      </td>
                      <td className="px-4 py-2 text-right font-mono-num">{fmt((ins?.costo || 0) * r.cantidad)}</td>
                      <td className="px-4 py-2 text-right"><button onClick={() => removeLinea(r.idInsumo)} className="text-stone-400 hover:text-red-600"><Icon name="trash" className="w-4 h-4" /></button></td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          <div className="mt-4 bg-white border border-stone-200 rounded-xl p-4 flex flex-wrap items-end gap-3">
            <div>
              <label className="text-xs uppercase tracking-wide text-stone-500">Agregar insumo</label>
              <select value={nuevoInsumo} onChange={(e) => setNuevoInsumo(e.target.value)} className="input mt-1">
                {insumos.map((i) => <option key={i.id} value={i.id}>{i.nombre}</option>)}
              </select>
            </div>
            <div>
              <label className="text-xs uppercase tracking-wide text-stone-500">Cantidad (kg/un.)</label>
              <input type="number" step="0.01" value={cantidad} onChange={(e) => setCantidad(e.target.value)} className="input mt-1 w-28" />
            </div>
            <button onClick={addLinea} className="flex items-center gap-2 bg-amber-700 hover:bg-amber-800 text-white px-4 py-2 rounded-lg text-sm font-medium">
              <Icon name="plus" className="w-4 h-4" /> Agregar
            </button>
          </div>
        </div>
      </div>
      <StyleHelper />
    </div>
  );
}

/* ============ COMPONENTES: VENTAS ============ */
/* ============ COMPONENTES: PEDIDOS ============ */
function Pedidos({ productos, setProductos, insumos, setInsumos, recetas, pedidos, setPedidos, setVentas, costoProducto }) {
  const [idProducto, setIdProducto] = useState(productos[0]?.id || "");
  const [cantidad, setCantidad] = useState(1);
  const [cliente, setCliente] = useState("");
  const [medioPago, setMedioPago] = useState("Efectivo");
  const [fecha, setFecha] = useState(todayISO());
  const [verEntregados, setVerEntregados] = useState(false);

  const producto = productos.find((p) => p.id === idProducto);
  const pendientes = pedidos.filter((p) => p.estado === "Pendiente");
  const entregados = pedidos.filter((p) => p.estado === "Entregado");

  const crearPedido = () => {
    if (!producto || cantidad <= 0) return;
    setPedidos((prev) => [
      {
        id: uid("PED"),
        fecha,
        cliente: cliente || "Consumidor final",
        idProducto,
        nombreProducto: producto.nombre,
        cantidad: Number(cantidad),
        medioPago,
        estado: "Pendiente",
      },
      ...prev,
    ]);
    setCantidad(1);
    setCliente("");
  };

  const confirmarEntrega = (pedido) => {
    const costoUnit = costoProducto(pedido.idProducto);
    const total = (productos.find((p) => p.id === pedido.idProducto)?.precio || 0) * pedido.cantidad;
    const ganancia = ((productos.find((p) => p.id === pedido.idProducto)?.precio || 0) - costoUnit) * pedido.cantidad;

    // Pasa a Ventas
    setVentas((prev) => [
      { id: uid("VTA"), fecha: todayISO(), cliente: pedido.cliente, idProducto: pedido.idProducto, nombreProducto: pedido.nombreProducto, cantidad: pedido.cantidad, medioPago: pedido.medioPago, total, ganancia },
      ...prev,
    ]);

    // Descuenta stock de producto
    setProductos((prev) => prev.map((p) => (p.id === pedido.idProducto ? { ...p, stock: p.stock - pedido.cantidad } : p)));

    // Descuenta stock de insumos según receta
    const itemsReceta = recetas.filter((r) => r.idProducto === pedido.idProducto);
    if (itemsReceta.length) {
      setInsumos((prev) =>
        prev.map((ins) => {
          const linea = itemsReceta.find((r) => r.idInsumo === ins.id);
          return linea ? { ...ins, stockActual: Math.max(0, ins.stockActual - linea.cantidad * pedido.cantidad) } : ins;
        })
      );
    }

    // Marca el pedido como entregado
    setPedidos((prev) => prev.map((p) => (p.id === pedido.id ? { ...p, estado: "Entregado", fechaEntrega: todayISO() } : p)));
  };

  const cancelarPedido = (id) => {
    if (confirm("¿Cancelar este pedido? No se va a registrar como venta.")) {
      setPedidos((prev) => prev.filter((p) => p.id !== id));
    }
  };

  return (
    <div>
      <header className="mb-6 flex items-center justify-between">
        <div>
          <h1 className="font-display text-3xl font-semibold text-stone-900">Pedidos</h1>
          <p className="text-stone-500 mt-1">Cargá el pedido primero; recién cuando confirmes la entrega pasa a Ventas y se descuenta stock.</p>
        </div>
        <BotonExportar
          onClick={() =>
            exportarExcel("goat-pedidos", {
              Pedidos: pedidos.map((p) => ({
                Fecha: p.fecha,
                Producto: p.nombreProducto,
                Cliente: p.cliente,
                "Medio de pago": p.medioPago,
                Cantidad: p.cantidad,
                Estado: p.estado,
                "Fecha entrega": p.fechaEntrega || "",
              })),
            })
          }
        />
      </header>

      <div className="bg-white rounded-2xl border border-stone-200 p-5 mb-6 grid md:grid-cols-5 gap-3 items-end">
        <Field label="Producto">
          <select value={idProducto} onChange={(e) => setIdProducto(e.target.value)} className="input">
            {productos.map((p) => <option key={p.id} value={p.id}>{p.nombre} (stock: {p.stock})</option>)}
          </select>
        </Field>
        <Field label="Cantidad"><input type="number" min="1" value={cantidad} onChange={(e) => setCantidad(e.target.value)} className="input" /></Field>
        <Field label="Cliente"><input value={cliente} onChange={(e) => setCliente(e.target.value)} placeholder="Opcional" className="input" /></Field>
        <Field label="Medio de pago">
          <select value={medioPago} onChange={(e) => setMedioPago(e.target.value)} className="input">
            <option>Efectivo</option><option>Transferencia</option><option>Tarjeta débito</option><option>Tarjeta crédito</option><option>Mercado Pago</option>
          </select>
        </Field>
        <Field label="Fecha del pedido"><input type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} className="input" /></Field>
      </div>
      <div className="mb-6 flex items-center justify-between bg-amber-50 border border-amber-200 rounded-xl px-4 py-3">
        <div className="text-sm text-stone-600">
          Total estimado: <span className="font-mono-num font-semibold">{fmt((producto?.precio || 0) * cantidad)}</span>
        </div>
        <button onClick={crearPedido} className="flex items-center gap-2 bg-amber-700 hover:bg-amber-800 text-white px-4 py-2 rounded-lg text-sm font-medium">
          <Icon name="plus" className="w-4 h-4" /> Crear pedido
        </button>
      </div>

      <h2 className="font-semibold text-stone-800 mb-2">Pendientes de entrega ({pendientes.length})</h2>
      <div className="bg-white rounded-2xl border border-stone-200 overflow-hidden mb-8">
        <table className="w-full text-sm">
          <thead className="bg-stone-100 text-stone-600 text-xs uppercase tracking-wide">
            <tr>
              <th className="text-left px-4 py-3">Fecha</th>
              <th className="text-left px-4 py-3">Producto</th>
              <th className="text-left px-4 py-3">Cliente</th>
              <th className="text-left px-4 py-3">Medio de pago</th>
              <th className="text-right px-4 py-3">Cant.</th>
              <th className="px-4 py-3 w-56"></th>
            </tr>
          </thead>
          <tbody className="divide-y divide-stone-100">
            {pendientes.length === 0 && (
              <tr><td colSpan={6} className="px-4 py-6 text-center text-stone-400">No hay pedidos pendientes.</td></tr>
            )}
            {pendientes.map((p) => (
              <tr key={p.id} className="hover:bg-stone-50">
                <td className="px-4 py-2 font-mono-num text-stone-500">{p.fecha}</td>
                <td className="px-4 py-2">{p.nombreProducto}</td>
                <td className="px-4 py-2">{p.cliente}</td>
                <td className="px-4 py-2">{p.medioPago}</td>
                <td className="px-4 py-2 text-right font-mono-num">{p.cantidad}</td>
                <td className="px-4 py-2 text-right whitespace-nowrap">
                  <button onClick={() => confirmarEntrega(p)} className="inline-flex items-center gap-1 bg-emerald-700 hover:bg-emerald-800 text-white px-3 py-1.5 rounded-lg text-xs font-medium mr-2">
                    <Icon name="check" className="w-3.5 h-3.5" /> Confirmar entrega
                  </button>
                  <button onClick={() => cancelarPedido(p.id)} className="text-stone-400 hover:text-red-600">
                    <Icon name="trash" className="w-4 h-4" />
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <button onClick={() => setVerEntregados((v) => !v)} className="text-sm text-stone-500 hover:text-stone-700 flex items-center gap-1 mb-2">
        {verEntregados ? "Ocultar" : "Ver"} entregados ({entregados.length})
      </button>
      {verEntregados && (
        <div className="bg-white rounded-2xl border border-stone-200 overflow-hidden">
          <table className="w-full text-sm">
            <thead className="bg-stone-100 text-stone-600 text-xs uppercase tracking-wide">
              <tr>
                <th className="text-left px-4 py-3">Fecha pedido</th>
                <th className="text-left px-4 py-3">Fecha entrega</th>
                <th className="text-left px-4 py-3">Producto</th>
                <th className="text-left px-4 py-3">Cliente</th>
                <th className="text-right px-4 py-3">Cant.</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-stone-100">
              {entregados.map((p) => (
                <tr key={p.id}>
                  <td className="px-4 py-2 font-mono-num text-stone-500">{p.fecha}</td>
                  <td className="px-4 py-2 font-mono-num text-stone-500">{p.fechaEntrega}</td>
                  <td className="px-4 py-2">{p.nombreProducto}</td>
                  <td className="px-4 py-2">{p.cliente}</td>
                  <td className="px-4 py-2 text-right font-mono-num">{p.cantidad}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
      <StyleHelper />
    </div>
  );
}

/* ============ COMPONENTES: VENTAS ============ */
function Ventas({ ventas }) {
  return (
    <div>
      <header className="mb-6 flex items-center justify-between">
        <div>
          <h1 className="font-display text-3xl font-semibold text-stone-900">Ventas</h1>
          <p className="text-stone-500 mt-1">Historial de ventas confirmadas (pedidos ya entregados).</p>
        </div>
        <BotonExportar
          onClick={() =>
            exportarExcel("goat-ventas", {
              Ventas: ventas.map((v) => ({
                Fecha: v.fecha,
                Producto: v.nombreProducto,
                Cliente: v.cliente,
                "Medio de pago": v.medioPago,
                Cantidad: v.cantidad,
                Total: v.total,
                Ganancia: Number((v.ganancia || 0).toFixed(2)),
              })),
            })
          }
        />
      </header>

      <div className="bg-white rounded-2xl border border-stone-200 overflow-hidden">
        <table className="w-full text-sm">
          <thead className="bg-stone-100 text-stone-600 text-xs uppercase tracking-wide">
            <tr>
              <th className="text-left px-4 py-3">Fecha</th>
              <th className="text-left px-4 py-3">Producto</th>
              <th className="text-left px-4 py-3">Cliente</th>
              <th className="text-left px-4 py-3">Medio de pago</th>
              <th className="text-right px-4 py-3">Cant.</th>
              <th className="text-right px-4 py-3">Total</th>
              <th className="text-right px-4 py-3">Ganancia</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-stone-100">
            {ventas.length === 0 && (
              <tr><td colSpan={7} className="px-4 py-6 text-center text-stone-400">Todavía no hay ventas confirmadas. Confirmá la entrega de un pedido para que aparezca acá.</td></tr>
            )}
            {ventas.map((v) => (
              <tr key={v.id} className="hover:bg-stone-50">
                <td className="px-4 py-2 font-mono-num text-stone-500">{v.fecha}</td>
                <td className="px-4 py-2">{v.nombreProducto}</td>
                <td className="px-4 py-2">{v.cliente}</td>
                <td className="px-4 py-2">{v.medioPago}</td>
                <td className="px-4 py-2 text-right font-mono-num">{v.cantidad}</td>
                <td className="px-4 py-2 text-right font-mono-num">{fmt(v.total)}</td>
                <td className="px-4 py-2 text-right font-mono-num text-emerald-700">{fmt(v.ganancia)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <StyleHelper />
    </div>
  );
}

/* ============ COMPONENTES: CALCULADORA PVP ============ */
function CalculadoraPVP({ productos, recetas, costoProducto }) {
  const [idProducto, setIdProducto] = useState(productos[0]?.id || "");
  const [margen, setMargen] = useState(30);
  const producto = productos.find((p) => p.id === idProducto);
  const items = recetas.filter((r) => r.idProducto === idProducto);
  const costo = costoProducto(idProducto);
  const pvpSugerido = costo / (1 - margen / 100);
  const gananciaSugerida = pvpSugerido - costo;
  const margenActual = producto && producto.precio > 0 ? ((producto.precio - costo) / producto.precio) * 100 : null;

  return (
    <div>
      <header className="mb-6">
        <h1 className="font-display text-3xl font-semibold text-stone-900">Calculadora de PVP</h1>
        <p className="text-stone-500 mt-1">Precio de venta sugerido según el costo de la receta y el margen que quieras ganar.</p>
      </header>

      <div className="grid md:grid-cols-2 gap-6">
        <div>
          <label className="text-xs uppercase tracking-wide text-stone-500 font-medium">Producto</label>
          <select value={idProducto} onChange={(e) => setIdProducto(e.target.value)} className="input mt-1 w-full">
            {productos.map((p) => <option key={p.id} value={p.id}>{p.nombre}</option>)}
          </select>

          <div className="mt-4 bg-white rounded-2xl border border-stone-200 overflow-hidden">
            <table className="w-full text-sm">
              <thead className="bg-stone-100 text-stone-600 text-xs uppercase"><tr><th className="text-left px-4 py-2">Insumo</th><th className="text-right px-4 py-2">Costo</th></tr></thead>
              <tbody className="divide-y divide-stone-100">
                {items.length === 0 && <tr><td colSpan={2} className="px-4 py-4 text-center text-stone-400">Este producto no tiene receta cargada. Andá a la pestaña Recetas.</td></tr>}
                {items.map((r) => (
                  <tr key={r.idInsumo}>
                    <td className="px-4 py-2">{r.nombreInsumo || r.idInsumo}</td>
                    <td className="px-4 py-2 text-right font-mono-num">{fmt(costoProducto ? (recetas.find(x=>x===r) && 0) : 0)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        <div className="bg-white rounded-2xl border border-stone-200 p-5">
          <label className="text-xs uppercase tracking-wide text-stone-500 font-medium">Margen deseado</label>
          <div className="flex items-center gap-3 mt-1">
            <input type="range" min="0" max="80" value={margen} onChange={(e) => setMargen(Number(e.target.value))} className="flex-1" />
            <span className="font-mono-num w-14 text-right">{margen}%</span>
          </div>

          <div className="mt-6 space-y-2">
            <Row label="Costo total (receta)" value={fmt(costo)} />
            <Row label="PVP sugerido" value={fmt(pvpSugerido)} big />
            <Row label="Ganancia por unidad" value={fmt(gananciaSugerida)} accent="emerald" />
            <div className="border-t border-stone-100 my-3" />
            <Row label="Precio actual cargado" value={fmt(producto?.precio || 0)} />
            <Row label="Margen actual" value={margenActual != null ? fmtNum(margenActual, 1) + "%" : "—"} accent={margenActual != null && margenActual < margen ? "orange" : "emerald"} />
          </div>
        </div>
      </div>
      <StyleHelper />
    </div>
  );
}

function Row({ label, value, big, accent }) {
  const colors = { emerald: "text-emerald-700", orange: "text-orange-600" };
  return (
    <div className="flex justify-between items-center">
      <span className="text-sm text-stone-600">{label}</span>
      <span className={`font-mono-num ${big ? "text-xl font-semibold" : ""} ${accent ? colors[accent] : ""}`}>{value}</span>
    </div>
  );
}

function Field({ label, children }) {
  return (
    <div>
      <label className="text-xs uppercase tracking-wide text-stone-500 block mb-1">{label}</label>
      {children}
    </div>
  );
}

function StyleHelper() {
  return (
    <style>{`
      .input { border: 1px solid #d6d3d1; border-radius: 0.5rem; padding: 0.5rem 0.75rem; font-size: 0.875rem; width: 100%; background: white; }
      .input:focus { outline: none; box-shadow: 0 0 0 2px #b45309; border-color: #b45309; }
    `}</style>
  );
}
