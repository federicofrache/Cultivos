import React, { useState, useEffect, useRef } from "react";
import { auth, db, storage } from "./firebase";
import {
  onAuthStateChanged, signInWithEmailAndPassword, createUserWithEmailAndPassword,
  signOut, sendPasswordResetEmail,
} from "firebase/auth";
import {
  collection, onSnapshot, addDoc, updateDoc, deleteDoc, doc, query, where, setDoc,
} from "firebase/firestore";
import { ref, uploadBytes, getDownloadURL } from "firebase/storage";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
} from "recharts";
import {
  Sprout, Snowflake, Sun, Wheat, Receipt, Mic, Camera, MapPin, TrendingUp,
  TrendingDown, Plus, Trash2, Loader2, LogOut, ChevronRight, ChevronLeft,
  Truck, DollarSign, FileText, AlertCircle, CheckCircle2, Paperclip, Pencil, X, Package, Boxes, Copy, Search,
  Download, RotateCcw, Trash, LayoutDashboard, Menu, Users,
} from "lucide-react";

/* ------------------------------------------------------------------ */
/*  Constantes de dominio                                              */
/* ------------------------------------------------------------------ */
const CULTIVOS_VERANO = ["Soja", "Maíz", "Sorgo"];
const CULTIVOS_INVIERNO = ["Colza", "Carinata", "Trigo", "Cebada", "Lupino", "Camelina"];
const CAT_COLOR = { verano: "#C68A2E", invierno: "#3D6E8C" };
const UNIDADES_INSUMO = ["Litros", "Unidades", "Kg", "Bolsas"];
const abrevUnidad = (u) => ({ Litros: "L", Unidades: "u.", Kg: "kg", Bolsas: "b." }[u] || u || "L");
const CATEGORIAS_GASTO = ["Insumos", "Servicio", "Renta", "Seguro", "Asesoramiento", "Otro"];
const GASTO_CAT_COLOR = {
  Insumo: "#5B4B8A", Insumos: "#5B4B8A", Servicio: "#8A6D3B", Renta: "#8C3D3D", Seguro: "#3D6E8C", Asesoramiento: "#5C7A4E", Otro: "#7A7267",
};
const PALETA_SOCIOS = ["#8A6D3B", "#3D6E8C", "#5C7A4E", "#8C3D3D", "#5B4B8A", "#C68A2E", "#4E7A7A", "#9E5B2E"];
function colorPorTexto(texto) {
  if (texto === "Sin asignar") return "#B5AF98";
  let hash = 0;
  for (let i = 0; i < texto.length; i++) hash = texto.charCodeAt(i) + ((hash << 5) - hash);
  return PALETA_SOCIOS[Math.abs(hash) % PALETA_SOCIOS.length];
}

const fmt = (n, decimals = 0) => {
  if (n === null || n === undefined || isNaN(n)) return "-";
  return new Intl.NumberFormat("es-UY", { minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(n);
};
const fmtUSD = (n) => `U$S ${fmt(n, 0)}`;
const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36);

// Calcula el costo de consumir `litrosNuevos` de un insumo, tomando primero de las
// compras más antiguas que todavía tengan saldo (FIFO), salteando lo ya consumido
// por otros gastos anteriores (`litrosYaConsumidos`).
function costoFIFO(comprasDelInsumo, litrosYaConsumidos, litrosNuevos) {
  const ordenadas = [...comprasDelInsumo].sort((a, b) => (a.fecha || "").localeCompare(b.fecha || "") || String(a.id).localeCompare(String(b.id)));
  let restanteYaConsumido = litrosYaConsumidos;
  let restanteNuevo = litrosNuevos;
  let costoTotal = 0;
  let litrosAsignados = 0;
  for (const c of ordenadas) {
    let disponibleAqui = Number(c.litros || 0);
    if (restanteYaConsumido > 0) {
      const salteado = Math.min(disponibleAqui, restanteYaConsumido);
      disponibleAqui -= salteado;
      restanteYaConsumido -= salteado;
    }
    if (disponibleAqui <= 0 || restanteNuevo <= 0) continue;
    const tomar = Math.min(disponibleAqui, restanteNuevo);
    const precioPorLitro = c.litros ? Number(c.precio || 0) / Number(c.litros) : 0;
    costoTotal += tomar * precioPorLitro;
    litrosAsignados += tomar;
    restanteNuevo -= tomar;
  }
  if (restanteNuevo > 0 && ordenadas.length) {
    const ultima = ordenadas[ordenadas.length - 1];
    const precioPorLitro = ultima.litros ? Number(ultima.precio || 0) / Number(ultima.litros) : 0;
    costoTotal += restanteNuevo * precioPorLitro;
    litrosAsignados += restanteNuevo;
  }
  return { costoTotal, costoPromedioEfectivo: litrosAsignados ? costoTotal / litrosAsignados : 0 };
}

/* ------------------------------------------------------------------ */
/*  Claude API (voz e imagen de factura)                               */
/* ------------------------------------------------------------------ */
async function askClaudeJSON(content) {
  const res = await fetch("/api/ask-claude", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ content }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data?.error || "Error consultando la IA");
  const text = (data.content || []).map((b) => b.text || "").join("\n");
  return JSON.parse(text.replace(/```json/g, "").replace(/```/g, "").trim());
}
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result.split(",")[1]);
    r.onerror = reject;
    r.readAsDataURL(file);
  });
}

function exportarExcel(nombreArchivo, hojas) {
  const wb = XLSX.utils.book_new();
  hojas.forEach((h) => {
    const ws = XLSX.utils.json_to_sheet(h.filas);
    XLSX.utils.book_append_sheet(wb, ws, h.nombre.slice(0, 31));
  });
  XLSX.writeFile(wb, `${nombreArchivo}.xlsx`);
}

const GLOBAL_STYLES = `
  @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');
  :root{
    --bg:#F1EEE1; --panel:#FBFAF4; --ink:#26241C; --soil:#2C3A24; --soil-light:#3E4F32;
    --gold:#C68A2E; --frost:#3D6E8C; --rust:#8C3D3D; --line:#DEDACB;
    --font-display:'Fraunces',serif; --font-body:'Inter',sans-serif; --font-mono:'IBM Plex Mono',monospace;
  }
  body{margin:0;}
  .cc-h{font-family:var(--font-display);}
  .cc-mono{font-family:var(--font-mono);}
  .cc-card{background:var(--panel); border:1px solid var(--line); border-radius:10px;}
  .cc-input{width:100%; border:1.5px solid #C9C3AC; border-radius:8px; padding:12px 14px; background:#fff; font-family:var(--font-body); font-size:15.5px; transition:border-color .12s ease;}
  .cc-input:hover{border-color:#A69B7A;}
  .cc-input:focus{outline:none; border-color:var(--gold); box-shadow:0 0 0 3px rgba(198,138,46,0.18);}
  .cc-input:disabled{background:#F1EFE6; color:#8A8570; border-color:var(--line);}
  label{font-size:13px !important; font-weight:600; color:#5A5647 !important; display:block; margin-bottom:3px;}
  .cc-btn{border-radius:7px; padding:9px 16px; font-weight:600; font-size:14px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; border:none;}
  .cc-btn-primary{background:var(--soil); color:#fff;}
  .cc-btn-primary:hover{background:var(--soil-light);}
  .cc-btn-ghost{background:transparent; border:1px solid var(--line); color:var(--ink);}
  .cc-btn-ghost:hover{background:rgba(0,0,0,.03);}
  .cc-chip{font-size:11px; font-weight:600; padding:3px 9px; border-radius:99px; letter-spacing:.02em;}
  .cc-sub{display:flex; align-items:center; gap:6px; padding:8px 14px; font-size:13px; font-weight:600; border-radius:8px 8px 0 0; cursor:pointer;}

  html, body { overflow-x: hidden; }
  .flex{display:flex;} .flex-col{flex-direction:column;} .flex-wrap{flex-wrap:wrap;}
  .items-center{align-items:center;} .items-start{align-items:flex-start;} .items-end{align-items:flex-end;}
  .justify-between{justify-content:space-between;} .justify-center{justify-content:center;} .justify-end{justify-content:flex-end;}
  .grid{display:grid;}
  .gap-1{gap:4px;} .gap-2{gap:8px;} .gap-3{gap:12px;} .gap-4{gap:16px;}
  .space-y-2>*+*{margin-top:8px;} .space-y-3>*+*{margin-top:12px;} .space-y-4>*+*{margin-top:16px;} .space-y-5>*+*{margin-top:20px;} .space-y-6>*+*{margin-top:24px;}
  .mb-1{margin-bottom:4px;} .mb-2{margin-bottom:8px;} .mb-3{margin-bottom:12px;} .mb-4{margin-bottom:16px;} .mb-5{margin-bottom:20px;}
  .mt-2{margin-top:8px;} .mt-3{margin-top:12px;}
  .p-4{padding:16px;} .p-6{padding:24px;}
  .px-3{padding-left:12px;padding-right:12px;} .px-4{padding-left:16px;padding-right:16px;} .px-6{padding-left:24px;padding-right:24px;}
  .py-2{padding-top:8px;padding-bottom:8px;} .py-4{padding-top:16px;padding-bottom:16px;} .py-6{padding-top:24px;padding-bottom:24px;} .py-16{padding-top:64px;padding-bottom:64px;}
  .max-w-6xl{max-width:72rem;} .mx-auto{margin-left:auto;margin-right:auto;}
  .w-full{width:100%;} .text-center{text-align:center;} .text-right{text-align:right;}
  .cursor-pointer{cursor:pointer;}
  .animate-spin{animation:cc-spin 1s linear infinite;} @keyframes cc-spin{to{transform:rotate(360deg);}}

  .cc-card.overflow-hidden{overflow-x:auto; -webkit-overflow-scrolling:touch;}
  .cc-card.overflow-hidden table{min-width:520px;}

  @media (max-width: 640px){
    .cc-input{font-size:16px;}
    .px-6{padding-left:14px;padding-right:14px;}
    .py-4{padding-top:12px;padding-bottom:12px;}
  }

  .cc-nav-desktop{display:flex;}
  .cc-nav-toggle{display:none;}
  @media (max-width: 760px){
    .cc-nav-desktop{display:none;}
    .cc-nav-toggle{display:inline-flex !important;}
  }
`;

/* ------------------------------------------------------------------ */
/*  Login                                                               */
/* ------------------------------------------------------------------ */
function Login() {
  const [email, setEmail] = useState("");
  const [pass, setPass] = useState("");
  const [modo, setModo] = useState("ingresar");
  const [error, setError] = useState("");
  const [cargando, setCargando] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    setError(""); setCargando(true);
    try {
      if (modo === "ingresar") await signInWithEmailAndPassword(auth, email, pass);
      else await createUserWithEmailAndPassword(auth, email, pass);
    } catch (err) {
      const msgs = {
        "auth/invalid-credential": "Usuario o contraseña incorrectos.",
        "auth/email-already-in-use": "Ese email ya tiene una cuenta creada.",
        "auth/weak-password": "La contraseña debe tener al menos 6 caracteres.",
        "auth/invalid-email": "El email no es válido.",
      };
      setError(msgs[err.code] || "Ocurrió un error. Probá de nuevo.");
    } finally { setCargando(false); }
  };
  const recuperar = async () => {
    if (!email) { setError("Escribí tu email arriba y volvé a tocar 'Olvidé mi contraseña'."); return; }
    try { await sendPasswordResetEmail(auth, email); setError("Te enviamos un email para restablecer la contraseña."); }
    catch { setError("No pudimos enviar el email de recuperación."); }
  };

  return (
    <div style={{ background: "var(--bg)", minHeight: "100vh" }} className="flex items-center justify-center">
      <style>{GLOBAL_STYLES}</style>
      <form onSubmit={submit} className="cc-card p-6" style={{ width: 340 }}>
        <div className="flex items-center gap-2 justify-center mb-1"><Wheat color="var(--gold)" size={26} /></div>
        <div className="cc-h text-center" style={{ fontSize: 20, fontWeight: 600 }}>Campo & Costo</div>
        <div className="text-center mb-5" style={{ fontSize: 12, color: "#8A8570" }}>Campañas · Cultivos · Gastos e Ingresos</div>
        <label style={{ fontSize: 12, color: "#8A8570" }}>Email</label>
        <input className="cc-input mb-3" type="email" required value={email} onChange={(e) => setEmail(e.target.value)} />
        <label style={{ fontSize: 12, color: "#8A8570" }}>Contraseña</label>
        <input className="cc-input mb-3" type="password" required minLength={6} value={pass} onChange={(e) => setPass(e.target.value)} />
        {error && <div style={{ fontSize: 12.5, color: "var(--rust)", marginBottom: 10 }}>{error}</div>}
        <button className="cc-btn cc-btn-primary w-full justify-center" disabled={cargando} type="submit">
          {cargando ? <Loader2 size={18} className="animate-spin" /> : null}{modo === "ingresar" ? "Ingresar" : "Crear cuenta"}
        </button>
        <div className="flex justify-between mt-3" style={{ fontSize: 12 }}>
          <button type="button" onClick={() => setModo(modo === "ingresar" ? "crear" : "ingresar")} style={{ color: "var(--frost)" }}>
            {modo === "ingresar" ? "Crear una cuenta nueva" : "Ya tengo cuenta"}
          </button>
          {modo === "ingresar" && <button type="button" onClick={recuperar} style={{ color: "#8A8570" }}>Olvidé mi contraseña</button>}
        </div>
      </form>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  App principal                                                       */
/* ------------------------------------------------------------------ */
export default function App() {
  const [user, setUser] = useState(undefined);
  const [campaniasRaw, setCampaniasRaw] = useState([]);
  const [cultivosRaw, setCultivosRaw] = useState([]);
  const [camposRaw, setCamposRaw] = useState([]);
  const [lotesRaw, setLotesRaw] = useState([]);
  const [insumosComprasRaw, setInsumosComprasRaw] = useState([]);
  const [puntosStockRaw, setPuntosStockRaw] = useState([]);
  const [gastosRaw, setGastosRaw] = useState([]);
  const [ventasRaw, setVentasRaw] = useState([]);
  const [remitosRaw, setRemitosRaw] = useState([]);
  const [usuariosRaw, setUsuariosRaw] = useState([]);
  const [nav, setNav] = useState({ view: "campanias", campaniaId: null, cultivoId: null, campoId: null });
  const [menuMovilAbierto, setMenuMovilAbierto] = useState(false);

  useEffect(() => onAuthStateChanged(auth, (u) => setUser(u)), []);

  useEffect(() => {
    if (!user) return;
    const u1 = onSnapshot(collection(db, "campanias"), (s) => setCampaniasRaw(s.docs.map((d) => ({ id: d.id, ...d.data() }))));
    const u2 = onSnapshot(collection(db, "cultivos"), (s) => setCultivosRaw(s.docs.map((d) => ({ id: d.id, ...d.data() }))));
    const u3 = onSnapshot(collection(db, "lotes"), (s) => setLotesRaw(s.docs.map((d) => ({ id: d.id, ...d.data() }))));
    const u4 = onSnapshot(collection(db, "insumos_compras"), (s) => setInsumosComprasRaw(s.docs.map((d) => ({ id: d.id, ...d.data() }))));
    const u5 = onSnapshot(collection(db, "gastos"), (s) => setGastosRaw(s.docs.map((d) => ({ id: d.id, ...d.data() }))));
    const u6 = onSnapshot(collection(db, "ventas"), (s) => setVentasRaw(s.docs.map((d) => ({ id: d.id, ...d.data() }))));
    const u7 = onSnapshot(collection(db, "remitos"), (s) => setRemitosRaw(s.docs.map((d) => ({ id: d.id, ...d.data() }))));
    const u8 = onSnapshot(collection(db, "campos"), (s) => setCamposRaw(s.docs.map((d) => ({ id: d.id, ...d.data() }))));
    const u9 = onSnapshot(collection(db, "usuarios"), (s) => setUsuariosRaw(s.docs.map((d) => ({ id: d.id, ...d.data() }))));
    const u10 = onSnapshot(collection(db, "puntos_stock"), (s) => setPuntosStockRaw(s.docs.map((d) => ({ id: d.id, ...d.data() }))));
    return () => { u1(); u2(); u3(); u4(); u5(); u6(); u7(); u8(); u9(); u10(); };
  }, [user]);

  const miUsuario = usuariosRaw.find((u) => u.id === user?.email);
  const miRol = miUsuario?.rol || "admin";
  const puedeEditar = miRol !== "lectura";
  const esAdmin = miRol === "admin";
  // "default" = el espacio compartido de siempre (el de todo tu equipo). Si a un usuario
  // le asignás otro valor en "Usuarios", pasa a tener su propio espacio, separado del resto.
  const miWorkspaceId = miUsuario?.workspaceId || "default";
  const esDeMiWorkspace = (x) => (x.workspaceId || "default") === miWorkspaceId;

  const campanias = campaniasRaw.filter((x) => !x.eliminado && esDeMiWorkspace(x));
  const campaniaIdsActivas = new Set(campanias.map((c) => c.id));
  const cultivos = cultivosRaw.filter((x) => !x.eliminado && campaniaIdsActivas.has(x.campaniaId));
  const cultivoIdsActivos = new Set(cultivos.map((c) => c.id));
  const campos = camposRaw.filter((x) => !x.eliminado && esDeMiWorkspace(x));
  const campoIdsActivos = new Set(campos.map((c) => c.id));
  const lotes = lotesRaw.filter((x) => !x.eliminado && esDeMiWorkspace(x) && (!x.campoId || campoIdsActivos.has(x.campoId)));
  const insumosCompras = insumosComprasRaw.filter((x) => !x.eliminado && esDeMiWorkspace(x));
  const puntosStock = puntosStockRaw.filter((x) => !x.eliminado && esDeMiWorkspace(x));
  const gastosTodos = gastosRaw.filter((x) => !x.eliminado && cultivoIdsActivos.has(x.cultivoId));
  const ventasTodas = ventasRaw.filter((x) => !x.eliminado && cultivoIdsActivos.has(x.cultivoId));
  const remitosTodos = remitosRaw.filter((x) => !x.eliminado && cultivoIdsActivos.has(x.cultivoId));

  if (user === undefined) {
    return <div style={{ background: "var(--bg)", minHeight: "100vh" }} className="flex items-center justify-center"><style>{GLOBAL_STYLES}</style><Loader2 className="animate-spin" color="var(--soil)" size={30} /></div>;
  }
  if (!user) return <Login />;

  const softDeleteApi = (coleccion) => ({
    add: (data) => addDoc(collection(db, coleccion), { ...data, workspaceId: miWorkspaceId, creadoPor: user.email, creadoEn: new Date().toISOString() }),
    update: (id, data) => updateDoc(doc(db, coleccion, id), { ...data, modificadoPor: user.email, modificadoEn: new Date().toISOString() }),
    remove: (id) => updateDoc(doc(db, coleccion, id), { eliminado: true, eliminadoEn: new Date().toISOString(), eliminadoPor: user.email }),
    restaurar: (id) => updateDoc(doc(db, coleccion, id), { eliminado: false, eliminadoEn: null, eliminadoPor: null }),
    eliminarDefinitivo: (id) => deleteDoc(doc(db, coleccion, id)),
  });

  const campaniasApi = softDeleteApi("campanias");
  const cultivosApi = softDeleteApi("cultivos");
  const camposApi = softDeleteApi("campos");
  const lotesApi = softDeleteApi("lotes");
  const insumosApi = softDeleteApi("insumos_compras");
  const puntosStockApi = softDeleteApi("puntos_stock");
  const ventasApiGlobal = softDeleteApi("ventas");
  const remitosApiGlobal = softDeleteApi("remitos");
  const gastosApiGlobal = softDeleteApi("gastos");

  // Clave de agrupación: mismo insumo en distintos puntos de stock se llevan por separado
  const clavePunto = (nombre, puntoStockId) => `${nombre}||${puntoStockId || "sin_punto"}`;

  const stockInsumos = (() => {
    const mapa = {};
    insumosCompras.forEach((c) => {
      const k = clavePunto(c.nombre, c.puntoStockId);
      if (!mapa[k]) mapa[k] = { nombre: c.nombre, puntoStockId: c.puntoStockId || null, puntoStockNombre: c.puntoStockNombre || "Sin punto asignado", litrosComprados: 0, costoComprado: 0, litrosConsumidos: 0, unidad: c.unidad || "Litros" };
      mapa[k].litrosComprados += Number(c.litros || 0);
      mapa[k].costoComprado += Number(c.precio || 0);
      mapa[k].unidad = c.unidad || "Litros";
    });
    gastosTodos.forEach((g) => {
      if (!g.insumoNombre) return;
      const k = clavePunto(g.insumoNombre, g.puntoStockId);
      if (mapa[k]) mapa[k].litrosConsumidos += Number(g.litrosUsados || 0);
    });
    return Object.values(mapa).map((i) => ({
      ...i,
      disponible: i.litrosComprados - i.litrosConsumidos,
      costoPromedioPorLitro: i.litrosComprados ? i.costoComprado / i.litrosComprados : 0,
    }));
  })();

  const campaniaActual = campanias.find((c) => c.id === nav.campaniaId);
  const cultivoActual = cultivos.find((c) => c.id === nav.cultivoId);
  const campoActual = nav.campoId === "__sin_campo__" ? { id: "__sin_campo__", nombre: "Sin campo asignado" } : campos.find((c) => c.id === nav.campoId);

  return (
    <div style={{ background: "var(--bg)", fontFamily: "var(--font-body)", color: "var(--ink)", minHeight: "100vh" }}>
      <style>{GLOBAL_STYLES}</style>
      <header style={{ background: "var(--soil)" }} className="px-6 py-4">
        <div className="flex items-center justify-between max-w-6xl mx-auto">
          <button className="flex items-center gap-3" onClick={() => setNav({ view: "campanias", campaniaId: null, cultivoId: null })}>
            <span style={{ display: "inline-flex", alignItems: "center", justifyContent: "center", width: 40, height: 40, borderRadius: "50%", background: "rgba(198,138,46,0.18)" }}><Wheat color="var(--gold)" size={22} /></span>
            <div style={{ textAlign: "left" }}>
              <div className="cc-h" style={{ color: "#fff", fontSize: 20, fontWeight: 600, lineHeight: 1 }}>Campo & Costo</div>
              <div style={{ color: "#B8C2AC", fontSize: 12 }}>Campañas · Cultivos · Gastos e Ingresos</div>
            </div>
          </button>

          <div className="cc-nav-desktop items-center gap-3">
            <button onClick={() => setNav({ view: "resumen_general", campaniaId: null, cultivoId: null })} className="cc-btn" style={{ background: "transparent", border: "1px solid #4C5A40", color: "#D8DECB", padding: "6px 12px", fontSize: 12.5 }}>
              <span style={{ display: "inline-flex", alignItems: "center", justifyContent: "center", width: 22, height: 22, borderRadius: "50%", background: "var(--frost)" }}><LayoutDashboard size={13} color="#fff" /></span> Resumen general
            </button>
            <button onClick={() => setNav({ view: "insumos", campaniaId: null, cultivoId: null })} className="cc-btn" style={{ background: "transparent", border: "1px solid #4C5A40", color: "#D8DECB", padding: "6px 12px", fontSize: 12.5 }}>
              <span style={{ display: "inline-flex", alignItems: "center", justifyContent: "center", width: 22, height: 22, borderRadius: "50%", background: "var(--gold)" }}><Package size={13} color="#fff" /></span> Insumos
            </button>
            <button onClick={() => setNav({ view: "campos", campaniaId: null, cultivoId: null, campoId: null })} className="cc-btn" style={{ background: "transparent", border: "1px solid #4C5A40", color: "#D8DECB", padding: "6px 12px", fontSize: 12.5 }}>
              <span style={{ display: "inline-flex", alignItems: "center", justifyContent: "center", width: 22, height: 22, borderRadius: "50%", background: "var(--soil-light)" }}><MapPin size={13} color="#fff" /></span> Lotes
            </button>
            <button onClick={() => setNav({ view: "papelera", campaniaId: null, cultivoId: null })} className="cc-btn" style={{ background: "transparent", border: "1px solid #4C5A40", color: "#D8DECB", padding: "6px 12px", fontSize: 12.5 }}>
              <span style={{ display: "inline-flex", alignItems: "center", justifyContent: "center", width: 22, height: 22, borderRadius: "50%", background: "var(--rust)" }}><Trash size={13} color="#fff" /></span> Papelera
            </button>
            {esAdmin && (
              <button onClick={() => setNav({ view: "usuarios", campaniaId: null, cultivoId: null })} className="cc-btn" style={{ background: "transparent", border: "1px solid #4C5A40", color: "#D8DECB", padding: "6px 12px", fontSize: 12.5 }}>
                <span style={{ display: "inline-flex", alignItems: "center", justifyContent: "center", width: 22, height: 22, borderRadius: "50%", background: "var(--soil-light)" }}><Users size={13} color="#fff" /></span> Usuarios
              </button>
            )}
            <span style={{ color: "#D8DECB", fontSize: 12.5 }}>{user.email}</span>
            <button onClick={() => signOut(auth)} className="cc-btn" style={{ background: "transparent", border: "1px solid #4C5A40", color: "#D8DECB", padding: "6px 12px", fontSize: 12.5 }}><LogOut size={16} /> Salir</button>
          </div>

          <button className="cc-nav-toggle" onClick={() => setMenuMovilAbierto((v) => !v)} style={{ background: "transparent", border: "1px solid #4C5A40", color: "#fff", padding: "8px", borderRadius: 7 }}>
            {menuMovilAbierto ? <X size={20} /> : <Menu size={20} />}
          </button>
        </div>

        {menuMovilAbierto && (
          <div className="flex flex-col gap-2 mt-3" style={{ maxWidth: "72rem", margin: "12px auto 0" }}>
            {[
              { label: "Resumen general", icon: LayoutDashboard, color: "var(--frost)", onClick: () => setNav({ view: "resumen_general", campaniaId: null, cultivoId: null }) },
              { label: "Insumos", icon: Package, color: "var(--gold)", onClick: () => setNav({ view: "insumos", campaniaId: null, cultivoId: null }) },
              { label: "Lotes", icon: MapPin, color: "var(--soil-light)", onClick: () => setNav({ view: "campos", campaniaId: null, cultivoId: null, campoId: null }) },
              { label: "Papelera", icon: Trash, color: "var(--rust)", onClick: () => setNav({ view: "papelera", campaniaId: null, cultivoId: null }) },
              ...(esAdmin ? [{ label: "Usuarios", icon: Users, color: "var(--soil-light)", onClick: () => setNav({ view: "usuarios", campaniaId: null, cultivoId: null }) }] : []),
            ].map((it) => {
              const Icon = it.icon;
              return (
                <button key={it.label} onClick={() => { it.onClick(); setMenuMovilAbierto(false); }} className="cc-btn" style={{ background: "transparent", border: "1px solid #4C5A40", color: "#D8DECB", padding: "10px 12px", fontSize: 13.5, justifyContent: "flex-start" }}>
                  <span style={{ display: "inline-flex", alignItems: "center", justifyContent: "center", width: 24, height: 24, borderRadius: "50%", background: it.color }}><Icon size={14} color="#fff" /></span> {it.label}
                </button>
              );
            })}
            <div style={{ borderTop: "1px solid #4C5A40", margin: "4px 0" }} />
            <div style={{ color: "#D8DECB", fontSize: 12.5, padding: "0 4px" }}>{user.email}</div>
            <button onClick={() => signOut(auth)} className="cc-btn" style={{ background: "transparent", border: "1px solid #4C5A40", color: "#D8DECB", padding: "10px 12px", fontSize: 13.5, justifyContent: "flex-start" }}><LogOut size={16} /> Salir</button>
          </div>
        )}
      </header>

      <main className="max-w-6xl mx-auto px-6 py-6">
        <Breadcrumb nav={nav} setNav={setNav} campania={campaniaActual} cultivo={cultivoActual} campo={campoActual} />

        {!puedeEditar && nav.view !== "papelera" && (
          <div className="flex items-center gap-2 mb-4" style={{ background: "#FDF3E0", border: "1px solid var(--gold)", borderRadius: 8, padding: "8px 12px", fontSize: 12.5, color: "#7A5A1E" }}>
            <AlertCircle size={16} /> Estás con acceso de <b>solo lectura</b>: podés ver todo, pero no crear ni borrar datos.
          </div>
        )}

        {nav.view === "campanias" && <CampaniasView campanias={campanias} api={campaniasApi} cultivosApi={cultivosApi} cultivos={cultivos} onOpen={(id) => setNav({ view: "cultivos", campaniaId: id, cultivoId: null })} puedeEditar={puedeEditar} />}

        {nav.view === "cultivos" && campaniaActual && (
          <CultivosDeCampania campania={campaniaActual} cultivos={cultivos.filter((c) => c.campaniaId === campaniaActual.id)} api={cultivosApi}
            onOpen={(id) => setNav({ view: "cultivo", campaniaId: campaniaActual.id, cultivoId: id })} puedeEditar={puedeEditar} />
        )}

        {nav.view === "cultivo" && cultivoActual && <CultivoDetail cultivo={cultivoActual} lotes={lotes} lotesApi={lotesApi} cultivosApi={cultivosApi} user={user} stockInsumos={stockInsumos} insumosCompras={insumosCompras} gastosTodos={gastosTodos} campos={campos} puedeEditar={puedeEditar} miWorkspaceId={miWorkspaceId} />}

        {nav.view === "campos" && <CamposView campos={campos} api={camposApi} lotesApi={lotesApi} lotes={lotes} onOpen={(id) => setNav({ view: "lotes", campaniaId: null, cultivoId: null, campoId: id })} puedeEditar={puedeEditar} />}

        {nav.view === "lotes" && campoActual && (
          <LotesView campo={campoActual} lotes={campoActual.id === "__sin_campo__" ? lotes.filter((l) => !l.campoId) : lotes.filter((l) => l.campoId === campoActual.id)} api={lotesApi} sinCampo={campoActual.id === "__sin_campo__"} campos={campos} puedeEditar={puedeEditar} />
        )}

        {nav.view === "insumos" && <InsumosView compras={insumosCompras} api={insumosApi} stockInsumos={stockInsumos} user={user} puedeEditar={puedeEditar} puntosStock={puntosStock} puntosStockApi={puntosStockApi} />}

        {nav.view === "resumen_general" && (
          <ResumenGeneralView campanias={campanias} cultivos={cultivos} gastos={gastosTodos} ventas={ventasTodas} remitos={remitosTodos} lotes={lotes} insumosCompras={insumosCompras}
            onOpenCultivo={(cultivoId, campaniaId) => setNav({ view: "cultivo", campaniaId, cultivoId })} />
        )}

        {nav.view === "usuarios" && esAdmin && <UsuariosView usuarios={usuariosRaw} miEmail={user.email} />}

        {nav.view === "papelera" && (
          <PapeleraView
            grupos={[
              { titulo: "Campañas", items: campaniasRaw.filter((x) => x.eliminado && esDeMiWorkspace(x)), api: campaniasApi, campo: (x) => x.nombre || x.anio },
              { titulo: "Cultivos", items: cultivosRaw.filter((x) => x.eliminado && esDeMiWorkspace(x)), api: cultivosApi, campo: (x) => x.nombre },
              { titulo: "Campos", items: camposRaw.filter((x) => x.eliminado && esDeMiWorkspace(x)), api: camposApi, campo: (x) => x.nombre },
              { titulo: "Lotes", items: lotesRaw.filter((x) => x.eliminado && esDeMiWorkspace(x)), api: lotesApi, campo: (x) => x.nombre },
              { titulo: "Insumos (compras)", items: insumosComprasRaw.filter((x) => x.eliminado && esDeMiWorkspace(x)), api: insumosApi, campo: (x) => `${x.nombre} — ${fmt(x.litros, 1)} ${abrevUnidad(x.unidad)}` },
              { titulo: "Puntos de stock", items: puntosStockRaw.filter((x) => x.eliminado && esDeMiWorkspace(x)), api: puntosStockApi, campo: (x) => x.nombre },
              { titulo: "Gastos", items: gastosRaw.filter((x) => x.eliminado && esDeMiWorkspace(x)), api: gastosApiGlobal, campo: (x) => `${x.origen} — ${fmtUSD(x.monto)}` },
              { titulo: "Ventas", items: ventasRaw.filter((x) => x.eliminado && esDeMiWorkspace(x)), api: ventasApiGlobal, campo: (x) => `${x.origen} — ${fmt(x.toneladas, 2)} tn` },
              { titulo: "Remitos", items: remitosRaw.filter((x) => x.eliminado && esDeMiWorkspace(x)), api: remitosApiGlobal, campo: (x) => `Remito ${x.remito}` },
            ]}
            puedeEditar={puedeEditar}
          />
        )}
      </main>
    </div>
  );
}

function Breadcrumb({ nav, setNav, campania, cultivo, campo }) {
  if (["campanias", "campos", "insumos", "resumen_general", "papelera", "usuarios"].includes(nav.view)) return null;
  return (
    <div className="flex items-center gap-1 mb-4" style={{ fontSize: 13, color: "#8A8570" }}>
      {nav.view === "lotes" ? (
        <>
          <button onClick={() => setNav({ view: "campos", campaniaId: null, cultivoId: null, campoId: null })} style={{ color: "var(--frost)", fontWeight: 600 }}>Campos</button>
          {campo && (
            <>
              <ChevronRight size={16} />
              <span style={{ color: "var(--ink)", fontWeight: 600 }}>{campo.nombre}</span>
            </>
          )}
        </>
      ) : (
        <>
          <button onClick={() => setNav({ view: "campanias", campaniaId: null, cultivoId: null })} style={{ color: "var(--frost)", fontWeight: 600 }}>Campañas</button>
          {campania && (
            <>
              <ChevronRight size={16} />
              <button onClick={() => setNav({ view: "cultivos", campaniaId: campania.id, cultivoId: null })} style={{ color: nav.view === "cultivos" ? "var(--ink)" : "var(--frost)", fontWeight: 600 }}>
                {campania.nombre || campania.anio}
              </button>
            </>
          )}
          {cultivo && (
            <>
              <ChevronRight size={16} />
              <span style={{ color: "var(--ink)", fontWeight: 600 }}>{cultivo.nombre}</span>
            </>
          )}
        </>
      )}
    </div>
  );
}

function EmptyState({ icon: Icon, title, text }) {
  return (
    <div className="cc-card flex flex-col items-center text-center py-16 px-6">
      <Icon size={32} color="var(--gold)" style={{ marginBottom: 10 }} />
      <div className="cc-h" style={{ fontSize: 17, fontWeight: 600 }}>{title}</div>
      <div style={{ color: "#8A8570", fontSize: 13, maxWidth: 380, marginTop: 4 }}>{text}</div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Campañas                                                            */
/* ------------------------------------------------------------------ */
function CampaniasView({ campanias, api, cultivosApi, cultivos, onOpen, puedeEditar = true }) {
  const [anio, setAnio] = useState(new Date().getFullYear());
  const [nombre, setNombre] = useState("");

  const crear = () => {
    if (!anio || Number(anio) < 1900 || Number(anio) > 2100) { alert("Ingresá un año válido."); return; }
    api.add({ anio: Number(anio), nombre: nombre.trim() || `Campaña ${anio}` }); setNombre("");
  };
  const eliminar = (id) => {
    if (!confirm("Esta campaña y todos sus cultivos se moverán a la papelera (se puede restaurar después). ¿Continuar?")) return;
    api.remove(id);
    cultivos.filter((cu) => cu.campaniaId === id).forEach((cu) => cultivosApi.remove(cu.id));
  };

  return (
    <div className="space-y-5">
      {puedeEditar && (
        <div className="cc-card p-4">
          <div className="cc-h" style={{ fontSize: 16, fontWeight: 600, marginBottom: 10 }}>Nueva campaña</div>
          <div className="flex gap-3 flex-wrap items-end">
            <div style={{ width: 140 }}><label style={{ fontSize: 12, color: "#8A8570" }}>Año</label><input className="cc-input" type="number" value={anio} onChange={(e) => setAnio(e.target.value)} /></div>
            <div style={{ flex: 1, minWidth: 180 }}><label style={{ fontSize: 12, color: "#8A8570" }}>Nombre (opcional)</label><input className="cc-input" value={nombre} onChange={(e) => setNombre(e.target.value)} placeholder={`Campaña ${anio}`} /></div>
            <button className="cc-btn cc-btn-primary" onClick={crear}><Plus size={18} /> Crear</button>
          </div>
        </div>
      )}

      {campanias.length === 0 ? (
        <EmptyState icon={Sprout} title="No hay campañas todavía" text="Creá tu primera campaña (por ejemplo, por año) y luego agregale los cultivos que sembraste." />
      ) : (
        <div className="grid gap-4" style={{ gridTemplateColumns: "repeat(auto-fill, minmax(240px, 1fr))" }}>
          {[...campanias].sort((a, b) => b.anio - a.anio).map((c) => {
            const nCultivos = cultivos.filter((cu) => cu.campaniaId === c.id).length;
            return (
              <div key={c.id} className="cc-card p-4 cursor-pointer" onClick={() => onOpen(c.id)}>
                <div className="flex items-start justify-between">
                  <div>
                    <div className="cc-h" style={{ fontSize: 18, fontWeight: 600 }}>{c.nombre || c.anio}</div>
                    <div style={{ fontSize: 12, color: "#8A8570" }}>{nCultivos} cultivo{nCultivos !== 1 ? "s" : ""}</div>
                  </div>
                  {puedeEditar && <button onClick={(e) => { e.stopPropagation(); eliminar(c.id); }}><Trash2 size={17} color="var(--rust)" /></button>}
                </div>
                <div className="flex items-center gap-1 mt-3" style={{ color: "var(--frost)", fontSize: 12.5, fontWeight: 600 }}>Ver cultivos <ChevronRight size={16} /></div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Cultivos de una campaña                                             */
/* ------------------------------------------------------------------ */
function CultivosDeCampania({ campania, cultivos, api, onOpen, puedeEditar = true }) {
  const [categoria, setCategoria] = useState("verano");
  const [nombre, setNombre] = useState(CULTIVOS_VERANO[0]);
  const [nombreCustom, setNombreCustom] = useState("");
  const opciones = categoria === "verano" ? CULTIVOS_VERANO : CULTIVOS_INVIERNO;

  const crear = () => { api.add({ campaniaId: campania.id, categoria, tipo: nombre, nombre: nombreCustom.trim() || nombre }); setNombreCustom(""); };
  const eliminar = (id) => { if (confirm("Este cultivo se moverá a la papelera (se puede restaurar después). ¿Continuar?")) api.remove(id); };

  return (
    <div className="space-y-5">
      {puedeEditar && (
        <div className="cc-card p-4">
          <div className="cc-h" style={{ fontSize: 16, fontWeight: 600, marginBottom: 10 }}>Agregar cultivo a {campania.nombre || campania.anio}</div>
          <div className="flex gap-3 flex-wrap items-end">
            <div style={{ width: 160 }}>
              <label style={{ fontSize: 12, color: "#8A8570" }}>Categoría</label>
              <select className="cc-input" value={categoria} onChange={(e) => { setCategoria(e.target.value); setNombre(e.target.value === "verano" ? CULTIVOS_VERANO[0] : CULTIVOS_INVIERNO[0]); }}>
                <option value="verano">Verano</option><option value="invierno">Invierno</option>
              </select>
            </div>
            <div style={{ width: 180 }}>
              <label style={{ fontSize: 12, color: "#8A8570" }}>Cultivo</label>
              <select className="cc-input" value={nombre} onChange={(e) => setNombre(e.target.value)}>{opciones.map((c) => <option key={c} value={c}>{c}</option>)}</select>
            </div>
            <button className="cc-btn cc-btn-primary" onClick={crear}><Plus size={18} /> Agregar</button>
          </div>
          <div style={{ marginTop: 12, maxWidth: 320 }}>
            <label style={{ fontSize: 12, color: "#8A8570" }}>Nombre para este cultivo (opcional)</label>
            <input className="cc-input" value={nombreCustom} onChange={(e) => setNombreCustom(e.target.value)} placeholder={`Ej: ${nombre} 1ra, ${nombre} Norte...`} />
            <div style={{ fontSize: 11.5, color: "#8A8570", marginTop: 4 }}>Si lo dejás vacío, se usa "{nombre}". Útil si vas a tener más de un lote de {nombre} en la misma campaña.</div>
          </div>
        </div>
      )}

      {cultivos.length === 0 ? (
        <EmptyState icon={Sprout} title="Todavía no hay cultivos en esta campaña" text="Agregá los cultivos sembrados (ej: Soja, Trigo) para empezar a cargar sus gastos e ingresos." />
      ) : (
        <div className="grid gap-4" style={{ gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))" }}>
          {cultivos.map((c) => (
            <div key={c.id} className="cc-card p-4 cursor-pointer" onClick={() => onOpen(c.id)}>
              <div className="flex items-start justify-between">
                <div>
                  <div className="cc-h" style={{ fontSize: 17, fontWeight: 600 }}>{c.nombre}</div>
                  <span className="cc-chip" style={{ background: CAT_COLOR[c.categoria] + "22", color: CAT_COLOR[c.categoria] }}>{c.categoria === "verano" ? "Verano" : "Invierno"}</span>
                </div>
                <div className="flex items-center gap-2">
                  {c.categoria === "verano" ? <Sun color="var(--gold)" size={21} /> : <Snowflake color="var(--frost)" size={21} />}
                  {puedeEditar && <button onClick={(e) => { e.stopPropagation(); eliminar(c.id); }}><Trash2 size={16} color="var(--rust)" /></button>}
                </div>
              </div>
              <div className="flex items-center gap-1 mt-3" style={{ color: "var(--frost)", fontSize: 12.5, fontWeight: 600 }}>Gastos e ingresos <ChevronRight size={16} /></div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Detalle de un cultivo: Resumen / Gastos / Ventas / Remitos          */
/* ------------------------------------------------------------------ */
function CultivoDetail({ cultivo, lotes, lotesApi, cultivosApi, user, stockInsumos, insumosCompras, gastosTodos, campos, puedeEditar = true, miWorkspaceId }) {
  const [tab, setTab] = useState("resumen");
  const [gastos, setGastos] = useState([]);
  const [ventas, setVentas] = useState([]);
  const [remitos, setRemitos] = useState([]);

  useEffect(() => {
    const q1 = query(collection(db, "gastos"), where("cultivoId", "==", cultivo.id));
    const q2 = query(collection(db, "ventas"), where("cultivoId", "==", cultivo.id));
    const q3 = query(collection(db, "remitos"), where("cultivoId", "==", cultivo.id));
    const u1 = onSnapshot(q1, (s) => setGastos(s.docs.map((d) => ({ id: d.id, ...d.data() })).filter((x) => !x.eliminado)));
    const u2 = onSnapshot(q2, (s) => setVentas(s.docs.map((d) => ({ id: d.id, ...d.data() })).filter((x) => !x.eliminado)));
    const u3 = onSnapshot(q3, (s) => setRemitos(s.docs.map((d) => ({ id: d.id, ...d.data() })).filter((x) => !x.eliminado)));
    return () => { u1(); u2(); u3(); };
  }, [cultivo.id]);

  const gastosApi = {
    add: (d) => addDoc(collection(db, "gastos"), { ...d, workspaceId: miWorkspaceId, creadoPor: user.email, creadoEn: new Date().toISOString() }),
    update: (id, d) => updateDoc(doc(db, "gastos", id), { ...d, modificadoPor: user.email, modificadoEn: new Date().toISOString() }),
    remove: (id) => updateDoc(doc(db, "gastos", id), { eliminado: true, eliminadoEn: new Date().toISOString(), eliminadoPor: user.email }),
  };
  const ventasApi = {
    add: (d) => addDoc(collection(db, "ventas"), { ...d, workspaceId: miWorkspaceId, creadoPor: user.email, creadoEn: new Date().toISOString() }),
    remove: (id) => updateDoc(doc(db, "ventas", id), { eliminado: true, eliminadoEn: new Date().toISOString(), eliminadoPor: user.email }),
  };
  const remitosApi = {
    add: (d) => addDoc(collection(db, "remitos"), { ...d, workspaceId: miWorkspaceId, creadoPor: user.email, creadoEn: new Date().toISOString() }),
    remove: (id) => updateDoc(doc(db, "remitos", id), { eliminado: true, eliminadoEn: new Date().toISOString(), eliminadoPor: user.email }),
  };

  const totalGastos = gastos.reduce((s, g) => s + Number(g.monto || 0), 0);
  const totalIngresos = ventas.reduce((s, v) => s + Number(v.toneladas || 0) * Number(v.dolaresPorTonelada || 0), 0);
  const totalToneladasVentas = ventas.reduce((s, v) => s + Number(v.toneladas || 0), 0);
  const totalKgSL = remitos.reduce((s, r) => s + Number(r.kgSL || 0), 0);
  const totalTonRemitos = totalKgSL / 1000;

  const loteIds = cultivo.loteIds || [];
  const superficie = lotes.filter((l) => loteIds.includes(l.id)).reduce((s, l) => s + Number(l.hectareas || 0), 0);
  const costoPorHa = superficie ? totalGastos / superficie : null;
  const ingresoPorHa = superficie ? totalIngresos / superficie : null;
  const balancePorHa = costoPorHa !== null && ingresoPorHa !== null ? ingresoPorHa - costoPorHa : null;
  const precioPromedioVenta = totalToneladasVentas ? totalIngresos / totalToneladasVentas : null;
  const rindeEquilibrio = costoPorHa !== null && precioPromedioVenta ? costoPorHa / precioPromedioVenta : null;
  const rendimientoReal = superficie && totalTonRemitos ? totalTonRemitos / superficie : null;

  const TABS = [
    { id: "resumen", label: "Resumen", icon: TrendingUp, color: "var(--frost)" },
    { id: "lotes", label: "Lotes", icon: MapPin, color: "var(--gold)" },
    { id: "gastos", label: "Gastos", icon: Receipt, color: "var(--rust)" },
    { id: "ventas", label: "Ingresos · Ventas", icon: DollarSign, color: "var(--soil-light)" },
    { id: "remitos", label: "Ingresos · Remitos", icon: Truck, color: "var(--soil-light)" },
  ];

  return (
    <div>
      <div className="flex gap-1 mb-4 flex-wrap">
        {TABS.map((t) => {
          const Icon = t.icon; const active = tab === t.id;
          return (
            <button key={t.id} onClick={() => setTab(t.id)} className="cc-sub" style={{ background: active ? "var(--panel)" : "transparent", color: active ? "var(--soil)" : "#8A8570", border: active ? "1px solid var(--line)" : "1px solid transparent", borderBottom: active ? "1px solid var(--panel)" : "none" }}>
              <span style={{ display: "inline-flex", alignItems: "center", justifyContent: "center", width: 24, height: 24, borderRadius: "50%", background: active ? t.color : t.color + "22" }}>
                <Icon size={14} color={active ? "#fff" : t.color} />
              </span>
              {t.label}
            </button>
          );
        })}
      </div>

      {tab === "resumen" && (
        <div className="space-y-4">
          <div className="grid gap-4" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(200px,1fr))" }}>
            <StatCard label="Superficie" value={`${fmt(superficie, 1)} ha`} icon={MapPin} color="var(--frost)" />
            <StatCard label="Total gastos" value={fmtUSD(totalGastos)} icon={Receipt} color="var(--rust)" />
            <StatCard label="Total ingresos (ventas)" value={fmtUSD(totalIngresos)} icon={DollarSign} color="var(--soil-light)" />
            <StatCard label="Margen" value={fmtUSD(totalIngresos - totalGastos)} icon={totalIngresos - totalGastos >= 0 ? TrendingUp : TrendingDown} color={totalIngresos - totalGastos >= 0 ? "var(--soil-light)" : "var(--rust)"} />
          </div>

          <PresupuestoCard cultivo={cultivo} cultivosApi={cultivosApi} totalGastos={totalGastos} />

          {!superficie ? (
            <div className="flex items-start gap-2" style={{ background: "#FDF3E0", border: "1px solid var(--gold)", borderRadius: 8, padding: "10px 12px" }}>
              <AlertCircle size={16} color="#7A5A1E" style={{ marginTop: 1, flexShrink: 0 }} />
              <div style={{ fontSize: 13, color: "#7A5A1E" }}>Todavía no asociaste lotes a este cultivo — andá a la pestaña "Lotes" para cargar la superficie y ver los indicadores por hectárea.</div>
            </div>
          ) : (
            <div className="cc-card p-4">
              <div className="cc-h" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Indicadores por hectárea</div>
              <div className="grid gap-4" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(180px,1fr))" }}>
                <MiniStat label="Gasto / ha" value={fmtUSD(costoPorHa)} />
                <MiniStat label="Ingreso / ha" value={fmtUSD(ingresoPorHa)} />
                <MiniStat label="Balance / ha" value={fmtUSD(balancePorHa)} />
                <MiniStat label="Precio prom. venta" value={precioPromedioVenta ? fmtUSD(precioPromedioVenta) + "/tn" : "-"} />
                <MiniStat label="Rendimiento real" value={rendimientoReal !== null ? `${fmt(rendimientoReal, 2)} tn/ha` : "-"} sub={rendimientoReal !== null ? `${fmt(rendimientoReal * 1000, 0)} kg/ha` : "Cargá remitos para calcularlo"} />
                <MiniStat label="Rinde de equilibrio" value={rindeEquilibrio !== null ? `${fmt(rindeEquilibrio, 2)} tn/ha` : "-"} />
              </div>
              {rindeEquilibrio === null && <div style={{ fontSize: 11.5, color: "#8A8570", marginTop: 6 }}>Cargá al menos una venta para calcular el rinde de equilibrio.</div>}
              {rendimientoReal !== null && rindeEquilibrio !== null && (
                <div style={{ fontSize: 11.5, color: rendimientoReal >= rindeEquilibrio ? "var(--soil-light)" : "var(--rust)", marginTop: 6, fontWeight: 600 }}>
                  {rendimientoReal >= rindeEquilibrio
                    ? `Vas ${fmt(rendimientoReal - rindeEquilibrio, 2)} tn/ha por encima del rinde de equilibrio.`
                    : `Vas ${fmt(rindeEquilibrio - rendimientoReal, 2)} tn/ha por debajo del rinde de equilibrio.`}
                </div>
              )}
            </div>
          )}

          <div className="cc-card p-4">
            <div className="cc-h" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Gastos por tipo</div>
            {gastos.length === 0 ? (
              <div style={{ fontSize: 13, color: "#8A8570" }}>Todavía no hay gastos cargados en este cultivo.</div>
            ) : (
              <div className="space-y-3">
                {Object.entries(
                  gastos.reduce((acc, g) => {
                    const cat = g.categoriaGasto || (g.insumoNombre ? "Insumo" : "Otro");
                    acc[cat] = (acc[cat] || 0) + Number(g.monto || 0);
                    return acc;
                  }, {})
                )
                  .sort((a, b) => b[1] - a[1])
                  .map(([cat, monto]) => {
                    const pct = totalGastos ? (monto / totalGastos) * 100 : 0;
                    const color = GASTO_CAT_COLOR[cat] || GASTO_CAT_COLOR.Otro;
                    return (
                      <div key={cat}>
                        <div className="flex items-center justify-between" style={{ fontSize: 13 }}>
                          <span style={{ fontWeight: 600, color: "#5A5647" }}>{cat}</span>
                          <span className="cc-mono">{fmtUSD(monto)} <span style={{ color: "#8A8570" }}>({fmt(pct, 1)}%)</span></span>
                        </div>
                        <div style={{ background: "#EEEADA", borderRadius: 99, height: 8, marginTop: 4, overflow: "hidden" }}>
                          <div style={{ width: `${Math.min(pct, 100)}%`, background: color, height: "100%", borderRadius: 99 }} />
                        </div>
                      </div>
                    );
                  })}
              </div>
            )}
          </div>

          <div className="cc-card p-4">
            <div className="cc-h" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Aporte por socio</div>
            {gastos.length === 0 ? (
              <div style={{ fontSize: 13, color: "#8A8570" }}>Todavía no hay gastos cargados en este cultivo.</div>
            ) : (
              <div className="space-y-3">
                {Object.entries(
                  gastos.reduce((acc, g) => {
                    const socio = g.socio && g.socio.trim() ? g.socio.trim() : "Sin asignar";
                    acc[socio] = (acc[socio] || 0) + Number(g.monto || 0);
                    return acc;
                  }, {})
                )
                  .sort((a, b) => b[1] - a[1])
                  .map(([socio, monto]) => {
                    const pct = totalGastos ? (monto / totalGastos) * 100 : 0;
                    const color = colorPorTexto(socio);
                    return (
                      <div key={socio}>
                        <div className="flex items-center justify-between" style={{ fontSize: 13 }}>
                          <span style={{ fontWeight: 600, color: socio === "Sin asignar" ? "#8A8570" : "#5A5647" }}>{socio}</span>
                          <span className="cc-mono">{fmtUSD(monto)} <span style={{ color: "#8A8570" }}>({fmt(pct, 1)}%)</span></span>
                        </div>
                        <div style={{ background: "#EEEADA", borderRadius: 99, height: 8, marginTop: 4, overflow: "hidden" }}>
                          <div style={{ width: `${Math.min(pct, 100)}%`, background: color, height: "100%", borderRadius: 99 }} />
                        </div>
                      </div>
                    );
                  })}
              </div>
            )}
          </div>

          <div className="cc-card p-4">
            <div className="cc-h" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Conciliación de toneladas</div>
            <div className="grid gap-4" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(180px,1fr))" }}>
              <MiniStat label="Toneladas (ventas)" value={`${fmt(totalToneladasVentas, 2)} tn`} />
              <MiniStat label="Kg SL acumulados (remitos)" value={`${fmt(totalKgSL, 0)} kg`} />
              <MiniStat label="Equivalente en toneladas" value={`${fmt(totalTonRemitos, 2)} tn`} />
            </div>
            <div className="flex items-center gap-2 mt-3" style={{ fontSize: 13 }}>
              {Math.abs(totalTonRemitos - totalToneladasVentas) < 0.05 && totalToneladasVentas > 0 ? (
                <><CheckCircle2 size={18} color="var(--soil-light)" /><span style={{ color: "var(--soil-light)" }}>Los remitos coinciden con las toneladas vendidas.</span></>
              ) : (
                <><AlertCircle size={18} color="var(--rust)" /><span style={{ color: "var(--rust)" }}>
                  Diferencia de {fmt(Math.abs(totalTonRemitos - totalToneladasVentas), 2)} tn entre remitos y ventas.
                </span></>
              )}
            </div>
          </div>
        </div>
      )}

      {tab === "lotes" && <CultivoLotesTab cultivo={cultivo} lotes={lotes} lotesApi={lotesApi} cultivosApi={cultivosApi} superficie={superficie} campos={campos} puedeEditar={puedeEditar} />}

      {tab === "gastos" && <GastosTab cultivo={cultivo} gastos={gastos} api={gastosApi} user={user} stockInsumos={stockInsumos} insumosCompras={insumosCompras} gastosTodos={gastosTodos} superficie={superficie} puedeEditar={puedeEditar} />}
      {tab === "ventas" && <VentasTab cultivo={cultivo} ventas={ventas} api={ventasApi} puedeEditar={puedeEditar} />}
      {tab === "remitos" && <RemitosTab cultivo={cultivo} remitos={remitos} api={remitosApi} lotes={lotes} campos={campos} totalToneladasVentas={totalToneladasVentas} puedeEditar={puedeEditar} />}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Lotes asociados a un cultivo                                        */
/* ------------------------------------------------------------------ */
function CultivoLotesTab({ cultivo, lotes, lotesApi, cultivosApi, superficie, campos = [], puedeEditar = true }) {
  const [nombre, setNombre] = useState("");
  const [hectareas, setHectareas] = useState("");
  const [filtroCampo, setFiltroCampo] = useState("todos");
  const loteIds = cultivo.loteIds || [];

  const toggle = async (loteId) => {
    const nuevos = loteIds.includes(loteId) ? loteIds.filter((id) => id !== loteId) : [...loteIds, loteId];
    await cultivosApi.update(cultivo.id, { loteIds: nuevos });
  };

  const agregarNuevoLote = async () => {
    if (!nombre.trim() || !hectareas) { alert("Completá nombre y hectáreas del lote."); return; }
    if (Number(hectareas) <= 0) { alert("Las hectáreas tienen que ser un número mayor a 0."); return; }
    const ref = await lotesApi.add({ nombre: nombre.trim(), hectareas: Number(hectareas) });
    await cultivosApi.update(cultivo.id, { loteIds: [...loteIds, ref.id] });
    setNombre(""); setHectareas("");
  };

  const lotesFiltrados = filtroCampo === "todos" ? lotes : filtroCampo === "__sin_campo__" ? lotes.filter((l) => !l.campoId) : lotes.filter((l) => l.campoId === filtroCampo);
  const gruposOrden = [...campos.map((c) => c.id), "__sin_campo__"];
  const nombreCampo = (id) => (id === "__sin_campo__" ? "Sin campo asignado" : (campos.find((c) => c.id === id)?.nombre || "Campo eliminado"));
  const grupos = {};
  lotesFiltrados.forEach((l) => {
    const key = l.campoId && campos.some((c) => c.id === l.campoId) ? l.campoId : "__sin_campo__";
    if (!grupos[key]) grupos[key] = [];
    grupos[key].push(l);
  });

  return (
    <div className="space-y-5">
      {puedeEditar && (
        <div className="cc-card p-4">
          <div className="cc-h" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Agregar un lote nuevo y sumarlo a este cultivo</div>
          <div className="flex gap-3 flex-wrap items-end">
            <div style={{ flex: 1, minWidth: 180 }}>
              <label style={{ fontSize: 12, color: "#8A8570" }}>Nombre del lote</label>
              <input className="cc-input" value={nombre} onChange={(e) => setNombre(e.target.value)} placeholder="Ej: Lote 4 - La Loma" />
            </div>
            <div style={{ width: 150 }}>
              <label style={{ fontSize: 12, color: "#8A8570" }}>Hectáreas</label>
              <input className="cc-input" type="number" value={hectareas} onChange={(e) => setHectareas(e.target.value)} placeholder="93" />
            </div>
            <button className="cc-btn cc-btn-primary" onClick={agregarNuevoLote}><Plus size={17} /> Agregar y sumar</button>
          </div>
          <div style={{ fontSize: 11.5, color: "#8A8570", marginTop: 6 }}>Se agrega sin campo asignado; podés ordenarlo después desde "Lotes" en el menú superior.</div>
        </div>
      )}

      <div className="cc-card p-4">
        <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
          <div className="cc-h" style={{ fontSize: 15, fontWeight: 600 }}>Lotes de este cultivo</div>
          <div className="cc-mono" style={{ fontSize: 15, fontWeight: 700, color: "var(--frost)" }}>{fmt(superficie, 1)} ha totales</div>
        </div>

        {campos.length > 0 && (
          <div className="mb-3" style={{ maxWidth: 260 }}>
            <label style={{ fontSize: 12, color: "#8A8570" }}>Filtrar por campo</label>
            <select className="cc-input" value={filtroCampo} onChange={(e) => setFiltroCampo(e.target.value)}>
              <option value="todos">Todos los campos</option>
              {campos.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}
              {lotes.some((l) => !l.campoId) && <option value="__sin_campo__">Sin campo asignado</option>}
            </select>
          </div>
        )}

        {lotes.length === 0 ? (
          <div style={{ fontSize: 13, color: "#8A8570" }}>Todavía no hay lotes cargados. Agregá uno arriba.</div>
        ) : lotesFiltrados.length === 0 ? (
          <div style={{ fontSize: 13, color: "#8A8570" }}>No hay lotes en ese campo.</div>
        ) : (
          <div className="space-y-4">
            {gruposOrden.filter((k) => grupos[k]?.length).map((key) => (
              <div key={key}>
                {campos.length > 0 && <div style={{ fontSize: 12, fontWeight: 700, color: "#8A8570", marginBottom: 6, textTransform: "uppercase", letterSpacing: ".03em" }}>{nombreCampo(key)}</div>}
                <div className="flex flex-wrap gap-2">
                  {grupos[key].map((l) => {
                    const sel = loteIds.includes(l.id);
                    return (
                      <button key={l.id} onClick={() => puedeEditar && toggle(l.id)} disabled={!puedeEditar} className="cc-btn"
                        style={{ padding: "8px 12px", fontSize: 13, background: sel ? "var(--soil)" : "#fff", color: sel ? "#fff" : "var(--ink)", border: "1px solid var(--line)" }}>
                        {sel ? <CheckCircle2 size={16} /> : <MapPin size={16} />} {l.nombre} ({fmt(l.hectareas, 1)} ha)
                      </button>
                    );
                  })}
                </div>
              </div>
            ))}
          </div>
        )}
        <div style={{ fontSize: 12, color: "#8A8570", marginTop: 10 }}>Tocá un lote para sumarlo o sacarlo de este cultivo. Un mismo lote puede pertenecer a distintos cultivos si corresponde (por ejemplo, cambió de campaña).</div>
      </div>
    </div>
  );
}


/* ------------------------------------------------------------------ */
/*  Presupuesto vs. real                                                */
/* ------------------------------------------------------------------ */
function PresupuestoCard({ cultivo, cultivosApi, totalGastos }) {
  const [editando, setEditando] = useState(false);
  const [valor, setValor] = useState(cultivo.presupuesto ?? "");

  const guardar = async () => {
    if (valor && Number(valor) <= 0) { alert("El presupuesto tiene que ser un número mayor a 0."); return; }
    await cultivosApi.update(cultivo.id, { presupuesto: valor ? Number(valor) : null });
    setEditando(false);
  };

  if (!cultivo.presupuesto && !editando) {
    return (
      <div className="cc-card p-4 flex items-center justify-between flex-wrap gap-2">
        <div style={{ fontSize: 13, color: "#8A8570" }}>Todavía no cargaste un presupuesto estimado para este cultivo.</div>
        <button className="cc-btn cc-btn-ghost" onClick={() => setEditando(true)}><Plus size={17} /> Cargar presupuesto</button>
      </div>
    );
  }

  if (editando) {
    return (
      <div className="cc-card p-4">
        <div className="cc-h" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Presupuesto estimado (U$S)</div>
        <div className="flex gap-2 items-end flex-wrap">
          <div style={{ width: 180 }}><input className="cc-input" type="number" value={valor} onChange={(e) => setValor(e.target.value)} placeholder="Ej: 45000" /></div>
          <button className="cc-btn cc-btn-primary" onClick={guardar}><CheckCircle2 size={17} /> Guardar</button>
          <button className="cc-btn cc-btn-ghost" onClick={() => { setEditando(false); setValor(cultivo.presupuesto ?? ""); }}><X size={17} /> Cancelar</button>
        </div>
      </div>
    );
  }

  const pct = cultivo.presupuesto ? (totalGastos / cultivo.presupuesto) * 100 : 0;
  const color = pct >= 100 ? "var(--rust)" : pct >= 80 ? "var(--gold)" : "var(--soil-light)";

  return (
    <div className="cc-card p-4">
      <div className="flex items-center justify-between mb-2 flex-wrap gap-2">
        <div className="cc-h" style={{ fontSize: 15, fontWeight: 600 }}>Presupuesto vs. real</div>
        <button style={{ fontSize: 12, color: "var(--frost)" }} onClick={() => setEditando(true)}>Editar</button>
      </div>
      <div className="flex items-center justify-between" style={{ fontSize: 13 }}>
        <span className="cc-mono">{fmtUSD(totalGastos)} gastado</span>
        <span className="cc-mono" style={{ color: "#8A8570" }}>de {fmtUSD(cultivo.presupuesto)} ({fmt(pct, 0)}%)</span>
      </div>
      <div style={{ background: "#EEEADA", borderRadius: 99, height: 10, marginTop: 6, overflow: "hidden" }}>
        <div style={{ width: `${Math.min(pct, 100)}%`, background: color, height: "100%", borderRadius: 99 }} />
      </div>
      {pct >= 100 && <div style={{ fontSize: 12, color: "var(--rust)", marginTop: 6 }}>Ya superaste el presupuesto estimado.</div>}
      {pct >= 80 && pct < 100 && <div style={{ fontSize: 12, color: "#7A5A1E", marginTop: 6 }}>Te estás acercando al presupuesto estimado.</div>}
    </div>
  );
}

function StatCard({ label, value, icon: Icon, color }) {
  return (
    <div className="cc-card p-4 flex items-center gap-3">
      <div style={{ background: color + "1A", borderRadius: "50%", width: 38, height: 38, display: "flex", alignItems: "center", justifyContent: "center" }}><Icon size={19} color={color} /></div>
      <div><div style={{ fontSize: 11, color: "#8A8570", textTransform: "uppercase", letterSpacing: ".03em" }}>{label}</div><div className="cc-mono" style={{ fontSize: 17, fontWeight: 600 }}>{value}</div></div>
    </div>
  );
}
function MiniStat({ label, value, sub }) {
  return <div><div style={{ fontSize: 10.5, color: "#8A8570", textTransform: "uppercase", letterSpacing: ".03em" }}>{label}</div><div className="cc-mono" style={{ fontSize: 15, fontWeight: 500 }}>{value}</div>{sub && <div style={{ fontSize: 10.5, color: "#8A8570" }}>{sub}</div>}</div>;
}

/* ------------------------------------------------------------------ */
/*  Gastos                                                               */
/* ------------------------------------------------------------------ */
const emptyGasto = (email) => ({ origen: "", monto: "", detalle: "", fecha: "", usuario: email, categoriaGasto: "Servicio", socio: "" });

function GastosTab({ cultivo, gastos, api, user, stockInsumos, insumosCompras, gastosTodos, superficie, puedeEditar = true }) {
  const [form, setForm] = useState(emptyGasto(user.email));
  const [editId, setEditId] = useState(null);
  const [tipo, setTipo] = useState("general"); // "general" | "insumo"
  const [insumoSel, setInsumoSel] = useState("");
  const [litrosUsados, setLitrosUsados] = useState("");
  const [archivos, setArchivos] = useState([]);
  const [previews, setPreviews] = useState([]);
  const [subiendo, setSubiendo] = useState(false);
  const [extrayendo, setExtrayendo] = useState(false);
  const [grabando, setGrabando] = useState(false);
  const [transcripcion, setTranscripcion] = useState("");
  const [interpretando, setInterpretando] = useState(false);
  const [error, setError] = useState("");
  const recRef = useRef(null);
  const set = (k, v) => setForm({ ...form, [k]: v });
  const [busqueda, setBusqueda] = useState("");
  const [filtroCategoria, setFiltroCategoria] = useState("todas");
  const [mensaje, setMensaje] = useState("");
  const avisar = (texto) => { setMensaje(texto); setTimeout(() => setMensaje(""), 2500); };
  const [moneda, setMoneda] = useState("USD");
  const [montoPesos, setMontoPesos] = useState("");
  const [cotizacion, setCotizacion] = useState("");
  const [cotizandoAuto, setCotizandoAuto] = useState(false);
  const [modoMonto, setModoMonto] = useState("total"); // "total" | "porHa"
  const [valorPorHa, setValorPorHa] = useState("");

  useEffect(() => {
    let cancelado = false;
    setCotizandoAuto(true);
    fetch("https://open.er-api.com/v6/latest/USD")
      .then((r) => r.json())
      .then((d) => { if (!cancelado && d?.rates?.UYU) setCotizacion(String(Math.round(d.rates.UYU * 100) / 100)); })
      .catch(() => {})
      .finally(() => { if (!cancelado) setCotizandoAuto(false); });
    return () => { cancelado = true; };
  }, []);

  useEffect(() => {
    if (moneda === "UYU" && montoPesos && Number(cotizacion) > 0) {
      setForm((f) => ({ ...f, monto: (Number(montoPesos) / Number(cotizacion)).toFixed(2) }));
    }
  }, [moneda, montoPesos, cotizacion]);

  // Cuando el modo es "Valor por hectárea", el monto total se recalcula solo (valor/ha × superficie del cultivo)
  useEffect(() => {
    if (modoMonto === "porHa" && valorPorHa && superficie) {
      setForm((f) => ({ ...f, monto: (Number(valorPorHa) * superficie).toFixed(2) }));
    }
  }, [modoMonto, valorPorHa, superficie]);

  const origenesSugeridos = Array.from(new Set(gastos.map((g) => g.origen).filter(Boolean)));
  const sociosSugeridos = Array.from(new Set(gastos.map((g) => g.socio).filter(Boolean)));
  const insumoElegido = stockInsumos.find((i) => `${i.nombre}||${i.puntoStockId || "sin_punto"}` === insumoSel);
  const insumoNombreSel = insumoElegido?.nombre || "";
  const puntoStockIdSel = insumoElegido?.puntoStockId || null;
  const comprasDeEsteInsumo = insumosCompras.filter((c) => c.nombre === insumoNombreSel && (c.puntoStockId || null) === puntoStockIdSel);
  const litrosYaConsumidosPorOtros = gastosTodos
    .filter((g) => g.insumoNombre === insumoNombreSel && (g.puntoStockId || null) === puntoStockIdSel && g.id !== editId)
    .reduce((s, g) => s + Number(g.litrosUsados || 0), 0);
  const fifo = insumoElegido ? costoFIFO(comprasDeEsteInsumo, litrosYaConsumidosPorOtros, Number(litrosUsados || 0)) : { costoTotal: 0, costoPromedioEfectivo: 0 };
  const montoCalculado = fifo.costoTotal;

  const editar = (g) => {
    setEditId(g.id);
    if (g.insumoNombre) {
      setTipo("insumo"); setInsumoSel(`${g.insumoNombre}||${g.puntoStockId || "sin_punto"}`); setLitrosUsados(String(g.litrosUsados ?? ""));
      setForm({ origen: g.origen || "", monto: g.monto ?? "", detalle: g.detalle || "", fecha: g.fecha || "", usuario: g.usuario || user.email, categoriaGasto: "Insumo", socio: g.socio || "" });
    } else {
      setTipo("general"); setInsumoSel(""); setLitrosUsados("");
      setForm({ origen: g.origen || "", monto: g.monto ?? "", detalle: g.detalle || "", fecha: g.fecha || "", usuario: g.usuario || user.email, categoriaGasto: g.categoriaGasto || "Servicio", socio: g.socio || "" });
    }
    if (g.modoMonto === "porHa" && g.valorPorHectarea) { setModoMonto("porHa"); setValorPorHa(String(g.valorPorHectarea)); }
    else { setModoMonto("total"); setValorPorHa(""); }
    setArchivos([]); setPreviews([]); setTranscripcion(""); setError("");
    window.scrollTo({ top: 0, behavior: "smooth" });
  };
  const duplicar = (g) => {
    setEditId(null);
    if (g.insumoNombre) {
      setTipo("insumo"); setInsumoSel(`${g.insumoNombre}||${g.puntoStockId || "sin_punto"}`); setLitrosUsados(String(g.litrosUsados ?? ""));
      setForm({ origen: g.origen || "", monto: "", detalle: g.detalle || "", fecha: "", usuario: user.email, categoriaGasto: "Insumo", socio: g.socio || "" });
    } else {
      setTipo("general"); setInsumoSel(""); setLitrosUsados("");
      setForm({ origen: g.origen || "", monto: String(g.monto ?? ""), detalle: g.detalle || "", fecha: "", usuario: user.email, categoriaGasto: g.categoriaGasto || "Servicio", socio: g.socio || "" });
    }
    if (g.modoMonto === "porHa" && g.valorPorHectarea) { setModoMonto("porHa"); setValorPorHa(String(g.valorPorHectarea)); }
    else { setModoMonto("total"); setValorPorHa(""); }
    setArchivos([]); setPreviews([]); setTranscripcion(""); setError("");
    window.scrollTo({ top: 0, behavior: "smooth" });
  };
  const cancelarEdicion = () => {
    setEditId(null); setForm(emptyGasto(user.email)); setArchivos([]); setPreviews([]); setTranscripcion("");
    setTipo("general"); setInsumoSel(""); setLitrosUsados(""); setMoneda("USD"); setMontoPesos("");
    setModoMonto("total"); setValorPorHa("");
  };

  const onFile = (e) => {
    const files = Array.from(e.target.files || []);
    if (!files.length) return;
    setArchivos(files);
    setPreviews(files.map((f) => (f.type.startsWith("image/") ? URL.createObjectURL(f) : null)));
  };
  const quitarArchivo = (i) => {
    setArchivos(archivos.filter((_, idx) => idx !== i));
    setPreviews(previews.filter((_, idx) => idx !== i));
  };

  const extraerDeFactura = async () => {
    const primera = archivos[0];
    if (!primera || !primera.type.startsWith("image/")) { setError("La extracción automática funciona con fotos (JPG/PNG) — toma la primera imagen adjunta."); return; }
    setExtrayendo(true); setError("");
    try {
      const b64 = await fileToBase64(primera);
      const data = await askClaudeJSON([
        { type: "image", source: { type: "base64", media_type: primera.type, data: b64 } },
        { type: "text", text: `Esta imagen es una factura de un proveedor agropecuario uruguayo. Devolvé SOLO un JSON (sin texto extra, sin markdown) con: {"origen":"nombre del proveedor/emisor","monto":number (total de la factura en dólares),"detalle":"resumen breve de qué incluye la factura, 1 línea","fecha":"YYYY-MM-DD o vacío"}.` },
      ]);
      setForm({ ...form, ...data, usuario: user.email });
    } catch (e) { setError("No se pudo leer la factura automáticamente. Completá los campos a mano."); }
    finally { setExtrayendo(false); }
  };

  const grabar = () => {
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SR) { setError("Este navegador no soporta dictado por voz."); return; }
    const rec = new SR(); rec.lang = "es-UY"; rec.continuous = true; rec.interimResults = true;
    rec.onresult = (ev) => { let t = ""; for (let i = 0; i < ev.results.length; i++) t += ev.results[i][0].transcript + " "; setTranscripcion(t); };
    rec.onend = () => setGrabando(false);
    rec.start(); recRef.current = rec; setGrabando(true); setError("");
  };
  const detener = () => { recRef.current?.stop(); setGrabando(false); };
  const interpretarVoz = async () => {
    if (!transcripcion.trim()) return;
    setInterpretando(true); setError("");
    try {
      const data = await askClaudeJSON(`Convertí esta nota de voz de un productor agropecuario en JSON con el esquema {"origen":"string","monto":number,"detalle":"string","fecha":"YYYY-MM-DD o vacío"}. Solo el JSON, sin texto extra. Texto: """${transcripcion}"""`);
      setForm({ ...form, ...data, usuario: user.email });
    } catch (e) { setError("No se pudo interpretar el audio."); }
    finally { setInterpretando(false); }
  };

  const guardar = async () => {
    if (tipo === "insumo") {
      if (!insumoSel || !litrosUsados || !form.fecha) { alert("Elegí el insumo, los litros usados y la fecha."); return; }
      if (Number(litrosUsados) <= 0) { alert("Los litros usados tienen que ser un número mayor a 0."); return; }
      if (!editId && Number(litrosUsados) > (insumoElegido?.disponible || 0)) {
        if (!confirm(`Solo hay ${fmt(insumoElegido?.disponible || 0, 1)} ${abrevUnidad(insumoElegido?.unidad)} disponibles de ${insumoNombreSel} en ${insumoElegido?.puntoStockNombre || "ese punto"}. ¿Registrar igual el consumo?`)) return;
      }
    } else if (modoMonto === "porHa") {
      if (!superficie) { alert("Este cultivo todavía no tiene hectáreas cargadas — andá a la pestaña \"Lotes\" para asociarle lotes, o cargá el gasto por monto total."); return; }
      if (!valorPorHa || Number(valorPorHa) <= 0) { alert("Ingresá un valor por hectárea mayor a 0."); return; }
      if (!form.origen || !form.fecha) { alert("Completá al menos origen y fecha."); return; }
    } else if (moneda === "UYU" && (!montoPesos || !cotizacion || Number(montoPesos) <= 0 || Number(cotizacion) <= 0)) { alert("Completá el monto en pesos y la cotización del dólar (tienen que ser mayores a 0)."); return; }
    else if (!form.origen || !form.monto || !form.fecha) { alert("Completá al menos origen, monto y fecha."); return; }
    else if (Number(form.monto) <= 0) { alert("El monto tiene que ser un número mayor a 0."); return; }

    let facturaUrls, facturaNombres;
    if (archivos.length) {
      setSubiendo(true);
      try {
        const subidos = await Promise.all(archivos.map(async (a) => {
          const path = `facturas/${cultivo.id}/${Date.now()}_${a.name}`;
          const r = ref(storage, path);
          await uploadBytes(r, a);
          return { url: await getDownloadURL(r), nombre: a.name };
        }));
        facturaUrls = subidos.map((s) => s.url);
        facturaNombres = subidos.map((s) => s.nombre);
      } catch (e) { alert("No se pudieron subir todas las facturas adjuntas, pero se guardará el gasto."); }
      setSubiendo(false);
    }

    let datos;
    if (tipo === "insumo") {
      datos = {
        origen: form.origen || insumoNombreSel, monto: montoCalculado, detalle: form.detalle || `Consumo de ${insumoNombreSel} — ${litrosUsados} L`,
        fecha: form.fecha, usuario: form.usuario || user.email, insumoNombre: insumoNombreSel, litrosUsados: Number(litrosUsados), costoPorLitro: fifo.costoPromedioEfectivo,
        puntoStockId: puntoStockIdSel, puntoStockNombre: insumoElegido?.puntoStockNombre || "",
        categoriaGasto: "Insumo", socio: form.socio || "", modoMonto: null, valorPorHectarea: null,
      };
    } else {
      const montoFinal = modoMonto === "porHa" ? Number(valorPorHa) * superficie : Number(form.monto);
      datos = {
        origen: form.origen, monto: montoFinal, detalle: form.detalle, fecha: form.fecha, usuario: form.usuario || user.email,
        insumoNombre: null, litrosUsados: null, categoriaGasto: form.categoriaGasto || "Otro", socio: form.socio || "",
        modoMonto: modoMonto === "porHa" ? "porHa" : null,
        valorPorHectarea: modoMonto === "porHa" ? Number(valorPorHa) : null,
        hectareasUsadas: modoMonto === "porHa" ? superficie : null,
      };
      if (moneda === "UYU" && modoMonto !== "porHa") { datos.montoPesos = Number(montoPesos); datos.cotizacionUsada = Number(cotizacion); }
    }
    if (archivos.length) { datos.facturaUrls = facturaUrls; datos.facturaNombres = facturaNombres; }

    if (editId) {
      await api.update(editId, datos);
    } else {
      await api.add({ cultivoId: cultivo.id, ...datos, facturaUrls: facturaUrls || [], facturaNombres: facturaNombres || [] });
    }
    const eraEdicion = !!editId;
    setEditId(null); setForm(emptyGasto(user.email)); setArchivos([]); setPreviews([]); setTranscripcion("");
    setTipo("general"); setInsumoSel(""); setLitrosUsados(""); setMoneda("USD"); setMontoPesos("");
    setModoMonto("total"); setValorPorHa("");
    avisar(eraEdicion ? "Cambios guardados ✓" : "Gasto guardado ✓");
  };

  const eliminar = (id) => { if (confirm("Este gasto se moverá a la papelera. ¿Continuar?")) api.remove(id); };

  return (
    <div className="space-y-5">
      <div className="flex items-center gap-2" style={{ fontSize: 13, color: "#5A5647" }}>
        <MapPin size={16} color="var(--frost)" />
        Superficie de este cultivo: <b className="cc-mono">{fmt(superficie, 1)} ha</b>
        {!superficie && <span style={{ color: "#8A8570" }}>— asociá lotes en la pestaña "Lotes" para calcular costo/ha.</span>}
      </div>

      {puedeEditar && (
      <div className="cc-card p-4 space-y-4">
        <div className="flex flex-wrap gap-2">
          <label className="cc-btn cc-btn-ghost" style={{ cursor: "pointer" }}><Paperclip size={17} /> {archivos.length ? `${archivos.length} factura(s) elegida(s)` : "Adjuntar factura(s) (imagen o PDF)"}<input type="file" accept="image/*,.pdf" multiple onChange={onFile} style={{ display: "none" }} /></label>
          {archivos[0]?.type?.startsWith("image/") && <button className="cc-btn cc-btn-ghost" onClick={extraerDeFactura} disabled={extrayendo}>{extrayendo ? <Loader2 size={17} className="animate-spin" /> : <Camera size={17} />} Extraer datos con IA (1ra foto)</button>}
          <button className="cc-btn" style={{ background: grabando ? "var(--rust)" : "var(--soil)", color: "#fff" }} onClick={grabando ? detener : grabar}><Mic size={17} /> {grabando ? "Detener" : "Dictar por voz"}</button>
        </div>

        {archivos.length > 0 && (
          <div className="flex flex-wrap gap-2">
            {archivos.map((a, i) => (
              <div key={i} className="flex items-center gap-2" style={{ background: "#F7F5EC", border: "1px solid var(--line)", borderRadius: 8, padding: "6px 10px" }}>
                {previews[i] ? <img src={previews[i]} alt={a.name} style={{ width: 32, height: 32, objectFit: "cover", borderRadius: 4 }} /> : <FileText size={17} color="var(--frost)" />}
                <span style={{ fontSize: 12 }}>{a.name}</span>
                <button onClick={() => quitarArchivo(i)}><X size={15} color="var(--rust)" /></button>
              </div>
            ))}
          </div>
        )}

        {(transcripcion || grabando) && (
          <div>
            <textarea className="cc-input" rows={2} placeholder="Ej: flete a Nueva Palmira, ochocientos dólares, quince de marzo" value={transcripcion} onChange={(e) => setTranscripcion(e.target.value)} />
            <button className="cc-btn cc-btn-ghost mt-2" onClick={interpretarVoz} disabled={interpretando}>{interpretando ? <Loader2 size={17} className="animate-spin" /> : <Sprout size={17} />} Interpretar con IA</button>
          </div>
        )}

        {error && <div style={{ color: "var(--rust)", fontSize: 13, display: "flex", gap: 6, alignItems: "center" }}><AlertCircle size={17} />{error}</div>}

        {editId && (
          <div className="flex items-center gap-2" style={{ background: "#FDF3E0", border: "1px solid var(--gold)", borderRadius: 8, padding: "6px 10px", fontSize: 12.5, color: "#7A5A1E" }}>
            <Pencil size={16} /> Editando un gasto ya guardado.
          </div>
        )}

        <div className="flex gap-2">
          <button className="cc-btn" onClick={() => setTipo("general")} style={{ background: tipo === "general" ? "var(--soil)" : "#fff", color: tipo === "general" ? "#fff" : "var(--ink)", border: "1px solid var(--line)" }}>Gasto general</button>
          <button className="cc-btn" onClick={() => setTipo("insumo")} style={{ background: tipo === "insumo" ? "var(--soil)" : "#fff", color: tipo === "insumo" ? "#fff" : "var(--ink)", border: "1px solid var(--line)" }}><Boxes size={17} /> Insumo de stock</button>
        </div>

        {tipo === "insumo" ? (
          <div className="grid gap-3" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(150px,1fr))" }}>
            <div>
              <label style={{ fontSize: 12, color: "#8A8570" }}>Insumo (punto de stock)</label>
              <select className="cc-input" value={insumoSel} onChange={(e) => setInsumoSel(e.target.value)}>
                <option value="">Elegir...</option>
                {[...stockInsumos].sort((a, b) => a.nombre.localeCompare(b.nombre) || a.puntoStockNombre.localeCompare(b.puntoStockNombre)).map((i) => (
                  <option key={`${i.nombre}||${i.puntoStockId || "sin_punto"}`} value={`${i.nombre}||${i.puntoStockId || "sin_punto"}`}>
                    {i.nombre} — {i.puntoStockNombre} ({fmt(i.disponible, 1)} {abrevUnidad(i.unidad)} disp.)
                  </option>
                ))}
              </select>
              {stockInsumos.length === 0 && <div style={{ fontSize: 11.5, color: "#8A8570" }}>No hay insumos cargados todavía — cargalos desde "Insumos" en el menú superior.</div>}
            </div>
            <div><label style={{ fontSize: 12, color: "#8A8570" }}>Cantidad usada{insumoElegido ? ` (${abrevUnidad(insumoElegido.unidad)})` : ""}</label><input className="cc-input" type="number" value={litrosUsados} onChange={(e) => setLitrosUsados(e.target.value)} /></div>
            <div><label style={{ fontSize: 12, color: "#8A8570" }}>Fecha</label><input className="cc-input" type="date" value={form.fecha} onChange={(e) => set("fecha", e.target.value)} /></div>
            <div><label style={{ fontSize: 12, color: "#8A8570" }}>Costo estimado (FIFO)</label><input className="cc-input" value={`U$S ${fmt(montoCalculado, 2)}`} disabled /></div>
            <div>
              <label style={{ fontSize: 12, color: "#8A8570" }}>Socio que aporta (opcional)</label>
              <input className="cc-input" list="socios-gastos" value={form.socio} onChange={(e) => set("socio", e.target.value)} placeholder="Ej: Juan Pérez" />
              <datalist id="socios-gastos">{sociosSugeridos.map((s) => <option key={s} value={s} />)}</datalist>
            </div>
            <div style={{ gridColumn: "1 / -1" }}><label style={{ fontSize: 12, color: "#8A8570" }}>Detalle (opcional)</label><input className="cc-input" value={form.detalle} onChange={(e) => set("detalle", e.target.value)} placeholder={insumoNombreSel ? `Consumo de ${insumoNombreSel}` : ""} /></div>
          </div>
        ) : (
          <div className="grid gap-3" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(150px,1fr))" }}>
            <div>
              <label style={{ fontSize: 12, color: "#8A8570" }}>Categoría</label>
              <select className="cc-input" value={form.categoriaGasto} onChange={(e) => set("categoriaGasto", e.target.value)}>
                {CATEGORIAS_GASTO.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </div>
            <div>
              <label style={{ fontSize: 12, color: "#8A8570" }}>Origen</label>
              <input className="cc-input" list="origenes-gastos" value={form.origen} onChange={(e) => set("origen", e.target.value)} placeholder="Proveedor, gasoil, renta..." />
              <datalist id="origenes-gastos">{origenesSugeridos.map((o) => <option key={o} value={o} />)}</datalist>
            </div>
            <div>
              <label style={{ fontSize: 12, color: "#8A8570" }}>¿Cómo cargás el monto?</label>
              <div className="flex gap-1">
                <button type="button" onClick={() => setModoMonto("total")} className="cc-btn" style={{ padding: "8px 12px", fontSize: 12.5, background: modoMonto === "total" ? "var(--soil)" : "#fff", color: modoMonto === "total" ? "#fff" : "var(--ink)", border: "1px solid var(--line)" }}>Monto total</button>
                <button type="button" onClick={() => setModoMonto("porHa")} className="cc-btn" style={{ padding: "8px 12px", fontSize: 12.5, background: modoMonto === "porHa" ? "var(--soil)" : "#fff", color: modoMonto === "porHa" ? "#fff" : "var(--ink)", border: "1px solid var(--line)" }}>U$S / hectárea</button>
              </div>
            </div>

            {modoMonto === "porHa" ? (
              <>
                <div>
                  <label style={{ fontSize: 12, color: "#8A8570" }}>Valor por hectárea (U$S/ha)</label>
                  <input className="cc-input" type="number" value={valorPorHa} onChange={(e) => setValorPorHa(e.target.value)} placeholder="Ej: 25" />
                </div>
                <div style={{ fontSize: 12.5, color: "#8A8570", display: "flex", alignItems: "end", paddingBottom: 6 }}>
                  {superficie ? (
                    <span>{fmt(superficie, 1)} ha × {valorPorHa ? fmtUSD(Number(valorPorHa)) : "U$S -"} = <b style={{ color: "var(--ink)", marginLeft: 4 }}>{fmtUSD(Number(form.monto) || 0)}</b></span>
                  ) : (
                    <span style={{ color: "var(--rust)" }}>Este cultivo no tiene hectáreas cargadas — andá a "Lotes" primero.</span>
                  )}
                </div>
              </>
            ) : moneda === "USD" ? (
              <div><label style={{ fontSize: 12, color: "#8A8570" }}>Monto (U$S)</label><input className="cc-input" type="number" value={form.monto} onChange={(e) => set("monto", e.target.value)} /></div>
            ) : (
              <>
                <div><label style={{ fontSize: 12, color: "#8A8570" }}>Monto ($ pesos)</label><input className="cc-input" type="number" value={montoPesos} onChange={(e) => setMontoPesos(e.target.value)} /></div>
                <div>
                  <label style={{ fontSize: 12, color: "#8A8570" }}>Cotización del dólar {cotizandoAuto && "(buscando...)"}</label>
                  <input className="cc-input" type="number" value={cotizacion} onChange={(e) => setCotizacion(e.target.value)} />
                </div>
                <div style={{ fontSize: 12.5, color: "#8A8570", display: "flex", alignItems: "end", paddingBottom: 6 }}>Equivale a: <b style={{ color: "var(--ink)", marginLeft: 4 }}>{fmtUSD(Number(form.monto) || 0)}</b></div>
              </>
            )}

            {modoMonto !== "porHa" && (
              <div>
                <label style={{ fontSize: 12, color: "#8A8570" }}>Moneda</label>
                <div className="flex gap-1">
                  <button type="button" onClick={() => setMoneda("USD")} className="cc-btn" style={{ padding: "8px 12px", fontSize: 12.5, background: moneda === "USD" ? "var(--soil)" : "#fff", color: moneda === "USD" ? "#fff" : "var(--ink)", border: "1px solid var(--line)" }}>Dólares</button>
                  <button type="button" onClick={() => setMoneda("UYU")} className="cc-btn" style={{ padding: "8px 12px", fontSize: 12.5, background: moneda === "UYU" ? "var(--gold)" : "#fff", color: moneda === "UYU" ? "#fff" : "var(--ink)", border: "1px solid var(--line)" }}>Pesos</button>
                </div>
              </div>
            )}
            <div><label style={{ fontSize: 12, color: "#8A8570" }}>Fecha</label><input className="cc-input" type="date" value={form.fecha} onChange={(e) => set("fecha", e.target.value)} /></div>
            <div><label style={{ fontSize: 12, color: "#8A8570" }}>Usuario</label><input className="cc-input" value={form.usuario} disabled /></div>
            <div>
              <label style={{ fontSize: 12, color: "#8A8570" }}>Socio que aporta (opcional)</label>
              <input className="cc-input" list="socios-gastos" value={form.socio} onChange={(e) => set("socio", e.target.value)} placeholder="Ej: Juan Pérez" />
              <datalist id="socios-gastos">{sociosSugeridos.map((s) => <option key={s} value={s} />)}</datalist>
            </div>
            <div style={{ gridColumn: "1 / -1" }}><label style={{ fontSize: 12, color: "#8A8570" }}>Detalle</label><input className="cc-input" value={form.detalle} onChange={(e) => set("detalle", e.target.value)} placeholder="Descripción del gasto" /></div>
          </div>
        )}
        {editId && archivos.length === 0 && <div style={{ fontSize: 12, color: "#8A8570" }}>Las facturas adjuntas actuales se mantienen salvo que subas nuevas arriba.</div>}
        <div className="flex gap-2 items-center flex-wrap">
          <button className="cc-btn cc-btn-primary" onClick={guardar} disabled={subiendo}>
            {subiendo ? <Loader2 size={18} className="animate-spin" /> : editId ? <Pencil size={18} /> : <Plus size={18} />} {editId ? "Guardar cambios" : "Guardar gasto"}
          </button>
          {editId && <button className="cc-btn cc-btn-ghost" onClick={cancelarEdicion}><X size={18} /> Cancelar</button>}
          {mensaje && <span style={{ color: "var(--soil-light)", fontWeight: 700, fontSize: 13.5 }}>{mensaje}</span>}
        </div>
      </div>
      )}

      {gastos.length > 0 && (
        <div className="flex items-center gap-3 flex-wrap">
          <div style={{ maxWidth: 320, flex: 1 }}>
            <input className="cc-input" value={busqueda} onChange={(e) => setBusqueda(e.target.value)} placeholder="Buscar por origen, detalle o usuario..." />
          </div>
          <div style={{ width: 190 }}>
            <select className="cc-input" value={filtroCategoria} onChange={(e) => setFiltroCategoria(e.target.value)}>
              <option value="todas">Todas las categorías</option>
              <option value="Insumo">Insumo</option>
              {CATEGORIAS_GASTO.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          <button className="cc-btn cc-btn-ghost" onClick={() => exportarExcel(`gastos_${cultivo.nombre}`, [{ nombre: "Gastos", filas: gastos.map((g) => ({ Fecha: g.fecha, Categoría: g.categoriaGasto || (g.insumoNombre ? "Insumo" : "Otro"), Origen: g.origen, Socio: g.socio || "", Detalle: g.detalle, Insumo: g.insumoNombre || "", "Litros usados": g.litrosUsados || "", "U$S/ha": g.valorPorHectarea || "", Usuario: g.usuario, Monto: g.monto })) }])}>
            <Download size={17} /> Exportar Excel
          </button>
        </div>
      )}

      {gastos.length === 0 ? <EmptyState icon={Receipt} title="Sin gastos cargados" text="Cargá el primer gasto de este cultivo, escrito, por voz o desde una foto de factura." /> : (
        <div className="cc-card overflow-hidden">
          <div className="flex items-center gap-4 px-3 py-2" style={{ fontSize: 11.5, color: "#8A8570", borderBottom: "1px solid var(--line)" }}>
            <span className="flex items-center gap-1"><span style={{ width: 10, height: 10, borderRadius: "50%", background: "#FDF3E0", border: "1px solid var(--gold)", display: "inline-block" }} /> Programado (fecha futura)</span>
            <span className="flex items-center gap-1"><span style={{ width: 10, height: 10, borderRadius: "50%", background: "#fff", border: "1px solid var(--line)", display: "inline-block" }} /> Realizado</span>
          </div>
          <table className="w-full" style={{ fontSize: 13 }}>
            <thead><tr style={{ background: "#EEEADA", textAlign: "left" }}><th className="px-3 py-2">Fecha</th><th className="px-3 py-2">Categoría</th><th className="px-3 py-2">Origen</th><th className="px-3 py-2">Socio</th><th className="px-3 py-2">Detalle</th><th className="px-3 py-2">Usuario</th><th className="px-3 py-2 text-right">Monto</th><th className="px-3 py-2"></th><th className="px-3 py-2"></th><th className="px-3 py-2"></th><th className="px-3 py-2"></th></tr></thead>
            <tbody>
              {[...gastos]
                .filter((g) => {
                  const q = busqueda.trim().toLowerCase();
                  const catG = g.categoriaGasto || (g.insumoNombre ? "Insumo" : "Otro");
                  if (filtroCategoria !== "todas" && catG !== filtroCategoria) return false;
                  if (!q) return true;
                  return [g.origen, g.detalle, g.usuario].some((v) => (v || "").toLowerCase().includes(q));
                })
                .sort((a, b) => (b.fecha || "").localeCompare(a.fecha || "")).map((g) => {
                const hoyStr = new Date().toISOString().slice(0, 10);
                const esProgramado = g.fecha && g.fecha > hoyStr;
                return (
                <tr key={g.id} style={{ borderTop: "1px solid var(--line)", background: editId === g.id ? "#FDF3E0" : esProgramado ? "#FDF8EC" : "transparent" }}>
                  <td className="px-3 py-2 cc-mono" style={{ color: esProgramado ? "#B8860B" : "var(--ink)", fontWeight: esProgramado ? 700 : 400 }}>
                    {g.fecha}
                    {esProgramado && <div style={{ fontSize: 10, fontWeight: 600, color: "var(--gold)" }}>Programado</div>}
                  </td>
                  <td className="px-3 py-2">
                    <span className="cc-chip" style={{ background: (GASTO_CAT_COLOR[g.categoriaGasto] || GASTO_CAT_COLOR.Otro) + "22", color: GASTO_CAT_COLOR[g.categoriaGasto] || GASTO_CAT_COLOR.Otro }}>
                      {g.categoriaGasto || (g.insumoNombre ? "Insumo" : "Otro")}
                    </span>
                  </td>
                  <td className="px-3 py-2">
                    {g.origen}
                    {g.insumoNombre && <span className="cc-chip" style={{ background: "#EDE7F6", color: "#5B4B8A", marginLeft: 6 }}>{fmt(g.litrosUsados, 1)} L{g.puntoStockNombre ? ` · ${g.puntoStockNombre}` : ""}</span>}
                    {g.modoMonto === "porHa" && <span className="cc-chip" style={{ background: "#E8F0EA", color: "var(--soil-light)", marginLeft: 6 }}>{fmtUSD(g.valorPorHectarea)}/ha</span>}
                  </td>
                  <td className="px-3 py-2" style={{ color: "#5A5647" }}>{g.socio || "-"}</td>
                  <td className="px-3 py-2" style={{ color: "#5A5647" }}>{g.detalle}</td>
                  <td className="px-3 py-2" style={{ color: "#8A8570", fontSize: 12 }}>
                    {g.usuario}
                    {g.modificadoPor && g.modificadoPor !== g.usuario && <div style={{ fontStyle: "italic", fontSize: 10.5 }}>editado por {g.modificadoPor}</div>}
                  </td>
                  <td className="px-3 py-2 text-right cc-mono">
                    {fmtUSD(g.monto)}
                    {g.montoPesos ? <div style={{ fontSize: 10.5, color: "#8A8570", fontWeight: 400 }}>$ {fmt(g.montoPesos, 0)} @ {fmt(g.cotizacionUsada, 2)}</div> : null}
                    {g.modoMonto === "porHa" ? <div style={{ fontSize: 10.5, color: "#8A8570", fontWeight: 400 }}>{fmt(g.hectareasUsadas, 1)} ha</div> : null}
                  </td>
                  <td className="px-3 py-2">
                    <div className="flex gap-1">
                      {(g.facturaUrls && g.facturaUrls.length ? g.facturaUrls : g.facturaUrl ? [g.facturaUrl] : []).map((u, i) => (
                        <a key={i} href={u} target="_blank" rel="noreferrer"><FileText size={17} color="var(--frost)" /></a>
                      ))}
                    </div>
                  </td>
                  <td className="px-3 py-2 text-right"><button onClick={() => duplicar(g)} title="Duplicar"><Copy size={16} color="#8A8570" /></button></td>
                  <td className="px-3 py-2 text-right">{puedeEditar && <button onClick={() => editar(g)} title="Editar"><Pencil size={16} color="var(--frost)" /></button>}</td>
                  <td className="px-3 py-2 text-right">{puedeEditar && <button onClick={() => eliminar(g.id)} title="Eliminar"><Trash2 size={16} color="var(--rust)" /></button>}</td>
                </tr>
                );
              })}
            </tbody>
            <tfoot><tr style={{ borderTop: "2px solid var(--line)", fontWeight: 700 }}><td className="px-3 py-2" colSpan={6}>Total</td><td className="px-3 py-2 text-right cc-mono">{fmtUSD(gastos.reduce((s, g) => s + Number(g.monto || 0), 0))}</td><td colSpan={4}></td></tr></tfoot>
          </table>
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Ingresos · Ventas                                                    */
/* ------------------------------------------------------------------ */
function VentasTab({ cultivo, ventas, api, puedeEditar = true }) {
  const [form, setForm] = useState({ fecha: "", origen: "", toneladas: "", dolaresPorTonelada: "" });
  const set = (k, v) => setForm({ ...form, [k]: v });
  const [mensaje, setMensaje] = useState("");
  const origenesSugeridos = Array.from(new Set(ventas.map((v) => v.origen).filter(Boolean)));

  const guardar = () => {
    if (!form.fecha || !form.toneladas || !form.dolaresPorTonelada) { alert("Completá fecha, toneladas y U$S/ton."); return; }
    if (Number(form.toneladas) <= 0 || Number(form.dolaresPorTonelada) <= 0) { alert("Toneladas y U$S/ton tienen que ser números mayores a 0."); return; }
    api.add({ cultivoId: cultivo.id, fecha: form.fecha, origen: form.origen, toneladas: Number(form.toneladas), dolaresPorTonelada: Number(form.dolaresPorTonelada) });
    setForm({ fecha: "", origen: "", toneladas: "", dolaresPorTonelada: "" });
    setMensaje("Venta guardada ✓"); setTimeout(() => setMensaje(""), 2500);
  };
  const eliminar = (id) => { if (confirm("Esta venta se moverá a la papelera. ¿Continuar?")) api.remove(id); };
  const totalTon = ventas.reduce((s, v) => s + Number(v.toneladas || 0), 0);
  const totalUSD = ventas.reduce((s, v) => s + Number(v.toneladas || 0) * Number(v.dolaresPorTonelada || 0), 0);

  return (
    <div className="space-y-5">
      {puedeEditar && (
      <div className="cc-card p-4 space-y-3">
        <div className="grid gap-3" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(150px,1fr))" }}>
          <div><label style={{ fontSize: 12, color: "#8A8570" }}>Fecha</label><input className="cc-input" type="date" value={form.fecha} onChange={(e) => set("fecha", e.target.value)} /></div>
          <div>
            <label style={{ fontSize: 12, color: "#8A8570" }}>Origen</label>
            <input className="cc-input" list="origenes-ventas" value={form.origen} onChange={(e) => set("origen", e.target.value)} placeholder="Exportadora / acopiador" />
            <datalist id="origenes-ventas">{origenesSugeridos.map((o) => <option key={o} value={o} />)}</datalist>
          </div>
          <div><label style={{ fontSize: 12, color: "#8A8570" }}>Toneladas</label><input className="cc-input" type="number" value={form.toneladas} onChange={(e) => set("toneladas", e.target.value)} /></div>
          <div><label style={{ fontSize: 12, color: "#8A8570" }}>U$S / tonelada</label><input className="cc-input" type="number" value={form.dolaresPorTonelada} onChange={(e) => set("dolaresPorTonelada", e.target.value)} /></div>
        </div>
        <div className="flex items-center gap-3">
          <button className="cc-btn cc-btn-primary" onClick={guardar}><Plus size={18} /> Guardar venta</button>
          {mensaje && <span style={{ color: "var(--soil-light)", fontWeight: 700, fontSize: 13.5 }}>{mensaje}</span>}
        </div>
      </div>
      )}

      {ventas.length > 0 && (
        <div className="flex justify-end">
          <button className="cc-btn cc-btn-ghost" onClick={() => exportarExcel(`ventas_${cultivo.nombre}`, [{ nombre: "Ventas", filas: ventas.map((v) => ({ Fecha: v.fecha, Origen: v.origen, Toneladas: v.toneladas, "U$S/ton": v.dolaresPorTonelada, Total: v.toneladas * v.dolaresPorTonelada })) }])}>
            <Download size={17} /> Exportar Excel
          </button>
        </div>
      )}

      {ventas.length === 0 ? <EmptyState icon={DollarSign} title="Sin ventas cargadas" text="Cargá las ventas de este cultivo: fecha, origen, toneladas y precio por tonelada." /> : (
        <div className="cc-card overflow-hidden">
          <table className="w-full" style={{ fontSize: 13 }}>
            <thead><tr style={{ background: "#EEEADA", textAlign: "left" }}><th className="px-3 py-2">Fecha</th><th className="px-3 py-2">Origen</th><th className="px-3 py-2 text-right">Toneladas</th><th className="px-3 py-2 text-right">U$S/ton</th><th className="px-3 py-2 text-right">Total</th><th className="px-3 py-2"></th></tr></thead>
            <tbody>
              {[...ventas].sort((a, b) => (b.fecha || "").localeCompare(a.fecha || "")).map((v) => (
                <tr key={v.id} style={{ borderTop: "1px solid var(--line)" }}>
                  <td className="px-3 py-2 cc-mono">{v.fecha}</td><td className="px-3 py-2">{v.origen}</td>
                  <td className="px-3 py-2 text-right cc-mono">{fmt(v.toneladas, 2)}</td>
                  <td className="px-3 py-2 text-right cc-mono">{fmtUSD(v.dolaresPorTonelada)}</td>
                  <td className="px-3 py-2 text-right cc-mono">{fmtUSD(v.toneladas * v.dolaresPorTonelada)}</td>
                  <td className="px-3 py-2 text-right">{puedeEditar && <button onClick={() => eliminar(v.id)}><Trash2 size={16} color="var(--rust)" /></button>}</td>
                </tr>
              ))}
            </tbody>
            <tfoot><tr style={{ borderTop: "2px solid var(--line)", fontWeight: 700 }}><td className="px-3 py-2" colSpan={2}>Total</td><td className="px-3 py-2 text-right cc-mono">{fmt(totalTon, 2)}</td><td></td><td className="px-3 py-2 text-right cc-mono">{fmtUSD(totalUSD)}</td><td></td></tr></tfoot>
          </table>
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Ingresos · Remitos                                                   */
/* ------------------------------------------------------------------ */
function RemitosTab({ cultivo, remitos, api, lotes, campos = [], totalToneladasVentas, puedeEditar = true }) {
  const [form, setForm] = useState({ fecha: "", remito: "", campoId: "", loteId: "", destino: "", kgTolva: "", kgBrutos: "", kgSL: "", humedad: "" });
  const set = (k, v) => setForm({ ...form, [k]: v });
  const [mensaje, setMensaje] = useState("");

  const lotesDelCampo = form.campoId ? lotes.filter((l) => l.campoId === form.campoId) : [];

  const elegirCampo = (campoId) => setForm({ ...form, campoId, loteId: "" });

  const guardar = () => {
    if (!form.fecha || !form.remito || !form.kgSL) { alert("Completá al menos fecha, N° de remito y Kg SL."); return; }
    if (Number(form.kgSL) <= 0) { alert("Los Kg SL tienen que ser un número mayor a 0."); return; }
    if (form.humedad && (Number(form.humedad) < 0 || Number(form.humedad) > 100)) { alert("La humedad tiene que ser un porcentaje entre 0 y 100."); return; }
    const campoNombre = campos.find((c) => c.id === form.campoId)?.nombre || "";
    const loteNombre = lotes.find((l) => l.id === form.loteId)?.nombre || "";
    api.add({
      cultivoId: cultivo.id, fecha: form.fecha, remito: form.remito,
      campoId: form.campoId || null, campo: campoNombre, loteId: form.loteId || null, lote: loteNombre,
      destino: form.destino,
      kgTolva: Number(form.kgTolva || 0), kgBrutos: Number(form.kgBrutos || 0), kgSL: Number(form.kgSL || 0), humedad: Number(form.humedad || 0),
    });
    setForm({ fecha: "", remito: "", campoId: "", loteId: "", destino: "", kgTolva: "", kgBrutos: "", kgSL: "", humedad: "" });
    setMensaje("Remito guardado ✓"); setTimeout(() => setMensaje(""), 2500);
  };
  const eliminar = (id) => { if (confirm("Este remito se moverá a la papelera. ¿Continuar?")) api.remove(id); };
  const totalKgSL = remitos.reduce((s, r) => s + Number(r.kgSL || 0), 0);

  return (
    <div className="space-y-5">
      {puedeEditar && (
      <div className="cc-card p-4 space-y-3">
        <div className="grid gap-3" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(130px,1fr))" }}>
          <div><label style={{ fontSize: 12, color: "#8A8570" }}>Fecha</label><input className="cc-input" type="date" value={form.fecha} onChange={(e) => set("fecha", e.target.value)} /></div>
          <div><label style={{ fontSize: 12, color: "#8A8570" }}>N° Remito</label><input className="cc-input" value={form.remito} onChange={(e) => set("remito", e.target.value)} /></div>
          <div>
            <label style={{ fontSize: 12, color: "#8A8570" }}>Campo</label>
            <select className="cc-input" value={form.campoId} onChange={(e) => elegirCampo(e.target.value)}>
              <option value="">Elegir campo...</option>
              {campos.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}
            </select>
          </div>
          <div>
            <label style={{ fontSize: 12, color: "#8A8570" }}>Lote</label>
            <select className="cc-input" value={form.loteId} onChange={(e) => set("loteId", e.target.value)} disabled={!form.campoId}>
              <option value="">{form.campoId ? "Elegir lote..." : "Elegí un campo primero"}</option>
              {lotesDelCampo.map((l) => <option key={l.id} value={l.id}>{l.nombre}{l.hectareas ? ` (${fmt(l.hectareas, 1)} ha)` : ""}</option>)}
            </select>
            {form.campoId && lotesDelCampo.length === 0 && <div style={{ fontSize: 11.5, color: "#8A8570" }}>Ese campo todavía no tiene lotes cargados.</div>}
          </div>
          <div><label style={{ fontSize: 12, color: "#8A8570" }}>Destino</label><input className="cc-input" value={form.destino} onChange={(e) => set("destino", e.target.value)} placeholder="Silo / planta" /></div>
          <div><label style={{ fontSize: 12, color: "#8A8570" }}>Kg tolva</label><input className="cc-input" type="number" value={form.kgTolva} onChange={(e) => set("kgTolva", e.target.value)} /></div>
          <div><label style={{ fontSize: 12, color: "#8A8570" }}>Kg brutos</label><input className="cc-input" type="number" value={form.kgBrutos} onChange={(e) => set("kgBrutos", e.target.value)} /></div>
          <div><label style={{ fontSize: 12, color: "#8A8570" }}>Kg SL</label><input className="cc-input" type="number" value={form.kgSL} onChange={(e) => set("kgSL", e.target.value)} /></div>
          <div><label style={{ fontSize: 12, color: "#8A8570" }}>Humedad (%)</label><input className="cc-input" type="number" value={form.humedad} onChange={(e) => set("humedad", e.target.value)} /></div>
        </div>
        <div className="flex items-center gap-3">
          <button className="cc-btn cc-btn-primary" onClick={guardar}><Plus size={18} /> Guardar remito</button>
          {mensaje && <span style={{ color: "var(--soil-light)", fontWeight: 700, fontSize: 13.5 }}>{mensaje}</span>}
        </div>
      </div>
      )}

      {remitos.length > 0 && (
        <div className="flex justify-end">
          <button className="cc-btn cc-btn-ghost" onClick={() => exportarExcel(`remitos_${cultivo.nombre}`, [{ nombre: "Remitos", filas: remitos.map((r) => ({ Fecha: r.fecha, Remito: r.remito, Campo: r.campo, Lote: r.lote, Destino: r.destino, "Kg tolva": r.kgTolva, "Kg brutos": r.kgBrutos, "Kg SL": r.kgSL, "Humedad %": r.humedad })) }])}>
            <Download size={17} /> Exportar Excel
          </button>
        </div>
      )}

      {remitos.length === 0 ? <EmptyState icon={Truck} title="Sin remitos cargados" text="Cargá cada remito de entrega de grano con sus kilos seco y limpio (Kg SL)." /> : (
        <div className="cc-card overflow-hidden">
          <table className="w-full" style={{ fontSize: 12.5 }}>
            <thead><tr style={{ background: "#EEEADA", textAlign: "left" }}><th className="px-3 py-2">Fecha</th><th className="px-3 py-2">Remito</th><th className="px-3 py-2">Campo</th><th className="px-3 py-2">Lote</th><th className="px-3 py-2">Destino</th><th className="px-3 py-2 text-right">Kg tolva</th><th className="px-3 py-2 text-right">Kg brutos</th><th className="px-3 py-2 text-right">Kg SL</th><th className="px-3 py-2 text-right">Hum.%</th><th className="px-3 py-2"></th></tr></thead>
            <tbody>
              {[...remitos].sort((a, b) => (b.fecha || "").localeCompare(a.fecha || "")).map((r) => (
                <tr key={r.id} style={{ borderTop: "1px solid var(--line)" }}>
                  <td className="px-3 py-2 cc-mono">{r.fecha}</td><td className="px-3 py-2">{r.remito}</td><td className="px-3 py-2">{r.campo || "-"}</td><td className="px-3 py-2">{r.lote || "-"}</td><td className="px-3 py-2">{r.destino}</td>
                  <td className="px-3 py-2 text-right cc-mono">{fmt(r.kgTolva)}</td><td className="px-3 py-2 text-right cc-mono">{fmt(r.kgBrutos)}</td>
                  <td className="px-3 py-2 text-right cc-mono">{fmt(r.kgSL)}</td><td className="px-3 py-2 text-right cc-mono">{fmt(r.humedad, 1)}</td>
                  <td className="px-3 py-2 text-right">{puedeEditar && <button onClick={() => eliminar(r.id)}><Trash2 size={16} color="var(--rust)" /></button>}</td>
                </tr>
              ))}
            </tbody>
            <tfoot><tr style={{ borderTop: "2px solid var(--line)", fontWeight: 700 }}><td className="px-3 py-2" colSpan={7}>Total Kg SL</td><td className="px-3 py-2 text-right cc-mono">{fmt(totalKgSL)}</td><td colSpan={2}></td></tr></tfoot>
          </table>
          <div className="px-3 py-2" style={{ fontSize: 12, color: "#8A8570", borderTop: "1px solid var(--line)" }}>
            Equivalente a {fmt(totalKgSL / 1000, 2)} tn — {Math.abs(totalKgSL / 1000 - totalToneladasVentas) < 0.05 && totalToneladasVentas > 0 ? "coincide con las ventas cargadas." : `las ventas cargadas suman ${fmt(totalToneladasVentas, 2)} tn.`}
          </div>
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Usuarios y roles (solo admins)                                       */
/* ------------------------------------------------------------------ */
function UsuariosView({ usuarios, miEmail }) {
  const [email, setEmail] = useState("");
  const [rol, setRol] = useState("lectura");
  const [workspaceId, setWorkspaceId] = useState("");
  const [editandoWorkspace, setEditandoWorkspace] = useState(null);
  const [workspaceEditado, setWorkspaceEditado] = useState("");

  const agregar = async () => {
    const e = email.trim().toLowerCase();
    if (!e || !e.includes("@")) { alert("Ingresá un email válido."); return; }
    if (e === miEmail.toLowerCase()) { alert("No podés cambiarte el rol a vos mismo desde acá (para evitar quedarte afuera por error)."); return; }
    await setDoc(doc(db, "usuarios", e), { email: e, rol, workspaceId: workspaceId.trim() || "default" });
    setEmail(""); setWorkspaceId("");
  };

  const cambiarRol = async (id, nuevoRol) => {
    if (id === miEmail.toLowerCase()) { alert("No podés cambiarte el rol a vos mismo desde acá."); return; }
    const actual = usuarios.find((u) => u.id === id);
    await setDoc(doc(db, "usuarios", id), { email: id, rol: nuevoRol, workspaceId: actual?.workspaceId || "default" });
  };

  const guardarWorkspace = async (id) => {
    if (id === miEmail.toLowerCase()) { alert("No podés cambiarte tu propio espacio de trabajo desde acá."); return; }
    const actual = usuarios.find((u) => u.id === id);
    await setDoc(doc(db, "usuarios", id), { email: id, rol: actual?.rol || "lectura", workspaceId: workspaceEditado.trim() || "default" });
    setEditandoWorkspace(null);
  };

  const quitar = async (id) => {
    if (id === miEmail.toLowerCase()) { alert("No podés quitarte a vos mismo de la lista."); return; }
    if (confirm("Este usuario va a quedar como editor normal, en el espacio de trabajo compartido. ¿Continuar?")) await deleteDoc(doc(db, "usuarios", id));
  };

  return (
    <div className="space-y-5">
      <div>
        <div className="cc-h" style={{ fontSize: 20, fontWeight: 600 }}>Usuarios</div>
        <div style={{ fontSize: 12.5, color: "#8A8570" }}>Administrá quién puede editar datos y en qué espacio de trabajo entra cada uno. Si un email no aparece en esta lista, tiene acceso completo de edición en el espacio compartido por defecto.</div>
      </div>

      <div className="cc-card p-4">
        <div className="cc-h" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Asignar rol y espacio a un usuario</div>
        <div className="flex gap-3 flex-wrap items-end">
          <div style={{ flex: 1, minWidth: 200 }}>
            <label style={{ fontSize: 12, color: "#8A8570" }}>Email del usuario</label>
            <input className="cc-input" type="email" value={email} onChange={(e) => setEmail(e.target.value)} placeholder="persona@ejemplo.com" />
          </div>
          <div style={{ width: 170 }}>
            <label style={{ fontSize: 12, color: "#8A8570" }}>Rol</label>
            <select className="cc-input" value={rol} onChange={(e) => setRol(e.target.value)}>
              <option value="lectura">Solo lectura</option>
              <option value="admin">Editor / Admin</option>
            </select>
          </div>
          <div style={{ width: 220 }}>
            <label style={{ fontSize: 12, color: "#8A8570" }}>Espacio de trabajo (opcional)</label>
            <input className="cc-input" value={workspaceId} onChange={(e) => setWorkspaceId(e.target.value)} placeholder="Dejalo vacío para el compartido" />
          </div>
          <button className="cc-btn cc-btn-primary" onClick={agregar}><Plus size={18} /> Guardar</button>
        </div>
        <div style={{ fontSize: 11.5, color: "#8A8570", marginTop: 8 }}>
          La persona tiene que crear su cuenta con este mismo email desde la pantalla de ingreso. Si dejás el "Espacio de trabajo" vacío, comparte campañas, cultivos y todo lo demás con el resto del equipo (como hasta ahora).
          Si le ponés un nombre propio (ej: <code>juan-privado</code>), esa persona va a ver una app completamente vacía y separada — sus campañas, gastos y todo lo que cargue no lo va a ver nadie más, ni vos, salvo que compartas ese mismo nombre de espacio con otro usuario.
        </div>
      </div>

      {usuarios.length === 0 ? (
        <EmptyState icon={Users} title="Todos comparten el mismo espacio" text="Todavía no le asignaste un rol o espacio de trabajo distinto a nadie." />
      ) : (
        <div className="cc-card overflow-hidden">
          <table className="w-full" style={{ fontSize: 13 }}>
            <thead><tr style={{ background: "#EEEADA", textAlign: "left" }}><th className="px-3 py-2">Email</th><th className="px-3 py-2">Rol</th><th className="px-3 py-2">Espacio de trabajo</th><th className="px-3 py-2"></th></tr></thead>
            <tbody>
              {usuarios.map((u) => (
                <tr key={u.id} style={{ borderTop: "1px solid var(--line)" }}>
                  <td className="px-3 py-2">{u.id}{u.id === miEmail.toLowerCase() && <span className="cc-chip" style={{ marginLeft: 6, background: "#EEEADA" }}>vos</span>}</td>
                  <td className="px-3 py-2">
                    <select className="cc-input" style={{ padding: "5px 8px", fontSize: 12.5, width: 160 }} value={u.rol} disabled={u.id === miEmail.toLowerCase()} onChange={(e) => cambiarRol(u.id, e.target.value)}>
                      <option value="lectura">Solo lectura</option>
                      <option value="admin">Editor / Admin</option>
                    </select>
                  </td>
                  <td className="px-3 py-2">
                    {editandoWorkspace === u.id ? (
                      <div className="flex gap-1 items-center">
                        <input className="cc-input" style={{ padding: "5px 8px", fontSize: 12.5, width: 160 }} value={workspaceEditado} onChange={(e) => setWorkspaceEditado(e.target.value)} placeholder="default" />
                        <button onClick={() => guardarWorkspace(u.id)} title="Guardar"><CheckCircle2 size={17} color="var(--soil-light)" /></button>
                        <button onClick={() => setEditandoWorkspace(null)} title="Cancelar"><X size={17} color="var(--rust)" /></button>
                      </div>
                    ) : (
                      <button
                        onClick={() => { if (u.id !== miEmail.toLowerCase()) { setEditandoWorkspace(u.id); setWorkspaceEditado(u.workspaceId && u.workspaceId !== "default" ? u.workspaceId : ""); } }}
                        style={{ color: u.workspaceId && u.workspaceId !== "default" ? "var(--frost)" : "#8A8570", fontWeight: u.workspaceId && u.workspaceId !== "default" ? 700 : 400 }}
                        disabled={u.id === miEmail.toLowerCase()}
                      >
                        {u.workspaceId && u.workspaceId !== "default" ? u.workspaceId : "Compartido (default)"}
                      </button>
                    )}
                  </td>
                  <td className="px-3 py-2 text-right">{u.id !== miEmail.toLowerCase() && <button onClick={() => quitar(u.id)}><Trash2 size={16} color="var(--rust)" /></button>}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}


/* ------------------------------------------------------------------ */
/*  Campos (establecimientos)                                           */
/* ------------------------------------------------------------------ */
function CamposView({ campos, api, lotesApi, lotes, onOpen, puedeEditar = true }) {
  const [nombre, setNombre] = useState("");
  const lotesSinCampo = lotes.filter((l) => !l.campoId);

  const crear = () => { if (!nombre.trim()) return; api.add({ nombre: nombre.trim() }); setNombre(""); };
  const eliminar = (id) => {
    if (!confirm("Este campo y todos sus lotes se moverán a la papelera (se puede restaurar después). ¿Continuar?")) return;
    api.remove(id);
    lotes.filter((l) => l.campoId === id).forEach((l) => lotesApi.remove(l.id));
  };

  return (
    <div className="space-y-5">
      {puedeEditar && (
        <div className="cc-card p-4">
          <div className="cc-h" style={{ fontSize: 16, fontWeight: 600, marginBottom: 10 }}>Nuevo campo</div>
          <div className="flex gap-3 flex-wrap items-end">
            <div style={{ flex: 1, minWidth: 180 }}>
              <label style={{ fontSize: 12, color: "#8A8570" }}>Nombre</label>
              <input className="cc-input" value={nombre} onChange={(e) => setNombre(e.target.value)} placeholder="Ej: Establecimiento La Loma" />
            </div>
            <button className="cc-btn cc-btn-primary" onClick={crear}><Plus size={18} /> Agregar</button>
          </div>
        </div>
      )}

      {campos.length === 0 && lotesSinCampo.length === 0 ? (
        <EmptyState icon={MapPin} title="Todavía no hay campos cargados" text="Creá un campo (establecimiento) para después agruparle sus lotes." />
      ) : (
        <div className="grid gap-4" style={{ gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))" }}>
          {lotesSinCampo.length > 0 && (
            <div className="cc-card p-4 cursor-pointer" style={{ border: "1.5px dashed var(--line)" }} onClick={() => onOpen("__sin_campo__")}>
              <div className="flex items-start justify-between">
                <div>
                  <div className="cc-h" style={{ fontSize: 17, fontWeight: 600, color: "#8A8570" }}>Sin campo asignado</div>
                  <div style={{ fontSize: 12, color: "#8A8570" }}>{lotesSinCampo.length} lote{lotesSinCampo.length !== 1 ? "s" : ""} suelto{lotesSinCampo.length !== 1 ? "s" : ""}</div>
                </div>
              </div>
              <div className="flex items-center gap-1 mt-3" style={{ color: "var(--frost)", fontSize: 12.5, fontWeight: 600 }}>Ver y ordenar <ChevronRight size={16} /></div>
            </div>
          )}
          {campos.map((c) => {
            const nLotes = lotes.filter((l) => l.campoId === c.id).length;
            const hasTotal = lotes.filter((l) => l.campoId === c.id).reduce((s, l) => s + Number(l.hectareas || 0), 0);
            return (
              <div key={c.id} className="cc-card p-4 cursor-pointer" onClick={() => onOpen(c.id)}>
                <div className="flex items-start justify-between">
                  <div>
                    <div className="cc-h" style={{ fontSize: 17, fontWeight: 600 }}>{c.nombre}</div>
                    <div style={{ fontSize: 12, color: "#8A8570" }}>{nLotes} lote{nLotes !== 1 ? "s" : ""}{hasTotal ? ` · ${fmt(hasTotal, 1)} ha` : ""}</div>
                  </div>
                  {puedeEditar && <button onClick={(e) => { e.stopPropagation(); eliminar(c.id); }}><Trash2 size={16} color="var(--rust)" /></button>}
                </div>
                <div className="flex items-center gap-1 mt-3" style={{ color: "var(--frost)", fontSize: 12.5, fontWeight: 600 }}>Ver lotes <ChevronRight size={16} /></div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Lotes de un campo                                                    */
/* ------------------------------------------------------------------ */
function LotesView({ campo, lotes, api, sinCampo, campos = [], puedeEditar = true }) {
  const [nombre, setNombre] = useState("");
  const [hectareas, setHectareas] = useState("");
  const [editId, setEditId] = useState(null);
  const [editNombre, setEditNombre] = useState("");
  const [editHectareas, setEditHectareas] = useState("");

  const guardar = () => {
    if (!nombre.trim()) return;
    if (hectareas && Number(hectareas) <= 0) { alert("Las hectáreas tienen que ser un número mayor a 0."); return; }
    api.add({ nombre: nombre.trim(), hectareas: hectareas ? Number(hectareas) : null, campoId: campo.id }); setNombre(""); setHectareas("");
  };
  const eliminar = (id) => { if (confirm("Este lote se moverá a la papelera. ¿Continuar?")) api.remove(id); };
  const mover = (id, campoId) => { if (campoId) api.update(id, { campoId }); };

  const empezarEdicion = (l) => { setEditId(l.id); setEditNombre(l.nombre); setEditHectareas(l.hectareas ?? ""); };
  const cancelarEdicion = () => { setEditId(null); setEditNombre(""); setEditHectareas(""); };
  const guardarEdicion = async (id) => {
    if (!editNombre.trim()) { alert("El nombre no puede quedar vacío."); return; }
    if (editHectareas && Number(editHectareas) <= 0) { alert("Las hectáreas tienen que ser un número mayor a 0."); return; }
    await api.update(id, { nombre: editNombre.trim(), hectareas: editHectareas ? Number(editHectareas) : null });
    cancelarEdicion();
  };

  return (
    <div className="space-y-5">
      {sinCampo ? (
        <div className="cc-card p-4" style={{ background: "#FDF3E0", border: "1px solid var(--gold)" }}>
          <div style={{ fontSize: 13, color: "#7A5A1E" }}>Estos lotes son de antes de tener "Campos" y no pertenecen a ninguno. Asignalos a un campo desde el selector de cada fila, o borralos si ya no los usás.</div>
        </div>
      ) : puedeEditar ? (
        <div className="cc-card p-4">
          <div className="cc-h" style={{ fontSize: 16, fontWeight: 600, marginBottom: 10 }}>Nuevo lote en {campo.nombre}</div>
          <div className="flex gap-3 flex-wrap items-end">
            <div style={{ flex: 1, minWidth: 180 }}><label style={{ fontSize: 12, color: "#8A8570" }}>Nombre</label><input className="cc-input" value={nombre} onChange={(e) => setNombre(e.target.value)} placeholder="Lote 4 - La Loma" /></div>
            <div style={{ width: 140 }}><label style={{ fontSize: 12, color: "#8A8570" }}>Hectáreas (opcional)</label><input className="cc-input" type="number" value={hectareas} onChange={(e) => setHectareas(e.target.value)} /></div>
            <button className="cc-btn cc-btn-primary" onClick={guardar}><Plus size={18} /> Agregar</button>
          </div>
        </div>
      ) : null}
      {lotes.length === 0 ? <EmptyState icon={MapPin} title="No hay lotes cargados en este campo" text="Los lotes son de referencia: van a aparecer como sugerencia al cargar el 'Campo' de cada remito." /> : (
        <div className="cc-card overflow-hidden">
          <table className="w-full" style={{ fontSize: 13 }}>
            <thead><tr style={{ background: "#EEEADA", textAlign: "left" }}><th className="px-4 py-2">Lote</th><th className="px-4 py-2">Hectáreas</th>{sinCampo && puedeEditar && <th className="px-4 py-2">Asignar a campo</th>}<th className="px-4 py-2" colSpan={2}></th></tr></thead>
            <tbody>{lotes.map((l) => {
              const enEdicion = editId === l.id;
              return (
              <tr key={l.id} style={{ borderTop: "1px solid var(--line)", background: enEdicion ? "#FDF3E0" : "transparent" }}>
                {enEdicion ? (
                  <>
                    <td className="px-4 py-2"><input className="cc-input" style={{ padding: "5px 8px", fontSize: 12.5 }} value={editNombre} onChange={(e) => setEditNombre(e.target.value)} /></td>
                    <td className="px-4 py-2"><input className="cc-input" style={{ padding: "5px 8px", fontSize: 12.5, width: 100 }} type="number" value={editHectareas} onChange={(e) => setEditHectareas(e.target.value)} /></td>
                    {sinCampo && puedeEditar && <td className="px-4 py-2"></td>}
                    <td className="px-4 py-2 text-right" colSpan={2}>
                      <div className="flex gap-2 justify-end">
                        <button onClick={() => guardarEdicion(l.id)} title="Guardar"><CheckCircle2 size={17} color="var(--soil-light)" /></button>
                        <button onClick={cancelarEdicion} title="Cancelar"><X size={17} color="var(--rust)" /></button>
                      </div>
                    </td>
                  </>
                ) : (
                  <>
                    <td className="px-4 py-2">{l.nombre}</td>
                    <td className="px-4 py-2 cc-mono">{l.hectareas ? `${fmt(l.hectareas, 1)} ha` : "-"}</td>
                    {sinCampo && puedeEditar && (
                      <td className="px-4 py-2">
                        <select className="cc-input" style={{ padding: "5px 8px", fontSize: 12.5 }} defaultValue="" onChange={(e) => mover(l.id, e.target.value)}>
                          <option value="" disabled>Elegir campo…</option>
                          {campos.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}
                        </select>
                      </td>
                    )}
                    <td className="px-4 py-2 text-right">{puedeEditar && <button onClick={() => empezarEdicion(l)} title="Editar"><Pencil size={16} color="var(--frost)" /></button>}</td>
                    <td className="px-4 py-2 text-right">{puedeEditar && <button onClick={() => eliminar(l.id)} title="Eliminar"><Trash2 size={17} color="var(--rust)" /></button>}</td>
                  </>
                )}
              </tr>
            );
            })}</tbody>
          </table>
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Insumos (compras + stock, referencia global)                        */
/* ------------------------------------------------------------------ */
const emptyItemInsumo = () => ({ nombre: "", litros: "", unidad: "Litros", precioUnitario: "" });

function InsumosView({ compras, api, stockInsumos, user, puedeEditar = true, puntosStock, puntosStockApi }) {
  const [fecha, setFecha] = useState("");
  const [origen, setOrigen] = useState("");
  const [socio, setSocio] = useState("");
  const [numeroFactura, setNumeroFactura] = useState("");
  const [puntoStockId, setPuntoStockId] = useState("");
  const [items, setItems] = useState([emptyItemInsumo()]);
  const [archivo, setArchivo] = useState(null);
  const [preview, setPreview] = useState(null);
  const [mediaType, setMediaType] = useState("image/jpeg");
  const [imgB64, setImgB64] = useState(null);
  const [extrayendo, setExtrayendo] = useState(false);
  const [subiendo, setSubiendo] = useState(false);
  const [error, setError] = useState("");
  const [mensaje, setMensaje] = useState("");
  const [busqueda, setBusqueda] = useState("");
  const [nuevoPunto, setNuevoPunto] = useState("");
  const [gestionandoPuntos, setGestionandoPuntos] = useState(false);

  const origenesSugeridos = Array.from(new Set(compras.map((c) => c.origen).filter(Boolean)));
  const nombresSugeridos = Array.from(new Set(compras.map((c) => c.nombre).filter(Boolean)));
  const sociosSugeridos = Array.from(new Set(compras.map((c) => c.socio).filter(Boolean)));

  const setItem = (i, next) => setItems(items.map((it, idx) => (idx === i ? next : it)));
  const agregarItem = () => setItems([...items, emptyItemInsumo()]);
  const quitarItem = (i) => setItems(items.filter((_, idx) => idx !== i));

  const onFile = async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    setArchivo(f);
    setMediaType(f.type || "image/jpeg");
    if (f.type.startsWith("image/")) { setPreview(URL.createObjectURL(f)); setImgB64(await fileToBase64(f)); }
    else { setPreview(null); setImgB64(null); }
  };

  const extraer = async () => {
    if (!imgB64) { setError("La extracción automática funciona con fotos (JPG/PNG) de la factura."); return; }
    setExtrayendo(true); setError("");
    try {
      const data = await askClaudeJSON([
        { type: "image", source: { type: "base64", media_type: mediaType, data: imgB64 } },
        { type: "text", text: `Esta imagen es una factura de compra de insumos agropecuarios (agroquímicos, fertilizantes, combustible, etc.) en Uruguay. Devolvé SOLO un JSON (sin texto extra, sin markdown) con: {"origen":"nombre del proveedor","fecha":"YYYY-MM-DD o vacío","numeroFactura":"número de factura si se ve en la imagen, sino vacío","items":[{"nombre":"nombre del insumo","litros":number (cantidad total, en la unidad que corresponda),"unidad":"Litros, Unidades, Kg o Bolsas — la que más se ajuste al insumo","precio":number (precio TOTAL pagado por esa cantidad, en dólares)}]}. Incluí un objeto en "items" por cada insumo distinto de la factura.` },
      ]);
      setOrigen(data.origen || origen);
      if (data.fecha) setFecha(data.fecha);
      if (data.numeroFactura) setNumeroFactura(data.numeroFactura);
      if (Array.isArray(data.items) && data.items.length) setItems(data.items.map((it) => {
        const litros = it.litros ?? "";
        const unidad = UNIDADES_INSUMO.includes(it.unidad) ? it.unidad : "Litros";
        const precioUnitario = litros && it.precio ? (Number(it.precio) / Number(litros)) : "";
        return { nombre: it.nombre || "", litros, unidad, precioUnitario };
      }));
    } catch (e) { setError("No se pudo leer la factura automáticamente. Completá los ítems a mano."); }
    finally { setExtrayendo(false); }
  };

  const guardar = async () => {
    const validos = items.filter((it) => it.nombre && it.litros && it.precioUnitario);
    if (!fecha || !origen || !validos.length) { alert("Completá fecha, origen y al menos un insumo con nombre, cantidad y precio por unidad."); return; }
    if (validos.some((it) => Number(it.litros) <= 0 || Number(it.precioUnitario) <= 0)) { alert("La cantidad y el precio por unidad de cada insumo tienen que ser números mayores a 0."); return; }
    let facturaUrl = null, facturaNombre = null;
    if (archivo) {
      setSubiendo(true);
      try {
        const path = `insumos/${Date.now()}_${archivo.name}`;
        const r = ref(storage, path);
        await uploadBytes(r, archivo);
        facturaUrl = await getDownloadURL(r);
        facturaNombre = archivo.name;
      } catch (e) { alert("No se pudo subir la factura adjunta, pero se guardará la compra sin el archivo."); }
      setSubiendo(false);
    }
    const puntoNombre = puntosStock.find((p) => p.id === puntoStockId)?.nombre || "";
    await Promise.all(validos.map((it) => api.add({
      fecha, origen, nombre: it.nombre, litros: Number(it.litros), unidad: it.unidad || "Litros",
      precio: Number(it.litros) * Number(it.precioUnitario), precioUnitario: Number(it.precioUnitario),
      facturaUrl, facturaNombre, numeroFactura: numeroFactura || "", usuario: user.email, socio: socio || "",
      puntoStockId: puntoStockId || null, puntoStockNombre: puntoNombre || "Sin punto asignado",
    })));
    setFecha(""); setOrigen(""); setSocio(""); setNumeroFactura(""); setItems([emptyItemInsumo()]); setArchivo(null); setPreview(null); setImgB64(null);
    setMensaje("Compra guardada ✓"); setTimeout(() => setMensaje(""), 2500);
  };

  const eliminar = (id) => { if (confirm("Esta compra se moverá a la papelera (se descontará del stock hasta que la restaures). ¿Continuar?")) api.remove(id); };

  return (
    <div className="space-y-6">
      <div>
        <div className="cc-h" style={{ fontSize: 20, fontWeight: 600 }}>Insumos</div>
        <div style={{ fontSize: 12.5, color: "#8A8570" }}>Registrá las compras de insumos. El stock se descuenta solo cuando lo consumís desde "Gastos" en cada cultivo.</div>
      </div>

      <div className="cc-card p-4 space-y-4" style={{ display: puedeEditar ? undefined : "none" }}>
        <div className="flex flex-wrap gap-2">
          <label className="cc-btn cc-btn-ghost" style={{ cursor: "pointer" }}>
            <Paperclip size={17} /> {archivo ? archivo.name : "Adjuntar factura (imagen o PDF)"}
            <input type="file" accept="image/*,.pdf" capture="environment" onChange={onFile} style={{ display: "none" }} />
          </label>
          {imgB64 && <button className="cc-btn cc-btn-ghost" onClick={extraer} disabled={extrayendo}>{extrayendo ? <Loader2 size={17} className="animate-spin" /> : <Camera size={17} />} Extraer datos con IA</button>}
        </div>
        {preview && <img src={preview} alt="Factura" style={{ maxWidth: 140, borderRadius: 8, border: "1px solid var(--line)" }} />}
        {error && <div style={{ color: "var(--rust)", fontSize: 13, display: "flex", gap: 6, alignItems: "center" }}><AlertCircle size={17} />{error}</div>}

        <div className="grid gap-3" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(160px,1fr))" }}>
          <div><label style={{ fontSize: 12, color: "#8A8570" }}>Fecha</label><input className="cc-input" type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} /></div>
          <div>
            <label style={{ fontSize: 12, color: "#8A8570" }}>Origen (proveedor)</label>
            <input className="cc-input" list="origenes-insumos" value={origen} onChange={(e) => setOrigen(e.target.value)} placeholder="Agromotora, Barraca..." />
            <datalist id="origenes-insumos">{origenesSugeridos.map((o) => <option key={o} value={o} />)}</datalist>
          </div>
          <div>
            <label style={{ fontSize: 12, color: "#8A8570" }}>Socio que aporta (opcional)</label>
            <input className="cc-input" list="socios-insumos" value={socio} onChange={(e) => setSocio(e.target.value)} placeholder="Ej: Juan Pérez" />
            <datalist id="socios-insumos">{sociosSugeridos.map((s) => <option key={s} value={s} />)}</datalist>
          </div>
          <div>
            <label style={{ fontSize: 12, color: "#8A8570" }}>N° de factura (opcional)</label>
            <input className="cc-input" value={numeroFactura} onChange={(e) => setNumeroFactura(e.target.value)} placeholder="Ej: A-0001234" />
          </div>
          <div>
            <label style={{ fontSize: 12, color: "#8A8570" }}>Punto de stock</label>
            <select className="cc-input" value={puntoStockId} onChange={(e) => setPuntoStockId(e.target.value)}>
              <option value="">Sin punto asignado</option>
              {puntosStock.map((p) => <option key={p.id} value={p.id}>{p.nombre}</option>)}
            </select>
          </div>
        </div>

        <div>
          <button className="cc-btn cc-btn-ghost" onClick={() => setGestionandoPuntos((v) => !v)}><Boxes size={17} /> {gestionandoPuntos ? "Ocultar" : "Gestionar"} puntos de stock</button>
          {gestionandoPuntos && (
            <div className="mt-3 p-3" style={{ background: "#F7F5EC", borderRadius: 8, border: "1px solid var(--line)" }}>
              <div className="flex gap-2 flex-wrap items-end mb-3">
                <div style={{ flex: 1, minWidth: 160 }}>
                  <label style={{ fontSize: 12, color: "#8A8570" }}>Nombre del punto (ej: Depósito Central, Silo Norte)</label>
                  <input className="cc-input" value={nuevoPunto} onChange={(e) => setNuevoPunto(e.target.value)} />
                </div>
                <button className="cc-btn cc-btn-primary" onClick={async () => { if (!nuevoPunto.trim()) return; await puntosStockApi.add({ nombre: nuevoPunto.trim() }); setNuevoPunto(""); }}><Plus size={17} /> Agregar</button>
              </div>
              {puntosStock.length === 0 ? (
                <div style={{ fontSize: 13, color: "#8A8570" }}>Todavía no hay puntos de stock cargados.</div>
              ) : (
                <div className="flex flex-wrap gap-2">
                  {puntosStock.map((p) => (
                    <span key={p.id} className="cc-chip" style={{ background: "#EEEADA", display: "flex", alignItems: "center", gap: 6, padding: "6px 10px" }}>
                      {p.nombre}
                      <button onClick={() => { if (confirm(`¿Eliminar el punto "${p.nombre}"? Se moverá a la papelera.`)) puntosStockApi.remove(p.id); }}><X size={13} color="var(--rust)" /></button>
                    </span>
                  ))}
                </div>
              )}
            </div>
          )}
        </div>

        <div>
          <div style={{ fontSize: 12, color: "#8A8570", marginBottom: 4 }}>Insumos de esta factura</div>
          <div className="space-y-2">
            {items.map((it, i) => {
              const total = it.litros && it.precioUnitario ? Number(it.litros) * Number(it.precioUnitario) : null;
              return (
                <div key={i} className="flex gap-2 items-end flex-wrap">
                  <div style={{ flex: 2, minWidth: 160 }}>
                    <input className="cc-input" list="nombres-insumos" placeholder="Nombre (ej: Glifosato)" value={it.nombre} onChange={(e) => setItem(i, { ...it, nombre: e.target.value })} />
                  </div>
                  <div style={{ width: 110 }}><input className="cc-input" type="number" placeholder="Cantidad" value={it.litros} onChange={(e) => setItem(i, { ...it, litros: e.target.value })} /></div>
                  <div style={{ width: 110 }}>
                    <select className="cc-input" value={it.unidad || "Litros"} onChange={(e) => setItem(i, { ...it, unidad: e.target.value })}>
                      {UNIDADES_INSUMO.map((u) => <option key={u} value={u}>{u}</option>)}
                    </select>
                  </div>
                  <div style={{ width: 150 }}><input className="cc-input" type="number" placeholder={`U$S por ${abrevUnidad(it.unidad)}`} value={it.precioUnitario} onChange={(e) => setItem(i, { ...it, precioUnitario: e.target.value })} /></div>
                  {total !== null && <div style={{ fontSize: 12.5, color: "#8A8570", minWidth: 110 }}>Total: <b style={{ color: "var(--ink)" }}>{fmtUSD(total)}</b></div>}
                  {items.length > 1 && <button onClick={() => quitarItem(i)}><X size={19} color="var(--rust)" /></button>}
                </div>
              );
            })}
            <datalist id="nombres-insumos">{nombresSugeridos.map((n) => <option key={n} value={n} />)}</datalist>
          </div>
          <button className="cc-btn cc-btn-ghost mt-2" onClick={agregarItem}><Plus size={17} /> Agregar otro insumo a esta factura</button>
        </div>

        <div className="flex items-center gap-3">
          <button className="cc-btn cc-btn-primary" onClick={guardar} disabled={subiendo}>{subiendo ? <Loader2 size={18} className="animate-spin" /> : <Plus size={18} />} Guardar compra</button>
          {mensaje && <span style={{ color: "var(--soil-light)", fontWeight: 700, fontSize: 13.5 }}>{mensaje}</span>}
        </div>
      </div>

      <div>
        <div className="cc-h" style={{ fontSize: 16, fontWeight: 600, marginBottom: 10 }}>Stock actual</div>

        {stockInsumos.some((i) => i.disponible <= 0) && (
          <div className="flex items-start gap-2 mb-3" style={{ background: "#FBEAEA", border: "1px solid var(--rust)", borderRadius: 8, padding: "10px 12px" }}>
            <AlertCircle size={16} color="var(--rust)" style={{ marginTop: 1, flexShrink: 0 }} />
            <div style={{ fontSize: 13, color: "#7A2E2E" }}>
              <b>Stock agotado o negativo:</b> {stockInsumos.filter((i) => i.disponible <= 0).map((i) => i.nombre).join(", ")}. Revisá si falta cargar alguna compra.
            </div>
          </div>
        )}

        {stockInsumos.length === 0 ? <EmptyState icon={Boxes} title="Sin insumos cargados" text="Cuando registres una compra, el stock disponible va a aparecer acá." /> : (
          <div className="cc-card overflow-hidden">
            <table className="w-full" style={{ fontSize: 13 }}>
              <thead><tr style={{ background: "#EEEADA", textAlign: "left" }}><th className="px-3 py-2">Insumo</th><th className="px-3 py-2">Punto de stock</th><th className="px-3 py-2 text-right">Comprado</th><th className="px-3 py-2 text-right">Consumido</th><th className="px-3 py-2 text-right">Disponible</th><th className="px-3 py-2 text-right">Costo prom. / unidad</th></tr></thead>
              <tbody>
                {[...stockInsumos].sort((a, b) => a.nombre.localeCompare(b.nombre) || a.puntoStockNombre.localeCompare(b.puntoStockNombre)).map((i) => (
                  <tr key={`${i.nombre}||${i.puntoStockId || "sin_punto"}`} style={{ borderTop: "1px solid var(--line)", background: i.disponible <= 0 ? "#FBEAEA" : "transparent" }}>
                    <td className="px-3 py-2">{i.nombre} {i.disponible < 0 && <AlertCircle size={13} color="var(--rust)" style={{ display: "inline", marginLeft: 4, verticalAlign: "-2px" }} />}</td>
                    <td className="px-3 py-2" style={{ color: "#5A5647" }}>{i.puntoStockNombre}</td>
                    <td className="px-3 py-2 text-right cc-mono">{fmt(i.litrosComprados, 1)} {abrevUnidad(i.unidad)}</td>
                    <td className="px-3 py-2 text-right cc-mono">{fmt(i.litrosConsumidos, 1)} {abrevUnidad(i.unidad)}</td>
                    <td className="px-3 py-2 text-right cc-mono" style={{ color: i.disponible <= 0 ? "var(--rust)" : "var(--soil-light)", fontWeight: 700 }}>{fmt(i.disponible, 1)} {abrevUnidad(i.unidad)}</td>
                    <td className="px-3 py-2 text-right cc-mono">{fmtUSD(i.costoPromedioPorLitro)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      <div>
        <div className="flex items-center justify-between mb-2 flex-wrap gap-2">
          <div className="cc-h" style={{ fontSize: 16, fontWeight: 600 }}>Historial de compras</div>
          <div className="flex items-center gap-2">
            {compras.length > 0 && <input className="cc-input" style={{ maxWidth: 280 }} value={busqueda} onChange={(e) => setBusqueda(e.target.value)} placeholder="Buscar por insumo u origen..." />}
            {compras.length > 0 && (
              <button className="cc-btn cc-btn-ghost" onClick={() => exportarExcel("compras_insumos", [{ nombre: "Compras", filas: compras.map((c) => ({ Fecha: c.fecha, Origen: c.origen, "N° factura": c.numeroFactura || "", Socio: c.socio || "", "Punto de stock": c.puntoStockNombre || "Sin punto asignado", Insumo: c.nombre, Cantidad: c.litros, Unidad: c.unidad || "Litros", "Precio por unidad": c.precioUnitario ?? (c.litros ? c.precio / c.litros : ""), "Precio total": c.precio })) }, { nombre: "Stock", filas: stockInsumos.map((i) => ({ Insumo: i.nombre, "Punto de stock": i.puntoStockNombre, Unidad: i.unidad, Comprado: i.litrosComprados, Consumido: i.litrosConsumidos, Disponible: i.disponible, "Costo prom/unidad": i.costoPromedioPorLitro })) }])}>
                <Download size={17} /> Exportar Excel
              </button>
            )}
          </div>
        </div>
        {compras.length === 0 ? <EmptyState icon={Receipt} title="Sin compras registradas" text="Registrá tu primera compra de insumos arriba." /> : (
          <div className="cc-card overflow-hidden">
            <table className="w-full" style={{ fontSize: 12.5 }}>
              <thead><tr style={{ background: "#EEEADA", textAlign: "left" }}><th className="px-3 py-2">Fecha</th><th className="px-3 py-2">Origen</th><th className="px-3 py-2">N° factura</th><th className="px-3 py-2">Socio</th><th className="px-3 py-2">Punto de stock</th><th className="px-3 py-2">Insumo</th><th className="px-3 py-2 text-right">Litros</th><th className="px-3 py-2 text-right">Precio</th><th className="px-3 py-2"></th><th className="px-3 py-2"></th></tr></thead>
              <tbody>
                {[...compras]
                  .filter((c) => {
                    const q = busqueda.trim().toLowerCase();
                    if (!q) return true;
                    return [c.nombre, c.origen, c.numeroFactura].some((v) => (v || "").toLowerCase().includes(q));
                  })
                  .sort((a, b) => (b.fecha || "").localeCompare(a.fecha || "")).map((c) => (
                  <tr key={c.id} style={{ borderTop: "1px solid var(--line)" }}>
                    <td className="px-3 py-2 cc-mono">{c.fecha}</td>
                    <td className="px-3 py-2">{c.origen}</td>
                    <td className="px-3 py-2 cc-mono" style={{ color: "#5A5647" }}>{c.numeroFactura || "-"}</td>
                    <td className="px-3 py-2" style={{ color: "#5A5647" }}>{c.socio || "-"}</td>
                    <td className="px-3 py-2" style={{ color: "#5A5647" }}>{c.puntoStockNombre || "Sin punto asignado"}</td>
                    <td className="px-3 py-2">{c.nombre}</td>
                    <td className="px-3 py-2 text-right cc-mono">{fmt(c.litros, 1)} {abrevUnidad(c.unidad)}</td>
                    <td className="px-3 py-2 text-right cc-mono">{fmtUSD(c.precio)}</td>
                    <td className="px-3 py-2">{c.facturaUrl && <a href={c.facturaUrl} target="_blank" rel="noreferrer"><FileText size={17} color="var(--frost)" /></a>}</td>
                    <td className="px-3 py-2 text-right">{puedeEditar && <button onClick={() => eliminar(c.id)}><Trash2 size={16} color="var(--rust)" /></button>}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {compras.length > 0 && (
        <div>
          <div className="cc-h" style={{ fontSize: 16, fontWeight: 600, marginBottom: 10 }}>Aporte por socio (insumos)</div>
          <div className="cc-card p-4 space-y-3">
            {Object.entries(
              compras.reduce((acc, c) => {
                const s = c.socio && c.socio.trim() ? c.socio.trim() : "Sin asignar";
                acc[s] = (acc[s] || 0) + Number(c.precio || 0);
                return acc;
              }, {})
            )
              .sort((a, b) => b[1] - a[1])
              .map(([s, monto]) => {
                const totalCompras = compras.reduce((sum, c) => sum + Number(c.precio || 0), 0);
                const pct = totalCompras ? (monto / totalCompras) * 100 : 0;
                const color = colorPorTexto(s);
                return (
                  <div key={s}>
                    <div className="flex items-center justify-between" style={{ fontSize: 13 }}>
                      <span style={{ fontWeight: 600, color: s === "Sin asignar" ? "#8A8570" : "#5A5647" }}>{s}</span>
                      <span className="cc-mono">{fmtUSD(monto)} <span style={{ color: "#8A8570" }}>({fmt(pct, 1)}%)</span></span>
                    </div>
                    <div style={{ background: "#EEEADA", borderRadius: 99, height: 8, marginTop: 4, overflow: "hidden" }}>
                      <div style={{ width: `${Math.min(pct, 100)}%`, background: color, height: "100%", borderRadius: 99 }} />
                    </div>
                  </div>
                );
              })}
          </div>
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Resumen general (todas las campañas y cultivos)                     */
/* ------------------------------------------------------------------ */
function ResumenGeneralView({ campanias, cultivos, gastos, ventas, remitos, lotes, insumosCompras, onOpenCultivo }) {
  const filas = cultivos.map((c) => {
    const campania = campanias.find((cp) => cp.id === c.campaniaId);
    const gastosCultivo = gastos.filter((g) => g.cultivoId === c.id);
    const ventasCultivo = ventas.filter((v) => v.cultivoId === c.id);
    const remitosCultivo = remitos.filter((r) => r.cultivoId === c.id);
    const totalGastos = gastosCultivo.reduce((s, g) => s + Number(g.monto || 0), 0);
    const totalIngresos = ventasCultivo.reduce((s, v) => s + Number(v.toneladas || 0) * Number(v.dolaresPorTonelada || 0), 0);
    const totalTon = remitosCultivo.reduce((s, r) => s + Number(r.kgSL || 0), 0) / 1000;
    const superficie = lotes.filter((l) => (c.loteIds || []).includes(l.id)).reduce((s, l) => s + Number(l.hectareas || 0), 0);
    const porSocio = {};
    gastosCultivo.forEach((g) => {
      const s = g.socio && g.socio.trim() ? g.socio.trim() : "Sin asignar";
      porSocio[s] = (porSocio[s] || 0) + Number(g.monto || 0);
    });
    return {
      cultivoId: c.id, campaniaId: c.campaniaId, cultivoNombre: c.nombre, tipoCultivo: c.tipo || c.nombre, categoria: c.categoria,
      campaniaNombre: campania ? (campania.nombre || campania.anio) : "—", anio: campania?.anio || 0,
      superficie, totalGastos, totalIngresos, margen: totalIngresos - totalGastos, rendimiento: superficie && totalTon ? totalTon / superficie : null,
      costoPorHa: superficie ? totalGastos / superficie : null, margenPorHa: superficie ? (totalIngresos - totalGastos) / superficie : null,
      porSocio,
    };
  }).sort((a, b) => b.anio - a.anio || a.cultivoNombre.localeCompare(b.cultivoNombre));

  const totalGastosGeneral = filas.reduce((s, f) => s + f.totalGastos, 0);
  const totalIngresosGeneral = filas.reduce((s, f) => s + f.totalIngresos, 0);
  const margenGeneral = totalIngresosGeneral - totalGastosGeneral;

  const porCultivoNombre = {};
  filas.forEach((f) => { (porCultivoNombre[f.tipoCultivo] ||= []).push(f); });
  const comparables = Object.entries(porCultivoNombre).filter(([, arr]) => arr.length > 1);

  const cultivoIdsActivos = new Set(cultivos.map((c) => c.id));
  const gastosActivos = gastos.filter((g) => cultivoIdsActivos.has(g.cultivoId));
  const aportesSocio = {};
  gastosActivos.forEach((g) => {
    const s = g.socio && g.socio.trim() ? g.socio.trim() : "Sin asignar";
    aportesSocio[s] = (aportesSocio[s] || 0) + Number(g.monto || 0);
  });
  insumosCompras.forEach((c) => {
    const s = c.socio && c.socio.trim() ? c.socio.trim() : "Sin asignar";
    aportesSocio[s] = (aportesSocio[s] || 0) + Number(c.precio || 0);
  });
  const totalAportes = Object.values(aportesSocio).reduce((s, v) => s + v, 0);

  const sociosGlobalExport = Array.from(new Set(gastosActivos.map((g) => (g.socio && g.socio.trim()) ? g.socio.trim() : "Sin asignar")));
  const exportar = () => exportarExcel("resumen_general_campo_costo", [
    { nombre: "Resumen", filas: filas.map((f) => ({ Campaña: f.campaniaNombre, Cultivo: f.cultivoNombre, Categoría: f.categoria, "Superficie (ha)": f.superficie, "Rinde (tn/ha)": f.rendimiento, "Total gastos": f.totalGastos, "Total ingresos": f.totalIngresos, Margen: f.margen, "Gasto/ha": f.costoPorHa, "Margen/ha": f.margenPorHa })) },
    { nombre: "Aporte por socio por cultivo", filas: filas.map((f) => { const fila = { Campaña: f.campaniaNombre, Cultivo: f.cultivoNombre }; sociosGlobalExport.forEach((s) => { fila[s] = f.porSocio[s] || 0; }); fila["Total"] = f.totalGastos; return fila; }) },
    { nombre: "Aporte por socio", filas: Object.entries(aportesSocio).map(([socio, monto]) => ({ Socio: socio, Monto: monto, "% del total": totalAportes ? (monto / totalAportes) * 100 : 0 })) },
  ]);

  if (!cultivos.length) {
    return <EmptyState icon={LayoutDashboard} title="Todavía no hay datos para resumir" text="Cuando crees campañas y cultivos con gastos e ingresos, acá vas a ver el resumen de todo junto." />;
  }

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between flex-wrap gap-2">
        <div>
          <div className="cc-h" style={{ fontSize: 20, fontWeight: 600 }}>Resumen general</div>
          <div style={{ fontSize: 12.5, color: "#8A8570" }}>Todas las campañas y cultivos juntos, de un vistazo.</div>
        </div>
        <button className="cc-btn cc-btn-ghost" onClick={exportar}><Download size={17} /> Exportar Excel</button>
      </div>

      <div className="grid gap-4" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(200px,1fr))" }}>
        <StatCard label="Total gastos (todo)" value={fmtUSD(totalGastosGeneral)} icon={Receipt} color="var(--rust)" />
        <StatCard label="Total ingresos (todo)" value={fmtUSD(totalIngresosGeneral)} icon={DollarSign} color="var(--soil-light)" />
        <StatCard label="Margen (todo)" value={fmtUSD(margenGeneral)} icon={margenGeneral >= 0 ? TrendingUp : TrendingDown} color={margenGeneral >= 0 ? "var(--soil-light)" : "var(--rust)"} />
      </div>

      <div className="grid gap-4" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(320px,1fr))" }}>
        <div className="cc-card p-4">
          <div className="cc-h" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Gastos vs. ingresos por cultivo</div>
          <div style={{ width: "100%", overflowX: filas.length > 8 ? "auto" : "visible" }}>
            <div style={{ minWidth: filas.length > 8 ? filas.length * 70 : "100%", height: 260 }}>
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={filas.map((f) => ({ nombre: `${f.cultivoNombre} · ${f.campaniaNombre}`, Gastos: Math.round(f.totalGastos), Ingresos: Math.round(f.totalIngresos) }))} margin={{ top: 4, right: 8, left: 0, bottom: 40 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="var(--line)" />
                  <XAxis dataKey="nombre" angle={-35} textAnchor="end" interval={0} height={70} tick={{ fontSize: 10.5, fill: "#8A8570" }} />
                  <YAxis tick={{ fontSize: 10.5, fill: "#8A8570" }} />
                  <Tooltip formatter={(v) => fmtUSD(v)} contentStyle={{ fontSize: 12.5, borderRadius: 8 }} />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                  <Bar dataKey="Gastos" fill="#8C3D3D" radius={[4, 4, 0, 0]} />
                  <Bar dataKey="Ingresos" fill="#3E4F32" radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>

        <div className="cc-card p-4">
          <div className="cc-h" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Gastos por categoría</div>
          {(() => {
            const porCategoria = {};
            gastosActivos.forEach((g) => {
              const cat = g.categoriaGasto || (g.insumoNombre ? "Insumo" : "Otro");
              porCategoria[cat] = (porCategoria[cat] || 0) + Number(g.monto || 0);
            });
            const dataCat = Object.entries(porCategoria).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
            const colores = ["#8C3D3D", "#C68A2E", "#3D6E8C", "#5C7A4E", "#5B4B8A", "#8A6D3B", "#7A7267"];
            if (!dataCat.length) return <div style={{ fontSize: 13, color: "#8A8570" }}>Todavía no hay gastos cargados.</div>;
            return (
              <div style={{ height: 260 }}>
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie data={dataCat} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={85} label={({ name, percent }) => `${name} ${(percent * 100).toFixed(0)}%`} labelLine={false} style={{ fontSize: 11 }}>
                      {dataCat.map((_, i) => <Cell key={i} fill={colores[i % colores.length]} />)}
                    </Pie>
                    <Tooltip formatter={(v) => fmtUSD(v)} contentStyle={{ fontSize: 12.5, borderRadius: 8 }} />
                  </PieChart>
                </ResponsiveContainer>
              </div>
            );
          })()}
        </div>
      </div>

      <div className="cc-card overflow-hidden">
        <table className="w-full" style={{ fontSize: 13 }}>
          <thead><tr style={{ background: "#EEEADA", textAlign: "left" }}>
            <th className="px-3 py-2">Campaña</th><th className="px-3 py-2">Cultivo</th><th className="px-3 py-2 text-right">Ha</th>
            <th className="px-3 py-2 text-right">Rinde tn/ha</th>
            <th className="px-3 py-2 text-right">Gastos</th><th className="px-3 py-2 text-right">Ingresos</th><th className="px-3 py-2 text-right">Margen</th>
          </tr></thead>
          <tbody>
            {filas.map((f) => (
              <tr key={f.cultivoId} style={{ borderTop: "1px solid var(--line)", cursor: "pointer" }} onClick={() => onOpenCultivo(f.cultivoId, f.campaniaId)}>
                <td className="px-3 py-2">{f.campaniaNombre}</td>
                <td className="px-3 py-2">
                  <span className="cc-chip" style={{ background: CAT_COLOR[f.categoria] + "22", color: CAT_COLOR[f.categoria], marginRight: 6 }}>{f.categoria === "verano" ? "V" : "I"}</span>
                  {f.cultivoNombre}
                </td>
                <td className="px-3 py-2 text-right cc-mono">{f.superficie ? fmt(f.superficie, 1) : "-"}</td>
                <td className="px-3 py-2 text-right cc-mono">{f.rendimiento !== null ? fmt(f.rendimiento, 2) : "-"}</td>
                <td className="px-3 py-2 text-right cc-mono">{fmtUSD(f.totalGastos)}</td>
                <td className="px-3 py-2 text-right cc-mono">{fmtUSD(f.totalIngresos)}</td>
                <td className="px-3 py-2 text-right cc-mono" style={{ color: f.margen >= 0 ? "var(--soil-light)" : "var(--rust)", fontWeight: 700 }}>{fmtUSD(f.margen)}</td>
              </tr>
            ))}
          </tbody>
          <tfoot><tr style={{ borderTop: "2px solid var(--line)", fontWeight: 700 }}>
            <td className="px-3 py-2" colSpan={4}>Total</td>
            <td className="px-3 py-2 text-right cc-mono">{fmtUSD(totalGastosGeneral)}</td>
            <td className="px-3 py-2 text-right cc-mono">{fmtUSD(totalIngresosGeneral)}</td>
            <td className="px-3 py-2 text-right cc-mono">{fmtUSD(margenGeneral)}</td>
          </tr></tfoot>
        </table>
      </div>

      {comparables.length > 0 && (
        <div>
          <div className="cc-h" style={{ fontSize: 17, fontWeight: 600, marginBottom: 4 }}>Comparar entre campañas</div>
          <div style={{ fontSize: 12.5, color: "#8A8570", marginBottom: 10 }}>El mismo cultivo, año a año.</div>
          <div className="space-y-4">
            {comparables.map(([nombre, arr]) => (
              <div key={nombre} className="cc-card overflow-hidden">
                <div className="px-3 py-2" style={{ background: "#EEEADA", fontWeight: 600, fontSize: 13.5 }}>{nombre}</div>
                <table className="w-full" style={{ fontSize: 13 }}>
                  <thead><tr style={{ textAlign: "left", color: "#8A8570" }}>
                    <th className="px-3 py-2">Campaña</th><th className="px-3 py-2 text-right">Ha</th><th className="px-3 py-2 text-right">Rinde tn/ha</th>
                    <th className="px-3 py-2 text-right">Gasto/ha</th><th className="px-3 py-2 text-right">Margen/ha</th><th className="px-3 py-2 text-right">Margen total</th>
                  </tr></thead>
                  <tbody>
                    {[...arr].sort((a, b) => b.anio - a.anio).map((f) => (
                      <tr key={f.cultivoId} style={{ borderTop: "1px solid var(--line)", cursor: "pointer" }} onClick={() => onOpenCultivo(f.cultivoId, f.campaniaId)}>
                        <td className="px-3 py-2">{f.campaniaNombre}</td>
                        <td className="px-3 py-2 text-right cc-mono">{f.superficie ? fmt(f.superficie, 1) : "-"}</td>
                        <td className="px-3 py-2 text-right cc-mono">{f.rendimiento !== null ? fmt(f.rendimiento, 2) : "-"}</td>
                        <td className="px-3 py-2 text-right cc-mono">{f.costoPorHa !== null ? fmtUSD(f.costoPorHa) : "-"}</td>
                        <td className="px-3 py-2 text-right cc-mono" style={{ color: f.margenPorHa >= 0 ? "var(--soil-light)" : "var(--rust)" }}>{f.margenPorHa !== null ? fmtUSD(f.margenPorHa) : "-"}</td>
                        <td className="px-3 py-2 text-right cc-mono" style={{ fontWeight: 700, color: f.margen >= 0 ? "var(--soil-light)" : "var(--rust)" }}>{fmtUSD(f.margen)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ))}
          </div>
        </div>
      )}

      {(() => {
        const sociosGlobal = Array.from(new Set(gastosActivos.map((g) => (g.socio && g.socio.trim()) ? g.socio.trim() : "Sin asignar")))
          .sort((a, b) => (a === "Sin asignar" ? 1 : b === "Sin asignar" ? -1 : a.localeCompare(b)));
        if (!sociosGlobal.length) return null;
        return (
          <div>
            <div className="cc-h" style={{ fontSize: 17, fontWeight: 600, marginBottom: 4 }}>Aporte por socio, por cultivo</div>
            <div style={{ fontSize: 12.5, color: "#8A8570", marginBottom: 10 }}>Cuánto puso cada socio, discriminado campaña por campaña.</div>
            <div className="cc-card overflow-hidden">
              <table className="w-full" style={{ fontSize: 13 }}>
                <thead><tr style={{ background: "#EEEADA", textAlign: "left" }}>
                  <th className="px-3 py-2">Campaña</th><th className="px-3 py-2">Cultivo</th>
                  {sociosGlobal.map((s) => <th key={s} className="px-3 py-2 text-right">{s}</th>)}
                  <th className="px-3 py-2 text-right">Total</th>
                </tr></thead>
                <tbody>
                  {filas.filter((f) => f.totalGastos > 0).map((f) => (
                    <tr key={f.cultivoId} style={{ borderTop: "1px solid var(--line)", cursor: "pointer" }} onClick={() => onOpenCultivo(f.cultivoId, f.campaniaId)}>
                      <td className="px-3 py-2">{f.campaniaNombre}</td>
                      <td className="px-3 py-2">{f.cultivoNombre}</td>
                      {sociosGlobal.map((s) => (
                        <td key={s} className="px-3 py-2 text-right cc-mono" style={{ color: f.porSocio[s] ? "var(--ink)" : "#C9C3AC" }}>
                          {f.porSocio[s] ? fmtUSD(f.porSocio[s]) : "-"}
                        </td>
                      ))}
                      <td className="px-3 py-2 text-right cc-mono" style={{ fontWeight: 700 }}>{fmtUSD(f.totalGastos)}</td>
                    </tr>
                  ))}
                </tbody>
                <tfoot><tr style={{ borderTop: "2px solid var(--line)", fontWeight: 700 }}>
                  <td className="px-3 py-2" colSpan={2}>Total</td>
                  {sociosGlobal.map((s) => <td key={s} className="px-3 py-2 text-right cc-mono">{fmtUSD(aportesSocio[s] || 0)}</td>)}
                  <td className="px-3 py-2 text-right cc-mono">{fmtUSD(totalGastosGeneral)}</td>
                </tr></tfoot>
              </table>
            </div>
          </div>
        );
      })()}

      {totalAportes > 0 && (
        <div>
          <div className="cc-h" style={{ fontSize: 17, fontWeight: 600, marginBottom: 10 }}>Aporte por socio (todo junto, incluye insumos comprados sin consumir)</div>
          <div className="cc-card p-4 space-y-3">
            {Object.entries(aportesSocio).sort((a, b) => b[1] - a[1]).map(([socio, monto]) => {
              const pct = totalAportes ? (monto / totalAportes) * 100 : 0;
              const color = colorPorTexto(socio);
              return (
                <div key={socio}>
                  <div className="flex items-center justify-between" style={{ fontSize: 13 }}>
                    <span style={{ fontWeight: 600, color: socio === "Sin asignar" ? "#8A8570" : "#5A5647" }}>{socio}</span>
                    <span className="cc-mono">{fmtUSD(monto)} <span style={{ color: "#8A8570" }}>({fmt(pct, 1)}%)</span></span>
                  </div>
                  <div style={{ background: "#EEEADA", borderRadius: 99, height: 8, marginTop: 4, overflow: "hidden" }}>
                    <div style={{ width: `${Math.min(pct, 100)}%`, background: color, height: "100%", borderRadius: 99 }} />
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Papelera                                                            */
/* ------------------------------------------------------------------ */
function PapeleraView({ grupos, puedeEditar = true }) {
  const hayAlgo = grupos.some((g) => g.items.length > 0);
  if (!hayAlgo) {
    return <EmptyState icon={Trash} title="La papelera está vacía" text="Todo lo que borres desde cualquier parte de la app va a aparecer acá antes de eliminarse para siempre." />;
  }
  return (
    <div className="space-y-6">
      <div>
        <div className="cc-h" style={{ fontSize: 20, fontWeight: 600 }}>Papelera</div>
        <div style={{ fontSize: 12.5, color: "#8A8570" }}>Lo que borres queda acá. Podés restaurarlo o eliminarlo para siempre.</div>
      </div>
      {grupos.filter((g) => g.items.length > 0).map((g) => (
        <div key={g.titulo}>
          <div className="cc-h" style={{ fontSize: 15, fontWeight: 600, marginBottom: 8 }}>{g.titulo} ({g.items.length})</div>
          <div className="cc-card overflow-hidden">
            <table className="w-full" style={{ fontSize: 13 }}>
              <thead><tr style={{ background: "#EEEADA", textAlign: "left" }}><th className="px-3 py-2">Detalle</th><th className="px-3 py-2">Eliminado por</th><th className="px-3 py-2">Cuándo</th><th className="px-3 py-2"></th></tr></thead>
              <tbody>
                {g.items.map((it) => (
                  <tr key={it.id} style={{ borderTop: "1px solid var(--line)" }}>
                    <td className="px-3 py-2">{g.campo(it)}</td>
                    <td className="px-3 py-2" style={{ color: "#8A8570", fontSize: 12 }}>{it.eliminadoPor || "-"}</td>
                    <td className="px-3 py-2 cc-mono" style={{ fontSize: 12 }}>{it.eliminadoEn ? new Date(it.eliminadoEn).toLocaleString("es-UY") : "-"}</td>
                    <td className="px-3 py-2 text-right">
                      {puedeEditar ? (
                        <div className="flex gap-2 justify-end">
                          <button className="cc-btn cc-btn-ghost" style={{ padding: "5px 10px", fontSize: 12 }} onClick={() => g.api.restaurar(it.id)}><RotateCcw size={15} /> Restaurar</button>
                          <button className="cc-btn" style={{ padding: "5px 10px", fontSize: 12, background: "var(--rust)", color: "#fff" }}
                            onClick={() => { if (confirm("Esto lo borra para siempre, sin poder recuperarlo. ¿Continuar?")) g.api.eliminarDefinitivo(it.id); }}>
                            <Trash2 size={15} /> Eliminar para siempre
                          </button>
                        </div>
                      ) : null}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      ))}
    </div>
  );
}
