import React, { useState, useEffect, useRef } from "react";
import { auth, db, storage } from "./firebase";
import {
  onAuthStateChanged, signInWithEmailAndPassword, createUserWithEmailAndPassword,
  signOut, sendPasswordResetEmail,
} from "firebase/auth";
import {
  collection, onSnapshot, addDoc, updateDoc, deleteDoc, doc, query, where,
} from "firebase/firestore";
import { ref, uploadBytes, getDownloadURL } from "firebase/storage";
import {
  Sprout, Snowflake, Sun, Wheat, Receipt, Mic, Camera, MapPin, TrendingUp,
  TrendingDown, Plus, Trash2, Loader2, LogOut, ChevronRight, ChevronLeft,
  Truck, DollarSign, FileText, AlertCircle, CheckCircle2, Paperclip, Pencil, X, Package, Boxes,
} from "lucide-react";

/* ------------------------------------------------------------------ */
/*  Constantes de dominio                                              */
/* ------------------------------------------------------------------ */
const CULTIVOS_VERANO = ["Soja", "Maíz", "Sorgo"];
const CULTIVOS_INVIERNO = ["Colza", "Carinata", "Trigo", "Cebada", "Lupino", "Camelina"];
const CAT_COLOR = { verano: "#C68A2E", invierno: "#3D6E8C" };

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
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ model: "claude-sonnet-4-6", max_tokens: 1000, messages: [{ role: "user", content }] }),
  });
  const data = await res.json();
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
        <div className="flex items-center gap-2 justify-center mb-1"><Wheat color="var(--gold)" size={24} /></div>
        <div className="cc-h text-center" style={{ fontSize: 20, fontWeight: 600 }}>Campo & Costo</div>
        <div className="text-center mb-5" style={{ fontSize: 12, color: "#8A8570" }}>Campañas · Cultivos · Gastos e Ingresos</div>
        <label style={{ fontSize: 12, color: "#8A8570" }}>Email</label>
        <input className="cc-input mb-3" type="email" required value={email} onChange={(e) => setEmail(e.target.value)} />
        <label style={{ fontSize: 12, color: "#8A8570" }}>Contraseña</label>
        <input className="cc-input mb-3" type="password" required minLength={6} value={pass} onChange={(e) => setPass(e.target.value)} />
        {error && <div style={{ fontSize: 12.5, color: "var(--rust)", marginBottom: 10 }}>{error}</div>}
        <button className="cc-btn cc-btn-primary w-full justify-center" disabled={cargando} type="submit">
          {cargando ? <Loader2 size={15} className="animate-spin" /> : null}{modo === "ingresar" ? "Ingresar" : "Crear cuenta"}
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
  const [campanias, setCampanias] = useState([]);
  const [cultivos, setCultivos] = useState([]);
  const [lotes, setLotes] = useState([]);
  const [insumosCompras, setInsumosCompras] = useState([]);
  const [gastosTodos, setGastosTodos] = useState([]);
  const [nav, setNav] = useState({ view: "campanias", campaniaId: null, cultivoId: null });

  useEffect(() => onAuthStateChanged(auth, (u) => setUser(u)), []);

  useEffect(() => {
    if (!user) return;
    const u1 = onSnapshot(collection(db, "campanias"), (s) => setCampanias(s.docs.map((d) => ({ id: d.id, ...d.data() }))));
    const u2 = onSnapshot(collection(db, "cultivos"), (s) => setCultivos(s.docs.map((d) => ({ id: d.id, ...d.data() }))));
    const u3 = onSnapshot(collection(db, "lotes"), (s) => setLotes(s.docs.map((d) => ({ id: d.id, ...d.data() }))));
    const u4 = onSnapshot(collection(db, "insumos_compras"), (s) => setInsumosCompras(s.docs.map((d) => ({ id: d.id, ...d.data() }))));
    const u5 = onSnapshot(collection(db, "gastos"), (s) => setGastosTodos(s.docs.map((d) => ({ id: d.id, ...d.data() }))));
    return () => { u1(); u2(); u3(); u4(); u5(); };
  }, [user]);

  if (user === undefined) {
    return <div style={{ background: "var(--bg)", minHeight: "100vh" }} className="flex items-center justify-center"><style>{GLOBAL_STYLES}</style><Loader2 className="animate-spin" color="var(--soil)" size={28} /></div>;
  }
  if (!user) return <Login />;

  const campaniasApi = {
    add: (data) => addDoc(collection(db, "campanias"), data),
    remove: (id) => deleteDoc(doc(db, "campanias", id)),
  };
  const cultivosApi = {
    add: (data) => addDoc(collection(db, "cultivos"), data),
    remove: (id) => deleteDoc(doc(db, "cultivos", id)),
  };
  const lotesApi = {
    add: (data) => addDoc(collection(db, "lotes"), data),
    remove: (id) => deleteDoc(doc(db, "lotes", id)),
  };
  const insumosApi = {
    add: (data) => addDoc(collection(db, "insumos_compras"), data),
    remove: (id) => deleteDoc(doc(db, "insumos_compras", id)),
  };

  // Stock disponible por insumo = total comprado - total consumido en gastos de cultivos
  const stockInsumos = (() => {
    const mapa = {};
    insumosCompras.forEach((c) => {
      const k = c.nombre;
      if (!mapa[k]) mapa[k] = { nombre: k, litrosComprados: 0, costoComprado: 0, litrosConsumidos: 0 };
      mapa[k].litrosComprados += Number(c.litros || 0);
      mapa[k].costoComprado += Number(c.precio || 0);
    });
    gastosTodos.forEach((g) => {
      if (g.insumoNombre && mapa[g.insumoNombre]) mapa[g.insumoNombre].litrosConsumidos += Number(g.litrosUsados || 0);
    });
    return Object.values(mapa).map((i) => ({
      ...i,
      disponible: i.litrosComprados - i.litrosConsumidos,
      costoPromedioPorLitro: i.litrosComprados ? i.costoComprado / i.litrosComprados : 0,
    }));
  })();

  const campaniaActual = campanias.find((c) => c.id === nav.campaniaId);
  const cultivoActual = cultivos.find((c) => c.id === nav.cultivoId);

  return (
    <div style={{ background: "var(--bg)", fontFamily: "var(--font-body)", color: "var(--ink)", minHeight: "100vh" }}>
      <style>{GLOBAL_STYLES}</style>
      <header style={{ background: "var(--soil)" }} className="px-6 py-4">
        <div className="flex items-center justify-between max-w-6xl mx-auto">
          <button className="flex items-center gap-3" onClick={() => setNav({ view: "campanias", campaniaId: null, cultivoId: null })}>
            <Wheat color="var(--gold)" size={26} />
            <div style={{ textAlign: "left" }}>
              <div className="cc-h" style={{ color: "#fff", fontSize: 20, fontWeight: 600, lineHeight: 1 }}>Campo & Costo</div>
              <div style={{ color: "#B8C2AC", fontSize: 12 }}>Campañas · Cultivos · Gastos e Ingresos</div>
            </div>
          </button>
          <div className="flex items-center gap-3">
            <button onClick={() => setNav({ view: "insumos", campaniaId: null, cultivoId: null })} className="cc-btn" style={{ background: "transparent", border: "1px solid #4C5A40", color: "#D8DECB", padding: "6px 12px", fontSize: 12.5 }}>
              <Package size={13} /> Insumos
            </button>
            <button onClick={() => setNav({ view: "lotes", campaniaId: null, cultivoId: null })} className="cc-btn" style={{ background: "transparent", border: "1px solid #4C5A40", color: "#D8DECB", padding: "6px 12px", fontSize: 12.5 }}>
              <MapPin size={13} /> Lotes
            </button>
            <span style={{ color: "#D8DECB", fontSize: 12.5 }}>{user.email}</span>
            <button onClick={() => signOut(auth)} className="cc-btn" style={{ background: "transparent", border: "1px solid #4C5A40", color: "#D8DECB", padding: "6px 12px", fontSize: 12.5 }}><LogOut size={13} /> Salir</button>
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-6 py-6">
        <Breadcrumb nav={nav} setNav={setNav} campania={campaniaActual} cultivo={cultivoActual} />

        {nav.view === "campanias" && <CampaniasView campanias={campanias} api={campaniasApi} cultivos={cultivos} onOpen={(id) => setNav({ view: "cultivos", campaniaId: id, cultivoId: null })} />}

        {nav.view === "cultivos" && campaniaActual && (
          <CultivosDeCampania campania={campaniaActual} cultivos={cultivos.filter((c) => c.campaniaId === campaniaActual.id)} api={cultivosApi}
            onOpen={(id) => setNav({ view: "cultivo", campaniaId: campaniaActual.id, cultivoId: id })} />
        )}

        {nav.view === "cultivo" && cultivoActual && <CultivoDetail cultivo={cultivoActual} lotes={lotes} user={user} stockInsumos={stockInsumos} insumosCompras={insumosCompras} gastosTodos={gastosTodos} />}

        {nav.view === "lotes" && <LotesView lotes={lotes} api={lotesApi} />}

        {nav.view === "insumos" && <InsumosView compras={insumosCompras} api={insumosApi} stockInsumos={stockInsumos} user={user} />}
      </main>
    </div>
  );
}

function Breadcrumb({ nav, setNav, campania, cultivo }) {
  if (nav.view === "campanias" || nav.view === "lotes" || nav.view === "insumos") return null;
  return (
    <div className="flex items-center gap-1 mb-4" style={{ fontSize: 13, color: "#8A8570" }}>
      <button onClick={() => setNav({ view: "campanias", campaniaId: null, cultivoId: null })} style={{ color: "var(--frost)", fontWeight: 600 }}>Campañas</button>
      {campania && (
        <>
          <ChevronRight size={13} />
          <button onClick={() => setNav({ view: "cultivos", campaniaId: campania.id, cultivoId: null })} style={{ color: nav.view === "cultivos" ? "var(--ink)" : "var(--frost)", fontWeight: 600 }}>
            {campania.nombre || campania.anio}
          </button>
        </>
      )}
      {cultivo && (
        <>
          <ChevronRight size={13} />
          <span style={{ color: "var(--ink)", fontWeight: 600 }}>{cultivo.nombre}</span>
        </>
      )}
    </div>
  );
}

function EmptyState({ icon: Icon, title, text }) {
  return (
    <div className="cc-card flex flex-col items-center text-center py-16 px-6">
      <Icon size={30} color="var(--gold)" style={{ marginBottom: 10 }} />
      <div className="cc-h" style={{ fontSize: 17, fontWeight: 600 }}>{title}</div>
      <div style={{ color: "#8A8570", fontSize: 13, maxWidth: 380, marginTop: 4 }}>{text}</div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Campañas                                                            */
/* ------------------------------------------------------------------ */
function CampaniasView({ campanias, api, cultivos, onOpen }) {
  const [anio, setAnio] = useState(new Date().getFullYear());
  const [nombre, setNombre] = useState("");

  const crear = () => { api.add({ anio: Number(anio), nombre: nombre.trim() || `Campaña ${anio}` }); setNombre(""); };
  const eliminar = (id) => { if (confirm("¿Eliminar esta campaña? Los cultivos asociados quedarán huérfanos.")) api.remove(id); };

  return (
    <div className="space-y-5">
      <div className="cc-card p-4">
        <div className="cc-h" style={{ fontSize: 16, fontWeight: 600, marginBottom: 10 }}>Nueva campaña</div>
        <div className="flex gap-3 flex-wrap items-end">
          <div style={{ width: 140 }}><label style={{ fontSize: 12, color: "#8A8570" }}>Año</label><input className="cc-input" type="number" value={anio} onChange={(e) => setAnio(e.target.value)} /></div>
          <div style={{ flex: 1, minWidth: 180 }}><label style={{ fontSize: 12, color: "#8A8570" }}>Nombre (opcional)</label><input className="cc-input" value={nombre} onChange={(e) => setNombre(e.target.value)} placeholder={`Campaña ${anio}`} /></div>
          <button className="cc-btn cc-btn-primary" onClick={crear}><Plus size={15} /> Crear</button>
        </div>
      </div>

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
                  <button onClick={(e) => { e.stopPropagation(); eliminar(c.id); }}><Trash2 size={14} color="var(--rust)" /></button>
                </div>
                <div className="flex items-center gap-1 mt-3" style={{ color: "var(--frost)", fontSize: 12.5, fontWeight: 600 }}>Ver cultivos <ChevronRight size={13} /></div>
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
function CultivosDeCampania({ campania, cultivos, api, onOpen }) {
  const [categoria, setCategoria] = useState("verano");
  const [nombre, setNombre] = useState(CULTIVOS_VERANO[0]);
  const opciones = categoria === "verano" ? CULTIVOS_VERANO : CULTIVOS_INVIERNO;

  const crear = () => api.add({ campaniaId: campania.id, categoria, nombre });
  const eliminar = (id) => { if (confirm("¿Eliminar este cultivo? Se pierde el acceso a sus gastos e ingresos cargados.")) api.remove(id); };

  return (
    <div className="space-y-5">
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
          <button className="cc-btn cc-btn-primary" onClick={crear}><Plus size={15} /> Agregar</button>
        </div>
      </div>

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
                  {c.categoria === "verano" ? <Sun color="var(--gold)" size={18} /> : <Snowflake color="var(--frost)" size={18} />}
                  <button onClick={(e) => { e.stopPropagation(); eliminar(c.id); }}><Trash2 size={13} color="var(--rust)" /></button>
                </div>
              </div>
              <div className="flex items-center gap-1 mt-3" style={{ color: "var(--frost)", fontSize: 12.5, fontWeight: 600 }}>Gastos e ingresos <ChevronRight size={13} /></div>
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
function CultivoDetail({ cultivo, lotes, user, stockInsumos, insumosCompras, gastosTodos }) {
  const [tab, setTab] = useState("resumen");
  const [gastos, setGastos] = useState([]);
  const [ventas, setVentas] = useState([]);
  const [remitos, setRemitos] = useState([]);

  useEffect(() => {
    const q1 = query(collection(db, "gastos"), where("cultivoId", "==", cultivo.id));
    const q2 = query(collection(db, "ventas"), where("cultivoId", "==", cultivo.id));
    const q3 = query(collection(db, "remitos"), where("cultivoId", "==", cultivo.id));
    const u1 = onSnapshot(q1, (s) => setGastos(s.docs.map((d) => ({ id: d.id, ...d.data() }))));
    const u2 = onSnapshot(q2, (s) => setVentas(s.docs.map((d) => ({ id: d.id, ...d.data() }))));
    const u3 = onSnapshot(q3, (s) => setRemitos(s.docs.map((d) => ({ id: d.id, ...d.data() }))));
    return () => { u1(); u2(); u3(); };
  }, [cultivo.id]);

  const gastosApi = { add: (d) => addDoc(collection(db, "gastos"), d), update: (id, d) => updateDoc(doc(db, "gastos", id), d), remove: (id) => deleteDoc(doc(db, "gastos", id)) };
  const ventasApi = { add: (d) => addDoc(collection(db, "ventas"), d), remove: (id) => deleteDoc(doc(db, "ventas", id)) };
  const remitosApi = { add: (d) => addDoc(collection(db, "remitos"), d), remove: (id) => deleteDoc(doc(db, "remitos", id)) };

  const totalGastos = gastos.reduce((s, g) => s + Number(g.monto || 0), 0);
  const totalIngresos = ventas.reduce((s, v) => s + Number(v.toneladas || 0) * Number(v.dolaresPorTonelada || 0), 0);
  const totalToneladasVentas = ventas.reduce((s, v) => s + Number(v.toneladas || 0), 0);
  const totalKgSL = remitos.reduce((s, r) => s + Number(r.kgSL || 0), 0);
  const totalTonRemitos = totalKgSL / 1000;

  const TABS = [
    { id: "resumen", label: "Resumen", icon: TrendingUp },
    { id: "gastos", label: "Gastos", icon: Receipt },
    { id: "ventas", label: "Ingresos · Ventas", icon: DollarSign },
    { id: "remitos", label: "Ingresos · Remitos", icon: Truck },
  ];

  return (
    <div>
      <div className="flex gap-1 mb-4 flex-wrap">
        {TABS.map((t) => {
          const Icon = t.icon; const active = tab === t.id;
          return (
            <button key={t.id} onClick={() => setTab(t.id)} className="cc-sub" style={{ background: active ? "var(--panel)" : "transparent", color: active ? "var(--soil)" : "#8A8570", border: active ? "1px solid var(--line)" : "1px solid transparent", borderBottom: active ? "1px solid var(--panel)" : "none" }}>
              <Icon size={14} /> {t.label}
            </button>
          );
        })}
      </div>

      {tab === "resumen" && (
        <div className="space-y-4">
          <div className="grid gap-4" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(200px,1fr))" }}>
            <StatCard label="Total gastos" value={fmtUSD(totalGastos)} icon={Receipt} color="var(--rust)" />
            <StatCard label="Total ingresos (ventas)" value={fmtUSD(totalIngresos)} icon={DollarSign} color="var(--soil-light)" />
            <StatCard label="Margen" value={fmtUSD(totalIngresos - totalGastos)} icon={totalIngresos - totalGastos >= 0 ? TrendingUp : TrendingDown} color={totalIngresos - totalGastos >= 0 ? "var(--soil-light)" : "var(--rust)"} />
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
                <><CheckCircle2 size={15} color="var(--soil-light)" /><span style={{ color: "var(--soil-light)" }}>Los remitos coinciden con las toneladas vendidas.</span></>
              ) : (
                <><AlertCircle size={15} color="var(--rust)" /><span style={{ color: "var(--rust)" }}>
                  Diferencia de {fmt(Math.abs(totalTonRemitos - totalToneladasVentas), 2)} tn entre remitos y ventas.
                </span></>
              )}
            </div>
          </div>
        </div>
      )}

      {tab === "gastos" && <GastosTab cultivo={cultivo} gastos={gastos} api={gastosApi} user={user} stockInsumos={stockInsumos} insumosCompras={insumosCompras} gastosTodos={gastosTodos} />}
      {tab === "ventas" && <VentasTab cultivo={cultivo} ventas={ventas} api={ventasApi} />}
      {tab === "remitos" && <RemitosTab cultivo={cultivo} remitos={remitos} api={remitosApi} lotes={lotes} totalToneladasVentas={totalToneladasVentas} />}
    </div>
  );
}

function StatCard({ label, value, icon: Icon, color }) {
  return (
    <div className="cc-card p-4 flex items-center gap-3">
      <div style={{ background: color + "1A", borderRadius: 8, padding: 8 }}><Icon size={18} color={color} /></div>
      <div><div style={{ fontSize: 11, color: "#8A8570", textTransform: "uppercase", letterSpacing: ".03em" }}>{label}</div><div className="cc-mono" style={{ fontSize: 17, fontWeight: 600 }}>{value}</div></div>
    </div>
  );
}
function MiniStat({ label, value }) {
  return <div><div style={{ fontSize: 10.5, color: "#8A8570", textTransform: "uppercase", letterSpacing: ".03em" }}>{label}</div><div className="cc-mono" style={{ fontSize: 15, fontWeight: 500 }}>{value}</div></div>;
}

/* ------------------------------------------------------------------ */
/*  Gastos                                                               */
/* ------------------------------------------------------------------ */
const emptyGasto = (email) => ({ origen: "", monto: "", detalle: "", fecha: "", usuario: email });

function GastosTab({ cultivo, gastos, api, user, stockInsumos, insumosCompras, gastosTodos }) {
  const [form, setForm] = useState(emptyGasto(user.email));
  const [editId, setEditId] = useState(null);
  const [tipo, setTipo] = useState("general"); // "general" | "insumo"
  const [insumoSel, setInsumoSel] = useState("");
  const [litrosUsados, setLitrosUsados] = useState("");
  const [archivo, setArchivo] = useState(null);
  const [preview, setPreview] = useState(null);
  const [subiendo, setSubiendo] = useState(false);
  const [extrayendo, setExtrayendo] = useState(false);
  const [grabando, setGrabando] = useState(false);
  const [transcripcion, setTranscripcion] = useState("");
  const [interpretando, setInterpretando] = useState(false);
  const [error, setError] = useState("");
  const recRef = useRef(null);
  const set = (k, v) => setForm({ ...form, [k]: v });

  const origenesSugeridos = Array.from(new Set(gastos.map((g) => g.origen).filter(Boolean)));
  const insumoElegido = stockInsumos.find((i) => i.nombre === insumoSel);
  const comprasDeEsteInsumo = insumosCompras.filter((c) => c.nombre === insumoSel);
  const litrosYaConsumidosPorOtros = gastosTodos
    .filter((g) => g.insumoNombre === insumoSel && g.id !== editId)
    .reduce((s, g) => s + Number(g.litrosUsados || 0), 0);
  const fifo = insumoSel ? costoFIFO(comprasDeEsteInsumo, litrosYaConsumidosPorOtros, Number(litrosUsados || 0)) : { costoTotal: 0, costoPromedioEfectivo: 0 };
  const montoCalculado = fifo.costoTotal;

  const editar = (g) => {
    setEditId(g.id);
    if (g.insumoNombre) {
      setTipo("insumo"); setInsumoSel(g.insumoNombre); setLitrosUsados(String(g.litrosUsados ?? ""));
      setForm({ origen: g.origen || "", monto: g.monto ?? "", detalle: g.detalle || "", fecha: g.fecha || "", usuario: g.usuario || user.email });
    } else {
      setTipo("general"); setInsumoSel(""); setLitrosUsados("");
      setForm({ origen: g.origen || "", monto: g.monto ?? "", detalle: g.detalle || "", fecha: g.fecha || "", usuario: g.usuario || user.email });
    }
    setArchivo(null); setPreview(null); setTranscripcion(""); setError("");
    window.scrollTo({ top: 0, behavior: "smooth" });
  };
  const cancelarEdicion = () => {
    setEditId(null); setForm(emptyGasto(user.email)); setArchivo(null); setPreview(null); setTranscripcion("");
    setTipo("general"); setInsumoSel(""); setLitrosUsados("");
  };

  const onFile = (e) => {
    const f = e.target.files[0];
    if (!f) return;
    setArchivo(f);
    if (f.type.startsWith("image/")) setPreview(URL.createObjectURL(f)); else setPreview(null);
  };

  const extraerDeFactura = async () => {
    if (!archivo || !archivo.type.startsWith("image/")) { setError("La extracción automática funciona con fotos (JPG/PNG) de la factura."); return; }
    setExtrayendo(true); setError("");
    try {
      const b64 = await fileToBase64(archivo);
      const data = await askClaudeJSON([
        { type: "image", source: { type: "base64", media_type: archivo.type, data: b64 } },
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
      if (!editId && Number(litrosUsados) > (insumoElegido?.disponible || 0)) {
        if (!confirm(`Solo hay ${fmt(insumoElegido?.disponible || 0, 1)} litros disponibles de ${insumoSel}. ¿Registrar igual el consumo?`)) return;
      }
    } else if (!form.origen || !form.monto || !form.fecha) { alert("Completá al menos origen, monto y fecha."); return; }

    let facturaUrl, facturaNombre;
    if (archivo) {
      setSubiendo(true);
      try {
        const path = `facturas/${cultivo.id}/${Date.now()}_${archivo.name}`;
        const r = ref(storage, path);
        await uploadBytes(r, archivo);
        facturaUrl = await getDownloadURL(r);
        facturaNombre = archivo.name;
      } catch (e) { alert("No se pudo subir la factura adjunta, pero se guardará el gasto sin el archivo."); }
      setSubiendo(false);
    }

    let datos;
    if (tipo === "insumo") {
      datos = {
        origen: form.origen || insumoSel, monto: montoCalculado, detalle: form.detalle || `Consumo de ${insumoSel} — ${litrosUsados} L`,
        fecha: form.fecha, usuario: form.usuario || user.email, insumoNombre: insumoSel, litrosUsados: Number(litrosUsados), costoPorLitro: fifo.costoPromedioEfectivo,
      };
    } else {
      datos = { origen: form.origen, monto: Number(form.monto), detalle: form.detalle, fecha: form.fecha, usuario: form.usuario || user.email, insumoNombre: null, litrosUsados: null };
    }
    if (archivo) { datos.facturaUrl = facturaUrl; datos.facturaNombre = facturaNombre; }

    if (editId) {
      await api.update(editId, datos);
    } else {
      await api.add({ cultivoId: cultivo.id, ...datos, facturaUrl: facturaUrl || null, facturaNombre: facturaNombre || null });
    }
    setEditId(null); setForm(emptyGasto(user.email)); setArchivo(null); setPreview(null); setTranscripcion("");
    setTipo("general"); setInsumoSel(""); setLitrosUsados("");
  };

  const eliminar = (id) => { if (confirm("¿Eliminar este gasto?")) api.remove(id); };

  return (
    <div className="space-y-5">
      <div className="cc-card p-4 space-y-4">
        <div className="flex flex-wrap gap-2">
          <label className="cc-btn cc-btn-ghost" style={{ cursor: "pointer" }}><Paperclip size={14} /> {archivo ? archivo.name : "Adjuntar factura (imagen o PDF)"}<input type="file" accept="image/*,.pdf" capture="environment" onChange={onFile} style={{ display: "none" }} /></label>
          {archivo?.type?.startsWith("image/") && <button className="cc-btn cc-btn-ghost" onClick={extraerDeFactura} disabled={extrayendo}>{extrayendo ? <Loader2 size={14} className="animate-spin" /> : <Camera size={14} />} Extraer datos con IA</button>}
          <button className="cc-btn" style={{ background: grabando ? "var(--rust)" : "var(--soil)", color: "#fff" }} onClick={grabando ? detener : grabar}><Mic size={14} /> {grabando ? "Detener" : "Dictar por voz"}</button>
        </div>

        {preview && <img src={preview} alt="Factura" style={{ maxWidth: 140, borderRadius: 8, border: "1px solid var(--line)" }} />}

        {(transcripcion || grabando) && (
          <div>
            <textarea className="cc-input" rows={2} placeholder="Ej: flete a Nueva Palmira, ochocientos dólares, quince de marzo" value={transcripcion} onChange={(e) => setTranscripcion(e.target.value)} />
            <button className="cc-btn cc-btn-ghost mt-2" onClick={interpretarVoz} disabled={interpretando}>{interpretando ? <Loader2 size={14} className="animate-spin" /> : <Sprout size={14} />} Interpretar con IA</button>
          </div>
        )}

        {error && <div style={{ color: "var(--rust)", fontSize: 13, display: "flex", gap: 6, alignItems: "center" }}><AlertCircle size={14} />{error}</div>}

        {editId && (
          <div className="flex items-center gap-2" style={{ background: "#FDF3E0", border: "1px solid var(--gold)", borderRadius: 8, padding: "6px 10px", fontSize: 12.5, color: "#7A5A1E" }}>
            <Pencil size={13} /> Editando un gasto ya guardado.
          </div>
        )}

        <div className="flex gap-2">
          <button className="cc-btn" onClick={() => setTipo("general")} style={{ background: tipo === "general" ? "var(--soil)" : "#fff", color: tipo === "general" ? "#fff" : "var(--ink)", border: "1px solid var(--line)" }}>Gasto general</button>
          <button className="cc-btn" onClick={() => setTipo("insumo")} style={{ background: tipo === "insumo" ? "var(--soil)" : "#fff", color: tipo === "insumo" ? "#fff" : "var(--ink)", border: "1px solid var(--line)" }}><Boxes size={14} /> Insumo de stock</button>
        </div>

        {tipo === "insumo" ? (
          <div className="grid gap-3" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(150px,1fr))" }}>
            <div>
              <label style={{ fontSize: 12, color: "#8A8570" }}>Insumo</label>
              <select className="cc-input" value={insumoSel} onChange={(e) => setInsumoSel(e.target.value)}>
                <option value="">Elegir...</option>
                {stockInsumos.map((i) => <option key={i.nombre} value={i.nombre}>{i.nombre} ({fmt(i.disponible, 1)} L disp.)</option>)}
              </select>
              {stockInsumos.length === 0 && <div style={{ fontSize: 11.5, color: "#8A8570" }}>No hay insumos cargados todavía — cargalos desde "Insumos" en el menú superior.</div>}
            </div>
            <div><label style={{ fontSize: 12, color: "#8A8570" }}>Litros usados</label><input className="cc-input" type="number" value={litrosUsados} onChange={(e) => setLitrosUsados(e.target.value)} /></div>
            <div><label style={{ fontSize: 12, color: "#8A8570" }}>Fecha</label><input className="cc-input" type="date" value={form.fecha} onChange={(e) => set("fecha", e.target.value)} /></div>
            <div><label style={{ fontSize: 12, color: "#8A8570" }}>Costo estimado (FIFO)</label><input className="cc-input" value={`U$S ${fmt(montoCalculado, 2)}`} disabled /></div>
            <div style={{ gridColumn: "1 / -1" }}><label style={{ fontSize: 12, color: "#8A8570" }}>Detalle (opcional)</label><input className="cc-input" value={form.detalle} onChange={(e) => set("detalle", e.target.value)} placeholder={insumoSel ? `Consumo de ${insumoSel}` : ""} /></div>
          </div>
        ) : (
          <div className="grid gap-3" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(150px,1fr))" }}>
            <div>
              <label style={{ fontSize: 12, color: "#8A8570" }}>Origen</label>
              <input className="cc-input" list="origenes-gastos" value={form.origen} onChange={(e) => set("origen", e.target.value)} placeholder="Proveedor, gasoil, renta..." />
              <datalist id="origenes-gastos">{origenesSugeridos.map((o) => <option key={o} value={o} />)}</datalist>
            </div>
            <div><label style={{ fontSize: 12, color: "#8A8570" }}>Monto (U$S)</label><input className="cc-input" type="number" value={form.monto} onChange={(e) => set("monto", e.target.value)} /></div>
            <div><label style={{ fontSize: 12, color: "#8A8570" }}>Fecha</label><input className="cc-input" type="date" value={form.fecha} onChange={(e) => set("fecha", e.target.value)} /></div>
            <div><label style={{ fontSize: 12, color: "#8A8570" }}>Usuario</label><input className="cc-input" value={form.usuario} disabled /></div>
            <div style={{ gridColumn: "1 / -1" }}><label style={{ fontSize: 12, color: "#8A8570" }}>Detalle</label><input className="cc-input" value={form.detalle} onChange={(e) => set("detalle", e.target.value)} placeholder="Descripción del gasto" /></div>
          </div>
        )}
        {editId && archivo === null && <div style={{ fontSize: 12, color: "#8A8570" }}>La factura adjunta actual se mantiene salvo que subas una nueva arriba.</div>}
        <div className="flex gap-2">
          <button className="cc-btn cc-btn-primary" onClick={guardar} disabled={subiendo}>
            {subiendo ? <Loader2 size={15} className="animate-spin" /> : editId ? <Pencil size={15} /> : <Plus size={15} />} {editId ? "Guardar cambios" : "Guardar gasto"}
          </button>
          {editId && <button className="cc-btn cc-btn-ghost" onClick={cancelarEdicion}><X size={15} /> Cancelar</button>}
        </div>
      </div>

      {gastos.length === 0 ? <EmptyState icon={Receipt} title="Sin gastos cargados" text="Cargá el primer gasto de este cultivo, escrito, por voz o desde una foto de factura." /> : (
        <div className="cc-card overflow-hidden">
          <table className="w-full" style={{ fontSize: 13 }}>
            <thead><tr style={{ background: "#EEEADA", textAlign: "left" }}><th className="px-3 py-2">Fecha</th><th className="px-3 py-2">Origen</th><th className="px-3 py-2">Detalle</th><th className="px-3 py-2">Usuario</th><th className="px-3 py-2 text-right">Monto</th><th className="px-3 py-2"></th><th className="px-3 py-2"></th><th className="px-3 py-2"></th></tr></thead>
            <tbody>
              {[...gastos].sort((a, b) => (b.fecha || "").localeCompare(a.fecha || "")).map((g) => (
                <tr key={g.id} style={{ borderTop: "1px solid var(--line)", background: editId === g.id ? "#FDF3E0" : "transparent" }}>
                  <td className="px-3 py-2 cc-mono">{g.fecha}</td>
                  <td className="px-3 py-2">
                    {g.origen}
                    {g.insumoNombre && <span className="cc-chip" style={{ background: "#EDE7F6", color: "#5B4B8A", marginLeft: 6 }}>{fmt(g.litrosUsados, 1)} L</span>}
                  </td>
                  <td className="px-3 py-2" style={{ color: "#5A5647" }}>{g.detalle}</td>
                  <td className="px-3 py-2" style={{ color: "#8A8570", fontSize: 12 }}>{g.usuario}</td>
                  <td className="px-3 py-2 text-right cc-mono">{fmtUSD(g.monto)}</td>
                  <td className="px-3 py-2">{g.facturaUrl && <a href={g.facturaUrl} target="_blank" rel="noreferrer"><FileText size={14} color="var(--frost)" /></a>}</td>
                  <td className="px-3 py-2 text-right"><button onClick={() => editar(g)}><Pencil size={13} color="var(--frost)" /></button></td>
                  <td className="px-3 py-2 text-right"><button onClick={() => eliminar(g.id)}><Trash2 size={13} color="var(--rust)" /></button></td>
                </tr>
              ))}
            </tbody>
            <tfoot><tr style={{ borderTop: "2px solid var(--line)", fontWeight: 700 }}><td className="px-3 py-2" colSpan={4}>Total</td><td className="px-3 py-2 text-right cc-mono">{fmtUSD(gastos.reduce((s, g) => s + Number(g.monto || 0), 0))}</td><td colSpan={3}></td></tr></tfoot>
          </table>
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Ingresos · Ventas                                                    */
/* ------------------------------------------------------------------ */
function VentasTab({ cultivo, ventas, api }) {
  const [form, setForm] = useState({ fecha: "", origen: "", toneladas: "", dolaresPorTonelada: "" });
  const set = (k, v) => setForm({ ...form, [k]: v });
  const origenesSugeridos = Array.from(new Set(ventas.map((v) => v.origen).filter(Boolean)));

  const guardar = () => {
    if (!form.fecha || !form.toneladas || !form.dolaresPorTonelada) { alert("Completá fecha, toneladas y U$S/ton."); return; }
    api.add({ cultivoId: cultivo.id, fecha: form.fecha, origen: form.origen, toneladas: Number(form.toneladas), dolaresPorTonelada: Number(form.dolaresPorTonelada) });
    setForm({ fecha: "", origen: "", toneladas: "", dolaresPorTonelada: "" });
  };
  const eliminar = (id) => { if (confirm("¿Eliminar esta venta?")) api.remove(id); };
  const totalTon = ventas.reduce((s, v) => s + Number(v.toneladas || 0), 0);
  const totalUSD = ventas.reduce((s, v) => s + Number(v.toneladas || 0) * Number(v.dolaresPorTonelada || 0), 0);

  return (
    <div className="space-y-5">
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
        <button className="cc-btn cc-btn-primary" onClick={guardar}><Plus size={15} /> Guardar venta</button>
      </div>

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
                  <td className="px-3 py-2 text-right"><button onClick={() => eliminar(v.id)}><Trash2 size={13} color="var(--rust)" /></button></td>
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
function RemitosTab({ cultivo, remitos, api, lotes, totalToneladasVentas }) {
  const [form, setForm] = useState({ fecha: "", remito: "", campo: "", destino: "", kgTolva: "", kgBrutos: "", kgSL: "", humedad: "" });
  const set = (k, v) => setForm({ ...form, [k]: v });

  const guardar = () => {
    if (!form.fecha || !form.remito || !form.kgSL) { alert("Completá al menos fecha, N° de remito y Kg SL."); return; }
    api.add({
      cultivoId: cultivo.id, fecha: form.fecha, remito: form.remito, campo: form.campo, destino: form.destino,
      kgTolva: Number(form.kgTolva || 0), kgBrutos: Number(form.kgBrutos || 0), kgSL: Number(form.kgSL || 0), humedad: Number(form.humedad || 0),
    });
    setForm({ fecha: "", remito: "", campo: "", destino: "", kgTolva: "", kgBrutos: "", kgSL: "", humedad: "" });
  };
  const eliminar = (id) => { if (confirm("¿Eliminar este remito?")) api.remove(id); };
  const totalKgSL = remitos.reduce((s, r) => s + Number(r.kgSL || 0), 0);

  return (
    <div className="space-y-5">
      <div className="cc-card p-4 space-y-3">
        <div className="grid gap-3" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(130px,1fr))" }}>
          <div><label style={{ fontSize: 12, color: "#8A8570" }}>Fecha</label><input className="cc-input" type="date" value={form.fecha} onChange={(e) => set("fecha", e.target.value)} /></div>
          <div><label style={{ fontSize: 12, color: "#8A8570" }}>N° Remito</label><input className="cc-input" value={form.remito} onChange={(e) => set("remito", e.target.value)} /></div>
          <div>
            <label style={{ fontSize: 12, color: "#8A8570" }}>Campo</label>
            <input className="cc-input" list="lotes-list" value={form.campo} onChange={(e) => set("campo", e.target.value)} placeholder="Lote / establecimiento" />
            <datalist id="lotes-list">{lotes.map((l) => <option key={l.id} value={l.nombre} />)}</datalist>
          </div>
          <div><label style={{ fontSize: 12, color: "#8A8570" }}>Destino</label><input className="cc-input" value={form.destino} onChange={(e) => set("destino", e.target.value)} placeholder="Silo / planta" /></div>
          <div><label style={{ fontSize: 12, color: "#8A8570" }}>Kg tolva</label><input className="cc-input" type="number" value={form.kgTolva} onChange={(e) => set("kgTolva", e.target.value)} /></div>
          <div><label style={{ fontSize: 12, color: "#8A8570" }}>Kg brutos</label><input className="cc-input" type="number" value={form.kgBrutos} onChange={(e) => set("kgBrutos", e.target.value)} /></div>
          <div><label style={{ fontSize: 12, color: "#8A8570" }}>Kg SL</label><input className="cc-input" type="number" value={form.kgSL} onChange={(e) => set("kgSL", e.target.value)} /></div>
          <div><label style={{ fontSize: 12, color: "#8A8570" }}>Humedad (%)</label><input className="cc-input" type="number" value={form.humedad} onChange={(e) => set("humedad", e.target.value)} /></div>
        </div>
        <button className="cc-btn cc-btn-primary" onClick={guardar}><Plus size={15} /> Guardar remito</button>
      </div>

      {remitos.length === 0 ? <EmptyState icon={Truck} title="Sin remitos cargados" text="Cargá cada remito de entrega de grano con sus kilos seco y limpio (Kg SL)." /> : (
        <div className="cc-card overflow-hidden">
          <table className="w-full" style={{ fontSize: 12.5 }}>
            <thead><tr style={{ background: "#EEEADA", textAlign: "left" }}><th className="px-3 py-2">Fecha</th><th className="px-3 py-2">Remito</th><th className="px-3 py-2">Campo</th><th className="px-3 py-2">Destino</th><th className="px-3 py-2 text-right">Kg tolva</th><th className="px-3 py-2 text-right">Kg brutos</th><th className="px-3 py-2 text-right">Kg SL</th><th className="px-3 py-2 text-right">Hum.%</th><th className="px-3 py-2"></th></tr></thead>
            <tbody>
              {[...remitos].sort((a, b) => (b.fecha || "").localeCompare(a.fecha || "")).map((r) => (
                <tr key={r.id} style={{ borderTop: "1px solid var(--line)" }}>
                  <td className="px-3 py-2 cc-mono">{r.fecha}</td><td className="px-3 py-2">{r.remito}</td><td className="px-3 py-2">{r.campo}</td><td className="px-3 py-2">{r.destino}</td>
                  <td className="px-3 py-2 text-right cc-mono">{fmt(r.kgTolva)}</td><td className="px-3 py-2 text-right cc-mono">{fmt(r.kgBrutos)}</td>
                  <td className="px-3 py-2 text-right cc-mono">{fmt(r.kgSL)}</td><td className="px-3 py-2 text-right cc-mono">{fmt(r.humedad, 1)}</td>
                  <td className="px-3 py-2 text-right"><button onClick={() => eliminar(r.id)}><Trash2 size={13} color="var(--rust)" /></button></td>
                </tr>
              ))}
            </tbody>
            <tfoot><tr style={{ borderTop: "2px solid var(--line)", fontWeight: 700 }}><td className="px-3 py-2" colSpan={6}>Total Kg SL</td><td className="px-3 py-2 text-right cc-mono">{fmt(totalKgSL)}</td><td colSpan={2}></td></tr></tfoot>
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
/*  Lotes (referencia global)                                           */
/* ------------------------------------------------------------------ */
function LotesView({ lotes, api }) {
  const [nombre, setNombre] = useState("");
  const [hectareas, setHectareas] = useState("");
  const guardar = () => { if (!nombre.trim()) return; api.add({ nombre: nombre.trim(), hectareas: hectareas ? Number(hectareas) : null }); setNombre(""); setHectareas(""); };
  const eliminar = (id) => { if (confirm("¿Eliminar este lote?")) api.remove(id); };

  return (
    <div className="space-y-5">
      <div className="cc-card p-4">
        <div className="cc-h" style={{ fontSize: 16, fontWeight: 600, marginBottom: 10 }}>Nuevo lote</div>
        <div className="flex gap-3 flex-wrap items-end">
          <div style={{ flex: 1, minWidth: 180 }}><label style={{ fontSize: 12, color: "#8A8570" }}>Nombre</label><input className="cc-input" value={nombre} onChange={(e) => setNombre(e.target.value)} placeholder="Lote 4 - La Loma" /></div>
          <div style={{ width: 140 }}><label style={{ fontSize: 12, color: "#8A8570" }}>Hectáreas (opcional)</label><input className="cc-input" type="number" value={hectareas} onChange={(e) => setHectareas(e.target.value)} /></div>
          <button className="cc-btn cc-btn-primary" onClick={guardar}><Plus size={15} /> Agregar</button>
        </div>
      </div>
      {lotes.length === 0 ? <EmptyState icon={MapPin} title="No hay lotes cargados" text="Los lotes son de referencia: van a aparecer como sugerencia al cargar el 'Campo' de cada remito." /> : (
        <div className="cc-card overflow-hidden">
          <table className="w-full" style={{ fontSize: 13 }}>
            <thead><tr style={{ background: "#EEEADA", textAlign: "left" }}><th className="px-4 py-2">Lote</th><th className="px-4 py-2">Hectáreas</th><th className="px-4 py-2"></th></tr></thead>
            <tbody>{lotes.map((l) => (
              <tr key={l.id} style={{ borderTop: "1px solid var(--line)" }}><td className="px-4 py-2">{l.nombre}</td><td className="px-4 py-2 cc-mono">{l.hectareas ? `${fmt(l.hectareas, 1)} ha` : "-"}</td><td className="px-4 py-2 text-right"><button onClick={() => eliminar(l.id)}><Trash2 size={14} color="var(--rust)" /></button></td></tr>
            ))}</tbody>
          </table>
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Insumos (compras + stock, referencia global)                        */
/* ------------------------------------------------------------------ */
const emptyItemInsumo = () => ({ nombre: "", litros: "", precio: "" });

function InsumosView({ compras, api, stockInsumos, user }) {
  const [fecha, setFecha] = useState("");
  const [origen, setOrigen] = useState("");
  const [items, setItems] = useState([emptyItemInsumo()]);
  const [archivo, setArchivo] = useState(null);
  const [preview, setPreview] = useState(null);
  const [mediaType, setMediaType] = useState("image/jpeg");
  const [imgB64, setImgB64] = useState(null);
  const [extrayendo, setExtrayendo] = useState(false);
  const [subiendo, setSubiendo] = useState(false);
  const [error, setError] = useState("");

  const origenesSugeridos = Array.from(new Set(compras.map((c) => c.origen).filter(Boolean)));
  const nombresSugeridos = Array.from(new Set(compras.map((c) => c.nombre).filter(Boolean)));

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
        { type: "text", text: `Esta imagen es una factura de compra de insumos agropecuarios (agroquímicos, fertilizantes, combustible, etc.) en Uruguay. Devolvé SOLO un JSON (sin texto extra, sin markdown) con: {"origen":"nombre del proveedor","fecha":"YYYY-MM-DD o vacío","items":[{"nombre":"nombre del insumo","litros":number (cantidad total en litros o unidad equivalente),"precio":number (precio TOTAL pagado por esa cantidad, en dólares)}]}. Incluí un objeto en "items" por cada insumo distinto de la factura.` },
      ]);
      setOrigen(data.origen || origen);
      if (data.fecha) setFecha(data.fecha);
      if (Array.isArray(data.items) && data.items.length) setItems(data.items.map((it) => ({ nombre: it.nombre || "", litros: it.litros ?? "", precio: it.precio ?? "" })));
    } catch (e) { setError("No se pudo leer la factura automáticamente. Completá los ítems a mano."); }
    finally { setExtrayendo(false); }
  };

  const guardar = async () => {
    const validos = items.filter((it) => it.nombre && it.litros && it.precio);
    if (!fecha || !origen || !validos.length) { alert("Completá fecha, origen y al menos un insumo con nombre, litros y precio."); return; }
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
    await Promise.all(validos.map((it) => api.add({
      fecha, origen, nombre: it.nombre, litros: Number(it.litros), precio: Number(it.precio),
      facturaUrl, facturaNombre, usuario: user.email,
    })));
    setFecha(""); setOrigen(""); setItems([emptyItemInsumo()]); setArchivo(null); setPreview(null); setImgB64(null);
  };

  const eliminar = (id) => { if (confirm("¿Eliminar esta compra? Esto también reduce el stock disponible registrado.")) api.remove(id); };

  return (
    <div className="space-y-6">
      <div>
        <div className="cc-h" style={{ fontSize: 20, fontWeight: 600 }}>Insumos</div>
        <div style={{ fontSize: 12.5, color: "#8A8570" }}>Registrá las compras de insumos. El stock se descuenta solo cuando lo consumís desde "Gastos" en cada cultivo.</div>
      </div>

      <div className="cc-card p-4 space-y-4">
        <div className="flex flex-wrap gap-2">
          <label className="cc-btn cc-btn-ghost" style={{ cursor: "pointer" }}>
            <Paperclip size={14} /> {archivo ? archivo.name : "Adjuntar factura (imagen o PDF)"}
            <input type="file" accept="image/*,.pdf" capture="environment" onChange={onFile} style={{ display: "none" }} />
          </label>
          {imgB64 && <button className="cc-btn cc-btn-ghost" onClick={extraer} disabled={extrayendo}>{extrayendo ? <Loader2 size={14} className="animate-spin" /> : <Camera size={14} />} Extraer datos con IA</button>}
        </div>
        {preview && <img src={preview} alt="Factura" style={{ maxWidth: 140, borderRadius: 8, border: "1px solid var(--line)" }} />}
        {error && <div style={{ color: "var(--rust)", fontSize: 13, display: "flex", gap: 6, alignItems: "center" }}><AlertCircle size={14} />{error}</div>}

        <div className="grid gap-3" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(160px,1fr))" }}>
          <div><label style={{ fontSize: 12, color: "#8A8570" }}>Fecha</label><input className="cc-input" type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} /></div>
          <div>
            <label style={{ fontSize: 12, color: "#8A8570" }}>Origen (proveedor)</label>
            <input className="cc-input" list="origenes-insumos" value={origen} onChange={(e) => setOrigen(e.target.value)} placeholder="Agromotora, Barraca..." />
            <datalist id="origenes-insumos">{origenesSugeridos.map((o) => <option key={o} value={o} />)}</datalist>
          </div>
        </div>

        <div>
          <div style={{ fontSize: 12, color: "#8A8570", marginBottom: 4 }}>Insumos de esta factura</div>
          <div className="space-y-2">
            {items.map((it, i) => (
              <div key={i} className="flex gap-2 items-end flex-wrap">
                <div style={{ flex: 2, minWidth: 160 }}>
                  <input className="cc-input" list="nombres-insumos" placeholder="Nombre (ej: Glifosato)" value={it.nombre} onChange={(e) => setItem(i, { ...it, nombre: e.target.value })} />
                </div>
                <div style={{ width: 130 }}><input className="cc-input" type="number" placeholder="Litros totales" value={it.litros} onChange={(e) => setItem(i, { ...it, litros: e.target.value })} /></div>
                <div style={{ width: 130 }}><input className="cc-input" type="number" placeholder="Precio total U$S" value={it.precio} onChange={(e) => setItem(i, { ...it, precio: e.target.value })} /></div>
                {items.length > 1 && <button onClick={() => quitarItem(i)}><X size={16} color="var(--rust)" /></button>}
              </div>
            ))}
            <datalist id="nombres-insumos">{nombresSugeridos.map((n) => <option key={n} value={n} />)}</datalist>
          </div>
          <button className="cc-btn cc-btn-ghost mt-2" onClick={agregarItem}><Plus size={14} /> Agregar otro insumo a esta factura</button>
        </div>

        <button className="cc-btn cc-btn-primary" onClick={guardar} disabled={subiendo}>{subiendo ? <Loader2 size={15} className="animate-spin" /> : <Plus size={15} />} Guardar compra</button>
      </div>

      <div>
        <div className="cc-h" style={{ fontSize: 16, fontWeight: 600, marginBottom: 10 }}>Stock actual</div>
        {stockInsumos.length === 0 ? <EmptyState icon={Boxes} title="Sin insumos cargados" text="Cuando registres una compra, el stock disponible va a aparecer acá." /> : (
          <div className="cc-card overflow-hidden">
            <table className="w-full" style={{ fontSize: 13 }}>
              <thead><tr style={{ background: "#EEEADA", textAlign: "left" }}><th className="px-3 py-2">Insumo</th><th className="px-3 py-2 text-right">Comprado</th><th className="px-3 py-2 text-right">Consumido</th><th className="px-3 py-2 text-right">Disponible</th><th className="px-3 py-2 text-right">Costo prom. / L</th></tr></thead>
              <tbody>
                {stockInsumos.sort((a, b) => a.nombre.localeCompare(b.nombre)).map((i) => (
                  <tr key={i.nombre} style={{ borderTop: "1px solid var(--line)" }}>
                    <td className="px-3 py-2">{i.nombre}</td>
                    <td className="px-3 py-2 text-right cc-mono">{fmt(i.litrosComprados, 1)} L</td>
                    <td className="px-3 py-2 text-right cc-mono">{fmt(i.litrosConsumidos, 1)} L</td>
                    <td className="px-3 py-2 text-right cc-mono" style={{ color: i.disponible < 0 ? "var(--rust)" : "var(--soil-light)", fontWeight: 700 }}>{fmt(i.disponible, 1)} L</td>
                    <td className="px-3 py-2 text-right cc-mono">{fmtUSD(i.costoPromedioPorLitro)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      <div>
        <div className="cc-h" style={{ fontSize: 16, fontWeight: 600, marginBottom: 10 }}>Historial de compras</div>
        {compras.length === 0 ? <EmptyState icon={Receipt} title="Sin compras registradas" text="Registrá tu primera compra de insumos arriba." /> : (
          <div className="cc-card overflow-hidden">
            <table className="w-full" style={{ fontSize: 12.5 }}>
              <thead><tr style={{ background: "#EEEADA", textAlign: "left" }}><th className="px-3 py-2">Fecha</th><th className="px-3 py-2">Origen</th><th className="px-3 py-2">Insumo</th><th className="px-3 py-2 text-right">Litros</th><th className="px-3 py-2 text-right">Precio</th><th className="px-3 py-2"></th><th className="px-3 py-2"></th></tr></thead>
              <tbody>
                {[...compras].sort((a, b) => (b.fecha || "").localeCompare(a.fecha || "")).map((c) => (
                  <tr key={c.id} style={{ borderTop: "1px solid var(--line)" }}>
                    <td className="px-3 py-2 cc-mono">{c.fecha}</td>
                    <td className="px-3 py-2">{c.origen}</td>
                    <td className="px-3 py-2">{c.nombre}</td>
                    <td className="px-3 py-2 text-right cc-mono">{fmt(c.litros, 1)} L</td>
                    <td className="px-3 py-2 text-right cc-mono">{fmtUSD(c.precio)}</td>
                    <td className="px-3 py-2">{c.facturaUrl && <a href={c.facturaUrl} target="_blank" rel="noreferrer"><FileText size={14} color="var(--frost)" /></a>}</td>
                    <td className="px-3 py-2 text-right"><button onClick={() => eliminar(c.id)}><Trash2 size={13} color="var(--rust)" /></button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}
