import { useState, useEffect, useRef } from "react";
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js";
import { getFirestore, collection, onSnapshot, doc, setDoc, deleteDoc } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";
import * as XLSX from "xlsx";

// ── Firebase config ───────────────────────────────────────────────────────────
const firebaseConfig = {
  apiKey: "AIzaSyBmim1GqYlMvyulfkthxmNPSLH9hEzFQNw",
  authDomain: "palacio-del-pie.firebaseapp.com",
  projectId: "palacio-del-pie",
  storageBucket: "palacio-del-pie.firebasestorage.app",
  messagingSenderId: "1068124921123",
  appId: "1:1068124921123:web:5c90ee74961562450c8096"
};
const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

// ── Contraseña ────────────────────────────────────────────────────────────────
const ACCESS_PASSWORD = "palaciodelpie";

const CIRUGIAS_GRUPOS = {
  "Hallux": ["Hallux valgus (osteotomía chevron)","Hallux valgus (osteotomía Scarf)","Hallux valgus (osteotomía Lapidus)","Hallux rigidus (queilectomía)","Hallux rigidus (artrodesis 1ª MTF)","Sesamoidectomía"],
  "Dedos menores": ["Dedo en martillo (proximal)","Dedo en garra","Dedo en maza","Metatarsalgia (osteotomía Weil)","Neuroma de Morton","Sinovitis 2ª MTF","Amputación dedo"],
  "Tendones": ["Reparación tendón Aquiles (aguda)","Reparación tendón Aquiles (crónica)","Tenotomía / alargamiento Aquiles","Reparación tendón tibial posterior","Reparación tendón peroneo","Tenolisis","Transferencia tendinosa"],
  "Pie": ["Fasciotomía plantar (endoscópica)","Fasciotomía plantar (abierta)","Osteotomía calcáneo (Dwyer)","Osteotomía calcáneo (Evans)","Osteotomía calcáneo (medializante)","Artrodesis mediopié","Artrodesis de Lisfranc","Corrección pie plano adulto","Corrección pie cavo","Resección espolón calcáneo","Resección exostosis dorsal","Bursectomía retrocalcánea"],
  "Tobillo": ["Artrodesis tobillo","Artroscopia tobillo (diagnóstica)","Artroscopia tobillo (terapéutica)","Ligamentoplastia tobillo (Broström)","Ligamentoplastia tobillo (Broström-Gould)","Reparación sindesmosis","Osteocondral tobillo (mosaicoplastia)","Osteocondral tobillo (microfracturas)","Prótesis total de tobillo","Revisión prótesis tobillo"],
  "Fracturas": ["Fractura maleolo lateral","Fractura maleolo medial","Fractura bimaleolar","Fractura trimaleolar","Fractura calcáneo (RAFI)","Fractura astrágalo","Fractura 5º metatarsiano","Fractura otros metatarsianos","Fractura falange","Retirada material osteosíntesis"],
  "Otras": ["Ganglio / quiste sinovial","Tumor partes blandas","Desbridamiento infección","Artrodesis subtalar","Artrorisis subtalar","Otra intervención"],
};

const ESTADOS = ["Pendiente fecha","Programado","Operado","Alta"];
const USUARIOS = [
  { nombre: "Dr. Contreras", icono: "👨‍⚕️" },
  { nombre: "Dr. Díez Saralegui", icono: "👨‍⚕️" },
  { nombre: "Dra. Jiménez", icono: "👩‍⚕️" },
  { nombre: "Irene", icono: "🗂️" },
];
const CIRUJANOS = USUARIOS.filter(u => u.nombre !== "Irene").map(u => u.nombre);

const initialForm = {
  nombre: "", apellidos: "", numHistoria: "", edad: "",
  telefono: "", email: "", cirugia: "", codigoCirugia: "",
  fechaCirugia: "", estado: "Pendiente fecha", pagado: false,
  montoPagado: "", montoTotal: "", notas: "",
  creadoPor: "", fechaCreacion: new Date().toISOString(),
  adjuntos: [],
};

const generateId = () => Date.now().toString(36) + Math.random().toString(36).slice(2);
const formatDate = (iso) => { if (!iso) return "—"; return new Date(iso).toLocaleDateString("es-ES", { day: "2-digit", month: "2-digit", year: "numeric" }); };

async function extractFromImage(base64Data, mediaType) {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ model: "claude-sonnet-4-20250514", max_tokens: 600, messages: [{ role: "user", content: [{ type: "image", source: { type: "base64", media_type: mediaType, data: base64Data } }, { type: "text", text: `Analiza esta imagen de una historia clínica española. Extrae: nombre, apellidos, numHistoria, edad, telefono, email. Devuelve SOLO JSON válido sin markdown. Campos no encontrados como "". Formato: {"nombre":"","apellidos":"","numHistoria":"","edad":"","telefono":"","email":""}` }] }] })
  });
  if (!res.ok) { const e = await res.json().catch(() => ({})); throw new Error(e?.error?.message || `HTTP ${res.status}`); }
  const data = await res.json();
  const raw = (data.content || []).map(b => b.text || "").join("").trim();
  return JSON.parse(raw.replace(/^```json\s*/i,"").replace(/^```\s*/i,"").replace(/```\s*$/i,"").trim());
}

function exportExcel(pacientes) {
  const headers = ["Nombre","Apellidos","Nº Historia","Edad","Teléfono","Email","Cirugía","Código","Fecha Cirugía","Estado","Pagado","Total (€)","Pagado (€)","Cirujano","Notas"];
  const rows = pacientes.map(p => [
    p.nombre||"", p.apellidos||"", p.numHistoria||"",
    p.edad ? Number(p.edad) : "",
    p.telefono||"", p.email||"",
    p.cirugia||"", p.codigoCirugia||"",
    p.fechaCirugia ? formatDate(p.fechaCirugia) : "",
    p.estado||"", p.pagado?"Sí":"No",
    p.montoTotal ? Number(p.montoTotal) : "",
    p.montoPagado ? Number(p.montoPagado) : "",
    p.creadoPor||"", (p.notas||"").replace(/\n/g," ")
  ]);
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws["!cols"] = [{wch:15},{wch:20},{wch:13},{wch:6},{wch:14},{wch:26},{wch:36},{wch:16},{wch:14},{wch:14},{wch:8},{wch:10},{wch:10},{wch:20},{wch:40}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Pacientes");
  XLSX.writeFile(wb, `palacio-del-pie-${new Date().toISOString().slice(0,10)}.xlsx`);
}

export default function App() {
  const [autenticado, setAutenticado] = useState(false);
  const [password, setPassword] = useState("");
  const [passwordError, setPasswordError] = useState(false);
  const [showPassword, setShowPassword] = useState(false);
  const [pacientes, setPacientes] = useState([]);
  const [vista, setVista] = useState("lista");
  const [selected, setSelected] = useState(null);
  const [form, setForm] = useState(initialForm);
  const [usuario, setUsuario] = useState("");
  const [filtro, setFiltro] = useState("");
  const [filtroPago, setFiltroPago] = useState("todos");
  const [filtroEstado, setFiltroEstado] = useState("todos");
  const [filtroCirujano, setFiltroCirujano] = useState("todos");
  const [filtroMes, setFiltroMes] = useState("todos");
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [toast, setToast] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [scanState, setScanState] = useState("idle");
  const [scanPreview, setScanPreview] = useState(null);
  const [scanError, setScanError] = useState("");
  const scanRef = useRef();
  const adjuntoRef = useRef();

  // ── Escucha en tiempo real de Firestore ───────────────────────────────────
  useEffect(() => {
    const unsub = onSnapshot(collection(db, "pacientes"), (snap) => {
      const data = snap.docs.map(d => ({ id: d.id, ...d.data() }));
      data.sort((a, b) => new Date(b.fechaCreacion) - new Date(a.fechaCreacion));
      setPacientes(data);
      setLoading(false);
    }, (err) => {
      console.error("Firestore error:", err);
      setLoading(false);
    });
    return () => unsub();
  }, []);

  async function saveData(paciente) {
    setSaving(true);
    try { await setDoc(doc(db, "pacientes", paciente.id), paciente); }
    catch (e) { showToast("Error al guardar", "error"); console.error(e); }
    setSaving(false);
  }

  async function deleteData(id) {
    try { await deleteDoc(doc(db, "pacientes", id)); }
    catch (e) { showToast("Error al eliminar", "error"); console.error(e); }
  }

  function showToast(msg, type = "ok") { setToast({ msg, type }); setTimeout(() => setToast(null), 4000); }

  function handleNuevo() { setForm({ ...initialForm, adjuntos: [], creadoPor: usuario }); setScanPreview(null); setScanState("idle"); setScanError(""); setVista("nuevo"); }
  function handleEditar(p) { setForm({ ...p, adjuntos: p.adjuntos || [] }); setSelected(p); setScanPreview(null); setScanState("idle"); setScanError(""); setVista("editar"); }
  function handleVer(p) { setSelected(p); setVista("detalle"); }

  async function handleGuardar() {
    if (!form.nombre?.trim() || !form.apellidos?.trim()) { showToast("Nombre y apellidos son obligatorios", "error"); return; }
    if (!form.cirugia) { showToast("Selecciona el tipo de cirugía", "error"); return; }
    const paciente = vista === "nuevo" ? { ...form, id: generateId(), fechaCreacion: new Date().toISOString() } : { ...form };
    await saveData(paciente);
    showToast(vista === "nuevo" ? "Paciente añadido ✓" : "Paciente actualizado ✓");
    setVista("lista");
  }

  async function handleDelete(id) {
    await deleteData(id);
    setConfirmDelete(null); setVista("lista"); showToast("Paciente eliminado");
  }

  async function togglePago(p) {
    await saveData({ ...p, pagado: !p.pagado });
    showToast(!p.pagado ? "Marcado como pagado ✓" : "Marcado como pendiente");
  }

  function handleScanClick() { scanRef.current?.click(); }

  async function handleScanFile(e) {
    const file = e.target.files?.[0]; e.target.value = "";
    if (!file) return;
    if (!["image/jpeg","image/png","image/webp","image/gif"].includes(file.type)) { showToast("Usa JPG, PNG o WEBP", "error"); return; }
    if (file.size > 5*1024*1024) { showToast("Imagen demasiado grande (máx. 5 MB)", "error"); return; }
    setScanState("loading"); setScanError(""); setScanPreview(null);
    try {
      const dataUrl = await readFileAsDataURL(file);
      setScanPreview(dataUrl);
      const extracted = await extractFromImage(dataUrl.split(",")[1], file.type);
      let filled = 0;
      setForm(prev => { const u={...prev}; ["nombre","apellidos","numHistoria","edad","telefono","email"].forEach(k=>{ if(extracted[k]&&!prev[k]){u[k]=extracted[k];filled++;} }); return u; });
      if (filled === 0) { setScanState("error"); setScanError("No se detectaron datos. Rellena manualmente."); }
      else { setScanState("done"); showToast(`${filled} campos extraídos ✓`); }
    } catch (err) { setScanState("error"); setScanError(`Error: ${err.message||"inténtalo de nuevo"}`); }
  }

  function handleAdjuntoClick() { adjuntoRef.current?.click(); }

  async function handleAdjuntoFile(e) {
    const files = Array.from(e.target.files||[]); e.target.value = "";
    const nuevos = [];
    for (const file of files) {
      if (file.size > 8*1024*1024) { showToast(`${file.name} supera 8 MB`,"error"); continue; }
      try { nuevos.push({ id: generateId(), nombre: file.name, tipo: file.type, dataUrl: await readFileAsDataURL(file), fecha: new Date().toISOString() }); }
      catch { showToast(`Error leyendo ${file.name}`,"error"); }
    }
    if (nuevos.length) { setForm(prev=>({...prev,adjuntos:[...(prev.adjuntos||[]),...nuevos]})); showToast(`${nuevos.length} archivo(s) adjuntado(s) ✓`); }
  }

  function removeAdjunto(id) { setForm(prev=>({...prev,adjuntos:prev.adjuntos.filter(a=>a.id!==id)})); }
  function downloadAdjunto(adj) { const a=document.createElement("a"); a.href=adj.dataUrl; a.download=adj.nombre; a.click(); }
  function readFileAsDataURL(file) { return new Promise((res,rej)=>{ const r=new FileReader(); r.onload=ev=>res(ev.target.result); r.onerror=()=>rej(new Error("Error")); r.readAsDataURL(file); }); }

  const filtrados = pacientes.filter(p => {
    const txt = `${p.nombre} ${p.apellidos} ${p.numHistoria} ${p.codigoCirugia} ${p.cirugia}`.toLowerCase();
    return txt.includes(filtro.toLowerCase()) && (filtroPago==="todos"||(filtroPago==="pagado"?p.pagado:!p.pagado)) && (filtroEstado==="todos"||p.estado===filtroEstado) && (filtroCirujano==="todos"||p.creadoPor===filtroCirujano) && (filtroMes==="todos"||(p.fechaCirugia&&p.fechaCirugia.slice(0,7)===filtroMes));
  });

  const stats = { total: pacientes.length, programados: pacientes.filter(p=>p.estado==="Programado").length, operados: pacientes.filter(p=>p.estado==="Operado"||p.estado==="Alta").length, pendientesPago: pacientes.filter(p=>!p.pagado).length };
  const mesesDisponibles = [...new Set(pacientes.filter(p=>p.fechaCirugia).map(p=>p.fechaCirugia.slice(0,7)))].sort().reverse();

  // ── PANTALLA CONTRASEÑA ───────────────────────────────────────────────────
  if (!autenticado) return (
    <div style={S.loginWrap}>
      <div style={S.loginCard}>
        <div style={S.brand2}><span style={S.brandFoot}>🦶</span><div><div style={S.brandTitle}>Palacio del Pie</div><div style={S.brandSub}>Gestión quirúrgica</div></div></div>
        <p style={S.loginLabel}>Introduce la contraseña</p>
        <div style={{display:"flex",gap:8,marginBottom:8}}>
          <input style={{flex:1,padding:"10px 14px",border:`1.5px solid ${passwordError?"#ef4444":"#e2e8f0"}`,borderRadius:9,fontSize:15,outline:"none",letterSpacing:"0.1em"}} type={showPassword?"text":"password"} value={password} onChange={e=>{setPassword(e.target.value);setPasswordError(false);}} onKeyDown={e=>{if(e.key==="Enter"){if(password===ACCESS_PASSWORD){setAutenticado(true);}else{setPasswordError(true);setPassword("");}}}} placeholder="Contraseña..." autoFocus />
          <button style={{padding:"8px 12px",background:"#f8fafc",border:"1.5px solid #e2e8f0",borderRadius:9,cursor:"pointer",fontSize:16}} onClick={()=>setShowPassword(v=>!v)}>{showPassword?"🙈":"👁"}</button>
        </div>
        {passwordError && <div style={{color:"#ef4444",fontSize:13,fontWeight:600,marginBottom:10}}>Contraseña incorrecta</div>}
        <button style={{width:"100%",padding:11,background:"#6366f1",color:"#fff",border:"none",borderRadius:10,fontWeight:700,fontSize:15,cursor:"pointer",marginTop:4}} onClick={()=>{if(password===ACCESS_PASSWORD){setAutenticado(true);}else{setPasswordError(true);setPassword("");}}}>Acceder</button>
      </div>
    </div>
  );

  // ── PANTALLA USUARIO ──────────────────────────────────────────────────────
  if (!usuario) return (
    <div style={S.loginWrap}>
      <div style={S.loginCard}>
        <div style={S.brand2}><span style={S.brandFoot}>🦶</span><div><div style={S.brandTitle}>Palacio del Pie</div><div style={S.brandSub}>Gestión quirúrgica</div></div></div>
        <p style={S.loginLabel}>¿Quién eres?</p>
        <div style={S.usuarioGrid}>
          {USUARIOS.map(u=>(<button key={u.nombre} style={S.usuarioBtn} onClick={()=>setUsuario(u.nombre)}><span style={{fontSize:28}}>{u.icono}</span><span style={{fontSize:13,fontWeight:700}}>{u.nombre}</span></button>))}
        </div>
        <button style={{marginTop:20,background:"none",border:"none",color:"#94a3b8",fontSize:12,cursor:"pointer"}} onClick={()=>setAutenticado(false)}>← Cambiar contraseña</button>
      </div>
    </div>
  );

  if (loading) return (<div style={S.loadingWrap}><div style={S.spinner}/><p style={{color:"#64748b",marginTop:16}}>Cargando pacientes...</p></div>);

  const isForm = vista==="nuevo"||vista==="editar";

  return (
    <div style={S.app}>
      {toast && <div style={{...S.toast,background:toast.type==="error"?"#ef4444":"#10b981"}}>{toast.msg}</div>}
      {confirmDelete && (<div style={S.overlay}><div style={S.confirmCard}><p style={{fontWeight:700,fontSize:16,marginBottom:6}}>¿Eliminar este paciente?</p><p style={{color:"#64748b",fontSize:14,marginBottom:22}}>Esta acción no se puede deshacer.</p><div style={{display:"flex",gap:10}}><button style={S.btnDanger} onClick={()=>handleDelete(confirmDelete)}>Eliminar</button><button style={S.btnSecondary} onClick={()=>setConfirmDelete(null)}>Cancelar</button></div></div></div>)}

      <header style={S.header}>
        <div style={S.headerLeft}><span style={{fontSize:22}}>🦶</span><span style={S.brandH}>Palacio del Pie</span>{saving&&<span style={S.savingBadge}>Guardando...</span>}</div>
        <div style={S.headerRight}><span style={S.userBadge}>{usuario}</span><button style={S.btnLogout} onClick={()=>setUsuario("")}>Salir</button></div>
      </header>

      <main style={S.main}>
        {vista==="lista" && (<>
          <div style={S.statsRow}>
            {[{label:"Total pacientes",value:stats.total,color:"#6366f1"},{label:"Programados",value:stats.programados,color:"#f59e0b"},{label:"Operados / Alta",value:stats.operados,color:"#10b981"},{label:"Pago pendiente",value:stats.pendientesPago,color:"#ef4444"}].map(s=>(
              <div key={s.label} style={{...S.statCard,borderTop:`3px solid ${s.color}`}}><div style={{...S.statNum,color:s.color}}>{s.value}</div><div style={S.statLabel}>{s.label}</div></div>
            ))}
          </div>
          <div style={S.toolbar}>
            <input style={S.search} placeholder="Buscar por nombre, nº historia, código..." value={filtro} onChange={e=>setFiltro(e.target.value)}/>
            <select style={S.select} value={filtroCirujano} onChange={e=>setFiltroCirujano(e.target.value)}><option value="todos">Todos los cirujanos</option>{CIRUJANOS.map(c=><option key={c}>{c}</option>)}</select>
            <select style={S.select} value={filtroEstado} onChange={e=>setFiltroEstado(e.target.value)}><option value="todos">Todos los estados</option>{ESTADOS.map(e=><option key={e}>{e}</option>)}</select>
            <select style={S.select} value={filtroPago} onChange={e=>setFiltroPago(e.target.value)}><option value="todos">Todos los pagos</option><option value="pagado">Pagados</option><option value="pendiente">Pago pendiente</option></select>
            <select style={S.select} value={filtroMes} onChange={e=>setFiltroMes(e.target.value)}><option value="todos">Todos los meses</option>{mesesDisponibles.map(m=>{const[y,mo]=m.split("-");const lbl=new Date(+y,+mo-1,1).toLocaleDateString("es-ES",{month:"long",year:"numeric"});return<option key={m} value={m}>{lbl.charAt(0).toUpperCase()+lbl.slice(1)}</option>;})}</select>
            <button style={S.btnExport} onClick={()=>exportExcel(filtrados)}>⬇ Exportar Excel</button>
            <button style={S.btnPrimary} onClick={handleNuevo}>+ Nuevo paciente</button>
          </div>
          {filtrados.length===0 ? (<div style={S.empty}><div style={{fontSize:48}}>🏥</div><p style={{color:"#94a3b8",marginTop:8}}>No hay pacientes con estos filtros.</p></div>) : (
            <div style={S.tableWrap}>
              <table style={S.table}>
                <thead><tr>{["Paciente","Nº Historia","Edad","Cirugía","F. Cirugía","Estado","Cirujano","Pago",""].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
                <tbody>
                  {filtrados.map(p=>(
                    <tr key={p.id} style={S.tr}>
                      <td style={S.td}><div style={{fontWeight:700,color:"#1e293b"}}>{p.nombre} {p.apellidos}</div><div style={{fontSize:12,color:"#94a3b8"}}>{p.telefono}</div></td>
                      <td style={S.td}><span style={S.mono}>{p.numHistoria||"—"}</span></td>
                      <td style={S.td}>{p.edad?`${p.edad}a`:"—"}</td>
                      <td style={S.td}><span style={S.cirugiaBadge}>{p.cirugia}</span></td>
                      <td style={S.td}>{formatDate(p.fechaCirugia)}</td>
                      <td style={S.td}><span style={{...S.estadoBadge,...estadoColor(p.estado)}}>{p.estado}</span></td>
                      <td style={S.td}><span style={S.cirujanoChip}>{p.creadoPor||"—"}</span></td>
                      <td style={S.td}><button style={{...S.pagoBadge,...(p.pagado?S.pagadoOn:S.pagadoOff)}} onClick={()=>togglePago(p)}>{p.pagado?"✓ Pagado":"⏳ Pendiente"}</button></td>
                      <td style={S.td}><div style={{display:"flex",gap:5}}><button style={S.btnIcon} onClick={()=>handleVer(p)}>👁</button><button style={S.btnIcon} onClick={()=>handleEditar(p)}>✏️</button><button style={S.btnIcon} onClick={()=>setConfirmDelete(p.id)}>🗑</button></div></td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <div style={S.tableFooter}>{filtrados.length} paciente{filtrados.length!==1?"s":""} mostrado{filtrados.length!==1?"s":""}</div>
            </div>
          )}
        </>)}

        {vista==="detalle" && selected && (
          <div style={S.formCard}>
            <div style={S.formHeader}><button style={S.backBtn} onClick={()=>setVista("lista")}>← Volver</button><h2 style={S.formTitle}>{selected.nombre} {selected.apellidos}</h2><button style={S.btnPrimary} onClick={()=>handleEditar(selected)}>Editar</button></div>
            <div style={S.detalleGrid}>
              <DetalleItem label="Nº Historia" value={selected.numHistoria} accent/><DetalleItem label="Edad" value={selected.edad?`${selected.edad} años`:null}/><DetalleItem label="Teléfono" value={selected.telefono}/><DetalleItem label="Email" value={selected.email}/>
              <DetalleItem label="Cirugía" value={selected.cirugia} accent/><DetalleItem label="Código" value={selected.codigoCirugia}/><DetalleItem label="Fecha cirugía" value={formatDate(selected.fechaCirugia)}/><DetalleItem label="Estado" value={selected.estado}/>
              <DetalleItem label="Cirujano" value={selected.creadoPor}/><DetalleItem label="Pago" value={selected.pagado?"✅ Pagado":"⏳ Pendiente"}/><DetalleItem label="Total" value={selected.montoTotal?`${selected.montoTotal} €`:null}/><DetalleItem label="Pagado" value={selected.montoPagado?`${selected.montoPagado} €`:null}/>
            </div>
            {selected.notas&&<div style={S.notasBox}><div style={S.notasLabel}>Notas clínicas</div><div style={S.notasText}>{selected.notas}</div></div>}
            {(selected.adjuntos||[]).length>0&&<div style={{marginTop:20}}><div style={S.sectionTitle}>Documentos adjuntos</div><div style={S.adjuntosGrid}>{selected.adjuntos.map(a=><AdjuntoCard key={a.id} adj={a} onDownload={()=>downloadAdjunto(a)}/>)}</div></div>}
          </div>
        )}

        {isForm && (
          <div style={S.formCard}>
            <div style={S.formHeader}><button style={S.backBtn} onClick={()=>setVista("lista")}>← Volver</button><h2 style={S.formTitle}>{vista==="nuevo"?"Nuevo paciente":"Editar paciente"}</h2></div>
            <div style={S.scanSection}>
              <div style={{flex:1}}>
                <div style={S.scanTitle}>📷 Escaneado automático desde foto</div>
                <div style={S.scanDesc}>Haz una foto a la portada de la historia clínica y la IA extraerá los datos automáticamente.</div>
                <input ref={scanRef} type="file" accept="image/jpeg,image/png,image/webp" capture="environment" style={{display:"none"}} onChange={handleScanFile}/>
                <button style={{...S.btnScan,opacity:scanState==="loading"?0.65:1,cursor:scanState==="loading"?"wait":"pointer"}} onClick={handleScanClick} disabled={scanState==="loading"}>
                  {scanState==="loading"?"⏳ Analizando...":scanState==="done"?"✓ Extraído — subir otra":"📷 Subir foto de la historia"}
                </button>
                {scanState==="error"&&<div style={S.scanMsg}>{scanError}</div>}
                {scanState==="done"&&<div style={{...S.scanMsg,color:"#16a34a"}}>✓ Revisa los campos y completa lo que falte.</div>}
              </div>
              {scanPreview&&<img src={scanPreview} alt="" style={S.scanPreview}/>}
            </div>
            <SectionTitle>Datos personales</SectionTitle>
            <div style={S.formGrid}>
              <FormField label="Nombre *" value={form.nombre} onChange={v=>setForm({...form,nombre:v})}/>
              <FormField label="Apellidos *" value={form.apellidos} onChange={v=>setForm({...form,apellidos:v})}/>
              <FormField label="Nº Historia clínica" value={form.numHistoria} onChange={v=>setForm({...form,numHistoria:v})} placeholder="123456"/>
              <FormField label="Edad (años)" value={form.edad} onChange={v=>setForm({...form,edad:v})} type="number" placeholder="58"/>
              <FormField label="Teléfono" value={form.telefono} onChange={v=>setForm({...form,telefono:v})} type="tel"/>
              <FormField label="Email" value={form.email} onChange={v=>setForm({...form,email:v})} type="email"/>
            </div>
            <SectionTitle>Datos quirúrgicos</SectionTitle>
            <div style={S.formGrid}>
              <div style={S.fieldWrap}><label style={S.label}>Tipo de cirugía *</label><select style={S.input} value={form.cirugia} onChange={e=>setForm({...form,cirugia:e.target.value})}><option value="">Seleccionar...</option>{Object.entries(CIRUGIAS_GRUPOS).map(([g,items])=><optgroup key={g} label={`── ${g}`}>{items.map(c=><option key={c}>{c}</option>)}</optgroup>)}</select></div>
              <FormField label="Código cirugía" value={form.codigoCirugia} onChange={v=>setForm({...form,codigoCirugia:v})} placeholder="HAV-2025-001"/>
              <FormField label="Fecha cirugía" value={form.fechaCirugia} onChange={v=>setForm({...form,fechaCirugia:v})} type="date"/>
              <div style={S.fieldWrap}><label style={S.label}>Estado</label><select style={S.input} value={form.estado} onChange={e=>setForm({...form,estado:e.target.value})}>{ESTADOS.map(e=><option key={e}>{e}</option>)}</select></div>
              <div style={S.fieldWrap}><label style={S.label}>Cirujano responsable</label><select style={S.input} value={form.creadoPor||usuario} onChange={e=>setForm({...form,creadoPor:e.target.value})}>{USUARIOS.map(u=><option key={u.nombre}>{u.nombre}</option>)}</select></div>
            </div>
            <SectionTitle>Información de pago</SectionTitle>
            <div style={S.formGrid}>
              <FormField label="Monto total (€)" value={form.montoTotal} onChange={v=>setForm({...form,montoTotal:v})} type="number"/>
              <FormField label="Monto pagado (€)" value={form.montoPagado} onChange={v=>setForm({...form,montoPagado:v})} type="number"/>
              <div style={S.fieldWrap}><label style={S.label}>Estado pago</label><button style={{...S.toggleBtn,...(form.pagado?S.toggleOn:S.toggleOff)}} onClick={()=>setForm({...form,pagado:!form.pagado})}>{form.pagado?"✓ Pagado":"⏳ Pendiente de pago"}</button></div>
            </div>
            <SectionTitle>Notas clínicas</SectionTitle>
            <textarea style={{...S.input,minHeight:80,resize:"vertical"}} value={form.notas} onChange={e=>setForm({...form,notas:e.target.value})} placeholder="Alergias, observaciones, indicaciones postoperatorias..."/>
            <SectionTitle>Documentos adjuntos</SectionTitle>
            <input ref={adjuntoRef} type="file" multiple accept=".pdf,.jpg,.jpeg,.png,.webp,.doc,.docx" style={{display:"none"}} onChange={handleAdjuntoFile}/>
            <button style={S.btnAdjunto} onClick={handleAdjuntoClick}>📎 Adjuntar archivos (PDF, imágenes, Word...)</button>
            {(form.adjuntos||[]).length>0&&<div style={S.adjuntosGrid}>{form.adjuntos.map(a=><AdjuntoCard key={a.id} adj={a} onDownload={()=>downloadAdjunto(a)} onRemove={()=>removeAdjunto(a.id)}/>)}</div>}
            <div style={S.formActions}>
              <button style={S.btnPrimary} onClick={handleGuardar}>{saving?"Guardando...":vista==="nuevo"?"Guardar paciente":"Actualizar paciente"}</button>
              <button style={S.btnSecondary} onClick={()=>setVista("lista")}>Cancelar</button>
              {vista==="editar"&&<button style={S.btnDanger} onClick={()=>setConfirmDelete(form.id)}>Eliminar paciente</button>}
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

function SectionTitle({children}){return <div style={S.sectionTitle}>{children}</div>;}
function DetalleItem({label,value,accent}){return(<div style={S.detalleItem}><div style={S.detalleLabel}>{label}</div><div style={{...S.detalleValue,color:accent?"#6366f1":"#1e293b",fontWeight:accent?700:500}}>{value||"—"}</div></div>);}
function FormField({label,value,onChange,type="text",placeholder}){return(<div style={S.fieldWrap}><label style={S.label}>{label}</label><input style={S.input} type={type} value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder||""}/></div>);}
function AdjuntoCard({adj,onDownload,onRemove}){const isImg=adj.tipo?.startsWith("image/");const icon=adj.tipo==="application/pdf"?"📄":isImg?"🖼️":"📎";return(<div style={S.adjuntoCard}>{isImg&&<img src={adj.dataUrl} alt={adj.nombre} style={S.adjuntoThumb}/>}<div style={S.adjuntoInfo}><span style={S.adjuntoIcon}>{icon}</span><span style={S.adjuntoNombre} title={adj.nombre}>{adj.nombre}</span></div><div style={{display:"flex",gap:4}}><button style={S.adjuntoBtn} onClick={onDownload}>⬇</button>{onRemove&&<button style={{...S.adjuntoBtn,color:"#dc2626"}} onClick={onRemove}>✕</button>}</div></div>);}
function estadoColor(e){return({"Pendiente fecha":{background:"#fef3c7",color:"#92400e"},"Programado":{background:"#dbeafe",color:"#1e40af"},"Operado":{background:"#dcfce7",color:"#166534"},"Alta":{background:"#f0fdf4",color:"#15803d"}})[e]||{};}

const S = {
  app:{minHeight:"100vh",background:"#f8fafc",fontFamily:"'DM Sans','Segoe UI',sans-serif"},
  loginWrap:{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"linear-gradient(135deg,#e0e7ff 0%,#f0f9ff 100%)"},
  loginCard:{background:"#fff",borderRadius:20,padding:"44px 40px",boxShadow:"0 20px 60px rgba(99,102,241,.15)",textAlign:"center",maxWidth:420,width:"100%"},
  brand2:{display:"flex",alignItems:"center",gap:14,justifyContent:"center",marginBottom:28},
  brandFoot:{fontSize:48},brandTitle:{fontSize:26,fontWeight:900,color:"#1e293b",letterSpacing:"-0.5px",textAlign:"left"},
  brandSub:{fontSize:13,color:"#6366f1",fontWeight:600,textAlign:"left"},
  loginLabel:{fontWeight:600,color:"#475569",marginBottom:16},
  usuarioGrid:{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12},
  usuarioBtn:{padding:"18px 10px",border:"2px solid #e2e8f0",borderRadius:12,background:"#fff",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:8},
  loadingWrap:{minHeight:"100vh",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center"},
  spinner:{width:36,height:36,border:"3px solid #e2e8f0",borderTop:"3px solid #6366f1",borderRadius:"50%",animation:"spin .8s linear infinite"},
  header:{background:"#fff",borderBottom:"1px solid #e2e8f0",padding:"0 24px",height:56,display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:10},
  headerLeft:{display:"flex",alignItems:"center",gap:10},brandH:{fontWeight:900,fontSize:18,color:"#1e293b",letterSpacing:"-0.4px"},
  savingBadge:{fontSize:12,color:"#6366f1",background:"#eef2ff",padding:"2px 8px",borderRadius:20},
  headerRight:{display:"flex",alignItems:"center",gap:10},
  userBadge:{fontSize:13,fontWeight:600,color:"#475569",background:"#f1f5f9",padding:"4px 12px",borderRadius:20},
  btnLogout:{fontSize:12,color:"#94a3b8",background:"none",border:"none",cursor:"pointer"},
  main:{maxWidth:1240,margin:"0 auto",padding:"24px 16px"},
  statsRow:{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:22},
  statCard:{background:"#fff",borderRadius:12,padding:"16px 18px",boxShadow:"0 1px 4px rgba(0,0,0,.06)"},
  statNum:{fontSize:30,fontWeight:800,lineHeight:1},statLabel:{fontSize:12,color:"#64748b",marginTop:4},
  toolbar:{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap"},
  search:{flex:1,minWidth:180,padding:"8px 13px",border:"1.5px solid #e2e8f0",borderRadius:9,fontSize:14,outline:"none",background:"#fff"},
  select:{padding:"8px 12px",border:"1.5px solid #e2e8f0",borderRadius:9,fontSize:13,background:"#fff",cursor:"pointer"},
  btnPrimary:{padding:"8px 18px",background:"#6366f1",color:"#fff",border:"none",borderRadius:9,fontWeight:700,fontSize:13,cursor:"pointer"},
  btnSecondary:{padding:"8px 18px",background:"#f1f5f9",color:"#475569",border:"none",borderRadius:9,fontWeight:600,fontSize:13,cursor:"pointer"},
  btnDanger:{padding:"8px 18px",background:"#fee2e2",color:"#dc2626",border:"none",borderRadius:9,fontWeight:600,fontSize:13,cursor:"pointer"},
  btnExport:{padding:"8px 16px",background:"#f0fdf4",color:"#15803d",border:"1.5px solid #bbf7d0",borderRadius:9,fontWeight:700,fontSize:13,cursor:"pointer"},
  btnIcon:{padding:"5px 9px",background:"#f8fafc",border:"1px solid #e2e8f0",borderRadius:7,cursor:"pointer",fontSize:13},
  tableWrap:{background:"#fff",borderRadius:12,boxShadow:"0 1px 4px rgba(0,0,0,.06)",overflow:"auto"},
  table:{width:"100%",borderCollapse:"collapse",fontSize:13},
  th:{padding:"11px 14px",textAlign:"left",fontWeight:700,color:"#64748b",fontSize:11,textTransform:"uppercase",letterSpacing:".05em",borderBottom:"1px solid #f1f5f9",whiteSpace:"nowrap"},
  tr:{borderBottom:"1px solid #f8fafc"},td:{padding:"11px 14px",verticalAlign:"middle"},
  mono:{fontFamily:"monospace",fontSize:12,color:"#475569"},
  cirugiaBadge:{fontSize:11,background:"#eef2ff",color:"#4338ca",padding:"2px 7px",borderRadius:20,fontWeight:600},
  estadoBadge:{fontSize:11,padding:"2px 8px",borderRadius:20,fontWeight:700},
  cirujanoChip:{fontSize:11,background:"#f1f5f9",color:"#475569",padding:"2px 7px",borderRadius:20,fontWeight:600},
  pagoBadge:{fontSize:11,padding:"3px 9px",borderRadius:20,fontWeight:700,border:"none",cursor:"pointer"},
  pagadoOn:{background:"#dcfce7",color:"#15803d"},pagadoOff:{background:"#fee2e2",color:"#dc2626"},
  tableFooter:{padding:"10px 16px",fontSize:12,color:"#94a3b8",borderTop:"1px solid #f1f5f9"},
  empty:{textAlign:"center",padding:"70px 20px",background:"#fff",borderRadius:12},
  overlay:{position:"fixed",inset:0,background:"rgba(0,0,0,.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:100},
  confirmCard:{background:"#fff",borderRadius:16,padding:"30px 26px",maxWidth:340,width:"90%",boxShadow:"0 20px 60px rgba(0,0,0,.2)"},
  toast:{position:"fixed",bottom:22,right:22,color:"#fff",padding:"11px 18px",borderRadius:11,fontWeight:600,zIndex:200,boxShadow:"0 4px 20px rgba(0,0,0,.15)",fontSize:13},
  formCard:{background:"#fff",borderRadius:14,padding:"26px",boxShadow:"0 1px 4px rgba(0,0,0,.06)",maxWidth:900,margin:"0 auto"},
  formHeader:{display:"flex",alignItems:"center",gap:14,marginBottom:22},
  formTitle:{fontSize:21,fontWeight:800,color:"#1e293b",flex:1,margin:0},
  backBtn:{background:"none",border:"none",color:"#6366f1",fontWeight:700,cursor:"pointer",fontSize:14,padding:0},
  scanSection:{background:"linear-gradient(135deg,#eef2ff,#f0f9ff)",border:"1.5px dashed #a5b4fc",borderRadius:13,padding:"18px 20px",marginBottom:22,display:"flex",alignItems:"center",gap:18},
  scanTitle:{fontWeight:700,color:"#4338ca",fontSize:14,marginBottom:3},
  scanDesc:{fontSize:12,color:"#64748b",marginBottom:10},
  btnScan:{padding:"8px 16px",background:"#6366f1",color:"#fff",border:"none",borderRadius:8,fontWeight:700,fontSize:12},
  scanMsg:{fontSize:12,color:"#dc2626",marginTop:7},
  scanPreview:{width:80,height:80,objectFit:"cover",borderRadius:9,border:"2px solid #a5b4fc",flexShrink:0},
  sectionTitle:{fontWeight:700,color:"#6366f1",fontSize:11,textTransform:"uppercase",letterSpacing:".08em",marginBottom:12,marginTop:22,paddingBottom:5,borderBottom:"1px solid #eef2ff"},
  formGrid:{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(230px,1fr))",gap:"12px 18px"},
  fieldWrap:{display:"flex",flexDirection:"column",gap:4},
  label:{fontSize:11,fontWeight:700,color:"#64748b",textTransform:"uppercase",letterSpacing:".04em"},
  input:{padding:"8px 11px",border:"1.5px solid #e2e8f0",borderRadius:8,fontSize:13,outline:"none",background:"#fafafa",width:"100%",boxSizing:"border-box"},
  toggleBtn:{padding:"8px 16px",border:"none",borderRadius:8,fontWeight:700,cursor:"pointer",fontSize:13,textAlign:"left"},
  toggleOn:{background:"#dcfce7",color:"#15803d"},toggleOff:{background:"#fee2e2",color:"#dc2626"},
  btnAdjunto:{padding:"8px 16px",background:"#f8fafc",color:"#475569",border:"1.5px dashed #cbd5e1",borderRadius:8,fontWeight:600,fontSize:13,cursor:"pointer",marginBottom:10},
  adjuntosGrid:{display:"flex",flexWrap:"wrap",gap:8,marginTop:8},
  adjuntoCard:{display:"flex",alignItems:"center",gap:8,background:"#f8fafc",border:"1px solid #e2e8f0",borderRadius:9,padding:"7px 10px",maxWidth:280},
  adjuntoThumb:{width:32,height:32,objectFit:"cover",borderRadius:5,flexShrink:0},
  adjuntoInfo:{flex:1,display:"flex",alignItems:"center",gap:5,overflow:"hidden"},
  adjuntoIcon:{fontSize:16,flexShrink:0},
  adjuntoNombre:{fontSize:12,color:"#334155",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"},
  adjuntoBtn:{padding:"3px 7px",background:"#fff",border:"1px solid #e2e8f0",borderRadius:6,cursor:"pointer",fontSize:12,flexShrink:0},
  formActions:{display:"flex",gap:10,marginTop:28,flexWrap:"wrap"},
  detalleGrid:{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(190px,1fr))",gap:12,marginBottom:18},
  detalleItem:{padding:"11px 13px",background:"#f8fafc",borderRadius:9},
  detalleLabel:{fontSize:10,fontWeight:700,color:"#94a3b8",textTransform:"uppercase",letterSpacing:".06em",marginBottom:3},
  detalleValue:{fontSize:14,color:"#1e293b"},
  notasBox:{background:"#fef9c3",border:"1px solid #fde68a",borderRadius:9,padding:"12px 14px",marginTop:6},
  notasLabel:{fontWeight:700,color:"#92400e",fontSize:11,marginBottom:5},
  notasText:{color:"#451a03",fontSize:13,lineHeight:1.6},
};
