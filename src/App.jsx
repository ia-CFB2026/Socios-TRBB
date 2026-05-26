import { useState, useEffect, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";

// ─────────────────────────────────────────────
// ⚙️  CONFIGURACIÓN — pegar tu URL de Apps Script
// ─────────────────────────────────────────────
const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxQwJtu11HTN4HaGnbDviyUQbhcWAE3Pd3ZS60jUIvg7cvNKhU_E7E7ii8IAfl0mIu4Xw/exec"; // ← pegar URL del deploy aquí
const APP_VERSION = "v3.0";

// Nombres de columna exactos (con \n real)
const COL = {
  SOCIO:    "N° Socio",
  APELLIDO: "Apellido",
  NOMBRE:   "Nombre",
  FNAC:     "Fecha Nac.",
  FINGRESO: "Fecha\nIngreso",
  TDOC:     "Tipo Doc.",
  NDOC:     "N° Doc",
  CALLE:    "Domicilio Calle",
  NUMERO:   "Domicilio Numero",
  LOCAL:    "Localidad",
  SEXO:     "Sexo",
  TEL:      "Telefono",
  TSOCIO:   "Tipo\nSocio",
  CONTACT:  "Contactado\npor",
  VOTA:     "Vota?",
  AQUIEN:   "A quien?",
  COMMENT:  "Comentarios",
};

const TIPO_SOCIO_OPTS = ["Activo", "Vitalicio", "Fuerza"];
const CONTACTADO_OPTS = [
  "Alexis","Claudio","Cristian","Iñaki","Jeronimo",
  "Juan Manuel","Martin Miguel","Nestor","Ruben","Santiago","Thiago","Walter"
];
const VOTA_OPTS    = ["Si", "No", "No sabe"];
const AQUIEN_OPTS  = ["Nosotros", "Los otros", "No sabe"];

const BADGE = {
  Activo:    "badge-activo",
  Vitalicio: "badge-vitalicio",
  Fuerza:    "badge-fuerza",
};

function normalize(str) {
  return (str || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

export default function App() {
  const [rows, setRows]       = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError]     = useState(null);
  const [search, setSearch]   = useState("");
  const [filterTipo, setFilterTipo]       = useState("Todos");
  const [filterContact, setFilterContact] = useState("Todos");
  const [editRowId, setEditRowId] = useState(null);
  const [editData, setEditData]   = useState({});
  const [savingRow, setSavingRow] = useState(null);
  const [flashRow, setFlashRow]   = useState(null);
  const [flashError, setFlashError] = useState(null);
  const [headerOpen, setHeaderOpen] = useState(true);

  // ── fetch ──
  const fetchData = useCallback(async () => {
    if (!SCRIPT_URL) { setError("⚙️ Falta configurar SCRIPT_URL en App.jsx"); setLoading(false); return; }
    try {
      setLoading(true);
      const res  = await fetch(SCRIPT_URL);
      const json = await res.json();
      if (!json.ok) throw new Error(json.error);
      setRows(json.rows);
      setError(null);
    } catch (e) {
      setError("Error al cargar datos: " + e.message);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => { fetchData(); }, [fetchData]);

  // ── contactado options ──
  const allContactados = useMemo(() => {
    const fromData = rows.map(r => r[COL.CONTACT]).filter(v => v && !CONTACTADO_OPTS.includes(v));
    return [...new Set([...CONTACTADO_OPTS, ...fromData])].sort();
  }, [rows]);

  // ── filtros ──
  const filtered = useMemo(() => {
    let result = rows;
    if (filterTipo !== "Todos")
      result = result.filter(r => r[COL.TSOCIO] === filterTipo);
    if (filterContact === "Sin contactar")
      result = result.filter(r => !r[COL.CONTACT]);
    else if (filterContact !== "Todos")
      result = result.filter(r => r[COL.CONTACT] === filterContact);
    if (search.trim()) {
      const q = normalize(search);
      result = result.filter(r => Object.values(r).some(v => normalize(String(v)).includes(q)));
    }
    return result;
  }, [rows, filterTipo, filterContact, search]);

  // ── stats ──
  const stats = useMemo(() => {
    const total = rows.length;
    const contactados = rows.filter(r => r[COL.CONTACT]);
    const byPerson = {};
    contactados.forEach(r => { const p = r[COL.CONTACT]; byPerson[p] = (byPerson[p] || 0) + 1; });

    const votaSi      = rows.filter(r => r[COL.VOTA] === "Si").length;
    const votaNosotros = rows.filter(r => r[COL.AQUIEN] === "Nosotros").length;
    const votaOtros    = rows.filter(r => r[COL.AQUIEN] === "Los otros").length;

    return { total, contactCount: contactados.length, byPerson, votaSi, votaNosotros, votaOtros };
  }, [rows]);

  // ── edit ──
  function startEdit(row) { setEditRowId(row._row); setEditData({ ...row }); }
  function cancelEdit()   { setEditRowId(null); setEditData({}); }

  async function saveEdit() {
    setSavingRow(editRowId);
    try {
      const params = new URLSearchParams();
      params.append("rowIndex", editData._row);
      Object.entries(editData).forEach(([k, v]) => { if (k !== "_row") params.append(k, v ?? ""); });
      const res  = await fetch(SCRIPT_URL, { method: "POST", body: params });
      const json = await res.json();
      if (!json.ok) throw new Error(json.error);
      setRows(prev => prev.map(r => r._row === editData._row ? { ...editData } : r));
      setFlashRow(editRowId);
      setTimeout(() => setFlashRow(null), 1200);
      setEditRowId(null);
      setEditData({});
    } catch (e) {
      setFlashError("Error al guardar: " + e.message);
      setTimeout(() => setFlashError(null), 3000);
    } finally {
      setSavingRow(null);
    }
  }

  // ── exportar ──
  function exportExcel() {
    const data = filtered.map(r => {
      const obj = {};
      Object.keys(COL).forEach(k => { obj[COL[k]] = r[COL[k]] || ""; });
      return obj;
    });
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Socios");
    XLSX.writeFile(wb, "socios_TFBB.xlsx");
  }
  function exportPDF() { window.print(); }

  const pct = stats.total ? Math.round((stats.contactCount / stats.total) * 100) : 0;
  const contactNames = useMemo(() => [...new Set(rows.map(r => r[COL.CONTACT]).filter(Boolean))].sort(), [rows]);

  if (!SCRIPT_URL) return (
    <div className="config-warning">
      <h2>⚙️ Configuración pendiente</h2>
      <p>Abrí <code>App.jsx</code> y pegá tu URL de Google Apps Script en <code>SCRIPT_URL</code>.</p>
    </div>
  );

  return (
    <div className="app">

      {/* ── HEADER ── */}
      <header className="app-header">
        {/* Barra siempre visible: título + toggle */}
        <div className="header-topbar">
          <div className="header-title">
            <span>🎯</span>
            <div>
              <h1>Club de Tiro TFBB</h1>
              <span className="subtitle">Gestión de Socios  ·  <span className="version-tag">{APP_VERSION}</span></span>
            </div>
          </div>
          <div className="header-topbar-right">
            <div className="export-btns">
              <button className="btn btn-excel"  onClick={exportExcel}>⬇ Excel</button>
              <button className="btn btn-pdf"    onClick={exportPDF}>🖨 PDF</button>
              <button className="btn btn-reload" onClick={fetchData} title="Recargar">↻</button>
            </div>
            <button
              className="btn-collapse"
              onClick={() => setHeaderOpen(o => !o)}
              title={headerOpen ? "Minimizar panel" : "Expandir panel"}
            >
              {headerOpen ? "▲" : "▼"}
            </button>
          </div>
        </div>

        {/* Contenido colapsable */}
        {headerOpen && (
          <div className="header-collapsible">

            {/* Progress contactados */}
            <div className="progress-section">
              <div className="progress-label">
                <span>Contactados: <strong>{stats.contactCount}/{stats.total}</strong> ({pct}%)</span>
                <div className="progress-by-person">
                  {Object.entries(stats.byPerson).sort((a,b) => b[1]-a[1]).map(([name, n]) => (
                    <span key={name} className="person-badge">{name} <em>{n}</em></span>
                  ))}
                </div>
              </div>
              <div className="progress-bar">
                <div className="progress-fill" style={{ width: pct + "%" }} />
              </div>
            </div>

            {/* Contadores de votos */}
            <div className="vote-counters">
              <div className="vote-card vote-si">
                <span className="vote-num">{stats.votaSi}</span>
                <span className="vote-label">Votan</span>
              </div>
              <div className="vote-card vote-nosotros">
                <span className="vote-num">{stats.votaNosotros}</span>
                <span className="vote-label">Nosotros</span>
              </div>
              <div className="vote-card vote-otros">
                <span className="vote-num">{stats.votaOtros}</span>
                <span className="vote-label">Los otros</span>
              </div>
              <div className="vote-card vote-total">
                <span className="vote-num">{stats.total}</span>
                <span className="vote-label">Total socios</span>
              </div>
            </div>

            {/* Búsqueda */}
            <input
              className="search-input"
              type="text"
              placeholder="🔍 Buscar por cualquier campo..."
              value={search}
              onChange={e => setSearch(e.target.value)}
            />

            {/* Filtros */}
            <div className="filters-row">
              <div className="chips-group">
                {["Todos", ...TIPO_SOCIO_OPTS].map(t => (
                  <button key={t}
                    className={`chip chip-tipo ${filterTipo === t ? "active" : ""} ${t !== "Todos" ? "chip-" + t.toLowerCase() : ""}`}
                    onClick={() => setFilterTipo(t)}
                  >{t}</button>
                ))}
              </div>
              <div className="chips-group chips-contact">
                {["Todos", "Sin contactar", ...contactNames].map(t => (
                  <button key={t}
                    className={`chip chip-contact ${filterContact === t ? "active" : ""}`}
                    onClick={() => setFilterContact(t)}
                  >{t}</button>
                ))}
              </div>
            </div>

            <div className="results-count">
              Mostrando <strong>{filtered.length}</strong> de {rows.length} socios
            </div>
          </div>
        )}
      </header>

      {flashError && <div className="flash-error">{flashError}</div>}

      {/* ── TABLA ── */}
      <div className="table-wrapper">
        {loading ? (
          <div className="loading">Cargando socios...</div>
        ) : error ? (
          <div className="error-msg">{error}</div>
        ) : (
          <table className="socios-table">
            <thead>
              <tr>
                <th>N° Socio</th>
                <th>Apellido</th>
                <th>Nombre</th>
                <th>Tipo</th>
                <th>Teléfono</th>
                <th>Localidad</th>
                <th>F. Ingreso</th>
                <th>Contactado por</th>
                <th>Vota?</th>
                <th>A quién?</th>
                <th>Comentarios</th>
                <th className="col-actions sticky-right">Acciones</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map(row => {
                const isEditing = editRowId === row._row;
                const isFlash   = flashRow  === row._row;
                const isSaving  = savingRow === row._row;
                const ed = isEditing ? editData : row;

                // helper select
                const Sel = ({ col, opts }) => (
                  <select className="cell-select"
                    value={ed[col] || ""}
                    onChange={e => setEditData(p => ({ ...p, [col]: e.target.value }))}
                  >
                    <option value="">—</option>
                    {opts.map(o => <option key={o}>{o}</option>)}
                  </select>
                );

                // helper text input
                const Inp = ({ col, w }) => (
                  <input className={`cell-input${w ? " w-" + w : ""}`}
                    value={ed[col] || ""}
                    onChange={e => setEditData(p => ({ ...p, [col]: e.target.value }))}
                  />
                );

                return (
                  <tr key={row._row}
                    className={(isEditing ? "row-editing " : "") + (isFlash ? "row-flash " : "")}
                  >
                    <td>{ed[COL.SOCIO]}</td>

                    {/* Apellido */}
                    <td>{isEditing ? <Inp col={COL.APELLIDO} /> : ed[COL.APELLIDO]}</td>

                    {/* Nombre */}
                    <td>{isEditing ? <Inp col={COL.NOMBRE} /> : ed[COL.NOMBRE]}</td>

                    {/* Tipo Socio */}
                    <td>
                      {isEditing
                        ? <Sel col={COL.TSOCIO} opts={TIPO_SOCIO_OPTS} />
                        : <span className={`badge ${BADGE[ed[COL.TSOCIO]] || ""}`}>{ed[COL.TSOCIO]}</span>
                      }
                    </td>

                    {/* Teléfono */}
                    <td>{isEditing ? <Inp col={COL.TEL} /> : ed[COL.TEL]}</td>

                    {/* Localidad */}
                    <td>{isEditing ? <Inp col={COL.LOCAL} /> : ed[COL.LOCAL]}</td>

                    {/* F. Ingreso (solo lectura) */}
                    <td>{ed[COL.FINGRESO]}</td>

                    {/* Contactado por */}
                    <td>
                      {isEditing
                        ? <select className="cell-select"
                            value={ed[COL.CONTACT] || ""}
                            onChange={e => setEditData(p => ({ ...p, [COL.CONTACT]: e.target.value }))}
                          >
                            <option value="">— Sin contactar —</option>
                            {allContactados.map(o => <option key={o}>{o}</option>)}
                          </select>
                        : ed[COL.CONTACT]
                          ? <span className="badge badge-contactado">{ed[COL.CONTACT]}</span>
                          : <span className="no-contact">—</span>
                      }
                    </td>

                    {/* Vota? */}
                    <td>
                      {isEditing
                        ? <Sel col={COL.VOTA} opts={VOTA_OPTS} />
                        : ed[COL.VOTA]
                          ? <span className={`badge badge-vota-${(ed[COL.VOTA]||"").toLowerCase().replace(" ","")}`}>{ed[COL.VOTA]}</span>
                          : <span className="no-contact">—</span>
                      }
                    </td>

                    {/* A quién? */}
                    <td>
                      {isEditing
                        ? <Sel col={COL.AQUIEN} opts={AQUIEN_OPTS} />
                        : ed[COL.AQUIEN]
                          ? <span className={`badge badge-aquien-${(ed[COL.AQUIEN]||"").toLowerCase().replace(" ","")}`}>{ed[COL.AQUIEN]}</span>
                          : <span className="no-contact">—</span>
                      }
                    </td>

                    {/* Comentarios */}
                    <td>
                      {isEditing
                        ? <input className="cell-input w-160"
                            value={ed[COL.COMMENT] || ""}
                            onChange={e => setEditData(p => ({ ...p, [COL.COMMENT]: e.target.value }))}
                          />
                        : <span className="comment-text">{ed[COL.COMMENT]}</span>
                      }
                    </td>

                    {/* Acciones sticky */}
                    <td className="col-actions sticky-right">
                      {isEditing ? (
                        <div className="action-btns">
                          <button className="btn-icon btn-save" onClick={saveEdit} disabled={isSaving} title="Guardar">
                            {isSaving ? "…" : "✓"}
                          </button>
                          <button className="btn-icon btn-cancel" onClick={cancelEdit} title="Cancelar">✗</button>
                        </div>
                      ) : (
                        <button className="btn-icon btn-edit" onClick={() => startEdit(row)} title="Editar">✏️</button>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>

      <style>{`
        @media print {
          .app-header .export-btns,
          .app-header .search-input,
          .app-header .filters-row,
          .col-actions { display: none !important; }
          .table-wrapper { overflow: visible !important; }
          .socios-table { font-size: 9px; }
        }
      `}</style>
    </div>
  );
}
