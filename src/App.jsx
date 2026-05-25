import { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

// ─────────────────────────────────────────────
// ⚙️  CONFIGURACIÓN — pegar tu URL de Apps Script
// ─────────────────────────────────────────────
const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxzGStt5G_JVMV0tWSr2POk3tMTRm-E8NVVArndlRY74Jf_yApoxl114azZofH7Q8Zy6Q/exec"; // ← pegar URL del deploy aquí

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
};

const TIPO_SOCIO_OPTS = ["Activo", "Vitalicio", "Fuerza"];
const CONTACTADO_OPTS = [
  "Alexis","Claudio","Cristian","Iñaki","Jeronimo",
  "Juan Manuel","Nestor","Ruben","Santiago","Thiago","Walter"
];

const BADGE = {
  Activo:    "badge-activo",
  Vitalicio: "badge-vitalicio",
  Fuerza:    "badge-fuerza",
};

// ─── utilidades ───────────────────────────────
function normalize(str) {
  return (str || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

// ═════════════════════════════════════════════
export default function App() {
  const [rows, setRows]         = useState([]);
  const [loading, setLoading]   = useState(true);
  const [error, setError]       = useState(null);
  const [search, setSearch]     = useState("");
  const [filterTipo, setFilterTipo]       = useState("Todos");
  const [filterContact, setFilterContact] = useState("Todos");
  const [editRowId, setEditRowId]   = useState(null); // _row being edited
  const [editData, setEditData]     = useState({});
  const [savingRow, setSavingRow]   = useState(null);
  const [flashRow, setFlashRow]     = useState(null);
  const [flashError, setFlashError] = useState(null);

  // ── fetch data ──
  const fetchData = useCallback(async () => {
    if (!SCRIPT_URL) {
      setError("⚠️ Falta configurar SCRIPT_URL en App.jsx");
      setLoading(false);
      return;
    }
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

  // ── contactado options (dinámico: los de la lista + los que ya existen en data)
  const allContactados = useMemo(() => {
    const fromData = rows
      .map(r => r[COL.CONTACT])
      .filter(v => v && !CONTACTADO_OPTS.includes(v));
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
      result = result.filter(r =>
        Object.values(r).some(v => normalize(String(v)).includes(q))
      );
    }
    return result;
  }, [rows, filterTipo, filterContact, search]);

  // ── progress ──
  const contactStats = useMemo(() => {
    const total = rows.length;
    const contactados = rows.filter(r => r[COL.CONTACT]);
    const byPerson = {};
    contactados.forEach(r => {
      const p = r[COL.CONTACT];
      byPerson[p] = (byPerson[p] || 0) + 1;
    });
    return { total, count: contactados.length, byPerson };
  }, [rows]);

  // ── edit handlers ──
  function startEdit(row) {
    setEditRowId(row._row);
    setEditData({ ...row });
  }
  function cancelEdit() {
    setEditRowId(null);
    setEditData({});
  }
  async function saveEdit() {
    setSavingRow(editRowId);
    try {
      const params = new URLSearchParams();
      params.append("rowIndex", editData._row);
      Object.entries(editData).forEach(([k, v]) => {
        if (k !== "_row") params.append(k, v ?? "");
      });
      const res  = await fetch(SCRIPT_URL, { method: "POST", body: params });
      const json = await res.json();
      if (!json.ok) throw new Error(json.error);

      // actualizar local
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

  // ── exportar Excel ──
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

  // ── exportar PDF ──
  function exportPDF() { window.print(); }

  // ─── render ────────────────────────────────
  const pct = contactStats.total
    ? Math.round((contactStats.count / contactStats.total) * 100)
    : 0;

  // lista de nombres que contactaron (para filtro chips)
  const contactNames = useMemo(() => {
    return [...new Set(rows.map(r => r[COL.CONTACT]).filter(Boolean))].sort();
  }, [rows]);

  if (!SCRIPT_URL) return (
    <div className="config-warning">
      <h2>⚙️ Configuración pendiente</h2>
      <p>Abrí <code>App.jsx</code> y pegá tu URL de Google Apps Script en la constante <code>SCRIPT_URL</code>.</p>
    </div>
  );

  return (
    <div className="app">
      <header className="app-header">
        <div className="header-top">
          <div>
            <h1>🎯 Club de Tiro TFBB</h1>
            <span className="subtitle">Gestión de Socios</span>
          </div>
          <div className="export-btns">
            <button className="btn btn-excel" onClick={exportExcel}>⬇ Excel</button>
            <button className="btn btn-pdf"   onClick={exportPDF}>🖨 PDF</button>
            <button className="btn btn-reload" onClick={fetchData} title="Recargar">↻</button>
          </div>
        </div>

        {/* Progress bar */}
        <div className="progress-section">
          <div className="progress-label">
            <span>Contactados: <strong>{contactStats.count}/{contactStats.total}</strong> ({pct}%)</span>
            <div className="progress-by-person">
              {Object.entries(contactStats.byPerson).sort((a,b) => b[1]-a[1]).map(([name, n]) => (
                <span key={name} className="person-badge">{name} <em>{n}</em></span>
              ))}
            </div>
          </div>
          <div className="progress-bar">
            <div className="progress-fill" style={{ width: pct + "%" }} />
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

        {/* Filtros Tipo Socio */}
        <div className="filters-row">
          <div className="chips-group">
            {["Todos", ...TIPO_SOCIO_OPTS].map(t => (
              <button
                key={t}
                className={`chip chip-tipo ${filterTipo === t ? "active" : ""} ${t !== "Todos" ? "chip-" + t.toLowerCase() : ""}`}
                onClick={() => setFilterTipo(t)}
              >{t}</button>
            ))}
          </div>

          {/* Filtros Contactado */}
          <div className="chips-group chips-contact">
            {["Todos", "Sin contactar", ...contactNames].map(t => (
              <button
                key={t}
                className={`chip chip-contact ${filterContact === t ? "active" : ""}`}
                onClick={() => setFilterContact(t)}
              >{t}</button>
            ))}
          </div>
        </div>

        <div className="results-count">
          Mostrando <strong>{filtered.length}</strong> de {rows.length} socios
        </div>
      </header>

      {/* Flash error */}
      {flashError && <div className="flash-error">{flashError}</div>}

      {/* Tabla */}
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
                <th>Contactado por</th>
                <th>F. Ingreso</th>
                <th className="col-actions sticky-right">Acciones</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map(row => {
                const isEditing = editRowId === row._row;
                const isFlash   = flashRow  === row._row;
                const isSaving  = savingRow === row._row;
                const ed = isEditing ? editData : row;

                return (
                  <tr
                    key={row._row}
                    className={
                      (isEditing ? "row-editing " : "") +
                      (isFlash   ? "row-flash "  : "")
                    }
                  >
                    <td>{ed[COL.SOCIO]}</td>

                    {/* Campos editables */}
                    {[COL.APELLIDO, COL.NOMBRE].map(col => (
                      <td key={col}>
                        {isEditing
                          ? <input className="cell-input" value={ed[col] || ""} onChange={e => setEditData(p => ({...p, [col]: e.target.value}))} />
                          : ed[col]
                        }
                      </td>
                    ))}

                    {/* Tipo Socio */}
                    <td>
                      {isEditing
                        ? <select className="cell-select" value={ed[COL.TSOCIO] || ""} onChange={e => setEditData(p => ({...p, [COL.TSOCIO]: e.target.value}))}>
                            <option value="">—</option>
                            {TIPO_SOCIO_OPTS.map(o => <option key={o}>{o}</option>)}
                          </select>
                        : <span className={`badge ${BADGE[ed[COL.TSOCIO]] || ""}`}>{ed[COL.TSOCIO]}</span>
                      }
                    </td>

                    {/* Teléfono */}
                    <td>
                      {isEditing
                        ? <input className="cell-input" value={ed[COL.TEL] || ""} onChange={e => setEditData(p => ({...p, [COL.TEL]: e.target.value}))} />
                        : ed[COL.TEL]
                      }
                    </td>

                    {/* Localidad */}
                    <td>
                      {isEditing
                        ? <input className="cell-input" value={ed[COL.LOCAL] || ""} onChange={e => setEditData(p => ({...p, [COL.LOCAL]: e.target.value}))} />
                        : ed[COL.LOCAL]
                      }
                    </td>

                    {/* Contactado por */}
                    <td>
                      {isEditing
                        ? <select className="cell-select" value={ed[COL.CONTACT] || ""} onChange={e => setEditData(p => ({...p, [COL.CONTACT]: e.target.value}))}>
                            <option value="">— Sin contactar —</option>
                            {allContactados.map(o => <option key={o}>{o}</option>)}
                          </select>
                        : ed[COL.CONTACT]
                          ? <span className="badge badge-contactado">{ed[COL.CONTACT]}</span>
                          : <span className="no-contact">—</span>
                      }
                    </td>

                    {/* Campos solo lectura adicionales */}
                    <td>{ed[COL.FINGRESO]}</td>
                    
                    {/* Acciones sticky */}
                    <td className="col-actions sticky-right">
                      {isEditing ? (
                        <div className="action-btns">
                          <button
                            className="btn-icon btn-save"
                            onClick={saveEdit}
                            disabled={isSaving}
                            title="Guardar"
                          >{isSaving ? "…" : "✓"}</button>
                          <button
                            className="btn-icon btn-cancel"
                            onClick={cancelEdit}
                            title="Cancelar"
                          >✗</button>
                        </div>
                      ) : (
                        <button
                          className="btn-icon btn-edit"
                          onClick={() => startEdit(row)}
                          title="Editar"
                        >✏️</button>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>

      {/* Print styles inline */}
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
