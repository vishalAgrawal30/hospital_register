import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ── ANTIBIOTIC DEFINITIONS ───────────────────────────────────────────────────
const ABX1 = [
  { code: "AMC", name: "Amoxicillin/Clavulanate" },
  { code: "PIT", name: "Piperacillin/Tazobactam" },
  { code: "CXM", name: "Cefuroxime" },
  { code: "CTR", name: "Ceftriaxone" },
  { code: "CFS", name: "Cefoperazone/Sulbactam" },
  { code: "CPM", name: "Cefepime" },
  { code: "ETP", name: "Ertapenem" },
  { code: "IPM", name: "Imipenem" },
  { code: "MRP", name: "Meropenem" },
  { code: "CL", name: "Colistin" },
  { code: "TIG", name: "Tigecycline" },
];
const ABX2 = [
  { code: "GEN", name: "Gentamicin" },
  { code: "AK", name: "Amikacin" },
  { code: "AT", name: "Aztreonam" },
  { code: "NET", name: "Netilmicin" },
  { code: "TOB", name: "Tobramycin" },
  { code: "CIP", name: "Ciprofloxacin" },
  { code: "LE", name: "Levofloxacin" },
  { code: "MI", name: "Minocycline" },
  { code: "CLS", name: "Clindamycin" },
  { code: "COT", name: "Co-trimoxazole" },
];
const ALL_ABX = [...ABX1, ...ABX2];

const ORGANISMS = [
  "NG", "NPG", "ESCHERICHIA COLI", "KLEBSIELLA PNEUMONIAE", "KLEBSIELLA PHNEMONIA",
  "PSEUDOMONAS AERUGINOSA", "PSEUDROMONAS AERGINOSA", "PSEUDOMONAS AERUGINERA",
  "ACINETOBACTER BAUMANNII", "STAPHYLOCOCCUS AUREUS", "ENTEROBACTER CLOACAE",
  "ENTEROCOCCUS FAECALIS", "CANDIDA SPECIES", "STREPTOCOCCUS SPP.",
  "PROTEUS MIRABILIS", "CITROBACTER FREUNDII",
];

const WARD_GROUPS = {
  "General Wards": ["MMW", "MOW", "MOE", "MSW", "MNW", "FMW", "FSW", "FNS"],
  "Surgical": ["GM", "GS", "GSF", "GSM", "G.SUR.", "ENT", "N.SUR", "N.SX", "NSX", "UROLOGY", "URO", "URTO"],
  "ICU / Special": ["MICU", "MICW", "SICU", "PICU", "NICU", "NEURO", "EHEMO", "FEMODI"],
  "Obstetrics / Gynec": ["OBST", "OBS.", "GYNEC", "PISW"],
  "Paediatrics": ["PAEDIA", "PEDIA"],
  "Emergency / OPD": ["CASUL", "OPD", "EM", "EMW", "EW", "EWIC", "EWK"],
  "Other": ["OTHER"],
};

const SAMPLES = [
  "PUS", "PUSDS", "PUSKS", "Sputum", "SPUTUM", "Urine", "URINE", "CSF", "FLUID",
  "PLEURAL FLUID", "ASCITIC", "ASC", "TISSUE", "BLOOD", "SECRETION", "ET", "AFB",
  "STOOL", "GUSTRIG ASPIRATE", "TRECHEAL", "G/S", "PAS", "OTHER",
];

const GRAM_STAIN_OPTIONS = [
  "", "Gram positive cocci seen", "Gram negative bacilli seen",
  "Gram positive cocci in pairs seen", "Gram positive cocci in clusters seen",
  "No organisms seen", "Gram negative cocci seen", "Gram positive bacilli seen",
];

const todayStr = () => new Date().toISOString().slice(0, 10);

const fmtDate = (d) => {
  if (!d) return "—";
  return new Date(d + "T00:00:00").toLocaleDateString("en-IN", {
    day: "2-digit", month: "short", year: "numeric",
  });
};

const loadRecords = () => {
  try { return JSON.parse(localStorage.getItem("labRec") || "[]"); } catch { return []; }
};
const saveRecords = (recs) => localStorage.setItem("labRec", JSON.stringify(recs));

// ── STYLES (CSS-in-JS via style tag injection) ───────────────────────────────
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;600&display=swap');

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root {
  --pri: #1565C0; --pri-light: #E3F2FD; --pri-dark: #0D47A1;
  --success: #2E7D32; --success-light: #E8F5E9;
  --warn: #E65100; --warn-light: #FFF3E0;
  --danger: #C62828; --danger-light: #FFEBEE;
  --g50: #F8F9FA; --g100: #F1F3F4; --g200: #E8EAED;
  --g300: #DADCE0; --g500: #9AA0A6; --g700: #5F6368; --g900: #202124;
  --border: #DADCE0; --rad: 8px; --rad-lg: 12px;
  --shadow: 0 1px 3px rgba(0,0,0,.10),0 1px 2px rgba(0,0,0,.06);
  --shadow-md: 0 4px 16px rgba(0,0,0,.10);
  --font: 'DM Sans', system-ui, sans-serif;
  --mono: 'IBM Plex Mono', monospace;
}

body { font-family: var(--font); background: #EEF2F7; color: var(--g900); min-height: 100vh; font-size: 14px; }

/* HEADER */
.hdr { background: linear-gradient(135deg,#1565C0,#0D47A1); color: #fff; padding: 0 16px; height: 56px; display: flex; align-items: center; justify-content: space-between; position: sticky; top: 0; z-index: 100; box-shadow: 0 2px 8px rgba(0,0,0,.22); }
.hdr-ico { width: 34px; height: 34px; background: rgba(255,255,255,.18); border-radius: 8px; display: flex; align-items: center; justify-content: center; font-size: 18px; }
.hdr-title { font-size: 15px; font-weight: 700; line-height: 1.2; }
.hdr-sub { font-size: 11px; opacity: .72; }
.hdr-btn { display: inline-flex; align-items: center; gap: 5px; padding: 7px 12px; border-radius: 6px; font-size: 12px; font-weight: 600; cursor: pointer; border: 1px solid rgba(255,255,255,.35); background: rgba(255,255,255,.14); color: #fff; font-family: var(--font); }
.hdr-btn:hover { background: rgba(255,255,255,.25); }

/* TABS */
.tabs { background: #fff; border-bottom: 1px solid var(--border); display: flex; padding: 0 12px; overflow-x: auto; scrollbar-width: none; }
.tabs::-webkit-scrollbar { display: none; }
.tab { padding: 13px 16px; font-size: 13px; font-weight: 500; color: var(--g700); cursor: pointer; border-bottom: 2.5px solid transparent; white-space: nowrap; display: flex; align-items: center; gap: 5px; transition: color .15s, border-color .15s, background .15s; font-family: var(--font); background: none; border-top: none; border-left: none; border-right: none; }
.tab:hover { color: var(--pri); background: var(--pri-light); }
.tab.active { color: var(--pri); border-bottom-color: var(--pri); background: var(--pri-light); }
.tab-badge { background: var(--pri); color: #fff; border-radius: 12px; padding: 1px 6px; font-size: 11px; font-weight: 700; }

/* MAIN */
.main { padding: 16px; max-width: 1200px; margin: 0 auto; }

/* CARD */
.card { background: #fff; border-radius: var(--rad-lg); border: 1px solid var(--border); padding: 16px 18px; box-shadow: var(--shadow); margin-bottom: 14px; }
.card-hdr { display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px; padding-bottom: 12px; border-bottom: 1px solid var(--g100); gap: 8px; flex-wrap: wrap; }
.card-title { font-size: 14px; font-weight: 700; color: var(--g900); display: flex; align-items: center; gap: 8px; }
.card-ico { width: 26px; height: 26px; background: var(--pri-light); border-radius: 6px; display: flex; align-items: center; justify-content: center; font-size: 13px; flex-shrink: 0; }

/* FORM */
.form-grid { display: flex; flex-direction: column; gap: 14px; }
.form-row { display: grid; gap: 12px; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); }
.form-group { display: flex; flex-direction: column; gap: 4px; }
.form-group.span2 { grid-column: span 2; }
.form-group.full { grid-column: 1 / -1; }
.field-lbl { font-size: 11px; font-weight: 700; color: var(--g700); text-transform: uppercase; letter-spacing: .5px; }
.req { color: var(--danger); margin-left: 2px; }
input, select, textarea { width: 100%; padding: 9px 11px; border: 1.5px solid var(--g300); border-radius: var(--rad); font-size: 13px; color: var(--g900); background: #fff; transition: border .15s, box-shadow .15s; font-family: var(--font); -webkit-appearance: none; appearance: none; }
input:focus, select:focus, textarea:focus { outline: none; border-color: var(--pri); box-shadow: 0 0 0 3px rgba(21,101,192,.13); }
select { background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3E%3Cpath fill='%239AA0A6' d='M6 8L1 3h10z'/%3E%3C/svg%3E"); background-repeat: no-repeat; background-position: right 10px center; padding-right: 28px; cursor: pointer; }
textarea { resize: vertical; min-height: 64px; line-height: 1.5; }
.err-msg { font-size: 11px; color: var(--danger); }
.input-err { border-color: var(--danger) !important; }

/* SECTION DIVIDER */
.sec-div { display: flex; align-items: center; gap: 8px; margin: 16px 0 12px; color: var(--g700); font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .7px; }
.sec-div::before, .sec-div::after { content: ''; flex: 1; height: 1px; background: var(--g200); }

/* AST */
.ast-legend { display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 12px; background: var(--g50); padding: 8px 12px; border-radius: var(--rad); border: 1px solid var(--g200); }
.ast-legend-item { display: flex; align-items: center; gap: 5px; font-size: 12px; font-weight: 600; }
.ast-legend-dot { width: 10px; height: 10px; border-radius: 50%; }
.ast-group-label { font-size: 12px; font-weight: 700; color: var(--g700); margin: 10px 0 8px; padding-left: 2px; display: flex; align-items: center; gap: 6px; }
.ast-group-label::after { content: ''; flex: 1; height: 1px; background: var(--g200); margin-left: 6px; }
.ast-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(110px, 1fr)); gap: 8px; margin-bottom: 4px; }
.ast-card { border: 1.5px solid var(--g200); border-radius: var(--rad); background: var(--g50); overflow: hidden; transition: border-color .15s, background .15s; }
.ast-card:hover { border-color: var(--pri); }
.ast-card-hdr { background: var(--g100); padding: 5px 8px; display: flex; align-items: center; justify-content: space-between; gap: 4px; }
.ast-abx-code { font-size: 12px; font-weight: 800; color: var(--g900); font-family: var(--mono); }
.ast-abx-name { font-size: 9px; color: var(--g500); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 70px; }
.ast-card-body { padding: 8px; }
.ast-row { display: flex; align-items: center; }
.ast-row select { padding: 4px 6px; font-size: 12px; height: 30px; }
.ast-sir-select { width: 100%; font-weight: 700; text-align: center; }
.ast-s { border-color: var(--success) !important; background: var(--success-light) !important; }
.ast-s .ast-card-hdr { background: #C8E6C9; }
.ast-r { border-color: var(--danger) !important; background: var(--danger-light) !important; }
.ast-r .ast-card-hdr { background: #FFCDD2; }
.ast-i { border-color: #F57C00 !important; background: var(--warn-light) !important; }
.ast-i .ast-card-hdr { background: #FFE0B2; }

/* BUTTONS */
.btn { display: inline-flex; align-items: center; gap: 5px; padding: 9px 16px; border-radius: var(--rad); font-size: 13px; font-weight: 600; cursor: pointer; border: 1.5px solid transparent; transition: all .15s; white-space: nowrap; font-family: var(--font); }
.btn:active { transform: scale(.97); }
.btn-primary { background: var(--pri); color: #fff; border-color: var(--pri); }
.btn-primary:hover { background: var(--pri-dark); }
.btn-success { background: var(--success); color: #fff; }
.btn-success:hover { filter: brightness(.92); }
.btn-outline { background: #fff; color: var(--pri); border-color: var(--pri); }
.btn-outline:hover { background: var(--pri-light); }
.btn-danger { background: var(--danger); color: #fff; border-color: var(--danger); }
.btn-danger:hover { filter: brightness(.9); }
.btn-ghost { background: transparent; color: var(--g700); border-color: var(--g300); }
.btn-ghost:hover { background: var(--g100); }
.btn-sm { padding: 6px 11px; font-size: 12px; }
.btn-lg { padding: 11px 22px; font-size: 14px; }
.btn-group { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 18px; padding-top: 16px; border-top: 1px solid var(--g100); }

/* STATS */
.stats-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(130px, 1fr)); gap: 10px; margin-bottom: 14px; }
.stat-card { background: #fff; border: 1px solid var(--border); border-radius: var(--rad-lg); padding: 12px 14px; box-shadow: var(--shadow); }
.stat-val { font-size: 24px; font-weight: 800; line-height: 1; margin-bottom: 3px; font-family: var(--mono); }
.stat-lbl { font-size: 10px; color: var(--g500); font-weight: 600; text-transform: uppercase; letter-spacing: .5px; }
.stat-blue .stat-val { color: var(--pri); }
.stat-green .stat-val { color: var(--success); }
.stat-orange .stat-val { color: #E65100; }
.stat-red .stat-val { color: var(--danger); }

/* TABLE */
.table-wrap { overflow-x: auto; border-radius: var(--rad); border: 1px solid var(--border); }
table { width: 100%; border-collapse: collapse; font-size: 12px; }
th { background: var(--g50); padding: 9px 10px; text-align: left; font-size: 10px; font-weight: 800; color: var(--g700); text-transform: uppercase; letter-spacing: .5px; border-bottom: 1px solid var(--border); white-space: nowrap; }
td { padding: 9px 10px; border-bottom: 1px solid var(--g100); vertical-align: middle; max-width: 180px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
tr:last-child td { border-bottom: none; }
tr:hover td { background: #F5F8FF; }
.badge { display: inline-flex; align-items: center; padding: 2px 7px; border-radius: 20px; font-size: 11px; font-weight: 700; }
.badge-blue { background: #E3F2FD; color: #1565C0; }
.badge-green { background: #E8F5E9; color: #2E7D32; }
.badge-orange { background: #FFF3E0; color: #E65100; }
.badge-red { background: #FFEBEE; color: #C62828; }
.badge-gray { background: #F1F3F4; color: #5F6368; }

/* MOBILE CARDS */
.mob-card { border: 1px solid var(--border); border-radius: var(--rad); background: #fff; padding: 12px; margin-bottom: 8px; box-shadow: var(--shadow); }
.mob-card-top { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 8px; }
.mob-card-name { font-size: 14px; font-weight: 700; }
.mob-card-date { font-size: 11px; color: var(--g500); }
.mob-card-meta { display: flex; gap: 5px; flex-wrap: wrap; margin-bottom: 8px; }
.mob-card-actions { display: flex; gap: 6px; }

/* SEARCH */
.search-bar { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 14px; }
.search-wrap { position: relative; flex: 1; min-width: 180px; }
.search-wrap input { padding-left: 32px; }
.search-wrap::before { content: '🔍'; position: absolute; left: 9px; top: 50%; transform: translateY(-50%); font-size: 13px; }
.filter-row { display: flex; gap: 8px; flex-wrap: wrap; width: 100%; margin-bottom: 14px; }
.filter-row select { flex: 1; min-width: 120px; }

/* TOAST */
.toast-con { position: fixed; bottom: 20px; right: 16px; display: flex; flex-direction: column; gap: 8px; z-index: 9999; max-width: calc(100vw - 32px); pointer-events: none; }
.toast { padding: 11px 16px; border-radius: var(--rad); color: #fff; font-size: 13px; font-weight: 600; box-shadow: var(--shadow-md); display: flex; align-items: center; gap: 8px; width: 300px; max-width: 100%; animation: toastin .2s ease; font-family: var(--font); }
.toast-success { background: #2E7D32; }
.toast-error { background: #C62828; }
.toast-info { background: #1565C0; }
@keyframes toastin { from { transform: translateX(120px); opacity: 0; } to { transform: translateX(0); opacity: 1; } }

/* EMPTY */
.empty { text-align: center; padding: 40px 20px; color: var(--g500); }
.empty-ico { font-size: 36px; margin-bottom: 10px; }
.empty-ttl { font-size: 14px; font-weight: 700; color: var(--g700); margin-bottom: 5px; }

/* MODAL OVERLAY */
.modal-overlay { position: fixed; inset: 0; background: rgba(0,0,0,.42); z-index: 200; display: flex; align-items: flex-end; justify-content: center; }
@media (min-width: 600px) { .modal-overlay { align-items: center; padding: 20px; } }
.modal { background: #fff; width: 100%; max-width: 820px; max-height: 92vh; overflow-y: auto; box-shadow: 0 20px 60px rgba(0,0,0,.22); border-radius: 16px 16px 0 0; }
@media (min-width: 600px) { .modal { border-radius: var(--rad-lg); } }
.modal-hdr { padding: 16px 18px; border-bottom: 1px solid var(--border); display: flex; align-items: center; justify-content: space-between; position: sticky; top: 0; background: #fff; z-index: 1; }
.modal-body { padding: 18px; }
.modal-ftr { padding: 14px 18px; border-top: 1px solid var(--border); display: flex; gap: 8px; justify-content: flex-end; position: sticky; bottom: 0; background: #fff; }
.close-btn { width: 30px; height: 30px; border: none; background: var(--g100); border-radius: 6px; cursor: pointer; font-size: 16px; display: flex; align-items: center; justify-content: center; }
.close-btn:hover { background: var(--g200); }

/* REPORT GRID */
.report-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }

/* RESPONSIVE */
@media (max-width: 640px) {
  .hdr-title { font-size: 13px; }
  .hdr-sub { display: none; }
  .main { padding: 10px; }
  .card { padding: 13px 14px; }
  .form-row { grid-template-columns: 1fr; }
  .form-group.span2 { grid-column: 1; }
  .ast-grid { grid-template-columns: repeat(auto-fill, minmax(140px, 1fr)); }
  .btn-group { flex-direction: column; }
  .btn-group .btn { width: 100%; justify-content: center; }
  .report-grid { grid-template-columns: 1fr; }
  .stats-row { grid-template-columns: 1fr 1fr; }
  .table-wrap { display: none; }
  .mob-records { display: block !important; }
}
@media (min-width: 641px) { .mob-records { display: none !important; } }
@media (max-width: 400px) {
  .ast-grid { grid-template-columns: 1fr 1fr; }
  .stats-row { grid-template-columns: 1fr 1fr; }
}
@media print { .hdr, .tabs, .btn, .search-bar { display: none !important; } .main { padding: 0; } .card { box-shadow: none; } }

/* IROW */
.irow { display: flex; justify-content: space-between; align-items: center; padding: 6px 0; border-bottom: 1px solid var(--g100); font-size: 13px; gap: 8px; }
.irow-lbl { color: var(--g500); font-weight: 600; flex-shrink: 0; }
.irow-val { text-align: right; }

/* LOGIN SCREEN */
.login-overlay {
  display: flex;
  align-items: center;
  justify-content: center;
  min-height: 100vh;
  background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
  padding: 20px;
}
.login-card {
  background: rgba(255, 255, 255, 0.95);
  backdrop-filter: blur(10px);
  border-radius: 16px;
  width: 100%;
  max-width: 400px;
  padding: 32px 28px;
  box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.3), 0 10px 10px -5px rgba(0, 0, 0, 0.3);
  border: 1px solid rgba(255, 255, 255, 0.25);
  animation: loginCardIn 0.5s cubic-bezier(0.16, 1, 0.3, 1) both;
}
@keyframes loginCardIn {
  from { transform: translateY(30px); opacity: 0; }
  to { transform: translateY(0); opacity: 1; }
}
.login-logo {
  font-size: 42px;
  margin-bottom: 12px;
  text-align: center;
  animation: logoPulse 2s infinite ease-in-out;
}
@keyframes logoPulse {
  0%, 100% { transform: scale(1); }
  50% { transform: scale(1.08); }
}
.login-title {
  font-size: 20px;
  font-weight: 700;
  color: #0f172a;
  text-align: center;
  margin-bottom: 6px;
  letter-spacing: -0.5px;
}
.login-subtitle {
  font-size: 13px;
  color: #64748b;
  text-align: center;
  margin-bottom: 24px;
}
.login-form {
  display: flex;
  flex-direction: column;
}
.login-error {
  background: #fef2f2;
  border: 1px solid #fee2e2;
  color: #991b1b;
  font-size: 12px;
  font-weight: 600;
  padding: 10px 12px;
  border-radius: 8px;
  margin-bottom: 16px;
  display: flex;
  align-items: center;
  gap: 6px;
  animation: shake 0.3s ease-in-out;
}
@keyframes shake {
  0%, 100% { transform: translateX(0); }
  25% { transform: translateX(-4px); }
  75% { transform: translateX(4px); }
}
`;

// ── HELPERS ──────────────────────────────────────────────────────────────────
const OrgBadge = ({ org }) => {
  if (!org) return <span className="badge badge-gray">—</span>;
  if (org === "NG" || org === "NPG") return <span className="badge badge-gray">{org}</span>;
  return <span className="badge badge-orange">{org}</span>;
};
const ZnBadge = ({ zn }) => {
  if (!zn) return <span style={{ color: "var(--g500)" }}>—</span>;
  const pos = zn.toUpperCase().includes("AFB SEEN") && !zn.toUpperCase().includes("NOT") && !zn.toUpperCase().includes("BOT") && !zn.toUpperCase().includes("NOR");
  return pos ? <span className="badge badge-red">AFB+</span> : <span className="badge badge-green">AFB−</span>;
};
const KohBadge = ({ koh }) => {
  if (!koh) return <span style={{ color: "var(--g500)" }}>—</span>;
  return koh.toUpperCase().startsWith("NO") ? <span className="badge badge-green">KOH−</span> : <span className="badge badge-red">KOH+</span>;
};

const IRow = ({ label, children }) => (
  <div className="irow">
    <span className="irow-lbl">{label}</span>
    <span className="irow-val">{children}</span>
  </div>
);

// ── TOAST ────────────────────────────────────────────────────────────────────
let _toastId = 0;
function useToasts() {
  const [toasts, setToasts] = useState([]);
  const toast = useCallback((msg, type = "info") => {
    const id = ++_toastId;
    setToasts((t) => [...t, { id, msg, type }]);
    setTimeout(() => setToasts((t) => t.filter((x) => x.id !== id)), 3200);
  }, []);
  return { toasts, toast };
}

// ── AST CARD ─────────────────────────────────────────────────────────────────
const ASTCard = ({ ab, sir, onChange }) => {
  const cls = "ast-card" + (sir === "S" ? " ast-s" : sir === "R" ? " ast-r" : sir === "I" ? " ast-i" : "");
  return (
    <div className={cls} title={ab.name}>
      <div className="ast-card-hdr">
        <span className="ast-abx-code">{ab.code}</span>
        <span className="ast-abx-name">{ab.name}</span>
      </div>
      <div className="ast-card-body">
        <div className="ast-row">
          <select className="ast-sir-select" value={sir} onChange={(e) => onChange(ab.code, e.target.value)}>
            <option value="">—</option>
            <option value="S">S</option>
            <option value="I">I</option>
            <option value="R">R</option>
          </select>
        </div>
      </div>
    </div>
  );
};

// ── BAR CHART ────────────────────────────────────────────────────────────────
const BarChart = ({ data }) => {
  if (!data.length) return <div className="empty" style={{ padding: "24px 0" }}>No data</div>;
  const mx = Math.max(...data.map((d) => d[1]));
  const cols = ["#1565C0", "#1976D2", "#1E88E5", "#2196F3", "#42A5F5", "#64B5F6", "#90CAF9"];
  return (
    <div>
      {data.map(([k, v], i) => (
        <div key={k} style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 0", borderBottom: "1px solid var(--g100)" }}>
          <div style={{ minWidth: 110, maxWidth: 110, fontSize: 11, color: "var(--g700)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={k}>{k}</div>
          <div style={{ flex: 1, height: 16, background: "var(--g100)", borderRadius: 4, overflow: "hidden" }}>
            <div style={{ width: `${Math.round(v / mx * 100)}%`, height: "100%", background: cols[i] || "#90CAF9", borderRadius: 4 }} />
          </div>
          <div style={{ minWidth: 24, textAlign: "right", fontSize: 12, fontWeight: 700, color: "var(--g700)" }}>{v}</div>
        </div>
      ))}
    </div>
  );
};

const countBy = (arr, fn) => {
  const c = {};
  arr.forEach((r) => { const k = fn(r) || "Unknown"; c[k] = (c[k] || 0) + 1; });
  return Object.entries(c).sort((a, b) => b[1] - a[1]).slice(0, 10);
};

// ── VIEW MODAL ───────────────────────────────────────────────────────────────
const ViewModal = ({ record, onClose, onEdit, onDelete }) => {
  if (!record) return null;
  const astEntries = Object.entries(record.ast || {});

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <div className="modal-hdr">
          <div className="card-title">
            {record.patientName}{record.ward ? ` — ${record.ward}` : ""}
          </div>
          <button className="close-btn" onClick={onClose}>✕</button>
        </div>
        <div className="modal-body">
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 18 }}>
            <div>
              <div className="sec-div" style={{ marginTop: 0 }}>Patient Info</div>
              <IRow label="Date">{fmtDate(record.date)}</IRow>
              <IRow label="IPD/OPD"><span className="badge badge-blue">{record.ipdopd || "—"}</span></IRow>
              <IRow label="Name"><strong>{record.patientName}</strong></IRow>
              <IRow label="Ward"><span className="badge badge-gray">{record.ward || "—"}</span></IRow>
              <IRow label="Sample">{record.sample || "—"}</IRow>
            </div>
            <div>
              <div className="sec-div" style={{ marginTop: 0 }}>Microbiology</div>
              <IRow label="Organism"><OrgBadge org={record.organism} /></IRow>
              <IRow label="ZN Stain">{record.znStain || "—"}</IRow>
              <IRow label="KOH">{record.koh || "—"}</IRow>
            </div>
            {record.gramStain && (
              <div style={{ gridColumn: "1 / -1" }}>
                <div className="sec-div">Gram Stain</div>
                <div style={{ background: "var(--g50)", padding: 10, borderRadius: "var(--rad)", fontSize: 13, color: "var(--g700)", lineHeight: 1.5 }}>
                  {record.gramStain}
                </div>
              </div>
            )}
            <div style={{ gridColumn: "1 / -1" }}>
              <div className="sec-div">AST Results</div>
              {astEntries.length === 0 ? (
                <span style={{ color: "var(--g500)", fontSize: 13 }}>No AST data recorded</span>
              ) : (
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                  {astEntries.map(([code, val]) => {
                    const sir = typeof val === "object" ? val.sir : val;
                    const cls = sir === "S" ? "badge-green" : sir === "R" ? "badge-red" : "badge-orange";
                    return (
                      <span key={code} className={`badge ${cls}`} style={{ gap: 4 }}>
                        {code}{sir && <b> {sir}</b>}
                      </span>
                    );
                  })}
                </div>
              )}
            </div>
            {record.remarks && (
              <div style={{ gridColumn: "1 / -1" }}>
                <div className="sec-div">Remarks</div>
                <div style={{ fontSize: 13, color: "var(--g700)" }}>{record.remarks}</div>
              </div>
            )}
          </div>
        </div>
        <div className="modal-ftr">
          <button className="btn btn-danger btn-sm" onClick={() => onDelete(record.id)}>🗑️ Delete</button>
          <button className="btn btn-outline btn-sm" onClick={() => onEdit(record.id)}>✏️ Edit</button>
          <button className="btn btn-ghost btn-sm" onClick={onClose}>Close</button>
        </div>
      </div>
    </div>
  );
};

// ── ENTRY PANE ───────────────────────────────────────────────────────────────
const emptyAst = () => {
  const a = {};
  ALL_ABX.forEach((ab) => { a[ab.code] = ""; });
  return a;
};

const EntryPane = ({ onSave, editingRecord, onCancelEdit, toast }) => {
  const [date, setDate] = useState(todayStr());
  const [ipdopd, setIpdopd] = useState("");
  const [name, setName] = useState("");
  const [ward, setWard] = useState("");
  const [sample, setSample] = useState("");
  const [gramStain, setGramStain] = useState("");
  const [znStain, setZnStain] = useState("");
  const [koh, setKoh] = useState("");
  const [organism, setOrganism] = useState("");
  const [remarks, setRemarks] = useState("");
  const [ast, setAst] = useState(emptyAst());
  const [errors, setErrors] = useState({});
  const isEditing = !!editingRecord;

  useEffect(() => {
    if (editingRecord) {
      setDate(editingRecord.date || "");
      setIpdopd(editingRecord.ipdopd || "");
      setName(editingRecord.patientName || "");
      setWard(editingRecord.ward || "");
      setSample(editingRecord.sample || "");
      setGramStain(editingRecord.gramStain || "");
      setZnStain(editingRecord.znStain || "");
      setKoh(editingRecord.koh || "");
      setOrganism(editingRecord.organism || "");
      setRemarks(editingRecord.remarks || "");
      const a = emptyAst();
      Object.entries(editingRecord.ast || {}).forEach(([code, val]) => {
        if (code in a) {
          a[code] = typeof val === "object" ? (val.sir || "") : (val || "");
        }
      });
      setAst(a);
    }
  }, [editingRecord]);

  const clearForm = () => {
    setDate(todayStr()); setIpdopd(""); setName(""); setWard(""); setSample("");
    setGramStain(""); setZnStain(""); setKoh(""); setOrganism(""); setRemarks("");
    setAst(emptyAst()); setErrors({});
    if (onCancelEdit) onCancelEdit();
  };

  const handleAstChange = (code, value) => {
    setAst((prev) => ({ ...prev, [code]: value }));
  };

  const validate = () => {
    const e = {};
    if (!date) e.date = "Date is required";
    if (!ipdopd.trim()) e.ipdopd = "IPD/OPD No. required";
    if (!name.trim()) e.name = "Patient name required";
    if (!ward) e.ward = "Ward is required";
    if (!sample) e.sample = "Sample type required";
    setErrors(e);
    return Object.keys(e).length === 0;
  };

  const handleSubmit = () => {
    if (!validate()) { toast("Please fix the errors above", "error"); return; }
    const astClean = {};
    Object.entries(ast).forEach(([code, val]) => {
      if (val) astClean[code] = val;
    });
    onSave({
      id: editingRecord?.id || Date.now().toString(),
      date, ipdopd: ipdopd.trim(), patientName: name.trim(), ward, sample,
      gramStain: gramStain.trim(), znStain, koh, organism: organism.trim(),
      ast: astClean, remarks: remarks.trim(),
    });
    clearForm();
  };

  return (
    <div className="card">
      <div className="card-hdr">
        <div className="card-title">
          <div className="card-ico">{isEditing ? "✏️" : "📝"}</div>
          {isEditing ? `Editing: ${editingRecord.patientName}` : "Patient & Sample Information"}
        </div>
        <div style={{ fontSize: 11, color: "var(--g500)" }}><span style={{ color: "var(--danger)" }}>*</span> required</div>
      </div>

      <div className="form-grid">
        {/* Row 1 */}
        <div className="form-row">
          <div className="form-group">
            <label className="field-lbl">Date <span className="req">*</span></label>
            <input type="date" value={date} onChange={(e) => setDate(e.target.value)} className={errors.date ? "input-err" : ""} />
            {errors.date && <span className="err-msg">{errors.date}</span>}
          </div>
          <div className="form-group">
            <label className="field-lbl">IPD/OPD No. <span className="req">*</span></label>
            <input type="text" value={ipdopd} onChange={(e) => setIpdopd(e.target.value)} placeholder="e.g. 205 or IPD" className={errors.ipdopd ? "input-err" : ""} />
            {errors.ipdopd && <span className="err-msg">{errors.ipdopd}</span>}
          </div>
          <div className="form-group span2">
            <label className="field-lbl">Patient Name <span className="req">*</span></label>
            <input type="text" value={name} onChange={(e) => setName(e.target.value)} placeholder="Full patient name" autoComplete="off" className={errors.name ? "input-err" : ""} />
            {errors.name && <span className="err-msg">{errors.name}</span>}
          </div>
        </div>

        {/* Row 2 */}
        <div className="form-row">
          <div className="form-group">
            <label className="field-lbl">Ward <span className="req">*</span></label>
            <select value={ward} onChange={(e) => setWard(e.target.value)} className={errors.ward ? "input-err" : ""}>
              <option value="">-- Select --</option>
              {Object.entries(WARD_GROUPS).map(([group, wards]) => (
                <optgroup key={group} label={group}>
                  {wards.map((w) => <option key={w}>{w}</option>)}
                </optgroup>
              ))}
            </select>
            {errors.ward && <span className="err-msg">{errors.ward}</span>}
          </div>
          <div className="form-group">
            <label className="field-lbl">Sample Type <span className="req">*</span></label>
            <select value={sample} onChange={(e) => setSample(e.target.value)} className={errors.sample ? "input-err" : ""}>
              <option value="">-- Select --</option>
              {SAMPLES.map((s) => <option key={s}>{s}</option>)}
            </select>
            {errors.sample && <span className="err-msg">{errors.sample}</span>}
          </div>
          <div className="form-group span2">
            <label className="field-lbl">Gram Stain Result</label>
            <input type="text" value={gramStain} onChange={(e) => setGramStain(e.target.value)} placeholder="Describe findings..." list="gram-list" />
            <datalist id="gram-list">
              {GRAM_STAIN_OPTIONS.filter(Boolean).map((o) => <option key={o} value={o} />)}
            </datalist>
          </div>
        </div>

        {/* Row 3 */}
        <div className="form-row">
          <div className="form-group">
            <label className="field-lbl">ZN Stain</label>
            <select value={znStain} onChange={(e) => setZnStain(e.target.value)}>
              <option value="">-- Not done --</option>
              <option>AFB SEEN</option>
              <option>AFB NOT SEEN</option>
              <option>AFB not seen</option>
              <option>AFB BOT SEEN</option>
              <option>AFB NOR SEEN</option>
            </select>
          </div>
          <div className="form-group">
            <label className="field-lbl">KOH</label>
            <select value={koh} onChange={(e) => setKoh(e.target.value)}>
              <option value="">-- Not done --</option>
              <option>No Fungal element seen</option>
              <option>NO FUNGAL ELEMENT SEEN</option>
              <option>Fungal elements seen</option>
            </select>
          </div>
          <div className="form-group span2">
            <label className="field-lbl">Organism Isolated</label>
            <input type="text" value={organism} onChange={(e) => setOrganism(e.target.value)} list="org-list" placeholder="Type or select organism" autoComplete="off" />
            <datalist id="org-list">
              {ORGANISMS.map((o) => <option key={o} value={o} />)}
            </datalist>
          </div>
        </div>

        {/* AST Section */}
        <div>
          <div className="sec-div">🧪 Antibiotic Sensitivity Testing (AST)</div>
          <div className="ast-legend">
            <div className="ast-legend-item"><div className="ast-legend-dot" style={{ background: "var(--success)" }} /> S = Sensitive</div>
            <div className="ast-legend-item"><div className="ast-legend-dot" style={{ background: "#F57C00" }} /> I = Intermediate</div>
            <div className="ast-legend-item"><div className="ast-legend-dot" style={{ background: "var(--danger)" }} /> R = Resistant</div>
            <div style={{ flex: 1 }} />
          </div>
          <div className="ast-group-label">🔵 Beta-lactams &amp; Carbapenems</div>
          <div className="ast-grid">
            {ABX1.map((ab) => (
              <ASTCard key={ab.code} ab={ab} sir={ast[ab.code] || ""} onChange={handleAstChange} />
            ))}
          </div>
          <div className="ast-group-label" style={{ marginTop: 14 }}>🟠 Aminoglycosides, Fluoroquinolones &amp; Others</div>
          <div className="ast-grid">
            {ABX2.map((ab) => (
              <ASTCard key={ab.code} ab={ab} sir={ast[ab.code] || ""} onChange={handleAstChange} />
            ))}
          </div>
        </div>

        {/* Remarks */}
        <div className="form-group">
          <label className="field-lbl">Remarks / Notes</label>
          <textarea value={remarks} onChange={(e) => setRemarks(e.target.value)} placeholder="Any additional notes or comments..." />
        </div>
      </div>

      <div className="btn-group">
        <button className="btn btn-primary btn-lg" onClick={handleSubmit}>
          💾 {isEditing ? "Update Entry" : "Save Entry"}
        </button>
        <button className="btn btn-ghost" onClick={clearForm}>🗑️ Clear</button>
        {isEditing && <button className="btn btn-ghost" onClick={clearForm}>✕ Cancel Edit</button>}
      </div>
    </div>
  );
};

// ── RECORDS PANE ─────────────────────────────────────────────────────────────
const RecordsPane = ({ records, onView, onEdit, onExport }) => {
  const [search, setSearch] = useState("");
  const [filterWard, setFilterWard] = useState("");
  const [filterSample, setFilterSample] = useState("");
  const [filterMonth, setFilterMonth] = useState("");

  const allWards = [...new Set(records.map((r) => r.ward).filter(Boolean))].sort();
  const allSamples = [...new Set(records.map((r) => r.sample).filter(Boolean))].sort();
  const allMonths = [...new Set(records.map((r) => (r.date || "").slice(0, 7)).filter(Boolean))].sort().reverse();

  const filtered = records.filter((r) => {
    const q = search.toLowerCase();
    return (
      (!q || [r.patientName, r.ward, r.organism, r.ipdopd, r.sample].some((f) => (f || "").toLowerCase().includes(q))) &&
      (!filterWard || r.ward === filterWard) &&
      (!filterSample || r.sample === filterSample) &&
      (!filterMonth || (r.date || "").startsWith(filterMonth))
    );
  });

  const sorted = [...filtered].sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  const pos = records.filter((r) => r.organism && r.organism !== "NG" && r.organism !== "NPG").length;

  return (
    <div>
      <div className="stats-row">
        <div className="stat-card stat-blue"><div className="stat-val">{records.length}</div><div className="stat-lbl">Total Entries</div></div>
        <div className="stat-card stat-green"><div className="stat-val">{pos}</div><div className="stat-lbl">Culture Positive</div></div>
        <div className="stat-card stat-orange"><div className="stat-val">{new Set(records.map((r) => r.ward)).size}</div><div className="stat-lbl">Wards</div></div>
        <div className="stat-card stat-red"><div className="stat-val">{records.length ? Math.round(pos / records.length * 100) : 0}%</div><div className="stat-lbl">Positivity</div></div>
      </div>

      <div className="card">
        <div className="card-hdr">
          <div className="card-title"><div className="card-ico">📋</div> Patient Records</div>
          <button className="btn btn-success btn-sm" onClick={onExport}>📥 Export</button>
        </div>

        <div className="search-bar">
          <div className="search-wrap">
            <input type="text" value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search name, ward, organism…" />
          </div>
        </div>
        <div className="filter-row">
          <select value={filterWard} onChange={(e) => setFilterWard(e.target.value)}>
            <option value="">All Wards</option>
            {allWards.map((w) => <option key={w}>{w}</option>)}
          </select>
          <select value={filterSample} onChange={(e) => setFilterSample(e.target.value)}>
            <option value="">All Samples</option>
            {allSamples.map((s) => <option key={s}>{s}</option>)}
          </select>
          <select value={filterMonth} onChange={(e) => setFilterMonth(e.target.value)}>
            <option value="">All Months</option>
            {allMonths.map((m) => (
              <option key={m} value={m}>
                {new Date(m + "-01").toLocaleDateString("en-IN", { month: "long", year: "numeric" })}
              </option>
            ))}
          </select>
        </div>

        {/* Desktop table */}
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>#</th><th>Date</th><th>IPD/OPD</th><th>Patient Name</th>
                <th>Ward</th><th>Sample</th><th>Organism</th><th>ZN</th><th>KOH</th><th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {sorted.map((r, i) => (
                <tr key={r.id}>
                  <td style={{ color: "var(--g500)" }}>{i + 1}</td>
                  <td style={{ fontWeight: 600 }}>{fmtDate(r.date)}</td>
                  <td><span className="badge badge-blue">{r.ipdopd || "—"}</span></td>
                  <td style={{ fontWeight: 600 }}>{r.patientName}</td>
                  <td><span className="badge badge-gray">{r.ward || "—"}</span></td>
                  <td>{r.sample || "—"}</td>
                  <td><OrgBadge org={r.organism} /></td>
                  <td><ZnBadge zn={r.znStain} /></td>
                  <td><KohBadge koh={r.koh} /></td>
                  <td style={{ whiteSpace: "nowrap" }}>
                    <button className="btn btn-outline btn-sm" onClick={() => onView(r.id)}>👁️ View</button>
                    <button className="btn btn-sm" style={{ background: "#E3F2FD", color: "#1565C0", border: "none", marginLeft: 4 }} onClick={() => onEdit(r.id)}>✏️</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* Mobile cards */}
        <div className="mob-records" style={{ display: "none" }}>
          {sorted.map((r) => (
            <div key={r.id} className="mob-card">
              <div className="mob-card-top">
                <div>
                  <div className="mob-card-name">{r.patientName}</div>
                  <div className="mob-card-date">{fmtDate(r.date)} &nbsp;·&nbsp; <span className="badge badge-blue" style={{ fontSize: 10 }}>{r.ipdopd || "—"}</span></div>
                </div>
                <OrgBadge org={r.organism} />
              </div>
              <div className="mob-card-meta">
                <span className="badge badge-gray">{r.ward || "—"}</span>
                <span className="badge badge-gray">{r.sample || "—"}</span>
                {r.znStain && <ZnBadge zn={r.znStain} />}
                {r.koh && <KohBadge koh={r.koh} />}
              </div>
              <div className="mob-card-actions">
                <button className="btn btn-outline btn-sm" onClick={() => onView(r.id)}>👁️ View</button>
                <button className="btn btn-ghost btn-sm" onClick={() => onEdit(r.id)}>✏️ Edit</button>
              </div>
            </div>
          ))}
        </div>

        {filtered.length === 0 && (
          <div className="empty">
            <div className="empty-ico">🔬</div>
            <div className="empty-ttl">No records found</div>
            <div style={{ fontSize: 12 }}>Add entries or adjust search filters</div>
          </div>
        )}
        {filtered.length > 0 && (
          <div style={{ padding: "10px 0 0", fontSize: 12, color: "var(--g500)" }}>
            Showing {filtered.length} of {records.length} records
          </div>
        )}
      </div>
    </div>
  );
};

// ── REPORT PANE ───────────────────────────────────────────────────────────────
const ReportPane = ({ records }) => {
  const pos = records.filter((r) => r.organism && r.organism !== "NG" && r.organism !== "NPG");
  const afb = records.filter((r) => {
    const z = (r.znStain || "").toUpperCase();
    return z.includes("AFB SEEN") && !z.includes("NOT") && !z.includes("BOT") && !z.includes("NOR");
  });
  return (
    <div>
      <div className="stats-row">
        <div className="stat-card stat-blue"><div className="stat-val">{records.length}</div><div className="stat-lbl">Total</div></div>
        <div className="stat-card stat-green"><div className="stat-val">{pos.length}</div><div className="stat-lbl">Culture +ve</div></div>
        <div className="stat-card stat-orange"><div className="stat-val">{records.length ? Math.round(pos.length / records.length * 100) : 0}%</div><div className="stat-lbl">Positivity</div></div>
        <div className="stat-card stat-red"><div className="stat-val">{afb.length}</div><div className="stat-lbl">AFB +ve</div></div>
      </div>
      <div className="report-grid">
        <div className="card">
          <div className="card-hdr"><div className="card-title"><div className="card-ico">🦠</div> Top Organisms</div></div>
          <BarChart data={countBy(pos, (r) => r.organism)} />
        </div>
        <div className="card">
          <div className="card-hdr"><div className="card-title"><div className="card-ico">🏥</div> By Ward</div></div>
          <BarChart data={countBy(records, (r) => r.ward)} />
        </div>
        <div className="card">
          <div className="card-hdr"><div className="card-title"><div className="card-ico">🧫</div> By Sample</div></div>
          <BarChart data={countBy(records, (r) => r.sample)} />
        </div>
        <div className="card">
          <div className="card-hdr"><div className="card-title"><div className="card-ico">📅</div> By Date</div></div>
          <BarChart data={countBy(records, (r) => r.date)} />
        </div>
      </div>
    </div>
  );
};

// ── EXCEL EXPORT ──────────────────────────────────────────────────────────────
function exportExcel(records, toast) {
  if (!records.length) { toast("No records to export", "error"); return; }

  const wb = XLSX.utils.book_new();
  const abxCodes = ALL_ABX.map((a) => a.code);
  const abxNames = ALL_ABX.map((a) => a.name);

  // Row 1: main title (merged across all columns)
  // Row 2: section sub-headers (Patient Info | Microbiology | Antibiotic Sensitivity Testing | )
  // Row 3: column headers
  const totalCols = 9 + abxCodes.length + 1; // 9 fixed + abx + remarks

  const makeRow = (r) => [
    r.date ? new Date(r.date + "T00:00:00").toLocaleDateString("en-IN", { day: "2-digit", month: "2-digit", year: "numeric" }) : "",
    r.ipdopd || "", r.patientName || "", r.ward || "", r.sample || "",
    r.gramStain || "", r.znStain || "", r.koh || "", r.organism || "",
    ...abxCodes.map((code) => {
      const v = (r.ast || {})[code];
      if (!v) return "";
      return typeof v === "object" ? (v.sir || "") : v;
    }),
    r.remarks || "",
  ];

  // Section sub-header row (row 2)
  // const sectionRow = [
  //   "Patient Information", "", "", "", "",          // cols 0-4 (Date, IPD, Name, Ward, Sample)
  //   "Microbiology", "", "", "",                     // cols 5-8 (Gram, ZN, KOH, Organism)
  //   ...abxCodes.map(() => "Antibiotic Sensitivity Testing (S / I / R)").map((v, i) => i === 0 ? v : ""), // first abx col has label, rest blank
  //   "Remarks",
  // ];

  // Column header row (row 3)
  const colHdrs = [
    "Date", "IPD / OPD No.", "Patient Name", "Ward", "Sample Type",
    "Gram Stain Result", "ZN Stain", "KOH", "Organism Isolated",
    ...abxCodes.map((code, i) => `${code}\n(${abxNames[i]})`),
    "Remarks / Notes",
  ];

  const colWidths = [
    { wch: 13 }, { wch: 12 }, { wch: 26 }, { wch: 10 }, { wch: 16 },
    { wch: 38 }, { wch: 18 }, { wch: 26 }, { wch: 28 },
    ...abxCodes.map(() => ({ wch: 14 })),
    { wch: 28 },
  ];

  const buildSheet = (recs, sheetTitle) => {
    const sortedRecs = [...recs].sort((a, b) => (a.date || "").localeCompare(b.date || ""));
    const dataRows = sortedRecs.map(makeRow);
    const aoa = [
      colHdrs,                                             // Row 1: Column Headers
      Array(totalCols).fill(""),                           // Row 2: Empty Row
      ...dataRows,
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = colWidths;

    // Style: column header row (row 1, index 0 - set background to black)
    for (let c = 0; c < totalCols; c++) {
      const ref = XLSX.utils.encode_cell({ r: 0, c });
      if (!ws[ref]) ws[ref] = { v: "" };
      ws[ref].s = {
        font: { bold: true, sz: 9, color: { rgb: "FFFFFF" } },
        fill: { fgColor: { rgb: "000000" } }, // Black background
        alignment: { horizontal: "center", vertical: "center", wrapText: true },
        border: {
          top: { style: "thin", color: { rgb: "455A64" } },
          bottom: { style: "medium", color: { rgb: "90A4AE" } },
        },
      };
    }

    // Style data rows: alternate row shading + SIR color coding
    dataRows.forEach((row, rowIdx) => {
      const r = rowIdx + 2; // offset by 2 header rows (Row 0, Row 1)
      const isAlt = rowIdx % 2 === 1;
      for (let c = 0; c < totalCols; c++) {
        const ref = XLSX.utils.encode_cell({ r, c });
        if (!ws[ref]) ws[ref] = { v: "" };
        const isAbxCol = c >= 9 && c < 9 + abxCodes.length;
        const sirVal = isAbxCol ? (row[c] || "") : "";
        let fillColor = isAlt ? "F5F8FF" : "FFFFFF";
        let fontColor = "202124";
        let bold = false;
        if (isAbxCol && sirVal) {
          if (sirVal === "S") { fillColor = "E8F5E9"; fontColor = "1B5E20"; bold = true; }
          else if (sirVal === "R") { fillColor = "FFEBEE"; fontColor = "B71C1C"; bold = true; }
          else if (sirVal === "I") { fillColor = "FFF3E0"; fontColor = "E65100"; bold = true; }
        }
        ws[ref].s = {
          font: { sz: 10, color: { rgb: fontColor }, bold },
          fill: { fgColor: { rgb: fillColor } },
          alignment: { vertical: "center", wrapText: c === 5 },
          border: {
            bottom: { style: "thin", color: { rgb: "E8EAED" } },
            right: { style: "thin", color: { rgb: "E8EAED" } },
          },
        };
      }
    });

    ws["!rows"] = [{ hpt: 36 }, { hpt: 20 }, ...dataRows.map(() => ({ hpt: 18 }))];
    return ws;
  };

  // Group records by month
  const byMonth = {};
  records.forEach((r) => {
    const m = r.date ? r.date.slice(0, 7) : "Unknown";
    if (!byMonth[m]) byMonth[m] = [];
    byMonth[m].push(r);
  });

  // Add one sheet per month — sheet name: "January 2025", "February 2025" etc.
  Object.entries(byMonth).sort().forEach(([month, recs]) => {
    let sheetName, sheetTitle;
    if (month === "Unknown") {
      sheetName = "Unknown";
      sheetTitle = "Hospital Laboratory Register — Unknown Date";
    } else {
      const d = new Date(month + "-01");
      const monthFull = d.toLocaleDateString("en-IN", { month: "long" });
      const year = d.getFullYear();
      sheetName = `${monthFull} ${year}`;                    // e.g. "January 2025"
      sheetTitle = `Hospital Laboratory Register — ${monthFull} ${year}`;
    }
    const ws = buildSheet(recs, sheetTitle);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  });

  // Summary "All Records" sheet
  const allTitle = `Hospital Laboratory Register — All Records (Exported ${new Date().toLocaleDateString("en-IN", { day: "2-digit", month: "long", year: "numeric" })})`;
  const wsAll = buildSheet(records, allTitle);
  XLSX.utils.book_append_sheet(wb, wsAll, "All Records");

  XLSX.writeFile(wb, `Lab_Register_${todayStr()}.xlsx`);
  toast("Excel exported successfully 📥", "success");
}

// ── APP ───────────────────────────────────────────────────────────────────────
export default function App() {
  const [tab, setTab] = useState("entry");
  const [records, setRecords] = useState(loadRecords);
  const [viewId, setViewId] = useState(null);
  const [editingRecord, setEditingRecord] = useState(null);
  const { toasts, toast } = useToasts();
  const styleInjected = useRef(false);

  // Authentication State
  const [isLoggedIn, setIsLoggedIn] = useState(() => {
    return localStorage.getItem("isLoggedIn") === "true";
  });
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [loginError, setLoginError] = useState("");

  const handleLogin = () => {
    if (username === "DV1404" && password === "1430") {
      setIsLoggedIn(true);
      localStorage.setItem("isLoggedIn", "true");
      setLoginError("");
    } else {
      setLoginError("Invalid username or password");
    }
  };

  const handleLogout = () => {
    if (window.confirm("Are you sure you want to logout?")) {
      setIsLoggedIn(false);
      localStorage.removeItem("isLoggedIn");
      setUsername("");
      setPassword("");
    }
  };

  useEffect(() => {
    if (!styleInjected.current) {
      const s = document.createElement("style");
      s.textContent = CSS;
      document.head.appendChild(s);
      styleInjected.current = true;
    }
  }, []);

  useEffect(() => { saveRecords(records); }, [records]);

  const handleSave = (data) => {
    setRecords((prev) => {
      const idx = prev.findIndex((r) => r.id === data.id);
      if (idx >= 0) {
        const next = [...prev];
        next[idx] = { ...next[idx], ...data, updatedAt: new Date().toISOString() };
        toast("Record updated ✓", "success");
        return next;
      }
      toast("Entry saved! 🎉", "success");
      return [...prev, { ...data, createdAt: new Date().toISOString() }];
    });
    setEditingRecord(null);
  };

  const handleEdit = (id) => {
    const r = records.find((x) => x.id === id);
    if (!r) return;
    setEditingRecord(r);
    setViewId(null);
    setTab("entry");
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  const handleDelete = (id) => {
    if (!window.confirm("Delete this record permanently?")) return;
    setRecords((prev) => prev.filter((r) => r.id !== id));
    setViewId(null);
    toast("Record deleted", "error");
  };

  const viewRecord = records.find((r) => r.id === viewId);

  useEffect(() => {
    const handler = (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === "Enter") document.querySelector(".btn-primary")?.click();
      if (e.key === "Escape") setViewId(null);
    };
    document.addEventListener("keydown", handler);
    return () => document.removeEventListener("keydown", handler);
  }, []);

  if (!isLoggedIn) {
    return (
      <div className="login-overlay">
        <div className="login-card">
          <div className="login-logo">🔬</div>
          <h2 className="login-title">Hospital Lab Portal</h2>
          <p className="login-subtitle">Enter your credentials to access the register</p>

          <div className="login-form">
            <div className="form-group" style={{ marginBottom: 16 }}>
              <label className="field-lbl">Username</label>
              <input
                type="text"
                value={username}
                onChange={(e) => { setUsername(e.target.value); setLoginError(""); }}
                // placeholder="e.g. DV1404"
                className={loginError ? "input-err" : ""}
                onKeyDown={(e) => e.key === "Enter" && handleLogin()}
              />
            </div>

            <div className="form-group" style={{ marginBottom: 20 }}>
              <label className="field-lbl">Password</label>
              <input
                type="password"
                value={password}
                onChange={(e) => { setPassword(e.target.value); setLoginError(""); }}
                // placeholder="••••"
                className={loginError ? "input-err" : ""}
                onKeyDown={(e) => e.key === "Enter" && handleLogin()}
              />
            </div>

            {loginError && (
              <div className="login-error">
                ⚠️ {loginError}
              </div>
            )}

            <button className="btn btn-primary btn-lg" style={{ width: "100%", justifyContent: "center", marginTop: 8 }} onClick={handleLogin}>
              🔑 Secure Login
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <>
      {/* HEADER */}
      <header className="hdr">
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div className="hdr-ico">🔬</div>
          <div>
            <div className="hdr-title">Hospital Laboratory Register</div>
            <div className="hdr-sub">Microbiology &amp; Culture Report System</div>
          </div>
        </div>
        <div>
          <button className="hdr-btn" onClick={() => exportExcel(records, toast)}>📥 Export</button>
          <button className="hdr-btn" style={{ marginLeft: 8, background: "rgba(198,40,40,0.18)", borderColor: "rgba(198,40,40,0.35)" }} onClick={handleLogout}>🚪 Logout</button>
        </div>
      </header>

      {/* TABS */}
      <div className="tabs">
        <button className={`tab${tab === "entry" ? " active" : ""}`} onClick={() => setTab("entry")}>✏️ New Entry</button>
        <button className={`tab${tab === "records" ? " active" : ""}`} onClick={() => setTab("records")}>
          📋 Records <span className="tab-badge">{records.length}</span>
        </button>
        <button className={`tab${tab === "report" ? " active" : ""}`} onClick={() => setTab("report")}>📊 Summary</button>
      </div>

      {/* MAIN */}
      <div className="main">
        {tab === "entry" && (
          <EntryPane
            onSave={handleSave}
            editingRecord={editingRecord}
            onCancelEdit={() => setEditingRecord(null)}
            toast={toast}
          />
        )}
        {tab === "records" && (
          <RecordsPane
            records={records}
            onView={setViewId}
            onEdit={handleEdit}
            onExport={() => exportExcel(records, toast)}
          />
        )}
        {tab === "report" && <ReportPane records={records} />}
      </div>

      {/* MODAL */}
      {viewId && (
        <ViewModal
          record={viewRecord}
          onClose={() => setViewId(null)}
          onEdit={handleEdit}
          onDelete={handleDelete}
        />
      )}

      {/* TOASTS */}
      <div className="toast-con">
        {toasts.map((t) => (
          <div key={t.id} className={`toast toast-${t.type}`}>
            {t.type === "success" ? "✅" : t.type === "error" ? "❌" : "ℹ️"} {t.msg}
          </div>
        ))}
      </div>
    </>
  );
}
