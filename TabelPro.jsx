import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  Users, CalendarDays, Ban, AlarmClock, Repeat, Timer, BarChart3,
  Search, Plus, History, MessageSquare, Download, FileSpreadsheet,
  ChevronLeft, ChevronRight, X, Phone, Cake, Briefcase, Trash2,
  Pencil, Send, LogOut, Building2, Check, Clock, KeyRound, Lock,
  ShieldCheck, RotateCcw, Eye, EyeOff, Sun, Moon, Plane
} from "lucide-react";

/* ============================================================
   ТАБЕЛЬ PRO — облік робочого часу
   • ролі та кабінети  • тема світла/темна
   • лікарняні/відпустки з картки → авто в табелі
   • спільне сховище (синхронізація між пристроями)
============================================================ */

const MONTHS = ["Січень","Лютий","Березень","Квітень","Травень","Червень","Липень","Серпень","Вересень","Жовтень","Листопад","Грудень"];
const MONTHS_SHORT = ["Січ","Лют","Бер","Кві","Тра","Чер","Лип","Сер","Вер","Жов","Лис","Гру"];
const WD = ["НД","ПН","ВТ","СР","ЧТ","ПТ","СБ"];

const STATUSES = {
  present:  { label: "Вчасно",      dot: "#16a34a", bg: "#dcfce7", text: "#166534" },
  late:     { label: "Запізнення",  dot: "#d97706", bg: "#fef3c7", text: "#92400e" },
  early:    { label: "Рано пішов",  dot: "#ea580c", bg: "#ffedd5", text: "#9a3412" },
  absent:   { label: "Відсутній",   dot: "#dc2626", bg: "#fee2e2", text: "#991b1b" },
  overtime: { label: "Надурочно",   dot: "#2563eb", bg: "#dbeafe", text: "#1e40af" },
  vacation: { label: "Відпустка",   dot: "#9333ea", bg: "#f3e8ff", text: "#6b21a8" },
  sick:     { label: "Лікарняний",  dot: "#0891b2", bg: "#cffafe", text: "#155e63" },
};
const EXCUSED = { vacation: 1, sick: 1 };

const SHIFTS = [
  { id: "s1", label: "8:00–17:00", start: "08:00", end: "17:00" },
  { id: "s2", label: "7:30–20:30", start: "07:30", end: "20:30" },
];

const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
const pad = (n) => String(n).padStart(2, "0");
const isoDay = (y, m, d) => `${y}-${pad(m + 1)}-${pad(d)}`;

/* ---------- тема (CSS-змінні) ---------- */
const THEME_CSS = `
.tp-root{--bg:linear-gradient(180deg,#f5f8fd 0%,#e9eef7 100%);--bg-solid:#eef2f7;--panel:#ffffff;--card:#ffffff;--elev:#ffffff;--text:#0f172a;--soft:#5b6678;--muted:#98a2b5;--border:#e7ebf3;--hover:#f1f4fa;--input:#ffffff;--inputbd:#e0e6f0;--th:#f7f9fc;--weekend:#eef1f7;--rowhover:#f8fafd;--today:#eef1ff;--overlay:rgba(15,23,42,.45);--brand1:#6366f1;--brand2:#8b5cf6;--shadow:0 1px 2px rgba(16,24,40,.04),0 10px 26px -14px rgba(16,24,40,.16);color-scheme:light;}
.tp-root.dark{--bg:linear-gradient(180deg,#0b1226 0%,#070c18 100%);--bg-solid:#0a0f1f;--panel:#0f1626;--card:#141c30;--elev:#172037;--text:#e8edf7;--soft:#9aa6bd;--muted:#6b7794;--border:#222c45;--hover:#1a2238;--input:#161e33;--inputbd:#2c3650;--th:#131b2d;--weekend:#0e1525;--rowhover:#161e31;--today:#1b2750;--overlay:rgba(0,0,0,.62);--brand1:#818cf8;--brand2:#a78bfa;--shadow:0 1px 2px rgba(0,0,0,.3),0 14px 32px -18px rgba(0,0,0,.65);color-scheme:dark;}
.s-app{background:var(--bg);color:var(--text);}
.s-panel{background:var(--panel);}
.s-card{background:var(--card);}
.s-elev{background:var(--elev);}
.s-soft{color:var(--soft);}
.s-muted{color:var(--muted);}
.s-bd{border-color:var(--border);}
.s-hover:hover{background:var(--hover);}
.s-th{background:var(--th);}
.s-weekend{background:var(--weekend);}
.s-rowhover:hover{background:var(--rowhover);}
.s-input{background:var(--input);border:1px solid var(--inputbd);color:var(--text);transition:border-color .15s,box-shadow .15s;}
.s-input:focus{border-color:var(--brand1);outline:none;box-shadow:0 0 0 3px rgba(99,102,241,.20);}
.s-input::placeholder{color:var(--muted);}
.s-grad{background:linear-gradient(135deg,var(--brand1),var(--brand2));color:#fff;}
.s-grad:hover{filter:brightness(1.06);}
.s-shadow{box-shadow:var(--shadow);}
.s-lift{transition:transform .16s ease,box-shadow .16s ease;}
.s-lift:hover{transform:translateY(-2px);box-shadow:0 14px 32px -16px rgba(16,24,40,.28);}
.s-glow{box-shadow:0 8px 24px -8px rgba(99,102,241,.55);}
.f9{font-size:9px}.f10{font-size:10px}.f11{font-size:11px;line-height:1.25}.f12{font-size:12px}
`;

/* ---------- сховище ---------- */
const memStore = {};
const hasStorage = typeof window !== "undefined" && window.storage && typeof window.storage.get === "function";
async function loadKey(key, fallback, shared = true) {
  if (!hasStorage) { const k = (shared ? "S:" : "P:") + key; return k in memStore ? memStore[k] : fallback; }
  try { const r = await window.storage.get(key, shared); if (r && typeof r.value === "string") return JSON.parse(r.value); return fallback; }
  catch (e) { return fallback; }
}
async function saveKey(key, value, shared = true) {
  if (!hasStorage) { memStore[(shared ? "S:" : "P:") + key] = value; return true; }
  try { await window.storage.set(key, JSON.stringify(value), shared); return true; }
  catch (e) { memStore[(shared ? "S:" : "P:") + key] = value; return false; }
}
function hashPin(pin) {
  let h = 2166136261 >>> 0; const s = "tp::" + String(pin);
  for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619) >>> 0; }
  return h.toString(16);
}

/* ---------- час ---------- */
function parseMin(t) { if (!t || typeof t !== "string") return null; const [h, m] = t.split(":").map(Number); if (Number.isNaN(h) || Number.isNaN(m)) return null; return h * 60 + m; }
function dayHours(rec) { if (!rec) return 0; const a = parseMin(rec.in), b = parseMin(rec.out); if (a == null || b == null || b <= a) return 0; return (b - a) / 60; }
const daysInMonth = (y, m) => new Date(y, m + 1, 0).getDate();
function workingDays(y, m) { let n = 0; const d = daysInMonth(y, m); for (let i = 1; i <= d; i++) { const wd = new Date(y, m, i).getDay(); if (wd !== 0 && wd !== 6) n++; } return n; }
const fmt = (n) => { const r = Math.round(n * 10) / 10; return Number.isInteger(r) ? String(r) : r.toFixed(1); };
function leaveOn(emp, dateStr) {
  if (!emp || !Array.isArray(emp.leaves)) return null;
  for (const l of emp.leaves) { if (l.from && l.to && dateStr >= l.from && dateStr <= l.to) return l.type; }
  return null;
}

/* ---------- міграція ---------- */
async function migrateOldData() {
  const newMeta = await loadKey("tp_meta", null, true);
  if (newMeta) return;
  const oldMeta = await loadKey("td_meta", null, false);
  if (!oldMeta || !Array.isArray(oldMeta.departments) || !oldMeta.departments.length) return;
  await saveKey("tp_meta", oldMeta, true);
  const y = new Date().getFullYear();
  for (const d of oldMeta.departments) {
    const emp = await loadKey("td_emp_" + d.id, [], false);
    if (emp.length) await saveKey("tp_emp_" + d.id, emp, true);
    for (let m = 0; m < 12; m++) {
      const att = await loadKey(`td_att_${d.id}_${y}_${m}`, null, false);
      if (att && Object.keys(att).length) await saveKey(`tp_att_${d.id}_${y}_${m}`, att, true);
    }
  }
  const chat = await loadKey("td_chat", [], false); if (chat.length) await saveKey("tp_chat", chat, true);
  const log = await loadKey("td_log", [], false); if (log.length) await saveKey("tp_log", log, true);
}

/* ============================================================ */
export default function TabelPro() {
  const now = new Date();
  const [booted, setBooted] = useState(false);
  const [theme, setTheme] = useState("light");
  const [departments, setDepartments] = useState([]);
  const [users, setUsers] = useState({});
  const [session, setSession] = useState(null);

  const [deptId, setDeptId] = useState(null);
  const [year, setYear] = useState(now.getFullYear());
  const [month, setMonth] = useState(now.getMonth());
  const [employees, setEmployees] = useState([]);
  const [attendance, setAttendance] = useState({});
  const [log, setLog] = useState([]);
  const [chat, setChat] = useState([]);

  const [search, setSearch] = useState("");
  const [newName, setNewName] = useState("");
  const [editing, setEditing] = useState(null);
  const [profile, setProfile] = useState(null);
  const [showChat, setShowChat] = useState(false);
  const [showHistory, setShowHistory] = useState(false);
  const [showReport, setShowReport] = useState(false);
  const [showAccess, setShowAccess] = useState(false);
  const [addDeptOpen, setAddDeptOpen] = useState(false);
  const [newDeptName, setNewDeptName] = useState("");
  const [clock, setClock] = useState(new Date());
  const [toast, setToast] = useState(null);

  const role = session?.role;
  const isAdmin = role === "admin";
  const isHR = role === "hr";
  const isManager = role === "dept";
  const canSwitch = isAdmin || isHR;

  useEffect(() => { const t = setInterval(() => setClock(new Date()), 30000); return () => clearInterval(t); }, []);

  useEffect(() => {
    (async () => {
      try { await migrateOldData(); } catch (e) {}
      setTheme(await loadKey("tp_theme", "light", false));
      let meta = await loadKey("tp_meta", null, true);
      if (!meta || !Array.isArray(meta.departments) || meta.departments.length === 0) {
        const seed = [{ id: uid(), name: "Відділ 1" }, { id: uid(), name: "Відділ 2" }, { id: uid(), name: "Відділ 3" }];
        meta = { departments: seed };
        await saveKey("tp_meta", meta, true);
        const demo = [
          { id: uid(), name: "Бондар Юлія",     start: "08:00", end: "17:00", phone: "", birthday: "", position: "Менеджер", leaves: [] },
          { id: uid(), name: "Шевченко Дмитро", start: "08:00", end: "17:00", phone: "", birthday: "", position: "Спеціаліст", leaves: [] },
          { id: uid(), name: "Ткач Наталя",     start: "08:00", end: "17:00", phone: "", birthday: "", position: "Бухгалтер", leaves: [] },
        ];
        await saveKey("tp_emp_" + seed[2].id, demo, true);
      }
      setDepartments(meta.departments);
      setUsers(await loadKey("tp_users", {}, true));
      setLog(await loadKey("tp_log", [], true));
      setChat(await loadKey("tp_chat", [], true));
      const sess = await loadKey("tp_session", null, false);
      if (sess && validSession(sess, meta.departments)) { setSession(sess); setDeptId(sess.role === "dept" ? sess.deptId : meta.departments[0].id); }
      setBooted(true);
    })();
  }, []);

  function validSession(s, depts) {
    if (!s || !s.role) return false;
    if (s.role === "dept") return depts.some((d) => d.id === s.deptId);
    return s.role === "admin" || s.role === "hr";
  }

  useEffect(() => { if (deptId) (async () => setEmployees(await loadKey("tp_emp_" + deptId, [], true)))(); }, [deptId]);
  const attKey = deptId ? `tp_att_${deptId}_${year}_${month}` : null;
  useEffect(() => { if (attKey) (async () => setAttendance(await loadKey(attKey, {}, true)))(); }, [attKey]);

  const persistDepartments = useCallback(async (list) => { setDepartments(list); await saveKey("tp_meta", { departments: list }, true); }, []);
  const persistEmployees = useCallback(async (list) => { setEmployees(list); if (deptId) await saveKey("tp_emp_" + deptId, list, true); }, [deptId]);
  const persistAttendance = useCallback(async (data) => { setAttendance(data); if (attKey) await saveKey(attKey, data, true); }, [attKey]);
  const persistUsers = useCallback(async (u) => { setUsers(u); await saveKey("tp_users", u, true); }, []);
  const pushLog = useCallback((text) => { setLog((prev) => { const next = [{ id: uid(), text, ts: Date.now() }, ...prev].slice(0, 200); saveKey("tp_log", next, true); return next; }); }, []);
  const flash = (msg) => { setToast(msg); setTimeout(() => setToast(null), 2200); };
  const toggleTheme = async () => { const t = theme === "dark" ? "light" : "dark"; setTheme(t); await saveKey("tp_theme", t, false); };

  async function startSession(sess) { setSession(sess); setDeptId(sess.role === "dept" ? sess.deptId : departments[0]?.id); await saveKey("tp_session", sess, false); }
  async function handleLogin(target, pin) {
    const key = target.role === "dept" ? target.deptId : target.role;
    const u = users[key];
    if (!u || !u.pin) return "Доступ ще не налаштовано керуючим";
    if (hashPin(pin) !== u.pin) return "Невірний PIN-код";
    const label = target.role === "admin" ? "Керуючий" : target.role === "hr" ? "Відділ кадрів" : (departments.find((d) => d.id === target.deptId)?.name || "");
    await startSession({ role: target.role, deptId: target.role === "dept" ? target.deptId : null, label });
    return null;
  }
  async function handleCreateAdmin(pin) { const next = { ...users, admin: { pin: hashPin(pin) } }; await persistUsers(next); await startSession({ role: "admin", deptId: null, label: "Керуючий" }); }
  async function logout() { setSession(null); await saveKey("tp_session", null, false); }

  const deptName = useMemo(() => departments.find((d) => d.id === deptId)?.name || "", [departments, deptId]);
  const visibleDepts = canSwitch ? departments : departments.filter((d) => d.id === session?.deptId);
  const visibleEmployees = useMemo(() => { const q = search.trim().toLowerCase(); return q ? employees.filter((e) => e.name.toLowerCase().includes(q)) : employees; }, [employees, search]);

  /* ---- розрахунок з урахуванням лікарняних/відпусток ---- */
  const computeEmp = useCallback((emp) => {
    const rec = attendance[emp.id] || {};
    const dN = daysInMonth(year, month);
    let worked = 0, normDays = 0, absent = 0, late = 0, overtime = 0, excused = 0;
    for (let d = 1; d <= dN; d++) {
      const wd = new Date(year, month, d).getDay();
      const isWork = wd !== 0 && wd !== 6;
      const manual = rec[d];
      let status = manual ? manual.status : null;
      if (!manual) { const lv = leaveOn(emp, isoDay(year, month, d)); if (lv) status = lv; }
      if (manual) {
        worked += dayHours(manual);
        if (manual.status === "absent") absent++;
        if (manual.status === "late") late++;
        if (manual.status === "overtime") overtime += dayHours(manual);
      }
      if (isWork) { if (status && EXCUSED[status]) excused++; else normDays++; }
    }
    const norm = normDays * 8;
    return { worked, norm, absent, late, overtime, excused, balance: worked - norm };
  }, [attendance, year, month]);

  const stats = useMemo(() => {
    const wDays = workingDays(year, month);
    let worked = 0, absent = 0, late = 0, overtime = 0, norm = 0, excused = 0;
    for (const emp of employees) { const c = computeEmp(emp); worked += c.worked; absent += c.absent; late += c.late; overtime += c.overtime; norm += c.norm; excused += c.excused; }
    return { wDays, norm, worked, absent, late, overtime, excused, balance: worked - norm };
  }, [employees, computeEmp, year, month]);

  /* ---- дії ---- */
  const addEmployee = async () => { const name = newName.trim(); if (!name) return; await persistEmployees([...employees, { id: uid(), name, start: "08:00", end: "17:00", phone: "", birthday: "", position: "", leaves: [] }]); setNewName(""); pushLog(`${session.label}: додано «${name}»`); flash("Співробітника додано"); };
  const removeEmployee = async (id) => { const e = employees.find((x) => x.id === id); await persistEmployees(employees.filter((x) => x.id !== id)); pushLog(`${session.label}: видалено «${e?.name || ""}»`); setProfile(null); };
  const saveProfile = async (data) => { await persistEmployees(employees.map((e) => (e.id === data.id ? data : e))); pushLog(`${session.label}: оновлено картку «${data.name}»`); setProfile(null); flash("Дані збережено"); };
  const saveCell = async (empId, day, rec) => {
    const next = { ...attendance }; const empRec = { ...(next[empId] || {}) };
    const clean = rec && (rec.in || rec.out || rec.status || rec.note);
    if (clean) empRec[day] = rec; else delete empRec[day];
    if (Object.keys(empRec).length) next[empId] = empRec; else delete next[empId];
    await persistAttendance(next); setEditing(null);
  };
  const addDepartment = async () => { const name = newDeptName.trim(); if (!name) return; const d = { id: uid(), name }; await persistDepartments([...departments, d]); setNewDeptName(""); setAddDeptOpen(false); setDeptId(d.id); pushLog(`Створено відділ «${name}»`); flash("Відділ створено"); };
  const removeDepartment = async (id) => {
    if (departments.length <= 1) { flash("Має лишитись хоча б один відділ"); return; }
    const d = departments.find((x) => x.id === id); const list = departments.filter((x) => x.id !== id);
    await persistDepartments(list); if (deptId === id) setDeptId(list[0].id);
    const nu = { ...users }; delete nu[id]; await persistUsers(nu); pushLog(`Видалено відділ «${d?.name || ""}»`);
  };
  const sendMessage = async (text) => {
    const t = text.trim(); if (!t) return;
    const ctxDeptId = isManager ? session.deptId : deptId;
    const ctxName = departments.find((d) => d.id === ctxDeptId)?.name || "";
    const msg = { id: uid(), deptId: ctxDeptId, deptName: ctxName, fromLabel: session.label, text: t, ts: Date.now() };
    const next = [...chat, msg]; setChat(next); await saveKey("tp_chat", next, true);
  };
  const setRolePin = async (key, pin) => { const next = { ...users, [key]: { pin: hashPin(pin) } }; await persistUsers(next); flash("PIN збережено"); };
  const clearRolePin = async (key) => { const next = { ...users }; delete next[key]; await persistUsers(next); flash("Доступ скинуто"); };

  const exportExcel = () => {
    const dN = daysInMonth(year, month);
    const head = ["Працівник", "Посада", "Телефон", "ДР"];
    for (let d = 1; d <= dN; d++) head.push(`${d} ${WD[new Date(year, month, d).getDay()]}`);
    head.push("Відпрацьовано", "Норма", "Баланс");
    const rows = [head];
    for (const emp of employees) {
      const rec = attendance[emp.id] || {};
      const row = [emp.name, emp.position || "", emp.phone || "", emp.birthday || ""];
      for (let d = 1; d <= dN; d++) {
        const manual = rec[d];
        if (manual) { const h = dayHours(manual); row.push(h > 0 ? fmt(h) : (manual.status ? STATUSES[manual.status]?.label || "" : "")); }
        else { const lv = leaveOn(emp, isoDay(year, month, d)); row.push(lv ? STATUSES[lv].label : ""); }
      }
      const c = computeEmp(emp);
      row.push(fmt(c.worked), c.norm, fmt(c.balance));
      rows.push(row);
    }
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 22 }, { wch: 14 }, { wch: 14 }, { wch: 12 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `${MONTHS_SHORT[month]} ${year}`);
    XLSX.writeFile(wb, `Табель_${deptName.replace(/[^\wа-яіїєґ\d]+/gi, "_")}_${MONTHS[month]}_${year}.xlsx`);
    pushLog(`${session.label}: експорт Excel ${MONTHS[month]} ${year}`); flash("Файл Excel сформовано");
  };

  /* ============================ РЕНДЕР ============================ */
  if (!booted) {
    return (<div className="min-h-screen flex items-center justify-center" style={{ background: "#eef2f7", color: "#64748b" }}><div className="flex items-center gap-3"><div className="w-5 h-5 rounded-full border-2 border-indigo-500 border-t-transparent animate-spin" /> Завантаження…</div></div>);
  }
  if (!session) {
    return (<><style>{THEME_CSS}</style><LoginScreen departments={departments} users={users} onLogin={handleLogin} onCreateAdmin={handleCreateAdmin} /></>);
  }

  const dN = daysInMonth(year, month);
  const dayList = Array.from({ length: dN }, (_, i) => i + 1);

  return (
    <div className={`tp-root ${theme === "dark" ? "dark" : ""} s-app min-h-screen flex`} style={{ fontFamily: "Inter, ui-sans-serif, system-ui, sans-serif" }}>
      <style>{THEME_CSS}</style>

      {/* SIDEBAR */}
      <aside className="w-60 shrink-0 s-panel border-r s-bd flex flex-col">
        <div className="px-5 py-4 border-b s-bd flex items-center gap-2">
          <div className="w-8 h-8 rounded-xl s-grad s-glow flex items-center justify-center text-white"><CalendarDays size={18} /></div>
          <div className="font-semibold tracking-tight">Табель <span className="text-indigo-500">PRO</span></div>
        </div>
        <div className="px-4 pt-4 pb-2 f11 font-semibold tracking-wider s-muted">{canSwitch ? "ПІДРОЗДІЛИ" : "МІЙ ВІДДІЛ"}</div>
        <nav className="flex-1 overflow-y-auto px-2 space-y-1">
          {visibleDepts.map((d) => (
            <div key={d.id} className="group flex items-center">
              <button onClick={() => { if (canSwitch) { setDeptId(d.id); setSearch(""); } }} className={`flex-1 flex items-center gap-2 px-3 py-2 rounded-lg text-sm text-left transition ${d.id === deptId ? "s-grad s-glow" : "s-soft s-hover"} ${!canSwitch ? "cursor-default" : ""}`}>
                <span className="w-1.5 h-1.5 rounded-full" style={{ background: d.id === deptId ? "#fff" : "var(--muted)" }} />{d.name}
              </button>
              {isAdmin && departments.length > 1 && (<button onClick={() => removeDepartment(d.id)} className="opacity-0 group-hover:opacity-100 p-1 s-muted hover:text-red-500" title="Видалити відділ"><Trash2 size={14} /></button>)}
            </div>
          ))}
          {isAdmin && (addDeptOpen ? (
            <div className="px-2 py-2 space-y-2">
              <input autoFocus value={newDeptName} onChange={(e) => setNewDeptName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addDepartment()} placeholder="Назва відділу" className="w-full px-3 py-2 text-sm rounded-lg s-input" />
              <div className="flex gap-2"><button onClick={addDepartment} className="flex-1 py-1.5 text-sm rounded-lg bg-indigo-600 text-white hover:bg-indigo-700">Додати</button><button onClick={() => { setAddDeptOpen(false); setNewDeptName(""); }} className="px-3 py-1.5 text-sm rounded-lg border s-bd s-soft">Скасувати</button></div>
            </div>
          ) : (<button onClick={() => setAddDeptOpen(true)} className="w-full flex items-center gap-2 px-3 py-2 mt-1 rounded-lg text-sm text-indigo-500 s-hover"><Plus size={16} /> Додати відділ</button>))}
        </nav>
        <div className="p-3 border-t s-bd space-y-1">
          {isAdmin && (<button onClick={() => setShowAccess(true)} className="w-full flex items-center gap-2 px-3 py-2 rounded-lg text-sm text-indigo-500 s-hover"><ShieldCheck size={16} /> Доступи та PIN</button>)}
          <button onClick={() => setShowReport(true)} className="w-full flex items-center gap-2 px-3 py-2 rounded-lg text-sm s-soft s-hover"><BarChart3 size={16} /> Звіт за місяць</button>
          <button onClick={exportExcel} className="w-full flex items-center gap-2 px-3 py-2 rounded-lg text-sm text-emerald-600 s-hover"><FileSpreadsheet size={16} /> Експорт у Excel</button>
          {(isAdmin || isHR) && (<button onClick={() => setShowHistory(true)} className="w-full flex items-center gap-2 px-3 py-2 rounded-lg text-sm s-soft s-hover"><History size={16} /> Історія дій</button>)}
        </div>
        <div className="p-3 border-t s-bd">
          <div className="px-2 mb-2"><div className="f11 s-muted">Ви увійшли як</div><div className="text-sm font-semibold flex items-center gap-1.5">{isAdmin ? <KeyRound size={14} className="text-amber-500" /> : isHR ? <Users size={14} className="text-sky-500" /> : <Building2 size={14} className="text-indigo-500" />}{session.label}</div></div>
          <button onClick={logout} className="w-full flex items-center gap-2 px-3 py-2 rounded-lg text-sm s-soft hover:text-red-600 s-hover"><LogOut size={16} /> Вийти</button>
        </div>
      </aside>

      {/* MAIN */}
      <main className="flex-1 min-w-0 flex flex-col">
        <header className="h-14 s-panel s-shadow border-b s-bd flex items-center gap-3 px-4 relative z-10">
          <div className="flex items-center gap-2 px-3 py-1.5 rounded-lg s-card border s-bd text-sm font-mono"><Clock size={14} className="s-muted" /> {pad(clock.getDate())}.{pad(clock.getMonth() + 1)} · {pad(clock.getHours())}:{pad(clock.getMinutes())}</div>
          <div className="flex items-center gap-2 font-semibold"><Building2 size={16} className="text-indigo-500" /> {deptName}</div>
          <div className="flex items-center gap-1 ml-2">
            <button onClick={() => setYear((y) => y - 1)} className="p-1.5 rounded-lg s-hover s-soft"><ChevronLeft size={16} /></button>
            <span className="px-2 font-semibold tabular-nums">{year}</span>
            <button onClick={() => setYear((y) => y + 1)} className="p-1.5 rounded-lg s-hover s-soft"><ChevronRight size={16} /></button>
          </div>
          <div className="ml-auto flex items-center gap-2">
            <button onClick={toggleTheme} className="p-2 rounded-lg s-hover s-soft" title="Перемкнути тему">{theme === "dark" ? <Sun size={16} /> : <Moon size={16} />}</button>
            <button onClick={() => setShowChat(true)} className="relative flex items-center gap-2 px-3 py-1.5 rounded-lg bg-violet-500 text-white hover:bg-violet-600 text-sm font-medium"><MessageSquare size={16} /> Чат {chat.length > 0 && <span className="f10 bg-white text-violet-600 rounded-full px-1.5 py-0.5">{chat.length}</span>}</button>
            <button onClick={exportExcel} className="flex items-center gap-2 px-3 py-1.5 rounded-lg s-grad text-sm font-medium"><Download size={16} /> Звіт</button>
          </div>
        </header>

        <div className="s-panel border-b s-bd px-4 overflow-x-auto">
          <div className="flex gap-1 py-2">{MONTHS.map((m, i) => (<button key={m} onClick={() => setMonth(i)} className={`px-3 py-1.5 rounded-lg text-sm whitespace-nowrap transition ${i === month ? "s-grad font-semibold" : "s-soft s-hover"}`}>{m}</button>))}</div>
        </div>

        <div className="flex-1 overflow-auto p-4 space-y-4">
          <div className="grid grid-cols-2 md:grid-cols-4 xl:grid-cols-7 gap-3">
            <StatCard icon={<Users size={16} />} label="Працівників" value={employees.length} color="#6366f1" />
            <StatCard icon={<CalendarDays size={16} />} label="Роб. днів" value={stats.wDays} color="#64748b" />
            <StatCard icon={<Ban size={16} />} label="Відсутніх" value={stats.absent} color="#ef4444" />
            <StatCard icon={<AlarmClock size={16} />} label="Запізнень" value={stats.late} color="#f59e0b" />
            <StatCard icon={<Plane size={16} />} label="Відп./лік." value={stats.excused} color="#9333ea" />
            <StatCard icon={<Repeat size={16} />} label="Відпрацьовано" value={fmt(stats.worked) + "г"} color="#10b981" />
            <StatCard icon={<BarChart3 size={16} />} label={`Баланс · норма ${stats.norm}г`} value={(stats.balance >= 0 ? "+" : "") + fmt(stats.balance) + "г"} color={stats.balance >= 0 ? "#10b981" : "#ef4444"} />
          </div>

          <div className="flex flex-wrap items-center gap-3">
            <div className="relative flex-1" style={{ minWidth: 220 }}>
              <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 s-muted" />
              <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Пошук співробітника…" className="w-full pl-9 pr-9 py-2.5 rounded-xl s-input text-sm" />
              {search && <button onClick={() => setSearch("")} className="absolute right-3 top-1/2 -translate-y-1/2 s-muted hover:text-red-500"><X size={16} /></button>}
            </div>
            <div className="flex items-center gap-2">
              <input value={newName} onChange={(e) => setNewName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addEmployee()} placeholder="ПІБ нового працівника…" className="px-3 py-2.5 rounded-xl s-input text-sm w-56" />
              <button onClick={addEmployee} className="flex items-center gap-2 px-4 py-2.5 rounded-xl s-grad text-sm font-medium s-glow"><Plus size={16} /> Додати</button>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-x-4 gap-y-1 text-xs s-soft">{Object.entries(STATUSES).map(([k, s]) => (<span key={k} className="flex items-center gap-1.5"><span className="w-2.5 h-2.5 rounded-full" style={{ background: s.dot }} /> {s.label}</span>))}</div>

          <div className="s-card s-shadow rounded-2xl border s-bd overflow-hidden">
            <div className="overflow-x-auto">
              <table className="border-collapse text-sm">
                <thead>
                  <tr className="s-th border-b s-bd">
                    <th className="sticky left-0 z-20 s-th text-left px-4 py-3 font-semibold s-soft border-r s-bd" style={{ minWidth: 200 }}>ПРАЦІВНИК</th>
                    {dayList.map((d) => {
                      const wd = new Date(year, month, d).getDay(); const weekend = wd === 0 || wd === 6;
                      const today = year === now.getFullYear() && month === now.getMonth() && d === now.getDate();
                      return (<th key={d} className={`px-1 py-2 text-center ${weekend ? "s-weekend" : ""}`} style={{ minWidth: 58, background: today ? "var(--today)" : undefined }}><div className={`font-bold ${today ? "text-indigo-500" : ""}`}>{d}</div><div className="f10" style={{ color: weekend ? "#f87171" : "var(--muted)" }}>{WD[wd]}</div></th>);
                    })}
                    <th className="px-3 py-2 text-center s-th border-l s-bd font-semibold s-soft" style={{ minWidth: 120 }}>ПІДСУМОК</th>
                  </tr>
                </thead>
                <tbody>
                  {visibleEmployees.length === 0 && (<tr><td colSpan={dN + 2} className="px-4 py-12 text-center s-muted">{employees.length === 0 ? "Ще немає співробітників. Додайте першого вище." : "Нічого не знайдено."}</td></tr>)}
                  {visibleEmployees.map((emp) => {
                    const rec = attendance[emp.id] || {}; const c = computeEmp(emp);
                    return (
                      <tr key={emp.id} className="border-b s-bd s-rowhover">
                        <td className="sticky left-0 z-10 s-card px-4 py-2 border-r s-bd align-top">
                          <button onClick={() => setProfile(emp)} className="text-left group">
                            <div className="font-medium group-hover:text-indigo-500 flex items-center gap-1.5">{emp.name} <Pencil size={12} className="opacity-0 group-hover:opacity-100 s-muted" /></div>
                            <div className="f11 s-muted">{emp.start}–{emp.end}{emp.position ? ` · ${emp.position}` : ""}</div>
                          </button>
                        </td>
                        {dayList.map((d) => {
                          const wd = new Date(year, month, d).getDay(); const weekend = wd === 0 || wd === 6;
                          const manual = rec[d];
                          const lv = manual ? null : leaveOn(emp, isoDay(year, month, d));
                          const eff = manual || (lv ? { status: lv, auto: true } : null);
                          return (<td key={d} className={`px-0.5 py-1 text-center ${weekend ? "s-weekend" : ""}`}><DayCell rec={eff} onClick={() => setEditing({ empId: emp.id, day: d, emp })} /></td>);
                        })}
                        <td className="px-3 py-2 text-center s-th border-l s-bd align-top">
                          <div className="text-xs s-soft">відпр. <b>{fmt(c.worked)}г</b></div>
                          {c.excused > 0 && <div className="f11 s-muted">відп/лік: {c.excused}</div>}
                          <div className="mt-1 inline-flex items-center gap-1 px-2 py-0.5 rounded-md text-xs font-medium" style={{ background: c.balance >= 0 ? "rgba(16,185,129,.14)" : "rgba(239,68,68,.14)", color: c.balance >= 0 ? "#10b981" : "#ef4444" }}>{c.balance >= 0 ? "+" : ""}{fmt(c.balance)}г</div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
          <p className="text-xs s-muted px-1">{isManager ? "Ви відмічаєте лише свій відділ. " : ""}Відпустки та лікарняні вносяться у картці співробітника й заповнюються в табелі автоматично. Дані синхронізуються між пристроями.</p>
        </div>
      </main>

      {/* МОДАЛКИ */}
      {editing && (() => {
        const manual = attendance[editing.empId]?.[editing.day];
        const emp = employees.find((e) => e.id === editing.empId) || editing.emp;
        const lv = manual ? null : leaveOn(emp, isoDay(year, month, editing.day));
        return (<DayEditor info={editing} value={manual || (lv ? { status: lv } : undefined)} autoLeave={!!lv} onClose={() => setEditing(null)} onSave={(rec) => saveCell(editing.empId, editing.day, rec)} dateLabel={`${editing.day} ${MONTHS[month]} ${year}`} />);
      })()}
      {profile && <ProfileEditor emp={profile} onClose={() => setProfile(null)} onSave={saveProfile} onDelete={() => removeEmployee(profile.id)} />}
      {showChat && <ChatPanel chat={chat} departments={departments} currentDeptId={isManager ? session.deptId : deptId} onSend={sendMessage} onClose={() => setShowChat(false)} />}
      {showHistory && <HistoryModal log={log} onClose={() => setShowHistory(false)} />}
      {showReport && <ReportModal onClose={() => setShowReport(false)} deptName={deptName} month={MONTHS[month]} year={year} employees={employees} stats={stats} compute={computeEmp} onExport={exportExcel} />}
      {showAccess && <AccessPanel onClose={() => setShowAccess(false)} departments={departments} users={users} onSetPin={setRolePin} onClearPin={clearRolePin} />}

      {toast && <div className="fixed bottom-5 left-1/2 -translate-x-1/2 z-50 flex items-center gap-2 px-4 py-2.5 rounded-xl text-white text-sm shadow-lg" style={{ background: "#0f172a" }}><Check size={16} className="text-emerald-400" /> {toast}</div>}
    </div>
  );
}

/* ============================================================
   ЕКРАН ВХОДУ (завжди темний)
============================================================ */
function LoginScreen({ departments, users, onLogin, onCreateAdmin }) {
  const [sel, setSel] = useState(null);
  const [pin, setPin] = useState(""); const [pin2, setPin2] = useState("");
  const [err, setErr] = useState(""); const [busy, setBusy] = useState(false); const [showPin, setShowPin] = useState(false);
  const tiles = [
    { role: "admin", label: "Керуючий", icon: <KeyRound size={22} /> },
    { role: "hr", label: "Відділ кадрів", icon: <Users size={22} /> },
    ...departments.map((d) => ({ role: "dept", deptId: d.id, label: d.name, icon: <Building2 size={22} /> })),
  ];
  const keyOf = (t) => (t.role === "dept" ? t.deptId : t.role);
  const adminPinSet = !!users.admin?.pin;
  const isCreateAdmin = sel?.role === "admin" && !adminPinSet;
  const selPinSet = sel ? !!users[keyOf(sel)]?.pin : false;
  const locked = sel && sel.role !== "admin" && !selPinSet;
  const pick = (t) => { setSel(t); setPin(""); setPin2(""); setErr(""); };
  const submit = async () => {
    if (!sel || busy) return; setErr(""); setBusy(true);
    try {
      if (isCreateAdmin) { if (pin.length < 4) { setErr("PIN має містити щонайменше 4 цифри"); return; } if (pin !== pin2) { setErr("PIN-коди не співпадають"); return; } await onCreateAdmin(pin); return; }
      if (locked) return;
      const e = await onLogin(sel, pin); if (e) setErr(e);
    } finally { setBusy(false); }
  };
  const canEnter = sel && !locked && !busy && pin.length >= 4;
  return (
    <div className="min-h-screen flex items-center justify-center p-4" style={{ background: "radial-gradient(120% 100% at 15% 0%, #1e2a5e 0%, #0c1230 45%, #070b1f 100%)", fontFamily: "Inter, ui-sans-serif, system-ui, sans-serif" }}>
      <div className="w-full max-w-md rounded-3xl p-6" style={{ background: "rgba(255,255,255,.05)", border: "1px solid rgba(255,255,255,.1)", backdropFilter: "blur(16px)", boxShadow: "0 30px 80px -20px rgba(0,0,0,.6)" }}>
        <div className="flex items-center gap-3 mb-6">
          <div className="w-12 h-12 rounded-2xl bg-gradient-to-br from-indigo-500 to-violet-600 flex items-center justify-center text-white" style={{ boxShadow: "0 10px 30px -6px rgba(99,102,241,.7)" }}><CalendarDays size={24} /></div>
          <div><div className="text-2xl font-bold text-white tracking-tight">Табель <span className="text-indigo-300">PRO</span></div><div className="text-sm" style={{ color: "#94a3b8" }}>Система обліку робочого часу</div></div>
        </div>
        {!adminPinSet && (<div className="mb-4 text-xs rounded-xl px-3 py-2" style={{ color: "#c7d2fe", background: "rgba(99,102,241,.1)", border: "1px solid rgba(99,102,241,.2)" }}>Перший запуск: оберіть «Керуючий» і створіть PIN-код адміністратора.</div>)}
        <div className="grid grid-cols-3 gap-2.5 mb-4">
          {tiles.map((t) => {
            const active = sel && keyOf(sel) === keyOf(t);
            const noPin = t.role !== "admin" && !users[keyOf(t)]?.pin;
            return (<button key={keyOf(t)} onClick={() => pick(t)} className="relative rounded-2xl px-2 py-4 flex flex-col items-center gap-2 transition" style={{ background: active ? "rgba(99,102,241,.2)" : "rgba(255,255,255,.05)", border: active ? "1px solid #818cf8" : "1px solid rgba(255,255,255,.1)", boxShadow: active ? "0 0 0 3px rgba(129,140,248,.3)" : "none" }}>
              <span style={{ color: active ? "#c7d2fe" : "#cbd5e1" }}>{t.icon}</span>
              <span className="f12 font-medium text-center leading-tight" style={{ color: "#e2e8f0" }}>{t.label}</span>
              {noPin && <Lock size={11} className="absolute top-2 right-2" style={{ color: "#64748b" }} />}
            </button>);
          })}
        </div>
        {sel && (<div className="mb-3">
          {locked ? (<div className="text-sm rounded-xl px-3 py-2.5 flex items-center gap-2" style={{ color: "#fde68a", background: "rgba(245,158,11,.1)", border: "1px solid rgba(245,158,11,.2)" }}><Lock size={15} /> Доступ для «{sel.label}» ще не налаштовано керуючим.</div>) : (
            <div className="space-y-2">
              <PinInput value={pin} onChange={setPin} show={showPin} onToggle={() => setShowPin((s) => !s)} placeholder={isCreateAdmin ? "Новий PIN (мін. 4 цифри)" : "Введіть PIN-код"} onEnter={submit} autoFocus />
              {isCreateAdmin && <PinInput value={pin2} onChange={setPin2} show={showPin} placeholder="Повторіть PIN" onEnter={submit} />}
              {err && <div className="text-sm" style={{ color: "#fca5a5" }}>{err}</div>}
            </div>)}
        </div>)}
        <button onClick={submit} disabled={!canEnter} className="w-full py-3 rounded-2xl font-semibold flex items-center justify-center gap-2 transition text-white" style={{ background: canEnter ? "linear-gradient(90deg,#6366f1,#8b5cf6)" : "rgba(255,255,255,.1)", color: canEnter ? "#fff" : "#64748b", cursor: canEnter ? "pointer" : "not-allowed" }}>{isCreateAdmin ? "Створити PIN та увійти" : "Увійти"} →</button>
      </div>
    </div>
  );
}
function PinInput({ value, onChange, show, onToggle, placeholder, onEnter, autoFocus }) {
  return (<div className="relative">
    <input autoFocus={autoFocus} type={show ? "text" : "password"} inputMode="numeric" maxLength={6} value={value} onChange={(e) => onChange(e.target.value.replace(/\D/g, ""))} onKeyDown={(e) => e.key === "Enter" && onEnter && onEnter()} placeholder={placeholder} className="w-full px-4 py-3 rounded-2xl tracking-widest text-white" style={{ background: "rgba(255,255,255,.05)", border: "1px solid rgba(255,255,255,.15)", outline: "none" }} />
    {onToggle && <button onClick={onToggle} className="absolute right-3 top-1/2 -translate-y-1/2" style={{ color: "#94a3b8" }}>{show ? <EyeOff size={16} /> : <Eye size={16} />}</button>}
  </div>);
}

/* ============================================================
   ПАНЕЛЬ ДОСТУПІВ
============================================================ */
function AccessPanel({ onClose, departments, users, onSetPin, onClearPin }) {
  const rows = [
    { key: "admin", label: "Керуючий", icon: <KeyRound size={16} className="text-amber-500" />, fixed: true },
    { key: "hr", label: "Відділ кадрів", icon: <Users size={16} className="text-sky-500" /> },
    ...departments.map((d) => ({ key: d.id, label: d.name, icon: <Building2 size={16} className="text-indigo-500" /> })),
  ];
  return (<Overlay onClose={onClose}><div className="s-elev rounded-2xl shadow-2xl overflow-hidden flex flex-col" style={{ width: 480, maxWidth: "92vw", maxHeight: "80vh" }}>
    <div className="px-5 py-4 border-b s-bd flex items-center justify-between"><div className="flex items-center gap-2 font-semibold"><ShieldCheck size={18} className="text-indigo-500" /> Доступи та PIN-коди</div><button onClick={onClose} className="p-1.5 rounded-lg s-hover s-muted"><X size={18} /></button></div>
    <div className="p-5 overflow-y-auto space-y-2"><p className="text-xs s-muted mb-2">Призначте PIN кожному кабінету. Керівник входить лише у свій відділ і відмічає тільки його.</p>{rows.map((r) => <AccessRow key={r.key} row={r} hasPin={!!users[r.key]?.pin} onSet={(pin) => onSetPin(r.key, pin)} onClear={() => onClearPin(r.key)} />)}</div>
  </div></Overlay>);
}
function AccessRow({ row, hasPin, onSet, onClear }) {
  const [editing, setEditing] = useState(false); const [pin, setPin] = useState("");
  const save = () => { if (pin.length >= 4) { onSet(pin); setPin(""); setEditing(false); } };
  return (<div className="flex items-center gap-3 px-3 py-2.5 rounded-xl border s-bd s-th">
    <div className="flex items-center gap-2 flex-1 font-medium">{row.icon}{row.label}</div>
    {editing ? (<div className="flex items-center gap-2"><input autoFocus type="text" inputMode="numeric" maxLength={6} value={pin} onChange={(e) => setPin(e.target.value.replace(/\D/g, ""))} onKeyDown={(e) => e.key === "Enter" && save()} placeholder="PIN ≥4" className="w-24 px-2 py-1.5 text-sm rounded-lg s-input tracking-widest" /><button onClick={save} className="px-2.5 py-1.5 text-sm rounded-lg bg-indigo-600 text-white hover:bg-indigo-700">OK</button><button onClick={() => { setEditing(false); setPin(""); }} className="s-muted"><X size={16} /></button></div>) : (
      <div className="flex items-center gap-2"><span className="text-xs px-2 py-1 rounded-md" style={{ background: hasPin ? "rgba(16,185,129,.14)" : "var(--hover)", color: hasPin ? "#10b981" : "var(--muted)" }}>{hasPin ? "PIN встановлено" : "не задано"}</span><button onClick={() => setEditing(true)} className="text-xs px-2.5 py-1.5 rounded-lg border s-bd s-soft s-hover">{hasPin ? "Змінити" : "Задати"}</button>{hasPin && !row.fixed && <button onClick={onClear} className="p-1.5 rounded-lg s-muted hover:text-red-500" title="Скинути"><RotateCcw size={14} /></button>}</div>)}
  </div>);
}

/* ============================================================
   КОМПОНЕНТИ
============================================================ */
function StatCard({ icon, label, value, color }) {
  return (<div className="s-card s-shadow s-lift rounded-2xl border s-bd px-3 py-3 flex items-center gap-3"><div className="w-9 h-9 rounded-xl flex items-center justify-center" style={{ color, background: color + "22" }}>{icon}</div><div className="min-w-0"><div className="f11 s-muted truncate">{label}</div><div className="text-lg font-bold leading-tight tabular-nums">{value}</div></div></div>);
}

function DayCell({ rec, onClick }) {
  if (!rec) return <button onClick={onClick} className="w-full h-12 rounded-lg border border-dashed s-bd s-muted hover:border-indigo-300 hover:text-indigo-400 transition flex items-center justify-center">+</button>;
  const h = dayHours(rec); const st = rec.status ? STATUSES[rec.status] : null;
  return (<button onClick={onClick} className="relative w-full h-12 rounded-lg border transition flex flex-col items-center justify-center gap-0.5 hover:ring-2 hover:ring-indigo-200" style={{ background: st ? st.bg : "#eef2ff", borderColor: st ? st.dot + "55" : "#c7d2fe" }}>
    {rec.auto && <span className="absolute top-1 right-1 w-1.5 h-1.5 rounded-full" style={{ background: st ? st.dot : "#6366f1" }} title="З картки" />}
    {rec.in || rec.out ? (<div className="f10 leading-tight font-medium" style={{ color: st ? st.text : "#3730a3" }}>{rec.in || "—"}<br />{rec.out || "—"}</div>) : (<div className="f10 font-semibold leading-none" style={{ color: st ? st.text : "#3730a3" }}>{st ? st.label : ""}</div>)}
    {h > 0 && <div className="f9 font-bold" style={{ color: st ? st.text : "#3730a3" }}>{fmt(h)}г</div>}
  </button>);
}

function DayEditor({ info, value, autoLeave, onClose, onSave, dateLabel }) {
  const [tin, setTin] = useState(value?.in || "");
  const [tout, setTout] = useState(value?.out || "");
  const [status, setStatus] = useState(value?.status || "");
  const [note, setNote] = useState(value?.note || "");
  const applyShift = () => { setTin(info.emp?.start || "08:00"); setTout(info.emp?.end || "17:00"); if (!status) setStatus("present"); };
  const hours = dayHours({ in: tin, out: tout });
  return (<Overlay onClose={onClose}><div className="s-elev rounded-2xl shadow-2xl overflow-hidden" style={{ width: 380, maxWidth: "92vw" }}>
    <div className="px-5 py-4 border-b s-bd flex items-center justify-between"><div><div className="font-semibold">{info.emp?.name}</div><div className="text-xs s-muted">{dateLabel}</div></div><button onClick={onClose} className="p-1.5 rounded-lg s-hover s-muted"><X size={18} /></button></div>
    <div className="p-5 space-y-4">
      {autoLeave && <div className="text-xs rounded-lg px-3 py-2 s-th s-soft">Цей день заповнено автоматично з картки. Збереження створить ручний запис, «Очистити» поверне авто.</div>}
      <div className="grid grid-cols-2 gap-3">
        <label className="block"><span className="text-xs s-soft">Прихід</span><input type="time" value={tin} onChange={(e) => setTin(e.target.value)} className="mt-1 w-full px-3 py-2 rounded-lg s-input" /></label>
        <label className="block"><span className="text-xs s-soft">Вихід</span><input type="time" value={tout} onChange={(e) => setTout(e.target.value)} className="mt-1 w-full px-3 py-2 rounded-lg s-input" /></label>
      </div>
      <div className="flex flex-wrap gap-1.5">
        <button onClick={applyShift} className="px-2.5 py-1.5 rounded-lg text-xs font-medium border s-bd s-soft s-hover">Графік {info.emp?.start}–{info.emp?.end}</button>
        {SHIFTS.map((s) => (<button key={s.id} onClick={() => { setTin(s.start); setTout(s.end); if (!status) setStatus("present"); }} className="px-2.5 py-1.5 rounded-lg text-xs font-medium border s-bd s-soft s-hover">{s.label}</button>))}
      </div>
      {hours > 0 && <div className="text-sm s-soft">Відпрацьовано: <b>{fmt(hours)} год</b></div>}
      <div><span className="text-xs s-soft">Статус</span><div className="mt-1.5 grid grid-cols-3 gap-1.5">{Object.entries(STATUSES).map(([k, s]) => (<button key={k} onClick={() => setStatus(status === k ? "" : k)} className="px-2 py-1.5 rounded-lg text-xs font-medium border transition" style={{ background: status === k ? s.bg : "transparent", borderColor: status === k ? s.dot : "var(--border)", color: status === k ? s.text : "var(--soft)" }}>{s.label}</button>))}</div></div>
      <label className="block"><span className="text-xs s-soft">Примітка</span><input value={note} onChange={(e) => setNote(e.target.value)} placeholder="напр. деталі дня…" className="mt-1 w-full px-3 py-2 rounded-lg s-input text-sm" /></label>
    </div>
    <div className="px-5 py-4 s-th flex items-center gap-2"><button onClick={() => onSave({ in: tin, out: tout, status, note })} className="flex-1 py-2.5 rounded-lg bg-indigo-600 text-white font-medium hover:bg-indigo-700">Зберегти</button><button onClick={() => onSave(null)} className="px-4 py-2.5 rounded-lg border s-bd s-soft s-hover">Очистити</button></div>
  </div></Overlay>);
}

function ProfileEditor({ emp, onClose, onSave, onDelete }) {
  const [d, setD] = useState({ leaves: [], ...emp });
  const [lt, setLt] = useState("vacation"); const [lf, setLf] = useState(""); const [lto, setLto] = useState("");
  const upd = (k, v) => setD((p) => ({ ...p, [k]: v }));
  const setShift = (start, end) => setD((p) => ({ ...p, start, end }));
  const initMatch = SHIFTS.find((s) => s.start === (emp.start || "") && s.end === (emp.end || ""));
  const [shiftMode, setShiftMode] = useState(initMatch ? initMatch.id : "manual");
  const pickPreset = (s) => { setShiftMode(s.id); setShift(s.start, s.end); };
  const addLeave = () => { if (!lf || !lto) return; const from = lf <= lto ? lf : lto, to = lf <= lto ? lto : lf; setD((p) => ({ ...p, leaves: [...(p.leaves || []), { id: uid(), type: lt, from, to }] })); setLf(""); setLto(""); };
  const removeLeave = (id) => setD((p) => ({ ...p, leaves: (p.leaves || []).filter((x) => x.id !== id) }));
  return (<Overlay onClose={onClose}><div className="s-elev rounded-2xl shadow-2xl overflow-hidden flex flex-col" style={{ width: 420, maxWidth: "92vw", maxHeight: "88vh" }}>
    <div className="px-5 py-4 border-b s-bd flex items-center justify-between"><div className="font-semibold">Картка співробітника</div><button onClick={onClose} className="p-1.5 rounded-lg s-hover s-muted"><X size={18} /></button></div>
    <div className="p-5 space-y-3 overflow-y-auto">
      <Field label="ПІБ"><input value={d.name} onChange={(e) => upd("name", e.target.value)} className="mt-1 w-full px-3 py-2 rounded-lg s-input text-sm" /></Field>
      <Field label="Посада" icon={<Briefcase size={14} />}><input value={d.position || ""} onChange={(e) => upd("position", e.target.value)} placeholder="напр. Менеджер" className="mt-1 w-full px-3 py-2 rounded-lg s-input text-sm" /></Field>
      <Field label="Телефон" icon={<Phone size={14} />}><input value={d.phone || ""} onChange={(e) => upd("phone", e.target.value)} placeholder="+380…" className="mt-1 w-full px-3 py-2 rounded-lg s-input text-sm" /></Field>
      <Field label="Дата народження" icon={<Cake size={14} />}><input type="date" value={d.birthday || ""} onChange={(e) => upd("birthday", e.target.value)} className="mt-1 w-full px-3 py-2 rounded-lg s-input text-sm" /></Field>
      <div>
        <span className="text-xs s-soft flex items-center gap-1.5"><Clock size={14} /> Графік роботи</span>
        <div className="mt-1.5 flex flex-wrap gap-1.5">
          {SHIFTS.map((s) => (<button key={s.id} onClick={() => pickPreset(s)} className="px-3 py-1.5 rounded-lg text-xs font-medium border transition" style={{ background: shiftMode === s.id ? "#6366f1" : "transparent", color: shiftMode === s.id ? "#fff" : "var(--soft)", borderColor: shiftMode === s.id ? "#6366f1" : "var(--border)" }}>{s.label}</button>))}
          <button onClick={() => setShiftMode("manual")} className="px-3 py-1.5 rounded-lg text-xs font-medium border transition" style={{ background: shiftMode === "manual" ? "#6366f1" : "transparent", color: shiftMode === "manual" ? "#fff" : "var(--soft)", borderColor: shiftMode === "manual" ? "#6366f1" : "var(--border)" }}>Вручну</button>
        </div>
        {shiftMode === "manual" && (
          <div className="mt-2 grid grid-cols-2 gap-3">
            <Field label="Початок"><input type="time" value={d.start} onChange={(e) => setShift(e.target.value, d.end)} className="mt-1 w-full px-3 py-2 rounded-lg s-input text-sm" /></Field>
            <Field label="Кінець"><input type="time" value={d.end} onChange={(e) => setShift(d.start, e.target.value)} className="mt-1 w-full px-3 py-2 rounded-lg s-input text-sm" /></Field>
          </div>
        )}
        <div className="mt-1.5 f11 s-muted">Обрано: {d.start || "—"}–{d.end || "—"} · {(() => { let h = dayHours({ in: d.start, out: d.end }); return h > 0 ? fmt(h) + " год" : ""; })()}</div>
      </div>

      <div className="pt-2 border-t s-bd">
        <span className="text-xs s-soft flex items-center gap-1.5"><Plane size={14} /> Лікарняні та відпустки</span>
        <div className="mt-2 space-y-1.5">
          {(d.leaves || []).map((l) => (<div key={l.id} className="flex items-center gap-2 text-sm px-2 py-1.5 rounded-lg s-th"><span className="px-2 py-0.5 rounded-md text-xs font-medium" style={{ background: STATUSES[l.type].bg, color: STATUSES[l.type].text }}>{STATUSES[l.type].label}</span><span className="s-soft">{l.from} → {l.to}</span><button onClick={() => removeLeave(l.id)} className="ml-auto s-muted hover:text-red-500"><X size={14} /></button></div>))}
          {(!d.leaves || d.leaves.length === 0) && <div className="text-xs s-muted">Ще не додано. Внесені періоди з'являться у табелі автоматично.</div>}
        </div>
        <div className="mt-2 flex flex-wrap items-end gap-2">
          <select value={lt} onChange={(e) => setLt(e.target.value)} className="px-2 py-2 rounded-lg s-input text-sm"><option value="vacation">Відпустка</option><option value="sick">Лікарняний</option></select>
          <label className="text-xs s-muted">з<input type="date" value={lf} onChange={(e) => setLf(e.target.value)} className="block mt-0.5 px-2 py-2 rounded-lg s-input text-sm" /></label>
          <label className="text-xs s-muted">по<input type="date" value={lto} onChange={(e) => setLto(e.target.value)} className="block mt-0.5 px-2 py-2 rounded-lg s-input text-sm" /></label>
          <button onClick={addLeave} className="px-3 py-2 rounded-lg bg-indigo-600 text-white text-sm hover:bg-indigo-700">Додати</button>
        </div>
      </div>
    </div>
    <div className="px-5 py-4 s-th flex items-center gap-2"><button onClick={() => onSave(d)} className="flex-1 py-2.5 rounded-lg bg-indigo-600 text-white font-medium hover:bg-indigo-700">Зберегти</button><button onClick={onDelete} className="px-3 py-2.5 rounded-lg border border-red-300 text-red-500 hover:bg-red-50 flex items-center gap-1.5"><Trash2 size={15} /> Видалити</button></div>
  </div></Overlay>);
}
function Field({ label, icon, children }) { return (<label className="block"><span className="text-xs s-soft flex items-center gap-1.5">{icon}{label}</span>{children}</label>); }

function ChatPanel({ chat, departments, currentDeptId, onSend, onClose }) {
  const [text, setText] = useState(""); const [filter, setFilter] = useState("all"); const endRef = useRef(null);
  useEffect(() => { endRef.current?.scrollIntoView({ behavior: "smooth" }); }, [chat, filter]);
  const list = filter === "all" ? chat : chat.filter((m) => m.deptId === filter);
  const send = () => { onSend(text); setText(""); };
  return (<div className="fixed inset-0 z-40 flex justify-end"><div className="absolute inset-0" style={{ background: "var(--overlay)" }} onClick={onClose} />
    <div className="relative h-full flex flex-col shadow-2xl s-panel" style={{ width: 420, maxWidth: "100%" }}>
      <div className="px-5 py-4 border-b s-bd flex items-center justify-between"><div className="flex items-center gap-2 font-semibold"><MessageSquare size={18} className="text-violet-500" /> Чат відділів</div><button onClick={onClose} className="p-1.5 rounded-lg s-hover s-muted"><X size={18} /></button></div>
      <div className="px-4 py-2 border-b s-bd flex gap-1 overflow-x-auto"><button onClick={() => setFilter("all")} className={`px-3 py-1 rounded-full text-xs whitespace-nowrap ${filter === "all" ? "bg-violet-500 text-white" : "s-th s-soft"}`}>Усі</button>{departments.map((d) => <button key={d.id} onClick={() => setFilter(d.id)} className={`px-3 py-1 rounded-full text-xs whitespace-nowrap ${filter === d.id ? "bg-violet-500 text-white" : "s-th s-soft"}`}>{d.name}</button>)}</div>
      <div className="flex-1 overflow-y-auto p-4 space-y-3 s-th">
        {list.length === 0 && <div className="text-center s-muted text-sm mt-10">Повідомлень ще немає.<br />Напишіть перше нижче.</div>}
        {list.map((m) => { const mine = m.deptId === currentDeptId; return (<div key={m.id} className={`flex flex-col ${mine ? "items-end" : "items-start"}`}><div className="f11 s-muted mb-0.5 px-1">{m.fromLabel || m.deptName} · {new Date(m.ts).toLocaleString("uk-UA", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" })}</div><div className="px-3 py-2 rounded-2xl text-sm" style={{ maxWidth: "80%", background: mine ? "#7c3aed" : "var(--card)", color: mine ? "#fff" : "var(--text)", border: mine ? "none" : "1px solid var(--border)" }}>{m.text}</div></div>); })}
        <div ref={endRef} />
      </div>
      <div className="p-3 border-t s-bd flex items-center gap-2"><input value={text} onChange={(e) => setText(e.target.value)} onKeyDown={(e) => e.key === "Enter" && send()} placeholder="Повідомлення…" className="flex-1 px-3 py-2.5 rounded-xl s-input text-sm" /><button onClick={send} className="w-10 h-10 rounded-xl bg-violet-500 text-white flex items-center justify-center hover:bg-violet-600"><Send size={16} /></button></div>
    </div>
  </div>);
}

function HistoryModal({ log, onClose }) {
  return (<Overlay onClose={onClose}><div className="s-elev rounded-2xl shadow-2xl overflow-hidden flex flex-col" style={{ width: 460, maxWidth: "92vw", maxHeight: "70vh" }}>
    <div className="px-5 py-4 border-b s-bd flex items-center justify-between"><div className="flex items-center gap-2 font-semibold"><History size={18} className="s-soft" /> Історія дій</div><button onClick={onClose} className="p-1.5 rounded-lg s-hover s-muted"><X size={18} /></button></div>
    <div className="overflow-y-auto p-4 space-y-2">{log.length === 0 && <div className="text-center s-muted text-sm py-8">Поки що порожньо.</div>}{log.map((e) => (<div key={e.id} className="flex items-start gap-3 text-sm"><div className="w-1.5 h-1.5 rounded-full bg-indigo-400 mt-2" /><div className="flex-1"><div>{e.text}</div><div className="f11 s-muted">{new Date(e.ts).toLocaleString("uk-UA")}</div></div></div>))}</div>
  </div></Overlay>);
}

function ReportModal({ onClose, deptName, month, year, employees, stats, compute, onExport }) {
  return (<Overlay onClose={onClose}><div className="s-elev rounded-2xl shadow-2xl overflow-hidden flex flex-col" style={{ width: 560, maxWidth: "92vw", maxHeight: "80vh" }}>
    <div className="px-5 py-4 border-b s-bd flex items-center justify-between"><div><div className="font-semibold">Звіт · {deptName}</div><div className="text-xs s-muted">{month} {year}</div></div><button onClick={onClose} className="p-1.5 rounded-lg s-hover s-muted"><X size={18} /></button></div>
    <div className="p-5 overflow-y-auto">
      <div className="grid grid-cols-3 gap-3 mb-4"><MiniStat label="Відпрацьовано" value={fmt(stats.worked) + "г"} /><MiniStat label="Норма" value={stats.norm + "г"} /><MiniStat label="Баланс" value={(stats.balance >= 0 ? "+" : "") + fmt(stats.balance) + "г"} red={stats.balance < 0} /></div>
      <table className="w-full text-sm"><thead><tr className="text-left s-muted border-b s-bd"><th className="py-2">Працівник</th><th className="py-2 text-right">Відпрац.</th><th className="py-2 text-right">Відп/лік</th><th className="py-2 text-right">Баланс</th></tr></thead><tbody>
        {employees.map((e) => { const c = compute(e); return (<tr key={e.id} className="border-b s-bd"><td className="py-2">{e.name}<div className="f11 s-muted">{e.position}</div></td><td className="py-2 text-right">{fmt(c.worked)}г</td><td className="py-2 text-right">{c.excused}</td><td className="py-2 text-right font-medium" style={{ color: c.balance >= 0 ? "#10b981" : "#ef4444" }}>{c.balance >= 0 ? "+" : ""}{fmt(c.balance)}г</td></tr>); })}
        {employees.length === 0 && <tr><td colSpan={4} className="py-6 text-center s-muted">Немає даних</td></tr>}
      </tbody></table>
    </div>
    <div className="px-5 py-4 s-th"><button onClick={onExport} className="w-full py-2.5 rounded-lg bg-emerald-600 text-white font-medium hover:bg-emerald-700 flex items-center justify-center gap-2"><FileSpreadsheet size={16} /> Завантажити Excel</button></div>
  </div></Overlay>);
}
function MiniStat({ label, value, red }) { return (<div className="rounded-xl s-th px-3 py-3 text-center"><div className="f11 s-muted">{label}</div><div className="text-lg font-bold" style={{ color: red ? "#ef4444" : "var(--text)" }}>{value}</div></div>); }

function Overlay({ children, onClose }) { return (<div className="fixed inset-0 z-40 flex items-center justify-center p-4"><div className="absolute inset-0" style={{ background: "var(--overlay)" }} onClick={onClose} /><div className="relative">{children}</div></div>); }
