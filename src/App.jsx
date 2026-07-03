import { useState, useMemo, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

const SUPABASE_URL = "https://uvmdgmhcwykvgfkoocdw.supabase.co";
const SUPABASE_KEY = "sb_publishable_xf2imykHcuQV3eUxp_LQ2g_9Apt--9E";
const AUTH_URL = SUPABASE_URL + "/auth/v1";
const REST_URL = SUPABASE_URL + "/rest/v1";

const authFetch = async (path, body, token) => {
  const res = await fetch(AUTH_URL + path, {
    method: "POST",
    headers: { "apikey": SUPABASE_KEY, "Content-Type": "application/json", ...(token ? { "Authorization": "Bearer " + token } : {}) },
    body: JSON.stringify(body),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error_description || data.msg || "Error de autenticación");
  return data;
};

const signUp = (email, pass) => authFetch("/signup", { email, password: pass });
const signIn = (email, pass) => authFetch("/token?grant_type=password", { email, password: pass });
const signOut = (token) => authFetch("/logout", {}, token);
const resetPass = (email) => authFetch("/recover", { email });

const db = async (path, opts, token) => {
  const options = opts || {};
  const res = await fetch(REST_URL + "/" + path, {
    headers: {
      "apikey": SUPABASE_KEY,
      "Authorization": "Bearer " + token,
      "Content-Type": "application/json",
      "Prefer": options.prefer !== undefined ? options.prefer : "return=representation",
    },
    method: options.method || "GET",
    body: options.body,
  });
  if (res.status === 204) return [];
  const text = await res.text();
  if (!res.ok) throw new Error(text);
  return text ? JSON.parse(text) : [];
};

const DB = {
  getAll: (t, token) => db(t + "?order=id.asc", { method: "GET" }, token),
  insert: (t, data, token) => db(t, { method: "POST", body: JSON.stringify(data) }, token),
  update: (t, id, data, token) => db(t + "?id=eq." + id, { method: "PATCH", body: JSON.stringify(data) }, token),
  remove: (t, id, token) => db(t + "?id=eq." + id, { method: "DELETE", prefer: "" }, token),
};

const CATEGORIES = {
  income: ["Salario","Freelance","Ventas","Inversiones","Arriendo cobrado","Bono","Regalo","Otro ingreso"],
  expense: ["Comida","Transporte","Arriendo","Servicios","Salud","Entretenimiento","Ropa","Educación","Suscripciones","Gym","Farmacia","Otro gasto"],
  cost: ["Materiales","Proveedores","Nómina","Marketing","Equipo","Logística","Impuestos","Otro costo"],
};
const TYPE_COLORS = { income: "#10b981", expense: "#ef4444", cost: "#f59e0b" };
const TYPE_LABELS = { income: "Ingreso", expense: "Gasto", cost: "Costo" };
const TYPE_ICONS = { income: "💰", expense: "💸", cost: "🏭" };
const MONTHS_ES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

const fmt = (n, currency) => new Intl.NumberFormat("es-CO", { style: "currency", currency: currency || "COP", minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(n || 0);
const fmtDate = (d) => { if (!d) return "—"; return new Date(d + "T12:00:00").toLocaleDateString("es-CO", { day: "2-digit", month: "short", year: "numeric" }); };
const today = () => new Date().toISOString().slice(0, 10);
const thisMonth = () => new Date().toISOString().slice(0, 7);

function Modal({ children, onClose, title }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", zIndex: 200, display: "flex", alignItems: "flex-end" }} onClick={onClose}>
      <div onClick={e => e.stopPropagation()} style={{ background: "#fff", borderRadius: "22px 22px 0 0", width: "100%", maxHeight: "92vh", overflowY: "auto", paddingBottom: 36 }}>
        <div style={{ position: "sticky", top: 0, background: "#fff", display: "flex", justifyContent: "space-between", alignItems: "center", padding: "18px 20px 14px", borderBottom: "1px solid #f0f0f0", zIndex: 1 }}>
          <span style={{ fontWeight: 800, fontSize: 17 }}>{title}</span>
          <button onClick={onClose} style={{ background: "#f3f4f6", border: "none", borderRadius: "50%", width: 32, height: 32, fontSize: 18, cursor: "pointer" }}>×</button>
        </div>
        <div style={{ padding: "16px 20px" }}>{children}</div>
      </div>
    </div>
  );
}

function Field({ label, children }) {
  return (
    <div style={{ marginBottom: 14 }}>
      <label style={{ display: "block", fontSize: 13, color: "#6b7280", marginBottom: 4, fontWeight: 700 }}>{label}</label>
      {children}
    </div>
  );
}

const inputSx = { width: "100%", padding: "11px 14px", borderRadius: 12, border: "1.5px solid #e5e7eb", fontSize: 15, outline: "none", boxSizing: "border-box", background: "#fafafa", fontFamily: "inherit" };
const Inp = (p) => <input style={inputSx} {...p} />;
const Sel = ({ children, ...p }) => <select style={{ ...inputSx, appearance: "none" }} {...p}>{children}</select>;
const Textarea = (p) => <textarea style={{ ...inputSx, resize: "none", minHeight: 64 }} {...p} />;

function Btn({ children, onClick, variant, small, full, disabled }) {
  const v = variant || "primary";
  const bg = { primary: "#4f46e5", success: "#10b981", danger: "#ef4444", ghost: "#f3f4f6", warning: "#f59e0b", dark: "#1f2937" };
  const cl = { primary: "#fff", success: "#fff", danger: "#fff", ghost: "#374151", warning: "#fff", dark: "#fff" };
  return (
    <button disabled={disabled} onClick={onClick}
      style={{ background: disabled ? "#d1d5db" : bg[v], color: disabled ? "#9ca3af" : cl[v], border: "none", borderRadius: 12, padding: small ? "7px 14px" : "13px 18px", fontSize: small ? 13 : 15, fontWeight: 700, cursor: disabled ? "not-allowed" : "pointer", width: full ? "100%" : "auto", marginBottom: full ? 10 : 0 }}>
      {children}
    </button>
  );
}

function Badge({ type }) {
  const bg = { income: "#d1fae5", expense: "#fee2e2", cost: "#fef3c7" };
  const cl = { income: "#065f46", expense: "#991b1b", cost: "#92400e" };
  return <span style={{ background: bg[type], color: cl[type], borderRadius: 8, padding: "2px 9px", fontSize: 11, fontWeight: 800 }}>{TYPE_LABELS[type]}</span>;
}

function ProgressBar({ value, max, color, height }) {
  const c = color || "#4f46e5";
  const h = height || 10;
  const pct = Math.min((value / max) * 100, 100);
  const col = pct >= 90 ? "#ef4444" : pct >= 70 ? "#f59e0b" : c;
  return (
    <div style={{ background: "#f3f4f6", borderRadius: 99, height: h, overflow: "hidden" }}>
      <div style={{ background: col, height: "100%", width: pct + "%", borderRadius: 99, transition: "width 0.5s" }} />
    </div>
  );
}

function Toast({ msg, onDone }) {
  useEffect(() => { const t = setTimeout(onDone, 2800); return () => clearTimeout(t); }, []);
  return <div style={{ position: "fixed", top: 60, left: "50%", transform: "translateX(-50%)", background: "#1f2937", color: "#fff", padding: "12px 24px", borderRadius: 14, fontSize: 14, fontWeight: 600, zIndex: 999, boxShadow: "0 4px 20px rgba(0,0,0,0.3)", whiteSpace: "nowrap" }}>{msg}</div>;
}

function Spinner() {
  return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", minHeight: "100vh", background: "#f5f5f7", gap: 16 }}>
      <div style={{ width: 48, height: 48, border: "4px solid #e5e7eb", borderTop: "4px solid #4f46e5", borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />
      <style>{"@keyframes spin{to{transform:rotate(360deg)}}"}</style>
    </div>
  );
}

function GoalCard({ g, profile, onDelete, onAdd, fmtFn, fmtDate }) {
  const [inputAmt, setInputAmt] = useState("");
  const pct = Math.min((g.saved / g.target) * 100, 100);
  const done = pct >= 100;
  const handle = () => { const n = parseFloat(inputAmt); if (!n || n <= 0) return; onAdd(g.id, n); setInputAmt(""); };
  return (
    <div style={{ background: "#f9fafb", borderRadius: 14, padding: 14, marginBottom: 12 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
        <div>
          <div style={{ fontWeight: 800, fontSize: 15 }}>{g.name} {done ? "🎉" : ""}</div>
          {g.deadline && <div style={{ fontSize: 11, color: "#9ca3af" }}>Fecha límite: {fmtDate(g.deadline)}</div>}
        </div>
        <button onClick={() => onDelete(g.id)} style={{ background: "none", border: "none", fontSize: 14, cursor: "pointer", opacity: 0.5 }}>🗑</button>
      </div>
      <ProgressBar value={g.saved} max={g.target} color={g.color} height={12} />
      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 6, fontSize: 12 }}>
        <span style={{ fontWeight: 700, color: g.color }}>{fmtFn(g.saved, profile.currency)} ahorrados</span>
        <span style={{ color: "#9ca3af" }}>de {fmtFn(g.target, profile.currency)} ({pct.toFixed(0)}%)</span>
      </div>
      {!done && (
        <div style={{ marginTop: 12 }}>
          <div style={{ fontSize: 12, color: "#6b7280", fontWeight: 700, marginBottom: 6 }}>Agregar ahorro:</div>
          <div style={{ display: "flex", gap: 8 }}>
            <input type="number" placeholder="Monto a agregar..." value={inputAmt} onChange={e => setInputAmt(e.target.value)} onKeyDown={e => e.key === "Enter" && handle()}
              style={{ flex: 1, padding: "10px 14px", borderRadius: 12, border: "1.5px solid " + g.color, fontSize: 14, outline: "none", background: "#fff", fontFamily: "inherit" }} />
            <button onClick={handle} disabled={!inputAmt || parseFloat(inputAmt) <= 0}
              style={{ background: inputAmt && parseFloat(inputAmt) > 0 ? g.color : "#d1d5db", color: "#fff", border: "none", borderRadius: 12, padding: "10px 18px", fontSize: 18, fontWeight: 800, cursor: inputAmt && parseFloat(inputAmt) > 0 ? "pointer" : "not-allowed" }}>+</button>
          </div>
        </div>
      )}
    </div>
  );
}

function LoginScreen({ onLogin }) {
  const [mode, setMode] = useState("login");
  const [email, setEmail] = useState("");
  const [pass, setPass] = useState("");
  const [pass2, setPass2] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [msg, setMsg] = useState("");

  const handle = async () => {
    setError(""); setMsg("");
    if (!email) return setError("Ingresa tu correo.");
    if (mode !== "reset" && !pass) return setError("Ingresa tu contraseña.");
    if (mode === "register" && pass !== pass2) return setError("Las contraseñas no coinciden.");
    if (mode !== "reset" && pass.length < 6) return setError("La contraseña debe tener al menos 6 caracteres.");
    setLoading(true);
    try {
      if (mode === "login") { const data = await signIn(email, pass); onLogin(data); }
      else if (mode === "register") { const data = await signUp(email, pass); if (data.access_token) { onLogin(data); } else setMsg("✅ Revisa tu correo para confirmar tu cuenta."); }
      else { await resetPass(email); setMsg("✅ Te enviamos un correo para restablecer tu contraseña."); }
    } catch (e) { setError(e.message); }
    setLoading(false);
  };

  return (
    <div style={{ minHeight: "100vh", background: "linear-gradient(145deg,#4f46e5 0%,#7c3aed 60%,#a855f7 100%)", display: "flex", alignItems: "flex-end", justifyContent: "center", fontFamily: "-apple-system,BlinkMacSystemFont,'SF Pro Display',sans-serif" }}>
      <div style={{ background: "#fff", borderRadius: "28px 28px 0 0", width: "100%", maxWidth: 430, padding: "32px 28px 48px" }}>
        <div style={{ textAlign: "center", marginBottom: 28 }}>
          <div style={{ fontSize: 44, marginBottom: 8 }}>💰</div>
          <div style={{ fontSize: 24, fontWeight: 900, color: "#1f2937" }}>Mis Finanzas</div>
          <div style={{ fontSize: 14, color: "#9ca3af", marginTop: 4 }}>{mode === "login" ? "Inicia sesión" : mode === "register" ? "Crea tu cuenta" : "Recuperar contraseña"}</div>
        </div>
        {error && <div style={{ background: "#fee2e2", color: "#991b1b", borderRadius: 12, padding: "12px 16px", marginBottom: 14, fontSize: 14, fontWeight: 600 }}>⚠️ {error}</div>}
        {msg && <div style={{ background: "#d1fae5", color: "#065f46", borderRadius: 12, padding: "12px 16px", marginBottom: 14, fontSize: 14, fontWeight: 600 }}>{msg}</div>}
        <Field label="Correo electrónico"><Inp type="email" placeholder="tucorreo@email.com" value={email} onChange={e => setEmail(e.target.value)} /></Field>
        {mode !== "reset" && <Field label="Contraseña"><Inp type="password" placeholder="Mínimo 6 caracteres" value={pass} onChange={e => setPass(e.target.value)} onKeyDown={e => e.key === "Enter" && handle()} /></Field>}
        {mode === "register" && <Field label="Confirmar contraseña"><Inp type="password" placeholder="Repite tu contraseña" value={pass2} onChange={e => setPass2(e.target.value)} /></Field>}
        <Btn onClick={handle} variant="primary" full disabled={loading}>{loading ? "⏳ Procesando..." : mode === "login" ? "Iniciar sesión" : mode === "register" ? "Crear cuenta" : "Enviar correo"}</Btn>
        <div style={{ display: "flex", flexDirection: "column", gap: 10, marginTop: 8, alignItems: "center" }}>
          {mode === "login" && <>
            <button onClick={() => { setMode("register"); setError(""); setMsg(""); }} style={{ background: "none", border: "none", color: "#4f46e5", fontWeight: 700, fontSize: 14, cursor: "pointer" }}>¿No tienes cuenta? Regístrate</button>
            <button onClick={() => { setMode("reset"); setError(""); setMsg(""); }} style={{ background: "none", border: "none", color: "#9ca3af", fontWeight: 600, fontSize: 13, cursor: "pointer" }}>¿Olvidaste tu contraseña?</button>
          </>}
          {mode !== "login" && <button onClick={() => { setMode("login"); setError(""); setMsg(""); }} style={{ background: "none", border: "none", color: "#4f46e5", fontWeight: 700, fontSize: 14, cursor: "pointer" }}>← Volver a iniciar sesión</button>}
        </div>
      </div>
    </div>
  );
}

export default function App() {
  const [session, setSession] = useState(null);
  const [authReady, setAuthReady] = useState(false);
  const [loading, setLoading] = useState(true);
  const [transactions, setTransactions] = useState([]);
  const [contacts, setContacts] = useState([]);
  const [budgets, setBudgets] = useState([]);
  const [goals, setGoals] = useState([]);
  const [debts, setDebts] = useState([]);
  const [profile, setProfile] = useState({ name: "", currency: "COP", monthly_saving_goal: 0 });
  const [tab, setTab] = useState("dashboard");
  const [filterType, setFilterType] = useState("all");
  const [filterMonth, setFilterMonth] = useState(thisMonth());
  const [search, setSearch] = useState("");
  const [showAdd, setShowAdd] = useState(false);
  const [showDetail, setShowDetail] = useState(null);
  const [showAddBudget, setShowAddBudget] = useState(false);
  const [showAddGoal, setShowAddGoal] = useState(false);
  const [showAddDebt, setShowAddDebt] = useState(false);
  const [showAddContact, setShowAddContact] = useState(false);
  const [showProfile, setShowProfile] = useState(false);
  const [showContactDetail, setShowContactDetail] = useState(null);
  const [toast, setToast] = useState(null);
  const [form, setForm] = useState({ type: "income", amount: "", category: "Salario", description: "", date: today(), contact: "", note: "" });
  const [budgetForm, setBudgetForm] = useState({ category: "Comida", limit: "" });
  const [goalForm, setGoalForm] = useState({ name: "", target: "", saved: "0", deadline: "", color: "#6366f1" });
  const [debtForm, setDebtForm] = useState({ type: "owed", person: "", amount: "", reason: "", date: today() });
  const [contactForm, setContactForm] = useState({ name: "", phone: "", notes: "" });
  const [profileForm, setProfileForm] = useState(profile);

  const token = session && session.access_token;
  const userId = session && session.user && session.user.id;
  const notify = (msg) => setToast(msg);

  useEffect(() => {
    const stored = localStorage.getItem("fin_session");
    if (stored) { try { setSession(JSON.parse(stored)); } catch (e) {} }
    setAuthReady(true);
  }, []);

  useEffect(() => {
    if (!token || !userId) { setLoading(false); return; }
    loadAll();
  }, [token, userId]);

  const loadAll = async () => {
    setLoading(true);
    try {
      const [tx, ct, bg, gl, db_, pf] = await Promise.all([
        DB.getAll("transactions", token), DB.getAll("contacts", token), DB.getAll("budgets", token),
        DB.getAll("goals", token), DB.getAll("debts", token), DB.getAll("profile", token),
      ]);
      setTransactions(tx); setContacts(ct); setBudgets(bg); setGoals(gl); setDebts(db_);
      if (pf.length > 0) { setProfile(pf[0]); setProfileForm(pf[0]); }
    } catch (e) { notify("⚠️ Error: " + e.message); }
    setLoading(false);
  };

  const handleLogin = (data) => { localStorage.setItem("fin_session", JSON.stringify(data)); setSession(data); };
  const handleLogout = async () => {
    try { await signOut(token); } catch (e) {}
    localStorage.removeItem("fin_session");
    setSession(null); setTransactions([]); setContacts([]); setBudgets([]); setGoals([]); setDebts([]);
    setProfile({ name: "", currency: "COP", monthly_saving_goal: 0 });
  };

  const monthTx = useMemo(() => transactions.filter(t => t.date.startsWith(filterMonth)), [transactions, filterMonth]);
  const totals = useMemo(() => ({
    income: monthTx.filter(t => t.type === "income").reduce((s, t) => s + t.amount, 0),
    expense: monthTx.filter(t => t.type === "expense").reduce((s, t) => s + t.amount, 0),
    cost: monthTx.filter(t => t.type === "cost").reduce((s, t) => s + t.amount, 0),
  }), [monthTx]);
  const balance = totals.income - totals.expense - totals.cost;
  const savingPct = totals.income > 0 ? (balance / totals.income) * 100 : 0;

  const filtered = useMemo(() => monthTx
    .filter(t => filterType === "all" || t.type === filterType)
    .filter(t => !search || (t.description && t.description.toLowerCase().includes(search.toLowerCase())) || (t.category && t.category.toLowerCase().includes(search.toLowerCase())))
    .sort((a, b) => new Date(b.date) - new Date(a.date)),
    [monthTx, filterType, search]);

  const byCategory = useMemo(() => {
    const map = {};
    monthTx.forEach(t => { if (!map[t.category]) map[t.category] = { total: 0, type: t.type }; map[t.category].total += t.amount; });
    return Object.entries(map).sort((a, b) => b[1].total - a[1].total);
  }, [monthTx]);

  const budgetUsage = (cat) => monthTx.filter(t => t.category === cat && t.type !== "income").reduce((s, t) => s + t.amount, 0);

  const trend = useMemo(() => {
    const months = [];
    for (let i = 5; i >= 0; i--) {
      const d = new Date(); d.setMonth(d.getMonth() - i);
      const key = d.toISOString().slice(0, 7);
      const txs = transactions.filter(t => t.date.startsWith(key));
      months.push({ label: MONTHS_ES[d.getMonth()].slice(0, 3), income: txs.filter(t => t.type === "income").reduce((s, t) => s + t.amount, 0), expense: txs.filter(t => t.type !== "income").reduce((s, t) => s + t.amount, 0) });
    }
    return months;
  }, [transactions]);
  const maxTrend = Math.max(...trend.flatMap(m => [m.income, m.expense]), 1);

  const contactTotal = (name) => transactions.filter(t => t.contact === name && t.type === "income").reduce((s, t) => s + t.amount, 0);
  const debtTotals = useMemo(() => ({
    owed: debts.filter(d => d.type === "owed" && !d.paid).reduce((s, d) => s + d.amount, 0),
    owing: debts.filter(d => d.type === "owing" && !d.paid).reduce((s, d) => s + d.amount, 0),
  }), [debts]);

  const availMonths = useMemo(() => {
    const months = new Set(transactions.map(t => t.date.slice(0, 7)));
    months.add(thisMonth());
    return [...months].sort((a, b) => b.localeCompare(a)).slice(0, 12);
  }, [transactions]);

  const addTx = async () => {
    if (!form.amount || !form.description) return;
    const row = { type: form.type, amount: parseFloat(form.amount), category: form.category, description: form.description, date: form.date, contact: form.contact || "", note: form.note || "", user_id: userId };
    try { const [saved] = await DB.insert("transactions", row, token); setTransactions(p => [...p, saved]); setForm({ type: "income", amount: "", category: "Salario", description: "", date: today(), contact: "", note: "" }); setShowAdd(false); notify("✅ Guardado en la nube"); } catch (e) { notify("⚠️ Error: " + e.message); }
  };
  const deleteTx = async (id) => { try { await DB.remove("transactions", id, token); setTransactions(p => p.filter(t => t.id !== id)); setShowDetail(null); notify("🗑 Eliminado"); } catch (e) { notify("⚠️ Error: " + e.message); } };
  const addBudget = async () => {
    if (!budgetForm.limit) return;
    try { const [saved] = await DB.insert("budgets", { category: budgetForm.category, limit: parseFloat(budgetForm.limit), user_id: userId }, token); setBudgets(p => [...p, saved]); setBudgetForm({ category: "Comida", limit: "" }); setShowAddBudget(false); notify("✅ Presupuesto guardado"); } catch (e) { notify("⚠️ Error: " + e.message); }
  };
  const deleteBudget = async (id) => { await DB.remove("budgets", id, token); setBudgets(p => p.filter(b => b.id !== id)); };
  const addGoal = async () => {
    if (!goalForm.name || !goalForm.target) return;
    try { const [saved] = await DB.insert("goals", { name: goalForm.name, target: parseFloat(goalForm.target), saved: parseFloat(goalForm.saved || 0), deadline: goalForm.deadline || null, color: goalForm.color, user_id: userId }, token); setGoals(p => [...p, saved]); setGoalForm({ name: "", target: "", saved: "0", deadline: "", color: "#6366f1" }); setShowAddGoal(false); notify("🎯 Meta guardada"); } catch (e) { notify("⚠️ Error: " + e.message); }
  };
  const deleteGoal = async (id) => { await DB.remove("goals", id, token); setGoals(p => p.filter(g => g.id !== id)); };
  const addToGoal = async (id, amt) => {
    const g = goals.find(g => g.id === id); if (!g) return;
    const newSaved = Math.min(g.saved + amt, g.target);
    try { await DB.update("goals", id, { saved: newSaved }, token); setGoals(p => p.map(g => g.id === id ? { ...g, saved: newSaved } : g)); notify("💰 Ahorro agregado"); } catch (e) { notify("⚠️ Error: " + e.message); }
  };
  const addDebt = async () => {
    if (!debtForm.person || !debtForm.amount) return;
    try { const [saved] = await DB.insert("debts", { type: debtForm.type, person: debtForm.person, amount: parseFloat(debtForm.amount), reason: debtForm.reason || "", date: debtForm.date, paid: false, user_id: userId }, token); setDebts(p => [...p, saved]); setDebtForm({ type: "owed", person: "", amount: "", reason: "", date: today() }); setShowAddDebt(false); notify("✅ Deuda registrada"); } catch (e) { notify("⚠️ Error: " + e.message); }
  };
  const toggleDebt = async (id) => { const d = debts.find(d => d.id === id); await DB.update("debts", id, { paid: !d.paid }, token); setDebts(p => p.map(d => d.id === id ? { ...d, paid: !d.paid } : d)); notify("✅ Actualizado"); };
  const deleteDebt = async (id) => { await DB.remove("debts", id, token); setDebts(p => p.filter(d => d.id !== id)); };
  const addContact = async () => {
    if (!contactForm.name) return;
    try { const [saved] = await DB.insert("contacts", { name: contactForm.name, phone: contactForm.phone || "", notes: contactForm.notes || "", user_id: userId }, token); setContacts(p => [...p, saved]); setContactForm({ name: "", phone: "", notes: "" }); setShowAddContact(false); notify("👤 Contacto guardado"); } catch (e) { notify("⚠️ Error: " + e.message); }
  };
  const saveProfile = async () => {
    try {
      const data = { name: profileForm.name, currency: profileForm.currency, monthly_saving_goal: parseFloat(profileForm.monthly_saving_goal) || 0, user_id: userId };
      if (profile.id) { await DB.update("profile", profile.id, data, token); setProfile({ ...data, id: profile.id }); }
      else { const [saved] = await DB.insert("profile", data, token); setProfile(saved); }
      setShowProfile(false); notify("✅ Perfil actualizado");
    } catch (e) { notify("⚠️ Error: " + e.message); }
  };

  const exportExcel = useCallback(() => {
    const wb = XLSX.utils.book_new();
    const ws1 = XLSX.utils.aoa_to_sheet([["Fecha","Tipo","Categoría","Descripción","Monto","Persona","Nota"], ...transactions.sort((a,b)=>b.date.localeCompare(a.date)).map(t=>[fmtDate(t.date),TYPE_LABELS[t.type],t.category,t.description,t.type==="income"?t.amount:-t.amount,t.contact||"",t.note||""])]);
    XLSX.utils.book_append_sheet(wb, ws1, "Movimientos");
    XLSX.writeFile(wb, "MisFinanzas_" + thisMonth() + ".xlsx");
    notify("📊 Excel exportado");
  }, [transactions]);

  const currency = (profile && profile.currency) || "COP";
  const mo = filterMonth.split("-");
  const monthLabel = MONTHS_ES[parseInt(mo[1]) - 1] + " " + mo[0];

  if (!authReady) return <Spinner />;
  if (!session) return <LoginScreen onLogin={handleLogin} />;
  if (loading) return <Spinner />;

  return (
    <div style={{ fontFamily: "-apple-system,BlinkMacSystemFont,'SF Pro Display',sans-serif", background: "#f5f5f7", minHeight: "100vh", maxWidth: "100%", margin: "0 auto", position: "relative", paddingBottom: 88 }}>
      {toast && <Toast msg={toast} onDone={() => setToast(null)} />}

      <div style={{ background: "linear-gradient(145deg,#4f46e5 0%,#7c3aed 60%,#a855f7 100%)", padding: "52px 20px 28px", color: "#fff", position: "relative" }}>
        <button onClick={() => { setProfileForm(profile); setShowProfile(true); }} style={{ position: "absolute", top: 16, right: 16, background: "rgba(255,255,255,0.2)", border: "none", borderRadius: "50%", width: 36, height: 36, fontSize: 18, cursor: "pointer" }}>⚙️</button>
        <div style={{ fontSize: 13, opacity: 0.75 }}>Hola, {profile.name || "👋"}</div>
        <div style={{ fontSize: 12, opacity: 0.6, marginBottom: 8 }}>{monthLabel}</div>
        <div style={{ fontSize: 38, fontWeight: 900, letterSpacing: -1 }}>{fmt(balance, currency)}</div>
        <div style={{ fontSize: 12, opacity: 0.7, marginTop: 4 }}>{balance >= 0 ? "✅ Ahorro del " + savingPct.toFixed(1) + "% este mes" : "⚠️ Gastas más de lo que ingresas"}</div>
        <div style={{ fontSize: 11, opacity: 0.5, marginTop: 8 }}>☁️ {session.user.email}</div>
      </div>

      <div style={{ background: "#fff", padding: "10px 16px", borderBottom: "1px solid #f0f0f0", display: "flex", gap: 8, overflowX: "auto", scrollbarWidth: "none" }}>
        {availMonths.map(m => {
          const parts = m.split("-");
          return <button key={m} onClick={() => setFilterMonth(m)} style={{ background: filterMonth === m ? "#4f46e5" : "#f3f4f6", color: filterMonth === m ? "#fff" : "#374151", border: "none", borderRadius: 20, padding: "6px 14px", fontSize: 13, fontWeight: 700, whiteSpace: "nowrap", cursor: "pointer", flexShrink: 0 }}>{MONTHS_ES[parseInt(parts[1]) - 1].slice(0, 3)} {parts[0]}</button>;
        })}
      </div>

      <div style={{ display: "flex", gap: 10, padding: "14px 16px 0" }}>
        {["income","expense","cost"].map(type => (
          <div key={type} style={{ background: "#fff", borderRadius: 16, padding: "12px 14px", flex: 1, boxShadow: "0 1px 6px rgba(0,0,0,0.06)" }}>
            <div style={{ fontSize: 18 }}>{TYPE_ICONS[type]}</div>
            <div style={{ fontSize: 10, color: "#9ca3af", fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.4, marginTop: 2 }}>{TYPE_LABELS[type]}</div>
            <div style={{ fontSize: 14, fontWeight: 800, color: TYPE_COLORS[type], marginTop: 2 }}>{fmt(totals[type], currency)}</div>
          </div>
        ))}
      </div>

      <div style={{ padding: "14px 16px 0" }}>
        {tab === "dashboard" && (
          <>
            {budgets.filter(b => { const u = budgetUsage(b.category); return u / b.limit >= 0.7; }).map(b => {
              const u = budgetUsage(b.category); const pct = (u / b.limit * 100).toFixed(0); const over = u >= b.limit;
              return <div key={b.id} style={{ background: over ? "#fee2e2" : "#fef3c7", borderRadius: 14, padding: "12px 16px", marginBottom: 10, display: "flex", alignItems: "center", gap: 10 }}><span style={{ fontSize: 22 }}>{over ? "🚨" : "⚠️"}</span><div><div style={{ fontWeight: 700, fontSize: 14, color: over ? "#991b1b" : "#92400e" }}>{over ? "Presupuesto de " + b.category + " superado" : b.category + " al " + pct + "% del límite"}</div><div style={{ fontSize: 12, color: "#6b7280" }}>{fmt(u, currency)} / {fmt(b.limit, currency)}</div></div></div>;
            })}
            {(debtTotals.owed > 0 || debtTotals.owing > 0) && (
              <div style={{ background: "#fff", borderRadius: 16, padding: 16, marginBottom: 14, boxShadow: "0 1px 6px rgba(0,0,0,0.06)" }}>
                <div style={{ fontWeight: 800, marginBottom: 10 }}>💼 Deudas pendientes</div>
                <div style={{ display: "flex", gap: 10 }}>
                  <div style={{ flex: 1, background: "#d1fae5", borderRadius: 12, padding: 12 }}><div style={{ fontSize: 11, color: "#065f46", fontWeight: 700 }}>ME DEBEN</div><div style={{ fontWeight: 800, color: "#059669", fontSize: 18 }}>{fmt(debtTotals.owed, currency)}</div></div>
                  <div style={{ flex: 1, background: "#fee2e2", borderRadius: 12, padding: 12 }}><div style={{ fontSize: 11, color: "#991b1b", fontWeight: 700 }}>DEBO</div><div style={{ fontWeight: 800, color: "#dc2626", fontSize: 18 }}>{fmt(debtTotals.owing, currency)}</div></div>
                </div>
              </div>
            )}
            <div style={{ background: "#fff", borderRadius: 16, overflow: "hidden", boxShadow: "0 1px 6px rgba(0,0,0,0.06)", marginBottom: 14 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "14px 16px 10px" }}>
                <span style={{ fontWeight: 800 }}>Últimos movimientos</span>
                <button onClick={() => setTab("transactions")} style={{ background: "none", border: "none", color: "#4f46e5", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>Ver todos</button>
              </div>
              {transactions.length === 0 && <div style={{ padding: "16px 16px 20px", color: "#9ca3af", fontSize: 14 }}>¡Agrega tu primer movimiento!</div>}
              {transactions.slice().sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 6).map(t => (
                <div key={t.id} onClick={() => setShowDetail(t)} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "11px 16px", borderTop: "1px solid #f9f9f9", cursor: "pointer" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                    <div style={{ width: 38, height: 38, borderRadius: 11, background: TYPE_COLORS[t.type] + "18", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>{TYPE_ICONS[t.type]}</div>
                    <div><div style={{ fontWeight: 700, fontSize: 14 }}>{t.description}</div><div style={{ fontSize: 11, color: "#9ca3af" }}>{t.category} · {fmtDate(t.date)}</div></div>
                  </div>
                  <div style={{ fontWeight: 800, color: TYPE_COLORS[t.type], fontSize: 14 }}>{t.type === "income" ? "+" : "-"}{fmt(t.amount, currency)}</div>
                </div>
              ))}
            </div>
            <Btn onClick={exportExcel} variant="dark" full>📊 Exportar a Excel</Btn>
            <div style={{ marginTop: 10 }}><Btn onClick={handleLogout} variant="ghost" full>🚪 Cerrar sesión</Btn></div>
          </>
        )}

        {tab === "transactions" && (
          <>
            <div style={{ position: "relative", marginBottom: 12 }}>
              <span style={{ position: "absolute", left: 14, top: "50%", transform: "translateY(-50%)", fontSize: 16 }}>🔍</span>
              <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Buscar..." style={{ ...inputSx, paddingLeft: 40 }} />
            </div>
            <div style={{ display: "flex", gap: 8, marginBottom: 14, overflowX: "auto", scrollbarWidth: "none" }}>
              {["all","income","expense","cost"].map(t => <button key={t} onClick={() => setFilterType(t)} style={{ background: filterType === t ? "#4f46e5" : "#f3f4f6", color: filterType === t ? "#fff" : "#374151", border: "none", borderRadius: 20, padding: "7px 14px", fontSize: 13, fontWeight: 700, whiteSpace: "nowrap", cursor: "pointer", flexShrink: 0 }}>{t === "all" ? "Todos" : TYPE_LABELS[t]}</button>)}
            </div>
            <div style={{ background: "#fff", borderRadius: 16, overflow: "hidden", boxShadow: "0 1px 6px rgba(0,0,0,0.06)" }}>
              {filtered.length === 0 && <div style={{ padding: 32, textAlign: "center", color: "#9ca3af" }}><div style={{ fontSize: 32, marginBottom: 8 }}>🔎</div>Sin resultados</div>}
              {filtered.map((t, i) => (
                <div key={t.id} onClick={() => setShowDetail(t)} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "13px 16px", borderTop: i > 0 ? "1px solid #f3f4f6" : "none", cursor: "pointer" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                    <div style={{ width: 42, height: 42, borderRadius: 13, background: TYPE_COLORS[t.type] + "18", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18 }}>{TYPE_ICONS[t.type]}</div>
                    <div><div style={{ fontWeight: 700, fontSize: 14 }}>{t.description}</div><div style={{ fontSize: 11, color: "#9ca3af" }}>{t.category} · {fmtDate(t.date)}</div>{t.contact && <div style={{ fontSize: 11, color: "#4f46e5" }}>👤 {t.contact}</div>}</div>
                  </div>
                  <div style={{ textAlign: "right" }}><div style={{ fontWeight: 800, color: TYPE_COLORS[t.type], fontSize: 14 }}>{t.type === "income" ? "+" : "-"}{fmt(t.amount, currency)}</div><Badge type={t.type} /></div>
                </div>
              ))}
            </div>
            <div style={{ marginTop: 14 }}><Btn onClick={exportExcel} variant="dark" full>📊 Exportar a Excel</Btn></div>
          </>
        )}

        {tab === "analytics" && (
          <>
            <div style={{ background: "#fff", borderRadius: 16, padding: "16px 16px 20px", marginBottom: 14, boxShadow: "0 1px 6px rgba(0,0,0,0.06)" }}>
              <div style={{ fontWeight: 800, marginBottom: 16 }}>📊 Tendencia 6 meses</div>
              <div style={{ display: "flex", alignItems: "flex-end", gap: 6, height: 130 }}>
                {trend.map((m, i) => (
                  <div key={i} style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center" }}>
                    <div style={{ width: "100%", display: "flex", gap: 2, height: 110, alignItems: "flex-end", justifyContent: "center" }}>
                      <div style={{ flex: 1, background: "#10b981", borderRadius: "4px 4px 0 0", height: (m.income / maxTrend * 100) + "%", minHeight: 2 }} />
                      <div style={{ flex: 1, background: "#ef4444", borderRadius: "4px 4px 0 0", height: (m.expense / maxTrend * 100) + "%", minHeight: 2 }} />
                    </div>
                    <div style={{ fontSize: 10, color: "#9ca3af", fontWeight: 700, marginTop: 4 }}>{m.label}</div>
                  </div>
                ))}
              </div>
            </div>
            <div style={{ background: "#fff", borderRadius: 16, padding: 16, boxShadow: "0 1px 6px rgba(0,0,0,0.06)" }}>
              <div style={{ fontWeight: 800, marginBottom: 14 }}>🗂 Por categoría</div>
              {byCategory.map(([cat, data]) => {
                const max = Math.max(...byCategory.map(c => c[1].total), 1);
                return <div key={cat} style={{ marginBottom: 14 }}><div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 5 }}><div style={{ display: "flex", alignItems: "center", gap: 8 }}><span style={{ fontSize: 14, fontWeight: 700 }}>{cat}</span><Badge type={data.type} /></div><span style={{ fontSize: 14, fontWeight: 800, color: TYPE_COLORS[data.type] }}>{fmt(data.total, currency)}</span></div><ProgressBar value={data.total} max={max} color={TYPE_COLORS[data.type]} /></div>;
              })}
            </div>
          </>
        )}

        {tab === "budgets" && (
          <>
            <div style={{ background: "#fff", borderRadius: 16, padding: 16, marginBottom: 14, boxShadow: "0 1px 6px rgba(0,0,0,0.06)" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
                <span style={{ fontWeight: 800, fontSize: 16 }}>🗓 Presupuestos</span>
                <Btn onClick={() => setShowAddBudget(true)} variant="primary" small>+ Agregar</Btn>
              </div>
              {budgets.length === 0 && <div style={{ color: "#9ca3af", fontSize: 14 }}>Sin presupuestos aún.</div>}
              {budgets.map(b => {
                const used = budgetUsage(b.category); const pct = used / b.limit * 100; const over = pct >= 100;
                return <div key={b.id} style={{ marginBottom: 16 }}><div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}><span style={{ fontWeight: 700, fontSize: 14 }}>{b.category}</span><div style={{ display: "flex", alignItems: "center", gap: 8 }}><span style={{ fontSize: 13, fontWeight: 700, color: over ? "#ef4444" : pct >= 70 ? "#f59e0b" : "#374151" }}>{fmt(used, currency)} / {fmt(b.limit, currency)}</span><button onClick={() => deleteBudget(b.id)} style={{ background: "none", border: "none", fontSize: 14, cursor: "pointer", opacity: 0.4 }}>🗑</button></div></div><ProgressBar value={used} max={b.limit} /><div style={{ fontSize: 11, color: over ? "#ef4444" : pct >= 70 ? "#f59e0b" : "#9ca3af", marginTop: 4, textAlign: "right" }}>{over ? "Superado por " + fmt(used - b.limit, currency) : "Disponible: " + fmt(b.limit - used, currency)}</div></div>;
              })}
            </div>
            <div style={{ background: "#fff", borderRadius: 16, padding: 16, marginBottom: 14, boxShadow: "0 1px 6px rgba(0,0,0,0.06)" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
                <span style={{ fontWeight: 800, fontSize: 16 }}>🏆 Metas de ahorro</span>
                <Btn onClick={() => setShowAddGoal(true)} variant="success" small>+ Nueva</Btn>
              </div>
              {goals.length === 0 && <div style={{ color: "#9ca3af", fontSize: 14 }}>Sin metas aún.</div>}
              {goals.map(g => <GoalCard key={g.id} g={g} profile={profile} onDelete={deleteGoal} onAdd={addToGoal} fmtFn={fmt} fmtDate={fmtDate} />)}
            </div>
            <div style={{ background: "#fff", borderRadius: 16, padding: 16, boxShadow: "0 1px 6px rgba(0,0,0,0.06)" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
                <span style={{ fontWeight: 800, fontSize: 16 }}>💼 Deudas y préstamos</span>
                <Btn onClick={() => setShowAddDebt(true)} variant="warning" small>+ Registrar</Btn>
              </div>
              {debts.length === 0 && <div style={{ color: "#9ca3af", fontSize: 14 }}>Sin deudas registradas.</div>}
              {["owed","owing"].map(dtype => {
                const list = debts.filter(d => d.type === dtype);
                if (list.length === 0) return null;
                return <div key={dtype}><div style={{ fontSize: 13, fontWeight: 800, color: dtype === "owed" ? "#059669" : "#dc2626", marginBottom: 8 }}>{dtype === "owed" ? "🤝 Me deben" : "⚠️ Yo debo"}</div>{list.map(d => <div key={d.id} style={{ background: d.paid ? "#f9fafb" : dtype === "owed" ? "#f0fdf4" : "#fff1f2", borderRadius: 12, padding: 12, marginBottom: 10, opacity: d.paid ? 0.6 : 1 }}><div style={{ display: "flex", justifyContent: "space-between" }}><div><div style={{ fontWeight: 700, fontSize: 14, textDecoration: d.paid ? "line-through" : "none" }}>{d.person}</div><div style={{ fontSize: 12, color: "#6b7280" }}>{d.reason || "Sin descripción"} · {fmtDate(d.date)}</div></div><div style={{ textAlign: "right" }}><div style={{ fontWeight: 800, color: dtype === "owed" ? "#059669" : "#dc2626", fontSize: 16 }}>{fmt(d.amount, currency)}</div><div style={{ fontSize: 11, color: d.paid ? "#059669" : "#f59e0b", fontWeight: 700 }}>{d.paid ? "✅ Liquidado" : "🕐 Pendiente"}</div></div></div><div style={{ display: "flex", gap: 8, marginTop: 10 }}><button onClick={() => toggleDebt(d.id)} style={{ flex: 1, background: d.paid ? "#f3f4f6" : "#d1fae5", color: d.paid ? "#6b7280" : "#059669", border: "none", borderRadius: 10, padding: "7px 0", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>{d.paid ? "↩ Reabrir" : "✅ Liquidar"}</button><button onClick={() => deleteDebt(d.id)} style={{ background: "#fee2e2", color: "#dc2626", border: "none", borderRadius: 10, padding: "7px 14px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>🗑</button></div></div>)}</div>;
              })}
            </div>
          </>
        )}

        {tab === "contacts" && (
          <>
            <Btn onClick={() => setShowAddContact(true)} variant="primary" full>+ Agregar persona</Btn>
            {contacts.length === 0 && <div style={{ color: "#9ca3af", textAlign: "center", marginTop: 20 }}>Sin contactos aún.</div>}
            {contacts.map(c => {
              const total = contactTotal(c.name);
              return <div key={c.id} onClick={() => setShowContactDetail(c)} style={{ background: "#fff", borderRadius: 16, padding: 16, marginBottom: 12, boxShadow: "0 1px 6px rgba(0,0,0,0.06)", cursor: "pointer", display: "flex", alignItems: "center", gap: 14 }}><div style={{ width: 50, height: 50, borderRadius: "50%", background: "linear-gradient(135deg,#4f46e5,#7c3aed)", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontWeight: 800, fontSize: 20, flexShrink: 0 }}>{c.name.charAt(0).toUpperCase()}</div><div style={{ flex: 1, minWidth: 0 }}><div style={{ fontWeight: 800, fontSize: 15 }}>{c.name}</div>{c.phone && <div style={{ fontSize: 12, color: "#9ca3af" }}>📞 {c.phone}</div>}{c.notes && <div style={{ fontSize: 12, color: "#6b7280", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{c.notes}</div>}</div><div style={{ textAlign: "right" }}><div style={{ fontWeight: 800, color: "#10b981", fontSize: 14 }}>{fmt(total, currency)}</div></div></div>;
            })}
          </>
        )}
      </div>

      <div style={{ position: "fixed", bottom: 0, left: "50%", transform: "translateX(-50%)", width: "100%", maxWidth: 430, background: "#fff", borderTop: "1px solid #f0f0f0", display: "flex", padding: "8px 0 22px", zIndex: 100 }}>
        {[{id:"dashboard",icon:"🏠",label:"Inicio"},{id:"transactions",icon:"💳",label:"Movimientos"},{id:"analytics",icon:"📈",label:"Análisis"},{id:"budgets",icon:"🎯",label:"Metas"},{id:"contacts",icon:"👥",label:"Personas"}].map(n => (
          <button key={n.id} onClick={() => setTab(n.id)} style={{ flex: 1, background: "none", border: "none", display: "flex", flexDirection: "column", alignItems: "center", gap: 2, cursor: "pointer", padding: "4px 0", color: tab === n.id ? "#4f46e5" : "#9ca3af" }}>
            <span style={{ fontSize: 22 }}>{n.icon}</span>
            <span style={{ fontSize: 9, fontWeight: 700 }}>{n.label}</span>
          </button>
        ))}
      </div>

      <button onClick={() => setShowAdd(true)} style={{ position: "fixed", bottom: 92, right: "calc(50% - 200px)", background: "linear-gradient(135deg,#4f46e5,#7c3aed)", color: "#fff", border: "none", borderRadius: "50%", width: 58, height: 58, fontSize: 30, cursor: "pointer", boxShadow: "0 6px 20px rgba(79,70,229,0.45)", zIndex: 110, lineHeight: 1 }}>+</button>

      {showAdd && (
        <Modal title="Nuevo movimiento" onClose={() => setShowAdd(false)}>
          <Field label="Tipo"><div style={{ display: "flex", gap: 8, marginBottom: 4 }}>{["income","expense","cost"].map(t => <button key={t} onClick={() => setForm({ ...form, type: t, category: CATEGORIES[t][0] })} style={{ flex: 1, background: form.type === t ? TYPE_COLORS[t] + "20" : "#f3f4f6", color: form.type === t ? TYPE_COLORS[t] : "#374151", border: form.type === t ? "2px solid " + TYPE_COLORS[t] : "2px solid transparent", borderRadius: 12, padding: "10px 0", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>{TYPE_ICONS[t]} {TYPE_LABELS[t]}</button>)}</div></Field>
          <Field label="Monto"><Inp type="number" placeholder="0" value={form.amount} onChange={e => setForm({ ...form, amount: e.target.value })} /></Field>
          <Field label="Categoría"><Sel value={form.category} onChange={e => setForm({ ...form, category: e.target.value })}>{CATEGORIES[form.type].map(c => <option key={c}>{c}</option>)}</Sel></Field>
          <Field label="Descripción"><Inp placeholder="Ej. Pago de cliente..." value={form.description} onChange={e => setForm({ ...form, description: e.target.value })} /></Field>
          <Field label="Fecha"><Inp type="date" value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} /></Field>
          <Field label="Nota (opcional)"><Textarea placeholder="Detalles extras..." value={form.note} onChange={e => setForm({ ...form, note: e.target.value })} /></Field>
          {form.type === "income" && <Field label="Persona (opcional)"><Sel value={form.contact} onChange={e => setForm({ ...form, contact: e.target.value })}><option value="">Sin persona</option>{contacts.map(c => <option key={c.id}>{c.name}</option>)}</Sel></Field>}
          <Btn onClick={addTx} variant="success" full disabled={!form.amount || !form.description}>Guardar en la nube ☁️</Btn>
        </Modal>
      )}

      {showDetail && (
        <Modal title="Detalle" onClose={() => setShowDetail(null)}>
          <div style={{ textAlign: "center", marginBottom: 24 }}><div style={{ fontSize: 52, marginBottom: 8 }}>{TYPE_ICONS[showDetail.type]}</div><div style={{ fontSize: 34, fontWeight: 900, color: TYPE_COLORS[showDetail.type] }}>{fmt(showDetail.amount, currency)}</div><Badge type={showDetail.type} /></div>
          {[["Descripción",showDetail.description],["Categoría",showDetail.category],["Fecha",fmtDate(showDetail.date)],showDetail.contact?["Persona","👤 " + showDetail.contact]:null,showDetail.note?["Nota",showDetail.note]:null].filter(Boolean).map(([k,v]) => <div key={k} style={{ display: "flex", justifyContent: "space-between", padding: "11px 0", borderBottom: "1px solid #f3f4f6" }}><span style={{ color: "#9ca3af", fontSize: 14 }}>{k}</span><span style={{ fontWeight: 700, fontSize: 14, maxWidth: "60%", textAlign: "right" }}>{v}</span></div>)}
          <div style={{ marginTop: 20 }}><Btn onClick={() => deleteTx(showDetail.id)} variant="danger" full>🗑 Eliminar</Btn></div>
        </Modal>
      )}

      {showAddBudget && (
        <Modal title="Nuevo presupuesto" onClose={() => setShowAddBudget(false)}>
          <Field label="Categoría"><Sel value={budgetForm.category} onChange={e => setBudgetForm({ ...budgetForm, category: e.target.value })}>{[...CATEGORIES.expense,...CATEGORIES.cost].map(c => <option key={c}>{c}</option>)}</Sel></Field>
          <Field label="Límite mensual"><Inp type="number" placeholder="0" value={budgetForm.limit} onChange={e => setBudgetForm({ ...budgetForm, limit: e.target.value })} /></Field>
          <Btn onClick={addBudget} variant="primary" full disabled={!budgetForm.limit}>Guardar</Btn>
        </Modal>
      )}

      {showAddGoal && (
        <Modal title="Nueva meta" onClose={() => setShowAddGoal(false)}>
          <Field label="Nombre"><Inp placeholder="Ej. Fondo de emergencia" value={goalForm.name} onChange={e => setGoalForm({ ...goalForm, name: e.target.value })} /></Field>
          <Field label="Meta total"><Inp type="number" placeholder="0" value={goalForm.target} onChange={e => setGoalForm({ ...goalForm, target: e.target.value })} /></Field>
          <Field label="Ya tengo ahorrado"><Inp type="number" placeholder="0" value={goalForm.saved} onChange={e => setGoalForm({ ...goalForm, saved: e.target.value })} /></Field>
          <Field label="Fecha límite (opcional)"><Inp type="date" value={goalForm.deadline} onChange={e => setGoalForm({ ...goalForm, deadline: e.target.value })} /></Field>
          <Field label="Color"><div style={{ display: "flex", gap: 10 }}>{["#6366f1","#10b981","#f59e0b","#ef4444","#8b5cf6","#06b6d4"].map(c => <div key={c} onClick={() => setGoalForm({ ...goalForm, color: c })} style={{ width: 32, height: 32, borderRadius: "50%", background: c, cursor: "pointer", border: goalForm.color === c ? "3px solid #1f2937" : "3px solid transparent" }} />)}</div></Field>
          <Btn onClick={addGoal} variant="success" full disabled={!goalForm.name || !goalForm.target}>Guardar meta</Btn>
        </Modal>
      )}

      {showAddDebt && (
        <Modal title="Registrar deuda" onClose={() => setShowAddDebt(false)}>
          <Field label="¿Quién debe?"><div style={{ display: "flex", gap: 8 }}>{["owed","owing"].map(t => <button key={t} onClick={() => setDebtForm({ ...debtForm, type: t })} style={{ flex: 1, background: debtForm.type === t ? (t === "owed" ? "#d1fae5" : "#fee2e2") : "#f3f4f6", color: debtForm.type === t ? (t === "owed" ? "#059669" : "#dc2626") : "#374151", border: "none", borderRadius: 12, padding: "11px 0", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>{t === "owed" ? "🤝 Me deben" : "⚠️ Yo debo"}</button>)}</div></Field>
          <Field label="Persona"><Inp placeholder="Nombre" value={debtForm.person} onChange={e => setDebtForm({ ...debtForm, person: e.target.value })} /></Field>
          <Field label="Monto"><Inp type="number" placeholder="0" value={debtForm.amount} onChange={e => setDebtForm({ ...debtForm, amount: e.target.value })} /></Field>
          <Field label="Razón"><Inp placeholder="Ej. Préstamo personal" value={debtForm.reason} onChange={e => setDebtForm({ ...debtForm, reason: e.target.value })} /></Field>
          <Field label="Fecha"><Inp type="date" value={debtForm.date} onChange={e => setDebtForm({ ...debtForm, date: e.target.value })} /></Field>
          <Btn onClick={addDebt} variant="warning" full disabled={!debtForm.person || !debtForm.amount}>Registrar</Btn>
        </Modal>
      )}

      {showAddContact && (
        <Modal title="Nueva persona" onClose={() => setShowAddContact(false)}>
          <Field label="Nombre"><Inp placeholder="Nombre completo" value={contactForm.name} onChange={e => setContactForm({ ...contactForm, name: e.target.value })} /></Field>
          <Field label="Teléfono"><Inp placeholder="300-000-0000" value={contactForm.phone} onChange={e => setContactForm({ ...contactForm, phone: e.target.value })} /></Field>
          <Field label="Notas"><Textarea placeholder="¿Quién es?" value={contactForm.notes} onChange={e => setContactForm({ ...contactForm, notes: e.target.value })} /></Field>
          <Btn onClick={addContact} variant="primary" full disabled={!contactForm.name}>Guardar</Btn>
        </Modal>
      )}

      {showContactDetail && (
        <Modal title={showContactDetail.name} onClose={() => setShowContactDetail(null)}>
          <div style={{ textAlign: "center", marginBottom: 20 }}><div style={{ width: 72, height: 72, borderRadius: "50%", background: "linear-gradient(135deg,#4f46e5,#7c3aed)", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontWeight: 900, fontSize: 28, margin: "0 auto 12px" }}>{showContactDetail.name.charAt(0).toUpperCase()}</div><div style={{ fontWeight: 800, fontSize: 20 }}>{showContactDetail.name}</div>{showContactDetail.phone && <div style={{ color: "#9ca3af" }}>📞 {showContactDetail.phone}</div>}</div>
          <div style={{ background: "#f0fdf4", borderRadius: 14, padding: 14, marginBottom: 16, textAlign: "center" }}><div style={{ fontSize: 11, color: "#059669", fontWeight: 700 }}>TOTAL RECIBIDO</div><div style={{ fontSize: 30, fontWeight: 900, color: "#059669" }}>{fmt(contactTotal(showContactDetail.name), currency)}</div></div>
          {showContactDetail.notes && <div style={{ background: "#fafafa", borderRadius: 12, padding: 12, fontSize: 14, color: "#6b7280", marginBottom: 14 }}>📝 {showContactDetail.notes}</div>}
          <div style={{ fontWeight: 800, marginBottom: 10 }}>Movimientos vinculados</div>
          {transactions.filter(t => t.contact === showContactDetail.name).length === 0 ? <div style={{ color: "#9ca3af", fontSize: 14 }}>Sin movimientos aún.</div> : transactions.filter(t => t.contact === showContactDetail.name).sort((a, b) => new Date(b.date) - new Date(a.date)).map(t => <div key={t.id} style={{ display: "flex", justifyContent: "space-between", padding: "11px 0", borderBottom: "1px solid #f3f4f6" }}><div><div style={{ fontWeight: 700, fontSize: 14 }}>{t.description}</div><div style={{ fontSize: 12, color: "#9ca3af" }}>{fmtDate(t.date)}</div></div><div style={{ fontWeight: 800, color: "#10b981" }}>+{fmt(t.amount, currency)}</div></div>)}
        </Modal>
      )}

      {showProfile && (
        <Modal title="Mi perfil" onClose={() => setShowProfile(false)}>
          <Field label="Mi nombre"><Inp placeholder="Tu nombre" value={profileForm.name} onChange={e => setProfileForm({ ...profileForm, name: e.target.value })} /></Field>
          <Field label="Moneda"><Sel value={profileForm.currency} onChange={e => setProfileForm({ ...profileForm, currency: e.target.value })}>{["COP","USD","EUR","MXN","ARS","CLP","PEN","BRL"].map(c => <option key={c}>{c}</option>)}</Sel></Field>
          <Field label="Meta de ahorro mensual"><Inp type="number" placeholder="0" value={profileForm.monthly_saving_goal} onChange={e => setProfileForm({ ...profileForm, monthly_saving_goal: e.target.value })} /></Field>
          <Btn onClick={saveProfile} variant="primary" full>Guardar perfil ☁️</Btn>
          <div style={{ marginTop: 10 }}><Btn onClick={handleLogout} variant="danger" full>🚪 Cerrar sesión</Btn></div>
        </Modal>
      )}
    </div>
  );
}
