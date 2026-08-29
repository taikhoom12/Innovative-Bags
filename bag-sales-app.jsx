import React, { useState, useEffect, useRef, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  LayoutDashboard, ClipboardList, PlusCircle, Package, FileSpreadsheet,
  Bell, BarChart3, Settings as SettingsIcon, Search, Printer, X, Check,
  ChevronDown, ChevronUp, Trash2, Pencil, Eye, Copy, AlertCircle,
  CheckCircle2, Clock, Download, MessageSquare, ArrowRight, Upload, Plus,
  ArrowUp, ArrowDown, Truck, CreditCard, Phone, User, ChevronsUpDown,
  Building2, Database, SlidersHorizontal,
} from "lucide-react";
import {
  LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer, PieChart, Pie, Cell,
} from "recharts";

/* ============================== DESIGN TOKENS ============================== */
const T = {
  primary: "#243447",
  primaryHover: "#1B2836",
  bg: "#F7F7F5",
  surface: "#FFFFFF",
  surface2: "#F1F3F2",
  textPrimary: "#1F2933",
  textSecondary: "#667085",
  textTertiary: "#98A2B3",
  border: "#E4E7EC",
  borderStrong: "#D0D5DD",
  success: "#3E7C63",
  successBg: "#E9F1EC",
  warning: "#B9853B",
  warningBg: "#F8EFDE",
  error: "#B85C5C",
  errorBg: "#F6E8E8",
  accent: "#2D6A8F",
  accentBg: "#E7EEF3",
  accentHover: "#245A7A",
};
const SPACE = { xs: 4, sm: 8, md: 12, lg: 16, xl: 20, xxl: 24, xxxl: 32 };
const RADIUS = { input: 8, button: 8, card: 12, container: 15 };
const EASE = "cubic-bezier(0.4, 0, 0.2, 1)";

const STATUS_META = {
  Draft: { color: T.textSecondary, dot: T.textTertiary },
  Confirmed: { color: T.accent, dot: T.accent },
  Processing: { color: T.warning, dot: T.warning },
  Ready: { color: T.accent, dot: T.accent },
  Delivered: { color: T.success, dot: T.success },
  Cancelled: { color: T.error, dot: T.error },
};
const PAYMENT_META = {
  Unpaid: { color: T.error, dot: T.error },
  Partial: { color: T.warning, dot: T.warning },
  Paid: { color: T.success, dot: T.success },
};

const FONTS = `@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@600;700;800&family=Inter:wght@400;500;600;700&display=swap');`;

/* ============================== HELPERS ============================== */
const uid = (p = "") => p + Math.random().toString(36).slice(2, 9);
const fmt = (n) => "Rs. " + Math.round(Number(n) || 0).toLocaleString("en-PK");
const fmtDate = (d) => {
  if (!d) return "-";
  const dt = new Date(d);
  if (isNaN(dt)) return d;
  return dt.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
};
const todayISO = () => new Date().toISOString().slice(0, 10);
const addDays = (n) => { const d = new Date(); d.setDate(d.getDate() + n); return d.toISOString().slice(0, 10); };
const norm = (s) => (s || "").toString().toLowerCase().trim().replace(/\s+/g, " ");
const monthLabel = () => new Date().toLocaleDateString("en-US", { month: "short", year: "numeric" });
const greeting = () => { const h = new Date().getHours(); return h < 12 ? "Good morning" : h < 17 ? "Good afternoon" : "Good evening"; };

function getBounds(period) {
  const now = new Date();
  let start, end;
  if (period === "Today") { start = new Date(now); start.setHours(0, 0, 0, 0); end = new Date(now); end.setHours(23, 59, 59, 999); }
  else if (period === "This Week") { start = new Date(now); start.setDate(now.getDate() - now.getDay()); start.setHours(0, 0, 0, 0); end = now; }
  else if (period === "This Month") { start = new Date(now.getFullYear(), now.getMonth(), 1); end = now; }
  else if (period === "Last Month") { start = new Date(now.getFullYear(), now.getMonth() - 1, 1); end = new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59); }
  else if (period === "This Year") { start = new Date(now.getFullYear(), 0, 1); end = now; }
  else return null;
  return { start, end };
}
function getPrevBounds(b) {
  const dur = b.end.getTime() - b.start.getTime();
  const prevEnd = new Date(b.start.getTime() - 1);
  const prevStart = new Date(prevEnd.getTime() - dur);
  return { start: prevStart, end: prevEnd };
}
function inBounds(dateStr, b) {
  if (!b) return true;
  const d = new Date(dateStr);
  return d >= b.start && d <= b.end;
}
function deltaLabel(cur, prev) {
  if (prev <= 0) return cur > 0 ? { text: "New this period", tone: "success" } : { text: "No prior data", tone: "neutral" };
  const pct = ((cur - prev) / prev) * 100;
  const tone = pct >= 0 ? "success" : "error";
  return { text: `${pct >= 0 ? "+" : ""}${pct.toFixed(1)}% vs previous period`, tone };
}

/* ============================== SEED DATA ============================== */
function seedProducts() {
  const names = [
    ["Classic Tote Black", "BT001", 2500, 250, "Tote"],
    ["Classic Tote Brown", "BT002", 2500, 300, "Tote"],
    ["Classic Tote Beige", "BT003", 2600, 300, "Tote"],
    ["Travel Bag Large", "TB001", 4500, 500, "Travel"],
    ["Travel Bag Medium", "TB002", 3800, 400, "Travel"],
    ["Evening Clutch Gold", "CL001", 1800, 350, "Clutch"],
    ["Evening Clutch Silver", "CL002", 1800, 350, "Clutch"],
    ["Evening Clutch Rose", "CL003", 1900, 380, "Clutch"],
    ["Sling Bag 105 Black", "SB105-BK", 2100, 220, "Sling"],
    ["Sling Bag 105 Brown", "SB105-BR", 2100, 220, "Sling"],
    ["Sling Bag Mini Tan", "SB110-TN", 1600, 200, "Sling"],
    ["Ladies Handbag Maroon", "HB201", 3200, 400, "Tote"],
    ["Ladies Handbag Navy", "HB202", 3200, 400, "Tote"],
    ["Canvas Backpack Grey", "BP301", 2800, 260, "Backpack"],
    ["Canvas Backpack Olive", "BP302", 2800, 260, "Backpack"],
    ["Quilted Crossbody Pink", "CB401", 2300, 300, "Sling"],
    ["Quilted Crossbody Cream", "CB402", 2300, 300, "Sling"],
    ["Office Tote Charcoal", "OT501", 3600, 420, "Tote"],
    ["Weekend Duffel Tan", "DF601", 4200, 480, "Travel"],
    ["Kids Mini Backpack", "BP310", 1400, 150, "Backpack"],
  ];
  return names.map(([name, sku, price, commission, category]) => ({
    id: uid("p_"), name, sku, sellingPrice: price, commission, category,
    supplier: "", notes: "", active: true, createdAt: todayISO(), updatedAt: todayISO(),
  }));
}
function seedCustomers() {
  const list = [
    ["Ali Traders", "03001234567", "Shop 12, Tariq Road, Karachi"],
    ["Sana Boutique", "03211234567", "Anarkali Bazaar, Lahore"],
    ["Hina Fashions", "03331234567", "Model Town, Lahore"],
    ["Zara Collections", "03451234567", "Gulshan-e-Iqbal, Karachi"],
    ["Ayesha Malik", "03011234567", "DHA Phase 5, Karachi"],
    ["Faiza Store", "03151234567", "Satellite Town, Rawalpindi"],
    ["Noor Handbags", "03021234567", "F-10 Markaz, Islamabad"],
    ["Bushra Boutique", "03361234567", "Cantt, Multan"],
  ];
  return list.map(([name, phone, address]) => ({ id: uid("c_"), name, phone, whatsapp: phone, address, notes: "" }));
}
function seedOrdersAndReminders(products, customers) {
  const statuses = ["Delivered", "Delivered", "Confirmed", "Processing", "Delivered", "Ready", "Cancelled", "Delivered", "Delivered", "Confirmed", "Processing", "Delivered", "Ready", "Delivered", "Confirmed"];
  const payments = ["Paid", "Paid", "Unpaid", "Partial", "Paid", "Unpaid", "Paid", "Paid", "Partial", "Unpaid", "Paid", "Paid", "Unpaid", "Paid", "Partial"];
  const orders = [];
  for (let i = 0; i < 15; i++) {
    const cust = customers[i % customers.length];
    const itemCount = 1 + (i % 3);
    const items = [];
    for (let j = 0; j < itemCount; j++) {
      const prod = products[(i * 3 + j) % products.length];
      const qty = 1 + ((i + j) % 4);
      items.push({ id: uid("oi_"), productId: prod.id, name: prod.name, sku: prod.sku, unitPrice: prod.sellingPrice, commissionPerUnit: prod.commission, quantity: qty, totalPrice: prod.sellingPrice * qty, totalCommission: prod.commission * qty });
    }
    const subtotal = items.reduce((s, it) => s + it.totalPrice, 0);
    const totalCommission = items.reduce((s, it) => s + it.totalCommission, 0);
    const daysAgo = 40 - i * 2;
    orders.push({
      id: uid("o_"), orderNumber: "ORD-" + (1030 + i), customerId: cust.id, customerName: cust.name, phone: cust.phone,
      address: cust.address, date: addDays(-daysAgo), status: statuses[i], paymentStatus: payments[i], items,
      subtotal, total: subtotal, totalCommission, notes: "", deliveryDate: addDays(-daysAgo + 3), reminderDate: null, createdAt: addDays(-daysAgo),
    });
  }
  const reminders = [];
  const reminderDefs = [
    { idx: 0, type: "Delivery", offset: 0 }, { idx: 2, type: "Payment follow-up", offset: -1 },
    { idx: 3, type: "Customer confirmation", offset: 1 }, { idx: 5, type: "Payment follow-up", offset: 2 },
    { idx: 7, type: "Delivery", offset: 4 },
  ];
  reminderDefs.forEach((r) => {
    const ord = orders[r.idx];
    reminders.push({ id: uid("r_"), orderId: ord.id, orderNumber: ord.orderNumber, customerName: ord.customerName, reminderType: r.type, reminderDate: addDays(r.offset), status: "Pending", notes: "" });
  });
  return { orders, reminders };
}

/* ============================== UI PRIMITIVES ============================== */
function GlobalStyle() {
  return (
    <style>{`
      ${FONTS}
      .bagcrm * { box-sizing: border-box; }
      .bagcrm { -webkit-font-smoothing: antialiased; }
      .bagcrm ::-webkit-scrollbar { width: 8px; height: 8px; }
      .bagcrm ::-webkit-scrollbar-thumb { background: ${T.borderStrong}; border-radius: 8px; }
      .bagcrm table { border-collapse: collapse; width: 100%; }
      .bagcrm th { text-align:left; font-size:11px; text-transform:uppercase; letter-spacing:0.04em; color:${T.textSecondary}; font-weight:600; padding:10px 12px; border-bottom:1px solid ${T.border}; background:${T.surface2}; white-space:nowrap; }
      .bagcrm th.sortable { cursor:pointer; user-select:none; }
      .bagcrm th.sortable:hover { color:${T.textPrimary}; }
      .bagcrm td { font-size:13.5px; padding:11px 12px; border-bottom:1px solid ${T.border}; color:${T.textPrimary}; }
      .bagcrm tr.rowhover:hover td { background:${T.surface2}; }
      .bagcrm button, .bagcrm input, .bagcrm select, .bagcrm textarea { font-family: 'Inter', sans-serif; }
      .bagcrm *:focus-visible { outline: 2px solid ${T.accent}; outline-offset: 2px; }
      .bagcrm button:focus-visible { outline-offset: 2px; }
      .skeleton { background: linear-gradient(90deg, ${T.surface2} 25%, #E9EBE9 37%, ${T.surface2} 63%); background-size: 400% 100%; animation: skeleton-loading 1.4s ease infinite; border-radius: 6px; }
      @keyframes skeleton-loading { 0% { background-position: 100% 50%; } 100% { background-position: 0 50%; } }
      @media (prefers-reduced-motion: reduce) {
        .bagcrm * { animation-duration: 0.001ms !important; transition-duration: 0.001ms !important; }
      }
    `}</style>
  );
}

function Btn({ children, onClick, variant = "primary", size = "md", icon: Icon, type = "button", disabled, style, title }) {
  const sizes = { sm: "6px 11px", md: "9px 15px", lg: "11px 20px" };
  const fontSizes = { sm: 12.5, md: 13.5, lg: 14.5 };
  const variants = {
    primary: { bg: T.primary, bgHover: T.primaryHover, color: "#fff", border: "transparent" },
    accent: { bg: T.accent, bgHover: T.accentHover, color: "#fff", border: "transparent" },
    outline: { bg: T.surface, bgHover: T.surface2, color: T.textPrimary, border: T.border },
    ghost: { bg: "transparent", bgHover: T.surface2, color: T.textSecondary, border: "transparent" },
    danger: { bg: T.surface, bgHover: T.errorBg, color: T.error, border: T.border },
  };
  const v = variants[variant];
  const [hover, setHover] = useState(false);
  const [pressed, setPressed] = useState(false);
  return (
    <button
      type={type} title={title} disabled={disabled} onClick={onClick}
      onMouseEnter={() => setHover(true)} onMouseLeave={() => { setHover(false); setPressed(false); }}
      onMouseDown={() => setPressed(true)} onMouseUp={() => setPressed(false)}
      style={{
        display: "inline-flex", alignItems: "center", gap: 6, fontWeight: 600,
        fontSize: fontSizes[size], borderRadius: RADIUS.button, padding: sizes[size],
        cursor: disabled ? "not-allowed" : "pointer", border: `1px solid ${v.border}`,
        background: disabled ? T.surface2 : (hover ? v.bgHover : v.bg),
        color: disabled ? T.textTertiary : v.color,
        transition: `background ${EASE} 150ms, transform ${EASE} 120ms, box-shadow ${EASE} 150ms`,
        transform: pressed && !disabled ? "scale(0.98)" : "scale(1)",
        boxShadow: variant === "primary" && hover && !disabled ? "0 2px 6px rgba(36,52,71,0.25)" : "none",
        whiteSpace: "nowrap", ...style,
      }}>
      {Icon && <Icon size={size === "sm" ? 13 : 15} strokeWidth={2} />}
      {children}
    </button>
  );
}

function IconBtn({ icon: Icon, onClick, title, danger }) {
  const [hover, setHover] = useState(false);
  const [pressed, setPressed] = useState(false);
  return (
    <button title={title} aria-label={title} onClick={onClick}
      onMouseEnter={() => setHover(true)} onMouseLeave={() => { setHover(false); setPressed(false); }}
      onMouseDown={() => setPressed(true)} onMouseUp={() => setPressed(false)}
      style={{
        width: 30, height: 30, display: "inline-flex", alignItems: "center", justifyContent: "center",
        borderRadius: 7, border: "none", cursor: "pointer",
        background: hover ? (danger ? T.errorBg : T.surface2) : "transparent",
        color: hover ? (danger ? T.error : T.textPrimary) : T.textSecondary,
        transition: `background 150ms ${EASE}, transform 120ms ${EASE}`,
        transform: pressed ? "scale(0.94)" : "scale(1)",
      }}>
      <Icon size={15} strokeWidth={2} />
    </button>
  );
}

function Card({ children, style, pad = 18 }) {
  return <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: RADIUS.card, padding: pad, ...style }}>{children}</div>;
}

function Field({ label, children, hint, required }) {
  return (
    <label style={{ display: "block", marginBottom: SPACE.md }}>
      {label && <div style={{ fontSize: 12, fontWeight: 600, color: T.textSecondary, marginBottom: 5, letterSpacing: "0.01em" }}>{label}{required && <span style={{ color: T.error }}> *</span>}</div>}
      {children}
      {hint && <div style={{ fontSize: 11.5, color: T.textTertiary, marginTop: 4 }}>{hint}</div>}
    </label>
  );
}

const inputBase = {
  width: "100%", boxSizing: "border-box", fontSize: 13.5, padding: "9px 12px",
  borderRadius: RADIUS.input, border: `1px solid ${T.border}`, background: T.surface,
  color: T.textPrimary, outline: "none", transition: `border-color 150ms ${EASE}, box-shadow 150ms ${EASE}`,
};
function TInput(props) {
  const [focus, setFocus] = useState(false);
  return <input {...props} onFocus={(e) => { setFocus(true); props.onFocus && props.onFocus(e); }} onBlur={(e) => { setFocus(false); props.onBlur && props.onBlur(e); }}
    style={{ ...inputBase, borderColor: focus ? T.accent : T.border, boxShadow: focus ? `0 0 0 3px ${T.accentBg}` : "none", ...props.style }} />;
}
function TTextarea(props) {
  const [focus, setFocus] = useState(false);
  return <textarea {...props} onFocus={(e) => { setFocus(true); props.onFocus && props.onFocus(e); }} onBlur={(e) => { setFocus(false); props.onBlur && props.onBlur(e); }}
    style={{ ...inputBase, resize: "vertical", borderColor: focus ? T.accent : T.border, boxShadow: focus ? `0 0 0 3px ${T.accentBg}` : "none", ...props.style }} />;
}
function TSelect({ value, onChange, options, style, ...rest }) {
  return (
    <select value={value} onChange={onChange} {...rest} style={{ ...inputBase, cursor: "pointer", ...style }}>
      {options.map((o) => (typeof o === "string" ? <option key={o} value={o}>{o}</option> : <option key={o.value} value={o.value}>{o.label}</option>))}
    </select>
  );
}

function StatusBadge({ status, kind = "status" }) {
  const meta = kind === "status" ? STATUS_META[status] : PAYMENT_META[status];
  if (!meta) return <span>{status}</span>;
  return (
    <span style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 12, fontWeight: 600, color: meta.color }}>
      <span style={{ width: 6, height: 6, borderRadius: 3, background: meta.dot, flexShrink: 0 }} />
      {status}
    </span>
  );
}

function Modal({ open, onClose, title, description, children, width = 560 }) {
  if (!open) return null;
  return (
    <div onClick={onClose} role="dialog" aria-modal="true" style={{ position: "fixed", inset: 0, background: "rgba(31,41,51,0.4)", zIndex: 100, display: "flex", alignItems: "flex-start", justifyContent: "center", padding: "6vh 16px", overflowY: "auto" }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: T.surface, borderRadius: RADIUS.container, width: "100%", maxWidth: width, boxShadow: "0 24px 64px rgba(31,41,51,0.22)", border: `1px solid ${T.border}` }}>
        <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", padding: "18px 22px", borderBottom: `1px solid ${T.border}` }}>
          <div>
            <div style={{ fontFamily: "Manrope, sans-serif", fontWeight: 700, fontSize: 16.5, color: T.textPrimary }}>{title}</div>
            {description && <div style={{ fontSize: 12.5, color: T.textSecondary, marginTop: 3 }}>{description}</div>}
          </div>
          <IconBtn icon={X} onClick={onClose} title="Close" />
        </div>
        <div style={{ padding: 22 }}>{children}</div>
      </div>
    </div>
  );
}

function Toast({ toasts }) {
  return (
    <div style={{ position: "fixed", bottom: 20, right: 20, zIndex: 300, display: "flex", flexDirection: "column", gap: 8 }}>
      {toasts.map((t) => (
        <div key={t.id} style={{
          display: "flex", alignItems: "center", gap: 9, background: T.primary, color: "#fff", padding: "11px 16px",
          borderRadius: RADIUS.button, fontSize: 13, fontWeight: 600, boxShadow: "0 10px 28px rgba(31,41,51,0.28)", minWidth: 230,
          borderLeft: `3px solid ${t.type === "error" ? T.error : T.success}`,
        }}>
          {t.type === "error" ? <AlertCircle size={15} /> : <CheckCircle2 size={15} />}
          {t.message}
        </div>
      ))}
    </div>
  );
}

function EmptyState({ icon: Icon, title, body, action }) {
  return (
    <div style={{ textAlign: "center", padding: "64px 20px", color: T.textSecondary }}>
      {Icon && <div style={{ width: 44, height: 44, borderRadius: 12, background: T.surface2, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 14px" }}><Icon size={20} color={T.textTertiary} /></div>}
      <div style={{ fontFamily: "Manrope, sans-serif", fontWeight: 700, fontSize: 15, color: T.textPrimary, marginBottom: 6 }}>{title}</div>
      <div style={{ fontSize: 13, marginBottom: 18, maxWidth: 360, marginLeft: "auto", marginRight: "auto", lineHeight: 1.5 }}>{body}</div>
      {action}
    </div>
  );
}

function SkeletonBlock({ w = "100%", h = 14, style }) {
  return <div className="skeleton" style={{ width: w, height: h, ...style }} />;
}

/* Searchable select with keyboard navigation */
function SearchSelect({ options, placeholder, onSelect, getLabel, getSub, emptyAction }) {
  const [q, setQ] = useState("");
  const [open, setOpen] = useState(false);
  const [activeIdx, setActiveIdx] = useState(0);
  const ref = useRef(null);
  useEffect(() => {
    function onDoc(e) { if (ref.current && !ref.current.contains(e.target)) setOpen(false); }
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, []);
  const filtered = useMemo(() => {
    const n = norm(q);
    if (!n) return options.slice(0, 8);
    return options.filter((o) => norm(getLabel(o) + " " + (getSub ? getSub(o) : "")).includes(n)).slice(0, 8);
  }, [q, options]);
  useEffect(() => setActiveIdx(0), [q, open]);
  const choose = (o) => { onSelect(o); setQ(""); setOpen(false); };
  const onKeyDown = (e) => {
    if (!open) return;
    if (e.key === "ArrowDown") { e.preventDefault(); setActiveIdx((i) => Math.min(i + 1, filtered.length - 1)); }
    else if (e.key === "ArrowUp") { e.preventDefault(); setActiveIdx((i) => Math.max(i - 1, 0)); }
    else if (e.key === "Enter") { e.preventDefault(); if (filtered[activeIdx]) choose(filtered[activeIdx]); }
    else if (e.key === "Escape") { setOpen(false); }
  };
  return (
    <div ref={ref} style={{ position: "relative" }}>
      <div style={{ position: "relative" }}>
        <Search size={14} style={{ position: "absolute", left: 11, top: 11.5, color: T.textTertiary }} />
        <TInput value={q} placeholder={placeholder} onFocus={() => setOpen(true)} onKeyDown={onKeyDown}
          onChange={(e) => { setQ(e.target.value); setOpen(true); }} style={{ paddingLeft: 32 }}
          role="combobox" aria-expanded={open} aria-autocomplete="list" />
      </div>
      {open && (
        <div style={{ position: "absolute", top: "calc(100% + 4px)", left: 0, right: 0, background: T.surface, border: `1px solid ${T.border}`, borderRadius: RADIUS.input, boxShadow: "0 12px 32px rgba(31,41,51,0.16)", zIndex: 40, maxHeight: 280, overflowY: "auto" }} role="listbox">
          {filtered.length === 0 ? (
            <div style={{ padding: "18px 14px", textAlign: "center" }}>
              <div style={{ fontSize: 12.5, color: T.textSecondary, marginBottom: emptyAction ? 8 : 0 }}>No matching products</div>
              {emptyAction}
            </div>
          ) : filtered.map((o, i) => (
            <div key={i} role="option" aria-selected={i === activeIdx} onClick={() => choose(o)} onMouseEnter={() => setActiveIdx(i)}
              style={{ padding: "9px 13px", cursor: "pointer", borderBottom: i < filtered.length - 1 ? `1px solid ${T.surface2}` : "none", background: i === activeIdx ? T.surface2 : "transparent" }}>
              <div style={{ fontSize: 13, fontWeight: 600, color: T.textPrimary }}>{getLabel(o)}</div>
              {getSub && <div style={{ fontSize: 11.5, color: T.textSecondary }}>{getSub(o)}</div>}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

/* Sortable table header helper */
function SortTh({ label, sortKey, sort, setSort, width }) {
  const active = sort.key === sortKey;
  return (
    <th className="sortable" style={{ width }} onClick={() => setSort((s) => ({ key: sortKey, dir: s.key === sortKey && s.dir === "asc" ? "desc" : "asc" }))}>
      <span style={{ display: "inline-flex", alignItems: "center", gap: 3 }}>
        {label}
        {active ? (sort.dir === "asc" ? <ChevronUp size={12} /> : <ChevronDown size={12} />) : <ChevronsUpDown size={11} style={{ opacity: 0.35 }} />}
      </span>
    </th>
  );
}
function sortRows(rows, sort) {
  if (!sort.key) return rows;
  const arr = [...rows];
  arr.sort((a, b) => {
    let av = a[sort.key], bv = b[sort.key];
    if (typeof av === "string") { av = av.toLowerCase(); bv = (bv || "").toLowerCase(); }
    if (av < bv) return sort.dir === "asc" ? -1 : 1;
    if (av > bv) return sort.dir === "asc" ? 1 : -1;
    return 0;
  });
  return arr;
}

/* ============================== WHATSAPP PARSER (unchanged logic) ============================== */
function scoreMatch(a, b) {
  a = norm(a); b = norm(b);
  if (a === b) return 100;
  if (b.includes(a) || a.includes(b)) return 80;
  const aw = a.split(" "), bw = b.split(" ");
  const overlap = aw.filter((w) => bw.includes(w)).length;
  return (overlap / Math.max(aw.length, bw.length)) * 60;
}
function findBestProduct(text, products) {
  let best = null, bestScore = 0;
  products.forEach((p) => {
    const s = Math.max(scoreMatch(text, p.name), p.sku && norm(text).includes(norm(p.sku)) ? 90 : 0);
    if (s > bestScore) { bestScore = s; best = p; }
  });
  return bestScore >= 45 ? best : null;
}
function parseWhatsAppOrder(raw, products) {
  const lines = raw.split("\n").map((l) => l.trim()).filter(Boolean);
  const phoneRe = /(\+?\d[\d\s-]{8,14}\d)/;
  const itemRe = /^(\d+)\s*[xX×]\s*(.+)$/;
  let phone = null, customerName = null;
  const itemLines = [];
  const otherLines = [];
  lines.forEach((line) => {
    const pm = line.match(phoneRe);
    const im = line.match(itemRe);
    if (im) { itemLines.push({ qty: parseInt(im[1], 10), text: im[2].trim() }); return; }
    if (pm && !phone) { phone = pm[1].replace(/\s+/g, ""); return; }
    otherLines.push(line);
  });
  const noteKeywords = /(deliver|delivery|payment|paid|cash|advance|cod|address|note)/i;
  const nameLine = otherLines.find((l) => !noteKeywords.test(l));
  customerName = nameLine || "Unknown Customer";
  const deliveryLine = otherLines.find((l) => /deliver/i.test(l));
  const paymentLine = otherLines.find((l) => /(payment|paid|cash|advance|cod)/i.test(l));
  const addressLine = otherLines.find((l) => /address/i.test(l));
  const extraNotes = otherLines.filter((l) => l !== nameLine && l !== deliveryLine && l !== paymentLine && l !== addressLine);
  const items = itemLines.map((il) => {
    const match = findBestProduct(il.text, products);
    return { id: uid("pi_"), raw: il.text, quantity: il.qty, productId: match ? match.id : null, matchedName: match ? match.name : null };
  });
  return {
    customerName, phone, address: addressLine ? addressLine.replace(/address[:\-]?/i, "").trim() : "",
    deliveryNote: deliveryLine || "", paymentNote: paymentLine || "", notes: extraNotes.join(" · "), items,
  };
}

/* ============================== PRINT SLIP (unchanged logic, refreshed visuals) ============================== */
function buildSlipHtml(order, settings) {
  const rows = order.items.map((it) => `
    <tr>
      <td style="padding:9px 6px;border-bottom:1px solid #E4E7EC;">${it.name}</td>
      <td style="padding:9px 6px;border-bottom:1px solid #E4E7EC;text-align:center;">${it.quantity}</td>
      <td style="padding:9px 6px;border-bottom:1px solid #E4E7EC;text-align:right;">${fmt(it.unitPrice)}</td>
      <td style="padding:9px 6px;border-bottom:1px solid #E4E7EC;text-align:right;">${fmt(it.totalPrice)}</td>
    </tr>`).join("");
  return `
  <div style="font-family:Inter,Arial,sans-serif;color:#1F2933;max-width:640px;margin:0 auto;">
    <div style="display:flex;justify-content:space-between;align-items:flex-start;border-bottom:2px solid #243447;padding-bottom:16px;margin-bottom:20px;">
      <div>
        <div style="font-size:20px;font-weight:800;font-family:Manrope,Arial,sans-serif;color:#243447;">${settings.businessName || "Bag Sales"}</div>
        <div style="font-size:12px;color:#667085;margin-top:3px;">${settings.address || ""}</div>
        <div style="font-size:12px;color:#667085;">${settings.phone || ""}</div>
      </div>
      <div style="text-align:right;">
        <div style="font-size:13px;font-weight:700;color:#243447;">Order Slip</div>
        <div style="font-size:12px;color:#667085;">#${order.orderNumber}</div>
        <div style="font-size:12px;color:#667085;">${fmtDate(order.date)}</div>
      </div>
    </div>
    <div style="display:flex;gap:24px;margin-bottom:18px;font-size:13px;">
      <div><div style="color:#667085;font-size:11px;">Customer</div><div style="font-weight:700;">${order.customerName}</div></div>
      <div><div style="color:#667085;font-size:11px;">Phone</div><div style="font-weight:700;">${order.phone || "-"}</div></div>
      <div><div style="color:#667085;font-size:11px;">Address</div><div style="font-weight:700;">${order.address || "-"}</div></div>
    </div>
    <table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:16px;">
      <thead><tr style="background:#F1F3F2;">
        <th style="text-align:left;padding:9px 6px;color:#667085;font-size:11px;text-transform:uppercase;">Item</th>
        <th style="text-align:center;padding:9px 6px;color:#667085;font-size:11px;text-transform:uppercase;">Qty</th>
        <th style="text-align:right;padding:9px 6px;color:#667085;font-size:11px;text-transform:uppercase;">Unit Price</th>
        <th style="text-align:right;padding:9px 6px;color:#667085;font-size:11px;text-transform:uppercase;">Total</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>
    <div style="display:flex;justify-content:flex-end;">
      <div style="width:240px;font-size:13px;">
        <div style="display:flex;justify-content:space-between;padding:5px 0;"><span>Total Quantity</span><b>${order.items.reduce((s, i) => s + i.quantity, 0)}</b></div>
        <div style="display:flex;justify-content:space-between;padding:6px 0;border-top:1px solid #E4E7EC;margin-top:4px;font-size:15px;"><span>Final Amount</span><b>${fmt(order.total)}</b></div>
        <div style="display:flex;justify-content:space-between;padding:5px 0;"><span>Payment Status</span><b>${order.paymentStatus}</b></div>
      </div>
    </div>
    <div style="margin-top:30px;padding-top:14px;border-top:1px solid #E4E7EC;font-size:11px;color:#667085;text-align:center;">
      Thank you for your order — ${settings.businessName || "Bag Sales"}
    </div>
  </div>`;
}
function printOrderSlip(order, settings) {
  const w = window.open("", "_blank", "width=800,height=900");
  if (!w) return;
  w.document.write(`<html><head><title>Slip ${order.orderNumber}</title></head><body style="padding:24px;">${buildSlipHtml(order, settings)}</body></html>`);
  w.document.close();
  setTimeout(() => { w.focus(); w.print(); }, 300);
}

/* ============================== NAV ============================== */
const NAV = [
  { key: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { key: "orders", label: "Orders", icon: ClipboardList },
  { key: "neworder", label: "New Order", icon: PlusCircle },
  { key: "products", label: "Products", icon: Package },
  { key: "import", label: "Import Excel", icon: FileSpreadsheet },
  { key: "reminders", label: "Reminders", icon: Bell },
  { key: "reports", label: "Reports", icon: BarChart3 },
  { key: "settings", label: "Settings", icon: SettingsIcon },
];

/* ============================== MAIN APP ============================== */
export default function BagSalesApp() {
  const [loaded, setLoaded] = useState(false);
  const [products, setProducts] = useState([]);
  const [customers, setCustomers] = useState([]);
  const [orders, setOrders] = useState([]);
  const [reminders, setReminders] = useState([]);
  const [settings, setSettings] = useState({ businessName: "Peach & Tote Bags", phone: "0300-1234567", address: "Karachi, Pakistan" });
  const [page, setPage] = useState("dashboard");
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [toasts, setToasts] = useState([]);
  const [orderDraft, setOrderDraft] = useState(null);

  const pushToast = useCallback((message, type = "success") => {
    const id = uid();
    setToasts((t) => [...t, { id, message, type }]);
    setTimeout(() => setToasts((t) => t.filter((x) => x.id !== id)), 2800);
  }, []);

  useEffect(() => {
    let cancelled = false;
    async function load() {
      try {
        const res = await window.storage.get("bagcrm-data-v1", false);
        if (res && res.value) {
          const data = JSON.parse(res.value);
          if (!cancelled) {
            setProducts(data.products && data.products.length ? data.products : seedProducts());
            setCustomers(data.customers && data.customers.length ? data.customers : seedCustomers());
            setOrders(data.orders || []);
            setReminders(data.reminders || []);
            setSettings(data.settings || { businessName: "Peach & Tote Bags", phone: "0300-1234567", address: "Karachi, Pakistan" });
            setLoaded(true);
            return;
          }
        }
      } catch (e) { /* fall through to seed */ }
      const p = seedProducts(); const c = seedCustomers();
      const { orders: o, reminders: r } = seedOrdersAndReminders(p, c);
      if (!cancelled) { setProducts(p); setCustomers(c); setOrders(o); setReminders(r); setLoaded(true); }
    }
    load();
    return () => { cancelled = true; };
  }, []);

  useEffect(() => {
    if (!loaded) return;
    const t = setTimeout(() => {
      const data = { products, customers, orders, reminders, settings };
      window.storage.set("bagcrm-data-v1", JSON.stringify(data), false).catch(() => {});
    }, 500);
    return () => clearTimeout(t);
  }, [products, customers, orders, reminders, settings, loaded]);

  const goNewOrder = (draft) => { setOrderDraft(draft || null); setPage("neworder"); setSidebarOpen(false); };

  if (!loaded) {
    return (
      <div className="bagcrm" style={{ minHeight: 480, background: T.bg, borderRadius: RADIUS.container, padding: 24, display: "flex", gap: 16 }}>
        <GlobalStyle />
        <SkeletonBlock w={200} h={480} />
        <div style={{ flex: 1 }}>
          <SkeletonBlock w={220} h={22} style={{ marginBottom: 20 }} />
          <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 12, marginBottom: 16 }}>
            {[0, 1, 2, 3].map((i) => <SkeletonBlock key={i} h={86} />)}
          </div>
          <SkeletonBlock h={220} />
        </div>
      </div>
    );
  }

  return (
    <div className="bagcrm" style={{ fontFamily: "Inter, sans-serif", background: T.bg, minHeight: 640, color: T.textPrimary, display: "flex", borderRadius: RADIUS.container, overflow: "hidden", border: `1px solid ${T.border}` }}>
      <GlobalStyle />

      {/* Sidebar */}
      <div style={{ width: 216, background: T.primary, color: "#fff", flexShrink: 0, display: "flex", flexDirection: "column", padding: "20px 12px" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 9, padding: "0 8px 22px" }}>
          <div style={{ width: 30, height: 30, borderRadius: 8, background: T.accent, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 800, fontFamily: "Manrope, sans-serif", fontSize: 13 }}>
            {(settings.businessName || "B").charAt(0)}
          </div>
          <div style={{ fontFamily: "Manrope, sans-serif", fontWeight: 700, fontSize: 14, lineHeight: 1.2, letterSpacing: "-0.01em" }}>{settings.businessName}</div>
        </div>
        <nav style={{ display: "flex", flexDirection: "column", gap: 1 }}>
          {NAV.map((n) => {
            const active = page === n.key;
            const Icon = n.icon;
            return (
              <button key={n.key} onClick={() => { setPage(n.key); setSidebarOpen(false); if (n.key !== "neworder") setOrderDraft(null); }}
                style={{
                  display: "flex", alignItems: "center", gap: 10, padding: "9px 11px", borderRadius: 7, cursor: "pointer",
                  fontSize: 13, fontWeight: 500, transition: `background 150ms ${EASE}, color 150ms ${EASE}`, textAlign: "left",
                  background: active ? "rgba(255,255,255,0.09)" : "transparent",
                  color: active ? "#fff" : "rgba(255,255,255,0.62)", border: "none",
                  borderLeft: active ? `2px solid ${T.accent}` : "2px solid transparent", position: "relative",
                }}
                onMouseEnter={(e) => { if (!active) e.currentTarget.style.background = "rgba(255,255,255,0.05)"; }}
                onMouseLeave={(e) => { if (!active) e.currentTarget.style.background = "transparent"; }}>
                <Icon size={15} strokeWidth={1.75} />
                {n.label}
              </button>
            );
          })}
        </nav>
        <div style={{ marginTop: "auto", paddingTop: 14, borderTop: "1px solid rgba(255,255,255,0.1)" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 9, padding: "8px 8px" }}>
            <div style={{ width: 26, height: 26, borderRadius: 7, background: "rgba(255,255,255,0.1)", display: "flex", alignItems: "center", justifyContent: "center" }}>
              <User size={13} color="rgba(255,255,255,0.7)" />
            </div>
            <div style={{ fontSize: 12, color: "rgba(255,255,255,0.75)", fontWeight: 500 }}>Business owner</div>
          </div>
        </div>
      </div>

      {/* Main */}
      <div style={{ flex: 1, minWidth: 0, display: "flex", flexDirection: "column" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "14px 24px", borderBottom: `1px solid ${T.border}`, background: T.surface }}>
          <div>
            <div style={{ fontFamily: "Manrope, sans-serif", fontWeight: 700, fontSize: 17, letterSpacing: "-0.01em" }}>{NAV.find((n) => n.key === page)?.label}</div>
            {page === "dashboard" && <div style={{ fontSize: 12, color: T.textSecondary, marginTop: 1 }}>Sales overview · {monthLabel()}</div>}
          </div>
          <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
            <Btn size="sm" variant="outline" icon={Upload} onClick={() => setPage("import")}>Import Excel</Btn>
            <Btn size="sm" variant="primary" icon={Plus} onClick={() => goNewOrder(null)}>New Order</Btn>
          </div>
        </div>
        <div style={{ padding: 24, overflowY: "auto", flex: 1 }}>
          {page === "dashboard" && <Dashboard products={products} orders={orders} reminders={reminders} setReminders={setReminders} goNewOrder={goNewOrder} setPage={setPage} settings={settings} />}
          {page === "orders" && <OrdersPage orders={orders} setOrders={setOrders} settings={settings} pushToast={pushToast} goEdit={(o) => goNewOrder({ editOrder: o })} />}
          {page === "neworder" && <NewOrderPage products={products} customers={customers} setCustomers={setCustomers} orders={orders} setOrders={setOrders} reminders={reminders} setReminders={setReminders} settings={settings} draft={orderDraft} pushToast={pushToast} onDone={() => setPage("orders")} setPage={setPage} />}
          {page === "products" && <ProductsPage products={products} setProducts={setProducts} orders={orders} pushToast={pushToast} />}
          {page === "import" && <ImportExcelPage products={products} setProducts={setProducts} pushToast={pushToast} />}
          {page === "reminders" && <RemindersPage reminders={reminders} setReminders={setReminders} orders={orders} />}
          {page === "reports" && <ReportsPage orders={orders} products={products} />}
          {page === "settings" && <SettingsPage settings={settings} setSettings={setSettings} pushToast={pushToast} products={products} orders={orders} customers={customers} setProducts={setProducts} setOrders={setOrders} setCustomers={setCustomers} setReminders={setReminders} />}
        </div>
      </div>
      <Toast toasts={toasts} />
    </div>
  );
}

/* ============================== DASHBOARD ============================== */
const PERIODS = ["Today", "This Week", "This Month", "Last Month", "This Year", "All Time"];
const CHART_COLORS = { line: T.accent, bar: T.primary, grid: T.border };
const PIE_COLORS = { Draft: T.textTertiary, Confirmed: T.accent, Processing: T.warning, Ready: "#6B8CAE", Delivered: T.success, Cancelled: T.error };

function KpiCard({ label, value, delta, valueColor }) {
  return (
    <Card pad={18} style={{ minWidth: 0 }}>
      <div style={{ fontSize: 11.5, fontWeight: 600, color: T.textSecondary, textTransform: "uppercase", letterSpacing: "0.04em" }}>{label}</div>
      <div style={{ fontFamily: "Manrope, sans-serif", fontWeight: 700, fontSize: 27, marginTop: 8, color: valueColor || T.textPrimary, letterSpacing: "-0.01em" }}>{value}</div>
      {delta && (
        <div style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 12, marginTop: 6, color: delta.tone === "success" ? T.success : delta.tone === "error" ? T.error : T.textTertiary, fontWeight: 500 }}>
          {delta.tone === "success" && <ArrowUp size={12} />}
          {delta.tone === "error" && <ArrowDown size={12} />}
          {delta.text}
        </div>
      )}
    </Card>
  );
}

function Dashboard({ products, orders, reminders, setReminders, goNewOrder, setPage, settings }) {
  const [period, setPeriod] = useState("This Month");
  const bounds = getBounds(period);
  const prevBounds = bounds ? getPrevBounds(bounds) : null;
  const filtered = useMemo(() => orders.filter((o) => inBounds(o.date, bounds)), [orders, period]);
  const prevFiltered = useMemo(() => (prevBounds ? orders.filter((o) => inBounds(o.date, prevBounds)) : []), [orders, period]);

  const totalSales = filtered.reduce((s, o) => s + o.total, 0);
  const prevSales = prevFiltered.reduce((s, o) => s + o.total, 0);
  const totalOrders = filtered.length;
  const prevOrders = prevFiltered.length;
  const commission = filtered.reduce((s, o) => s + o.totalCommission, 0);
  const prevCommission = prevFiltered.reduce((s, o) => s + o.totalCommission, 0);
  const pending = filtered.filter((o) => ["Draft", "Confirmed", "Processing", "Ready"].includes(o.status)).length;
  const completed = filtered.filter((o) => o.status === "Delivered").length;
  const outstanding = filtered.filter((o) => o.paymentStatus !== "Paid").reduce((s, o) => s + o.total * (o.paymentStatus === "Partial" ? 0.5 : 1), 0);
  const bagsSold = filtered.reduce((s, o) => s + o.items.reduce((a, i) => a + i.quantity, 0), 0);

  const salesByDay = useMemo(() => {
    const map = {};
    filtered.forEach((o) => { const k = fmtDate(o.date); map[k] = (map[k] || 0) + o.total; });
    return Object.entries(map).map(([date, total]) => ({ date, total })).slice(-14);
  }, [filtered]);
  const topBags = useMemo(() => {
    const map = {};
    filtered.forEach((o) => o.items.forEach((i) => { map[i.name] = (map[i.name] || 0) + i.quantity; }));
    return Object.entries(map).map(([name, qty]) => ({ name, qty })).sort((a, b) => b.qty - a.qty).slice(0, 6);
  }, [filtered]);
  const statusBreakdown = useMemo(() => {
    const map = {};
    filtered.forEach((o) => { map[o.status] = (map[o.status] || 0) + 1; });
    return Object.entries(map).map(([status, count]) => ({ status, count }));
  }, [filtered]);

  const todaysReminders = reminders.filter((r) => r.status === "Pending" && new Date(r.reminderDate) <= new Date(addDays(1)));
  const recent = [...orders].sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 8);

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 20, flexWrap: "wrap" }}>
        <div style={{ fontFamily: "Manrope, sans-serif", fontWeight: 700, fontSize: 19, letterSpacing: "-0.01em" }}>{greeting()}</div>
        <div style={{ marginLeft: "auto", display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          <Btn icon={Plus} onClick={() => goNewOrder(null)}>New Order</Btn>
          <Btn icon={MessageSquare} variant="outline" onClick={() => goNewOrder({ openWhatsApp: true })}>Paste WhatsApp Order</Btn>
          <Btn icon={Upload} variant="ghost" onClick={() => setPage("import")}>Import Excel</Btn>
          <TSelect value={period} onChange={(e) => setPeriod(e.target.value)} options={PERIODS} style={{ width: 150 }} />
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 14, marginBottom: 16 }}>
        <KpiCard label="Total Sales" value={fmt(totalSales)} delta={deltaLabel(totalSales, prevSales)} />
        <KpiCard label="Total Commission" value={fmt(commission)} delta={deltaLabel(commission, prevCommission)} valueColor={T.success} />
        <KpiCard label="Orders" value={totalOrders} delta={{ text: `${completed} completed`, tone: "neutral" }} />
        <KpiCard label="Pending" value={pending} delta={{ text: pending > 0 ? "Needs attention" : "All clear", tone: pending > 0 ? "error" : "success" }} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14, marginBottom: 16 }}>
        <KpiCard label="Bags Sold" value={bagsSold} />
        <KpiCard label="Outstanding" value={fmt(outstanding)} valueColor={outstanding > 0 ? T.warning : T.textPrimary} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1.3fr 1fr", gap: 14, marginBottom: 16 }}>
        <Card>
          <div style={{ fontWeight: 600, fontFamily: "Manrope, sans-serif", fontSize: 14, marginBottom: 14 }}>Sales over time</div>
          <ResponsiveContainer width="100%" height={200}>
            <LineChart data={salesByDay}>
              <CartesianGrid stroke={CHART_COLORS.grid} vertical={false} strokeDasharray="0" />
              <XAxis dataKey="date" tick={{ fontSize: 10.5, fill: T.textTertiary }} axisLine={{ stroke: T.border }} tickLine={false} />
              <YAxis tick={{ fontSize: 10.5, fill: T.textTertiary }} width={42} axisLine={false} tickLine={false} />
              <Tooltip formatter={(v) => fmt(v)} contentStyle={{ fontSize: 12, borderRadius: 8, border: `1px solid ${T.border}`, boxShadow: "0 8px 20px rgba(31,41,51,0.1)" }} />
              <Line type="monotone" dataKey="total" stroke={CHART_COLORS.line} strokeWidth={2} dot={{ r: 2.5, fill: CHART_COLORS.line }} activeDot={{ r: 4 }} />
            </LineChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <div style={{ fontWeight: 600, fontFamily: "Manrope, sans-serif", fontSize: 14, marginBottom: 14 }}>Order status</div>
          <ResponsiveContainer width="100%" height={168}>
            <PieChart>
              <Pie data={statusBreakdown} dataKey="count" nameKey="status" innerRadius={44} outerRadius={70} paddingAngle={2} stroke="none">
                {statusBreakdown.map((s, i) => <Cell key={i} fill={PIE_COLORS[s.status] || T.textTertiary} />)}
              </Pie>
              <Tooltip contentStyle={{ fontSize: 12, borderRadius: 8, border: `1px solid ${T.border}` }} />
            </PieChart>
          </ResponsiveContainer>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 10, marginTop: 6, justifyContent: "center" }}>
            {statusBreakdown.map((s) => (
              <div key={s.status} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 11.5, color: T.textSecondary }}>
                <div style={{ width: 7, height: 7, borderRadius: 3.5, background: PIE_COLORS[s.status] }} />{s.status}
              </div>
            ))}
          </div>
        </Card>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 16 }}>
        <Card>
          <div style={{ fontWeight: 600, fontFamily: "Manrope, sans-serif", fontSize: 14, marginBottom: 14 }}>Top selling bags</div>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={topBags} layout="vertical" margin={{ left: 6 }}>
              <CartesianGrid stroke={CHART_COLORS.grid} horizontal={false} />
              <XAxis type="number" tick={{ fontSize: 10.5, fill: T.textTertiary }} axisLine={{ stroke: T.border }} tickLine={false} />
              <YAxis type="category" dataKey="name" width={128} tick={{ fontSize: 10.5, fill: T.textSecondary }} axisLine={false} tickLine={false} />
              <Tooltip contentStyle={{ fontSize: 12, borderRadius: 8, border: `1px solid ${T.border}` }} cursor={{ fill: T.surface2 }} />
              <Bar dataKey="qty" fill={CHART_COLORS.bar} radius={[0, 4, 4, 0]} barSize={14} />
            </BarChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <div style={{ display: "flex", alignItems: "center", marginBottom: 12 }}>
            <div style={{ fontWeight: 600, fontFamily: "Manrope, sans-serif", fontSize: 14 }}>Today's reminders</div>
          </div>
          {todaysReminders.length === 0 ? (
            <div style={{ fontSize: 12.5, color: T.textSecondary, padding: "26px 0", textAlign: "center" }}>No reminders due. You're all caught up.</div>
          ) : todaysReminders.map((r) => (
            <div key={r.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "9px 0", borderBottom: `1px solid ${T.border}` }}>
              <Clock size={14} color={T.warning} />
              <div style={{ flex: 1 }}>
                <div style={{ fontSize: 12.5, fontWeight: 600 }}>{r.orderNumber} — {r.reminderType}</div>
                <div style={{ fontSize: 11, color: T.textSecondary }}>{r.customerName} · {fmtDate(r.reminderDate)}</div>
              </div>
              <Btn size="sm" variant="ghost" onClick={() => setReminders((rs) => rs.map((x) => x.id === r.id ? { ...x, status: "Done" } : x))}>Mark done</Btn>
            </div>
          ))}
        </Card>
      </div>

      <Card pad={0}>
        <div style={{ fontWeight: 600, fontFamily: "Manrope, sans-serif", fontSize: 14, padding: "16px 18px 0" }}>Recent orders</div>
        <div style={{ overflowX: "auto" }}>
          <table>
            <thead><tr><th>Order</th><th>Customer</th><th>Items</th><th>Total</th><th>Status</th><th>Date</th></tr></thead>
            <tbody>
              {recent.map((o) => (
                <tr key={o.id} className="rowhover">
                  <td style={{ fontWeight: 600 }}>{o.orderNumber}</td>
                  <td>{o.customerName}</td>
                  <td>{o.items.reduce((s, i) => s + i.quantity, 0)}</td>
                  <td>{fmt(o.total)}</td>
                  <td><StatusBadge status={o.status} /></td>
                  <td style={{ color: T.textSecondary }}>{fmtDate(o.date)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

/* ============================== ORDERS PAGE ============================== */
function OrdersPage({ orders, setOrders, settings, pushToast, goEdit }) {
  const [q, setQ] = useState("");
  const [statusFilter, setStatusFilter] = useState("All");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [sort, setSort] = useState({ key: "date", dir: "desc" });
  const [viewOrder, setViewOrder] = useState(null);
  const [deleteTarget, setDeleteTarget] = useState(null);

  let filtered = orders.filter((o) => {
    const matchesQ = !q || norm(o.orderNumber + " " + o.customerName + " " + o.phone).includes(norm(q));
    const matchesStatus = statusFilter === "All" || o.status === statusFilter;
    const matchesFrom = !dateFrom || o.date >= dateFrom;
    const matchesTo = !dateTo || o.date <= dateTo;
    return matchesQ && matchesStatus && matchesFrom && matchesTo;
  });
  filtered = sortRows(filtered, sort);

  const setStatus = (id, status) => setOrders((os) => os.map((o) => o.id === id ? { ...o, status } : o));
  const setPayment = (id, paymentStatus) => setOrders((os) => os.map((o) => o.id === id ? { ...o, paymentStatus } : o));
  const duplicate = (o) => {
    const copy = { ...o, id: uid("o_"), orderNumber: "ORD-" + (1000 + Math.floor(Math.random() * 9000)), date: todayISO(), status: "Draft", paymentStatus: "Unpaid" };
    setOrders((os) => [copy, ...os]);
    pushToast("Order duplicated as " + copy.orderNumber);
  };
  const doDelete = () => { setOrders((os) => os.filter((o) => o.id !== deleteTarget.id)); pushToast("Order deleted"); setDeleteTarget(null); };

  if (orders.length === 0) {
    return <EmptyState icon={ClipboardList} title="No orders yet" body="Create your first order to start tracking sales and commission." action={<Btn icon={Plus} onClick={() => goEdit(null)}>New Order</Btn>} />;
  }

  return (
    <div>
      <Card pad={14} style={{ marginBottom: 14 }}>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
          <div style={{ position: "relative", flex: "1 1 220px", maxWidth: 300 }}>
            <Search size={14} style={{ position: "absolute", left: 11, top: 11.5, color: T.textTertiary }} />
            <TInput placeholder="Search order #, customer, phone" value={q} onChange={(e) => setQ(e.target.value)} style={{ paddingLeft: 32 }} />
          </div>
          <TSelect value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} options={["All", ...Object.keys(STATUS_META)]} style={{ width: 150 }} />
          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <TInput type="date" value={dateFrom} onChange={(e) => setDateFrom(e.target.value)} style={{ width: 138 }} />
            <span style={{ fontSize: 12, color: T.textTertiary }}>to</span>
            <TInput type="date" value={dateTo} onChange={(e) => setDateTo(e.target.value)} style={{ width: 138 }} />
          </div>
          {(q || statusFilter !== "All" || dateFrom || dateTo) && (
            <Btn size="sm" variant="ghost" onClick={() => { setQ(""); setStatusFilter("All"); setDateFrom(""); setDateTo(""); }}>Clear filters</Btn>
          )}
          <div style={{ marginLeft: "auto", fontSize: 12, color: T.textSecondary }}>{filtered.length} of {orders.length} orders</div>
        </div>
      </Card>

      <Card pad={0}>
        <div style={{ overflowX: "auto" }}>
          <table>
            <thead><tr>
              <SortTh label="Order #" sortKey="orderNumber" sort={sort} setSort={setSort} />
              <SortTh label="Date" sortKey="date" sort={sort} setSort={setSort} />
              <SortTh label="Customer" sortKey="customerName" sort={sort} setSort={setSort} />
              <th>Items</th>
              <SortTh label="Total" sortKey="total" sort={sort} setSort={setSort} />
              <th>Commission</th><th>Payment</th><th>Status</th><th></th>
            </tr></thead>
            <tbody>
              {filtered.map((o) => (
                <tr key={o.id} className="rowhover">
                  <td style={{ fontWeight: 600, cursor: "pointer", color: T.accent }} onClick={() => setViewOrder(o)}>{o.orderNumber}</td>
                  <td style={{ color: T.textSecondary }}>{fmtDate(o.date)}</td>
                  <td>{o.customerName}</td>
                  <td>{o.items.reduce((s, i) => s + i.quantity, 0)}</td>
                  <td style={{ fontWeight: 600 }}>{fmt(o.total)}</td>
                  <td style={{ color: T.success }}>{fmt(o.totalCommission)}</td>
                  <td>
                    <select value={o.paymentStatus} onChange={(e) => setPayment(o.id, e.target.value)}
                      style={{ border: "none", background: "transparent", fontSize: 12, fontWeight: 600, cursor: "pointer", color: PAYMENT_META[o.paymentStatus].color }}>
                      {Object.keys(PAYMENT_META).map((s) => <option key={s} value={s}>{s}</option>)}
                    </select>
                  </td>
                  <td>
                    <select value={o.status} onChange={(e) => setStatus(o.id, e.target.value)}
                      style={{ border: `1px solid ${T.border}`, background: T.surface, color: STATUS_META[o.status].color, fontSize: 12, fontWeight: 600, borderRadius: 6, padding: "4px 8px", cursor: "pointer" }}>
                      {Object.keys(STATUS_META).map((s) => <option key={s} value={s}>{s}</option>)}
                    </select>
                  </td>
                  <td>
                    <div style={{ display: "flex", gap: 2 }}>
                      <IconBtn icon={Eye} title="View" onClick={() => setViewOrder(o)} />
                      <IconBtn icon={Pencil} title="Edit" onClick={() => goEdit(o)} />
                      <IconBtn icon={Printer} title="Print" onClick={() => printOrderSlip(o, settings)} />
                      <IconBtn icon={Copy} title="Duplicate" onClick={() => duplicate(o)} />
                      <IconBtn icon={Trash2} title="Delete" danger onClick={() => setDeleteTarget(o)} />
                    </div>
                  </td>
                </tr>
              ))}
              {filtered.length === 0 && (
                <tr><td colSpan={9}><div style={{ padding: "32px 0", textAlign: "center", color: T.textSecondary, fontSize: 13 }}>No orders match these filters.</div></td></tr>
              )}
            </tbody>
          </table>
        </div>
      </Card>

      <Modal open={!!viewOrder} onClose={() => setViewOrder(null)} title={viewOrder ? "Order " + viewOrder.orderNumber : ""} width={620}>
        {viewOrder && (
          <div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12, marginBottom: 18 }}>
              <div><div style={{ fontSize: 11, color: T.textSecondary }}>Customer</div><div style={{ fontWeight: 600, fontSize: 13.5 }}>{viewOrder.customerName}</div></div>
              <div><div style={{ fontSize: 11, color: T.textSecondary }}>Phone</div><div style={{ fontWeight: 600, fontSize: 13.5 }}>{viewOrder.phone || "-"}</div></div>
              <div><div style={{ fontSize: 11, color: T.textSecondary }}>Address</div><div style={{ fontWeight: 600, fontSize: 13.5 }}>{viewOrder.address || "-"}</div></div>
            </div>
            <table style={{ marginBottom: 16 }}>
              <thead><tr><th>Item</th><th>Qty</th><th>Price</th><th>Total</th><th>Commission</th></tr></thead>
              <tbody>{viewOrder.items.map((it) => (
                <tr key={it.id}><td>{it.name}</td><td>{it.quantity}</td><td>{fmt(it.unitPrice)}</td><td>{fmt(it.totalPrice)}</td><td style={{ color: T.success }}>{fmt(it.totalCommission)}</td></tr>
              ))}</tbody>
            </table>
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13.5, marginBottom: 6 }}><span>Sale amount</span><b>{fmt(viewOrder.total)}</b></div>
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13.5, marginBottom: 6 }}><span>Commission / profit</span><b style={{ color: T.success }}>{fmt(viewOrder.totalCommission)}</b></div>
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13.5, marginBottom: 18 }}><span>Status</span><StatusBadge status={viewOrder.status} /></div>
            <div style={{ display: "flex", gap: 8 }}>
              <Btn icon={Printer} onClick={() => printOrderSlip(viewOrder, settings)}>Print slip</Btn>
              <Btn variant="outline" icon={Pencil} onClick={() => { setViewOrder(null); goEdit(viewOrder); }}>Edit</Btn>
            </div>
          </div>
        )}
      </Modal>

      <Modal open={!!deleteTarget} onClose={() => setDeleteTarget(null)} title="Delete order?" description="This action cannot be undone." width={380}>
        {deleteTarget && (
          <div>
            <div style={{ fontSize: 13.5, marginBottom: 18, color: T.textSecondary }}>Order <b style={{ color: T.textPrimary }}>{deleteTarget.orderNumber}</b> and its history will be permanently removed.</div>
            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
              <Btn variant="ghost" onClick={() => setDeleteTarget(null)}>Cancel</Btn>
              <Btn variant="danger" onClick={doDelete}>Delete order</Btn>
            </div>
          </div>
        )}
      </Modal>
    </div>
  );
}

/* ============================== NEW ORDER PAGE ============================== */
function NewOrderPage({ products, customers, setCustomers, orders, setOrders, reminders, setReminders, settings, draft, pushToast, onDone, setPage }) {
  const editingOrder = draft && draft.editOrder;
  const [tab, setTab] = useState(draft && draft.openWhatsApp ? "wa" : "manual");
  const [customer, setCustomer] = useState(editingOrder ? { name: editingOrder.customerName, phone: editingOrder.phone, address: editingOrder.address } : { name: "", phone: "", whatsapp: "", address: "", notes: "" });
  const [items, setItems] = useState(editingOrder ? editingOrder.items.map((i) => ({ ...i })) : []);
  const [status, setStatus] = useState(editingOrder ? editingOrder.status : "Confirmed");
  const [paymentStatus, setPaymentStatus] = useState(editingOrder ? editingOrder.paymentStatus : "Unpaid");
  const [deliveryDate, setDeliveryDate] = useState(editingOrder ? editingOrder.deliveryDate : addDays(3));
  const [waText, setWaText] = useState("");
  const [parsed, setParsed] = useState(null);
  const [lastCreated, setLastCreated] = useState(null);

  const addProduct = (p) => {
    setItems((its) => {
      const existing = its.find((i) => i.productId === p.id);
      if (existing) return its.map((i) => i.productId === p.id ? { ...i, quantity: i.quantity + 1, totalPrice: (i.quantity + 1) * i.unitPrice, totalCommission: (i.quantity + 1) * i.commissionPerUnit } : i);
      return [...its, { id: uid("oi_"), productId: p.id, name: p.name, sku: p.sku, unitPrice: p.sellingPrice, commissionPerUnit: p.commission, quantity: 1, totalPrice: p.sellingPrice, totalCommission: p.commission }];
    });
  };
  const updateQty = (id, qty) => setItems((its) => its.map((i) => i.id === id ? { ...i, quantity: qty, totalPrice: qty * i.unitPrice, totalCommission: qty * i.commissionPerUnit } : i));
  const removeItem = (id) => setItems((its) => its.filter((i) => i.id !== id));

  const subtotal = items.reduce((s, i) => s + i.totalPrice, 0);
  const totalQty = items.reduce((s, i) => s + i.quantity, 0);
  const totalCommission = items.reduce((s, i) => s + i.totalCommission, 0);

  const resetForm = () => {
    setCustomer({ name: "", phone: "", whatsapp: "", address: "", notes: "" });
    setItems([]); setStatus("Confirmed"); setPaymentStatus("Unpaid"); setDeliveryDate(addDays(3));
    setWaText(""); setParsed(null);
  };

  const saveOrder = () => {
    if (!customer.name.trim()) { pushToast("Customer name is required", "error"); return; }
    if (items.length === 0) { pushToast("Add at least one bag to the order", "error"); return; }
    let custRecord = customers.find((c) => norm(c.phone) === norm(customer.phone) && customer.phone);
    if (!custRecord) {
      custRecord = { id: uid("c_"), name: customer.name, phone: customer.phone, whatsapp: customer.whatsapp || customer.phone, address: customer.address, notes: customer.notes };
      setCustomers((cs) => [...cs, custRecord]);
    }
    if (editingOrder) {
      const updated = { ...editingOrder, customerId: custRecord.id, customerName: customer.name, phone: customer.phone, address: customer.address, items, subtotal, total: subtotal, totalCommission, status, paymentStatus, deliveryDate };
      setOrders((os) => os.map((o) => o.id === editingOrder.id ? updated : o));
      pushToast("Order " + updated.orderNumber + " updated");
      setLastCreated(updated);
    } else {
      const orderNumber = "ORD-" + (1000 + orders.length + Math.floor(Math.random() * 40));
      const newOrder = { id: uid("o_"), orderNumber, customerId: custRecord.id, customerName: customer.name, phone: customer.phone, address: customer.address, date: todayISO(), status, paymentStatus, items, subtotal, total: subtotal, totalCommission, notes: customer.notes, deliveryDate, reminderDate: null, createdAt: todayISO() };
      setOrders((os) => [newOrder, ...os]);
      if (deliveryDate) setReminders((rs) => [...rs, { id: uid("r_"), orderId: newOrder.id, orderNumber, customerName: customer.name, reminderType: "Delivery", reminderDate: deliveryDate, status: "Pending", notes: "" }]);
      pushToast("Order " + orderNumber + " created");
      setLastCreated(newOrder);
      resetForm();
    }
  };

  const analyzeWa = () => {
    if (!waText.trim()) { pushToast("Paste a WhatsApp order first", "error"); return; }
    setParsed(parseWhatsAppOrder(waText, products));
  };
  const applyParsed = () => {
    setCustomer({ name: parsed.customerName, phone: parsed.phone || "", whatsapp: parsed.phone || "", address: parsed.address || "", notes: [parsed.deliveryNote, parsed.paymentNote, parsed.notes].filter(Boolean).join(" · ") });
    const newItems = parsed.items.filter((i) => i.productId).map((i) => {
      const p = products.find((pp) => pp.id === i.productId);
      return { id: uid("oi_"), productId: p.id, name: p.name, sku: p.sku, unitPrice: p.sellingPrice, commissionPerUnit: p.commission, quantity: i.quantity, totalPrice: p.sellingPrice * i.quantity, totalCommission: p.commission * i.quantity };
    });
    setItems(newItems);
    setTab("manual");
    pushToast("Order details filled in — review and save");
  };
  const fixUnmatched = (parsedItemId, product) => setParsed((p) => ({ ...p, items: p.items.map((i) => i.id === parsedItemId ? { ...i, productId: product.id, matchedName: product.name } : i) }));
  const unmatchedCount = parsed ? parsed.items.filter((i) => !i.productId).length : 0;

  return (
    <div>
      {!editingOrder && (
        <div style={{ display: "flex", gap: 3, marginBottom: 20, background: T.surface2, padding: 3, borderRadius: 9, width: "fit-content" }}>
          {[{ k: "manual", l: "Manual order" }, { k: "wa", l: "Paste WhatsApp order" }].map((t) => (
            <button key={t.k} onClick={() => setTab(t.k)} style={{
              padding: "7px 16px", borderRadius: 7, fontSize: 13, fontWeight: 600, cursor: "pointer", border: "none",
              background: tab === t.k ? T.surface : "transparent", color: tab === t.k ? T.primary : T.textSecondary,
              boxShadow: tab === t.k ? "0 1px 2px rgba(31,41,51,0.08)" : "none", transition: `all 150ms ${EASE}`,
            }}>{t.l}</button>
          ))}
        </div>
      )}

      {tab === "wa" && !editingOrder && (
        <Card style={{ marginBottom: 16, maxWidth: 720 }}>
          {!parsed ? (
            <>
              <div style={{ fontWeight: 600, fontFamily: "Manrope, sans-serif", marginBottom: 4, fontSize: 15 }}>Paste WhatsApp order</div>
              <div style={{ fontSize: 12.5, color: T.textSecondary, marginBottom: 12 }}>Copy the final order message from WhatsApp and paste it below. We'll detect the customer, items and notes automatically.</div>
              <TTextarea rows={9} value={waText} onChange={(e) => setWaText(e.target.value)}
                placeholder={"Ali Traders\n03001234567\n\n2 x Classic Tote Black\n3 x Travel Bag Large\n\nDeliver tomorrow\nPayment cash"} />
              <div style={{ marginTop: 12 }}><Btn icon={ArrowRight} onClick={analyzeWa}>Analyze order</Btn></div>
            </>
          ) : (
            <div>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
                <div style={{ fontWeight: 600, fontFamily: "Manrope, sans-serif", fontSize: 15 }}>Review parsed order</div>
                {unmatchedCount > 0 && <div style={{ fontSize: 12, fontWeight: 600, color: T.warning, background: T.warningBg, padding: "3px 10px", borderRadius: 20 }}>{unmatchedCount} item{unmatchedCount > 1 ? "s" : ""} need{unmatchedCount === 1 ? "s" : ""} review</div>}
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 14 }}>
                <div style={{ display: "flex", gap: 8, alignItems: "center" }}><User size={14} color={T.textTertiary} /><div><div style={{ fontSize: 11, color: T.textSecondary }}>Customer</div><div style={{ fontWeight: 600, fontSize: 13.5 }}>{parsed.customerName}</div></div></div>
                <div style={{ display: "flex", gap: 8, alignItems: "center" }}><Phone size={14} color={T.textTertiary} /><div><div style={{ fontSize: 11, color: T.textSecondary }}>Phone</div><div style={{ fontWeight: 600, fontSize: 13.5 }}>{parsed.phone || "Not found"}</div></div></div>
              </div>
              {(parsed.deliveryNote || parsed.paymentNote) && (
                <div style={{ display: "flex", gap: 18, marginBottom: 14, fontSize: 12.5, color: T.textSecondary }}>
                  {parsed.deliveryNote && <div style={{ display: "flex", gap: 6, alignItems: "center" }}><Truck size={13} />{parsed.deliveryNote}</div>}
                  {parsed.paymentNote && <div style={{ display: "flex", gap: 6, alignItems: "center" }}><CreditCard size={13} />{parsed.paymentNote}</div>}
                </div>
              )}
              <div style={{ fontSize: 11, fontWeight: 600, color: T.textSecondary, marginBottom: 8, textTransform: "uppercase", letterSpacing: "0.04em" }}>Items detected</div>
              {parsed.items.map((it) => (
                <div key={it.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "9px 0", borderBottom: `1px solid ${T.border}` }}>
                  {it.productId ? <CheckCircle2 size={16} color={T.success} /> : <AlertCircle size={16} color={T.warning} />}
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 13, fontWeight: 500 }}>{it.matchedName || it.raw} <span style={{ color: T.textSecondary, fontWeight: 400 }}>× {it.quantity}</span></div>
                    {!it.productId && <div style={{ fontSize: 11.5, color: T.warning }}>Could not match "{it.raw}" — select the correct bag</div>}
                  </div>
                  {!it.productId && (
                    <div style={{ width: 220 }}>
                      <SearchSelect options={products} placeholder="Select correct bag" getLabel={(p) => p.name} getSub={(p) => p.sku} onSelect={(p) => fixUnmatched(it.id, p)} />
                    </div>
                  )}
                </div>
              ))}
              <div style={{ display: "flex", justifyContent: "space-between", marginTop: 16, fontSize: 14, fontWeight: 600 }}>
                <span>Total</span>
                <span>{fmt(parsed.items.filter(i => i.productId).reduce((s, i) => { const p = products.find(pp => pp.id === i.productId); return s + (p ? p.sellingPrice * i.quantity : 0); }, 0))}</span>
              </div>
              <div style={{ display: "flex", gap: 8, marginTop: 18 }}>
                <Btn variant="ghost" onClick={() => setParsed(null)}>Start over</Btn>
                <Btn onClick={applyParsed} disabled={parsed.items.every((i) => !i.productId)}>Use this order</Btn>
              </div>
            </div>
          )}
        </Card>
      )}

      {(tab === "manual" || editingOrder) && (
        <div style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr", gap: 18 }}>
          <div>
            <Card style={{ marginBottom: 16 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
                <div style={{ width: 20, height: 20, borderRadius: 10, background: T.accentBg, color: T.accent, fontSize: 11, fontWeight: 700, display: "flex", alignItems: "center", justifyContent: "center" }}>1</div>
                <div style={{ fontWeight: 600, fontFamily: "Manrope, sans-serif", fontSize: 14 }}>Customer information</div>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                <Field label="Customer name" required><TInput value={customer.name} onChange={(e) => setCustomer({ ...customer, name: e.target.value })} placeholder="e.g. Ali Traders" /></Field>
                <Field label="Phone number"><TInput value={customer.phone} onChange={(e) => setCustomer({ ...customer, phone: e.target.value })} placeholder="03xxxxxxxxx" /></Field>
              </div>
              <Field label="Address"><TInput value={customer.address} onChange={(e) => setCustomer({ ...customer, address: e.target.value })} placeholder="Delivery address" /></Field>
              <Field label="Notes"><TTextarea rows={2} value={customer.notes} onChange={(e) => setCustomer({ ...customer, notes: e.target.value })} placeholder="Optional notes" /></Field>
            </Card>

            <Card>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
                <div style={{ width: 20, height: 20, borderRadius: 10, background: T.accentBg, color: T.accent, fontSize: 11, fontWeight: 700, display: "flex", alignItems: "center", justifyContent: "center" }}>2</div>
                <div style={{ fontWeight: 600, fontFamily: "Manrope, sans-serif", fontSize: 14 }}>Order items</div>
              </div>
              <SearchSelect options={products.filter((p) => p.active !== false)} placeholder="Search bag by name, SKU or category…"
                getLabel={(p) => p.name} getSub={(p) => p.sku + " · " + fmt(p.sellingPrice)} onSelect={addProduct}
                emptyAction={<Btn size="sm" variant="outline" onClick={() => setPage("products")}>Add a new bag</Btn>} />
              <div style={{ marginTop: 14 }}>
                {items.length === 0 ? (
                  <div style={{ fontSize: 12.5, color: T.textSecondary, padding: "18px 0", textAlign: "center" }}>No bags added yet — search above to add one.</div>
                ) : items.map((it) => (
                  <div key={it.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 0", borderBottom: `1px solid ${T.border}` }}>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontSize: 13, fontWeight: 500 }}>{it.name}</div>
                      <div style={{ fontSize: 11.5, color: T.textSecondary }}>{it.sku} · {fmt(it.unitPrice)} each · commission {fmt(it.commissionPerUnit)}/unit</div>
                    </div>
                    <input type="number" min={1} value={it.quantity} onChange={(e) => updateQty(it.id, Math.max(1, parseInt(e.target.value || "1", 10)))}
                      style={{ ...inputBase, width: 62, padding: "6px 8px", textAlign: "center" }} />
                    <div style={{ width: 92, textAlign: "right", fontWeight: 600, fontSize: 13.5 }}>{fmt(it.totalPrice)}</div>
                    <IconBtn icon={Trash2} danger onClick={() => removeItem(it.id)} title="Remove" />
                  </div>
                ))}
              </div>
            </Card>
          </div>

          <div>
            <Card style={{ marginBottom: 16 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
                <div style={{ width: 20, height: 20, borderRadius: 10, background: T.accentBg, color: T.accent, fontSize: 11, fontWeight: 700, display: "flex", alignItems: "center", justifyContent: "center" }}>3</div>
                <div style={{ fontWeight: 600, fontFamily: "Manrope, sans-serif", fontSize: 14 }}>Review pricing</div>
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13.5, marginBottom: 9 }}><span style={{ color: T.textSecondary }}>Subtotal</span><b>{fmt(subtotal)}</b></div>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13.5, marginBottom: 9 }}><span style={{ color: T.textSecondary }}>Total quantity</span><b>{totalQty}</b></div>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13.5, marginBottom: 9, color: T.success }}><span>Your commission</span><b>{fmt(totalCommission)}</b></div>
              <div style={{ height: 1, background: T.border, margin: "12px 0" }} />
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 17, fontWeight: 700 }}><span>Final total</span><span>{fmt(subtotal)}</span></div>
            </Card>
            <Card style={{ marginBottom: 16 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
                <div style={{ width: 20, height: 20, borderRadius: 10, background: T.accentBg, color: T.accent, fontSize: 11, fontWeight: 700, display: "flex", alignItems: "center", justifyContent: "center" }}>4</div>
                <div style={{ fontWeight: 600, fontFamily: "Manrope, sans-serif", fontSize: 14 }}>Delivery &amp; payment</div>
              </div>
              <Field label="Order status"><TSelect value={status} onChange={(e) => setStatus(e.target.value)} options={Object.keys(STATUS_META)} /></Field>
              <Field label="Payment status"><TSelect value={paymentStatus} onChange={(e) => setPaymentStatus(e.target.value)} options={Object.keys(PAYMENT_META)} /></Field>
              <Field label="Delivery date"><TInput type="date" value={deliveryDate} onChange={(e) => setDeliveryDate(e.target.value)} /></Field>
            </Card>
            <Btn size="lg" style={{ width: "100%", justifyContent: "center" }} onClick={saveOrder}>{editingOrder ? "Save changes" : "Confirm & create order"}</Btn>
            {lastCreated && (
              <div style={{ marginTop: 12 }}>
                <Btn variant="outline" icon={Printer} style={{ width: "100%", justifyContent: "center" }} onClick={() => printOrderSlip(lastCreated, settings)}>Print / generate PDF — {lastCreated.orderNumber}</Btn>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

/* ============================== PRODUCTS PAGE ============================== */
function ProductsPage({ products, setProducts, orders, pushToast }) {
  const [q, setQ] = useState("");
  const [sort, setSort] = useState({ key: "name", dir: "asc" });
  const [modalProduct, setModalProduct] = useState(undefined);

  const salesByProduct = useMemo(() => {
    const map = {};
    orders.forEach((o) => o.items.forEach((i) => { map[i.productId] = (map[i.productId] || 0) + i.quantity; }));
    return map;
  }, [orders]);

  let filtered = products.filter((p) => norm(p.name + " " + p.sku + " " + p.category).includes(norm(q)));
  filtered = sortRows(filtered, sort);

  const save = (form) => {
    if (!form.name.trim() || !form.sku.trim()) { pushToast("Bag name and SKU are required", "error"); return; }
    if (modalProduct) {
      setProducts((ps) => ps.map((p) => p.id === modalProduct.id ? { ...p, ...form, updatedAt: todayISO() } : p));
      pushToast("Bag updated");
    } else {
      if (products.some((p) => norm(p.sku) === norm(form.sku))) { pushToast("A bag with this SKU already exists", "error"); return; }
      setProducts((ps) => [...ps, { id: uid("p_"), ...form, active: true, createdAt: todayISO(), updatedAt: todayISO() }]);
      pushToast("Bag added");
    }
    setModalProduct(undefined);
  };
  const toggleActive = (id) => setProducts((ps) => ps.map((p) => p.id === id ? { ...p, active: !p.active } : p));
  const remove = (id) => { setProducts((ps) => ps.filter((p) => p.id !== id)); pushToast("Bag removed"); };

  if (products.length === 0) {
    return <EmptyState icon={Package} title="No products imported yet" body="Add bags manually or import your Excel product sheet to get started." action={<Btn onClick={() => setModalProduct(null)}>Add a bag</Btn>} />;
  }

  return (
    <div>
      <div style={{ display: "flex", gap: 10, marginBottom: 14, alignItems: "center" }}>
        <div style={{ position: "relative", flex: "1 1 260px", maxWidth: 320 }}>
          <Search size={14} style={{ position: "absolute", left: 11, top: 11.5, color: T.textTertiary }} />
          <TInput placeholder="Search name, SKU or category" value={q} onChange={(e) => setQ(e.target.value)} style={{ paddingLeft: 32 }} />
        </div>
        <div style={{ marginLeft: "auto", fontSize: 12, color: T.textSecondary }}>{filtered.length} of {products.length} bags</div>
        <Btn icon={Plus} onClick={() => setModalProduct(null)}>Add bag</Btn>
      </div>
      <Card pad={0}>
        <div style={{ overflowX: "auto" }}>
          <table>
            <thead><tr>
              <SortTh label="Bag name" sortKey="name" sort={sort} setSort={setSort} />
              <SortTh label="SKU" sortKey="sku" sort={sort} setSort={setSort} />
              <th>Category</th>
              <SortTh label="Selling price" sortKey="sellingPrice" sort={sort} setSort={setSort} />
              <SortTh label="Commission" sortKey="commission" sort={sort} setSort={setSort} />
              <th>Units sold</th><th>Status</th><th></th>
            </tr></thead>
            <tbody>
              {filtered.map((p) => (
                <tr key={p.id} className="rowhover">
                  <td style={{ fontWeight: 500 }}>{p.name}</td>
                  <td style={{ color: T.textSecondary }}>{p.sku}</td>
                  <td>{p.category}</td>
                  <td>{fmt(p.sellingPrice)}</td>
                  <td style={{ color: T.success, fontWeight: 600 }}>{fmt(p.commission)}</td>
                  <td>{salesByProduct[p.id] || 0}</td>
                  <td>
                    <span onClick={() => toggleActive(p.id)} style={{ cursor: "pointer" }}>
                      <StatusBadge status={p.active === false ? "Cancelled" : "Delivered"} />
                    </span>
                  </td>
                  <td>
                    <div style={{ display: "flex", gap: 2 }}>
                      <IconBtn icon={Pencil} onClick={() => setModalProduct(p)} title="Edit" />
                      <IconBtn icon={Trash2} danger onClick={() => remove(p.id)} title="Delete" />
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
      <ProductModal open={modalProduct !== undefined} product={modalProduct} onClose={() => setModalProduct(undefined)} onSave={save} />
    </div>
  );
}
function ProductModal({ open, product, onClose, onSave }) {
  const [form, setForm] = useState({ name: "", sku: "", sellingPrice: "", commission: "", category: "Tote", supplier: "", notes: "" });
  useEffect(() => { if (open) setForm(product ? { ...product } : { name: "", sku: "", sellingPrice: "", commission: "", category: "Tote", supplier: "", notes: "" }); }, [open, product]);
  return (
    <Modal open={open} onClose={onClose} title={product ? "Edit bag" : "Add bag"} width={480}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 0 }}>
        <div style={{ gridColumn: "1 / -1" }}><Field label="Bag name" required><TInput value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></Field></div>
        <Field label="SKU / code" required><TInput value={form.sku} onChange={(e) => setForm({ ...form, sku: e.target.value })} /></Field>
        <Field label="Category"><TSelect value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })} options={["Tote", "Clutch", "Sling", "Travel", "Backpack", "Other"]} /></Field>
        <Field label="Selling price (Rs.)"><TInput type="number" value={form.sellingPrice} onChange={(e) => setForm({ ...form, sellingPrice: Number(e.target.value) })} /></Field>
        <Field label="Commission (Rs.)"><TInput type="number" value={form.commission} onChange={(e) => setForm({ ...form, commission: Number(e.target.value) })} /></Field>
        <div style={{ gridColumn: "1 / -1" }}><Field label="Supplier (optional)"><TInput value={form.supplier} onChange={(e) => setForm({ ...form, supplier: e.target.value })} /></Field></div>
        <div style={{ gridColumn: "1 / -1" }}><Field label="Notes (optional)"><TTextarea rows={2} value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} /></Field></div>
      </div>
      <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 8 }}>
        <Btn variant="ghost" onClick={onClose}>Cancel</Btn>
        <Btn onClick={() => onSave(form)}>{product ? "Save changes" : "Add bag"}</Btn>
      </div>
    </Modal>
  );
}

/* ============================== IMPORT EXCEL PAGE ============================== */
function detectCol(headers, candidates) {
  const nh = headers.map((h) => norm(h));
  for (const c of candidates) { const i = nh.findIndex((h) => h === c); if (i >= 0) return i; }
  for (const c of candidates) { const i = nh.findIndex((h) => h.includes(c)); if (i >= 0) return i; }
  return -1;
}
function ImportExcelPage({ products, setProducts, pushToast }) {
  const fileRef = useRef(null);
  const [rows, setRows] = useState(null);
  const [errors, setErrors] = useState([]);
  const [summary, setSummary] = useState(null);
  const [fileName, setFileName] = useState("");
  const [stage, setStage] = useState("upload"); // upload -> preview -> imported
  const [lastImportCount, setLastImportCount] = useState(0);

  const onFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setFileName(file.name);
    setStage("upload");
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        if (data.length < 2) { pushToast("The file appears to be empty", "error"); return; }
        const headers = data[0].map(String);
        const nameIdx = detectCol(headers, ["bag name", "name", "product name", "product"]);
        const skuIdx = detectCol(headers, ["code", "sku", "bag code/sku", "bag code"]);
        const priceIdx = detectCol(headers, ["selling price", "price", "sale price"]);
        const commIdx = detectCol(headers, ["commission", "profit", "my profit/commission", "profit/commission"]);
        const catIdx = detectCol(headers, ["category", "type"]);
        const supIdx = detectCol(headers, ["supplier"]);
        const noteIdx = detectCol(headers, ["notes", "note"]);
        if (nameIdx === -1 || priceIdx === -1) {
          setErrors([{ row: "Header", reason: "Could not find required columns 'Bag Name' and 'Selling Price'. Check your column headers." }]);
          setRows([]); setSummary(null); setStage("preview");
          return;
        }
        const parsedRows = []; const errs = [];
        for (let i = 1; i < data.length; i++) {
          const r = data[i];
          if (!r || r.every((c) => c === "" || c === undefined)) continue;
          const name = String(r[nameIdx] || "").trim();
          const sku = skuIdx >= 0 ? String(r[skuIdx] || "").trim() : "";
          const price = priceIdx >= 0 ? Number(r[priceIdx]) : NaN;
          const commission = commIdx >= 0 ? Number(r[commIdx]) : 0;
          const category = catIdx >= 0 ? String(r[catIdx] || "").trim() : "Other";
          if (!name) { errs.push({ row: i + 1, reason: "Missing bag name" }); continue; }
          if (isNaN(price) || price <= 0) { errs.push({ row: i + 1, reason: "Missing or invalid selling price" }); continue; }
          parsedRows.push({ name, sku: sku || ("AUTO-" + uid()), sellingPrice: price, commission: isNaN(commission) ? 0 : commission, category: category || "Other", supplier: supIdx >= 0 ? String(r[supIdx] || "") : "", notes: noteIdx >= 0 ? String(r[noteIdx] || "") : "", _isUpdate: products.some((p) => norm(p.sku) === norm(sku) && sku) });
        }
        setRows(parsedRows);
        setErrors(errs);
        setSummary({ newCount: parsedRows.filter((r) => !r._isUpdate).length, updateCount: parsedRows.filter((r) => r._isUpdate).length, errorCount: errs.length });
        setStage("preview");
      } catch (err) { pushToast("Could not read this file. Make sure it's a valid .xlsx or .xls file.", "error"); }
    };
    reader.readAsArrayBuffer(file);
  };

  const confirmImport = () => {
    setProducts((ps) => {
      const bySku = new Map(ps.map((p) => [norm(p.sku), p]));
      rows.forEach((r) => {
        const existing = bySku.get(norm(r.sku));
        if (existing) Object.assign(existing, { name: r.name, sellingPrice: r.sellingPrice, commission: r.commission, category: r.category, supplier: r.supplier, notes: r.notes, updatedAt: todayISO() });
        else { const np = { id: uid("p_"), name: r.name, sku: r.sku, sellingPrice: r.sellingPrice, commission: r.commission, category: r.category, supplier: r.supplier, notes: r.notes, active: true, createdAt: todayISO(), updatedAt: todayISO() }; ps = [...ps, np]; bySku.set(norm(r.sku), np); }
      });
      return [...ps];
    });
    pushToast(`Imported ${rows.length} products (${summary.newCount} new, ${summary.updateCount} updated)`);
    setLastImportCount(rows.length);
    setStage("imported");
    if (fileRef.current) fileRef.current.value = "";
  };
  const startOver = () => { setRows(null); setSummary(null); setErrors([]); setFileName(""); setStage("upload"); };

  const steps = ["Upload", "Analyze", "Preview", "Confirm"];
  const stepIndex = stage === "upload" ? (fileName ? 1 : 0) : stage === "preview" ? 2 : 3;

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 18 }}>
        {steps.map((s, i) => (
          <React.Fragment key={s}>
            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <div style={{ width: 20, height: 20, borderRadius: 10, fontSize: 11, fontWeight: 700, display: "flex", alignItems: "center", justifyContent: "center", background: i <= stepIndex ? T.accent : T.surface2, color: i <= stepIndex ? "#fff" : T.textTertiary }}>{i + 1}</div>
              <div style={{ fontSize: 12.5, fontWeight: 500, color: i <= stepIndex ? T.textPrimary : T.textTertiary }}>{s}</div>
            </div>
            {i < steps.length - 1 && <div style={{ width: 24, height: 1, background: T.border }} />}
          </React.Fragment>
        ))}
      </div>

      {stage === "imported" ? (
        <Card style={{ textAlign: "center", padding: "48px 20px" }}>
          <div style={{ width: 44, height: 44, borderRadius: 22, background: T.successBg, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 14px" }}>
            <CheckCircle2 size={22} color={T.success} />
          </div>
          <div style={{ fontFamily: "Manrope, sans-serif", fontWeight: 700, fontSize: 16, marginBottom: 6 }}>Import complete</div>
          <div style={{ fontSize: 13, color: T.textSecondary, marginBottom: 18 }}>{lastImportCount} products were added or updated in your catalogue.</div>
          <Btn variant="outline" onClick={startOver}>Import another file</Btn>
        </Card>
      ) : (
        <>
          <Card style={{ marginBottom: 16 }}>
            <div style={{ fontWeight: 600, fontFamily: "Manrope, sans-serif", fontSize: 14.5, marginBottom: 6 }}>Import bag data from Excel</div>
            <div style={{ fontSize: 12.5, color: T.textSecondary, marginBottom: 16 }}>Upload a .xlsx or .xls file with columns for bag name, code/SKU, selling price and commission. Existing SKUs are updated automatically — new SKUs are added as new bags.</div>
            <div style={{ border: `1.5px dashed ${T.borderStrong}`, borderRadius: RADIUS.card, padding: "28px 20px", textAlign: "center", background: T.surface2 }}>
              <Upload size={20} color={T.textSecondary} style={{ marginBottom: 8 }} />
              <div style={{ fontSize: 13, fontWeight: 500, marginBottom: 12, color: T.textPrimary }}>{fileName || "Select an Excel file to import"}</div>
              <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={onFile} style={{ display: "none" }} id="excel-input" />
              <Btn variant="outline" onClick={() => fileRef.current && fileRef.current.click()}>Choose file</Btn>
            </div>
          </Card>

          {errors.length > 0 && (
            <Card style={{ marginBottom: 16, borderColor: T.error }}>
              <div style={{ fontWeight: 600, color: T.error, fontSize: 13, marginBottom: 8 }}>Import errors ({errors.length})</div>
              {errors.map((e, i) => <div key={i} style={{ fontSize: 12.5, color: T.textSecondary, padding: "3px 0" }}>Row {e.row} — {e.reason}</div>)}
            </Card>
          )}

          {rows && rows.length > 0 && summary && (
            <>
              <Card style={{ marginBottom: 16 }}>
                <div style={{ fontWeight: 600, fontFamily: "Manrope, sans-serif", fontSize: 14.5, marginBottom: 14 }}>Import summary</div>
                <div style={{ display: "flex", gap: 28, flexWrap: "wrap" }}>
                  <div><div style={{ fontSize: 11, color: T.textSecondary }}>New products</div><div style={{ fontWeight: 700, fontSize: 21, color: T.success, fontFamily: "Manrope, sans-serif" }}>{summary.newCount}</div></div>
                  <div><div style={{ fontSize: 11, color: T.textSecondary }}>Updated products</div><div style={{ fontWeight: 700, fontSize: 21, color: T.accent, fontFamily: "Manrope, sans-serif" }}>{summary.updateCount}</div></div>
                  <div><div style={{ fontSize: 11, color: T.textSecondary }}>Errors</div><div style={{ fontWeight: 700, fontSize: 21, color: T.error, fontFamily: "Manrope, sans-serif" }}>{summary.errorCount}</div></div>
                  <div><div style={{ fontSize: 11, color: T.textSecondary }}>Total rows read</div><div style={{ fontWeight: 700, fontSize: 21, fontFamily: "Manrope, sans-serif" }}>{rows.length}</div></div>
                </div>
                <div style={{ display: "flex", gap: 8, marginTop: 18 }}>
                  <Btn onClick={confirmImport}>Confirm import</Btn>
                  <Btn variant="ghost" onClick={startOver}>Cancel</Btn>
                </div>
              </Card>
              <Card pad={0}>
                <div style={{ fontWeight: 600, fontFamily: "Manrope, sans-serif", fontSize: 14, padding: "16px 18px 0" }}>Preview ({rows.length} rows)</div>
                <div style={{ overflowX: "auto", maxHeight: 340, overflowY: "auto" }}>
                  <table>
                    <thead><tr><th>Bag name</th><th>SKU</th><th>Price</th><th>Commission</th><th>Category</th><th>Action</th></tr></thead>
                    <tbody>{rows.map((r, i) => (
                      <tr key={i}><td>{r.name}</td><td>{r.sku}</td><td>{fmt(r.sellingPrice)}</td><td>{fmt(r.commission)}</td><td>{r.category}</td>
                        <td>{r._isUpdate ? <span style={{ color: T.accent, fontWeight: 600, fontSize: 11.5 }}>Update</span> : <span style={{ color: T.success, fontWeight: 600, fontSize: 11.5 }}>New</span>}</td></tr>
                    ))}</tbody>
                  </table>
                </div>
              </Card>
            </>
          )}
        </>
      )}
    </div>
  );
}

/* ============================== REMINDERS PAGE ============================== */
function RemindersPage({ reminders, setReminders, orders }) {
  const markDone = (id) => setReminders((rs) => rs.map((r) => r.id === id ? { ...r, status: "Done" } : r));
  const snooze = (id) => setReminders((rs) => rs.map((r) => r.id === id ? { ...r, reminderDate: addDays(2), status: "Pending" } : r));

  if (reminders.length === 0) {
    return <EmptyState icon={Bell} title="No reminders yet" body="Reminders you add to an order — delivery, payment follow-up, or confirmation — will show up here." />;
  }

  const pending = reminders.filter((r) => r.status === "Pending");
  const dueToday = pending.filter((r) => new Date(r.reminderDate) <= new Date(addDays(0.9999))).sort((a, b) => new Date(a.reminderDate) - new Date(b.reminderDate));
  const upcoming = pending.filter((r) => new Date(r.reminderDate) > new Date(addDays(0.9999))).sort((a, b) => new Date(a.reminderDate) - new Date(b.reminderDate));
  const completed = reminders.filter((r) => r.status === "Done");

  const Group = ({ title, list, tone }) => list.length > 0 && (
    <div style={{ marginBottom: 22 }}>
      <div style={{ fontSize: 11.5, fontWeight: 600, color: T.textSecondary, textTransform: "uppercase", letterSpacing: "0.04em", marginBottom: 10 }}>{title} <span style={{ color: T.textTertiary }}>({list.length})</span></div>
      <Card pad={0}>
        {list.map((r, i) => {
          const overdue = tone === "due" && new Date(r.reminderDate) < new Date(addDays(-0.001));
          return (
            <div key={r.id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "13px 16px", borderBottom: i < list.length - 1 ? `1px solid ${T.border}` : "none" }}>
              <div style={{ width: 7, height: 7, borderRadius: 3.5, background: tone === "due" ? (overdue ? T.error : T.warning) : tone === "done" ? T.success : T.accent, flexShrink: 0 }} />
              <div style={{ flex: 1 }}>
                <div style={{ fontSize: 13, fontWeight: 500 }}>{r.orderNumber} — {r.reminderType}</div>
                <div style={{ fontSize: 11.5, color: T.textSecondary }}>{r.customerName} · {fmtDate(r.reminderDate)}{overdue && <span style={{ color: T.error, fontWeight: 600 }}> · Overdue</span>}</div>
              </div>
              {r.status === "Pending" ? (
                <div style={{ display: "flex", gap: 6 }}>
                  <Btn size="sm" variant="ghost" onClick={() => snooze(r.id)}>Snooze</Btn>
                  <Btn size="sm" variant="outline" icon={Check} onClick={() => markDone(r.id)}>Mark complete</Btn>
                </div>
              ) : <StatusBadge status="Delivered" />}
            </div>
          );
        })}
      </Card>
    </div>
  );

  return (
    <div>
      <Group title="Due today" list={dueToday} tone="due" />
      <Group title="Upcoming" list={upcoming} tone="upcoming" />
      <Group title="Completed" list={completed} tone="done" />
    </div>
  );
}

/* ============================== REPORTS PAGE ============================== */
function ReportsPage({ orders, products }) {
  const [from, setFrom] = useState(addDays(-90));
  const [to, setTo] = useState(todayISO());
  const filtered = orders.filter((o) => o.date >= from && o.date <= to);

  const totalSales = filtered.reduce((s, o) => s + o.total, 0);
  const totalCommission = filtered.reduce((s, o) => s + o.totalCommission, 0);
  const totalBags = filtered.reduce((s, o) => s + o.items.reduce((a, i) => a + i.quantity, 0), 0);
  const avgOrder = filtered.length ? totalSales / filtered.length : 0;

  const topProducts = useMemo(() => {
    const map = {};
    filtered.forEach((o) => o.items.forEach((i) => {
      if (!map[i.name]) map[i.name] = { name: i.name, sku: i.sku, qty: 0, sales: 0, commission: 0 };
      map[i.name].qty += i.quantity; map[i.name].sales += i.totalPrice; map[i.name].commission += i.totalCommission;
    }));
    return Object.values(map).sort((a, b) => b.sales - a.sales).slice(0, 10);
  }, [filtered]);

  const exportExcel = () => {
    const wsData = filtered.map((o) => ({ "Order #": o.orderNumber, Date: o.date, Customer: o.customerName, Phone: o.phone, "Total Items": o.items.reduce((s, i) => s + i.quantity, 0), "Sale Amount": o.total, Commission: o.totalCommission, Status: o.status, Payment: o.paymentStatus }));
    const ws = XLSX.utils.json_to_sheet(wsData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    XLSX.writeFile(wb, `sales-report-${from}-to-${to}.xlsx`);
  };

  return (
    <div>
      <Card pad={14} style={{ marginBottom: 16 }}>
        <div style={{ display: "flex", gap: 12, alignItems: "flex-end", flexWrap: "wrap" }}>
          <Field label="From"><TInput type="date" value={from} onChange={(e) => setFrom(e.target.value)} style={{ width: 160 }} /></Field>
          <Field label="To"><TInput type="date" value={to} onChange={(e) => setTo(e.target.value)} style={{ width: 160 }} /></Field>
          <Btn variant="outline" icon={Download} onClick={exportExcel} style={{ marginBottom: 12 }}>Export to Excel</Btn>
          <div style={{ marginLeft: "auto", fontSize: 12, color: T.textSecondary, marginBottom: 12 }}>{filtered.length} orders in range</div>
        </div>
      </Card>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14, marginBottom: 16 }}>
        <KpiCard label="Total sales" value={fmt(totalSales)} />
        <KpiCard label="Total commission" value={fmt(totalCommission)} valueColor={T.success} />
        <KpiCard label="Bags sold" value={totalBags} />
        <KpiCard label="Orders" value={filtered.length} />
        <KpiCard label="Avg order value" value={fmt(avgOrder)} />
      </div>

      <Card pad={0}>
        <div style={{ fontWeight: 600, fontFamily: "Manrope, sans-serif", fontSize: 14, padding: "16px 18px 0" }}>Top products</div>
        <div style={{ overflowX: "auto" }}>
          <table>
            <thead><tr><th>Bag</th><th>SKU</th><th>Units sold</th><th>Sales</th><th>Commission</th></tr></thead>
            <tbody>{topProducts.map((p) => (
              <tr key={p.name}><td style={{ fontWeight: 500 }}>{p.name}</td><td style={{ color: T.textSecondary }}>{p.sku}</td><td>{p.qty}</td><td>{fmt(p.sales)}</td><td style={{ color: T.success, fontWeight: 600 }}>{fmt(p.commission)}</td></tr>
            ))}</tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

/* ============================== SETTINGS PAGE ============================== */
function SettingsSection({ icon: Icon, title, description, children }) {
  return (
    <Card style={{ marginBottom: 16 }}>
      <div style={{ display: "flex", gap: 10, marginBottom: 16 }}>
        <div style={{ width: 32, height: 32, borderRadius: 8, background: T.surface2, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}><Icon size={15} color={T.textSecondary} /></div>
        <div>
          <div style={{ fontWeight: 600, fontFamily: "Manrope, sans-serif", fontSize: 14.5 }}>{title}</div>
          {description && <div style={{ fontSize: 12, color: T.textSecondary, marginTop: 2 }}>{description}</div>}
        </div>
      </div>
      {children}
    </Card>
  );
}
function SettingsPage({ settings, setSettings, pushToast, products, orders, customers, setProducts, setOrders, setCustomers, setReminders }) {
  const [form, setForm] = useState(settings);
  useEffect(() => setForm(settings), [settings]);
  const save = () => { setSettings(form); pushToast("Settings saved"); };

  const exportAll = () => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(products.map(({ id, ...p }) => p)), "Products");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(orders.map((o) => ({ orderNumber: o.orderNumber, customer: o.customerName, date: o.date, total: o.total, commission: o.totalCommission, status: o.status, payment: o.paymentStatus }))), "Orders");
    XLSX.writeFile(wb, "bagcrm-full-export.xlsx");
    pushToast("Data exported");
  };

  return (
    <div style={{ maxWidth: 520 }}>
      <SettingsSection icon={Building2} title="Business" description="Appears on the header of every printed order slip.">
        <Field label="Business name"><TInput value={form.businessName} onChange={(e) => setForm({ ...form, businessName: e.target.value })} /></Field>
        <Field label="Phone"><TInput value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} /></Field>
        <Field label="Address"><TInput value={form.address} onChange={(e) => setForm({ ...form, address: e.target.value })} /></Field>
        <Btn onClick={save}>Save changes</Btn>
      </SettingsSection>

      <SettingsSection icon={SlidersHorizontal} title="Preferences" description="Fixed for this workspace.">
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "6px 0", borderBottom: `1px solid ${T.border}` }}><span style={{ color: T.textSecondary }}>Currency</span><span style={{ fontWeight: 600 }}>Rs. (PKR)</span></div>
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "10px 0" }}><span style={{ color: T.textSecondary }}>Commission basis</span><span style={{ fontWeight: 600 }}>Per bag, set in Products</span></div>
      </SettingsSection>

      <SettingsSection icon={Database} title="Data" description="Your products, orders and customers are saved automatically.">
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "6px 0", borderBottom: `1px solid ${T.border}` }}><span style={{ color: T.textSecondary }}>Products</span><span style={{ fontWeight: 600 }}>{products.length}</span></div>
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "6px 0", borderBottom: `1px solid ${T.border}` }}><span style={{ color: T.textSecondary }}>Orders</span><span style={{ fontWeight: 600 }}>{orders.length}</span></div>
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "10px 0 16px" }}><span style={{ color: T.textSecondary }}>Customers</span><span style={{ fontWeight: 600 }}>{customers.length}</span></div>
        <Btn variant="outline" icon={Download} onClick={exportAll}>Export all data to Excel</Btn>
      </SettingsSection>
    </div>
  );
}
