import { useState, useEffect, useRef } from "react";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, AreaChart, Area } from "recharts";
import { save } from "@tauri-apps/plugin-dialog";
import { writeTextFile } from "@tauri-apps/plugin-fs";

const GOLD = "#C9A84C";
const GOLD_LIGHT = "#E2C57A";
const GOLD_DIM = "#8A6D2E";

const todayDate = () => new Date().toISOString().split("T")[0];
const rupees = (value) => `₹${(Number(value) || 0).toLocaleString("en-IN")}`;
const escapeHtml = (value) => String(value ?? "").replace(/[&<>"']/g, ch => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch]));

function buildExcelHtml(title, rows, totals = {}) {
  const headers = ["S.No.", "Date/Time", "Customer", "Product", "Category", "Size", "Color", "Qty", "Cost Price", "Sell Price", "Profit", "Product Total", "Bill Total"];
  const bodyRows = rows.map(r => `
    <tr>
      <td>${escapeHtml(r.serial_no)}</td>
      <td>${escapeHtml(r.created_at)}</td>
      <td>${escapeHtml(r.customer_name)}</td>
      <td>${escapeHtml(r.product_name)}</td>
      <td>${escapeHtml(r.category)}</td>
      <td>${escapeHtml(r.size)}</td>
      <td>${escapeHtml(r.color)}</td>
      <td>${escapeHtml(r.qty)}</td>
      <td>${Number(r.cost || 0).toFixed(2)}</td>
      <td>${Number(r.price || 0).toFixed(2)}</td>
      <td>${Number(r.profit || 0).toFixed(2)}</td>
      <td>${Number(r.line_total || 0).toFixed(2)}</td>
      <td>${Number(r.sale_total || 0).toFixed(2)}</td>
    </tr>`).join("");
  const profit = rows.reduce((sum, r) => sum + (Number(r.profit) || 0), 0);
  const html = `
    <html>
      <head><meta charset="utf-8" /></head>
      <body>
        <h2>${escapeHtml(title)}</h2>
        <p>Total sales: ${Number(totals.sales || 0).toFixed(2)}</p>
        <p>Total product subtotals: ${Number(totals.subtotal || 0).toFixed(2)}</p>
        <p>Total profit: ${profit.toFixed(2)}</p>
        <table border="1">
          <thead><tr>${headers.map(h => `<th>${escapeHtml(h)}</th>`).join("")}</tr></thead>
          <tbody>${bodyRows}</tbody>
        </table>
      </body>
    </html>`;
  return html;
}

function buildReceiptHtml(receipt, includeButton = true) {
  const receiptText = [
    "Aura Fits Receipt",
    `Receipt #${receipt.saleId}`,
    `Customer: ${receipt.customerName || "Walk-in Customer"}`,
    `Date: ${receipt.date}`,
    `Payment: ${receipt.payment}`,
    ...receipt.items.map(item => `${item.name} x${item.qty} - ${rupees(item.lineTotal)}`),
    `Total: ${rupees(receipt.total)}`,
  ].join("\n");
  const shareText = encodeURIComponent(receiptText);
  const mailSubject = encodeURIComponent(`Aura Fits Receipt #${receipt.saleId}`);
  const itemRows = receipt.items.map(item => `
    <tr>
      <td>${escapeHtml(item.name)}</td>
      <td style="text-align:center">${item.qty}</td>
      <td style="text-align:right">${rupees(item.unitPrice)}</td>
      <td style="text-align:right">${rupees(item.lineTotal)}</td>
    </tr>`).join("");
  return `
    <html>
      <head>
        <title>Aura Fits Receipt</title>
        <style>
          body { font-family: Arial, sans-serif; color: #111; padding: 18px; }
          h1 { text-align: center; font-size: 22px; margin: 0; letter-spacing: 4px; }
          .muted { color: #666; font-size: 12px; }
          table { width: 100%; border-collapse: collapse; margin-top: 14px; font-size: 12px; }
          th, td { border-bottom: 1px solid #ddd; padding: 7px 4px; }
          th { text-align: left; }
          .line { display: flex; justify-content: space-between; margin-top: 8px; font-size: 13px; }
          .total { font-size: 18px; font-weight: 700; border-top: 1px solid #111; padding-top: 10px; }
          .share { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 8px; margin-top: 16px; }
          .share a, button { border: 1px solid #111; border-radius: 6px; color: #111; display: block; font-size: 12px; padding: 9px; text-align: center; text-decoration: none; }
          @media print { button, .share { display: none; } body { padding: 0; } }
        </style>
      </head>
      <body>
        <h1>AURA FITS</h1>
        <p class="muted" style="text-align:center">Receipt #${escapeHtml(receipt.saleId)} · ${escapeHtml(receipt.date)}</p>
        <p class="muted">Customer: ${escapeHtml(receipt.customerName || "Walk-in Customer")}</p>
        <p class="muted">Payment: ${escapeHtml(receipt.payment)}</p>
        <table>
          <thead><tr><th>Product</th><th>Qty</th><th style="text-align:right">Rate</th><th style="text-align:right">Subtotal</th></tr></thead>
          <tbody>${itemRows}</tbody>
        </table>
        <div class="line"><span>Subtotal</span><span>${rupees(receipt.subtotal)}</span></div>
        <div class="line"><span>Discount (${receipt.discount}%)</span><span>-${rupees(receipt.discountAmt)}</span></div>
        <div class="line total"><span>Total</span><span>${rupees(receipt.total)}</span></div>
        <p class="muted" style="text-align:center;margin-top:24px">Thank you for shopping with us.</p>
        <div class="share">
          <a href="https://wa.me/?text=${shareText}" target="_blank">WhatsApp</a>
          <a href="mailto:?subject=${mailSubject}&body=${shareText}">Email</a>
          <a href="sms:?body=${shareText}">SMS</a>
        </div>
        ${includeButton ? '<button onclick="window.print()" style="width:100%;padding:10px;margin-top:16px">Print Receipt</button>' : ''}
      </body>
    </html>
  `;
}

function printReceipt(receipt) {
  const printWindow = window.open("", "_blank", "width=420,height=650");
  if (!printWindow) return;
  printWindow.document.write(buildReceiptHtml(receipt));
  printWindow.document.close();
  printWindow.focus();
}

const globalStyles = `
  @import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@300;400;600&family=DM+Sans:wght@300;400;500;600&display=swap');
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: #0A0A0A; color: #E8E4D9; font-family: 'DM Sans', sans-serif; }
  ::-webkit-scrollbar { width: 4px; }
  ::-webkit-scrollbar-track { background: #111; }
  ::-webkit-scrollbar-thumb { background: #333; border-radius: 2px; }
  input, select, textarea {
    background: #1a1a1a; border: 1px solid #2a2a2a; color: #E8E4D9;
    border-radius: 8px; padding: 10px 14px; font-family: 'DM Sans', sans-serif;
    font-size: 14px; outline: none; transition: border-color 0.2s; width: 100%;
  }
  input:focus, select:focus, textarea:focus { border-color: ${GOLD}; }
  select option { background: #1a1a1a; }
  .toast-enter { animation: toastIn 0.3s ease; }
  @keyframes toastIn { from { opacity: 0; transform: translateX(100%); } to { opacity: 1; transform: translateX(0); } }
  .fade-in { animation: fadeIn 0.25s ease; }
  @keyframes fadeIn { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
`;

// ─── SHARED UI COMPONENTS ────────────────────────────────────────────────────

function Toast({ toasts }) {
  return (
    <div style={{ position: "fixed", top: 20, right: 20, zIndex: 9999, display: "flex", flexDirection: "column", gap: 8 }}>
      {toasts.map(t => (
        <div key={t.id} className="toast-enter" style={{
          background: t.type === "success" ? "#1a2d1a" : t.type === "error" ? "#2d1a1a" : "#1a1d2d",
          border: `1px solid ${t.type === "success" ? "#2d5c2d" : t.type === "error" ? "#5c2d2d" : "#2d3d5c"}`,
          color: "#E8E4D9", padding: "12px 18px", borderRadius: 10, fontSize: 13, maxWidth: 300,
          display: "flex", alignItems: "center", gap: 8
        }}>
          <span style={{ color: t.type === "success" ? "#5fa05f" : t.type === "error" ? "#c05f5f" : GOLD }}>
            {t.type === "success" ? "✓" : t.type === "error" ? "✕" : "ℹ"}
          </span>
          {t.message}
        </div>
      ))}
    </div>
  );
}

function Card({ children, style = {}, onClick }) {
  return (
    <div onClick={onClick} style={{
      background: "#141414", border: "1px solid #222", borderRadius: 14,
      padding: 20, transition: "border-color 0.2s", cursor: onClick ? "pointer" : "default", ...style
    }}
      onMouseEnter={e => onClick && (e.currentTarget.style.borderColor = "#333")}
      onMouseLeave={e => onClick && (e.currentTarget.style.borderColor = "#222")}
    >{children}</div>
  );
}

function GoldButton({ children, onClick, style = {}, variant = "primary", size = "md" }) {
  return (
    <button style={{
      background: variant === "primary" ? `linear-gradient(135deg, ${GOLD} 0%, ${GOLD_LIGHT} 50%, ${GOLD} 100%)` : "transparent",
      border: `1px solid ${GOLD}`, color: variant === "primary" ? "#0A0A0A" : GOLD,
      borderRadius: 8, cursor: "pointer", fontFamily: "'DM Sans', sans-serif",
      fontWeight: 600, fontSize: size === "sm" ? 12 : 14,
      padding: size === "sm" ? "6px 14px" : "10px 22px",
      transition: "all 0.2s", letterSpacing: "0.02em", ...style
    }} onClick={onClick}
      onMouseEnter={e => { e.currentTarget.style.opacity = "0.85"; }}
      onMouseLeave={e => { e.currentTarget.style.opacity = "1"; }}
    >{children}</button>
  );
}

function StatCard({ label, value, sub, color = GOLD, icon }) {
  return (
    <Card>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <div>
          <p style={{ fontSize: 11, color: "#666", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 6 }}>{label}</p>
          <p style={{ fontSize: 24, fontWeight: 600, color, fontFamily: "'Cormorant Garamond', serif", letterSpacing: "0.02em" }}>{value}</p>
          {sub && <p style={{ fontSize: 12, color: "#555", marginTop: 4 }}>{sub}</p>}
        </div>
        {icon && <div style={{ fontSize: 20, opacity: 0.4 }}>{icon}</div>}
      </div>
    </Card>
  );
}

function Badge({ children, color = GOLD }) {
  return (
    <span style={{
      background: `${color}18`, border: `1px solid ${color}40`, color,
      borderRadius: 6, padding: "2px 8px", fontSize: 11, fontWeight: 500
    }}>{children}</span>
  );
}

function Modal({ title, onClose, children, width = 520 }) {
  return (
    <div style={{
      position: "fixed", inset: 0, background: "rgba(0,0,0,0.8)", zIndex: 1000,
      display: "flex", alignItems: "center", justifyContent: "center"
    }} onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="fade-in" style={{
        background: "#141414", border: "1px solid #2a2a2a", borderRadius: 16,
        width, maxWidth: "95vw", maxHeight: "90vh", overflow: "auto", padding: 28
      }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 22 }}>
          <h3 style={{ fontFamily: "'Cormorant Garamond', serif", fontSize: 20, color: GOLD, fontWeight: 400, letterSpacing: "0.05em" }}>{title}</h3>
          <button onClick={onClose} style={{ background: "none", border: "none", color: "#666", cursor: "pointer", fontSize: 20 }}>×</button>
        </div>
        {children}
      </div>
    </div>
  );
}

function FormField({ label, children }) {
  return (
    <div style={{ marginBottom: 16 }}>
      <label style={{ display: "block", fontSize: 12, color: "#888", marginBottom: 6, letterSpacing: "0.05em", textTransform: "uppercase" }}>{label}</label>
      {children}
    </div>
  );
}

function Table({ headers, rows, onEdit, onDelete }) {
  return (
    <div style={{ overflowX: "auto" }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
        <thead>
          <tr>
            {headers.map(h => (
              <th key={h} style={{ textAlign: "left", padding: "10px 14px", color: "#555", fontWeight: 500, fontSize: 11, letterSpacing: "0.08em", textTransform: "uppercase", borderBottom: "1px solid #1e1e1e" }}>{h}</th>
            ))}
            {(onEdit || onDelete) && <th style={{ padding: "10px 14px", borderBottom: "1px solid #1e1e1e" }}></th>}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, i) => (
            <tr key={i} style={{ borderBottom: "1px solid #161616" }}
              onMouseEnter={e => e.currentTarget.style.background = "#181818"}
              onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
              {row.map((cell, j) => (
                <td key={j} style={{ padding: "12px 14px", color: "#C8C4B8" }}>{cell}</td>
              ))}
              {(onEdit || onDelete) && (
                <td style={{ padding: "12px 14px", textAlign: "right" }}>
                  <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                    {onEdit && <button onClick={() => onEdit(i)} style={{ background: "none", border: "1px solid #2a2a2a", color: "#888", borderRadius: 6, padding: "4px 10px", cursor: "pointer", fontSize: 11 }}>Edit</button>}
                    {onDelete && <button onClick={() => onDelete(i)} style={{ background: "none", border: "1px solid #3d1a1a", color: "#c05f5f", borderRadius: 6, padding: "4px 10px", cursor: "pointer", fontSize: 11 }}>Delete</button>}
                  </div>
                </td>
              )}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// ─── LOGIN ───────────────────────────────────────────────────────────────────
const SECURITY_QUESTIONS = [
  "What was the name of your first pet?",
  "What city were you born in?",
  "What is your mother's maiden name?",
  "What was the name of your primary school?",
  "What was your childhood nickname?",
];

function LoginPage({ onLogin }) {
  const [screen, setScreen] = useState("login");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [showPass, setShowPass] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  const [signupName, setSignupName] = useState("");
  const [signupEmail, setSignupEmail] = useState("");
  const [signupPass, setSignupPass] = useState("");
  const [signupPass2, setSignupPass2] = useState("");
  const [signupQ, setSignupQ] = useState(SECURITY_QUESTIONS[0]);
  const [signupAnswer, setSignupAnswer] = useState("");

  const [forgotEmail, setForgotEmail] = useState("");
  const [forgotQ, setForgotQ] = useState("");
  const [forgotAnswer, setForgotAnswer] = useState("");
  const [forgotUser, setForgotUser] = useState(null);
  const [newPass, setNewPass] = useState("");
  const [newPass2, setNewPass2] = useState("");

  const handleLogin = async () => {
    if (!email || !password) { setError("Please fill in all fields"); return; }
    setLoading(true); setError("");
    const result = await window.db.login({ email, password });
    if (result.ok) { onLogin(result.name); }
    else { setError("Invalid email or password"); setLoading(false); }
  };

  const handleSignup = async () => {
    setError("");
    if (!signupName || !signupEmail || !signupPass || !signupAnswer) { setError("Please fill in all fields"); return; }
    if (signupPass !== signupPass2) { setError("Passwords do not match"); return; }
    if (signupPass.length < 6) { setError("Password must be at least 6 characters"); return; }
    const result = await window.db.signup({ name: signupName, email: signupEmail, password: signupPass, securityQuestion: signupQ, securityAnswer: signupAnswer });
    if (result.ok) { setScreen("login"); setEmail(signupEmail); setPassword(""); setError(""); }
    else { setError(result.error || "Signup failed"); }
  };

  const handleForgotLookup = async () => {
    setError("");
    const user = await window.db.findUser(forgotEmail);
    if (!user) { setError("No account found with that email"); return; }
    setForgotUser({ ...user, email: forgotEmail });
    setForgotQ(user.securityQuestion);
  };

  const handleForgotVerify = () => {
    setError("");
    if (forgotAnswer.toLowerCase().trim() !== forgotUser.securityAnswer) { setError("Incorrect answer. Please try again."); return; }
    setScreen("resetPass");
  };

  const handleResetPass = async () => {
    setError("");
    if (!newPass || !newPass2) { setError("Please fill in both fields"); return; }
    if (newPass !== newPass2) { setError("Passwords do not match"); return; }
    if (newPass.length < 6) { setError("Password must be at least 6 characters"); return; }
    await window.db.resetPassword({ email: forgotUser.email, newPassword: newPass });
    setScreen("login"); setEmail(forgotUser.email); setPassword(""); setError("");
    setForgotEmail(""); setForgotAnswer(""); setForgotUser(null);
  };

  const Logo = () => (
    <div style={{ textAlign: "center", marginBottom: 36 }}>
      <div style={{ display: "inline-flex", alignItems: "center", justifyContent: "center", width: 64, height: 64, background: `${GOLD}18`, border: `1px solid ${GOLD}40`, borderRadius: 16, marginBottom: 16 }}>
        <svg width="36" height="36" viewBox="0 0 36 36" fill="none">
          <path d="M18 4 L32 28 L4 28 Z" stroke={GOLD} strokeWidth="2.2" fill="none" strokeLinejoin="round"/>
          <path d="M18 12 L26 26 L10 26 Z" fill={`${GOLD}30`} strokeLinejoin="round"/>
          <circle cx="18" cy="19" r="2.5" fill={GOLD}/>
        </svg>
      </div>
      <h1 style={{ fontFamily: "'Cormorant Garamond', serif", fontSize: 30, color: "#E8E4D9", fontWeight: 300, letterSpacing: "0.25em" }}>AURA FITS</h1>
      <p style={{ color: "#444", fontSize: 11, letterSpacing: "0.2em", textTransform: "uppercase", marginTop: 4 }}>Management Console</p>
    </div>
  );

  return (
    <div style={{ minHeight: "100vh", background: "#0A0A0A", display: "flex", alignItems: "center", justifyContent: "center", position: "relative", overflow: "hidden" }}>
      <div style={{ position: "absolute", inset: 0, background: `radial-gradient(ellipse at 20% 50%, ${GOLD}08 0%, transparent 60%), radial-gradient(ellipse at 80% 20%, ${GOLD}05 0%, transparent 50%)` }} />
      <div className="fade-in" style={{ width: screen === "signup" ? 480 : 420, position: "relative", zIndex: 1 }}>
        <Logo />

        {screen === "login" && (
          <div style={{ background: "#111", border: "1px solid #1e1e1e", borderRadius: 20, padding: 36 }}>
            <h2 style={{ fontSize: 16, color: "#888", fontWeight: 400, marginBottom: 28 }}>Welcome back</h2>
            <FormField label="Email"><input type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder="you@aura.fits" /></FormField>
            <FormField label="Password">
              <div style={{ position: "relative" }}>
                <input type={showPass ? "text" : "password"} value={password} onChange={e => setPassword(e.target.value)} placeholder="••••••••" onKeyDown={e => e.key === "Enter" && handleLogin()} />
                <button onClick={() => setShowPass(!showPass)} style={{ position: "absolute", right: 12, top: "50%", transform: "translateY(-50%)", background: "none", border: "none", color: "#555", cursor: "pointer", fontSize: 13 }}>{showPass ? "Hide" : "Show"}</button>
              </div>
            </FormField>
            <div style={{ textAlign: "right", marginBottom: 16 }}>
              <button onClick={() => { setScreen("forgot"); setError(""); setForgotEmail(""); setForgotUser(null); }} style={{ background: "none", border: "none", color: GOLD_DIM, fontSize: 12, cursor: "pointer" }}>Forgot password?</button>
            </div>
            {error && <p style={{ fontSize: 12, color: "#c05f5f", marginBottom: 14, textAlign: "center" }}>{error}</p>}
            <GoldButton onClick={handleLogin} style={{ width: "100%", padding: "13px" }}>{loading ? "Authenticating..." : "Sign In"}</GoldButton>
            <div style={{ textAlign: "center", marginTop: 20 }}>
              <span style={{ fontSize: 13, color: "#444" }}>Don't have an account? </span>
              <button onClick={() => { setScreen("signup"); setError(""); }} style={{ background: "none", border: "none", color: GOLD, fontSize: 13, cursor: "pointer", fontWeight: 500 }}>Create Account</button>
            </div>
          </div>
        )}

        {screen === "signup" && (
          <div style={{ background: "#111", border: "1px solid #1e1e1e", borderRadius: 20, padding: 36 }}>
            <h2 style={{ fontSize: 16, color: "#888", fontWeight: 400, marginBottom: 24 }}>Create a new account</h2>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
              <FormField label="Full Name"><input value={signupName} onChange={e => setSignupName(e.target.value)} placeholder="Your Name" /></FormField>
              <FormField label="Email"><input type="email" value={signupEmail} onChange={e => setSignupEmail(e.target.value)} placeholder="you@email.com" /></FormField>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
              <FormField label="Password"><input type="password" value={signupPass} onChange={e => setSignupPass(e.target.value)} placeholder="Min. 6 characters" /></FormField>
              <FormField label="Confirm Password"><input type="password" value={signupPass2} onChange={e => setSignupPass2(e.target.value)} placeholder="Repeat password" /></FormField>
            </div>
            <div style={{ borderTop: "1px solid #1e1e1e", paddingTop: 18, marginTop: 4 }}>
              <p style={{ fontSize: 11, color: GOLD_DIM, letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 14 }}>Security Question</p>
              <FormField label="Choose a Question">
                <select value={signupQ} onChange={e => setSignupQ(e.target.value)}>
                  {SECURITY_QUESTIONS.map(q => <option key={q} value={q}>{q}</option>)}
                </select>
              </FormField>
              <FormField label="Your Answer"><input value={signupAnswer} onChange={e => setSignupAnswer(e.target.value)} placeholder="Answer (case-insensitive)" /></FormField>
            </div>
            {error && <p style={{ fontSize: 12, color: "#c05f5f", marginBottom: 12, textAlign: "center" }}>{error}</p>}
            <GoldButton onClick={handleSignup} style={{ width: "100%", padding: "13px", marginTop: 4 }}>Create Account</GoldButton>
            <div style={{ textAlign: "center", marginTop: 16 }}>
              <button onClick={() => { setScreen("login"); setError(""); }} style={{ background: "none", border: "none", color: "#555", fontSize: 13, cursor: "pointer" }}>← Back to Sign In</button>
            </div>
          </div>
        )}

        {screen === "forgot" && (
          <div style={{ background: "#111", border: "1px solid #1e1e1e", borderRadius: 20, padding: 36 }}>
            <h2 style={{ fontSize: 16, color: "#888", fontWeight: 400, marginBottom: 24 }}>Reset your password</h2>
            <FormField label="Enter your email">
              <input type="email" value={forgotEmail} onChange={e => { setForgotEmail(e.target.value); setForgotUser(null); setError(""); }} placeholder="you@email.com" />
            </FormField>
            {!forgotUser && (
              <>
                {error && <p style={{ fontSize: 12, color: "#c05f5f", marginBottom: 12, textAlign: "center" }}>{error}</p>}
                <GoldButton onClick={handleForgotLookup} style={{ width: "100%", padding: "13px" }}>Find My Account</GoldButton>
              </>
            )}
            {forgotUser && (
              <>
                <div style={{ background: "#0d0d0d", border: "1px solid #1e1e1e", borderRadius: 10, padding: "14px 16px", marginBottom: 16 }}>
                  <p style={{ fontSize: 11, color: "#555", marginBottom: 6, textTransform: "uppercase" }}>Security Question</p>
                  <p style={{ fontSize: 13, color: "#C8C4B8" }}>{forgotQ}</p>
                </div>
                <FormField label="Your Answer">
                  <input value={forgotAnswer} onChange={e => { setForgotAnswer(e.target.value); setError(""); }} placeholder="Type your answer" />
                </FormField>
                {error && <p style={{ fontSize: 12, color: "#c05f5f", marginBottom: 12, textAlign: "center" }}>{error}</p>}
                <GoldButton onClick={handleForgotVerify} style={{ width: "100%", padding: "13px" }}>Verify Answer</GoldButton>
              </>
            )}
            <div style={{ textAlign: "center", marginTop: 16 }}>
              <button onClick={() => { setScreen("login"); setError(""); }} style={{ background: "none", border: "none", color: "#555", fontSize: 13, cursor: "pointer" }}>← Back to Sign In</button>
            </div>
          </div>
        )}

        {screen === "resetPass" && (
          <div style={{ background: "#111", border: "1px solid #1e1e1e", borderRadius: 20, padding: 36 }}>
            <h2 style={{ fontSize: 16, color: "#888", fontWeight: 400, marginBottom: 24 }}>Set a new password</h2>
            <FormField label="New Password"><input type="password" value={newPass} onChange={e => setNewPass(e.target.value)} placeholder="Min. 6 characters" /></FormField>
            <FormField label="Confirm New Password"><input type="password" value={newPass2} onChange={e => setNewPass2(e.target.value)} placeholder="Repeat password" /></FormField>
            {error && <p style={{ fontSize: 12, color: "#c05f5f", marginBottom: 12, textAlign: "center" }}>{error}</p>}
            <GoldButton onClick={handleResetPass} style={{ width: "100%", padding: "13px" }}>Update Password</GoldButton>
          </div>
        )}

        <p style={{ textAlign: "center", color: "#333", fontSize: 11, marginTop: 24, letterSpacing: "0.08em" }}>AURA FITS · PRIVATE MANAGEMENT SYSTEM</p>
      </div>
    </div>
  );
}

// ─── SIDEBAR ─────────────────────────────────────────────────────────────────
const NAV_ITEMS = [
  { id: "dashboard", label: "Dashboard", icon: "◈" },
  { id: "pos", label: "New Sale", icon: "⊕" },
  { id: "inventory", label: "Inventory", icon: "◫" },
  { id: "categories", label: "Categories", icon: "⊞" },
  { id: "reports", label: "Reports & Insights", icon: "▦" },
  { id: "statement", label: "Statement", icon: "≡" },
  { id: "expenses", label: "Expenses", icon: "◉" },
  { id: "settings", label: "Settings", icon: "⊛" },
];

function Sidebar({ active, setActive, collapsed, setCollapsed, userName }) {
  return (
    <div style={{
      width: collapsed ? 64 : 220, background: "#0D0D0D", borderRight: "1px solid #1a1a1a",
      height: "100vh", position: "fixed", top: 0, left: 0, transition: "width 0.25s ease",
      display: "flex", flexDirection: "column", zIndex: 100, overflow: "hidden"
    }}>
      <div style={{ padding: "20px 0", borderBottom: "1px solid #1a1a1a", display: "flex", alignItems: "center", justifyContent: collapsed ? "center" : "space-between", paddingLeft: collapsed ? 0 : 20, paddingRight: collapsed ? 0 : 16 }}>
        {!collapsed && (
          <div>
            <p style={{ fontFamily: "'Cormorant Garamond', serif", fontSize: 16, color: GOLD, letterSpacing: "0.1em" }}>AURA</p>
            <p style={{ fontSize: 9, color: "#444", letterSpacing: "0.12em", textTransform: "uppercase" }}>Fits</p>
          </div>
        )}
        <button onClick={() => setCollapsed(!collapsed)} style={{ background: "none", border: "none", color: "#555", cursor: "pointer", fontSize: 18, width: 32, height: 32, display: "flex", alignItems: "center", justifyContent: "center", borderRadius: 6 }}
          onMouseEnter={e => e.currentTarget.style.color = GOLD}
          onMouseLeave={e => e.currentTarget.style.color = "#555"}
        >{collapsed ? "›" : "‹"}</button>
      </div>
      <nav style={{ flex: 1, paddingTop: 12, overflow: "auto" }}>
        {NAV_ITEMS.map(item => {
          const isActive = active === item.id;
          return (
            <div key={item.id} onClick={() => setActive(item.id)} title={collapsed ? item.label : ""} style={{
              display: "flex", alignItems: "center", gap: 12,
              padding: collapsed ? "12px 0" : "12px 20px",
              justifyContent: collapsed ? "center" : "flex-start",
              cursor: "pointer", transition: "all 0.15s",
              background: isActive ? `${GOLD}10` : "transparent",
              borderLeft: isActive ? `2px solid ${GOLD}` : "2px solid transparent", marginBottom: 2
            }}
              onMouseEnter={e => !isActive && (e.currentTarget.style.background = "#141414")}
              onMouseLeave={e => !isActive && (e.currentTarget.style.background = "transparent")}
            >
              <span style={{ fontSize: 16, color: isActive ? GOLD : "#444", minWidth: 20, textAlign: "center" }}>{item.icon}</span>
              {!collapsed && <span style={{ fontSize: 13, color: isActive ? GOLD : "#777", fontWeight: isActive ? 500 : 400, whiteSpace: "nowrap" }}>{item.label}</span>}
            </div>
          );
        })}
      </nav>
      <div style={{ padding: "16px 20px", borderTop: "1px solid #1a1a1a" }}>
        {!collapsed && (
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ width: 28, height: 28, borderRadius: "50%", background: `${GOLD}20`, border: `1px solid ${GOLD}40`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11, color: GOLD, fontWeight: 600 }}>{(userName || "A")[0].toUpperCase()}</div>
            <div>
              <p style={{ fontSize: 12, color: "#C8C4B8" }}>{userName || "Admin"}</p>
              <p style={{ fontSize: 10, color: "#444" }}>Owner</p>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── DASHBOARD ───────────────────────────────────────────────────────────────
function Dashboard() {
  const [stats, setStats] = useState({ today: { total: 0, count: 0 }, month: { total: 0 }, weekly: [] });
  const [products, setProducts] = useState([]);
  const [expenses, setExpenses] = useState([]);
  const tooltipStyle = { background: "#141414", border: "1px solid #222", borderRadius: 8, color: "#E8E4D9", fontSize: 12 };

  useEffect(() => {
    window.db.getSalesStats().then(setStats);
    window.db.getProducts().then(setProducts);
    window.db.getExpenses().then(setExpenses);
  }, []);

  const lowStock = products.filter(p => p.stock <= 5);
  const monthExpenses = expenses.reduce((s, e) => s + e.amount, 0);
  const todayStr = new Date().toLocaleDateString("en-IN", { weekday: "long", day: "numeric", month: "long", year: "numeric" });

  // Build weekly chart data (Sun=0..Sat=6)
  const dayNames = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  const weeklyChart = dayNames.map((day, i) => {
    const found = (stats.weekly || []).find(w => Number(w.dow) === i);
    return { day, sales: found ? found.sales : 0 };
  });

  return (
    <div className="fade-in">
      <div style={{ marginBottom: 28 }}>
        <h1 style={{ fontFamily: "'Cormorant Garamond', serif", fontSize: 28, fontWeight: 300, color: "#E8E4D9" }}>Dashboard</h1>
        <p style={{ color: "#555", fontSize: 13, marginTop: 4 }}>{todayStr}</p>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: 14, marginBottom: 24 }}>
        <StatCard label="Today's Sales" value={`₹${(stats.today?.total || 0).toLocaleString()}`} sub={`${stats.today?.count || 0} transactions`} icon="◈" />
        <StatCard label="Monthly Revenue" value={`₹${(stats.month?.total || 0).toLocaleString()}`} sub={new Date().toLocaleString("default", { month: "long", year: "numeric" })} color="#E8E4D9" icon="▦" />
        <StatCard label="Total Products" value={products.length} sub="in inventory" color="#E8E4D9" icon="◫" />
        <StatCard label="Monthly Expenses" value={`₹${monthExpenses.toLocaleString()}`} color="#c08060" icon="◉" />
        <StatCard label="Low Stock Items" value={lowStock.length} sub="Needs restock" color="#c05f5f" icon="⚠" />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
        <Card>
          <p style={{ fontSize: 11, color: "#555", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 16 }}>Weekly Sales (Last 7 Days)</p>
          <ResponsiveContainer width="100%" height={180}>
            <BarChart data={weeklyChart}>
              <CartesianGrid strokeDasharray="3 3" stroke="#1e1e1e" />
              <XAxis dataKey="day" tick={{ fill: "#555", fontSize: 11 }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fill: "#555", fontSize: 11 }} axisLine={false} tickLine={false} tickFormatter={v => `₹${(v / 1000).toFixed(0)}k`} />
              <Tooltip contentStyle={tooltipStyle} formatter={v => [`₹${v.toLocaleString()}`, ""]} />
              <Bar dataKey="sales" fill={GOLD} radius={[4, 4, 0, 0]} opacity={0.85} />
            </BarChart>
          </ResponsiveContainer>
        </Card>

        <Card>
          <p style={{ fontSize: 11, color: "#555", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 16 }}>⚠ Low Stock Alert</p>
          {products.length === 0
            ? <p style={{ color: "#c08060", fontSize: 13, padding: "20px 0" }}>No products in inventory yet.</p>
            : lowStock.length === 0
            ? <p style={{ color: "#5fa05f", fontSize: 13, padding: "20px 0" }}>All products well stocked ✓</p>
            : <Table
                headers={["Product", "Brand", "Stock"]}
                rows={lowStock.map(p => [
                  p.name, p.brand,
                  <span style={{ color: p.stock <= 3 ? "#c05f5f" : "#c08060", fontWeight: 500 }}>Only {p.stock} left</span>
                ])}
              />
          }
        </Card>
      </div>

      <Card>
        <p style={{ fontSize: 11, color: "#555", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 16 }}>Inventory Overview</p>
        <Table
          headers={["Product", "Category", "Brand", "Cost", "Stock"]}
          rows={products.slice(0, 8).map(p => [
            <span style={{ color: "#E8E4D9", fontWeight: 500 }}>{p.name}</span>,
            <Badge>{p.category}</Badge>,
            p.brand,
            <span style={{ color: GOLD }}>₹{p.cost.toLocaleString()}</span>,
            <span style={{ color: p.stock <= 3 ? "#c05f5f" : p.stock <= 7 ? "#c08060" : "#5fa05f", fontWeight: 500 }}>{p.stock}</span>
          ])}
        />
      </Card>
    </div>
  );
}

// ─── POS ─────────────────────────────────────────────────────────────────────
function POSPage({ addToast }) {
  const [products, setProducts] = useState([]);
  const [categories, setCategories] = useState([]);
  const [search, setSearch] = useState("");
  const [catFilter, setCatFilter] = useState("All");
  const [customerName, setCustomerName] = useState("");
  const [cart, setCart] = useState([]);
  const [discount, setDiscount] = useState(0);
  const [payment, setPayment] = useState("Cash");
  const [lastReceipt, setLastReceipt] = useState(null);
  const searchRef = useRef(null);

  useEffect(() => {
    window.db.getProducts().then(setProducts);
    window.db.getCategories().then(setCategories);
    searchRef.current?.focus();
  }, []);

  const filtered = products.filter(p =>
    (catFilter === "All" || p.category === catFilter) &&
    p.name.toLowerCase().includes(search.toLowerCase())
  );

  const addItem = (product) => {
    setCart(prev => {
      const ex = prev.find(i => i.id === product.id);
      if (ex) return prev.map(i => i.id === product.id ? { ...i, qty: i.qty + 1 } : i);
      return [...prev, { ...product, qty: 1, sellPrice: "" }];
    });
  };

  const updateQty = (id, delta) => {
    setCart(prev => prev.map(i => i.id === id ? { ...i, qty: Math.max(0, i.qty + delta) } : i).filter(i => i.qty > 0));
  };

  const updateSellPrice = (id, val) => {
    setCart(prev => prev.map(i => i.id === id ? { ...i, sellPrice: val } : i));
  };

  const subtotal = cart.reduce((s, i) => s + (Number(i.sellPrice) || 0) * i.qty, 0);
  const discountAmt = Math.round(subtotal * discount / 100);
  const total = subtotal - discountAmt;

  const checkout = async () => {
    if (!cart.length) { addToast("Cart is empty", "error"); return; }
    if (cart.some(i => !Number(i.sellPrice))) { addToast("Enter sale price for every product", "error"); return; }
    const cartToSave = cart.map(i => ({ ...i, price: Number(i.sellPrice) || 0 }));
    const receiptCustomer = customerName.trim() || "Walk-in Customer";
    const result = await window.db.recordSale({ cart: cartToSave, total, discount, payment, customerName: receiptCustomer });
    if (result.ok) {
      setLastReceipt({
        saleId: result.saleId,
        date: new Date().toLocaleString("en-IN"),
        customerName: receiptCustomer,
        payment,
        items: cart.map(i => {
          const unitPrice = Number(i.sellPrice) || 0;
          return { name: i.name, qty: i.qty, unitPrice, lineTotal: unitPrice * i.qty };
        }),
        subtotal,
        discount,
        discountAmt,
        total,
      });
      addToast(`Sale of ₹${total.toLocaleString()} recorded via ${payment}`, "success");
      setCart([]); setDiscount(0); setCustomerName("");
      // Refresh products to show updated stock
      window.db.getProducts().then(setProducts);
    } else {
      addToast("Failed to record sale", "error");
    }
  };

  return (
    <div className="fade-in" style={{ display: "grid", gridTemplateColumns: "1fr 420px", gap: 16, height: "calc(100vh - 40px)" }}>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <input ref={searchRef} value={search} onChange={e => setSearch(e.target.value)} placeholder="Search products..." style={{ height: 46, maxWidth: 520, flex: "0 0 auto" }} />
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {["All", ...categories.map(c => c.name)].map(cat => (
            <button key={cat} onClick={() => setCatFilter(cat)} style={{
              background: catFilter === cat ? `${GOLD}20` : "transparent",
              border: `1px solid ${catFilter === cat ? GOLD : "#222"}`,
              color: catFilter === cat ? GOLD : "#555",
              borderRadius: 20, padding: "4px 14px", cursor: "pointer", fontSize: 12
            }}>{cat}</button>
          ))}
        </div>
        <div style={{ flex: 1, overflowY: "auto", display: "flex", flexDirection: "column", gap: 4 }}>
          {filtered.map(p => (
            <div key={p.id} onClick={() => addItem(p)} style={{
              background: "#111", border: "1px solid #1e1e1e", borderRadius: 10, padding: "12px 16px",
              cursor: "pointer", display: "flex", alignItems: "center", gap: 12
            }}
              onMouseEnter={e => { e.currentTarget.style.borderColor = GOLD + "50"; e.currentTarget.style.background = "#141414"; }}
              onMouseLeave={e => { e.currentTarget.style.borderColor = "#1e1e1e"; e.currentTarget.style.background = "#111"; }}
            >
              <div style={{ flex: 1 }}>
                <p style={{ fontSize: 13, color: "#E8E4D9", marginBottom: 2 }}>{p.name}</p>
                <p style={{ fontSize: 11, color: "#555" }}>{p.brand} · {p.category} · {p.size} · {p.color}</p>
              </div>
              <div style={{ textAlign: "right" }}>
                <p style={{ fontSize: 14, color: GOLD, fontWeight: 600 }}>Cost ₹{p.cost.toLocaleString()}</p>
                <p style={{ fontSize: 11, color: p.stock <= 3 ? "#c05f5f" : "#555" }}>{p.stock} in stock</p>
              </div>
            </div>
          ))}
        </div>
      </div>

      <div style={{ display: "flex", flexDirection: "column", background: "#111", borderRadius: 14, border: "1px solid #1e1e1e", overflow: "hidden" }}>
        <div style={{ padding: "16px 20px", borderBottom: "1px solid #1e1e1e" }}>
          <p style={{ fontFamily: "'Cormorant Garamond', serif", fontSize: 18, color: GOLD }}>Current Sale</p>
          <p style={{ fontSize: 11, color: "#555", marginTop: 2 }}>{cart.length} item{cart.length !== 1 ? "s" : ""}</p>
          <input value={customerName} onChange={e => setCustomerName(e.target.value)} placeholder="Customer name" style={{ marginTop: 12, height: 40 }} />
        </div>
        <div style={{ flex: 1, overflowY: "auto", padding: "12px 16px" }}>
          {cart.length === 0
            ? <div style={{ textAlign: "center", padding: "40px 0", color: "#444" }}><p style={{ fontSize: 24, marginBottom: 8 }}>◫</p><p style={{ fontSize: 13 }}>Add products to start sale</p></div>
            : cart.map(item => (
              <div key={item.id} style={{ display: "grid", gridTemplateColumns: "1fr 92px 118px", alignItems: "center", gap: 10, marginBottom: 12, padding: "12px 0", borderBottom: "1px solid #1a1a1a" }}>
                <div style={{ flex: 1 }}>
                  <p style={{ fontSize: 12, color: "#C8C4B8" }}>{item.name}</p>
                  <p style={{ fontSize: 11, color: "#555" }}>Cost ₹{item.cost.toLocaleString()} each</p>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <button onClick={() => updateQty(item.id, -1)} style={{ width: 26, height: 26, background: "#1a1a1a", border: "1px solid #2a2a2a", color: "#888", borderRadius: 6, cursor: "pointer", fontSize: 16 }}>−</button>
                  <span style={{ fontSize: 13, color: "#E8E4D9", minWidth: 20, textAlign: "center" }}>{item.qty}</span>
                  <button onClick={() => updateQty(item.id, 1)} style={{ width: 26, height: 26, background: "#1a1a1a", border: "1px solid #2a2a2a", color: "#888", borderRadius: 6, cursor: "pointer", fontSize: 16 }}>+</button>
                </div>
                <div style={{ display: "flex", flexDirection: "column", alignItems: "stretch", gap: 4 }}>
                  <span style={{ fontSize: 10, color: "#666", textTransform: "uppercase", letterSpacing: "0.08em" }}>Sell Price</span>
                  <input type="number" value={item.sellPrice} onChange={e => updateSellPrice(item.id, e.target.value)} placeholder="Sale price"
                    style={{ width: "100%", padding: "6px 8px", textAlign: "right", fontSize: 12, color: GOLD, background: "#1a1a1a", border: "1px solid #333", borderRadius: 6 }} />
                  <span style={{ fontSize: 11, color: "#555" }}>× {item.qty} = ₹{((Number(item.sellPrice) || 0) * item.qty).toLocaleString()}</span>
                </div>
              </div>
            ))
          }
        </div>
        <div style={{ padding: "16px 20px", borderTop: "1px solid #1e1e1e" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
            <span style={{ fontSize: 12, color: "#666", flex: 1 }}>Discount %</span>
            <input type="number" value={discount} onChange={e => setDiscount(Math.max(0, Number(e.target.value)))} style={{ width: 80, padding: "6px 10px", textAlign: "right" }} min={0} placeholder="0" />
          </div>
          {discount > 0 && (
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
              <span style={{ fontSize: 12, color: "#c05f5f" }}>Discount</span>
              <span style={{ fontSize: 13, color: "#c05f5f" }}>−₹{Math.round(subtotal * discount / 100).toLocaleString()}</span>
            </div>
          )}
          <div style={{ display: "flex", justifyContent: "space-between", padding: "12px 0", borderTop: "1px solid #1e1e1e" }}>
            <span style={{ fontSize: 15, color: "#E8E4D9", fontWeight: 500 }}>Total</span>
            <span style={{ fontSize: 20, color: GOLD, fontWeight: 700, fontFamily: "'Cormorant Garamond', serif" }}>₹{total.toLocaleString()}</span>
          </div>
          <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
            {["Cash", "UPI", "Card"].map(m => (
              <button key={m} onClick={() => setPayment(m)} style={{
                flex: 1, padding: "8px 0", background: payment === m ? `${GOLD}20` : "transparent",
                border: `1px solid ${payment === m ? GOLD : "#2a2a2a"}`,
                color: payment === m ? GOLD : "#666", borderRadius: 8, cursor: "pointer", fontSize: 12
              }}>{m}</button>
            ))}
          </div>
          <GoldButton onClick={checkout} style={{ width: "100%", padding: "14px", fontSize: 15 }}>✓ Checkout — ₹{total.toLocaleString()}</GoldButton>
          {cart.length > 0 && (
            <button onClick={() => setCart([])} style={{ width: "100%", background: "none", border: "none", color: "#444", cursor: "pointer", fontSize: 12, marginTop: 10, padding: "6px" }}>Clear Cart</button>
          )}
        </div>
      </div>
      {lastReceipt && (
        <Modal title={`Receipt #${lastReceipt.saleId}`} onClose={() => setLastReceipt(null)} width={460}>
          <div style={{ background: "#0d0d0d", border: "1px solid #1e1e1e", borderRadius: 10, padding: 14, marginBottom: 16 }}>
            <p style={{ fontSize: 12, color: "#666", marginBottom: 4 }}>{lastReceipt.date} · {lastReceipt.payment}</p>
            {lastReceipt.items.map((item, i) => (
              <div key={i} style={{ display: "grid", gridTemplateColumns: "1fr 48px 90px", gap: 10, padding: "8px 0", borderBottom: i === lastReceipt.items.length - 1 ? "none" : "1px solid #1a1a1a" }}>
                <span style={{ color: "#C8C4B8", fontSize: 13 }}>{item.name}</span>
                <span style={{ color: "#777", fontSize: 13, textAlign: "center" }}>x{item.qty}</span>
                <span style={{ color: GOLD, fontSize: 13, textAlign: "right" }}>{rupees(item.lineTotal)}</span>
              </div>
            ))}
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, color: "#888", marginBottom: 8 }}><span>Subtotal</span><span>{rupees(lastReceipt.subtotal)}</span></div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, color: "#c05f5f", marginBottom: 8 }}><span>Discount</span><span>-{rupees(lastReceipt.discountAmt)}</span></div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 18, color: GOLD, marginBottom: 18, borderTop: "1px solid #222", paddingTop: 12 }}><span>Total</span><span>{rupees(lastReceipt.total)}</span></div>
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
            <GoldButton variant="outline" onClick={() => setLastReceipt(null)}>Close</GoldButton>
            <GoldButton variant="outline" onClick={async () => {
              const filePath = await save({
                defaultPath: `aura-fits-receipt-${lastReceipt.saleId}.html`,
                filters: [{ name: "Receipt", extensions: ["html"] }],
              });
              if (!filePath) return;
              await writeTextFile(filePath, buildReceiptHtml(lastReceipt, false));
              addToast("Receipt saved", "success");
            }}>Save Receipt</GoldButton>
            <GoldButton onClick={() => printReceipt(lastReceipt)}>Print Receipt</GoldButton>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ─── INVENTORY ───────────────────────────────────────────────────────────────
function InventoryPage({ addToast }) {
  const [products, setProducts] = useState([]);
  const [categories, setCategories] = useState([]);
  const [search, setSearch] = useState("");
  const [catFilter, setCatFilter] = useState("All");
  const [showModal, setShowModal] = useState(false);
  const [editProduct, setEditProduct] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [form, setForm] = useState({ name: "", category: "T-Shirts", brand: "", size: "", color: "", cost: "", stock: "" });

  const load = () => {
    window.db.getProducts().then(setProducts);
    window.db.getCategories().then(setCategories);
  };
  useEffect(() => {
    load();
  }, []);

  const filtered = products.filter(p =>
    (catFilter === "All" || p.category === catFilter) &&
    (p.name.toLowerCase().includes(search.toLowerCase()) || p.brand.toLowerCase().includes(search.toLowerCase()))
  );

  const openAdd = () => { setForm({ name: "", category: categories[0]?.name || "T-Shirts", brand: "", size: "", color: "", cost: "", stock: "" }); setEditProduct(null); setShowModal(true); };
  const openEdit = (i) => { setForm({ ...filtered[i] }); setEditProduct(filtered[i]); setShowModal(true); };

  const save = async () => {
    if (!form.name || !form.cost || !form.stock) { addToast("Please fill required fields", "error"); return; }
    const p = { ...form, cost: Number(form.cost), price: Number(form.cost), stock: Number(form.stock) };
    if (editProduct) {
      await window.db.updateProduct({ ...p, id: editProduct.id });
      addToast("Product updated", "success");
    } else {
      await window.db.addProduct(p);
      addToast("Product added", "success");
    }
    setShowModal(false); load();
  };

  const doDelete = async (i) => {
    await window.db.deleteProduct(filtered[i].id);
    setConfirmDelete(null); addToast("Product deleted", "info"); load();
  };

  return (
    <div className="fade-in">
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24 }}>
        <div>
          <h1 style={{ fontFamily: "'Cormorant Garamond', serif", fontSize: 28, fontWeight: 300, color: "#E8E4D9" }}>Inventory</h1>
          <p style={{ color: "#555", fontSize: 13, marginTop: 4 }}>{products.length} products</p>
        </div>
        <GoldButton onClick={openAdd}>+ Add Product</GoldButton>
      </div>

      <Card style={{ marginBottom: 16 }}>
        <div style={{ display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
          <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search by name or brand..." style={{ maxWidth: 280 }} />
          <select value={catFilter} onChange={e => setCatFilter(e.target.value)} style={{ maxWidth: 180 }}>
            <option>All</option>
            {categories.map(c => <option key={c.id}>{c.name}</option>)}
          </select>
          <span style={{ fontSize: 12, color: "#555" }}>{filtered.length} results</span>
        </div>
      </Card>

      <Card>
        <Table
          headers={["Product Name", "Category", "Brand", "Size", "Color", "Cost", "Stock"]}
          rows={filtered.map(p => [
            <span style={{ color: "#E8E4D9", fontWeight: 500 }}>{p.name}</span>,
            <Badge>{p.category}</Badge>,
            p.brand, p.size, p.color,
            `₹${p.cost.toLocaleString()}`,
            <span style={{ color: p.stock <= 3 ? "#c05f5f" : p.stock <= 7 ? "#c08060" : "#5fa05f", fontWeight: 500 }}>
              {p.stock <= 5 ? `⚠ Only ${p.stock}` : p.stock}
            </span>
          ])}
          onEdit={openEdit}
          onDelete={i => setConfirmDelete(i)}
        />
      </Card>

      {showModal && (
        <Modal title={editProduct ? "Edit Product" : "Add New Product"} onClose={() => setShowModal(false)}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", columnGap: 12, rowGap: 14, marginBottom: 22 }}>
            {[
              ["name","Product Name","text","1 / -1"],
              ["category","Category","select","1 / 2"],
              ["brand","Brand","text","2 / 3"],
              ["size","Size","text","1 / 2"],
              ["color","Color","text","2 / 3"],
              ["cost","Cost Price (₹)","number","1 / 2"],
              ["stock","Stock Qty","number","2 / 3"],
            ].map(([key, lbl, type, span]) => (
              <div key={key} style={{ gridColumn: span }}>
                <FormField label={lbl}>
                  {type === "select"
                    ? <select value={form[key]} onChange={e => setForm(f => ({ ...f, [key]: e.target.value }))}>{categories.map(c => <option key={c.id}>{c.name}</option>)}</select>
                    : <input type={type} value={form[key]} onChange={e => setForm(f => ({ ...f, [key]: e.target.value }))} />
                  }
                </FormField>
              </div>
            ))}
          </div>
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
            <GoldButton variant="outline" onClick={() => setShowModal(false)}>Cancel</GoldButton>
            <GoldButton onClick={save}>{editProduct ? "Save Changes" : "Add Product"}</GoldButton>
          </div>
        </Modal>
      )}

      {confirmDelete !== null && (
        <Modal title="Confirm Delete" onClose={() => setConfirmDelete(null)} width={380}>
          <p style={{ fontSize: 13, color: "#888", marginBottom: 24 }}>Delete <strong style={{ color: "#E8E4D9" }}>{filtered[confirmDelete]?.name}</strong>? This cannot be undone.</p>
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
            <GoldButton variant="outline" onClick={() => setConfirmDelete(null)}>Cancel</GoldButton>
            <button onClick={() => doDelete(confirmDelete)} style={{ background: "#2d1a1a", border: "1px solid #5c2d2d", color: "#c05f5f", borderRadius: 8, padding: "10px 22px", cursor: "pointer", fontSize: 14 }}>Delete</button>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ─── CATEGORIES ──────────────────────────────────────────────────────────────
function CategoriesPage({ addToast }) {
  const [cats, setCats] = useState([]);
  const [showModal, setShowModal] = useState(false);
  const [editCat, setEditCat] = useState(null);
  const [form, setForm] = useState({ name: "", icon: "👔" });

  const load = () => window.db.getCategories().then(setCats);
  useEffect(() => {
    load();
  }, []);

  const openAdd = () => { setForm({ name: "", icon: "👔" }); setEditCat(null); setShowModal(true); };
  const openEdit = (c) => { setForm({ name: c.name, icon: c.icon }); setEditCat(c); setShowModal(true); };

  const save = async () => {
    if (!form.name) { addToast("Category name required", "error"); return; }
    if (editCat) { await window.db.updateCategory({ ...form, id: editCat.id }); addToast("Category updated", "success"); }
    else { await window.db.addCategory(form); addToast("Category added", "success"); }
    setShowModal(false); load();
  };

  const doDelete = async (id) => {
    await window.db.deleteCategory(id); addToast("Deleted", "info"); load();
  };

  return (
    <div className="fade-in">
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 28 }}>
        <div>
          <h1 style={{ fontFamily: "'Cormorant Garamond', serif", fontSize: 28, fontWeight: 300, color: "#E8E4D9" }}>Categories</h1>
          <p style={{ color: "#555", fontSize: 13, marginTop: 4 }}>{cats.length} categories</p>
        </div>
        <GoldButton onClick={openAdd}>+ Add Category</GoldButton>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14 }}>
        {cats.map(c => (
          <Card key={c.id}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 12 }}>
              <span style={{ fontSize: 28 }}>{c.icon}</span>
              <div style={{ display: "flex", gap: 6 }}>
                <GoldButton variant="outline" size="sm" onClick={() => openEdit(c)}>Edit</GoldButton>
                <button onClick={() => doDelete(c.id)} style={{ background: "none", border: "1px solid #3d1a1a", color: "#c05f5f", borderRadius: 6, padding: "4px 10px", cursor: "pointer", fontSize: 11 }}>Del</button>
              </div>
            </div>
            <p style={{ fontSize: 15, color: "#E8E4D9", marginBottom: 4 }}>{c.name}</p>
          </Card>
        ))}
      </div>

      {showModal && (
        <Modal title={editCat ? "Edit Category" : "Add Category"} onClose={() => setShowModal(false)} width={380}>
          <FormField label="Category Name"><input value={form.name} onChange={e => setForm(f => ({ ...f, name: e.target.value }))} placeholder="e.g. Blazers" /></FormField>
          <FormField label="Emoji Icon"><input value={form.icon} onChange={e => setForm(f => ({ ...f, icon: e.target.value }))} placeholder="👔" /></FormField>
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
            <GoldButton variant="outline" onClick={() => setShowModal(false)}>Cancel</GoldButton>
            <GoldButton onClick={save}>{editCat ? "Save Changes" : "Add"}</GoldButton>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ─── REPORTS ─────────────────────────────────────────────────────────────────
function ReportsPage() {
  const [todaySales, setTodaySales] = useState([]);
  const [stats, setStats] = useState({ today: { total: 0, count: 0 }, month: { total: 0 }, weekly: [] });
  const [expenses, setExpenses] = useState([]);
  const tooltipStyle = { background: "#141414", border: "1px solid #222", borderRadius: 8, color: "#E8E4D9", fontSize: 12 };

  useEffect(() => {
    window.db.getTodaySales().then(setTodaySales);
    window.db.getSalesStats().then(setStats);
    window.db.getExpenses().then(setExpenses);
  }, []);

  const monthExpenses = expenses.reduce((s, e) => s + e.amount, 0);
  const profit = (stats.month?.total || 0) - monthExpenses;

  const dayNames = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  const weeklyChart = dayNames.map((day, i) => {
    const found = (stats.weekly || []).find(w => Number(w.dow) === i);
    return { day, sales: found ? found.sales : 0 };
  });

  return (
    <div className="fade-in">
      <div style={{ marginBottom: 28 }}>
        <h1 style={{ fontFamily: "'Cormorant Garamond', serif", fontSize: 28, fontWeight: 300, color: "#E8E4D9" }}>Reports & Insights</h1>
        <p style={{ color: "#555", fontSize: 13, marginTop: 4 }}>Live data from your database</p>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14, marginBottom: 20 }}>
        <StatCard label="Today's Revenue" value={`₹${(stats.today?.total || 0).toLocaleString()}`} sub={`${stats.today?.count || 0} sales`} />
        <StatCard label="Monthly Revenue" value={`₹${(stats.month?.total || 0).toLocaleString()}`} color="#E8E4D9" />
        <StatCard label="Monthly Expenses" value={`₹${monthExpenses.toLocaleString()}`} color="#c08060" />
        <StatCard label="Est. Profit" value={`₹${profit.toLocaleString()}`} color={profit >= 0 ? "#5fa05f" : "#c05f5f"} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
        <Card>
          <p style={{ fontSize: 11, color: "#555", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 16 }}>Weekly Sales</p>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={weeklyChart}>
              <CartesianGrid strokeDasharray="3 3" stroke="#1e1e1e" />
              <XAxis dataKey="day" tick={{ fill: "#555", fontSize: 11 }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fill: "#555", fontSize: 11 }} axisLine={false} tickLine={false} tickFormatter={v => `₹${(v / 1000).toFixed(0)}k`} />
              <Tooltip contentStyle={tooltipStyle} formatter={v => [`₹${v.toLocaleString()}`, ""]} />
              <Bar dataKey="sales" fill={GOLD} radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </Card>

        <Card>
          <p style={{ fontSize: 11, color: "#555", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 14 }}>Today's Sales Log</p>
          {todaySales.length === 0
            ? <p style={{ color: "#555", fontSize: 13, padding: "20px 0" }}>No sales recorded today yet.</p>
            : <Table
                headers={["Time", "Payment", "Items", "Total"]}
                rows={todaySales.map(s => [
                  s.created_at?.slice(11, 16) || "—",
                  <Badge color={s.payment === "Cash" ? "#5fa05f" : s.payment === "UPI" ? "#5f8fa0" : GOLD_DIM}>{s.payment}</Badge>,
                  <span style={{ color: "#888" }}>{s.items_summary || "—"}</span>,
                  <span style={{ color: GOLD, fontWeight: 500 }}>₹{s.total.toLocaleString()}</span>
                ])}
              />
          }
        </Card>
      </div>
    </div>
  );
}

// ─── STATEMENT ───────────────────────────────────────────────────────────────
function StatementPage({ addToast }) {
  const [fromDate, setFromDate] = useState(todayDate());
  const [toDate, setToDate] = useState(todayDate());
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(false);

  const load = async () => {
    if (!fromDate || !toDate) { addToast("Choose both dates", "error"); return; }
    if (fromDate > toDate) { addToast("From date cannot be after To date", "error"); return; }
    setLoading(true);
    try {
      setRows(await window.db.getStatement(fromDate, toDate));
    } catch (error) {
      addToast(error?.message || "Failed to load statement", "error");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    load();
  }, []);

  const saleIds = new Set(rows.map(r => r.sale_id));
  const totalSales = [...saleIds].reduce((sum, id) => {
    const sale = rows.find(r => r.sale_id === id);
    return sum + (sale?.sale_total || 0);
  }, 0);
  const productSubtotals = rows.reduce((sum, r) => sum + (r.line_total || 0), 0);
  const totalProfit = rows.reduce((sum, r) => sum + (r.profit || 0), 0);

  const exportExcel = async () => {
    if (!rows.length) { addToast("No statement data to export", "error"); return; }
    const filePath = await save({
      defaultPath: `aura-fits-statement-${fromDate}-to-${toDate}.xls`,
      filters: [{ name: "Excel Workbook", extensions: ["xls"] }],
    });
    if (!filePath) return;
    const html = buildExcelHtml(`Aura Fits Statement (${fromDate} to ${toDate})`, rows, { sales: totalSales, subtotal: productSubtotals });
    await writeTextFile(filePath, html);
    addToast("Statement exported for Excel", "success");
  };

  return (
    <div className="fade-in">
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24 }}>
        <div>
          <h1 style={{ fontFamily: "'Cormorant Garamond', serif", fontSize: 28, fontWeight: 300, color: "#E8E4D9" }}>Statement</h1>
          <p style={{ color: "#555", fontSize: 13, marginTop: 4 }}>Sales export by date range</p>
        </div>
        <GoldButton onClick={exportExcel}>Export Excel</GoldButton>
      </div>

      <Card style={{ marginBottom: 16 }}>
        <div style={{ display: "grid", gridTemplateColumns: "180px 180px auto", gap: 12, alignItems: "end" }}>
          <FormField label="From Date"><input type="date" value={fromDate} onChange={e => setFromDate(e.target.value)} /></FormField>
          <FormField label="To Date"><input type="date" value={toDate} onChange={e => setToDate(e.target.value)} /></FormField>
          <GoldButton onClick={load} style={{ marginBottom: 16, width: 120 }}>{loading ? "Loading..." : "Apply"}</GoldButton>
        </div>
      </Card>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 14, marginBottom: 16 }}>
        <StatCard label="Bills" value={saleIds.size} sub="in selected range" />
        <StatCard label="Sales Total" value={rupees(totalSales)} color="#E8E4D9" />
        <StatCard label="Profit" value={rupees(totalProfit)} color={totalProfit >= 0 ? "#5fa05f" : "#c05f5f"} />
      </div>

      <Card>
        {rows.length === 0
          ? <p style={{ color: "#555", fontSize: 13, padding: "20px 0" }}>No sales found for this date range.</p>
          : <Table
              headers={["S.No.", "Date/Time", "Customer", "Product", "Category", "Size", "Color", "Cost", "Sell", "Profit", "Total"]}
              rows={rows.map(r => [
                r.serial_no,
                r.created_at?.slice(0, 16) || "",
                r.customer_name || "Walk-in Customer",
                r.product_name,
                r.category,
                r.size,
                r.color,
                rupees(r.cost),
                rupees(r.price),
                <span style={{ color: r.profit >= 0 ? "#5fa05f" : "#c05f5f", fontWeight: 500 }}>{rupees(r.profit)}</span>,
                <span style={{ color: GOLD, fontWeight: 500 }}>{rupees(r.line_total)}</span>,
              ])}
            />
        }
      </Card>
    </div>
  );
}

// ─── EXPENSES ────────────────────────────────────────────────────────────────
function ExpensesPage({ addToast }) {
  const [expenses, setExpenses] = useState([]);
  const [showModal, setShowModal] = useState(false);
  const [editExpense, setEditExpense] = useState(null);
  const [form, setForm] = useState({ name: "", category: "Rent", amount: "", date: "", notes: "" });

  const load = () => window.db.getExpenses().then(setExpenses);
  useEffect(() => {
    load();
  }, []);

  const total = expenses.reduce((s, e) => s + e.amount, 0);
  const catBreak = expenses.reduce((acc, e) => { acc[e.category] = (acc[e.category] || 0) + e.amount; return acc; }, {});

  const openAdd = () => { setForm({ name: "", category: "Rent", amount: "", date: new Date().toISOString().split("T")[0], notes: "" }); setEditExpense(null); setShowModal(true); };
  const openEdit = (i) => { setForm({ ...expenses[i] }); setEditExpense(expenses[i]); setShowModal(true); };

  const save = async () => {
    if (!form.name || !form.amount) { addToast("Fill required fields", "error"); return; }
    const e = { ...form, amount: Number(form.amount) };
    if (editExpense) { await window.db.updateExpense({ ...e, id: editExpense.id }); addToast("Expense updated", "success"); }
    else { await window.db.addExpense(e); addToast("Expense added", "success"); }
    setShowModal(false); load();
  };

  const doDelete = async (i) => {
    await window.db.deleteExpense(expenses[i].id); addToast("Expense deleted", "info"); load();
  };

  return (
    <div className="fade-in">
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24 }}>
        <div>
          <h1 style={{ fontFamily: "'Cormorant Garamond', serif", fontSize: 28, fontWeight: 300, color: "#E8E4D9" }}>Expenses</h1>
          <p style={{ color: "#555", fontSize: 13, marginTop: 4 }}>{new Date().toLocaleString("default", { month: "long", year: "numeric" })}</p>
        </div>
        <GoldButton onClick={openAdd}>+ Add Expense</GoldButton>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14, marginBottom: 20 }}>
        <StatCard label="Total Expenses" value={`₹${total.toLocaleString()}`} color="#c08060" />
        {Object.entries(catBreak).map(([cat, amt]) => (
          <StatCard key={cat} label={cat} value={`₹${amt.toLocaleString()}`} color="#666" />
        ))}
      </div>

      <Card>
        <Table
          headers={["Expense Name", "Category", "Amount", "Date", "Notes"]}
          rows={expenses.map(e => [
            <span style={{ color: "#E8E4D9", fontWeight: 500 }}>{e.name}</span>,
            <Badge color="#666">{e.category}</Badge>,
            <span style={{ color: "#c08060", fontWeight: 600 }}>₹{e.amount.toLocaleString()}</span>,
            e.date,
            <span style={{ color: "#555", fontSize: 12 }}>{e.notes}</span>
          ])}
          onEdit={openEdit}
          onDelete={doDelete}
        />
      </Card>

      {showModal && (
        <Modal title={editExpense ? "Edit Expense" : "Add Expense"} onClose={() => setShowModal(false)}>
          <FormField label="Expense Name"><input value={form.name} onChange={e => setForm(f => ({ ...f, name: e.target.value }))} /></FormField>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="Category">
              <select value={form.category} onChange={e => setForm(f => ({ ...f, category: e.target.value }))}>
                {["Rent", "Electricity", "Packaging", "Other"].map(c => <option key={c}>{c}</option>)}
              </select>
            </FormField>
            <FormField label="Amount (₹)"><input type="number" value={form.amount} onChange={e => setForm(f => ({ ...f, amount: e.target.value }))} /></FormField>
          </div>
          <FormField label="Date"><input type="date" value={form.date} onChange={e => setForm(f => ({ ...f, date: e.target.value }))} /></FormField>
          <FormField label="Notes"><textarea value={form.notes} onChange={e => setForm(f => ({ ...f, notes: e.target.value }))} rows={2} style={{ resize: "none" }} /></FormField>
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
            <GoldButton variant="outline" onClick={() => setShowModal(false)}>Cancel</GoldButton>
            <GoldButton onClick={save}>{editExpense ? "Save" : "Add Expense"}</GoldButton>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ─── SETTINGS ────────────────────────────────────────────────────────────────
function SettingsPage({ addToast, onLogout }) {
  const [settings, setSettings] = useState({ shopName: "Aura Fits", currency: "INR", address: "", phone: "" });
  const save = () => addToast("Settings saved", "success");

  return (
    <div className="fade-in">
      <h1 style={{ fontFamily: "'Cormorant Garamond', serif", fontSize: 28, fontWeight: 300, color: "#E8E4D9", marginBottom: 28 }}>Settings</h1>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
        <Card>
          <p style={{ fontSize: 13, color: GOLD, fontWeight: 500, marginBottom: 18 }}>Shop Information</p>
          <FormField label="Shop Name"><input value={settings.shopName} onChange={e => setSettings(s => ({ ...s, shopName: e.target.value }))} /></FormField>
          <FormField label="Address"><input value={settings.address} onChange={e => setSettings(s => ({ ...s, address: e.target.value }))} /></FormField>
          <FormField label="Phone"><input value={settings.phone} onChange={e => setSettings(s => ({ ...s, phone: e.target.value }))} /></FormField>
          <FormField label="Currency">
            <select value={settings.currency} onChange={e => setSettings(s => ({ ...s, currency: e.target.value }))}>
              <option value="INR">INR — Indian Rupee (₹)</option>
              <option value="USD">USD — US Dollar ($)</option>
              <option value="EUR">EUR — Euro (€)</option>
            </select>
          </FormField>
          <GoldButton onClick={save} style={{ marginTop: 8 }}>Save Changes</GoldButton>
        </Card>

        <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
          <Card>
            <p style={{ fontSize: 13, color: GOLD, fontWeight: 500, marginBottom: 16 }}>Database Location</p>
            <p style={{ fontSize: 12, color: "#555", marginBottom: 8 }}>Your data is stored locally on this computer.</p>
            <p style={{ fontSize: 11, color: "#444", fontFamily: "monospace" }}>%APPDATA%\aura-fits\aurafits.db</p>
          </Card>
          <Card>
            <p style={{ fontSize: 13, color: "#c05f5f", fontWeight: 500, marginBottom: 16 }}>Danger Zone</p>
            <p style={{ fontSize: 12, color: "#555", marginBottom: 14 }}>Sign out of the management console.</p>
            <button onClick={onLogout} style={{ background: "#2d1a1a", border: "1px solid #5c2d2d", color: "#c05f5f", borderRadius: 8, padding: "10px 22px", cursor: "pointer", fontSize: 13 }}>Sign Out</button>
          </Card>
        </div>
      </div>
    </div>
  );
}

// ─── MAIN APP ─────────────────────────────────────────────────────────────────
export default function App() {
  const [loggedIn, setLoggedIn] = useState(false);
  const [userName, setUserName] = useState("");
  const [activePage, setActivePage] = useState("dashboard");
  const [collapsed, setCollapsed] = useState(false);
  const [toasts, setToasts] = useState([]);

  useEffect(() => {
    const style = document.createElement("style");
    style.textContent = globalStyles;
    document.head.appendChild(style);
    return () => document.head.removeChild(style);
  }, []);

  const addToast = (message, type = "info") => {
    const id = Date.now();
    setToasts(prev => [...prev, { id, message, type }]);
    setTimeout(() => setToasts(prev => prev.filter(t => t.id !== id)), 3000);
  };

  if (!loggedIn) return <LoginPage onLogin={(name) => { setLoggedIn(true); setUserName(name); }} />;

  const sideW = collapsed ? 64 : 220;
  const pages = {
    dashboard: <Dashboard />,
    pos: <POSPage addToast={addToast} />,
    inventory: <InventoryPage addToast={addToast} />,
    categories: <CategoriesPage addToast={addToast} />,
    reports: <ReportsPage />,
    statement: <StatementPage addToast={addToast} />,
    expenses: <ExpensesPage addToast={addToast} />,
    settings: <SettingsPage addToast={addToast} onLogout={() => setLoggedIn(false)} />,
  };

  return (
    <div style={{ minHeight: "100vh", background: "#0A0A0A", fontFamily: "'DM Sans', sans-serif" }}>
      <Toast toasts={toasts} />
      <Sidebar active={activePage} setActive={setActivePage} collapsed={collapsed} setCollapsed={setCollapsed} userName={userName} />
      <div style={{ position: "fixed", top: 0, left: sideW, right: 0, height: 52, background: "#0D0D0D", borderBottom: "1px solid #1a1a1a", display: "flex", alignItems: "center", justifyContent: "space-between", padding: "0 28px", zIndex: 99, transition: "left 0.25s ease" }}>
        <p style={{ fontSize: 12, color: "#444", letterSpacing: "0.08em" }}>{NAV_ITEMS.find(n => n.id === activePage)?.label?.toUpperCase()}</p>
        <div style={{ display: "flex", alignItems: "center", gap: 20 }}>
          <p style={{ fontSize: 11, color: "#666" }}>{new Date().toLocaleDateString("en-IN", { weekday: "short", day: "numeric", month: "short", year: "numeric" })}</p>
          <div style={{ width: 30, height: 30, borderRadius: "50%", background: `${GOLD}20`, border: `1px solid ${GOLD}40`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, color: GOLD, fontWeight: 600 }}>{(userName || "A")[0].toUpperCase()}</div>
        </div>
      </div>
      <main style={{ marginLeft: sideW, paddingTop: 52, minHeight: "100vh", transition: "margin-left 0.25s ease" }}>
        <div style={{ padding: 28, maxWidth: 1400 }}>
          {pages[activePage]}
        </div>
      </main>
    </div>
  );
}
