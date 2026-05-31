import { useState, useEffect, useRef, useCallback } from "react";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, AreaChart, Area } from "recharts";
import { save } from "@tauri-apps/plugin-dialog";
import { writeFile } from "@tauri-apps/plugin-fs";

const GOLD = "#C9A84C";
const GOLD_LIGHT = "#E2C57A";
const GOLD_DIM = "#8A6D2E";

const todayDate = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
};
const rupees = (value) => `₹${(Number(value) || 0).toLocaleString("en-IN")}`;
const pdfAmt = (value) => `Rs.${(Number(value)||0).toFixed(0).replace(/\B(?=(\d{3})+(?!\d))/g,",")}`;
const escapeHtml = (value) => String(value ?? "").replace(/[&<>"']/g, ch => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch]));

// Groups raw statement rows by sale_id into wide sale records
function groupRowsIntoSales(rows) {
  const salesMap = {};
  rows.forEach(r => {
    if (!salesMap[r.sale_id]) {
      salesMap[r.sale_id] = {
        sale_id: r.sale_id,
        created_at: r.created_at,
        customer_name: r.customer_name,
        customer_phone: r.customer_phone,
        payment: r.payment,
        sale_total: r.sale_total,
        amount_paid: r.amount_paid,
        balance: r.balance,
        discount: r.discount,
        items: [],
      };
    }
    if (r.product_name) {
      salesMap[r.sale_id].items.push({
        product_name: r.product_name,
        category: r.category,
        size: r.size,
        color: r.color,
        qty: r.qty,
        cost: r.cost,
        price: r.price,
        line_total: r.line_total,
        profit: r.profit,
      });
    }
  });
  return Object.values(salesMap).sort((a, b) => a.sale_id - b.sale_id).map((s, idx) => ({ ...s, serial_no: idx + 1 }));
}

function ri(v) { return Math.round(Number(v) || 0); }
// Profit = (sell price - cost price) * qty — always use this formula
function calcProfit(items) { return items.reduce((t, i) => t + (ri(i.price) - ri(i.cost)) * ri(i.qty), 0); }

async function buildExcelXlsx(rows) {
  // Dynamically load SheetJS if not already present
  if (!window.XLSX) {
    await new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
      s.onload = resolve; s.onerror = reject;
      document.head.appendChild(s);
    });
  }
  const XLSX = window.XLSX;
  const sales = groupRowsIntoSales(rows);

  const headers = ["S.No.", "Bill #", "Date/Time", "Customer", "Phone", "Payment",
    "Products", "Cost Price", "Sell Price", "Total Price", "Paid", "Profit", "Due Balance"];

  const data = [headers];
  let totals = { cost: 0, sell: 0, total: 0, paid: 0, profit: 0, due: 0 };

  for (const s of sales) {
    const products = s.items.map(i => `${i.product_name} (x${i.qty})`).join(", ");
    const costParts = s.items.map(i => ri(i.cost) * ri(i.qty));
    const sellParts = s.items.map(i => ri(i.price) * ri(i.qty));
    const costPrice = costParts.reduce((a, b) => a + b, 0);
    const sellPrice = sellParts.reduce((a, b) => a + b, 0);
    const costStr = costParts.length > 1 ? `${costParts.join("+")}=${costPrice}` : String(costPrice);
    const sellStr = sellParts.length > 1 ? `${sellParts.join("+")}=${sellPrice}` : String(sellPrice);
    const profit = calcProfit(s.items);
    const total = ri(s.sale_total);
    const paid = ri(s.amount_paid);
    const due = ri(s.balance);
    data.push([s.serial_no, s.sale_id, s.created_at, s.customer_name, s.customer_phone || "",
      s.payment, products, costStr, sellStr, total, paid, profit, due]);
    totals.cost += costPrice; totals.sell += sellPrice; totals.total += total;
    totals.paid += paid; totals.profit += profit; totals.due += due;
  }

  data.push(["TOTALS", "", "", "", "", "", "",
    totals.cost, totals.sell, totals.total, totals.paid, totals.profit, totals.due]);

  const ws = XLSX.utils.aoa_to_sheet(data);

  // Style header row bold + gold background
  const range = XLSX.utils.decode_range(ws["!ref"]);
  for (let C = range.s.c; C <= range.e.c; C++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c: C })];
    if (cell) cell.s = { font: { bold: true }, fill: { fgColor: { rgb: "C9A84C" } } };
  }
  // Style totals row bold
  const lastRow = data.length - 1;
  for (let C = range.s.c; C <= range.e.c; C++) {
    const cell = ws[XLSX.utils.encode_cell({ r: lastRow, c: C })];
    if (cell) cell.s = { font: { bold: true }, fill: { fgColor: { rgb: "F5F5F5" } } };
  }

  // Set column widths
  ws["!cols"] = [6, 8, 18, 18, 14, 10, 40, 12, 12, 12, 12, 12, 12].map(w => ({ wch: w }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Statement");

  // Return as Uint8Array binary
  return XLSX.write(wb, { bookType: "xlsx", type: "array" });
}

function buildCsv(rows) {
  const sales = groupRowsIntoSales(rows);
  const escape = v => `"${String(v ?? "").replace(/"/g, '""')}"`;
  const headers = ["S.No.", "Bill #", "Date/Time", "Customer", "Phone", "Payment",
    "Products", "Cost Price", "Sell Price", "Total Price", "Paid", "Profit", "Due Balance"].map(escape).join(",");

  const dataRows = sales.map(s => {
    const products = s.items.map(i => `${i.product_name} (x${i.qty})`).join(", ");
    const costParts = s.items.map(i => ri(i.cost) * ri(i.qty));
    const sellParts = s.items.map(i => ri(i.price) * ri(i.qty));
    const costPrice = costParts.reduce((a, b) => a + b, 0);
    const sellPrice = sellParts.reduce((a, b) => a + b, 0);
    const costStr = costParts.length > 1 ? `${costParts.join("+")}=${costPrice}` : String(costPrice);
    const sellStr = sellParts.length > 1 ? `${sellParts.join("+")}=${sellPrice}` : String(sellPrice);
    const profit = calcProfit(s.items);
    return [
      s.serial_no, s.sale_id, s.created_at, s.customer_name, s.customer_phone || "", s.payment,
      products, costStr, sellStr, ri(s.sale_total), ri(s.amount_paid), profit, ri(s.balance),
    ].map(escape).join(",");
  });

  const totals = sales.reduce((acc, s) => {
    acc.cost += s.items.reduce((t, i) => t + ri(i.cost) * ri(i.qty), 0);
    acc.sell += s.items.reduce((t, i) => t + ri(i.price) * ri(i.qty), 0);
    acc.total += ri(s.sale_total);
    acc.paid += ri(s.amount_paid);
    acc.profit += calcProfit(s.items);
    acc.due += ri(s.balance);
    return acc;
  }, { cost: 0, sell: 0, total: 0, paid: 0, profit: 0, due: 0 });

  const totalsRow = ["TOTALS", "", "", "", "", "", "",
    totals.cost, totals.sell, totals.total, totals.paid, totals.profit, totals.due].map(escape).join(",");

  return [headers, ...dataRows, totalsRow].join("\r\n");
}

async function buildStatementPdf(sales, fromDate, toDate) {
  if (!window.jspdf) {
    await new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js";
      s.onload = resolve; s.onerror = reject;
      document.head.appendChild(s);
    });
  }
  if (!window.jspdf.jsPDF.autoTable) {
    await new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = "https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js";
      s.onload = resolve; s.onerror = reject;
      document.head.appendChild(s);
    });
  }
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "mm", format: "a4", orientation: "landscape" });

  doc.setFont("helvetica", "bold");
  doc.setFontSize(16);
  doc.text("AURA FITS - Sales Statement", 14, 16);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  doc.setTextColor(120);
  doc.text(`Period: ${fromDate} to ${toDate}    Generated: ${new Date().toLocaleString("en-IN")}`, 14, 23);
  doc.setTextColor(0);

  const grouped = groupRowsIntoSales(sales);
  const tableRows = grouped.map(s => {
    const products = s.items.map(i => `${i.product_name} (x${i.qty})`).join(", ");
    const costParts = s.items.map(i => ri(i.cost) * ri(i.qty));
    const sellParts = s.items.map(i => ri(i.price) * ri(i.qty));
    const costPrice = costParts.reduce((a, b) => a + b, 0);
    const sellPrice = sellParts.reduce((a, b) => a + b, 0);
    const costStr = costParts.length > 1 ? `${costParts.join("+")}=${costPrice}` : String(costPrice);
    const sellStr = sellParts.length > 1 ? `${sellParts.join("+")}=${sellPrice}` : String(sellPrice);
    const profit = calcProfit(s.items);
    return [
      s.serial_no, `#${s.sale_id}`, (s.created_at || "").slice(0, 16),
      s.customer_name || "Walk-in", s.customer_phone || "-", s.payment,
      products,
      `Rs.${costStr}`, `Rs.${sellStr}`, `Rs.${ri(s.sale_total)}`,
      `Rs.${ri(s.amount_paid)}`, `Rs.${profit}`,
      s.balance > 0 ? `Rs.${ri(s.balance)}` : "Cleared",
    ];
  });

  const totals = grouped.reduce((acc, s) => {
    acc.cost += s.items.reduce((t, i) => t + ri(i.cost) * ri(i.qty), 0);
    acc.sell += s.items.reduce((t, i) => t + ri(i.price) * ri(i.qty), 0);
    acc.total += ri(s.sale_total);
    acc.paid += ri(s.amount_paid);
    acc.profit += calcProfit(s.items);
    acc.due += ri(s.balance);
    return acc;
  }, { cost: 0, sell: 0, total: 0, paid: 0, profit: 0, due: 0 });

  tableRows.push([
    "", "TOTALS", "", "", "", "", "",
    `Rs.${totals.cost}`, `Rs.${totals.sell}`, `Rs.${totals.total}`,
    `Rs.${totals.paid}`, `Rs.${totals.profit}`, `Rs.${totals.due}`,
  ]);

  doc.autoTable({
    startY: 28,
    head: [["S.No.", "Bill", "Date/Time", "Customer", "Phone", "Payment",
      "Products", "Cost", "Sell", "Total", "Paid", "Profit", "Due Balance"]],
    body: tableRows,
    styles: { fontSize: 7, cellPadding: 2 },
    headStyles: { fillColor: [30, 25, 10], textColor: [201, 168, 76], fontStyle: "bold" },
    alternateRowStyles: { fillColor: [245, 245, 245] },
    columnStyles: {
      0: { cellWidth: 8 }, 1: { cellWidth: 12 }, 2: { cellWidth: 26 },
      3: { cellWidth: 25 }, 4: { cellWidth: 20 }, 5: { cellWidth: 16 },
      6: { cellWidth: 48 }, 7: { cellWidth: 18 }, 8: { cellWidth: 18 },
      9: { cellWidth: 18 }, 10: { cellWidth: 16 }, 11: { cellWidth: 16 }, 12: { cellWidth: 18 },
    },
    didParseCell: (data) => {
      if (data.row.index === tableRows.length - 1) {
        data.cell.styles.fontStyle = "bold";
        data.cell.styles.fillColor = [240, 235, 220];
      }
    },
  });
  return doc.output("arraybuffer");
}



function buildReceiptText(receipt) {
  return [
    "AURA FITS",
    `Receipt #${receipt.saleId}`,
    `Customer: ${receipt.customerName || "Walk-in Customer"}`,
    receipt.customerPhone ? `Phone: ${receipt.customerPhone}` : "",
    `Date: ${receipt.date}`,
    `Payment: ${receipt.payment}`,
    "---",
    ...receipt.items.map(item => `${item.name} x${item.qty}  ${rupees(item.lineTotal)}`),
    "---",
    `Total: ${rupees(receipt.total)}`,
    receipt.amountPaid > 0 ? `Amount Paid: ${rupees(receipt.amountPaid)}` : "",
    receipt.balance > 0 ? `Balance Due: ${rupees(receipt.balance)}` : "",
    "---",
    "Thank you for shopping with us!",
  ].filter(Boolean).join("\n");
}

async function generateReceiptPdf(receipt) {
  if (!window.jspdf) {
    await new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js";
      s.onload = resolve; s.onerror = reject;
      document.head.appendChild(s);
    });
  }
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "mm", format: [80, 220], orientation: "portrait" });
  const lm = 6; const rm = 74; let y = 10;

  // Header
  doc.setFont("helvetica", "bold");
  doc.setFontSize(18);
  doc.text("AURA FITS", 40, y, { align: "center" }); y += 7;
  doc.setFontSize(8); doc.setFont("helvetica", "normal");
  doc.setTextColor(100);
  doc.text(`Receipt #${receipt.saleId}`, 40, y, { align: "center" }); y += 4;
  doc.text(receipt.date, 40, y, { align: "center" }); y += 6;
  doc.setTextColor(0);

  // Divider
  doc.setDrawColor(180); doc.line(lm, y, rm, y); y += 5;

  // Customer info
  const infoRows = [
    ["Customer", receipt.customerName || "Walk-in Customer"],
    ...(receipt.customerPhone ? [["Phone", receipt.customerPhone]] : []),
    ["Payment", receipt.payment],
  ];
  infoRows.forEach(([label, val]) => {
    doc.setFont("helvetica", "bold"); doc.setFontSize(8);
    doc.text(label + ":", lm, y);
    doc.setFont("helvetica", "normal");
    doc.text(String(val), lm + 20, y); y += 5;
  });
  y += 1; doc.line(lm, y, rm, y); y += 5;

  // Items header
  doc.setFont("helvetica", "bold"); doc.setFontSize(8);
  doc.text("Item", lm, y);
  doc.text("Qty", 50, y, { align: "center" });
  doc.text("Amount", rm, y, { align: "right" }); y += 4;
  doc.setFont("helvetica", "normal"); doc.setFontSize(8);
  doc.setDrawColor(220); doc.line(lm, y, rm, y); y += 3;

  // Items
  (receipt.items || []).forEach(item => {
    const name = String(item.name || "").length > 24 ? String(item.name || "").slice(0, 23) + "…" : String(item.name || "");
    doc.text(name, lm, y);
    doc.text(String(item.qty), 50, y, { align: "center" });
    doc.text(pdfAmt(item.lineTotal), rm, y, { align: "right" });
    y += 5;
  });

  y += 1; doc.setDrawColor(180); doc.line(lm, y, rm, y); y += 5;

  // Totals
  doc.setFontSize(8);
  // Total box
  doc.setFillColor(240, 240, 240); doc.roundedRect(lm, y - 2, rm - lm, 9, 1, 1, "F");
  doc.setFont("helvetica", "bold"); doc.setFontSize(10);
  doc.text("TOTAL", lm + 2, y + 4);
  doc.text(pdfAmt(receipt.total), rm - 2, y + 4, { align: "right" });
  y += 13; doc.setFont("helvetica", "normal"); doc.setFontSize(8);

  if (receipt.amountPaid > 0) {
    doc.text("Amount Paid", lm, y); doc.text(pdfAmt(receipt.amountPaid), rm, y, { align: "right" }); y += 5;
  }
  if (receipt.balance > 0) {
    doc.setTextColor(180, 60, 60); doc.setFont("helvetica", "bold");
    doc.text("Balance Due", lm, y); doc.text(pdfAmt(receipt.balance), rm, y, { align: "right" });
    y += 5; doc.setTextColor(0); doc.setFont("helvetica", "normal");
  }

  y += 3; doc.setDrawColor(180); doc.line(lm, y, rm, y); y += 6;
  doc.setFontSize(7.5); doc.setTextColor(120);
  doc.text("Thank you for shopping with us!", 40, y, { align: "center" });
  doc.setTextColor(0);
  return doc;
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

function GoldButton({ children, onClick, style = {}, variant = "primary", size = "md", disabled = false }) {
  return (
    <button style={{
      background: variant === "primary" ? `linear-gradient(135deg, ${GOLD} 0%, ${GOLD_LIGHT} 50%, ${GOLD} 100%)` : "transparent",
      border: `1px solid ${GOLD}`, color: variant === "primary" ? "#0A0A0A" : GOLD,
      borderRadius: 8, cursor: disabled ? "not-allowed" : "pointer", fontFamily: "'DM Sans', sans-serif",
      fontWeight: 600, fontSize: size === "sm" ? 12 : 14,
      padding: size === "sm" ? "6px 14px" : "10px 22px",
      transition: "all 0.2s", letterSpacing: "0.02em", opacity: disabled ? 0.6 : 1, ...style
    }} onClick={disabled ? undefined : onClick} disabled={disabled}
      onMouseEnter={e => { if (!disabled) e.currentTarget.style.opacity = "0.85"; }}
      onMouseLeave={e => { if (!disabled) e.currentTarget.style.opacity = disabled ? "0.6" : "1"; }}
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
    try {
      const result = await window.db.login({ email, password });
      if (result.ok) { onLogin(result.name); }
      else { setError("Invalid email or password"); setLoading(false); }
    } catch (e) { setError("Login failed. Please try again."); setLoading(false); }
  };

  const handleSignup = async () => {
    setError("");
    if (!signupName || !signupEmail || !signupPass || !signupAnswer) { setError("Please fill in all fields"); return; }
    if (signupPass !== signupPass2) { setError("Passwords do not match"); return; }
    if (signupPass.length < 6) { setError("Password must be at least 6 characters"); return; }
    try {
      const result = await window.db.signup({ name: signupName, email: signupEmail, password: signupPass, securityQuestion: signupQ, securityAnswer: signupAnswer });
      if (result.ok) { setScreen("login"); setEmail(signupEmail); setPassword(""); setError(""); }
      else { setError(result.error || "Signup failed"); }
    } catch (e) { setError("Signup failed. Please try again."); }
  };

  const handleForgotLookup = async () => {
    setError("");
    try {
      const user = await window.db.findUser(forgotEmail);
      if (!user) { setError("No account found with that email"); return; }
      setForgotUser({ ...user, email: forgotEmail });
      setForgotQ(user.securityQuestion);
    } catch (e) { setError("Failed to look up account. Try again."); }
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
  { id: "pending", label: "Pending Balances", icon: "⚠" },
  { id: "expenses", label: "Shop Expenses", icon: "◉" },
  { id: "personal", label: "Personal Expenses", icon: "👤" },
  { id: "settings", label: "Settings", icon: "⊛" },
];

function Sidebar({ active, setActive, collapsed, setCollapsed, userName, pendingCount }) {
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
          const badge = item.id === "pending" && pendingCount > 0 ? pendingCount : null;
          return (
            <div key={item.id} onClick={() => setActive(item.id)} title={collapsed ? item.label : ""} style={{
              display: "flex", alignItems: "center", gap: 12,
              padding: collapsed ? "12px 0" : "12px 20px",
              justifyContent: collapsed ? "center" : "flex-start",
              cursor: "pointer", transition: "all 0.15s",
              background: isActive ? `${GOLD}10` : "transparent",
              borderLeft: isActive ? `2px solid ${GOLD}` : "2px solid transparent", marginBottom: 2,
              position: "relative",
            }}
              onMouseEnter={e => !isActive && (e.currentTarget.style.background = "#141414")}
              onMouseLeave={e => !isActive && (e.currentTarget.style.background = "transparent")}
            >
              <span style={{ fontSize: 16, color: item.id === "pending" && pendingCount > 0 ? "#c05f5f" : isActive ? GOLD : "#444", minWidth: 20, textAlign: "center" }}>{item.icon}</span>
              {!collapsed && <span style={{ fontSize: 13, color: isActive ? GOLD : item.id === "pending" && pendingCount > 0 ? "#c08060" : "#777", fontWeight: isActive ? 500 : 400, whiteSpace: "nowrap", flex: 1 }}>{item.label}</span>}
              {!collapsed && badge && <span style={{ background: "#c05f5f", color: "#fff", borderRadius: 10, fontSize: 10, fontWeight: 700, padding: "1px 6px", minWidth: 18, textAlign: "center" }}>{badge}</span>}
              {collapsed && badge && <span style={{ position: "absolute", top: 6, right: 6, background: "#c05f5f", color: "#fff", borderRadius: "50%", fontSize: 9, fontWeight: 700, width: 14, height: 14, display: "flex", alignItems: "center", justifyContent: "center" }}>{badge}</span>}
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
  const [stats, setStats] = useState({ today: { total: 0, count: 0 }, month: { total: 0 }, weekly: [], real_profit: 0 });
  const [products, setProducts] = useState([]);
  const [expenses, setExpenses] = useState([]);
  const [todaySaleRows, setTodaySaleRows] = useState([]);
  const [monthSaleRows, setMonthSaleRows] = useState([]);
  const tooltipStyle = { background: "#141414", border: "1px solid #222", borderRadius: 8, color: "#E8E4D9", fontSize: 12 };

  useEffect(() => {
    window.db.getSalesStats().then(setStats).catch(() => {});
    window.db.getProducts().then(setProducts).catch(() => {});
    window.db.getExpenses().then(setExpenses).catch(() => {});
    // Fetch today's and this month's statement to compute correct profit
    const today = new Date().toISOString().slice(0, 10);
    const monthStart = today.slice(0, 8) + "01";
    window.db.getStatement(today, today).then(setTodaySaleRows).catch(() => {});
    window.db.getStatement(monthStart, today).then(setMonthSaleRows).catch(() => {});
  }, []);

  // Correct profit = (sell price - cost price) * qty, summed per item
  const computeProfit = (rows) => {
    const seen = new Set();
    return rows.reduce((total, r) => {
      if (!r.product_name || seen.has(`${r.sale_id}-${r.product_id}`)) return total;
      seen.add(`${r.sale_id}-${r.product_id}`);
      return total + (ri(r.price) - ri(r.cost)) * ri(r.qty);
    }, 0);
  };
  const todayProfit = computeProfit(todaySaleRows);
  const monthProfit = computeProfit(monthSaleRows);
  const lowStock = products.filter(p => p.stock <= 5);
  const thisMonth = new Date().toISOString().slice(0, 7);
  const monthExpenses = expenses.filter(e => (e.date || "").slice(0, 7) === thisMonth).reduce((s, e) => s + e.amount, 0);
  const todayStr = new Date().toLocaleDateString("en-IN", { weekday: "long", day: "numeric", month: "long", year: "numeric" });

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

      <p style={{ fontSize: 11, color: "#555", letterSpacing: "0.1em", textTransform: "uppercase", marginBottom: 10 }}>Today</p>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14, marginBottom: 20 }}>
        <StatCard label="Today's Revenue" value={`₹${(stats.today?.total || 0).toLocaleString()}`} sub="collected today" icon="💰" />
        <StatCard label="Today's Profit" value={`₹${todayProfit.toLocaleString("en-IN", { maximumFractionDigits: 0 })}`} color={todayProfit >= 0 ? "#5fa05f" : "#c05f5f"} sub="sell − cost" icon="📈" />
        <StatCard label="Today's Sales" value={stats.today?.count || 0} sub="transactions" color="#E8E4D9" icon="🧾" />
        <StatCard label="Low Stock" value={lowStock.length} sub={lowStock.length > 0 ? "needs restock" : "all good ✓"} color={lowStock.length > 0 ? "#c05f5f" : "#5fa05f"} icon="⚠" />
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
          {lowStock.length === 0
            ? <p style={{ color: "#5fa05f", fontSize: 13, padding: "20px 0" }}>All products well stocked ✓</p>
            : <Table headers={["Product", "Brand", "Stock"]} rows={lowStock.map(p => [p.name, p.brand, <span style={{ color: p.stock <= 3 ? "#c05f5f" : "#c08060", fontWeight: 500 }}>Only {p.stock} left</span>])} />
          }
        </Card>
      </div>
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
  const [customerPhone, setCustomerPhone] = useState("");
  const [amountPaid, setAmountPaid] = useState("");
  const [cart, setCart] = useState([]);
  const [payment, setPayment] = useState("Cash");
  const [lastReceipt, setLastReceipt] = useState(null);
  const [checkingOut, setCheckingOut] = useState(false);
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
    if (product.stock <= 0) { addToast(`${product.name} is out of stock`, "error"); return; }
    setCart(prev => {
      const ex = prev.find(i => i.id === product.id);
      if (ex) {
        if (ex.qty >= product.stock) { addToast(`Only ${product.stock} in stock`, "error"); return prev; }
        return prev.map(i => i.id === product.id ? { ...i, qty: i.qty + 1 } : i);
      }
      return [...prev, { ...product, qty: 1, sellPrice: "" }];
    });
  };

  const updateQty = (id, delta) => {
    setCart(prev => {
      const product = products.find(p => p.id === id);
      return prev.map(i => {
        if (i.id !== id) return i;
        const newQty = i.qty + delta;
        if (delta > 0 && product && newQty > product.stock) {
          addToast(`Only ${product.stock} in stock`, "error");
          return i;
        }
        return { ...i, qty: Math.max(0, newQty) };
      }).filter(i => i.qty > 0);
    });
  };

  const updateSellPrice = (id, val) => {
    setCart(prev => prev.map(i => {
      if (i.id !== id) return i;
      const numVal = Number(val);
      const belowCost = val !== "" && numVal < i.cost;
      return { ...i, sellPrice: val, belowCost };
    }));
  };

  const total = cart.reduce((s, i) => s + (Number(i.sellPrice) || 0) * i.qty, 0);

  const checkout = async () => {
    if (checkingOut) return;
    if (!cart.length) { addToast("Cart is empty", "error"); return; }
    if (cart.some(i => !Number(i.sellPrice))) { addToast("Enter sale price for every product", "error"); return; }
    if (cart.some(i => Number(i.sellPrice) < i.cost)) { addToast("Sell price cannot be below cost price — fix highlighted items", "error"); return; }
    const phoneVal = customerPhone.trim();
    if (phoneVal && !/^\d{7,15}$/.test(phoneVal)) { addToast("Phone number must be 7–15 digits only", "error"); return; }
    setCheckingOut(true);
    try {
      const cartToSave = cart.map(i => ({ ...i, price: Number(i.sellPrice) || 0 }));
      const receiptCustomer = customerName.trim() || "Walk-in Customer";
      const paid = Number(amountPaid) || 0;
      const balance = Math.max(0, total - paid);
      const result = await window.db.recordSale({ cart: cartToSave, total, discount: 0, payment, customerName: receiptCustomer, customerPhone: customerPhone.trim(), amountPaid: paid, balance });
      if (result.ok) {
        setLastReceipt({
          saleId: result.saleId,
          date: new Date().toLocaleString("en-IN"),
          customerName: receiptCustomer,
          customerPhone: customerPhone.trim(),
          payment,
          items: cart.map(i => {
            const unitPrice = Number(i.sellPrice) || 0;
            return { name: i.name, qty: i.qty, unitPrice, lineTotal: unitPrice * i.qty };
          }),
          total,
          amountPaid: paid,
          balance,
        });
        addToast(`Sale of ₹${total.toLocaleString()} recorded via ${payment}`, "success");
        setCart([]); setCustomerName(""); setCustomerPhone(""); setAmountPaid(""); setPayment("Cash");
        window.db.getProducts().then(setProducts);
      } else {
        addToast("Failed to record sale", "error");
      }
    } catch (e) {
      addToast("Failed to record sale: " + (e?.message || String(e)), "error");
    } finally {
      setCheckingOut(false);
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
          <input value={customerPhone} onChange={e => setCustomerPhone(e.target.value)} placeholder="Customer phone (optional)" style={{ marginTop: 8, height: 40 }} type="tel" />
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
                  <span style={{ fontSize: 10, color: item.belowCost ? "#c05f5f" : "#666", textTransform: "uppercase", letterSpacing: "0.08em" }}>{item.belowCost ? "⚠ Below Cost!" : "Sell Price"}</span>
                  <input type="number" value={item.sellPrice} onChange={e => updateSellPrice(item.id, e.target.value)} placeholder="Sale price"
                    style={{ width: "100%", padding: "6px 8px", textAlign: "right", fontSize: 12, color: item.belowCost ? "#c05f5f" : GOLD, background: item.belowCost ? "#2d1a1a" : "#1a1a1a", border: `1px solid ${item.belowCost ? "#c05f5f" : "#333"}`, borderRadius: 6 }} />
                  {item.belowCost
                    ? <span style={{ fontSize: 11, color: "#c05f5f", fontWeight: 600 }}>Min: ₹{item.cost.toLocaleString()}</span>
                    : <span style={{ fontSize: 11, color: "#555" }}>× {item.qty} = ₹{((Number(item.sellPrice) || 0) * item.qty).toLocaleString()}</span>
                  }
                </div>
              </div>
            ))
          }
        </div>
        <div style={{ padding: "16px 20px", borderTop: "1px solid #1e1e1e" }}>
          <div style={{ display: "flex", justifyContent: "space-between", padding: "12px 0", borderTop: "1px solid #1e1e1e" }}>
            <span style={{ fontSize: 15, color: "#E8E4D9", fontWeight: 500 }}>Total</span>
            <span style={{ fontSize: 20, color: GOLD, fontWeight: 700, fontFamily: "'Cormorant Garamond', serif" }}>₹{total.toLocaleString()}</span>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
            <span style={{ fontSize: 12, color: "#666", flex: 1 }}>Amount Paid ₹</span>
            <input type="number" value={amountPaid} onChange={e => setAmountPaid(e.target.value)} style={{ width: 110, padding: "6px 10px", textAlign: "right" }} min={0} placeholder="Full amount" />
          </div>
          {amountPaid !== "" && Number(amountPaid) < total && (
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8, background: "#2d1a1a", borderRadius: 8, padding: "8px 12px", border: "1px solid #5c2d2d" }}>
              <span style={{ fontSize: 13, color: "#c05f5f", fontWeight: 600 }}>Balance Due</span>
              <span style={{ fontSize: 14, color: "#c05f5f", fontWeight: 700 }}>₹{Math.max(0, total - Number(amountPaid)).toLocaleString()}</span>
            </div>
          )}
          {amountPaid !== "" && Number(amountPaid) >= total && (
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8, background: "#1a2d1a", borderRadius: 8, padding: "8px 12px", border: "1px solid #2d5c2d" }}>
              <span style={{ fontSize: 13, color: "#5fa05f", fontWeight: 600 }}>Change</span>
              <span style={{ fontSize: 14, color: "#5fa05f", fontWeight: 700 }}>₹{Math.max(0, Number(amountPaid) - total).toLocaleString()}</span>
            </div>
          )}
          <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
            {["Cash", "UPI", "Card"].map(m => (
              <button key={m} onClick={() => setPayment(m)} style={{
                flex: 1, padding: "8px 0", background: payment === m ? `${GOLD}20` : "transparent",
                border: `1px solid ${payment === m ? GOLD : "#2a2a2a"}`,
                color: payment === m ? GOLD : "#666", borderRadius: 8, cursor: "pointer", fontSize: 12
              }}>{m}</button>
            ))}
          </div>
          <GoldButton onClick={checkout} disabled={checkingOut} style={{ width: "100%", padding: "14px", fontSize: 15, opacity: checkingOut ? 0.6 : 1 }}>{checkingOut ? "Processing..." : `✓ Checkout — ₹${total.toLocaleString()}`}</GoldButton>
          {cart.length > 0 && (
            <button onClick={() => setCart([])} style={{ width: "100%", background: "none", border: "none", color: "#444", cursor: "pointer", fontSize: 12, marginTop: 10, padding: "6px" }}>Clear Cart</button>
          )}
        </div>
      </div>
      {lastReceipt && (
        <Modal title={`Receipt #${lastReceipt.saleId}`} onClose={() => setLastReceipt(null)} width={460}>
          <div style={{ background: "#0d0d0d", border: "1px solid #1e1e1e", borderRadius: 10, padding: 14, marginBottom: 16 }}>
            <p style={{ fontSize: 12, color: "#666", marginBottom: 2 }}>{lastReceipt.date} · {lastReceipt.payment}</p>
            {lastReceipt.customerPhone && <p style={{ fontSize: 12, color: "#555", marginBottom: 4 }}>📞 {lastReceipt.customerPhone}</p>}
            {lastReceipt.items.map((item, i) => (
              <div key={i} style={{ display: "grid", gridTemplateColumns: "1fr 48px 90px", gap: 10, padding: "8px 0", borderBottom: i === lastReceipt.items.length - 1 ? "none" : "1px solid #1a1a1a" }}>
                <span style={{ color: "#C8C4B8", fontSize: 13 }}>{item.name}</span>
                <span style={{ color: "#777", fontSize: 13, textAlign: "center" }}>x{item.qty}</span>
                <span style={{ color: GOLD, fontSize: 13, textAlign: "right" }}>{rupees(item.lineTotal)}</span>
              </div>
            ))}
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 18, color: GOLD, marginBottom: 8, borderTop: "1px solid #222", paddingTop: 12 }}><span>Total</span><span>{rupees(lastReceipt.total)}</span></div>
          {lastReceipt.amountPaid > 0 && (
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, color: "#888", marginBottom: 4 }}><span>Amount Paid</span><span>{rupees(lastReceipt.amountPaid)}</span></div>
          )}
          {lastReceipt.balance > 0 && (
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 14, color: "#c05f5f", fontWeight: 700, background: "#2d1a1a", borderRadius: 8, padding: "8px 12px", marginBottom: 8, border: "1px solid #5c2d2d" }}>
              <span>Balance Due</span><span>{rupees(lastReceipt.balance)}</span>
            </div>
          )}
          <div style={{ marginBottom: 14 }} />
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
            <GoldButton variant="outline" onClick={() => setLastReceipt(null)}>Close</GoldButton>
            <GoldButton onClick={async () => {
              try {
                const doc = await generateReceiptPdf(lastReceipt);
                doc.save(`receipt-${lastReceipt.saleId}.pdf`);
                addToast("Receipt downloaded ✓", "success");
              } catch (e) {
                addToast("PDF failed: " + (e?.message || String(e)), "error");
              }
            }}>⬇ Download Receipt</GoldButton>
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
  const [form, setForm] = useState({ name: "", category: "T-Shirts", brand: "", size: "", color: "", cost: "", price: "", stock: "" });

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

  const openAdd = () => { setForm({ name: "", category: categories[0]?.name || "T-Shirts", brand: "", size: "", color: "", cost: "", price: "", stock: "" }); setEditProduct(null); setShowModal(true); };
  const openEdit = (i) => { setForm({ ...filtered[i], price: String(filtered[i].price || filtered[i].cost || "") }); setEditProduct(filtered[i]); setShowModal(true); };

  // localStorage cache for initial stock (fallback when DB doesn't store it)
  const getInitialStock = (p) => {
    // If DB returned initial_stock, use it
    if (p.initial_stock && p.initial_stock > 0) return p.initial_stock;
    // Otherwise read from localStorage cache
    const cached = localStorage.getItem(`aura_initstock_${p.id}`);
    return cached ? Number(cached) : p.stock;
  };

  const save = async () => {
    if (!form.name || !form.cost || !form.stock) { addToast("Please fill required fields", "error"); return; }
    const stockVal = Number(form.stock);
    const costVal = Number(form.cost);
    const p = { ...form, cost: costVal, price: costVal, stock: stockVal };
    if (editProduct) {
      await window.db.updateProduct({ ...p, id: editProduct.id, initial_stock: editProduct.initial_stock || editProduct.stock });
      addToast("Product updated", "success");
      setShowModal(false); load();
    } else {
      await window.db.addProduct({ ...p, initial_stock: stockVal });
      addToast("Product added", "success");
      setShowModal(false);
      // Reload and cache initial_stock for the new product by ID
      const updated = await window.db.getProducts().catch(() => []);
      const existingIds = new Set(products.map(pr => pr.id));
      const newProd = updated.find(pr => !existingIds.has(pr.id));
      if (newProd) localStorage.setItem(`aura_initstock_${newProd.id}`, String(stockVal));
      setProducts(updated);
      window.db.getCategories().then(setCategories);
    }
  };

  const doDelete = async (i) => {
    await window.db.deleteProduct(filtered[i].id);
    setConfirmDelete(null); addToast("Product deleted", "info"); load();
  };

  const totalStock = products.reduce((s, p) => s + p.stock, 0);
  const totalValue = products.reduce((s, p) => s + p.cost * p.stock, 0);
  const outOfStock = products.filter(p => p.stock <= 0).length;
  const lowStock = products.filter(p => p.stock > 0 && p.stock <= 5).length;

  return (
    <div className="fade-in">
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24 }}>
        <div>
          <h1 style={{ fontFamily: "'Cormorant Garamond', serif", fontSize: 28, fontWeight: 300, color: "#E8E4D9" }}>Inventory</h1>
          <p style={{ color: "#555", fontSize: 13, marginTop: 4 }}>{products.length} products · {totalStock} units</p>
        </div>
        <GoldButton onClick={openAdd}>+ Add Product</GoldButton>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14, marginBottom: 20 }}>
        <StatCard label="Total Products" value={products.length} sub="in inventory" color={GOLD} icon="◫" />
        <StatCard label="Total Stock" value={totalStock} sub="units available" color="#5fa05f" icon="📦" />
        <StatCard label="Inventory Value" value={rupees(totalValue)} sub="at cost price" color="#E8E4D9" icon="💰" />
        <StatCard label="Needs Attention" value={outOfStock + lowStock} sub={`${outOfStock} out · ${lowStock} low`} color={outOfStock > 0 ? "#c05f5f" : "#c08060"} icon="⚠" />
      </div>

      <div style={{ marginBottom: 16, display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
        <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search by name or brand..." style={{ maxWidth: 260 }} />
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {["All", ...categories.map(c => c.name)].map(cat => (
            <button key={cat} onClick={() => setCatFilter(cat)} style={{
              padding: "6px 14px", borderRadius: 20, fontSize: 12, cursor: "pointer", fontWeight: catFilter === cat ? 600 : 400,
              background: catFilter === cat ? `${GOLD}20` : "transparent",
              border: `1px solid ${catFilter === cat ? GOLD : "#2a2a2a"}`,
              color: catFilter === cat ? GOLD : "#666",
              transition: "all 0.15s",
            }}>{cat === "All" ? `All (${products.length})` : `${categories.find(c=>c.name===cat)?.icon||""} ${cat} (${products.filter(p=>p.category===cat).length})`}</button>
          ))}
        </div>
        <span style={{ fontSize: 12, color: "#555", marginLeft: 4 }}>{filtered.length} results</span>
      </div>

      <Card>
        <Table
          headers={["Product Name", "Category", "Brand", "Size", "Color", "Cost Price", "Initial", "Stock"]}
          rows={filtered.map(p => [
            <span style={{ color: "#E8E4D9", fontWeight: 500 }}>{p.name}</span>,
            <Badge>{p.category}</Badge>,
            p.brand, p.size, p.color,
            `₹${p.cost.toLocaleString()}`,
            <span style={{ color: "#555", fontSize: 12 }}>{getInitialStock(p)}</span>,
            <span style={{ color: p.stock <= 0 ? "#c05f5f" : p.stock <= 5 ? "#c08060" : "#5fa05f", fontWeight: 500 }}>
              {p.stock <= 0 ? "✕ Out" : p.stock <= 5 ? `⚠ ${p.stock}` : p.stock}
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
  const [monthRows, setMonthRows] = useState([]);
  const tooltipStyle = { background: "#141414", border: "1px solid #222", borderRadius: 8, color: "#E8E4D9", fontSize: 12 };

  useEffect(() => {
    window.db.getTodaySales().then(setTodaySales).catch(() => {});
    window.db.getSalesStats().then(setStats).catch(() => {});
    window.db.getExpenses().then(setExpenses).catch(() => {});
    const today = new Date().toISOString().slice(0, 10);
    const monthStart = today.slice(0, 8) + "01";
    window.db.getStatement(monthStart, today).then(setMonthRows).catch(() => {});
  }, []);

  const thisMonth = new Date().toISOString().slice(0, 7);
  const monthExpenses = expenses.filter(e => (e.date || "").slice(0, 7) === thisMonth).reduce((s, e) => s + e.amount, 0);
  // Correct profit = (sell - cost) * qty per item, no duplicates
  const profit = (() => {
    const seen = new Set();
    return monthRows.reduce((total, r) => {
      if (!r.product_name || seen.has(`${r.sale_id}-${r.product_id}`)) return total;
      seen.add(`${r.sale_id}-${r.product_id}`);
      return total + (ri(r.price) - ri(r.cost)) * ri(r.qty);
    }, 0);
  })();

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

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14, marginBottom: 20 }}>
        <StatCard label="Today's Revenue" value={`₹${(stats.today?.total || 0).toLocaleString()}`} sub={`${stats.today?.count || 0} sales`} />
        <StatCard label="Monthly Revenue" value={`₹${(stats.month?.total || 0).toLocaleString()}`} color="#E8E4D9" />
        <StatCard label="Monthly Expenses" value={`₹${monthExpenses.toLocaleString()}`} color="#c08060" />
        <StatCard label="Monthly Profit" value={`₹${profit.toLocaleString("en-IN", { maximumFractionDigits: 0 })}`} color={profit >= 0 ? "#5fa05f" : "#c05f5f"} />
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
function StatementPage({ addToast, onRefreshPending }) {
  const [tab, setTab] = useState("sales"); // "sales" | "pending"
  const [fromDate, setFromDate] = useState(todayDate());
  const [toDate, setToDate] = useState(todayDate());
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(false);
  const [editSale, setEditSale] = useState(null);
  const [editForm, setEditForm] = useState(null);
  const [products, setProducts] = useState([]);
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [pending, setPending] = useState([]);
  const [payModal, setPayModal] = useState(null);
  const [search, setSearch] = useState("");

  const load = async () => {
    if (!fromDate || !toDate) { addToast("Choose both dates", "error"); return; }
    if (fromDate > toDate) { addToast("From date cannot be after To date", "error"); return; }
    setLoading(true);
    try { setRows(await window.db.getStatement(fromDate, toDate)); }
    catch (e) { addToast(e?.message || "Failed to load statement", "error"); }
    finally { setLoading(false); }
  };

  const loadPending = () => window.db.getPendingBalances().then(setPending).catch(() => {});

  useEffect(() => { load(); window.db.getProducts().then(setProducts); loadPending(); }, []);

  // Group rows by sale_id
  const salesMap = {};
  rows.forEach(r => {
    if (!salesMap[r.sale_id]) salesMap[r.sale_id] = { ...r, items: [] };
    if (r.product_name) salesMap[r.sale_id].items.push(r);
  });
  const sales = Object.values(salesMap).sort((a, b) => b.sale_id - a.sale_id);
  const saleIds = new Set(rows.map(r => r.sale_id));
  const totalSales = [...saleIds].reduce((sum, id) => sum + (salesMap[id]?.sale_total || 0), 0);
  const totalProfit = (() => {
    const sales = groupRowsIntoSales(rows);
    return sales.reduce((t, s) => t + calcProfit(s.items), 0);
  })();
  const totalBalance = [...saleIds].reduce((sum, id) => sum + (salesMap[id]?.balance || 0), 0);

  const filteredSales = search.trim()
    ? sales.filter(s => (s.customer_name || "").toLowerCase().includes(search.toLowerCase()) || (s.customer_phone || "").includes(search))
    : sales;

  const openEdit = (sale) => {
    setEditSale(sale);
    setEditForm({
      customerName: sale.customer_name,
      customerPhone: sale.customer_phone || "",
      payment: sale.payment,
      amountPaid: sale.amount_paid || 0,
      balance: sale.balance || 0,
      cart: sale.items.map(i => ({
        id: i.product_id || null, name: i.product_name, category: i.category,
        size: i.size, color: i.color, cost: i.cost, qty: i.qty, price: i.price,
        sellPrice: String(i.price ?? ""),
      })),
    });
  };

  const saveEdit = async () => {
    if (!editForm.cart.length) { addToast("Add at least one product", "error"); return; }
    const cart = editForm.cart.map(i => ({ ...i, price: Number(i.sellPrice) || i.price }));
    const total = cart.reduce((s, i) => s + (Number(i.sellPrice) || i.price) * i.qty, 0);
    const paid = Number(editForm.amountPaid) || 0;
    const balance = Math.max(0, total - paid);
    try {
      await window.db.updateSale({ id: editSale.sale_id, customerName: editForm.customerName, customerPhone: editForm.customerPhone, payment: editForm.payment, discount: 0, amountPaid: paid, balance, total, cart });
      addToast("Sale updated", "success");
      setEditSale(null); setEditForm(null); load(); loadPending(); onRefreshPending?.();
    } catch (e) { addToast("Failed to update: " + (e?.message || String(e)), "error"); }
  };

  const doDelete = async (id) => {
    try {
      await window.db.deleteSale(id);
      addToast("Sale deleted, stock restored", "success");
      setConfirmDeleteId(null); load(); loadPending();
    } catch (e) { addToast("Failed to delete: " + (e?.message || String(e)), "error"); }
  };

  const [showExportMenu, setShowExportMenu] = useState(false);
  const exportRef = useRef(null);
  useEffect(() => {
    if (!showExportMenu) return;
    const handler = (e) => { if (exportRef.current && !exportRef.current.contains(e.target)) setShowExportMenu(false); };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [showExportMenu]);

  const doExport = async (format) => {
    setShowExportMenu(false);
    if (!rows.length) { addToast("No data to export", "error"); return; }
    if (format === "pdf") {
      try {
        const filePath = await save({
          defaultPath: `aura-fits-statement-${fromDate}-to-${toDate}.pdf`,
          filters: [{ name: "PDF Document", extensions: ["pdf"] }],
        });
        if (!filePath) return;
        const pdfBytes = await buildStatementPdf(rows, fromDate, toDate);
        await writeFile(filePath, new Uint8Array(pdfBytes));
        addToast("Statement PDF saved ✓", "success");
      } catch (e) { addToast("PDF failed: " + (e?.message || String(e)), "error"); }
      return;
    }
    const isCSV = format === "csv";
    try {
      if (isCSV) {
        const filePath = await save({
          defaultPath: `aura-fits-statement-${fromDate}-to-${toDate}.csv`,
          filters: [{ name: "CSV File", extensions: ["csv"] }],
        });
        if (!filePath) return;
        const content = buildCsv(rows);
        await writeFile(filePath, new TextEncoder().encode(content));
        addToast("Exported as .CSV", "success");
      } else {
        const filePath = await save({
          defaultPath: `aura-fits-statement-${fromDate}-to-${toDate}.xlsx`,
          filters: [{ name: "Excel Workbook", extensions: ["xlsx"] }],
        });
        if (!filePath) return;
        addToast("Building Excel file…", "success");
        const xlsxBytes = await buildExcelXlsx(rows);
        await writeFile(filePath, new Uint8Array(xlsxBytes));
        addToast("Exported as .XLSX ✓", "success");
      }
    } catch (e) {
      addToast("Export failed: " + (e?.message || String(e)), "error");
    }
  };

  const tabBtn = (id, label, badge) => (
    <button onClick={() => setTab(id)} style={{
      padding: "8px 20px", borderRadius: 8, fontSize: 13, cursor: "pointer", fontWeight: tab === id ? 600 : 400,
      background: tab === id ? `${GOLD}20` : "transparent",
      border: `1px solid ${tab === id ? GOLD : "#2a2a2a"}`,
      color: tab === id ? GOLD : "#666", display: "flex", alignItems: "center", gap: 8,
    }}>
      {label}
      {badge > 0 && <span style={{ background: "#c05f5f", color: "#fff", borderRadius: 10, fontSize: 10, fontWeight: 700, padding: "1px 7px" }}>{badge}</span>}
    </button>
  );

  return (
    <div className="fade-in">
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24 }}>
        <div>
          <h1 style={{ fontFamily: "'Cormorant Garamond', serif", fontSize: 28, fontWeight: 300, color: "#E8E4D9" }}>Statement</h1>
          <p style={{ color: "#555", fontSize: 13, marginTop: 4 }}>Sales records & pending balances</p>
        </div>
        <div style={{ position: "relative" }}>
          <GoldButton onClick={() => setShowExportMenu(m => !m)}>Export ▾</GoldButton>
          {showExportMenu && (
            <div ref={exportRef} style={{ position: "absolute", right: 0, top: "110%", background: "#1a1a1a", border: "1px solid #2a2a2a", borderRadius: 10, zIndex: 200, minWidth: 210, overflow: "hidden", boxShadow: "0 8px 32px rgba(0,0,0,0.5)" }}>
              <button onClick={() => doExport("pdf")} style={{ width: "100%", background: "none", border: "none", color: "#E8E4D9", padding: "12px 18px", textAlign: "left", cursor: "pointer", fontSize: 13, display: "flex", alignItems: "center", gap: 10 }}
                onMouseEnter={e => e.currentTarget.style.background = "#222"}
                onMouseLeave={e => e.currentTarget.style.background = "none"}>
                <span style={{ fontSize: 18 }}>📄</span>
                <div><p style={{ fontWeight: 600, marginBottom: 2 }}>PDF (.pdf)</p><p style={{ fontSize: 11, color: "#555" }}>Printable wide-format report</p></div>
              </button>
              <div style={{ height: 1, background: "#222" }} />
              <button onClick={() => doExport("xls")} style={{ width: "100%", background: "none", border: "none", color: "#E8E4D9", padding: "12px 18px", textAlign: "left", cursor: "pointer", fontSize: 13, display: "flex", alignItems: "center", gap: 10 }}
                onMouseEnter={e => e.currentTarget.style.background = "#222"}
                onMouseLeave={e => e.currentTarget.style.background = "none"}>
                <span style={{ fontSize: 18 }}>📊</span>
                <div><p style={{ fontWeight: 600, marginBottom: 2 }}>Excel (.xlsx)</p><p style={{ fontSize: 11, color: "#555" }}>For Windows — Microsoft Excel</p></div>
              </button>
              <div style={{ height: 1, background: "#222" }} />
              <button onClick={() => doExport("csv")} style={{ width: "100%", background: "none", border: "none", color: "#E8E4D9", padding: "12px 18px", textAlign: "left", cursor: "pointer", fontSize: 13, display: "flex", alignItems: "center", gap: 10 }}
                onMouseEnter={e => e.currentTarget.style.background = "#222"}
                onMouseLeave={e => e.currentTarget.style.background = "none"}>
                <span style={{ fontSize: 18 }}>🍎</span>
                <div><p style={{ fontWeight: 600, marginBottom: 2 }}>CSV (.csv)</p><p style={{ fontSize: 11, color: "#555" }}>For Mac — Numbers / Excel</p></div>
              </button>
            </div>
          )}
        </div>
      </div>

      {/* Tabs */}
      <div style={{ display: "flex", gap: 8, marginBottom: 20 }}>
        {tabBtn("sales", "All Sales")}
        {tabBtn("pending", "Pending Balances", pending.length)}
      </div>

      {tab === "sales" && (<>
        <Card style={{ marginBottom: 16 }}>
          <div style={{ display: "grid", gridTemplateColumns: "180px 180px 1fr auto", gap: 12, alignItems: "end" }}>
            <FormField label="From Date"><input type="date" value={fromDate} onChange={e => setFromDate(e.target.value)} /></FormField>
            <FormField label="To Date"><input type="date" value={toDate} onChange={e => setToDate(e.target.value)} /></FormField>
            <FormField label="Search Customer"><input value={search} onChange={e => setSearch(e.target.value)} placeholder="Name or phone..." /></FormField>
            <GoldButton onClick={load} style={{ marginBottom: 16, width: 120 }}>{loading ? "Loading..." : "Apply"}</GoldButton>
          </div>
        </Card>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14, marginBottom: 16 }}>
          <StatCard label="Bills" value={saleIds.size} sub="in selected range" />
          <StatCard label="Sales Total" value={rupees(totalSales)} color="#E8E4D9" />
          <StatCard label="Profit" value={rupees(totalProfit)} color={totalProfit >= 0 ? "#5fa05f" : "#c05f5f"} />
          <StatCard label="Pending Balance" value={rupees(totalBalance)} color={totalBalance > 0 ? "#c05f5f" : "#5fa05f"} />
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          {filteredSales.length === 0
            ? <Card><p style={{ color: "#555", fontSize: 13, padding: "20px 0" }}>{search ? `No sales found for "${search}"` : "No sales found for this date range."}</p></Card>
            : filteredSales.map(sale => (
              <Card key={sale.sale_id} style={{ padding: 0, overflow: "hidden" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", padding: "14px 18px", borderBottom: "1px solid #1a1a1a", background: "#111" }}>
                  <div>
                    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 4 }}>
                      <span style={{ fontSize: 12, color: GOLD, fontWeight: 600 }}>Bill #{sale.sale_id}</span>
                      <Badge color={sale.payment === "Cash" ? "#5fa05f" : sale.payment === "UPI" ? "#5f8fa0" : GOLD_DIM}>{sale.payment}</Badge>
                      {sale.balance > 0 && <Badge color="#c05f5f">⚠ Balance: {rupees(sale.balance)}</Badge>}
                    </div>
                    <p style={{ fontSize: 12, color: "#555" }}>{sale.created_at?.slice(0, 16) || ""}</p>
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <p style={{ fontSize: 18, color: GOLD, fontWeight: 700 }}>{rupees(sale.sale_total)}</p>
                    <div style={{ display: "flex", gap: 6, marginTop: 6, justifyContent: "flex-end" }}>
                      <button onClick={async () => {
                        try {
                          const receiptData = {
                            saleId: sale.sale_id,
                            date: sale.created_at?.slice(0, 16) || "",
                            customerName: sale.customer_name || "Walk-in Customer",
                            customerPhone: sale.customer_phone || "",
                            payment: sale.payment,
                            items: sale.items.map(i => ({ name: i.product_name, qty: i.qty, lineTotal: i.line_total })),
                            total: sale.sale_total,
                            amountPaid: sale.amount_paid || 0,
                            balance: sale.balance || 0,
                          };
                          const doc = await generateReceiptPdf(receiptData);
                          doc.save(`receipt-${sale.sale_id}.pdf`);
                          addToast("Receipt downloaded ✓", "success");
                        } catch(e) { addToast("PDF failed: " + (e?.message||String(e)), "error"); }
                      }} style={{ background: "none", border: "1px solid #2a2a2a", color: GOLD, borderRadius: 6, padding: "4px 12px", cursor: "pointer", fontSize: 11 }}>⬇ Receipt</button>
                      <button onClick={() => openEdit(sale)} style={{ background: "none", border: "1px solid #2a2a2a", color: "#888", borderRadius: 6, padding: "4px 12px", cursor: "pointer", fontSize: 11 }}>Edit</button>
                      <button onClick={() => setConfirmDeleteId(sale.sale_id)} style={{ background: "none", border: "1px solid #3d1a1a", color: "#c05f5f", borderRadius: 6, padding: "4px 12px", cursor: "pointer", fontSize: 11 }}>Delete</button>
                    </div>
                  </div>
                </div>
                <div style={{ display: "flex", gap: 24, padding: "10px 18px", borderBottom: "1px solid #161616", background: "#0f0f0f" }}>
                  <div><span style={{ fontSize: 10, color: "#555", textTransform: "uppercase", letterSpacing: "0.08em" }}>Customer</span><p style={{ fontSize: 13, color: "#C8C4B8", marginTop: 2 }}>{sale.customer_name || "Walk-in"}</p></div>
                  {sale.customer_phone && <div><span style={{ fontSize: 10, color: "#555", textTransform: "uppercase", letterSpacing: "0.08em" }}>Phone</span><p style={{ fontSize: 13, color: "#C8C4B8", marginTop: 2 }}>📞 {sale.customer_phone}</p></div>}
                  {sale.amount_paid > 0 && <div><span style={{ fontSize: 10, color: "#555", textTransform: "uppercase", letterSpacing: "0.08em" }}>Paid</span><p style={{ fontSize: 13, color: "#5fa05f", marginTop: 2 }}>{rupees(sale.amount_paid)}</p></div>}
                  {sale.balance > 0 && <div><span style={{ fontSize: 10, color: "#555", textTransform: "uppercase", letterSpacing: "0.08em" }}>Balance Due</span><p style={{ fontSize: 13, color: "#c05f5f", fontWeight: 700, marginTop: 2 }}>{rupees(sale.balance)}</p></div>}
                </div>
                <div style={{ padding: "0 18px" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                    <thead><tr>{["Product","Category","Size","Color","Qty","Cost","Sell","Total","Profit"].map(h=><th key={h} style={{ textAlign:"left", padding:"8px 6px", color:"#444", fontWeight:500, fontSize:10, letterSpacing:"0.08em", textTransform:"uppercase", borderBottom:"1px solid #1a1a1a" }}>{h}</th>)}</tr></thead>
                    <tbody>
                      {sale.items.map((item,i)=>(
                        <tr key={i} style={{ borderBottom:"1px solid #111" }}>
                          <td style={{ padding:"8px 6px", color:"#C8C4B8", fontWeight:500 }}>{item.product_name}</td>
                          <td style={{ padding:"8px 6px" }}><Badge>{item.category}</Badge></td>
                          <td style={{ padding:"8px 6px", color:"#666" }}>{item.size}</td>
                          <td style={{ padding:"8px 6px", color:"#666" }}>{item.color}</td>
                          <td style={{ padding:"8px 6px", color:"#888" }}>{item.qty}</td>
                          <td style={{ padding:"8px 6px", color:"#666" }}>{rupees(item.cost)}</td>
                          <td style={{ padding:"8px 6px", color:"#C8C4B8" }}>{rupees(item.price)}</td>
                          <td style={{ padding:"8px 6px", color:GOLD, fontWeight:600 }}>{rupees(item.line_total)}</td>
                          <td style={{ padding:"8px 6px", color:item.profit>=0?"#5fa05f":"#c05f5f", fontWeight:500 }}>{rupees(item.profit)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </Card>
            ))
          }
        </div>
      </>)}

      {tab === "pending" && (
        <div>
          {pending.length === 0
            ? <Card><p style={{ color: "#5fa05f", fontSize: 14, padding: "28px 0", textAlign: "center" }}>✓ No pending balances — all customers are settled!</p></Card>
            : <>
              <div style={{ background: "#2d1a1a", border: "1px solid #5c2d2d", borderRadius: 10, padding: "12px 18px", marginBottom: 16, display: "flex", alignItems: "center", gap: 10 }}>
                <span style={{ fontSize: 18 }}>⚠</span>
                <div>
                  <p style={{ fontSize: 13, color: "#c05f5f", fontWeight: 600 }}>{pending.length} customer{pending.length > 1 ? "s" : ""} with pending balance</p>
                  <p style={{ fontSize: 12, color: "#944" }}>Total outstanding: {rupees(pending.reduce((s,p)=>s+p.balance,0))}</p>
                </div>
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {pending.map(p => (
                  <Card key={p.sale_id} style={{ padding: 0, overflow: "hidden", border: "1px solid #3d1a1a" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "14px 18px", background: "#111" }}>
                      <div>
                        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 4 }}>
                          <span style={{ fontSize: 13, color: "#E8E4D9", fontWeight: 600 }}>{p.customer_name}</span>
                          {p.customer_phone && <span style={{ fontSize: 12, color: "#555" }}>📞 {p.customer_phone}</span>}
                          <Badge color="#c05f5f">Bill #{p.sale_id}</Badge>
                        </div>
                        <p style={{ fontSize: 11, color: "#555" }}>{p.created_at?.slice(0, 16)} · {p.payment}</p>
                        <p style={{ fontSize: 12, color: "#666", marginTop: 4 }}>{p.items_summary}</p>
                      </div>
                      <div style={{ textAlign: "right" }}>
                        <p style={{ fontSize: 11, color: "#555" }}>Bill Total: {rupees(p.sale_total)}</p>
                        <p style={{ fontSize: 11, color: "#5fa05f" }}>Paid: {rupees(p.amount_paid)}</p>
                        <p style={{ fontSize: 18, color: "#c05f5f", fontWeight: 700, marginTop: 4 }}>{rupees(p.balance)} due</p>
                        <button onClick={() => setPayModal(p)} style={{ marginTop: 8, background: `${GOLD}20`, border: `1px solid ${GOLD}`, color: GOLD, borderRadius: 6, padding: "6px 14px", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>Update Payment</button>
                      </div>
                    </div>
                  </Card>
                ))}
              </div>
            </>
          }
        </div>
      )}

      {/* Update Payment Modal */}
      {payModal && (
        <Modal title={`Update Payment — Bill #${payModal.sale_id}`} onClose={() => setPayModal(null)} width={420}>
          <div style={{ background: "#0d0d0d", borderRadius: 10, padding: 14, marginBottom: 16, border: "1px solid #1e1e1e" }}>
            <p style={{ fontSize: 13, color: "#C8C4B8" }}>{payModal.customer_name} {payModal.customer_phone ? `· 📞 ${payModal.customer_phone}` : ""}</p>
            <p style={{ fontSize: 12, color: "#555", marginTop: 4 }}>{payModal.items_summary}</p>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12, marginTop: 12 }}>
              <div><p style={{ fontSize: 10, color: "#555", textTransform: "uppercase" }}>Bill Total</p><p style={{ fontSize: 14, color: "#E8E4D9", fontWeight: 600, marginTop: 2 }}>{rupees(payModal.sale_total)}</p></div>
              <div><p style={{ fontSize: 10, color: "#555", textTransform: "uppercase" }}>Already Paid</p><p style={{ fontSize: 14, color: "#5fa05f", fontWeight: 600, marginTop: 2 }}>{rupees(payModal.amount_paid)}</p></div>
              <div><p style={{ fontSize: 10, color: "#555", textTransform: "uppercase" }}>Balance Due</p><p style={{ fontSize: 14, color: "#c05f5f", fontWeight: 700, marginTop: 2 }}>{rupees(payModal.balance)}</p></div>
            </div>
          </div>
          <PaymentUpdater sale={payModal} addToast={addToast} onDone={() => { setPayModal(null); loadPending(); load(); }} />
        </Modal>
      )}

      {/* Edit Sale Modal */}
      {editSale && editForm && (
        <Modal title={`Edit Bill #${editSale.sale_id}`} onClose={() => { setEditSale(null); setEditForm(null); }} width={600}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 16 }}>
            <FormField label="Customer Name"><input value={editForm.customerName} onChange={e => setEditForm(f => ({ ...f, customerName: e.target.value }))} /></FormField>
            <FormField label="Customer Phone"><input value={editForm.customerPhone} onChange={e => setEditForm(f => ({ ...f, customerPhone: e.target.value }))} type="tel" /></FormField>
            <FormField label="Payment">
              <select value={editForm.payment} onChange={e => setEditForm(f => ({ ...f, payment: e.target.value }))}>
                {["Cash", "UPI", "Card"].map(m => <option key={m}>{m}</option>)}
              </select>
            </FormField>
            <FormField label="Amount Paid ₹"><input type="number" value={editForm.amountPaid} onChange={e => setEditForm(f => ({ ...f, amountPaid: e.target.value }))} min={0} /></FormField>
          </div>
          <p style={{ fontSize: 12, color: GOLD, marginBottom: 10, fontWeight: 600 }}>Products in this sale</p>
          {editForm.cart.map((item, i) => (
            <div key={i} style={{ display: "grid", gridTemplateColumns: "1fr 80px 80px auto", gap: 8, marginBottom: 8, alignItems: "center" }}>
              <span style={{ fontSize: 13, color: "#C8C4B8" }}>{item.name}</span>
              <input type="number" value={item.qty ?? 1} min={1} onChange={e => setEditForm(f => ({ ...f, cart: f.cart.map((c,j) => j===i ? {...c, qty: Number(e.target.value)} : c) }))} placeholder="Qty" style={{ textAlign: "center" }} />
              <input type="number" value={item.sellPrice ?? item.price} onChange={e => setEditForm(f => ({ ...f, cart: f.cart.map((c,j) => j===i ? {...c, sellPrice: e.target.value} : c) }))} placeholder="Price" style={{ textAlign: "right" }} />
              <button onClick={() => setEditForm(f => ({ ...f, cart: f.cart.filter((_,j) => j!==i) }))} style={{ background: "none", border: "1px solid #3d1a1a", color: "#c05f5f", borderRadius: 6, padding: "6px 10px", cursor: "pointer", fontSize: 11 }}>✕</button>
            </div>
          ))}
          <FormField label="Add Product">
            <select onChange={e => {
              const p = products.find(p => p.id === Number(e.target.value));
              if (!p) return;
              setEditForm(f => ({ ...f, cart: [...f.cart, { id: p.id, name: p.name, category: p.category, size: p.size, color: p.color, cost: p.cost, qty: 1, price: p.price, sellPrice: String(p.price) }] }));
              e.target.value = "";
            }} defaultValue="">
              <option value="">— Select product to add —</option>
              {products.map(p => <option key={p.id} value={p.id}>{p.name} (Stock: {p.stock})</option>)}
            </select>
          </FormField>
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 16 }}>
            <GoldButton variant="outline" onClick={() => { setEditSale(null); setEditForm(null); }}>Cancel</GoldButton>
            <GoldButton onClick={saveEdit}>Save Changes</GoldButton>
          </div>
        </Modal>
      )}

      {confirmDeleteId !== null && (
        <Modal title="Delete Sale" onClose={() => setConfirmDeleteId(null)} width={380}>
          <p style={{ fontSize: 13, color: "#888", marginBottom: 8 }}>Delete Bill <strong style={{ color: "#E8E4D9" }}>#{confirmDeleteId}</strong>?</p>
          <p style={{ fontSize: 12, color: "#5fa05f", marginBottom: 24 }}>✓ All products will be returned to inventory automatically.</p>
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
            <GoldButton variant="outline" onClick={() => setConfirmDeleteId(null)}>Cancel</GoldButton>
            <button onClick={() => doDelete(confirmDeleteId)} style={{ background: "#2d1a1a", border: "1px solid #5c2d2d", color: "#c05f5f", borderRadius: 8, padding: "10px 22px", cursor: "pointer", fontSize: 14 }}>Delete & Restore Stock</button>
          </div>
        </Modal>
      )}
    </div>
  );
}

// Payment updater sub-component
function PaymentUpdater({ sale, addToast, onDone }) {
  const [newPayment, setNewPayment] = useState("");
  const remaining = sale.balance;
  const newPaid = Number(newPayment) || 0;
  const newBalance = Math.max(0, remaining - newPaid);
  const totalPaid = sale.amount_paid + newPaid;

  const submit = async () => {
    if (!newPayment || newPaid <= 0) { addToast("Enter amount received", "error"); return; }
    if (newPaid > remaining) { addToast("Amount exceeds balance due", "error"); return; }
    try {
      await window.__TAURI__.core.invoke("update_balance", { sale_id: sale.sale_id, amount_paid: totalPaid, balance: newBalance });
      addToast(newBalance === 0 ? "Balance fully cleared ✓" : `Payment recorded. Remaining: ₹${newBalance.toLocaleString()}`, "success");
      onDone();
    } catch (e) { addToast("Failed: " + (e?.message || String(e)), "error"); }
  };

  return (
    <div>
      <FormField label="Amount Received Now ₹">
        <input type="number" value={newPayment} onChange={e => setNewPayment(e.target.value)} placeholder={`Up to ₹${remaining.toLocaleString()}`} min={0} max={remaining} autoFocus />
      </FormField>
      {newPayment && (
        <div style={{ background: newBalance === 0 ? "#1a2d1a" : "#2d1a1a", border: `1px solid ${newBalance === 0 ? "#2d5c2d" : "#5c2d2d"}`, borderRadius: 8, padding: "10px 14px", marginBottom: 16 }}>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, marginBottom: 4 }}>
            <span style={{ color: "#888" }}>Total Paid After</span>
            <span style={{ color: "#5fa05f", fontWeight: 600 }}>{rupees(totalPaid)}</span>
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 14 }}>
            <span style={{ color: "#888" }}>Remaining Balance</span>
            <span style={{ color: newBalance === 0 ? "#5fa05f" : "#c05f5f", fontWeight: 700 }}>{newBalance === 0 ? "✓ Cleared" : rupees(newBalance)}</span>
          </div>
        </div>
      )}
      <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
        <GoldButton onClick={submit}>Record Payment</GoldButton>
      </div>
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
  useEffect(() => {
    try { const s = localStorage.getItem("aura_settings"); if (s) setSettings(JSON.parse(s)); } catch {}
  }, []);
  const save = () => {
    try { localStorage.setItem("aura_settings", JSON.stringify(settings)); } catch {}
    addToast("Settings saved", "success");
  };

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

// ─── PENDING BALANCES ─────────────────────────────────────────────────────────
function PendingBalancesPage({ addToast, onRefreshPending }) {
  const [pending, setPending] = useState([]);
  const [loading, setLoading] = useState(false);
  const [updateModal, setUpdateModal] = useState(null); // { sale }
  const [newPayment, setNewPayment] = useState("");

  const load = async () => {
    setLoading(true);
    try { setPending(await window.db.getPendingBalances()); }
    catch (e) { addToast("Failed to load pending balances", "error"); }
    finally { setLoading(false); }
  };

  useEffect(() => { load(); }, []);

  const openUpdate = (sale) => {
    setUpdateModal(sale);
    setNewPayment("");
  };

  const savePayment = async () => {
    if (!updateModal) return;
    const extra = Number(newPayment) || 0;
    if (extra <= 0) { addToast("Enter a valid payment amount", "error"); return; }
    const newPaid = (updateModal.amount_paid || 0) + extra;
    const newBalance = Math.max(0, (updateModal.sale_total || 0) - newPaid);
    try {
      await window.__TAURI__.core.invoke("update_balance", { sale_id: updateModal.sale_id, amount_paid: newPaid, balance: newBalance });
      addToast(newBalance === 0 ? "Balance fully cleared ✓" : `Balance updated. Remaining: ₹${newBalance.toLocaleString()}`, "success");
      setUpdateModal(null); setNewPayment(""); load(); onRefreshPending?.();
    } catch (e) { addToast("Failed to update: " + (e?.message || String(e)), "error"); }
  };

  const totalPending = pending.reduce((s, p) => s + p.balance, 0);

  return (
    <div className="fade-in">
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24 }}>
        <div>
          <h1 style={{ fontFamily: "'Cormorant Garamond', serif", fontSize: 28, fontWeight: 300, color: "#E8E4D9" }}>Pending Balances</h1>
          <p style={{ color: "#555", fontSize: 13, marginTop: 4 }}>Customers who owe you money</p>
        </div>
        <GoldButton variant="outline" onClick={load}>{loading ? "Loading..." : "Refresh"}</GoldButton>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 14, marginBottom: 20 }}>
        <StatCard label="Pending Customers" value={pending.length} color="#c05f5f" icon="⚠" />
        <StatCard label="Total Due Amount" value={rupees(totalPending)} color="#c05f5f" />
        <StatCard label="Status" value={pending.length === 0 ? "All Clear ✓" : "Action Needed"} color={pending.length === 0 ? "#5fa05f" : "#c08060"} />
      </div>

      {pending.length === 0 ? (
        <Card>
          <div style={{ textAlign: "center", padding: "40px 0" }}>
            <p style={{ fontSize: 32, marginBottom: 12 }}>✓</p>
            <p style={{ color: "#5fa05f", fontSize: 15, fontWeight: 500 }}>No pending balances</p>
            <p style={{ color: "#555", fontSize: 13, marginTop: 6 }}>All customers are fully paid up</p>
          </div>
        </Card>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {pending.map(sale => (
            <Card key={sale.sale_id} style={{ borderColor: "#3d1a1a" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                <div style={{ flex: 1 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6 }}>
                    <span style={{ fontSize: 13, color: "#E8E4D9", fontWeight: 600 }}>{sale.customer_name}</span>
                    {sale.customer_phone && <span style={{ fontSize: 12, color: "#555" }}>📞 {sale.customer_phone}</span>}
                    <Badge color="#c05f5f">Bill #{sale.sale_id}</Badge>
                  </div>
                  <p style={{ fontSize: 12, color: "#555", marginBottom: 6 }}>{sale.created_at?.slice(0, 16)} · {sale.payment}</p>
                  <p style={{ fontSize: 12, color: "#666", marginBottom: 10 }}>{sale.items_summary}</p>
                  <div style={{ display: "flex", gap: 20 }}>
                    <div>
                      <p style={{ fontSize: 10, color: "#555", textTransform: "uppercase", letterSpacing: "0.08em" }}>Bill Total</p>
                      <p style={{ fontSize: 14, color: "#C8C4B8", fontWeight: 500 }}>{rupees(sale.sale_total)}</p>
                    </div>
                    <div>
                      <p style={{ fontSize: 10, color: "#555", textTransform: "uppercase", letterSpacing: "0.08em" }}>Paid</p>
                      <p style={{ fontSize: 14, color: "#5fa05f", fontWeight: 500 }}>{rupees(sale.amount_paid)}</p>
                    </div>
                    <div>
                      <p style={{ fontSize: 10, color: "#c05f5f", textTransform: "uppercase", letterSpacing: "0.08em" }}>Balance Due</p>
                      <p style={{ fontSize: 18, color: "#c05f5f", fontWeight: 700 }}>{rupees(sale.balance)}</p>
                    </div>
                  </div>
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 8, marginLeft: 20 }}>
                  <GoldButton onClick={() => openUpdate(sale)}>Collect Payment</GoldButton>
                  <button onClick={async () => {
                    const reminderReceipt = {
                      saleId: sale.sale_id,
                      date: sale.created_at?.slice(0, 16) || "",
                      customerName: sale.customer_name,
                      customerPhone: sale.customer_phone || "",
                      payment: sale.payment,
                      items: (sale.items_summary || "").split(",").map(s => ({ name: s.trim(), qty: 1, lineTotal: 0 })),
                      subtotal: sale.sale_total,
                      discount: 0,
                      discountAmt: 0,
                      total: sale.sale_total,
                      amountPaid: sale.amount_paid,
                      balance: sale.balance,
                    };
                    try {
                      const doc = await generateReceiptPdf(reminderReceipt);
                      doc.save(`receipt-bill-${sale.sale_id}.pdf`);
                      addToast("Receipt downloaded ✓", "success");
                    } catch (e) {
                      addToast("Failed: " + (e?.message || String(e)), "error");
                    }
                  }} style={{ background: "none", border: "1px solid #2a2a2a", color: "#888", borderRadius: 8, padding: "8px 14px", cursor: "pointer", fontSize: 12 }}>
                    ⬇ Download Receipt
                  </button>
                </div>
              </div>
            </Card>
          ))}
        </div>
      )}

      {updateModal && (
        <Modal title={`Collect Payment — Bill #${updateModal.sale_id}`} onClose={() => setUpdateModal(null)} width={420}>
          <div style={{ background: "#0d0d0d", borderRadius: 10, padding: 14, marginBottom: 18 }}>
            <p style={{ fontSize: 13, color: "#C8C4B8", fontWeight: 600, marginBottom: 4 }}>{updateModal.customer_name}</p>
            {updateModal.customer_phone && <p style={{ fontSize: 12, color: "#555", marginBottom: 8 }}>📞 {updateModal.customer_phone}</p>}
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
              <span style={{ fontSize: 12, color: "#666" }}>Bill Total</span>
              <span style={{ fontSize: 13, color: "#C8C4B8" }}>{rupees(updateModal.sale_total)}</span>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
              <span style={{ fontSize: 12, color: "#666" }}>Already Paid</span>
              <span style={{ fontSize: 13, color: "#5fa05f" }}>{rupees(updateModal.amount_paid)}</span>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", borderTop: "1px solid #222", paddingTop: 8 }}>
              <span style={{ fontSize: 13, color: "#c05f5f", fontWeight: 600 }}>Balance Due</span>
              <span style={{ fontSize: 16, color: "#c05f5f", fontWeight: 700 }}>{rupees(updateModal.balance)}</span>
            </div>
          </div>
          <FormField label="Amount Received Now (₹)">
            <input type="number" value={newPayment} onChange={e => setNewPayment(e.target.value)} placeholder={`Max: ${updateModal.balance}`} autoFocus min={1} max={updateModal.balance} />
          </FormField>
          {newPayment && Number(newPayment) > 0 && (
            <div style={{ background: "#1a2d1a", border: "1px solid #2d5c2d", borderRadius: 8, padding: "10px 14px", marginBottom: 14 }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                <span style={{ fontSize: 12, color: "#888" }}>New Total Paid</span>
                <span style={{ fontSize: 13, color: "#5fa05f" }}>{rupees((updateModal.amount_paid || 0) + Number(newPayment))}</span>
              </div>
              <div style={{ display: "flex", justifyContent: "space-between" }}>
                <span style={{ fontSize: 12, color: "#888" }}>Remaining Balance</span>
                <span style={{ fontSize: 13, color: Math.max(0, updateModal.balance - Number(newPayment)) === 0 ? "#5fa05f" : "#c05f5f", fontWeight: 600 }}>
                  {Math.max(0, updateModal.balance - Number(newPayment)) === 0 ? "✓ Fully Paid!" : rupees(Math.max(0, updateModal.balance - Number(newPayment)))}
                </span>
              </div>
            </div>
          )}
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
            <GoldButton variant="outline" onClick={() => setUpdateModal(null)}>Cancel</GoldButton>
            <GoldButton onClick={savePayment}>Save Payment</GoldButton>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ─── PERSONAL EXPENSES ───────────────────────────────────────────────────────
function PersonalExpensesPage({ addToast }) {
  const [expenses, setExpenses] = useState([]);
  const [showModal, setShowModal] = useState(false);
  const [editExpense, setEditExpense] = useState(null);
  const [personFilter, setPersonFilter] = useState("All");
  const [form, setForm] = useState({ person: "", name: "", amount: "", date: todayDate(), notes: "" });

  const load = () => window.__TAURI__.core.invoke("get_personal_expenses").then(setExpenses).catch(() => {});
  useEffect(() => { load(); }, []);

  const persons = ["All", ...Array.from(new Set(expenses.map(e => e.person).filter(Boolean)))];
  const filtered = personFilter === "All" ? expenses : expenses.filter(e => e.person === personFilter);
  const total = filtered.reduce((s, e) => s + e.amount, 0);
  const byPerson = expenses.reduce((acc, e) => { acc[e.person] = (acc[e.person] || 0) + e.amount; return acc; }, {});

  const openAdd = () => { setForm({ person: "", name: "", amount: "", date: todayDate(), notes: "" }); setEditExpense(null); setShowModal(true); };
  const openEdit = (i) => { setForm({ ...filtered[i] }); setEditExpense(filtered[i]); setShowModal(true); };
  const doDelete = async (i) => { await window.__TAURI__.core.invoke("delete_personal_expense", { id: filtered[i].id }); addToast("Deleted", "info"); load(); };

  const save = async () => {
    if (!form.person || !form.name || !form.amount) { addToast("Fill all required fields", "error"); return; }
    const e = { ...form, amount: Number(form.amount) };
    if (editExpense) { await window.__TAURI__.core.invoke("update_personal_expense", { expense: { ...e, id: editExpense.id } }); addToast("Updated", "success"); }
    else { await window.__TAURI__.core.invoke("add_personal_expense", { expense: e }); addToast("Added", "success"); }
    setShowModal(false); load();
  };

  const today = todayDate();
  const thisMonth = today.slice(0, 7);
  const uniquePersons = Array.from(new Set(expenses.map(e => e.person).filter(Boolean)));

  return (
    <div className="fade-in">
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24 }}>
        <div>
          <h1 style={{ fontFamily: "'Cormorant Garamond', serif", fontSize: 28, fontWeight: 300, color: "#E8E4D9" }}>Personal Expenses</h1>
          <p style={{ color: "#555", fontSize: 13, marginTop: 4 }}>Track individual spending for each person</p>
        </div>
        <GoldButton onClick={openAdd}>+ Add Expense</GoldButton>
      </div>

      {/* Grand total */}
      {expenses.length > 0 && (
        <div style={{ marginBottom: 16 }}>
          <Card style={{ border: `1px solid ${GOLD_DIM}`, background: "#111" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div>
                <p style={{ fontSize: 11, color: "#555", textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 4 }}>Total Personal Expenses</p>
                <p style={{ fontSize: 28, color: "#c08060", fontWeight: 700, fontFamily: "'Cormorant Garamond', serif" }}>{rupees(expenses.reduce((s, e) => s + e.amount, 0))}</p>
              </div>
              <div style={{ textAlign: "right" }}>
                <p style={{ fontSize: 11, color: "#555", marginBottom: 4 }}>{uniquePersons.length} people · {expenses.length} transactions</p>
                <p style={{ fontSize: 12, color: "#666" }}>For display & record only</p>
              </div>
            </div>
          </Card>
        </div>
      )}

      {/* Person total cards — sorted by name, click to view transactions */}
      {uniquePersons.length > 0 && (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(180px, 1fr))", gap: 12, marginBottom: 20 }}>
          {[...uniquePersons].sort((a, b) => a.localeCompare(b)).map(person => {
            const allTimeAmt = expenses.filter(e => e.person === person).reduce((s, e) => s + e.amount, 0);
            const txnCount = expenses.filter(e => e.person === person).length;
            const isActive = personFilter === person;
            return (
              <Card key={person}
                style={{ cursor: "pointer", border: `1px solid ${isActive ? GOLD : GOLD_DIM}`, background: isActive ? `${GOLD}08` : "#141414", transition: "border-color 0.2s" }}
                onClick={() => setPersonFilter(isActive ? "All" : person)}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                  <span style={{ fontSize: 14, color: isActive ? GOLD : "#E8E4D9", fontWeight: 600 }}>{person}</span>
                  <span style={{ fontSize: 10, color: "#555" }}>{txnCount} txns</span>
                </div>
                <p style={{ fontSize: 20, color: "#c08060", fontWeight: 700, fontFamily: "'Cormorant Garamond', serif" }}>{rupees(allTimeAmt)}</p>
                <p style={{ fontSize: 10, color: "#555", marginTop: 4 }}>{isActive ? "▲ Showing transactions" : "Click to view"}</p>
              </Card>
            );
          })}
        </div>
      )}

      {/* Filter pills */}
      <div style={{ display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap", alignItems: "center" }}>
        {["All", ...uniquePersons].map(p => (
          <button key={p} onClick={() => setPersonFilter(p)} style={{
            padding: "6px 16px", borderRadius: 20, fontSize: 12, cursor: "pointer",
            background: personFilter === p ? `${GOLD}20` : "transparent",
            border: `1px solid ${personFilter === p ? GOLD : "#2a2a2a"}`,
            color: personFilter === p ? GOLD : "#666",
          }}>{p}{p !== "All" && ` · ${rupees(byPerson[p] || 0)}`}</button>
        ))}
      </div>

      {/* Person detail header when filtered */}
      {personFilter !== "All" && (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginBottom: 16 }}>
          <StatCard label="Today" value={rupees(expenses.filter(e => e.person === personFilter && e.date === today).reduce((s,e) => s+e.amount, 0))} color="#c08060" />
          <StatCard label="This Month" value={rupees(expenses.filter(e => e.person === personFilter && (e.date||"").slice(0,7) === thisMonth).reduce((s,e) => s+e.amount, 0))} color="#c08060" />
          <StatCard label="All Time" value={rupees(byPerson[personFilter] || 0)} color="#E8E4D9" />
        </div>
      )}

      <Card>
        <Table
          headers={["Person", "Expense", "Amount", "Date", "Notes"]}
          rows={filtered.map(e => [
            <Badge color="#8A6D2E">{e.person}</Badge>,
            <span style={{ color: "#E8E4D9", fontWeight: 500 }}>{e.name}</span>,
            <span style={{ color: "#c08060", fontWeight: 600 }}>{rupees(e.amount)}</span>,
            e.date,
            <span style={{ color: "#555", fontSize: 12 }}>{e.notes}</span>,
          ])}
          onEdit={openEdit}
          onDelete={doDelete}
        />
      </Card>

      {showModal && (
        <Modal title={editExpense ? "Edit Personal Expense" : "Add Personal Expense"} onClose={() => setShowModal(false)}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <FormField label="Person Name *"><input value={form.person} onChange={e => setForm(f => ({ ...f, person: e.target.value }))} placeholder="e.g. Junaid, Partner" /></FormField>
            <FormField label="Expense Name *"><input value={form.name} onChange={e => setForm(f => ({ ...f, name: e.target.value }))} placeholder="e.g. Food, Travel" /></FormField>
            <FormField label="Amount (₹) *"><input type="number" value={form.amount} onChange={e => setForm(f => ({ ...f, amount: e.target.value }))} /></FormField>
            <FormField label="Date"><input type="date" value={form.date} onChange={e => setForm(f => ({ ...f, date: e.target.value }))} /></FormField>
          </div>
          <FormField label="Notes"><textarea value={form.notes} onChange={e => setForm(f => ({ ...f, notes: e.target.value }))} rows={2} style={{ resize: "none" }} /></FormField>
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
            <GoldButton variant="outline" onClick={() => setShowModal(false)}>Cancel</GoldButton>
            <GoldButton onClick={save}>{editExpense ? "Save" : "Add"}</GoldButton>
          </div>
        </Modal>
      )}
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
  const [pendingCount, setPendingCount] = useState(0);

  useEffect(() => {
    const style = document.createElement("style");
    style.textContent = globalStyles;
    document.head.appendChild(style);
    return () => document.head.removeChild(style);
  }, []);

  const refreshPending = () => {
    window.db?.getPendingBalances?.().then(p => setPendingCount(p.length)).catch(() => {});
  };

  useEffect(() => {
    if (loggedIn) { refreshPending(); const t = setInterval(refreshPending, 30000); return () => clearInterval(t); }
  }, [loggedIn]);

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
    statement: <StatementPage addToast={addToast} onRefreshPending={refreshPending} />,
    pending: <PendingBalancesPage addToast={addToast} onRefreshPending={refreshPending} />,
    expenses: <ExpensesPage addToast={addToast} />,
    personal: <PersonalExpensesPage addToast={addToast} />,
    settings: <SettingsPage addToast={addToast} onLogout={() => setLoggedIn(false)} />,
  };

  return (
    <div style={{ minHeight: "100vh", background: "#0A0A0A", fontFamily: "'DM Sans', sans-serif" }}>
      <Toast toasts={toasts} />
      <Sidebar active={activePage} setActive={setActivePage} collapsed={collapsed} setCollapsed={setCollapsed} userName={userName} pendingCount={pendingCount} />
      <div style={{ position: "fixed", top: 0, left: sideW, right: 0, height: 52, background: "#0D0D0D", borderBottom: "1px solid #1a1a1a", display: "flex", alignItems: "center", justifyContent: "space-between", padding: "0 28px", zIndex: 99, transition: "left 0.25s ease" }}>
        <p style={{ fontSize: 12, color: "#444", letterSpacing: "0.08em" }}>{NAV_ITEMS.find(n => n.id === activePage)?.label?.toUpperCase()}</p>
        <div style={{ display: "flex", alignItems: "center", gap: 20 }}>
          {pendingCount > 0 && (
            <div onClick={() => setActivePage("pending")} style={{ display: "flex", alignItems: "center", gap: 6, background: "#2d1a1a", border: "1px solid #5c2d2d", borderRadius: 8, padding: "4px 10px", cursor: "pointer" }}>
              <span style={{ color: "#c05f5f", fontSize: 12 }}>⚠</span>
              <span style={{ color: "#c05f5f", fontSize: 12, fontWeight: 600 }}>{pendingCount} pending balance{pendingCount > 1 ? "s" : ""}</span>
            </div>
          )}
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