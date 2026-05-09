import { useState } from "react";

const PASSWORD = "under11";

const CUCINA = [
  { id: "sal",    nome: "PANINO CON SALAMELLA",                          prezzo: 4.5 },
  { id: "salveg", nome: "PANINO CON SALAMELLA E VERDURE GRIGLIATE",      prezzo: 5.0 },
  { id: "wur",    nome: "PANINO CON WURSTEL",                            prezzo: 3.0 },
  { id: "wurcr",  nome: "PANINO CON WURSTEL E CRAUTI",                   prezzo: 3.0 },
  { id: "wurveg", nome: "PANINO CON WURSTEL E VERDURE GRIGLIATE",        prezzo: 3.5 },
  { id: "salme",  nome: "PANINO CON SALAME",                             prezzo: 3.5 },
  { id: "cotto",  nome: "PANINO CON PROSCIUTTO COTTO",                   prezzo: 3.5 },
  { id: "crudo",  nome: "PANINO CON PROSCIUTTO CRUDO",                   prezzo: 3.5 },
  { id: "veg",    nome: "PANINO CON VERDURE GRIGLIATE",                  prezzo: 3.0 },
  { id: "toast",  nome: "TOAST",                                         prezzo: 2.5 },
  { id: "pat",    nome: "PATATINE",                                      prezzo: 2.5 },
];

const BAR = [
  { id: "acqn",    nome: "ACQUA NATURALE",   prezzo: 0.5 },
  { id: "acqf",    nome: "ACQUA FRIZZANTE",  prezzo: 0.5 },
  { id: "coca",    nome: "COCA COLA",        prezzo: 1.5 },
  { id: "cocaz",   nome: "COCA COLA ZERO",   prezzo: 1.5 },
  { id: "fanta",   nome: "FANTA",            prezzo: 1.5 },
  { id: "sprite",  nome: "SPRITE",           prezzo: 1.5 },
  { id: "chin",    nome: "CHINOTTO",         prezzo: 1.5 },
  { id: "tepesca", nome: "TE' ALLA PESCA",   prezzo: 1.5 },
  { id: "telim",   nome: "TE' AL LIMONE",    prezzo: 1.5 },
  { id: "birra",   nome: "BIRRA",            prezzo: 2.5 },
];

const SQUADRA_FIXED = "UNDER 11";
const fmt = (n) => n.toFixed(2).replace(".", ",") + " €";

// ─── Excel generation ─────────────────────────────────────────────────────────

async function loadExcelJS() {
  if (window.ExcelJS) return window.ExcelJS;
  return new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js";
    s.onload = () => resolve(window.ExcelJS);
    s.onerror = reject;
    document.head.appendChild(s);
  });
}

function thin() {
  const s = { style: "thin", color: { argb: "FF000000" } };
  return { top: s, bottom: s, left: s, right: s };
}

function buildSheet(ws, items, titolo, qtys, oggi) {
  ws.getColumn(1).width = 57.71;
  ws.getColumn(2).width = 10;
  ws.getColumn(3).width = 9.14;
  ws.getColumn(4).width = 29.71;
  let r = 1;

  ws.mergeCells(r, 1, r, 4);
  const c1 = ws.getCell(r, 1);
  c1.value = titolo;
  c1.font = { name: "Calibri", bold: true, size: 26 };
  c1.alignment = { horizontal: "center", vertical: "center", wrapText: true };
  c1.border = thin();
  ws.getRow(r).height = 67.5; r++;

  ws.mergeCells(r, 1, r, 4);
  ws.getRow(r).height = 9; r++;

  const cDataL = ws.getCell(r, 1);
  cDataL.value = "DATA";
  cDataL.font = { name: "Calibri", bold: true, size: 16 };
  cDataL.alignment = { horizontal: "right", vertical: "center" };
  cDataL.border = thin();
  ws.mergeCells(r, 2, r, 4);
  const cDataV = ws.getCell(r, 2);
  cDataV.value = oggi;
  cDataV.font = { name: "Calibri", size: 11 };
  cDataV.alignment = { horizontal: "center", vertical: "center" };
  cDataV.border = thin();
  ws.getRow(r).height = 27; r++;

  const cSocL = ws.getCell(r, 1);
  cSocL.value = "SOCIETA' SPORTIVA";
  cSocL.font = { name: "Calibri", bold: true, size: 16 };
  cSocL.alignment = { horizontal: "right", vertical: "center" };
  cSocL.border = thin();
  ws.mergeCells(r, 2, r, 4);
  const cSocV = ws.getCell(r, 2);
  cSocV.value = SQUADRA_FIXED;
  cSocV.font = { name: "Calibri", bold: true, size: 11 };
  cSocV.alignment = { horizontal: "center", vertical: "center" };
  cSocV.border = thin();
  ws.getRow(r).height = 27; r++;

  const cO1 = ws.getCell(r, 1);
  cO1.value = "Orario inizio partita: _______________";
  cO1.font = { name: "Calibri", bold: true, size: 16 };
  cO1.alignment = { horizontal: "left", vertical: "center" };
  cO1.border = thin();
  const cO2 = ws.getCell(r, 2);
  cO2.value = "Orario arrivo a cena: ______________";
  cO2.font = { name: "Calibri", bold: true, size: 16 };
  cO2.alignment = { horizontal: "left", vertical: "center" };
  cO2.border = thin();
  ws.getCell(r, 3).border = thin();
  ws.getCell(r, 4).border = thin();
  ws.getRow(r).height = 27; r++;

  ws.getCell(r, 1).border = thin();
  [["N.", 2], ["€", 3], ["TOTALE € ", 4]].forEach(([h, col]) => {
    const c = ws.getCell(r, col);
    c.value = h;
    c.font = { name: "Calibri", bold: true, size: 12 };
    c.alignment = { horizontal: "center", vertical: "center" };
    c.border = thin();
  });
  ws.getRow(r).height = 15.75; r++;

  items.forEach((item) => {
    const qty = qtys[item.id] || null;
    const cA = ws.getCell(r, 1);
    cA.value = item.nome;
    cA.font = { name: "Calibri", bold: true, size: 14 };
    cA.alignment = { vertical: "center" };
    cA.border = thin();
    const cB = ws.getCell(r, 2);
    cB.value = qty;
    cB.font = { name: "Calibri", size: 11 };
    cB.alignment = { vertical: "center" };
    cB.border = thin();
    const cC = ws.getCell(r, 3);
    cC.value = item.prezzo;
    cC.font = { name: "Calibri", bold: true, size: 14 };
    cC.alignment = { horizontal: "center", vertical: "center" };
    cC.border = thin();
    const cD = ws.getCell(r, 4);
    cD.value = qty ? qty * item.prezzo : null;
    cD.font = { name: "Calibri", size: 11 };
    cD.alignment = { vertical: "center" };
    cD.border = thin();
    ws.getRow(r).height = 21; r++;
  });

  ws.mergeCells(r, 1, r, 3);
  const cTL = ws.getCell(r, 1);
  cTL.value = "TOTALE SQUADRA";
  cTL.font = { name: "Calibri", bold: true, size: 16 };
  cTL.alignment = { horizontal: "right", vertical: "center" };
  cTL.border = { top: thin().top, bottom: thin().bottom, left: thin().left };
  const totale = items.reduce((s, it) => s + (qtys[it.id] || 0) * it.prezzo, 0);
  const cTV = ws.getCell(r, 4);
  cTV.value = totale > 0 ? totale : null;
  cTV.font = { name: "Calibri", size: 11 };
  cTV.alignment = { vertical: "center" };
  cTV.border = thin();
  ws.getRow(r).height = 30; r++;

  ws.mergeCells(r, 1, r, 4);
  ws.getRow(r).height = 9;
}

async function generaExcel(ordini) {
  const ExcelJS = await loadExcelJS();
  const qtyCucina = {};
  const qtyBar = {};
  ordini.forEach((o) => {
    o.cucina.forEach(({ id, qty }) => { qtyCucina[id] = (qtyCucina[id] || 0) + qty; });
    o.bar.forEach(({ id, qty }) => { qtyBar[id] = (qtyBar[id] || 0) + qty; });
  });
  const oggi = new Date().toLocaleDateString("it-IT");
  const wb = new ExcelJS.Workbook();
  buildSheet(wb.addWorksheet("Prenotazioni CENA squadre"), CUCINA, "TORNEI HAITIAMOLI 2026\nPRENOTAZIONE CENA SQUADRA", qtyCucina, oggi);
  buildSheet(wb.addWorksheet("Prenotazioni BAR squadre"), BAR, "TORNEI HAITIAMOLI 2026\nPRENOTAZIONE BAR SQUADRA", qtyBar, oggi);
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `Haitiamoli_2026_Prenotazioni_${oggi.replace(/\//g, "-")}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}

// ─── UI Components ────────────────────────────────────────────────────────────

function Counter({ value, onChange }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 7 }}>
      <button onClick={() => onChange(Math.max(0, value - 1))} style={btnSt("#1a237e", "#fff")}>−</button>
      <span style={{ minWidth: 20, textAlign: "center", fontWeight: 700, fontSize: 15 }}>{value}</span>
      <button onClick={() => onChange(value + 1)} style={btnSt("#f9a825", "#1a237e")}>+</button>
    </div>
  );
}
function btnSt(bg, color) {
  return { background: bg, color, border: "none", borderRadius: 6, width: 28, height: 28, fontSize: 17, cursor: "pointer", fontWeight: 900, lineHeight: 1 };
}

function Section({ title, icon, items, qty, setQty }) {
  const sub = items.reduce((s, it) => s + (qty[it.id] || 0) * it.prezzo, 0);
  return (
    <div style={{ marginBottom: 18 }}>
      <div style={{ background: "#1a237e", color: "#fff", borderRadius: 10, padding: "8px 14px", marginBottom: 8, display: "flex", alignItems: "center", gap: 8, fontSize: 14, fontWeight: 800 }}>
        <span>{icon}</span>{title}
        {sub > 0 && <span style={{ marginLeft: "auto", color: "#f9a825" }}>{fmt(sub)}</span>}
      </div>
      {items.map((it) => (
        <div key={it.id} style={{ display: "flex", alignItems: "center", padding: "6px 4px", borderBottom: "1px solid #e8eaf6", gap: 6 }}>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: "#1a237e", lineHeight: 1.2 }}>{it.nome}</div>
            <div style={{ fontSize: 11, color: "#aaa" }}>{fmt(it.prezzo)}</div>
          </div>
          <Counter value={qty[it.id] || 0} onChange={(v) => setQty((p) => ({ ...p, [it.id]: v }))} />
          <div style={{ minWidth: 54, textAlign: "right", fontSize: 12, fontWeight: 700, color: (qty[it.id] || 0) > 0 ? "#f9a825" : "transparent" }}>
            {fmt((qty[it.id] || 0) * it.prezzo)}
          </div>
        </div>
      ))}
    </div>
  );
}

// ─── Login Screen ─────────────────────────────────────────────────────────────

function LoginScreen({ onAccess }) {
  const [pwd, setPwd] = useState("");
  const [errore, setErrore] = useState(false);

  function handleLogin() {
    if (pwd.trim().toLowerCase() === PASSWORD) {
      onAccess();
    } else {
      setErrore(true);
      setPwd("");
      setTimeout(() => setErrore(false), 2000);
    }
  }

  return (
    <div style={{
      minHeight: "100vh",
      background: "#1a237e",
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      justifyContent: "center",
      padding: 24,
    }}>
      {/* Logo area */}
      <div style={{ textAlign: "center", marginBottom: 36 }}>
        <div style={{ fontSize: 11, letterSpacing: 4, color: "#f9a825", fontWeight: 700, marginBottom: 4 }}>
          POLISPORTIVA ASSISI
        </div>
        <div style={{ fontSize: 42, fontWeight: 900, letterSpacing: 3, color: "#fff", lineHeight: 1 }}>
          HAITIAMOLI
        </div>
        <div style={{ fontSize: 22, fontWeight: 700, color: "#f9a825", marginTop: 2 }}>2026</div>
        <div style={{ fontSize: 13, color: "rgba(255,255,255,0.55)", marginTop: 6, fontStyle: "italic" }}>
          Prenotazione Cibo & Bevande
        </div>
      </div>

      {/* Card login */}
      <div style={{
        background: "#fff",
        borderRadius: 16,
        padding: "28px 24px",
        width: "100%",
        maxWidth: 340,
        boxShadow: "0 8px 32px rgba(0,0,0,0.3)",
      }}>
        <div style={{ fontWeight: 800, color: "#1a237e", fontSize: 15, marginBottom: 16, textAlign: "center" }}>
          🔒 Inserisci la password
        </div>

        <input
          type="password"
          placeholder="Password"
          value={pwd}
          onChange={(e) => { setPwd(e.target.value); setErrore(false); }}
          onKeyDown={(e) => e.key === "Enter" && handleLogin()}
          style={{
            width: "100%",
            padding: "11px 14px",
            borderRadius: 10,
            border: errore ? "2px solid #c62828" : "2px solid #c5cae9",
            fontSize: 15,
            boxSizing: "border-box",
            outline: "none",
            color: "#1a237e",
            textAlign: "center",
            letterSpacing: 2,
            transition: "border 0.2s",
          }}
          autoFocus
        />

        {errore && (
          <div style={{ color: "#c62828", fontSize: 12, textAlign: "center", marginTop: 8, fontWeight: 700 }}>
            ❌ Password errata
          </div>
        )}

        <button
          onClick={handleLogin}
          style={{
            width: "100%",
            marginTop: 14,
            padding: "12px",
            borderRadius: 10,
            border: "none",
            background: "#f9a825",
            color: "#1a237e",
            fontSize: 15,
            fontWeight: 900,
            cursor: "pointer",
            letterSpacing: 0.5,
          }}
        >
          ACCEDI
        </button>
      </div>

      <div style={{ marginTop: 20, fontSize: 11, color: "rgba(255,255,255,0.3)" }}>
        Solo per uso interno squadra
      </div>
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────

export default function App() {
  const [autenticato, setAutenticato] = useState(false);
  const [nome, setNome] = useState("");
  const [note, setNote] = useState("");
  const [qtyCucina, setQtyCucina] = useState({});
  const [qtyBar, setQtyBar] = useState({});
  const [ordini, setOrdini] = useState([]);
  const [view, setView] = useState("form");
  const [flash, setFlash] = useState(null);
  const [loadingXls, setLoadingXls] = useState(false);

  if (!autenticato) return <LoginScreen onAccess={() => setAutenticato(true)} />;

  const totCucina = CUCINA.reduce((s, it) => s + (qtyCucina[it.id] || 0) * it.prezzo, 0);
  const totBar = BAR.reduce((s, it) => s + (qtyBar[it.id] || 0) * it.prezzo, 0);
  const totale = totCucina + totBar;
  const canSend = nome.trim() && totale > 0;

  function showFlash(msg, color = "#43a047") {
    setFlash({ msg, color });
    setTimeout(() => setFlash(null), 2800);
  }

  function handleInvia() {
    if (!canSend) return;
    setOrdini((prev) => [...prev, {
      nome: nome.trim(), note: note.trim(),
      cucina: CUCINA.filter((i) => qtyCucina[i.id] > 0).map((i) => ({ id: i.id, nome: i.nome, qty: qtyCucina[i.id], prezzo: i.prezzo })),
      bar: BAR.filter((i) => qtyBar[i.id] > 0).map((i) => ({ id: i.id, nome: i.nome, qty: qtyBar[i.id], prezzo: i.prezzo })),
      totale, ts: new Date().toLocaleTimeString("it-IT"),
    }]);
    setNome(""); setNote(""); setQtyCucina({}); setQtyBar({});
    showFlash("✅ Ordine aggiunto!");
  }

  async function handleEstrai() {
    if (!ordini.length) { showFlash("⚠️ Nessun ordine da estrarre", "#e65100"); return; }
    setLoadingXls(true);
    try {
      await generaExcel(ordini);
      showFlash("📥 Excel scaricato!", "#1565c0");
    } catch (e) {
      showFlash("❌ Errore generazione", "#c62828");
      console.error(e);
    }
    setLoadingXls(false);
  }

  const grandTotal = ordini.reduce((s, o) => s + o.totale, 0);

  return (
    <div style={{ minHeight: "100vh", background: "linear-gradient(160deg,#e8eaf6 0%,#fffde7 100%)", fontFamily: "'Segoe UI',Arial,sans-serif", paddingBottom: 50 }}>

      {/* Header */}
      <div style={{ background: "#1a237e", textAlign: "center", padding: "16px 16px 12px" }}>
        <div style={{ fontSize: 10, letterSpacing: 3, color: "#f9a825", fontWeight: 700 }}>POLISPORTIVA ASSISI</div>
        <div style={{ fontSize: 28, fontWeight: 900, letterSpacing: 2, color: "#fff" }}>HAITIAMOLI</div>
        <div style={{ fontSize: 17, fontWeight: 700, color: "#f9a825" }}>2026</div>
        <div style={{ fontSize: 12, color: "rgba(255,255,255,0.75)", fontStyle: "italic", marginTop: 2 }}>
          Prenotazione — <strong style={{ color: "#f9a825" }}>{SQUADRA_FIXED}</strong>
        </div>
        <div style={{ display: "flex", justifyContent: "center", gap: 8, marginTop: 10 }}>
          {[["form", "📋 Ordina"], ["riepilogo", `📊 Riepilogo (${ordini.length})`]].map(([v, label]) => (
            <button key={v} onClick={() => setView(v)} style={{ background: view === v ? "#f9a825" : "rgba(255,255,255,0.15)", color: view === v ? "#1a237e" : "#fff", border: "none", borderRadius: 20, padding: "5px 14px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
              {label}
            </button>
          ))}
        </div>
      </div>

      <div style={{ maxWidth: 520, margin: "0 auto", padding: "14px 13px 0" }}>

        {flash && (
          <div style={{ background: flash.color, color: "#fff", borderRadius: 10, padding: "10px 16px", marginBottom: 12, fontWeight: 700, textAlign: "center", fontSize: 13 }}>
            {flash.msg}
          </div>
        )}

        {/* ── FORM ── */}
        {view === "form" && (
          <>
            <div style={{ background: "#fff", borderRadius: 12, padding: 14, marginBottom: 14, boxShadow: "0 2px 8px rgba(0,0,0,0.07)" }}>
              <div style={{ fontWeight: 700, color: "#1a237e", marginBottom: 8, fontSize: 13 }}>👤 Chi ordina</div>
              <input placeholder="Nome e Cognome *" value={nome} onChange={(e) => setNome(e.target.value)} style={INP} />
              <input placeholder="Note (allergie, preferenze…)" value={note} onChange={(e) => setNote(e.target.value)} style={{ ...INP, marginTop: 7 }} />
            </div>

            <div style={{ background: "#fff", borderRadius: 12, padding: 14, marginBottom: 14, boxShadow: "0 2px 8px rgba(0,0,0,0.07)" }}>
              <Section title="Menù Cucina" icon="🍔" items={CUCINA} qty={qtyCucina} setQty={setQtyCucina} />
              <Section title="Menù Bar" icon="🥤" items={BAR} qty={qtyBar} setQty={setQtyBar} />
            </div>

            <div style={{ background: "#1a237e", borderRadius: 12, padding: "12px 16px", display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
              <div>
                <div style={{ color: "#f9a825", fontSize: 10, fontWeight: 700, letterSpacing: 1 }}>TOTALE</div>
                <div style={{ color: "#fff", fontSize: 22, fontWeight: 900 }}>{fmt(totale)}</div>
              </div>
              <div style={{ textAlign: "right", fontSize: 11, color: "rgba(255,255,255,0.6)" }}>
                {totCucina > 0 && <div>Cucina: {fmt(totCucina)}</div>}
                {totBar > 0 && <div>Bar: {fmt(totBar)}</div>}
              </div>
            </div>

            <div style={{ display: "flex", gap: 8 }}>
              <button onClick={() => { setQtyCucina({}); setQtyBar({}); }} style={{ flex: 1, padding: "11px", borderRadius: 10, border: "2px solid #1a237e", background: "#fff", color: "#1a237e", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                🗑 Azzera
              </button>
              <button onClick={handleInvia} disabled={!canSend} style={{ flex: 2, padding: "11px", borderRadius: 10, border: "none", background: canSend ? "#f9a825" : "#ddd", color: "#1a237e", fontSize: 13, fontWeight: 900, cursor: canSend ? "pointer" : "not-allowed" }}>
                ✅ AGGIUNGI ORDINE
              </button>
            </div>
          </>
        )}

        {/* ── RIEPILOGO ── */}
        {view === "riepilogo" && (
          <>
            {ordini.length === 0 ? (
              <div style={{ textAlign: "center", padding: 40, color: "#bbb", fontSize: 14 }}>Nessun ordine ancora</div>
            ) : (
              <>
                {ordini.map((o, i) => (
                  <div key={i} style={{ background: "#fff", borderRadius: 12, padding: 12, marginBottom: 10, boxShadow: "0 2px 8px rgba(0,0,0,0.07)" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 5 }}>
                      <div>
                        <div style={{ fontWeight: 700, color: "#1a237e", fontSize: 13 }}>{o.nome}</div>
                        <div style={{ fontSize: 10, color: "#bbb" }}>ore {o.ts}</div>
                      </div>
                      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <span style={{ fontWeight: 900, color: "#f9a825", fontSize: 16 }}>{fmt(o.totale)}</span>
                        <button onClick={() => setOrdini((p) => p.filter((_, j) => j !== i))} style={{ background: "#ffebee", color: "#c62828", border: "none", borderRadius: 6, padding: "2px 7px", cursor: "pointer", fontSize: 11, fontWeight: 700 }}>✕</button>
                      </div>
                    </div>
                    {[...o.cucina, ...o.bar].map((row, j) => (
                      <div key={j} style={{ fontSize: 12, color: "#444", display: "flex", justifyContent: "space-between" }}>
                        <span>{row.qty}× {row.nome}</span>
                        <span style={{ color: "#1a237e", fontWeight: 600 }}>{fmt(row.qty * row.prezzo)}</span>
                      </div>
                    ))}
                    {o.note && <div style={{ marginTop: 4, fontSize: 11, color: "#f57c00", fontStyle: "italic" }}>📝 {o.note}</div>}
                  </div>
                ))}

                <div style={{ background: "#1a237e", borderRadius: 12, padding: "11px 16px", display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
                  <div style={{ color: "#fff", fontWeight: 700, fontSize: 13 }}>Totale {ordini.length} ordini</div>
                  <div style={{ color: "#f9a825", fontWeight: 900, fontSize: 18 }}>{fmt(grandTotal)}</div>
                </div>

                <button
                  onClick={handleEstrai}
                  disabled={loadingXls}
                  style={{
                    width: "100%", padding: "15px 0", borderRadius: 12, border: "none",
                    background: loadingXls ? "#aaa" : "#2e7d32",
                    color: "#fff", fontSize: 15, fontWeight: 900,
                    cursor: loadingXls ? "wait" : "pointer",
                    letterSpacing: 0.5, boxShadow: "0 4px 14px rgba(46,125,50,0.4)",
                    display: "flex", alignItems: "center", justifyContent: "center", gap: 10,
                  }}
                >
                  {loadingXls ? "⏳ Generazione…" : "📥 ESTRAI EXCEL"}
                </button>
                <div style={{ textAlign: "center", fontSize: 11, color: "#999", marginTop: 6 }}>
                  Genera il file nel formato ufficiale Haitiamoli 2026 con le quantità aggregate
                </div>
              </>
            )}
          </>
        )}
      </div>
    </div>
  );
}

const INP = {
  width: "100%", padding: "9px 11px", borderRadius: 8,
  border: "1.5px solid #c5cae9", fontSize: 13,
  boxSizing: "border-box", outline: "none", color: "#1a237e",
};
