import React, { useState, useEffect, useMemo, useCallback } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  Cell,
  PieChart,
  Pie,
  Legend,
  ComposedChart,
  Line,
} from "recharts";
import {
  LayoutDashboard,
  UploadCloud,
  Tags,
  Landmark,
  Plus,
  Trash2,
  Pencil,
  Check,
  X,
  ChevronRight,
  FileSpreadsheet,
  AlertCircle,
  ArrowUpRight,
  ArrowDownRight,
  Sparkles,
} from "lucide-react";

// ---------------------------------------------------------------------------
// Costanti e helper
// ---------------------------------------------------------------------------

const TX_KEY = "finance-transactions";
const CFG_KEY = "finance-config";

const DEFAULT_CATEGORIES = [
  "Stipendio",
  "Spesa",
  "Bollette",
  "Casa",
  "Mutuo",
  "Trasporti",
  "Ristoranti",
  "Vino",
  "Abbonamenti",
  "Svago",
  "Salute",
  "Shopping",
  "Telefono",
  "Figlio",
  "Risparmio",
  "Ricarica conto",
  "Altro",
];

const DEFAULT_RULES = [
  { keyword: "esselunga", category: "Spesa" },
  { keyword: "coop", category: "Spesa" },
  { keyword: "conad", category: "Spesa" },
  { keyword: "carrefour", category: "Spesa" },
  { keyword: "eurospin", category: "Spesa" },
  { keyword: "lidl", category: "Spesa" },
  { keyword: "enel", category: "Bollette" },
  { keyword: "eni", category: "Bollette" },
  { keyword: "iren", category: "Bollette" },
  { keyword: "a2a", category: "Bollette" },
  { keyword: "hera", category: "Bollette" },
  { keyword: "tim ", category: "Bollette" },
  { keyword: "vodafone", category: "Bollette" },
  { keyword: "fastweb", category: "Bollette" },
  { keyword: "wind", category: "Bollette" },
  { keyword: "ho-mobile", category: "Telefono" },
  { keyword: "ho mobile", category: "Telefono" },
  { keyword: "iliad", category: "Telefono" },
  { keyword: "affitto", category: "Casa" },
  { keyword: "mutuo", category: "Mutuo" },
  { keyword: "condominio", category: "Casa" },
  { keyword: "trenitalia", category: "Trasporti" },
  { keyword: "italo", category: "Trasporti" },
  { keyword: "autostrade", category: "Trasporti" },
  { keyword: "telepass", category: "Trasporti" },
  { keyword: "atm", category: "Trasporti" },
  { keyword: "esso", category: "Trasporti" },
  { keyword: "ristorante", category: "Ristoranti" },
  { keyword: "pizzeria", category: "Ristoranti" },
  { keyword: "deliveroo", category: "Ristoranti" },
  { keyword: "glovo", category: "Ristoranti" },
  { keyword: "justeat", category: "Ristoranti" },
  { keyword: "netflix", category: "Abbonamenti" },
  { keyword: "spotify", category: "Abbonamenti" },
  { keyword: "cinema", category: "Svago" },
  { keyword: "farmacia", category: "Salute" },
  { keyword: "ospedale", category: "Salute" },
  { keyword: "amazon", category: "Shopping" },
  { keyword: "zara", category: "Shopping" },
  { keyword: "decathlon", category: "Shopping" },
  { keyword: "stipendio", category: "Stipendio" },
  { keyword: "salario", category: "Stipendio" },
];

const DEFAULT_CONFIG = {
  categories: DEFAULT_CATEGORIES,
  rules: DEFAULT_RULES,
  budgets: {},
  accounts: [],
};

function parseAmount(raw) {
  if (raw === null || raw === undefined) return NaN;
  let s = String(raw).trim();
  if (!s) return NaN;
  s = s.replace(/^['"]+|['"]+$/g, "").trim();
  s = s.replace(/[€\s]/g, "");
  if (/^\(.*\)$/.test(s)) {
    s = "-" + s.slice(1, -1);
  }
  const hasComma = s.includes(",");
  const hasDot = s.includes(".");
  if (hasComma && hasDot) {
    // formato italiano: 1.234,56
    if (s.lastIndexOf(",") > s.lastIndexOf(".")) {
      s = s.replace(/\./g, "").replace(",", ".");
    } else {
      s = s.replace(/,/g, "");
    }
  } else if (hasComma) {
    s = s.replace(",", ".");
  }
  const n = parseFloat(s);
  return n;
}

function parseDateFlexible(raw) {
  if (raw === null || raw === undefined) return null;
  if (typeof raw === "number") {
    // seriale data Excel
    const d = XLSX.SSF ? XLSX.SSF.parse_date_code(raw) : null;
    if (d) return new Date(Date.UTC(d.y, d.m - 1, d.d));
  }
  const s = String(raw).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
  m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if (m) {
    let [, d, mo, y] = m;
    if (y.length === 2) y = "20" + y;
    return new Date(Date.UTC(+y, +mo - 1, +d));
  }
  const t = Date.parse(s);
  if (!isNaN(t)) return new Date(t);
  return null;
}

function monthKey(date) {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}`;
}

function formatMonthLabel(key) {
  const [y, m] = key.split("-");
  const d = new Date(Date.UTC(+y, +m - 1, 1));
  const label = d.toLocaleDateString("it-IT", { month: "long", year: "numeric", timeZone: "UTC" });
  return label.charAt(0).toUpperCase() + label.slice(1);
}

function formatMonthShort(key) {
  const [y, m] = key.split("-");
  const d = new Date(Date.UTC(+y, +m - 1, 1));
  return d.toLocaleDateString("it-IT", { month: "short", timeZone: "UTC" }).replace(".", "");
}

function formatCurrency(n) {
  if (n === null || n === undefined || isNaN(n)) return "—";
  return new Intl.NumberFormat("it-IT", { style: "currency", currency: "EUR" }).format(n);
}

function formatDate(date) {
  return date.toLocaleDateString("it-IT", { day: "2-digit", month: "2-digit", year: "numeric", timeZone: "UTC" });
}

function txId(t) {
  return `${t.account}|${t.date}|${t.description}|${t.amount.toFixed(2)}`;
}

function categorize(description, rules) {
  const desc = description.toLowerCase();
  for (const r of rules) {
    if (r.keyword && desc.includes(r.keyword.toLowerCase())) return r.category;
  }
  return "Altro";
}

function spendByCategory(txs) {
  const map = {};
  for (const t of txs) {
    if (t.amount < 0) map[t.category] = (map[t.category] || 0) + -t.amount;
  }
  return map;
}

function totalsOf(txs) {
  let entrate = 0,
    uscite = 0;
  for (const t of txs) {
    if (t.amount >= 0) entrate += t.amount;
    else uscite += -t.amount;
  }
  return { entrate, uscite, saldo: entrate - uscite };
}

// Analisi testuale del mese: confronta il mese selezionato con quello
// precedente e con la media dei 3 mesi prima, individuando le categorie
// cresciute di più e quelle fuori budget. Nessuna chiamata esterna:
// è un'analisi calcolata sui tuoi dati, direttamente nel browser.
function analyzeMonth(transactions, selectedMonth, months, config) {
  const idx = months.indexOf(selectedMonth);
  const prevMonth = idx >= 0 && idx + 1 < months.length ? months[idx + 1] : null;
  const priorMonths = months.slice(idx + 1, idx + 4); // fino a 3 mesi precedenti

  const curTx = transactions.filter((t) => t.month === selectedMonth);
  const curTotals = totalsOf(curTx);
  const curByCat = spendByCategory(curTx);

  const lines = [];

  if (!prevMonth) {
    lines.push(
      `Questo è il primo mese con dati: ${formatCurrency(curTotals.uscite)} di uscite e ${formatCurrency(
        curTotals.entrate
      )} di entrate. Da qui in poi potrò confrontarlo con i mesi successivi.`
    );
  } else {
    const prevTx = transactions.filter((t) => t.month === prevMonth);
    const prevTotals = totalsOf(prevTx);
    const prevByCat = spendByCategory(prevTx);

    const deltaUscite = curTotals.uscite - prevTotals.uscite;
    const deltaLabel = formatMonthLabel(prevMonth);
    if (Math.abs(deltaUscite) < 1) {
      lines.push(`Hai speso circa come a ${deltaLabel}: ${formatCurrency(curTotals.uscite)} di uscite totali.`);
    } else if (deltaUscite > 0) {
      lines.push(
        `Hai speso ${formatCurrency(deltaUscite)} in più rispetto a ${deltaLabel} (${formatCurrency(
          curTotals.uscite
        )} contro ${formatCurrency(prevTotals.uscite)}).`
      );
    } else {
      lines.push(
        `Hai speso ${formatCurrency(-deltaUscite)} in meno rispetto a ${deltaLabel} (${formatCurrency(
          curTotals.uscite
        )} contro ${formatCurrency(prevTotals.uscite)}) — ottimo.`
      );
    }

    // media dei mesi precedenti (esclusi quello corrente), per categoria
    const priorTx = transactions.filter((t) => priorMonths.includes(t.month));
    const priorCount = priorMonths.filter((m) => transactions.some((t) => t.month === m)).length || 1;
    const priorByCat = spendByCategory(priorTx);
    const avgPriorByCat = {};
    Object.keys(priorByCat).forEach((c) => (avgPriorByCat[c] = priorByCat[c] / priorCount));

    const cats = new Set([...Object.keys(curByCat), ...Object.keys(avgPriorByCat)]);
    const increases = Array.from(cats)
      .filter((c) => c !== "Stipendio")
      .map((c) => {
        const cur = curByCat[c] || 0;
        const avg = avgPriorByCat[c] || 0;
        return { category: c, cur, avg, delta: cur - avg };
      })
      .filter((r) => r.cur >= 15 && r.delta > 10)
      .sort((a, b) => b.delta - a.delta)
      .slice(0, 3);

    if (increases.length) {
      const parts = increases.map(
        (r) =>
          `${r.category} (${formatCurrency(r.cur)}${
            r.avg > 0 ? `, contro una media di ${formatCurrency(r.avg)}` : ", categoria nuova o quasi"
          })`
      );
      lines.push(`Le categorie cresciute di più rispetto alla media dei mesi precedenti sono: ${parts.join("; ")}.`);
    }

    const overBudget = Object.keys(config.budgets)
      .map((c) => ({ category: c, budget: config.budgets[c] || 0, speso: curByCat[c] || 0 }))
      .filter((r) => r.budget > 0 && r.speso > r.budget)
      .sort((a, b) => b.speso - b.budget - (a.speso - a.budget))
      .slice(0, 3);

    if (overBudget.length) {
      const parts = overBudget.map(
        (r) => `${r.category} di ${formatCurrency(r.speso - r.budget)} oltre il budget previsto`
      );
      lines.push(`Sei andato fuori budget su: ${parts.join("; ")}.`);
    }

    const suggestion = increases[0] || overBudget[0];
    if (suggestion) {
      const catName = suggestion.category;
      lines.push(
        `Se cerchi dove tagliare, ${catName} è il punto da cui partire il prossimo mese — è la voce che ha pesato di più sull'aumento della spesa.`
      );
    } else if (deltaUscite <= 0) {
      lines.push("Nessuna categoria fuori controllo questo mese: la spesa è sotto controllo su tutta la linea.");
    }
  }

  return lines;
}

// ---------------------------------------------------------------------------
// UI di base
// ---------------------------------------------------------------------------

const INK = "#1C1B18";
const BG = "#F2EEE4";
const SIDEBAR = "#17160F";
const SIDEBAR_ACTIVE = "#F2EEE4";
const CARD = "#FFFFFF";
const LINE = "#E6E1D3";
const GREEN = "#2E8F6F";
const GREEN_SOFT = "#E4F1EA";
const RED = "#C2513F";
const RED_SOFT = "#F6E5E1";
const TEAL = "#3FA79E";
const GOLD = "#C6924A";
const MUTED = "#8B8676";
const ERROR_BG = "#F6E5E1";
const SHADOW = "0 1px 2px rgba(28,27,24,0.04), 0 6px 16px rgba(28,27,24,0.05)";

const PIE_PALETTE = ["#3FA79E", "#C6924A", "#2E8F6F", "#8B7FC7", "#C2513F", "#5C8FBF", "#B08968", "#7CA37E"];

const AVATAR_PALETTE = ["#3FA79E", "#C6924A", "#8B7FC7", "#2E8F6F", "#C2513F", "#5C8FBF"];

function Card({ children, className = "", style = {} }) {
  return (
    <div
      className={`rounded-2xl ${className}`}
      style={{ background: CARD, boxShadow: SHADOW, ...style }}
    >
      {children}
    </div>
  );
}

function NavItem({ icon: Icon, label, active, onClick }) {
  return (
    <button
      onClick={onClick}
      className="w-full flex items-center gap-3 px-3.5 py-2.5 text-left rounded-xl transition-colors"
      style={{
        color: active ? INK : "#B7B2A0",
        background: active ? SIDEBAR_ACTIVE : "transparent",
        fontFamily: "'IBM Plex Sans', system-ui, sans-serif",
        fontWeight: active ? 600 : 400,
        fontSize: "14px",
      }}
    >
      <Icon size={16} strokeWidth={1.9} />
      <span>{label}</span>
    </button>
  );
}

function StatCard({ label, value, dark }) {
  return (
    <Card
      className="px-5 py-4"
      style={dark ? { background: SIDEBAR, boxShadow: SHADOW } : {}}
    >
      <div
        className="text-xs"
        style={{ color: dark ? "#B7B2A0" : MUTED, fontWeight: 500 }}
      >
        {label}
      </div>
      <div
        className="mt-1.5 text-2xl"
        style={{
          color: dark ? "#FFFFFF" : INK,
          fontWeight: 600,
          fontVariantNumeric: "tabular-nums",
          letterSpacing: "-0.01em",
        }}
      >
        {value}
      </div>
    </Card>
  );
}

// ---------------------------------------------------------------------------
// Componente principale
// ---------------------------------------------------------------------------

export default function FinanceTracker() {
  const [loading, setLoading] = useState(true);
  const [transactions, setTransactions] = useState([]);
  const [config, setConfig] = useState(DEFAULT_CONFIG);
  const [view, setView] = useState("dashboard");
  const [selectedMonth, setSelectedMonth] = useState(null);
  const [viewMode, setViewMode] = useState("mese"); // mese | anno
  const [selectedYear, setSelectedYear] = useState(null);
  const [accountFilter, setAccountFilter] = useState("Tutti");
  const [saveError, setSaveError] = useState("");

  // ---- caricamento iniziale --------------------------------------------
  useEffect(() => {
    (async () => {
      try {
        let tx = [];
        let cfg = DEFAULT_CONFIG;
        try {
          const r = await window.storage.get(TX_KEY, false);
          if (r && r.value) tx = JSON.parse(r.value);
        } catch (e) {
          /* nessuna transazione salvata ancora */
        }
        try {
          const r = await window.storage.get(CFG_KEY, false);
          if (r && r.value) {
            const saved = JSON.parse(r.value);
            const mergedCategories = Array.from(new Set([...(saved.categories || []), ...DEFAULT_CATEGORIES]));
            cfg = { ...DEFAULT_CONFIG, ...saved, categories: mergedCategories };
          }
        } catch (e) {
          /* nessuna config salvata ancora, uso default */
        }
        setTransactions(tx);
        setConfig(cfg);
      } finally {
        setLoading(false);
      }
    })();
  }, []);

  const persistTransactions = useCallback(async (next) => {
    setTransactions(next);
    try {
      const res = await window.storage.set(TX_KEY, JSON.stringify(next), false);
      if (!res) setSaveError("Salvataggio non riuscito. Riprova.");
      else setSaveError("");
    } catch (e) {
      setSaveError("Salvataggio non riuscito. Riprova.");
    }
  }, []);

  const persistConfig = useCallback(async (next) => {
    setConfig(next);
    try {
      const res = await window.storage.set(CFG_KEY, JSON.stringify(next), false);
      if (!res) setSaveError("Salvataggio non riuscito. Riprova.");
      else setSaveError("");
    } catch (e) {
      setSaveError("Salvataggio non riuscito. Riprova.");
    }
  }, []);

  // ---- dati derivati -------------------------------------------------
  const months = useMemo(() => {
    const s = new Set(transactions.map((t) => t.month));
    return Array.from(s).sort().reverse();
  }, [transactions]);

  useEffect(() => {
    if (!selectedMonth && months.length) setSelectedMonth(months[0]);
  }, [months, selectedMonth]);

  const years = useMemo(() => {
    const s = new Set(months.map((m) => m.slice(0, 4)));
    return Array.from(s).sort().reverse();
  }, [months]);

  useEffect(() => {
    if (!selectedYear && years.length) setSelectedYear(years[0]);
  }, [years, selectedYear]);

  const periodTx = useMemo(() => {
    if (viewMode === "anno") {
      return transactions.filter(
        (t) => t.month.slice(0, 4) === selectedYear && (accountFilter === "Tutti" || t.account === accountFilter)
      );
    }
    return transactions.filter(
      (t) => t.month === selectedMonth && (accountFilter === "Tutti" || t.account === accountFilter)
    );
  }, [transactions, selectedMonth, selectedYear, viewMode, accountFilter]);

  const monthTx = periodTx; // nome storico usato altrove nel componente

  const totals = useMemo(() => {
    let entrate = 0,
      uscite = 0;
    for (const t of periodTx) {
      if (t.amount >= 0) entrate += t.amount;
      else uscite += -t.amount;
    }
    return { entrate, uscite, saldo: entrate - uscite };
  }, [periodTx]);

  const byCategory = useMemo(() => {
    const map = {};
    for (const t of periodTx) {
      if (t.amount < 0) {
        map[t.category] = (map[t.category] || 0) + -t.amount;
      }
    }
    const monthsInPeriod =
      viewMode === "anno"
        ? new Set(periodTx.map((t) => t.month)).size || 1
        : 1;
    const cats = new Set([...config.categories, ...Object.keys(map)]);
    return Array.from(cats)
      .filter((c) => c !== "Stipendio")
      .map((c) => {
        const speso = map[c] || 0;
        const budget = (config.budgets[c] || 0) * monthsInPeriod;
        return { category: c, speso, budget, scostamento: budget - speso };
      })
      .filter((r) => r.speso > 0 || r.budget > 0)
      .sort((a, b) => b.speso - a.speso);
  }, [periodTx, config, viewMode]);

  // Andamento mensile: serie storica di entrate/uscite/saldo per ogni mese,
  // usata dal grafico di confronto tra mesi (sia in vista Mese che Anno).
  const monthlySeries = useMemo(() => {
    const relevantTx = accountFilter === "Tutti" ? transactions : transactions.filter((t) => t.account === accountFilter);
    if (viewMode === "anno" && selectedYear) {
      return Array.from({ length: 12 }, (_, i) => {
        const m = `${selectedYear}-${String(i + 1).padStart(2, "0")}`;
        const t = totalsOf(relevantTx.filter((tx) => tx.month === m));
        return { month: m, label: formatMonthShort(m), Entrate: Math.round(t.entrate), Uscite: Math.round(t.uscite), Saldo: Math.round(t.saldo) };
      });
    }
    const ascMonths = months.slice().reverse();
    const idx = ascMonths.indexOf(selectedMonth);
    const windowMonths = idx >= 0 ? ascMonths.slice(Math.max(0, idx - 5), idx + 1) : ascMonths.slice(-6);
    return windowMonths.map((m) => {
      const t = totalsOf(relevantTx.filter((tx) => tx.month === m));
      return { month: m, label: formatMonthShort(m), Entrate: Math.round(t.entrate), Uscite: Math.round(t.uscite), Saldo: Math.round(t.saldo) };
    });
  }, [transactions, months, selectedMonth, selectedYear, viewMode, accountFilter]);

  if (loading) {
    return (
      <div className="flex items-center justify-center h-96" style={{ background: BG, color: MUTED }}>
        Caricamento…
      </div>
    );
  }

  return (
    <div
      className="w-full min-h-[640px] flex flex-col md:flex-row p-3 md:p-4 gap-4"
      style={{
        background: BG,
        fontFamily: "'IBM Plex Sans', system-ui, sans-serif",
        color: INK,
      }}
    >
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&display=swap');
      `}</style>

      {/* Sidebar */}
      <div
        className="md:w-56 shrink-0 rounded-2xl flex md:flex-col"
        style={{ background: SIDEBAR, boxShadow: SHADOW }}
      >
        <div className="px-5 pt-6 pb-5 hidden md:block">
          <div className="text-[10px] uppercase" style={{ color: "#8A8574", letterSpacing: "0.12em" }}>
            Portale personale
          </div>
          <div className="text-lg mt-0.5 font-semibold" style={{ color: "#FFFFFF" }}>
            Expense Tracker
          </div>
        </div>
        <nav className="p-2.5 md:p-2.5 flex md:flex-col gap-1 overflow-x-auto md:overflow-visible w-full">
          <NavItem icon={LayoutDashboard} label="Bilancio" active={view === "dashboard"} onClick={() => setView("dashboard")} />
          <NavItem icon={UploadCloud} label="Importa estratto" active={view === "import"} onClick={() => setView("import")} />
          <NavItem icon={Tags} label="Categorie e budget" active={view === "rules"} onClick={() => setView("rules")} />
          <NavItem icon={Landmark} label="Conti" active={view === "accounts"} onClick={() => setView("accounts")} />
        </nav>
      </div>

      {/* Contenuto */}
      <div className="flex-1 min-w-0">
        {saveError && (
          <div className="px-4 py-2 mb-3 text-sm flex items-center gap-2 rounded-xl" style={{ background: ERROR_BG, color: RED }}>
            <AlertCircle size={14} /> {saveError}
          </div>
        )}
        {view === "dashboard" && (
          <Dashboard
            transactions={transactions}
            monthTx={monthTx}
            months={months}
            years={years}
            selectedMonth={selectedMonth}
            setSelectedMonth={setSelectedMonth}
            selectedYear={selectedYear}
            setSelectedYear={setSelectedYear}
            viewMode={viewMode}
            setViewMode={setViewMode}
            monthlySeries={monthlySeries}
            accountFilter={accountFilter}
            setAccountFilter={setAccountFilter}
            accounts={config.accounts}
            totals={totals}
            byCategory={byCategory}
            config={config}
            onUpdateTransaction={(updated) => {
              const next = transactions.map((t) => (t.id === updated.id ? updated : t));
              persistTransactions(next);
            }}
            onDeleteTransaction={(id) => {
              persistTransactions(transactions.filter((t) => t.id !== id));
            }}
            onAddTransaction={(tx, accountUsed) => {
              persistTransactions([...transactions, tx]);
              if (accountUsed && !config.accounts.includes(accountUsed)) {
                persistConfig({ ...config, accounts: [...config.accounts, accountUsed] });
              }
            }}
          />
        )}
        {view === "import" && (
          <ImportView
            config={config}
            onImport={(newTx, accountUsed) => {
              const existingIds = new Set(transactions.map((t) => t.id));
              const toAdd = newTx.filter((t) => !existingIds.has(t.id));
              persistTransactions([...transactions, ...toAdd]);
              if (accountUsed && !config.accounts.includes(accountUsed)) {
                persistConfig({ ...config, accounts: [...config.accounts, accountUsed] });
              }
              setView("dashboard");
              const ms = Array.from(new Set(toAdd.map((t) => t.month))).sort().reverse();
              if (ms.length) setSelectedMonth(ms[0]);
            }}
          />
        )}
        {view === "rules" && (
          <RulesView config={config} onChange={(next) => persistConfig(next)} />
        )}
        {view === "accounts" && (
          <AccountsView
            config={config}
            transactions={transactions}
            onChangeConfig={(next) => persistConfig(next)}
          />
        )}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Dashboard
// ---------------------------------------------------------------------------

function Dashboard({
  transactions,
  monthTx,
  months,
  years,
  selectedMonth,
  setSelectedMonth,
  selectedYear,
  setSelectedYear,
  viewMode,
  setViewMode,
  monthlySeries,
  accountFilter,
  setAccountFilter,
  accounts,
  totals,
  byCategory,
  config,
  onUpdateTransaction,
  onDeleteTransaction,
  onAddTransaction,
}) {
  const [showManualForm, setShowManualForm] = useState(false);
  const [showAnalysis, setShowAnalysis] = useState(false);

  const saveManual = (tx, accountUsed) => {
    onAddTransaction(tx, accountUsed);
    setShowManualForm(false);
  };

  if (!months.length) {
    return (
      <Card className="p-10 max-w-md">
        <div className="text-2xl mb-2 font-semibold">Nessun movimento ancora</div>
        <p style={{ color: MUTED, fontSize: "14.5px", lineHeight: 1.6 }}>
          Comincia caricando l'estratto conto del mese da "Importa estratto", oppure aggiungi subito un movimento a
          mano (utile per le spese in contanti).
        </p>
        {showManualForm ? (
          <div className="mt-5">
            <ManualEntryForm accounts={accounts} categories={config.categories} onSave={saveManual} onCancel={() => setShowManualForm(false)} />
          </div>
        ) : (
          <button
            onClick={() => setShowManualForm(true)}
            className="mt-5 px-3 py-1.5 text-sm rounded-lg flex items-center gap-1 text-white font-medium"
            style={{ background: TEAL }}
          >
            <Plus size={13} /> Aggiungi movimento a mano
          </button>
        )}
      </Card>
    );
  }

  const chartData = byCategory.slice(0, 7).map((r) => ({ name: r.category, Speso: Math.round(r.speso), Budget: Math.round(r.budget) }));
  const pieData = byCategory.filter((r) => r.speso > 0).map((r) => ({ name: r.category, value: Math.round(r.speso) }));
  const totalSpeso = pieData.reduce((s, d) => s + d.value, 0);
  const totalBudget = byCategory.reduce((s, r) => s + (r.budget || 0), 0);
  const tooltipStyle = { fontSize: 12, background: CARD, border: `1px solid ${LINE}`, color: INK, borderRadius: 8 };
  const periodLabel = viewMode === "anno" ? selectedYear : formatMonthLabel(selectedMonth);

  return (
    <div>
      <div className="flex flex-wrap items-end justify-between gap-4 mb-5">
        <div>
          <div className="text-2xl font-semibold">Bilancio</div>
          <div className="flex items-center gap-2 mt-1.5">
            <div className="flex rounded-lg overflow-hidden" style={{ border: `1px solid ${LINE}` }}>
              <button
                onClick={() => setViewMode("mese")}
                className="px-2.5 py-1 text-xs font-medium"
                style={{ background: viewMode === "mese" ? INK : "transparent", color: viewMode === "mese" ? CARD : MUTED }}
              >
                Mese
              </button>
              <button
                onClick={() => setViewMode("anno")}
                className="px-2.5 py-1 text-xs font-medium"
                style={{ background: viewMode === "anno" ? INK : "transparent", color: viewMode === "anno" ? CARD : MUTED }}
              >
                Anno
              </button>
            </div>
            {viewMode === "mese" ? (
              <select
                value={selectedMonth || ""}
                onChange={(e) => setSelectedMonth(e.target.value)}
                className="text-sm bg-transparent outline-none"
                style={{ color: MUTED }}
              >
                {months.map((m) => (
                  <option key={m} value={m}>
                    {formatMonthLabel(m)}
                  </option>
                ))}
              </select>
            ) : (
              <select
                value={selectedYear || ""}
                onChange={(e) => setSelectedYear(e.target.value)}
                className="text-sm bg-transparent outline-none"
                style={{ color: MUTED }}
              >
                {years.map((y) => (
                  <option key={y} value={y}>
                    {y}
                  </option>
                ))}
              </select>
            )}
          </div>
        </div>
        <div className="flex items-center gap-2">
          <select
            value={accountFilter}
            onChange={(e) => setAccountFilter(e.target.value)}
            className="text-sm px-3 py-1.5 rounded-xl"
            style={{ border: `1px solid ${LINE}`, color: INK, background: CARD, boxShadow: SHADOW }}
          >
            <option>Tutti</option>
            {accounts.map((a) => (
              <option key={a}>{a}</option>
            ))}
          </select>
          <button
            onClick={() => setShowAnalysis((v) => !v)}
            className="text-sm px-3 py-1.5 rounded-xl flex items-center gap-1.5 font-medium"
            style={{ background: showAnalysis ? GOLD : CARD, color: showAnalysis ? "#fff" : INK, boxShadow: SHADOW }}
          >
            <Sparkles size={14} /> Analizza
          </button>
          <button
            onClick={() => setShowManualForm((v) => !v)}
            className="text-sm px-3 py-1.5 rounded-xl flex items-center gap-1.5 font-medium text-white"
            style={{ background: TEAL, boxShadow: SHADOW }}
          >
            <Plus size={14} /> Movimento
          </button>
        </div>
      </div>

      {showAnalysis && (
        <div className="mb-4">
          <AiAnalysisPanel transactions={transactions} selectedMonth={selectedMonth} months={months} config={config} />
        </div>
      )}

      {showManualForm && (
        <div className="mb-4">
          <ManualEntryForm
            accounts={accounts}
            categories={config.categories}
            onSave={saveManual}
            onCancel={() => setShowManualForm(false)}
          />
        </div>
      )}

      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <StatCard label="Saldo netto" value={formatCurrency(totals.saldo)} dark />
        <StatCard label="Entrate" value={formatCurrency(totals.entrate)} />
        <StatCard label="Uscite" value={formatCurrency(totals.uscite)} />
      </div>

      {monthlySeries.length > 1 && (
        <Card className="p-5 mt-4">
          <div className="text-sm font-medium mb-3">
            {viewMode === "anno" ? `Andamento mensile — ${selectedYear}` : "Andamento degli ultimi mesi"}
          </div>
          <ResponsiveContainer width="100%" height={200}>
            <ComposedChart data={monthlySeries} margin={{ left: 4, right: 8, top: 4, bottom: 0 }}>
              <CartesianGrid stroke={LINE} vertical={false} />
              <XAxis dataKey="label" tick={{ fontSize: 11, fill: MUTED }} axisLine={{ stroke: LINE }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: MUTED }} axisLine={{ stroke: LINE }} tickLine={false} width={44} />
              <Tooltip formatter={(v) => formatCurrency(v)} contentStyle={tooltipStyle} labelStyle={{ color: INK }} itemStyle={{ color: INK }} cursor={{ fill: BG }} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="Entrate" fill={GREEN} radius={[4, 4, 0, 0]} barSize={14} />
              <Bar dataKey="Uscite" fill={RED} radius={[4, 4, 0, 0]} barSize={14} />
              <Line type="monotone" dataKey="Saldo" stroke={GOLD} strokeWidth={2.5} dot={{ r: 3, fill: GOLD }} />
            </ComposedChart>
          </ResponsiveContainer>
        </Card>
      )}

      {(chartData.length > 0 || pieData.length > 0) && (
        <div className="mt-4 grid grid-cols-1 lg:grid-cols-5 gap-4">
          {chartData.length > 0 && (
            <Card className="p-5 lg:col-span-3">
              <div className="text-sm font-medium mb-3">Spesa per categoria rispetto al budget</div>
              <ResponsiveContainer width="100%" height={Math.max(180, chartData.length * 32)}>
                <BarChart data={chartData} layout="vertical" margin={{ left: 4, right: 16, top: 0, bottom: 0 }}>
                  <CartesianGrid stroke={LINE} horizontal={false} />
                  <XAxis type="number" tick={{ fontSize: 11, fill: MUTED }} axisLine={{ stroke: LINE }} tickLine={false} />
                  <YAxis type="category" dataKey="name" tick={{ fontSize: 12, fill: INK }} width={92} axisLine={{ stroke: LINE }} tickLine={false} />
                  <Tooltip formatter={(v) => formatCurrency(v)} contentStyle={tooltipStyle} labelStyle={{ color: INK }} itemStyle={{ color: INK }} cursor={{ fill: BG }} />
                  <Bar dataKey="Budget" fill={LINE} radius={[0, 6, 6, 0]} barSize={9} />
                  <Bar dataKey="Speso" radius={[0, 6, 6, 0]} barSize={9}>
                    {chartData.map((d, i) => (
                      <Cell key={i} fill={d.Budget > 0 && d.Speso > d.Budget ? RED : TEAL} />
                    ))}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </Card>
          )}
          {pieData.length > 0 && (
            <Card className="p-5 lg:col-span-2">
              <div className="flex items-center justify-between mb-1">
                <div className="text-sm font-medium">Ripartizione spesa</div>
                {totalBudget > 0 && <span className="text-xs" style={{ color: MUTED }}>di {formatCurrency(totalBudget)}</span>}
              </div>
              <div className="relative">
                <ResponsiveContainer width="100%" height={200}>
                  <PieChart>
                    <Pie
                      data={pieData}
                      dataKey="value"
                      nameKey="name"
                      innerRadius="62%"
                      outerRadius="92%"
                      paddingAngle={2}
                      stroke={CARD}
                      strokeWidth={2}
                    >
                      {pieData.map((d, i) => (
                        <Cell key={i} fill={PIE_PALETTE[i % PIE_PALETTE.length]} />
                      ))}
                    </Pie>
                    <Tooltip formatter={(v) => formatCurrency(v)} contentStyle={tooltipStyle} labelStyle={{ color: INK }} itemStyle={{ color: INK }} />
                  </PieChart>
                </ResponsiveContainer>
                <div className="absolute inset-0 flex flex-col items-center justify-center pointer-events-none">
                  <div className="text-lg font-semibold" style={{ fontVariantNumeric: "tabular-nums" }}>{formatCurrency(totalSpeso)}</div>
                  <div className="text-[11px]" style={{ color: MUTED }}>speso</div>
                </div>
              </div>
              <div className="mt-2 space-y-1.5 max-h-32 overflow-y-auto pr-1">
                {byCategory
                  .filter((r) => r.speso > 0)
                  .map((r, i) => (
                    <div key={r.category} className="flex items-center justify-between text-xs">
                      <span className="flex items-center gap-1.5" style={{ color: INK }}>
                        <span className="w-2 h-2 rounded-full inline-block" style={{ background: PIE_PALETTE[i % PIE_PALETTE.length] }} />
                        {r.category}
                      </span>
                      <span style={{ fontVariantNumeric: "tabular-nums", color: MUTED }}>{formatCurrency(r.speso)}</span>
                    </div>
                  ))}
              </div>
            </Card>
          )}
        </div>
      )}

      <Card className="mt-4 p-5">
        <div className="text-sm font-medium mb-3">
          Scostamento dal budget{viewMode === "anno" ? " (budget annuale stimato)" : ""}
        </div>
        <table className="w-full text-sm" style={{ borderCollapse: "collapse" }}>
          <thead>
            <tr style={{ borderBottom: `1px solid ${LINE}`, color: MUTED, textAlign: "left" }}>
              <th className="py-2 font-medium">Categoria</th>
              <th className="py-2 font-medium text-right">Speso</th>
              <th className="py-2 font-medium text-right">Budget</th>
              <th className="py-2 font-medium text-right">Scostamento</th>
            </tr>
          </thead>
          <tbody>
            {byCategory.map((r) => (
              <tr key={r.category} style={{ borderBottom: `1px solid ${LINE}` }}>
                <td className="py-2">{r.category}</td>
                <td className="py-2 text-right" style={{ fontVariantNumeric: "tabular-nums" }}>{formatCurrency(r.speso)}</td>
                <td className="py-2 text-right" style={{ fontVariantNumeric: "tabular-nums", color: MUTED }}>
                  {r.budget ? formatCurrency(r.budget) : "—"}
                </td>
                <td className="py-2 text-right">
                  {r.budget ? (
                    <span
                      className="px-2 py-0.5 rounded-full text-xs font-medium"
                      style={{
                        fontVariantNumeric: "tabular-nums",
                        color: r.scostamento >= 0 ? GREEN : RED,
                        background: r.scostamento >= 0 ? GREEN_SOFT : RED_SOFT,
                      }}
                    >
                      {r.scostamento >= 0 ? "+" : ""}
                      {formatCurrency(r.scostamento)}
                    </span>
                  ) : (
                    <span style={{ color: MUTED }}>—</span>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </Card>

      <Card className="mt-4 p-5">
        <div className="text-sm font-medium mb-3">
          Movimenti {viewMode === "anno" ? `del ${selectedYear}` : "del mese"} ({monthTx.length})
        </div>
        <div className="overflow-x-auto">
          <div className="divide-y" style={{ borderColor: LINE }}>
            {monthTx
              .slice()
              .sort((a, b) => b.date.localeCompare(a.date))
              .map((t, i) => (
                <div key={t.id} className="flex items-center gap-3 py-2.5" style={{ borderBottom: `1px solid ${LINE}` }}>
                  <div
                    className="w-8 h-8 rounded-full flex items-center justify-center text-xs font-semibold shrink-0"
                    style={{ background: AVATAR_PALETTE[i % AVATAR_PALETTE.length] + "26", color: AVATAR_PALETTE[i % AVATAR_PALETTE.length] }}
                  >
                    {t.description.trim().charAt(0).toUpperCase() || "?"}
                  </div>
                  <div className="min-w-0 flex-1">
                    <div className="text-sm truncate" title={t.description}>{t.description}</div>
                    <div className="text-xs" style={{ color: MUTED }}>
                      {formatDate(new Date(t.date))} · {t.account}
                    </div>
                  </div>
                  <select
                    value={t.category}
                    onChange={(e) => onUpdateTransaction({ ...t, category: e.target.value })}
                    className="text-xs rounded-lg outline-none shrink-0"
                    style={{ border: `1px solid ${LINE}`, padding: "3px 6px", color: MUTED }}
                  >
                    {config.categories.map((c) => (
                      <option key={c}>{c}</option>
                    ))}
                  </select>
                  <div
                    className="text-sm font-medium text-right shrink-0 w-24"
                    style={{ fontVariantNumeric: "tabular-nums", color: t.amount >= 0 ? GREEN : RED }}
                  >
                    {t.amount >= 0 ? "+" : ""}
                    {formatCurrency(t.amount)}
                  </div>
                  <button onClick={() => onDeleteTransaction(t.id)} style={{ color: MUTED }} title="Elimina" className="shrink-0">
                    <Trash2 size={13} />
                  </button>
                </div>
              ))}
          </div>
        </div>
      </Card>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Import
// ---------------------------------------------------------------------------

function ImportView({ config, onImport }) {
  const [step, setStep] = useState(1);
  const [fileName, setFileName] = useState("");
  const [rawRows, setRawRows] = useState([]); // array di array, prima riga = header
  const [account, setAccount] = useState(config.accounts[0] || "");
  const [newAccountName, setNewAccountName] = useState("");
  const [dateCol, setDateCol] = useState(0);
  const [descCol, setDescCol] = useState(1);
  const [amountMode, setAmountMode] = useState("single"); // single | double
  const [amountCol, setAmountCol] = useState(2);
  const [invertSign, setInvertSign] = useState(false);
  const [inCol, setInCol] = useState(2);
  const [outCol, setOutCol] = useState(3);
  const [error, setError] = useState("");
  const [preview, setPreview] = useState([]);

  const handleFile = (file) => {
    setError("");
    setFileName(file.name);
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext === "csv" || ext === "txt") {
      Papa.parse(file, {
        complete: (res) => {
          const rows = res.data.filter((r) => r.some((c) => String(c).trim() !== ""));
          if (!rows.length) {
            setError("Il file sembra vuoto.");
            return;
          }
          setRawRows(rows);
          guessColumns(rows[0]);
          setStep(2);
        },
        error: () => setError("Non riesco a leggere questo CSV."),
      });
    } else if (ext === "xlsx" || ext === "xls") {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(e.target.result, { type: "array" });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
          const cleaned = rows.filter((r) => r && r.some((c) => String(c ?? "").trim() !== ""));
          if (!cleaned.length) {
            setError("Il file sembra vuoto.");
            return;
          }
          setRawRows(cleaned);
          guessColumns(cleaned[0]);
          setStep(2);
        } catch (err) {
          setError("Non riesco a leggere questo file Excel.");
        }
      };
      reader.readAsArrayBuffer(file);
    } else {
      setError("Formato non supportato: usa CSV, XLS o XLSX.");
    }
  };

  const guessColumns = (headerRow) => {
    const lower = headerRow.map((h) => String(h ?? "").toLowerCase());
    const findIdx = (words) => lower.findIndex((h) => words.some((w) => h.includes(w)));
    const d = findIdx(["data", "date"]);
    const desc = findIdx(["descrizione", "causale", "description", "dettaglio"]);
    const amt = findIdx(["importo", "amount"]);
    const inc = findIdx(["entrat", "accrediti", "avere"]);
    const out = findIdx(["uscit", "addebiti", "dare"]);
    if (d >= 0) setDateCol(d);
    if (desc >= 0) setDescCol(desc);
    if (amt >= 0) {
      setAmountMode("single");
      setAmountCol(amt);
    } else if (inc >= 0 && out >= 0) {
      setAmountMode("double");
      setInCol(inc);
      setOutCol(out);
    }
  };

  const dataRows = rawRows.slice(1);
  const header = rawRows[0] || [];

  const buildPreview = () => {
    const acc = account === "__new__" ? newAccountName.trim() : account;
    if (!acc) {
      setError("Indica un conto per questa importazione.");
      return;
    }
    const rows = [];
    for (const r of dataRows) {
      const dateRaw = r[dateCol];
      const desc = String(r[descCol] ?? "").trim();
      const date = parseDateFlexible(dateRaw);
      if (!date || !desc) continue;
      let amount;
      if (amountMode === "single") {
        amount = parseAmount(r[amountCol]);
        if (invertSign) amount = -amount;
      } else {
        const inV = parseAmount(r[inCol]) || 0;
        const outV = parseAmount(r[outCol]) || 0;
        amount = (isNaN(inV) ? 0 : inV) - (isNaN(outV) ? 0 : outV);
      }
      if (isNaN(amount)) continue;
      const category = categorize(desc, config.rules);
      const iso = date.toISOString().slice(0, 10);
      rows.push({
        id: "",
        date: iso,
        month: monthKey(date),
        description: desc,
        account: acc,
        amount,
        category,
      });
    }
    rows.forEach((r) => (r.id = txId(r)));
    if (!rows.length) {
      setError("Nessuna riga valida trovata con questo mapping. Controlla le colonne selezionate.");
      return;
    }
    setError("");
    setPreview(rows);
    setStep(3);
  };

  const updatePreviewCategory = (id, category) => {
    setPreview((p) => p.map((r) => (r.id === id ? { ...r, category } : r)));
  };

  const confirmImport = () => {
    const acc = account === "__new__" ? newAccountName.trim() : account;
    onImport(preview, acc);
  };

  const colOptions = header.map((h, i) => (
    <option key={i} value={i}>
      {String(h ?? `Colonna ${i + 1}`)}
    </option>
  ));

  return (
    <Card className="p-6 md:p-8 max-w-3xl">
      <div className="text-2xl font-semibold mb-1">Importa estratto conto</div>
      <p className="text-sm mb-6" style={{ color: MUTED }}>
        Passo {step} di 3
      </p>

      {error && (
        <div className="mb-4 px-3 py-2 text-sm flex items-center gap-2 rounded-xl" style={{ background: ERROR_BG, color: RED }}>
          <AlertCircle size={14} /> {error}
        </div>
      )}

      {step === 1 && (
        <label
          className="flex flex-col items-center justify-center gap-2 py-14 cursor-pointer rounded-2xl"
          style={{ border: `1.5px dashed ${LINE}`, background: BG }}
        >
          <FileSpreadsheet size={28} strokeWidth={1.3} style={{ color: MUTED }} />
          <span style={{ color: INK }}>Trascina qui il file oppure clicca per scegliere</span>
          <span className="text-xs" style={{ color: MUTED }}>CSV, XLS o XLSX scaricato dalla tua banca</span>
          <input
            type="file"
            accept=".csv,.xls,.xlsx,.txt"
            className="hidden"
            onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])}
          />
        </label>
      )}

      {step === 2 && (
        <div className="space-y-5">
          <div className="text-sm" style={{ color: MUTED }}>
            File: {fileName} · {dataRows.length} righe trovate
          </div>

          <div>
            <div className="text-xs uppercase mb-1" style={{ color: MUTED, letterSpacing: "0.06em" }}>Conto</div>
            <div className="flex gap-2 flex-wrap items-center">
              <select value={account} onChange={(e) => setAccount(e.target.value)} className="px-2 py-1.5 text-sm rounded-lg" style={{ border: `1px solid ${LINE}` }}>
                <option value="">Seleziona…</option>
                {config.accounts.map((a) => (
                  <option key={a} value={a}>{a}</option>
                ))}
                <option value="__new__">+ Nuovo conto</option>
              </select>
              {account === "__new__" && (
                <input
                  value={newAccountName}
                  onChange={(e) => setNewAccountName(e.target.value)}
                  placeholder="Nome del conto (es. Carta di credito)"
                  className="px-2 py-1.5 text-sm rounded-lg"
                  style={{ border: `1px solid ${LINE}` }}
                />
              )}
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <div>
              <div className="text-xs uppercase mb-1" style={{ color: MUTED, letterSpacing: "0.06em" }}>Colonna data</div>
              <select value={dateCol} onChange={(e) => setDateCol(+e.target.value)} className="w-full px-2 py-1.5 text-sm rounded-lg" style={{ border: `1px solid ${LINE}` }}>
                {colOptions}
              </select>
            </div>
            <div>
              <div className="text-xs uppercase mb-1" style={{ color: MUTED, letterSpacing: "0.06em" }}>Colonna descrizione</div>
              <select value={descCol} onChange={(e) => setDescCol(+e.target.value)} className="w-full px-2 py-1.5 text-sm rounded-lg" style={{ border: `1px solid ${LINE}` }}>
                {colOptions}
              </select>
            </div>
          </div>

          <div>
            <div className="text-xs uppercase mb-1" style={{ color: MUTED, letterSpacing: "0.06em" }}>Formato importo</div>
            <div className="flex gap-4 text-sm mb-2">
              <label className="flex items-center gap-1.5">
                <input type="radio" checked={amountMode === "single"} onChange={() => setAmountMode("single")} />
                Colonna unica con segno
              </label>
              <label className="flex items-center gap-1.5">
                <input type="radio" checked={amountMode === "double"} onChange={() => setAmountMode("double")} />
                Entrate / Uscite separate
              </label>
            </div>
            {amountMode === "single" ? (
              <div className="flex gap-3 items-center flex-wrap">
                <select value={amountCol} onChange={(e) => setAmountCol(+e.target.value)} className="px-2 py-1.5 text-sm rounded-lg" style={{ border: `1px solid ${LINE}` }}>
                  {colOptions}
                </select>
                <label className="flex items-center gap-1.5 text-sm" style={{ color: MUTED }}>
                  <input type="checkbox" checked={invertSign} onChange={(e) => setInvertSign(e.target.checked)} />
                  Inverti segno (positivo = uscita)
                </label>
              </div>
            ) : (
              <div className="flex gap-3 flex-wrap">
                <div>
                  <div className="text-xs mb-1" style={{ color: MUTED }}>Entrate</div>
                  <select value={inCol} onChange={(e) => setInCol(+e.target.value)} className="px-2 py-1.5 text-sm rounded-lg" style={{ border: `1px solid ${LINE}` }}>
                    {colOptions}
                  </select>
                </div>
                <div>
                  <div className="text-xs mb-1" style={{ color: MUTED }}>Uscite</div>
                  <select value={outCol} onChange={(e) => setOutCol(+e.target.value)} className="px-2 py-1.5 text-sm rounded-lg" style={{ border: `1px solid ${LINE}` }}>
                    {colOptions}
                  </select>
                </div>
              </div>
            )}
          </div>

          <button onClick={buildPreview} className="px-4 py-2 text-sm rounded-xl text-white font-medium" style={{ background: TEAL }}>
            Continua →
          </button>
        </div>
      )}

      {step === 3 && (
        <div>
          <div className="text-sm mb-4" style={{ color: MUTED }}>
            {preview.length} movimenti pronti per l'importazione. Puoi correggere la categoria prima di confermare.
          </div>
          <div className="max-h-96 overflow-y-auto mb-5">
            <table className="w-full text-sm" style={{ borderCollapse: "collapse" }}>
              <thead className="sticky top-0" style={{ background: CARD }}>
                <tr style={{ borderBottom: `1px solid ${LINE}`, color: MUTED, textAlign: "left" }}>
                  <th className="py-2 font-normal">Data</th>
                  <th className="py-2 font-normal">Descrizione</th>
                  <th className="py-2 font-normal">Categoria</th>
                  <th className="py-2 font-normal text-right">Importo</th>
                </tr>
              </thead>
              <tbody>
                {preview.map((r) => (
                  <tr key={r.id} style={{ borderBottom: `1px solid ${LINE}` }}>
                    <td className="py-1.5 whitespace-nowrap" style={{ color: MUTED }}>{formatDate(new Date(r.date))}</td>
                    <td className="py-1.5 max-w-[240px] truncate" title={r.description}>{r.description}</td>
                    <td className="py-1.5">
                      <select
                        value={r.category}
                        onChange={(e) => updatePreviewCategory(r.id, e.target.value)}
                        className="bg-transparent text-sm outline-none"
                        style={{ border: `1px solid ${LINE}`, padding: "2px 4px" }}
                      >
                        {config.categories.map((c) => (
                          <option key={c}>{c}</option>
                        ))}
                      </select>
                    </td>
                    <td className="py-1.5 text-right" style={{ fontVariantNumeric: "tabular-nums", color: r.amount >= 0 ? GREEN : RED }}>
                      {formatCurrency(r.amount)}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="flex gap-2">
            <button onClick={() => setStep(2)} className="px-4 py-2 text-sm rounded-xl" style={{ border: `1px solid ${LINE}`, color: INK }}>
              ← Indietro
            </button>
            <button onClick={confirmImport} className="px-4 py-2 text-sm rounded-xl text-white font-medium" style={{ background: TEAL }}>
              Conferma importazione
            </button>
          </div>
        </div>
      )}
    </Card>
  );
}

// ---------------------------------------------------------------------------
// Categorie & budget
// ---------------------------------------------------------------------------

function RulesView({ config, onChange }) {
  const [newCategory, setNewCategory] = useState("");
  const [newRuleKeyword, setNewRuleKeyword] = useState("");
  const [newRuleCategory, setNewRuleCategory] = useState(config.categories[0] || "");

  const setBudget = (cat, value) => {
    onChange({ ...config, budgets: { ...config.budgets, [cat]: value === "" ? undefined : parseFloat(value) || 0 } });
  };

  const addCategory = () => {
    const c = newCategory.trim();
    if (!c || config.categories.includes(c)) return;
    onChange({ ...config, categories: [...config.categories, c] });
    setNewCategory("");
  };

  const removeCategory = (c) => {
    const { [c]: _, ...restBudgets } = config.budgets;
    onChange({
      ...config,
      categories: config.categories.filter((x) => x !== c),
      budgets: restBudgets,
      rules: config.rules.filter((r) => r.category !== c),
    });
  };

  const addRule = () => {
    const k = newRuleKeyword.trim();
    if (!k) return;
    onChange({ ...config, rules: [{ keyword: k, category: newRuleCategory }, ...config.rules] });
    setNewRuleKeyword("");
  };

  const removeRule = (idx) => {
    onChange({ ...config, rules: config.rules.filter((_, i) => i !== idx) });
  };

  return (
    <div className="max-w-3xl space-y-4">
      <Card className="p-6 md:p-8">
        <div className="text-2xl font-semibold mb-4">Categorie e budget mensile</div>
        <table className="w-full text-sm mb-3" style={{ borderCollapse: "collapse" }}>
          <thead>
            <tr style={{ borderBottom: `1px solid ${LINE}`, color: MUTED, textAlign: "left" }}>
              <th className="py-2 font-medium">Categoria</th>
              <th className="py-2 font-medium text-right">Budget mensile</th>
              <th className="py-2 font-medium w-8"></th>
            </tr>
          </thead>
          <tbody>
            {config.categories.map((c) => (
              <tr key={c} style={{ borderBottom: `1px solid ${LINE}` }}>
                <td className="py-2">{c}</td>
                <td className="py-2 text-right">
                  <input
                    type="number"
                    value={config.budgets[c] ?? ""}
                    onChange={(e) => setBudget(c, e.target.value)}
                    placeholder="0"
                    className="w-24 text-right bg-transparent outline-none rounded-lg"
                    style={{ border: `1px solid ${LINE}`, padding: "3px 6px", fontVariantNumeric: "tabular-nums" }}
                  />
                </td>
                <td className="py-2 text-right">
                  <button onClick={() => removeCategory(c)} style={{ color: MUTED }} title="Elimina categoria">
                    <Trash2 size={13} />
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        <div className="flex gap-2">
          <input
            value={newCategory}
            onChange={(e) => setNewCategory(e.target.value)}
            placeholder="Nuova categoria"
            className="px-2 py-1.5 text-sm rounded-lg"
            style={{ border: `1px solid ${LINE}` }}
          />
          <button onClick={addCategory} className="px-3 py-1.5 text-sm rounded-lg flex items-center gap-1 text-white font-medium" style={{ background: TEAL }}>
            <Plus size={13} /> Aggiungi
          </button>
        </div>
      </Card>

      <Card className="p-6 md:p-8">
        <div className="text-2xl font-semibold mb-1">Regole di categorizzazione</div>
        <p className="text-sm mb-4" style={{ color: MUTED }}>
          Se la descrizione di un movimento contiene la parola chiave, viene assegnata automaticamente la categoria.
          La prima regola che corrisponde vince.
        </p>
        <div className="flex gap-2 mb-4 flex-wrap">
          <input
            value={newRuleKeyword}
            onChange={(e) => setNewRuleKeyword(e.target.value)}
            placeholder="Parola chiave (es. esselunga)"
            className="px-2 py-1.5 text-sm rounded-lg"
            style={{ border: `1px solid ${LINE}` }}
          />
          <select value={newRuleCategory} onChange={(e) => setNewRuleCategory(e.target.value)} className="px-2 py-1.5 text-sm rounded-lg" style={{ border: `1px solid ${LINE}` }}>
            {config.categories.map((c) => (
              <option key={c}>{c}</option>
            ))}
          </select>
          <button onClick={addRule} className="px-3 py-1.5 text-sm rounded-lg flex items-center gap-1 text-white font-medium" style={{ background: TEAL }}>
            <Plus size={13} /> Aggiungi regola
          </button>
        </div>
        <div className="max-h-80 overflow-y-auto">
          <table className="w-full text-sm" style={{ borderCollapse: "collapse" }}>
            <tbody>
              {config.rules.map((r, i) => (
                <tr key={i} style={{ borderBottom: `1px solid ${LINE}` }}>
                  <td className="py-1.5" style={{ color: MUTED }}>"{r.keyword}"</td>
                  <td className="py-1.5">
                    <ChevronRight size={12} style={{ display: "inline", color: MUTED, marginRight: 4 }} />
                    {r.category}
                  </td>
                  <td className="py-1.5 text-right">
                    <button onClick={() => removeRule(i)} style={{ color: MUTED }}>
                      <Trash2 size={13} />
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Conti
// ---------------------------------------------------------------------------

function AccountsView({ config, transactions, onChangeConfig }) {
  const [newAccount, setNewAccount] = useState("");

  const addAccount = () => {
    const a = newAccount.trim();
    if (!a || config.accounts.includes(a)) return;
    onChangeConfig({ ...config, accounts: [...config.accounts, a] });
    setNewAccount("");
  };

  const removeAccount = (a) => {
    onChangeConfig({ ...config, accounts: config.accounts.filter((x) => x !== a) });
  };

  const countFor = (a) => transactions.filter((t) => t.account === a).length;

  return (
    <Card className="p-6 md:p-8 max-w-xl">
      <div className="text-2xl font-semibold mb-4">Conti e carte</div>
      <table className="w-full text-sm mb-4" style={{ borderCollapse: "collapse" }}>
        <tbody>
          {config.accounts.map((a) => (
            <tr key={a} style={{ borderBottom: `1px solid ${LINE}` }}>
              <td className="py-2">{a}</td>
              <td className="py-2 text-right" style={{ color: MUTED }}>{countFor(a)} movimenti</td>
              <td className="py-2 text-right">
                <button onClick={() => removeAccount(a)} style={{ color: MUTED }} title="Rimuovi conto">
                  <Trash2 size={13} />
                </button>
              </td>
            </tr>
          ))}
          {!config.accounts.length && (
            <tr>
              <td className="py-2" style={{ color: MUTED }}>
                Nessun conto ancora — se ne crea uno automaticamente al primo import.
              </td>
            </tr>
          )}
        </tbody>
      </table>
      <div className="flex gap-2">
        <input
          value={newAccount}
          onChange={(e) => setNewAccount(e.target.value)}
          placeholder="Nome conto (es. Conto Corrente)"
          className="px-2 py-1.5 text-sm rounded-lg"
          style={{ border: `1px solid ${LINE}` }}
        />
        <button onClick={addAccount} className="px-3 py-1.5 text-sm rounded-lg flex items-center gap-1 text-white font-medium" style={{ background: TEAL }}>
          <Plus size={13} /> Aggiungi
        </button>
      </div>
    </Card>
  );
}

// ---------------------------------------------------------------------------
// Movimento manuale (es. spese in contanti)
// ---------------------------------------------------------------------------

function ManualEntryForm({ accounts, categories, onSave, onCancel }) {
  const [date, setDate] = useState(() => new Date().toISOString().slice(0, 10));
  const [description, setDescription] = useState("");
  const [account, setAccount] = useState(accounts[0] || "Contanti");
  const [newAccountName, setNewAccountName] = useState("");
  const [category, setCategory] = useState(categories[0] || "Altro");
  const [kind, setKind] = useState("uscita"); // uscita | entrata
  const [amount, setAmount] = useState("");
  const [error, setError] = useState("");

  const handleSave = () => {
    const desc = description.trim();
    const acc = account === "__new__" ? newAccountName.trim() : account;
    const amt = parseFloat(String(amount).replace(",", "."));
    if (!desc) {
      setError("Inserisci una descrizione.");
      return;
    }
    if (!acc) {
      setError("Indica un conto (es. Contanti).");
      return;
    }
    if (isNaN(amt) || amt <= 0) {
      setError("Inserisci un importo valido.");
      return;
    }
    const d = parseDateFlexible(date);
    if (!d) {
      setError("Data non valida.");
      return;
    }
    const signedAmount = kind === "entrata" ? amt : -amt;
    const base = {
      date: d.toISOString().slice(0, 10),
      month: monthKey(d),
      description: desc,
      account: acc,
      amount: signedAmount,
      category,
    };
    // suffisso univoco: un movimento manuale non deve mai essere scartato
    // come "duplicato" di un import CSV successivo.
    const tx = { ...base, id: txId(base) + "-m" + Date.now() };
    onSave(tx, account === "__new__" ? acc : null);
  };

  return (
    <Card className="p-5 max-w-2xl">
      <div className="text-sm font-medium mb-3">Aggiungi movimento a mano</div>
      {error && (
        <div className="mb-3 px-3 py-2 text-sm flex items-center gap-2 rounded-xl" style={{ background: ERROR_BG, color: RED }}>
          <AlertCircle size={14} /> {error}
        </div>
      )}
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-3">
        <div>
          <div className="text-xs uppercase mb-1" style={{ color: MUTED, letterSpacing: "0.06em" }}>Data</div>
          <input
            type="date"
            value={date}
            onChange={(e) => setDate(e.target.value)}
            className="w-full px-2 py-1.5 text-sm rounded-lg"
            style={{ border: `1px solid ${LINE}` }}
          />
        </div>
        <div>
          <div className="text-xs uppercase mb-1" style={{ color: MUTED, letterSpacing: "0.06em" }}>Descrizione</div>
          <input
            value={description}
            onChange={(e) => setDescription(e.target.value)}
            placeholder="es. Caffè, mercatino, regalo…"
            className="w-full px-2 py-1.5 text-sm rounded-lg"
            style={{ border: `1px solid ${LINE}` }}
          />
        </div>
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-3">
        <div>
          <div className="text-xs uppercase mb-1" style={{ color: MUTED, letterSpacing: "0.06em" }}>Conto</div>
          <div className="flex gap-2 flex-wrap">
            <select
              value={account}
              onChange={(e) => setAccount(e.target.value)}
              className="px-2 py-1.5 text-sm rounded-lg flex-1 min-w-[120px]"
              style={{ border: `1px solid ${LINE}` }}
            >
              {!accounts.length && <option value="Contanti">Contanti</option>}
              {accounts.map((a) => (
                <option key={a} value={a}>{a}</option>
              ))}
              <option value="__new__">+ Nuovo conto (es. Contanti)</option>
            </select>
          </div>
          {account === "__new__" && (
            <input
              value={newAccountName}
              onChange={(e) => setNewAccountName(e.target.value)}
              placeholder="es. Contanti"
              className="w-full mt-2 px-2 py-1.5 text-sm rounded-lg"
              style={{ border: `1px solid ${LINE}` }}
            />
          )}
        </div>
        <div>
          <div className="text-xs uppercase mb-1" style={{ color: MUTED, letterSpacing: "0.06em" }}>Categoria</div>
          <select
            value={category}
            onChange={(e) => setCategory(e.target.value)}
            className="w-full px-2 py-1.5 text-sm rounded-lg"
            style={{ border: `1px solid ${LINE}` }}
          >
            {categories.map((c) => (
              <option key={c}>{c}</option>
            ))}
          </select>
        </div>
      </div>
      <div className="flex flex-wrap items-end gap-3 mb-4">
        <div>
          <div className="text-xs uppercase mb-1" style={{ color: MUTED, letterSpacing: "0.06em" }}>Tipo</div>
          <div className="flex gap-1 rounded-lg overflow-hidden" style={{ border: `1px solid ${LINE}` }}>
            <button
              type="button"
              onClick={() => setKind("uscita")}
              className="px-3 py-1.5 text-sm"
              style={{ background: kind === "uscita" ? RED_SOFT : "transparent", color: kind === "uscita" ? RED : MUTED }}
            >
              Uscita
            </button>
            <button
              type="button"
              onClick={() => setKind("entrata")}
              className="px-3 py-1.5 text-sm"
              style={{ background: kind === "entrata" ? GREEN_SOFT : "transparent", color: kind === "entrata" ? GREEN : MUTED }}
            >
              Entrata
            </button>
          </div>
        </div>
        <div>
          <div className="text-xs uppercase mb-1" style={{ color: MUTED, letterSpacing: "0.06em" }}>Importo (€)</div>
          <input
            value={amount}
            onChange={(e) => setAmount(e.target.value)}
            placeholder="0,00"
            inputMode="decimal"
            className="w-32 px-2 py-1.5 text-sm rounded-lg"
            style={{ border: `1px solid ${LINE}`, fontVariantNumeric: "tabular-nums" }}
          />
        </div>
      </div>
      <div className="flex gap-2">
        <button onClick={onCancel} className="px-4 py-2 text-sm rounded-xl" style={{ border: `1px solid ${LINE}`, color: INK }}>
          Annulla
        </button>
        <button onClick={handleSave} className="px-4 py-2 text-sm rounded-xl text-white font-medium" style={{ background: TEAL }}>
          Salva movimento
        </button>
      </div>
    </Card>
  );
}

// ---------------------------------------------------------------------------
// Analisi del mese con AI (Claude), con fallback automatico senza AI
// ---------------------------------------------------------------------------

const AI_KEY_STORAGE_KEY = "anthropic_api_key";
const FIXED_CATEGORY_HINT =
  "Mutuo, Casa, Bollette, Telefono, Assicurazioni, Risparmio, Ricarica conto";

async function callClaude(apiKey, prompt) {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "anthropic-dangerous-direct-browser-access": "true",
    },
    body: JSON.stringify({
      model: "claude-haiku-4-5-20251001",
      max_tokens: 500,
      messages: [{ role: "user", content: prompt }],
    }),
  });
  if (!res.ok) {
    let detail = "";
    try {
      const j = await res.json();
      detail = j?.error?.message || "";
    } catch (e) {
      /* risposta non JSON */
    }
    if (res.status === 401) throw new Error("Chiave API non valida. Controllala e riprova.");
    throw new Error(detail || `Errore dell'API (codice ${res.status}).`);
  }
  const data = await res.json();
  return (data.content || [])
    .filter((b) => b.type === "text")
    .map((b) => b.text)
    .join("\n")
    .trim();
}

function buildAnalysisPrompt(transactions, selectedMonth, months, config) {
  const idx = months.indexOf(selectedMonth);
  const prevMonth = idx >= 0 && idx + 1 < months.length ? months[idx + 1] : null;

  const curTx = transactions.filter((t) => t.month === selectedMonth);
  const curByCat = spendByCategory(curTx);
  const curTotals = totalsOf(curTx);

  const curLines = Object.keys(curByCat)
    .sort((a, b) => curByCat[b] - curByCat[a])
    .map((c) => `${c}: speso ${curByCat[c].toFixed(0)}€${config.budgets[c] ? `, budget ${config.budgets[c]}€` : ""}`)
    .join("\n");

  let prevBlock = "Nessun mese precedente disponibile per il confronto.";
  if (prevMonth) {
    const prevTx = transactions.filter((t) => t.month === prevMonth);
    const prevByCat = spendByCategory(prevTx);
    const prevLines = Object.keys(prevByCat)
      .sort((a, b) => prevByCat[b] - prevByCat[a])
      .map((c) => `${c}: ${prevByCat[c].toFixed(0)}€`)
      .join("\n");
    prevBlock = `Mese precedente (${formatMonthLabel(prevMonth)}), uscite totali ${totalsOf(prevTx).uscite.toFixed(
      0
    )}€:\n${prevLines}`;
  }

  return `Sei un assistente che aiuta una persona a leggere le proprie spese personali su un portale che si è costruita da sola.

Mese in esame: ${formatMonthLabel(selectedMonth)}. Entrate totali ${curTotals.entrate.toFixed(
    0
  )}€, uscite totali ${curTotals.uscite.toFixed(0)}€, saldo ${curTotals.saldo.toFixed(0)}€.

Spesa per categoria questo mese:
${curLines || "(nessuna spesa registrata)"}

${prevBlock}

Istruzioni:
- Le seguenti categorie sono spese fisse, difficili o non sensate da ridurre nel breve periodo: ${FIXED_CATEGORY_HINT}. Non suggerire mai di tagliarle, a meno che l'importo di questo mese non sia chiaramente anomalo rispetto al solito (es. raddoppiato senza motivo).
- Concentra i consigli di risparmio solo su categorie discrezionali (es. Ristoranti, Shopping, Svago, Vino, Abbonamenti, Trasporti, Spesa se sopra la media).
- Scrivi in italiano, tono diretto e amichevole, come se parlassi direttamente con la persona. Niente markdown, niente elenchi puntati: 4-5 frasi scorrevoli.
- Includi: un giudizio complessivo sul mese rispetto al precedente, e un consiglio pratico e specifico su una o due categorie discrezionali su cui vale la pena agire (se ce ne sono).`;
}

function AiAnalysisPanel({ transactions, selectedMonth, months, config }) {
  const [apiKey, setApiKey] = useState(undefined); // undefined = ancora in caricamento
  const [keyInput, setKeyInput] = useState("");
  const [mode, setMode] = useState("ai"); // ai | offline
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState("");
  const [error, setError] = useState("");

  useEffect(() => {
    (async () => {
      try {
        const r = await window.storage.get(AI_KEY_STORAGE_KEY, false);
        setApiKey(r && r.value ? r.value : null);
      } catch (e) {
        setApiKey(null);
      }
    })();
  }, []);

  const runAnalysis = async (key) => {
    setLoading(true);
    setError("");
    setResult("");
    try {
      const prompt = buildAnalysisPrompt(transactions, selectedMonth, months, config);
      const text = await callClaude(key, prompt);
      setResult(text);
    } catch (e) {
      setError(e.message || "Analisi non riuscita.");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (apiKey && mode === "ai") runAnalysis(apiKey);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [apiKey, selectedMonth]);

  const saveKey = async () => {
    const k = keyInput.trim();
    if (!k) return;
    try {
      await window.storage.set(AI_KEY_STORAGE_KEY, k, false);
    } catch (e) {
      /* se il salvataggio fallisce, la chiave resta comunque in memoria per questa sessione */
    }
    setApiKey(k);
  };

  const forgetKey = async () => {
    try {
      await window.storage.delete(AI_KEY_STORAGE_KEY, false);
    } catch (e) {
      /* ignorabile */
    }
    setApiKey(null);
    setResult("");
  };

  const offlineLines = mode === "offline" ? analyzeMonth(transactions, selectedMonth, months, config) : [];

  return (
    <Card className="p-5">
      <div className="flex items-center justify-between mb-2">
        <div className="text-sm font-medium flex items-center gap-1.5">
          <Sparkles size={14} style={{ color: GOLD }} /> Come è andato {formatMonthLabel(selectedMonth).toLowerCase()}
        </div>
        {apiKey && (
          <button onClick={forgetKey} className="text-xs" style={{ color: MUTED }}>
            Rimuovi chiave salvata
          </button>
        )}
      </div>

      {apiKey === undefined && <div className="text-sm" style={{ color: MUTED }}>Caricamento…</div>}

      {apiKey === null && mode === "ai" && (
        <div>
          <p className="text-sm mb-3" style={{ color: MUTED, lineHeight: 1.5 }}>
            Per un'analisi vera (che capisce quali spese sono fisse e quali no) serve una tua chiave API Anthropic —
            gratuita da creare, il costo di queste analisi è di pochi centesimi. La chiave resta salvata solo nel tuo
            account, privata.
          </p>
          <div className="flex gap-2 flex-wrap mb-2">
            <input
              type="password"
              value={keyInput}
              onChange={(e) => setKeyInput(e.target.value)}
              placeholder="sk-ant-…"
              className="px-2 py-1.5 text-sm rounded-lg flex-1 min-w-[200px]"
              style={{ border: `1px solid ${LINE}` }}
            />
            <button onClick={saveKey} className="px-3 py-1.5 text-sm rounded-lg text-white font-medium" style={{ background: TEAL }}>
              Salva e analizza
            </button>
          </div>
          <div className="text-xs" style={{ color: MUTED }}>
            La crei su{" "}
            <a href="https://console.anthropic.com/settings/keys" target="_blank" rel="noreferrer" style={{ color: TEAL }}>
              console.anthropic.com/settings/keys
            </a>
            . In alternativa,{" "}
            <button onClick={() => setMode("offline")} style={{ color: TEAL, textDecoration: "underline" }}>
              usa l'analisi automatica senza AI
            </button>
            .
          </div>
        </div>
      )}

      {apiKey && mode === "ai" && (
        <div>
          {loading && <div className="text-sm" style={{ color: MUTED }}>Sto leggendo i tuoi movimenti…</div>}
          {error && (
            <div className="mb-3 px-3 py-2 text-sm flex items-center gap-2 rounded-xl" style={{ background: ERROR_BG, color: RED }}>
              <AlertCircle size={14} /> {error}
            </div>
          )}
          {!loading && result && (
            <p className="text-sm" style={{ color: INK, lineHeight: 1.6, whiteSpace: "pre-line" }}>
              {result}
            </p>
          )}
          {!loading && (
            <button onClick={() => runAnalysis(apiKey)} className="mt-3 text-xs" style={{ color: MUTED }}>
              Rigenera analisi
            </button>
          )}
        </div>
      )}

      {mode === "offline" && (
        <div>
          <ul className="space-y-1.5 text-sm mb-2" style={{ color: INK, lineHeight: 1.55 }}>
            {offlineLines.map((line, i) => (
              <li key={i} className="flex gap-2">
                <span style={{ color: GOLD }}>—</span>
                <span>{line}</span>
              </li>
            ))}
          </ul>
          <button onClick={() => setMode("ai")} className="text-xs" style={{ color: TEAL }}>
            Prova invece l'analisi con AI
          </button>
        </div>
      )}
    </Card>
  );
}
