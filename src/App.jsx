import { useState, useMemo, useEffect } from "react";
import {
  AreaChart, Area, PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid,
  Tooltip, ResponsiveContainer, Legend, BarChart, Bar, LineChart, Line
} from "recharts";
import * as XLSX from "xlsx";

// ---------- Tax engine (FY 2026-27, new regime, individual < 60) ----------
const SLABS = [
  { upto: 400000, rate: 0 },
  { upto: 800000, rate: 0.05 },
  { upto: 1200000, rate: 0.10 },
  { upto: 1600000, rate: 0.15 },
  { upto: 2000000, rate: 0.20 },
  { upto: 2400000, rate: 0.25 },
  { upto: Infinity, rate: 0.30 },
];

function slabTax(income) {
  let tax = 0, prev = 0;
  for (const s of SLABS) {
    if (income > prev) {
      tax += (Math.min(income, s.upto) - prev) * s.rate;
      prev = s.upto;
    } else break;
  }
  // Section 87A rebate: taxable income up to 12L -> nil tax (new regime)
  if (income <= 1200000) return 0;
  const cess = tax * 0.04;
  return tax + cess;
}

function effectiveSlabRate(income) {
  if (income <= 0) return 0;
  return slabTax(income) / income;
}

const EQUITY_EXEMPTION = 125000; // Sec 112A annual LTCG exemption
const EQUITY_LTCG_RATE = 0.125;
const GOLD_LTCG_RATE = 0.125;

const ASSET_META = {
  fd:     { label: "Fixed Deposit",        color: "#C79A45", taxType: "slab" },
  scss:   { label: "SCSS (Senior Citizen)", color: "#8FA6C7", taxType: "slab" },
  debtmf: { label: "Debt Mutual Fund",      color: "#7C93A6", taxType: "slab" },
  equitymf:{ label: "Equity MF (SWP)",      color: "#5FA777", taxType: "equityLTCG" },
  gold:   { label: "Gold Fund / ETF",       color: "#D6B36B", taxType: "goldLTCG" },
};

function runSimulation({ corpus, monthlyExpense, inflation, years, otherIncome, allocations }) {
  const rows = [];
  let bal = corpus;
  let depletionYear = null;

  for (let y = 1; y <= years; y++) {
    if (bal <= 0) { rows.push({ year: y, corpus: 0, expense: 0 }); continue; }

    const expense = monthlyExpense * 12 * Math.pow(1 + inflation / 100, y - 1);

    // Pre-tax growth per asset (on current balance, proportionally split)
    let slabIncome = otherIncome; // other pension/annuity income taxed at slab, used for bracket calc
    let equityGain = 0, goldGain = 0;
    let investmentGrowth = 0; // growth from the allocated corpus only (excludes otherIncome)

    for (const key of Object.keys(allocations)) {
      const a = allocations[key];
      const assetValue = bal * (a.pct / 100);
      const growth = assetValue * (a.returnRate / 100);
      investmentGrowth += growth;
      const meta = ASSET_META[key];
      if (meta.taxType === "slab") slabIncome += growth;
      if (meta.taxType === "equityLTCG") equityGain += growth;
      if (meta.taxType === "goldLTCG") goldGain += growth;
    }

    // Tax on slab-taxed income (FD/SCSS/Debt MF interest + other income), using the
    // blended bracket so otherIncome correctly pushes interest into a higher slab
    const slabRate = effectiveSlabRate(slabIncome);
    const slabTaxAmt = slabIncome * slabRate;
    const investmentSlabTaxAmt = (slabIncome - otherIncome) * slabRate; // portion of slab tax from investments only
    const otherIncomeTaxAmt = otherIncome * slabRate;

    // Tax on equity LTCG (assume all gains realized are long-term, 1.25L exemption/yr)
    const taxableEquity = Math.max(0, equityGain - EQUITY_EXEMPTION);
    const equityTaxAmt = taxableEquity * EQUITY_LTCG_RATE;

    // Tax on gold LTCG (no exemption)
    const goldTaxAmt = goldGain * GOLD_LTCG_RATE;

    const totalTax = slabTaxAmt + equityTaxAmt + goldTaxAmt;
    // otherIncome is real cash that arrives every year — credit it to the corpus (net of its own tax)
    const preTaxGrowth = investmentGrowth + otherIncome;
    const investmentPostTaxGrowth = investmentGrowth - investmentSlabTaxAmt - equityTaxAmt - goldTaxAmt;
    const otherIncomeNet = otherIncome - otherIncomeTaxAmt;
    const postTaxGrowth = investmentPostTaxGrowth + otherIncomeNet;

    bal = bal + postTaxGrowth - expense;
    if (bal <= 0 && depletionYear === null) depletionYear = y;

    rows.push({
      year: y,
      corpus: Math.max(0, Math.round(bal)),
      expense: Math.round(expense),
      preTaxGrowth: Math.round(preTaxGrowth),
      investmentGrowth: Math.round(investmentGrowth),
      investmentPostTaxGrowth: Math.round(investmentPostTaxGrowth),
      tax: Math.round(totalTax),
      postTaxGrowth: Math.round(postTaxGrowth),
    });
  }
  return { rows, depletionYear };
}

function inr(n) {
  if (n === undefined || n === null || isNaN(n)) return "₹0";
  const abs = Math.abs(n);
  if (abs >= 1e7) return "₹" + (n / 1e7).toFixed(2) + " Cr";
  if (abs >= 1e5) return "₹" + (n / 1e5).toFixed(2) + " L";
  return "₹" + Math.round(n).toLocaleString("en-IN");
}

const DEFAULT_ALLOC = {
  fd:      { pct: 20, returnRate: 7.0 },
  scss:    { pct: 15, returnRate: 8.2 },
  debtmf:  { pct: 20, returnRate: 7.5 },
  equitymf:{ pct: 35, returnRate: 11.0 },
  gold:    { pct: 10, returnRate: 8.0 },
};

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

const DEFAULT_EXPENSES = [
  { id: "e1", label: "Groceries & household", amount: 20000 },
  { id: "e2", label: "Utilities (electricity, water, gas)", amount: 6000 },
  { id: "e3", label: "Rent / maintenance / property tax", amount: 10000 },
  { id: "e4", label: "Healthcare & insurance premiums", amount: 15000 },
  { id: "e5", label: "Transport & fuel", amount: 6000 },
  { id: "e6", label: "Leisure, travel & dining out", amount: 15000 },
  { id: "e7", label: "Miscellaneous / family support", amount: 8000 },
];

const DEFAULT_INCOME = [
  { id: "i1", label: "Pension", amount: 0 },
  { id: "i2", label: "Rental income", amount: 0 },
];

export default function RetirementPlanner() {
  const STORAGE_KEY = "retirement-planner-state";

  // Load from localStorage or use defaults
  const loadState = () => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      return saved ? JSON.parse(saved) : null;
    } catch {
      return null;
    }
  };

  const initialState = loadState();

  const [corpus, setCorpus] = useState(initialState?.corpus ?? 30000000);
  const [expenseItems, setExpenseItems] = useState(initialState?.expenseItems ?? DEFAULT_EXPENSES);
  const [inflation, setInflation] = useState(initialState?.inflation ?? 6);
  const [years, setYears] = useState(initialState?.years ?? 30);
  const [incomeItems, setIncomeItems] = useState(initialState?.incomeItems ?? DEFAULT_INCOME);
  const [alloc, setAlloc] = useState(initialState?.alloc ?? DEFAULT_ALLOC);
  const [actuals, setActuals] = useState(initialState?.actuals ?? {});
  const [selectedMonth, setSelectedMonth] = useState(initialState?.selectedMonth ?? new Date().getMonth());
  const [collapsed, setCollapsed] = useState(initialState?.collapsed ?? {
    expenses: false, income: false, allocation: false, budgetVsActual: false,
  });
  // Autosave is ON only if we actually found saved data on load (i.e. it was left on last time).
  // If nothing was in localStorage, default to ON so first-time users get autosave by default.
  const [autosave, setAutosave] = useState(true);

  function toggleCollapse(key) {
    setCollapsed((prev) => ({ ...prev, [key]: !prev[key] }));
  }

  // Save to localStorage whenever any state changes
  const saveState = () => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify({
        corpus, expenseItems, inflation, years, incomeItems, alloc, actuals, selectedMonth, collapsed,
      }));
    } catch (e) {
      console.warn("Failed to save to localStorage:", e);
    }
  };

  function toggleAutosave() {
    setAutosave((prev) => {
      const next = !prev;
      if (!next) {
        // Turning autosave OFF: wipe any previously saved data
        try {
          localStorage.removeItem(STORAGE_KEY);
        } catch (e) {
          console.warn("Failed to clear localStorage:", e);
        }
      } else {
        // Turning autosave ON: immediately persist current state
        saveState();
      }
      return next;
    });
  }

  // Save on every state change, but only while autosave is enabled
  useEffect(() => {
    if (autosave) saveState();
  }, [corpus, expenseItems, inflation, years, incomeItems, alloc, actuals, selectedMonth, collapsed, autosave]);

  const monthlyExpense = expenseItems.reduce((s, e) => s + Number(e.amount || 0), 0);
  const monthlyOtherIncome = incomeItems.reduce((s, i) => s + Number(i.amount || 0), 0);
  const otherIncome = monthlyOtherIncome * 12; // annualized, for the tax/simulation engine

  function updateIncome(id, field, value) {
    setIncomeItems((prev) => prev.map((i) => (i.id === id ? { ...i, [field]: field === "amount" ? Number(value) : value } : i)));
  }
  function addIncome() {
    setIncomeItems((prev) => [...prev, { id: `i${Date.now()}`, label: "New income source", amount: 0 }]);
  }
  function removeIncome(id) {
    setIncomeItems((prev) => prev.filter((i) => i.id !== id));
  }

  function updateActual(monthIdx, categoryId, value) {
    setActuals((prev) => ({
      ...prev,
      [monthIdx]: { ...(prev[monthIdx] || {}), [categoryId]: Number(value) },
    }));
  }
  function copyBudgetAsActual(monthIdx) {
    const copy = {};
    expenseItems.forEach((e) => { copy[e.id] = Number(e.amount); });
    setActuals((prev) => ({ ...prev, [monthIdx]: copy }));
  }
  function getActual(monthIdx, categoryId) {
    return actuals[monthIdx]?.[categoryId] ?? "";
  }

  const budgetVsActualRows = expenseItems.map((e) => {
    const actual = Number(actuals[selectedMonth]?.[e.id] || 0);
    return { id: e.id, label: e.label, budget: Number(e.amount), actual, variance: actual - Number(e.amount) };
  });
  const monthTotalBudget = monthlyExpense;
  const monthTotalActual = budgetVsActualRows.reduce((s, r) => s + r.actual, 0);
  const monthTotalVariance = monthTotalActual - monthTotalBudget;

  const trendData = MONTHS.map((m, i) => {
    const hasData = !!actuals[i] && Object.keys(actuals[i]).length > 0;
    const actualTotal = hasData ? Object.values(actuals[i]).reduce((s, v) => s + Number(v || 0), 0) : null;
    return { month: m.slice(0, 3), budget: monthlyExpense, actual: actualTotal };
  });

  function updateExpense(id, field, value) {
    setExpenseItems((prev) => prev.map((e) => (e.id === id ? { ...e, [field]: field === "amount" ? Number(value) : value } : e)));
  }
  function addExpense() {
    setExpenseItems((prev) => [...prev, { id: `e${Date.now()}`, label: "New expense", amount: 0 }]);
  }
  function removeExpense(id) {
    setExpenseItems((prev) => prev.filter((e) => e.id !== id));
  }

  const totalPct = Object.values(alloc).reduce((s, a) => s + Number(a.pct), 0);

  const { rows, depletionYear } = useMemo(
    () => runSimulation({ corpus, monthlyExpense, inflation, years, otherIncome, allocations: alloc }),
    [corpus, monthlyExpense, inflation, years, otherIncome, alloc]
  );

  const pieData = Object.keys(alloc).map((k) => ({
    name: ASSET_META[k].label, value: Number(alloc[k].pct), color: ASSET_META[k].color,
  }));

  const blendedPreTax = Object.keys(alloc).reduce((s, k) => s + (alloc[k].pct / 100) * alloc[k].returnRate, 0);
  const year1 = rows[0] || {};
  const blendedPostTax = year1.investmentGrowth ? blendedPreTax * (year1.investmentPostTaxGrowth / year1.investmentGrowth) : blendedPreTax;

  function updateAlloc(key, field, value) {
    setAlloc((prev) => ({ ...prev, [key]: { ...prev[key], [field]: Number(value) } }));
  }

  const safeWithdrawalRate = ((monthlyExpense * 12) / corpus * 100).toFixed(2);

  function downloadExcel() {
    const wb = XLSX.utils.book_new();

    // --- Sheet 1: Inputs ---
    const inputsData = [
      ["Retirement Corpus Planner — Inputs", ""],
      ["Generated on", new Date().toLocaleDateString("en-IN")],
      [],
      ["Parameter", "Value"],
      ["Retirement corpus (₹)", corpus],
      ["Monthly household expense, today's value (₹)", monthlyExpense],
      ["Expected inflation (% p.a.)", inflation],
      ["Projection horizon (years)", years],
      ["Other monthly income — pension/rent/etc. (₹)", monthlyOtherIncome],
      [],
      ["Derived figures", ""],
      ["Blended pre-tax return (%)", Number(blendedPreTax.toFixed(2))],
      ["Blended post-tax return, Year 1 (%)", Number(blendedPostTax.toFixed(2))],
      ["Year-1 withdrawal rate (%)", Number(safeWithdrawalRate)],
      ["Corpus depletes in year", depletionYear || `Lasts full ${years}-yr horizon`],
    ];
    const wsInputs = XLSX.utils.aoa_to_sheet(inputsData);
    wsInputs["!cols"] = [{ wch: 42 }, { wch: 22 }];
    XLSX.utils.book_append_sheet(wb, wsInputs, "Inputs");

    // --- Sheet 2: Monthly expense breakdown ---
    const expenseHeader = ["Expense item", "Monthly amount (₹)", "Annual amount (₹)"];
    const expenseRows = expenseItems.map((e) => [e.label, Number(e.amount), Number(e.amount) * 12]);
    const wsExpense = XLSX.utils.aoa_to_sheet([
      expenseHeader, ...expenseRows, [],
      ["Total", monthlyExpense, monthlyExpense * 12],
    ]);
    wsExpense["!cols"] = [{ wch: 34 }, { wch: 18 }, { wch: 18 }];
    XLSX.utils.book_append_sheet(wb, wsExpense, "Monthly Expenses");

    // --- Sheet 3: Monthly income breakdown ---
    const incomeHeader = ["Income source", "Monthly amount (₹)", "Annual amount (₹)"];
    const incomeRows = incomeItems.map((i) => [i.label, Number(i.amount), Number(i.amount) * 12]);
    const wsIncome = XLSX.utils.aoa_to_sheet([
      incomeHeader, ...incomeRows, [],
      ["Total", monthlyOtherIncome, otherIncome],
    ]);
    wsIncome["!cols"] = [{ wch: 34 }, { wch: 18 }, { wch: 18 }];
    XLSX.utils.book_append_sheet(wb, wsIncome, "Monthly Income");

    // --- Sheet 4: Budget vs Actual (all months with data) ---
    const bvaHeader = ["Month", "Expense item", "Budgeted (₹)", "Actual (₹)", "Variance (₹)"];
    const bvaRows = [];
    MONTHS.forEach((m, i) => {
      if (actuals[i] && Object.keys(actuals[i]).length > 0) {
        expenseItems.forEach((e) => {
          const actual = Number(actuals[i][e.id] || 0);
          bvaRows.push([m, e.label, Number(e.amount), actual, actual - Number(e.amount)]);
        });
        const totalActual = expenseItems.reduce((s, e) => s + Number(actuals[i][e.id] || 0), 0);
        bvaRows.push([m, "TOTAL", monthlyExpense, totalActual, totalActual - monthlyExpense]);
        bvaRows.push([]);
      }
    });
    if (bvaRows.length === 0) bvaRows.push(["No actuals entered yet", "", "", "", ""]);
    const wsBva = XLSX.utils.aoa_to_sheet([bvaHeader, ...bvaRows]);
    wsBva["!cols"] = [{ wch: 12 }, { wch: 34 }, { wch: 16 }, { wch: 16 }, { wch: 16 }];
    XLSX.utils.book_append_sheet(wb, wsBva, "Budget vs Actual");

    // --- Sheet 5: Allocation & tax treatment ---
    const allocHeader = ["Asset", "Allocation (%)", "Expected return (% p.a.)", "Amount allocated (₹)", "Tax treatment (FY 2026-27)"];
    const taxNotes = {
      fd: "Interest taxed yearly at slab rate",
      scss: "Interest taxed yearly at slab rate; ₹30L cap/person",
      debtmf: "Gains taxed at slab rate regardless of holding period (Sec 50AA)",
      equitymf: "LTCG @ 12.5% above ₹1.25L exemption/yr (Sec 112A); STCG @ 20% if <12mo",
      gold: "LTCG @ 12.5%, no exemption threshold (Sec 112)",
    };
    const allocRows = Object.keys(alloc).map((k) => [
      ASSET_META[k].label,
      alloc[k].pct,
      alloc[k].returnRate,
      Math.round(corpus * (alloc[k].pct / 100)),
      taxNotes[k],
    ]);
    const wsAlloc = XLSX.utils.aoa_to_sheet([allocHeader, ...allocRows, [], ["Total", totalPct, "", corpus, ""]]);
    wsAlloc["!cols"] = [{ wch: 22 }, { wch: 14 }, { wch: 20 }, { wch: 18 }, { wch: 55 }];
    XLSX.utils.book_append_sheet(wb, wsAlloc, "Allocation");

    // --- Sheet 6: Year-by-year projection ---
    const projHeader = ["Year", "Expense for the year (₹)", "Pre-tax growth (₹)", "Tax paid (₹)", "Post-tax growth (₹)", "Corpus balance, year-end (₹)"];
    const projRows = rows.map((r) => [r.year, r.expense || 0, r.preTaxGrowth || 0, r.tax || 0, r.postTaxGrowth || 0, r.corpus]);
    const wsProj = XLSX.utils.aoa_to_sheet([projHeader, ...projRows]);
    wsProj["!cols"] = [{ wch: 6 }, { wch: 20 }, { wch: 18 }, { wch: 14 }, { wch: 18 }, { wch: 22 }];
    XLSX.utils.book_append_sheet(wb, wsProj, "Year-by-year Projection");

    const filename = `retirement-corpus-plan-${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(wb, filename);
  }

  return (
    <div style={styles.page}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:wght@500;600;700&family=IBM+Plex+Mono:wght@400;500;600&family=Inter:wght@400;500;600&display=swap');
        * { box-sizing: border-box; }
        input[type=range] { -webkit-appearance: none; height: 4px; background: #3a4557; border-radius: 2px; }
        input[type=range]::-webkit-slider-thumb { -webkit-appearance: none; width: 14px; height: 14px; border-radius: 50%; background: #C79A45; cursor: pointer; border: 2px solid #F7F3E9; }
        input[type=number] { font-family: 'IBM Plex Mono', monospace; }
        ::selection { background: #C79A45; color: #12203a; }
      `}</style>

      <header style={styles.header}>
        <div style={styles.topRow}>
          <div style={styles.ledgerTab}>LEDGER · FY 2026-27</div>
          <button
            onClick={toggleAutosave}
            style={{ ...styles.autosaveBtn, ...(autosave ? styles.autosaveBtnOn : styles.autosaveBtnOff) }}
            aria-pressed={autosave}
          >
            <span style={{ ...styles.autosaveDot, background: autosave ? "#5FA777" : "#7C93A6" }} />
            Autosave: {autosave ? "On" : "Off"}
          </button>
        </div>
        <h1 style={styles.title}>Retirement Corpus Planner</h1>
        <p style={styles.subtitle}>
          Allocate the corpus across FD, SCSS, Debt & Equity MF and Gold — see the
          post-tax, inflation-adjusted drawdown play out, year by ledger year.
        </p>
        {!autosave && (
          <p style={styles.autosaveWarning}>
            Autosave is off — your changes will not be saved, and any previously saved data has been cleared.
            Turn autosave back on to persist your inputs in this browser.
          </p>
        )}
        <button onClick={downloadExcel} style={styles.downloadBtn}>
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ marginRight: 8 }}>
            <path d="M12 3v12m0 0l-4-4m4 4l4-4M4 17v2a2 2 0 002 2h12a2 2 0 002-2v-2" strokeLinecap="round" strokeLinejoin="round" />
          </svg>
          Download as Excel
        </button>
        <button
          onClick={() => {
            if (window.confirm("Reset all data to defaults? This cannot be undone.")) {
              localStorage.removeItem(STORAGE_KEY);
              window.location.reload();
            }
          }}
          style={styles.resetBtn}
        >
          Reset to defaults
        </button>
      </header>

      <div style={styles.grid}>
        {/* LEFT: inputs */}
        <section style={styles.card}>
          <h2 style={styles.cardTitle}>01 · The Corpus</h2>

          <Field label="Retirement corpus" value={corpus} onChange={setCorpus} display={inr(corpus)} min={1000000} max={100000000} step={100000} />
          <Field label="Expected inflation" value={inflation} onChange={setInflation} display={inflation + "%"} min={2} max={10} step={0.5} />
          <Field label="Projection horizon" value={years} onChange={setYears} display={years + " yrs"} min={5} max={40} step={1} />

          <div style={styles.divider} />
          <div style={styles.expenseHeaderRow}>
            <div style={styles.sectionTitleRow} onClick={() => toggleCollapse("expenses")}>
              <CollapseChevron collapsed={collapsed.expenses} />
              <h2 style={{ ...styles.cardTitle, margin: 0 }}>02 · Monthly Household Expenses</h2>
            </div>
            <span style={styles.expenseTotal}>{inr(monthlyExpense)}/mo</span>
          </div>

          {!collapsed.expenses && (
            <>
              {expenseItems.map((e) => (
                <ExpenseRow
                  key={e.id}
                  label={e.label}
                  amount={e.amount}
                  onLabel={(v) => updateExpense(e.id, "label", v)}
                  onAmount={(v) => updateExpense(e.id, "amount", v)}
                  onRemove={() => removeExpense(e.id)}
                />
              ))}
              <button onClick={addExpense} style={styles.addExpenseBtn}>+ Add expense line</button>
            </>
          )}

          <div style={styles.divider} />
          <div style={styles.expenseHeaderRow}>
            <div style={styles.sectionTitleRow} onClick={() => toggleCollapse("income")}>
              <CollapseChevron collapsed={collapsed.income} />
              <h2 style={{ ...styles.cardTitle, margin: 0 }}>03 · Other Monthly Income</h2>
            </div>
            <span style={styles.expenseTotal}>{inr(monthlyOtherIncome)}/mo</span>
          </div>

          {!collapsed.income && (
            <>
              <p style={styles.helperText}>Pension, rental income, annuity payouts, part-time consulting, etc.</p>
              {incomeItems.map((i) => (
                <ExpenseRow
                  key={i.id}
                  label={i.label}
                  amount={i.amount}
                  onLabel={(v) => updateIncome(i.id, "label", v)}
                  onAmount={(v) => updateIncome(i.id, "amount", v)}
                  onRemove={() => removeIncome(i.id)}
                />
              ))}
              <button onClick={addIncome} style={styles.addExpenseBtn}>+ Add income source</button>
            </>
          )}

          <div style={styles.divider} />
          <div style={styles.sectionTitleRow} onClick={() => toggleCollapse("allocation")}>
            <CollapseChevron collapsed={collapsed.allocation} />
            <h2 style={{ ...styles.cardTitle, margin: 0 }}>04 · The Allocation</h2>
          </div>

          {!collapsed.allocation && (
            <>
              <div style={{ fontSize: 12, fontFamily: "'IBM Plex Mono', monospace", color: totalPct === 100 ? "#5FA777" : "#C1594B", margin: "12px 0 10px" }}>
                Total allocated: {totalPct}% {totalPct !== 100 && "— adjust to sum to 100%"}
              </div>

              {Object.keys(alloc).map((k) => (
                <AllocRow
                  key={k}
                  label={ASSET_META[k].label}
                  color={ASSET_META[k].color}
                  pct={alloc[k].pct}
                  returnRate={alloc[k].returnRate}
                  onPct={(v) => updateAlloc(k, "pct", v)}
                  onReturn={(v) => updateAlloc(k, "returnRate", v)}
                />
              ))}
            </>
          )}
        </section>

        {/* RIGHT: outputs */}
        <section style={styles.outputCol}>
          <div style={styles.statRow}>
            <Stat label="Blended pre-tax return" value={blendedPreTax.toFixed(2) + "%"} />
            <Stat label="Blended post-tax return" value={blendedPostTax.toFixed(2) + "%"} accent />
            <Stat label="Year-1 withdrawal rate" value={safeWithdrawalRate + "%"} />
            <Stat
              label="Corpus lasts"
              value={depletionYear ? `${depletionYear} yrs` : `${years}+ yrs`}
              warn={!!depletionYear}
            />
          </div>

          <div style={styles.stickyCard}>
            <h2 style={styles.cardTitle}>05 · Corpus Over Time</h2>
            <ResponsiveContainer width="100%" height={280}>
              <AreaChart data={rows} margin={{ top: 10, right: 20, left: 0, bottom: 0 }}>
                <defs>
                  <linearGradient id="corpusFill" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor="#C79A45" stopOpacity={0.5} />
                    <stop offset="100%" stopColor="#C79A45" stopOpacity={0.02} />
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="2 4" stroke="#2a3548" vertical={false} />
                <XAxis dataKey="year" stroke="#7C93A6" fontSize={11} tickLine={false} label={{ value: "Year", position: "insideBottom", offset: -2, fill: "#7C93A6", fontSize: 11 }} />
                <YAxis stroke="#7C93A6" fontSize={11} tickFormatter={(v) => inr(v)} width={70} tickLine={false} />
                <Tooltip
                  contentStyle={{ background: "#1B2333", border: "1px solid #3a4557", borderRadius: 6, fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }}
                  labelFormatter={(l) => `Year ${l}`}
                  formatter={(v, n) => [inr(v), n === "corpus" ? "Corpus balance" : n]}
                />
                <Area type="monotone" dataKey="corpus" stroke="#C79A45" strokeWidth={2} fill="url(#corpusFill)" />
              </AreaChart>
            </ResponsiveContainer>
          </div>
          {depletionYear && (
            <p style={{ ...styles.warning, marginTop: -8 }}>
              ⚠ At this allocation and withdrawal rate, the corpus is projected to deplete in year {depletionYear}.
              Consider a lower withdrawal rate, higher equity allocation, or reduced expenses.
            </p>
          )}

          <div style={styles.card}>
            <div style={styles.expenseHeaderRow}>
              <div style={styles.sectionTitleRow} onClick={() => toggleCollapse("budgetVsActual")}>
                <CollapseChevron collapsed={collapsed.budgetVsActual} />
                <h2 style={{ ...styles.cardTitle, margin: 0 }}>06 · Budget vs Actual Spend</h2>
              </div>
              <select
                value={selectedMonth}
                onChange={(e) => setSelectedMonth(Number(e.target.value))}
                onClick={(e) => e.stopPropagation()}
                style={styles.monthSelect}
              >
                {MONTHS.map((m, i) => <option key={m} value={i}>{m}</option>)}
              </select>
            </div>

            {!collapsed.budgetVsActual && (
              <>
                <button onClick={() => copyBudgetAsActual(selectedMonth)} style={styles.addExpenseBtn}>
                  Copy budget as actual for {MONTHS[selectedMonth]}
                </button>

                <table style={{ ...styles.table, marginTop: 14 }}>
                  <thead>
                    <tr>
                      <th style={styles.th}>Category</th>
                      <th style={styles.th}>Budget</th>
                      <th style={styles.th}>Actual</th>
                      <th style={styles.th}>Variance</th>
                    </tr>
                  </thead>
                  <tbody>
                    {budgetVsActualRows.map((r) => (
                      <tr key={r.id}>
                        <td style={styles.td}>{r.label}</td>
                        <td style={styles.td}>{inr(r.budget)}</td>
                        <td style={styles.td}>
                          <input
                            type="number"
                            value={getActual(selectedMonth, r.id)}
                            onChange={(e) => updateActual(selectedMonth, r.id, e.target.value)}
                            placeholder="0"
                            style={styles.actualInput}
                          />
                        </td>
                        <td style={{ ...styles.td, color: r.variance > 0 ? "#C1594B" : r.variance < 0 ? "#5FA777" : "#7C93A6" }}>
                          {r.variance > 0 ? "+" : ""}{inr(r.variance)}
                        </td>
                      </tr>
                    ))}
                    <tr>
                      <td style={{ ...styles.td, fontWeight: 600, borderTop: "1px solid #2a3548" }}>Total</td>
                      <td style={{ ...styles.td, fontWeight: 600, borderTop: "1px solid #2a3548" }}>{inr(monthTotalBudget)}</td>
                      <td style={{ ...styles.td, fontWeight: 600, borderTop: "1px solid #2a3548" }}>{inr(monthTotalActual)}</td>
                      <td style={{ ...styles.td, fontWeight: 600, borderTop: "1px solid #2a3548", color: monthTotalVariance > 0 ? "#C1594B" : monthTotalVariance < 0 ? "#5FA777" : "#7C93A6" }}>
                        {monthTotalVariance > 0 ? "+" : ""}{inr(monthTotalVariance)}
                      </td>
                    </tr>
                  </tbody>
                </table>

                <ResponsiveContainer width="100%" height={200} style={{ marginTop: 18 }}>
                  <BarChart data={budgetVsActualRows} margin={{ left: 0, right: 10 }}>
                    <CartesianGrid strokeDasharray="2 4" stroke="#2a3548" vertical={false} />
                    <XAxis dataKey="label" stroke="#7C93A6" fontSize={9} tickLine={false} interval={0} angle={-20} textAnchor="end" height={60} />
                    <YAxis stroke="#7C93A6" fontSize={10} tickFormatter={(v) => inr(v)} width={60} tickLine={false} />
                    <Tooltip contentStyle={{ background: "#1B2333", border: "1px solid #3a4557", borderRadius: 6, fontSize: 12 }} formatter={(v) => inr(v)} />
                    <Legend wrapperStyle={{ fontSize: 11 }} />
                    <Bar dataKey="budget" name="Budget" fill="#7C93A6" radius={[3, 3, 0, 0]} />
                    <Bar dataKey="actual" name="Actual" fill="#C79A45" radius={[3, 3, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>

                <h3 style={styles.subCardTitle}>Year trend — total budget vs total actual</h3>
                <ResponsiveContainer width="100%" height={180}>
                  <LineChart data={trendData} margin={{ left: 0, right: 10 }}>
                    <CartesianGrid strokeDasharray="2 4" stroke="#2a3548" vertical={false} />
                    <XAxis dataKey="month" stroke="#7C93A6" fontSize={10} tickLine={false} />
                    <YAxis stroke="#7C93A6" fontSize={10} tickFormatter={(v) => inr(v)} width={60} tickLine={false} />
                    <Tooltip contentStyle={{ background: "#1B2333", border: "1px solid #3a4557", borderRadius: 6, fontSize: 12 }} formatter={(v) => (v === null ? "No data" : inr(v))} />
                    <Legend wrapperStyle={{ fontSize: 11 }} />
                    <Line type="monotone" dataKey="budget" name="Budget" stroke="#7C93A6" strokeWidth={2} dot={false} />
                    <Line type="monotone" dataKey="actual" name="Actual" stroke="#C79A45" strokeWidth={2} connectNulls={false} dot={{ r: 3 }} />
                  </LineChart>
                </ResponsiveContainer>
                <p style={styles.footnote}>Months with no actuals entered show a gap in the "Actual" line rather than zero.</p>
              </>
            )}
          </div>

          <div style={styles.twoCol}>
            <div style={styles.card}>
              <h2 style={styles.cardTitle}>07 · Allocation Mix</h2>
              <ResponsiveContainer width="100%" height={220}>
                <PieChart>
                  <Pie data={pieData} dataKey="value" nameKey="name" innerRadius={45} outerRadius={80} paddingAngle={2}>
                    {pieData.map((entry, i) => <Cell key={i} fill={entry.color} stroke="#12203a" />)}
                  </Pie>
                  <Tooltip contentStyle={{ background: "#1B2333", border: "1px solid #3a4557", borderRadius: 6, fontSize: 12 }} formatter={(v, n) => [v + "%", n]} />
                  <Legend wrapperStyle={{ fontSize: 11, fontFamily: "Inter, sans-serif" }} />
                </PieChart>
              </ResponsiveContainer>
            </div>

            <div style={styles.card}>
              <h2 style={styles.cardTitle}>08 · Year 1 Tax Drag</h2>
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={[
                  { name: "Pre-tax growth", value: year1.preTaxGrowth || 0 },
                  { name: "Tax paid", value: year1.tax || 0 },
                  { name: "Post-tax growth", value: year1.postTaxGrowth || 0 },
                ]} margin={{ left: 0, right: 10 }}>
                  <CartesianGrid strokeDasharray="2 4" stroke="#2a3548" vertical={false} />
                  <XAxis dataKey="name" stroke="#7C93A6" fontSize={10} tickLine={false} />
                  <YAxis stroke="#7C93A6" fontSize={10} tickFormatter={(v) => inr(v)} width={60} tickLine={false} />
                  <Tooltip contentStyle={{ background: "#1B2333", border: "1px solid #3a4557", borderRadius: 6, fontSize: 12 }} formatter={(v) => inr(v)} />
                  <Bar dataKey="value" fill="#5FA777" radius={[3, 3, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          <div style={styles.card}>
            <h2 style={styles.cardTitle}>09 · How Each Asset Is Taxed (FY 2026-27)</h2>
            <table style={styles.table}>
              <thead>
                <tr>
                  <th style={styles.th}>Asset</th>
                  <th style={styles.th}>Tax treatment</th>
                  <th style={styles.th}>Effective drag</th>
                </tr>
              </thead>
              <tbody>
                <tr><td style={styles.td}>Fixed Deposit</td><td style={styles.td}>Interest taxed yearly at your income-tax slab rate</td><td style={styles.td}>High (up to 30%)</td></tr>
                <tr><td style={styles.td}>SCSS</td><td style={styles.td}>Interest taxed yearly at slab rate; ₹30L cap per person</td><td style={styles.td}>High (up to 30%)</td></tr>
                <tr><td style={styles.td}>Debt Mutual Fund</td><td style={styles.td}>Gains taxed at slab rate regardless of holding period (Sec 50AA)</td><td style={styles.td}>High (up to 30%)</td></tr>
                <tr><td style={styles.td}>Equity MF (via SWP)</td><td style={styles.td}>LTCG @ 12.5% above ₹1.25L exemption/yr (Sec 112A); STCG @ 20% if held &lt;12mo</td><td style={styles.td}>Low</td></tr>
                <tr><td style={styles.td}>Gold Fund / ETF</td><td style={styles.td}>LTCG @ 12.5%, no exemption threshold (Sec 112)</td><td style={styles.td}>Moderate</td></tr>
              </tbody>
            </table>
            <p style={styles.footnote}>
              Rates per Finance (No.2) Act 2024, unchanged in Budget 2025 and Budget 2026 for FY 2026-27.
              New tax regime slabs used; ₹12L taxable income is effectively tax-free after Sec 87A rebate.
              This is a planning model, not tax or investment advice — actual fund-level gains, indexation history,
              and surcharge can change the numbers. Verify with a CA/CFP before acting.
            </p>
          </div>
        </section>
      </div>
    </div>
  );
}

function CollapseChevron({ collapsed }) {
  return (
    <svg
      width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#C79A45" strokeWidth="2.5"
      style={{ transform: collapsed ? "rotate(-90deg)" : "rotate(0deg)", transition: "transform 0.15s ease", flexShrink: 0 }}
    >
      <path d="M6 9l6 6 6-6" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

function Field({ label, value, onChange, display, min, max, step }) {
  return (
    <div style={styles.field}>
      <div style={styles.fieldLabelRow}>
        <label style={styles.fieldLabel}>{label}</label>
        <span style={styles.fieldValue}>{display}</span>
      </div>
      <input type="range" min={min} max={max} step={step} value={value} onChange={(e) => onChange(Number(e.target.value))} style={{ width: "100%" }} />
    </div>
  );
}

function ExpenseRow({ label, amount, onLabel, onAmount, onRemove }) {
  return (
    <div style={styles.expenseRow}>
      <input
        type="text"
        value={label}
        onChange={(e) => onLabel(e.target.value)}
        style={styles.expenseLabelInput}
        placeholder="Expense label"
      />
      <div style={styles.allocInputWrap}>
        <span style={styles.unit}>₹</span>
        <input
          type="number"
          value={amount}
          onChange={(e) => onAmount(e.target.value)}
          style={{ ...styles.smallInput, width: 68 }}
        />
      </div>
      <button onClick={onRemove} style={styles.removeBtn} aria-label={`Remove ${label}`}>×</button>
    </div>
  );
}

function AllocRow({ label, color, pct, returnRate, onPct, onReturn }) {
  return (
    <div style={styles.allocRow}>
      <div style={styles.allocLabel}><span style={{ ...styles.dot, background: color }} />{label}</div>
      <div style={styles.allocInputs}>
        <div style={styles.allocInputWrap}>
          <input type="number" value={pct} onChange={(e) => onPct(e.target.value)} style={styles.smallInput} />
          <span style={styles.unit}>%</span>
        </div>
        <div style={styles.allocInputWrap}>
          <input type="number" step="0.1" value={returnRate} onChange={(e) => onReturn(e.target.value)} style={styles.smallInput} />
          <span style={styles.unit}>% p.a.</span>
        </div>
      </div>
    </div>
  );
}

function Stat({ label, value, accent, warn }) {
  return (
    <div style={styles.stat}>
      <div style={styles.statLabel}>{label}</div>
      <div style={{ ...styles.statValue, color: warn ? "#C1594B" : accent ? "#5FA777" : "#F7F3E9" }}>{value}</div>
    </div>
  );
}

const styles = {
  page: { minHeight: "100%", background: "#12203a", color: "#F7F3E9", fontFamily: "'Inter', sans-serif", padding: "28px 24px 60px", },
  header: { maxWidth: 1100, margin: "0 auto 28px" },
  topRow: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, marginBottom: 12 },
  ledgerTab: { display: "inline-block", fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, letterSpacing: "0.12em", color: "#C79A45", border: "1px solid #C79A45", borderRadius: 3, padding: "3px 8px" },
  autosaveBtn: { display: "inline-flex", alignItems: "center", gap: 8, background: "transparent", border: "1px solid #3a4557", borderRadius: 20, padding: "6px 14px 6px 10px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, fontWeight: 600, letterSpacing: "0.02em", cursor: "pointer", transition: "border-color 0.15s ease, color 0.15s ease" },
  autosaveBtnOn: { color: "#5FA777", borderColor: "rgba(95,167,119,0.45)" },
  autosaveBtnOff: { color: "#7C93A6", borderColor: "#3a4557" },
  autosaveDot: { width: 7, height: 7, borderRadius: "50%", display: "inline-block" },
  autosaveWarning: { fontSize: 12.5, color: "#E0A088", background: "rgba(193,89,75,0.12)", border: "1px solid rgba(193,89,75,0.35)", borderRadius: 6, padding: "8px 12px", marginTop: 14, maxWidth: 620, lineHeight: 1.5 },
  title: { fontFamily: "'Source Serif 4', serif", fontSize: 34, fontWeight: 600, margin: "0 0 8px", letterSpacing: "-0.01em" },
  subtitle: { color: "#9FB0C7", fontSize: 15, maxWidth: 620, lineHeight: 1.5, margin: 0 },
  downloadBtn: { display: "inline-flex", alignItems: "center", marginTop: 18, background: "#C79A45", color: "#12203a", border: "none", borderRadius: 6, padding: "10px 18px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, fontWeight: 600, letterSpacing: "0.02em", cursor: "pointer" },
  resetBtn: { display: "inline-flex", alignItems: "center", marginTop: 18, marginLeft: 12, background: "transparent", color: "#7C93A6", border: "1px solid #3a4557", borderRadius: 6, padding: "10px 18px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, fontWeight: 600, letterSpacing: "0.02em", cursor: "pointer" },
  grid: { maxWidth: 1100, margin: "0 auto", display: "grid", gridTemplateColumns: "340px 1fr", gap: 20, alignItems: "start" },
  card: { background: "#182338", border: "1px solid #2a3548", borderRadius: 10, padding: "18px 20px", marginBottom: 20 },
  stickyCard: { background: "#182338", border: "1px solid #2a3548", borderRadius: 10, padding: "18px 20px", marginBottom: 20, position: "sticky", top: 12, zIndex: 30, boxShadow: "0 8px 24px rgba(0,0,0,0.35)" },
  sectionTitleRow: { display: "flex", alignItems: "center", gap: 8, cursor: "pointer", userSelect: "none" },
  cardTitle: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, letterSpacing: "0.04em", color: "#C79A45", margin: "0 0 16px", textTransform: "uppercase" },
  divider: { height: 1, background: "#2a3548", margin: "20px 0" },
  field: { marginBottom: 18 },
  fieldLabelRow: { display: "flex", justifyContent: "space-between", marginBottom: 6 },
  fieldLabel: { fontSize: 13, color: "#C6D2E0" },
  fieldValue: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: "#C79A45" },
  outputCol: { minWidth: 0 },
  statRow: { display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12, marginBottom: 20 },
  stat: { background: "#182338", border: "1px solid #2a3548", borderRadius: 10, padding: "14px 16px" },
  statLabel: { fontSize: 11, color: "#7C93A6", marginBottom: 6, lineHeight: 1.3 },
  statValue: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 20, fontWeight: 600 },
  warning: { fontSize: 13, color: "#E0A088", background: "rgba(193,89,75,0.12)", border: "1px solid rgba(193,89,75,0.35)", borderRadius: 6, padding: "10px 14px", marginTop: 12 },
  twoCol: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 },
  allocRow: { display: "flex", alignItems: "center", justifyContent: "space-between", padding: "8px 0", borderBottom: "1px solid #232e42" },
  expenseHeaderRow: { display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 },
  expenseTotal: { fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: "#5FA777", fontWeight: 600 },
  helperText: { fontSize: 11.5, color: "#7C93A6", margin: "0 0 10px", lineHeight: 1.4 },
  expenseRow: { display: "flex", alignItems: "center", gap: 8, padding: "6px 0", borderBottom: "1px solid #232e42" },
  expenseLabelInput: { flex: 1, background: "transparent", border: "none", borderBottom: "1px dashed #2a3548", color: "#E6ECF3", fontSize: 12.5, padding: "4px 2px", outline: "none", minWidth: 0 },
  removeBtn: { background: "none", border: "none", color: "#7C93A6", fontSize: 18, lineHeight: 1, cursor: "pointer", padding: "0 2px" },
  addExpenseBtn: { width: "100%", marginTop: 8, background: "transparent", border: "1px dashed #3a4557", color: "#C79A45", borderRadius: 6, padding: "8px 0", fontSize: 12.5, fontFamily: "'IBM Plex Mono', monospace", cursor: "pointer" },
  monthSelect: { background: "#0F1928", border: "1px solid #2a3548", color: "#F7F3E9", borderRadius: 6, padding: "6px 10px", fontSize: 12.5, fontFamily: "'IBM Plex Mono', monospace", outline: "none" },
  actualInput: { width: 90, background: "#0F1928", border: "1px solid #2a3548", borderRadius: 5, color: "#F7F3E9", fontSize: 12.5, padding: "5px 8px", outline: "none" },
  subCardTitle: { fontSize: 12, color: "#7C93A6", margin: "18px 0 6px", fontFamily: "'IBM Plex Mono', monospace", letterSpacing: "0.02em" },
  allocLabel: { display: "flex", alignItems: "center", fontSize: 13, color: "#E6ECF3", gap: 8 },
  dot: { width: 8, height: 8, borderRadius: "50%", display: "inline-block" },
  allocInputs: { display: "flex", gap: 10 },
  allocInputWrap: { display: "flex", alignItems: "center", gap: 4, background: "#0F1928", border: "1px solid #2a3548", borderRadius: 5, padding: "2px 6px" },
  smallInput: { width: 44, background: "transparent", border: "none", color: "#F7F3E9", fontSize: 12, textAlign: "right", outline: "none" },
  unit: { fontSize: 10, color: "#7C93A6" },
  table: { width: "100%", borderCollapse: "collapse", fontSize: 13 },
  th: { textAlign: "left", color: "#7C93A6", fontWeight: 500, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.03em", padding: "6px 10px 10px 0", borderBottom: "1px solid #2a3548" },
  td: { padding: "9px 10px 9px 0", borderBottom: "1px solid #202a3d", color: "#D8E1EA", verticalAlign: "top" },
  footnote: { fontSize: 11.5, color: "#7C93A6", lineHeight: 1.6, marginTop: 14 },
};
