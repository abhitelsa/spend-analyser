const DOMAINS = [
  "Food",
  "Grocery",
  "Medicine",
  "Petrol",
  "Shopping",
  "Travel",
  "Entertainment",
  "Mobile tsky ott",
  "Loan, Rent",
  "Investment",
  "Transfers",
  "Misc",
];

const DOMAIN_COLORS = {
  "Food": "#c45c26",
  "Grocery": "#2f6b4f",
  "Medicine": "#9b2f2f",
  "Petrol": "#b08900",
  "Shopping": "#7a4e7a",
  "Travel": "#3d5a80",
  "Entertainment": "#8b5cf6",
  "Mobile tsky ott": "#1c7c7d",
  "Loan, Rent": "#8d6e63",
  "Investment": "#4a6fa5",
  "Transfers": "#6b4f3a",
  "Misc": "#857f72",
};

let rulesFromJson = [];

async function loadRulesJson() {
  try {
    const res = await fetch("rules.json");
    if (res.ok) {
      rulesFromJson = await res.json();
      console.log(`Loaded ${rulesFromJson.length} rules from rules.json`);
      updateRulesBadge();
      reapplyCategorization();
      if (allTxns.length) render();
    }
  } catch (e) {
    console.warn("Could not load rules.json (likely running via file:// protocol). Custom rules will still work.", e);
  }
}

const STORAGE_KEY = "spend_analyser_custom_rules";

function getCustomRules() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch (e) {
    return [];
  }
}

function saveCustomRules(rules) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(rules));
  } catch (e) {
    console.error("Failed to save rules to localStorage", e);
  }
  updateRulesBadge();
}

function addCustomRule(keyword, domain) {
  const kw = String(keyword || "").trim();
  if (!kw) return;
  let rules = getCustomRules();
  rules = rules.filter((r) => r.keyword.toLowerCase() !== kw.toLowerCase());
  rules.unshift({ keyword: kw, domain });
  saveCustomRules(rules);
  reapplyCategorization();
  render();
}

function deleteCustomRule(keyword) {
  let rules = getCustomRules();
  rules = rules.filter((r) => r.keyword.toLowerCase() !== keyword.toLowerCase());
  saveCustomRules(rules);
  reapplyCategorization();
  render();
}

const els = {
  dropzone: document.getElementById("dropzone"),
  fileInput: document.getElementById("file-input"),
  status: document.getElementById("status"),
  dashboard: document.getElementById("dashboard"),
  excludeTransfers: document.getElementById("exclude-transfers"),
  debitsOnly: document.getElementById("debits-only"),
  modeFilter: document.getElementById("mode-filter"),
  miscFilter: document.getElementById("misc-filter"),
  miscCount: document.getElementById("misc-count"),
  exportExcelBtn: document.getElementById("export-excel-btn"),
  rulesBtn: document.getElementById("rules-btn"),
  rulesCount: document.getElementById("rules-count"),
  rulesModal: document.getElementById("rules-modal"),
  closeRulesBtn: document.getElementById("close-rules-btn"),
  addRuleForm: document.getElementById("add-rule-form"),
  ruleKeyword: document.getElementById("rule-keyword"),
  ruleDomain: document.getElementById("rule-domain"),
  rulesList: document.getElementById("rules-list"),
  downloadRulesBtn: document.getElementById("download-rules-btn"),
  kpis: document.getElementById("kpis"),
  merchants: document.getElementById("merchants"),
  instruments: document.getElementById("instruments"),
  txnBody: document.getElementById("txn-body"),
  search: document.getElementById("search"),
  transfersChart: document.getElementById("transfers-chart"),
  investmentsChart: document.getElementById("investments-chart"),
  clearBtn: document.getElementById("clear-btn"),
};

let allTxns = [];
let statementMeta = { account: "", duration: "", fileName: "", isConsolidated: false };
let filterMiscOnly = false;

function inr(n) {
  return new Intl.NumberFormat("en-IN", {
    style: "currency",
    currency: "INR",
    maximumFractionDigits: 0,
  }).format(n || 0);
}

function inrExact(n) {
  return new Intl.NumberFormat("en-IN", {
    style: "currency",
    currency: "INR",
    maximumFractionDigits: 2,
  }).format(n || 0);
}

function normalizeHeader(h) {
  return String(h || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function parseAmount(v) {
  if (typeof v === "number") return v;
  const n = Number(String(v || "").replace(/[^0-9.-]/g, ""));
  return Number.isFinite(n) ? n : 0;
}

function parseDate(v) {
  if (v instanceof Date && !Number.isNaN(v.getTime())) return v;
  if (typeof v === "number" && v > 20000 && v < 60000) {
    // Excel serial date number
    return new Date((v - (25567 + 2)) * 86400 * 1000);
  }
  const s = String(v || "").trim();
  const d = new Date(s);
  if (!Number.isNaN(d.getTime())) return d;
  return null;
}

function formatInstrument(raw) {
  if (!raw) return "Unknown";
  const s = String(raw).trim();
  if (s.includes("653029")) return "cc-pp-hdfc";
  if (s.includes("2387")) return "hdfc-pp";
  if (s.includes("4526")) return "sbi-pp";
  if (s.includes("9928")) return "bank-9928";
  return s.replace(/^Paid by\s+/i, "").replace(/^Credited to\s+/i, "");
}

function categorize(details, type) {
  const text = `${details} ${type}`.toLowerCase();

  // 1. Custom rules saved in localStorage take top priority
  const customRules = getCustomRules();
  for (const r of customRules) {
    if (text.includes(r.keyword.toLowerCase())) {
      return r.domain;
    }
  }

  // 2. Rules loaded from rules.json
  for (const r of rulesFromJson) {
    if (r.keyword && text.includes(r.keyword.toLowerCase())) {
      return r.domain;
    }
  }

  // 3. Fallback to Misc
  return "Misc";
}

function reapplyCategorization() {
  for (const t of allTxns) {
    if (!t._manualOverride) {
      t.category = categorize(t.details, t.type);
    }
  }
}

function isTransferLike(txn) {
  return (
    txn.category === "Transfers" ||
    txn.category === "Investment" ||
    txn.status === "Transfer" ||
    txn.status === "Investment" ||
    txn.status === "CC Bill Repayment" ||
    txn.status === "ATM Withdrawal"
  );
}

function findHeaderRow(rows) {
  for (let i = 0; i < Math.min(rows.length, 25); i++) {
    const cells = rows[i].map(normalizeHeader);
    if (cells.includes("date") && cells.some((c) => c.includes("amount"))) return i;
    if (cells.includes("purpose") && cells.some((c) => c.includes("domain") || c.includes("mode"))) return i;
  }
  return -1;
}

function parseWorkbook(wb, fileName) {
  let targetSheetName = wb.SheetNames[0];
  if (wb.SheetNames.includes("Expenditure")) {
    targetSheetName = "Expenditure";
  } else if (wb.SheetNames.includes("Form Responses 5")) {
    targetSheetName = "Form Responses 5";
  }

  const sheet = wb.Sheets[targetSheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
  const headerIdx = findHeaderRow(rows);
  if (headerIdx < 0) throw new Error("Could not find a recognized header row with Date and Amount/Purpose.");

  const headers = rows[headerIdx].map(normalizeHeader);
  const col = (aliases) => headers.findIndex((h) => aliases.some((a) => h === a || h.includes(a)));

  const iDate = col(["date", "timestamp"]);
  const iTime = col(["time"]);
  const iDetails = col(["purpose", "transaction details", "details", "description", "narration", "merchant"]);
  const iId = col(["transaction id", "refid", "id"]);
  const iType = col(["transaction type", "type"]);
  const iInstrument = col(["mode", "payment mode", "credit/debit instrument", "instrument"]);
  const iDomain = col(["domain", "category"]);
  const iStatus = col(["status"]);
  const iSource = col(["source"]);
  const iAmount = col(["amount"]);
  const iDirection = col(["credit/debit", "dr/cr"]);

  const isConsolidated = iDomain >= 0 || targetSheetName === "Expenditure" || targetSheetName === "Form Responses 5";

  const txns = [];
  let txnIdCounter = 1;
  const modesFound = new Set();

  for (let r = headerIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.every((c) => String(c).trim() === "")) continue;
    const first = String(row[0] || "");
    if (/automatically generated|disclaimer|terms-conditions/i.test(first)) break;

    const details = String(row[iDetails] || "").trim();
    const amount = parseAmount(row[iAmount]);
    const date = parseDate(row[iDate]);
    if (!details || !amount || !date) continue;

    let direction = String(row[iType] || row[iDirection] || "").toUpperCase();
    if (!direction.includes("CREDIT") && !direction.includes("DEBIT")) {
      direction = amount < 0 ? "DEBIT" : "DEBIT";
    } else if (direction.includes("CREDIT")) {
      direction = "CREDIT";
    } else {
      direction = "DEBIT";
    }

    const rawInstrument = String(row[iInstrument] || "").trim();
    const mode = rawInstrument ? (isConsolidated ? rawInstrument : formatInstrument(rawInstrument)) : "Unknown";
    modesFound.add(mode);

    let assignedDomain = isConsolidated && row[iDomain] ? String(row[iDomain]).trim() : "";
    if (!assignedDomain) {
      assignedDomain = categorize(details, direction);
    }

    const status = isConsolidated && row[iStatus] ? String(row[iStatus]).trim() : (assignedDomain === "Transfers" ? "Transfer" : "Living Expense");
    const source = isConsolidated && row[iSource] ? String(row[iSource]).trim() : "Statement";

    const txn = {
      _id: txnIdCounter++,
      date,
      time: String(row[iTime] || "").trim(),
      details,
      id: String(row[iId] || "").trim(),
      type: direction,
      rawInstrument,
      instrument: mode,
      mode,
      category: assignedDomain,
      status,
      source,
      amount: Math.abs(amount),
      _manualOverride: isConsolidated,
    };

    txns.push(txn);
  }

  // Populate mode filter dropdown
  if (els.modeFilter) {
    const currentMode = els.modeFilter.value || "all";
    els.modeFilter.innerHTML = '<option value="all">All Payment Modes</option>' +
      [...modesFound].sort().map((m) => `<option value="${m}">${m}</option>`).join("");
    if (modesFound.has(currentMode)) {
      els.modeFilter.value = currentMode;
    }
  }

  const account = String(rows[0]?.[0] || "");
  const duration = String(rows[1]?.[1] || rows[1]?.[0] || "");
  return { txns, meta: { account, duration, fileName, isConsolidated } };
}

function filteredTxns() {
  const selectedMode = els.modeFilter ? els.modeFilter.value : "all";

  return allTxns.filter((t) => {
    if (els.debitsOnly.checked && t.type !== "DEBIT") return false;
    if (els.excludeTransfers.checked && isTransferLike(t)) return false;
    if (selectedMode !== "all" && t.mode !== selectedMode && t.instrument !== selectedMode) return false;
    if (filterMiscOnly && t.category !== "Misc") return false;
    return true;
  });
}

function groupSum(txns, keyFn) {
  const map = new Map();
  for (const t of txns) {
    const k = keyFn(t);
    map.set(k, (map.get(k) || 0) + t.amount);
  }
  return [...map.entries()].sort((a, b) => b[1] - a[1]);
}

function dayKey(d) {
  return d.toISOString().slice(0, 10);
}

function renderKpis(txns) {
  const spend = txns.filter((t) => t.type === "DEBIT").reduce((s, t) => s + t.amount, 0);
  const credit = allTxns.filter((t) => t.type === "CREDIT").reduce((s, t) => s + t.amount, 0);
  const days = new Set(txns.map((t) => dayKey(t.date))).size || 1;
  const avg = spend / days;
  const biggest = txns.reduce((m, t) => (t.amount > (m?.amount || 0) ? t : m), null);

  // Reconciliation data if available in dataset
  const atmTotal = allTxns.filter((t) => t.status === "ATM Withdrawal").reduce((s, t) => s + t.amount, 0);
  const ccBillsTotal = allTxns.filter((t) => t.status === "CC Bill Repayment").reduce((s, t) => s + t.amount, 0);

  const kpiItems = [
    ["Spend in view", inr(spend), `${txns.length} transactions`],
    ["Credits (all)", inr(credit), "Incoming funds & refunds"],
    ["Avg / active day", inr(avg), `${days} days with activity`],
    ["Largest in view", biggest ? inr(biggest.amount) : "—", biggest ? biggest.details : ""],
  ];

  if (statementMeta.isConsolidated && (atmTotal > 0 || ccBillsTotal > 0)) {
    kpiItems.push(
      ["CC Bill Repayments", inr(ccBillsTotal), "Filtered from living spend"],
      ["ATM Cash Drawn", inr(atmTotal), "Cash seeded from bank"]
    );
  }

  els.kpis.innerHTML = kpiItems
    .map(
      ([label, value, hint]) =>
        `<article class="kpi"><div class="label">${label}</div><div class="value">${value}</div><div class="hint">${hint}</div></article>`
    )
    .join("");
}

function renderRanks(el, rows, limit = 8) {
  el.innerHTML =
    rows
      .slice(0, limit)
      .map(([name, amt]) => `<li><span>${name}</span><span class="amt">${inr(amt)}</span></li>`)
      .join("") || "<li>No data</li>";
}

function renderTable(txns) {
  const q = els.search.value.trim().toLowerCase();
  const rows = txns.filter((t) => {
    if (!q) return true;
    return `${t.details} ${t.category} ${t.id} ${t.instrument} ${t.source || ""}`.toLowerCase().includes(q);
  });

  els.txnBody.innerHTML = rows
    .map((t) => {
      const when = t.date.toLocaleDateString("en-IN", { day: "2-digit", month: "short" });
      const cls = t.type === "CREDIT" ? "credit" : "debit";
      const sign = t.type === "CREDIT" ? "+" : "−";
      const optionsHtml = DOMAINS.map(
        (d) => `<option value="${d}" ${t.category === d ? "selected" : ""}>${d}</option>`
      ).join("");

      const sourceBadge = t.source ? `<span style="font-size:11px;color:var(--muted);margin-left:6px;">[${t.source}]</span>` : "";

      return `<tr>
        <td>${when}<div class="meta">${t.time}</div></td>
        <td>${t.details}${sourceBadge}<div class="meta">${t.instrument || t.id}</div></td>
        <td>
          <select class="category-select ${t.category === "Misc" ? "is-misc" : ""}" data-id="${t._id}">
            ${optionsHtml}
          </select>
        </td>
        <td>${t.type}</td>
        <td class="num ${cls}">${sign}${inrExact(t.amount)}</td>
      </tr>`;
    })
    .join("");
}

function drawBars(canvas, pairs) {
  const ctx = canvas.getContext("2d");
  const dpr = window.devicePixelRatio || 1;
  const cssW = canvas.clientWidth || 640;
  const cssH = 220;
  canvas.width = cssW * dpr;
  canvas.height = cssH * dpr;
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, cssW, cssH);
  if (!pairs.length) return;

  const pad = { t: 12, r: 8, b: 36, l: 8 };
  const w = cssW - pad.l - pad.r;
  const h = cssH - pad.t - pad.b;
  const max = Math.max(...pairs.map((p) => p[1]), 1);
  const gap = 4;
  const barW = Math.max(2, (w - gap * pairs.length) / pairs.length);

  pairs.forEach(([label, value], i) => {
    const x = pad.l + i * (barW + gap);
    const bh = (value / max) * h;
    const y = pad.t + h - bh;
    ctx.fillStyle = "#c45c26";
    ctx.globalAlpha = 0.85;
    ctx.fillRect(x, y, barW, bh);
    ctx.globalAlpha = 1;
    if (pairs.length <= 16 || i % Math.ceil(pairs.length / 10) === 0) {
      ctx.fillStyle = "#6b6458";
      ctx.font = "10px Instrument Sans, sans-serif";
      ctx.save();
      ctx.translate(x + barW / 2, cssH - 8);
      ctx.rotate(-0.6);
      ctx.fillText(label.slice(5), 0, 0);
      ctx.restore();
    }
  });
}

function drawDonut(canvas, pairs) {
  const ctx = canvas.getContext("2d");
  const dpr = window.devicePixelRatio || 1;
  const cssW = canvas.clientWidth || 400;
  const cssH = 220;
  canvas.width = cssW * dpr;
  canvas.height = cssH * dpr;
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, cssW, cssH);
  if (!pairs.length) return;

  const total = pairs.reduce((s, p) => s + p[1], 0) || 1;
  const cx = 90;
  const cy = cssH / 2;
  const r = 68;
  let a0 = -Math.PI / 2;
  pairs.forEach(([cat, value]) => {
    const a1 = a0 + (value / total) * Math.PI * 2;
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.arc(cx, cy, r, a0, a1);
    ctx.closePath();
    ctx.fillStyle = DOMAIN_COLORS[cat] || "#857f72";
    ctx.fill();
    a0 = a1;
  });
  ctx.beginPath();
  ctx.fillStyle = "#fffdf8";
  ctx.arc(cx, cy, 38, 0, Math.PI * 2);
  ctx.fill();

  ctx.font = "12px Instrument Sans, sans-serif";
  pairs.slice(0, 7).forEach(([name, amt], i) => {
    const y = 24 + i * 26;
    ctx.fillStyle = DOMAIN_COLORS[name] || "#857f72";
    ctx.fillRect(190, y - 8, 10, 10);
    ctx.fillStyle = "#1c1914";
    ctx.fillText(`${name} (${inr(amt)})`, 208, y);
  });
}

function updateRulesBadge() {
  const customCount = getCustomRules().length;
  const totalCount = rulesFromJson.length + customCount;
  if (els.rulesCount) els.rulesCount.textContent = totalCount;
}

function updateMiscBadge() {
  const count = allTxns.filter((t) => t.type === "DEBIT" && t.category === "Misc").length;
  if (els.miscCount) els.miscCount.textContent = count;
}

// Render Transfers chart
function renderTransfersChart() {
  const transfersTxns = filteredTxns().filter(t => t.category === "Transfers" && t.type === "DEBIT");
  const byDayMap = new Map();
  transfersTxns.forEach(t => {
    const k = dayKey(t.date);
    byDayMap.set(k, (byDayMap.get(k) || 0) + t.amount);
  });
  const pairs = [...byDayMap.entries()].sort((a, b) => a[0].localeCompare(b[0]));
  drawBars(els.transfersChart, pairs);
}

// Render Investments chart
function renderInvestmentsChart() {
  const invTxns = filteredTxns().filter(t => t.category === "Investment" && t.type === "DEBIT");
  const byDayMap = new Map();
  invTxns.forEach(t => {
    const k = dayKey(t.date);
    byDayMap.set(k, (byDayMap.get(k) || 0) + t.amount);
  });
  const pairs = [...byDayMap.entries()].sort((a, b) => a[0].localeCompare(b[0]));
  drawBars(els.investmentsChart, pairs);
}

// Call new render functions
const originalRender = render;
function render() {
  originalRender();
  renderTransfersChart();
  renderInvestmentsChart();
}

  const txns = filteredTxns();
  renderKpis(txns);

  const byDayMap = new Map();
  txns
    .filter((t) => t.type === "DEBIT")
    .forEach((t) => {
      const k = dayKey(t.date);
      byDayMap.set(k, (byDayMap.get(k) || 0) + t.amount);
    });
  const byDay = [...byDayMap.entries()].sort((a, b) => a[0].localeCompare(b[0]));
  drawBars(els.dailyChart, byDay);

  const cats = groupSum(
    txns.filter((t) => t.type === "DEBIT"),
    (t) => t.category
  );
  drawDonut(els.categoryChart, cats);

  const merchants = groupSum(
    txns.filter((t) => t.type === "DEBIT"),
    (t) => t.details.replace(/^Paid to\s+/i, "")
  );
  renderRanks(els.merchants, merchants);

  const instruments = groupSum(
    txns.filter((t) => t.type === "DEBIT"),
    (t) => t.instrument || "Unknown"
  );
  renderRanks(els.instruments, instruments);
  renderTable(txns);

  updateMiscBadge();
  updateRulesBadge();

  const { fileName, duration, account, isConsolidated } = statementMeta;
  els.status.hidden = false;
  const tag = isConsolidated ? " [Consolidated Statement]" : "";
  els.status.textContent = `${fileName}${tag} · ${txns.length} shown of ${allTxns.length} · ${duration || account || "All transactions ready"}`;
}

function renderRulesModalList() {
  const rules = getCustomRules();
  const baseCount = rulesFromJson.length;
  let html = "";
  if (baseCount > 0) {
    html += `<li class="rule-item" style="background:#f4f6f8;font-size:12px;color:var(--muted);border-radius:6px;padding:8px 12px;margin-bottom:8px;">
      <span>📁 <strong>${baseCount} base rules</strong> active from <code>rules.json</code></span>
    </li>`;
  }
  if (!rules.length) {
    html += `<li class="rule-item" style="color:var(--muted)">No custom overrides added yet. Add a merchant keyword above or click any category in the table!</li>`;
    els.rulesList.innerHTML = html;
    return;
  }
  html += rules
    .map(
      (r) => `<li class="rule-item">
        <span><strong>${r.keyword}</strong> &rarr; <span class="pill">${r.domain}</span></span>
        <button type="button" class="delete-rule-btn" data-kw="${r.keyword}" title="Delete rule">&times;</button>
      </li>`
    )
    .join("");
  els.rulesList.innerHTML = html;
}

function downloadRulesJson() {
  const customRules = getCustomRules();
  const mergedMap = new Map();

  for (const r of rulesFromJson) {
    if (r && r.keyword) {
      mergedMap.set(r.keyword.toLowerCase().trim(), { keyword: r.keyword.trim(), domain: r.domain });
    }
  }

  for (const r of customRules) {
    if (r && r.keyword) {
      mergedMap.set(r.keyword.toLowerCase().trim(), { keyword: r.keyword.trim(), domain: r.domain });
    }
  }

  const mergedRules = Array.from(mergedMap.values());
  mergedRules.sort((a, b) => {
    if (a.domain === b.domain) return a.keyword.localeCompare(b.keyword);
    return a.domain.localeCompare(b.domain);
  });

  const jsonStr = JSON.stringify(mergedRules, null, 2);
  const blob = new Blob([jsonStr], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "rules.json";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function openRulesModal() {
  if (!els.ruleDomain.children.length) {
    els.ruleDomain.innerHTML = DOMAINS.map((d) => `<option value="${d}">${d}</option>`).join("");
  }
  renderRulesModalList();
  els.rulesModal.showModal();
}

// -------------------------------------------------------------
// EXPORT UPDATED DATASET BACK TO EXCEL
// -------------------------------------------------------------
function exportUpdatedExcel() {
  if (!allTxns.length) {
    alert("No transactions to export. Load a file first!");
    return;
  }

  const exportRows = allTxns.map((t) => ({
    "Date": dayKey(t.date),
    "Amount": t.amount,
    "Purpose": t.details,
    "Mode": t.mode || t.instrument,
    "Domain": t.category,
    "Source": t.source || "Manual/Statement",
    "Status": t.status || (isTransferLike(t) ? "Transfer" : "Living Expense"),
    "Type": t.type,
    "RefId": t.id || "",
  }));

  const wbOut = XLSX.utils.book_new();
  const wsExp = XLSX.utils.json_to_sheet(exportRows);
  XLSX.utils.book_append_sheet(wbOut, wsExp, "Expenditure");

  // Generate Bal-Summary Sheet
  const livingTxns = allTxns.filter((t) => t.type === "DEBIT" && !isTransferLike(t));
  const domainTotals = groupSum(livingTxns, (t) => t.category);
  const modeTotals = groupSum(livingTxns, (t) => t.mode || t.instrument);

  const balRows = [
    { "Section": "Mode Breakdown", "Key": "", "Value": "" },
    ...modeTotals.map(([mode, sum]) => ({ "Section": "", "Key": mode, "Value": sum })),
    { "Section": "", "Key": "", "Value": "" },
    { "Section": "Domain Breakdown", "Key": "", "Value": "" },
    ...domainTotals.map(([domain, sum]) => ({ "Section": "", "Key": domain, "Value": sum })),
  ];
  const wsBal = XLSX.utils.json_to_sheet(balRows);
  XLSX.utils.book_append_sheet(wbOut, wsBal, "Bal-Summary");

  const baseName = (statementMeta.fileName || "Spend-Analyser-Export").replace(/\.[^/.]+$/, "");
  const outFileName = `${baseName}_Updated.xlsx`;

  XLSX.writeFile(wbOut, outFileName);
}

async function consolidateAndLoadFiles(files) {
  if (typeof XLSX === "undefined") {
    throw new Error("Excel/CSV library failed to load. Check your internet connection and retry.");
  }

  const allConsolidatedTxns = [];
  const utrSet = new Set();
  const modesFound = new Set();
  let txnCounter = 1;

  // Read all workbooks
  const parsedFiles = [];
  for (const f of files) {
    const buf = await f.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array", cellDates: true });
    parsedFiles.push({ name: f.name.toLowerCase(), wb });
  }

  // 1. Find and parse PhonePe
  const pp = parsedFiles.find((p) => p.name.includes("phonepe") || p.name.includes("statement"));
  if (pp) {
    const sheet = pp.wb.Sheets[pp.wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
    const headerIdx = findHeaderRow(rows);
    if (headerIdx >= 0) {
      const headers = rows[headerIdx].map(normalizeHeader);
      const col = (aliases) => headers.findIndex((h) => aliases.some((a) => h === a || h.includes(a)));
      const iDate = col(["date"]);
      const iTime = col(["time"]);
      const iDetails = col(["transaction details", "details", "description", "merchant"]);
      const iId = col(["transaction id"]);
      const iUtr = col(["utr"]);
      const iType = col(["transaction type", "type"]);
      const iInstrument = col(["credit/debit instrument", "instrument"]);
      const iAmount = col(["amount"]);
      const iDirection = col(["credit/debit", "dr/cr"]);

      for (let r = headerIdx + 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || row.every((c) => String(c).trim() === "")) continue;
        const details = String(row[iDetails] || "").trim();
        const amount = parseAmount(row[iAmount]);
        const date = parseDate(row[iDate]);
        if (!details || !amount || !date) continue;

        let direction = String(row[iType] || row[iDirection] || "").toUpperCase();
        direction = direction.includes("CREDIT") ? "CREDIT" : "DEBIT";

        const rawUtr = String(row[iUtr] || "").replace(/[^0-9]/g, "");
        if (rawUtr.length >= 6) utrSet.add(rawUtr);

        const rawInstr = String(row[iInstrument] || "");
        const mode = formatInstrument(rawInstr);
        modesFound.add(mode);

        const domain = categorize(details, direction);
        const status = domain === "Transfers" ? "Transfer" : (domain === "Investment" ? "Investment" : (direction === "CREDIT" ? "Income / Credit" : "Living Expense"));

        allConsolidatedTxns.push({
          _id: txnCounter++,
          date,
          time: String(row[iTime] || "").trim(),
          details: details.replace(/^Paid to\s+/i, "").replace(/^Payment to\s+/i, ""),
          id: String(row[iId] || "").trim(),
          type: direction,
          rawInstrument: rawInstr,
          instrument: mode,
          mode,
          category: domain,
          status,
          source: "PhonePe",
          amount: Math.abs(amount),
          _manualOverride: false,
        });
      }
    }
  }

  // 2. Find and parse ICICI Credit Card
  const icici = parsedFiles.find((p) => p.name.includes("icici") || p.name.includes("creditcard"));
  if (icici) {
    const sheet = icici.wb.Sheets[icici.wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.length < 4) continue;
      const d = parseDate(row[0]);
      const details = String(row[1] || "").trim();
      const amt = parseAmount(row[3]);
      const sign = String(row[4] || "").trim().toUpperCase();
      if (!d || !details || details.includes("Transaction Details")) continue;

      const type = sign === "CR" ? "CREDIT" : "DEBIT";
      let domain = "Misc";
      let status = "Living Expense";
      const lowD = details.toLowerCase();

      if (details.includes("BBPS Payment received")) {
        domain = "Transfers";
        status = "CC Bill Repayment";
      } else {
        domain = categorize(details, type);
        status = domain === "Transfers" ? "Transfer" : (domain === "Investment" ? "Investment" : (type === "CREDIT" ? "Refund / Credit" : "Living Expense"));
      }

      modesFound.add("cc-icici");
      allConsolidatedTxns.push({
        _id: txnCounter++,
        date: d,
        time: "",
        details,
        id: "",
        type,
        rawInstrument: "ICICI Credit Card",
        instrument: "cc-icici",
        mode: "cc-icici",
        category: domain,
        status,
        source: "ICICI-CreditCard",
        amount: Math.abs(amt),
        _manualOverride: false,
      });
    }
  }

  // 3. Find and parse HDFC Bank (Unique transactions only)
  const hdfc = parsedFiles.find((p) => p.name.includes("hdfc"));
  if (hdfc) {
    const sheet = hdfc.wb.Sheets[hdfc.wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.length < 3) continue;
      const d = parseDate(row[0]);
      const narr = String(row[1] || "").trim();
      if (!d || !narr || narr.includes("Narration") || narr.includes("***")) continue;

      const wAmt = parseAmount(row[3]);
      const dAmt = parseAmount(row[4]);
      const amt = wAmt > 0 ? wAmt : dAmt;
      if (!amt) continue;
      const type = wAmt > 0 ? "DEBIT" : "CREDIT";

      // Deduplicate: check if UTR is in PhonePe
      let isDuplicate = false;
      for (const u of utrSet) {
        if (narr.includes(u)) {
          isDuplicate = true;
          break;
        }
      }
      if (isDuplicate) continue;

      let mode = "hdfc-pp";
      let domain = "Misc";
      let status = "Living Expense";
      let purpose = narr;

      if (/^ATW-/i.test(narr)) {
        mode = "Cash";
        purpose = "ATM Cash Withdrawal";
        status = "ATM Withdrawal";
      } else if (/CRED CLUB/i.test(narr)) {
        mode = "hdfc-pp";
        domain = "Transfers";
        purpose = "Credit Card Bill (CRED)";
        status = "CC Bill Repayment";
      } else if (/NATIONAL HIGHWAYS AU/i.test(narr)) {
        domain = "Travel";
        purpose = "NHAI Annual Pass";
      } else if (/ABHISHEK/i.test(narr)) {
        domain = "Transfers";
        status = type === "CREDIT" ? "Income / Credit" : "Transfer";
      } else {
        domain = categorize(purpose, type);
      }

      modesFound.add(mode);
      allConsolidatedTxns.push({
        _id: txnCounter++,
        date: d,
        time: "",
        details: purpose,
        id: "",
        type,
        rawInstrument: mode,
        instrument: mode,
        mode,
        category: domain,
        status,
        source: "HDFC-Bank",
        amount: Math.abs(amt),
        _manualOverride: false,
      });
    }
  }

  if (!allConsolidatedTxns.length) {
    throw new Error("Could not find recognizable transactions across the uploaded statements.");
  }

  allTxns = allConsolidatedTxns.sort((a, b) => b.date - a.date || b.amount - a.amount);
  statementMeta = {
    account: "Multi-Account",
    duration: "Consolidated View",
    fileName: `Consolidated (${files.length} statements)`,
    isConsolidated: true,
  };

  // Populate mode filter dropdown
  if (els.modeFilter) {
    els.modeFilter.innerHTML = '<option value="all">All Payment Modes</option>' +
      [...modesFound].sort().map((m) => `<option value="${m}">${m}</option>`).join("");
    els.modeFilter.value = "all";
  }

  els.dashboard.hidden = false;
  render();
}

async function handleIncomingFiles(fileList) {
  if (!fileList || !fileList.length) return;
  try {
    if (fileList.length === 1) {
      setBusy(`Reading ${fileList[0].name}…`);
      await loadFile(fileList[0]);
    } else {
      setBusy(`Consolidating ${fileList.length} statements…`);
      await consolidateAndLoadFiles(Array.from(fileList));
    }
  } catch (err) {
    els.dashboard.hidden = true;
    setBusy(err.message || String(err));
  }
}

async function loadFile(file) {
  if (typeof XLSX === "undefined") {
    throw new Error("Excel/CSV library failed to load. Check your internet connection and retry.");
  }
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const { txns, meta } = parseWorkbook(wb, file.name);
  if (!txns.length) throw new Error("No transactions found in that file.");
  allTxns = txns.sort((a, b) => b.date - a.date || b.amount - a.amount);
  statementMeta = meta;
  els.dashboard.hidden = false;
  render();
}

function setBusy(msg) {
  els.status.hidden = false;
  els.status.textContent = msg;
}

// Drag and drop & file input
els.dropzone.addEventListener("click", () => els.fileInput.click());
els.dropzone.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") els.fileInput.click();
});
els.fileInput.addEventListener("change", () => handleIncomingFiles(els.fileInput.files));

["dragenter", "dragover"].forEach((ev) => {
  els.dropzone.addEventListener(ev, (e) => {
    e.preventDefault();
    els.dropzone.classList.add("dragover");
  });
});
["dragleave", "drop"].forEach((ev) => {
  els.dropzone.addEventListener(ev, (e) => {
    e.preventDefault();
    els.dropzone.classList.remove("dragover");
  });
});
els.dropzone.addEventListener("drop", (e) => {
  e.preventDefault();
  els.dropzone.classList.remove("dragover");
  handleIncomingFiles(e.dataTransfer.files);
});

// Interactive in-table recategorization
els.txnBody.addEventListener("change", (e) => {
  if (!e.target.classList.contains("category-select")) return;
  const id = Number(e.target.dataset.id);
  const newDomain = e.target.value;
  const txn = allTxns.find((t) => t._id === id);
  if (!txn) return;

  txn.category = newDomain;
  txn._manualOverride = true;

  // Clean merchant or recipient name
  const merchant = txn.details.replace(/^Paid to\s+/i, "").replace(/^Payment to\s+/i, "").trim();

  // If merchant has text, ask user if they want to remember this rule for all transactions with this merchant
  if (merchant) {
    const applyToAll = window.confirm(
      `Save rule to always categorize "${merchant}" as "${newDomain}"?\n\n(Saved locally in your browser)`
    );
    if (applyToAll) {
      addCustomRule(merchant, newDomain);
      return;
    }
  }

  render();
});

// Toolbar controls
els.excludeTransfers.addEventListener("change", render);
els.debitsOnly.addEventListener("change", render);
if (els.modeFilter) els.modeFilter.addEventListener("change", render);
els.search.addEventListener("input", () => renderTable(filteredTxns()));

els.miscFilter.addEventListener("click", () => {
  filterMiscOnly = !filterMiscOnly;
  els.miscFilter.classList.toggle("active", filterMiscOnly);
  render();
});

if (els.exportExcelBtn) {
  els.exportExcelBtn.addEventListener("click", exportUpdatedExcel);
}

els.rulesBtn.addEventListener("click", openRulesModal);
els.closeRulesBtn.addEventListener("click", () => els.rulesModal.close());
els.rulesModal.addEventListener("click", (e) => {
  if (e.target === els.rulesModal) els.rulesModal.close();
});

els.addRuleForm.addEventListener("submit", (e) => {
  e.preventDefault();
  const kw = els.ruleKeyword.value.trim();
  const cat = els.ruleDomain.value;
  if (kw && cat) {
    addCustomRule(kw, cat);
    els.ruleKeyword.value = "";
    renderRulesModalList();
  }
});

els.rulesList.addEventListener("click", (e) => {
  if (!e.target.classList.contains("delete-rule-btn")) return;
  const kw = e.target.dataset.kw;
  if (kw) {
    deleteCustomRule(kw);
    renderRulesModalList();
  }
});

els.clearBtn.addEventListener("click", () => {
  allTxns = [];
  els.fileInput.value = "";
  els.dashboard.hidden = true;
  els.status.hidden = true;
  filterMiscOnly = false;
  els.miscFilter.classList.remove("active");
  if (els.modeFilter) els.modeFilter.value = "all";
});

window.addEventListener("resize", () => {
  if (allTxns.length) render();
});

if (els.downloadRulesBtn) {
  els.downloadRulesBtn.addEventListener("click", downloadRulesJson);
}

// Initialize on page load
loadRulesJson();
updateRulesBadge();
