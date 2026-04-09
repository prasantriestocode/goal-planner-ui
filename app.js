const goals = [
  { id: "g1", name: "ABC's Education", years: 0, amount: 0, provision: 0, inflationType: "education", kind: "goal" },
  { id: "g2", name: "ABC's Marriage", years: 0, amount: 0, provision: 0, inflationType: "marriage", kind: "goal" },
  { id: "g3", name: "PQR's Education", years: 0, amount: 0, provision: 0, inflationType: "education", kind: "goal" },
  { id: "g4", name: "PQR's Marriage", years: 0, amount: 0, provision: 0, inflationType: "marriage", kind: "goal" },
];

const model = {
  name: "",
  planDate: "",
  dob: "",
  city: "",
  state: "",
  spouseDob: "",
  child1Dob: "",
  child2Dob: "",
  inflationRate: 0,
  educationInflationRate: 0,
  marriageInflationRate: 0,
  preRetRate: 0,
  postRetRate: 0,
  cashInGrowthRate: 0,
  retirementAge: 0,
  lifeExpectancy: 0,
  debtRate: 0,
  incomeMain: 0,
  incomeSpouse: 0,
  expHousehold: 0,
  expLifestyle: 0,
  expEducation: 0,
  expVehicle: 0,
  expMediclaim: 0,
  expUtilities: 0,
  expCarInsurance: 0,
  expMisc: 0,
  expLifeIns: 0,
  expVacation: 0,
  expRent: 0,
  expCreditCard: 0,
  expTravel: 0,
  expProfFees: 0,
  expPpfMonthly: 0,
  assetHome: 0,
  assetCar: 0,
  assetGold: 0,
  invLiquidMf: 0,
  invSavings: 0,
  invShares: 0,
  invEquityMf: 0,
  invDebtMf: 0,
  invBonds: 0,
  invPostal: 0,
  invPpf: 0,
  invUlip: 0,
  loanHome: 0,
  loanCar: 0,
  loanOther: 0,
  currentSipPm: 0,
  networthNotes: "",
  // ── New fields ──────────────────────────────
  invEpf: 0,
  invElss: 0,
  invBankFd: 0,
  invCash: 0,
  wizardCompleted: false,
  willStatus: "",
  willLastUpdated: "",
  nominationsUpdated: "",
  retirementMonthlyExp: 0,
  familyHistoryCriticalIllness: "",
  familyHistoryDescription: "",
};

const latestState = {
  goalSummary: null,
  networth: null,
  cashflow: null,
};

let additionalProperties = [];
let adminPortfolio = {
  asOfDate: "",
  equityRows: [],
  unifiRows: [],
  iciciRows: [],
};

let lifeInsuranceRows = [];
let healthInsuranceRows = [];
let carInsuranceRows = [];
let propertyInsuranceRows = [];
let customExpenses = [];
let cashflowOverrides = {}; // { [year]: { cashIn?, cashOut? } }
let children = [];
let customAssets = { physical: [], equity: [], debt: [], retirement: [], cash: [] };
let customLiabilities = [];
// Per-asset class growth rate overrides (% as number, e.g. 12 = 12%).
// null → falls back to model.preRetRate / model.debtRate as appropriate.
// NOTE: equity is intentionally locked at 12 (not null) so it stays
// independent of preRetRate — the Cash Flow growth rate is driven by the
// ROI table's weighted average, not by preRetRate directly.
let assetGrowthRates = {
  realEstate: 8.0,
  equity:     12,     // locked at 12% — independent of preRetRate
  debtSaving: null,   // null → model.debtRate
  debtMf:     null,
  bondsFd:    null,
  other:      null,
  ppf:        7.9,
  epf:        8.2,
  gold:       7.0,
};
let wizardCurrentStep = 0;

let auth = null;
let db = null;
let currentUser = null;
let currentRole = null;
let currentPlanId = null;
let autosaveTimer = null;
let isHydrating = false;

const defaultGoals = JSON.parse(JSON.stringify(goals));
const defaultModel = JSON.parse(JSON.stringify(model));

const inr = new Intl.NumberFormat("en-IN", { maximumFractionDigits: 0 });
const pct = (n) => `${(n * 100).toFixed(2)}%`;

function byId(id) {
  return document.getElementById(id);
}

function deepClone(obj) {
  return JSON.parse(JSON.stringify(obj));
}

function setStatus(msg) {
  const el = byId("authStatus");
  if (el) el.textContent = msg;
}

/* ── Toast notification system ───────────────────────────────── */
function showToast(msg, type = "success", duration = 3200) {
  const container = byId("toast-container");
  if (!container) return;
  const icons = { success: "✓", error: "✕", info: "ℹ", warning: "⚠" };
  const toast = document.createElement("div");
  toast.className = `toast-item toast-${type}`;
  toast.innerHTML = `
    <span class="toast-icon">${icons[type] || "ℹ"}</span>
    <span class="toast-msg">${msg}</span>
    <button class="toast-close" aria-label="Dismiss">✕</button>`;
  container.appendChild(toast);
  const dismiss = () => {
    toast.classList.add("toast-out");
    toast.addEventListener("animationend", () => toast.remove(), { once: true });
  };
  toast.querySelector(".toast-close").addEventListener("click", dismiss);
  setTimeout(dismiss, duration);
}

/* ── Skeleton loader helpers ─────────────────────────────────── */
function showSkeleton() {
  const el = byId("skeleton-overlay");
  const hero = byId("db-hero");
  if (el)   el.classList.add("visible");
  if (hero) hero.style.display = "none";
}
function hideSkeleton() {
  const el = byId("skeleton-overlay");
  const hero = byId("db-hero");
  if (el)   el.classList.remove("visible");
  if (hero) hero.style.display = "";
}

function isAdmin() {
  return currentRole === "admin";
}

function getAdminEmail() {
  return (window.firebaseConfig?.adminEmail || "").trim().toLowerCase();
}

function isAdminEmail(email) {
  return (email || "").trim().toLowerCase() === getAdminEmail();
}

function yearsBetween(start, end) {
  if (!start || !end) return 0;
  const s = new Date(start);
  const e = new Date(end);
  if (Number.isNaN(s.valueOf()) || Number.isNaN(e.valueOf())) return 0;
  return Math.max(0, Math.floor((e - s) / (365.25 * 24 * 60 * 60 * 1000)));
}

function pmt(rate, periods, pv, fv = 0) {
  if (!periods || periods <= 0) return 0;
  if (rate === 0) return -(pv + fv) / periods;
  const factor = (1 + rate) ** periods;
  return (-(fv + pv * factor) * rate) / (factor - 1);
}

function requiredMonthlyFromGap(gap, annualRate, months) {
  if (gap <= 0 || months <= 0) return 0;
  const r = annualRate / 12;
  if (r === 0) return gap / months;
  // Monthly installment to accumulate `gap` as future value with type=1 (beginning-of-month).
  return (gap * r) / (((1 + r) ** months - 1) * (1 + r));
}

function formatRs(n) {
  return inr.format(Math.round(n || 0));
}

function escHtml(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function npv(rate, values) {
  let out = 0;
  for (let i = 0; i < values.length; i += 1) {
    out += values[i] / (1 + rate) ** (i + 1);
  }
  return out;
}

function rateForGoal(goal, data) {
  if (goal.inflationType === "education") return data.educationInflationRate / 100;
  if (goal.inflationType === "marriage") return data.marriageInflationRate / 100;
  return data.inflationRate / 100;
}

function resetToDefaults() {
  Object.keys(model).forEach((k) => {
    model[k] = deepClone(defaultModel[k]);
  });
  goals.splice(0, goals.length, ...deepClone(defaultGoals));
  additionalProperties = [];
  adminPortfolio = { asOfDate: "", equityRows: [], unifiRows: [], iciciRows: [] };
  lifeInsuranceRows = [];
  healthInsuranceRows = [];
  carInsuranceRows = [];
  propertyInsuranceRows = [];
  customExpenses = [];
  cashflowOverrides = {};
  children = [];
  customAssets = { physical: [], equity: [], debt: [], retirement: [], cash: [] };
  customLiabilities = [];
}

function applyPlanData(planData = {}) {
  isHydrating = true;
  const incomingModel = planData.model || {};
  Object.keys(model).forEach((k) => {
    if (Object.prototype.hasOwnProperty.call(incomingModel, k)) model[k] = incomingModel[k];
  });
  const incomingGoals = Array.isArray(planData.goals) ? planData.goals : deepClone(defaultGoals);
  goals.splice(0, goals.length, ...incomingGoals);
  additionalProperties = Array.isArray(planData.additionalProperties) ? planData.additionalProperties : [];
  adminPortfolio = planData.adminPortfolio || { asOfDate: "", equityRows: [], unifiRows: [], iciciRows: [] };
  model.networthNotes = planData.networthNotes || model.networthNotes || "";
  lifeInsuranceRows = Array.isArray(planData.lifeInsuranceRows) ? planData.lifeInsuranceRows : [];
  healthInsuranceRows = Array.isArray(planData.healthInsuranceRows) ? planData.healthInsuranceRows : [];
  carInsuranceRows = Array.isArray(planData.carInsuranceRows) ? planData.carInsuranceRows : [];
  propertyInsuranceRows = Array.isArray(planData.propertyInsuranceRows) ? planData.propertyInsuranceRows : [];
  customExpenses = Array.isArray(planData.customExpenses) ? planData.customExpenses : [];
  if (planData.customAssets && typeof planData.customAssets === "object") {
    customAssets = {
      physical: Array.isArray(planData.customAssets.physical) ? planData.customAssets.physical : [],
      equity: Array.isArray(planData.customAssets.equity) ? planData.customAssets.equity : [],
      debt: Array.isArray(planData.customAssets.debt) ? planData.customAssets.debt : [],
      retirement: Array.isArray(planData.customAssets.retirement) ? planData.customAssets.retirement : [],
      cash: Array.isArray(planData.customAssets.cash) ? planData.customAssets.cash : [],
    };
  } else {
    customAssets = { physical: [], equity: [], debt: [], retirement: [], cash: [] };
  }
  customLiabilities = Array.isArray(planData.customLiabilities) ? planData.customLiabilities : [];
  cashflowOverrides = (planData.cashflowOverrides && typeof planData.cashflowOverrides === "object") ? planData.cashflowOverrides : {};
  children = Array.isArray(planData.children) ? planData.children : [];
  if (planData.assetGrowthRates && typeof planData.assetGrowthRates === "object") {
    Object.assign(assetGrowthRates, planData.assetGrowthRates);
  }

  bindAllInputValues();
  renderGoalInputRows();
  renderPropertyRows();
  renderAdminNetworthSheet();
  recalc();
  isHydrating = false;
}

async function initFirebase() {
  if (!window.firebase || !window.firebaseConfig || window.firebaseConfig.apiKey === "REPLACE_ME") {
    setStatus("Firebase not configured. Fill firebase-config.js");
    return;
  }
  if (!firebase.apps.length) firebase.initializeApp(window.firebaseConfig);
  auth = firebase.auth();
  db = firebase.firestore();

  auth.onAuthStateChanged(async (user) => {
    currentUser = user;
    if (!user) {
      currentRole = null;
      currentPlanId = null;
      byId("adminPanel").hidden = true;
      setStatus("Not logged in.");
      resetToDefaults();
      bindAllInputValues();
      renderGoalInputRows();
      renderPropertyRows();
      renderAdminNetworthSheet();
      recalc();
      setAppLocked(true);
      return;
    }
    try {
      const userDoc = await db.collection("users").doc(user.uid).get();
      const fallbackRole = isAdminEmail(user.email) ? "admin" : "investor";
      const userData = userDoc.exists ? userDoc.data() : { role: fallbackRole, investorName: model.name };
      currentRole = userData.role || fallbackRole;
      byId("adminPanel").hidden = !isAdmin();
      // Inject admin option dynamically when admin logs in; investors never see it.
      const roleSelect = byId("authRole");
      if (isAdmin() && !roleSelect.querySelector('option[value="admin"]')) {
        const opt = document.createElement("option");
        opt.value = "admin";
        opt.textContent = "Admin";
        roleSelect.appendChild(opt);
      }
      roleSelect.value = currentRole;
      setStatus(`Logged in as ${user.email} (${currentRole})`);
      applyRoleVisibility();
      setAppLocked(false);

      if (isAdmin()) {
        try {
          await loadInvestorList();
        } catch (e) {
          setStatus(`Admin read blocked by Firestore rules: ${e.message}`);
          currentPlanId = user.uid;
          await loadPlan(currentPlanId);
        }
      } else {
        currentPlanId = user.uid;
        await loadPlan(currentPlanId);
      }
    } catch (e) {
      setStatus(`Login data load failed: ${e.message}`);
    }
  });
}

async function signup() {
  if (!auth || !db) return;
  const email = byId("authEmail").value.trim();
  const password = byId("authPassword").value;
  if (!email || !password) return alert("Enter email and password.");

  // Only admin email is allowed to self-signup
  // Investors must be created by admin via the "Add Investor" form
  if (!isAdminEmail(email)) {
    throw new Error("Investor accounts can only be created by the admin. Please contact support to request an account.");
  }

  const role = "admin";
  const cred = await auth.createUserWithEmailAndPassword(email, password);
  await db.collection("users").doc(cred.user.uid).set({
    email,
    role,
    investorName: model.name || "",
    createdAt: firebase.firestore.FieldValue.serverTimestamp(),
  });
  await db.collection("investorPlans").doc(cred.user.uid).set({
    investorName: model.name || "",
    model,
    goals,
    additionalProperties,
    networthNotes: model.networthNotes || "",
    adminPortfolio,
    customAssets,
    customLiabilities,
    updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
  });
}

async function login() {
  if (!auth) {
    setStatus("Firebase is initializing... Please wait a moment and try again.");
    throw new Error("Firebase not ready yet");
  }
  const email = byId("authEmail").value.trim();
  const password = byId("authPassword").value;
  if (!email || !password) return alert("Enter email and password.");
  await auth.signInWithEmailAndPassword(email, password);
  // Role is determined from Firestore in onAuthStateChanged — no blocking check needed here.
}

async function logout() {
  if (auth) await auth.signOut();
}

async function createInvestor(email, investorName, password) {
  if (!auth || !db) {
    throw new Error("Firebase not initialized");
  }
  if (!email || !investorName || !password) {
    throw new Error("Email, investor name, and password are required");
  }

  try {
    // 1. Create Firebase Auth user
    const cred = await auth.createUserWithEmailAndPassword(email, password);
    const uid = cred.user.uid;

    // 2. Create users collection document with investor role
    await db.collection("users").doc(uid).set({
      email,
      role: "investor",
      investorName,
      createdAt: firebase.firestore.FieldValue.serverTimestamp(),
    });

    // 3. Create empty investorPlans document
    await db.collection("investorPlans").doc(uid).set({
      investorName,
      model: JSON.parse(JSON.stringify(model)), // deep copy
      goals: JSON.parse(JSON.stringify(goals)), // deep copy
      additionalProperties: [],
      networthNotes: "",
      adminPortfolio: {
        asOfDate: "",
        equityRows: [],
        unifiRows: [],
        iciciRows: [],
      },
      customAssets: { physical: [], equity: [], debt: [], retirement: [], cash: [] },
      customLiabilities: [],
      customExpenses: [],
      lifeInsuranceRows: [],
      healthInsuranceRows: [],
      carInsuranceRows: [],
      propertyInsuranceRows: [],
      cashflowOverrides: {},
      children: [],
      assetGrowthRates: {},
      updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
    });

    // Return success with created credentials for admin to share
    return {
      uid,
      email,
      investorName,
      password,
      message: `Investor created successfully. Share these credentials:\nEmail: ${email}\nPassword: ${password}`
    };
  } catch (error) {
    // If user creation succeeded but subsequent Firestore writes failed,
    // we should clean up. For now, rethrow the error.
    throw new Error(`Failed to create investor: ${error.message}`);
  }
}

async function loadPlan(planId) {
  if (!db || !planId) return;
  currentPlanId = planId;
  const doc = await db.collection("investorPlans").doc(planId).get();
  if (!doc.exists) {
    resetToDefaults();
    bindAllInputValues();
    recalc();
    return;
  }
  applyPlanData(doc.data());
  if (isAdmin()) byId("adminInvestorName").value = doc.data().investorName || "";
  // Show wizard on first login for investors
  if (!isAdmin() && !model.wizardCompleted) {
    setTimeout(openWizard, 500);
  }
  // Lock My Page if wizard has been completed
  if (model.wizardCompleted) {
    setTimeout(lockMyPage, 100);
  }
}

async function saveCurrentPlan() {
  if (!db || !currentUser || !currentPlanId) return;
  const investorName = isAdmin() ? byId("adminInvestorName").value.trim() || model.name : model.name;
  const payload = {
    investorName,
    model,
    goals,
    additionalProperties,
    networthNotes: model.networthNotes || "",
    adminPortfolio,
    lifeInsuranceRows,
    healthInsuranceRows,
    carInsuranceRows,
    propertyInsuranceRows,
    customExpenses,
    cashflowOverrides,
    children,
    customAssets,
    customLiabilities,
    assetGrowthRates,
    updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
  };
  await db.collection("investorPlans").doc(currentPlanId).set(payload, { merge: true });
  setStatus(`Saved at ${new Date().toLocaleTimeString()}`);
  showToast(`Plan saved at ${new Date().toLocaleTimeString()}`, "success");
}

function scheduleAutosave() {
  if (!currentUser || !currentPlanId || isHydrating) return;
  clearTimeout(autosaveTimer);
  autosaveTimer = setTimeout(() => {
    saveCurrentPlan().catch((e) => {
      setStatus(`Save failed: ${e.message}`);
      showToast(`Save failed: ${e.message}`, "error");
    });
  }, 800);
}

async function loadInvestorList() {
  if (!db) return;
  const snap = await db.collection("investorPlans").orderBy("updatedAt", "desc").get();
  const sel = byId("investorSelect");
  sel.innerHTML = "";
  snap.forEach((doc) => {
    const opt = document.createElement("option");
    const d = doc.data();
    opt.value = doc.id;
    opt.textContent = d.investorName ? `${d.investorName} (${doc.id.slice(0, 6)})` : doc.id;
    sel.appendChild(opt);
  });
  if (sel.options.length) {
    currentPlanId = sel.value;
    await loadPlan(sel.value);
  }
}

function applyRoleVisibility() {
  document.querySelectorAll(".admin-only").forEach((el) => {
    el.hidden = !isAdmin();
  });
  const asOf = byId("adminAsOfDate");
  if (asOf) asOf.disabled = !isAdmin();
}

function setAppLocked(locked) {
  document.body.classList.toggle("app-locked", locked);
  const tabs = byId("sheetTabs");
  if (tabs) tabs.hidden = locked;
  document.querySelectorAll(".sheet").forEach((s) => {
    s.hidden = locked;
  });
  const lockTargets = [
    "downloadExcelBtn", "downloadPdfBtn", "savePlanBtn", "logoutBtn",
    "mobExcelBtn", "mobPdfBtn", "mobWizardBtn"
  ];
  lockTargets.forEach((id) => {
    const el = byId(id);
    if (el) el.disabled = locked;
  });
}

function bindInput(id) {
  const el = byId(id);
  if (!el) return;
  el.value = model[id] ?? "";
  el.addEventListener("input", () => {
    model[id] = el.type === "number" ? Number(el.value || 0) : el.value;
    recalc();
  });
}

function bindAllInputValues() {
  const ids = [
    "name",
    "planDate",
    "dob",
    "city",
    "state",
    "spouseDob",
    "child1Dob",
    "child2Dob",
    "inflationRate",
    "educationInflationRate",
    "marriageInflationRate",
    "preRetRate",
    "postRetRate",
    "cashInGrowthRate",
    "retirementAge",
    "lifeExpectancy",
    "debtRate",
    "incomeMain",
    "incomeSpouse",
    "expHousehold",
    "expLifestyle",
    "expEducation",
    "expVehicle",
    "expMediclaim",
    "expUtilities",
    "expCarInsurance",
    "expMisc",
    "assetHome",
    "assetCar",
    "assetGold",
    "invLiquidMf",
    "invSavings",
    "invShares",
    "invEquityMf",
    "invDebtMf",
    "invBonds",
    "invPostal",
    "invPpf",
    "invEpf",
    "invElss",
    "invBankFd",
    "invCash",
    "invUlip",
    "loanHome",
    "loanCar",
    "loanOther",
    "currentSipPm",
  ];
  ids.forEach((id) => {
    const el = byId(id);
    if (!el) return;
    el.value = model[id] ?? "";
  });
  const notes = byId("networthNotes");
  if (notes) notes.value = model.networthNotes || "";
  // Will fields
  const ws = byId("db-willStatus");    if (ws) ws.value = model.willStatus || "";
  const wlu = byId("db-willLastUpdated"); if (wlu) wlu.value = model.willLastUpdated || "";
  const nu = byId("db-nominationsUpdated"); if (nu) nu.value = model.nominationsUpdated || "";
}

function renderPropertyRows() {
  const body = byId("propertyBody");
  if (!body) return;
  body.innerHTML = "";
  additionalProperties.forEach((p, idx) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><input data-prop-idx="${idx}" data-prop-key="name" value="${escHtml(p.name || "")}"></td>
      <td><input type="number" data-prop-idx="${idx}" data-prop-key="value" value="${p.value || 0}"></td>
      <td><input type="number" data-prop-idx="${idx}" data-prop-key="ownership" value="${p.ownership ?? 100}"></td>
      <td><input data-prop-idx="${idx}" data-prop-key="loanLinked" value="${escHtml(p.loanLinked || "")}"></td>
      <td><button type="button" data-del-prop="${idx}">Delete</button></td>
    `;
    body.appendChild(tr);
  });

  body.querySelectorAll("input[data-prop-idx]").forEach((el) => {
    el.addEventListener("change", () => {
      const idx = Number(el.dataset.propIdx);
      const key = el.dataset.propKey;
      const raw = el.value;
      additionalProperties[idx][key] = key === "value" || key === "ownership" ? Number(raw || 0) : raw;
      recalc();
    });
  });
  body.querySelectorAll("button[data-del-prop]").forEach((btn) => {
    btn.addEventListener("click", () => {
      additionalProperties.splice(Number(btn.dataset.delProp), 1);
      renderPropertyRows();
      recalc();
    });
  });
}

function bindStaticUiEvents() {
  byId("addPropertyBtn")?.addEventListener("click", () => {
    additionalProperties.push({ name: "", value: 0, ownership: 100, loanLinked: "" });
    renderPropertyRows();
    recalc();
  });
  byId("networthNotes")?.addEventListener("input", (e) => {
    model.networthNotes = e.target.value;
    scheduleAutosave();
  });
  byId("savePlanBtn")?.addEventListener("click", () => saveCurrentPlan().catch((e) => { setStatus(e.message); showToast(e.message, "error"); }));
  byId("logoutBtn")?.addEventListener("click", () => logout().catch((e) => { setStatus(e.message); showToast(e.message, "error"); }));
  byId("loginBtn")?.addEventListener("click", async () => {
    try {
      const btn = byId("loginBtn");
      const origText = btn.textContent;
      btn.disabled = true;
      btn.textContent = "Signing in...";
      await login();
      btn.textContent = origText;
    } catch (e) {
      setStatus(e.message || "Login failed. Please try again.");
      const btn = byId("loginBtn");
      btn.disabled = false;
      btn.textContent = "Sign In";
    }
  });
  byId("signupBtn")?.addEventListener("click", () => signup().catch((e) => setStatus(e.message)));
  byId("createInvestorBtn")?.addEventListener("click", async () => {
    if (!isAdmin()) return showToast("Only admins can create investors", "error");
    const email = byId("newInvestorEmail")?.value.trim();
    const investorName = byId("newInvestorName")?.value.trim();
    const password = byId("newInvestorPassword")?.value;
    const msgEl = byId("createInvestorMsg");

    if (!email || !investorName || !password) {
      if (msgEl) {
        msgEl.style.color = "#dc2626";
        msgEl.textContent = "All fields are required";
        msgEl.style.display = "block";
      }
      return;
    }

    try {
      const btn = byId("createInvestorBtn");
      if (btn) {
        btn.disabled = true;
        btn.textContent = "Creating...";
      }

      const result = await createInvestor(email, investorName, password);

      // Clear form and show success message
      if (byId("newInvestorEmail")) byId("newInvestorEmail").value = "";
      if (byId("newInvestorName")) byId("newInvestorName").value = "";
      if (byId("newInvestorPassword")) byId("newInvestorPassword").value = "";

      if (msgEl) {
        msgEl.style.color = "#059669";
        msgEl.innerHTML = `✓ ${result.message}<br><code style="background:#f0f4f8;padding:0.25rem 0.5rem;border-radius:4px;font-size:0.8rem;">Password: ${result.password}</code>`;
        msgEl.style.display = "block";
      }

      // Refresh investor list
      await loadInvestorList();
      showToast("Investor created successfully", "success", 3000);

      // Clear message after 5 seconds
      setTimeout(() => {
        if (msgEl) msgEl.style.display = "none";
      }, 5000);
    } catch (error) {
      if (msgEl) {
        msgEl.style.color = "#dc2626";
        msgEl.textContent = error.message || "Failed to create investor";
        msgEl.style.display = "block";
      }
      showToast(error.message || "Failed to create investor", "error");
    } finally {
      const btn = byId("createInvestorBtn");
      if (btn) {
        btn.disabled = false;
        btn.textContent = "Create Investor";
      }
    }
  });
  byId("investorSelect")?.addEventListener("change", async (e) => {
    showSkeleton();
    try {
      await loadPlan(e.target.value);
      showToast("Investor data loaded", "info", 2000);
    } catch (err) {
      showToast(`Load failed: ${err.message}`, "error");
    } finally {
      hideSkeleton();
    }
  });
  byId("adminInvestorName")?.addEventListener("change", (e) => {
    model.name = e.target.value;
    const nameEl = byId("name");
    if (nameEl) nameEl.value = model.name;
    scheduleAutosave();
  });
  byId("adminAsOfDate")?.addEventListener("change", (e) => {
    if (!isAdmin()) return;
    adminPortfolio.asOfDate = e.target.value;
    scheduleAutosave();
  });
  byId("addEquityRowBtn")?.addEventListener("click", () => {
    if (!isAdmin()) return;
    adminPortfolio.equityRows.push({ name: "", currentValue: 0 });
    renderAdminNetworthSheet();
    scheduleAutosave();
  });
  byId("addUnifiRowBtn")?.addEventListener("click", () => {
    if (!isAdmin()) return;
    adminPortfolio.unifiRows.push({ name: "", currentValue: 0 });
    renderAdminNetworthSheet();
    scheduleAutosave();
  });
  byId("addIciciRowBtn")?.addEventListener("click", () => {
    if (!isAdmin()) return;
    adminPortfolio.iciciRows.push({ name: "", currentValue: 0 });
    renderAdminNetworthSheet();
    scheduleAutosave();
  });

  // Insurance add buttons on dashboard
  byId("addLifeInsBtn")?.addEventListener("click", () => {
    lifeInsuranceRows.push({ policyName: "", company: "", sumAssured: 0, annualPrem: 0, surrenderVal: 0 });
    renderInsuranceTables();
    scheduleAutosave();
  });
  byId("addHealthInsBtn")?.addEventListener("click", () => {
    healthInsuranceRows.push({ policyName: "", company: "", sumAssured: 0, annualPrem: 0, members: "" });
    renderInsuranceTables();
    scheduleAutosave();
  });
  byId("addCarInsBtn")?.addEventListener("click", () => {
    carInsuranceRows.push({ policyName: "", company: "", idv: 0, annualPrem: 0, expiry: "" });
    renderInsuranceTables();
    scheduleAutosave();
  });
  byId("addPropertyInsBtn")?.addEventListener("click", () => {
    propertyInsuranceRows.push({ policyName: "", company: "", propertyName: "", cover: 0, annualPrem: 0 });
    renderInsuranceTables();
    scheduleAutosave();
  });

  // Cashflow table — pressing Enter on a growth/cashIn/cashOut cell commits the value
  // (number inputs only fire "change" on blur, not on Enter, so we trigger it manually)
  byId("cashflowBody")?.addEventListener("keydown", (e) => {
    if (e.key !== "Enter") return;
    const inp = e.target;
    if (inp.classList.contains("cf-edit")) {
      e.preventDefault();
      inp.dispatchEvent(new Event("change", { bubbles: true }));
    }
  });

  // Cashflow table — editable growth rate, cash in, cash out
  byId("cashflowBody")?.addEventListener("change", (e) => {
    const inp = e.target;
    if (!inp.classList.contains("cf-edit")) return;
    const year = Number(inp.dataset.cfYear);
    const key  = inp.dataset.cfKey;
    const val  = Number(inp.value || 0);
    if (key === "growth") {
      const isPostRet = inp.dataset.cfPostRet === "1";
      if (isPostRet) {
        // Post-ret: change global post-ret rate and sync My Page input
        model.postRetRate = val;
        const syncEl = byId("postRetRate");
        if (syncEl) syncEl.value = val;
      } else {
        // Pre-ret: store as a per-year override (blended ROI stays as the base default)
        if (!cashflowOverrides[year]) cashflowOverrides[year] = {};
        cashflowOverrides[year].growth = val;
      }
    } else {
      if (!cashflowOverrides[year]) cashflowOverrides[year] = {};
      cashflowOverrides[year][key] = val;
    }
    scheduleAutosave();
    recalc();
  });

  // Will & nominations fields on dashboard
  byId("db-willStatus")?.addEventListener("change", (e) => {
    model.willStatus = e.target.value;
    scheduleAutosave();
  });
  byId("db-willLastUpdated")?.addEventListener("change", (e) => {
    model.willLastUpdated = e.target.value;
    scheduleAutosave();
  });
  byId("db-nominationsUpdated")?.addEventListener("change", (e) => {
    model.nominationsUpdated = e.target.value;
    scheduleAutosave();
  });
}

function renderGoalInputRows() {
  const body = byId("goalBody");
  if (!body) return;
  body.innerHTML = "";
  goals.forEach((g) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${escHtml(g.name)}</td>
      <td><input type="number" min="0" value="${g.years}" data-key="${g.id}:years" /></td>
      <td><input type="number" min="0" value="${g.amount}" data-key="${g.id}:amount" /></td>
      <td><input type="number" min="0" value="${g.provision}" data-key="${g.id}:provision" /></td>
    `;
    body.appendChild(row);
  });

  body.querySelectorAll("input").forEach((input) => {
    input.addEventListener("change", () => {
      const [id, key] = input.dataset.key.split(":");
      const goal = goals.find((g) => g.id === id);
      goal[key] = Number(input.value || 0);
      recalc();
    });
  });
}

function computeGoalOutput() {
  const startYear = new Date(model.planDate).getFullYear();
  return goals.map((g) => {
    const inflation = rateForGoal(g, model);
    const projectedValue = g.amount * (1 + inflation) ** g.years;
    const gap = Math.max(0, projectedValue - g.provision);
    const monthlyRate = model.preRetRate / 100 / 12;
    const sip = Math.max(0, pmt(monthlyRate, g.years * 12, -g.provision, gap));
    return {
      ...g,
      inflation,
      projectedValue,
      targetYear: startYear + g.years,
      gap,
      sip,
    };
  });
}

function computeCashflow(goalOutput, requiredSip, monthlyInflow, monthlyOutflow) {
  const startYear = new Date(model.planDate).getFullYear();
  const currentAge = yearsBetween(model.dob, model.planDate);
  const years = Math.max(0, model.lifeExpectancy - currentAge);
  const nonRetirementGoals = goalOutput.filter((g) => g.kind !== "retirement");
  const retireAfterYears = Math.max(0, model.retirementAge - currentAge);
  const retirementMap = new Map();
  const retirementYears = Math.max(0, model.lifeExpectancy - model.retirementAge);
  // Retirement cost excludes children's education expense.
  // If investor explicitly set retirementMonthlyExp, use that; otherwise derive from expenses.
  const customExpTotal = customExpenses.reduce((s, e) => s + (Number(e.amount)||0), 0);
  const computedRetirementBase =
    model.expHousehold +
    model.expLifestyle +
    model.expVehicle +
    model.expMediclaim +
    model.expUtilities +
    model.expCarInsurance +
    model.expMisc +
    model.expLifeIns +
    model.expVacation +
    model.expRent +
    model.expCreditCard +
    model.expTravel +
    model.expProfFees +
    model.expPpfMonthly +
    customExpTotal;
  const retirementBaseOutflow = (model.retirementMonthlyExp > 0)
    ? model.retirementMonthlyExp
    : computedRetirementBase;
  const t6 = retirementBaseOutflow * 12 * (1 + model.inflationRate / 100) ** retireAfterYears;
  const retirementSeries = [];
  for (let i = 0; i < retirementYears; i += 1) {
    // Excel-style T6:T31 progression: first retirement year = T6, then each year grows by inflation.
    retirementSeries.push(t6 * (1 + model.inflationRate / 100) ** i);
  }
  retirementSeries.forEach((amt, idx) => {
    const y = startYear + retireAfterYears + 1 + idx;
    retirementMap.set(y, amt);
  });

  let opening =
    model.invLiquidMf +
    model.invSavings +
    model.invShares +
    model.invEquityMf +
    model.invDebtMf +
    (model.invElss || 0) +
    model.invBonds +
    model.invPostal +
    model.invPpf +
    (model.invEpf || 0) +
    model.invUlip +
    (model.invBankFd || 0) +
    (model.invCash || 0);
  const annualSurplus = Math.max(0, (monthlyInflow - monthlyOutflow) * 12);
  // Keep Goal-Sheet linkage but ensure Input inflow/outflow changes are reflected immediately.
  let cashIn = (model.currentSipPm || 0) * 12;
  const rows = [];

  for (let i = 0; i <= years; i += 1) {
    const year = startYear + i;
    const age = currentAge + i;
    const ov = cashflowOverrides[year] || {};
    // Pre-ret growth default = portfolio weighted average (from ROI table).
    // Post-ret growth default = postRetRate.
    // Either can be overridden per-year via cashflowOverrides[year].growth.
    const blendedPreRet = computeBlendedRoi();
    const baseGrowth = age < model.retirementAge ? blendedPreRet : model.postRetRate / 100;
    const growth = (ov.growth !== undefined) ? (ov.growth / 100) : baseGrowth;
    const baseCashIn = age <= model.retirementAge ? cashIn : 0;
    const effectiveCashIn = (ov.cashIn !== undefined) ? ov.cashIn : baseCashIn;
    const lumpSum = ov.lumpSum || 0;
    const fvEnd = opening * (1 + growth) + effectiveCashIn + lumpSum;
    const goalHit = nonRetirementGoals.filter((g) => g.targetYear === year);
    const retireOut = retirementMap.get(year) || 0;
    const computedCashOut = goalHit.reduce((sum, g) => sum + g.projectedValue, 0) + retireOut;
    const cashOut = (ov.cashOut !== undefined) ? ov.cashOut : computedCashOut;
    const goalText = [...goalHit.map((g) => g.name), ...(retireOut > 0 ? ["Retirement"] : [])].join(" & ");
    const clBal = fvEnd - cashOut;

    rows.push({
      no: i,
      year,
      age,
      opBal: opening,
      cashIn: effectiveCashIn,
      lumpSum,
      growth,
      fvEnd,
      cashOut,
      clBal,
      goals: goalText,
    });

    opening = clBal;
    cashIn *= 1.10; // 10% SIP step-up year on year
  }

  return rows;
}

function renderGoalSheet(goalOutput) {
  const customExpTotal = customExpenses.reduce((s, e) => s + (Number(e.amount)||0), 0);
  const monthlyOutflow =
    model.expHousehold +
    model.expLifestyle +
    model.expEducation +
    model.expVehicle +
    model.expMediclaim +
    model.expUtilities +
    model.expCarInsurance +
    model.expMisc +
    model.expLifeIns +
    model.expVacation +
    model.expRent +
    model.expCreditCard +
    model.expTravel +
    model.expProfFees +
    model.expPpfMonthly +
    customExpTotal;
  // Retirement current cost: use explicit retirementMonthlyExp if set, otherwise
  // derive from expenses (excluding education — per Excel pattern).
  const computedRetirementBaseGS =
    model.expHousehold +
    model.expLifestyle +
    model.expVehicle +
    model.expMediclaim +
    model.expUtilities +
    model.expCarInsurance +
    model.expMisc;
  const retirementBaseOutflow = (model.retirementMonthlyExp > 0)
    ? model.retirementMonthlyExp
    : computedRetirementBaseGS;
  const currentAge = yearsBetween(model.dob, model.planDate);
  const retireAfterYears = Math.max(0, model.retirementAge - currentAge);
  const retirementYears = Math.max(0, model.lifeExpectancy - model.retirementAge);
  const t6 = retirementBaseOutflow * 12 * (1 + model.inflationRate / 100) ** retireAfterYears;
  const retirementSeries = [];
  for (let i = 0; i < retirementYears; i += 1) {
    retirementSeries.push(t6 * (1 + model.inflationRate / 100) ** i);
  }
  const retirementCorpus = npv(model.inflationRate / 100, retirementSeries);

  const targetBody = byId("goalTargetBody");
  const strategyBody = byId("goalStrategyBody");
  targetBody.innerHTML = "";
  strategyBody.innerHTML = "";

  const enriched = goalOutput.map((g) => {
    const corpus = g.projectedValue;
    const gap = Math.max(0, corpus - g.provision);
    const pm = requiredMonthlyFromGap(gap, model.preRetRate / 100, g.years * 12);
    return { ...g, corpus, gap, pm, py: pm * 12 };
  });

  // Retirement in Goal-Sheet is automated from outflow sum fields (Excel F102 pattern).
  const retirementCurrCost = retirementBaseOutflow;
  const retirementYearsLeft = Math.max(0, retireAfterYears);
  const retirementTargetYear = new Date(model.planDate).getFullYear() + retirementYearsLeft;
  const retirementProvision =
    model.invLiquidMf +
    model.invSavings +
    model.invShares +
    model.invEquityMf +
    model.invDebtMf +
    (model.invElss || 0) +
    model.invBonds +
    model.invPostal +
    model.invPpf +
    (model.invEpf || 0) +
    model.invUlip +
    (model.invBankFd || 0) +
    (model.invCash || 0);
  const retirementGap = Math.max(0, retirementCorpus - retirementProvision);
  const retirementPm = requiredMonthlyFromGap(retirementGap, model.preRetRate / 100, retirementYearsLeft * 12);
  enriched.push({
    id: "retirement",
    name: "Retirement",
    targetYear: retirementTargetYear,
    years: retirementYearsLeft,
    amount: retirementCurrCost,
    inflation: model.inflationRate / 100,
    projectedValue: retirementCurrCost * (1 + model.inflationRate / 100) ** retirementYearsLeft,
    corpus: retirementCorpus,
    provision: retirementProvision,
    gap: retirementGap,
    pm: retirementPm,
    py: retirementPm * 12,
  });

  enriched.forEach((g, idx) => {
    const isRetirement = g.id === "retirement";
    const targetRow = document.createElement("tr");
    targetRow.innerHTML = `
      <td>${idx + 1}</td>
      <td>${escHtml(g.name)}</td>
      <td>${g.targetYear}</td>
      <td>${
        isRetirement
          ? `${g.years}`
          : `<input type="number" min="0" value="${g.years}" data-goal-id="${g.id}" data-goal-key="years">`
      }</td>
      <td>${
        isRetirement
          ? `${formatRs(g.amount)}`
          : `<input type="number" min="0" value="${Math.round(g.amount)}" data-goal-id="${g.id}" data-goal-key="amount">`
      }</td>
      <td>${Math.round(g.inflation * 100)}%</td>
      <td>${formatRs(g.projectedValue)}</td>
      <td>${formatRs(g.corpus)}</td>
    `;
    targetBody.appendChild(targetRow);

    const strategyRow = document.createElement("tr");
    strategyRow.innerHTML = `
      <td>${idx + 1}</td>
      <td>${escHtml(g.name)}</td>
      <td>${g.targetYear}</td>
      <td>${g.years}</td>
      <td>${
        isRetirement
          ? `${formatRs(g.provision)}`
          : `<input type="number" min="0" value="${Math.round(g.provision)}" data-goal-id="${g.id}" data-goal-key="provision">`
      }</td>
      <td>${formatRs(g.gap)}</td>
      <td>${formatRs(g.pm)}</td>
      <td>${formatRs(g.py)}</td>
    `;
    strategyBody.appendChild(strategyRow);
  });

  const totalGoalCorpus = enriched.reduce((sum, g) => sum + g.corpus, 0);
  const grossPm = enriched.reduce((sum, g) => sum + g.pm, 0);
  const requiredSip = Math.max(0, grossPm - model.currentSipPm);
  { const _e1 = byId("goalTargetTotal"); if (_e1) _e1.textContent = formatRs(totalGoalCorpus); }
  { const _e2 = byId("lessCurrentSipPm"); if (_e2) _e2.textContent = formatRs(model.currentSipPm); }
  { const _e3 = byId("lessCurrentSipPy"); if (_e3) _e3.textContent = formatRs(model.currentSipPm * 12); }
  { const _e4 = byId("requiredSip"); if (_e4) _e4.textContent = formatRs(requiredSip); }
  { const _e5 = byId("requiredSipYearly"); if (_e5) _e5.textContent = formatRs(requiredSip * 12); }

  document.querySelectorAll("#sheet-goal input[data-goal-id]").forEach((input) => {
    input.addEventListener("change", () => {
      const goal = goals.find((x) => x.id === input.dataset.goalId);
      if (!goal) return;
      goal[input.dataset.goalKey] = Number(input.value || 0);
      renderGoalInputRows();
      recalc();
    });
  });

  return { totalGoalCorpus, requiredSip, goalStrategyRows: enriched };
}

function renderGoalPie(goalOutput) {
  const svg = byId("goalPieChart");
  const legend = byId("goalPieLegend");
  if (!svg || !legend) return;

  const total = goalOutput.reduce((sum, g) => sum + g.py, 0);
  const data = total > 0 ? goalOutput.map((g) => ({ name: g.name, value: g.py })) : [];
  const colors = ["#3c78d8", "#cc4125", "#91c33b", "#674ea7", "#32a7c7", "#f1c232"];

  const W = 720, H = 340;
  // Shift pie left so right-side labels have more room
  const cx = 240, cy = 170, r = 115;

  // ── 1. Compute raw positions for each slice ──────────────────────
  let startAngle = -Math.PI / 2;
  const items = data.map((item, idx) => {
    const frac = item.value / total;
    const endAngle = startAngle + frac * 2 * Math.PI;
    const mid = startAngle + frac * Math.PI;       // midpoint angle
    const largeArc = frac > 0.5 ? 1 : 0;
    const color = colors[idx % colors.length];
    const share = Math.round(frac * 100);
    const onRight = Math.cos(mid) >= 0;

    // Elbow point sits just outside the slice edge
    const elbowR = r + 22;
    const elbowX = cx + elbowR * Math.cos(mid);
    const elbowY = cy + elbowR * Math.sin(mid);

    // Horizontal end of leader line (left or right rail)
    const railX = onRight ? cx + r + 80 : cx - r - 80;

    const result = {
      idx, item, frac, share, color, largeArc, mid, onRight,
      // Slice arc endpoints
      x1: cx + r * Math.cos(startAngle), y1: cy + r * Math.sin(startAngle),
      x2: cx + r * Math.cos(endAngle),   y2: cy + r * Math.sin(endAngle),
      // Leader line points
      edgeX: cx + (r - 4) * Math.cos(mid),
      edgeY: cy + (r - 4) * Math.sin(mid),
      elbowX, elbowY,
      railX,
      // Label y starts at elbow height — collision resolution adjusts this
      labelY: elbowY,
    };
    startAngle = endAngle;
    return result;
  });

  // ── 2. Collision resolution — spread labels that are too close ───
  const MIN_GAP = 28;   // minimum px between consecutive label baselines
  const PAD_TOP = 14, PAD_BOT = H - 14;

  ["right", "left"].forEach(side => {
    const grp = items
      .filter(it => side === "right" ? it.onRight : !it.onRight)
      .sort((a, b) => a.labelY - b.labelY);

    // Forward pass — push down
    for (let i = 1; i < grp.length; i++) {
      if (grp[i].labelY - grp[i - 1].labelY < MIN_GAP)
        grp[i].labelY = grp[i - 1].labelY + MIN_GAP;
    }
    // Backward pass — push up
    for (let i = grp.length - 2; i >= 0; i--) {
      if (grp[i + 1].labelY - grp[i].labelY < MIN_GAP)
        grp[i].labelY = grp[i + 1].labelY - MIN_GAP;
    }
    // Clamp within SVG bounds
    grp.forEach(it => {
      it.labelY = Math.max(PAD_TOP + 12, Math.min(PAD_BOT, it.labelY));
    });
  });

  // ── 3. Build SVG markup ─────────────────────────────────────────
  let slices = "";
  items.forEach(it => {
    slices += `<path d="M ${cx} ${cy} L ${it.x1.toFixed(1)} ${it.y1.toFixed(1)} A ${r} ${r} 0 ${it.largeArc} 1 ${it.x2.toFixed(1)} ${it.y2.toFixed(1)} Z"
      fill="${it.color}" stroke="#fff" stroke-width="1.5"/>`;
  });

  let leaders = "";
  items.forEach(it => {
    const tx = it.onRight ? it.railX + 5 : it.railX - 5;
    const anchor = it.onRight ? "start" : "end";
    // Elbow leader: slice-edge → elbow → horizontal to rail at adjusted labelY
    leaders += `<polyline points="${it.edgeX.toFixed(1)},${it.edgeY.toFixed(1)} ${it.elbowX.toFixed(1)},${it.elbowY.toFixed(1)} ${it.railX.toFixed(1)},${it.labelY.toFixed(1)}"
      fill="none" stroke="#94a3b8" stroke-width="0.9"/>`;
    leaders += `<text x="${tx}" y="${(it.labelY - 3).toFixed(1)}" font-size="12"
      text-anchor="${anchor}" fill="#1e293b" font-weight="500">${escHtml(it.item.name)}</text>`;
    leaders += `<text x="${tx}" y="${(it.labelY + 12).toFixed(1)}" font-size="11"
      text-anchor="${anchor}" fill="#6b7280">${it.share}%</text>`;
  });

  svg.innerHTML = `
    <rect width="${W}" height="${H}" fill="#f8fafc" rx="6"/>
    ${slices}
    ${leaders}
  `;

  legend.innerHTML = "";
  data.forEach((d, i) => {
    const item = document.createElement("p");
    item.innerHTML = `
      <span style="display:inline-block;width:10px;height:10px;background:${colors[i % colors.length]};border-radius:2px;margin-right:6px;"></span>
      ${escHtml(d.name)}: ${formatRs(d.value)}
    `;
    legend.appendChild(item);
  });
}

function renderNetworth() {
  const rows = [
    { label: "Home", amount: model.assetHome },
    { label: "Car", amount: model.assetCar },
    { label: "Gold", amount: model.assetGold },
    { label: "Liquid MF", amount: model.invLiquidMf },
    { label: "Savings Bank", amount: model.invSavings },
    { label: "Shares", amount: model.invShares },
    { label: "Equity MF", amount: model.invEquityMf },
    { label: "Debt MF", amount: model.invDebtMf },
    { label: "ELSS / Tax Saver MF", amount: model.invElss || 0 },
    { label: "Bonds", amount: model.invBonds },
    { label: "Bank FD", amount: model.invBankFd || 0 },
    { label: "Cash", amount: model.invCash || 0 },
    { label: "Postal Deposits", amount: model.invPostal },
    { label: "PPF", amount: model.invPpf },
    { label: "EPF", amount: model.invEpf || 0 },
    { label: "ULIP", amount: model.invUlip },
  ];
  additionalProperties.forEach((p) => {
    const effective = Number(p.value || 0) * (Number(p.ownership ?? 100) / 100);
    rows.push({ label: `Property: ${p.name || "Unnamed"}`, amount: effective });
  });
  // Add custom assets from all categories
  if (customAssets) {
    Object.keys(customAssets).forEach(cat => {
      (customAssets[cat] || []).forEach(item => {
        if (item.name || item.value) {
          rows.push({ label: item.name || "Custom Asset", amount: Number(item.value || 0) });
        }
      });
    });
  }
  const totalAssets = rows.reduce((sum, r) => sum + r.amount, 0);
  // Add custom liabilities to total
  const customLiabTotal = (customLiabilities || []).reduce((s, l) => s + Number(l.value || 0), 0);
  const totalLiabilities = model.loanHome + model.loanCar + model.loanOther + customLiabTotal;
  const netWorth = totalAssets - totalLiabilities;

  { const _e6 = byId("totalAssets"); if (_e6) _e6.textContent = formatRs(totalAssets); }
  { const _e7 = byId("totalLiabilities"); if (_e7) _e7.textContent = formatRs(totalLiabilities); }
  { const _e8 = byId("netWorth"); if (_e8) _e8.textContent = formatRs(netWorth); }

  const body = byId("networthBody");
  if (!body) return;
  body.innerHTML = "";
  rows.forEach((r) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${escHtml(r.label)}</td><td>${formatRs(r.amount)}</td><td>${pct(totalAssets ? r.amount / totalAssets : 0)}</td>`;
    body.appendChild(tr);
  });

  renderNetworthPie(rows, totalAssets);
  return { rows, totalAssets, totalLiabilities, netWorth };
}

function renderNetworthPie(rows, totalAssets) {
  const svg = byId("networthPieChart");
  const legend = byId("networthPieLegend");
  if (!svg || !legend) return;

  const data = rows.filter((r) => r.amount > 0);
  legend.innerHTML = "";
  if (!data.length || totalAssets <= 0) {
    svg.innerHTML = `<rect width="720" height="320" fill="#d0d0d0"></rect><text x="360" y="165" text-anchor="middle" font-size="14">No asset data</text>`;
    return;
  }

  const colors = [
    "#3c78d8",
    "#cc4125",
    "#91c33b",
    "#674ea7",
    "#32a7c7",
    "#f1c232",
    "#a64d79",
    "#6aa84f",
    "#e69138",
    "#4a86e8",
    "#8e7cc3",
    "#45818e",
  ];

  const cx = 260;
  const cy = 160;
  const r = 115;
  let startAngle = -Math.PI / 2;
  let slices = "";
  let labels = "";

  data.forEach((item, idx) => {
    const frac = item.amount / totalAssets;
    const endAngle = startAngle + frac * Math.PI * 2;
    const x1 = cx + r * Math.cos(startAngle);
    const y1 = cy + r * Math.sin(startAngle);
    const x2 = cx + r * Math.cos(endAngle);
    const y2 = cy + r * Math.sin(endAngle);
    const largeArc = frac > 0.5 ? 1 : 0;
    const color = colors[idx % colors.length];

    slices += `<path d="M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${largeArc} 1 ${x2} ${y2} Z" fill="${color}" stroke="#fff" stroke-width="1.2"></path>`;

    const mid = startAngle + (endAngle - startAngle) / 2;
    const tx = cx + (r + 22) * Math.cos(mid);
    const ty = cy + (r + 22) * Math.sin(mid);
    const share = Math.round(frac * 100);
    if (share >= 4) {
      labels += `<text x="${tx}" y="${ty}" font-size="11" text-anchor="middle">${share}%</text>`;
    }

    startAngle = endAngle;
  });

  svg.innerHTML = `
    <rect width="720" height="320" fill="#d0d0d0"></rect>
    ${slices}
    ${labels}
  `;

  data.forEach((d, i) => {
    const p = document.createElement("p");
    p.innerHTML = `<span style="display:inline-block;width:10px;height:10px;background:${colors[i % colors.length]};margin-right:6px;"></span>${escHtml(d.label)}: ${formatRs(d.amount)}`;
    legend.appendChild(p);
  });
}

// ── Pure helper: compute portfolio blended ROI from assetGrowthRates + model ──
// Returns a decimal (e.g. 0.092 = 9.2%).  Safe to call before any DOM render.
function computeBlendedRoi() {
  const gr = assetGrowthRates;
  const equityR  = (gr.equity    !== null && gr.equity    !== undefined) ? gr.equity    : (model.preRetRate  || 12);
  const debtR    = (gr.debtSaving!== null && gr.debtSaving!== undefined) ? gr.debtSaving: (model.debtRate    || 7);
  const debtMfR  = (gr.debtMf    !== null && gr.debtMf    !== undefined) ? gr.debtMf   : (model.debtRate    || 7);
  const bondsFdR = (gr.bondsFd   !== null && gr.bondsFd   !== undefined) ? gr.bondsFd  : (model.debtRate    || 7);
  const otherR   = (gr.other     !== null && gr.other     !== undefined) ? gr.other    : (model.debtRate    || 7);

  const addlPropVal = (additionalProperties || []).reduce(
    (s, p) => s + Number(p.value || 0) * (Number(p.ownership || 100) / 100), 0);
  const realEstateVal = (model.assetHome || 0) + addlPropVal;

  const rows = [
    { r: gr.realEstate || 8,  a: realEstateVal },
    { r: equityR,              a: (model.invShares||0) + (model.invEquityMf||0) + (model.invElss||0) },
    { r: debtR,                a: (model.invSavings||0) + (model.invLiquidMf||0) + (model.invUlip||0) },
    { r: debtMfR,              a: (model.invDebtMf||0) },
    { r: bondsFdR,             a: (model.invBonds||0) + (model.invBankFd||0) },
    { r: otherR,               a: (model.invPostal||0) + (model.invCash||0) },
    { r: gr.ppf || 7.9,        a: (model.invPpf||0) },
    { r: gr.epf || 8.2,        a: (model.invEpf||0) },
    { r: gr.gold || 7.0,       a: (model.assetGold||0) },
  ];

  // Custom wizard assets
  ["physical", "equity", "debt"].forEach((cat) => {
    (customAssets[cat] || []).forEach((ca, i) => {
      const key = `custom_${cat}_${i}`;
      const defRate = cat === "equity" ? equityR : (cat === "physical" ? (gr.realEstate || 8) : debtR);
      rows.push({
        r: (gr[key] !== null && gr[key] !== undefined) ? gr[key] : defRate,
        a: Number(ca.value || 0),
      });
    });
  });

  const total = rows.reduce((s, r) => s + r.a, 0);
  if (!total) return (model.preRetRate || 12) / 100;   // fallback if no assets
  return rows.reduce((s, r) => s + (r.a / total) * (r.r / 100), 0);
}

function renderRoiTable() {
  const gr = assetGrowthRates;
  const equityR  = (gr.equity    !== null && gr.equity    !== undefined) ? gr.equity    : (model.preRetRate  || 12);
  const debtR    = (gr.debtSaving!== null && gr.debtSaving!== undefined) ? gr.debtSaving: (model.debtRate    || 7);
  const debtMfR  = (gr.debtMf    !== null && gr.debtMf    !== undefined) ? gr.debtMf   : (model.debtRate    || 7);
  const bondsFdR = (gr.bondsFd   !== null && gr.bondsFd   !== undefined) ? gr.bondsFd  : (model.debtRate    || 7);
  const otherR   = (gr.other     !== null && gr.other     !== undefined) ? gr.other    : (model.debtRate    || 7);

  const addlPropVal = (additionalProperties || []).reduce(
    (s, p) => s + Number(p.value || 0) * (Number(p.ownership || 100) / 100), 0);
  const realEstateVal = (model.assetHome || 0) + addlPropVal;

  const roiRows = [
    { key: "realEstate",  p: "Real Estate",                  r: gr.realEstate || 8,  a: realEstateVal },
    { key: "equity",      p: "Equity (Shares + MF + ELSS)",  r: equityR,              a: (model.invShares||0) + (model.invEquityMf||0) + (model.invElss||0) },
    { key: "debtSaving",  p: "Debt – Saving / Liquid / ULIP",r: debtR,                a: (model.invSavings||0) + (model.invLiquidMf||0) + (model.invUlip||0) },
    { key: "debtMf",      p: "Debt MF",                      r: debtMfR,              a: (model.invDebtMf||0) },
    { key: "bondsFd",     p: "Bonds & FDs",                  r: bondsFdR,             a: (model.invBonds||0) + (model.invBankFd||0) },
    { key: "other",       p: "Other Investment",             r: otherR,               a: (model.invPostal||0) + (model.invCash||0) },
    { key: "ppf",         p: "PPF",                          r: gr.ppf || 7.9,        a: (model.invPpf||0) },
    { key: "epf",         p: "EPF",                          r: gr.epf || 8.2,        a: (model.invEpf||0) },
    { key: "gold",        p: "Gold",                         r: gr.gold || 7.0,       a: (model.assetGold||0) },
  ];

  ["physical", "equity", "debt"].forEach((cat) => {
    (customAssets[cat] || []).forEach((ca, i) => {
      const key = `custom_${cat}_${i}`;
      const defRate = cat === "equity" ? equityR : (cat === "physical" ? (gr.realEstate || 8) : debtR);
      roiRows.push({
        key,
        p: `${ca.name || "Custom"} (${cat})`,
        r: (gr[key] !== null && gr[key] !== undefined) ? gr[key] : defRate,
        a: Number(ca.value || 0),
      });
    });
  });

  const total = roiRows.reduce((s, r) => s + r.a, 0);
  const body = byId("roiBody");
  if (!body) return;
  body.innerHTML = "";
  let blendedRoi = 0;

  roiRows.forEach((r) => {
    const w   = total ? r.a / total : 0;
    const roi = w * (r.r / 100);
    blendedRoi += roi;
    const isEquity = r.key === "equity";
    // Equity rate is locked at 12% — not editable (used as the standard equity assumption)
    const rateCell = isEquity
      ? `<span class="roi-locked-badge" title="Locked at 12% — standard equity return assumption">12.0% 🔒</span>`
      : `<input type="number" class="roi-rate-inp" data-roi-key="${escHtml(r.key)}"
           value="${Number(r.r).toFixed(1)}" step="0.1" min="0" max="50"
           style="width:56px;text-align:right;padding:2px 4px;">`;
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td style="white-space:nowrap">${escHtml(r.p)}</td>
      <td>${rateCell}</td>
      <td>${formatRs(r.a)}</td>
      <td>${pct(w)}</td>
      <td>${pct(roi)}</td>`;
    body.appendChild(tr);
  });

  const totalRow = document.createElement("tr");
  totalRow.style.fontWeight = "bold";
  totalRow.innerHTML = `<td>Blended Total</td><td>—</td><td>${formatRs(total)}</td><td>100%</td><td>${pct(blendedRoi)}</td>`;
  body.appendChild(totalRow);

  // Banner: blended rate IS the Cash Flow pre-ret growth rate
  const avgEl = byId("roiWeightedAvg");
  if (avgEl) {
    const blendedPct = (blendedRoi * 100).toFixed(1);
    avgEl.innerHTML = `✦ Blended return <strong>${blendedPct}%</strong> — used as Cash Flow pre-retirement growth rate &nbsp;
      <span style="color:#64748b;font-size:0.9em;">(Pre-ret rate in My Page / Goal Sheet SIP calc: ${model.preRetRate || 12}%)</span>`;
  }

  // Rate input change → re-run full recalc so CF table/chart update immediately
  body.querySelectorAll(".roi-rate-inp").forEach((inp) => {
    inp.addEventListener("change", () => {
      assetGrowthRates[inp.dataset.roiKey] = Number(inp.value || 0);
      recalc();
      scheduleAutosave();
    });
  });

  latestState.blendedRoi = blendedRoi;
}

function renderCashflowTable(rows) {
  const body = byId("cashflowBody");
  if (!body) return;
  body.innerHTML = "";
  const retAge = model.retirementAge || 60;

  rows.forEach((r) => {
    const isPreRet = r.age <= retAge;
    const ov = cashflowOverrides[r.year] || {};
    const tr = document.createElement("tr");
    if (!isPreRet) tr.classList.add("cf-post-ret");

    // Highlight overridden cells
    const ciClass = ov.cashIn  !== undefined ? "cf-edit cf-overridden" : "cf-edit";
    const coClass = ov.cashOut !== undefined ? "cf-edit cf-overridden" : "cf-edit";

    const lsClass = ov.lumpSum ? "cf-edit cf-overridden" : "cf-edit";
    // Pre-ret growth = weighted average from ROI table (read-only; edit rates there).
    // Post-ret growth = postRetRate (editable here — changes global post-ret rate).
    const grClass = ov.growth !== undefined ? "cf-edit cf-growth-inp cf-overridden" : "cf-edit cf-growth-inp";
    const growthCell = `<input type="number" class="${grClass}"
        data-cf-year="${r.year}" data-cf-key="growth" data-cf-post-ret="${isPreRet ? "0" : "1"}"
        value="${(r.growth * 100).toFixed(1)}" step="0.1">`;
    tr.innerHTML = `
      <td>${r.no}</td>
      <td>${r.year}</td>
      <td>${r.age}</td>
      <td>${formatRs(r.opBal)}</td>
      <td><input type="number" class="${ciClass}" data-cf-year="${r.year}" data-cf-key="cashIn"
          value="${Math.round(r.cashIn)}" ${!isPreRet ? "disabled" : ""}></td>
      <td><input type="number" class="${lsClass}" data-cf-year="${r.year}" data-cf-key="lumpSum"
          value="${Math.round(r.lumpSum || 0)}" placeholder="0"></td>
      <td>${growthCell}</td>
      <td>${formatRs(r.fvEnd)}</td>
      <td><input type="number" class="${coClass}" data-cf-year="${r.year}" data-cf-key="cashOut"
          value="${Math.round(r.cashOut)}"></td>
      <td class="${r.clBal < 0 ? "cf-negative" : ""}">${formatRs(r.clBal)}</td>
      <td>${escHtml(r.goals)}</td>
    `;
    body.appendChild(tr);
  });
}

// Compact axis label: "1.6Cr", "45L", "12K" — fits in tight left padding
function fmtChartVal(n) {
  const abs = Math.abs(n);
  if (abs >= 1e7) return `${(n / 1e7).toFixed(1)}Cr`;
  if (abs >= 1e5) return `${(n / 1e5).toFixed(1)}L`;
  if (abs >= 1e3) return `${(n / 1e3).toFixed(0)}K`;
  return String(Math.round(n));
}

function renderCashflowChart(rows, targetId = "cashflowChart") {
  const svg = byId(targetId);
  if (!svg || !rows.length) return;

  // Read dimensions from the SVG's own viewBox so both the
  // main chart (900×300) and the dashboard mini chart (700×220)
  // use the correct coordinate space and nothing gets clipped.
  const vb = svg.viewBox.baseVal;
  const w  = (vb && vb.width)  || 900;
  const h  = (vb && vb.height) || 300;

  const p = { top: 16, right: 16, bottom: 40, left: 64 };
  const innerW = w - p.left - p.right;
  const innerH = h - p.top  - p.bottom;

  const vals = rows.map(r => r.clBal);
  const maxY  = Math.max(...vals, 1);
  const minY  = Math.min(...vals, 0);   // handles negative balance
  const range = maxY - minY || 1;

  const xStep = innerW / Math.max(rows.length - 1, 1);
  const xp = i => p.left + i * xStep;
  const yp = v => p.top  + ((maxY - v) / range) * innerH;

  const points = rows.map((r, i) => `${xp(i).toFixed(1)},${yp(r.clBal).toFixed(1)}`).join(" ");

  // Y-axis grid + labels (6 ticks: 0 … maxY)
  const TICKS = 5;
  let grid = "";
  for (let i = 0; i <= TICKS; i++) {
    const val = minY + (range / TICKS) * i;
    const yy  = yp(val).toFixed(1);
    grid += `<line x1="${p.left}" y1="${yy}" x2="${w - p.right}" y2="${yy}"
               stroke="#e2e8f0" stroke-width="1" stroke-dasharray="3,3"/>`;
    grid += `<text x="${p.left - 7}" y="${(+yy + 4).toFixed(1)}"
               text-anchor="end" font-size="13" fill="#94a3b8">${fmtChartVal(val)}</text>`;
  }

  // Red zero-line when balance can go negative
  const zeroLine = minY < 0
    ? `<line x1="${p.left}" y1="${yp(0).toFixed(1)}" x2="${w - p.right}" y2="${yp(0).toFixed(1)}"
         stroke="#ef4444" stroke-width="1.5" stroke-dasharray="4,3"/>`
    : "";

  // X-axis age labels — thin out to avoid overlap
  // Cap at 10 visible labels; always include first and last.
  const maxLabels = w < 750 ? 7 : 10;
  const labelEvery = Math.max(1, Math.ceil(rows.length / maxLabels));
  const labels = rows.map((r, i) => {
    if (i % labelEvery !== 0 && i !== rows.length - 1) return "";
    return `<text x="${xp(i).toFixed(1)}" y="${h - 5}"
              text-anchor="middle" font-size="10" fill="#94a3b8">${r.age}</text>`;
  }).join("");

  // Dot markers only for shorter series (≤ 40 rows) to avoid noise
  const markers = rows.length <= 40
    ? rows.map((r, i) =>
        `<circle cx="${xp(i).toFixed(1)}" cy="${yp(r.clBal).toFixed(1)}"
           r="2.5" fill="#2563eb" opacity="0.85"/>`
      ).join("")
    : "";

  svg.innerHTML = `
    <rect width="${w}" height="${h}" fill="#f8fafc" rx="6"/>
    ${grid}
    ${zeroLine}
    <polyline points="${points}" fill="none" stroke="#2563eb" stroke-width="2.5"
      stroke-linejoin="round" stroke-linecap="round"/>
    ${markers}
    ${labels}
  `;

  // Event markers only on the full Cash Flow sheet (not the dashboard mini-chart)
  if (targetId === "cashflowChart") {
    addCfEventMarkers(svg, rows, xp, yp, p, w, h);
  }
}

// ── Cash Flow event markers ───────────────────────────────────────────────────
// Builds significant-event markers (arrows + tooltips) from the cashflow rows
// and appends them as live DOM elements so hover listeners work.
function addCfEventMarkers(svg, rows, xp, yp, pad, svgW, svgH) {
  const ns = "http://www.w3.org/2000/svg";

  // ── 1. Collect events ───────────────────────────────────────────────────────
  const events = [];
  let retirementMarked = false;

  for (let i = 0; i < rows.length; i++) {
    const r    = rows[i];
    const prev = rows[i - 1];

    // Goal payouts (non-retirement goals in this year)
    const goalNames = r.goals
      ? r.goals.split(" & ").filter(g => g.trim() !== "Retirement" && g.trim() !== "")
      : [];
    if (goalNames.length && r.cashOut > 0) {
      events.push({
        i,
        label : goalNames.join(" & "),
        sub   : `−${fmtChartVal(r.cashOut)}  ·  Age ${r.age}`,
        color : "#ef4444",
      });
    }

    // Retirement begins — first year with a retirement outflow
    if (!retirementMarked && r.goals && r.goals.includes("Retirement") && r.cashOut > 0) {
      retirementMarked = true;
      events.push({
        i,
        label : "Retirement Begins",
        sub   : `−${fmtChartVal(r.cashOut)}/yr  ·  Age ${r.age}`,
        color : "#f59e0b",
      });
    }

  }

  if (!events.length) return;

  // ── 2. Tooltip div ──────────────────────────────────────────────────────────
  const wrap = svg.parentElement;
  wrap.style.position = "relative";
  const oldTip = wrap.querySelector(".cf-event-tooltip");
  if (oldTip) oldTip.remove();
  const tooltip = document.createElement("div");
  tooltip.className = "cf-event-tooltip";
  wrap.appendChild(tooltip);

  // ── 3. Render each marker ───────────────────────────────────────────────────
  events.forEach(ev => {
    const cx = xp(ev.i);
    const cy = yp(rows[ev.i].clBal);

    // If data point is near the top, put the arrow below it; else above
    const below   = cy < pad.top + 60;
    const dir     = below ? 1 : -1;          // +1 = downward arrow, -1 = upward
    const arrowLen = 20;
    const tipY    = cy + dir * arrowLen;     // arrowhead tip (touching the line)
    const baseY   = tipY + dir * 14;         // arrow shaft base / label anchor

    const g = document.createElementNS(ns, "g");
    g.setAttribute("class", "cf-event-marker");
    g.style.cursor = "pointer";

    // Dashed vertical guide line
    const vl = document.createElementNS(ns, "line");
    vl.setAttribute("x1", cx);  vl.setAttribute("y1", pad.top);
    vl.setAttribute("x2", cx);  vl.setAttribute("y2", svgH - pad.bottom);
    vl.setAttribute("stroke", ev.color);
    vl.setAttribute("stroke-width", "1");
    vl.setAttribute("stroke-dasharray", "4 3");
    vl.setAttribute("opacity", "0.4");
    g.appendChild(vl);

    // Arrow shaft (from baseY toward the data point)
    const shaft = document.createElementNS(ns, "line");
    shaft.setAttribute("x1", cx);  shaft.setAttribute("y1", baseY);
    shaft.setAttribute("x2", cx);  shaft.setAttribute("y2", tipY + dir * -6);
    shaft.setAttribute("stroke", ev.color);
    shaft.setAttribute("stroke-width", "2");
    g.appendChild(shaft);

    // Arrowhead (triangle pointing at the data point)
    const head = document.createElementNS(ns, "polygon");
    head.setAttribute("points", `${cx},${tipY} ${cx-4.5},${tipY+dir*-9} ${cx+4.5},${tipY+dir*-9}`);
    head.setAttribute("fill", ev.color);
    g.appendChild(head);

    // White circle on the data-point for emphasis
    const dot = document.createElementNS(ns, "circle");
    dot.setAttribute("cx", cx);  dot.setAttribute("cy", cy);
    dot.setAttribute("r", "5");
    dot.setAttribute("fill", "#fff");
    dot.setAttribute("stroke", ev.color);
    dot.setAttribute("stroke-width", "2");
    g.appendChild(dot);

    // Short label (first 2 words, fits in tight space)
    const shortLabel = ev.label.split(/[\s&\/]+/).slice(0, 2).join(" ");
    const lbl = document.createElementNS(ns, "text");
    lbl.setAttribute("x", cx);
    lbl.setAttribute("y", below ? baseY + 13 : baseY - 3);
    lbl.setAttribute("text-anchor", "middle");
    lbl.setAttribute("fill", ev.color);
    lbl.setAttribute("font-size", "11");
    lbl.setAttribute("font-weight", "700");
    lbl.setAttribute("letter-spacing", "0.025em");
    lbl.textContent = shortLabel;
    g.appendChild(lbl);

    // Tooltip on hover
    g.addEventListener("mouseenter", () => {
      tooltip.innerHTML =
        `<strong>${escHtml(ev.label)}</strong><span>${escHtml(ev.sub)}</span>`;
      tooltip.style.display = "block";

      const svgRect  = svg.getBoundingClientRect();
      const wrapRect = wrap.getBoundingClientRect();
      const scaleX   = svgRect.width  / svgW;
      const scaleY   = svgRect.height / svgH;

      // Position tooltip centred on x, above or below the arrow base
      let left = cx * scaleX + (svgRect.left - wrapRect.left) - 60;
      let top  = below
        ? (baseY + 18) * scaleY
        : (baseY - 44) * scaleY;

      // Keep tooltip inside the wrap horizontally
      left = Math.max(0, Math.min(left, wrapRect.width - 130));
      tooltip.style.left = left + "px";
      tooltip.style.top  = top  + "px";
    });
    g.addEventListener("mouseleave", () => { tooltip.style.display = "none"; });

    svg.appendChild(g);
  });

  // ── 4. Legend strip below chart ─────────────────────────────────────────────
  const oldLegend = wrap.querySelector(".cf-event-legend");
  if (oldLegend) oldLegend.remove();
  const legend = document.createElement("div");
  legend.className = "cf-event-legend";
  const legendItems = [
    { color: "#ef4444", label: "Goal payout"      },
    { color: "#f59e0b", label: "Retirement begins" },
  ];
  legend.innerHTML = legendItems.map(li =>
    `<span class="cf-leg-item">
       <span class="cf-leg-dot" style="background:${li.color}"></span>${escHtml(li.label)}
     </span>`
  ).join("");
  wrap.appendChild(legend);
}

function renderBreakup(goalStrategyRows) {
  const body = byId("breakupBody");
  if (!body) return;
  body.innerHTML = "";
  goalStrategyRows.forEach((g, i) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${i + 1}</td>
      <td>${escHtml(g.name)}</td>
      <td>${formatRs(g.corpus)}</td>
      <td>${formatRs(g.provision)}</td>
      <td>${formatRs(g.gap)}</td>
      <td>${formatRs(g.pm)}</td>
    `;
    body.appendChild(tr);
  });
}

function rowCurrentValue(r) {
  // Graceful fallback for old equity rows that stored units * nav instead of currentValue.
  if (r.currentValue !== undefined) return Number(r.currentValue || 0);
  if (r.units !== undefined && r.nav !== undefined) return Number(r.units || 0) * Number(r.nav || 0);
  return 0;
}

function rowName(r) {
  // Support both new "name" key and legacy "schemeName" / "investorName" keys.
  return r.name || r.schemeName || r.investorName || "";
}

function renderAdminPortfolioRows(bodyId, rows, type) {
  const body = byId(bodyId);
  if (!body) return;
  body.innerHTML = "";
  rows.forEach((r, idx) => {
    const cv = rowCurrentValue(r);
    const nm = rowName(r);
    const disabled = isAdmin() ? "" : "disabled";
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td><input data-admin-type="${type}" data-admin-idx="${idx}" data-key="name"
          value="${escHtml(nm)}" placeholder="Enter name" ${disabled}></td>
      <td><input type="number" data-admin-type="${type}" data-admin-idx="${idx}" data-key="currentValue"
          value="${cv}" placeholder="0" ${disabled}></td>
      <td class="admin-only">${isAdmin() ? `<button type="button" data-admin-del="${type}:${idx}">Delete</button>` : ""}</td>
    `;
    body.appendChild(tr);
  });
}

function renderAdminNetworthSheet() {
  { const _e9 = byId("adminAsOfDate"); if (_e9) _e9.value = adminPortfolio.asOfDate || ""; }
  const eq = adminPortfolio.equityRows || [];
  const uf = adminPortfolio.unifiRows || [];
  const ic = adminPortfolio.iciciRows || [];
  renderAdminPortfolioRows("adminEquityBody", eq, "equity");
  renderAdminPortfolioRows("adminUnifiBody", uf, "unifi");
  renderAdminPortfolioRows("adminIciciBody", ic, "icici");

  // Use rowCurrentValue() so both old (units*nav) and new (currentValue) formats work.
  const totalEq = eq.reduce((s, r) => s + rowCurrentValue(r), 0);
  const totalUf = uf.reduce((s, r) => s + rowCurrentValue(r), 0);
  const totalIc = ic.reduce((s, r) => s + rowCurrentValue(r), 0);

  // Update per-section subtotal cells.
  const stocksTotalEl = byId("stocksTotal");
  if (stocksTotalEl) stocksTotalEl.textContent = formatRs(totalEq);
  const pmsTotalEl = byId("pmsTotal");
  if (pmsTotalEl) pmsTotalEl.textContent = formatRs(totalUf);
  const mfTotalEl = byId("mfTotal");
  if (mfTotalEl) mfTotalEl.textContent = formatRs(totalIc);

  { const _e10 = byId("adminTotalPortfolio"); if (_e10) _e10.value = formatRs(totalEq + totalUf + totalIc); }

  document.querySelectorAll("input[data-admin-type]").forEach((el) => {
    el.addEventListener("change", () => {
      if (!isAdmin()) return;
      const t = el.dataset.adminType;
      const idx = Number(el.dataset.adminIdx);
      const key = el.dataset.key;
      const collection = t === "equity" ? adminPortfolio.equityRows : t === "unifi" ? adminPortfolio.unifiRows : adminPortfolio.iciciRows;
      collection[idx][key] = ["costValue", "units", "nav", "sipAmt", "currentValue", "xirr"].includes(key)
        ? Number(el.value || 0)
        : el.value;
      // Keep legacy helper fields in sync when editing via new simplified keys.
      if (key === "name") { collection[idx].schemeName = el.value; }
      renderAdminNetworthSheet();
      scheduleAutosave();
    });
  });
  document.querySelectorAll("button[data-admin-del]").forEach((btn) => {
    btn.addEventListener("click", () => {
      if (!isAdmin()) return;
      const [t, idxRaw] = btn.dataset.adminDel.split(":");
      const idx = Number(idxRaw);
      const collection = t === "equity" ? adminPortfolio.equityRows : t === "unifi" ? adminPortfolio.unifiRows : adminPortfolio.iciciRows;
      collection.splice(idx, 1);
      renderAdminNetworthSheet();
      scheduleAutosave();
    });
  });
}

// ═══════════════════════════════════════════════════════════
// DASHBOARD
// ═══════════════════════════════════════════════════════════

// ── Balance sheet donut pie chart ────────────────────────────
function renderBalancePie(rows) {
  const container = byId("db-balancePie");
  if (!container || !rows.length) return;
  const SZ = 280, CX = SZ / 2, CY = SZ / 2, R = 120, IR = 72;
  const total = rows.reduce((s, r) => s + r.value, 0) || 1;
  let angle = -Math.PI / 2;
  const slices = rows.map(r => {
    const sweep = (r.value / total) * 2 * Math.PI;
    const end   = angle + sweep;
    const la    = sweep > Math.PI ? 1 : 0;
    const cos1  = Math.cos(angle), sin1 = Math.sin(angle);
    const cos2  = Math.cos(end),   sin2 = Math.sin(end);
    const path  = [
      `M ${(CX + IR*cos1).toFixed(2)} ${(CY + IR*sin1).toFixed(2)}`,
      `L ${(CX + R*cos1).toFixed(2)}  ${(CY + R*sin1).toFixed(2)}`,
      `A ${R} ${R} 0 ${la} 1 ${(CX + R*cos2).toFixed(2)} ${(CY + R*sin2).toFixed(2)}`,
      `L ${(CX + IR*cos2).toFixed(2)} ${(CY + IR*sin2).toFixed(2)}`,
      `A ${IR} ${IR} 0 ${la} 0 ${(CX + IR*cos1).toFixed(2)} ${(CY + IR*sin1).toFixed(2)} Z`,
    ].join(" ");
    const pct = ((r.value / total) * 100).toFixed(1);
    angle = end;
    return { ...r, path, pct };
  });

  const svgPaths = slices.map(s =>
    `<path d="${s.path}" fill="${s.color}" stroke="#fff" stroke-width="1.5" opacity="0.92">
       <title>${s.label}: ${formatRs(s.value)} (${s.pct}%)</title></path>`
  ).join("");

  const legend = slices.map(s =>
    `<div class="bp-row">
       <span class="bp-dot" style="background:${s.color}"></span>
       <span class="bp-lbl">${s.label}</span>
       <span class="bp-pct">${s.pct}%</span>
     </div>`
  ).join("");

  container.innerHTML = `
    <div class="bp-wrap">
      <svg viewBox="0 0 ${SZ} ${SZ}" class="bp-svg">
        ${svgPaths}
        <text x="${CX}" y="${CY - 10}" text-anchor="middle" font-size="13" fill="#9ca3af">Assets</text>
        <text x="${CX}" y="${CY + 14}" text-anchor="middle" font-size="17" fill="#111827" font-weight="700">${fmtChartVal(total)}</text>
      </svg>
      <div class="bp-legend">${legend}</div>
    </div>`;
}

function renderDashboard() {
  if (!byId("sheet-dashboard")) return;

  // ── Hero welcome card ────────────────────────────────────
  const dbName = byId("db-name");
  if (dbName) dbName.textContent = `Welcome back, ${model.name || "—"}`;

  // Avatar initials (up to 2 chars from name)
  const heroAvatar = byId("db-hero-avatar");
  if (heroAvatar) {
    const words = (model.name || "").trim().split(/\s+/).filter(Boolean);
    heroAvatar.textContent = words.length >= 2
      ? (words[0][0] + words[words.length - 1][0]).toUpperCase()
      : (model.name || "?").slice(0, 2).toUpperCase();
  }

  const dbSub = byId("db-subtitle");
  if (dbSub) {
    const planYear = model.planDate ? new Date(model.planDate).getFullYear() : "—";
    const yearsToRet = model.retirementAge && model.planDate
      ? Math.max(0, (model.retirementAge - (new Date().getFullYear() - new Date(model.planDate).getFullYear() + (model.retirementAge ? 0 : 0))))
      : null;
    const currentAge = model.dob
      ? Math.floor((new Date() - new Date(model.dob)) / (365.25 * 24 * 3600 * 1000))
      : null;
    const retYearsLeft = (currentAge && model.retirementAge)
      ? Math.max(0, model.retirementAge - currentAge)
      : null;
    dbSub.textContent = retYearsLeft !== null
      ? `Plan Date: ${planYear}  ·  Retirement in ${retYearsLeft} year${retYearsLeft !== 1 ? "s" : ""}`
      : `Plan Date: ${planYear}  ·  Retirement Age: ${model.retirementAge || "—"}`;
  }

  // Hero chips
  const chipsEl = byId("db-hero-chips");
  if (chipsEl) {
    const chips = [];
    if (model.retirementAge)  chips.push(`Ret. Age ${model.retirementAge}`);
    if (model.lifeExpectancy) chips.push(`Life Exp. ${model.lifeExpectancy}`);
    if (model.city)           chips.push(`📍 ${model.city}`);
    chipsEl.innerHTML = chips.map(c => `<span class="db-hero-chip">${escHtml(c)}</span>`).join("");
  }

  // Populate navbar center pill (shows on mobile)
  const navUser = byId("navbarUser");
  if (navUser) {
    const label = model.name || (currentUser && currentUser.email) || "";
    navUser.textContent = label;
    navUser.hidden = !label;
  }

  // ── Balance Sheet numbers ────────────────────────────────
  const stocksVal = (adminPortfolio.equityRows || []).reduce((s, r) => s + rowCurrentValue(r), 0);
  const pmsVal    = (adminPortfolio.unifiRows || []).reduce((s, r) => s + rowCurrentValue(r), 0);
  const mfAdminVal= (adminPortfolio.iciciRows || []).reduce((s, r) => s + rowCurrentValue(r), 0);
  const mfInputVal= (model.invEquityMf||0) + (model.invDebtMf||0) + (model.invElss||0) + (model.invLiquidMf||0);
  // Use admin portfolio MF if populated, else use input sheet values
  const mfVal = mfAdminVal > 0 ? mfAdminVal : mfInputVal;

  const financialAssets = stocksVal + pmsVal + mfVal
    + (model.invEpf||0) + (model.invPpf||0) + (model.invSavings||0) + (model.invShares||0)
    + (model.invBonds||0) + (model.invBankFd||0) + (model.invCash||0)
    + (model.invPostal||0) + (model.invUlip||0);

  const propVal = (model.assetHome||0) + additionalProperties.reduce(
    (s, p) => s + Number(p.value||0) * (Number(p.ownership||100) / 100), 0);
  const physicalAssets = propVal + (model.assetCar||0) + (model.assetGold||0);

  // Include custom assets from all categories in dashboard total
  const customAssetsDashTotal = customAssets ? Object.values(customAssets).reduce((s, arr) =>
    s + (arr || []).reduce((s2, item) => s2 + Number(item.value || 0), 0), 0) : 0;
  const totalAssets = financialAssets + physicalAssets + customAssetsDashTotal;
  const customLiabDashTotal = (customLiabilities || []).reduce((s, l) => s + Number(l.value || 0), 0);
  const totalLiabilities = (model.loanHome||0) + (model.loanCar||0) + (model.loanOther||0) + customLiabDashTotal;
  const netWorth = totalAssets - totalLiabilities;

  const db = byId;
  if (db("db-totalAssets"))      db("db-totalAssets").textContent      = formatRs(totalAssets);
  if (db("db-totalLiabilities")) db("db-totalLiabilities").textContent = formatRs(totalLiabilities);
  if (db("db-netWorth"))         db("db-netWorth").textContent         = formatRs(netWorth);
  // Colour net worth card depending on sign
  const nwCard = byId("db-netWorthCard");
  if (nwCard) {
    nwCard.className = "cards-article " + (netWorth >= 0 ? "card-brand" : "card-danger");
  }

  // ── Assets breakdown bars ────────────────────────────────
  const breakdownRows = [
    { label: "Stocks",         value: stocksVal,          color: "#3b82f6" },
    { label: "PMS",            value: pmsVal,             color: "#6366f1" },
    { label: "Mutual Funds",   value: mfVal,              color: "#10b981",
      sub: [
        { label: "Equity MF",    value: model.invEquityMf||0 },
        { label: "Debt MF",      value: model.invDebtMf||0  },
        { label: "ELSS",         value: model.invElss||0    },
        { label: "Liquid/Hybrid",value: model.invLiquidMf||0},
      ].filter(s => s.value > 0),
    },
    { label: "EPF",            value: model.invEpf||0,   color: "#f59e0b" },
    { label: "PPF",            value: model.invPpf||0,   color: "#f97316" },
    { label: "Shares",         value: model.invShares||0, color: "#ec4899" },
    { label: "Savings/FD",     value: model.invSavings||0,color: "#8b5cf6" },
    { label: "Bonds",          value: model.invBonds||0,  color: "#06b6d4" },
    { label: "Postal/NSC",     value: model.invPostal||0, color: "#84cc16" },
    { label: "ULIP",           value: model.invUlip||0,   color: "#a78bfa" },
    { label: "Property",       value: propVal,            color: "#0ea5e9" },
    { label: "Car",            value: model.assetCar||0,  color: "#64748b" },
    { label: "Gold",           value: model.assetGold||0, color: "#fbbf24" },
  ].filter(r => r.value > 0);

  const bContainer = byId("db-assetsBreakdown");
  if (bContainer) {
    const bTotal = breakdownRows.reduce((s, r) => s + r.value, 0) || 1;
    bContainer.innerHTML = breakdownRows.map(r => {
      const pctW = ((r.value / bTotal) * 100).toFixed(1);
      const subHtml = r.sub && r.sub.length
        ? `<div class="bs-sub">${r.sub.map(s => `
            <div class="bs-sub-row">
              <span class="bs-sub-label">${s.label}</span>
              <span class="bs-sub-val">${formatRs(s.value)}</span>
            </div>`).join("")}</div>`
        : "";
      return `
        <div class="bs-row">
          <div class="bs-row-top">
            <span class="bs-dot" style="background:${r.color}"></span>
            <span class="bs-lbl">${r.label}</span>
            <span class="bs-val">${formatRs(r.value)}</span>
            <span class="bs-pct">${pctW}%</span>
          </div>
          <div class="bs-bar-wrap">
            <div class="bs-bar" style="width:${pctW}%;background:${r.color}17;border-left:3px solid ${r.color}"></div>
          </div>
          ${subHtml}
        </div>`;
    }).join("");
  }

  // ── Balance sheet pie ────────────────────────────────────
  renderBalancePie(breakdownRows);

  // ── Will fields ─────────────────────────────────────────
  const ws = byId("db-willStatus");      if (ws) ws.value = model.willStatus || "";
  const wl = byId("db-willLastUpdated"); if (wl) wl.value = model.willLastUpdated || "";
  const nu = byId("db-nominationsUpdated"); if (nu) nu.value = model.nominationsUpdated || "";

  // ── Insurance tables ─────────────────────────────────────
  renderInsuranceTables();

  // ── Goals Summary ────────────────────────────────────────
  const gs = latestState.goalSummary;
  if (gs) {
    if (byId("db-totalCorpus")) byId("db-totalCorpus").textContent = formatRs(gs.totalCorpus || 0);
    if (byId("db-sipRequired")) byId("db-sipRequired").textContent = formatRs(gs.requiredSip || 0);
    if (byId("db-currentSip"))  byId("db-currentSip").textContent  = formatRs(model.currentSipPm || 0);

    // KPI trend badges on Goals Summary cards
    const sipCoverage = gs.requiredSip > 0
      ? Math.round((model.currentSipPm / gs.requiredSip) * 100) : 0;
    const sipBadgeEl = byId("db-sipRequired")?.closest("article");
    if (sipBadgeEl && !sipBadgeEl.querySelector(".kpi-trend")) {
      const badge = document.createElement("span");
      badge.className = `kpi-trend ${sipCoverage >= 80 ? "up" : sipCoverage >= 40 ? "flat" : "down"}`;
      badge.textContent = sipCoverage >= 80 ? `↑ ${sipCoverage}% covered`
        : sipCoverage >= 40 ? `${sipCoverage}% covered`
        : `↓ ${sipCoverage}% covered`;
      sipBadgeEl.appendChild(badge);
    }

    const gb = byId("db-goalsBody");
    if (gb && gs.goalStrategyRows) {
      gb.innerHTML = gs.goalStrategyRows.map(g => {
        const pct = g.corpus > 0 ? Math.min(100, Math.round((g.provision / g.corpus) * 100)) : 0;
        const progClass = pct >= 75 ? "prog-good" : pct >= 40 ? "prog-warn" : "prog-danger";
        const lblColor  = pct >= 75 ? "var(--success)" : pct >= 40 ? "var(--warning)" : "var(--danger)";
        return `
          <tr>
            <td>
              <div class="goal-name-cell">${escHtml(g.name)}</div>
              <div class="goal-prog-wrap">
                <div class="goal-prog-bar ${progClass}" style="width:${pct}%"></div>
              </div>
              <div class="goal-prog-lbl" style="color:${lblColor}">${pct}% funded</div>
            </td>
            <td>${g.targetYear}</td>
            <td>${g.years}</td>
            <td>${formatRs(g.corpus)}</td>
            <td>${formatRs(g.pm)}</td>
          </tr>`;
      }).join("");
    }
  }

  // ── Cash Flow preview ────────────────────────────────────
  const cf = latestState.cashflow;
  if (cf && cf.length) {
    const retRow = cf.find(r => r.age >= (model.retirementAge || 60));
    if (retRow) {
      if (byId("db-retirementYear"))   byId("db-retirementYear").textContent   = retRow.year;
      if (byId("db-retirementCorpus")) byId("db-retirementCorpus").textContent = formatRs(retRow.clBal);
    }
    const retYrs = Math.max(0, (model.lifeExpectancy||85) - (model.retirementAge||60));
    if (byId("db-retirementYears")) byId("db-retirementYears").textContent = retYrs;
    renderCashflowChart(cf, "db-cashflowChart");
  }
}

// ── Insurance table rendering ────────────────────────────────
function renderInsuranceTables() {
  renderInsuranceBodyRows("db-lifeInsBody", lifeInsuranceRows, "life");
  renderInsuranceBodyRows("db-healthInsBody", healthInsuranceRows, "health");
  renderInsuranceBodyRows("db-carInsBody", carInsuranceRows, "car");
  renderInsuranceBodyRows("db-propInsBody", propertyInsuranceRows, "property");

  const lifeSA   = lifeInsuranceRows.reduce((s, r) => s + Number(r.sumAssured||0), 0);
  const lifePrem = lifeInsuranceRows.reduce((s, r) => s + Number(r.annualPrem||0), 0);
  const lifeSurr = lifeInsuranceRows.reduce((s, r) => s + Number(r.surrenderVal||0), 0);
  if (byId("db-lifeSA"))   byId("db-lifeSA").textContent   = formatRs(lifeSA);
  if (byId("db-lifePrem")) byId("db-lifePrem").textContent = formatRs(lifePrem);
  if (byId("db-lifeSurr")) byId("db-lifeSurr").textContent = formatRs(lifeSurr);

  const healthSA   = healthInsuranceRows.reduce((s, r) => s + Number(r.sumAssured||0), 0);
  const healthPrem = healthInsuranceRows.reduce((s, r) => s + Number(r.annualPrem||0), 0);
  if (byId("db-healthSA"))   byId("db-healthSA").textContent   = formatRs(healthSA);
  if (byId("db-healthPrem")) byId("db-healthPrem").textContent = formatRs(healthPrem);

  const carIDV  = carInsuranceRows.reduce((s, r) => s + Number(r.idv||0), 0);
  const carPrem = carInsuranceRows.reduce((s, r) => s + Number(r.annualPrem||0), 0);
  if (byId("db-carIDV"))  byId("db-carIDV").textContent  = formatRs(carIDV);
  if (byId("db-carPrem")) byId("db-carPrem").textContent = formatRs(carPrem);

  const propCover = propertyInsuranceRows.reduce((s, r) => s + Number(r.cover||0), 0);
  const propPrem  = propertyInsuranceRows.reduce((s, r) => s + Number(r.annualPrem||0), 0);
  if (byId("db-propCover")) byId("db-propCover").textContent = formatRs(propCover);
  if (byId("db-propPrem"))  byId("db-propPrem").textContent  = formatRs(propPrem);
}

function renderInsuranceBodyRows(bodyId, rows, type) {
  const body = byId(bodyId);
  if (!body) return;
  body.innerHTML = "";
  const inpClass = "ins-input";
  const numericKeys = ["sumAssured","annualPrem","surrenderVal","idv","cover"];

  rows.forEach((r, idx) => {
    const tr = document.createElement("tr");
    let cells = `<td>${idx + 1}</td>`;

    if (type === "life") {
      cells += `
        <td><input class="${inpClass}" data-ins-type="${type}" data-ins-idx="${idx}" data-ins-key="policyName" value="${escHtml(r.policyName||"")}" placeholder="Policy name"></td>
        <td><input class="${inpClass}" data-ins-type="${type}" data-ins-idx="${idx}" data-ins-key="company" value="${escHtml(r.company||"")}" placeholder="Company"></td>
        <td><input type="number" class="${inpClass}" data-ins-type="${type}" data-ins-idx="${idx}" data-ins-key="sumAssured" value="${r.sumAssured||0}"></td>
        <td><input type="number" class="${inpClass}" data-ins-type="${type}" data-ins-idx="${idx}" data-ins-key="annualPrem" value="${r.annualPrem||0}"></td>
        <td><input type="number" class="${inpClass}" data-ins-type="${type}" data-ins-idx="${idx}" data-ins-key="surrenderVal" value="${r.surrenderVal||0}"></td>`;
    } else if (type === "health") {
      cells += `
        <td><input class="${inpClass}" data-ins-type="${type}" data-ins-idx="${idx}" data-ins-key="policyName" value="${escHtml(r.policyName||"")}" placeholder="Policy name"></td>
        <td><input class="${inpClass}" data-ins-type="${type}" data-ins-idx="${idx}" data-ins-key="company" value="${escHtml(r.company||"")}" placeholder="Insurer"></td>
        <td><input type="number" class="${inpClass}" data-ins-type="${type}" data-ins-idx="${idx}" data-ins-key="sumAssured" value="${r.sumAssured||0}"></td>
        <td><input type="number" class="${inpClass}" data-ins-type="${type}" data-ins-idx="${idx}" data-ins-key="annualPrem" value="${r.annualPrem||0}"></td>
        <td><input class="${inpClass}" data-ins-type="${type}" data-ins-idx="${idx}" data-ins-key="members" value="${escHtml(r.members||"")}" placeholder="Members covered"></td>`;
    } else if (type === "car") {
      cells += `
        <td><input class="${inpClass}" data-ins-type="${type}" data-ins-idx="${idx}" data-ins-key="policyName" value="${escHtml(r.policyName||"")}" placeholder="Policy name"></td>
        <td><input class="${inpClass}" data-ins-type="${type}" data-ins-idx="${idx}" data-ins-key="company" value="${escHtml(r.company||"")}" placeholder="Insurer"></td>
        <td><input type="number" class="${inpClass}" data-ins-type="${type}" data-ins-idx="${idx}" data-ins-key="idv" value="${r.idv||0}"></td>
        <td><input type="number" class="${inpClass}" data-ins-type="${type}" data-ins-idx="${idx}" data-ins-key="annualPrem" value="${r.annualPrem||0}"></td>
        <td><input type="date" class="${inpClass}" data-ins-type="${type}" data-ins-idx="${idx}" data-ins-key="expiry" value="${r.expiry||""}"></td>`;
    } else if (type === "property") {
      cells += `
        <td><input class="${inpClass}" data-ins-type="${type}" data-ins-idx="${idx}" data-ins-key="policyName" value="${escHtml(r.policyName||"")}" placeholder="Policy name"></td>
        <td><input class="${inpClass}" data-ins-type="${type}" data-ins-idx="${idx}" data-ins-key="company" value="${escHtml(r.company||"")}" placeholder="Insurer"></td>
        <td><input class="${inpClass}" data-ins-type="${type}" data-ins-idx="${idx}" data-ins-key="propertyName" value="${escHtml(r.propertyName||"")}" placeholder="Property"></td>
        <td><input type="number" class="${inpClass}" data-ins-type="${type}" data-ins-idx="${idx}" data-ins-key="cover" value="${r.cover||0}"></td>
        <td><input type="number" class="${inpClass}" data-ins-type="${type}" data-ins-idx="${idx}" data-ins-key="annualPrem" value="${r.annualPrem||0}"></td>`;
    }
    cells += `<td><button type="button" data-ins-del="${type}:${idx}">Delete</button></td>`;
    tr.innerHTML = cells;
    body.appendChild(tr);
  });

  const arrMap = { life: lifeInsuranceRows, health: healthInsuranceRows, car: carInsuranceRows, property: propertyInsuranceRows };
  body.querySelectorAll("input[data-ins-key]").forEach(el => {
    el.addEventListener("change", () => {
      const arr = arrMap[el.dataset.insType];
      const i   = Number(el.dataset.insIdx);
      const k   = el.dataset.insKey;
      if (arr && arr[i]) {
        arr[i][k] = numericKeys.includes(k) ? Number(el.value||0) : el.value;
      }
      renderInsuranceTables();
      scheduleAutosave();
    });
  });
  body.querySelectorAll("button[data-ins-del]").forEach(btn => {
    btn.addEventListener("click", () => {
      const [t, idxRaw] = btn.dataset.insDel.split(":");
      const arr = arrMap[t];
      if (arr) arr.splice(Number(idxRaw), 1);
      renderInsuranceTables();
      scheduleAutosave();
    });
  });
}

// ═══════════════════════════════════════════════════════════
// 7-STEP QUESTIONNAIRE WIZARD
// ═══════════════════════════════════════════════════════════

const WIZARD_STEPS = [
  { title: "Your Family",          subtitle: "Tell us about your household" },
  { title: "Income",               subtitle: "Your monthly earnings" },
  { title: "Monthly Expenses",     subtitle: "What you spend each month" },
  { title: "Goals",                subtitle: "What you're working towards" },
  { title: "Retirement Planning",  subtitle: "When and how you want to retire" },
  { title: "Assets & Investments", subtitle: "What you own and have invested" },
  { title: "Insurance",            subtitle: "Your protection coverage" },
  { title: "Will & Estate",        subtitle: "Legal & nomination status" },
];

function openWizard() {
  const ov = byId("wizardOverlay");
  if (!ov) return;
  wizardCurrentStep = 0;
  ov.hidden = false;
  renderWizardStep(0);
}

function closeWizard() {
  const ov = byId("wizardOverlay");
  if (ov) ov.hidden = true;
}

function wizardFinish() {
  model.wizardCompleted = true;
  closeWizard();
  // Lock My Page after questionnaire completion
  lockMyPage();
  // Switch to dashboard tab
  const dashBtn = document.querySelector('#sheetTabs button[data-sheet="dashboard"]');
  if (dashBtn) dashBtn.click();
  saveCurrentPlan().catch(e => setStatus(`Save failed: ${e.message}`));
  // Push questionnaire results to Google Sheet
  pushToGoogleSheet();
}

function lockMyPage() {
  // List of input IDs on My Page that should be locked after wizard completion
  const myPageInputIds = [
    "inflationRate",
    "educationInflationRate",
    "marriageInflationRate",
    "preRetRate",
    "postRetRate",
    "cashInGrowthRate",
    "retirementAge",
    "lifeExpectancy",
    "debtRate",
    "retirementMonthlyExp",
  ];

  myPageInputIds.forEach(id => {
    const el = byId(id);
    if (!el) return;
    el.disabled = true;
    el.title = "This field is locked after questionnaire submission";
  });

  // Add a visual indicator
  const myPageSheet = byId("sheet-mypage");
  if (myPageSheet && !myPageSheet.querySelector(".page-locked-indicator")) {
    const indicator = document.createElement("div");
    indicator.className = "page-locked-indicator";
    indicator.innerHTML = "🔒 Locked after questionnaire submission";
    indicator.style.cssText = `
      background: #ecfdf5;
      border: 1px solid #a7f3d0;
      border-left: 4px solid #059669;
      border-radius: 6px;
      padding: 0.75rem 1rem;
      margin-bottom: 1rem;
      font-size: 0.9rem;
      color: #047857;
      font-weight: 500;
    `;
    const panel = myPageSheet.querySelector(".panel");
    if (panel) panel.insertBefore(indicator, panel.firstChild);
  }
}

// ── Google Sheet integration ───────────────────────────────────────────────
const GOOGLE_SHEET_WEBHOOK = "https://script.google.com/macros/s/AKfycbz89ghOAZTTbk85jUwJ_Z5xSpZo4_pVWnXs1IlYEHoB_rxW53ObYKZZxOLklRnfH9ec/exec";

function pushToGoogleSheet() {
  try {
    const user = auth.currentUser;

    // ── Personal & Contact Info ────────────────────────────────────────
    const data = {
      timestamp:        new Date().toISOString(),
      email:            user ? user.email    : "",
      investorName:     model.name           || "",
      planDate:         model.planDate       || "",
      dob:              model.dob            || "",
      spouseDob:        model.spouseDob      || "",
      child1Dob:        model.child1Dob      || "",
      child2Dob:        model.child2Dob      || "",
      city:             model.city           || "",
      state:            model.state          || "",

      // ── Income ─────────────────────────────────────────────────────
      incomeMain:       model.incomeMain     || 0,
      incomeSpouse:     model.incomeSpouse   || 0,

      // ── Expenses (Monthly) ─────────────────────────────────────────
      expHousehold:     model.expHousehold   || 0,
      expLifestyle:     model.expLifestyle   || 0,
      expEducation:     model.expEducation   || 0,
      expVehicle:       model.expVehicle     || 0,
      expMediclaim:     model.expMediclaim   || 0,
      expUtilities:     model.expUtilities   || 0,
      expCarInsurance:  model.expCarInsurance || 0,
      expMisc:          model.expMisc        || 0,
      expLifeIns:       model.expLifeIns     || 0,
      expVacation:      model.expVacation    || 0,
      expRent:          model.expRent        || 0,
      expCreditCard:    model.expCreditCard  || 0,
      expTravel:        model.expTravel      || 0,
      expProfFees:      model.expProfFees    || 0,
      expPpfMonthly:    model.expPpfMonthly  || 0,

      // ── Assets ─────────────────────────────────────────────────────
      assetHome:        model.assetHome      || 0,
      assetCar:         model.assetCar       || 0,
      assetGold:        model.assetGold      || 0,

      // ── Investments ────────────────────────────────────────────────
      invLiquidMf:      model.invLiquidMf    || 0,
      invSavings:       model.invSavings     || 0,
      invShares:        model.invShares      || 0,
      invEquityMf:      model.invEquityMf    || 0,
      invDebtMf:        model.invDebtMf      || 0,
      invBonds:         model.invBonds       || 0,
      invPostal:        model.invPostal      || 0,
      invPpf:           model.invPpf         || 0,
      invUlip:          model.invUlip        || 0,
      invEpf:           model.invEpf         || 0,
      invElss:          model.invElss        || 0,

      // ── Liabilities ────────────────────────────────────────────────
      loanHome:         model.loanHome       || 0,
      loanCar:          model.loanCar        || 0,
      loanOther:        model.loanOther      || 0,

      // ── Rates & Assumptions ────────────────────────────────────────
      inflationRate:    model.inflationRate  || 0,
      educationInflationRate: model.educationInflationRate || 0,
      marriageInflationRate:  model.marriageInflationRate  || 0,
      preRetRate:       model.preRetRate     || 0,
      postRetRate:      model.postRetRate    || 0,
      cashInGrowthRate: model.cashInGrowthRate || 0,
      debtRate:         model.debtRate       || 0,

      // ── Life Milestones ────────────────────────────────────────────
      retirementAge:    model.retirementAge  || 0,
      lifeExpectancy:   model.lifeExpectancy || 0,
      retirementMonthlyExp: model.retirementMonthlyExp || 0,

      // ── Current Plan State ─────────────────────────────────────────
      currentSipPm:     model.currentSipPm   || 0,

      // ── Goals (aggregated) ─────────────────────────────────────────
      goals:            (model.goals || []).map(g => g.name).join(" | "),
      numGoals:         (model.goals || []).length,

      // ── Additional Properties (dynamic) ────────────────────────────
      numAdditionalProperties: additionalProperties.length,
      additionalPropertiesJson: JSON.stringify(additionalProperties),

      // ── Life & Health Insurance ────────────────────────────────────
      numLifeInsuranceRows: lifeInsuranceRows.length,
      numHealthInsuranceRows: healthInsuranceRows.length,
      lifeInsuranceJson: JSON.stringify(lifeInsuranceRows),
      healthInsuranceJson: JSON.stringify(healthInsuranceRows),

      // ── Admin Portfolio (if applicable) ────────────────────────────
      adminPortfolioAsOfDate: adminPortfolio.asOfDate || "",
      numAdminEquityRows: (adminPortfolio.equityRows || []).length,
      numAdminUnifiRows: (adminPortfolio.unifiRows || []).length,
      numAdminIciciRows: (adminPortfolio.iciciRows || []).length,
      adminPortfolioJson: JSON.stringify(adminPortfolio),

      // ── Compliance & Documentation ─────────────────────────────────
      willStatus:       model.willStatus     || "",
      willLastUpdated:  model.willLastUpdated || "",
      nominationsUpdated: model.nominationsUpdated || "",

      // ── Health & Family History ────────────────────────────────────
      familyHistoryCriticalIllness: model.familyHistoryCriticalIllness || "",
      familyHistoryDescription: model.familyHistoryDescription || "",

      // ── Notes ──────────────────────────────────────────────────────
      networthNotes:    model.networthNotes  || "",

      // ── Wizard Status ──────────────────────────────────────────────
      wizardCompleted:  model.wizardCompleted ? "Yes" : "No",
    };

    // Apps Script requires no-cors; we fire-and-forget
    fetch(GOOGLE_SHEET_WEBHOOK, {
      method:  "POST",
      mode:    "no-cors",
      headers: { "Content-Type": "application/json" },
      body:    JSON.stringify(data),
    }).catch(err => console.warn("Sheet push failed:", err));
  } catch (err) {
    console.warn("pushToGoogleSheet error:", err);
  }
}

function renderWizardStep(n) {
  const step = WIZARD_STEPS[n];
  const titleEl    = byId("wizardTitle");
  const subtitleEl = byId("wizardSubtitle");
  const labelEl    = byId("wizardStepLabel");
  const fillEl     = byId("wizardProgressFill");
  const contentEl  = byId("wizardContent");
  const backBtn    = byId("wizardBackBtn");
  const nextBtn    = byId("wizardNextBtn");

  if (titleEl)    titleEl.textContent    = step.title;
  if (subtitleEl) subtitleEl.textContent = step.subtitle;
  if (labelEl)    labelEl.textContent    = `Step ${n + 1} of ${WIZARD_STEPS.length}`;
  if (fillEl)     fillEl.style.width     = `${((n + 1) / WIZARD_STEPS.length) * 100}%`;
  if (backBtn)    backBtn.hidden         = (n === 0);
  if (nextBtn)    nextBtn.textContent    = (n === WIZARD_STEPS.length - 1) ? "Save & Finish ✓" : "Next →";

  // Dots
  const dotsEl = byId("wizardDots");
  if (dotsEl) {
    dotsEl.innerHTML = WIZARD_STEPS.map((_, i) =>
      `<span class="wz-dot${i === n ? " active" : ""}"></span>`).join("");
  }

  const renderers = [wz1, wz2, wz3, wz4, wzRetirement, wz5, wz6, wz7];
  if (contentEl && renderers[n]) contentEl.innerHTML = renderers[n]();

  // Attach wizard-specific live listeners after render
  contentEl.querySelectorAll("input[data-wz], select[data-wz], textarea[data-wz]").forEach(el => {
    el.addEventListener("input", () => {
      const k = el.dataset.wz;
      model[k] = (el.type === "number" || el.type === "range") ? Number(el.value||0) : el.value;
      if (k === "dob" || k === "planDate") recalc();
      scheduleAutosave();
    });
  });

  // Step 5 (Assets): wire custom asset/liability add/edit/delete
  contentEl.querySelectorAll("[data-add-custom]").forEach(btn => {
    btn.addEventListener("click", () => {
      const type = btn.dataset.addCustom;  // "asset" or "liability"
      const cat = btn.dataset.addCat;      // "physical", "equity", etc.
      if (type === "asset") {
        if (!customAssets[cat]) customAssets[cat] = [];
        customAssets[cat].push({ name: "", value: 0 });
      } else {
        customLiabilities.push({ name: "", value: 0 });
      }
      renderWizardStep(wizardCurrentStep);
      scheduleAutosave();
    });
  });

  contentEl.querySelectorAll("[data-custom-key]").forEach(el => {
    el.addEventListener("input", () => {
      const type = el.dataset.customType;
      const cat = el.dataset.customCat;
      const idx = Number(el.dataset.customIdx);
      const key = el.dataset.customKey;
      const arr = type === "asset" ? customAssets[cat] : customLiabilities;
      if (arr && arr[idx]) {
        arr[idx][key] = key === "value" ? Number(el.value || 0) : el.value;
        recalc();
        scheduleAutosave();
      }
    });
  });

  contentEl.querySelectorAll("[data-custom-del-idx]").forEach(btn => {
    btn.addEventListener("click", () => {
      const type = btn.dataset.customDelType;
      const cat = btn.dataset.customDelCat;
      const idx = Number(btn.dataset.customDelIdx);
      if (type === "asset") {
        customAssets[cat].splice(idx, 1);
      } else {
        customLiabilities.splice(idx, 1);
      }
      renderWizardStep(wizardCurrentStep);
      recalc();
      scheduleAutosave();
    });
  });

  // Step 5 (Assets): wire Additional Properties add/edit/delete inside wizard
  if (n === 5) {
    byId("wzAddPropBtn")?.addEventListener("click", () => {
      additionalProperties.push({ name: "", value: 0, ownership: 100, loanLinked: "" });
      renderWizardStep(5);
      scheduleAutosave();
    });
    contentEl.querySelectorAll("input[data-prop-idx]").forEach(el => {
      el.addEventListener("change", () => {
        const idx = Number(el.dataset.propIdx);
        const key = el.dataset.propKey;
        additionalProperties[idx][key] = (key === "value" || key === "ownership") ? Number(el.value||0) : el.value;
        recalc();
        scheduleAutosave();
      });
    });
    contentEl.querySelectorAll(".wz-prop-del").forEach(btn => {
      btn.addEventListener("click", () => {
        additionalProperties.splice(Number(btn.dataset.delProp), 1);
        renderWizardStep(5);
        scheduleAutosave();
      });
    });
  }
}

// Helpers used across steps
function fi(id, label, type="text", placeholder="") {
  const val = type==="number" ? (model[id]||0) : (model[id]||"");
  return `<label>${label}<input data-wz="${id}" type="${type}" value="${escHtml(String(val))}" placeholder="${placeholder}"></label>`;
}

function fis(id, label, placeholder="") { return fi(id, label, "number", placeholder); }

// ── Step 1: Family ───────────────────────────────────────────
function wz1() {
  const numC = children.length;
  const childRows = Array.from({length: numC}, (_, i) => `
    <div class="wz-child-row">
      <span class="wz-child-num">${i+1}</span>
      <label>Name<input data-child-idx="${i}" data-child-key="name"
        value="${escHtml(children[i]?.name||"")}" placeholder="Child's name"></label>
      <label>Date of Birth<input type="date" data-child-idx="${i}" data-child-key="dob"
        value="${children[i]?.dob||""}"></label>
    </div>`).join("");

  return `
    <div class="wz-grid two">
      ${fi("name","Full Name","text","Your name")}
      ${fi("planDate","Plan Date","date")}
      ${fi("dob","Your Date of Birth","date")}
      <label>Age<input disabled value="${yearsBetween(model.dob,model.planDate)||"—"}"></label>
      ${fi("city","City","text","City")}
      ${fi("state","State","text","State")}
      ${fi("spouseDob","Spouse Date of Birth","date")}
      <label>Spouse Age<input disabled value="${yearsBetween(model.spouseDob,model.planDate)||"—"}"></label>
    </div>
    <div class="wz-section-label">Children</div>
    <div class="wz-row-center">
      <label>Number of Children
        <select id="wzNumChildren" onchange="wzSetNumChildren(+this.value)">
          ${[0,1,2,3,4,5].map(n=>`<option value="${n}"${n===numC?" selected":""}>${n}</option>`).join("")}
        </select>
      </label>
    </div>
    <div id="wzChildRows" class="wz-children">${childRows}</div>
  `;
}

// ── Step 2: Income ──────────────────────────────────────────
function wz2() {
  const isPrivate = !model.invEpf;
  return `
    <div class="wz-grid two">
      ${fis("incomeMain","Self Monthly Income (₹)","e.g. 100000")}
      ${fis("incomeSpouse","Spouse Monthly Income (₹)","0 if not applicable")}
      <label>EPF Monthly Contribution (₹)
        <input data-wz="invEpf" type="number" value="${model.invEpf||0}" placeholder="0">
        ${isPrivate ? '<span class="wz-note">EPF = 0 → Private sector assumed</span>' : ''}
      </label>
      <label>Other Income (₹/month)<input type="number" disabled value="—" placeholder="Add other income in assets"></label>
    </div>
    <div class="wz-info-box">
      <strong>Estimated Monthly Surplus:</strong>
      <span id="wzSurplus">${formatRs(Math.max(0,
        (model.incomeMain||0) + (model.incomeSpouse||0) -
        ((model.expHousehold||0)+(model.expLifestyle||0)+(model.expEducation||0)+
         (model.expVehicle||0)+(model.expMediclaim||0)+(model.expUtilities||0)+
         (model.expCarInsurance||0)+(model.expMisc||0)+
         (model.expLifeIns||0)+(model.expVacation||0)+(model.expRent||0)+
         (model.expCreditCard||0)+(model.expTravel||0)+(model.expProfFees||0)+(model.expPpfMonthly||0)+
         customExpenses.reduce((s,e)=>s+(Number(e.amount)||0),0))
      ))}</span> / month
    </div>
  `;
}

// ── Step 3: Expenses ────────────────────────────────────────
function wz3() {
  const customRows = customExpenses.map((e, i) => `
    <div class="wz-custom-exp-row">
      <label style="flex:2;">Expense Name
        <input type="text" data-custom-exp-idx="${i}" data-custom-exp-key="name"
          value="${escHtml(e.name||"")}" placeholder="e.g. Club membership">
      </label>
      <label>Amount (₹/month)
        <input type="number" data-custom-exp-idx="${i}" data-custom-exp-key="amount"
          value="${e.amount||0}" placeholder="0">
      </label>
      <label style="flex:2;">Note (optional)
        <input type="text" data-custom-exp-idx="${i}" data-custom-exp-key="note"
          value="${escHtml(e.note||"")}" placeholder="Any note about this expense">
      </label>
      <button type="button" onclick="wzDelCustomExp(${i})" style="margin-top:1.4rem;">✕</button>
    </div>`).join("");

  return `
    <div class="wz-section-label">Core Expenses</div>
    <div class="wz-grid two">
      ${fis("expHousehold","Household Expenses (₹/month)")}
      ${fis("expLifestyle","Lifestyle & Dining (₹/month)")}
      ${fis("expRent","Rent / EMI (₹/month)")}
      ${fis("expEducation","Children's Education (₹/month)")}
      ${fis("expVehicle","Vehicle Expenses (₹/month)")}
      ${fis("expUtilities","Utilities (₹/month)")}
    </div>

    <div class="wz-section-label">Insurance &amp; Finance</div>
    <div class="wz-grid two">
      ${fis("expMediclaim","Health Insurance Premium (₹/month)")}
      ${fis("expLifeIns","Life Insurance Premium (₹/month)")}
      ${fis("expCarInsurance","Car Insurance (₹/month)")}
      ${fis("expCreditCard","Credit Card Expenses (₹/month)")}
    </div>

    <div class="wz-section-label">Leisure &amp; Professional</div>
    <div class="wz-grid two">
      ${fis("expVacation","Vacation / Holidays (₹/month)")}
      ${fis("expTravel","Travel Expenses (₹/month)")}
      ${fis("expProfFees","Professional Fees (₹/month)")}
      ${fis("expMisc","Miscellaneous (₹/month)")}
    </div>

    <div class="wz-section-label">Savings &amp; Investments</div>
    <div class="wz-grid two">
      ${fis("expPpfMonthly","PPF Contribution (₹/month)")}
      ${fis("currentSipPm","Current SIPs (₹/month)")}
    </div>

    <div class="wz-section-label">Custom Expenses</div>
    <div id="wzCustomExpList">${customRows}</div>
    <button type="button" class="wz-add-btn" onclick="wzAddCustomExp()">+ Add Custom Expense</button>
  `;
}

// ── Step 4: Goals ───────────────────────────────────────────
function wz4() {
  // Ensure each child has education + marriage goals
  children.forEach((c, i) => {
    ["education","marriage"].forEach(type => {
      const existing = goals.find(g => g.childIdx === i && g.inflationType === type);
      if (!existing) {
        goals.push({
          id: `child-${i}-${type}`,
          name: `${c.name||`Child ${i+1}`}'s ${type.charAt(0).toUpperCase()+type.slice(1)}`,
          kind: "goal",
          inflationType: type,
          childIdx: i,
          years: 0, amount: 0, provision: 0,
        });
      } else if (c.name) {
        existing.name = `${c.name}'s ${type.charAt(0).toUpperCase()+type.slice(1)}`;
      }
    });
  });

  const childGoalHtml = children.map((c, i) => {
    const eduG = goals.find(g => g.childIdx === i && g.inflationType === "education") || {};
    const marG = goals.find(g => g.childIdx === i && g.inflationType === "marriage") || {};
    return `
      <div class="wz-goal-child-block">
        <div class="wz-goal-child-name">${c.name || `Child ${i+1}`}</div>
        <div class="wz-grid two">
          <label>Education — Years Away
            <input type="number" data-goal-id="${eduG.id}" data-goal-key="years"
              value="${eduG.years||0}" min="0">
          </label>
          <label>Education — Current Cost (₹)
            <input type="number" data-goal-id="${eduG.id}" data-goal-key="amount"
              value="${eduG.amount||0}" placeholder="0">
          </label>
          <label>Marriage — Years Away
            <input type="number" data-goal-id="${marG.id}" data-goal-key="years"
              value="${marG.years||0}" min="0">
          </label>
          <label>Marriage — Current Cost (₹)
            <input type="number" data-goal-id="${marG.id}" data-goal-key="amount"
              value="${marG.amount||0}" placeholder="0">
          </label>
        </div>
      </div>`;
  }).join("");

  const otherGoals = goals.filter(g => g.childIdx === undefined);
  const otherGoalHtml = otherGoals.map(g => `
    <div class="wz-goal-row">
      <input data-goal-id="${g.id}" data-goal-key="name" value="${escHtml(g.name)}" placeholder="Goal name">
      <label class="wz-inline">Years<input type="number" data-goal-id="${g.id}" data-goal-key="years" value="${g.years||0}" min="0"></label>
      <label class="wz-inline">Amount (₹)<input type="number" data-goal-id="${g.id}" data-goal-key="amount" value="${g.amount||0}"></label>
      <button type="button" onclick="wzDelGoal('${g.id}')">✕</button>
    </div>`).join("");

  return `
    ${children.length ? `<div class="wz-section-label">Child Goals</div>${childGoalHtml}` : ""}
    <div class="wz-section-label">Other Goals</div>
    <div id="wzOtherGoals">${otherGoalHtml}</div>
    <div class="actions" style="margin-top:.5rem;">
      <button type="button" onclick="wzAddGoal()">+ Add Goal</button>
    </div>
  `;
}

// ── Step Retirement: Retirement Planning ────────────────────
function wzRetirement() {
  // Calculate target retirement year for display
  const currentAge    = yearsBetween(model.dob, model.planDate) || 0;
  const planYear      = model.planDate ? new Date(model.planDate).getFullYear() : new Date().getFullYear();
  const retAge        = Number(model.retirementAge) || 60;
  const targetYear    = planYear + Math.max(0, retAge - currentAge);
  const yearsToRetire = Math.max(0, retAge - currentAge);

  return `
    <div class="wz-grid two">
      <label>Retirement Age
        <input data-wz="retirementAge" type="number" min="40" max="80"
          value="${model.retirementAge || 60}" placeholder="60">
      </label>
      <label>Life Expectancy
        <input data-wz="lifeExpectancy" type="number" min="60" max="100"
          value="${model.lifeExpectancy || 85}" placeholder="85">
      </label>
      <label>Target Retirement Year <span style="font-weight:400;color:var(--muted);font-size:0.75rem;">(auto)</span>
        <input type="text" disabled value="${targetYear}"
          style="background:var(--panel-alt);color:var(--muted);">
      </label>
      <label>Years to Retirement <span style="font-weight:400;color:var(--muted);font-size:0.75rem;">(auto)</span>
        <input type="text" disabled value="${yearsToRetire} years"
          style="background:var(--panel-alt);color:var(--muted);">
      </label>
    </div>

    <div class="wz-section-label" style="margin-top:0.5rem;">Retirement Lifestyle</div>
    <div class="wz-grid two">
      <label>Estimated Monthly Expense in Retirement (₹)
        <span style="font-size:0.75rem;color:var(--muted);font-weight:400;">In today's value — inflation will be applied</span>
        <input data-wz="retirementMonthlyExp" type="number" min="0"
          value="${model.retirementMonthlyExp || 0}" placeholder="e.g. 80000">
      </label>
      <label>Post-Retirement Return Rate (%)
        <span style="font-size:0.75rem;color:var(--muted);font-weight:400;">Investment return after retirement</span>
        <input data-wz="postRetRate" type="number" step="0.1"
          value="${model.postRetRate || 8}" placeholder="8">
      </label>
    </div>

    <div class="wz-info-box" style="margin-top:0.75rem;">
      <strong>Note:</strong> If you enter a monthly expense here it will be used as the retirement cost base.
      Leave it at <strong>0</strong> to auto-derive it from your current expense profile.
      Current expense-based estimate: <strong>${formatRs(
        (model.expHousehold||0)+(model.expLifestyle||0)+(model.expRent||0)+
        (model.expVehicle||0)+(model.expMediclaim||0)+(model.expUtilities||0)+
        (model.expCarInsurance||0)+(model.expMisc||0)+
        (model.expLifeIns||0)+(model.expVacation||0)+(model.expCreditCard||0)+
        (model.expTravel||0)+(model.expProfFees||0)+(model.expPpfMonthly||0)+
        customExpenses.reduce((s,e)=>s+(Number(e.amount)||0),0)
      )}/month</strong>
    </div>
  `;
}


function renderWzCustomItems(category, type) {
  const items = type === "asset" ? (customAssets[category] || []) : customLiabilities;
  return items.map((item, idx) => `
    <div class="wz-custom-row" style="display:flex;gap:0.5rem;align-items:center;margin-top:0.35rem;">
      <input type="text" data-custom-type="${type}" data-custom-cat="${category}" data-custom-idx="${idx}" data-custom-key="name"
             placeholder="Name" value="${escHtml(item.name||"")}" style="flex:1;">
      <input type="number" data-custom-type="${type}" data-custom-cat="${category}" data-custom-idx="${idx}" data-custom-key="value"
             placeholder="Value (₹)" value="${item.value||0}" style="width:140px;">
      <button type="button" class="wz-custom-del" data-custom-del-type="${type}" data-custom-del-cat="${category}" data-custom-del-idx="${idx}" title="Remove" style="background:#ef4444;color:#fff;border:none;border-radius:4px;padding:0.2rem 0.5rem;cursor:pointer;">✕</button>
    </div>`).join("");
}

// ── Step 5: Assets ──────────────────────────────────────────
function wz5() {
  const addlPropsHtml = additionalProperties.map((p, idx) => `
    <div class="wz-prop-row" data-prop-idx="${idx}">
      <input class="wz-prop-name" data-prop-idx="${idx}" data-prop-key="name"
             placeholder="Property name" value="${escHtml(p.name||"")}">
      <input class="wz-prop-val" type="number" data-prop-idx="${idx}" data-prop-key="value"
             placeholder="Current value (₹)" value="${p.value||0}">
      <input class="wz-prop-own" type="number" data-prop-idx="${idx}" data-prop-key="ownership"
             placeholder="Ownership %" value="${p.ownership??100}" min="0" max="100">
      <input class="wz-prop-loan" data-prop-idx="${idx}" data-prop-key="loanLinked"
             placeholder="Loan linked (optional)" value="${escHtml(p.loanLinked||"")}">
      <button type="button" class="wz-prop-del" data-del-prop="${idx}" title="Remove">✕</button>
    </div>`).join("");

  return `
    <div class="wz-section-label">
      Physical Assets
      <button type="button" class="wz-add-btn" data-add-custom="asset" data-add-cat="physical">+ Add Asset</button>
    </div>
    <div class="wz-grid three">
      ${fis("assetHome","Home / Property (₹)")}
      ${fis("assetCar","Car (₹)")}
      ${fis("assetGold","Gold & Jewellery (₹)")}
    </div>
    <div id="wzCustomPhysical">${renderWzCustomItems("physical","asset")}</div>

    <div class="wz-section-label" style="margin-top:0.5rem;">
      Additional Properties
      <button type="button" id="wzAddPropBtn" class="wz-add-btn">+ Add Property</button>
    </div>
    <div id="wzPropList">
      ${addlPropsHtml || '<p class="wz-note" style="margin:0.25rem 0 0.5rem;">No additional properties added.</p>'}
    </div>

    <div class="wz-section-label" style="margin-top:1rem;">
      Equity & Market Investments
      <button type="button" class="wz-add-btn" data-add-custom="asset" data-add-cat="equity">+ Add Asset</button>
    </div>
    <div class="wz-grid three">
      ${fis("invEquityMf","Equity Mutual Funds (₹)")}
      ${fis("invElss","ELSS / Tax Saver MF (₹)")}
      ${fis("invLiquidMf","Liquid / Hybrid MF (₹)")}
      ${fis("invShares","Shares & Securities (₹)")}
    </div>
    <div id="wzCustomEquity">${renderWzCustomItems("equity","asset")}</div>

    <div class="wz-section-label" style="margin-top:0.5rem;">
      Debt & Fixed Income
      <button type="button" class="wz-add-btn" data-add-custom="asset" data-add-cat="debt">+ Add Asset</button>
    </div>
    <div class="wz-grid three">
      ${fis("invDebtMf","Debt Mutual Funds (₹)")}
      ${fis("invSavings","Savings Bank Account (₹)")}
      ${fis("invBankFd","Bank Fixed Deposits (₹)")}
      ${fis("invBonds","Bonds (₹)")}
      ${fis("invPostal","Postal / NSC (₹)")}
    </div>
    <div id="wzCustomDebt">${renderWzCustomItems("debt","asset")}</div>

    <div class="wz-section-label" style="margin-top:0.5rem;">
      Retirement & Long-term
      <button type="button" class="wz-add-btn" data-add-custom="asset" data-add-cat="retirement">+ Add Asset</button>
    </div>
    <div class="wz-grid three">
      ${fis("invPpf","PPF — Current Value (₹)")}
      <label>EPF — Current Value (₹)
        <input data-wz="invEpf" type="number" value="${model.invEpf||0}" placeholder="0">
      </label>
      ${fis("invUlip","ULIP (₹)")}
    </div>
    <div id="wzCustomRetirement">${renderWzCustomItems("retirement","asset")}</div>

    <div class="wz-section-label" style="margin-top:0.5rem;">
      Cash
      <button type="button" class="wz-add-btn" data-add-custom="asset" data-add-cat="cash">+ Add Asset</button>
    </div>
    <div class="wz-grid three">
      ${fis("invCash","Cash in Hand (₹)")}
    </div>
    <div id="wzCustomCash">${renderWzCustomItems("cash","asset")}</div>

    <div class="wz-section-label" style="margin-top:1rem;">
      Liabilities
      <button type="button" class="wz-add-btn" data-add-custom="liability" data-add-cat="liabilities">+ Add Liability</button>
    </div>
    <div class="wz-grid three">
      ${fis("loanHome","Home Loan Outstanding (₹)")}
      ${fis("loanCar","Car Loan Outstanding (₹)")}
      ${fis("loanOther","Other Loans (₹)")}
    </div>
    <div id="wzCustomLiabilities">${renderWzCustomItems("liabilities","liability")}</div>
  `;
}

// ── Step 6: Insurance ───────────────────────────────────────
function wz6() {
  const lifeRows = lifeInsuranceRows.map((r, i) => `
    <tr>
      <td>${i+1}</td>
      <td><input data-wz-ins-type="life" data-wz-ins-idx="${i}" data-wz-ins-key="policyName"
          value="${escHtml(r.policyName||"")}" placeholder="Policy name"></td>
      <td><input data-wz-ins-type="life" data-wz-ins-idx="${i}" data-wz-ins-key="company"
          value="${escHtml(r.company||"")}" placeholder="Company"></td>
      <td><input type="number" data-wz-ins-type="life" data-wz-ins-idx="${i}" data-wz-ins-key="sumAssured"
          value="${r.sumAssured||0}"></td>
      <td><input type="number" data-wz-ins-type="life" data-wz-ins-idx="${i}" data-wz-ins-key="annualPrem"
          value="${r.annualPrem||0}"></td>
      <td><input type="number" data-wz-ins-type="life" data-wz-ins-idx="${i}" data-wz-ins-key="surrenderVal"
          value="${r.surrenderVal||0}"></td>
      <td><button type="button" onclick="wzDelIns('life',${i})">✕</button></td>
    </tr>`).join("");

  const healthRows = healthInsuranceRows.map((r, i) => `
    <tr>
      <td>${i+1}</td>
      <td><input data-wz-ins-type="health" data-wz-ins-idx="${i}" data-wz-ins-key="policyName"
          value="${escHtml(r.policyName||"")}" placeholder="Policy name"></td>
      <td><input data-wz-ins-type="health" data-wz-ins-idx="${i}" data-wz-ins-key="company"
          value="${escHtml(r.company||"")}" placeholder="Company"></td>
      <td><input type="number" data-wz-ins-type="health" data-wz-ins-idx="${i}" data-wz-ins-key="sumAssured"
          value="${r.sumAssured||0}"></td>
      <td><input type="number" data-wz-ins-type="health" data-wz-ins-idx="${i}" data-wz-ins-key="annualPrem"
          value="${r.annualPrem||0}"></td>
      <td><button type="button" onclick="wzDelIns('health',${i})">✕</button></td>
    </tr>`).join("");

  const carRows = carInsuranceRows.map((r, i) => `
    <tr>
      <td>${i+1}</td>
      <td><input data-wz-ins-type="car" data-wz-ins-idx="${i}" data-wz-ins-key="policyName"
          value="${escHtml(r.policyName||"")}" placeholder="Policy name"></td>
      <td><input data-wz-ins-type="car" data-wz-ins-idx="${i}" data-wz-ins-key="company"
          value="${escHtml(r.company||"")}" placeholder="Insurer"></td>
      <td><input type="number" data-wz-ins-type="car" data-wz-ins-idx="${i}" data-wz-ins-key="idv"
          value="${r.idv||0}" placeholder="0"></td>
      <td><input type="number" data-wz-ins-type="car" data-wz-ins-idx="${i}" data-wz-ins-key="annualPrem"
          value="${r.annualPrem||0}" placeholder="0"></td>
      <td><input type="date" data-wz-ins-type="car" data-wz-ins-idx="${i}" data-wz-ins-key="expiry"
          value="${r.expiry||""}"></td>
      <td><button type="button" onclick="wzDelIns('car',${i})">✕</button></td>
    </tr>`).join("");

  const propRows = propertyInsuranceRows.map((r, i) => `
    <tr>
      <td>${i+1}</td>
      <td><input data-wz-ins-type="property" data-wz-ins-idx="${i}" data-wz-ins-key="policyName"
          value="${escHtml(r.policyName||"")}" placeholder="Policy name"></td>
      <td><input data-wz-ins-type="property" data-wz-ins-idx="${i}" data-wz-ins-key="company"
          value="${escHtml(r.company||"")}" placeholder="Insurer"></td>
      <td><input data-wz-ins-type="property" data-wz-ins-idx="${i}" data-wz-ins-key="propertyName"
          value="${escHtml(r.propertyName||"")}" placeholder="Property name"></td>
      <td><input type="number" data-wz-ins-type="property" data-wz-ins-idx="${i}" data-wz-ins-key="cover"
          value="${r.cover||0}" placeholder="0"></td>
      <td><input type="number" data-wz-ins-type="property" data-wz-ins-idx="${i}" data-wz-ins-key="annualPrem"
          value="${r.annualPrem||0}" placeholder="0"></td>
      <td><button type="button" onclick="wzDelIns('property',${i})">✕</button></td>
    </tr>`).join("");

  return `
    <div class="wz-section-label">Life Insurance Policies</div>
    <div class="table-wrap">
      <table>
        <thead><tr><th>#</th><th>Policy Name</th><th>Company</th><th>Sum Assured (₹)</th><th>Annual Premium (₹)</th><th>Surrender Value (₹)</th><th></th></tr></thead>
        <tbody id="wzLifeBody">${lifeRows}</tbody>
      </table>
    </div>
    <div class="actions"><button type="button" onclick="wzAddIns('life')">+ Add Life Policy</button></div>

    <div class="wz-section-label" style="margin-top:1.25rem;">Health Insurance Policies</div>
    <div class="table-wrap">
      <table>
        <thead><tr><th>#</th><th>Policy Name</th><th>Company</th><th>Sum Assured (₹)</th><th>Annual Premium (₹)</th><th></th></tr></thead>
        <tbody id="wzHealthBody">${healthRows}</tbody>
      </table>
    </div>
    <div class="actions"><button type="button" onclick="wzAddIns('health')">+ Add Health Policy</button></div>

    <div class="wz-section-label" style="margin-top:1.25rem;">Car Insurance</div>
    <div class="table-wrap">
      <table>
        <thead><tr><th>#</th><th>Policy Name</th><th>Insurer</th><th>IDV (₹)</th><th>Annual Premium (₹)</th><th>Expiry</th><th></th></tr></thead>
        <tbody id="wzCarBody">${carRows}</tbody>
      </table>
    </div>
    <div class="actions"><button type="button" onclick="wzAddIns('car')">+ Add Car Policy</button></div>

    <div class="wz-section-label" style="margin-top:1.25rem;">Home / Property Insurance</div>
    <div class="table-wrap">
      <table>
        <thead><tr><th>#</th><th>Policy Name</th><th>Insurer</th><th>Property</th><th>Cover (₹)</th><th>Annual Premium (₹)</th><th></th></tr></thead>
        <tbody id="wzPropBody">${propRows}</tbody>
      </table>
    </div>
    <div class="actions"><button type="button" onclick="wzAddIns('property')">+ Add Property Policy</button></div>

    <div class="wz-section-label" style="margin-top:1.5rem;">Family Medical History</div>
    <div class="wz-grid two">
      <label>Family history of critical illness?
        <span style="font-size:0.75rem;color:var(--muted);font-weight:400;">Heart disease, cancer, diabetes, etc.</span>
        <select data-wz="familyHistoryCriticalIllness" onchange="wzToggleFamilyHistory(this.value)">
          <option value=""${!model.familyHistoryCriticalIllness ? " selected" : ""}>— Select —</option>
          <option value="yes"${model.familyHistoryCriticalIllness === "yes" ? " selected" : ""}>Yes</option>
          <option value="no"${model.familyHistoryCriticalIllness === "no" ? " selected" : ""}>No</option>
        </select>
      </label>
      <div id="wzFamilyHistoryDesc" style="${model.familyHistoryCriticalIllness === 'yes' ? '' : 'display:none;'}">
        <label>Brief description
          <span style="font-size:0.75rem;color:var(--muted);font-weight:400;">e.g. Father — heart disease at 55, Mother — diabetes</span>
          <textarea data-wz="familyHistoryDescription" rows="3"
            placeholder="Describe the condition(s) and family member(s)…"
            style="resize:vertical;">${escHtml(model.familyHistoryDescription || "")}</textarea>
        </label>
      </div>
    </div>
    ${model.familyHistoryCriticalIllness === "yes" ? `
    <div class="wz-info-box" style="margin-top:0.5rem;background:var(--warning-xlight);border-color:var(--warning-light);color:var(--warning);">
      <strong>Consider:</strong> A critical illness rider or standalone critical illness plan may be advisable given your family history. Discuss with your advisor.
    </div>` : ""}
  `;
}

// ── Step 7: Will & Estate ───────────────────────────────────
function wz7() {
  return `
    <div class="wz-grid two">
      <label>Do you have a Will?
        <select data-wz="willStatus">
          <option value="">— Select —</option>
          <option value="yes"${model.willStatus==="yes"?" selected":""}>Yes</option>
          <option value="no"${model.willStatus==="no"?" selected":""}>No</option>
        </select>
      </label>
      <label>Will Last Updated
        <input type="date" data-wz="willLastUpdated" value="${model.willLastUpdated||""}">
      </label>
      <label>Are Nominations Updated?
        <select data-wz="nominationsUpdated">
          <option value="">— Select —</option>
          <option value="yes"${model.nominationsUpdated==="yes"?" selected":""}>Yes</option>
          <option value="no"${model.nominationsUpdated==="no"?" selected":""}>No</option>
        </select>
      </label>
    </div>
    <div class="wz-info-box" style="margin-top:1rem;">
      <p>🔒 Your will and nomination status is visible only to you and your advisor. Keeping this updated ensures your financial wishes are honoured.</p>
    </div>
  `;
}

// ── Wizard navigation ───────────────────────────────────────
function wizardNext() {
  // Save any goal or child row inputs from the current step before moving
  saveCurrentWizardStepInputs();
  if (wizardCurrentStep < WIZARD_STEPS.length - 1) {
    wizardCurrentStep++;
    renderWizardStep(wizardCurrentStep);
  } else {
    wizardFinish();
  }
}

function wizardBack() {
  saveCurrentWizardStepInputs();
  if (wizardCurrentStep > 0) {
    wizardCurrentStep--;
    renderWizardStep(wizardCurrentStep);
  }
}

function saveCurrentWizardStepInputs() {
  const content = byId("wizardContent");
  if (!content) return;
  // Goal inputs
  content.querySelectorAll("input[data-goal-id]").forEach(el => {
    const g = goals.find(g => g.id === el.dataset.goalId);
    if (!g) return;
    const k = el.dataset.goalKey;
    g[k] = k === "years" || k === "amount" || k === "provision" ? Number(el.value||0) : el.value;
  });
  // Child inputs
  content.querySelectorAll("input[data-child-idx]").forEach(el => {
    const i = Number(el.dataset.childIdx);
    if (!children[i]) children[i] = { name: "", dob: "" };
    children[i][el.dataset.childKey] = el.value;
  });
  // Insurance inputs (wizard step 6 table)
  content.querySelectorAll("input[data-wz-ins-type]").forEach(el => {
    const t = el.dataset.wzInsType;
    const i = Number(el.dataset.wzInsIdx);
    const k = el.dataset.wzInsKey;
    const arrMap = { life: lifeInsuranceRows, health: healthInsuranceRows, car: carInsuranceRows, property: propertyInsuranceRows };
    const arr = arrMap[t];
    if (arr && arr[i]) {
      const numericKeys = ["sumAssured","annualPrem","surrenderVal","idv","cover"];
      arr[i][k] = numericKeys.includes(k) ? Number(el.value||0) : el.value;
    }
  });
  // Custom expenses (wizard step 3)
  content.querySelectorAll("[data-custom-exp-idx]").forEach(el => {
    const i = Number(el.dataset.customExpIdx);
    const k = el.dataset.customExpKey;
    if (!customExpenses[i]) customExpenses[i] = { name: "", amount: 0, note: "" };
    customExpenses[i][k] = k === "amount" ? Number(el.value||0) : el.value;
  });
  // Sync child[0] and child[1] DOBs to existing model fields for cashflow calculations
  if (children[0]) { model.child1Dob = children[0].dob || ""; }
  if (children[1]) { model.child2Dob = children[1].dob || ""; }
  scheduleAutosave();
  recalc();
}

// ── Wizard global helpers (called from inline HTML) ──────────
window.wzSetNumChildren = function(n) {
  while (children.length < n) children.push({ name: "", dob: "" });
  children.splice(n); // trim excess children
  // Remove any goals that belong to children beyond the new count
  for (let i = goals.length - 1; i >= 0; i--) {
    if (goals[i].childIdx !== undefined && goals[i].childIdx >= n) {
      goals.splice(i, 1);
    }
  }
  // Clear the DOB model fields for removed children slots
  if (n < 1) { model.child1Dob = ""; }
  if (n < 2) { model.child2Dob = ""; }
  recalc();
  renderWizardStep(wizardCurrentStep);
};

window.wzAddGoal = function() {
  goals.push({ id: `custom-${Date.now()}`, name: "New Goal", kind: "goal", inflationType: "general", years: 5, amount: 0, provision: 0 });
  renderWizardStep(wizardCurrentStep);
};

window.wzDelGoal = function(id) {
  const idx = goals.findIndex(g => g.id === id);
  if (idx !== -1) goals.splice(idx, 1);
  renderWizardStep(wizardCurrentStep);
};

window.wzAddIns = function(type) {
  if (type === "life")          lifeInsuranceRows.push({ policyName: "", company: "", sumAssured: 0, annualPrem: 0, surrenderVal: 0 });
  else if (type === "health")   healthInsuranceRows.push({ policyName: "", company: "", sumAssured: 0, annualPrem: 0 });
  else if (type === "car")      carInsuranceRows.push({ policyName: "", company: "", idv: 0, annualPrem: 0, expiry: "" });
  else if (type === "property") propertyInsuranceRows.push({ policyName: "", company: "", propertyName: "", cover: 0, annualPrem: 0 });
  renderWizardStep(wizardCurrentStep);
};

window.wzDelIns = function(type, idx) {
  const arrMap = { life: lifeInsuranceRows, health: healthInsuranceRows, car: carInsuranceRows, property: propertyInsuranceRows };
  if (arrMap[type]) arrMap[type].splice(idx, 1);
  renderWizardStep(wizardCurrentStep);
};

window.wzToggleFamilyHistory = function(val) {
  model.familyHistoryCriticalIllness = val;
  const descEl = byId("wzFamilyHistoryDesc");
  if (descEl) descEl.style.display = (val === "yes") ? "" : "none";
  // Show/hide the advisory info box by re-rendering only if it just changed to yes
  scheduleAutosave();
};

window.wzAddCustomExp = function() {
  customExpenses.push({ name: "", amount: 0, note: "" });
  renderWizardStep(wizardCurrentStep);
};

window.wzDelCustomExp = function(idx) {
  customExpenses.splice(idx, 1);
  renderWizardStep(wizardCurrentStep);
};

function initWizard() {
  byId("wizardNextBtn")?.addEventListener("click", wizardNext);
  byId("wizardBackBtn")?.addEventListener("click", wizardBack);
  byId("wizardCloseBtn")?.addEventListener("click", closeWizard);
  byId("openWizardBtn")?.addEventListener("click", openWizard);
  byId("mobWizardBtn")?.addEventListener("click", openWizard);
}

function recalc() {
  const monthlyInflow = model.incomeMain + model.incomeSpouse;
  { const _e11 = byId("age"); if (_e11) _e11.value = yearsBetween(model.dob, model.planDate); }
  { const _e12 = byId("spouseAge"); if (_e12) _e12.value = yearsBetween(model.spouseDob, model.planDate); }
  { const _e13 = byId("child1Age"); if (_e13) _e13.value = yearsBetween(model.child1Dob, model.planDate); }
  { const _e14 = byId("child2Age"); if (_e14) _e14.value = yearsBetween(model.child2Dob, model.planDate); }

  const customExpTotal = customExpenses.reduce((s, e) => s + (Number(e.amount)||0), 0);
  const monthlyOutflow =
    model.expHousehold +
    model.expLifestyle +
    model.expEducation +
    model.expVehicle +
    model.expMediclaim +
    model.expUtilities +
    model.expCarInsurance +
    model.expMisc +
    model.expLifeIns +
    model.expVacation +
    model.expRent +
    model.expCreditCard +
    model.expTravel +
    model.expProfFees +
    model.expPpfMonthly +
    customExpTotal;

  const goalOutput = computeGoalOutput();
  const goalSummary = renderGoalSheet(goalOutput);
  renderGoalPie(goalSummary.goalStrategyRows);
  const networthSummary = renderNetworth();
  renderRoiTable();
  const cfRows = computeCashflow(goalOutput, goalSummary.requiredSip, monthlyInflow, monthlyOutflow);
  renderCashflowTable(cfRows);
  renderCashflowChart(cfRows);
  // renderBreakup removed — Goal Sheet Breakup tab eliminated

  latestState.goalSummary = goalSummary;
  latestState.networth = networthSummary;
  latestState.cashflow = cfRows;
  renderAdminNetworthSheet();
  renderDashboard();
  updateInvTabTotals();
  scheduleAutosave();
}

function styleHeaderRow(row) {
  row.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC12800" } };
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
    cell.alignment = { vertical: "middle", horizontal: "center" };
  });
}

function styleGrid(ws) {
  ws.eachRow((row) => {
    row.eachCell((cell) => {
      if (!cell.border || !cell.border.top) {
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      }
      if (!cell.alignment) cell.alignment = { vertical: "middle" };
    });
  });
}

async function downloadWorkbook() {
  if (!window.ExcelJS) {
    alert("Excel export library not loaded. Please refresh and try again.");
    return;
  }
  const btn = byId("downloadExcelBtn");
  if (btn) btn.disabled = true;
  try {
    recalc();
    const wb = new window.ExcelJS.Workbook();
    wb.creator = "Goal Planner";
    wb.created = new Date();

    const goalData = latestState.goalSummary?.goalStrategyRows || [];
    const netData = latestState.networth?.rows || [];
    const cfData = latestState.cashflow || [];

    const wsGoal = wb.addWorksheet("Goal-Sheet");
    wsGoal.columns = [
      { header: "Sr.No.", key: "sr", width: 8 },
      { header: "Goal", key: "goal", width: 24 },
      { header: "Target yr.", key: "target", width: 12 },
      { header: "Yrs.", key: "yrs", width: 8 },
      { header: "Curr. Prov. (Rs.)", key: "prov", width: 18 },
      { header: "Gap (Rs.)", key: "gap", width: 18 },
      { header: "PM (Rs.)", key: "pm", width: 14 },
      { header: "PY (Rs.)", key: "py", width: 14 },
    ];
    styleHeaderRow(wsGoal.getRow(1));
    goalData.forEach((g, i) => {
      const rowNo = i + 2;
      wsGoal.addRow({
        sr: i + 1,
        goal: g.name,
        target: g.targetYear,
        yrs: g.years,
        prov: g.provision,
        gap: g.gap,
        pm: { formula: `IF(F${rowNo}<=0,0,PMT(${model.preRetRate / 100 / 12},D${rowNo}*12,0,-F${rowNo},1))` },
        py: { formula: `G${rowNo}*12` },
      });
    });
    wsGoal.addRow([]);
    const lessRow = wsGoal.addRow(["", "Less: Current SIPs", "", "", "", "", model.currentSipPm, model.currentSipPm * 12]);
    lessRow.font = { bold: true };
    const totalRow = wsGoal.addRow([
      "",
      "",
      "",
      "",
      "",
      "Total Investments Required",
      { formula: `SUM(G2:G${goalData.length + 1})-G${goalData.length + 3}` },
      { formula: `SUM(H2:H${goalData.length + 1})-H${goalData.length + 3}` },
    ]);
    totalRow.font = { bold: true };
    styleGrid(wsGoal);

    const wsNet = wb.addWorksheet("Networth Statement");
    wsNet.columns = [
      { header: "Particulars", key: "label", width: 24 },
      { header: "Amount", key: "amount", width: 16 },
      { header: "% of Assets", key: "pct", width: 14 },
    ];
    styleHeaderRow(wsNet.getRow(1));
    netData.forEach((r, i) => {
      const rowNo = i + 2;
      wsNet.addRow({
        label: r.label,
        amount: r.amount,
        pct: { formula: `IF(B${netData.length + 2}=0,0,B${rowNo}/B${netData.length + 2})` },
      });
    });
    const tAssetsRow = wsNet.addRow(["Total Assets", { formula: `SUM(B2:B${netData.length + 1})` }, 1]);
    tAssetsRow.font = { bold: true };
    const tLiabRow = wsNet.addRow(["Total Liabilities", latestState.networth?.totalLiabilities || 0, ""]);
    const nWRow = wsNet.addRow(["Net Worth", { formula: `B${netData.length + 2}-B${netData.length + 3}` }, ""]);
    tLiabRow.font = { bold: true };
    nWRow.font = { bold: true };
    styleGrid(wsNet);

    const wsCf = wb.addWorksheet("Cash Flow");
    wsCf.columns = [
      { header: "No.", key: "no", width: 8 },
      { header: "Year", key: "year", width: 10 },
      { header: "Age", key: "age", width: 8 },
      { header: "Op bal", key: "op", width: 16 },
      { header: "Cash In", key: "in", width: 16 },
      { header: "Lump Sum", key: "lumpSum", width: 16 },
      { header: "Growth", key: "growth", width: 10 },
      { header: "FV End", key: "fv", width: 16 },
      { header: "Cash Out", key: "out", width: 16 },
      { header: "Cl Bal", key: "cl", width: 16 },
      { header: "Goals", key: "goals", width: 28 },
    ];
    styleHeaderRow(wsCf.getRow(1));
    cfData.forEach((r) => {
      wsCf.addRow({
        no: r.no,
        year: r.year,
        age: r.age,
        op: r.opBal,
        in: r.cashIn,
        lumpSum: r.lumpSum || 0,
        growth: r.growth,
        fv: r.fvEnd,
        out: r.cashOut,
        cl: r.clBal,
        goals: r.goals,
      });
    });
    wsCf.getColumn("growth").numFmt = "0%";
    styleGrid(wsCf);

    // ── Questionnaire Data Sheet ───────────────────────────────────────────
    const wsQuest = wb.addWorksheet("Questionnaire Data");
    wsQuest.columns = [
      { header: "Field", key: "field", width: 30 },
      { header: "Value", key: "value", width: 40 },
    ];
    styleHeaderRow(wsQuest.getRow(1));

    const questData = [
      ["Investor Name", model.name || ""],
      ["Email", auth.currentUser?.email || ""],
      ["Plan Date", model.planDate || ""],
      ["DOB", model.dob || ""],
      ["Spouse DOB", model.spouseDob || ""],
      ["Child 1 DOB", model.child1Dob || ""],
      ["Child 2 DOB", model.child2Dob || ""],
      ["City", model.city || ""],
      ["State", model.state || ""],
      ["", ""],
      ["Income - Main", model.incomeMain || 0],
      ["Income - Spouse", model.incomeSpouse || 0],
      ["", ""],
      ["Expenses - Household", model.expHousehold || 0],
      ["Expenses - Lifestyle", model.expLifestyle || 0],
      ["Expenses - Education", model.expEducation || 0],
      ["Expenses - Vehicle", model.expVehicle || 0],
      ["Expenses - Mediclaim", model.expMediclaim || 0],
      ["Expenses - Utilities", model.expUtilities || 0],
      ["Expenses - Car Insurance", model.expCarInsurance || 0],
      ["Expenses - Misc", model.expMisc || 0],
      ["Expenses - Life Insurance", model.expLifeIns || 0],
      ["Expenses - Vacation", model.expVacation || 0],
      ["Expenses - Rent", model.expRent || 0],
      ["Expenses - Credit Card", model.expCreditCard || 0],
      ["Expenses - Travel", model.expTravel || 0],
      ["Expenses - Prof Fees", model.expProfFees || 0],
      ["Expenses - PPF Monthly", model.expPpfMonthly || 0],
      ["", ""],
      ["Assets - Home", model.assetHome || 0],
      ["Assets - Car", model.assetCar || 0],
      ["Assets - Gold", model.assetGold || 0],
      ["", ""],
      ["Investments - Liquid MF", model.invLiquidMf || 0],
      ["Investments - Savings", model.invSavings || 0],
      ["Investments - Shares", model.invShares || 0],
      ["Investments - Equity MF", model.invEquityMf || 0],
      ["Investments - Debt MF", model.invDebtMf || 0],
      ["Investments - Bonds", model.invBonds || 0],
      ["Investments - Postal", model.invPostal || 0],
      ["Investments - PPF", model.invPpf || 0],
      ["Investments - ULIP", model.invUlip || 0],
      ["Investments - EPF", model.invEpf || 0],
      ["Investments - ELSS", model.invElss || 0],
      ["Investments - Bank FD", model.invBankFd || 0],
      ["Investments - Cash", model.invCash || 0],
      ["", ""],
      ["Liabilities - Home Loan", model.loanHome || 0],
      ["Liabilities - Car Loan", model.loanCar || 0],
      ["Liabilities - Other Loan", model.loanOther || 0],
      ["", ""],
      ["Inflation Rate", `${(model.inflationRate * 100 || 0).toFixed(2)}%`],
      ["Education Inflation Rate", `${(model.educationInflationRate * 100 || 0).toFixed(2)}%`],
      ["Marriage Inflation Rate", `${(model.marriageInflationRate * 100 || 0).toFixed(2)}%`],
      ["Pre-Retirement Rate", `${(model.preRetRate * 100 || 0).toFixed(2)}%`],
      ["Post-Retirement Rate", `${(model.postRetRate * 100 || 0).toFixed(2)}%`],
      ["Cash In Growth Rate", `${(model.cashInGrowthRate * 100 || 0).toFixed(2)}%`],
      ["Debt Rate", `${(model.debtRate * 100 || 0).toFixed(2)}%`],
      ["", ""],
      ["Retirement Age", model.retirementAge || 0],
      ["Life Expectancy", model.lifeExpectancy || 0],
      ["Retirement Monthly Expense", model.retirementMonthlyExp || 0],
      ["Current SIP (Monthly)", model.currentSipPm || 0],
      ["", ""],
      ["Will Status", model.willStatus || "Not provided"],
      ["Nominations Updated", model.nominationsUpdated || "Not provided"],
      ["Net Worth Notes", model.networthNotes || ""],
    ];

    questData.forEach(([field, value]) => {
      wsQuest.addRow({ field, value });
    });
    styleGrid(wsQuest);

    const now = new Date();
    const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(
      2,
      "0"
    )}`;
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Goal_Planner_${stamp}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  } finally {
    if (btn) btn.disabled = false;
  }
}

async function downloadPDF() {
  if (!window.jspdf || !window.html2canvas) {
    alert("PDF export libraries not loaded. Please refresh and try again.");
    return;
  }

  const btn = byId("downloadPdfBtn");
  if (btn) btn.disabled = true;

  try {
    recalc();
    const { jsPDF } = window.jspdf;
    const html2canvas = window.html2canvas;
    const doc = new jsPDF("p", "mm", "a4");
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const margin = 18;
    const contentWidth = pageWidth - 2 * margin;
    const brandOrange = [193, 40, 0];
    const brandDark = [40, 40, 40];
    const lightGray = [250, 250, 250];
    const darkGray = [100, 100, 100];

    // ── Helper: Draw section header bar ──────────────────────────────────
    const drawSectionHeader = (title, yPos) => {
      doc.setFillColor(...brandOrange);
      doc.rect(margin, yPos - 6, contentWidth, 10, "F");
      doc.setFontSize(16);
      doc.setTextColor(255, 255, 255);
      doc.setFont(undefined, "bold");
      doc.text(title, margin + 5, yPos);
      doc.setFont(undefined, "normal");
      doc.setTextColor(0, 0, 0); // Reset text color to avoid bleed
      return yPos + 12;
    };

    // ── Page 1: Title & Summary ────────────────────────────────────────────
    let yPos = 18;

    // Brand header bar
    doc.setFillColor(...brandOrange);
    doc.rect(0, 0, pageWidth, 16, "F");
    doc.setFontSize(20);
    doc.setTextColor(255, 255, 255);
    doc.setFont(undefined, "bold");
    doc.text("GOAL PLANNER REPORT", pageWidth / 2, 11, { align: "center" });
    doc.setFont(undefined, "normal");

    yPos = 28;

    // Investor info
    doc.setFontSize(11);
    doc.setTextColor(...brandDark);
    doc.setFont(undefined, "bold");
    doc.text("Investor Information", margin, yPos);
    doc.setFont(undefined, "normal");
    yPos += 7;

    doc.setFontSize(10);
    doc.setTextColor(80, 80, 80);
    doc.text(`Name: ${model.name || "N/A"}`, margin, yPos);
    yPos += 5;
    doc.text(`Date: ${new Date().toLocaleDateString("en-IN", { year: "numeric", month: "long", day: "numeric" })}`, margin, yPos);
    yPos += 12;

    // Financial Summary - Key Metrics
    const summaryData = [
      { label: "Total Assets", value: latestState.networth?.totalAssets || 0 },
      { label: "Total Liabilities", value: latestState.networth?.totalLiabilities || 0 },
      { label: "Net Worth", value: latestState.networth?.netWorth || 0 },
      { label: "Goal Corpus Required", value: latestState.goalSummary?.totalGoalCorpus || 0 },
      { label: "Monthly SIP Required", value: latestState.goalSummary?.requiredSip || 0 },
    ];

    doc.setFillColor(...lightGray);
    doc.rect(margin, yPos - 1, contentWidth, 55, "F");

    doc.setFontSize(11);
    doc.setTextColor(...brandDark);
    doc.setFont(undefined, "bold");
    doc.text("Financial Summary", margin + 4, yPos + 4);
    doc.setFont(undefined, "normal");

    yPos += 10;
    doc.setFontSize(9);
    doc.setTextColor(80, 80, 80);

    summaryData.forEach((item) => {
      doc.text(item.label, margin + 4, yPos);
      doc.setFont(undefined, "bold");
      doc.text(formatRs(item.value), pageWidth - margin - 4, yPos, { align: "right" });
      doc.setFont(undefined, "normal");
      yPos += 8;
    });

    // ── Page 2: Goal-Sheet Table ───────────────────────────────────────────
    doc.addPage();
    yPos = 15;
    yPos = drawSectionHeader("Goal-Sheet: Achievement Strategy", yPos);

    const goalRows = latestState.goalSummary?.goalStrategyRows || [];
    const goalHeaders = ["No", "Goal Name", "Target Yr", "Years", "Provision", "Gap", "Monthly", "Yearly"];
    const goalColWidths = [8, 40, 16, 12, 25, 25, 22, 22];

    // Table header
    doc.setFontSize(9);
    doc.setTextColor(255, 255, 255);
    doc.setFont(undefined, "bold");
    doc.setFillColor(...brandOrange);
    let xPos = margin;
    goalHeaders.forEach((header, idx) => {
      doc.rect(xPos, yPos - 5, goalColWidths[idx], 7, "F");
      const textX = xPos + goalColWidths[idx] / 2;
      doc.text(header, textX, yPos - 1.5, { align: "center", fontSize: 8 });
      xPos += goalColWidths[idx];
    });
    yPos += 4;
    doc.setFont(undefined, "normal");

    // Table rows
    doc.setTextColor(60, 60, 60);
    let rowAlt = false;
    goalRows.forEach((goal, idx) => {
      if (yPos > pageHeight - 18) {
        doc.addPage();
        yPos = 15;
      }

      const rowData = [
        String(idx + 1),
        goal.name,
        String(goal.targetYear),
        String(goal.years),
        formatRsCompact(goal.provision),
        formatRsCompact(goal.gap),
        formatRsCompact(goal.pm),
        formatRsCompact(goal.py),
      ];

      // Alternating row background
      if (rowAlt) {
        doc.setFillColor(248, 248, 248);
        xPos = margin;
        goalColWidths.forEach((w) => {
          doc.rect(xPos, yPos - 4.5, w, 6, "F");
          xPos += w;
        });
      }

      // Draw borders
      xPos = margin;
      doc.setDrawColor(220, 220, 220);
      doc.setLineWidth(0.2);
      goalColWidths.forEach((w) => {
        doc.rect(xPos, yPos - 4.5, w, 6);
        xPos += w;
      });

      xPos = margin;
      rowData.forEach((cell, idx) => {
        const align = (idx === 1) ? "left" : "right";
        const offset = (idx === 1) ? 1 : goalColWidths[idx] - 2;
        doc.setFontSize(8);
        doc.text(String(cell), xPos + offset, yPos - 1, { align });
        xPos += goalColWidths[idx];
      });
      yPos += 6;
      rowAlt = !rowAlt;
    });

    yPos += 3;
    doc.setFontSize(9);
    doc.setTextColor(...brandDark);
    doc.setFont(undefined, "bold");
    doc.text(`Total Goal Corpus: ${formatRs(latestState.goalSummary?.totalGoalCorpus || 0)}`, margin, yPos);
    yPos += 5;
    doc.text(`Required Monthly SIP: ${formatRs(latestState.goalSummary?.requiredSip || 0)}`, margin, yPos);
    doc.setFont(undefined, "normal");

    // Goal Pie Chart — drawn natively on canvas
    yPos += 8;
    if (yPos > pageHeight - 70) { doc.addPage(); yPos = 15; }
    const goalPieData = (latestState.goalSummary?.goalStrategyRows || [])
      .filter(g => g.py > 0)
      .map(g => ({ name: g.name, value: g.py }));
    if (goalPieData.length) {
      doc.setFontSize(11);
      doc.setTextColor(...brandDark);
      doc.setFont(undefined, "bold");
      doc.text("Goal Allocation by Annual SIP", margin, yPos);
      doc.setFont(undefined, "normal");
      yPos += 6;
      const pieImg = pdfDrawPie(goalPieData, 480, 200);
      if (pieImg) {
        doc.addImage(pieImg, "PNG", margin, yPos, contentWidth, 80);
        yPos += 85;
      }
    }

    // ── Page 3: Networth Statement ─────────────────────────────────────────
    doc.addPage();
    yPos = 15;
    yPos = drawSectionHeader("Net Worth Statement", yPos);

    const networthRows = latestState.networth?.rows || [];
    const networthHeaders = ["Asset Class", "Value", "% of Total"];
    const networthColWidths = [65, 35, 35];

    // Table header
    doc.setFontSize(9);
    doc.setTextColor(255, 255, 255);
    doc.setFont(undefined, "bold");
    doc.setFillColor(...brandOrange);
    xPos = margin;
    networthHeaders.forEach((header, idx) => {
      doc.rect(xPos, yPos - 5, networthColWidths[idx], 7, "F");
      const textX = xPos + (idx === 0 ? 2 : networthColWidths[idx] / 2);
      const align = idx === 0 ? "left" : "center";
      doc.text(header, textX, yPos - 1.5, { align, fontSize: 8 });
      xPos += networthColWidths[idx];
    });
    yPos += 4;
    doc.setFont(undefined, "normal");

    // Table rows
    doc.setTextColor(60, 60, 60);
    rowAlt = false;
    networthRows.forEach((row) => {
      if (yPos > pageHeight - 18) {
        doc.addPage();
        yPos = 15;
      }

      const totalAssets = latestState.networth?.totalAssets || 1;
      const pct = totalAssets ? Math.round((row.amount / totalAssets) * 100) : 0;
      const rowData = [row.label, formatRsCompact(row.amount), `${pct}%`];

      // Alternating row background
      if (rowAlt) {
        doc.setFillColor(248, 248, 248);
        xPos = margin;
        networthColWidths.forEach((w) => {
          doc.rect(xPos, yPos - 4.5, w, 6, "F");
          xPos += w;
        });
      }

      // Draw borders
      xPos = margin;
      doc.setDrawColor(220, 220, 220);
      doc.setLineWidth(0.2);
      networthColWidths.forEach((w) => {
        doc.rect(xPos, yPos - 4.5, w, 6);
        xPos += w;
      });

      xPos = margin;
      doc.setFontSize(8);
      doc.text(rowData[0], xPos + 2, yPos - 1, { align: "left" });
      doc.text(rowData[1], xPos + networthColWidths[0] + 30, yPos - 1, { align: "right" });
      doc.text(rowData[2], xPos + networthColWidths[0] + networthColWidths[1] + 30, yPos - 1, { align: "right" });
      yPos += 6;
      rowAlt = !rowAlt;
    });

    // Summary boxes
    yPos += 5;
    const boxWidth = (contentWidth - 4) / 3;
    const summaryBoxes = [
      { label: "Total Assets", value: latestState.networth?.totalAssets || 0 },
      { label: "Total Liabilities", value: latestState.networth?.totalLiabilities || 0 },
      { label: "Net Worth", value: latestState.networth?.netWorth || 0 },
    ];

    summaryBoxes.forEach((box, idx) => {
      const xStart = margin + idx * (boxWidth + 2);
      doc.setDrawColor(...brandOrange);
      doc.setLineWidth(0.5);
      doc.rect(xStart, yPos, boxWidth, 14);
      doc.setFillColor(255, 255, 255);

      doc.setFontSize(8);
      doc.setTextColor(...darkGray);
      doc.text(box.label, xStart + boxWidth / 2, yPos + 4, { align: "center" });

      doc.setFontSize(10);
      doc.setFont(undefined, "bold");
      doc.setTextColor(...brandOrange);
      doc.text(formatRsCompact(box.value), xStart + boxWidth / 2, yPos + 10, { align: "center" });
      doc.setFont(undefined, "normal");
    });

    // Networth Pie Chart — drawn natively on canvas
    yPos += 20;
    if (yPos > pageHeight - 90) { doc.addPage(); yPos = 15; }
    const nwPieData = (latestState.networth?.rows || [])
      .filter(r => r.amount > 0)
      .map(r => ({ name: r.label, value: r.amount }));
    if (nwPieData.length) {
      doc.setFontSize(11);
      doc.setTextColor(...brandDark);
      doc.setFont(undefined, "bold");
      doc.text("Asset Allocation Breakdown", margin, yPos);
      doc.setFont(undefined, "normal");
      yPos += 6;
      const nwPieImg = pdfDrawPie(nwPieData, 480, 200);
      if (nwPieImg) {
        doc.addImage(nwPieImg, "PNG", margin, yPos, contentWidth, 80);
        yPos += 85;
      }
    }

    // ── Page 4: Cash Flow Projection ───────────────────────────────────────
    doc.addPage();
    yPos = 15;
    yPos = drawSectionHeader("Cash Flow Projection", yPos);

    const cfRows = latestState.cashflow || [];
    const cfHeaders = ["Year", "Age", "Opening Bal", "Cash In", "Lump Sum", "Growth %", "FV End", "Cash Out", "Closing Bal"];
    const cfColWidths = [14, 11, 20, 20, 13, 16, 20, 20, 19];  // 9 items — one per header

    // Table header
    doc.setFontSize(8);
    doc.setTextColor(255, 255, 255);
    doc.setFont(undefined, "bold");
    doc.setFillColor(...brandOrange);
    xPos = margin;
    cfHeaders.forEach((header, idx) => {
      doc.rect(xPos, yPos - 5, cfColWidths[idx], 7, "F");
      const textX = xPos + cfColWidths[idx] / 2;
      doc.text(header, textX, yPos - 1.5, { align: "center", fontSize: 7 });
      xPos += cfColWidths[idx];
    });
    yPos += 4;
    doc.setFont(undefined, "normal");

    // Table rows - show first 12 years
    doc.setTextColor(60, 60, 60);
    rowAlt = false;
    const cfDisplay = cfRows.slice(0, 12);
    cfDisplay.forEach((cf) => {
      const rowData = [
        String(cf.year),
        String(cf.age),
        formatRsCompact(cf.opBal),
        formatRsCompact(cf.cashIn),
        formatRsCompact(cf.lumpSum || 0),
        `${(cf.growth * 100).toFixed(1)}%`,
        formatRsCompact(cf.fvEnd),
        formatRsCompact(cf.cashOut),
        formatRsCompact(cf.clBal),
      ];

      // Alternating row background
      if (rowAlt) {
        doc.setFillColor(248, 248, 248);
        xPos = margin;
        cfColWidths.forEach((w) => {
          doc.rect(xPos, yPos - 4, w, 5.5, "F");
          xPos += w;
        });
      }

      // Draw borders
      xPos = margin;
      doc.setDrawColor(220, 220, 220);
      doc.setLineWidth(0.2);
      cfColWidths.forEach((w) => {
        doc.rect(xPos, yPos - 4, w, 5.5);
        xPos += w;
      });

      xPos = margin;
      doc.setFontSize(7);
      rowData.forEach((cell, idx) => {
        const align = idx === 0 || idx === 1 ? "center" : "right";
        const offset = idx === 0 || idx === 1 ? cfColWidths[idx] / 2 : cfColWidths[idx] - 1;
        doc.text(String(cell), xPos + offset, yPos - 1, { align });
        xPos += cfColWidths[idx];
      });
      yPos += 5.5;
      rowAlt = !rowAlt;
    });

    if (cfRows.length > 12) {
      yPos += 2;
      doc.setFontSize(8);
      doc.setTextColor(...darkGray);
      doc.text(`... projection continues for ${cfRows.length - 12} more years`, margin, yPos);
    }

    // ── Cash Flow Line Chart — drawn natively on canvas ────────────────────
    yPos += 10;
    if (yPos > pageHeight - 110) { doc.addPage(); yPos = 15; }
    if (cfRows.length > 1) {
      yPos = drawSectionHeader("Cash Flow Visualization", yPos);
      const cfChartImg = pdfDrawLineChart(cfRows, 700, 280);
      if (cfChartImg) {
        doc.addImage(cfChartImg, "PNG", margin, yPos, contentWidth, 110);
        yPos += 115;
      }
    }

    // Footer
    doc.setFontSize(8);
    doc.setTextColor(150, 150, 150);
    doc.text("Arthashastra Investments | Confidential", pageWidth / 2, pageHeight - 5, { align: "center" });

    // Save PDF
    const now = new Date();
    const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}`;
    doc.save(`Goal_Planner_Report_${stamp}.pdf`);

  } catch (err) {
    console.error("PDF export error:", err);
    alert("Error generating PDF. Please try again.");
  } finally {
    if (btn) btn.disabled = false;
  }
}

// Compact format for PDF tables (Cr, L, K, etc.)
function formatRsCompact(val) {
  const n = Math.abs(Number(val) || 0);
  if (n >= 10000000) return (val / 10000000).toFixed(1).replace(/\.0$/, "") + " Cr";
  if (n >= 100000) return (val / 100000).toFixed(1).replace(/\.0$/, "") + " L";
  if (n >= 1000) return (val / 1000).toFixed(0) + "K";
  return Math.round(val).toString();
}

// ── PDF native chart renderers ─────────────────────────────────────────────

const PDF_COLORS = [
  "#3c78d8","#cc4125","#6aa84f","#e69138","#674ea7",
  "#45818e","#f1c232","#a64d79","#32a7c7","#4a86e8",
  "#8e7cc3","#91c33b",
];

// Draw a pie chart + legend to an offscreen canvas, return PNG data URL
function pdfDrawPie(data, canvasW, canvasH) {
  try {
    const canvas = document.createElement("canvas");
    canvas.width = canvasW;
    canvas.height = canvasH;
    const ctx = canvas.getContext("2d");

    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvasW, canvasH);

    const total = data.reduce((s, d) => s + d.value, 0);
    if (!total) return null;

    const cx = canvasW * 0.32;
    const cy = canvasH / 2;
    const r  = Math.min(cx, cy) * 0.82;

    // Draw slices
    let angle = -Math.PI / 2;
    data.forEach((d, i) => {
      const sweep = (d.value / total) * 2 * Math.PI;
      ctx.beginPath();
      ctx.moveTo(cx, cy);
      ctx.arc(cx, cy, r, angle, angle + sweep);
      ctx.closePath();
      ctx.fillStyle = PDF_COLORS[i % PDF_COLORS.length];
      ctx.fill();
      ctx.strokeStyle = "#ffffff";
      ctx.lineWidth = 2;
      ctx.stroke();

      // Percentage label inside slice if large enough
      const pct = d.value / total;
      if (pct > 0.07) {
        const mid = angle + sweep / 2;
        const lx = cx + (r * 0.62) * Math.cos(mid);
        const ly = cy + (r * 0.62) * Math.sin(mid);
        ctx.fillStyle = "#ffffff";
        ctx.font = "bold 13px Arial";
        ctx.textAlign = "center";
        ctx.textBaseline = "middle";
        ctx.fillText(`${Math.round(pct * 100)}%`, lx, ly);
      }
      angle += sweep;
    });

    // Legend
    const legX = canvasW * 0.67;
    const lineH = Math.min(22, (canvasH - 16) / Math.min(data.length, 10));
    let legY = (canvasH - lineH * Math.min(data.length, 10)) / 2 + lineH * 0.7;

    ctx.textAlign = "left";
    ctx.textBaseline = "middle";
    data.slice(0, 10).forEach((d, i) => {
      const pct = Math.round((d.value / total) * 100);
      ctx.fillStyle = PDF_COLORS[i % PDF_COLORS.length];
      ctx.fillRect(legX, legY - 7, 14, 14);
      ctx.fillStyle = "#333333";
      const fSize = Math.max(9, Math.min(12, lineH - 3));
      ctx.font = `${fSize}px Arial`;
      const label = d.name.length > 18 ? d.name.slice(0, 16) + "…" : d.name;
      ctx.fillText(`${label}  ${pct}%`, legX + 19, legY);
      legY += lineH;
    });

    return canvas.toDataURL("image/png");
  } catch (e) {
    console.error("pdfDrawPie error:", e);
    return null;
  }
}

// Draw a line/area chart of closing balance over years, return PNG data URL
function pdfDrawLineChart(cfRows, canvasW, canvasH) {
  try {
    const canvas = document.createElement("canvas");
    canvas.width = canvasW;
    canvas.height = canvasH;
    const ctx = canvas.getContext("2d");

    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvasW, canvasH);

    const padL = 80, padR = 20, padT = 20, padB = 50;
    const cW = canvasW - padL - padR;
    const cH = canvasH - padT - padB;
    const vals = cfRows.map(r => r.clBal);
    const minV = Math.min(0, ...vals);
    const maxV = Math.max(...vals);
    const range = maxV - minV || 1;

    const toX = (i) => padL + (i / (cfRows.length - 1)) * cW;
    const toY = (v) => padT + cH - ((v - minV) / range) * cH;

    // Horizontal grid lines
    ctx.lineWidth = 0.8;
    for (let i = 0; i <= 5; i++) {
      const v = minV + (i / 5) * range;
      const y = toY(v);
      ctx.strokeStyle = "#eeeeee";
      ctx.beginPath(); ctx.moveTo(padL, y); ctx.lineTo(padL + cW, y); ctx.stroke();
      ctx.fillStyle = "#888";
      ctx.font = "11px Arial";
      ctx.textAlign = "right";
      ctx.textBaseline = "middle";
      ctx.fillText(formatRsCompact(v), padL - 6, y);
    }

    // Zero line if needed
    if (minV < 0) {
      const y0 = toY(0);
      ctx.strokeStyle = "#cccccc";
      ctx.lineWidth = 1.5;
      ctx.setLineDash([4, 3]);
      ctx.beginPath(); ctx.moveTo(padL, y0); ctx.lineTo(padL + cW, y0); ctx.stroke();
      ctx.setLineDash([]);
    }

    // Area fill
    ctx.beginPath();
    ctx.moveTo(toX(0), toY(vals[0]));
    vals.forEach((v, i) => ctx.lineTo(toX(i), toY(v)));
    ctx.lineTo(toX(vals.length - 1), toY(minV));
    ctx.lineTo(toX(0), toY(minV));
    ctx.closePath();
    ctx.fillStyle = "rgba(193,40,0,0.08)";
    ctx.fill();

    // Line
    ctx.beginPath();
    ctx.moveTo(toX(0), toY(vals[0]));
    vals.forEach((v, i) => ctx.lineTo(toX(i), toY(v)));
    ctx.strokeStyle = "#c12800";
    ctx.lineWidth = 2.5;
    ctx.lineJoin = "round";
    ctx.stroke();

    // X-axis year labels
    ctx.fillStyle = "#555";
    ctx.font = "11px Arial";
    ctx.textAlign = "center";
    ctx.textBaseline = "top";
    cfRows.forEach((r, i) => {
      if (i % Math.ceil(cfRows.length / 10) === 0 || i === cfRows.length - 1) {
        ctx.fillText(r.year, toX(i), padT + cH + 8);
        // Tick mark
        ctx.strokeStyle = "#cccccc";
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(toX(i), padT + cH);
        ctx.lineTo(toX(i), padT + cH + 5);
        ctx.stroke();
      }
    });

    // Axes
    ctx.strokeStyle = "#aaaaaa";
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    ctx.moveTo(padL, padT);
    ctx.lineTo(padL, padT + cH);
    ctx.lineTo(padL + cW, padT + cH);
    ctx.stroke();

    // Chart title
    ctx.fillStyle = "#333";
    ctx.font = "bold 13px Arial";
    ctx.textAlign = "center";
    ctx.textBaseline = "alphabetic";
    ctx.fillText("Portfolio Closing Balance Projection (₹)", padL + cW / 2, padT - 4);

    return canvas.toDataURL("image/png");
  } catch (e) {
    console.error("pdfDrawLineChart error:", e);
    return null;
  }
}

function initExportButton() {
  const btn = byId("downloadExcelBtn");
  if (!btn) return;
  btn.addEventListener("click", downloadWorkbook);

  const pdfBtn = byId("downloadPdfBtn");
  if (pdfBtn) {
    pdfBtn.addEventListener("click", downloadPDF);
  }

  // Mobile quick-action bar mirrors
  byId("mobExcelBtn")?.addEventListener("click", downloadWorkbook);
  byId("mobPdfBtn")?.addEventListener("click", downloadPDF);
}

function initTabs() {
  const tabs = document.querySelectorAll("#sheetTabs button");
  const sheets = document.querySelectorAll(".sheet");
  tabs.forEach((btn) => {
    btn.addEventListener("click", () => {
      tabs.forEach((t) => t.classList.remove("active"));
      sheets.forEach((s) => s.classList.remove("active"));
      btn.classList.add("active");
      byId(`sheet-${btn.dataset.sheet}`).classList.add("active");
    });
  });

  // ── Inner investment tabs (Net Worth section) ───────────────
  const invBar = byId("invTabBar");
  if (invBar) {
    invBar.addEventListener("click", (e) => {
      const btn = e.target.closest(".inv-tab-btn");
      if (!btn) return;
      const tab = btn.dataset.invTab;
      // Switch active button
      invBar.querySelectorAll(".inv-tab-btn").forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      // Switch active pane
      document.querySelectorAll(".inv-tab-pane").forEach(p => p.classList.remove("active"));
      const pane = byId(`inv-tab-${tab}`);
      if (pane) pane.classList.add("active");
    });
  }
}

/* Update investment tab totals — called after recalc */
function updateInvTabTotals() {
  const fmt = (v) => "₹ " + Math.round(v || 0).toLocaleString("en-IN");
  const marketTotal = (model.invLiquidMf||0) + (model.invEquityMf||0) + (model.invDebtMf||0)
    + (model.invElss||0) + (model.invUlip||0) + (model.invShares||0);
  const savingsTotal = (model.invSavings||0) + (model.invPpf||0) + (model.invEpf||0);
  const bondsTotal   = (model.invBankFd||0) + (model.invBonds||0) + (model.invPostal||0);
  const cashTotal    = (model.invCash||0);
  const el = (id) => byId(id);
  if (el("inv-total-market"))  el("inv-total-market").textContent  = fmt(marketTotal);
  if (el("inv-total-savings")) el("inv-total-savings").textContent = fmt(savingsTotal);
  if (el("inv-total-bonds"))   el("inv-total-bonds").textContent   = fmt(bondsTotal);
  if (el("inv-total-cash"))    el("inv-total-cash").textContent    = fmt(cashTotal);
}

[
  "name",
  "planDate",
  "dob",
  "city",
  "state",
  "spouseDob",
  "child1Dob",
  "child2Dob",
  "inflationRate",
  "educationInflationRate",
  "marriageInflationRate",
  "preRetRate",
  "postRetRate",
  "cashInGrowthRate",
  "retirementAge",
  "lifeExpectancy",
  "retirementMonthlyExp",
  "debtRate",
  "incomeMain",
  "incomeSpouse",
  "expHousehold",
  "expLifestyle",
  "expEducation",
  "expVehicle",
  "expMediclaim",
  "expUtilities",
  "expCarInsurance",
  "expMisc",
  "expLifeIns",
  "expVacation",
  "expRent",
  "expCreditCard",
  "expTravel",
  "expProfFees",
  "expPpfMonthly",
  "assetHome",
  "assetCar",
  "assetGold",
  "invLiquidMf",
  "invSavings",
  "invShares",
  "invEquityMf",
  "invDebtMf",
  "invBonds",
  "invPostal",
  "invPpf",
  "invEpf",
  "invElss",
  "invUlip",
  "loanHome",
  "loanCar",
  "loanOther",
  "currentSipPm",
].forEach(bindInput);

bindStaticUiEvents();
renderGoalInputRows();
renderPropertyRows();
initTabs();
initExportButton();
initWizard();
applyRoleVisibility();
recalc();
setAppLocked(true);
initFirebase().catch((e) => setStatus(`Firebase init failed: ${e.message}`));
