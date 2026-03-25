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
  cashflowOverrides = (planData.cashflowOverrides && typeof planData.cashflowOverrides === "object") ? planData.cashflowOverrides : {};
  children = Array.isArray(planData.children) ? planData.children : [];

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
  // Role is always determined by email — admin email gets admin, everyone else gets investor.
  const role = isAdminEmail(email) ? "admin" : "investor";
  if (!email || !password) return alert("Enter email and password.");
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
    updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
  });
}

async function login() {
  if (!auth) return;
  const email = byId("authEmail").value.trim();
  const password = byId("authPassword").value;
  if (!email || !password) return alert("Enter email and password.");
  await auth.signInWithEmailAndPassword(email, password);
  // Role is determined from Firestore in onAuthStateChanged — no blocking check needed here.
}

async function logout() {
  if (auth) await auth.signOut();
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
    updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
  };
  await db.collection("investorPlans").doc(currentPlanId).set(payload, { merge: true });
  setStatus(`Saved at ${new Date().toLocaleTimeString()}`);
}

function scheduleAutosave() {
  if (!currentUser || !currentPlanId || isHydrating) return;
  clearTimeout(autosaveTimer);
  autosaveTimer = setTimeout(() => {
    saveCurrentPlan().catch((e) => setStatus(`Save failed: ${e.message}`));
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
  const lockTargets = ["downloadExcelBtn", "savePlanBtn", "logoutBtn"];
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
  byId("savePlanBtn")?.addEventListener("click", () => saveCurrentPlan().catch((e) => setStatus(e.message)));
  byId("logoutBtn")?.addEventListener("click", () => logout().catch((e) => setStatus(e.message)));
  byId("loginBtn")?.addEventListener("click", () => login().catch((e) => setStatus(e.message)));
  byId("signupBtn")?.addEventListener("click", () => signup().catch((e) => setStatus(e.message)));
  byId("investorSelect")?.addEventListener("change", async (e) => {
    await loadPlan(e.target.value);
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
      // data-cf-post-ret="1" is stamped at render time for every row
      // whose age > retirementAge — read it directly to avoid any
      // re-derivation that could fail when dob/planDate are missing.
      const isPostRet = inp.dataset.cfPostRet === "1";
      if (isPostRet) {
        model.postRetRate = val;
        const syncEl = byId("postRetRate");
        if (syncEl) syncEl.value = val;
      } else {
        model.preRetRate = val;
        const syncEl = byId("preRetRate");
        if (syncEl) syncEl.value = val;
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
    model.invBonds +
    model.invPostal +
    model.invUlip;
  const annualSurplus = Math.max(0, (monthlyInflow - monthlyOutflow) * 12);
  // Keep Goal-Sheet linkage but ensure Input inflow/outflow changes are reflected immediately.
  let cashIn = Math.max(requiredSip * 12, annualSurplus);
  const rows = [];

  for (let i = 0; i <= years; i += 1) {
    const year = startYear + i;
    const age = currentAge + i;
    const ov = cashflowOverrides[year] || {};
    const growth = age < model.retirementAge ? model.preRetRate / 100 : model.postRetRate / 100;
    const baseCashIn = age <= model.retirementAge ? cashIn : 0;
    const effectiveCashIn = (ov.cashIn !== undefined) ? ov.cashIn : baseCashIn;
    const fvEnd = opening * (1 + growth) + effectiveCashIn;
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
      growth,
      fvEnd,
      cashOut,
      clBal,
      goals: goalText,
    });

    opening = clBal;
    cashIn *= 1 + model.cashInGrowthRate / 100;
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
  // Retirement current cost excludes children's education expense.
  const retirementBaseOutflow =
    model.expHousehold +
    model.expLifestyle +
    model.expVehicle +
    model.expMediclaim +
    model.expUtilities +
    model.expCarInsurance +
    model.expMisc;
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
    model.invBonds +
    model.invPostal +
    model.invPpf +
    model.invUlip;
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
  byId("goalTargetTotal").textContent = formatRs(totalGoalCorpus);
  byId("lessCurrentSipPm").textContent = formatRs(model.currentSipPm);
  byId("lessCurrentSipPy").textContent = formatRs(model.currentSipPm * 12);
  byId("requiredSip").textContent = formatRs(requiredSip);
  byId("requiredSipYearly").textContent = formatRs(requiredSip * 12);

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

  const cx = 280;
  const cy = 165;
  const r = 120;
  let startAngle = -Math.PI / 2;
  let slices = "";
  let labels = "";

  data.forEach((item, idx) => {
    const frac = item.value / total;
    const endAngle = startAngle + frac * Math.PI * 2;
    const x1 = cx + r * Math.cos(startAngle);
    const y1 = cy + r * Math.sin(startAngle);
    const x2 = cx + r * Math.cos(endAngle);
    const y2 = cy + r * Math.sin(endAngle);
    const largeArc = frac > 0.5 ? 1 : 0;
    const color = colors[idx % colors.length];
    slices += `<path d="M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${largeArc} 1 ${x2} ${y2} Z" fill="${color}" stroke="#ffffff" stroke-width="1.2"></path>`;

    const mid = startAngle + (endAngle - startAngle) / 2;
    const lineStartX = cx + (r - 6) * Math.cos(mid);
    const lineStartY = cy + (r - 6) * Math.sin(mid);
    const lineMidX = cx + (r + 18) * Math.cos(mid);
    const lineMidY = cy + (r + 18) * Math.sin(mid);
    const onRight = Math.cos(mid) >= 0;
    const lineEndX = lineMidX + (onRight ? 18 : -18);
    const lineEndY = lineMidY;
    const share = Math.round(frac * 100);
    labels += `<path d="M ${lineStartX} ${lineStartY} L ${lineMidX} ${lineMidY} L ${lineEndX} ${lineEndY}" fill="none" stroke="#333" stroke-width="1"></path>`;
    labels += `<text x="${lineEndX + (onRight ? 4 : -4)}" y="${lineEndY - 3}" font-size="12" text-anchor="${onRight ? "start" : "end"}">${escHtml(item.name)}</text>`;
    labels += `<text x="${lineEndX + (onRight ? 4 : -4)}" y="${lineEndY + 12}" font-size="12" text-anchor="${onRight ? "start" : "end"}">${share}%</text>`;
    startAngle = endAngle;
  });

  svg.innerHTML = `
    <rect width="720" height="340" fill="#d0d0d0"></rect>
    ${slices}
    ${labels}
  `;

  legend.innerHTML = "";
  data.forEach((d, i) => {
    const item = document.createElement("p");
    item.innerHTML = `
      <span style="display:inline-block;width:10px;height:10px;background:${colors[i % colors.length]};margin-right:6px;"></span>
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
    { label: "Bonds", amount: model.invBonds },
    { label: "Postal Deposits", amount: model.invPostal },
    { label: "PPF/EPF", amount: model.invPpf },
    { label: "ULIP", amount: model.invUlip },
  ];
  additionalProperties.forEach((p) => {
    const effective = Number(p.value || 0) * (Number(p.ownership ?? 100) / 100);
    rows.push({ label: `Property: ${p.name || "Unnamed"}`, amount: effective });
  });
  const totalAssets = rows.reduce((sum, r) => sum + r.amount, 0);
  const totalLiabilities = model.loanHome + model.loanCar + model.loanOther;
  const netWorth = totalAssets - totalLiabilities;

  byId("totalAssets").textContent = formatRs(totalAssets);
  byId("totalLiabilities").textContent = formatRs(totalLiabilities);
  byId("netWorth").textContent = formatRs(netWorth);

  const body = byId("networthBody");
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

function renderRoiTable() {
  const roiRows = [
    { p: "Real Estate Rate", r: 0.08, a: 0 },
    { p: "Equity(Shares+MF)", r: model.preRetRate / 100, a: model.invShares + model.invEquityMf },
    { p: "Debt - Saving/Liquid/ULIP", r: model.debtRate / 100, a: model.invSavings + model.invLiquidMf + model.invUlip },
    { p: "Debt MF", r: model.debtRate / 100, a: model.invDebtMf },
    { p: "Bonds & FDs", r: model.debtRate / 100, a: model.invBonds },
    { p: "Other investment", r: model.debtRate / 100, a: model.invPostal },
    { p: "PPF", r: 0.079, a: model.invPpf },
    { p: "Gold", r: 0.07, a: model.assetGold },
  ];
  const total = roiRows.reduce((s, r) => s + r.a, 0);
  const body = byId("roiBody");
  body.innerHTML = "";
  let totalRoi = 0;

  roiRows.forEach((r) => {
    const w = total ? r.a / total : 0;
    const roi = w * r.r;
    totalRoi += roi;
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${r.p}</td><td>${pct(r.r)}</td><td>${formatRs(r.a)}</td><td>${pct(w)}</td><td>${pct(roi)}</td>`;
    body.appendChild(tr);
  });

  const totalRow = document.createElement("tr");
  totalRow.innerHTML = `<th>Total</th><th></th><th>${formatRs(total)}</th><th></th><th>${pct(totalRoi)}</th>`;
  body.appendChild(totalRow);
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

    tr.innerHTML = `
      <td>${r.no}</td>
      <td>${r.year}</td>
      <td>${r.age}</td>
      <td>${formatRs(r.opBal)}</td>
      <td><input type="number" class="${ciClass}" data-cf-year="${r.year}" data-cf-key="cashIn"
          value="${Math.round(r.cashIn)}" ${!isPreRet ? "disabled" : ""}></td>
      <td><input type="number" class="cf-edit cf-growth-inp" data-cf-year="${r.year}" data-cf-key="growth" data-cf-post-ret="${r.age < retAge ? "0" : "1"}"
          value="${(r.growth * 100).toFixed(1)}" step="0.1"></td>
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

  const p = { top: 16, right: 16, bottom: 32, left: 56 };
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
    grid += `<text x="${p.left - 5}" y="${(+yy + 4).toFixed(1)}"
               text-anchor="end" font-size="10" fill="#94a3b8">${fmtChartVal(val)}</text>`;
  }

  // Red zero-line when balance can go negative
  const zeroLine = minY < 0
    ? `<line x1="${p.left}" y1="${yp(0).toFixed(1)}" x2="${w - p.right}" y2="${yp(0).toFixed(1)}"
         stroke="#ef4444" stroke-width="1.5" stroke-dasharray="4,3"/>`
    : "";

  // X-axis age labels — thin out to avoid overlap
  const labelEvery = Math.ceil(rows.length / (w < 750 ? 7 : 12));
  const labels = rows.map((r, i) => {
    if (i % labelEvery !== 0 && i !== rows.length - 1) return "";
    return `<text x="${xp(i).toFixed(1)}" y="${h - 6}"
              text-anchor="middle" font-size="9" fill="#94a3b8">${r.age}</text>`;
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
    lbl.setAttribute("font-size", "8.5");
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
  byId("adminAsOfDate").value = adminPortfolio.asOfDate || "";
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

  byId("adminTotalPortfolio").value = formatRs(totalEq + totalUf + totalIc);

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

  // ── Welcome banner ──────────────────────────────────────
  const dbName = byId("db-name");
  if (dbName) dbName.textContent = model.name || "—";
  const dbSub = byId("db-subtitle");
  if (dbSub) {
    const planYear = model.planDate ? new Date(model.planDate).getFullYear() : "—";
    dbSub.textContent = `Plan Date: ${planYear}  ·  Retirement Age: ${model.retirementAge || "—"}  ·  Life Expectancy: ${model.lifeExpectancy || "—"}`;
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
    + (model.invBonds||0) + (model.invPostal||0) + (model.invUlip||0);

  const propVal = (model.assetHome||0) + additionalProperties.reduce(
    (s, p) => s + Number(p.value||0) * (Number(p.ownership||100) / 100), 0);
  const physicalAssets = propVal + (model.assetCar||0) + (model.assetGold||0);

  const totalAssets = financialAssets + physicalAssets;
  const totalLiabilities = (model.loanHome||0) + (model.loanCar||0) + (model.loanOther||0);
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
    const gb = byId("db-goalsBody");
    if (gb && gs.goalStrategyRows) {
      gb.innerHTML = gs.goalStrategyRows.map(g => {
        const pct = g.corpus > 0 ? Math.min(100, Math.round((g.provision / g.corpus) * 100)) : 0;
        const barColor = pct >= 75 ? "var(--success)" : pct >= 40 ? "var(--warning)" : "var(--danger)";
        return `
          <tr>
            <td>
              <div class="goal-name-cell">${escHtml(g.name)}</div>
              <div class="goal-prog-wrap">
                <div class="goal-prog-bar" style="width:${pct}%;background:${barColor}"></div>
              </div>
              <div class="goal-prog-lbl" style="color:${barColor}">${pct}% funded</div>
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
  // Switch to dashboard tab
  const dashBtn = document.querySelector('#sheetTabs button[data-sheet="dashboard"]');
  if (dashBtn) dashBtn.click();
  saveCurrentPlan().catch(e => setStatus(`Save failed: ${e.message}`));
  // Push questionnaire results to Google Sheet
  pushToGoogleSheet();
}

// ── Google Sheet integration ───────────────────────────────────────────────
const GOOGLE_SHEET_WEBHOOK = "https://script.google.com/macros/s/AKfycbz89ghOAZTTbk85jUwJ_Z5xSpZo4_pVWnXs1IlYEHoB_rxW53ObYKZZxOLklRnfH9ec/exec";

function pushToGoogleSheet() {
  try {
    const user = auth.currentUser;
    const data = {
      timestamp:        new Date().toISOString(),
      investorName:     model.investorName   || "",
      email:            user ? user.email    : "",
      age:              model.currentAge     || "",
      retirementAge:    model.retirementAge  || "",
      monthlyIncome:    model.monthlyIncome  || "",
      monthlyExpenses:  model.monthlyExpenses || "",
      existingCorpus:   model.existingCorpus || "",
      riskProfile:      model.riskProfile    || "",
      goals:            (model.goals || []).map(g => g.name).join(", "),
      numGoals:         (model.goals || []).length,
      homeLoanLinked:   model.homeLoan       ? "Yes" : "No",
      netWorthHome:     model.networthHome   || "",
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

// ── Step 5: Assets ──────────────────────────────────────────
function wz5() {
  return `
    <div class="wz-section-label">Physical Assets</div>
    <div class="wz-grid three">
      ${fis("assetHome","Home / Property (₹)")}
      ${fis("assetCar","Car (₹)")}
      ${fis("assetGold","Gold & Jewellery (₹)")}
    </div>
    <div class="wz-section-label">Financial Investments</div>
    <div class="wz-grid three">
      ${fis("invEquityMf","Equity Mutual Funds (₹)")}
      ${fis("invDebtMf","Debt Mutual Funds (₹)")}
      ${fis("invElss","ELSS / Tax Saver MF (₹)")}
      ${fis("invLiquidMf","Liquid / Hybrid MF (₹)")}
      ${fis("invShares","Shares & Securities (₹)")}
      ${fis("invSavings","Savings / Bank FD (₹)")}
      ${fis("invBonds","Bonds (₹)")}
      ${fis("invPostal","Postal / NSC (₹)")}
      ${fis("invPpf","PPF (₹)")}
      <label>EPF — Current Value (₹)
        <input data-wz="invEpf" type="number" value="${model.invEpf||0}" placeholder="0">
      </label>
      ${fis("invUlip","ULIP (₹)")}
    </div>
    <div class="wz-section-label">Liabilities</div>
    <div class="wz-grid three">
      ${fis("loanHome","Home Loan Outstanding (₹)")}
      ${fis("loanCar","Car Loan Outstanding (₹)")}
      ${fis("loanOther","Other Loans (₹)")}
    </div>
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
}

function recalc() {
  const monthlyInflow = model.incomeMain + model.incomeSpouse;
  byId("age").value = yearsBetween(model.dob, model.planDate);
  byId("spouseAge").value = yearsBetween(model.spouseDob, model.planDate);
  byId("child1Age").value = yearsBetween(model.child1Dob, model.planDate);
  byId("child2Age").value = yearsBetween(model.child2Dob, model.planDate);

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
  renderBreakup(goalSummary.goalStrategyRows);

  latestState.goalSummary = goalSummary;
  latestState.networth = networthSummary;
  latestState.cashflow = cfRows;
  renderAdminNetworthSheet();
  renderDashboard();
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
        growth: r.growth,
        fv: r.fvEnd,
        out: r.cashOut,
        cl: r.clBal,
        goals: r.goals,
      });
    });
    wsCf.getColumn("growth").numFmt = "0%";
    styleGrid(wsCf);

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

function initExportButton() {
  const btn = byId("downloadExcelBtn");
  if (!btn) return;
  btn.addEventListener("click", downloadWorkbook);
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
