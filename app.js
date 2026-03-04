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
      byId("authRole").value = currentRole;
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
  const requestedRole = byId("authRole").value;
  const role = isAdminEmail(email) ? "admin" : "investor";
  if (requestedRole === "admin" && !isAdminEmail(email)) {
    throw new Error(`Only ${getAdminEmail()} can be admin.`);
  }
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
  const expectedRole = byId("authRole").value;
  if (!email || !password) return alert("Enter email and password.");
  if (expectedRole === "admin" && !isAdminEmail(email)) {
    throw new Error(`Only ${getAdminEmail()} can login as admin.`);
  }
  await auth.signInWithEmailAndPassword(email, password);
  if (!db || !auth.currentUser) return;
  const userDoc = await db.collection("users").doc(auth.currentUser.uid).get();
  const actualRole = userDoc.exists ? userDoc.data().role : isAdminEmail(email) ? "admin" : "investor";
  if (actualRole !== expectedRole) {
    await auth.signOut();
    alert(`This account is ${actualRole}. Please select ${actualRole} role before login.`);
  }
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
    "invUlip",
    "loanHome",
    "loanCar",
    "loanOther",
  ];
  ids.forEach((id) => {
    const el = byId(id);
    if (!el) return;
    el.value = model[id] ?? "";
  });
  const notes = byId("networthNotes");
  if (notes) notes.value = model.networthNotes || "";
}

function renderPropertyRows() {
  const body = byId("propertyBody");
  if (!body) return;
  body.innerHTML = "";
  additionalProperties.forEach((p, idx) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><input data-prop-idx="${idx}" data-prop-key="name" value="${p.name || ""}"></td>
      <td><input type="number" data-prop-idx="${idx}" data-prop-key="value" value="${p.value || 0}"></td>
      <td><input type="number" data-prop-idx="${idx}" data-prop-key="ownership" value="${p.ownership ?? 100}"></td>
      <td><input data-prop-idx="${idx}" data-prop-key="loanLinked" value="${p.loanLinked || ""}"></td>
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
    scheduleAutosave();
  });
  byId("adminAsOfDate")?.addEventListener("change", (e) => {
    if (!isAdmin()) return;
    adminPortfolio.asOfDate = e.target.value;
    scheduleAutosave();
  });
  byId("addEquityRowBtn")?.addEventListener("click", () => {
    if (!isAdmin()) return;
    adminPortfolio.equityRows.push({
      investorName: model.name || "",
      schemeName: "",
      type: "Equity",
      costValue: 0,
      units: 0,
      nav: 0,
      sipAmt: 0,
    });
    renderAdminNetworthSheet();
    scheduleAutosave();
  });
  byId("addUnifiRowBtn")?.addEventListener("click", () => {
    if (!isAdmin()) return;
    adminPortfolio.unifiRows.push({ investorName: model.name || "", schemeName: "", costValue: 0, currentValue: 0, dateOfInv: "", xirr: 0 });
    renderAdminNetworthSheet();
    scheduleAutosave();
  });
  byId("addIciciRowBtn")?.addEventListener("click", () => {
    if (!isAdmin()) return;
    adminPortfolio.iciciRows.push({ investorName: model.name || "", schemeName: "", costValue: 0, currentValue: 0, dateOfInv: "", xirr: 0 });
    renderAdminNetworthSheet();
    scheduleAutosave();
  });
}

function renderGoalInputRows() {
  const body = byId("goalBody");
  body.innerHTML = "";
  goals.forEach((g) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${g.name}</td>
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
  const retirementBaseOutflow =
    model.expHousehold +
    model.expLifestyle +
    model.expVehicle +
    model.expMediclaim +
    model.expUtilities +
    model.expCarInsurance +
    model.expMisc;
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
    const growth = age <= model.retirementAge ? model.preRetRate / 100 : model.postRetRate / 100;
    const effectiveCashIn = age <= model.retirementAge ? cashIn : 0;
    const fvEnd = opening * (1 + growth) + effectiveCashIn;
    const goalHit = nonRetirementGoals.filter((g) => g.targetYear === year);
    const retireOut = retirementMap.get(year) || 0;
    const cashOut = goalHit.reduce((sum, g) => sum + g.projectedValue, 0) + retireOut;
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
  const monthlyOutflow =
    model.expHousehold +
    model.expLifestyle +
    model.expEducation +
    model.expVehicle +
    model.expMediclaim +
    model.expUtilities +
    model.expCarInsurance +
    model.expMisc;
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
      <td>${g.name}</td>
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
      <td>${g.name}</td>
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
    labels += `<text x="${lineEndX + (onRight ? 4 : -4)}" y="${lineEndY - 3}" font-size="12" text-anchor="${onRight ? "start" : "end"}">${item.name}</text>`;
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
      ${d.name}: ${formatRs(d.value)}
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
    tr.innerHTML = `<td>${r.label}</td><td>${formatRs(r.amount)}</td><td>${pct(totalAssets ? r.amount / totalAssets : 0)}</td>`;
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
    p.innerHTML = `<span style="display:inline-block;width:10px;height:10px;background:${colors[i % colors.length]};margin-right:6px;"></span>${d.label}: ${formatRs(d.amount)}`;
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
  body.innerHTML = "";
  rows.forEach((r) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${r.no}</td>
      <td>${r.year}</td>
      <td>${r.age}</td>
      <td>${formatRs(r.opBal)}</td>
      <td>${formatRs(r.cashIn)}</td>
      <td>${Math.round(r.growth * 100)}%</td>
      <td>${formatRs(r.fvEnd)}</td>
      <td>${r.cashOut ? formatRs(r.cashOut) : ""}</td>
      <td>${formatRs(r.clBal)}</td>
      <td>${r.goals}</td>
    `;
    body.appendChild(tr);
  });
}

function renderCashflowChart(rows) {
  const svg = byId("cashflowChart");
  if (!svg || !rows.length) return;
  const w = 900;
  const h = 320;
  const p = { top: 12, right: 12, bottom: 38, left: 62 };
  const innerW = w - p.left - p.right;
  const innerH = h - p.top - p.bottom;
  const maxY = Math.max(...rows.map((r) => r.clBal), 1);
  const minY = 0;
  const xStep = innerW / Math.max(rows.length - 1, 1);
  const x = (i) => p.left + i * xStep;
  const y = (v) => p.top + ((maxY - v) / (maxY - minY || 1)) * innerH;
  const points = rows.map((r, i) => `${x(i)},${y(r.clBal)}`).join(" ");

  let grid = "";
  for (let i = 0; i <= 5; i += 1) {
    const val = (maxY / 5) * i;
    const yy = y(val);
    grid += `<line x1="${p.left}" y1="${yy}" x2="${w - p.right}" y2="${yy}" stroke="#8f8f8f" stroke-width="1"/>`;
    grid += `<text x="${p.left - 8}" y="${yy + 4}" text-anchor="end" font-size="11">${formatRs(val)}</text>`;
  }

  const markers = rows
    .map((r, i) => `<circle cx="${x(i)}" cy="${y(r.clBal)}" r="3" fill="#3e74b9"></circle>`)
    .join("");
  const labels = rows
    .map((r, i) => `<text x="${x(i)}" y="${h - 18}" text-anchor="middle" font-size="10">${r.age}</text>`)
    .join("");

  svg.innerHTML = `
    <rect width="${w}" height="${h}" fill="#dfdfdf"></rect>
    ${grid}
    <polyline points="${points}" fill="none" stroke="#3e74b9" stroke-width="3"></polyline>
    ${markers}
    ${labels}
  `;
}

function renderBreakup(goalOutput) {
  const body = byId("breakupBody");
  body.innerHTML = "";
  goalOutput.forEach((g, i) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${i + 1}</td>
      <td>${g.name}</td>
      <td>${formatRs(g.projectedValue)}</td>
      <td>${formatRs(g.provision)}</td>
      <td>${formatRs(g.gap)}</td>
      <td>${formatRs(g.sip)}</td>
    `;
    body.appendChild(tr);
  });
}

function renderAdminPortfolioRows(bodyId, rows, type) {
  const body = byId(bodyId);
  if (!body) return;
  body.innerHTML = "";
  rows.forEach((r, idx) => {
    const tr = document.createElement("tr");
    if (type === "equity") {
      const value = Number(r.units || 0) * Number(r.nav || 0);
      tr.innerHTML = `
        <td>${idx + 1}</td>
        <td><input data-admin-type="${type}" data-admin-idx="${idx}" data-key="investorName" value="${r.investorName || ""}" ${
          isAdmin() ? "" : "disabled"
        }></td>
        <td><input data-admin-type="${type}" data-admin-idx="${idx}" data-key="schemeName" value="${r.schemeName || ""}" ${
          isAdmin() ? "" : "disabled"
        }></td>
        <td><input data-admin-type="${type}" data-admin-idx="${idx}" data-key="type" value="${r.type || "Equity"}" ${
          isAdmin() ? "" : "disabled"
        }></td>
        <td><input type="number" data-admin-type="${type}" data-admin-idx="${idx}" data-key="costValue" value="${r.costValue || 0}" ${
          isAdmin() ? "" : "disabled"
        }></td>
        <td><input type="number" step="0.001" data-admin-type="${type}" data-admin-idx="${idx}" data-key="units" value="${
          r.units || 0
        }" ${isAdmin() ? "" : "disabled"}></td>
        <td><input type="number" step="0.01" data-admin-type="${type}" data-admin-idx="${idx}" data-key="nav" value="${
          r.nav || 0
        }" ${isAdmin() ? "" : "disabled"}></td>
        <td>${formatRs(value)}</td>
        <td><input type="number" data-admin-type="${type}" data-admin-idx="${idx}" data-key="sipAmt" value="${r.sipAmt || 0}" ${
          isAdmin() ? "" : "disabled"
        }></td>
        <td class="admin-only">${isAdmin() ? `<button type="button" data-admin-del="${type}:${idx}">Delete</button>` : ""}</td>
      `;
    } else {
      tr.innerHTML = `
        <td>${idx + 1}</td>
        <td><input data-admin-type="${type}" data-admin-idx="${idx}" data-key="investorName" value="${r.investorName || ""}" ${
          isAdmin() ? "" : "disabled"
        }></td>
        <td><input data-admin-type="${type}" data-admin-idx="${idx}" data-key="schemeName" value="${r.schemeName || ""}" ${
          isAdmin() ? "" : "disabled"
        }></td>
        <td><input type="number" data-admin-type="${type}" data-admin-idx="${idx}" data-key="costValue" value="${r.costValue || 0}" ${
          isAdmin() ? "" : "disabled"
        }></td>
        <td><input type="number" data-admin-type="${type}" data-admin-idx="${idx}" data-key="currentValue" value="${
          r.currentValue || 0
        }" ${isAdmin() ? "" : "disabled"}></td>
        <td><input type="date" data-admin-type="${type}" data-admin-idx="${idx}" data-key="dateOfInv" value="${
          r.dateOfInv || ""
        }" ${isAdmin() ? "" : "disabled"}></td>
        <td><input type="number" step="0.01" data-admin-type="${type}" data-admin-idx="${idx}" data-key="xirr" value="${
          r.xirr || 0
        }" ${isAdmin() ? "" : "disabled"}></td>
        <td class="admin-only">${isAdmin() ? `<button type="button" data-admin-del="${type}:${idx}">Delete</button>` : ""}</td>
      `;
    }
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

  const totalEq = eq.reduce((s, r) => s + Number(r.units || 0) * Number(r.nav || 0), 0);
  const totalUf = uf.reduce((s, r) => s + Number(r.currentValue || 0), 0);
  const totalIc = ic.reduce((s, r) => s + Number(r.currentValue || 0), 0);
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

function recalc() {
  const monthlyInflow = model.incomeMain + model.incomeSpouse;
  byId("age").value = yearsBetween(model.dob, model.planDate);
  byId("spouseAge").value = yearsBetween(model.spouseDob, model.planDate);
  byId("child1Age").value = yearsBetween(model.child1Dob, model.planDate);
  byId("child2Age").value = yearsBetween(model.child2Dob, model.planDate);

  const monthlyOutflow =
    model.expHousehold +
    model.expLifestyle +
    model.expEducation +
    model.expVehicle +
    model.expMediclaim +
    model.expUtilities +
    model.expCarInsurance +
    model.expMisc;

  const goalOutput = computeGoalOutput();
  const goalSummary = renderGoalSheet(goalOutput);
  renderGoalPie(goalSummary.goalStrategyRows);
  const networthSummary = renderNetworth();
  renderRoiTable();
  const cfRows = computeCashflow(goalOutput, goalSummary.requiredSip, monthlyInflow, monthlyOutflow);
  renderCashflowTable(cfRows);
  renderCashflowChart(cfRows);
  renderBreakup(goalOutput);

  latestState.goalSummary = goalSummary;
  latestState.networth = networthSummary;
  latestState.cashflow = cfRows;
  renderAdminNetworthSheet();
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
  "invUlip",
  "loanHome",
  "loanCar",
  "loanOther",
].forEach(bindInput);

bindStaticUiEvents();
renderGoalInputRows();
renderPropertyRows();
initTabs();
initExportButton();
applyRoleVisibility();
recalc();
setAppLocked(true);
initFirebase().catch((e) => setStatus(`Firebase init failed: ${e.message}`));
