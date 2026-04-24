const loginPanel = document.getElementById("loginPanel");
const appPanel = document.getElementById("appPanel");
const loginForm = document.getElementById("loginForm");
const loginMessage = document.getElementById("loginMessage");
const loginCopy = document.getElementById("loginCopy");
const loginButton = document.getElementById("loginButton");
const logoutButton = document.getElementById("logoutButton");
const menuToggleButton = document.getElementById("menuToggleButton");
const workspaceMenu = document.getElementById("workspaceMenu");
const form = document.getElementById("investmentForm");
const loadCompanyDetailsButton = document.getElementById("loadCompanyDetailsButton");
const formMessage = document.getElementById("formMessage");
const updatesList = document.getElementById("updatesList");
const authStatus = document.getElementById("authStatus");
const emailStatus = document.getElementById("emailStatus");
const recipientStatus = document.getElementById("recipientStatus");
const roleNotice = document.getElementById("roleNotice");
const refreshButton = document.getElementById("refreshButton");
const submitButton = document.getElementById("submitButton");
const downloadCsvButton = document.getElementById("downloadCsvButton");
const downloadExcelButton = document.getElementById("downloadExcelButton");
const downloadBackupButton = document.getElementById("downloadBackupButton");
const downloadFamilyOfficeWorkbookButton = document.getElementById(
  "downloadFamilyOfficeWorkbookButton"
);
const downloadReconciliationButton = document.getElementById("downloadReconciliationButton");
const saveAllReconciliationButton = document.getElementById("saveAllReconciliationButton");
const previewDigestButton = document.getElementById("previewDigestButton");
const sendDigestButton = document.getElementById("sendDigestButton");
const digestMessage = document.getElementById("digestMessage");
const digestPreview = document.getElementById("digestPreview");
const reconciliationMessage = document.getElementById("reconciliationMessage");
const importWorkbookFile = document.getElementById("importWorkbookFile");
const restoreBackupFile = document.getElementById("restoreBackupFile");
const importWorkbookMessage = document.getElementById("importWorkbookMessage");
const cancelEditButton = document.getElementById("cancelEditButton");
const editingInvestmentId = document.getElementById("editingInvestmentId");
const notesField = document.getElementById("notesField");
const deckSummaryField = document.getElementById("deckSummaryField");
const deckFile = document.getElementById("deckFile");
const deckMessage = document.getElementById("deckMessage");
const summarizeDeckButton = document.getElementById("summarizeDeckButton");
const deckDropZone = document.getElementById("deckDropZone");
const deckFileName = document.getElementById("deckFileName");
const capitalActivityList = document.getElementById("capitalActivityList");
const addCapitalActivityButton = document.getElementById("addCapitalActivityButton");
const documentFile = document.getElementById("documentFile");
const documentDropZone = document.getElementById("documentDropZone");
const documentMessage = document.getElementById("documentMessage");
const uploadedDocumentsList = document.getElementById("uploadedDocumentsList");
const companyDocumentFile = document.getElementById("companyDocumentFile");
const companyDocumentDropZone = document.getElementById("companyDocumentDropZone");
const companyDocumentMessage = document.getElementById("companyDocumentMessage");
const emailSummaryInput = document.getElementById("emailSummaryInput");
const summarizeEmailButton = document.getElementById("summarizeEmailButton");
const emailMessage = document.getElementById("emailMessage");
const dashboardCards = document.getElementById("dashboardCards");
const entityPerformanceCards = document.getElementById("entityPerformanceCards");
const dataQualitySummary = document.getElementById("dataQualitySummary");
const dataQualityList = document.getElementById("dataQualityList");
const entityDetailSection = document.getElementById("entityDetailSection");
const entityDetailTitle = document.getElementById("entityDetailTitle");
const entityDetailCopy = document.getElementById("entityDetailCopy");
const entityDetailSummary = document.getElementById("entityDetailSummary");
const entityDetailInvestments = document.getElementById("entityDetailInvestments");
const closeEntityDetailButton = document.getElementById("closeEntityDetailButton");
const entityFilter = document.getElementById("entityFilter");
const searchFilter = document.getElementById("searchFilter");
const statusFilter = document.getElementById("statusFilter");
const stageFilter = document.getElementById("stageFilter");
const ownerFilter = document.getElementById("ownerFilter");
const companySuggestions = document.getElementById("companySuggestions");
const companyPanel = document.getElementById("companyPanel");
const companyPanelTitle = document.getElementById("companyPanelTitle");
const companyPanelCopy = document.getElementById("companyPanelCopy");
const companySummary = document.getElementById("companySummary");
const companyHighlights = document.getElementById("companyHighlights");
const companyContactInfo = document.getElementById("companyContactInfo");
const companyPerformanceSummary = document.getElementById("companyPerformanceSummary");
const companyEntityPerformance = document.getElementById("companyEntityPerformance");
const companyOwnershipSummary = document.getElementById("companyOwnershipSummary");
const companyDeckSummaries = document.getElementById("companyDeckSummaries");
const companyDecisionLog = document.getElementById("companyDecisionLog");
const companyNextSteps = document.getElementById("companyNextSteps");
const companyFollowOnCapital = document.getElementById("companyFollowOnCapital");
const companyValuationHistory = document.getElementById("companyValuationHistory");
const companyTimeline = document.getElementById("companyTimeline");
const closeCompanyPanelButton = document.getElementById("closeCompanyPanelButton");
const taskForm = document.getElementById("taskForm");
const taskMessage = document.getElementById("taskMessage");
const saveTaskButton = document.getElementById("saveTaskButton");
const cancelTaskEditButton = document.getElementById("cancelTaskEditButton");
const editingTaskId = document.getElementById("editingTaskId");
const tasksList = document.getElementById("tasksList");
const companyTasks = document.getElementById("companyTasks");
const researchDeckFeed = document.getElementById("researchDeckFeed");
const researchNotesFeed = document.getElementById("researchNotesFeed");
const researchDocumentsFeed = document.getElementById("researchDocumentsFeed");
const researchDecisionFeed = document.getElementById("researchDecisionFeed");
const reconciliationList = document.getElementById("reconciliationList");
const workspaceViews = Array.from(document.querySelectorAll(".workspace-view"));
const workspaceMenuLinks = Array.from(document.querySelectorAll(".menu-link[data-view]"));

let currentUser = null;
let allInvestments = [];
let allCompanies = [];
let selectedCompany = "";
let selectedCompanyEntity = "";
let configuredEntities = [];
let companyPerformanceMap = new Map();
let entityPerformanceMap = new Map();
let uploadedDocuments = [];
let allTasks = [];
let activeWorkspaceView = "home";
let selectedEntity = "";
let digestStatus = {
  lastDigestSentAt: "",
  nextDigestDueAt: "",
  openReminderCount: 0
};
let dirtyReconciliationRows = new Set();
let savingAllReconciliation = false;

const moneyFieldNames = [
  "amount",
  "officialValue",
  "internalValue",
  "exitValue",
  "followOnCapitalAmount"
];

const CANONICAL_STATUSES = [
  "New Lead",
  "Under Review",
  "Approved",
  "Funded",
  "Active",
  "Partially Realized",
  "Realized",
  "Written Off",
  "Passed",
  "Closed / Archived"
];

const STATUS_ALIASES = {
  "new lead": "New Lead",
  newlead: "New Lead",
  pipeline: "New Lead",
  "under review": "Under Review",
  underreview: "Under Review",
  approved: "Approved",
  funded: "Funded",
  active: "Active",
  "partially realized": "Partially Realized",
  partiallyrealized: "Partially Realized",
  "partial realized": "Partially Realized",
  realized: "Realized",
  "written off": "Written Off",
  writtenoff: "Written Off",
  passed: "Passed",
  closed: "Closed / Archived",
  archived: "Closed / Archived",
  "closed archived": "Closed / Archived",
  "closed/archived": "Closed / Archived",
  "closed / archived": "Closed / Archived"
};

const ENTITY_ALIASES = {
  "Beaman Ventures": "Beaman Ventures",
  "Lee Beaman": "Lee Beaman",
  "Kat Trust": "Katherine Trust",
  "Nat Trust": "Natalie Trust",
  "Katherine Trust": "Katherine Trust",
  "Natalie Trust": "Natalie Trust"
};

function companyKey(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function normalizeStatusName(value) {
  const rawStatus = String(value || "").trim();
  if (!rawStatus) {
    return "";
  }

  const normalizedKey = rawStatus
    .toLowerCase()
    .replace(/[-_]+/g, " ")
    .replace(/\s*\/\s*/g, "/")
    .replace(/\s+/g, " ")
    .trim();

  return STATUS_ALIASES[normalizedKey] || STATUS_ALIASES[normalizedKey.replace(/\s+/g, "")] || rawStatus;
}

function statusEquals(left, right) {
  return normalizeStatusName(left) === normalizeStatusName(right);
}

function normalizeEntityName(value) {
  const raw = String(value || "").trim();
  return ENTITY_ALIASES[raw] || raw;
}

function entityKey(value) {
  return companyKey(normalizeEntityName(value));
}

function companyEntityKey(company, entity) {
  const normalizedCompany = companyKey(company);
  if (!normalizedCompany) {
    return "";
  }

  const normalizedEntity = entityKey(entity);
  return normalizedEntity ? `${normalizedCompany}::${normalizedEntity}` : normalizedCompany;
}

function getInvestmentPositionKey(investment) {
  return companyEntityKey(investment && investment.company, investment && investment.entity);
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function normalizeMoneyString(value) {
  const text = String(value || "").trim();
  if (!text) {
    return "";
  }

  const cleaned = text.replace(/[^0-9.]/g, "");
  const [integerPartRaw, decimalPartRaw = ""] = cleaned.split(".");
  const integerPart = integerPartRaw.replace(/^0+(?=\d)/, "") || integerPartRaw || "0";
  const decimalPart = decimalPartRaw.replace(/\./g, "").slice(0, 2);
  const withCommas = integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return decimalPart ? `${withCommas}.${decimalPart}` : withCommas;
}

function formatPhoneNumber(value) {
  const digits = String(value || "").replace(/\D/g, "").slice(0, 10);
  if (!digits) {
    return "";
  }

  if (digits.length <= 3) {
    return `(${digits}`;
  }

  if (digits.length <= 6) {
    return `(${digits.slice(0, 3)}) ${digits.slice(3)}`;
  }

  return `(${digits.slice(0, 3)}) ${digits.slice(3, 6)}-${digits.slice(6)}`;
}

function formatMoneyField(field) {
  if (!field) {
    return;
  }

  field.value = normalizeMoneyString(field.value);
}

function applyFormInputFormatting() {
  moneyFieldNames.forEach((name) => {
    const field = form.elements[name];
    if (field) {
      formatMoneyField(field);
    }
  });

  if (form.elements.contactPhone) {
    form.elements.contactPhone.value = formatPhoneNumber(form.elements.contactPhone.value);
  }
}

async function fetchJson(url, options) {
  const response = await fetch(url, options);
  const data = await response.json();

  if (!response.ok) {
    const error = new Error(data.error || "Request failed");
    error.status = response.status;
    throw error;
  }

  return data;
}

function setSignedInState(user) {
  currentUser = user;
  const isSignedIn = Boolean(user);

  loginPanel.classList.toggle("hidden", isSignedIn);
  appPanel.classList.toggle("hidden", !isSignedIn);
  logoutButton.classList.toggle("hidden", !isSignedIn);
  authStatus.textContent = isSignedIn
    ? `Signed in as ${user.email}${user.role ? ` • ${user.role}` : ""}`
    : "Please sign in to view updates";

  if (isSignedIn) {
    showWorkspaceView(activeWorkspaceView);
  }
}

function showWorkspaceView(viewName) {
  activeWorkspaceView = viewName || "home";

  workspaceViews.forEach((section) => {
    const matchesView = section.dataset.view === activeWorkspaceView;
    if (section.id === "companyPanel" && activeWorkspaceView === "portfolio") {
      if (!selectedCompany) {
        section.classList.add("hidden");
        return;
      }
    }
    if (section.id === "entityDetailSection" && activeWorkspaceView === "entity") {
      if (!selectedEntity) {
        section.classList.add("hidden");
        return;
      }
    }
    section.classList.toggle("hidden", !matchesView);
  });

  workspaceMenuLinks.forEach((link) => {
    link.classList.toggle("is-active", link.dataset.view === activeWorkspaceView);
  });
}

function canEditWorkspace() {
  return !currentUser || currentUser.role !== "viewer";
}

function selectedDeckFile() {
  return deckFile.files && deckFile.files[0] ? deckFile.files[0] : null;
}

function updateDeckFileLabel(file) {
  deckFileName.textContent = file
    ? `Selected: ${file.name}`
    : "No file selected yet";
}

function normalizeCapitalActivityRows(rows) {
  return (Array.isArray(rows) ? rows : [])
    .map((row) => ({
      date: String((row && row.date) || "").trim(),
      type: String((row && row.type) || "").trim(),
      amount: String((row && row.amount) || "").trim(),
      notes: String((row && row.notes) || "").trim()
    }))
    .filter((row) => row.date || row.type || row.amount || row.notes);
}

function buildLegacyCapitalActivityRows(investment) {
  const rows = [];
  const legacyInvestmentAmount = investment.capitalCallAmount || investment.amount || "";
  const legacyInvestmentDate = investment.capitalCallDate || investment.createdAt || "";

  if (legacyInvestmentDate || legacyInvestmentAmount) {
    rows.push({
      date: legacyInvestmentDate,
      type: "Investment Amount",
      amount: legacyInvestmentAmount,
      notes: ""
    });
  }

  if (investment.distributionDate || investment.distributionAmount) {
    rows.push({
      date: investment.distributionDate || "",
      type: "Distribution",
      amount: investment.distributionAmount || "",
      notes: ""
    });
  }

  return rows;
}

function isCommittedCapitalType(type) {
  return String(type || "").toLowerCase().includes("committed capital");
}

function renderCapitalActivityRows(rows = []) {
  const normalizedRows = normalizeCapitalActivityRows(rows);
  const rowsToRender = normalizedRows.length
    ? normalizedRows
    : [{ date: "", type: "Investment Amount", amount: "", notes: "" }];

  capitalActivityList.innerHTML = rowsToRender
    .map(
      (row, index) => `
        <article class="capital-activity-row" data-index="${index}">
          <div class="capital-activity-grid">
            <label>
              Date
              <input type="date" data-capital-field="date" value="${escapeHtml(row.date)}" />
            </label>
            <label>
              Type
              <select data-capital-field="type">
                <option value="Committed Capital" ${row.type === "Committed Capital" ? "selected" : ""}>Committed Capital</option>
                <option value="Investment Amount" ${row.type === "Investment Amount" ? "selected" : ""}>Investment Amount</option>
                <option value="Capital Call" ${row.type === "Capital Call" ? "selected" : ""}>Capital Call</option>
                <option value="Distribution" ${row.type === "Distribution" ? "selected" : ""}>Distribution</option>
                <option value="Dividend" ${row.type === "Dividend" ? "selected" : ""}>Dividend</option>
                <option value="Return of Capital" ${row.type === "Return of Capital" ? "selected" : ""}>Return of Capital</option>
                <option value="Partial Exit" ${row.type === "Partial Exit" ? "selected" : ""}>Partial Exit</option>
                <option value="Fee" ${row.type === "Fee" ? "selected" : ""}>Fee</option>
              </select>
            </label>
            <label>
              Amount
              <input type="text" inputmode="decimal" data-capital-field="amount" value="${escapeHtml(normalizeMoneyString(row.amount))}" placeholder="250,000" />
            </label>
          </div>
          <div class="capital-activity-row-footer">
            <label>
              Notes
              <input type="text" data-capital-field="notes" value="${escapeHtml(row.notes)}" placeholder="Optional note about this cash flow" />
            </label>
            <button type="button" class="secondary-button inline-action-button danger-button" data-action="remove-capital-activity" data-index="${index}">
              Remove
            </button>
          </div>
        </article>
      `
    )
    .join("");
}

function collectCapitalActivityRows() {
  return Array.from(capitalActivityList.querySelectorAll(".capital-activity-row"))
    .map((row) => ({
      date: row.querySelector('[data-capital-field="date"]')?.value || "",
      type: row.querySelector('[data-capital-field="type"]')?.value || "",
      amount: row.querySelector('[data-capital-field="amount"]')?.value || "",
      notes: row.querySelector('[data-capital-field="notes"]')?.value || ""
    }))
    .filter((row) => row.date || row.type || row.amount || row.notes);
}

function attachFormattedInputHandlers() {
  moneyFieldNames.forEach((name) => {
    const field = form.elements[name];
    if (!field || field.dataset.formattedBound === "true") {
      return;
    }

    field.dataset.formattedBound = "true";
    field.addEventListener("input", () => {
      formatMoneyField(field);
    });
    field.addEventListener("blur", () => {
      formatMoneyField(field);
    });
  });

  const phoneField = form.elements.contactPhone;
  if (phoneField && phoneField.dataset.formattedBound !== "true") {
    phoneField.dataset.formattedBound = "true";
    phoneField.addEventListener("input", () => {
      phoneField.value = formatPhoneNumber(phoneField.value);
    });
    phoneField.addEventListener("blur", () => {
      phoneField.value = formatPhoneNumber(phoneField.value);
    });
  }
}

function summarizeCapitalActivity(investment) {
  const rows = normalizeCapitalActivityRows(
    investment.capitalActivity && investment.capitalActivity.length
      ? investment.capitalActivity
      : buildLegacyCapitalActivityRows(investment)
  );

  if (!rows.length) {
    return "";
  }

  return rows
    .slice(0, 3)
    .map((row) => {
      const amount = row.amount ? `${investment.currency || "USD"} ${row.amount}` : "Amount not set";
      return `${row.type || "Activity"} ${amount}${row.date ? ` on ${row.date}` : ""}`;
    })
    .join(" • ");
}

function formatDisplayDate(value) {
  const parsed = parseDateValue(value, null);
  if (!parsed) {
    return "Date not set";
  }

  return parsed.toLocaleDateString(undefined, {
    year: "numeric",
    month: "short",
    day: "numeric"
  });
}

function renderUploadedDocuments() {
  if (!uploadedDocuments.length) {
    uploadedDocumentsList.innerHTML =
      '<p class="update-meta">No uploaded documents on this update yet.</p>';
    return;
  }

  uploadedDocumentsList.innerHTML = uploadedDocuments
    .map(
      (document) => `
        <article class="uploaded-document-card">
          <div>
            <p class="highlight-value">${escapeHtml(document.name)}</p>
            <p class="update-meta">${escapeHtml(formatDisplayDate(document.uploadedAt))}</p>
          </div>
          <div class="uploaded-document-actions">
            <a class="secondary-button inline-action-button" href="${escapeHtml(document.url)}" target="_blank" rel="noreferrer">Open</a>
            <button type="button" class="secondary-button inline-action-button danger-button" data-action="remove-document" data-document-id="${escapeHtml(document.id)}">Remove</button>
          </div>
        </article>
      `
    )
    .join("");
}

function renderRoleState() {
  const editable = canEditWorkspace();
  form.classList.toggle("hidden", !editable);
  taskForm.classList.toggle("hidden", !editable);
  sendDigestButton.classList.toggle("hidden", !editable);
  roleNotice.textContent = editable
    ? "Editors can add investments, tasks, documents, and research."
    : "Your account is view-only. You can review investments, research, and tasks, but editing is disabled.";
}

function currentFilters() {
  return {
    entity: entityFilter.value,
    search: searchFilter.value.trim().toLowerCase(),
    status: statusFilter.value,
    stage: stageFilter.value,
    owner: ownerFilter.value
  };
}

function findCompanyRecord(company, entity = "") {
  const key = companyEntityKey(company, entity);
  if (!key) {
    return null;
  }

  return getCompanyCollections(allInvestments).find((record) => record.key === key) || null;
}

function hydrateFormFromCompanyRecord(company) {
  const companyRecord = findCompanyRecord(company, form.elements.entity ? form.elements.entity.value : "");
  if (!companyRecord || !companyRecord.updates || !companyRecord.updates.length) {
    return false;
  }

  const latest = companyRecord.updates[0];
  const assignIfBlank = (fieldName, value) => {
    const field = form.elements[fieldName];
    if (!field || field.value || !value) {
      return;
    }
    field.value = value;
  };

  assignIfBlank("entity", latest.entity || "");
  assignIfBlank("currency", latest.currency || "USD");
  assignIfBlank("stage", latest.stage || "");
  assignIfBlank("status", normalizeStatusName(latest.status) || "");
  assignIfBlank("owner", latest.owner || "");
  assignIfBlank("nextStep", latest.nextStep || "");
  assignIfBlank("nextStepDueDate", latest.nextStepDueDate || "");
  assignIfBlank("contactName", latest.contactName || "");
  assignIfBlank("contactPosition", latest.contactPosition || "");
  assignIfBlank("contactEmail", latest.contactEmail || "");
  assignIfBlank("contactPhone", latest.contactPhone || "");
  assignIfBlank("ownershipPercent", latest.ownershipPercent || "");
  assignIfBlank("entityOwnershipPercent", latest.entityOwnershipPercent || "");
  assignIfBlank("ownershipNotes", latest.ownershipNotes || "");
  assignIfBlank("followOnCapitalAmount", latest.followOnCapitalAmount || "");
  assignIfBlank("followOnCapitalStatus", latest.followOnCapitalStatus || "");
  assignIfBlank("followOnCapitalNotes", latest.followOnCapitalNotes || "");

  if (!notesField.value && latest.notes) {
    notesField.value = latest.notes;
  }

  if (!deckSummaryField.value && latest.deckSummary) {
    deckSummaryField.value = latest.deckSummary;
  }

  if (!editingInvestmentId.value && !collectCapitalActivityRows().length) {
    const activityRows = normalizeCapitalActivityRows(
      latest.capitalActivity && latest.capitalActivity.length
        ? latest.capitalActivity
        : buildLegacyCapitalActivityRows(latest)
    );
    if (activityRows.length) {
      renderCapitalActivityRows(activityRows);
    }
  }

  applyFormInputFormatting();
  return true;
}

function filterInvestments(investments) {
  const filters = currentFilters();
  return investments.filter((investment) => {
    const searchHaystack = [
      investment.entity,
      investment.company,
      investment.notes,
      investment.deckSummary,
      investment.owner,
      investment.nextStep,
      investment.capitalCallAmount,
      investment.distributionAmount,
      investment.officialValue,
      investment.internalValue,
      investment.exitValue,
      investment.ownershipPercent,
      investment.entityOwnershipPercent,
      investment.ownershipNotes,
      investment.followOnCapitalStatus,
      investment.followOnCapitalNotes,
      investment.contactName,
      investment.contactPosition,
      investment.contactEmail,
      investment.contactPhone,
      investment.documentLinks,
      Array.isArray(investment.documents)
        ? investment.documents.map((document) => document.name).join(" ")
        : "",
      investment.decisionType,
      investment.decisionSummary,
      investment.submittedBy
    ]
      .join(" ")
      .toLowerCase();

    const matchesSearch = !filters.search || searchHaystack.includes(filters.search);
    const matchesEntity =
      !filters.entity || normalizeEntityName(investment.entity) === normalizeEntityName(filters.entity);
    const matchesStatus = !filters.status || statusEquals(investment.status, filters.status);
    const matchesStage = !filters.stage || investment.stage === filters.stage;
    const matchesOwner = !filters.owner || investment.owner === filters.owner;

    return matchesEntity && matchesSearch && matchesStatus && matchesStage && matchesOwner;
  });
}

function toNumber(value) {
  const cleaned = String(value ?? "")
    .trim()
    .replace(/[$,\s]/g, "")
    .replace(/[^\d.-]/g, "");
  const amount = Number(cleaned);
  return Number.isFinite(amount) ? amount : 0;
}

function formatMoney(value) {
  return `$${toNumber(value).toLocaleString()}`;
}

function formatPercent(value) {
  if (!Number.isFinite(value)) {
    return "N/A";
  }

  return `${(value * 100).toFixed(1)}%`;
}

function formatTurns(value) {
  if (!Number.isFinite(value)) {
    return "N/A";
  }

  return `${value.toFixed(2)}x`;
}

function parseDateValue(value, fallback) {
  const text = String(value || "").trim();
  if (!text) {
    return fallback || null;
  }

  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? fallback || null : parsed;
}

function yearFraction(startDate, endDate) {
  return (endDate.getTime() - startDate.getTime()) / (365 * 24 * 60 * 60 * 1000);
}

function xnpv(rate, cashFlows) {
  const firstDate = cashFlows[0].date;
  return cashFlows.reduce(
    (sum, cashFlow) =>
      sum + cashFlow.amount / Math.pow(1 + rate, yearFraction(firstDate, cashFlow.date)),
    0
  );
}

function dxnpv(rate, cashFlows) {
  const firstDate = cashFlows[0].date;
  return cashFlows.reduce((sum, cashFlow) => {
    const fraction = yearFraction(firstDate, cashFlow.date);
    if (fraction === 0) {
      return sum;
    }

    return sum - (fraction * cashFlow.amount) / Math.pow(1 + rate, fraction + 1);
  }, 0);
}

function calculateXirr(cashFlows) {
  if (!Array.isArray(cashFlows) || cashFlows.length < 2) {
    return null;
  }

  const sorted = cashFlows
    .filter((cashFlow) => cashFlow.date instanceof Date && Number.isFinite(cashFlow.amount))
    .sort((left, right) => left.date - right.date);

  const hasNegative = sorted.some((cashFlow) => cashFlow.amount < 0);
  const hasPositive = sorted.some((cashFlow) => cashFlow.amount > 0);
  if (sorted.length < 2 || !hasNegative || !hasPositive) {
    return null;
  }

  let rate = 0.15;
  for (let index = 0; index < 50; index += 1) {
    const value = xnpv(rate, sorted);
    const derivative = dxnpv(rate, sorted);
    if (Math.abs(value) < 1e-7) {
      return rate;
    }

    if (!Number.isFinite(derivative) || Math.abs(derivative) < 1e-10) {
      break;
    }

    const nextRate = rate - value / derivative;
    if (!Number.isFinite(nextRate) || nextRate <= -0.9999 || nextRate > 1000) {
      break;
    }

    rate = nextRate;
  }

  let low = -0.9999;
  let high = 1;
  let lowValue = xnpv(low, sorted);
  let highValue = xnpv(high, sorted);

  for (let attempt = 0; attempt < 12 && lowValue * highValue > 0; attempt += 1) {
    high *= 2;
    highValue = xnpv(high, sorted);
  }

  if (lowValue * highValue > 0) {
    return null;
  }

  for (let iteration = 0; iteration < 80; iteration += 1) {
    const mid = (low + high) / 2;
    const midValue = xnpv(mid, sorted);
    if (Math.abs(midValue) < 1e-7) {
      return mid;
    }

    if (lowValue * midValue <= 0) {
      high = mid;
    } else {
      low = mid;
      lowValue = midValue;
    }
  }

  return (low + high) / 2;
}

function pickLatestNumericValue(updates, fieldName) {
  const sorted = [...updates].sort((left, right) => {
    const rightDate = parseDateValue(right.valuationDate, parseDateValue(right.createdAt, new Date(0)));
    const leftDate = parseDateValue(left.valuationDate, parseDateValue(left.createdAt, new Date(0)));
    return rightDate - leftDate;
  });

  for (const update of sorted) {
    const amount = toNumber(update[fieldName]);
    if (amount > 0) {
      return {
        value: amount,
        date: parseDateValue(update.valuationDate, parseDateValue(update.createdAt, new Date()))
      };
    }
  }

  return { value: 0, date: parseDateValue(sorted[0] && sorted[0].createdAt, new Date()) };
}

function buildPerformanceInputs(updates) {
  const updateActivities = updates.map((update) => {
    const activityRows = normalizeCapitalActivityRows(
      update.capitalActivity && update.capitalActivity.length
        ? update.capitalActivity
        : buildLegacyCapitalActivityRows(update)
    );

    return {
      update,
      activities: activityRows.map((activity) => ({
        ...activity,
        fallbackDate: parseDateValue(update.createdAt, null)
      }))
    };
  });

  const reconciliationOverride = updateActivities.find(({ activities }) =>
    activities.some((activity) => {
      const type = String(activity.type || "").toLowerCase();
      const notes = String(activity.notes || "").trim();
      return (
        (type.includes("capital call") ||
          type.includes("investment amount") ||
          type.includes("fee")) &&
        notes === "Updated from entity reconciliation"
      );
    })
  );

  const normalizedActivities = updateActivities.flatMap(({ update, activities }) => {
    if (!reconciliationOverride || reconciliationOverride.update.id === update.id) {
      return activities;
    }

    return activities.filter((activity) => {
      const type = String(activity.type || "").toLowerCase();
      return !(
        type.includes("capital call") ||
        type.includes("investment amount") ||
        type.includes("fee")
      );
    });
  });

  const hasActualCalledCapital = normalizedActivities.some((activity) => {
    const type = String(activity.type || "").toLowerCase();
    return type.includes("capital call") || type.includes("fee");
  });
  const committedCapital = normalizedActivities.reduce((sum, activity) => {
    const amount = toNumber(activity.amount);
    return isCommittedCapitalType(activity.type) ? sum + amount : sum;
  }, 0);

  const effectiveActivities = normalizedActivities.filter((activity) => {
    const type = String(activity.type || "").toLowerCase();
    if (isCommittedCapitalType(type)) {
      return false;
    }
    if (hasActualCalledCapital && type.includes("investment amount")) {
      return false;
    }

    return true;
  });

  const investedCapital = effectiveActivities.reduce((sum, activity) => {
    const type = String(activity.type || "").toLowerCase();
    const amount = toNumber(activity.amount);
    return type.includes("capital call") ||
      type.includes("investment amount") ||
      type.includes("fee")
      ? sum + amount
      : sum;
  }, 0);
  const distributions = effectiveActivities.reduce((sum, activity) => {
    const type = String(activity.type || "").toLowerCase();
    const amount = toNumber(activity.amount);
    return !type.includes("capital call") &&
      !isCommittedCapitalType(type) &&
      !type.includes("investment amount") &&
      !type.includes("fee")
      ? sum + amount
      : sum;
  }, 0);
  const officialMark = pickLatestNumericValue(updates, "officialValue");
  const internalMark = pickLatestNumericValue(updates, "internalValue");
  const exitMark = pickLatestNumericValue(updates, "exitValue");

  const baseCashFlows = [];
  effectiveActivities.forEach((activity) => {
    const amount = toNumber(activity.amount);
    const date = parseDateValue(activity.date, activity.fallbackDate);
    const type = String(activity.type || "").toLowerCase();

    if (!amount || !date) {
      return;
    }

    if (type.includes("capital call") || type.includes("investment amount") || type.includes("fee")) {
      baseCashFlows.push({ date, amount: -amount });
      return;
    }

    baseCashFlows.push({ date, amount });
  });

  return {
    committedCapital,
    investedCapital,
    distributions,
    officialMark,
    internalMark,
    exitMark,
    baseCashFlows
  };
}

function buildPerformanceView(baseCashFlows, terminalMark, investedCapital, distributions) {
  const terminalValue = terminalMark && terminalMark.value ? terminalMark.value : 0;
  const terminalDate = terminalMark && terminalMark.date ? terminalMark.date : null;
  const cashFlows = [...baseCashFlows];
  if (terminalValue > 0 && terminalDate) {
    cashFlows.push({ date: terminalDate, amount: terminalValue });
  }

  return {
    xirr: calculateXirr(cashFlows),
    moic:
      investedCapital > 0 ? (distributions + terminalValue) / investedCapital : null
  };
}

function buildCompanyPerformance(updates) {
  const {
    committedCapital,
    investedCapital,
    distributions,
    officialMark,
    internalMark,
    exitMark,
    baseCashFlows
  } = buildPerformanceInputs(updates);

  return {
    committedCapital,
    investedCapital,
    distributions,
    officialValue: officialMark.value,
    internalValue: internalMark.value,
    exitValue: exitMark.value,
    official: buildPerformanceView(baseCashFlows, officialMark, investedCapital, distributions),
    internal: buildPerformanceView(baseCashFlows, internalMark, investedCapital, distributions),
    exit: buildPerformanceView(baseCashFlows, exitMark, investedCapital, distributions)
  };
}

function getCompanyCollections(investments) {
  if (allCompanies.length) {
    return allCompanies
      .flatMap((company) => {
        const groupedUpdates = new Map();

        (company.updates || []).forEach((investment) => {
          const key = getInvestmentPositionKey(investment);
          if (!key) {
            return;
          }

          if (!groupedUpdates.has(key)) {
            groupedUpdates.set(key, []);
          }

          groupedUpdates.get(key).push(investment);
        });

        return Array.from(groupedUpdates.entries()).map(([key, updates]) => {
          const sortedUpdates = [...updates].sort(
            (left, right) => new Date(right.createdAt) - new Date(left.createdAt)
          );
          const latest = sortedUpdates[0];
          const sameEntity = (row) =>
            normalizeEntityName(row && row.entity) === normalizeEntityName(latest.entity);

          return {
            key,
            companyKey: company.companyKey,
            latest,
            updates: sortedUpdates,
            tasks: (company.tasks || []).filter(sameEntity),
            documents: (company.documents || []).filter(
              (document) => !document.entity || sameEntity(document)
            ),
            researchEntries: (company.researchEntries || []).filter(
              (entry) => !entry.entity || sameEntity(entry)
            ),
            capitalActivities: (company.capitalActivities || []).filter(
              (entry) => !entry.entity || sameEntity(entry)
            ),
            followOnHistory: (company.followOnHistory || []).filter(
              (entry) => !entry.entity || sameEntity(entry)
            ),
            decisionLog: (company.decisionLog || []).filter(
              (entry) => !entry.entity || sameEntity(entry)
            ),
            valuationHistory: (company.valuationHistory || []).filter(
              (entry) => !entry.entity || sameEntity(entry)
            ),
            ownershipHistory: (company.ownershipHistory || []).filter(
              (entry) => !entry.entity || sameEntity(entry)
            ),
            performance: buildCompanyPerformance(sortedUpdates)
          };
        });
      })
      .filter((company) => company.latest);
  }

  const grouped = new Map();

  investments.forEach((investment) => {
    const key = getInvestmentPositionKey(investment);
    if (!key) {
      return;
    }

    if (!grouped.has(key)) {
      grouped.set(key, []);
    }

    grouped.get(key).push(investment);
  });

  return Array.from(grouped.entries())
    .map(([key, updates]) => {
      const sortedUpdates = [...updates].sort(
        (left, right) => new Date(right.createdAt) - new Date(left.createdAt)
      );
      return {
        key,
        latest: sortedUpdates[0],
        updates: sortedUpdates,
        performance: buildCompanyPerformance(sortedUpdates)
      };
    })
    .filter((company) => company.latest);
}

function buildAggregatePerformance(companyCollections) {
  const companyInputs = companyCollections.map((company) => ({
    company,
    inputs: buildPerformanceInputs(company.updates)
  }));

  const reportedAmount = calculateReportedAmountTotal(companyInputs);
  const investedCapital = companyInputs.reduce(
    (sum, { inputs }) => sum + inputs.investedCapital,
    0
  );
  const distributions = companyInputs.reduce(
    (sum, { inputs }) => sum + inputs.distributions,
    0
  );
  const officialValue = companyInputs.reduce(
    (sum, { inputs }) => sum + inputs.officialMark.value,
    0
  );
  const internalValue = companyInputs.reduce(
    (sum, { inputs }) => sum + inputs.internalMark.value,
    0
  );
  const exitValue = companyInputs.reduce(
    (sum, { inputs }) => sum + inputs.exitMark.value,
    0
  );

  const buildAggregateView = (markName) => {
    const cashFlows = [];
    let terminalTotal = 0;

    companyInputs.forEach(({ inputs }) => {
      cashFlows.push(...inputs.baseCashFlows);
      const terminalMark = inputs[markName];
      if (terminalMark.value > 0 && terminalMark.date) {
        cashFlows.push({ date: terminalMark.date, amount: terminalMark.value });
        terminalTotal += terminalMark.value;
      }
    });

    return {
      xirr: calculateXirr(cashFlows),
      moic:
        investedCapital > 0 ? (distributions + terminalTotal) / investedCapital : null
    };
  };

  return {
    reportedAmount,
    investedCapital,
    distributions,
    officialValue,
    internalValue,
    exitValue,
    official: buildAggregateView("officialMark"),
    internal: buildAggregateView("internalMark"),
    exit: buildAggregateView("exitMark")
  };
}

function calculateReportedAmountTotal(companyInputs) {
  return companyInputs.reduce((sum, { company }) => {
    const latest = company && company.latest ? company.latest : null;
    const latestStatus = String((latest && latest.status) || "").trim();
    const latestReportedAmount = toNumber(latest && latest.amount);
    const includeInReportedAmount = isCommittedStatus(latestStatus);

    return includeInReportedAmount ? sum + latestReportedAmount : sum;
  }, 0);
}

function isCommittedStatus(status) {
  return ["Approved", "Closed / Archived"].includes(normalizeStatusName(status));
}

function isPipelineStatus(status) {
  return ["New Lead", "Under Review"].includes(normalizeStatusName(status));
}

function isPipelineRow(row) {
  return (
    isPipelineStatus(row && row.latest && row.latest.status) ||
    isPipelineStatus(row && row.latest && row.latest.stage)
  );
}

function buildEntityRows(investments, entity) {
  const normalizedEntity = normalizeEntityName(entity);
  return getCompanyCollections(investments)
    .filter((company) => {
      if (!normalizedEntity) {
        return true;
      }

      return normalizeEntityName(company.latest.entity) === normalizedEntity;
    })
    .map((company) => {
      const performance = company.performance || buildCompanyPerformance(company.updates);
      const status = normalizeStatusName(company.latest && company.latest.status);
      const reportedAmount = toNumber(company.latest && company.latest.amount);
      const includeReportedAmount = isCommittedStatus(status);

      return {
        company,
        latest: company.latest,
        performance,
        reportedAmount,
        includedReportedAmount: includeReportedAmount ? reportedAmount : 0,
        includeReportedAmount
      };
    });
}

function sumEntityRows(rows, selector) {
  return rows.reduce((sum, row) => sum + selector(row), 0);
}

function buildEntityRowTotals(rows) {
  const entityCompanies = rows.map((row) => row.company);
  const aggregatePerformance = buildAggregatePerformance(entityCompanies);

  return {
    reportedAmount: sumEntityRows(rows, (row) => row.includedReportedAmount),
    investedCapital: sumEntityRows(rows, (row) => row.performance.investedCapital),
    distributions: sumEntityRows(rows, (row) => row.performance.distributions),
    officialValue: sumEntityRows(rows, (row) => row.performance.officialValue),
    internalValue: sumEntityRows(rows, (row) => row.performance.internalValue),
    exitValue: sumEntityRows(rows, (row) => row.performance.exitValue),
    official: aggregatePerformance.official,
    internal: aggregatePerformance.internal,
    exit: aggregatePerformance.exit
  };
}

function csvEscape(value) {
  const text = String(value === null || value === undefined ? "" : value);
  if (/[",\n\r]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }

  return text;
}

function downloadTextFile(filename, content, mimeType = "text/csv;charset=utf-8") {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function buildReconciliationCsv() {
  const headers = [
    "Entity",
    "Company",
    "Stage",
    "Status",
    "Owner",
    "Reported amount field",
    "Included committed capital",
    "Included in committed total?",
    "Called capital",
    "Official NAV",
    "Internal NAV",
    "Latest update date",
    "Latest update id"
  ];
  const rows = [];
  const entities = Array.from(
    new Set(
      configuredEntities
        .concat(
          getCompanyCollections(allInvestments)
            .map((company) => normalizeEntityName(company.latest.entity))
            .filter(Boolean)
        )
        .map(normalizeEntityName)
    )
  ).filter(Boolean);

  entities.forEach((entity) => {
    const entityRows = buildEntityRows(allInvestments, entity).sort((left, right) =>
      String(left.latest.company || "").localeCompare(String(right.latest.company || ""), undefined, {
        sensitivity: "base"
      })
    );

    entityRows.forEach((row) => {
      rows.push([
        entity,
        row.latest.company || "",
        row.latest.stage || "",
        normalizeStatusName(row.latest.status) || "",
        row.latest.owner || "",
        row.reportedAmount,
        row.includedReportedAmount,
        row.includeReportedAmount ? "Yes" : "No",
        row.performance.investedCapital,
        row.performance.officialValue,
        row.performance.internalValue,
        row.latest.updatedAt || row.latest.createdAt || "",
        row.latest.id || ""
      ]);
    });

    const totals = buildEntityRowTotals(entityRows);
    rows.push([
      entity,
      "SUBTOTAL",
      "",
      "",
      "",
      "",
      totals.reportedAmount,
      "",
      totals.investedCapital,
      totals.officialValue,
      totals.internalValue,
      "",
      ""
    ]);
  });

  return [headers, ...rows].map((row) => row.map(csvEscape).join(",")).join("\n");
}

function buildCompanyPerformanceMap(investments) {
  const grouped = new Map();

  investments.forEach((investment) => {
    const key = getInvestmentPositionKey(investment);
    if (!key) {
      return;
    }

    if (!grouped.has(key)) {
      grouped.set(key, []);
    }
    grouped.get(key).push(investment);
  });

  const performanceMap = new Map();
  grouped.forEach((updates, key) => {
    performanceMap.set(key, buildCompanyPerformance(updates));
  });

  return performanceMap;
}

function buildEntityPerformanceMap(investments) {
  const performanceMap = new Map();
  const companyCollections = getCompanyCollections(investments);
  const entityList = Array.from(
    new Set(
      configuredEntities
        .concat(
          companyCollections.map((company) => normalizeEntityName(company.latest.entity)).filter(Boolean)
        )
        .map(normalizeEntityName)
    )
  ).filter(Boolean);

  entityList.forEach((entity) => {
    const entityCompanies = companyCollections.filter(
      (company) => normalizeEntityName(company.latest.entity) === normalizeEntityName(entity)
    );
    performanceMap.set(entity, buildAggregatePerformance(entityCompanies));
  });

  return performanceMap;
}

function buildDashboardCards(investments) {
  const companySummaries = getCompanyCollections(investments);
  const allEntityRows = buildEntityRows(investments);
  const qualityAlerts = buildDataQualityAlerts();
  const pipelineRows = allEntityRows.filter(isPipelineRow);
  const openCount = pipelineRows.length;
  const openPipelineAmount = sumEntityRows(pipelineRows, (row) => row.reportedAmount);
  const approvedCount = companySummaries.filter((summary) =>
    statusEquals(summary.latest && summary.latest.status, "Approved")
  ).length;
  const openReminderCount = allTasks.filter(
    (task) =>
      task.autoManaged &&
      task.sourceKind === "next-step" &&
      String(task.status || "").trim() !== "Completed"
  ).length;
  const totalCommittedCapital = sumEntityRows(
    allEntityRows,
    (row) => row.includedReportedAmount
  );
  const totalInvestedCapital = sumEntityRows(
    allEntityRows,
    (row) => row.performance.investedCapital
  );
  const officialNav = sumEntityRows(allEntityRows, (row) => row.performance.officialValue);
  const internalNav = sumEntityRows(allEntityRows, (row) => row.performance.internalValue);

  const cards = [
    { label: "Updates", value: String(investments.length), action: "portfolio" },
    { label: "Companies", value: String(companySummaries.length), action: "portfolio" },
    { label: "Pipeline deals", value: String(openCount), action: "portfolio" },
    {
      label: "Pipeline amount",
      value: formatMoney(openPipelineAmount),
      action: "portfolio"
    },
    {
      label: "Approved",
      value: String(approvedCount),
      action: "portfolio",
      status: "Approved"
    },
    { label: "Open reminders", value: String(openReminderCount), action: "tasks" },
    { label: "Data alerts", value: String(qualityAlerts.length), action: "quality" },
    { label: "Total committed capital", value: formatMoney(totalCommittedCapital), action: "portfolio" },
    { label: "Called capital", value: formatMoney(totalInvestedCapital), action: "portfolio" },
    { label: "Official NAV", value: formatMoney(officialNav), action: "portfolio" },
    { label: "Internal NAV", value: formatMoney(internalNav), action: "portfolio" }
  ];

  const entityTotals = (configuredEntities.length ? configuredEntities.map(normalizeEntityName) : [])
    .map((entity) => ({
      label: entity,
      value: String(
        companySummaries.filter(
          (company) => normalizeEntityName(company.latest.entity) === normalizeEntityName(entity)
        ).length
      ),
      action: "entity",
      entity
    }));

  return cards.concat(entityTotals);
}

function daysSinceDate(value) {
  const parsed = parseDateValue(value, null);
  if (!parsed) {
    return null;
  }

  return Math.floor((Date.now() - parsed.getTime()) / (24 * 60 * 60 * 1000));
}

function addQualityAlert(alerts, row, severity, title, detail) {
  alerts.push({
    id: row.latest.id,
    company: row.latest.company || "Unnamed investment",
    entity: normalizeEntityName(row.latest.entity) || "No entity",
    severity,
    title,
    detail
  });
}

function buildDataQualityAlerts() {
  const alerts = [];
  const rows = buildEntityRows(allInvestments);

  rows.forEach((row) => {
    const latest = row.latest || {};
    const performance = row.performance || {};
    const normalizedStatus = normalizeStatusName(latest.status);
    const reportedAmount = row.reportedAmount;
    const calledCapital = performance.investedCapital || 0;
    const hasCommittedStatus = isCommittedStatus(normalizedStatus);
    const isPipeline = isPipelineRow(row);
    const staleValuationDays = daysSinceDate(latest.valuationDate);

    if (!latest.company) {
      addQualityAlert(alerts, row, "High", "Missing company name", "This row cannot be reconciled cleanly without a company name.");
    }

    if (!normalizeEntityName(latest.entity)) {
      addQualityAlert(alerts, row, "High", "Missing entity", "Assign this investment to Beaman Ventures, Lee Beaman, Katherine Trust, or Natalie Trust.");
    }

    if (!normalizedStatus) {
      addQualityAlert(alerts, row, "High", "Missing status", "Choose a status so the app knows whether this belongs in pipeline, committed capital, or archive views.");
    }

    if (latest.status && normalizedStatus !== String(latest.status || "").trim()) {
      addQualityAlert(
        alerts,
        row,
        "Medium",
        "Legacy status wording",
        `This record uses "${latest.status}". Save it once as "${normalizedStatus}" to clean up filters and exports.`
      );
    }

    if (hasCommittedStatus && !reportedAmount) {
      addQualityAlert(alerts, row, "High", "Committed deal missing reported amount", "Approved or closed deals should have a reported amount for committed capital totals.");
    }

    if (isPipeline && !reportedAmount) {
      addQualityAlert(alerts, row, "Medium", "Pipeline deal missing amount", "Pipeline deals can stay open, but adding an expected amount makes the pipeline total useful.");
    }

    if (reportedAmount > 0 && calledCapital > reportedAmount) {
      addQualityAlert(
        alerts,
        row,
        "High",
        "Called capital exceeds committed capital",
        `${formatMoney(calledCapital)} called against ${formatMoney(reportedAmount)} committed.`
      );
    }

    if ((performance.officialValue > 0 || performance.internalValue > 0) && !latest.valuationDate) {
      addQualityAlert(alerts, row, "Medium", "Valuation date missing", "Add a valuation date so stale marks and XIRR timing are easier to audit.");
    } else if (staleValuationDays !== null && staleValuationDays > 180) {
      addQualityAlert(
        alerts,
        row,
        "Medium",
        "Valuation may be stale",
        `Latest valuation date is ${formatDisplayDate(latest.valuationDate)}, about ${staleValuationDays} days ago.`
      );
    }

    if ((hasCommittedStatus || statusEquals(normalizedStatus, "Funded") || statusEquals(normalizedStatus, "Active")) && !latest.nextStep) {
      addQualityAlert(alerts, row, "Low", "No next step", "Add a next step if this investment needs an upcoming follow-up or valuation review.");
    }

    if ((hasCommittedStatus || statusEquals(normalizedStatus, "Funded") || statusEquals(normalizedStatus, "Active")) && !latest.contactName && !latest.contactEmail && !latest.contactPhone) {
      addQualityAlert(alerts, row, "Low", "Contact info missing", "Add the best contact for future updates.");
    }
  });

  const severityRank = { High: 0, Medium: 1, Low: 2 };
  return alerts.sort((left, right) => {
    const severityDelta = severityRank[left.severity] - severityRank[right.severity];
    if (severityDelta) {
      return severityDelta;
    }

    return left.company.localeCompare(right.company);
  });
}

function renderDashboard(investments) {
  const cards = buildDashboardCards(investments);
  dashboardCards.innerHTML = cards
    .map(
      (card) => `
        <article
          class="dashboard-card"
          ${card.action ? `data-dashboard-action="${escapeHtml(card.action)}"` : ""}
          ${card.entity ? `data-entity="${escapeHtml(card.entity)}"` : ""}
          ${card.status ? `data-status="${escapeHtml(card.status)}"` : ""}
        >
          <p class="dashboard-label">${escapeHtml(card.label)}</p>
          <p class="dashboard-value">${escapeHtml(card.value)}</p>
        </article>
      `
    )
    .join("");

  const entityCards = Array.from(entityPerformanceMap.entries())
    .map(([entity]) => ({ entity, rows: buildEntityRows(allInvestments, entity) }))
    .filter(
      ({ entity }) =>
        !currentFilters().entity ||
        normalizeEntityName(currentFilters().entity) === normalizeEntityName(entity)
    );

  entityPerformanceCards.innerHTML = entityCards
    .map(
      ({ entity, rows }) => {
        const totals = buildEntityRowTotals(rows);
        const metrics = [
          { label: "Total committed capital", value: formatMoney(totals.reportedAmount) },
          { label: "Called capital", value: formatMoney(totals.investedCapital) },
          {
            label: "Unfunded commitment",
            value: formatMoney(Math.max(totals.reportedAmount - totals.investedCapital, 0))
          },
          { label: "Official NAV", value: formatMoney(totals.officialValue) },
          { label: "Internal NAV", value: formatMoney(totals.internalValue) },
          { label: "Internal XIRR", value: formatPercent(totals.internal.xirr) },
          { label: "Internal MOIC", value: formatTurns(totals.internal.moic) }
        ];

        return `
        <article class="dashboard-card entity-performance-card" data-entity="${escapeHtml(entity)}">
          <div class="entity-performance-header">
            <div>
              <p class="dashboard-label">Entity</p>
              <h3>${escapeHtml(entity)}</h3>
            </div>
            <span class="entity-open-pill">Open entity</span>
          </div>
          <div class="entity-metric-grid">
            ${metrics
              .map(
                (metric) => `
                  <div class="entity-metric-box">
                    <p class="dashboard-label">${escapeHtml(metric.label)}</p>
                    <p class="dashboard-value">${escapeHtml(metric.value)}</p>
                  </div>
                `
              )
              .join("")}
          </div>
        </article>
      `;
      }
    )
    .join("");
}

function renderDataQuality() {
  const alerts = buildDataQualityAlerts();
  const highCount = alerts.filter((alert) => alert.severity === "High").length;
  const mediumCount = alerts.filter((alert) => alert.severity === "Medium").length;
  const lowCount = alerts.filter((alert) => alert.severity === "Low").length;

  dataQualitySummary.innerHTML = [
    { label: "Total alerts", value: String(alerts.length) },
    { label: "High priority", value: String(highCount) },
    { label: "Medium priority", value: String(mediumCount) },
    { label: "Low priority", value: String(lowCount) }
  ]
    .map(
      (item) => `
        <article class="dashboard-card">
          <p class="dashboard-label">${escapeHtml(item.label)}</p>
          <p class="dashboard-value">${escapeHtml(item.value)}</p>
        </article>
      `
    )
    .join("");

  dataQualityList.innerHTML = alerts.length
    ? alerts
        .map(
          (alert) => `
            <article class="quality-alert-card quality-alert-${escapeHtml(alert.severity.toLowerCase())}">
              <div class="update-head">
                <div>
                  <p class="dashboard-label">${escapeHtml(alert.entity)} • ${escapeHtml(alert.company)}</p>
                  <h3>${escapeHtml(alert.title)}</h3>
                </div>
                <span class="status-chip">${escapeHtml(alert.severity)}</span>
              </div>
              <p class="update-meta">${escapeHtml(alert.detail)}</p>
              <div class="card-actions">
                <button class="secondary-button card-action-button" type="button" data-action="edit-quality-investment" data-id="${escapeHtml(alert.id)}">Edit investment</button>
              </div>
            </article>
          `
        )
        .join("")
    : '<p class="update-meta">No data quality issues found. The reconciliation gremlins are quiet for now.</p>';
}

function renderEntityDetail() {
  if (!selectedEntity) {
    entityDetailSection.classList.add("hidden");
    entityDetailSummary.innerHTML = "";
    entityDetailInvestments.innerHTML = "";
    entityDetailTitle.textContent = "Entity detail";
    entityDetailCopy.textContent = "";
    return;
  }

  const investments = sortInvestmentsAlphabetically(
    allInvestments.filter(
      (investment) => normalizeEntityName(investment.entity) === normalizeEntityName(selectedEntity)
    )
  );
  const entityRows = buildEntityRows(allInvestments, selectedEntity);
  const performance = buildEntityRowTotals(entityRows);

  entityDetailSection.classList.remove("hidden");
  entityDetailTitle.textContent = selectedEntity;
  const investmentCount = entityRows.length;
  entityDetailCopy.textContent = `${investmentCount} investment${investmentCount === 1 ? "" : "s"} tracked under this entity.`;
  entityDetailSummary.innerHTML = [
    { label: "Total committed capital", value: formatMoney(performance.reportedAmount) },
    { label: "Called capital", value: formatMoney(performance.investedCapital) },
    {
      label: "Unfunded commitment",
      value: formatMoney(Math.max(performance.reportedAmount - performance.investedCapital, 0))
    },
    { label: "Distributions", value: formatMoney(performance.distributions) },
    { label: "Official NAV", value: formatMoney(performance.officialValue) },
    { label: "Internal NAV", value: formatMoney(performance.internalValue) },
    { label: "Official XIRR", value: formatPercent(performance.official.xirr) },
    { label: "Official MOIC", value: formatTurns(performance.official.moic) },
    { label: "Current investments", value: String(investmentCount) }
  ]
    .map(
      (item) => `
        <article class="dashboard-card">
          <p class="dashboard-label">${escapeHtml(item.label)}</p>
          <p class="dashboard-value">${escapeHtml(item.value)}</p>
        </article>
      `
    )
    .join("");

  entityDetailInvestments.innerHTML = investments.length
    ? investments
        .map((investment) => {
          const companyPerformance =
            companyPerformanceMap.get(getInvestmentPositionKey(investment)) ||
            buildCompanyPerformance([investment]);
          return `
            <article class="update-card">
              <div class="update-head">
                <button class="link-button company-link" type="button" data-company="${escapeHtml(investment.company)}" data-entity="${escapeHtml(investment.entity || "")}">
                  ${escapeHtml(investment.company)}
                </button>
                <span class="status-chip">${escapeHtml(normalizeStatusName(investment.status) || "Update")}</span>
              </div>
              <p class="update-meta">${escapeHtml(investment.stage || "Stage not set")} • Owner: ${escapeHtml(investment.owner || "Not set")}</p>
              <p class="update-meta">Official NAV ${escapeHtml(formatMoney(companyPerformance.officialValue))} • XIRR ${escapeHtml(formatPercent(companyPerformance.official.xirr))}</p>
              <p class="update-notes">${escapeHtml(summarizeText(investment.notes, "No notes provided."))}</p>
              <div class="card-actions">
                <button class="secondary-button card-action-button" type="button" data-action="view-company" data-company="${escapeHtml(investment.company)}" data-entity="${escapeHtml(investment.entity || "")}">View company</button>
                <button class="secondary-button card-action-button" type="button" data-action="edit" data-id="${investment.id}">Edit</button>
              </div>
            </article>
          `;
        })
        .join("")
    : '<p class="update-meta">No investments are assigned to this entity yet.</p>';
}

function renderFilterOptions() {
  const entities = Array.from(
    new Set(
      configuredEntities
        .concat(allInvestments.map((item) => normalizeEntityName(item.entity)).filter(Boolean))
        .map(normalizeEntityName)
    )
  ).sort();
  const statuses = CANONICAL_STATUSES;
  const stages = Array.from(new Set(allInvestments.map((item) => item.stage).filter(Boolean))).sort();
  const owners = Array.from(new Set(allInvestments.map((item) => item.owner).filter(Boolean))).sort();

  const assignOptions = (element, placeholder, values) => {
    const currentValue = element.value;
    element.innerHTML = [`<option value="">${placeholder}</option>`]
      .concat(values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`))
      .join("");
    if (element === statusFilter) {
      const normalizedCurrentValue = normalizeStatusName(currentValue);
      element.value = values.includes(normalizedCurrentValue) ? normalizedCurrentValue : "";
      return;
    }

    element.value = values.includes(currentValue) ? currentValue : "";
  };

  assignOptions(entityFilter, "All entities", entities);
  assignOptions(statusFilter, "All statuses", statuses);
  assignOptions(stageFilter, "All stages", stages);
  assignOptions(ownerFilter, "All owners", owners);
}

function renderCompanySuggestions() {
  const companies = Array.from(
    new Set(
      (allCompanies.length ? allCompanies.map((item) => item.company) : allInvestments.map((item) => item.company)).filter(
        Boolean
      )
    )
  ).sort((left, right) => left.localeCompare(right));

  companySuggestions.innerHTML = companies
    .map((company) => `<option value="${escapeHtml(company)}"></option>`)
    .join("");
}

function summarizeText(value, fallback) {
  const text = String(value || "").trim();
  if (!text) {
    return fallback;
  }

  return text.length > 220 ? `${text.slice(0, 217)}...` : text;
}

function sortInvestmentsAlphabetically(investments) {
  return [...investments].sort((left, right) =>
    String(left.company || "").localeCompare(String(right.company || ""), undefined, {
      sensitivity: "base"
    })
  );
}

function beginEditInvestment(investmentId) {
  const investment = allInvestments.find((item) => item.id === investmentId);
  if (!investment) {
    return;
  }

  editingInvestmentId.value = investment.id;
  form.elements.company.value = investment.company || "";
  form.elements.entity.value = investment.entity || "";
  form.elements.amount.value = investment.amount || "";
  form.elements.currency.value = investment.currency || "USD";
  form.elements.stage.value = investment.stage || "";
  form.elements.status.value = normalizeStatusName(investment.status) || "";
  form.elements.owner.value = investment.owner || "";
  form.elements.nextStep.value = investment.nextStep || "";
  form.elements.nextStepDueDate.value = investment.nextStepDueDate || "";
  form.elements.contactName.value = investment.contactName || "";
  form.elements.contactPosition.value = investment.contactPosition || "";
  form.elements.contactEmail.value = investment.contactEmail || "";
  form.elements.contactPhone.value = investment.contactPhone || "";
  form.elements.recipients.value = Array.isArray(investment.recipients)
    ? investment.recipients.join(", ")
    : "";
  notesField.value = investment.notes || "";
  deckSummaryField.value = investment.deckSummary || "";
  renderCapitalActivityRows(
    normalizeCapitalActivityRows(
      investment.capitalActivity && investment.capitalActivity.length
        ? investment.capitalActivity
        : buildLegacyCapitalActivityRows(investment)
    )
  );
  form.elements.valuationDate.value = investment.valuationDate || "";
  form.elements.officialValue.value = investment.officialValue || "";
  form.elements.internalValue.value = investment.internalValue || "";
  form.elements.exitValue.value = investment.exitValue || "";
  form.elements.ownershipPercent.value = investment.ownershipPercent || "";
  form.elements.entityOwnershipPercent.value = investment.entityOwnershipPercent || "";
  form.elements.ownershipNotes.value = investment.ownershipNotes || "";
  form.elements.followOnCapitalAmount.value = investment.followOnCapitalAmount || "";
  form.elements.followOnCapitalStatus.value = investment.followOnCapitalStatus || "";
  form.elements.followOnCapitalNotes.value = investment.followOnCapitalNotes || "";
  form.elements.documentLinks.value = investment.documentLinks || "";
  uploadedDocuments = Array.isArray(investment.documents) ? [...investment.documents] : [];
  renderUploadedDocuments();
  form.elements.decisionDate.value = investment.decisionDate || "";
  form.elements.decisionType.value = investment.decisionType || "";
  form.elements.decisionSummary.value = investment.decisionSummary || "";
  submitButton.textContent = "Save changes";
  cancelEditButton.classList.remove("hidden");
  formMessage.textContent = `Editing ${investment.company}.`;
  applyFormInputFormatting();
  showWorkspaceView("capture");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function resetFormToCreateMode() {
  form.reset();
  editingInvestmentId.value = "";
  submitButton.textContent = "Save update and send email";
  cancelEditButton.classList.add("hidden");
  updateDeckFileLabel(null);
  emailMessage.textContent = "";
  deckMessage.textContent = "";
  deckSummaryField.value = "";
  uploadedDocuments = [];
  renderUploadedDocuments();
  documentMessage.textContent = "";
  renderCapitalActivityRows([]);
  applyFormInputFormatting();
}

function beginEditTask(taskId) {
  const task = allTasks.find((item) => item.id === taskId);
  if (!task) {
    return;
  }

  editingTaskId.value = task.id;
  taskForm.elements.title.value = task.title || "";
  taskForm.elements.company.value = task.company || "";
  taskForm.elements.entity.value = task.entity || "";
  taskForm.elements.dueDate.value = task.dueDate || "";
  taskForm.elements.priority.value = task.priority || "Medium";
  taskForm.elements.category.value = task.category || "";
  taskForm.elements.assignee.value = task.assignee || "";
  taskForm.elements.status.value = task.status || "Open";
  taskForm.elements.description.value = task.description || "";
  saveTaskButton.textContent = "Save task changes";
  cancelTaskEditButton.classList.remove("hidden");
  taskMessage.textContent = `Editing task: ${task.title}`;
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function resetTaskForm() {
  taskForm.reset();
  editingTaskId.value = "";
  saveTaskButton.textContent = "Save task";
  cancelTaskEditButton.classList.add("hidden");
}

async function deleteTaskById(taskId) {
  if (!window.confirm("Delete this task?")) {
    return;
  }

  taskMessage.textContent = "Deleting task...";
  try {
    await fetchJson(`/api/tasks/${taskId}`, { method: "DELETE" });
    if (editingTaskId.value === taskId) {
      resetTaskForm();
    }
    await loadTasks();
    taskMessage.textContent = "Task deleted.";
  } catch (error) {
    taskMessage.textContent = error.message;
  }
}

async function deleteInvestmentById(investmentId) {
  if (!window.confirm("Delete this investment update?")) {
    return;
  }

  formMessage.textContent = "Deleting update...";

  try {
    await fetchJson(`/api/investments/${investmentId}`, { method: "DELETE" });
    if (editingInvestmentId.value === investmentId) {
      resetFormToCreateMode();
    }
    await loadUpdates();
    formMessage.textContent = "Investment update deleted.";
  } catch (error) {
    formMessage.textContent = error.message;
  }
}

function buildReconciliationCapitalActivity(investment, investedCapitalValue) {
  const normalizedAmount = normalizeMoneyString(investedCapitalValue);
  const existingRows = normalizeCapitalActivityRows(
    investment.capitalActivity && investment.capitalActivity.length
      ? investment.capitalActivity
      : buildLegacyCapitalActivityRows(investment)
  );

  const contributionRows = existingRows.filter((row) => {
    const type = String(row.type || "").toLowerCase();
    return (
      type.includes("capital call") ||
      type.includes("investment amount") ||
      type.includes("fee")
    );
  });

  const nonContributionRows = existingRows.filter((row) => {
    const type = String(row.type || "").toLowerCase();
    return !(
      type.includes("capital call") ||
      type.includes("investment amount") ||
      type.includes("fee")
    );
  });

  const contributionDate =
    contributionRows
      .map((row) => row.date)
      .filter(Boolean)
      .sort()[0] || investment.capitalCallDate || investment.createdAt || "";

  return normalizedAmount
    ? [
        {
          date: contributionDate,
          type: "Investment Amount",
          amount: normalizedAmount,
          notes: "Updated from entity reconciliation"
        }
      ].concat(nonContributionRows)
    : nonContributionRows;
}

function getReconciliationInputs(investmentId) {
  return Array.from(
    reconciliationList.querySelectorAll(
      `[data-edit-input="true"][data-id="${CSS.escape(investmentId)}"]`
    )
  );
}

function getReconciliationValues(investmentId) {
  const inputs = getReconciliationInputs(investmentId);
  return inputs.reduce((result, input) => {
    result[input.dataset.field] = input.value;
    return result;
  }, {});
}

function setReconciliationRowDirty(investmentId, isDirty = true) {
  const row = reconciliationList.querySelector(`[data-reconciliation-row-id="${CSS.escape(investmentId)}"]`);
  if (row) {
    row.classList.toggle("reconciliation-dirty-row", isDirty);
  }

  if (isDirty) {
    dirtyReconciliationRows.add(investmentId);
  } else {
    dirtyReconciliationRows.delete(investmentId);
  }

  if (saveAllReconciliationButton) {
    saveAllReconciliationButton.disabled = !dirtyReconciliationRows.size || savingAllReconciliation;
  }

  if (reconciliationMessage && !savingAllReconciliation) {
    reconciliationMessage.textContent = dirtyReconciliationRows.size
      ? `${dirtyReconciliationRows.size} unsaved reconciliation change${dirtyReconciliationRows.size === 1 ? "" : "s"}.`
      : "All reconciliation changes saved.";
  }
}

async function saveReconciliationRows(investmentIds) {
  const ids = Array.from(new Set(investmentIds)).filter(Boolean);
  if (!ids.length) {
    if (reconciliationMessage) {
      reconciliationMessage.textContent = "No reconciliation changes to save.";
    }
    return;
  }

  savingAllReconciliation = true;
  if (saveAllReconciliationButton) {
    saveAllReconciliationButton.disabled = true;
    saveAllReconciliationButton.textContent = "Saving...";
  }
  if (reconciliationMessage) {
    reconciliationMessage.textContent = `Saving ${ids.length} reconciliation change${ids.length === 1 ? "" : "s"}...`;
  }

  try {
    for (const investmentId of ids) {
      const values = getReconciliationValues(investmentId);
      await saveReconciliationRow(investmentId, values, { reload: false });
    }

    await loadUpdates();
    ids.forEach((investmentId) => setReconciliationRowDirty(investmentId, false));

    if (reconciliationMessage) {
      reconciliationMessage.textContent = `${ids.length} reconciliation change${ids.length === 1 ? "" : "s"} saved.`;
    }
  } finally {
    savingAllReconciliation = false;
    if (saveAllReconciliationButton) {
      saveAllReconciliationButton.textContent = "Save all changes";
      saveAllReconciliationButton.disabled = !dirtyReconciliationRows.size;
    }
  }
}

async function saveReconciliationRow(investmentId, values, options = {}) {
  const investment = allInvestments.find((item) => item.id === investmentId);
  if (!investment) {
    return;
  }

  const reportedAmount = normalizeMoneyString(values.reportedAmount);
  const investedCapital = normalizeMoneyString(values.investedCapital);
  const officialValue = normalizeMoneyString(values.officialValue);
  const internalValue = normalizeMoneyString(values.internalValue);

  const payload = {
    company: String(values.company || investment.company || "").trim(),
    entity: values.entity || investment.entity,
    amount: reportedAmount,
    currency: investment.currency || "USD",
    stage: values.stage || "",
    status: normalizeStatusName(values.status) || "",
    owner: values.owner || investment.owner || "",
    nextStep: investment.nextStep || "",
    notes: investment.notes || "",
    deckSummary: investment.deckSummary || "",
    capitalActivity: buildReconciliationCapitalActivity(investment, investedCapital),
    capitalCallDate: investment.capitalCallDate || "",
    capitalCallAmount: investment.capitalCallAmount || "",
    distributionDate: investment.distributionDate || "",
    distributionAmount: investment.distributionAmount || "",
    valuationDate: investment.valuationDate || "",
    officialValue,
    internalValue,
    exitValue: investment.exitValue || "",
    ownershipPercent: investment.ownershipPercent || "",
    entityOwnershipPercent: investment.entityOwnershipPercent || "",
    ownershipNotes: investment.ownershipNotes || "",
    followOnCapitalAmount: investment.followOnCapitalAmount || "",
    followOnCapitalStatus: investment.followOnCapitalStatus || "",
    followOnCapitalNotes: investment.followOnCapitalNotes || "",
    contactName: investment.contactName || "",
    contactPosition: investment.contactPosition || "",
    contactEmail: investment.contactEmail || "",
    contactPhone: investment.contactPhone || "",
    documentLinks: investment.documentLinks || "",
    documents: Array.isArray(investment.documents) ? investment.documents : [],
    decisionDate: investment.decisionDate || "",
    decisionType: investment.decisionType || "",
    decisionSummary: investment.decisionSummary || "",
    recipients: Array.isArray(investment.recipients) ? investment.recipients : []
  };

  await fetchJson(`/api/investments/${investmentId}`, {
    method: "PATCH",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });

  if (options.reload !== false) {
    await loadUpdates();
  }
}

function renderCompanyPanel() {
  if (!selectedCompany) {
    companyPanel.classList.add("hidden");
    companySummary.innerHTML = "";
    companyHighlights.innerHTML = "";
    companyContactInfo.innerHTML = "";
    companyPerformanceSummary.innerHTML = "";
    companyEntityPerformance.innerHTML = "";
    companyOwnershipSummary.innerHTML = "";
    companyDeckSummaries.innerHTML = "";
    companyDecisionLog.innerHTML = "";
    companyNextSteps.innerHTML = "";
    companyTasks.innerHTML = "";
    companyFollowOnCapital.innerHTML = "";
    companyValuationHistory.innerHTML = "";
    companyTimeline.innerHTML = "";
    companyDocumentMessage.textContent = "";
    return;
  }

  const companyRecord = findCompanyRecord(selectedCompany, selectedCompanyEntity);
  const companyUpdates = companyRecord
    ? [...companyRecord.updates]
    : allInvestments
        .filter(
          (investment) =>
            companyKey(investment.company) === companyKey(selectedCompany) &&
            normalizeEntityName(investment.entity) === normalizeEntityName(selectedCompanyEntity)
        )
        .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

  if (!companyUpdates.length) {
    companyPanel.classList.add("hidden");
    return;
  }

  companyDocumentMessage.textContent = "";

  const latest = companyUpdates[0];
  const earliest = companyUpdates[companyUpdates.length - 1];
  const totalAmount = companyUpdates.reduce((sum, investment) => sum + toNumber(investment.amount), 0);
  const uniqueOwners = Array.from(
    new Set(companyUpdates.map((investment) => investment.owner).filter(Boolean))
  );
  const uniqueStatuses = Array.from(
    new Set(companyUpdates.map((investment) => normalizeStatusName(investment.status)).filter(Boolean))
  );
  const nextSteps = Array.from(
    new Set(companyUpdates.map((investment) => investment.nextStep).filter(Boolean))
  );
  const deckSummaries = companyRecord
    ? companyRecord.researchEntries.filter((entry) => entry.type === "Deck Summary")
    : companyUpdates.filter((investment) => investment.deckSummary);
  const followOnUpdates = companyRecord
    ? companyRecord.followOnHistory
    : companyUpdates.filter(
        (investment) =>
          investment.followOnCapitalAmount ||
          investment.followOnCapitalStatus ||
          investment.followOnCapitalNotes
      );
  const decisionUpdates = companyRecord
    ? companyRecord.decisionLog
    : companyUpdates.filter(
        (investment) =>
          investment.documentLinks ||
          investment.decisionDate ||
          investment.decisionType ||
          investment.decisionSummary
      );
  const relatedTasks = companyRecord
    ? companyRecord.tasks
    : allTasks
        .filter(
          (task) =>
            companyKey(task.company) === companyKey(selectedCompany) &&
            normalizeEntityName(task.entity) === normalizeEntityName(selectedCompanyEntity)
        )
        .sort(
          (left, right) =>
            new Date(left.dueDate || left.createdAt) - new Date(right.dueDate || right.createdAt)
        );
  const valuationUpdates = companyRecord
    ? companyRecord.valuationHistory
    : companyUpdates.filter(
        (investment) =>
          investment.valuationDate ||
          investment.officialValue ||
          investment.internalValue ||
          investment.exitValue
      );
  const ownershipUpdates = companyRecord
    ? companyRecord.ownershipHistory
    : companyUpdates.filter(
        (investment) =>
          investment.ownershipPercent ||
          investment.entityOwnershipPercent ||
          investment.ownershipNotes
      );
  const perEntityCompanyPerformance = Array.from(
    new Set(companyUpdates.map((investment) => normalizeEntityName(investment.entity)).filter(Boolean))
  ).map((entity) => ({
    entity,
    performance: buildCompanyPerformance(
      companyUpdates.filter(
        (investment) => normalizeEntityName(investment.entity) === normalizeEntityName(entity)
      )
    )
  }));
  companyPanel.classList.remove("hidden");
  const performance =
    companyPerformanceMap.get(companyEntityKey(selectedCompany, selectedCompanyEntity)) ||
    buildCompanyPerformance(companyUpdates);
  companyPanelTitle.textContent = latest.company || selectedCompany;
  companyPanelCopy.textContent = companyRecord
    ? `${companyUpdates.length} update${companyUpdates.length === 1 ? "" : "s"} organized into structured research, capital, valuation, decision, and document records.`
    : `${companyUpdates.length} update${companyUpdates.length === 1 ? "" : "s"} saved for this company.`;
  companySummary.innerHTML = [
    { label: "Latest status", value: normalizeStatusName(latest.status) || "Not set" },
    { label: "Latest entity", value: normalizeEntityName(latest.entity) || "Not set" },
    { label: "Latest stage", value: latest.stage || "Not set" },
    { label: "Latest owner", value: latest.owner || "Not set" },
    { label: "Reported amount", value: formatMoney(totalAmount) }
  ]
    .map(
      (item) => `
        <article class="company-summary-card">
          <p class="dashboard-label">${escapeHtml(item.label)}</p>
          <p class="dashboard-value">${escapeHtml(item.value)}</p>
        </article>
      `
    )
    .join("");

  companyHighlights.innerHTML = [
    { label: "First entered", value: earliest.createdAt || "Unknown" },
    { label: "Latest update", value: latest.createdAt || "Unknown" },
    { label: "Submitted by", value: latest.submittedBy || "Unknown" },
    {
      label: "Entities used",
      value:
        Array.from(
          new Set(companyUpdates.map((investment) => normalizeEntityName(investment.entity)).filter(Boolean))
        ).join(", ") || "Not set"
    },
    { label: "Owners involved", value: uniqueOwners.length ? uniqueOwners.join(", ") : "Not set" },
    { label: "Statuses used", value: uniqueStatuses.length ? uniqueStatuses.join(", ") : "Not set" },
    { label: "Latest notes", value: latest.notes || "No notes provided." },
    {
      label: "Latest deck summary",
      value: summarizeText(latest.deckSummary, "No deck summary provided yet.")
    }
  ]
    .map(
      (item) => `
        <div class="highlight-row">
          <p class="dashboard-label">${escapeHtml(item.label)}</p>
          <p class="highlight-value">${escapeHtml(item.value)}</p>
        </div>
      `
    )
    .join("");

  companyContactInfo.innerHTML = [
    { label: "Name", value: latest.contactName || "Not set" },
    { label: "Position", value: latest.contactPosition || "Not set" },
    { label: "Email", value: latest.contactEmail || "Not set" },
    { label: "Phone", value: latest.contactPhone || "Not set" }
  ]
    .map(
      (item) => `
        <div class="highlight-row">
          <p class="dashboard-label">${escapeHtml(item.label)}</p>
          <p class="highlight-value">${escapeHtml(item.value)}</p>
        </div>
      `
    )
    .join("");

  companyPerformanceSummary.innerHTML = [
    { label: "Called capital", value: formatMoney(performance.investedCapital) },
    { label: "Distributions", value: formatMoney(performance.distributions) },
    { label: "Official value", value: formatMoney(performance.officialValue) },
    { label: "Internal value", value: formatMoney(performance.internalValue) },
    { label: "Official MOIC", value: formatTurns(performance.official.moic) },
    { label: "Internal MOIC", value: formatTurns(performance.internal.moic) },
    { label: "Official XIRR", value: formatPercent(performance.official.xirr) },
    { label: "Internal XIRR", value: formatPercent(performance.internal.xirr) },
    { label: "Exit XIRR", value: formatPercent(performance.exit.xirr) }
  ]
    .map(
      (item) => `
        <div class="highlight-row">
          <p class="dashboard-label">${escapeHtml(item.label)}</p>
          <p class="highlight-value">${escapeHtml(item.value)}</p>
        </div>
      `
    )
    .join("");

  companyEntityPerformance.innerHTML = perEntityCompanyPerformance.length
    ? perEntityCompanyPerformance
        .map(
          ({ entity, performance }) => `
            <div class="highlight-row">
              <p class="dashboard-label">${escapeHtml(entity)}</p>
              <p class="highlight-value">Invested ${escapeHtml(formatMoney(performance.investedCapital))} • Official NAV ${escapeHtml(
                formatMoney(performance.officialValue)
              )}</p>
              <p class="update-meta">XIRR ${escapeHtml(formatPercent(performance.official.xirr))} • MOIC ${escapeHtml(
                formatTurns(performance.official.moic)
              )}</p>
            </div>
          `
        )
        .join("")
    : '<p class="update-meta">No entity-level performance yet.</p>';

  companyOwnershipSummary.innerHTML = ownershipUpdates.length
    ? ownershipUpdates
        .slice(0, 6)
        .map(
          (investment) => `
            <div class="highlight-row">
              <p class="dashboard-label">${escapeHtml(investment.entity || "Entity not set")}</p>
              <p class="highlight-value">Total ${escapeHtml(investment.totalPercent || investment.ownershipPercent || "Not set")}${
                investment.totalPercent || investment.ownershipPercent ? "%" : ""
              } • Entity ${escapeHtml(investment.entityPercent || investment.entityOwnershipPercent || "Not set")}${
                investment.entityPercent || investment.entityOwnershipPercent ? "%" : ""
              }</p>
              <p class="update-meta">${escapeHtml(investment.notes || investment.ownershipNotes || "No ownership notes.")}</p>
            </div>
          `
        )
        .join("")
    : '<p class="update-meta">No ownership allocation notes yet.</p>';

  companyNextSteps.innerHTML = nextSteps.length
    ? `<ul class="company-list">${nextSteps
        .slice(0, 6)
        .map((nextStep) => `<li>${escapeHtml(nextStep)}</li>`)
        .join("")}</ul>`
    : '<p class="update-meta">No next steps recorded yet.</p>';

  companyTasks.innerHTML = relatedTasks.length
    ? relatedTasks
        .map(
          (task) => `
            <article class="timeline-card timeline-card-compact">
              <p class="dashboard-label">${escapeHtml(task.title)}</p>
              <p class="update-meta">${escapeHtml(task.status)} • Due ${escapeHtml(task.dueDate || "not set")} • ${escapeHtml(task.assignee || "Unassigned")}</p>
              <p class="highlight-value">${escapeHtml(task.description || "No task details.")}</p>
            </article>
          `
        )
        .join("")
    : '<p class="update-meta">No company tasks yet.</p>';

  companyDeckSummaries.innerHTML = deckSummaries.length
    ? deckSummaries
        .slice(0, 4)
        .map(
          (investment) => `
            <article class="timeline-card timeline-card-compact">
              <p class="dashboard-label">${escapeHtml(investment.date || investment.createdAt)}</p>
              <p class="highlight-value">${escapeHtml(investment.summary || investment.deckSummary)}</p>
            </article>
          `
        )
        .join("")
    : '<p class="update-meta">No deck summaries saved yet.</p>';

  companyDecisionLog.innerHTML = decisionUpdates.length
    ? decisionUpdates
        .slice(0, 6)
        .map(
          (investment) => `
            <article class="timeline-card timeline-card-compact">
              <p class="dashboard-label">${escapeHtml(investment.date || investment.decisionDate || investment.createdAt)}</p>
              <p class="highlight-value">${escapeHtml(investment.type || investment.decisionType || "Decision not set")}</p>
              <p class="update-meta">${escapeHtml(investment.summary || investment.decisionSummary || "No decision summary.")}</p>
              ${
                investment.documentLinks
                  ? `<p class="update-meta">${escapeHtml(investment.documentLinks)}</p>`
                  : ""
              }
              ${companyRecord && Array.isArray(companyRecord.documents) && companyRecord.documents.length
                  ? `<div class="document-pill-row">${companyRecord.documents
                      .map(
                        (document) =>
                          document.source === "company-vault" && canEditWorkspace()
                            ? `<span class="document-pill document-pill-with-action"><a class="document-pill-link" href="${escapeHtml(document.url)}" target="_blank" rel="noreferrer">${escapeHtml(document.name)}</a><button type="button" class="document-pill-remove" data-action="delete-company-document" data-document-id="${escapeHtml(document.id)}">Remove</button></span>`
                            : `<a class="document-pill" href="${escapeHtml(document.url)}" target="_blank" rel="noreferrer">${escapeHtml(document.name)}</a>`
                      )
                      .join("")}</div>`
                  : ""
              }
            </article>
          `
        )
        .join("")
    : '<p class="update-meta">No linked documents or decisions yet.</p>';

  companyFollowOnCapital.innerHTML = followOnUpdates.length
    ? followOnUpdates
        .slice(0, 4)
        .map(
          (investment) => `
            <div class="highlight-row">
              <p class="dashboard-label">${escapeHtml(investment.date || investment.createdAt)}</p>
              <p class="highlight-value">${
                investment.amount || investment.followOnCapitalAmount
                  ? `${escapeHtml(investment.currency || latest.currency)} ${escapeHtml(investment.amount || investment.followOnCapitalAmount)}`
                  : "Amount not set"
              } • ${escapeHtml(investment.type || investment.followOnCapitalStatus || "Status not set")}</p>
              <p class="update-meta">${escapeHtml(investment.notes || investment.followOnCapitalNotes || "No follow-on notes.")}</p>
            </div>
          `
        )
        .join("")
    : '<p class="update-meta">No follow-on capital entries yet.</p>';

  companyValuationHistory.innerHTML = valuationUpdates.length
    ? valuationUpdates
        .map(
          (investment) => `
            <article class="timeline-card timeline-card-compact">
              <p class="dashboard-label">${escapeHtml(
                investment.date || investment.valuationDate || investment.createdAt
              )}</p>
              <p class="highlight-value">Official ${escapeHtml(
                investment.officialValue ? `${investment.currency || latest.currency} ${investment.officialValue}` : "not set"
              )}</p>
              <p class="update-meta">Internal ${escapeHtml(
                investment.internalValue ? `${investment.currency || latest.currency} ${investment.internalValue}` : "not set"
              )} • Exit ${escapeHtml(
                investment.exitValue ? `${investment.currency || latest.currency} ${investment.exitValue}` : "not set"
              )}</p>
            </article>
          `
        )
        .join("")
    : '<p class="update-meta">No valuation history yet.</p>';

  companyTimeline.innerHTML = companyUpdates
    .map(
      (investment) => `
        <article class="timeline-card">
          <div class="update-head">
            <h3>${escapeHtml(normalizeStatusName(investment.status) || "Update")}</h3>
            <span class="status-chip">${escapeHtml(investment.stage || "No stage")}</span>
          </div>
          <p class="update-meta">
            ${escapeHtml(investment.createdAt)} • ${escapeHtml(investment.entity || "No entity")} • Owner: ${escapeHtml(investment.owner || "Not set")}
          </p>
          <p class="update-meta">
            Amount: ${
              investment.amount
                ? `${escapeHtml(investment.currency)} ${escapeHtml(investment.amount)}`
                : "Amount not specified"
            } • Submitted by: ${escapeHtml(investment.submittedBy || "Unknown")}
          </p>
          ${
            summarizeCapitalActivity(investment)
              ? `<p class="update-meta">Cash activity: ${escapeHtml(summarizeCapitalActivity(investment))}</p>`
              : ""
          }
          ${
            investment.officialValue || investment.internalValue || investment.exitValue
              ? `<p class="update-meta">Marks: Official ${
                  investment.officialValue
                    ? `${escapeHtml(investment.currency)} ${escapeHtml(investment.officialValue)}`
                    : "not set"
                } • Internal ${
                  investment.internalValue
                    ? `${escapeHtml(investment.currency)} ${escapeHtml(investment.internalValue)}`
                    : "not set"
                } • Exit ${
                  investment.exitValue
                    ? `${escapeHtml(investment.currency)} ${escapeHtml(investment.exitValue)}`
                    : "not set"
                }</p>`
              : ""
          }
          ${
            investment.ownershipPercent || investment.entityOwnershipPercent || investment.ownershipNotes
              ? `<p class="update-meta">Ownership: Total ${
                  escapeHtml(investment.ownershipPercent || "Not set")
                }${investment.ownershipPercent ? "%" : ""} • Entity ${escapeHtml(
                  investment.entityOwnershipPercent || "Not set"
                )}${investment.entityOwnershipPercent ? "%" : ""}</p>`
              : ""
          }
          <p class="update-meta">
            Next: ${escapeHtml(investment.nextStep || "No next step provided")}
          </p>
          <p class="update-notes">${escapeHtml(investment.notes || "No notes provided.")}</p>
          ${
            investment.deckSummary
              ? `<div class="update-subsection"><p class="dashboard-label">Deck summary</p><p class="update-notes">${escapeHtml(investment.deckSummary)}</p></div>`
              : ""
          }
          ${
            investment.followOnCapitalAmount || investment.followOnCapitalStatus || investment.followOnCapitalNotes
              ? `<div class="update-subsection"><p class="dashboard-label">Follow-on capital</p><p class="update-meta">${
                  investment.followOnCapitalAmount
                    ? `${escapeHtml(investment.currency)} ${escapeHtml(investment.followOnCapitalAmount)}`
                    : "Amount not set"
                } • ${escapeHtml(investment.followOnCapitalStatus || "Status not set")}</p><p class="update-meta">${escapeHtml(investment.followOnCapitalNotes || "No follow-on notes.")}</p></div>`
              : ""
          }
          ${
            investment.decisionDate || investment.decisionType || investment.decisionSummary || investment.documentLinks
              ? `<div class="update-subsection"><p class="dashboard-label">Decision log</p><p class="update-meta">${
                  escapeHtml(investment.decisionDate || "No decision date")
                } • ${escapeHtml(investment.decisionType || "No decision type")}</p><p class="update-meta">${escapeHtml(
                  investment.decisionSummary || "No decision summary."
                )}</p>${
                  investment.documentLinks
                    ? `<p class="update-meta">${escapeHtml(investment.documentLinks)}</p>`
                    : ""
                }${
                  Array.isArray(investment.documents) && investment.documents.length
                    ? `<div class="document-pill-row">${investment.documents
                        .map(
                          (document) =>
                            `<a class="document-pill" href="${escapeHtml(document.url)}" target="_blank" rel="noreferrer">${escapeHtml(document.name)}</a>`
                        )
                        .join("")}</div>`
                    : ""
                }</div>`
              : ""
          }
        </article>
      `
    )
    .join("");
}

function renderUpdates(investments) {
  if (!investments.length) {
    updatesList.innerHTML =
      '<p class="update-meta">No updates yet. Your first submission will show up here.</p>';
    return;
  }

  updatesList.innerHTML = sortInvestmentsAlphabetically(investments)
    .map((investment) => {
      const performance =
        companyPerformanceMap.get(getInvestmentPositionKey(investment)) ||
        buildCompanyPerformance([investment]);
      const amount = investment.amount
        ? `${escapeHtml(investment.currency)} ${escapeHtml(investment.amount)}`
        : "Amount not specified";

      return `
        <article class="update-card">
          <div class="update-head">
            <button class="link-button company-link" type="button" data-company="${escapeHtml(investment.company)}" data-entity="${escapeHtml(investment.entity || "")}">
              ${escapeHtml(investment.company)}
            </button>
            <span class="status-chip">${escapeHtml(normalizeStatusName(investment.status) || "Update")}</span>
          </div>
          <p class="update-meta">
            ${escapeHtml(investment.entity || "Entity not specified")} • ${amount} • ${escapeHtml(investment.stage || "Stage not specified")}
          </p>
          <p class="update-meta">
            Owner: ${escapeHtml(investment.owner || "Not set")} • Submitted by:
            ${escapeHtml(investment.submittedBy || "Unknown")}
          </p>
          <p class="update-meta">
            Official XIRR: ${escapeHtml(formatPercent(performance.official.xirr))} • Official MOIC: ${escapeHtml(
              formatTurns(performance.official.moic)
            )}
          </p>
          ${
            summarizeCapitalActivity(investment)
              ? `<p class="update-meta">Cash activity: ${escapeHtml(summarizeCapitalActivity(investment))}</p>`
              : ""
          }
          <p class="update-meta">
            Next: ${escapeHtml(investment.nextStep || "Not set")}
          </p>
          <p class="update-notes">${escapeHtml(investment.notes || "No notes provided.")}</p>
          ${
            investment.deckSummary
              ? `<p class="update-meta"><strong>Deck summary:</strong> ${escapeHtml(summarizeText(investment.deckSummary, ""))}</p>`
              : ""
          }
          ${
            investment.followOnCapitalAmount || investment.followOnCapitalStatus
              ? `<p class="update-meta"><strong>Follow-on:</strong> ${
                  investment.followOnCapitalAmount
                    ? `${escapeHtml(investment.currency)} ${escapeHtml(investment.followOnCapitalAmount)}`
                    : "Amount not set"
                } • ${escapeHtml(investment.followOnCapitalStatus || "Status not set")}</p>`
              : ""
          }
          ${
            investment.decisionType || investment.decisionSummary
              ? `<p class="update-meta"><strong>Decision:</strong> ${escapeHtml(
                  investment.decisionType || "Decision not set"
                )} • ${escapeHtml(investment.decisionSummary || "No summary")}</p>`
              : ""
          }
          ${
            Array.isArray(investment.documents) && investment.documents.length
              ? `<div class="document-pill-row">${investment.documents
                  .slice(0, 3)
                  .map(
                    (document) =>
                      `<a class="document-pill" href="${escapeHtml(document.url)}" target="_blank" rel="noreferrer">${escapeHtml(document.name)}</a>`
                  )
                  .join("")}</div>`
              : ""
          }
          <div class="card-actions">
            <button class="secondary-button card-action-button" type="button" data-action="view-company" data-company="${escapeHtml(investment.company)}" data-entity="${escapeHtml(investment.entity || "")}">View company</button>
            <button class="secondary-button card-action-button" type="button" data-action="edit" data-id="${investment.id}">Edit</button>
            <button class="secondary-button card-action-button danger-button" type="button" data-action="delete" data-id="${investment.id}">Delete</button>
          </div>
        </article>
      `;
    })
    .join("");
}

function renderTasks() {
  const filteredTasks = allTasks.filter((task) => {
    const filters = currentFilters();
    const matchesEntity = !filters.entity || task.entity === filters.entity;
    const matchesSearch =
      !filters.search ||
      [task.title, task.company, task.description, task.assignee, task.category]
        .join(" ")
        .toLowerCase()
        .includes(filters.search);
    const matchesOwner = !filters.owner || task.assignee === filters.owner;
    return matchesEntity && matchesSearch && matchesOwner;
  });

  if (!filteredTasks.length) {
    tasksList.innerHTML = '<p class="update-meta">No tasks yet.</p>';
    return;
  }

  tasksList.innerHTML = filteredTasks
    .map(
      (task) => `
        <article class="update-card">
          <div class="update-head">
            <h3>${escapeHtml(task.title)}</h3>
            <span class="status-chip">${escapeHtml(task.status)}</span>
          </div>
          <p class="update-meta">${escapeHtml(task.company || "General")} • ${escapeHtml(task.entity || "No entity")} • ${escapeHtml(task.priority)}</p>
          <p class="update-meta">Due ${escapeHtml(task.dueDate || "not set")} • Assignee: ${escapeHtml(task.assignee || "Not set")} • Category: ${escapeHtml(task.category || "General")}</p>
          <p class="update-notes">${escapeHtml(task.description || "No task details.")}</p>
          ${
            canEditWorkspace()
              ? `<div class="card-actions">
                  <button class="secondary-button card-action-button" type="button" data-action="edit-task" data-task-id="${escapeHtml(task.id)}">Edit</button>
                  <button class="secondary-button card-action-button danger-button" type="button" data-action="delete-task" data-task-id="${escapeHtml(task.id)}">Delete</button>
                </div>`
              : ""
          }
        </article>
      `
    )
    .join("");
}

function renderResearchLibrary(investments) {
  const latestDecks = investments.filter((investment) => investment.deckSummary).slice(0, 6);
  const latestNotes = investments.filter((investment) => investment.notes).slice(0, 6);
  const latestDocuments = investments
    .filter((investment) => Array.isArray(investment.documents) && investment.documents.length)
    .slice(0, 6);
  const latestDecisions = investments
    .filter((investment) => investment.decisionType || investment.decisionSummary)
    .slice(0, 6);

  researchDeckFeed.innerHTML = latestDecks.length
    ? latestDecks
        .map(
          (investment) => `
            <article class="timeline-card timeline-card-compact">
              <p class="dashboard-label">${escapeHtml(investment.company)} • ${escapeHtml(formatDisplayDate(investment.createdAt))}</p>
              <p class="highlight-value">${escapeHtml(summarizeText(investment.deckSummary, ""))}</p>
            </article>
          `
        )
        .join("")
    : '<p class="update-meta">No deck summaries yet.</p>';

  researchNotesFeed.innerHTML = latestNotes.length
    ? latestNotes
        .map(
          (investment) => `
            <article class="timeline-card timeline-card-compact">
              <p class="dashboard-label">${escapeHtml(investment.company)} • ${escapeHtml(normalizeStatusName(investment.status) || "Update")}</p>
              <p class="highlight-value">${escapeHtml(summarizeText(investment.notes, "No notes"))}</p>
            </article>
          `
        )
        .join("")
    : '<p class="update-meta">No note entries yet.</p>';

  researchDocumentsFeed.innerHTML = latestDocuments.length
    ? latestDocuments
        .map(
          (investment) => `
            <article class="timeline-card timeline-card-compact">
              <p class="dashboard-label">${escapeHtml(investment.company)}</p>
              <div class="document-pill-row">${investment.documents
                .map(
                  (document) =>
                    `<a class="document-pill" href="${escapeHtml(document.url)}" target="_blank" rel="noreferrer">${escapeHtml(document.name)}</a>`
                )
                .join("")}</div>
            </article>
          `
        )
        .join("")
    : '<p class="update-meta">No uploaded documents yet.</p>';

  researchDecisionFeed.innerHTML = latestDecisions.length
    ? latestDecisions
        .map(
          (investment) => `
            <article class="timeline-card timeline-card-compact">
              <p class="dashboard-label">${escapeHtml(investment.company)} • ${escapeHtml(
                investment.decisionDate || investment.createdAt
              )}</p>
              <p class="highlight-value">${escapeHtml(investment.decisionType || "Decision not set")}</p>
              <p class="update-meta">${escapeHtml(investment.decisionSummary || "No decision summary.")}</p>
            </article>
          `
        )
        .join("")
    : '<p class="update-meta">No decisions recorded yet.</p>';
}

function renderReconciliation() {
  dirtyReconciliationRows = new Set();
  if (reconciliationMessage && !savingAllReconciliation) {
    reconciliationMessage.textContent = canEditWorkspace()
      ? "No unsaved reconciliation changes."
      : "Review the rollup by entity.";
  }
  if (saveAllReconciliationButton) {
    saveAllReconciliationButton.disabled = true;
  }

  const companySummaries = getCompanyCollections(allInvestments);
  const entities = Array.from(
    new Set(
      configuredEntities
        .concat(companySummaries.map((company) => normalizeEntityName(company.latest.entity)).filter(Boolean))
        .map(normalizeEntityName)
    )
  ).filter(Boolean);

  if (!entities.length) {
    reconciliationList.innerHTML =
      '<p class="update-meta">No entity data is available yet.</p>';
    return;
  }

  const statusOptions = CANONICAL_STATUSES;
  const entityOptions = Array.from(
    new Set(configuredEntities.concat(allInvestments.map((investment) => normalizeEntityName(investment.entity)).filter(Boolean)))
  )
    .map(normalizeEntityName)
    .filter(Boolean)
    .sort((left, right) => left.localeCompare(right));

  reconciliationList.innerHTML = entities
    .map((entity) => {
      const entityRows = buildEntityRows(allInvestments, entity)
        .sort((left, right) =>
          String(left.latest.company || "").localeCompare(String(right.latest.company || ""), undefined, {
            sensitivity: "base"
          })
        );
      const entityPerformance = buildEntityRowTotals(entityRows);

      return `
        <section class="reconciliation-card">
          <div class="panel-header">
            <div>
              <h3>${escapeHtml(entity)}</h3>
              <p class="section-copy">${entityRows.length} investment${entityRows.length === 1 ? "" : "s"} included in this entity total.</p>
            </div>
          </div>
          <table class="reconciliation-table">
            <thead>
              <tr>
                <th>Company</th>
                <th>Entity</th>
                <th>Stage</th>
                <th>Status</th>
                <th>Owner</th>
                <th>Reported amount</th>
                <th>Included committed</th>
                <th>Called capital</th>
                <th>Official NAV</th>
                <th>Internal NAV</th>
                ${canEditWorkspace() ? "<th>Edit</th>" : ""}
              </tr>
            </thead>
            <tbody>
              ${entityRows.length
                ? entityRows
                    .map(
                      ({ latest, performance, includedReportedAmount, includeReportedAmount }) => `
                        <tr data-reconciliation-row-id="${escapeHtml(latest.id)}">
                          <td>
                            ${
                              canEditWorkspace()
                                ? `<input class="reconciliation-amount-input" type="text" value="${escapeHtml(
                                    latest.company || ""
                                  )}" data-edit-input="true" data-field="company" data-id="${escapeHtml(latest.id)}" aria-label="Company name for ${escapeHtml(latest.company)}" />
                                  <button class="link-button company-link" type="button" data-company="${escapeHtml(latest.company)}" data-entity="${escapeHtml(latest.entity || "")}">Open</button>`
                                : `<button class="link-button company-link" type="button" data-company="${escapeHtml(latest.company)}" data-entity="${escapeHtml(latest.entity || "")}">
                                    ${escapeHtml(latest.company)}
                                  </button>`
                            }
                          </td>
                          <td>
                            ${
                              canEditWorkspace()
                                ? `<select class="reconciliation-amount-input" data-edit-input="true" data-field="entity" data-id="${escapeHtml(latest.id)}" aria-label="Entity for ${escapeHtml(latest.company)}">
                                    ${entityOptions
                                      .map(
                                        (entityOption) =>
                                          `<option value="${escapeHtml(entityOption)}" ${
                                            normalizeEntityName(entityOption) === normalizeEntityName(latest.entity || "")
                                              ? "selected"
                                              : ""
                                          }>${escapeHtml(entityOption)}</option>`
                                      )
                                      .join("")}
                                  </select>`
                                : escapeHtml(normalizeEntityName(latest.entity) || "Not set")
                            }
                          </td>
                          <td>
                            ${
                              canEditWorkspace()
                                ? `<input class="reconciliation-amount-input" type="text" value="${escapeHtml(
                                    latest.stage || ""
                                  )}" data-edit-input="true" data-field="stage" data-id="${escapeHtml(latest.id)}" aria-label="Stage for ${escapeHtml(latest.company)}" />`
                                : escapeHtml(latest.stage || "Not set")
                            }
                          </td>
                          <td>
                            ${
                              canEditWorkspace()
                                ? `<select class="reconciliation-amount-input" data-edit-input="true" data-field="status" data-id="${escapeHtml(latest.id)}" aria-label="Status for ${escapeHtml(latest.company)}">
                                    ${[""].concat(statusOptions)
                                      .map(
                                        (status) =>
                                          `<option value="${escapeHtml(status)}" ${
                                            statusEquals(status, latest.status || "") ? "selected" : ""
                                          }>${escapeHtml(status || "Select status")}</option>`
                                      )
                                      .join("")}
                                  </select>`
                                : escapeHtml(normalizeStatusName(latest.status) || "Not set")
                            }
                          </td>
                          <td>
                            ${
                              canEditWorkspace()
                                ? `<input class="reconciliation-amount-input" type="text" value="${escapeHtml(
                                    latest.owner || ""
                                  )}" data-edit-input="true" data-field="owner" data-id="${escapeHtml(latest.id)}" aria-label="Owner for ${escapeHtml(latest.company)}" />`
                                : escapeHtml(latest.owner || "Not set")
                            }
                          </td>
                          <td>
                            ${
                              canEditWorkspace()
                                ? `<input class="reconciliation-amount-input" type="text" inputmode="decimal" value="${escapeHtml(
                                    normalizeMoneyString(latest.amount || "")
                                  )}" data-edit-input="true" data-money-input="true" data-field="reportedAmount" data-id="${escapeHtml(latest.id)}" aria-label="Reported amount for ${escapeHtml(latest.company)}" />`
                                : escapeHtml(formatMoney(toNumber(latest.amount)))
                            }
                          </td>
                          <td>
                            <span class="dashboard-label">${escapeHtml(formatMoney(includedReportedAmount))}</span>
                            ${
                              includeReportedAmount
                                ? ""
                                : '<p class="update-meta">Excluded from committed total</p>'
                            }
                          </td>
                          <td>
                            ${
                              canEditWorkspace()
                                ? `<input class="reconciliation-amount-input" type="text" inputmode="decimal" value="${escapeHtml(
                                    normalizeMoneyString(performance.investedCapital)
                                  )}" data-edit-input="true" data-money-input="true" data-field="investedCapital" data-id="${escapeHtml(latest.id)}" aria-label="Called capital for ${escapeHtml(latest.company)}" />`
                                : escapeHtml(formatMoney(performance.investedCapital))
                            }
                          </td>
                          <td>
                            ${
                              canEditWorkspace()
                                ? `<input class="reconciliation-amount-input" type="text" inputmode="decimal" value="${escapeHtml(
                                    normalizeMoneyString(performance.officialValue)
                                  )}" data-edit-input="true" data-money-input="true" data-field="officialValue" data-id="${escapeHtml(latest.id)}" aria-label="Official NAV for ${escapeHtml(latest.company)}" />`
                                : escapeHtml(formatMoney(performance.officialValue))
                            }
                          </td>
                          <td>
                            ${
                              canEditWorkspace()
                                ? `<input class="reconciliation-amount-input" type="text" inputmode="decimal" value="${escapeHtml(
                                    normalizeMoneyString(performance.internalValue)
                                  )}" data-edit-input="true" data-money-input="true" data-field="internalValue" data-id="${escapeHtml(latest.id)}" aria-label="Internal NAV for ${escapeHtml(latest.company)}" />`
                                : escapeHtml(formatMoney(performance.internalValue))
                            }
                          </td>
                          ${
                            canEditWorkspace()
                              ? `<td><button class="secondary-button card-action-button" type="button" data-action="save-reconciliation-amount" data-id="${escapeHtml(
                                  latest.id
                                )}">Save row</button></td>`
                              : ""
                          }
                        </tr>
                      `
                    )
                    .join("")
                : `<tr><td colspan="${canEditWorkspace() ? "11" : "10"}" class="update-meta">No investments are assigned to this entity.</td></tr>`}
            </tbody>
            <tfoot>
              <tr>
                <td colspan="5">Subtotal</td>
                <td></td>
                <td>${escapeHtml(formatMoney(entityPerformance.reportedAmount))}</td>
                <td>${escapeHtml(formatMoney(entityPerformance.investedCapital))}</td>
                <td>${escapeHtml(formatMoney(entityPerformance.officialValue))}</td>
                <td>${escapeHtml(formatMoney(entityPerformance.internalValue))}</td>
                ${canEditWorkspace() ? "<td></td>" : ""}
              </tr>
            </tfoot>
          </table>
        </section>
      `;
    })
    .join("");
}

function renderAll() {
  companyPerformanceMap = buildCompanyPerformanceMap(allInvestments);
  entityPerformanceMap = buildEntityPerformanceMap(allInvestments);
  renderRoleState();
  renderCompanySuggestions();
  renderFilterOptions();
  const filteredInvestments = filterInvestments(allInvestments);
  renderDashboard(allInvestments);
  renderDataQuality();
  renderResearchLibrary(allInvestments);
  renderUpdates(filteredInvestments);
  renderTasks();
  renderReconciliation();
  renderCompanyPanel();
  renderEntityDetail();
}

async function loadConfig() {
  const config = await fetchJson("/api/config");
  setSignedInState(config.user || null);

  emailStatus.textContent = config.emailConfigured
    ? "Email sending is configured"
    : "Email sending is not configured yet";

  recipientStatus.textContent = config.defaultRecipients.length
    ? `Default team emails: ${config.defaultRecipients.join(", ")}`
    : "No default team emails set";
  digestStatus = {
    lastDigestSentAt: config.lastDigestSentAt || "",
    nextDigestDueAt: config.nextDigestDueAt || "",
    openReminderCount: Number(config.openReminderCount || 0)
  };
  if (digestStatus.lastDigestSentAt) {
    recipientStatus.textContent += ` • Last digest sent ${formatDisplayDate(digestStatus.lastDigestSentAt)}`;
  }
  configuredEntities = Array.isArray(config.entities) ? config.entities : [];

  loginCopy.textContent =
    config.authMode === "individual"
      ? `Use your email and your personal workspace password to sign in. ${config.teamUserCount} team login${config.teamUserCount === 1 ? "" : "s"} configured. Add :viewer to a TEAM_USERS entry for view-only access.`
      : "Use your email and the shared workspace password to unlock updates.";

  if (!config.aiConfigured) {
    deckMessage.textContent =
      "Add OPENAI_API_KEY in Render to turn on deck summarization.";
  }

  if (!config.familyOfficeWorkbookAvailable) {
    importWorkbookMessage.textContent =
      "The family office workbook template is not on the server yet.";
  }

  if (!config.authConfigured) {
    loginMessage.textContent =
      "The server still needs SESSION_SECRET plus TEAM_PASSWORD or TEAM_USERS in Render.";
  }

  roleNotice.textContent = config.canEdit
    ? "Editors can add investments, tasks, documents, and research."
    : "Your account is view-only. You can review investments, research, and tasks, but editing is disabled.";
}

async function loadTasks() {
  try {
    const data = await fetchJson("/api/tasks");
    allTasks = data.tasks || [];
    renderAll();
  } catch (error) {
    if (error.status === 401) {
      setSignedInState(null);
      tasksList.innerHTML = "";
      return;
    }

    throw error;
  }
}

function renderDigestPreview(digest) {
  if (!digest) {
    digestPreview.classList.add("hidden");
    digestPreview.innerHTML = "";
    return;
  }

  const changedInvestments = Array.isArray(digest.changedInvestments)
    ? digest.changedInvestments
    : [];
  const changedByEntity = Array.isArray(digest.changedByEntity) ? digest.changedByEntity : [];
  const overdueTasks = Array.isArray(digest.overdueTasks) ? digest.overdueTasks : [];
  const upcomingTasks = Array.isArray(digest.upcomingTasks) ? digest.upcomingTasks : [];
  const dataAlerts = buildDataQualityAlerts(allInvestments);

  digestPreview.classList.remove("hidden");
  digestPreview.innerHTML = `
    <div class="digest-preview-grid">
      <div>
        <p class="dashboard-label">Window ${escapeHtml(formatDisplayDate(digest.windowStart))} to ${escapeHtml(formatDisplayDate(digest.generatedAt))}</p>
        <p class="highlight-value">${escapeHtml(digest.subject)}</p>
        <p class="update-meta">Biweekly briefing for the family office team, with portfolio changes, reminder pressure, and data hygiene checks in one place.</p>
      </div>

      <div class="digest-preview-summary">
        ${[
          { label: "Changed investments", value: digest.counts.changedInvestments },
          { label: "Open reminders", value: digest.counts.openNextStepTasks },
          { label: "Overdue reminders", value: digest.counts.overdueTasks },
          { label: "Upcoming reminders", value: digest.counts.upcomingTasks || 0 },
          { label: "Data alerts", value: dataAlerts.length }
        ]
          .map(
            (item) => `
              <article class="digest-preview-kpi">
                <p class="dashboard-label">${escapeHtml(item.label)}</p>
                <p class="dashboard-value">${escapeHtml(String(item.value))}</p>
              </article>
            `
          )
          .join("")}
      </div>

      <section class="digest-preview-section">
        <div>
          <p class="dashboard-label">Activity by entity</p>
          <p class="update-meta">A quick scan of where the latest motion happened.</p>
        </div>
        ${
          changedByEntity.length
            ? `<div class="digest-preview-list">${changedByEntity
                .map(
                  (group) => `
                    <article class="digest-preview-item">
                      <p class="highlight-value">${escapeHtml(group.entity)}</p>
                      <p class="update-meta">${escapeHtml(String(group.count))} update${group.count === 1 ? "" : "s"} • ${escapeHtml(
                        group.companies.slice(0, 4).join(", ") || "No companies listed"
                      )}${group.companies.length > 4 ? "..." : ""}</p>
                    </article>
                  `
                )
                .join("")}</div>`
            : '<div class="digest-preview-empty">No entity activity landed in this digest window.</div>'
        }
      </section>

      <div class="digest-preview-columns">
        <section class="digest-preview-section">
          <div>
            <p class="dashboard-label">Portfolio changes</p>
            <p class="update-meta">Recent investment records that were added or materially updated.</p>
          </div>
          ${
            changedInvestments.length
              ? `<div class="digest-preview-list">${changedInvestments
                  .slice(0, 8)
                  .map(
                    (investment) => `
                      <article class="digest-preview-item">
                        <p class="highlight-value">${escapeHtml(investment.company || "Unnamed investment")}</p>
                        <p class="update-meta">${escapeHtml(normalizeEntityName(investment.entity) || "No entity")} • ${escapeHtml(
                          normalizeStatusName(investment.status) || "Not set"
                        )} • ${escapeHtml(formatAmount(investment))}</p>
                        <p class="update-meta">Next step: ${escapeHtml(investment.nextStep || "None recorded")}</p>
                      </article>
                    `
                  )
                  .join("")}</div>`
              : '<div class="digest-preview-empty">No investment updates were entered in this period.</div>'
          }
        </section>

        <section class="digest-preview-section">
          <div>
            <p class="dashboard-label">Reminder pressure</p>
            <p class="update-meta">Items that need attention now, plus what is coming due soon.</p>
          </div>
          ${
            overdueTasks.length
              ? `<div class="digest-preview-list">${overdueTasks
                  .slice(0, 6)
                  .map(
                    (task) => `
                      <article class="digest-preview-item">
                        <p class="highlight-value">${escapeHtml(task.company || "General")}</p>
                        <p class="update-meta">${escapeHtml(task.title)} • overdue since ${escapeHtml(task.dueDate || "no due date")}</p>
                      </article>
                    `
                  )
                  .join("")}</div>`
              : '<div class="digest-preview-empty">No overdue next-step reminders.</div>'
          }
          ${
            upcomingTasks.length
              ? `<div class="digest-preview-list">${upcomingTasks
                  .slice(0, 6)
                  .map(
                    (task) => `
                      <article class="digest-preview-item">
                        <p class="highlight-value">${escapeHtml(task.company || "General")}</p>
                        <p class="update-meta">${escapeHtml(task.title)} • due ${escapeHtml(task.dueDate || "not set")}</p>
                      </article>
                    `
                  )
                  .join("")}</div>`
              : '<div class="digest-preview-empty">No upcoming reminders due in the next two weeks.</div>'
          }
        </section>
      </div>

      <section class="digest-preview-section">
        <div>
          <p class="dashboard-label">Email body preview</p>
          <p class="update-meta">The exact text that will go out if you send the digest now.</p>
        </div>
        <pre class="digest-preview-text">${escapeHtml(digest.text)}</pre>
      </section>
    </div>
  `;
}

function readFileAsBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = String(reader.result || "");
      const [, base64 = ""] = result.split(",");
      resolve(base64);
    };
    reader.onerror = () => reject(new Error("Could not read the PDF file."));
    reader.readAsDataURL(file);
  });
}

async function uploadSupportingDocuments(files) {
  if (!files || !files.length) {
    return;
  }

  documentMessage.textContent = `Uploading ${files.length} document${files.length === 1 ? "" : "s"}...`;
  const uploaded = [];

  for (const file of Array.from(files)) {
    const fileData = await readFileAsBase64(file);
    const result = await fetchJson("/api/upload-document", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        filename: file.name,
        fileData
      })
    });
    uploaded.push(result.document);
  }

  uploadedDocuments = uploadedDocuments.concat(uploaded);
  renderUploadedDocuments();
  documentMessage.textContent = `Uploaded ${uploaded.length} document${uploaded.length === 1 ? "" : "s"}.`;
}

async function uploadCompanyDocuments(files) {
  if (!selectedCompany || !files || !files.length) {
    return;
  }

  const companyRecord = findCompanyRecord(selectedCompany, selectedCompanyEntity);
  const entity = (companyRecord && companyRecord.profile && companyRecord.profile.entity) || "";
  companyDocumentMessage.textContent = `Uploading ${files.length} investment file${files.length === 1 ? "" : "s"}...`;

  for (const file of Array.from(files)) {
    const fileData = await readFileAsBase64(file);
    await fetchJson("/api/company-documents", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        company: selectedCompany,
        entity,
        filename: file.name,
        fileData
      })
    });
  }

  companyDocumentMessage.textContent = `Uploaded ${files.length} investment file${files.length === 1 ? "" : "s"}.`;
  await loadUpdates();
}

async function deleteCompanyDocumentById(documentId) {
  if (!window.confirm("Delete this investment file?")) {
    return;
  }

  companyDocumentMessage.textContent = "Deleting investment file...";

  try {
    await fetchJson(`/api/company-documents/${documentId}`, { method: "DELETE" });
    companyDocumentMessage.textContent = "Investment file deleted.";
    await loadUpdates();
  } catch (error) {
    companyDocumentMessage.textContent = error.message;
  }
}

deckFile.addEventListener("change", () => {
  updateDeckFileLabel(selectedDeckFile());
});

["dragenter", "dragover"].forEach((eventName) => {
  deckDropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    deckDropZone.classList.add("deck-drop-zone-active");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  deckDropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    deckDropZone.classList.remove("deck-drop-zone-active");
  });
});

deckDropZone.addEventListener("drop", (event) => {
  const files = event.dataTransfer && event.dataTransfer.files;
  if (!files || !files.length) {
    return;
  }

  const transferredFile = files[0];
  const transfer = new DataTransfer();
  transfer.items.add(transferredFile);
  deckFile.files = transfer.files;
  updateDeckFileLabel(transferredFile);
});

["dragenter", "dragover"].forEach((eventName) => {
  documentDropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    documentDropZone.classList.add("deck-drop-zone-active");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  documentDropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    documentDropZone.classList.remove("deck-drop-zone-active");
  });
});

documentDropZone.addEventListener("drop", async (event) => {
  const files = event.dataTransfer && event.dataTransfer.files;
  if (!files || !files.length) {
    return;
  }

  try {
    await uploadSupportingDocuments(files);
  } catch (error) {
    documentMessage.textContent = error.message;
  }
});

documentFile.addEventListener("change", async () => {
  const files = documentFile.files;
  if (!files || !files.length) {
    return;
  }

  try {
    await uploadSupportingDocuments(files);
    documentFile.value = "";
  } catch (error) {
    documentMessage.textContent = error.message;
  }
});

["dragenter", "dragover"].forEach((eventName) => {
  companyDocumentDropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    companyDocumentDropZone.classList.add("deck-drop-zone-active");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  companyDocumentDropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    companyDocumentDropZone.classList.remove("deck-drop-zone-active");
  });
});

companyDocumentDropZone.addEventListener("drop", async (event) => {
  const files = event.dataTransfer && event.dataTransfer.files;
  if (!files || !files.length) {
    return;
  }

  try {
    await uploadCompanyDocuments(files);
  } catch (error) {
    companyDocumentMessage.textContent = error.message;
  }
});

companyDocumentFile.addEventListener("change", async () => {
  const files = companyDocumentFile.files;
  if (!files || !files.length) {
    return;
  }

  try {
    await uploadCompanyDocuments(files);
    companyDocumentFile.value = "";
  } catch (error) {
    companyDocumentMessage.textContent = error.message;
  }
});

menuToggleButton.addEventListener("click", () => {
  workspaceMenu.classList.toggle("hidden");
});

workspaceMenu.addEventListener("click", (event) => {
  const target = event.target.closest("[data-view]");
  if (!target) {
    return;
  }

  showWorkspaceView(target.dataset.view);
  window.scrollTo({ top: 0, behavior: "smooth" });
  workspaceMenu.classList.add("hidden");
});

async function loadUpdates() {
  try {
    const data = await fetchJson("/api/investments");
    allInvestments = data.investments;
    allCompanies = data.companies || [];
    renderAll();
    setSignedInState(data.user || currentUser);
  } catch (error) {
    if (error.status === 401) {
      setSignedInState(null);
      updatesList.innerHTML = "";
      return;
    }

    throw error;
  }
}

taskForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  taskMessage.textContent = "Saving task...";
  saveTaskButton.disabled = true;

  const formData = new FormData(taskForm);
  const payload = {
    title: formData.get("title"),
    company: formData.get("company"),
    entity: formData.get("entity"),
    dueDate: formData.get("dueDate"),
    priority: formData.get("priority"),
    category: formData.get("category"),
    assignee: formData.get("assignee"),
    status: formData.get("status"),
    description: formData.get("description")
  };

  try {
    const editingId = editingTaskId.value;
    await fetchJson(editingId ? `/api/tasks/${editingId}` : "/api/tasks", {
      method: editingId ? "PATCH" : "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload)
    });

    taskMessage.textContent = editingId ? "Task updated." : "Task saved.";
    resetTaskForm();
    await loadTasks();
  } catch (error) {
    if (error.status === 401) {
      setSignedInState(null);
      taskMessage.textContent = "Your session expired. Please sign in again.";
      return;
    }

    taskMessage.textContent = error.message;
  } finally {
    saveTaskButton.disabled = false;
  }
});

loginForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  loginMessage.textContent = "Signing in...";
  loginButton.disabled = true;

  const formData = new FormData(loginForm);
  const payload = {
    email: formData.get("email"),
    password: formData.get("password")
  };

  try {
    const result = await fetchJson("/api/session", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload)
    });

    setSignedInState(result.user);
    loginForm.reset();
    loginMessage.textContent = "Signed in. Your workspace is ready.";
    await Promise.all([loadConfig(), loadUpdates()]);
  } catch (error) {
    loginMessage.textContent = error.message;
  } finally {
    loginButton.disabled = false;
  }
});

logoutButton.addEventListener("click", async () => {
  logoutButton.disabled = true;

  try {
    await fetchJson("/api/session", { method: "DELETE" });
    setSignedInState(null);
    updatesList.innerHTML = "";
    formMessage.textContent = "";
    loginMessage.textContent = "Signed out.";
  } catch (error) {
    loginMessage.textContent = error.message;
  } finally {
    logoutButton.disabled = false;
  }
});

summarizeDeckButton.addEventListener("click", async () => {
  const selectedFile = selectedDeckFile();
  if (!selectedFile) {
    deckMessage.textContent = "Choose a PDF deck first.";
    return;
  }

  if (!selectedFile.name.toLowerCase().endsWith(".pdf")) {
    deckMessage.textContent = "Please choose a PDF file.";
    return;
  }

  summarizeDeckButton.disabled = true;
  deckMessage.textContent = "Reading deck and generating summary...";

  try {
    const fileData = await readFileAsBase64(selectedFile);
    const companyValue = form.elements.company.value;
    const stageValue = form.elements.stage.value;
    const result = await fetchJson("/api/summarize-deck", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        filename: selectedFile.name,
        fileData,
        company: companyValue,
        stage: stageValue
      })
    });

    deckSummaryField.value = result.summary;
    deckMessage.textContent = "Deck summary added to the deck summary section.";
  } catch (error) {
    if (error.status === 401) {
      setSignedInState(null);
      deckMessage.textContent = "Your session expired. Please sign in again.";
      return;
    }

    deckMessage.textContent = error.message;
  } finally {
    summarizeDeckButton.disabled = false;
  }
});

summarizeEmailButton.addEventListener("click", async () => {
  const emailText = emailSummaryInput.value.trim();
  if (!emailText) {
    emailMessage.textContent = "Paste an email or thread first.";
    return;
  }

  summarizeEmailButton.disabled = true;
  emailMessage.textContent = "Summarizing email into notes...";

  try {
    const companyValue = form.elements.company.value;
    const stageValue = form.elements.stage.value;
    const result = await fetchJson("/api/summarize-email", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        emailText,
        company: companyValue,
        stage: stageValue
      })
    });

    notesField.value = result.summary;
    emailMessage.textContent = "Email summary added to notes.";
  } catch (error) {
    if (error.status === 401) {
      setSignedInState(null);
      emailMessage.textContent = "Your session expired. Please sign in again.";
      return;
    }

    emailMessage.textContent = error.message;
  } finally {
    summarizeEmailButton.disabled = false;
  }
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  formMessage.textContent = "Saving your update...";
  submitButton.disabled = true;

  const formData = new FormData(form);
  const recipients = String(formData.get("recipients") || "")
    .split(",")
    .map((email) => email.trim())
    .filter(Boolean);

  const payload = {
    company: formData.get("company"),
    entity: formData.get("entity"),
    amount: formData.get("amount"),
    currency: formData.get("currency"),
    stage: formData.get("stage"),
    status: formData.get("status"),
    owner: formData.get("owner"),
    nextStep: formData.get("nextStep"),
    nextStepDueDate: formData.get("nextStepDueDate"),
    contactName: formData.get("contactName"),
    contactPosition: formData.get("contactPosition"),
    contactEmail: formData.get("contactEmail"),
    contactPhone: formData.get("contactPhone"),
    recipients,
    notes: formData.get("notes"),
    deckSummary: formData.get("deckSummary"),
    capitalActivity: collectCapitalActivityRows(),
    valuationDate: formData.get("valuationDate"),
    officialValue: formData.get("officialValue"),
    internalValue: formData.get("internalValue"),
    exitValue: formData.get("exitValue"),
    ownershipPercent: formData.get("ownershipPercent"),
    entityOwnershipPercent: formData.get("entityOwnershipPercent"),
    ownershipNotes: formData.get("ownershipNotes"),
    followOnCapitalAmount: formData.get("followOnCapitalAmount"),
    followOnCapitalStatus: formData.get("followOnCapitalStatus"),
    followOnCapitalNotes: formData.get("followOnCapitalNotes"),
    documentLinks: formData.get("documentLinks"),
    documents: uploadedDocuments,
    decisionDate: formData.get("decisionDate"),
    decisionType: formData.get("decisionType"),
    decisionSummary: formData.get("decisionSummary")
  };

  try {
    const editingId = editingInvestmentId.value;
    const result = await fetchJson(editingId ? `/api/investments/${editingId}` : "/api/investments", {
      method: editingId ? "PATCH" : "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload)
    });

    formMessage.textContent = editingId
      ? "Investment update saved."
      : result.email.sent
        ? "Update saved and email sent to your team."
        : "Update saved. Email is not configured yet, so the team email preview is available on the server.";

    resetFormToCreateMode();
    await loadUpdates();
    await loadConfig();
  } catch (error) {
    if (error.status === 401) {
      setSignedInState(null);
      formMessage.textContent = "Your session expired. Please sign in again.";
      return;
    }

    formMessage.textContent = error.message;
  } finally {
    submitButton.disabled = false;
  }
});

refreshButton.addEventListener("click", () => {
  Promise.all([loadUpdates(), loadTasks()]).catch((error) => {
    formMessage.textContent = error.message;
  });
});

addCapitalActivityButton.addEventListener("click", () => {
  const rows = collectCapitalActivityRows();
  rows.push({ date: "", type: "Investment Amount", amount: "", notes: "" });
  renderCapitalActivityRows(rows);
});

capitalActivityList.addEventListener("click", (event) => {
  const target = event.target.closest("[data-action='remove-capital-activity']");
  if (!target) {
    return;
  }

  const index = Number(target.dataset.index);
  const rows = collectCapitalActivityRows().filter((_, rowIndex) => rowIndex !== index);
  renderCapitalActivityRows(rows);
});

capitalActivityList.addEventListener("input", (event) => {
  const amountField = event.target.closest('[data-capital-field="amount"]');
  if (!amountField) {
    return;
  }

  amountField.value = normalizeMoneyString(amountField.value);
});

attachFormattedInputHandlers();
applyFormInputFormatting();

loadCompanyDetailsButton.addEventListener("click", () => {
  const company = String(form.elements.company.value || "").trim();
  if (!company) {
    formMessage.textContent = "Enter a company name first, then load the latest saved details.";
    return;
  }

  const hydrated = hydrateFormFromCompanyRecord(company);
  formMessage.textContent = hydrated
    ? `Loaded the latest saved details for ${company}.`
    : "No saved company record was found to load.";
});

form.elements.company.addEventListener("change", () => {
  if (editingInvestmentId.value) {
    return;
  }

  const company = String(form.elements.company.value || "").trim();
  if (!company) {
    return;
  }

  hydrateFormFromCompanyRecord(company);
});

downloadCsvButton.addEventListener("click", () => {
  window.location.href = "/api/investments.csv";
});

downloadExcelButton.addEventListener("click", () => {
  window.location.href = "/api/investments.xlsx";
});

downloadFamilyOfficeWorkbookButton.addEventListener("click", () => {
  window.location.href = "/api/family-office-workbook.xlsx";
});

downloadReconciliationButton.addEventListener("click", () => {
  downloadTextFile("entity-reconciliation.csv", buildReconciliationCsv());
});

saveAllReconciliationButton.addEventListener("click", async () => {
  try {
    await saveReconciliationRows(Array.from(dirtyReconciliationRows));
  } catch (error) {
    if (reconciliationMessage) {
      reconciliationMessage.textContent = error.message;
    }
  }
});

downloadBackupButton.addEventListener("click", () => {
  window.location.href = "/api/backup-export";
});

previewDigestButton.addEventListener("click", async () => {
  digestMessage.textContent = "Building biweekly digest preview...";
  previewDigestButton.disabled = true;

  try {
    const result = await fetchJson("/api/biweekly-digest");
    renderDigestPreview(result.digest);
    digestMessage.textContent = digestStatus.lastDigestSentAt
      ? `Preview ready. Last digest sent ${formatDisplayDate(digestStatus.lastDigestSentAt)}.`
      : "Preview ready.";
  } catch (error) {
    if (error.status === 401) {
      setSignedInState(null);
      digestMessage.textContent = "Your session expired. Please sign in again.";
      return;
    }

    digestMessage.textContent = error.message;
  } finally {
    previewDigestButton.disabled = false;
  }
});

sendDigestButton.addEventListener("click", async () => {
  digestMessage.textContent = "Sending biweekly digest...";
  sendDigestButton.disabled = true;

  try {
    const result = await fetchJson("/api/biweekly-digest/send", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({})
    });
    digestMessage.textContent = result.message;
    await loadConfig();
    const preview = await fetchJson("/api/biweekly-digest");
    renderDigestPreview(preview.digest);
  } catch (error) {
    if (error.status === 401) {
      setSignedInState(null);
      digestMessage.textContent = "Your session expired. Please sign in again.";
      return;
    }

    digestMessage.textContent = error.message;
  } finally {
    sendDigestButton.disabled = false;
  }
});

importWorkbookFile.addEventListener("change", async () => {
  const selectedFile = importWorkbookFile.files && importWorkbookFile.files[0];
  if (!selectedFile) {
    return;
  }

  importWorkbookMessage.textContent = "Reading workbook and importing updates...";

  try {
    const fileData = await readFileAsBase64(selectedFile);
    const result = await fetchJson("/api/import-workbook", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        filename: selectedFile.name,
        fileData
      })
    });

    importWorkbookMessage.textContent = `${result.message} Sheets used: ${result.sheets.join(", ")}.`;
    importWorkbookFile.value = "";
    await loadUpdates();
  } catch (error) {
    if (error.status === 401) {
      setSignedInState(null);
      importWorkbookMessage.textContent = "Your session expired. Please sign in again.";
      return;
    }

    importWorkbookMessage.textContent = error.message;
  }
});

restoreBackupFile.addEventListener("change", async () => {
  const selectedFile = restoreBackupFile.files && restoreBackupFile.files[0];
  if (!selectedFile) {
    return;
  }

  if (!window.confirm("Restore this backup? This will replace the current app data, but a fresh backup will be created first.")) {
    restoreBackupFile.value = "";
    return;
  }

  importWorkbookMessage.textContent = "Reading backup and restoring data...";

  try {
    const rawText = await selectedFile.text();
    const payload = JSON.parse(rawText);
    const result = await fetchJson("/api/restore-backup", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ backup: payload })
    });

    importWorkbookMessage.textContent = result.message;
    restoreBackupFile.value = "";
    await Promise.all([loadUpdates(), loadTasks()]);
  } catch (error) {
    if (error.status === 401) {
      setSignedInState(null);
      importWorkbookMessage.textContent = "Your session expired. Please sign in again.";
      return;
    }

    importWorkbookMessage.textContent = error.message;
  }
});

cancelEditButton.addEventListener("click", () => {
  resetFormToCreateMode();
  formMessage.textContent = "Edit canceled.";
});

cancelTaskEditButton.addEventListener("click", () => {
  resetTaskForm();
  taskMessage.textContent = "Task edit canceled.";
});

[entityFilter, searchFilter, statusFilter, stageFilter, ownerFilter].forEach((element) => {
  element.addEventListener("input", renderAll);
  element.addEventListener("change", renderAll);
});

uploadedDocumentsList.addEventListener("click", (event) => {
  const action = event.target.closest("[data-action='remove-document']");
  if (!action) {
    return;
  }

  const documentId = action.dataset.documentId;
  uploadedDocuments = uploadedDocuments.filter((document) => document.id !== documentId);
  renderUploadedDocuments();
});

closeCompanyPanelButton.addEventListener("click", () => {
  selectedCompany = "";
  selectedCompanyEntity = "";
  renderCompanyPanel();
  showWorkspaceView("portfolio");
});

closeEntityDetailButton.addEventListener("click", () => {
  selectedEntity = "";
  renderEntityDetail();
  showWorkspaceView("home");
});

entityPerformanceCards.addEventListener("click", (event) => {
  const card = event.target.closest("[data-entity]");
  if (!card) {
    return;
  }

  selectedEntity = card.dataset.entity || "";
  renderEntityDetail();
  showWorkspaceView("entity");
  window.scrollTo({ top: 0, behavior: "smooth" });
});

dashboardCards.addEventListener("click", (event) => {
  const card = event.target.closest("[data-dashboard-action], [data-entity]");
  if (!card) {
    return;
  }

  const action = card.dataset.dashboardAction || "";
  const entity = card.dataset.entity || "";
  const status = card.dataset.status || "";

  if (action === "entity" && entity) {
    selectedEntity = entity;
    renderEntityDetail();
    showWorkspaceView("entity");
    window.scrollTo({ top: 0, behavior: "smooth" });
    return;
  }

  if (action === "portfolio") {
    if (statusFilter) {
      statusFilter.value = status || "";
    }
    renderDashboard(allInvestments);
    renderEntityDetail();
    renderCompanyPanel();
    showWorkspaceView("portfolio");
    window.scrollTo({ top: 0, behavior: "smooth" });
    return;
  }

  if (action === "tasks") {
    showWorkspaceView("tasks");
    window.scrollTo({ top: 0, behavior: "smooth" });
    return;
  }

  if (action === "quality") {
    showWorkspaceView("quality");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }
});

companyDecisionLog.addEventListener("click", (event) => {
  const target = event.target.closest("[data-action='delete-company-document']");
  if (!target) {
    return;
  }

  deleteCompanyDocumentById(target.dataset.documentId);
});

updatesList.addEventListener("click", (event) => {
  const target = event.target.closest("[data-action], [data-company]");
  if (!target) {
    return;
  }

  const company = target.dataset.company || "";
  const entity = target.dataset.entity || "";
  const action = target.dataset.action || "";
  const investmentId = target.dataset.id || "";

  if (company && (!action || action === "view-company")) {
    selectedCompany = company;
    selectedCompanyEntity = entity;
    renderCompanyPanel();
    showWorkspaceView("portfolio");
    window.scrollTo({ top: 0, behavior: "smooth" });
    return;
  }

  if (action === "edit" && investmentId) {
    beginEditInvestment(investmentId);
    return;
  }

  if (action === "delete" && investmentId) {
    deleteInvestmentById(investmentId);
  }
});

entityDetailInvestments.addEventListener("click", (event) => {
  const target = event.target.closest("[data-action], [data-company]");
  if (!target) {
    return;
  }

  const company = target.dataset.company || "";
  const entity = target.dataset.entity || "";
  const action = target.dataset.action || "";
  const investmentId = target.dataset.id || "";

  if (company && (!action || action === "view-company")) {
    selectedCompany = company;
    selectedCompanyEntity = entity;
    renderCompanyPanel();
    showWorkspaceView("portfolio");
    window.scrollTo({ top: 0, behavior: "smooth" });
    return;
  }

  if (action === "edit" && investmentId) {
    beginEditInvestment(investmentId);
    return;
  }

  if (action === "delete" && investmentId) {
    deleteInvestmentById(investmentId);
  }
});

reconciliationList.addEventListener("click", (event) => {
  const saveButton = event.target.closest("[data-action='save-reconciliation-amount']");
  if (saveButton) {
    const investmentId = saveButton.dataset.id || "";
    if (!investmentId) {
      return;
    }

    const inputs = getReconciliationInputs(investmentId);
    if (!inputs.length) {
      return;
    }

    const values = getReconciliationValues(investmentId);

    saveButton.disabled = true;
    saveButton.textContent = "Saving...";
    saveReconciliationRow(investmentId, values)
      .then(() => {
        setReconciliationRowDirty(investmentId, false);
        if (reconciliationMessage) {
          reconciliationMessage.textContent = "Reconciliation row saved.";
        }
      })
      .catch((error) => {
        if (reconciliationMessage) {
          reconciliationMessage.textContent = error.message;
        }
      })
      .finally(() => {
        saveButton.disabled = false;
        saveButton.textContent = "Save row";
      });
    return;
  }

  const target = event.target.closest("[data-company]");
  if (!target) {
    return;
  }

  const company = target.dataset.company || "";
  const entity = target.dataset.entity || "";
  if (!company) {
    return;
  }

  selectedCompany = company;
  selectedCompanyEntity = entity;
  renderCompanyPanel();
  showWorkspaceView("portfolio");
  window.scrollTo({ top: 0, behavior: "smooth" });
});

function handleReconciliationFieldEdit(event) {
  const editableTarget = event.target.closest("[data-edit-input='true']");
  if (!editableTarget) {
    return;
  }

  if (editableTarget.dataset.moneyInput === "true") {
    editableTarget.value = normalizeMoneyString(editableTarget.value);
  }

  const investmentId = editableTarget.dataset.id || "";
  if (investmentId) {
    setReconciliationRowDirty(investmentId, true);
  }
}

reconciliationList.addEventListener("input", handleReconciliationFieldEdit);
reconciliationList.addEventListener("change", handleReconciliationFieldEdit);

dataQualityList.addEventListener("click", (event) => {
  const target = event.target.closest("[data-action='edit-quality-investment']");
  if (!target) {
    return;
  }

  beginEditInvestment(target.dataset.id || "");
});

tasksList.addEventListener("click", (event) => {
  const target = event.target.closest("[data-action]");
  if (!target) {
    return;
  }

  const action = target.dataset.action;
  const taskId = target.dataset.taskId || "";

  if (action === "edit-task" && taskId) {
    beginEditTask(taskId);
    return;
  }

  if (action === "delete-task" && taskId) {
    deleteTaskById(taskId);
  }
});

Promise.all([loadConfig(), loadUpdates(), loadTasks()]).catch((error) => {
  if (error.status === 401) {
    setSignedInState(null);
    return;
  }

  loginMessage.textContent = error.message;
});

renderUploadedDocuments();
renderCapitalActivityRows([]);
