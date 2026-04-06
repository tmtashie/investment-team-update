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
const workspaceViews = Array.from(document.querySelectorAll(".workspace-view"));
const workspaceMenuLinks = Array.from(document.querySelectorAll(".menu-link[data-view]"));

let currentUser = null;
let allInvestments = [];
let allCompanies = [];
let selectedCompany = "";
let configuredEntities = [];
let companyPerformanceMap = new Map();
let entityPerformanceMap = new Map();
let uploadedDocuments = [];
let allTasks = [];
let activeWorkspaceView = "home";
let selectedEntity = "";

function companyKey(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
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

  if (investment.capitalCallDate || investment.capitalCallAmount) {
    rows.push({
      date: investment.capitalCallDate || "",
      type: "Capital Call",
      amount: investment.capitalCallAmount || "",
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

function renderCapitalActivityRows(rows = []) {
  const normalizedRows = normalizeCapitalActivityRows(rows);
  const rowsToRender = normalizedRows.length ? normalizedRows : [{ date: "", type: "Capital Call", amount: "", notes: "" }];

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
              <input type="number" min="0" step="0.01" data-capital-field="amount" value="${escapeHtml(row.amount)}" placeholder="250000" />
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

function findCompanyRecord(company) {
  const key = companyKey(company);
  return allCompanies.find((record) => record.companyKey === key) || null;
}

function hydrateFormFromCompanyRecord(company) {
  const companyRecord = findCompanyRecord(company);
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
  assignIfBlank("status", latest.status || "");
  assignIfBlank("owner", latest.owner || "");
  assignIfBlank("nextStep", latest.nextStep || "");
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
    const matchesEntity = !filters.entity || investment.entity === filters.entity;
    const matchesStatus = !filters.status || investment.status === filters.status;
    const matchesStage = !filters.stage || investment.stage === filters.stage;
    const matchesOwner = !filters.owner || investment.owner === filters.owner;

    return matchesEntity && matchesSearch && matchesStatus && matchesStage && matchesOwner;
  });
}

function toNumber(value) {
  const amount = Number(value);
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

function buildCompanyPerformance(updates) {
  const normalizedActivities = updates.flatMap((update) => {
    const activityRows = normalizeCapitalActivityRows(
      update.capitalActivity && update.capitalActivity.length
        ? update.capitalActivity
        : buildLegacyCapitalActivityRows(update)
    );

    return activityRows.map((activity) => ({
      ...activity,
      fallbackDate: parseDateValue(update.createdAt, null)
    }));
  });

  const investedCapital = normalizedActivities.reduce((sum, activity) => {
    const type = String(activity.type || "").toLowerCase();
    const amount = toNumber(activity.amount);
    return type.includes("capital call") || type.includes("fee") ? sum + amount : sum;
  }, 0);
  const distributions = normalizedActivities.reduce((sum, activity) => {
    const type = String(activity.type || "").toLowerCase();
    const amount = toNumber(activity.amount);
    return !type.includes("capital call") && !type.includes("fee") ? sum + amount : sum;
  }, 0);
  const officialMark = pickLatestNumericValue(updates, "officialValue");
  const internalMark = pickLatestNumericValue(updates, "internalValue");
  const exitMark = pickLatestNumericValue(updates, "exitValue");
  const valuationDate =
    officialMark.date || internalMark.date || exitMark.date || new Date();

  const baseCashFlows = [];
  normalizedActivities.forEach((activity) => {
    const amount = toNumber(activity.amount);
    const date = parseDateValue(activity.date, activity.fallbackDate);
    const type = String(activity.type || "").toLowerCase();

    if (!amount || !date) {
      return;
    }

    if (type.includes("capital call") || type.includes("fee")) {
      baseCashFlows.push({ date, amount: -amount });
      return;
    }

    baseCashFlows.push({ date, amount });
  });

  const buildView = (terminalValue) => {
    const cashFlows = [...baseCashFlows];
    if (terminalValue > 0 && valuationDate) {
      cashFlows.push({ date: valuationDate, amount: terminalValue });
    }

    return {
      xirr: calculateXirr(cashFlows),
      moic:
        investedCapital > 0 ? (distributions + terminalValue) / investedCapital : null
    };
  };

  return {
    investedCapital,
    distributions,
    officialValue: officialMark.value,
    internalValue: internalMark.value,
    exitValue: exitMark.value,
    official: buildView(officialMark.value),
    internal: buildView(internalMark.value),
    exit: buildView(exitMark.value)
  };
}

function buildCompanyPerformanceMap(investments) {
  const grouped = new Map();

  investments.forEach((investment) => {
    const key = companyKey(investment.company);
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
  const entityList = Array.from(
    new Set(configuredEntities.concat(investments.map((investment) => investment.entity).filter(Boolean)))
  ).filter(Boolean);

  entityList.forEach((entity) => {
    const entityUpdates = investments.filter((investment) => investment.entity === entity);
    performanceMap.set(entity, buildCompanyPerformance(entityUpdates));
  });

  return performanceMap;
}

function buildDashboardCards(investments) {
  const companySummaries = (allCompanies.length
    ? allCompanies.map((company) => ({
        key: company.companyKey,
        latest: company.updates[0],
        performance: buildCompanyPerformance(company.updates)
      }))
    : Array.from(new Set(investments.map((investment) => companyKey(investment.company)).filter(Boolean)))
        .map((key) => {
          const updates = investments.filter((investment) => companyKey(investment.company) === key);
          return {
            key,
            latest: updates.sort((left, right) => new Date(right.createdAt) - new Date(left.createdAt))[0],
            performance: buildCompanyPerformance(updates)
          };
        }))
    .filter(Boolean);
  const openCount = companySummaries.filter(
    (summary) => !["Passed", "Closed"].includes(summary.latest && summary.latest.status)
  ).length;
  const approvedCount = investments.filter((investment) => investment.status === "Approved").length;
  const totalInvestedCapital = companySummaries.reduce(
    (sum, summary) => sum + summary.performance.investedCapital,
    0
  );
  const officialNav = companySummaries.reduce(
    (sum, summary) => sum + summary.performance.officialValue,
    0
  );
  const internalNav = companySummaries.reduce(
    (sum, summary) => sum + summary.performance.internalValue,
    0
  );

  const cards = [
    { label: "Updates", value: String(investments.length), action: "portfolio" },
    { label: "Companies", value: String(companySummaries.length), action: "portfolio" },
    { label: "Active pipeline", value: String(openCount), action: "portfolio" },
    {
      label: "Approved",
      value: String(approvedCount),
      action: "portfolio",
      status: "Approved"
    },
    { label: "Invested capital", value: formatMoney(totalInvestedCapital), action: "portfolio" },
    { label: "Official NAV", value: formatMoney(officialNav), action: "portfolio" },
    { label: "Internal NAV", value: formatMoney(internalNav), action: "portfolio" }
  ];

  const entityTotals = (configuredEntities.length ? configuredEntities : [])
    .map((entity) => ({
      label: entity,
      value: String(investments.filter((investment) => investment.entity === entity).length),
      action: "entity",
      entity
    }));

  return cards.concat(entityTotals);
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
    .map(([entity, performance]) => ({ entity, performance }))
    .filter(({ entity }) => !currentFilters().entity || currentFilters().entity === entity);

  entityPerformanceCards.innerHTML = entityCards
    .map(
      ({ entity, performance }) => `
        <article class="dashboard-card entity-performance-card" data-entity="${escapeHtml(entity)}">
          <p class="dashboard-label">${escapeHtml(entity)}</p>
          <p class="dashboard-value">${escapeHtml(formatMoney(performance.investedCapital))}</p>
          <p class="update-meta">Official NAV: ${escapeHtml(formatMoney(performance.officialValue))}</p>
          <p class="update-meta">Official XIRR: ${escapeHtml(formatPercent(performance.official.xirr))}</p>
          <p class="update-meta">Official MOIC: ${escapeHtml(formatTurns(performance.official.moic))}</p>
        </article>
      `
    )
    .join("");
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

  const investments = allInvestments.filter((investment) => investment.entity === selectedEntity);
  const performance = entityPerformanceMap.get(selectedEntity) || buildCompanyPerformance(investments);

  entityDetailSection.classList.remove("hidden");
  entityDetailTitle.textContent = selectedEntity;
  entityDetailCopy.textContent = `${investments.length} investment update${investments.length === 1 ? "" : "s"} tracked under this entity.`;
  entityDetailSummary.innerHTML = [
    { label: "Invested capital", value: formatMoney(performance.investedCapital) },
    { label: "Distributions", value: formatMoney(performance.distributions) },
    { label: "Official NAV", value: formatMoney(performance.officialValue) },
    { label: "Official XIRR", value: formatPercent(performance.official.xirr) },
    { label: "Official MOIC", value: formatTurns(performance.official.moic) },
    { label: "Current investments", value: String(investments.length) }
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
            companyPerformanceMap.get(companyKey(investment.company)) || buildCompanyPerformance([investment]);
          return `
            <article class="update-card">
              <div class="update-head">
                <button class="link-button company-link" type="button" data-company="${escapeHtml(investment.company)}">
                  ${escapeHtml(investment.company)}
                </button>
                <span class="status-chip">${escapeHtml(investment.status || "Update")}</span>
              </div>
              <p class="update-meta">${escapeHtml(investment.stage || "Stage not set")} • Owner: ${escapeHtml(investment.owner || "Not set")}</p>
              <p class="update-meta">Official NAV ${escapeHtml(formatMoney(companyPerformance.officialValue))} • XIRR ${escapeHtml(formatPercent(companyPerformance.official.xirr))}</p>
              <p class="update-notes">${escapeHtml(summarizeText(investment.notes, "No notes provided."))}</p>
              <div class="card-actions">
                <button class="secondary-button card-action-button" type="button" data-action="view-company" data-company="${escapeHtml(investment.company)}">View company</button>
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
    new Set(configuredEntities.concat(allInvestments.map((item) => item.entity).filter(Boolean)))
  ).sort();
  const statuses = Array.from(new Set(allInvestments.map((item) => item.status).filter(Boolean))).sort();
  const stages = Array.from(new Set(allInvestments.map((item) => item.stage).filter(Boolean))).sort();
  const owners = Array.from(new Set(allInvestments.map((item) => item.owner).filter(Boolean))).sort();

  const assignOptions = (element, placeholder, values) => {
    const currentValue = element.value;
    element.innerHTML = [`<option value="">${placeholder}</option>`]
      .concat(values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`))
      .join("");
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
  form.elements.status.value = investment.status || "";
  form.elements.owner.value = investment.owner || "";
  form.elements.nextStep.value = investment.nextStep || "";
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

  const companyRecord = findCompanyRecord(selectedCompany);
  const companyUpdates = companyRecord
    ? [...companyRecord.updates]
    : allInvestments
        .filter((investment) => companyKey(investment.company) === companyKey(selectedCompany))
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
    new Set(companyUpdates.map((investment) => investment.status).filter(Boolean))
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
        .filter((task) => companyKey(task.company) === companyKey(selectedCompany))
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
    new Set(companyUpdates.map((investment) => investment.entity).filter(Boolean))
  ).map((entity) => ({
    entity,
    performance: buildCompanyPerformance(
      companyUpdates.filter((investment) => investment.entity === entity)
    )
  }));
  companyPanel.classList.remove("hidden");
  const performance =
    companyPerformanceMap.get(companyKey(selectedCompany)) || buildCompanyPerformance(companyUpdates);
  companyPanelTitle.textContent = latest.company || selectedCompany;
  companyPanelCopy.textContent = companyRecord
    ? `${companyUpdates.length} update${companyUpdates.length === 1 ? "" : "s"} organized into structured research, capital, valuation, decision, and document records.`
    : `${companyUpdates.length} update${companyUpdates.length === 1 ? "" : "s"} saved for this company.`;
  companySummary.innerHTML = [
    { label: "Latest status", value: latest.status || "Not set" },
    { label: "Latest entity", value: latest.entity || "Not set" },
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
    { label: "Entities used", value: Array.from(new Set(companyUpdates.map((investment) => investment.entity).filter(Boolean))).join(", ") || "Not set" },
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
    { label: "Invested capital", value: formatMoney(performance.investedCapital) },
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
            <h3>${escapeHtml(investment.status || "Update")}</h3>
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

  updatesList.innerHTML = investments
    .map((investment) => {
      const performance =
        companyPerformanceMap.get(companyKey(investment.company)) || buildCompanyPerformance([investment]);
      const amount = investment.amount
        ? `${escapeHtml(investment.currency)} ${escapeHtml(investment.amount)}`
        : "Amount not specified";

      return `
        <article class="update-card">
          <div class="update-head">
            <button class="link-button company-link" type="button" data-company="${escapeHtml(investment.company)}">
              ${escapeHtml(investment.company)}
            </button>
            <span class="status-chip">${escapeHtml(investment.status)}</span>
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
            <button class="secondary-button card-action-button" type="button" data-action="view-company" data-company="${escapeHtml(investment.company)}">View company</button>
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
              <p class="dashboard-label">${escapeHtml(investment.company)} • ${escapeHtml(investment.status || "Update")}</p>
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

function renderAll() {
  companyPerformanceMap = buildCompanyPerformanceMap(allInvestments);
  entityPerformanceMap = buildEntityPerformanceMap(allInvestments);
  renderRoleState();
  renderCompanySuggestions();
  renderFilterOptions();
  const filteredInvestments = filterInvestments(allInvestments);
  renderDashboard(filteredInvestments);
  renderResearchLibrary(allInvestments);
  renderUpdates(filteredInvestments);
  renderTasks();
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

  const companyRecord = findCompanyRecord(selectedCompany);
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
  rows.push({ date: "", type: "Capital Call", amount: "", notes: "" });
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

downloadBackupButton.addEventListener("click", () => {
  window.location.href = "/api/backup-export";
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
  const action = target.dataset.action || "";
  const investmentId = target.dataset.id || "";

  if (company && (!action || action === "view-company")) {
    selectedCompany = company;
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
  const action = target.dataset.action || "";
  const investmentId = target.dataset.id || "";

  if (company && (!action || action === "view-company")) {
    selectedCompany = company;
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
