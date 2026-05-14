const loginPanel = document.getElementById("loginPanel");
const appPanel = document.getElementById("appPanel");
const loginForm = document.getElementById("loginForm");
const loginMessage = document.getElementById("loginMessage");
const loginCopy = document.getElementById("loginCopy");
const loginButton = document.getElementById("loginButton");
const logoutButton = document.getElementById("logoutButton");
const menuToggleButton = document.getElementById("menuToggleButton");
const workspaceMenu = document.getElementById("workspaceMenu");
const brandSubtitle = document.querySelector(".brand-subtitle");
const heroTopline = document.querySelector(".hero-topline");
const heroTitle = document.querySelector(".brand-title");
const heroThesis = document.querySelector(".hero-thesis");
const statusRow = document.querySelector(".status-row");
const brandLockupEyebrow = document.querySelector(".brand-lockup .eyebrow");
const dashboardSection = document.getElementById("dashboardSection");
const entityPerformanceSection = document.getElementById("entityPerformanceSection");
const form = document.getElementById("investmentForm");
const loadCompanyDetailsButton = document.getElementById("loadCompanyDetailsButton");
const formMessage = document.getElementById("formMessage");
const updatesList = document.getElementById("updatesList");
const authStatus = document.getElementById("authStatus");
const emailStatus = document.getElementById("emailStatus");
const recipientStatus = document.getElementById("recipientStatus");
const heroCopy = document.querySelector(".hero-copy");
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
const aiAnalystSection = document.getElementById("aiAnalystSection");
const aiAnalystWidgetToggleButton = document.getElementById("aiAnalystWidgetToggleButton");
const aiAnalystWidgetMinimizeButton = document.getElementById("aiAnalystWidgetMinimizeButton");
const aiAnalystWidgetCloseButton = document.getElementById("aiAnalystWidgetCloseButton");
const aiAnalystCompanyField = document.getElementById("aiAnalystCompanyField");
const aiAnalystEntityField = document.getElementById("aiAnalystEntityField");
const aiAnalystPromptField = document.getElementById("aiAnalystPromptField");
const askAiAnalystButton = document.getElementById("askAiAnalystButton");
const clearAiAnalystButton = document.getElementById("clearAiAnalystButton");
const useAiAnalystInSummaryButton = document.getElementById("useAiAnalystInSummaryButton");
const aiAnalystMessage = document.getElementById("aiAnalystMessage");
const aiAnalystResponseCard = document.getElementById("aiAnalystResponseCard");
const aiAnalystContextLabel = document.getElementById("aiAnalystContextLabel");
const aiAnalystResponseBody = document.getElementById("aiAnalystResponseBody");
const companySuggestions = document.getElementById("companySuggestions");
const companyPanel = document.getElementById("companyPanel");
const companyPanelTitle = document.getElementById("companyPanelTitle");
const companyPanelCopy = document.getElementById("companyPanelCopy");
const openAiAnalystButton = document.getElementById("openAiAnalystButton");
const generateInvestmentSummaryButton = document.getElementById("generateInvestmentSummaryButton");
const investmentSummaryPanel = document.getElementById("investmentSummaryPanel");
const investmentSummaryDocument = document.getElementById("investmentSummaryDocument");
const printInvestmentSummaryButton = document.getElementById("printInvestmentSummaryButton");
const closeInvestmentSummaryButton = document.getElementById("closeInvestmentSummaryButton");
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
const reportUpdateTypeFilter = document.getElementById("reportUpdateTypeFilter");
const reportUpdatePeriodFilter = document.getElementById("reportUpdatePeriodFilter");
const addReportUpdateButton = document.getElementById("addReportUpdateButton");
const compareLatestReportUpdatesButton = document.getElementById(
  "compareLatestReportUpdatesButton"
);
const generateAllReportSummaryButton = document.getElementById(
  "generateAllReportSummaryButton"
);
const reportUpdateComposer = document.getElementById("reportUpdateComposer");
const reportUpdateDateField = document.getElementById("reportUpdateDateField");
const reportUpdatePeriodField = document.getElementById("reportUpdatePeriodField");
const reportUpdateTypeField = document.getElementById("reportUpdateTypeField");
const reportUpdateSourceTypeField = document.getElementById("reportUpdateSourceTypeField");
const reportUpdateTitleField = document.getElementById("reportUpdateTitleField");
const reportUpdateOriginalNotesField = document.getElementById("reportUpdateOriginalNotesField");
const reportUpdateAiSummaryField = document.getElementById("reportUpdateAiSummaryField");
const reportUpdateKeyMetricsField = document.getElementById("reportUpdateKeyMetricsField");
const reportUpdateKeyWinsField = document.getElementById("reportUpdateKeyWinsField");
const reportUpdateKeyRisksField = document.getElementById("reportUpdateKeyRisksField");
const reportUpdateActionItemsField = document.getElementById("reportUpdateActionItemsField");
const reportUpdateAttachmentField = document.getElementById("reportUpdateAttachmentField");
const summarizeReportUpdateDraftButton = document.getElementById(
  "summarizeReportUpdateDraftButton"
);
const saveReportUpdateButton = document.getElementById("saveReportUpdateButton");
const cancelReportUpdateButton = document.getElementById("cancelReportUpdateButton");
const reportUpdateMessage = document.getElementById("reportUpdateMessage");
const reportUpdateInsightPanel = document.getElementById("reportUpdateInsightPanel");
const reportUpdateInsightTitle = document.getElementById("reportUpdateInsightTitle");
const reportUpdateInsightBody = document.getElementById("reportUpdateInsightBody");
const useReportInsightInSummaryButton = document.getElementById("useReportInsightInSummaryButton");
const closeReportUpdateInsightButton = document.getElementById("closeReportUpdateInsightButton");
const reportUpdatesList = document.getElementById("reportUpdatesList");
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
let dismissedDataAlerts = {};
let dirtyReconciliationRows = new Set();
let savingAllReconciliation = false;
let latestCompanySummaryContext = null;
let latestAiAnalystResult = null;
let investmentSummaryAiNotes = null;
let latestReportInsight = null;
let aiAnalystWidgetMinimized = false;
let activePortfolioPreset = "";
let reportUpdateFilters = {
  type: "",
  period: ""
};

const REPORT_UPDATE_TYPES = [
  "Monthly",
  "Quarterly",
  "Annual",
  "Capital Call",
  "Legal",
  "Call Notes",
  "Other"
];

const REPORT_SOURCE_TYPES = ["PDF", "Email", "Call", "Manual Note", "Other"];
const DEFAULT_BRAND_SUBTITLE = "Family office investment workspace";
const DEFAULT_BRAND_EYEBROW = "BVB";
const DEFAULT_HERO_TITLE = "BVB";
const DASHBOARD_VIEWER_BRAND_SUBTITLE = "Private family office reporting dashboard";
const DEFAULT_HERO_COPY =
  "A cleaner family office dashboard for tracking investments, entity exposure, research, documents, decisions, and follow-on capital from one disciplined source of truth.";
const DASHBOARD_VIEWER_HERO_COPY =
  "Portfolio dashboard for internal family office review.";
const DEFAULT_HERO_THESIS = "Stewarding assets, growing wealth, and pursuing bold ideas.";
const DEFAULT_DASHBOARD_COPY =
  "Track pipeline size, marks, and company activity from the BVB home dashboard.";
const DASHBOARD_VIEWER_DASHBOARD_COPY =
  "Review top-line portfolio activity, capital exposure, and current marks across the family office.";
const DEFAULT_ENTITY_PERFORMANCE_COPY =
  "See committed capital, called capital, marks, and performance by family office entity.";
const DASHBOARD_VIEWER_ENTITY_PERFORMANCE_COPY =
  "Review committed capital, called capital, current marks, and performance by family office entity.";

function normalizeDismissedAlertMap(value) {
  if (!value || typeof value !== "object") {
    return {};
  }

  return Object.entries(value).reduce((entries, [alertKey, dismissedUntil]) => {
    const key = String(alertKey || "").trim();
    const parsed = parseDateValue(dismissedUntil, null);
    if (!key || !parsed) {
      return entries;
    }

    entries[key] = parsed.toISOString();
    return entries;
  }, {});
}

function isAlertDismissed(alertKey) {
  const key = String(alertKey || "").trim();
  if (!key) {
    return false;
  }

  const dismissedUntil = dismissedDataAlerts[key];
  const parsed = parseDateValue(dismissedUntil, null);
  if (!parsed) {
    delete dismissedDataAlerts[key];
    return false;
  }

  if (parsed.getTime() <= Date.now()) {
    delete dismissedDataAlerts[key];
    return false;
  }

  return true;
}

function getDashboardViewerGreeting(user) {
  const email = String((user && user.email) || "").trim().toLowerCase();
  if (!email) {
    return "Welcome";
  }
  if (email === "lee@leebeaman.com") {
    return "Welcome: Lee Beaman";
  }
  const namePart = email.split("@")[0].replace(/[._-]+/g, " ").trim();
  if (!namePart) {
    return "Welcome";
  }
  const titleCaseName = namePart
    .split(/\s+/)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join(" ");
  return `Welcome: ${titleCaseName}`;
}

function getDashboardViewerBrandSubtitle(user) {
  const email = String((user && user.email) || "").trim().toLowerCase();
  if (email === "lee@leebeaman.com") {
    return "Beaman Family Office";
  }
  return DASHBOARD_VIEWER_BRAND_SUBTITLE;
}

function getDashboardViewerHeroThesis(user) {
  const email = String((user && user.email) || "").trim().toLowerCase();
  if (email === "lee@leebeaman.com") {
    return "Welcome: Lee Beaman";
  }
  return getDashboardViewerGreeting(user);
}

function getDashboardViewerHeroCopy(user) {
  const email = String((user && user.email) || "").trim().toLowerCase();
  if (email === "lee@leebeaman.com") {
    return "Private portfolio summary for internal review.";
  }
  return "Private portfolio summary for internal review.";
}

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

async function fetchJson(url, options = {}) {
  const response = await fetch(url, {
    credentials: "include",
    ...options
  });
  const data = await response.json();

  if (!response.ok) {
    const error = new Error(data.error || "Request failed");
    error.status = response.status;
    throw error;
  }

  return data;
}

function addListener(element, eventName, handler, options) {
  if (!element) {
    return;
  }

  element.addEventListener(eventName, handler, options);
}

function setSignedInState(user) {
  currentUser = user;
  const isSignedIn = Boolean(user);
  const dashboardViewer = Boolean(user && user.role === "dashboard-viewer");

  loginPanel.classList.toggle("hidden", isSignedIn);
  appPanel.classList.toggle("hidden", !isSignedIn);
  logoutButton.classList.toggle("hidden", !isSignedIn);
  authStatus.textContent = isSignedIn
    ? dashboardViewer
      ? getDashboardViewerGreeting(user)
      : `Signed in as ${user.email}${user.role ? ` • ${user.role}` : ""}`
    : "Please sign in to view updates";
  brandSubtitle.textContent = dashboardViewer
    ? getDashboardViewerBrandSubtitle(user)
    : DEFAULT_BRAND_SUBTITLE;
  if (brandLockupEyebrow) {
    brandLockupEyebrow.textContent = dashboardViewer ? "" : DEFAULT_BRAND_EYEBROW;
  }
  const heroToplineEyebrow = heroTopline ? heroTopline.querySelector(".eyebrow") : null;
  if (heroToplineEyebrow) {
    heroToplineEyebrow.textContent = dashboardViewer ? "" : DEFAULT_BRAND_EYEBROW;
  }
  if (heroTitle) {
    heroTitle.textContent = dashboardViewer ? "" : DEFAULT_HERO_TITLE;
  }
  heroThesis.textContent = dashboardViewer
    ? getDashboardViewerHeroThesis(user)
    : DEFAULT_HERO_THESIS;
  heroCopy.textContent = dashboardViewer ? getDashboardViewerHeroCopy(user) : DEFAULT_HERO_COPY;

  workspaceMenuLinks.forEach((link) => {
    const viewName = link.dataset.view || "";
    link.classList.toggle("hidden", isSignedIn && !canAccessWorkspaceView(viewName));
  });

  if (isSignedIn) {
    if (!canAccessWorkspaceView(activeWorkspaceView)) {
      activeWorkspaceView = "home";
    }
    showWorkspaceView(activeWorkspaceView);
    syncAiAnalystWidgetAvailability();
  } else {
    setAiAnalystWidgetOpen(false);
  }
}

function showWorkspaceView(viewName) {
  const requestedView = viewName || "home";
  activeWorkspaceView = canAccessWorkspaceView(requestedView) ? requestedView : "home";

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
  return !currentUser || !["viewer", "dashboard-viewer"].includes(currentUser.role);
}

function isDashboardViewer() {
  return Boolean(currentUser && currentUser.role === "dashboard-viewer");
}

function canAccessWorkspaceView(viewName) {
  if (!isDashboardViewer()) {
    return true;
  }

  return viewName === "home";
}

function canOpenCompanyDetails() {
  return !isDashboardViewer();
}

function canUseAiAnalyst() {
  return Boolean(currentUser) && !isDashboardViewer();
}

function setAiAnalystWidgetMinimized(minimized) {
  aiAnalystWidgetMinimized = Boolean(minimized);
  if (!aiAnalystSection) {
    return;
  }

  aiAnalystSection.classList.toggle("ai-analyst-widget-panel-minimized", aiAnalystWidgetMinimized);
  if (aiAnalystWidgetMinimizeButton) {
    aiAnalystWidgetMinimizeButton.textContent = aiAnalystWidgetMinimized ? "Expand" : "Minimize";
    aiAnalystWidgetMinimizeButton.setAttribute(
      "aria-label",
      aiAnalystWidgetMinimized ? "Expand AI Analyst" : "Minimize AI Analyst"
    );
  }
}

function setAiAnalystWidgetOpen(isOpen, options = {}) {
  const shouldOpen = Boolean(isOpen) && canUseAiAnalyst();

  if (aiAnalystWidgetToggleButton) {
    aiAnalystWidgetToggleButton.classList.toggle("hidden", !canUseAiAnalyst());
    aiAnalystWidgetToggleButton.classList.toggle("is-open", shouldOpen);
    aiAnalystWidgetToggleButton.setAttribute("aria-expanded", shouldOpen ? "true" : "false");
  }

  if (!aiAnalystSection) {
    return;
  }

  aiAnalystSection.classList.toggle("hidden", !shouldOpen);
  aiAnalystSection.setAttribute("aria-hidden", shouldOpen ? "false" : "true");

  if (shouldOpen) {
    setAiAnalystWidgetMinimized(false);
    if (options.focusInput && aiAnalystPromptField) {
      aiAnalystPromptField.focus();
    }
  }
}

function syncAiAnalystWidgetAvailability() {
  if (!canUseAiAnalyst()) {
    setAiAnalystWidgetOpen(false);
    return;
  }

  if (aiAnalystWidgetToggleButton) {
    aiAnalystWidgetToggleButton.classList.remove("hidden");
  }
}

function clearAiAnalystState(resetFields = false) {
  latestAiAnalystResult = null;
  if (aiAnalystResponseCard) {
    aiAnalystResponseCard.classList.add("hidden");
  }
  if (aiAnalystContextLabel) {
    aiAnalystContextLabel.textContent = "Portfolio-wide analysis";
  }
  if (aiAnalystResponseBody) {
    aiAnalystResponseBody.textContent = "";
  }
  if (aiAnalystMessage) {
    aiAnalystMessage.textContent = "";
  }
  if (useAiAnalystInSummaryButton) {
    useAiAnalystInSummaryButton.classList.add("hidden");
  }
  if (resetFields) {
    if (aiAnalystCompanyField) {
      aiAnalystCompanyField.value = "";
    }
    if (aiAnalystEntityField) {
      aiAnalystEntityField.value = "";
    }
    if (aiAnalystPromptField) {
      aiAnalystPromptField.value = "";
    }
  }
}

function renderAiAnalystResponse(result) {
  latestAiAnalystResult = result || null;
  if (!result) {
    clearAiAnalystState();
    return;
  }

  if (aiAnalystContextLabel) {
    aiAnalystContextLabel.textContent = result.contextLabel || "Portfolio-wide analysis";
  }
  if (aiAnalystResponseBody) {
    aiAnalystResponseBody.textContent = result.answer || "";
  }
  if (aiAnalystResponseCard) {
    aiAnalystResponseCard.classList.remove("hidden");
  }
  if (useAiAnalystInSummaryButton) {
    useAiAnalystInSummaryButton.classList.toggle("hidden", !(result.company || selectedCompany));
  }
}

function applyAiAnalystTemplate(template) {
  if (!aiAnalystPromptField) {
    return;
  }

  const company = String(
    (aiAnalystCompanyField && aiAnalystCompanyField.value) || selectedCompany || "this investment"
  ).trim();
  const prompt = String(template || "")
    .replace(/\{\{company\}\}/g, company || "this investment")
    .trim();
  aiAnalystPromptField.value = prompt;
  aiAnalystPromptField.focus();
}

function openAiAnalystForSelectedCompany() {
  if (!aiAnalystSection || !canUseAiAnalyst()) {
    return;
  }

  if (selectedCompany && aiAnalystCompanyField) {
    aiAnalystCompanyField.value = selectedCompany;
  }
  if (selectedCompanyEntity && aiAnalystEntityField) {
    aiAnalystEntityField.value = normalizeEntityName(selectedCompanyEntity);
  }
  if (selectedCompany && aiAnalystPromptField && !aiAnalystPromptField.value.trim()) {
    aiAnalystPromptField.value = `Summarize ${selectedCompany} for Lee. Include the main risks, what is missing, and recommended next steps.`;
  }

  setAiAnalystWidgetOpen(true, { focusInput: !selectedCompany });
}

function closeReportInsightPanel() {
  latestReportInsight = null;
  if (reportUpdateInsightPanel) {
    reportUpdateInsightPanel.classList.add("hidden");
  }
  if (reportUpdateInsightBody) {
    reportUpdateInsightBody.innerHTML = "";
  }
}

function renderReportInsight(title, body, company = selectedCompany, entity = selectedCompanyEntity) {
  latestReportInsight = {
    title: String(title || "").trim(),
    body: String(body || "").trim(),
    company: String(company || "").trim(),
    entity: String(entity || "").trim()
  };

  if (reportUpdateInsightTitle) {
    reportUpdateInsightTitle.textContent = latestReportInsight.title || "Lee-ready output";
  }
  if (reportUpdateInsightBody) {
    reportUpdateInsightBody.innerHTML = `
      <article class="timeline-card timeline-card-compact">
        <pre class="report-insight-text">${escapeHtml(latestReportInsight.body || "Not Available")}</pre>
      </article>
    `;
  }
  if (reportUpdateInsightPanel) {
    reportUpdateInsightPanel.classList.remove("hidden");
    reportUpdateInsightPanel.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }
}

function buildInvestmentPatchPayload(investment, overrides = {}) {
  if (!investment) {
    return null;
  }

  return {
    company: investment.company,
    entity: investment.entity,
    amount: investment.amount,
    currency: investment.currency,
    stage: investment.stage,
    status: normalizeStatusName(investment.status) || investment.status,
    owner: investment.owner,
    nextStep: investment.nextStep,
    nextStepDueDate: investment.nextStepDueDate,
    notes: investment.notes,
    deckSummary: investment.deckSummary,
    capitalActivity: normalizeCapitalActivityRows(investment.capitalActivity),
    valuationDate: investment.valuationDate,
    officialValue: investment.officialValue,
    internalValue: investment.internalValue,
    exitValue: investment.exitValue,
    ownershipPercent: investment.ownershipPercent,
    entityOwnershipPercent: investment.entityOwnershipPercent,
    ownershipNotes: investment.ownershipNotes,
    followOnCapitalAmount: investment.followOnCapitalAmount,
    followOnCapitalStatus: investment.followOnCapitalStatus,
    followOnCapitalNotes: investment.followOnCapitalNotes,
    contactName: investment.contactName,
    contactPosition: investment.contactPosition,
    contactEmail: investment.contactEmail,
    contactPhone: investment.contactPhone,
    documentLinks: investment.documentLinks,
    documents: Array.isArray(investment.documents) ? investment.documents : [],
    decisionDate: investment.decisionDate,
    decisionType: investment.decisionType,
    decisionSummary: investment.decisionSummary,
    recipients: Array.isArray(investment.recipients) ? investment.recipients : [],
    reportUpdates: normalizeReportUpdateRows(investment.reportUpdates),
    ...overrides
  };
}

async function saveCompanyReportUpdates(reportUpdates, successMessage = "Report update saved.") {
  const companyRecord = findCompanyRecord(selectedCompany, selectedCompanyEntity);
  const latest = companyRecord && companyRecord.latest ? companyRecord.latest : null;
  if (!latest) {
    throw new Error("Open an investment record first.");
  }

  const payload = buildInvestmentPatchPayload(latest, {
    reportUpdates: normalizeReportUpdateRows(reportUpdates)
  });

  await fetchJson(`/api/investments/${latest.id}`, {
    method: "PATCH",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });

  if (reportUpdateMessage) {
    reportUpdateMessage.textContent = successMessage;
  }
  await loadUpdates();
}

function buildReportAnalystPrompt(mode, rows, companyRecord) {
  const companyName = (companyRecord && companyRecord.latest && companyRecord.latest.company) || selectedCompany || "this investment";
  const baseInstruction =
    mode === "compare"
      ? `Compare the two most recent updates for ${companyName}.`
      : mode === "lee-single"
        ? `Create a Lee-ready internal summary for one update on ${companyName}.`
        : mode === "lee-all"
          ? `Create a Lee-ready internal summary using all provided updates for ${companyName}.`
          : `Summarize the provided update for ${companyName}.`;

  const rowsText = normalizeReportUpdateRows(rows)
    .map(
      (row, index) => [
        `Update ${index + 1}:`,
        `Date received: ${row.date || "Not Available"}`,
        `Report period: ${row.reportPeriod || "Not Available"}`,
        `Update type: ${row.type || "Not Available"}`,
        `Source type: ${row.sourceType || "Not Available"}`,
        `Title: ${row.title || "Not Available"}`,
        `Original notes: ${row.originalNotes || "Not Available"}`,
        `AI summary: ${row.aiSummary || "Not Available"}`,
        `Key wins: ${row.keyWins || "Not Available"}`,
        `Key risks: ${row.keyRisks || "Not Available"}`,
        `Key metrics: ${row.keyMetrics || "Not Available"}`,
        `Action items: ${row.actionItems || "Not Available"}`,
        `Attachment: ${row.attachmentLink || "Not Available"}`
      ].join("\n")
    )
    .join("\n\n");

  const endingInstruction =
    mode === "compare"
      ? "Focus on what changed, new risks, what improved, what worsened, and what needs follow-up."
      : mode.startsWith("lee")
        ? "Use short internal-ready sections: Executive Summary, What Changed, Risks, Key Metrics, Next Steps, Verification Note."
        : "Summarize the key wins, risks, metrics, and follow-ups.";

  return [baseInstruction, "", rowsText, "", endingInstruction].join("\n");
}

async function runReportAnalystPrompt(mode, rows, companyRecord, options = {}) {
  const prompt = buildReportAnalystPrompt(mode, rows, companyRecord);
  const company =
    (companyRecord && companyRecord.latest && companyRecord.latest.company) || selectedCompany || "";
  const entity =
    normalizeEntityName(
      (companyRecord && companyRecord.latest && companyRecord.latest.entity) || selectedCompanyEntity || ""
    ) || "";

  if (reportUpdateMessage) {
    reportUpdateMessage.textContent = options.loadingMessage || "AI Analyst is reviewing the update...";
  }

  const result = await fetchJson("/api/ai-agent", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      question: prompt,
      company,
      entity
    })
  });

  return {
    answer: String((result && result.answer) || "").trim(),
    company: company || String((result && result.company) || "").trim(),
    entity: entity || normalizeEntityName(String((result && result.entity) || "").trim())
  };
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

function normalizeReportUpdateRows(rows) {
  return (Array.isArray(rows) ? rows : [])
    .map((row) => ({
      date: String((row && row.date) || "").trim(),
      reportPeriod: String((row && row.reportPeriod) || "").trim(),
      type: String((row && row.type) || "").trim(),
      sourceType: String((row && row.sourceType) || "").trim(),
      title: String((row && row.title) || "").trim(),
      originalNotes: String((row && row.originalNotes) || "").trim(),
      aiSummary: String((row && row.aiSummary) || "").trim(),
      keyWins: String((row && row.keyWins) || "").trim(),
      keyRisks: String((row && row.keyRisks) || "").trim(),
      keyMetrics: String((row && row.keyMetrics) || "").trim(),
      actionItems: String((row && row.actionItems) || "").trim(),
      attachmentLink: String((row && row.attachmentLink) || "").trim(),
      sourceUpdateId: String((row && row.sourceUpdateId) || "").trim()
    }))
    .filter((row) => Object.values(row).some(Boolean));
}

function formatDisplayDateOrText(value) {
  const text = String(value || "").trim();
  if (!text) {
    return "Not Available";
  }

  return formatDisplayDate(text);
}

function getReportUpdateFilters() {
  return {
    type: String((reportUpdateTypeFilter && reportUpdateTypeFilter.value) || "").trim(),
    period: String((reportUpdatePeriodFilter && reportUpdatePeriodFilter.value) || "").trim()
  };
}

function resetReportUpdateComposer() {
  if (!reportUpdateComposer) {
    return;
  }

  reportUpdateDateField.value = "";
  reportUpdatePeriodField.value = "";
  reportUpdateTypeField.value = "Monthly";
  reportUpdateSourceTypeField.value = "PDF";
  reportUpdateTitleField.value = "";
  reportUpdateOriginalNotesField.value = "";
  reportUpdateAiSummaryField.value = "";
  reportUpdateKeyMetricsField.value = "";
  reportUpdateKeyWinsField.value = "";
  reportUpdateKeyRisksField.value = "";
  reportUpdateActionItemsField.value = "";
  reportUpdateAttachmentField.value = "";
  reportUpdateMessage.textContent = "";
  reportUpdateComposer.classList.add("hidden");
}

function openReportUpdateComposer() {
  if (!reportUpdateComposer) {
    return;
  }

  reportUpdateComposer.classList.remove("hidden");
  reportUpdateDateField.value = new Date().toISOString().slice(0, 10);
  reportUpdateTitleField.focus();
}

function collectReportUpdateFormData() {
  return {
    date: String(reportUpdateDateField.value || "").trim(),
    reportPeriod: String(reportUpdatePeriodField.value || "").trim(),
    type: String(reportUpdateTypeField.value || "").trim(),
    sourceType: String(reportUpdateSourceTypeField.value || "").trim(),
    title: String(reportUpdateTitleField.value || "").trim(),
    originalNotes: String(reportUpdateOriginalNotesField.value || "").trim(),
    aiSummary: String(reportUpdateAiSummaryField.value || "").trim(),
    keyMetrics: String(reportUpdateKeyMetricsField.value || "").trim(),
    keyWins: String(reportUpdateKeyWinsField.value || "").trim(),
    keyRisks: String(reportUpdateKeyRisksField.value || "").trim(),
    actionItems: String(reportUpdateActionItemsField.value || "").trim(),
    attachmentLink: String(reportUpdateAttachmentField.value || "").trim()
  };
}

function buildReportUpdateField(label, value) {
  const text = String(value || "").trim();
  if (!text) {
    return "";
  }

  return `
    <div class="report-update-detail">
      <p class="dashboard-label">${escapeHtml(label)}</p>
      <p class="highlight-value">${escapeHtml(text)}</p>
    </div>
  `;
}

function renderReportUpdateFilters(rows) {
  const normalizedRows = normalizeReportUpdateRows(rows);
  const periods = Array.from(
    new Set(normalizedRows.map((row) => row.reportPeriod).filter(Boolean))
  ).sort((left, right) => right.localeCompare(left));
  const types = Array.from(new Set(normalizedRows.map((row) => row.type).filter(Boolean)));

  if (reportUpdateTypeFilter) {
    reportUpdateTypeFilter.innerHTML = [
      '<option value="">All update types</option>',
      ...REPORT_UPDATE_TYPES.filter((type) => types.includes(type)).map(
        (type) =>
          `<option value="${escapeHtml(type)}"${reportUpdateFilters.type === type ? " selected" : ""}>${escapeHtml(type)}</option>`
      ),
      ...types
        .filter((type) => !REPORT_UPDATE_TYPES.includes(type))
        .map(
          (type) =>
            `<option value="${escapeHtml(type)}"${reportUpdateFilters.type === type ? " selected" : ""}>${escapeHtml(type)}</option>`
        )
    ].join("");
  }

  if (reportUpdatePeriodFilter) {
    reportUpdatePeriodFilter.innerHTML = [
      '<option value="">All periods</option>',
      ...periods.map(
        (period) =>
          `<option value="${escapeHtml(period)}"${reportUpdateFilters.period === period ? " selected" : ""}>${escapeHtml(period)}</option>`
      )
    ].join("");
  }
}

function renderReportUpdatesSection(companyRecord) {
  const reportRows = companyRecord
    ? normalizeReportUpdateRows(companyRecord.reportUpdates).sort(
        (left, right) => new Date(right.date || 0).getTime() - new Date(left.date || 0).getTime()
      )
    : [];
  renderReportUpdateFilters(reportRows);

  const filteredRows = reportRows
    .map((row, index) => ({ row, index }))
    .filter(({ row }) => {
    if (reportUpdateFilters.type && row.type !== reportUpdateFilters.type) {
      return false;
    }
    if (reportUpdateFilters.period && row.reportPeriod !== reportUpdateFilters.period) {
      return false;
    }
    return true;
  });

  if (reportUpdatesList) {
    reportUpdatesList.innerHTML = filteredRows.length
      ? filteredRows
          .map(
            ({ row, index }) => `
              <article class="timeline-card report-update-card">
                <div class="update-head">
                  <div>
                    <h3>${escapeHtml(row.title || row.type || "Update")}</h3>
                    <p class="update-meta">
                      ${escapeHtml(formatDisplayDate(row.date || "") || row.date || "Not Available")} • ${escapeHtml(row.reportPeriod || "Period not set")} • ${escapeHtml(row.type || "Type not set")} • ${escapeHtml(row.sourceType || "Source not set")}
                    </p>
                  </div>
                  <span class="status-chip">${escapeHtml(row.type || "Update")}</span>
                </div>
                <div class="report-update-grid">
                  ${buildReportUpdateField("AI summary", row.aiSummary || "Not Available")}
                  ${buildReportUpdateField("Key wins", row.keyWins || "Not Available")}
                  ${buildReportUpdateField("Key risks", row.keyRisks || "Not Available")}
                  ${buildReportUpdateField("Key metrics", row.keyMetrics || "Not Available")}
                  ${buildReportUpdateField("Action items / follow-ups", row.actionItems || "Not Available")}
                  ${buildReportUpdateField("Attachment link or file reference", row.attachmentLink || "Not Available")}
                </div>
                ${
                  row.originalNotes
                    ? `<div class="update-subsection"><p class="dashboard-label">Original notes or pasted text</p><p class="update-notes">${escapeHtml(row.originalNotes)}</p></div>`
                    : ""
                }
                <div class="card-actions">
                  <button type="button" class="secondary-button card-action-button" data-action="summarize-report-update" data-index="${index}">
                    Summarize with AI
                  </button>
                  <button type="button" class="secondary-button card-action-button" data-action="generate-report-update-summary" data-index="${index}">
                    Generate Lee-ready summary
                  </button>
                </div>
              </article>
            `
          )
          .join("")
      : '<p class="update-meta">No saved updates or reports yet. Add your first monthly report, quarterly letter, capital call, or call note above.</p>';
  }

  if (compareLatestReportUpdatesButton) {
    compareLatestReportUpdatesButton.disabled = reportRows.length < 2;
  }
  if (generateAllReportSummaryButton) {
    generateAllReportSummaryButton.disabled = reportRows.length === 0;
  }
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
  const dashboardViewer = isDashboardViewer();
  const dashboardCopy = dashboardSection.querySelector(".section-copy");
  const entityPerformanceCopy = entityPerformanceSection.querySelector(".section-copy");
  form.classList.toggle("hidden", !editable);
  taskForm.classList.toggle("hidden", !editable);
  sendDigestButton.classList.toggle("hidden", !editable);
  menuToggleButton.classList.toggle("hidden", dashboardViewer);
  if (dashboardViewer) {
    workspaceMenu.classList.add("hidden");
  }
  heroTopline.classList.toggle("hidden", dashboardViewer);
  heroTitle.classList.toggle("hidden", dashboardViewer);
  brandLockupEyebrow.classList.toggle("hidden", dashboardViewer);
  authStatus.classList.toggle("hidden", dashboardViewer);
  emailStatus.classList.toggle("hidden", dashboardViewer);
  recipientStatus.classList.toggle("hidden", dashboardViewer);
  if (statusRow) {
    statusRow.classList.toggle("hidden", dashboardViewer);
  }
  roleNotice.classList.toggle("hidden", dashboardViewer);
  dashboardCopy.textContent = dashboardViewer
    ? DASHBOARD_VIEWER_DASHBOARD_COPY
    : DEFAULT_DASHBOARD_COPY;
  entityPerformanceCopy.textContent = dashboardViewer
    ? DASHBOARD_VIEWER_ENTITY_PERFORMANCE_COPY
    : DEFAULT_ENTITY_PERFORMANCE_COPY;
  syncAiAnalystWidgetAvailability();
  roleNotice.textContent = editable
    ? "Editors can add investments, tasks, documents, and research."
    : dashboardViewer
      ? "This account is configured for a clean read-only family dashboard."
      : "Your account is view-only. You can review investments, research, and tasks, but editing is disabled.";
}

function currentFilters() {
  return {
    entity: entityFilter.value,
    search: searchFilter.value.trim().toLowerCase(),
    status: statusFilter.value,
    stage: stageFilter.value,
    owner: ownerFilter.value,
    portfolioPreset: activePortfolioPreset
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
    const matchesPipelinePreset =
      filters.portfolioPreset !== "pipeline" ||
      isPipelineStatus(investment.status) ||
      isPipelineStatus(investment.stage);

    return (
      matchesEntity &&
      matchesSearch &&
      matchesStatus &&
      matchesStage &&
      matchesOwner &&
      matchesPipelinePreset
    );
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
            reportUpdates: (company.reportUpdates || []).filter(
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

function isEntityPerformanceStatus(status) {
  return [
    "Approved",
    "Funded",
    "Active",
    "Partially Realized",
    "Realized",
    "Written Off",
    "Closed / Archived"
  ].includes(normalizeStatusName(status));
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
      const includeInEntityPerformance = isEntityPerformanceStatus(status);

      return {
        company,
        latest: company.latest,
        performance,
        reportedAmount,
        includedReportedAmount: includeReportedAmount ? reportedAmount : 0,
        includeReportedAmount,
        includeInEntityPerformance
      };
    });
}

function sumEntityRows(rows, selector) {
  return rows.reduce((sum, row) => sum + selector(row), 0);
}

function buildEntityRowTotals(rows) {
  const includedRows = rows.filter((row) => row.includeInEntityPerformance);
  const entityCompanies = includedRows.map((row) => row.company);
  const aggregatePerformance = buildAggregatePerformance(entityCompanies);

  return {
    reportedAmount: sumEntityRows(includedRows, (row) => row.includedReportedAmount),
    investedCapital: sumEntityRows(includedRows, (row) => row.performance.investedCapital),
    distributions: sumEntityRows(includedRows, (row) => row.performance.distributions),
    officialValue: sumEntityRows(includedRows, (row) => row.performance.officialValue),
    internalValue: sumEntityRows(includedRows, (row) => row.performance.internalValue),
    exitValue: sumEntityRows(includedRows, (row) => row.performance.exitValue),
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
      (company) =>
        normalizeEntityName(company.latest.entity) === normalizeEntityName(entity) &&
        isEntityPerformanceStatus(company.latest.status)
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
  const openReminderCount = isDashboardViewer()
    ? Number(digestStatus.openReminderCount || 0)
    : allTasks.filter(
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

  let cards = [
    { label: "Updates", value: String(investments.length), action: "portfolio" },
    { label: "Companies", value: String(companySummaries.length), action: "portfolio" },
    { label: "Pipeline deals", value: String(openCount), action: "pipeline" },
    {
      label: "Pipeline amount",
      value: formatMoney(openPipelineAmount),
      action: "pipeline"
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

  if (isDashboardViewer()) {
    cards = cards.filter((card) => !["Open reminders", "Data alerts"].includes(card.label));
  }

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
  const alertId =
    row.latest.id ||
    companyEntityKey(row.latest.company, row.latest.entity) ||
    `${normalizeCompanyKey(row.latest.company)}::${Date.now()}`;
  const alertKey = `${alertId}::${normalizeCompanyKey(title)}`;
  alerts.push({
    id: alertId,
    alertKey,
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
  return alerts
    .filter((alert) => !isAlertDismissed(alert.alertKey))
    .sort((left, right) => {
      const severityDelta = severityRank[left.severity] - severityRank[right.severity];
      if (severityDelta) {
        return severityDelta;
      }

      return left.company.localeCompare(right.company);
    });
}

function renderDashboard(investments) {
  const cards = buildDashboardCards(investments);
  const restrictedDashboardActions = new Set(["portfolio", "tasks", "quality"]);
  dashboardCards.innerHTML = cards
    .map(
      (card) => `
        <article
          class="dashboard-card"
          ${card.action && !(isDashboardViewer() && restrictedDashboardActions.has(card.action))
            ? `data-dashboard-action="${escapeHtml(card.action)}"`
            : ""}
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
            ${isDashboardViewer() ? "" : '<span class="entity-open-pill">Open entity</span>'}
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
                ${
                  canEditWorkspace()
                    ? `<button class="secondary-button card-action-button" type="button" data-action="dismiss-quality-alert" data-alert-key="${escapeHtml(
                        alert.alertKey
                      )}">Dismiss for 30 days</button>`
                    : ""
                }
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
                ${
                  canOpenCompanyDetails()
                    ? `<button class="link-button company-link" type="button" data-company="${escapeHtml(investment.company)}" data-entity="${escapeHtml(investment.entity || "")}">
                        ${escapeHtml(investment.company)}
                      </button>`
                    : `<h3>${escapeHtml(investment.company)}</h3>`
                }
                <span class="status-chip">${escapeHtml(normalizeStatusName(investment.status) || "Update")}</span>
              </div>
              <p class="update-meta">${escapeHtml(investment.stage || "Stage not set")} • Owner: ${escapeHtml(investment.owner || "Not set")}</p>
              <p class="update-meta">Official NAV ${escapeHtml(formatMoney(companyPerformance.officialValue))} • XIRR ${escapeHtml(formatPercent(companyPerformance.official.xirr))}</p>
              <p class="update-notes">${escapeHtml(summarizeText(investment.notes, "No notes provided."))}</p>
              ${
                canOpenCompanyDetails() || canEditWorkspace()
                  ? `<div class="card-actions">
                      ${
                        canOpenCompanyDetails()
                          ? `<button class="secondary-button card-action-button" type="button" data-action="view-company" data-company="${escapeHtml(investment.company)}" data-entity="${escapeHtml(investment.entity || "")}">View company</button>`
                          : ""
                      }
                      ${
                        canEditWorkspace()
                          ? `<button class="secondary-button card-action-button" type="button" data-action="edit" data-id="${investment.id}">Edit</button>`
                          : ""
                      }
                    </div>`
                  : ""
              }
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

function formatSummaryField(value, formatter = null) {
  if (value === null || value === undefined || value === "") {
    return "Not Available";
  }

  if (formatter) {
    return formatter(value);
  }

  return String(value).trim() || "Not Available";
}

function joinSummarySentences(parts, fallback = "Not Available") {
  const cleaned = parts
    .map((part) => String(part || "").trim())
    .filter(Boolean);

  return cleaned.length ? cleaned.join(" ") : fallback;
}

function buildInvestmentExecutiveSummary(latest, performance, latestUpdateText) {
  const status = normalizeStatusName(latest.status) || "Not Available";
  const investedCapital = formatSummaryField(performance.investedCapital, formatMoney);
  const currentValue = formatSummaryField(
    performance.internalValue || performance.officialValue,
    formatMoney
  );
  const returnedCapital = formatSummaryField(performance.distributions, formatMoney);
  const updateSentence =
    latestUpdateText && latestUpdateText !== "Not Available"
      ? `The latest operating note is: ${latestUpdateText}`
      : "";

  return joinSummarySentences([
    `${latest.company || "This investment"} is currently marked as ${status}.`,
    `Called capital stands at ${investedCapital}, with ${returnedCapital} returned to date and current value/NAV estimated at ${currentValue}.`,
    updateSentence
  ]);
}

function buildInvestmentChangeSummary(earliest, latest, performance) {
  const changes = [];
  const earliestStatus = normalizeStatusName(earliest.status);
  const latestStatus = normalizeStatusName(latest.status);
  if (earliestStatus || latestStatus) {
    changes.push(
      `Status moved from ${earliestStatus || "Not Available"} to ${latestStatus || "Not Available"}.`
    );
  }

  const earliestAmount = toNumber(earliest.amount);
  const latestAmount = toNumber(latest.amount);
  if (earliestAmount || latestAmount) {
    changes.push(
      `Reported commitment shifted from ${formatMoney(earliestAmount)} to ${formatMoney(latestAmount || earliestAmount)}.`
    );
  }

  if (performance.distributions > 0) {
    changes.push(`Capital returned now totals ${formatMoney(performance.distributions)}.`);
  }

  if (performance.internalValue || performance.officialValue) {
    changes.push(
      `Current marks stand at official NAV ${formatMoney(performance.officialValue)} and internal NAV ${formatMoney(performance.internalValue)}.`
    );
  }

  return joinSummarySentences(changes);
}

function buildInvestmentSummaryContext() {
  if (!selectedCompany) {
    return null;
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
    return null;
  }

  const latest = companyUpdates[0];
  const earliest = companyUpdates[companyUpdates.length - 1];
  const performance =
    companyPerformanceMap.get(companyEntityKey(selectedCompany, selectedCompanyEntity)) ||
    buildCompanyPerformance(companyUpdates);
  const latestUpdateText = summarizeText(
    latest.notes || latest.deckSummary || latest.decisionSummary,
    "Not Available"
  );
  const currentValue = performance.internalValue || performance.officialValue || null;
  const totalValue = currentValue + performance.distributions;
  const gainLoss = totalValue - performance.investedCapital;
  const moic = performance.internal.moic ?? performance.official.moic;
  const irr = performance.internal.xirr ?? performance.official.xirr;
  const originalInvestmentAmount = performance.committedCapital || toNumber(latest.amount) || null;
  const overview = joinSummarySentences([
    latest.stage ? `${latest.stage} investment.` : "",
    latest.entity ? `Tracked under ${normalizeEntityName(latest.entity)}.` : "",
    latest.owner ? `Current owner: ${latest.owner}.` : "",
    latest.contactName ? `Primary contact: ${latest.contactName}.` : "",
    latest.notes ? `Latest context: ${summarizeText(latest.notes, "")}` : ""
  ]);
  const risks = joinSummarySentences(
    [
      latest.followOnCapitalNotes,
      latest.ownershipNotes,
      latest.decisionSummary && latest.decisionType ? `${latest.decisionType}: ${latest.decisionSummary}` : ""
    ],
    "Not Available"
  );
  const recommendation = joinSummarySentences(
    [
      latest.nextStep ? `Next step: ${latest.nextStep}.` : "",
      latest.followOnCapitalStatus ? `Follow-on status: ${latest.followOnCapitalStatus}.` : "",
      latest.nextStepDueDate ? `Reminder date: ${formatDisplayDate(latest.nextStepDueDate)}.` : ""
    ],
    "Not Available"
  );

  return {
    latest,
    earliest,
    performance,
    originalInvestmentAmount,
    latestUpdateText,
    currentValue,
    totalValue,
    gainLoss,
    moic,
    irr,
    overview,
    risks,
    recommendation,
    executiveSummary: buildInvestmentExecutiveSummary(latest, performance, latestUpdateText),
    changeSummary: buildInvestmentChangeSummary(earliest, latest, performance)
  };
}

function renderInvestmentSummary() {
  const context = buildInvestmentSummaryContext();
  latestCompanySummaryContext = context;
  if (!context) {
    investmentSummaryPanel.classList.add("hidden");
    investmentSummaryDocument.innerHTML = "";
    return;
  }

  const {
    latest,
    performance,
    originalInvestmentAmount,
    currentValue,
    totalValue,
    gainLoss,
    moic,
    irr
  } = context;

  const sections = [
    {
      title: "Executive Summary",
      body: context.executiveSummary
    },
    {
      title: "Company / Fund Overview",
      body: context.overview
    },
    {
      title: "Key Updates",
      body: context.latestUpdateText
    },
    {
      title: "What Has Changed Since Investment",
      body: context.changeSummary
    },
    {
      title: "Current Risks / Open Questions",
      body: context.risks
    },
    {
      title: "Next Steps / Recommendation",
      body: context.recommendation
    },
    {
      title: "Notes for Lee / Internal Discussion",
      body: joinSummarySentences(
        [latest.decisionSummary, latest.documentLinks, latest.notes],
        "Not Available"
      )
    }
  ];

  if (
    investmentSummaryAiNotes &&
    companyKey(investmentSummaryAiNotes.company) === companyKey(selectedCompany) &&
    normalizeEntityName(investmentSummaryAiNotes.entity) === normalizeEntityName(selectedCompanyEntity)
  ) {
    sections.push({
      title: "AI Analyst Notes",
      body: investmentSummaryAiNotes.answer || "Not Available"
    });
  }

  investmentSummaryDocument.innerHTML = `
    <header class="investment-summary-header">
      <div>
        <p class="investment-summary-eyebrow">Investment Summary</p>
        <h2>${escapeHtml(latest.company || selectedCompany)}</h2>
        <p class="investment-summary-meta">${escapeHtml(
          normalizeEntityName(latest.entity) || "Not Available"
        )} • ${escapeHtml(normalizeStatusName(latest.status) || "Not Available")} • Prepared ${escapeHtml(
          formatDisplayDate(new Date().toISOString())
        )}</p>
      </div>
    </header>
    <section class="investment-summary-grid">
      <article class="investment-summary-kpi">
        <p class="dashboard-label">Original investment amount</p>
        <p class="dashboard-value">${escapeHtml(formatSummaryField(originalInvestmentAmount, formatMoney))}</p>
      </article>
      <article class="investment-summary-kpi">
        <p class="dashboard-label">Current value / NAV</p>
        <p class="dashboard-value">${escapeHtml(formatSummaryField(currentValue, formatMoney))}</p>
      </article>
      <article class="investment-summary-kpi">
        <p class="dashboard-label">Capital returned</p>
        <p class="dashboard-value">${escapeHtml(
          formatSummaryField(performance.distributions, formatMoney)
        )}</p>
      </article>
      <article class="investment-summary-kpi">
        <p class="dashboard-label">Total value realized + unrealized</p>
        <p class="dashboard-value">${escapeHtml(formatSummaryField(totalValue, formatMoney))}</p>
      </article>
      <article class="investment-summary-kpi">
        <p class="dashboard-label">Gain / Loss</p>
        <p class="dashboard-value">${escapeHtml(formatSummaryField(gainLoss, formatMoney))}</p>
      </article>
      <article class="investment-summary-kpi">
        <p class="dashboard-label">MOIC</p>
        <p class="dashboard-value">${escapeHtml(formatSummaryField(moic, formatTurns))}</p>
      </article>
      <article class="investment-summary-kpi">
        <p class="dashboard-label">IRR</p>
        <p class="dashboard-value">${escapeHtml(formatSummaryField(irr, formatPercent))}</p>
      </article>
      <article class="investment-summary-kpi">
        <p class="dashboard-label">Investment status</p>
        <p class="dashboard-value">${escapeHtml(
          formatSummaryField(normalizeStatusName(latest.status))
        )}</p>
      </article>
    </section>
    <section class="investment-summary-sections">
      ${sections
        .map(
          (section) => `
            <article class="investment-summary-section">
              <h3>${escapeHtml(section.title)}</h3>
              <p>${escapeHtml(section.body)}</p>
            </article>
          `
        )
        .join("")}
    </section>
    <footer class="investment-summary-footer">
      Prepared for internal discussion purposes only. Figures should be verified against source documents before making investment decisions.
    </footer>
  `;
}

function openInvestmentSummary() {
  renderInvestmentSummary();
  if (!latestCompanySummaryContext) {
    return;
  }

  investmentSummaryPanel.classList.remove("hidden");
  investmentSummaryPanel.scrollIntoView({ behavior: "smooth", block: "start" });
}

function closeInvestmentSummary() {
  latestCompanySummaryContext = null;
  investmentSummaryPanel.classList.add("hidden");
  investmentSummaryDocument.innerHTML = "";
  document.body.classList.remove("print-investment-summary");
}

function printInvestmentSummary() {
  if (investmentSummaryPanel.classList.contains("hidden")) {
    openInvestmentSummary();
  }

  document.body.classList.add("print-investment-summary");
  window.print();
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
  const valuationDate = String(values.valuationDate || "").trim();
  const nextStepDueDate = String(values.nextStepDueDate || "").trim();

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
    valuationDate,
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
    nextStepDueDate,
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
    closeInvestmentSummary();
    resetReportUpdateComposer();
    closeReportInsightPanel();
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
    reportUpdatesList.innerHTML = "";
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
    closeInvestmentSummary();
    resetReportUpdateComposer();
    closeReportInsightPanel();
    reportUpdatesList.innerHTML = "";
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
  const companyDocumentCount = companyRecord
    ? (Array.isArray(companyRecord.documents) ? companyRecord.documents.length : 0)
    : companyUpdates.reduce(
        (count, investment) => count + (Array.isArray(investment.documents) ? investment.documents.length : 0),
        0
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
    ? `${companyUpdates.length} update${companyUpdates.length === 1 ? "" : "s"} organized into one operating file for research, capital history, decisions, reminders, and source material.`
    : `${companyUpdates.length} update${companyUpdates.length === 1 ? "" : "s"} saved for this company in one operating record.`;
  companySummary.innerHTML = [
    { label: "Latest status", value: normalizeStatusName(latest.status) || "Not set" },
    { label: "Latest entity", value: normalizeEntityName(latest.entity) || "Not set" },
    { label: "Latest stage", value: latest.stage || "Not set" },
    { label: "Latest owner", value: latest.owner || "Not set" },
    { label: "Reported amount", value: formatMoney(totalAmount) },
    { label: "Open reminders", value: String(relatedTasks.filter((task) => task.status !== "Completed").length) },
    { label: "Vault documents", value: String(companyDocumentCount) },
    { label: "Research entries", value: String(deckSummaries.length) }
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

  renderReportUpdatesSection(companyRecord);

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

  if (!investmentSummaryPanel.classList.contains("hidden")) {
    renderInvestmentSummary();
  }
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
                <th>Valuation date</th>
                <th>Next step reminder</th>
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
                          <td>
                            ${
                              canEditWorkspace()
                                ? `<input class="reconciliation-amount-input" type="date" value="${escapeHtml(
                                    latest.valuationDate || ""
                                  )}" data-edit-input="true" data-field="valuationDate" data-id="${escapeHtml(latest.id)}" aria-label="Valuation date for ${escapeHtml(latest.company)}" />`
                                : escapeHtml(formatDisplayDateOrText(latest.valuationDate || ""))
                            }
                          </td>
                          <td>
                            ${
                              canEditWorkspace()
                                ? `<input class="reconciliation-amount-input" type="date" value="${escapeHtml(
                                    latest.nextStepDueDate || ""
                                  )}" data-edit-input="true" data-field="nextStepDueDate" data-id="${escapeHtml(latest.id)}" aria-label="Next step reminder date for ${escapeHtml(latest.company)}" />`
                                : escapeHtml(formatDisplayDateOrText(latest.nextStepDueDate || ""))
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
                : `<tr><td colspan="${canEditWorkspace() ? "13" : "12"}" class="update-meta">No investments are assigned to this entity.</td></tr>`}
            </tbody>
            <tfoot>
              <tr>
                <td colspan="5">Subtotal</td>
                <td></td>
                <td>${escapeHtml(formatMoney(entityPerformance.reportedAmount))}</td>
                <td>${escapeHtml(formatMoney(entityPerformance.investedCapital))}</td>
                <td>${escapeHtml(formatMoney(entityPerformance.officialValue))}</td>
                <td>${escapeHtml(formatMoney(entityPerformance.internalValue))}</td>
                <td></td>
                <td></td>
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
  dismissedDataAlerts = normalizeDismissedAlertMap(config.dismissedDataAlerts);

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
      ? `Use your email and your personal workspace password to sign in. ${config.teamUserCount} team login${config.teamUserCount === 1 ? "" : "s"} configured. Add :viewer for full read-only access or :dashboard-viewer for dashboard-only access in TEAM_USERS.`
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

  renderAll();
}

async function dismissDataQualityAlert(alertKey) {
  const normalizedKey = String(alertKey || "").trim();
  if (!normalizedKey) {
    return;
  }

  try {
    const result = await fetchJson("/api/data-alerts/dismiss", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ alertKey: normalizedKey })
    });
    dismissedDataAlerts = normalizeDismissedAlertMap(
      (result && result.dismissedDataAlerts) || {
        ...dismissedDataAlerts,
        [normalizedKey]: result.dismissedUntil
      }
    );
    renderDataQuality();
    renderDashboard(allInvestments);
  } catch (error) {
    reconciliationMessage.textContent =
      error.message || "Could not dismiss the data alert right now.";
  }
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

    if (error.status === 403 && isDashboardViewer()) {
      allTasks = [];
      renderAll();
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

addListener(deckFile, "change", () => {
  updateDeckFileLabel(selectedDeckFile());
});

["dragenter", "dragover"].forEach((eventName) => {
  addListener(deckDropZone, eventName, (event) => {
    event.preventDefault();
    deckDropZone.classList.add("deck-drop-zone-active");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  addListener(deckDropZone, eventName, (event) => {
    event.preventDefault();
    deckDropZone.classList.remove("deck-drop-zone-active");
  });
});

addListener(deckDropZone, "drop", (event) => {
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
  addListener(documentDropZone, eventName, (event) => {
    event.preventDefault();
    documentDropZone.classList.add("deck-drop-zone-active");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  addListener(documentDropZone, eventName, (event) => {
    event.preventDefault();
    documentDropZone.classList.remove("deck-drop-zone-active");
  });
});

addListener(documentDropZone, "drop", async (event) => {
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

addListener(documentFile, "change", async () => {
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
  addListener(companyDocumentDropZone, eventName, (event) => {
    event.preventDefault();
    companyDocumentDropZone.classList.add("deck-drop-zone-active");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  addListener(companyDocumentDropZone, eventName, (event) => {
    event.preventDefault();
    companyDocumentDropZone.classList.remove("deck-drop-zone-active");
  });
});

addListener(companyDocumentDropZone, "drop", async (event) => {
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

addListener(companyDocumentFile, "change", async () => {
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

addListener(menuToggleButton, "click", () => {
  if (!workspaceMenu) {
    return;
  }

  workspaceMenu.classList.toggle("hidden");
});

addListener(workspaceMenu, "click", (event) => {
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

addListener(loginForm, "submit", async (event) => {
  event.preventDefault();
  if (loginMessage) {
    loginMessage.textContent = "Signing in...";
  }
  if (loginButton) {
    loginButton.disabled = true;
  }

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
    if (loginMessage) {
      loginMessage.textContent = "Signed in. Your workspace is ready.";
    }
    await Promise.all([loadConfig(), loadUpdates()]);
  } catch (error) {
    if (loginMessage) {
      loginMessage.textContent = error.message;
    }
  } finally {
    if (loginButton) {
      loginButton.disabled = false;
    }
  }
});

addListener(taskForm, "submit", async (event) => {
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

addListener(logoutButton, "click", async () => {
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

addListener(summarizeDeckButton, "click", async () => {
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

addListener(summarizeEmailButton, "click", async () => {
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

if (openAiAnalystButton) {
  openAiAnalystButton.addEventListener("click", () => {
    openAiAnalystForSelectedCompany();
  });
}

if (aiAnalystWidgetToggleButton) {
  aiAnalystWidgetToggleButton.addEventListener("click", () => {
    const shouldOpen = aiAnalystSection ? aiAnalystSection.classList.contains("hidden") : true;
    setAiAnalystWidgetOpen(shouldOpen, { focusInput: shouldOpen });
  });
}

if (aiAnalystWidgetMinimizeButton) {
  aiAnalystWidgetMinimizeButton.addEventListener("click", () => {
    setAiAnalystWidgetMinimized(!aiAnalystWidgetMinimized);
  });
}

if (aiAnalystWidgetCloseButton) {
  aiAnalystWidgetCloseButton.addEventListener("click", () => {
    setAiAnalystWidgetOpen(false);
  });
}

if (aiAnalystSection) {
  aiAnalystSection.addEventListener("click", (event) => {
    const templateButton = event.target.closest("[data-analyst-template]");
    if (!templateButton) {
      return;
    }

    applyAiAnalystTemplate(templateButton.dataset.analystTemplate || "");
  });
}

if (askAiAnalystButton) {
  askAiAnalystButton.addEventListener("click", async () => {
    const question = String((aiAnalystPromptField && aiAnalystPromptField.value) || "").trim();
    const company = String((aiAnalystCompanyField && aiAnalystCompanyField.value) || "").trim();
    const entity = normalizeEntityName(
      String((aiAnalystEntityField && aiAnalystEntityField.value) || "").trim()
    );

    if (!question) {
      if (aiAnalystMessage) {
        aiAnalystMessage.textContent = "Ask the AI Analyst a question first.";
      }
      if (aiAnalystPromptField) {
        aiAnalystPromptField.focus();
      }
      return;
    }

    askAiAnalystButton.disabled = true;
    clearAiAnalystState(false);
    if (aiAnalystMessage) {
      aiAnalystMessage.textContent = "Analyzing portfolio context...";
    }

    try {
      const result = await fetchJson("/api/ai-agent", {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          question,
          company,
          entity
        })
      });

      renderAiAnalystResponse(result);
      if (aiAnalystMessage) {
        aiAnalystMessage.textContent = "AI Analyst response ready.";
      }
    } catch (error) {
      if (error.status === 401) {
        setSignedInState(null);
        if (aiAnalystMessage) {
          aiAnalystMessage.textContent = "Your session expired. Please sign in again.";
        }
        return;
      }

      if (aiAnalystMessage) {
        aiAnalystMessage.textContent = error.message;
      }
    } finally {
      askAiAnalystButton.disabled = false;
    }
  });
}

if (clearAiAnalystButton) {
  clearAiAnalystButton.addEventListener("click", () => {
    clearAiAnalystState(true);
  });
}

if (useAiAnalystInSummaryButton) {
  useAiAnalystInSummaryButton.addEventListener("click", () => {
    if (!latestAiAnalystResult) {
      return;
    }

    investmentSummaryAiNotes = {
      company: latestAiAnalystResult.company || selectedCompany,
      entity: normalizeEntityName(latestAiAnalystResult.entity || selectedCompanyEntity),
      answer: latestAiAnalystResult.answer || ""
    };

    if (investmentSummaryAiNotes.company) {
      selectedCompany = investmentSummaryAiNotes.company;
    }
    selectedCompanyEntity = investmentSummaryAiNotes.entity || "";
    renderCompanyPanel();
    showWorkspaceView("portfolio");
    openInvestmentSummary();

    if (aiAnalystMessage) {
      aiAnalystMessage.textContent = "AI Analyst notes added to the investment summary.";
    }
  });
}

addListener(form, "submit", async (event) => {
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

addListener(refreshButton, "click", () => {
  Promise.all([loadUpdates(), loadTasks()]).catch((error) => {
    formMessage.textContent = error.message;
  });
});

if (addReportUpdateButton) {
  addReportUpdateButton.addEventListener("click", () => {
    if (!selectedCompany) {
      if (reportUpdateMessage) {
        reportUpdateMessage.textContent = "Open an investment detail page first.";
      }
      return;
    }

    reportUpdateMessage.textContent = "";
    openReportUpdateComposer();
  });
}

if (cancelReportUpdateButton) {
  cancelReportUpdateButton.addEventListener("click", () => {
    resetReportUpdateComposer();
  });
}

if (closeReportUpdateInsightButton) {
  closeReportUpdateInsightButton.addEventListener("click", () => {
    closeReportInsightPanel();
  });
}

if (useReportInsightInSummaryButton) {
  useReportInsightInSummaryButton.addEventListener("click", () => {
    if (!latestReportInsight || !latestReportInsight.body) {
      return;
    }

    investmentSummaryAiNotes = {
      company: latestReportInsight.company || selectedCompany,
      entity: normalizeEntityName(latestReportInsight.entity || selectedCompanyEntity),
      answer: latestReportInsight.body
    };

    if (investmentSummaryAiNotes.company) {
      selectedCompany = investmentSummaryAiNotes.company;
    }
    selectedCompanyEntity = investmentSummaryAiNotes.entity || "";
    renderCompanyPanel();
    openInvestmentSummary();

    if (reportUpdateMessage) {
      reportUpdateMessage.textContent = "Lee-ready notes added to the investment summary.";
    }
  });
}

[reportUpdateTypeFilter, reportUpdatePeriodFilter].forEach((element) => {
  if (!element) {
    return;
  }

  element.addEventListener("change", () => {
    reportUpdateFilters = getReportUpdateFilters();
    renderReportUpdatesSection(findCompanyRecord(selectedCompany, selectedCompanyEntity));
  });
});

if (summarizeReportUpdateDraftButton) {
  summarizeReportUpdateDraftButton.addEventListener("click", async () => {
    const companyRecord = findCompanyRecord(selectedCompany, selectedCompanyEntity);
    const draft = collectReportUpdateFormData();
    if (!draft.title && !draft.originalNotes && !draft.keyWins && !draft.keyRisks && !draft.keyMetrics) {
      if (reportUpdateMessage) {
        reportUpdateMessage.textContent = "Add a title, pasted text, or key update details first.";
      }
      return;
    }

    summarizeReportUpdateDraftButton.disabled = true;

    try {
      const result = await runReportAnalystPrompt("single", [draft], companyRecord, {
        loadingMessage: "Summarizing draft update..."
      });
      if (reportUpdateAiSummaryField) {
        reportUpdateAiSummaryField.value = result.answer || "";
      }
      if (reportUpdateMessage) {
        reportUpdateMessage.textContent = "AI summary added to the draft update.";
      }
    } catch (error) {
      if (error.status === 401) {
        setSignedInState(null);
        if (reportUpdateMessage) {
          reportUpdateMessage.textContent = "Your session expired. Please sign in again.";
        }
        return;
      }

      if (reportUpdateMessage) {
        reportUpdateMessage.textContent = error.message;
      }
    } finally {
      summarizeReportUpdateDraftButton.disabled = false;
    }
  });
}

if (saveReportUpdateButton) {
  saveReportUpdateButton.addEventListener("click", async () => {
    const companyRecord = findCompanyRecord(selectedCompany, selectedCompanyEntity);
    const currentRows = companyRecord ? normalizeReportUpdateRows(companyRecord.reportUpdates) : [];
    const draft = collectReportUpdateFormData();
    if (!draft.date) {
      draft.date = new Date().toISOString().slice(0, 10);
    }

    if (!draft.title && !draft.originalNotes && !draft.aiSummary) {
      if (reportUpdateMessage) {
        reportUpdateMessage.textContent = "Add a title, notes, or AI summary before saving.";
      }
      return;
    }

    saveReportUpdateButton.disabled = true;

    try {
      await saveCompanyReportUpdates(
        [
          {
            ...draft,
            sourceUpdateId:
              (companyRecord && companyRecord.latest && companyRecord.latest.id) || ""
          },
          ...currentRows
        ],
        "Report update saved."
      );
      resetReportUpdateComposer();
      closeReportInsightPanel();
    } catch (error) {
      if (error.status === 401) {
        setSignedInState(null);
        if (reportUpdateMessage) {
          reportUpdateMessage.textContent = "Your session expired. Please sign in again.";
        }
        return;
      }

      if (reportUpdateMessage) {
        reportUpdateMessage.textContent = error.message;
      }
    } finally {
      saveReportUpdateButton.disabled = false;
    }
  });
}

if (compareLatestReportUpdatesButton) {
  compareLatestReportUpdatesButton.addEventListener("click", async () => {
    const companyRecord = findCompanyRecord(selectedCompany, selectedCompanyEntity);
    const rows = companyRecord ? normalizeReportUpdateRows(companyRecord.reportUpdates) : [];
    if (rows.length < 2) {
      if (reportUpdateMessage) {
        reportUpdateMessage.textContent = "Add at least two updates before comparing changes.";
      }
      return;
    }

    compareLatestReportUpdatesButton.disabled = true;

    try {
      const result = await runReportAnalystPrompt("compare", rows.slice(0, 2), companyRecord, {
        loadingMessage: "Comparing the latest two updates..."
      });
      renderReportInsight("What changed since last update?", result.answer, result.company, result.entity);
      if (reportUpdateMessage) {
        reportUpdateMessage.textContent = "Comparison ready.";
      }
    } catch (error) {
      if (error.status === 401) {
        setSignedInState(null);
        if (reportUpdateMessage) {
          reportUpdateMessage.textContent = "Your session expired. Please sign in again.";
        }
        return;
      }

      if (reportUpdateMessage) {
        reportUpdateMessage.textContent = error.message;
      }
    } finally {
      compareLatestReportUpdatesButton.disabled = false;
    }
  });
}

if (generateAllReportSummaryButton) {
  generateAllReportSummaryButton.addEventListener("click", async () => {
    const companyRecord = findCompanyRecord(selectedCompany, selectedCompanyEntity);
    const rows = companyRecord ? normalizeReportUpdateRows(companyRecord.reportUpdates) : [];
    if (!rows.length) {
      if (reportUpdateMessage) {
        reportUpdateMessage.textContent = "Add at least one update before generating a summary.";
      }
      return;
    }

    generateAllReportSummaryButton.disabled = true;

    try {
      const result = await runReportAnalystPrompt("lee-all", rows, companyRecord, {
        loadingMessage: "Generating Lee-ready summary across all updates..."
      });
      renderReportInsight(
        "Lee-ready summary across all updates",
        result.answer,
        result.company,
        result.entity
      );
      if (reportUpdateMessage) {
        reportUpdateMessage.textContent = "Lee-ready summary is ready.";
      }
    } catch (error) {
      if (error.status === 401) {
        setSignedInState(null);
        if (reportUpdateMessage) {
          reportUpdateMessage.textContent = "Your session expired. Please sign in again.";
        }
        return;
      }

      if (reportUpdateMessage) {
        reportUpdateMessage.textContent = error.message;
      }
    } finally {
      generateAllReportSummaryButton.disabled = false;
    }
  });
}

addListener(addCapitalActivityButton, "click", () => {
  const rows = collectCapitalActivityRows();
  rows.push({ date: "", type: "Investment Amount", amount: "", notes: "" });
  renderCapitalActivityRows(rows);
});

addListener(capitalActivityList, "click", (event) => {
  const target = event.target.closest("[data-action='remove-capital-activity']");
  if (!target) {
    return;
  }

  const index = Number(target.dataset.index);
  const rows = collectCapitalActivityRows().filter((_, rowIndex) => rowIndex !== index);
  renderCapitalActivityRows(rows);
});

addListener(capitalActivityList, "input", (event) => {
  const amountField = event.target.closest('[data-capital-field="amount"]');
  if (!amountField) {
    return;
  }

  amountField.value = normalizeMoneyString(amountField.value);
});

attachFormattedInputHandlers();
applyFormInputFormatting();

addListener(loadCompanyDetailsButton, "click", () => {
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

if (form && form.elements && form.elements.company) {
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
}

addListener(downloadCsvButton, "click", () => {
  window.location.href = "/api/investments.csv";
});

addListener(downloadExcelButton, "click", () => {
  window.location.href = "/api/investments.xlsx";
});

addListener(downloadFamilyOfficeWorkbookButton, "click", () => {
  window.location.href = "/api/family-office-workbook.xlsx";
});

addListener(downloadReconciliationButton, "click", () => {
  downloadTextFile("entity-reconciliation.csv", buildReconciliationCsv());
});

addListener(saveAllReconciliationButton, "click", async () => {
  try {
    await saveReconciliationRows(Array.from(dirtyReconciliationRows));
  } catch (error) {
    if (reconciliationMessage) {
      reconciliationMessage.textContent = error.message;
    }
  }
});

addListener(downloadBackupButton, "click", () => {
  window.location.href = "/api/backup-export";
});

addListener(previewDigestButton, "click", async () => {
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

addListener(sendDigestButton, "click", async () => {
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

addListener(importWorkbookFile, "change", async () => {
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

addListener(restoreBackupFile, "change", async () => {
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

addListener(cancelEditButton, "click", () => {
  resetFormToCreateMode();
  formMessage.textContent = "Edit canceled.";
});

addListener(cancelTaskEditButton, "click", () => {
  resetTaskForm();
  taskMessage.textContent = "Task edit canceled.";
});

[entityFilter, searchFilter, statusFilter, stageFilter, ownerFilter].forEach((element) => {
  const clearPresetAndRender = () => {
    activePortfolioPreset = "";
    renderAll();
  };
  addListener(element, "input", clearPresetAndRender);
  addListener(element, "change", clearPresetAndRender);
});

addListener(uploadedDocumentsList, "click", (event) => {
  const action = event.target.closest("[data-action='remove-document']");
  if (!action) {
    return;
  }

  const documentId = action.dataset.documentId;
  uploadedDocuments = uploadedDocuments.filter((document) => document.id !== documentId);
  renderUploadedDocuments();
});

addListener(closeCompanyPanelButton, "click", () => {
  selectedCompany = "";
  selectedCompanyEntity = "";
  renderCompanyPanel();
  showWorkspaceView("portfolio");
});

addListener(generateInvestmentSummaryButton, "click", () => {
  openInvestmentSummary();
});

addListener(closeInvestmentSummaryButton, "click", () => {
  closeInvestmentSummary();
});

addListener(printInvestmentSummaryButton, "click", () => {
  printInvestmentSummary();
});

addListener(window, "afterprint", () => {
  document.body.classList.remove("print-investment-summary");
});

addListener(closeEntityDetailButton, "click", () => {
  selectedEntity = "";
  renderEntityDetail();
  showWorkspaceView("home");
});

addListener(entityPerformanceCards, "click", (event) => {
  if (isDashboardViewer()) {
    return;
  }

  const card = event.target.closest("[data-entity]");
  if (!card) {
    return;
  }

  selectedEntity = card.dataset.entity || "";
  renderEntityDetail();
  showWorkspaceView("entity");
  window.scrollTo({ top: 0, behavior: "smooth" });
});

addListener(dashboardCards, "click", (event) => {
  if (isDashboardViewer()) {
    return;
  }

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
    activePortfolioPreset = "";
    if (statusFilter) {
      statusFilter.value = status || "";
    }
    renderAll();
    showWorkspaceView("portfolio");
    window.scrollTo({ top: 0, behavior: "smooth" });
    return;
  }

  if (action === "pipeline") {
    activePortfolioPreset = "pipeline";
    if (statusFilter) {
      statusFilter.value = "";
    }
    renderAll();
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

addListener(companyDecisionLog, "click", (event) => {
  const target = event.target.closest("[data-action='delete-company-document']");
  if (!target) {
    return;
  }

  deleteCompanyDocumentById(target.dataset.documentId);
});

if (reportUpdatesList) {
  reportUpdatesList.addEventListener("click", async (event) => {
    const target = event.target.closest("[data-action]");
    if (!target) {
      return;
    }

    const action = target.dataset.action || "";
    const index = Number(target.dataset.index);
    const companyRecord = findCompanyRecord(selectedCompany, selectedCompanyEntity);
    const rows = companyRecord
      ? normalizeReportUpdateRows(companyRecord.reportUpdates).sort(
          (left, right) => new Date(right.date || 0).getTime() - new Date(left.date || 0).getTime()
        )
      : [];
    const row = Number.isFinite(index) ? rows[index] : null;

    if (!row) {
      if (reportUpdateMessage) {
        reportUpdateMessage.textContent = "Could not find that report update.";
      }
      return;
    }

    target.disabled = true;

    try {
      if (action === "summarize-report-update") {
        const result = await runReportAnalystPrompt("single", [row], companyRecord, {
          loadingMessage: "Summarizing selected update..."
        });
        const updatedRows = rows.map((entry, rowIndex) =>
          rowIndex === index ? { ...entry, aiSummary: result.answer || entry.aiSummary || "" } : entry
        );
        await saveCompanyReportUpdates(updatedRows, "AI summary saved to the update.");
        renderReportInsight("Update summary", result.answer, result.company, result.entity);
        return;
      }

      if (action === "generate-report-update-summary") {
        const result = await runReportAnalystPrompt("lee-single", [row], companyRecord, {
          loadingMessage: "Generating Lee-ready update summary..."
        });
        renderReportInsight("Lee-ready update summary", result.answer, result.company, result.entity);
        if (reportUpdateMessage) {
          reportUpdateMessage.textContent = "Lee-ready summary is ready.";
        }
      }
    } catch (error) {
      if (error.status === 401) {
        setSignedInState(null);
        if (reportUpdateMessage) {
          reportUpdateMessage.textContent = "Your session expired. Please sign in again.";
        }
        return;
      }

      if (reportUpdateMessage) {
        reportUpdateMessage.textContent = error.message;
      }
    } finally {
      target.disabled = false;
    }
  });
}

addListener(updatesList, "click", (event) => {
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

addListener(entityDetailInvestments, "click", (event) => {
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

addListener(reconciliationList, "click", (event) => {
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

addListener(reconciliationList, "input", handleReconciliationFieldEdit);
addListener(reconciliationList, "change", handleReconciliationFieldEdit);

addListener(dataQualityList, "click", (event) => {
  const dismissTarget = event.target.closest("[data-action='dismiss-quality-alert']");
  if (dismissTarget) {
    dismissDataQualityAlert(dismissTarget.dataset.alertKey || "");
    return;
  }

  const target = event.target.closest("[data-action='edit-quality-investment']");
  if (!target) {
    return;
  }

  beginEditInvestment(target.dataset.id || "");
});

addListener(tasksList, "click", (event) => {
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
