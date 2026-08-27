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
const importWorkbookLabel = document.querySelector("label[for='importWorkbookFile']");
const restoreBackupLabel = document.querySelector("label[for='restoreBackupFile']");
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
const publicStockSummary = document.getElementById("publicStockSummary");
const publicStockList = document.getElementById("publicStockList");
const publicStockEntityFilter = document.getElementById("publicStockEntityFilter");
const publicStockSearchFilter = document.getElementById("publicStockSearchFilter");
const refreshPublicStockPricesButton = document.getElementById("refreshPublicStockPricesButton");
const publicStockPriceMessage = document.getElementById("publicStockPriceMessage");
const cashSummary = document.getElementById("cashSummary");
const cashList = document.getElementById("cashList");
const cashEntityFilter = document.getElementById("cashEntityFilter");
const cashInstitutionFilter = document.getElementById("cashInstitutionFilter");
const fixedIncomeSummary = document.getElementById("fixedIncomeSummary");
const fixedIncomeLadder = document.getElementById("fixedIncomeLadder");
const fixedIncomeList = document.getElementById("fixedIncomeList");
const fixedIncomeEntityFilter = document.getElementById("fixedIncomeEntityFilter");
const fixedIncomeTypeFilter = document.getElementById("fixedIncomeTypeFilter");
const fixedIncomeSearchFilter = document.getElementById("fixedIncomeSearchFilter");
const fixedIncomeMaturityFilter = document.getElementById("fixedIncomeMaturityFilter");
const realEstateSummary = document.getElementById("realEstateSummary");
const realEstateList = document.getElementById("realEstateList");
const realEstateEntityFilter = document.getElementById("realEstateEntityFilter");
const realEstateTypeFilter = document.getElementById("realEstateTypeFilter");
const realEstateSearchFilter = document.getElementById("realEstateSearchFilter");
const addPublicSquareStockButton = document.getElementById("addPublicSquareStockButton");
const addSpaceXStockButton = document.getElementById("addSpaceXStockButton");
const realEstateDetailsPanel = document.getElementById("realEstateDetailsPanel");
const realEstatePropertyNameField = document.getElementById("realEstatePropertyNameField");
const realEstateAddressField = document.getElementById("realEstateAddressField");
const realEstateCityField = document.getElementById("realEstateCityField");
const realEstateStateField = document.getElementById("realEstateStateField");
const realEstateZipField = document.getElementById("realEstateZipField");
const realEstateEntityOwnerField = document.getElementById("realEstateEntityOwnerField");
const realEstatePropertyTypeField = document.getElementById("realEstatePropertyTypeField");
const realEstateOwnershipPercentField = document.getElementById("realEstateOwnershipPercentField");
const realEstateOwnershipNotesField = document.getElementById("realEstateOwnershipNotesField");
const realEstateAcquisitionDateField = document.getElementById("realEstateAcquisitionDateField");
const realEstatePurchasePriceField = document.getElementById("realEstatePurchasePriceField");
const realEstateCostBasisField = document.getElementById("realEstateCostBasisField");
const realEstateAppraisedValueField = document.getElementById("realEstateAppraisedValueField");
const realEstateAppraisalDateField = document.getElementById("realEstateAppraisalDateField");
const realEstateAppraiserField = document.getElementById("realEstateAppraiserField");
const realEstateAppraisalDocumentField = document.getElementById("realEstateAppraisalDocumentField");
const realEstateInternalValueOverrideField = document.getElementById("realEstateInternalValueOverrideField");
const realEstateInternalValueDateField = document.getElementById("realEstateInternalValueDateField");
const realEstateDebtField = document.getElementById("realEstateDebtField");
const realEstateLoanLenderField = document.getElementById("realEstateLoanLenderField");
const realEstateDebtInterestRateField = document.getElementById("realEstateDebtInterestRateField");
const realEstateDebtMaturityDateField = document.getElementById("realEstateDebtMaturityDateField");
const realEstateDebtServiceField = document.getElementById("realEstateDebtServiceField");
const realEstateLoanNotesField = document.getElementById("realEstateLoanNotesField");
const realEstateNoiField = document.getElementById("realEstateNoiField");
const realEstateRevenueField = document.getElementById("realEstateRevenueField");
const realEstateCapRateField = document.getElementById("realEstateCapRateField");
const realEstateOccupancyField = document.getElementById("realEstateOccupancyField");
const realEstateSquareFootageField = document.getElementById("realEstateSquareFootageField");
const realEstateAcreageField = document.getElementById("realEstateAcreageField");
const realEstateUnitsField = document.getElementById("realEstateUnitsField");
const realEstatePropertyTaxesField = document.getElementById("realEstatePropertyTaxesField");
const realEstateInsuranceField = document.getElementById("realEstateInsuranceField");
const realEstateOtherExpensesField = document.getElementById("realEstateOtherExpensesField");
const realEstateOperatingNotesField = document.getElementById("realEstateOperatingNotesField");
const realEstateValuePreview = document.getElementById("realEstateValuePreview");
const bondDetailsPanel = document.getElementById("bondDetailsPanel");
const bondIssuerField = document.getElementById("bondIssuerField");
const bondDescriptionField = document.getElementById("bondDescriptionField");
const bondTypeField = document.getElementById("bondTypeField");
const bondCusipField = document.getElementById("bondCusipField");
const bondEntityOwnerField = document.getElementById("bondEntityOwnerField");
const bondParValueField = document.getElementById("bondParValueField");
const bondPurchasePriceField = document.getElementById("bondPurchasePriceField");
const bondPurchaseDateField = document.getElementById("bondPurchaseDateField");
const bondCostBasisField = document.getElementById("bondCostBasisField");
const bondCouponRateField = document.getElementById("bondCouponRateField");
const bondCouponFrequencyField = document.getElementById("bondCouponFrequencyField");
const bondMaturityDateField = document.getElementById("bondMaturityDateField");
const bondCallDateField = document.getElementById("bondCallDateField");
const bondCallPriceField = document.getElementById("bondCallPriceField");
const bondCurrentPriceField = document.getElementById("bondCurrentPriceField");
const bondMarketPriceDateField = document.getElementById("bondMarketPriceDateField");
const bondMarketValueField = document.getElementById("bondMarketValueField");
const bondYieldToMaturityField = document.getElementById("bondYieldToMaturityField");
const bondYieldToCallField = document.getElementById("bondYieldToCallField");
const bondCurrentYieldField = document.getElementById("bondCurrentYieldField");
const bondCreditRatingField = document.getElementById("bondCreditRatingField");
const bondInsurerField = document.getElementById("bondInsurerField");
const bondTaxStatusField = document.getElementById("bondTaxStatusField");
const bondAccruedInterestField = document.getElementById("bondAccruedInterestField");
const bondValuePreview = document.getElementById("bondValuePreview");
const cashDetailsPanel = document.getElementById("cashDetailsPanel");
const cashAccountNameField = document.getElementById("cashAccountNameField");
const cashInstitutionField = document.getElementById("cashInstitutionField");
const cashAccountTypeField = document.getElementById("cashAccountTypeField");
const cashBalanceField = document.getElementById("cashBalanceField");
const cashBalanceDateField = document.getElementById("cashBalanceDateField");
const stockDetailsPanel = document.getElementById("stockDetailsPanel");
const stockValuePreview = document.getElementById("stockValuePreview");
const fetchStockQuoteButton = document.getElementById("fetchStockQuoteButton");
const stockQuoteMessage = document.getElementById("stockQuoteMessage");
const valuationHelperText = document.getElementById("valuationHelperText");
const entityDetailSection = document.getElementById("entityDetailSection");
const entityDetailTitle = document.getElementById("entityDetailTitle");
const entityDetailCopy = document.getElementById("entityDetailCopy");
const entityDetailSummary = document.getElementById("entityDetailSummary");
const entityDetailInvestments = document.getElementById("entityDetailInvestments");
const closeEntityDetailButton = document.getElementById("closeEntityDetailButton");
const xirrAuditEntitySelect = document.getElementById("xirrAuditEntitySelect");
const xirrAuditSummary = document.getElementById("xirrAuditSummary");
const xirrAuditTable = document.getElementById("xirrAuditTable");
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
const requestLatestUpdateButton = document.getElementById("requestLatestUpdateButton");
const updateRequestModal = document.getElementById("updateRequestModal");
const updateRequestForm = document.getElementById("updateRequestForm");
const updateRequestRecipientField = document.getElementById("updateRequestRecipientField");
const updateRequestSubjectField = document.getElementById("updateRequestSubjectField");
const updateRequestBodyField = document.getElementById("updateRequestBodyField");
const updateRequestMaterialsList = document.getElementById("updateRequestMaterialsList");
const updateRequestMessage = document.getElementById("updateRequestMessage");
const sendUpdateRequestButton = document.getElementById("sendUpdateRequestButton");
const closeUpdateRequestModalButton = document.getElementById("closeUpdateRequestModalButton");
const cancelUpdateRequestButton = document.getElementById("cancelUpdateRequestButton");
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
const investmentTimeline = document.getElementById("investmentTimeline");
const companyTimeline = document.getElementById("companyTimeline");
const reportUpdateTypeFilter = document.getElementById("reportUpdateTypeFilter");
const reportUpdatePeriodFilter = document.getElementById("reportUpdatePeriodFilter");
const addReportUpdateButton = document.getElementById("addReportUpdateButton");
const compareLatestReportUpdatesButton = document.getElementById(
  "compareLatestReportUpdatesButton"
);
const generateInvestmentHistorySummaryButton = document.getElementById(
  "generateInvestmentHistorySummaryButton"
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
const reportUpdateFileDropZone = document.getElementById("reportUpdateFileDropZone");
const reportUpdateFileInput = document.getElementById("reportUpdateFileInput");
const reportUpdateFileName = document.getElementById("reportUpdateFileName");
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
const aiUpdateInboxSummary = document.getElementById("aiUpdateInboxSummary");
const aiUpdateInboxMessage = document.getElementById("aiUpdateInboxMessage");
const aiUpdateInboxList = document.getElementById("aiUpdateInboxList");
const aiUpdateProposalDetail = document.getElementById("aiUpdateProposalDetail");
const openAiUpdateAnalyzerButton = document.getElementById("openAiUpdateAnalyzerButton");
const aiUpdateAnalyzerPanel = document.getElementById("aiUpdateAnalyzerPanel");
const aiUpdateAnalysisForm = document.getElementById("aiUpdateAnalysisForm");
const cancelAiUpdateAnalysisButton = document.getElementById("cancelAiUpdateAnalysisButton");
const runAiUpdateAnalysisButton = document.getElementById("runAiUpdateAnalysisButton");
const aiAnalysisInvestmentField = document.getElementById("aiAnalysisInvestmentField");
const aiAnalysisEntityField = document.getElementById("aiAnalysisEntityField");
const aiUpdateAnalysisReview = document.getElementById("aiUpdateAnalysisReview");
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
let allAiUpdateProposals = [];
let aiUpdateProposalCounts = {
  pending: 0,
  approved: 0,
  rejected: 0
};
let selectedAiUpdateProposalId = "";
let latestAiUpdateAnalysis = null;
let activeWorkspaceView = "home";
let selectedEntity = "";
let selectedXirrAuditEntity = "";
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
let activeUpdateRequestInvestment = null;
let investmentsLoaded = false;
let investmentsLoadError = "";
let reportUpdateFilters = {
  type: "",
  period: ""
};
let publicStockFilters = {
  entity: "",
  search: ""
};
let cashFilters = {
  entity: "",
  institution: ""
};
let fixedIncomeFilters = {
  entity: "",
  type: "",
  search: "",
  maturityYear: ""
};
let realEstateFilters = {
  entity: "",
  type: "",
  search: ""
};
let assetViewEntityFilterSource = {
  "public-stocks": "",
  "fixed-income": "",
  cash: "",
  "real-estate": ""
};
let publicStockQuoteRequestKeys = new Set();
let automaticPublicStockRefreshInFlight = false;

const REPORT_UPDATE_TYPES = [
  "Monthly",
  "Quarterly",
  "Annual",
  "Update Request",
  "Capital Call",
  "Legal",
  "Call Notes",
  "Other"
];

const REPORT_SOURCE_TYPES = ["PDF", "Email", "Call", "Manual Note", "Other"];
const PUBLIC_STOCK_PRESETS = [
  {
    company: "SpaceX",
    entity: "Beaman Ventures",
    assetType: "Public Stock",
    ticker: "SPCX",
    exchange: "NASDAQ",
    currency: "USD",
    stage: "Public Equity",
    status: "Active",
    notes: "SpaceX public stock position."
  },
  {
    company: "PublicSquare",
    entity: "Beaman Ventures",
    assetType: "Public Stock",
    ticker: "PSQH",
    exchange: "NYSE",
    currency: "USD",
    stage: "Public Equity",
    status: "Active",
    notes: "PublicSquare public stock position."
  }
];
const UPDATE_REQUEST_MATERIALS = [
  "Latest investor update",
  "Updated financials",
  "Current cash balance / runway",
  "Revenue / EBITDA metrics",
  "Pipeline updates",
  "Capital needs",
  "Major risks or changes",
  "Updated cap table"
];
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
  "followOnCapitalAmount",
  "costBasisPerShare",
  "marketPrice",
  "bondParValue",
  "bondPurchasePrice",
  "bondCostBasis",
  "bondCallPrice",
  "bondCurrentPrice",
  "bondMarketValue",
  "bondAccruedInterest",
  "realEstatePurchasePrice",
  "realEstateCostBasis",
  "realEstateAppraisedValue",
  "realEstateInternalValueOverride",
  "realEstateDebt",
  "realEstateDebtService",
  "realEstateNoi",
  "realEstateRevenue",
  "realEstateSquareFootage",
  "realEstateAcreage",
  "realEstatePropertyTaxes",
  "realEstateInsurance",
  "realEstateOtherExpenses"
];
const moneyFieldDecimalPlaces = {
  shareCount: 6,
  costBasisPerShare: 6,
  marketPrice: 6,
  bondPurchasePrice: 6,
  bondCallPrice: 6,
  bondCurrentPrice: 6
};

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
  "Lee Beaman IRA": "Lee Beaman IRA",
  "Lee Beaman Ira": "Lee Beaman IRA",
  "Lee's IRA": "Lee Beaman IRA",
  "Lees IRA": "Lee Beaman IRA",
  "Lee IRA": "Lee Beaman IRA",
  "Lee Beaman Individual Retirement Account": "Lee Beaman IRA",
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
  const aliasKey = Object.keys(ENTITY_ALIASES).find(
    (key) => key.toLowerCase() === raw.toLowerCase()
  );
  return ENTITY_ALIASES[raw] || (aliasKey ? ENTITY_ALIASES[aliasKey] : raw);
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

function normalizeMoneyString(value, decimalPlaces = 2) {
  const text = String(value || "").trim();
  if (!text) {
    return "";
  }

  const cleaned = text.replace(/[^0-9.]/g, "");
  const [integerPartRaw, decimalPartRaw = ""] = cleaned.split(".");
  const integerPart = integerPartRaw.replace(/^0+(?=\d)/, "") || integerPartRaw || "0";
  const decimalPart = decimalPartRaw.replace(/\./g, "").slice(0, decimalPlaces);
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

  field.value = normalizeMoneyString(field.value, moneyFieldDecimalPlaces[field.name] || 2);
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
  const roleLabel = user && (user.roleLabel || getRoleLabel(user.role));

  loginPanel.classList.toggle("hidden", isSignedIn);
  appPanel.classList.toggle("hidden", !isSignedIn);
  logoutButton.classList.toggle("hidden", !isSignedIn);
  authStatus.textContent = isSignedIn
    ? dashboardViewer
      ? getDashboardViewerGreeting(user)
      : `Signed in as ${user.email}${roleLabel ? ` • Role: ${roleLabel}` : ""}`
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

  maybeAutoRefreshPublicStockPrices();
}

function setAssetViewEntityFilter(viewName, entity, source = "manual") {
  const normalizedEntity = normalizeEntityName(entity);
  if (viewName === "public-stocks") {
    publicStockFilters.entity = normalizedEntity;
    assetViewEntityFilterSource["public-stocks"] = normalizedEntity ? source : "";
    renderPublicStocks();
    return;
  }
  if (viewName === "fixed-income") {
    fixedIncomeFilters.entity = normalizedEntity;
    assetViewEntityFilterSource["fixed-income"] = normalizedEntity ? source : "";
    renderFixedIncome();
    return;
  }
  if (viewName === "cash") {
    cashFilters.entity = normalizedEntity;
    assetViewEntityFilterSource.cash = normalizedEntity ? source : "";
    renderCash();
    return;
  }
  if (viewName === "real-estate") {
    realEstateFilters.entity = normalizedEntity;
    assetViewEntityFilterSource["real-estate"] = normalizedEntity ? source : "";
    renderRealEstate();
  }
}

function openAssetViewForEntity(viewName, entity) {
  setAssetViewEntityFilter(viewName, entity, "entity-card");
  showWorkspaceView(viewName);
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function resetEntityCardFilterForMenuView(viewName) {
  if (!assetViewEntityFilterSource[viewName] || assetViewEntityFilterSource[viewName] !== "entity-card") {
    return;
  }
  setAssetViewEntityFilter(viewName, "", "");
}

function maybeAutoRefreshPublicStockPrices() {
  if (activeWorkspaceView !== "public-stocks" || !investmentsLoaded || !canEditWorkspace()) {
    return;
  }
  refreshPublicStockPrices({ automatic: true });
}

function canEditWorkspace() {
  return !currentUser || !["viewer", "dashboard-viewer"].includes(currentUser.role);
}

function isMasterEditor() {
  return Boolean(currentUser && currentUser.role === "master-editor");
}

function isDashboardViewer() {
  return Boolean(currentUser && currentUser.role === "dashboard-viewer");
}

function getRoleLabel(role) {
  if (role === "master-editor") {
    return "Master Editor";
  }
  if (role === "dashboard-viewer") {
    return "Lee Dashboard";
  }
  return "Editor";
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
    assetType: investment.assetType,
    ticker: investment.ticker,
    exchange: investment.exchange,
    shareCount: investment.shareCount,
    costBasisPerShare: investment.costBasisPerShare,
    marketPrice: investment.marketPrice,
    marketPriceDate: investment.marketPriceDate,
    marketValue: investment.marketValue,
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
    reportingCadence: investment.reportingCadence,
    updateRequestStatus: investment.updateRequestStatus,
    lastUpdateRequestSentAt: investment.lastUpdateRequestSentAt,
    lastUpdateRequestSubject: investment.lastUpdateRequestSubject,
    lastUpdateRequestContact: investment.lastUpdateRequestContact,
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

function getLatestCompanyInvestment() {
  const companyRecord = findCompanyRecord(selectedCompany, selectedCompanyEntity);
  if (companyRecord && companyRecord.latest) {
    return companyRecord.latest;
  }

  return allInvestments.find(
    (investment) =>
      companyKey(investment.company) === companyKey(selectedCompany) &&
      normalizeEntityName(investment.entity) === normalizeEntityName(selectedCompanyEntity)
  );
}

function buildUpdateRequestDraft(investment, materials = []) {
  const companyName = investment && investment.company ? investment.company : "the company";
  const contactName = investment && investment.contactName ? investment.contactName : "there";
  const subject = `Request for Latest Update – ${companyName}`;
  const materialsLine = materials.length
    ? `\n\nSpecifically, it would be helpful to include: ${materials.join(", ")}.`
    : "";
  const body = [
    `Hi ${contactName},`,
    "",
    `I hope you’re doing well. I’m working through our investment updates and wanted to see if you could send over the latest update for ${companyName} when you have a chance.`,
    "",
    `If available, it would be helpful to include any recent investor materials, updated financials, current cash/runway, revenue or operating metrics, major developments, and any expected capital needs or key risks we should be aware of.${materialsLine}`,
    "",
    "Thanks,",
    "Tyler"
  ].join("\n");

  return { subject, body };
}

function getSelectedUpdateRequestMaterials() {
  return Array.from(
    updateRequestMaterialsList
      ? updateRequestMaterialsList.querySelectorAll("input[type='checkbox']:checked")
      : []
  ).map((input) => input.value);
}

function refreshUpdateRequestBody() {
  if (!activeUpdateRequestInvestment || !updateRequestBodyField) {
    return;
  }

  const draft = buildUpdateRequestDraft(
    activeUpdateRequestInvestment,
    getSelectedUpdateRequestMaterials()
  );
  updateRequestBodyField.value = draft.body;
}

function openUpdateRequestModal() {
  const investment = getLatestCompanyInvestment();
  activeUpdateRequestInvestment = investment || null;
  if (!investment) {
    if (reportUpdateMessage) {
      reportUpdateMessage.textContent = "Open an investment detail page first.";
    }
    return;
  }

  if (!investment.contactEmail) {
    if (reportUpdateMessage) {
      reportUpdateMessage.textContent = "Add a primary contact email before requesting an update.";
    }
    return;
  }

  const defaultMaterials = [
    "Latest investor update",
    "Updated financials",
    "Current cash balance / runway",
    "Revenue / EBITDA metrics",
    "Capital needs",
    "Major risks or changes"
  ];
  const draft = buildUpdateRequestDraft(investment, defaultMaterials);

  if (updateRequestRecipientField) {
    updateRequestRecipientField.value = investment.contactEmail || "";
  }
  if (updateRequestSubjectField) {
    updateRequestSubjectField.value = draft.subject;
  }
  if (updateRequestBodyField) {
    updateRequestBodyField.value = draft.body;
  }
  if (updateRequestMaterialsList) {
    updateRequestMaterialsList.innerHTML = UPDATE_REQUEST_MATERIALS.map(
      (material) => `
        <label class="checkbox-row">
          <input type="checkbox" value="${escapeHtml(material)}"${defaultMaterials.includes(material) ? " checked" : ""} />
          <span>${escapeHtml(material)}</span>
        </label>
      `
    ).join("");
  }
  if (updateRequestMessage) {
    updateRequestMessage.textContent = "";
  }
  if (updateRequestModal) {
    updateRequestModal.classList.remove("hidden");
  }
}

function closeUpdateRequestModal() {
  activeUpdateRequestInvestment = null;
  if (updateRequestModal) {
    updateRequestModal.classList.add("hidden");
  }
  if (updateRequestForm) {
    updateRequestForm.reset();
  }
  if (updateRequestMaterialsList) {
    updateRequestMaterialsList.innerHTML = "";
  }
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
        : mode === "history-summary"
          ? `Create an investment history summary for ${companyName} using all provided updates over time.`
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
      : mode === "history-summary"
        ? "Use these exact sections: Original Thesis, Key Developments Over Time, Current Status, Major Risks, Changes Since Investment, Recommended Next Step. Be concise, internal-ready, and do not make up facts. If information is missing, say what is missing."
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

function updateReportUpdateFileLabel(file) {
  if (!reportUpdateFileName) {
    return;
  }

  reportUpdateFileName.textContent = file
    ? `Selected: ${file.name}`
    : "PDF or text files can fill the summary fields below.";
}

function normalizeCapitalActivityRows(rows) {
  return (Array.isArray(rows) ? rows : [])
    .map((row) => ({
      date: String((row && row.date) || "").trim(),
      type: String((row && row.type) || "").trim(),
      amount: String((row && row.amount) || "").trim(),
      notes: String((row && row.notes) || "").trim(),
      legacyFallback: Boolean(row && row.legacyFallback)
    }))
    .filter(
      (row) => row.date || row.type || row.amount || row.notes || row.legacyFallback
    );
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
      contactEmailed: String((row && row.contactEmailed) || "").trim(),
      subjectLine: String((row && row.subjectLine) || "").trim(),
      responseStatus: String((row && row.responseStatus) || "").trim(),
      materialsRequested: Array.isArray(row && row.materialsRequested)
        ? row.materialsRequested.map((item) => String(item).trim()).filter(Boolean)
        : [],
      sourceUpdateId: String((row && row.sourceUpdateId) || "").trim()
    }))
    .filter((row) =>
      Object.values(row).some((value) =>
        Array.isArray(value) ? value.length > 0 : Boolean(value)
      )
    );
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
  updateReportUpdateFileLabel(null);
  if (reportUpdateFileInput) {
    reportUpdateFileInput.value = "";
  }
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

function getReportUpdateFileSourceType(file) {
  const name = String((file && file.name) || "").toLowerCase();
  if (name.endsWith(".pdf")) {
    return "PDF";
  }
  if (name.endsWith(".txt") || name.endsWith(".md")) {
    return "Manual Note";
  }
  return "Other";
}

function stripFileExtension(filename) {
  return String(filename || "").replace(/\.[^.]+$/, "");
}

function formatReportSummaryError(error) {
  const message = String((error && error.message) || "Report file summarization failed.");
  try {
    const parsed = JSON.parse(message.replace(/^AI Analyst failed:\s*/i, ""));
    return parsed.error && parsed.error.message
      ? `AI Analyst failed: ${parsed.error.message}`
      : message;
  } catch (_) {
    return message;
  }
}

function applyReportUpdateSummary(summary, file) {
  const clean = summary && typeof summary === "object" ? summary : {};
  if (reportUpdateAiSummaryField) {
    reportUpdateAiSummaryField.value = String(clean.aiSummary || "").trim();
  }
  if (reportUpdateKeyMetricsField) {
    reportUpdateKeyMetricsField.value = String(clean.keyMetrics || "").trim();
  }
  if (reportUpdateKeyWinsField) {
    reportUpdateKeyWinsField.value = String(clean.keyWins || "").trim();
  }
  if (reportUpdateKeyRisksField) {
    reportUpdateKeyRisksField.value = String(clean.keyRisks || "").trim();
  }
  if (reportUpdateActionItemsField) {
    reportUpdateActionItemsField.value = String(clean.actionItems || "").trim();
  }
  if (reportUpdateOriginalNotesField && clean.originalNotes) {
    reportUpdateOriginalNotesField.value = String(clean.originalNotes || "").trim();
  }
  if (reportUpdateTitleField && !String(reportUpdateTitleField.value || "").trim()) {
    reportUpdateTitleField.value = String(clean.title || stripFileExtension(file && file.name)).trim();
  }
  if (reportUpdateAttachmentField && file) {
    reportUpdateAttachmentField.value = file.name;
  }
  if (reportUpdateSourceTypeField && file) {
    reportUpdateSourceTypeField.value = getReportUpdateFileSourceType(file);
  }
}

async function summarizeReportUpdateFile(file) {
  if (!file) {
    return;
  }

  const lowerName = String(file.name || "").toLowerCase();
  if (!lowerName.match(/\.(pdf|txt|md)$/)) {
    if (reportUpdateMessage) {
      reportUpdateMessage.textContent = "Upload a PDF or text file for report summarization.";
    }
    return;
  }

  updateReportUpdateFileLabel(file);
  if (reportUpdateMessage) {
    reportUpdateMessage.textContent = `Summarizing ${file.name} into update fields...`;
  }

  if (summarizeReportUpdateDraftButton) {
    summarizeReportUpdateDraftButton.disabled = true;
  }

  try {
    const fileData = await readFileAsBase64(file);
    const companyRecord = findCompanyRecord(selectedCompany, selectedCompanyEntity);
    const latest = companyRecord && companyRecord.latest ? companyRecord.latest : {};
    const result = await fetchJson("/api/summarize-report-file", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        filename: file.name,
        fileData,
        company: latest.company || selectedCompany || "",
        entity: normalizeEntityName(latest.entity || selectedCompanyEntity || ""),
        reportPeriod: String(reportUpdatePeriodField.value || "").trim(),
        updateType: String(reportUpdateTypeField.value || "").trim()
      })
    });

    applyReportUpdateSummary(result.summary, file);
    if (reportUpdateMessage) {
      reportUpdateMessage.textContent = "File summarized into the update fields.";
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
      reportUpdateMessage.textContent = formatReportSummaryError(error);
    }
  } finally {
    if (summarizeReportUpdateDraftButton) {
      summarizeReportUpdateDraftButton.disabled = false;
    }
  }
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
                  ${buildReportUpdateField("Contact emailed", row.contactEmailed || "Not Available")}
                  ${buildReportUpdateField("Subject line", row.subjectLine || "Not Available")}
                  ${buildReportUpdateField("Response status", row.responseStatus || "Not Available")}
                  ${buildReportUpdateField(
                    "Requested materials",
                    Array.isArray(row.materialsRequested) && row.materialsRequested.length
                      ? row.materialsRequested.join(", ")
                      : "Not Available"
                  )}
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
  if (generateInvestmentHistorySummaryButton) {
    generateInvestmentHistorySummaryButton.disabled = reportRows.length === 0;
  }
  if (generateAllReportSummaryButton) {
    generateAllReportSummaryButton.disabled = reportRows.length === 0;
  }
}

function buildLegacyCapitalActivityRows(investment) {
  const rows = [];
  const legacyInvestmentAmount = investment.capitalCallAmount || investment.amount || "";
  const legacyInvestmentDate = investment.capitalCallDate || investment.createdAt || "";
  const pipelineInvestment =
    isPipelineStatus(investment.status) || isPipelineStatus(investment.stage);

  if (!pipelineInvestment && (legacyInvestmentDate || legacyInvestmentAmount)) {
    rows.push({
      date: legacyInvestmentDate,
      type: "Investment Amount",
      amount: legacyInvestmentAmount,
      notes: "",
      legacyFallback: true
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

function isContributionCapitalType(type) {
  const normalizedType = String(type || "").toLowerCase();
  return (
    normalizedType.includes("capital call") ||
    normalizedType.includes("investment amount") ||
    normalizedType.includes("fee")
  );
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
    field.addEventListener("blur", () => {
      formatMoneyField(field);
    });
    field.addEventListener("change", () => {
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
  const masterEditor = isMasterEditor();
  const dashboardCopy = dashboardSection.querySelector(".section-copy");
  const entityPerformanceCopy = entityPerformanceSection.querySelector(".section-copy");
  form.classList.toggle("hidden", !editable);
  taskForm.classList.toggle("hidden", !editable);
  sendDigestButton.classList.toggle("hidden", !editable);
  downloadBackupButton.classList.toggle("hidden", !masterEditor);
  if (importWorkbookLabel) {
    importWorkbookLabel.classList.toggle("hidden", !masterEditor);
  }
  if (restoreBackupLabel) {
    restoreBackupLabel.classList.toggle("hidden", !masterEditor);
  }
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
  roleNotice.textContent = masterEditor
    ? "Role: Master Editor. Full portfolio administration is enabled."
    : editable
      ? "Role: Editor. Lee Beaman IRA plus Lee Beaman public stock and fixed income holdings are restricted."
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
  assignIfBlank("assetType", latest.assetType || "Private Investment");
  assignIfBlank("ticker", latest.ticker || "");
  assignIfBlank("exchange", latest.exchange || "");
  assignIfBlank("shareCount", latest.shareCount || "");
  assignIfBlank("costBasisPerShare", latest.costBasisPerShare || "");
  assignIfBlank("marketPrice", latest.marketPrice || "");
  assignIfBlank("marketPriceDate", latest.marketPriceDate || "");
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
  assignIfBlank("reportingCadence", latest.reportingCadence || "");
  assignIfBlank("updateRequestStatus", latest.updateRequestStatus || "");
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
  updateStockDetailsVisibility();
  return true;
}

function filterInvestments(investments) {
  const filters = currentFilters();
  return investments.filter((investment) => {
    const searchHaystack = [
      investment.entity,
      investment.company,
      investment.assetType,
      investment.ticker,
      investment.exchange,
      investment.notes,
      investment.deckSummary,
      investment.owner,
      investment.nextStep,
      investment.capitalCallAmount,
      investment.distributionAmount,
      investment.officialValue,
      investment.internalValue,
      investment.exitValue,
      investment.shareCount,
      investment.costBasisPerShare,
      investment.marketPrice,
      investment.marketPriceDate,
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

function isStockInvestment(investment) {
  const assetType = String((investment && investment.assetType) || "").trim().toLowerCase();
  return assetType.includes("stock") || Boolean(String((investment && investment.ticker) || "").trim());
}

function isPublicStockInvestment(investment) {
  return String((investment && investment.assetType) || "").trim() === "Public Stock";
}

function isCashInvestment(investment) {
  return String((investment && investment.assetType) || "").trim() === "Cash";
}

function isBondInvestment(investment) {
  return String((investment && investment.assetType) || "").trim() === "Bond / Fixed Income";
}

function isRealEstateInvestment(investment) {
  return String((investment && investment.assetType) || "").trim() === "Real Estate";
}

function isStockAssetType(value) {
  return String(value || "").toLowerCase().includes("stock");
}

function isPublicStockAssetType(value) {
  return String(value || "").trim() === "Public Stock";
}

function isCashAssetType(value) {
  return String(value || "").trim() === "Cash";
}

function isBondAssetType(value) {
  return String(value || "").trim() === "Bond / Fixed Income";
}

function isRealEstateAssetType(value) {
  return String(value || "").trim() === "Real Estate";
}

function stockKey(investment) {
  const ticker = String((investment && investment.ticker) || "").trim().toUpperCase();
  return ticker || companyKey(investment && investment.company);
}

function getPublicStockRows(investments) {
  const savedStocks = sortInvestmentsAlphabetically(investments.filter(isStockInvestment));
  const savedKeys = new Set(savedStocks.map(stockKey).filter(Boolean));
  const watchlistRows = PUBLIC_STOCK_PRESETS.filter(
    (preset) => !savedKeys.has(stockKey(preset))
  ).map((preset) => ({
    ...preset,
    id: "",
    shareCount: "",
    costBasisPerShare: "",
    marketPrice: "",
    marketPriceDate: "",
    marketValue: "",
    amount: "",
    isWatchlistOnly: true
  }));

  return sortInvestmentsAlphabetically(savedStocks.concat(watchlistRows));
}

function getStockMarketValue(investment) {
  const shares = toNumber(investment && investment.shareCount);
  const price = toNumber(investment && investment.marketPrice);
  const manualValue = toNumber(investment && investment.marketValue);
  return shares && price ? shares * price : manualValue;
}

function getStockCostBasis(investment) {
  const shares = toNumber(investment && investment.shareCount);
  const costPerShare = toNumber(investment && investment.costBasisPerShare);
  return shares && costPerShare ? shares * costPerShare : toNumber(investment && investment.amount);
}

function getStockGainLoss(investment) {
  return getStockMetrics(investment).gainLoss;
}

function getStockMetrics(investment) {
  const shares = toNumber(investment && investment.shareCount);
  const costPerShare = toNumber(investment && investment.costBasisPerShare);
  const currentPrice = toNumber(investment && investment.marketPrice);
  const totalCostBasis = getStockCostBasis(investment);
  const marketValue = getStockMarketValue(investment);
  const gainLoss = marketValue - totalCostBasis;
  const gainLossPercent = totalCostBasis > 0 ? gainLoss / totalCostBasis : null;

  return {
    shares,
    costPerShare,
    currentPrice,
    totalCostBasis,
    marketValue,
    gainLoss,
    gainLossPercent,
    hasCostBasis: totalCostBasis > 0,
    hasMarketValue: marketValue > 0
  };
}

function getPublicStockMarketValue(investment) {
  const metrics = getStockMetrics(investment);
  return metrics.shares && metrics.currentPrice ? metrics.shares * metrics.currentPrice : 0;
}

function getPublicStockSyncedValuation(investment) {
  if (!isPublicStockInvestment(investment)) {
    return null;
  }

  const marketValue = getPublicStockMarketValue(investment);
  if (!marketValue) {
    return null;
  }

  const normalizedMarketValue = normalizeMoneyString(String(marketValue), 6);
  return {
    marketValue: normalizedMarketValue,
    valuationDate: String(investment.marketPriceDate || "").trim(),
    officialValue: normalizedMarketValue,
    internalValue: normalizedMarketValue
  };
}

function applyPublicStockValuationSync(payload) {
  if (!isPublicStockInvestment(payload)) {
    return payload;
  }

  const syncedValuation = getPublicStockSyncedValuation(payload);
  if (!syncedValuation) {
    return payload;
  }

  return {
    ...payload,
    marketValue: syncedValuation.marketValue,
    valuationDate: syncedValuation.valuationDate,
    officialValue: syncedValuation.officialValue,
    internalValue: syncedValuation.internalValue
  };
}

function applyCashValuationSync(payload) {
  if (!isCashInvestment(payload)) {
    return payload;
  }

  const balance = toNumber(payload.amount);
  const normalizedBalance = normalizeMoneyString(String(payload.amount || ""), 6);
  return {
    ...payload,
    ticker: "",
    exchange: "",
    shareCount: "",
    costBasisPerShare: "",
    marketPrice: "",
    marketPriceDate: "",
    marketValue: balance ? normalizedBalance : "",
    status: "Active",
    nextStep: "",
    nextStepDueDate: "",
    capitalActivity: [],
    capitalCallDate: "",
    capitalCallAmount: "",
    distributionDate: "",
    distributionAmount: "",
    valuationDate: payload.valuationDate || "",
    officialValue: balance ? normalizedBalance : "",
    internalValue: balance ? normalizedBalance : "",
    exitValue: "",
    ownershipPercent: "",
    entityOwnershipPercent: "",
    ownershipNotes: "",
    followOnCapitalAmount: "",
    followOnCapitalStatus: "",
    followOnCapitalNotes: "",
    deckSummary: "",
    decisionDate: "",
    decisionType: "",
    decisionSummary: ""
  };
}

function getBondMetrics(investment) {
  const parValue = toNumber(investment && investment.bondParValue);
  const purchasePrice = toNumber(investment && investment.bondPurchasePrice);
  const enteredCostBasis = toNumber(investment && investment.bondCostBasis);
  const currentPrice = toNumber(investment && investment.bondCurrentPrice);
  const couponRate = toNumber(investment && investment.bondCouponRate) / 100;
  const ytm = toNumber(investment && investment.bondYieldToMaturity) / 100;
  const ytc = toNumber(investment && investment.bondYieldToCall) / 100;
  const marketValue = parValue > 0 && currentPrice > 0 ? (parValue * currentPrice) / 100 : 0;
  const annualCouponIncome = parValue > 0 && couponRate > 0 ? parValue * couponRate : 0;
  const calculatedCostBasis = parValue > 0 && purchasePrice > 0 ? (parValue * purchasePrice) / 100 : 0;
  const costBasis = enteredCostBasis > 0 ? enteredCostBasis : calculatedCostBasis;
  const currentYield = marketValue > 0 && annualCouponIncome > 0 ? annualCouponIncome / marketValue : null;

  return {
    parValue,
    purchasePrice,
    currentPrice,
    marketValue,
    annualCouponIncome,
    couponRate: Number.isFinite(couponRate) && couponRate > 0 ? couponRate : null,
    costBasis,
    currentYield,
    ytm: Number.isFinite(ytm) && ytm !== 0 ? ytm : null,
    ytc: Number.isFinite(ytc) && ytc !== 0 ? ytc : null
  };
}

function applyBondValuationSync(payload) {
  if (!isBondInvestment(payload)) {
    return payload;
  }

  const metrics = getBondMetrics(payload);
  const normalizedMarketValue = metrics.marketValue > 0 ? normalizeMoneyString(String(metrics.marketValue), 6) : "";
  const normalizedCostBasis = metrics.costBasis > 0 ? normalizeMoneyString(String(metrics.costBasis), 6) : "";
  const currentYieldPercent =
    metrics.currentYield !== null ? String(Number((metrics.currentYield * 100).toFixed(6))) : "";

  return {
    ...payload,
    ticker: "",
    exchange: "",
    shareCount: "",
    costBasisPerShare: "",
    marketPrice: "",
    marketPriceDate: "",
    marketValue: normalizedMarketValue,
    amount: normalizedCostBasis || payload.amount || "",
    status: payload.status || "Active",
    capitalCallDate: "",
    capitalCallAmount: "",
    distributionDate: "",
    distributionAmount: "",
    valuationDate:
      payload.bondCurrentPrice && payload.bondMarketPriceDate
        ? payload.bondMarketPriceDate
        : payload.valuationDate || "",
    officialValue: normalizedMarketValue || payload.officialValue || "",
    internalValue: normalizedMarketValue || payload.internalValue || "",
    ownershipPercent: "",
    entityOwnershipPercent: "",
    ownershipNotes: "",
    followOnCapitalAmount: "",
    followOnCapitalStatus: "",
    followOnCapitalNotes: "",
    bondMarketValue: normalizedMarketValue,
    bondCostBasis: normalizedCostBasis || payload.bondCostBasis || "",
    bondCurrentYield: currentYieldPercent
  };
}

function applyAssetValuationSync(payload) {
  return applyRealEstateValuationSync(
    applyBondValuationSync(applyCashValuationSync(applyPublicStockValuationSync(payload)))
  );
}

function getBondRows(investments) {
  return sortInvestmentsAlphabetically(investments.filter(isBondInvestment));
}

function getBondIssuer(investment) {
  return String((investment && (investment.bondIssuer || investment.company)) || "").trim();
}

function getBondDescription(investment) {
  return String((investment && (investment.bondDescription || investment.notes)) || "").trim();
}

function getBondType(investment) {
  return String((investment && (investment.bondType || investment.stage)) || "Other").trim() || "Other";
}

function getBondMaturityYear(investment) {
  const maturity = String((investment && investment.bondMaturityDate) || "").trim();
  return /^\d{4}/.test(maturity) ? maturity.slice(0, 4) : "";
}

function buildFixedIncomeSummary(rows) {
  const positions = getBondRows(rows);
  const totals = positions.reduce(
    (summary, investment) => {
      const metrics = getBondMetrics(investment);
      summary.marketValue += metrics.marketValue;
      summary.parValue += metrics.parValue;
      summary.annualCouponIncome += metrics.annualCouponIncome;
      if (metrics.couponRate !== null && metrics.parValue > 0) {
        summary.weightedCouponNumerator += metrics.couponRate * metrics.parValue;
        summary.weightedCouponDenominator += metrics.parValue;
      }
      if (metrics.ytm !== null && metrics.marketValue > 0) {
        summary.weightedYtmNumerator += metrics.ytm * metrics.marketValue;
        summary.weightedYtmDenominator += metrics.marketValue;
      }
      return summary;
    },
    {
      positions,
      marketValue: 0,
      parValue: 0,
      annualCouponIncome: 0,
      weightedCouponNumerator: 0,
      weightedCouponDenominator: 0,
      weightedYtmNumerator: 0,
      weightedYtmDenominator: 0
    }
  );

  const nextMaturity = positions
    .map((investment) => investment.bondMaturityDate)
    .filter(Boolean)
    .sort((left, right) => parseDateValue(left, new Date(8640000000000000)) - parseDateValue(right, new Date(8640000000000000)))[0] || "";

  return {
    ...totals,
    weightedCoupon:
      totals.weightedCouponDenominator > 0
        ? totals.weightedCouponNumerator / totals.weightedCouponDenominator
        : null,
    weightedYtm:
      totals.weightedYtmDenominator > 0
        ? totals.weightedYtmNumerator / totals.weightedYtmDenominator
        : null,
    nextMaturity
  };
}

function buildFixedIncomeHoldingsSummary(entity = "") {
  return buildFixedIncomeSummary(
    getBondRows(allInvestments).filter(
      (investment) => !entity || normalizeEntityName(investment.entity) === normalizeEntityName(entity)
    )
  );
}

function buildMaturityLadder(rows) {
  const ladder = new Map();
  getBondRows(rows).forEach((investment) => {
    const year = getBondMaturityYear(investment);
    if (!year) {
      return;
    }
    ladder.set(year, (ladder.get(year) || 0) + getBondMetrics(investment).parValue);
  });
  return Array.from(ladder.entries()).sort(([left], [right]) => left.localeCompare(right));
}

function getRealEstateRows(investments) {
  return sortInvestmentsAlphabetically(investments.filter(isRealEstateInvestment));
}

function getRealEstatePropertyName(investment) {
  return String((investment && (investment.realEstatePropertyName || investment.company)) || "").trim();
}

function getRealEstatePropertyType(investment) {
  return String((investment && (investment.realEstatePropertyType || investment.stage)) || "Other").trim() || "Other";
}

function getRealEstateAddress(investment) {
  const parts = [
    investment && investment.realEstateAddress,
    investment && investment.realEstateCity,
    investment && investment.realEstateState,
    investment && investment.realEstateZip
  ]
    .map((part) => String(part || "").trim())
    .filter(Boolean);
  return parts.length ? parts.join(", ") : "";
}

function getRealEstateOwnershipRate(investment) {
  const entered = toNumber(investment && investment.realEstateOwnershipPercent);
  if (!entered) {
    return 1;
  }
  return Math.max(0, Math.min(entered, 100)) / 100;
}

function getRealEstateMetrics(investment) {
  const appraisedValue = toNumber(investment && investment.realEstateAppraisedValue);
  const ownershipRate = getRealEstateOwnershipRate(investment);
  const debt = toNumber(investment && investment.realEstateDebt);
  const noi = toNumber(investment && investment.realEstateNoi);
  const revenue = toNumber(investment && investment.realEstateRevenue);
  const purchasePrice = toNumber(investment && investment.realEstatePurchasePrice);
  const costBasis = toNumber(investment && investment.realEstateCostBasis);
  const internalValueOverride = toNumber(investment && investment.realEstateInternalValueOverride);
  const entityValue = appraisedValue * ownershipRate;
  const entityDebt = debt * ownershipRate;
  const netEquity = entityValue - entityDebt;
  const capRate = appraisedValue > 0 && noi > 0 ? noi / appraisedValue : null;
  const ltv = appraisedValue > 0 && debt > 0 ? debt / appraisedValue : null;
  const noiMargin = revenue > 0 && noi > 0 ? noi / revenue : null;
  const hasInternalValueOverride = String((investment && investment.realEstateInternalValueOverride) || "").trim() !== "";

  return {
    appraisedValue,
    ownershipRate,
    ownershipPercent: ownershipRate * 100,
    debt,
    entityValue,
    entityDebt,
    netEquity,
    noi,
    revenue,
    purchasePrice,
    costBasis,
    internalValueOverride,
    hasInternalValueOverride,
    capRate,
    ltv,
    noiMargin
  };
}

function applyRealEstateValuationSync(payload) {
  if (!isRealEstateInvestment(payload)) {
    return payload;
  }

  const metrics = getRealEstateMetrics(payload);
  const hasValuationInputs = metrics.appraisedValue > 0 || metrics.debt > 0;
  const normalizedNetEquity = hasValuationInputs ? normalizeMoneyString(String(metrics.netEquity), 6) : "";
  const normalizedInternalOverride = metrics.hasInternalValueOverride
    ? normalizeMoneyString(String(metrics.internalValueOverride), 6)
    : "";
  const capRatePercent =
    metrics.capRate !== null ? String(Number((metrics.capRate * 100).toFixed(6))) : "";
  const noiMarginPercent =
    metrics.noiMargin !== null ? String(Number((metrics.noiMargin * 100).toFixed(6))) : "";

  return {
    ...payload,
    ticker: "",
    exchange: "",
    shareCount: "",
    costBasisPerShare: "",
    marketPrice: "",
    marketPriceDate: "",
    marketValue: normalizedNetEquity,
    amount: payload.realEstateCostBasis || payload.realEstatePurchasePrice || payload.amount || "",
    stage: payload.realEstatePropertyType || payload.stage || "",
    owner: payload.realEstateEntityOwner || payload.owner || "",
    status: payload.status || "Active",
    capitalCallDate: "",
    capitalCallAmount: "",
    distributionDate: "",
    distributionAmount: "",
    capitalActivity: [],
    valuationDate: payload.realEstateAppraisalDate || payload.valuationDate || "",
    officialValue: normalizedNetEquity || payload.officialValue || "",
    internalValue: normalizedInternalOverride || normalizedNetEquity || payload.internalValue || "",
    ownershipPercent: payload.realEstateOwnershipPercent || payload.ownershipPercent || "",
    realEstateCapRate: capRatePercent,
    realEstateNoiMargin: noiMarginPercent
  };
}

function buildRealEstateSummary(rows) {
  const positions = getRealEstateRows(rows);
  const totals = positions.reduce(
    (summary, investment) => {
      const metrics = getRealEstateMetrics(investment);
      summary.appraisedValue += metrics.appraisedValue;
      summary.netEquity += metrics.netEquity;
      summary.debt += metrics.debt;
      summary.noi += metrics.noi;
      if (metrics.capRate !== null && metrics.appraisedValue > 0) {
        summary.weightedCapRateNumerator += metrics.capRate * metrics.appraisedValue;
        summary.weightedCapRateDenominator += metrics.appraisedValue;
      }
      if (metrics.ltv !== null && metrics.appraisedValue > 0) {
        summary.weightedLtvNumerator += metrics.ltv * metrics.appraisedValue;
        summary.weightedLtvDenominator += metrics.appraisedValue;
      }
      return summary;
    },
    {
      positions,
      appraisedValue: 0,
      netEquity: 0,
      debt: 0,
      noi: 0,
      weightedCapRateNumerator: 0,
      weightedCapRateDenominator: 0,
      weightedLtvNumerator: 0,
      weightedLtvDenominator: 0
    }
  );
  const latestAppraisalDate = positions
    .map((investment) => investment.realEstateAppraisalDate || investment.valuationDate)
    .filter(Boolean)
    .sort((left, right) => parseDateValue(right, new Date(0)) - parseDateValue(left, new Date(0)))[0] || "";

  return {
    ...totals,
    weightedCapRate:
      totals.weightedCapRateDenominator > 0
        ? totals.weightedCapRateNumerator / totals.weightedCapRateDenominator
        : null,
    weightedLtv:
      totals.weightedLtvDenominator > 0
        ? totals.weightedLtvNumerator / totals.weightedLtvDenominator
        : null,
    latestAppraisalDate
  };
}

function buildRealEstateHoldingsSummary(entity = "") {
  return buildRealEstateSummary(
    getRealEstateRows(allInvestments).filter(
      (investment) => !entity || normalizeEntityName(investment.entity) === normalizeEntityName(entity)
    )
  );
}

function getCashRows(investments) {
  return sortInvestmentsAlphabetically(investments.filter(isCashInvestment));
}

function getCashBalance(investment) {
  return toNumber(investment && (investment.amount || investment.internalValue || investment.officialValue));
}

function getCashAccountType(investment) {
  return String((investment && investment.stage) || "Other").trim() || "Other";
}

function getCashInstitution(investment) {
  return String((investment && investment.owner) || "").trim();
}

function buildCashHoldingsSummary(entity = "") {
  const positions = getCashRows(allInvestments).filter(
    (investment) => !entity || normalizeEntityName(investment.entity) === normalizeEntityName(entity)
  );
  const totalBalance = positions.reduce((sum, investment) => sum + getCashBalance(investment), 0);
  const largestAccount = positions.reduce((largest, investment) => {
    const balance = getCashBalance(investment);
    return !largest || balance > largest.balance
      ? {
          investment,
          balance
        }
      : largest;
  }, null);
  const latestBalanceDate = positions
    .map((investment) => investment.valuationDate)
    .filter(Boolean)
    .sort((left, right) => parseDateValue(right, new Date(0)) - parseDateValue(left, new Date(0)))[0] || "";

  return {
    positions,
    totalBalance,
    largestAccount,
    latestBalanceDate
  };
}

function getCashByEntity(investments = allInvestments) {
  const grouped = new Map();
  getCashRows(investments).forEach((investment) => {
    const entity = normalizeEntityName(investment.entity) || "No entity";
    grouped.set(entity, (grouped.get(entity) || 0) + getCashBalance(investment));
  });
  return Array.from(grouped.entries()).sort(([left], [right]) => left.localeCompare(right));
}

function publicStockValuationNeedsPatch(investment, quote) {
  const nextPayload = getQuotePatchPayload(investment, quote);
  return (
    normalizeMoneyString(investment.marketPrice || "", 6) !==
      normalizeMoneyString(nextPayload.marketPrice || "", 6) ||
    String(investment.marketPriceDate || "") !== String(nextPayload.marketPriceDate || "") ||
    normalizeMoneyString(investment.marketValue || "", 6) !==
      normalizeMoneyString(nextPayload.marketValue || "", 6) ||
    String(investment.valuationDate || "") !== String(nextPayload.valuationDate || "") ||
    normalizeMoneyString(investment.officialValue || "", 6) !==
      normalizeMoneyString(nextPayload.officialValue || "", 6) ||
    normalizeMoneyString(investment.internalValue || "", 6) !==
      normalizeMoneyString(nextPayload.internalValue || "", 6)
  );
}

function getPublicStockQuoteRequestKey(investment) {
  return String(investment.ticker || "").trim().toUpperCase();
}

function waitForPublicStockQuoteSlot(delayMs = 650) {
  return new Promise((resolve) => {
    window.setTimeout(resolve, delayMs);
  });
}

function groupPublicStockPositionsByTicker(positions) {
  const groups = new Map();
  positions.forEach((investment) => {
    const ticker = String(investment.ticker || "").trim().toUpperCase();
    if (!ticker) {
      return;
    }
    if (!groups.has(ticker)) {
      groups.set(ticker, []);
    }
    groups.get(ticker).push(investment);
  });
  return groups;
}

function buildPublicHoldingsSummary(entity) {
  const positions = allInvestments.filter(
    (investment) =>
      isPublicStockInvestment(investment) &&
      normalizeEntityName(investment.entity) === normalizeEntityName(entity)
  );
  const marketValue = positions.reduce((sum, investment) => sum + getStockMetrics(investment).marketValue, 0);
  const costBasis = positions.reduce((sum, investment) => sum + getStockMetrics(investment).totalCostBasis, 0);
  const gainLoss = marketValue - costBasis;
  const gainLossPercent = costBasis > 0 ? gainLoss / costBasis : null;

  return {
    positions,
    marketValue,
    costBasis,
    gainLoss,
    gainLossPercent,
    ytdChangePercent: null
  };
}

function formatSignedMoney(value) {
  const amount = toNumber(value);
  if (!amount) {
    return "$0";
  }
  return amount > 0 ? `+${formatMoney(amount)}` : `-${formatMoney(Math.abs(amount))}`;
}

function formatStockPercent(value) {
  return Number.isFinite(value) ? `${value >= 0 ? "+" : ""}${formatPercent(value)}` : "N/A";
}

function getStockPerformanceClass(value) {
  if (!Number.isFinite(value) || value === 0) {
    return "stock-performance-flat";
  }
  return value > 0 ? "stock-performance-positive" : "stock-performance-negative";
}

function getFilteredPublicStockRows(stocks) {
  const search = String(publicStockFilters.search || "").trim().toLowerCase();
  return stocks.filter((investment) => {
    const matchesEntity =
      !publicStockFilters.entity ||
      normalizeEntityName(investment.entity) === normalizeEntityName(publicStockFilters.entity);
    const matchesSearch =
      !search ||
      String(investment.company || "").toLowerCase().includes(search) ||
      String(investment.ticker || "").toLowerCase().includes(search);
    return matchesEntity && matchesSearch;
  });
}

function isPublicStockRow(investment) {
  return Boolean((investment && investment.isWatchlistOnly) || isPublicStockInvestment(investment));
}

function renderPublicStockFilterOptions(stocks) {
  if (!publicStockEntityFilter) {
    return;
  }

  const selectedEntity = publicStockFilters.entity;
  const entities = Array.from(
    new Set(
      configuredEntities
        .concat(stocks.map((investment) => normalizeEntityName(investment.entity)).filter(Boolean))
        .map(normalizeEntityName)
    )
  ).sort((left, right) => left.localeCompare(right));
  publicStockEntityFilter.innerHTML = [
    '<option value="">All entities</option>',
    ...entities.map((entity) => `<option value="${escapeHtml(entity)}">${escapeHtml(entity)}</option>`)
  ].join("");
  publicStockEntityFilter.value = entities.includes(selectedEntity) ? selectedEntity : "";
  publicStockFilters.entity = publicStockEntityFilter.value;
}

function getQuotePatchPayload(investment, quote) {
  return applyPublicStockValuationSync({
    company: investment.company || "",
    entity: investment.entity || "",
    assetType: investment.assetType || "",
    ticker: quote.symbol || investment.ticker || "",
    exchange: quote.exchangeName || investment.exchange || "",
    shareCount: investment.shareCount || "",
    costBasisPerShare: investment.costBasisPerShare || "",
    marketPrice: quote.price || investment.marketPrice || "",
    marketPriceDate: quote.priceDate || investment.marketPriceDate || "",
    marketValue: investment.marketValue || "",
    amount: investment.amount || "",
    currency: quote.currency || investment.currency || "USD",
    stage: investment.stage || "",
    status: investment.status || "",
    owner: investment.owner || "",
    nextStep: investment.nextStep || "",
    nextStepDueDate: investment.nextStepDueDate || "",
    notes: investment.notes || "",
    deckSummary: investment.deckSummary || "",
    capitalActivity: Array.isArray(investment.capitalActivity) ? investment.capitalActivity : [],
    capitalCallDate: investment.capitalCallDate || "",
    capitalCallAmount: investment.capitalCallAmount || "",
    distributionDate: investment.distributionDate || "",
    distributionAmount: investment.distributionAmount || "",
    valuationDate: investment.valuationDate || "",
    officialValue: investment.officialValue || "",
    internalValue: investment.internalValue || "",
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
    reportingCadence: investment.reportingCadence || "",
    updateRequestStatus: investment.updateRequestStatus || "",
    lastUpdateRequestSentAt: investment.lastUpdateRequestSentAt || "",
    lastUpdateRequestSubject: investment.lastUpdateRequestSubject || "",
    lastUpdateRequestContact: investment.lastUpdateRequestContact || "",
    documentLinks: investment.documentLinks || "",
    documents: Array.isArray(investment.documents) ? investment.documents : [],
    decisionDate: investment.decisionDate || "",
    decisionType: investment.decisionType || "",
    decisionSummary: investment.decisionSummary || "",
    reportUpdates: Array.isArray(investment.reportUpdates) ? investment.reportUpdates : [],
    recipients: Array.isArray(investment.recipients) ? investment.recipients : []
  });
}

function getStockTickerLabel(investment) {
  const ticker = String((investment && investment.ticker) || "").trim().toUpperCase();
  const exchange = String((investment && investment.exchange) || "").trim().toUpperCase();
  if (!ticker) {
    return "No public ticker";
  }
  return exchange ? `${exchange}: ${ticker}` : ticker;
}

function clearStockFields() {
  ["ticker", "exchange", "shareCount", "costBasisPerShare", "marketPrice", "marketPriceDate"].forEach(
    (fieldName) => {
      if (form && form.elements && form.elements[fieldName]) {
        form.elements[fieldName].value = "";
      }
    }
  );
}

function clearBondFields() {
  [
    "bondIssuer",
    "bondDescription",
    "bondType",
    "bondCusip",
    "bondEntityOwner",
    "bondParValue",
    "bondPurchasePrice",
    "bondPurchaseDate",
    "bondCostBasis",
    "bondCouponRate",
    "bondCouponFrequency",
    "bondMaturityDate",
    "bondCallDate",
    "bondCallPrice",
    "bondCurrentPrice",
    "bondMarketPriceDate",
    "bondMarketValue",
    "bondYieldToMaturity",
    "bondYieldToCall",
    "bondCurrentYield",
    "bondCreditRating",
    "bondInsurer",
    "bondTaxStatus",
    "bondAccruedInterest",
    "realEstatePropertyName",
    "realEstateAddress",
    "realEstateCity",
    "realEstateState",
    "realEstateZip",
    "realEstateEntityOwner",
    "realEstatePropertyType",
    "realEstateOwnershipPercent",
    "realEstateOwnershipNotes",
    "realEstateAcquisitionDate",
    "realEstatePurchasePrice",
    "realEstateCostBasis",
    "realEstateAppraisedValue",
    "realEstateAppraisalDate",
    "realEstateAppraiser",
    "realEstateAppraisalDocument",
    "realEstateInternalValueOverride",
    "realEstateInternalValueDate",
    "realEstateDebt",
    "realEstateLoanLender",
    "realEstateDebtInterestRate",
    "realEstateDebtMaturityDate",
    "realEstateDebtService",
    "realEstateLoanNotes",
    "realEstateNoi",
    "realEstateRevenue",
    "realEstateCapRate",
    "realEstateNoiMargin",
    "realEstateOccupancy",
    "realEstateSquareFootage",
    "realEstateAcreage",
    "realEstateUnits",
    "realEstatePropertyTaxes",
    "realEstateInsurance",
    "realEstateOtherExpenses",
    "realEstateOperatingNotes"
  ].forEach((fieldName) => {
    if (form && form.elements && form.elements[fieldName]) {
      form.elements[fieldName].value = "";
    }
  });
}

function clearRealEstateFields() {
  [
    "realEstatePropertyName",
    "realEstateAddress",
    "realEstateCity",
    "realEstateState",
    "realEstateZip",
    "realEstateEntityOwner",
    "realEstatePropertyType",
    "realEstateOwnershipPercent",
    "realEstateOwnershipNotes",
    "realEstateAcquisitionDate",
    "realEstatePurchasePrice",
    "realEstateCostBasis",
    "realEstateAppraisedValue",
    "realEstateAppraisalDate",
    "realEstateAppraiser",
    "realEstateAppraisalDocument",
    "realEstateInternalValueOverride",
    "realEstateInternalValueDate",
    "realEstateDebt",
    "realEstateLoanLender",
    "realEstateDebtInterestRate",
    "realEstateDebtMaturityDate",
    "realEstateDebtService",
    "realEstateLoanNotes",
    "realEstateNoi",
    "realEstateRevenue",
    "realEstateCapRate",
    "realEstateNoiMargin",
    "realEstateOccupancy",
    "realEstateSquareFootage",
    "realEstateAcreage",
    "realEstateUnits",
    "realEstatePropertyTaxes",
    "realEstateInsurance",
    "realEstateOtherExpenses",
    "realEstateOperatingNotes"
  ].forEach((fieldName) => {
    if (form && form.elements && form.elements[fieldName]) {
      form.elements[fieldName].value = "";
    }
  });
}

function updateStockDetailsVisibility(options = {}) {
  if (!stockDetailsPanel || !form || !form.elements || !form.elements.assetType) {
    return;
  }

  const showStockDetails = isStockAssetType(form.elements.assetType.value);
  const publicStockDetails = isPublicStockAssetType(form.elements.assetType.value);
  const cashDetails = isCashAssetType(form.elements.assetType.value);
  const bondDetails = isBondAssetType(form.elements.assetType.value);
  const realEstateDetails = isRealEstateAssetType(form.elements.assetType.value);
  stockDetailsPanel.classList.toggle("hidden", !showStockDetails);
  if (cashDetailsPanel) {
    cashDetailsPanel.classList.toggle("hidden", !cashDetails);
  }
  if (bondDetailsPanel) {
    bondDetailsPanel.classList.toggle("hidden", !bondDetails);
  }
  if (realEstateDetailsPanel) {
    realEstateDetailsPanel.classList.toggle("hidden", !realEstateDetails);
  }
  if (!showStockDetails && options.clearHiddenFields) {
    clearStockFields();
  }
  if (!bondDetails && options.clearHiddenFields) {
    clearBondFields();
  }
  if (!realEstateDetails && options.clearHiddenFields) {
    clearRealEstateFields();
  }
  if (cashDetails) {
    syncCashPanelFromForm();
    syncCashFormFields();
  }
  if (bondDetails) {
    syncBondPanelFromForm();
    syncBondFormFields();
  }
  if (realEstateDetails) {
    syncRealEstatePanelFromForm();
    syncRealEstateFormFields();
  }
  ["valuationDate", "officialValue", "internalValue"].forEach((fieldName) => {
    if (form.elements[fieldName]) {
      form.elements[fieldName].readOnly = publicStockDetails || cashDetails || bondDetails || realEstateDetails;
    }
  });
  if (form.elements.exitValue) {
    form.elements.exitValue.readOnly = cashDetails;
  }
  if (valuationHelperText) {
    valuationHelperText.textContent = cashDetails
      ? "For Cash, balance date drives valuation date and current balance drives official and internal value. Cash does not create IRR cash flows."
      : bondDetails
        ? "For Fixed Income, market price date drives valuation date and market value drives official and internal value. Exit value stays manual."
        : realEstateDetails
          ? "For Real Estate, appraisal date drives valuation date and net equity drives official and internal value. Exit value stays manual."
          : publicStockDetails
            ? "For Public Stocks, valuation date, official value, and internal value are driven by share count and current market price. Exit value stays manual."
            : "Use the latest valuation date when you update official, internal, or exit marks.";
  }
  syncPublicStockFormValuation();
  updateBondValuePreview();
  updateRealEstateValuePreview();
  updateStockValuePreview();
}

function syncRealEstatePanelFromForm() {
  if (!form || !form.elements || !isRealEstateAssetType(form.elements.assetType.value)) {
    return;
  }
  if (realEstatePropertyNameField && !realEstatePropertyNameField.value) {
    realEstatePropertyNameField.value = form.elements.company ? form.elements.company.value : "";
  }
  if (realEstatePropertyTypeField && !realEstatePropertyTypeField.value) {
    realEstatePropertyTypeField.value = form.elements.stage ? form.elements.stage.value : "";
  }
  if (realEstateEntityOwnerField && !realEstateEntityOwnerField.value) {
    realEstateEntityOwnerField.value = form.elements.owner ? form.elements.owner.value : "";
  }
}

function syncRealEstateFormFields() {
  if (!form || !form.elements || !isRealEstateAssetType(form.elements.assetType.value)) {
    return;
  }
  if (form.elements.company && realEstatePropertyNameField) {
    form.elements.company.value = realEstatePropertyNameField.value;
  }
  if (form.elements.stage && realEstatePropertyTypeField) {
    form.elements.stage.value = realEstatePropertyTypeField.value;
  }
  if (form.elements.owner && realEstateEntityOwnerField) {
    form.elements.owner.value = realEstateEntityOwnerField.value;
  }
  if (form.elements.status && !form.elements.status.value) {
    form.elements.status.value = "Active";
  }

  const synced = applyRealEstateValuationSync({
    assetType: "Real Estate",
    amount: form.elements.amount ? form.elements.amount.value : "",
    status: form.elements.status ? form.elements.status.value : "",
    valuationDate: form.elements.valuationDate ? form.elements.valuationDate.value : "",
    officialValue: form.elements.officialValue ? form.elements.officialValue.value : "",
    internalValue: form.elements.internalValue ? form.elements.internalValue.value : "",
    ownershipPercent: form.elements.ownershipPercent ? form.elements.ownershipPercent.value : "",
    realEstatePropertyType: realEstatePropertyTypeField ? realEstatePropertyTypeField.value : "",
    realEstateEntityOwner: realEstateEntityOwnerField ? realEstateEntityOwnerField.value : "",
    realEstateOwnershipPercent: realEstateOwnershipPercentField ? realEstateOwnershipPercentField.value : "",
    realEstatePurchasePrice: realEstatePurchasePriceField ? realEstatePurchasePriceField.value : "",
    realEstateCostBasis: realEstateCostBasisField ? realEstateCostBasisField.value : "",
    realEstateAppraisedValue: realEstateAppraisedValueField ? realEstateAppraisedValueField.value : "",
    realEstateAppraisalDate: realEstateAppraisalDateField ? realEstateAppraisalDateField.value : "",
    realEstateInternalValueOverride: realEstateInternalValueOverrideField
      ? realEstateInternalValueOverrideField.value
      : "",
    realEstateDebt: realEstateDebtField ? realEstateDebtField.value : "",
    realEstateNoi: realEstateNoiField ? realEstateNoiField.value : "",
    realEstateRevenue: realEstateRevenueField ? realEstateRevenueField.value : ""
  });

  if (realEstateCapRateField) {
    realEstateCapRateField.value = synced.realEstateCapRate || "";
  }
  if (form.elements.realEstateCapRate) {
    form.elements.realEstateCapRate.value = synced.realEstateCapRate || "";
  }
  if (form.elements.realEstateNoiMargin) {
    form.elements.realEstateNoiMargin.value = synced.realEstateNoiMargin || "";
  }
  if (form.elements.amount) {
    form.elements.amount.value = synced.amount || "";
  }
  if (form.elements.valuationDate) {
    form.elements.valuationDate.value = synced.valuationDate || "";
  }
  if (form.elements.officialValue) {
    form.elements.officialValue.value = synced.officialValue || "";
  }
  if (form.elements.internalValue) {
    form.elements.internalValue.value = synced.internalValue || "";
  }
  if (form.elements.ownershipPercent) {
    form.elements.ownershipPercent.value = synced.ownershipPercent || "";
  }
}

function updateRealEstateValuePreview() {
  if (!realEstateValuePreview || !form || !form.elements || !isRealEstateAssetType(form.elements.assetType.value)) {
    return;
  }
  const metrics = getRealEstateMetrics({
    realEstateOwnershipPercent: realEstateOwnershipPercentField ? realEstateOwnershipPercentField.value : "",
    realEstateAppraisedValue: realEstateAppraisedValueField ? realEstateAppraisedValueField.value : "",
    realEstateDebt: realEstateDebtField ? realEstateDebtField.value : "",
    realEstateNoi: realEstateNoiField ? realEstateNoiField.value : "",
    realEstateRevenue: realEstateRevenueField ? realEstateRevenueField.value : "",
    realEstateInternalValueOverride: realEstateInternalValueOverrideField
      ? realEstateInternalValueOverrideField.value
      : ""
  });
  const internalValueText = metrics.hasInternalValueOverride
    ? ` • internal override ${formatMoney(metrics.internalValueOverride)}`
    : "";
  realEstateValuePreview.textContent =
    metrics.appraisedValue > 0
      ? `Gross ${formatMoney(metrics.appraisedValue)} • entity gross ${formatMoney(metrics.entityValue)} • entity debt ${formatMoney(metrics.entityDebt)} • net equity ${formatMoney(metrics.netEquity)} • LTV ${formatPercent(metrics.ltv)} • cap rate ${formatPercent(metrics.capRate)} • NOI margin ${formatPercent(metrics.noiMargin)}${internalValueText}`
      : "Net equity will calculate from appraised value, ownership, and debt.";
}

function syncBondPanelFromForm() {
  if (!form || !form.elements || !isBondAssetType(form.elements.assetType.value)) {
    return;
  }
  if (bondIssuerField && !bondIssuerField.value) {
    bondIssuerField.value = form.elements.company ? form.elements.company.value : "";
  }
  if (bondTypeField && !bondTypeField.value) {
    bondTypeField.value = form.elements.stage ? form.elements.stage.value : "";
  }
  if (bondEntityOwnerField && !bondEntityOwnerField.value) {
    bondEntityOwnerField.value = form.elements.owner ? form.elements.owner.value : "";
  }
}

function syncBondFormFields() {
  if (!form || !form.elements || !isBondAssetType(form.elements.assetType.value)) {
    return;
  }
  if (form.elements.company && bondIssuerField) {
    form.elements.company.value = bondIssuerField.value;
  }
  if (form.elements.stage && bondTypeField) {
    form.elements.stage.value = bondTypeField.value;
  }
  if (form.elements.owner && bondEntityOwnerField) {
    form.elements.owner.value = bondEntityOwnerField.value;
  }
  if (form.elements.amount && bondCostBasisField) {
    form.elements.amount.value = bondCostBasisField.value;
  }
  if (form.elements.status && !form.elements.status.value) {
    form.elements.status.value = "Active";
  }

  const synced = applyBondValuationSync({
    assetType: "Bond / Fixed Income",
    amount: form.elements.amount ? form.elements.amount.value : "",
    status: form.elements.status ? form.elements.status.value : "",
    valuationDate: form.elements.valuationDate ? form.elements.valuationDate.value : "",
    officialValue: form.elements.officialValue ? form.elements.officialValue.value : "",
    internalValue: form.elements.internalValue ? form.elements.internalValue.value : "",
    bondParValue: bondParValueField ? bondParValueField.value : "",
    bondPurchasePrice: bondPurchasePriceField ? bondPurchasePriceField.value : "",
    bondCostBasis: bondCostBasisField ? bondCostBasisField.value : "",
    bondCouponRate: bondCouponRateField ? bondCouponRateField.value : "",
    bondCurrentPrice: bondCurrentPriceField ? bondCurrentPriceField.value : "",
    bondMarketPriceDate: bondMarketPriceDateField ? bondMarketPriceDateField.value : ""
  });

  if (bondCostBasisField && synced.bondCostBasis && !bondCostBasisField.value) {
    bondCostBasisField.value = synced.bondCostBasis;
  }
  if (bondMarketValueField) {
    bondMarketValueField.value = synced.bondMarketValue || "";
  }
  if (bondCurrentYieldField) {
    bondCurrentYieldField.value = synced.bondCurrentYield || "";
  }
  if (form.elements.amount) {
    form.elements.amount.value = synced.amount || "";
  }
  if (form.elements.valuationDate) {
    form.elements.valuationDate.value = synced.valuationDate || "";
  }
  if (form.elements.officialValue) {
    form.elements.officialValue.value = synced.officialValue || "";
  }
  if (form.elements.internalValue) {
    form.elements.internalValue.value = synced.internalValue || "";
  }
}

function updateBondValuePreview() {
  if (!bondValuePreview || !form || !form.elements || !isBondAssetType(form.elements.assetType.value)) {
    return;
  }
  const metrics = getBondMetrics({
    bondParValue: bondParValueField ? bondParValueField.value : "",
    bondPurchasePrice: bondPurchasePriceField ? bondPurchasePriceField.value : "",
    bondCostBasis: bondCostBasisField ? bondCostBasisField.value : "",
    bondCouponRate: bondCouponRateField ? bondCouponRateField.value : "",
    bondCurrentPrice: bondCurrentPriceField ? bondCurrentPriceField.value : ""
  });
  bondValuePreview.textContent =
    metrics.marketValue > 0
      ? `Market value ${formatMoney(metrics.marketValue)} • annual coupon ${formatMoney(metrics.annualCouponIncome)} • current yield ${formatPercent(metrics.currentYield)}`
      : "Market value and current yield will calculate from par value, price, and coupon.";
}

function syncCashPanelFromForm() {
  if (!form || !form.elements || !isCashAssetType(form.elements.assetType.value)) {
    return;
  }
  if (cashAccountNameField) {
    cashAccountNameField.value = form.elements.company ? form.elements.company.value : "";
  }
  if (cashInstitutionField) {
    cashInstitutionField.value = form.elements.owner ? form.elements.owner.value : "";
  }
  if (cashAccountTypeField) {
    cashAccountTypeField.value = form.elements.stage ? form.elements.stage.value : "";
  }
  if (cashBalanceField) {
    cashBalanceField.value = form.elements.amount ? form.elements.amount.value : "";
  }
  if (cashBalanceDateField) {
    cashBalanceDateField.value = form.elements.valuationDate ? form.elements.valuationDate.value : "";
  }
}

function syncCashFormFields() {
  if (!form || !form.elements || !isCashAssetType(form.elements.assetType.value)) {
    return;
  }
  if (form.elements.company && cashAccountNameField) {
    form.elements.company.value = cashAccountNameField.value;
  }
  if (form.elements.owner && cashInstitutionField) {
    form.elements.owner.value = cashInstitutionField.value;
  }
  if (form.elements.stage && cashAccountTypeField) {
    form.elements.stage.value = cashAccountTypeField.value;
  }
  if (form.elements.amount && cashBalanceField) {
    form.elements.amount.value = cashBalanceField.value;
  }
  if (form.elements.valuationDate && cashBalanceDateField) {
    form.elements.valuationDate.value = cashBalanceDateField.value;
  }
  if (form.elements.status) {
    form.elements.status.value = "Active";
  }
  const synced = applyCashValuationSync({
    assetType: "Cash",
    amount: form.elements.amount ? form.elements.amount.value : "",
    valuationDate: form.elements.valuationDate ? form.elements.valuationDate.value : "",
    status: "Active"
  });
  if (form.elements.officialValue) {
    form.elements.officialValue.value = synced.officialValue || "";
  }
  if (form.elements.internalValue) {
    form.elements.internalValue.value = synced.internalValue || "";
  }
  if (form.elements.exitValue) {
    form.elements.exitValue.value = "";
  }
}

function syncPublicStockFormValuation() {
  if (!form || !form.elements || !isPublicStockAssetType(form.elements.assetType.value)) {
    return;
  }

  const previewInvestment = {
    assetType: "Public Stock",
    shareCount: form.elements.shareCount ? form.elements.shareCount.value : "",
    marketPrice: form.elements.marketPrice ? form.elements.marketPrice.value : "",
    marketPriceDate: form.elements.marketPriceDate ? form.elements.marketPriceDate.value : ""
  };
  const syncedValuation = getPublicStockSyncedValuation(previewInvestment);
  if (!syncedValuation) {
    return;
  }

  if (form.elements.valuationDate) {
    form.elements.valuationDate.value = syncedValuation.valuationDate;
  }
  if (form.elements.officialValue) {
    form.elements.officialValue.value = syncedValuation.officialValue;
  }
  if (form.elements.internalValue) {
    form.elements.internalValue.value = syncedValuation.internalValue;
  }
}

function updateStockValuePreview() {
  if (!stockValuePreview || !form || !form.elements) {
    return;
  }
  syncPublicStockFormValuation();

  const previewInvestment = {
    amount: form.elements.amount ? form.elements.amount.value : "",
    shareCount: form.elements.shareCount ? form.elements.shareCount.value : "",
    costBasisPerShare: form.elements.costBasisPerShare ? form.elements.costBasisPerShare.value : "",
    marketPrice: form.elements.marketPrice ? form.elements.marketPrice.value : ""
  };
  const marketValue = getStockMarketValue(previewInvestment);
  const costBasis = getStockCostBasis(previewInvestment);
  const gainLoss = getStockGainLoss(previewInvestment);
  stockValuePreview.textContent = marketValue
    ? `Market value ${formatMoney(marketValue)}${costBasis ? ` • Cost basis ${formatMoney(costBasis)} • Gain/loss ${formatMoney(gainLoss)}` : ""}`
    : "Market value will calculate from shares and current price.";
}

async function fetchStockQuoteForForm() {
  if (!form || !form.elements || !form.elements.ticker) {
    return;
  }

  const ticker = String(form.elements.ticker.value || "").trim().toUpperCase();
  if (!ticker) {
    if (stockQuoteMessage) {
      stockQuoteMessage.textContent = "Enter a ticker first.";
    }
    return;
  }

  if (fetchStockQuoteButton) {
    fetchStockQuoteButton.disabled = true;
    fetchStockQuoteButton.textContent = "Updating...";
  }
  if (stockQuoteMessage) {
    stockQuoteMessage.textContent = `Looking up ${ticker}...`;
  }

  try {
    const quote = await fetchJson(`/api/stock-quote?ticker=${encodeURIComponent(ticker)}`);
    if (form.elements.marketPrice) {
      form.elements.marketPrice.value = normalizeMoneyString(String(quote.price || ""), 6);
    }
    if (form.elements.marketPriceDate) {
      form.elements.marketPriceDate.value = quote.priceDate || "";
    }
    if (form.elements.exchange && quote.exchangeName && !form.elements.exchange.value) {
      form.elements.exchange.value = quote.exchangeName;
    }
    updateStockValuePreview();
    if (stockQuoteMessage) {
      stockQuoteMessage.textContent = `Updated ${quote.symbol || ticker} at ${formatMoney(quote.price)}${quote.priceDate ? ` on ${quote.priceDate}` : ""}.`;
    }
  } catch (error) {
    if (error.status === 401) {
      setSignedInState(null);
      if (stockQuoteMessage) {
        stockQuoteMessage.textContent = "Your session expired. Please sign in again.";
      }
      return;
    }

    if (stockQuoteMessage) {
      stockQuoteMessage.textContent = error.message || "Stock price could not be loaded.";
    }
  } finally {
    if (fetchStockQuoteButton) {
      fetchStockQuoteButton.disabled = false;
      fetchStockQuoteButton.textContent = "Update price from ticker";
    }
  }
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

function getLatestCashFlowDate(cashFlows) {
  return (Array.isArray(cashFlows) ? cashFlows : []).reduce((latest, cashFlow) => {
    if (!(cashFlow.date instanceof Date) || !Number.isFinite(cashFlow.date.getTime())) {
      return latest;
    }

    return !latest || cashFlow.date > latest ? cashFlow.date : latest;
  }, null);
}

function isTerminalMarkAfterCashFlows(cashFlows, terminalMark) {
  if (!terminalMark || !terminalMark.date) {
    return false;
  }

  const latestCashFlowDate = getLatestCashFlowDate(cashFlows);
  return !latestCashFlowDate || terminalMark.date >= latestCashFlowDate;
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

function bisectXirrRoot(cashFlows, low, high) {
  let lowValue = xnpv(low, cashFlows);
  let highValue = xnpv(high, cashFlows);

  if (Math.abs(lowValue) < 1e-7) {
    return low;
  }
  if (Math.abs(highValue) < 1e-7) {
    return high;
  }
  if (lowValue * highValue > 0) {
    return null;
  }

  for (let iteration = 0; iteration < 80; iteration += 1) {
    const mid = (low + high) / 2;
    const midValue = xnpv(mid, cashFlows);
    if (Math.abs(midValue) < 1e-7) {
      return mid;
    }

    if (lowValue * midValue <= 0) {
      high = mid;
      highValue = midValue;
    } else {
      low = mid;
      lowValue = midValue;
    }
  }

  return (low + high) / 2;
}

function findXirrByBrackets(cashFlows) {
  const guesses = [
    -0.9999, -0.95, -0.75, -0.5, -0.25, -0.1, -0.05, 0, 0.05, 0.1, 0.15, 0.25,
    0.5, 0.75, 1, 1.5, 2, 3, 5, 10, 25, 50, 100, 250, 1000
  ];
  const roots = [];

  for (let index = 0; index < guesses.length - 1; index += 1) {
    const low = guesses[index];
    const high = guesses[index + 1];
    const lowValue = xnpv(low, cashFlows);
    const highValue = xnpv(high, cashFlows);
    if (!Number.isFinite(lowValue) || !Number.isFinite(highValue)) {
      continue;
    }

    if (Math.abs(lowValue) < 1e-7) {
      roots.push(low);
      continue;
    }

    if (lowValue * highValue <= 0) {
      const root = bisectXirrRoot(cashFlows, low, high);
      if (Number.isFinite(root)) {
        roots.push(root);
      }
    }
  }

  const uniqueRoots = roots.filter(
    (root, index) => roots.findIndex((candidate) => Math.abs(candidate - root) < 1e-6) === index
  );
  return uniqueRoots.sort((left, right) => Math.abs(left - 0.15) - Math.abs(right - 0.15))[0] ?? null;
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

  return findXirrByBrackets(sorted);
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
    const pipelineUpdate =
      isPipelineStatus(update.status) || isPipelineStatus(update.stage);
    const activityRows = isCashInvestment(update) || isBondInvestment(update) || isRealEstateInvestment(update)
      ? []
      : normalizeCapitalActivityRows(
          update.capitalActivity && update.capitalActivity.length
            ? update.capitalActivity
            : buildLegacyCapitalActivityRows(update)
        );

    return {
      update,
      pipelineUpdate,
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

  const normalizedActivities = updateActivities.flatMap(({ update, pipelineUpdate, activities }) => {
    if (!reconciliationOverride || reconciliationOverride.update.id === update.id) {
      return activities.filter(
        (activity) => !(pipelineUpdate && isContributionCapitalType(activity.type))
      );
    }

    return activities.filter((activity) => {
      if (pipelineUpdate && isContributionCapitalType(activity.type)) {
        return false;
      }

      return !isContributionCapitalType(activity.type);
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
    if (activity.legacyFallback && type.includes("investment amount")) {
      return false;
    }
    if (hasActualCalledCapital && type.includes("investment amount")) {
      return false;
    }

    return true;
  });

  const investedCapital = effectiveActivities.reduce((sum, activity) => {
    const amount = toNumber(activity.amount);
    return isContributionCapitalType(activity.type) ? sum + amount : sum;
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

    if (isContributionCapitalType(type)) {
      baseCashFlows.push({ date, amount: -amount });
      return;
    }

    baseCashFlows.push({ date, amount });
  });

  return {
    isCashPosition: updates.some(isCashInvestment),
    isBondPosition: updates.some(isBondInvestment),
    isRealEstatePosition: updates.some(isRealEstateInvestment),
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
  if (terminalValue > 0 && terminalDate && isTerminalMarkAfterCashFlows(baseCashFlows, terminalMark)) {
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
      if (inputs.isCashPosition || inputs.isBondPosition || inputs.isRealEstatePosition) {
        return;
      }
      cashFlows.push(...inputs.baseCashFlows);
      const terminalMark = inputs[markName];
      if (
        terminalMark.value > 0 &&
        terminalMark.date &&
        isTerminalMarkAfterCashFlows(inputs.baseCashFlows, terminalMark)
      ) {
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

function buildEntityCashFlowAuditRows(rows, markName = "internalMark") {
  return rows
    .filter((row) => row.includeInEntityPerformance)
    .flatMap((row) => {
      const companyName = row.latest && row.latest.company ? row.latest.company : "Unnamed investment";
      const inputs = buildPerformanceInputs(row.company.updates);
      const cashFlowRows = inputs.baseCashFlows.map((cashFlow) => ({
        company: companyName,
        date: cashFlow.date,
        type: cashFlow.amount < 0 ? "Contribution" : "Distribution",
        amount: cashFlow.amount,
        includedInXirr: true
      }));
      if (inputs.isCashPosition || inputs.isBondPosition || inputs.isRealEstatePosition) {
        return cashFlowRows;
      }
      const terminalMark = inputs[markName];
      const terminalIncluded =
        terminalMark &&
        terminalMark.value > 0 &&
        terminalMark.date &&
        isTerminalMarkAfterCashFlows(inputs.baseCashFlows, terminalMark);

      if (terminalMark && terminalMark.value > 0 && terminalMark.date) {
        cashFlowRows.push({
          company: companyName,
          date: terminalMark.date,
          type: "Terminal NAV",
          amount: terminalMark.value,
          includedInXirr: terminalIncluded
        });
      }

      return cashFlowRows;
    })
    .sort((left, right) => {
      const dateDelta = (left.date || new Date(0)) - (right.date || new Date(0));
      if (dateDelta) {
        return dateDelta;
      }

      return left.company.localeCompare(right.company);
    });
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
  const performanceRows = allEntityRows.filter(
    (row) => row && row.includeInEntityPerformance
  );
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
    performanceRows,
    (row) => row.includedReportedAmount
  );
  const totalInvestedCapital = sumEntityRows(
    performanceRows,
    (row) => row.performance.investedCapital
  );
  const officialNav = sumEntityRows(
    performanceRows,
    (row) => row.performance.officialValue
  );
  const internalNav = sumEntityRows(
    performanceRows,
    (row) => row.performance.internalValue
  );
  const updateRequestStats = buildUpdateRequestStats(investments);
  const publicStockRows = getPublicStockRows(investments).filter(isPublicStockRow);
  const savedPublicStockRows = publicStockRows.filter((investment) => !investment.isWatchlistOnly);
  const stockMarketValue = savedPublicStockRows.reduce((sum, investment) => sum + getStockMarketValue(investment), 0);
  const cashRows = getCashRows(investments);
  const totalCash = cashRows.reduce((sum, investment) => sum + getCashBalance(investment), 0);
  const fixedIncomeRows = getBondRows(investments);
  const fixedIncomeSummary = buildFixedIncomeSummary(fixedIncomeRows);
  const realEstateRows = getRealEstateRows(investments);
  const realEstateSummary = buildRealEstateSummary(realEstateRows);
  const totalUnfundedCommitments = Math.max(totalCommittedCapital - totalInvestedCapital, 0);
  const liquidityCoverage =
    totalUnfundedCommitments > 0 ? totalCash / totalUnfundedCommitments : null;
  const cashByEntityCards = getCashByEntity(investments).map(([entity, balance]) => ({
    label: `Cash - ${entity}`,
    value: formatMoney(balance),
    action: "cash"
  }));

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
    {
      label: "Update requests",
      value: `${updateRequestStats.sent} sent / ${updateRequestStats.awaiting} awaiting / ${updateRequestStats.followUp} follow-up`,
      action: "portfolio"
    },
    { label: "Public stocks", value: String(savedPublicStockRows.length), action: "public-stocks" },
    { label: "Stock market value", value: formatMoney(stockMarketValue), action: "public-stocks" },
    { label: "Fixed income positions", value: String(fixedIncomeRows.length), action: "fixed-income" },
    { label: "Fixed income market value", value: formatMoney(fixedIncomeSummary.marketValue), action: "fixed-income" },
    { label: "Real estate properties", value: String(realEstateRows.length), action: "real-estate" },
    { label: "Real estate net equity", value: formatMoney(realEstateSummary.netEquity), action: "real-estate" },
    { label: "Total cash", value: formatMoney(totalCash), action: "cash" },
    { label: "Total unfunded commitments", value: formatMoney(totalUnfundedCommitments), action: "portfolio" },
    { label: "Liquidity coverage", value: formatTurns(liquidityCoverage), action: "cash" },
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

  return cards.concat(cashByEntityCards, entityTotals);
}

function daysSinceDate(value) {
  const parsed = parseDateValue(value, null);
  if (!parsed) {
    return null;
  }

  return Math.floor((Date.now() - parsed.getTime()) / (24 * 60 * 60 * 1000));
}

function getEffectiveUpdateRequestStatus(investment) {
  const status = String((investment && investment.updateRequestStatus) || "").trim();
  if (status === "Requested") {
    const sentDaysAgo = daysSinceDate(investment && investment.lastUpdateRequestSentAt);
    if (sentDaysAgo !== null && sentDaysAgo >= 7) {
      return "Follow-up Needed";
    }
  }

  return status;
}

function buildUpdateRequestStats(investments) {
  const companySummaries = getCompanyCollections(investments);
  return companySummaries.reduce(
    (stats, company) => {
      const latest = company.latest || {};
      const status = getEffectiveUpdateRequestStatus(latest);
      if (latest.lastUpdateRequestSentAt || status) {
        stats.sent += 1;
      }
      if (status === "Requested") {
        stats.awaiting += 1;
      }
      if (status === "Follow-up Needed") {
        stats.followUp += 1;
      }
      return stats;
    },
    { sent: 0, awaiting: 0, followUp: 0 }
  );
}

function addQualityAlert(alerts, row, severity, title, detail) {
  const alertId =
    row.latest.id ||
    companyEntityKey(row.latest.company, row.latest.entity) ||
    `${companyKey(row.latest.company)}::${Date.now()}`;
  const alertKey = `${alertId}::${companyKey(title)}`;
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
      addQualityAlert(
        alerts,
        row,
        "High",
        "Missing entity",
        "Assign this investment to Beaman Ventures, Lee Beaman, Lee Beaman IRA, Katherine Trust, or Natalie Trust."
      );
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

    const performanceInputs = buildPerformanceInputs(row.company.updates);
    const latestCashFlowDate = getLatestCashFlowDate(performanceInputs.baseCashFlows);
    const valuationDate = parseDateValue(latest.valuationDate, null);
    if (
      latestCashFlowDate &&
      valuationDate &&
      latestCashFlowDate > valuationDate &&
      (performance.officialValue > 0 || performance.internalValue > 0)
    ) {
      addQualityAlert(
        alerts,
        row,
        "High",
        "Cash flow after valuation date",
        `Latest cash flow is ${formatDisplayDate(latestCashFlowDate)}, after the valuation date ${formatDisplayDate(valuationDate)}. Update the valuation date so XIRR can be calculated correctly.`
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
        const publicHoldings = buildPublicHoldingsSummary(entity);
        const cashHoldings = buildCashHoldingsSummary(entity);
        const fixedIncomeHoldings = buildFixedIncomeHoldingsSummary(entity);
        const realEstateHoldings = buildRealEstateHoldingsSummary(entity);
        const cashNavPercent =
          totals.internalValue > 0 ? cashHoldings.totalBalance / totals.internalValue : null;
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
        const isLeeBeamanIraEntity = normalizeEntityName(entity) === "Lee Beaman IRA";
        const publicHoldingsTitle = isLeeBeamanIraEntity ? entity : "Public equity";

        return `
        <section class="entity-performance-group" data-entity="${escapeHtml(entity)}" data-entity-group="${escapeHtml(entity)}">
          ${
            isLeeBeamanIraEntity
              ? ""
              : `
                <div class="entity-performance-group-header">
                  <div>
                    <p class="dashboard-label">Entity</p>
                    <h3>${escapeHtml(entity)}</h3>
                  </div>
                  ${isDashboardViewer() ? "" : '<span class="entity-open-pill">Open entity</span>'}
                </div>
              `
          }
          <div class="entity-performance-group-grid">
        ${
          isLeeBeamanIraEntity
            ? ""
            : `
              <article class="dashboard-card entity-performance-card entity-summary-card" data-entity="${escapeHtml(entity)}">
                <p class="dashboard-label">Entity summary</p>
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
            `
        }
        ${
          publicHoldings.positions.length
            ? `
              <article class="dashboard-card entity-performance-card entity-holdings-card public-holdings-card" data-holdings-action="public-stocks" data-entity="${escapeHtml(entity)}" role="button" tabindex="0">
                <div class="entity-performance-header">
                  <div>
                    <p class="dashboard-label">Public Holdings</p>
                    <h3>${escapeHtml(publicHoldingsTitle)}</h3>
                  </div>
                  <span class="status-chip">${escapeHtml(String(publicHoldings.positions.length))} position${publicHoldings.positions.length === 1 ? "" : "s"}</span>
                </div>
                <div class="entity-metric-grid">
                  ${[
                    { label: "Current public holdings value", value: formatMoney(publicHoldings.marketValue) },
                    { label: "Total public stock cost basis", value: formatMoney(publicHoldings.costBasis) },
                    { label: "Unrealized gain/loss $", value: formatSignedMoney(publicHoldings.gainLoss), performance: publicHoldings.gainLoss },
                    { label: "Unrealized gain/loss %", value: formatStockPercent(publicHoldings.gainLossPercent), performance: publicHoldings.gainLossPercent },
                    { label: "Public stock positions", value: String(publicHoldings.positions.length) },
                    { label: "YTD change %", value: formatStockPercent(publicHoldings.ytdChangePercent) }
                  ]
                    .map(
                      (metric) => `
                        <div class="entity-metric-box">
                          <p class="dashboard-label">${escapeHtml(metric.label)}</p>
                          <p class="dashboard-value ${metric.performance === undefined ? "" : escapeHtml(getStockPerformanceClass(metric.performance))}">${escapeHtml(metric.value)}</p>
                        </div>
                      `
                    )
                    .join("")}
                </div>
              </article>
            `
            : ""
        }
        ${
          fixedIncomeHoldings.positions.length
            ? `
              <article class="dashboard-card entity-performance-card entity-holdings-card fixed-income-holdings-card" data-holdings-action="fixed-income" data-entity="${escapeHtml(entity)}" role="button" tabindex="0">
                <div class="entity-performance-header">
                  <div>
                    <p class="dashboard-label">Fixed Income Holdings</p>
                    <h3>Fixed income</h3>
                  </div>
                  <span class="status-chip">${escapeHtml(String(fixedIncomeHoldings.positions.length))} position${fixedIncomeHoldings.positions.length === 1 ? "" : "s"}</span>
                </div>
                <div class="entity-metric-grid">
                  ${[
                    { label: "Current fixed income market value", value: formatMoney(fixedIncomeHoldings.marketValue) },
                    { label: "Total par value", value: formatMoney(fixedIncomeHoldings.parValue) },
                    { label: "Annual coupon income", value: formatMoney(fixedIncomeHoldings.annualCouponIncome) },
                    { label: "Weighted average coupon", value: formatPercent(fixedIncomeHoldings.weightedCoupon) },
                    { label: "Weighted average YTM", value: formatPercent(fixedIncomeHoldings.weightedYtm) },
                    { label: "Bond positions", value: String(fixedIncomeHoldings.positions.length) },
                    { label: "Next maturity", value: fixedIncomeHoldings.nextMaturity || "N/A" }
                  ]
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
            `
            : ""
        }
        ${
          realEstateHoldings.positions.length
            ? `
              <article class="dashboard-card entity-performance-card entity-holdings-card real-estate-holdings-card" data-holdings-action="real-estate" data-entity="${escapeHtml(entity)}" role="button" tabindex="0">
                <div class="entity-performance-header">
                  <div>
                    <p class="dashboard-label">Real Estate Holdings</p>
                    <h3>Real estate</h3>
                  </div>
                  <span class="status-chip">${escapeHtml(String(realEstateHoldings.positions.length))} propert${realEstateHoldings.positions.length === 1 ? "y" : "ies"}</span>
                </div>
                <div class="entity-metric-grid">
                  ${[
                    { label: "Appraised value", value: formatMoney(realEstateHoldings.appraisedValue) },
                    { label: "Net equity value", value: formatMoney(realEstateHoldings.netEquity) },
                    { label: "Total debt", value: formatMoney(realEstateHoldings.debt) },
                    { label: "Annual NOI", value: formatMoney(realEstateHoldings.noi) },
                    { label: "Weighted average cap rate", value: formatPercent(realEstateHoldings.weightedCapRate) },
                    { label: "Weighted average LTV", value: formatPercent(realEstateHoldings.weightedLtv) },
                    { label: "Properties", value: String(realEstateHoldings.positions.length) },
                    { label: "Latest appraisal date", value: realEstateHoldings.latestAppraisalDate || "N/A" }
                  ]
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
            `
            : ""
        }
        ${
          cashHoldings.positions.length
            ? `
              <article class="dashboard-card entity-performance-card entity-holdings-card cash-holdings-card" data-holdings-action="cash" data-entity="${escapeHtml(entity)}" role="button" tabindex="0">
                <div class="entity-performance-header">
                  <div>
                    <p class="dashboard-label">Cash Holdings</p>
                    <h3>Cash</h3>
                  </div>
                  <span class="status-chip">${escapeHtml(String(cashHoldings.positions.length))} account${cashHoldings.positions.length === 1 ? "" : "s"}</span>
                </div>
                <div class="entity-metric-grid">
                  ${[
                    { label: "Total cash balance", value: formatMoney(cashHoldings.totalBalance) },
                    { label: "Cash accounts", value: String(cashHoldings.positions.length) },
                    { label: "Cash as % of Internal NAV", value: formatStockPercent(cashNavPercent) },
                    {
                      label: "Largest cash account",
                      value: cashHoldings.largestAccount
                        ? `${cashHoldings.largestAccount.investment.company || "Unnamed account"} ${formatMoney(cashHoldings.largestAccount.balance)}`
                        : "N/A"
                    },
                    { label: "Latest balance date", value: cashHoldings.latestBalanceDate || "N/A" }
                  ]
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
            `
            : ""
        }
          </div>
        </section>
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
    { label: "Internal XIRR", value: formatPercent(performance.internal.xirr) },
    { label: "Internal MOIC", value: formatTurns(performance.internal.moic) },
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

  const investmentMarkup = investments.length
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

  entityDetailInvestments.innerHTML = investmentMarkup;
}

function getAllEntityNames() {
  return Array.from(
    new Set(
      configuredEntities
        .concat(allInvestments.map((item) => normalizeEntityName(item.entity)).filter(Boolean))
        .map(normalizeEntityName)
    )
  ).sort();
}

function renderXirrAuditOptions() {
  if (!xirrAuditEntitySelect) {
    return;
  }

  const entities = getAllEntityNames();
  if (!selectedXirrAuditEntity && entities.length) {
    selectedXirrAuditEntity = entities[0];
  }

  xirrAuditEntitySelect.innerHTML = ['<option value="">Select an entity</option>']
    .concat(entities.map((entity) => `<option value="${escapeHtml(entity)}">${escapeHtml(entity)}</option>`))
    .join("");
  xirrAuditEntitySelect.value = entities.includes(selectedXirrAuditEntity) ? selectedXirrAuditEntity : "";
  selectedXirrAuditEntity = xirrAuditEntitySelect.value;
}

function buildXirrAuditMarkup(cashFlowAuditRows) {
  return cashFlowAuditRows.length
    ? `
      <section class="cash-flow-audit">
        <div class="panel-header panel-header-stack">
          <div>
            <h3>XIRR cash-flow audit</h3>
            <p class="section-copy">These dated cash flows feed the entity Internal XIRR. Excluded NAV rows usually mean the valuation date is before the latest contribution.</p>
          </div>
        </div>
        <table class="reconciliation-table">
          <thead>
            <tr>
              <th>Date</th>
              <th>Company</th>
              <th>Type</th>
              <th>Amount</th>
              <th>XIRR use</th>
            </tr>
          </thead>
          <tbody>
            ${cashFlowAuditRows
              .map(
                (row) => `
                  <tr>
                    <td>${escapeHtml(formatDisplayDate(row.date))}</td>
                    <td>${escapeHtml(row.company)}</td>
                    <td>${escapeHtml(row.type)}</td>
                    <td>${escapeHtml(formatMoney(row.amount))}</td>
                    <td>${escapeHtml(row.includedInXirr ? "Included" : "Excluded: valuation before cash flow")}</td>
                  </tr>
                `
              )
              .join("")}
          </tbody>
        </table>
      </section>
    `
    : '<p class="update-meta">No dated cash flows are available for XIRR yet.</p>';
}

function renderXirrAudit() {
  if (!xirrAuditSummary || !xirrAuditTable) {
    return;
  }

  renderXirrAuditOptions();
  if (!selectedXirrAuditEntity) {
    xirrAuditSummary.innerHTML = "";
    xirrAuditTable.innerHTML = '<p class="update-meta">Select an entity to review XIRR cash flows.</p>';
    return;
  }

  const entityRows = buildEntityRows(allInvestments, selectedXirrAuditEntity);
  const performance = buildEntityRowTotals(entityRows);
  const cashFlowAuditRows = buildEntityCashFlowAuditRows(entityRows, "internalMark");

  xirrAuditSummary.innerHTML = [
    { label: "Entity", value: selectedXirrAuditEntity },
    { label: "Internal XIRR", value: formatPercent(performance.internal.xirr) },
    { label: "Internal NAV", value: formatMoney(performance.internalValue) },
    { label: "Cash-flow rows", value: String(cashFlowAuditRows.length) }
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
  xirrAuditTable.innerHTML = buildXirrAuditMarkup(cashFlowAuditRows);
}

function renderFilterOptions() {
  const entities = getAllEntityNames();
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

function renderConfiguredEntitySelects() {
  const renderSelect = (element, placeholder) => {
    if (!element) {
      return;
    }
    const currentValue = normalizeEntityName(element.value);
    element.innerHTML = [`<option value="">${placeholder}</option>`]
      .concat(
        configuredEntities.map(
          (entity) => `<option value="${escapeHtml(entity)}">${escapeHtml(entity)}</option>`
        )
      )
      .join("");
    element.value = configuredEntities.includes(currentValue) ? currentValue : "";
  };

  renderSelect(form && form.elements ? form.elements.entity : null, "Select entity");
  renderSelect(taskForm && taskForm.elements ? taskForm.elements.entity : null, "Select entity");
  renderSelect(aiAnalystEntityField, "Any entity");
  renderSelect(aiAnalysisEntityField, "Let AI match entity");
}

function renderAiAnalysisInvestmentOptions() {
  if (!aiAnalysisInvestmentField) {
    return;
  }

  const currentValue = aiAnalysisInvestmentField.value;
  aiAnalysisInvestmentField.innerHTML = ['<option value="">Let AI match investment</option>']
    .concat(
      sortInvestmentsAlphabetically(allInvestments).map(
        (investment) =>
          `<option value="${escapeHtml(investment.id)}">${escapeHtml(investment.company)}${investment.entity ? ` - ${escapeHtml(investment.entity)}` : ""}</option>`
      )
    )
    .join("");
  aiAnalysisInvestmentField.value = allInvestments.some((investment) => investment.id === currentValue)
    ? currentValue
    : "";
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
  form.elements.assetType.value = investment.assetType || "Private Investment";
  form.elements.ticker.value = investment.ticker || "";
  form.elements.exchange.value = investment.exchange || "";
  form.elements.shareCount.value = investment.shareCount || "";
  form.elements.costBasisPerShare.value = investment.costBasisPerShare || "";
  form.elements.marketPrice.value = investment.marketPrice || "";
  form.elements.marketPriceDate.value = investment.marketPriceDate || "";
  [
    "bondIssuer",
    "bondDescription",
    "bondType",
    "bondCusip",
    "bondEntityOwner",
    "bondParValue",
    "bondPurchasePrice",
    "bondPurchaseDate",
    "bondCostBasis",
    "bondCouponRate",
    "bondCouponFrequency",
    "bondMaturityDate",
    "bondCallDate",
    "bondCallPrice",
    "bondCurrentPrice",
    "bondMarketPriceDate",
    "bondMarketValue",
    "bondYieldToMaturity",
    "bondYieldToCall",
    "bondCurrentYield",
    "bondCreditRating",
    "bondInsurer",
    "bondTaxStatus",
    "bondAccruedInterest"
  ].forEach((fieldName) => {
    if (form.elements[fieldName]) {
      form.elements[fieldName].value = investment[fieldName] || "";
    }
  });
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
  form.elements.reportingCadence.value = investment.reportingCadence || "";
  form.elements.updateRequestStatus.value = investment.updateRequestStatus || "";
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
  updateStockDetailsVisibility();
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
  updateStockDetailsVisibility({ clearHiddenFields: true });
}

function prefillStockPosition(preset) {
  resetFormToCreateMode();
  const defaults = {
    company: "",
    entity: "Beaman Ventures",
    assetType: "Public Stock",
    ticker: "",
    exchange: "NASDAQ",
    currency: "USD",
    stage: "Public Equity",
    status: "Active",
    notes: ""
  };
  const stock = { ...defaults, ...preset };

  Object.entries(stock).forEach(([fieldName, value]) => {
    if (form.elements[fieldName]) {
      form.elements[fieldName].value = value;
    }
  });

  if (notesField && stock.notes) {
    notesField.value = stock.notes;
  }

  formMessage.textContent = `${stock.company} stock position ready. Add shares, cost basis, current price, and save.`;
  updateStockDetailsVisibility();
  showWorkspaceView("capture");
  window.scrollTo({ top: 0, behavior: "smooth" });
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
    assetType: investment.assetType || "",
    ticker: investment.ticker || "",
    exchange: investment.exchange || "",
    shareCount: investment.shareCount || "",
    costBasisPerShare: investment.costBasisPerShare || "",
    marketPrice: investment.marketPrice || "",
    marketPriceDate: investment.marketPriceDate || "",
    marketValue: investment.marketValue || "",
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
    reportingCadence: investment.reportingCadence || "",
    updateRequestStatus: investment.updateRequestStatus || "",
    lastUpdateRequestSentAt: investment.lastUpdateRequestSentAt || "",
    lastUpdateRequestSubject: investment.lastUpdateRequestSubject || "",
    lastUpdateRequestContact: investment.lastUpdateRequestContact || "",
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

function timelineDateValue(value, fallback = "") {
  return String(value || fallback || "").trim();
}

function timelineAmountValue(value, currency = "USD") {
  const amount = toNumber(value);
  return amount ? `${currency || "USD"} ${normalizeMoneyString(value)}` : "";
}

function addTimelineEvent(events, event) {
  const date = timelineDateValue(event.date, event.fallbackDate);
  const hasContent = date || event.title || event.description || event.amount;
  if (!hasContent) {
    return;
  }

  const timelineEvent = {
    date,
    type: String(event.type || "Update").trim() || "Update",
    title: String(event.title || event.type || "Timeline event").trim(),
    amount: String(event.amount || "").trim(),
    description: String(event.description || "").trim(),
    source: String(event.source || "").trim(),
    category: String(event.category || event.type || "Update").trim()
  };
  const eventKey = [
    timelineEvent.date,
    timelineEvent.type,
    timelineEvent.title,
    timelineEvent.amount,
    timelineEvent.description,
    timelineEvent.source
  ]
    .join("::")
    .toLowerCase();
  const alreadyIncluded = events.some((existingEvent) => {
    const existingKey = [
      existingEvent.date,
      existingEvent.type,
      existingEvent.title,
      existingEvent.amount,
      existingEvent.description,
      existingEvent.source
    ]
      .join("::")
      .toLowerCase();
    return existingKey === eventKey;
  });
  if (!alreadyIncluded) {
    events.push(timelineEvent);
  }
}

function buildInvestmentTimelineEvents(companyRecord, companyUpdates, relatedTasks) {
  const events = [];
  const companyDocuments = companyRecord && Array.isArray(companyRecord.documents)
    ? companyRecord.documents
    : [];

  companyUpdates.forEach((investment) => {
    const currency = investment.currency || "USD";
    addTimelineEvent(events, {
      date: investment.createdAt,
      type: "Update",
      title: investment.createdAt === (companyUpdates[companyUpdates.length - 1] || {}).createdAt
        ? "Initial investment record"
        : "Investment update saved",
      amount: timelineAmountValue(investment.amount, currency),
      description: summarizeText(investment.notes, investment.nextStep || investment.status || ""),
      source: investment.submittedBy ? `Submitted by ${investment.submittedBy}` : "Investment record"
    });

    normalizeCapitalActivityRows(
      investment.capitalActivity && investment.capitalActivity.length
        ? investment.capitalActivity
        : buildLegacyCapitalActivityRows(investment)
    ).forEach((activity) => {
      const normalizedType = String(activity.type || "").toLowerCase();
      const eventType = normalizedType.includes("distribution") ||
        normalizedType.includes("dividend") ||
        normalizedType.includes("return") ||
        normalizedType.includes("exit")
        ? "Distribution"
        : "Capital Call";
      addTimelineEvent(events, {
        date: activity.date,
        fallbackDate: investment.createdAt,
        type: eventType,
        title: activity.type || eventType,
        amount: timelineAmountValue(activity.amount, currency),
        description: activity.notes || "",
        source: "Capital activity"
      });
    });

    if (investment.valuationDate || investment.officialValue || investment.internalValue || investment.exitValue) {
      const marks = [
        investment.officialValue ? `Official ${formatMoney(investment.officialValue)}` : "",
        investment.internalValue ? `Internal ${formatMoney(investment.internalValue)}` : "",
        investment.exitValue ? `Exit ${formatMoney(investment.exitValue)}` : ""
      ].filter(Boolean);
      addTimelineEvent(events, {
        date: investment.valuationDate,
        fallbackDate: investment.createdAt,
        type: "Valuation",
        title: "Valuation updated",
        amount: marks.join(" / "),
        description: "Latest mark captured for this investment.",
        source: "Valuation"
      });
    }

    if (investment.ownershipPercent || investment.entityOwnershipPercent || investment.ownershipNotes) {
      addTimelineEvent(events, {
        date: investment.createdAt,
        type: "Ownership",
        title: "Ownership updated",
        amount: [
          investment.ownershipPercent ? `Total ${investment.ownershipPercent}%` : "",
          investment.entityOwnershipPercent ? `Entity ${investment.entityOwnershipPercent}%` : ""
        ].filter(Boolean).join(" / "),
        description: investment.ownershipNotes || "",
        source: "Ownership"
      });
    }

    if (investment.followOnCapitalAmount || investment.followOnCapitalStatus || investment.followOnCapitalNotes) {
      addTimelineEvent(events, {
        date: investment.createdAt,
        type: "Follow-on",
        title: investment.followOnCapitalStatus || "Follow-on capital update",
        amount: timelineAmountValue(investment.followOnCapitalAmount, currency),
        description: investment.followOnCapitalNotes || "",
        source: "Follow-on"
      });
    }

    if (investment.decisionDate || investment.decisionType || investment.decisionSummary) {
      addTimelineEvent(events, {
        date: investment.decisionDate,
        fallbackDate: investment.createdAt,
        type: "Decision",
        title: investment.decisionType || "Decision logged",
        description: investment.decisionSummary || "",
        source: "Decision log"
      });
    }

    if (investment.deckSummary) {
      addTimelineEvent(events, {
        date: investment.createdAt,
        type: "Research",
        title: "Deck summary saved",
        description: summarizeText(investment.deckSummary, ""),
        source: "AI research"
      });
    }

    if (investment.documentLinks) {
      addTimelineEvent(events, {
        date: investment.createdAt,
        type: "Document",
        title: "Document links added",
        description: summarizeText(investment.documentLinks, ""),
        source: "Linked documents"
      });
    }

    if (Array.isArray(investment.documents)) {
      investment.documents.forEach((document) => {
        addTimelineEvent(events, {
          date: document.uploadedAt,
          fallbackDate: investment.createdAt,
          type: "Document",
          title: document.name || "Document uploaded",
          description: document.url || "",
          source: "Uploaded document"
        });
      });
    }

    if (investment.lastUpdateRequestSentAt) {
      addTimelineEvent(events, {
        date: investment.lastUpdateRequestSentAt,
        type: "Update",
        title: "Update request sent",
        description: investment.lastUpdateRequestSubject || investment.lastUpdateRequestContact || "",
        source: investment.lastUpdateRequestContact || "Update request"
      });
    }
  });

  const structuredCollections = [
    { rows: companyRecord ? companyRecord.valuationHistory : [], type: "Valuation", source: "Valuation history" },
    { rows: companyRecord ? companyRecord.ownershipHistory : [], type: "Ownership", source: "Ownership history" },
    { rows: companyRecord ? companyRecord.followOnHistory : [], type: "Follow-on", source: "Follow-on history" },
    { rows: companyRecord ? companyRecord.decisionLog : [], type: "Decision", source: "Decision log" },
    { rows: companyRecord ? companyRecord.researchEntries : [], type: "Research", source: "Research" },
    { rows: companyRecord ? companyRecord.reportUpdates : [], type: "Update", source: "Investor update" }
  ];

  structuredCollections.forEach(({ rows, type, source }) => {
    (Array.isArray(rows) ? rows : []).forEach((row) => {
      addTimelineEvent(events, {
        date: row.date,
        type: row.type || type,
        title: row.title || row.type || type,
        amount: row.amount || [
          row.officialValue ? `Official ${formatMoney(row.officialValue)}` : "",
          row.internalValue ? `Internal ${formatMoney(row.internalValue)}` : "",
          row.exitValue ? `Exit ${formatMoney(row.exitValue)}` : ""
        ].filter(Boolean).join(" / "),
        description:
          row.aiSummary ||
          row.summary ||
          row.notes ||
          row.originalNotes ||
          row.keyMetrics ||
          row.actionItems ||
          "",
        source: row.sourceType || row.reportPeriod || source
      });
    });
  });

  companyDocuments.forEach((document) => {
    addTimelineEvent(events, {
      date: document.uploadedAt,
      type: "Document",
      title: document.name || "Document uploaded",
      description: document.notes || document.url || "",
      source: document.source || "Company vault"
    });
  });

  relatedTasks.forEach((task) => {
    addTimelineEvent(events, {
      date: task.completedAt || task.dueDate || task.createdAt,
      type: "Task",
      title: task.title || "Task",
      description: task.description || task.status || "",
      source: [task.category, task.status, task.assignee].filter(Boolean).join(" / ")
    });
  });

  return events.sort((left, right) => {
    const rightTime = parseDateValue(right.date, new Date(0)).getTime();
    const leftTime = parseDateValue(left.date, new Date(0)).getTime();
    return rightTime - leftTime;
  });
}

function renderInvestmentTimeline(events) {
  if (!investmentTimeline) {
    return;
  }

  investmentTimeline.innerHTML = events.length
    ? events
        .map(
          (event) => `
            <article class="investment-timeline-event investment-timeline-${escapeHtml(
              event.type.toLowerCase().replace(/[^a-z0-9]+/g, "-")
            )}">
              <div class="investment-timeline-marker">
                <span class="timeline-type-label">${escapeHtml(event.type)}</span>
                <span class="dashboard-label">${escapeHtml(formatDisplayDateOrText(event.date))}</span>
              </div>
              <div class="investment-timeline-body">
                <div class="update-head">
                  <h4>${escapeHtml(event.title)}</h4>
                  ${event.amount ? `<span class="status-chip">${escapeHtml(event.amount)}</span>` : ""}
                </div>
                ${event.description ? `<p class="update-meta">${escapeHtml(summarizeText(event.description, ""))}</p>` : ""}
                ${event.source ? `<p class="dashboard-label">${escapeHtml(event.source)}</p>` : ""}
              </div>
            </article>
          `
        )
        .join("")
    : '<p class="update-meta">No dated investment history is available yet.</p>';
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
    if (investmentTimeline) {
      investmentTimeline.innerHTML = "";
    }
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
  renderInvestmentTimeline(buildInvestmentTimelineEvents(companyRecord, companyUpdates, relatedTasks));
  const performance =
    companyPerformanceMap.get(companyEntityKey(selectedCompany, selectedCompanyEntity)) ||
    buildCompanyPerformance(companyUpdates);
  companyPanelTitle.textContent = latest.company || selectedCompany;
  companyPanelCopy.textContent = companyRecord
    ? `${companyUpdates.length} update${companyUpdates.length === 1 ? "" : "s"} organized into one operating file for research, capital history, decisions, reminders, and source material.`
    : `${companyUpdates.length} update${companyUpdates.length === 1 ? "" : "s"} saved for this company in one operating record.`;
  if (requestLatestUpdateButton) {
    requestLatestUpdateButton.classList.toggle("hidden", !canEditWorkspace());
    requestLatestUpdateButton.disabled = !latest.contactEmail;
    requestLatestUpdateButton.title = latest.contactEmail
      ? "Request the latest update from this investment contact"
      : "Add a primary contact email before requesting an update";
  }
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
    { label: "Phone", value: latest.contactPhone || "Not set" },
    { label: "Reporting cadence", value: latest.reportingCadence || "Unknown" },
    {
      label: "Update request status",
      value: getEffectiveUpdateRequestStatus(latest) || "Not requested"
    },
    {
      label: "Last request sent",
      value: latest.lastUpdateRequestSentAt
        ? formatDisplayDate(latest.lastUpdateRequestSentAt)
        : "Not sent"
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

function getAiProposalInvestmentName(proposal) {
  return (
    (proposal && proposal.investment && proposal.investment.company) ||
    "Unmatched investment"
  );
}

function getAiProposalEntity(proposal) {
  return (
    (proposal && proposal.investment && proposal.investment.entity) ||
    (proposal && proposal.entityId) ||
    "No entity"
  );
}

function formatConfidence(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed > 0 ? `${Math.round(parsed)}%` : "Not scored";
}

function summarizeJsonValue(value) {
  if (value === null || value === undefined || value === "") {
    return "Not provided";
  }
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  return JSON.stringify(value, null, 2);
}

function formatEvidenceStatus(status) {
  const normalized = String(status || "").trim().toLowerCase();
  if (normalized === "verified") {
    return "Verified";
  }
  if (normalized === "probable") {
    return "Probable";
  }
  return "Not verified";
}

function evidenceStatusClass(status) {
  const normalized = String(status || "").trim().toLowerCase();
  return normalized === "verified"
    ? "ai-evidence-verified"
    : normalized === "probable"
      ? "ai-evidence-probable"
      : "ai-evidence-unresolved";
}

function normalizeProposedChanges(value) {
  if (Array.isArray(value)) {
    return value;
  }
  if (value && typeof value === "object") {
    return Object.entries(value).map(([field, change]) =>
      change && typeof change === "object"
        ? { field, ...change }
        : { field, proposedValue: change }
    );
  }
  return [];
}

function getProposedChangeField(change) {
  return change.field || change.action || change.actionType || change.type || "Proposed change";
}

function getProposedChangeCurrentValue(change) {
  return change.currentValue !== undefined
    ? change.currentValue
    : change.current_value !== undefined
      ? change.current_value
      : change.current !== undefined
        ? change.current
        : "";
}

function getProposedChangeProposedValue(change) {
  return change.proposedValue !== undefined
    ? change.proposedValue
    : change.proposed_value !== undefined
      ? change.proposed_value
      : change.value !== undefined
        ? change.value
        : "";
}

function getProposedChangeEvidence(change) {
  return change.sourceEvidence || change.source_evidence || change.evidence || change.source || "";
}

function renderExtractedData(value) {
  if (!value || (typeof value === "object" && !Array.isArray(value) && !Object.keys(value).length)) {
    return '<p class="update-meta">No extracted information staged yet.</p>';
  }

  if (value && typeof value === "object" && !Array.isArray(value)) {
    return `
      <div class="ai-extracted-grid">
        ${Object.entries(value)
          .map(
            ([key, entry]) => `
              <article class="ai-extracted-item">
                <p class="dashboard-label">${escapeHtml(key)}</p>
                <pre>${escapeHtml(summarizeJsonValue(entry))}</pre>
              </article>
            `
          )
          .join("")}
      </div>
    `;
  }

  return `<pre class="ai-json-block">${escapeHtml(summarizeJsonValue(value))}</pre>`;
}

function renderProposedChanges(value) {
  const changes = normalizeProposedChanges(value);
  if (!changes.length) {
    return '<p class="update-meta">No proposed field-level changes staged yet.</p>';
  }

  return `
    <div class="ai-change-list">
      ${changes
        .map((change) => {
          const currentValue = getProposedChangeCurrentValue(change);
          const proposedValue = getProposedChangeProposedValue(change);
          return `
            <article class="ai-change-card">
              <div class="update-head">
                <h4>${escapeHtml(getProposedChangeField(change))}</h4>
                <span class="status-chip">${escapeHtml(formatConfidence(change.confidence))}</span>
              </div>
              <div class="ai-change-comparison">
                <div>
                  <p class="dashboard-label">Current</p>
                  <pre>${escapeHtml(summarizeJsonValue(currentValue))}</pre>
                </div>
                <div>
                  <p class="dashboard-label">Proposed</p>
                  <pre>${escapeHtml(summarizeJsonValue(proposedValue))}</pre>
                </div>
              </div>
              ${
                getProposedChangeEvidence(change)
                  ? `<p class="update-meta"><strong>Evidence:</strong> ${escapeHtml(getProposedChangeEvidence(change))}</p>`
                  : ""
              }
              ${
                change.notes
                  ? `<p class="update-meta"><strong>Notes:</strong> ${escapeHtml(change.notes)}</p>`
                  : ""
              }
            </article>
          `;
        })
        .join("")}
    </div>
  `;
}

function renderAiAnalysisReview() {
  if (!aiUpdateAnalysisReview) {
    return;
  }

  if (!latestAiUpdateAnalysis) {
    aiUpdateAnalysisReview.classList.add("hidden");
    aiUpdateAnalysisReview.innerHTML = "";
    return;
  }

  const analysis = latestAiUpdateAnalysis.analysis || {};
  const investmentMatch = analysis.investmentMatch || {};
  const entityMatch = analysis.entityMatch || {};
  const selectedInvestment = aiAnalysisInvestmentField && aiAnalysisInvestmentField.value
    ? allInvestments.find((investment) => investment.id === aiAnalysisInvestmentField.value)
    : null;
  const selectedEntity = aiAnalysisEntityField && aiAnalysisEntityField.value
    ? aiAnalysisEntityField.value
    : "";
  const effectiveInvestmentId = selectedInvestment
    ? selectedInvestment.id
    : investmentMatch.investmentId || "";
  const effectiveInvestmentName = selectedInvestment
    ? selectedInvestment.company
    : investmentMatch.investmentName || "Unmatched";
  const effectiveEntity = selectedEntity || entityMatch.entityName || entityMatch.entityId || "";
  const warnings = Array.isArray(analysis.warnings) ? analysis.warnings : [];
  const unresolved = Array.isArray(analysis.unresolved) ? analysis.unresolved : [];
  const facts = Array.isArray(analysis.extractedFacts) ? analysis.extractedFacts : [];
  const whatChanged = Array.isArray(analysis.whatChanged) ? analysis.whatChanged : [];
  const proposedChanges = Array.isArray(analysis.proposedChanges) ? analysis.proposedChanges : [];
  const needsMatch = !effectiveInvestmentId;

  aiUpdateAnalysisReview.classList.remove("hidden");
  aiUpdateAnalysisReview.innerHTML = `
    <section class="ai-detail-section">
      <div class="panel-header">
        <div>
          <h4>Analysis Review</h4>
          <p class="section-copy">Review this result before creating a pending inbox proposal.</p>
        </div>
        <span class="status-chip">${escapeHtml(formatDisplayDate(latestAiUpdateAnalysis.analyzedAt))}</span>
      </div>
      <div class="company-summary-grid">
        <article class="company-summary-card">
          <p class="dashboard-label">Matched investment</p>
          <p class="highlight-value">${escapeHtml(effectiveInvestmentName)}</p>
          <p class="update-meta">${escapeHtml(formatConfidence(investmentMatch.confidence))}</p>
          <p class="update-meta">${escapeHtml(investmentMatch.reason || "No reason provided.")}</p>
        </article>
        <article class="company-summary-card">
          <p class="dashboard-label">Matched entity</p>
          <p class="highlight-value">${escapeHtml(effectiveEntity || "Unresolved")}</p>
          <p class="update-meta">${escapeHtml(formatConfidence(entityMatch.confidence))}</p>
          <p class="update-meta">${escapeHtml(entityMatch.reason || "No reason provided.")}</p>
        </article>
      </div>
    </section>

    <section class="ai-detail-section">
      <h4>Warnings / Needs Review</h4>
      ${
        needsMatch || warnings.length || unresolved.length
          ? `<div class="ai-warning-list">${[
              needsMatch ? "Choose an investment before creating the proposal." : "",
              ...warnings,
              ...unresolved
            ]
              .filter(Boolean)
              .map((item) => `<p>${escapeHtml(item)}</p>`)
              .join("")}</div>`
          : '<p class="update-meta">No warnings returned.</p>'
      }
    </section>

    <section class="ai-detail-section">
      <h4>What Changed</h4>
      ${
        whatChanged.length
          ? `<ul class="ai-bullet-list">${whatChanged.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`
          : '<p class="update-meta">No material change summary returned.</p>'
      }
    </section>

    <section class="ai-detail-section">
      <h4>Extracted Facts</h4>
      ${
        facts.length
          ? `<div class="ai-fact-list">${facts
              .map(
                (fact) => `
                  <article class="ai-extracted-item">
                    <p class="dashboard-label">${escapeHtml(fact.category || "Fact")}</p>
                    <p class="highlight-value">${escapeHtml(fact.field || "Unlabeled fact")}: ${escapeHtml(summarizeJsonValue(fact.value))}</p>
                    <p class="update-meta">${escapeHtml([fact.unit, fact.period, fact.date, fact.factType].filter(Boolean).join(" • ") || "No period or type")}</p>
                    <p class="update-meta"><strong>Source evidence:</strong> ${escapeHtml(fact.sourceEvidence || "Not verified")}</p>
                    <div class="ai-evidence-row">
                      <span class="status-chip ${evidenceStatusClass(fact.evidenceStatus)}">${escapeHtml(formatEvidenceStatus(fact.evidenceStatus))}</span>
                      <span class="status-chip">${escapeHtml(formatConfidence(fact.confidence))}</span>
                    </div>
                  </article>
                `
              )
              .join("")}</div>`
          : '<p class="update-meta">No extracted facts returned.</p>'
      }
    </section>

    <section class="ai-detail-section">
      <h4>Proposed Changes</h4>
      ${
        proposedChanges.length
          ? `<div class="ai-change-list">${proposedChanges
              .map(
                (change) => `
                  <article class="ai-change-card ${change.riskLevel === "high" ? "ai-change-card-high-risk" : ""}">
                    <div class="update-head">
                      <h4>${escapeHtml(change.field || change.actionType || "Proposed change")}</h4>
                      <span class="status-chip">${escapeHtml(change.riskLevel || "medium")} risk</span>
                    </div>
                    <div class="ai-change-comparison">
                      <div><p class="dashboard-label">Current</p><pre>${escapeHtml(summarizeJsonValue(change.currentValue))}</pre></div>
                      <div><p class="dashboard-label">Proposed</p><pre>${escapeHtml(summarizeJsonValue(change.proposedValue))}</pre></div>
                    </div>
                    <p class="update-meta">${escapeHtml([change.period, change.date, formatConfidence(change.confidence)].filter(Boolean).join(" • "))}</p>
                    <p class="update-meta"><strong>Source evidence:</strong> ${escapeHtml(change.sourceEvidence || "Not verified")}</p>
                    <span class="status-chip ${evidenceStatusClass(change.evidenceStatus)}">${escapeHtml(formatEvidenceStatus(change.evidenceStatus))}</span>
                    ${change.notes ? `<p class="update-meta">${escapeHtml(change.notes)}</p>` : ""}
                  </article>
                `
              )
              .join("")}</div>`
          : '<p class="update-meta">No proposed changes returned.</p>'
      }
    </section>

    <div class="card-actions">
      <button type="button" data-action="create-ai-analysis-proposal"${needsMatch ? " disabled" : ""}>Create Proposal</button>
      <button type="button" class="secondary-button" data-action="focus-ai-analysis-investment">Change Investment</button>
      <button type="button" class="secondary-button" data-action="focus-ai-analysis-entity">Change Entity</button>
      <button type="button" class="secondary-button danger-button" data-action="cancel-ai-analysis">Cancel</button>
    </div>
  `;
}

function openAiUpdateAnalyzer() {
  if (!aiUpdateAnalyzerPanel) {
    return;
  }
  renderAiAnalysisInvestmentOptions();
  renderConfiguredEntitySelects();
  aiUpdateAnalyzerPanel.classList.remove("hidden");
  latestAiUpdateAnalysis = null;
  renderAiAnalysisReview();
  if (aiUpdateInboxMessage) {
    aiUpdateInboxMessage.textContent = "";
  }
}

function closeAiUpdateAnalyzer() {
  latestAiUpdateAnalysis = null;
  if (aiUpdateAnalysisForm) {
    aiUpdateAnalysisForm.reset();
  }
  if (aiUpdateAnalyzerPanel) {
    aiUpdateAnalyzerPanel.classList.add("hidden");
  }
  renderAiAnalysisReview();
}

function buildProposalPayloadFromAnalysis() {
  if (!latestAiUpdateAnalysis) {
    return null;
  }

  const analysis = latestAiUpdateAnalysis.analysis || {};
  const source = latestAiUpdateAnalysis.source || {};
  const selectedInvestment = aiAnalysisInvestmentField && aiAnalysisInvestmentField.value
    ? allInvestments.find((investment) => investment.id === aiAnalysisInvestmentField.value)
    : null;
  const investmentId = selectedInvestment
    ? selectedInvestment.id
    : (analysis.investmentMatch && analysis.investmentMatch.investmentId) || "";
  const selectedEntity = aiAnalysisEntityField && aiAnalysisEntityField.value
    ? aiAnalysisEntityField.value
    : "";
  const entityId = selectedEntity || (analysis.entityMatch && (analysis.entityMatch.entityId || analysis.entityMatch.entityName)) || "";

  if (!investmentId) {
    throw new Error("Choose an investment before creating the proposal.");
  }

  return {
    investmentId,
    entityId,
    sourceType: source.sourceType,
    sourceIdentifier: source.sourceIdentifier || [source.sender, source.subject, source.sourceDate].filter(Boolean).join(" | "),
    sourceDate: source.sourceDate,
    sender: source.sender,
    subject: source.subject,
    confidenceScore: analysis.investmentMatch ? analysis.investmentMatch.confidence : 0,
    matchReason: [
      analysis.investmentMatch && analysis.investmentMatch.reason,
      analysis.entityMatch && analysis.entityMatch.reason,
      selectedInvestment ? "Investment manually selected in the review form." : "",
      selectedEntity ? "Entity manually selected in the review form." : ""
    ]
      .filter(Boolean)
      .join(" "),
    summary: Array.isArray(analysis.whatChanged) && analysis.whatChanged.length
      ? analysis.whatChanged.map((item) => `• ${item}`).join("\n")
      : "AI analysis completed; review extracted facts and proposed changes.",
    extractedData: {
      facts: analysis.extractedFacts || [],
      warnings: analysis.warnings || [],
      unresolved: analysis.unresolved || [],
      candidates: analysis.candidates || [],
      analyzedAt: latestAiUpdateAnalysis.analyzedAt
    },
    proposedChanges: analysis.proposedChanges || [],
    documents: [],
    status: "pending"
  };
}

function renderAiUpdateInbox() {
  if (!aiUpdateInboxSummary || !aiUpdateInboxList || !aiUpdateProposalDetail) {
    return;
  }

  aiUpdateInboxSummary.innerHTML = ["pending", "approved", "rejected"]
    .map(
      (status) => `
        <article class="dashboard-card">
          <p class="dashboard-label">${escapeHtml(status)}</p>
          <p class="dashboard-value">${escapeHtml(String(aiUpdateProposalCounts[status] || 0))}</p>
        </article>
      `
    )
    .join("");

  const pending = allAiUpdateProposals.filter((proposal) => proposal.status === "pending");
  const reviewed = allAiUpdateProposals.filter((proposal) => proposal.status !== "pending");
  const proposals = pending.concat(reviewed).slice(0, 30);

  aiUpdateInboxList.innerHTML = proposals.length
    ? proposals
        .map(
          (proposal) => `
            <article class="update-card ai-update-card ${proposal.id === selectedAiUpdateProposalId ? "is-selected" : ""}">
              <div class="update-head">
                <button class="link-button company-link" type="button" data-action="view-ai-proposal" data-id="${escapeHtml(proposal.id)}">
                  ${escapeHtml(getAiProposalInvestmentName(proposal))}
                </button>
                <span class="status-chip">${escapeHtml(proposal.status)}</span>
              </div>
              <p class="update-meta">
                ${escapeHtml(getAiProposalEntity(proposal))} • Source ${escapeHtml(proposal.sourceDate || "date not set")} • ${escapeHtml(formatConfidence(proposal.confidenceScore))}
              </p>
              <p class="update-meta">
                ${escapeHtml(proposal.sender || "Sender not set")} • ${escapeHtml(proposal.subject || "No subject")}
              </p>
              <p class="update-notes">${escapeHtml(summarizeText(proposal.summary || "No summary staged.", ""))}</p>
              <p class="update-meta">Created ${escapeHtml(formatDisplayDate(proposal.createdAt))}</p>
            </article>
          `
        )
        .join("")
    : '<p class="update-meta">No AI-proposed updates are staged yet.</p>';

  renderAiUpdateProposalDetail();
}

function renderAiUpdateProposalDetail() {
  if (!aiUpdateProposalDetail) {
    return;
  }

  const proposal = allAiUpdateProposals.find((item) => item.id === selectedAiUpdateProposalId);
  if (!proposal) {
    aiUpdateProposalDetail.classList.add("hidden");
    aiUpdateProposalDetail.innerHTML = "";
    return;
  }

  aiUpdateProposalDetail.classList.remove("hidden");
  aiUpdateProposalDetail.innerHTML = `
    <div class="panel-header">
      <div>
        <h3>${escapeHtml(getAiProposalInvestmentName(proposal))}</h3>
        <p class="section-copy">Proposed update staged ${escapeHtml(formatDisplayDate(proposal.createdAt))}</p>
      </div>
      <button class="secondary-button" type="button" data-action="close-ai-proposal">Close</button>
    </div>

    <section class="ai-detail-section">
      <h4>Match</h4>
      <div class="company-summary-grid">
        <article class="company-summary-card"><p class="dashboard-label">Investment</p><p class="highlight-value">${escapeHtml(getAiProposalInvestmentName(proposal))}</p></article>
        <article class="company-summary-card"><p class="dashboard-label">Entity</p><p class="highlight-value">${escapeHtml(getAiProposalEntity(proposal))}</p></article>
        <article class="company-summary-card"><p class="dashboard-label">Confidence</p><p class="highlight-value">${escapeHtml(formatConfidence(proposal.confidenceScore))}</p></article>
      </div>
      <p class="update-meta">${escapeHtml(proposal.matchReason || "No match reason provided.")}</p>
    </section>

    <section class="ai-detail-section">
      <h4>Source</h4>
      <p class="update-meta">${escapeHtml(proposal.sender || "Sender not set")} • ${escapeHtml(proposal.subject || "No subject")}</p>
      <p class="update-meta">${escapeHtml(proposal.sourceType || "Source type not set")} • ${escapeHtml(proposal.sourceIdentifier || "No source identifier")} • ${escapeHtml(proposal.sourceDate || "Source date not set")}</p>
      ${
        Array.isArray(proposal.documents) && proposal.documents.length
          ? `<div class="document-pill-row">${proposal.documents
              .map((document) =>
                document.url
                  ? `<a class="document-pill" href="${escapeHtml(document.url)}" target="_blank" rel="noreferrer">${escapeHtml(document.name || document.url)}</a>`
                  : `<span class="document-pill">${escapeHtml(document.name || document.id || "Attachment")}</span>`
              )
              .join("")}</div>`
          : '<p class="update-meta">No attachment references.</p>'
      }
    </section>

    <section class="ai-detail-section">
      <h4>What Changed</h4>
      <p class="update-notes ai-summary-note">${escapeHtml(proposal.summary || "No summary staged.")}</p>
    </section>

    <section class="ai-detail-section">
      <h4>Extracted Information</h4>
      ${renderExtractedData(proposal.extractedData)}
    </section>

    <section class="ai-detail-section">
      <h4>Proposed Changes</h4>
      ${renderProposedChanges(proposal.proposedChanges)}
    </section>

    <section class="ai-detail-section">
      <h4>Review</h4>
      <p class="update-meta">Status: ${escapeHtml(proposal.status)}${
        proposal.reviewedBy
          ? ` • Reviewed by ${escapeHtml(proposal.reviewedBy)} on ${escapeHtml(formatDisplayDate(proposal.reviewedAt))}`
          : ""
      }</p>
      ${
        proposal.status === "pending" && canEditWorkspace()
          ? `<div class="card-actions">
              <button type="button" data-action="approve-ai-proposal" data-id="${escapeHtml(proposal.id)}">Approve</button>
              <button class="secondary-button danger-button" type="button" data-action="reject-ai-proposal" data-id="${escapeHtml(proposal.id)}">Reject</button>
              <button class="secondary-button" type="button" data-action="edit-ai-proposal-source" data-investment-id="${escapeHtml(proposal.investmentId)}">Edit live record first</button>
            </div>`
          : ""
      }
    </section>
  `;
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

function renderPublicStocks() {
  if (!publicStockSummary || !publicStockList) {
    return;
  }

  const stocks = getPublicStockRows(allInvestments).filter(isPublicStockRow);
  renderPublicStockFilterOptions(stocks);
  const filteredStocks = getFilteredPublicStockRows(stocks);
  const savedFilteredStocks = filteredStocks.filter((investment) => !investment.isWatchlistOnly);
  const totalMarketValue = savedFilteredStocks.reduce(
    (sum, investment) => sum + getStockMetrics(investment).marketValue,
    0
  );
  const totalCostBasis = savedFilteredStocks.reduce(
    (sum, investment) => sum + getStockMetrics(investment).totalCostBasis,
    0
  );
  const totalGainLoss = totalMarketValue - totalCostBasis;
  const totalGainLossPercent = totalCostBasis > 0 ? totalGainLoss / totalCostBasis : null;

  publicStockSummary.innerHTML = [
    { label: "Public equity market value", value: formatMoney(totalMarketValue) },
    { label: "Total cost basis", value: formatMoney(totalCostBasis) },
    { label: "Total unrealized gain/loss", value: formatSignedMoney(totalGainLoss) },
    { label: "Total unrealized gain/loss %", value: formatStockPercent(totalGainLossPercent) },
    { label: "Public equity positions", value: String(savedFilteredStocks.length) }
  ]
    .map(
      (card) => `
        <article class="dashboard-card ${card.label.includes("gain/loss") ? getStockPerformanceClass(totalGainLoss) : ""}">
          <p class="dashboard-label">${escapeHtml(card.label)}</p>
          <p class="dashboard-value">${escapeHtml(card.value)}</p>
        </article>
      `
    )
    .join("");

  if (publicStockSearchFilter && publicStockSearchFilter.value !== publicStockFilters.search) {
    publicStockSearchFilter.value = publicStockFilters.search;
  }

  publicStockList.innerHTML = filteredStocks.length
    ? filteredStocks
        .map((investment) => {
          const metrics = getStockMetrics(investment);
          const gainLossLabel = metrics.hasCostBasis
            ? formatSignedMoney(metrics.gainLoss)
            : "N/A";
          const gainLossPercentLabel = metrics.hasCostBasis
            ? formatStockPercent(metrics.gainLossPercent)
            : "N/A";
          return `
            <article class="update-card public-stock-card">
              <div class="update-head">
                ${
                  investment.isWatchlistOnly
                    ? `<h3>${escapeHtml(investment.company)}</h3>`
                    : `<button class="link-button company-link" type="button" data-company="${escapeHtml(investment.company)}" data-entity="${escapeHtml(investment.entity || "")}">
                        ${escapeHtml(investment.company)}
                      </button>`
                }
                <span class="status-chip">${escapeHtml(getStockTickerLabel(investment))}</span>
              </div>
              <p class="update-meta">
                ${escapeHtml(investment.entity || "Entity not specified")} • ${escapeHtml(investment.assetType || "Stock")} • ${escapeHtml(normalizeStatusName(investment.status) || "Status not set")}
              </p>
              <div class="stock-metric-grid">
                <div class="entity-metric-box">
                  <p class="dashboard-label">Shares</p>
                  <p class="highlight-value">${metrics.shares ? escapeHtml(metrics.shares.toLocaleString()) : "Not set"}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Average cost / share</p>
                  <p class="highlight-value">${metrics.costPerShare ? escapeHtml(formatMoney(metrics.costPerShare)) : "Not set"}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Total cost basis</p>
                  <p class="highlight-value">${metrics.totalCostBasis ? escapeHtml(formatMoney(metrics.totalCostBasis)) : "Not set"}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Current price</p>
                  <p class="highlight-value">${metrics.currentPrice ? escapeHtml(formatMoney(metrics.currentPrice)) : "Not set"}</p>
                  <p class="update-meta">${escapeHtml(investment.marketPriceDate || "No price date")}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Market value</p>
                  <p class="highlight-value">${escapeHtml(formatMoney(metrics.marketValue))}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Unrealized gain/loss $</p>
                  <p class="highlight-value ${escapeHtml(getStockPerformanceClass(metrics.hasCostBasis ? metrics.gainLoss : null))}">${escapeHtml(gainLossLabel)}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Unrealized gain/loss %</p>
                  <p class="highlight-value ${escapeHtml(getStockPerformanceClass(metrics.gainLossPercent))}">${escapeHtml(gainLossPercentLabel)}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Price as of</p>
                  <p class="highlight-value">${escapeHtml(investment.marketPriceDate || "Not set")}</p>
                </div>
              </div>
              <p class="update-notes">${escapeHtml(investment.notes || "No notes provided.")}</p>
              <div class="card-actions">
                ${
                  investment.isWatchlistOnly
                    ? `<button class="secondary-button card-action-button" type="button" data-action="add-stock" data-ticker="${escapeHtml(investment.ticker)}">Add position</button>`
                    : `<button class="secondary-button card-action-button" type="button" data-action="view-company" data-company="${escapeHtml(investment.company)}" data-entity="${escapeHtml(investment.entity || "")}">View company</button>
                      <button class="secondary-button card-action-button" type="button" data-action="edit" data-id="${escapeHtml(investment.id)}">Edit</button>`
                }
              </div>
            </article>
          `;
        })
        .join("")
    : '<p class="update-meta">No public stock positions match those filters.</p>';
}

function renderCashFilterOptions(cashRows) {
  const assignOptions = (element, placeholder, values, selectedValue) => {
    if (!element) {
      return "";
    }
    element.innerHTML = [`<option value="">${placeholder}</option>`]
      .concat(values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`))
      .join("");
    element.value = values.includes(selectedValue) ? selectedValue : "";
    return element.value;
  };

  const entities = Array.from(
    new Set(
      configuredEntities
        .concat(cashRows.map((investment) => normalizeEntityName(investment.entity)).filter(Boolean))
        .map(normalizeEntityName)
    )
  ).sort((left, right) => left.localeCompare(right));
  const institutions = Array.from(
    new Set(cashRows.map(getCashInstitution).filter(Boolean))
  ).sort((left, right) => left.localeCompare(right));

  cashFilters.entity = assignOptions(cashEntityFilter, "All entities", entities, cashFilters.entity);
  cashFilters.institution = assignOptions(
    cashInstitutionFilter,
    "All institutions",
    institutions,
    cashFilters.institution
  );
}

function getFilteredCashRows(cashRows) {
  return cashRows.filter((investment) => {
    const matchesEntity =
      !cashFilters.entity ||
      normalizeEntityName(investment.entity) === normalizeEntityName(cashFilters.entity);
    const matchesInstitution =
      !cashFilters.institution ||
      getCashInstitution(investment) === cashFilters.institution;
    return matchesEntity && matchesInstitution;
  });
}

function renderCash() {
  if (!cashSummary || !cashList) {
    return;
  }

  const cashRows = getCashRows(allInvestments);
  renderCashFilterOptions(cashRows);
  const filteredCashRows = getFilteredCashRows(cashRows);
  const totalBalance = filteredCashRows.reduce((sum, investment) => sum + getCashBalance(investment), 0);
  const entityCount = new Set(filteredCashRows.map((investment) => normalizeEntityName(investment.entity)).filter(Boolean)).size;
  const latestBalanceDate = filteredCashRows
    .map((investment) => investment.valuationDate)
    .filter(Boolean)
    .sort((left, right) => parseDateValue(right, new Date(0)) - parseDateValue(left, new Date(0)))[0] || "N/A";

  cashSummary.innerHTML = [
    { label: "Total cash balance", value: formatMoney(totalBalance) },
    { label: "Cash accounts", value: String(filteredCashRows.length) },
    { label: "Entities with cash", value: String(entityCount) },
    { label: "Latest balance date", value: latestBalanceDate }
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

  cashList.innerHTML = filteredCashRows.length
    ? filteredCashRows
        .map(
          (investment) => `
            <article class="update-card cash-account-card">
              <div class="update-head">
                <button class="link-button company-link" type="button" data-company="${escapeHtml(investment.company)}" data-entity="${escapeHtml(investment.entity || "")}">
                  ${escapeHtml(investment.company || "Unnamed cash account")}
                </button>
                <span class="status-chip">${escapeHtml(getCashAccountType(investment))}</span>
              </div>
              <p class="update-meta">
                ${escapeHtml(getCashInstitution(investment) || "Institution not set")} • ${escapeHtml(normalizeEntityName(investment.entity) || "Entity not specified")} • ${escapeHtml(investment.currency || "USD")}
              </p>
              <div class="stock-metric-grid">
                <div class="entity-metric-box">
                  <p class="dashboard-label">Current balance</p>
                  <p class="highlight-value">${escapeHtml(formatMoney(getCashBalance(investment)))}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Balance date</p>
                  <p class="highlight-value">${escapeHtml(investment.valuationDate || "Not set")}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Institution</p>
                  <p class="highlight-value">${escapeHtml(getCashInstitution(investment) || "Not set")}</p>
                </div>
              </div>
              <p class="update-notes">${escapeHtml(investment.notes || "No notes provided.")}</p>
              ${
                canEditWorkspace()
                  ? `<div class="card-actions">
                      <button class="secondary-button card-action-button" type="button" data-action="edit" data-id="${escapeHtml(investment.id)}">Edit</button>
                    </div>`
                  : ""
              }
            </article>
          `
        )
        .join("")
    : '<p class="update-meta">No cash accounts match those filters.</p>';
}

function renderFixedIncomeFilterOptions(bondRows) {
  const assignOptions = (element, placeholder, values, selectedValue) => {
    if (!element) {
      return "";
    }
    element.innerHTML = [`<option value="">${placeholder}</option>`]
      .concat(values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`))
      .join("");
    element.value = values.includes(selectedValue) ? selectedValue : "";
    return element.value;
  };

  const entities = Array.from(
    new Set(
      configuredEntities
        .concat(bondRows.map((investment) => normalizeEntityName(investment.entity)).filter(Boolean))
        .map(normalizeEntityName)
    )
  ).sort((left, right) => left.localeCompare(right));
  const types = Array.from(new Set(bondRows.map(getBondType).filter(Boolean))).sort((left, right) =>
    left.localeCompare(right)
  );
  const maturityYears = Array.from(new Set(bondRows.map(getBondMaturityYear).filter(Boolean))).sort();

  fixedIncomeFilters.entity = assignOptions(
    fixedIncomeEntityFilter,
    "All entities",
    entities,
    fixedIncomeFilters.entity
  );
  fixedIncomeFilters.type = assignOptions(
    fixedIncomeTypeFilter,
    "All bond types",
    types,
    fixedIncomeFilters.type
  );
  fixedIncomeFilters.maturityYear = assignOptions(
    fixedIncomeMaturityFilter,
    "All maturities",
    maturityYears,
    fixedIncomeFilters.maturityYear
  );
}

function getFilteredFixedIncomeRows(bondRows) {
  const search = String(fixedIncomeFilters.search || "").trim().toLowerCase();
  return bondRows.filter((investment) => {
    const matchesEntity =
      !fixedIncomeFilters.entity ||
      normalizeEntityName(investment.entity) === normalizeEntityName(fixedIncomeFilters.entity);
    const matchesType = !fixedIncomeFilters.type || getBondType(investment) === fixedIncomeFilters.type;
    const matchesMaturity =
      !fixedIncomeFilters.maturityYear ||
      getBondMaturityYear(investment) === fixedIncomeFilters.maturityYear;
    const matchesSearch =
      !search ||
      getBondIssuer(investment).toLowerCase().includes(search) ||
      String(investment.bondCusip || "").toLowerCase().includes(search) ||
      getBondDescription(investment).toLowerCase().includes(search);
    return matchesEntity && matchesType && matchesMaturity && matchesSearch;
  });
}

function renderFixedIncome() {
  if (!fixedIncomeSummary || !fixedIncomeList || !fixedIncomeLadder) {
    return;
  }

  const bondRows = getBondRows(allInvestments);
  renderFixedIncomeFilterOptions(bondRows);
  if (fixedIncomeSearchFilter && fixedIncomeSearchFilter.value !== fixedIncomeFilters.search) {
    fixedIncomeSearchFilter.value = fixedIncomeFilters.search;
  }

  const filteredRows = getFilteredFixedIncomeRows(bondRows);
  const summary = buildFixedIncomeSummary(filteredRows);
  const ladderRows = buildMaturityLadder(filteredRows);

  fixedIncomeSummary.innerHTML = [
    { label: "Fixed income market value", value: formatMoney(summary.marketValue) },
    { label: "Total par value", value: formatMoney(summary.parValue) },
    { label: "Annual coupon income", value: formatMoney(summary.annualCouponIncome) },
    { label: "Weighted average coupon", value: formatPercent(summary.weightedCoupon) },
    { label: "Weighted average YTM", value: formatPercent(summary.weightedYtm) },
    { label: "Bond positions", value: String(summary.positions.length) },
    { label: "Next maturity date", value: summary.nextMaturity || "N/A" }
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

  fixedIncomeLadder.innerHTML = ladderRows.length
    ? `
      <section class="fixed-income-ladder-card">
        <div class="update-head">
          <h3>Maturity ladder</h3>
          <span class="status-chip">Par value</span>
        </div>
        <div class="maturity-ladder-list">
          ${ladderRows
            .map(
              ([year, parValue]) => `
                <div class="maturity-ladder-row">
                  <span>${escapeHtml(year)}</span>
                  <strong>${escapeHtml(formatMoney(parValue))}</strong>
                </div>
              `
            )
            .join("")}
        </div>
      </section>
    `
    : '<p class="update-meta">No maturity dates are available for the selected bonds.</p>';

  fixedIncomeList.innerHTML = filteredRows.length
    ? filteredRows
        .map((investment) => {
          const metrics = getBondMetrics(investment);
          return `
            <article class="update-card fixed-income-card">
              <div class="update-head">
                <button class="link-button company-link" type="button" data-company="${escapeHtml(investment.company)}" data-entity="${escapeHtml(investment.entity || "")}">
                  ${escapeHtml(getBondIssuer(investment) || "Unnamed issuer")}
                </button>
                <span class="status-chip">${escapeHtml(getBondType(investment))}</span>
              </div>
              <p class="update-meta">
                ${escapeHtml(getBondDescription(investment) || "Description not set")} • ${escapeHtml(investment.bondCusip || "CUSIP not set")} • ${escapeHtml(normalizeEntityName(investment.entity) || "Entity not specified")}
              </p>
              <div class="stock-metric-grid">
                <div class="entity-metric-box">
                  <p class="dashboard-label">Par value</p>
                  <p class="highlight-value">${escapeHtml(formatMoney(metrics.parValue))}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Market value</p>
                  <p class="highlight-value">${escapeHtml(formatMoney(metrics.marketValue))}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Coupon</p>
                  <p class="highlight-value">${escapeHtml(formatPercent(metrics.couponRate))}</p>
                  <p class="update-meta">${escapeHtml(investment.bondCouponFrequency || "Frequency not set")}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Annual coupon income</p>
                  <p class="highlight-value">${escapeHtml(formatMoney(metrics.annualCouponIncome))}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Maturity</p>
                  <p class="highlight-value">${escapeHtml(investment.bondMaturityDate || "Not set")}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Call date</p>
                  <p class="highlight-value">${escapeHtml(investment.bondCallDate || "N/A")}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Current price</p>
                  <p class="highlight-value">${metrics.currentPrice ? escapeHtml(metrics.currentPrice.toLocaleString()) : "Not set"}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Yield to maturity</p>
                  <p class="highlight-value">${escapeHtml(formatPercent(metrics.ytm))}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Yield to call</p>
                  <p class="highlight-value">${escapeHtml(formatPercent(metrics.ytc))}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Credit rating</p>
                  <p class="highlight-value">${escapeHtml(investment.bondCreditRating || "Not set")}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Insurer</p>
                  <p class="highlight-value">${escapeHtml(investment.bondInsurer || "N/A")}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Tax status</p>
                  <p class="highlight-value">${escapeHtml(investment.bondTaxStatus || "Not set")}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Price as of</p>
                  <p class="highlight-value">${escapeHtml(investment.bondMarketPriceDate || "Not set")}</p>
                </div>
              </div>
              <p class="update-notes">${escapeHtml(investment.notes || "No notes provided.")}</p>
              ${
                canEditWorkspace()
                  ? `<div class="card-actions">
                      <button class="secondary-button card-action-button" type="button" data-action="edit" data-id="${escapeHtml(investment.id)}">Edit</button>
                    </div>`
                  : ""
              }
            </article>
          `;
        })
        .join("")
    : '<p class="update-meta">No fixed income positions match those filters.</p>';
}

function renderRealEstateFilterOptions(realEstateRows) {
  const assignOptions = (element, placeholder, values, selectedValue) => {
    if (!element) {
      return "";
    }
    element.innerHTML = [`<option value="">${placeholder}</option>`]
      .concat(values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`))
      .join("");
    element.value = values.includes(selectedValue) ? selectedValue : "";
    return element.value;
  };

  const entities = Array.from(
    new Set(
      configuredEntities
        .concat(realEstateRows.map((investment) => normalizeEntityName(investment.entity)).filter(Boolean))
        .map(normalizeEntityName)
    )
  ).sort((left, right) => left.localeCompare(right));
  const types = Array.from(new Set(realEstateRows.map(getRealEstatePropertyType).filter(Boolean))).sort(
    (left, right) => left.localeCompare(right)
  );

  realEstateFilters.entity = assignOptions(
    realEstateEntityFilter,
    "All entities",
    entities,
    realEstateFilters.entity
  );
  realEstateFilters.type = assignOptions(
    realEstateTypeFilter,
    "All property types",
    types,
    realEstateFilters.type
  );
}

function getFilteredRealEstateRows(realEstateRows) {
  const search = String(realEstateFilters.search || "").trim().toLowerCase();
  return realEstateRows.filter((investment) => {
    const matchesEntity =
      !realEstateFilters.entity ||
      normalizeEntityName(investment.entity) === normalizeEntityName(realEstateFilters.entity);
    const matchesType =
      !realEstateFilters.type || getRealEstatePropertyType(investment) === realEstateFilters.type;
    const matchesSearch =
      !search ||
      getRealEstatePropertyName(investment).toLowerCase().includes(search) ||
      getRealEstateAddress(investment).toLowerCase().includes(search);
    return matchesEntity && matchesType && matchesSearch;
  });
}

function renderRealEstate() {
  if (!realEstateSummary || !realEstateList) {
    return;
  }

  const realEstateRows = getRealEstateRows(allInvestments);
  renderRealEstateFilterOptions(realEstateRows);
  if (realEstateSearchFilter && realEstateSearchFilter.value !== realEstateFilters.search) {
    realEstateSearchFilter.value = realEstateFilters.search;
  }

  const filteredRows = getFilteredRealEstateRows(realEstateRows);
  const summary = buildRealEstateSummary(filteredRows);

  realEstateSummary.innerHTML = [
    { label: "Total appraised value", value: formatMoney(summary.appraisedValue) },
    { label: "Total debt", value: formatMoney(summary.debt) },
    { label: "Total net equity", value: formatMoney(summary.netEquity) },
    { label: "Annual NOI", value: formatMoney(summary.noi) },
    { label: "Weighted average cap rate", value: formatPercent(summary.weightedCapRate) },
    { label: "Weighted average LTV", value: formatPercent(summary.weightedLtv) },
    { label: "Properties", value: String(summary.positions.length) },
    { label: "Latest appraisal date", value: summary.latestAppraisalDate || "N/A" }
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

  realEstateList.innerHTML = filteredRows.length
    ? filteredRows
        .map((investment) => {
          const metrics = getRealEstateMetrics(investment);
          const address = getRealEstateAddress(investment);
          const sizeParts = [
            investment.realEstateSquareFootage ? `${investment.realEstateSquareFootage} sf` : "",
            investment.realEstateAcreage ? `${investment.realEstateAcreage} acres` : "",
            investment.realEstateUnits ? `${investment.realEstateUnits} units` : ""
          ].filter(Boolean);
          return `
            <article class="update-card real-estate-card">
              <div class="update-head">
                <button class="link-button company-link" type="button" data-company="${escapeHtml(investment.company)}" data-entity="${escapeHtml(investment.entity || "")}">
                  ${escapeHtml(getRealEstatePropertyName(investment) || "Unnamed property")}
                </button>
                <span class="status-chip">${escapeHtml(getRealEstatePropertyType(investment))}</span>
              </div>
              <p class="update-meta">
                ${escapeHtml(address || "Address not set")} • ${escapeHtml(normalizeEntityName(investment.entity) || "Entity not specified")}
              </p>
              <div class="stock-metric-grid">
                <div class="entity-metric-box">
                  <p class="dashboard-label">Gross appraised value</p>
                  <p class="highlight-value">${escapeHtml(formatMoney(metrics.appraisedValue))}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Appraisal date</p>
                  <p class="highlight-value">${escapeHtml(investment.realEstateAppraisalDate || investment.valuationDate || "Not set")}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Ownership %</p>
                  <p class="highlight-value">${escapeHtml(formatPercent(metrics.ownershipRate))}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Entity share gross</p>
                  <p class="highlight-value">${escapeHtml(formatMoney(metrics.entityValue))}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Entity share debt</p>
                  <p class="highlight-value">${escapeHtml(formatMoney(metrics.entityDebt))}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Net equity</p>
                  <p class="highlight-value">${escapeHtml(formatMoney(metrics.netEquity))}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Debt</p>
                  <p class="highlight-value">${escapeHtml(formatMoney(metrics.debt))}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">LTV</p>
                  <p class="highlight-value">${escapeHtml(formatPercent(metrics.ltv))}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">NOI</p>
                  <p class="highlight-value">${escapeHtml(formatMoney(metrics.noi))}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Cap rate</p>
                  <p class="highlight-value">${escapeHtml(formatPercent(metrics.capRate))}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">NOI margin</p>
                  <p class="highlight-value">${escapeHtml(formatPercent(metrics.noiMargin))}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Occupancy</p>
                  <p class="highlight-value">${escapeHtml(investment.realEstateOccupancy ? `${investment.realEstateOccupancy}%` : "N/A")}</p>
                </div>
                <div class="entity-metric-box">
                  <p class="dashboard-label">Size</p>
                  <p class="highlight-value">${escapeHtml(sizeParts.join(" / ") || "Not set")}</p>
                </div>
              </div>
              ${
                metrics.hasInternalValueOverride
                  ? `<p class="update-meta">Internal value override: ${escapeHtml(formatMoney(metrics.internalValueOverride))}${investment.realEstateInternalValueDate ? ` as of ${escapeHtml(investment.realEstateInternalValueDate)}` : ""}</p>`
                  : ""
              }
              <p class="update-notes">${escapeHtml(investment.notes || "No notes provided.")}</p>
              <div class="card-actions">
                <button class="secondary-button card-action-button" type="button" data-action="view" data-company="${escapeHtml(investment.company)}" data-entity="${escapeHtml(investment.entity || "")}">View Property</button>
                ${
                  canEditWorkspace()
                    ? `
                      <button class="secondary-button card-action-button" type="button" data-action="edit" data-id="${escapeHtml(investment.id)}">Edit</button>
                    `
                    : ""
                }
              </div>
            </article>
          `;
        })
        .join("")
    : '<p class="update-meta">No real estate positions match those filters.</p>';
}

async function refreshPublicStockPrices(options = {}) {
  if (!refreshPublicStockPricesButton) {
    return;
  }

  const automatic = Boolean(options.automatic);
  const force = Boolean(options.force);
  const positions = getFilteredPublicStockRows(getPublicStockRows(allInvestments).filter(isPublicStockRow)).filter(
    (investment) =>
      !investment.isWatchlistOnly &&
      isPublicStockInvestment(investment) &&
      investment.id &&
      investment.ticker &&
      (force || !publicStockQuoteRequestKeys.has(getPublicStockQuoteRequestKey(investment)))
  );
  const positionGroups = groupPublicStockPositionsByTicker(positions);
  const tickerGroups = Array.from(positionGroups.entries());
  if (!tickerGroups.length) {
    if (!automatic && publicStockPriceMessage) {
      publicStockPriceMessage.textContent = "No saved public stock positions with tickers are visible.";
    }
    return;
  }

  if (automaticPublicStockRefreshInFlight && automatic) {
    return;
  }
  automaticPublicStockRefreshInFlight = automatic;
  if (!automatic) {
    refreshPublicStockPricesButton.disabled = true;
    refreshPublicStockPricesButton.textContent = "Refreshing...";
  }
  if (publicStockPriceMessage) {
    publicStockPriceMessage.textContent = automatic
      ? `Refreshing latest public stock prices...`
      : `Refreshing ${tickerGroups.length} ticker${tickerGroups.length === 1 ? "" : "s"} across ${positions.length} position${positions.length === 1 ? "" : "s"}...`;
  }

  const updated = [];
  const checked = [];
  const failed = [];
  try {
    for (let groupIndex = 0; groupIndex < tickerGroups.length; groupIndex += 1) {
      const [ticker, tickerPositions] = tickerGroups[groupIndex];
      publicStockQuoteRequestKeys.add(ticker);
      if (groupIndex > 0) {
        await waitForPublicStockQuoteSlot();
      }
      try {
        const quote = await fetchJson(`/api/stock-quote?ticker=${encodeURIComponent(ticker)}`);
        let tickerUpdated = 0;
        tickerPositions.forEach((investment) => {
          if (publicStockValuationNeedsPatch(investment, quote)) {
            tickerUpdated += 1;
          }
        });
        if (!tickerUpdated) {
          checked.push(`${quote.symbol || ticker}${quote.priceDate ? ` (${quote.priceDate})` : ""}`);
          continue;
        }

        for (const investment of tickerPositions) {
          if (!publicStockValuationNeedsPatch(investment, quote)) {
            continue;
          }
          await fetchJson(`/api/investments/${investment.id}`, {
            method: "PATCH",
            headers: {
              "Content-Type": "application/json"
            },
            body: JSON.stringify(getQuotePatchPayload(investment, quote))
          });
        }
        updated.push(
          `${quote.symbol || ticker}${quote.priceDate ? ` (${quote.priceDate})` : ""} for ${tickerUpdated} position${tickerUpdated === 1 ? "" : "s"}`
        );
      } catch (error) {
        if (error.status === 401) {
          setSignedInState(null);
          throw error;
        }
        failed.push(`${ticker}: Price update temporarily unavailable — previous price retained`);
      }
    }

    if (updated.length) {
      await loadUpdates();
    } else {
      renderPublicStocks();
    }
    if (publicStockPriceMessage) {
      const successMessage = updated.length
        ? `Updated ${updated.length} position${updated.length === 1 ? "" : "s"}: ${updated.join(", ")}.`
        : checked.length
          ? `Prices already current for ${checked.length} position${checked.length === 1 ? "" : "s"}: ${checked.join(", ")}.`
          : "No prices were updated.";
      publicStockPriceMessage.textContent = failed.length
        ? `${successMessage} ${failed.length} quote warning${failed.length === 1 ? "" : "s"}: ${failed.join("; ")}. Last saved prices were kept.`
        : successMessage;
    }
  } finally {
    automaticPublicStockRefreshInFlight = false;
    if (!automatic) {
      refreshPublicStockPricesButton.disabled = false;
      refreshPublicStockPricesButton.textContent = "Refresh prices";
    }
  }
}

function renderAll() {
  companyPerformanceMap = buildCompanyPerformanceMap(allInvestments);
  entityPerformanceMap = buildEntityPerformanceMap(allInvestments);
  renderRoleState();
  renderCompanySuggestions();
  renderAiAnalysisInvestmentOptions();
  renderFilterOptions();
  const filteredInvestments = filterInvestments(allInvestments);
  renderDashboard(allInvestments);
  renderDataQuality();
  renderPublicStocks();
  renderFixedIncome();
  renderRealEstate();
  renderCash();
  renderResearchLibrary(allInvestments);
  renderUpdates(filteredInvestments);
  renderTasks();
  renderAiUpdateInbox();
  renderReconciliation();
  renderXirrAudit();
  renderCompanyPanel();
  renderEntityDetail();
}

function renderDataLoadError(message) {
  const text = message || "Investment data could not be loaded. Refresh or sign in again.";
  const errorCard = `
    <article class="dashboard-card quality-alert-card quality-alert-high">
      <p class="dashboard-label">Investment data did not load</p>
      <p class="update-meta">${escapeHtml(text)}</p>
    </article>
  `;

  if (dashboardCards) {
    dashboardCards.innerHTML = errorCard;
  }
  if (entityPerformanceCards) {
    entityPerformanceCards.innerHTML = errorCard;
  }
  if (updatesList) {
    updatesList.innerHTML = `<p class="update-meta">${escapeHtml(text)}</p>`;
  }
}

function renderDataLoadingState() {
  const loadingCard = `
    <article class="dashboard-card">
      <p class="dashboard-label">Loading</p>
      <p class="update-meta">Loading investment data...</p>
    </article>
  `;

  if (dashboardCards) {
    dashboardCards.innerHTML = loadingCard;
  }
  if (entityPerformanceCards) {
    entityPerformanceCards.innerHTML = loadingCard;
  }
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
  renderConfiguredEntitySelects();

  loginCopy.textContent =
    config.authMode === "individual"
      ? `Use your email and your personal workspace password to sign in. ${config.teamUserCount} team login${config.teamUserCount === 1 ? "" : "s"} configured. TEAM_USERS roles: master-editor, editor, or dashboard-viewer.`
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
    ? `Role: ${config.user && config.user.roleLabel ? config.user.roleLabel : getRoleLabel(config.user && config.user.role)}.`
    : "Your account is view-only. You can review investments, research, and tasks, but editing is disabled.";

  if (investmentsLoaded || !currentUser) {
    renderAll();
  } else if (currentUser) {
    renderDataLoadingState();
  }
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
    if (investmentsLoaded) {
      renderAll();
    }
  } catch (error) {
    if (error.status === 401) {
      setSignedInState(null);
      tasksList.innerHTML = "";
      return;
    }

    if (error.status === 403 && isDashboardViewer()) {
      allTasks = [];
      if (investmentsLoaded) {
        renderAll();
      }
      return;
    }

    taskMessage.textContent = error.message || "Tasks could not load.";
  }
}

async function loadAiUpdateProposals() {
  try {
    const data = await fetchJson("/api/ai-update-proposals");
    allAiUpdateProposals = Array.isArray(data.proposals) ? data.proposals : [];
    aiUpdateProposalCounts = {
      pending: Number((data.counts && data.counts.pending) || 0),
      approved: Number((data.counts && data.counts.approved) || 0),
      rejected: Number((data.counts && data.counts.rejected) || 0)
    };
    if (
      selectedAiUpdateProposalId &&
      !allAiUpdateProposals.some((proposal) => proposal.id === selectedAiUpdateProposalId)
    ) {
      selectedAiUpdateProposalId = "";
    }
    renderAiUpdateInbox();
  } catch (error) {
    if (error.status === 401) {
      setSignedInState(null);
      allAiUpdateProposals = [];
      renderAiUpdateInbox();
      return;
    }

    if (error.status === 403 && isDashboardViewer()) {
      allAiUpdateProposals = [];
      renderAiUpdateInbox();
      return;
    }

    if (aiUpdateInboxMessage) {
      aiUpdateInboxMessage.textContent = error.message || "AI Update Inbox could not load.";
    }
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
    reader.onerror = () => reject(new Error("Could not read the selected file."));
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

["dragenter", "dragover"].forEach((eventName) => {
  addListener(reportUpdateFileDropZone, eventName, (event) => {
    event.preventDefault();
    reportUpdateFileDropZone.classList.add("deck-drop-zone-active");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  addListener(reportUpdateFileDropZone, eventName, (event) => {
    event.preventDefault();
    reportUpdateFileDropZone.classList.remove("deck-drop-zone-active");
  });
});

addListener(reportUpdateFileDropZone, "drop", async (event) => {
  const files = event.dataTransfer && event.dataTransfer.files;
  const file = files && files[0] ? files[0] : null;
  await summarizeReportUpdateFile(file);
});

addListener(reportUpdateFileInput, "change", async () => {
  const file = reportUpdateFileInput.files && reportUpdateFileInput.files[0]
    ? reportUpdateFileInput.files[0]
    : null;
  await summarizeReportUpdateFile(file);
  if (reportUpdateFileInput) {
    reportUpdateFileInput.value = "";
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

  resetEntityCardFilterForMenuView(target.dataset.view);
  showWorkspaceView(target.dataset.view);
  window.scrollTo({ top: 0, behavior: "smooth" });
  workspaceMenu.classList.add("hidden");
});

async function loadUpdates() {
  try {
    const data = await fetchJson("/api/investments");
    allInvestments = Array.isArray(data.investments) ? data.investments : [];
    allCompanies = data.companies || [];
    setSignedInState(data.user || currentUser);
    investmentsLoaded = true;
    investmentsLoadError = "";
    renderAll();
    maybeAutoRefreshPublicStockPrices();
  } catch (error) {
    if (error.status === 401) {
      setSignedInState(null);
      investmentsLoaded = false;
      investmentsLoadError = "";
      updatesList.innerHTML = "";
      return;
    }

    investmentsLoaded = false;
    investmentsLoadError = error.message || "Investment data could not be loaded.";
    renderDataLoadError(investmentsLoadError);
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
    await loadConfig();
    await loadUpdates();
    if (investmentsLoaded) {
      await loadTasks();
      await loadAiUpdateProposals();
    }
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

  syncCashFormFields();
  syncBondFormFields();
  syncRealEstateFormFields();
  const formData = new FormData(form);
  const recipients = String(formData.get("recipients") || "")
    .split(",")
    .map((email) => email.trim())
    .filter(Boolean);

  const payload = applyAssetValuationSync({
    company: formData.get("company"),
    entity: formData.get("entity"),
    assetType: formData.get("assetType"),
    ticker: formData.get("ticker"),
    exchange: formData.get("exchange"),
    shareCount: formData.get("shareCount"),
    costBasisPerShare: formData.get("costBasisPerShare"),
    marketPrice: formData.get("marketPrice"),
    marketPriceDate: formData.get("marketPriceDate"),
    bondIssuer: formData.get("bondIssuer"),
    bondDescription: formData.get("bondDescription"),
    bondType: formData.get("bondType"),
    bondCusip: formData.get("bondCusip"),
    bondEntityOwner: formData.get("bondEntityOwner"),
    bondParValue: formData.get("bondParValue"),
    bondPurchasePrice: formData.get("bondPurchasePrice"),
    bondPurchaseDate: formData.get("bondPurchaseDate"),
    bondCostBasis: formData.get("bondCostBasis"),
    bondCouponRate: formData.get("bondCouponRate"),
    bondCouponFrequency: formData.get("bondCouponFrequency"),
    bondMaturityDate: formData.get("bondMaturityDate"),
    bondCallDate: formData.get("bondCallDate"),
    bondCallPrice: formData.get("bondCallPrice"),
    bondCurrentPrice: formData.get("bondCurrentPrice"),
    bondMarketPriceDate: formData.get("bondMarketPriceDate"),
    bondMarketValue: formData.get("bondMarketValue"),
    bondYieldToMaturity: formData.get("bondYieldToMaturity"),
    bondYieldToCall: formData.get("bondYieldToCall"),
    bondCurrentYield: formData.get("bondCurrentYield"),
    bondCreditRating: formData.get("bondCreditRating"),
    bondInsurer: formData.get("bondInsurer"),
    bondTaxStatus: formData.get("bondTaxStatus"),
    bondAccruedInterest: formData.get("bondAccruedInterest"),
    realEstatePropertyName: formData.get("realEstatePropertyName"),
    realEstateAddress: formData.get("realEstateAddress"),
    realEstateCity: formData.get("realEstateCity"),
    realEstateState: formData.get("realEstateState"),
    realEstateZip: formData.get("realEstateZip"),
    realEstateEntityOwner: formData.get("realEstateEntityOwner"),
    realEstatePropertyType: formData.get("realEstatePropertyType"),
    realEstateOwnershipPercent: formData.get("realEstateOwnershipPercent"),
    realEstateOwnershipNotes: formData.get("realEstateOwnershipNotes"),
    realEstateAcquisitionDate: formData.get("realEstateAcquisitionDate"),
    realEstatePurchasePrice: formData.get("realEstatePurchasePrice"),
    realEstateCostBasis: formData.get("realEstateCostBasis"),
    realEstateAppraisedValue: formData.get("realEstateAppraisedValue"),
    realEstateAppraisalDate: formData.get("realEstateAppraisalDate"),
    realEstateAppraiser: formData.get("realEstateAppraiser"),
    realEstateAppraisalDocument: formData.get("realEstateAppraisalDocument"),
    realEstateInternalValueOverride: formData.get("realEstateInternalValueOverride"),
    realEstateInternalValueDate: formData.get("realEstateInternalValueDate"),
    realEstateDebt: formData.get("realEstateDebt"),
    realEstateLoanLender: formData.get("realEstateLoanLender"),
    realEstateDebtInterestRate: formData.get("realEstateDebtInterestRate"),
    realEstateDebtMaturityDate: formData.get("realEstateDebtMaturityDate"),
    realEstateDebtService: formData.get("realEstateDebtService"),
    realEstateLoanNotes: formData.get("realEstateLoanNotes"),
    realEstateNoi: formData.get("realEstateNoi"),
    realEstateRevenue: formData.get("realEstateRevenue"),
    realEstateCapRate: formData.get("realEstateCapRate"),
    realEstateOccupancy: formData.get("realEstateOccupancy"),
    realEstateSquareFootage: formData.get("realEstateSquareFootage"),
    realEstateAcreage: formData.get("realEstateAcreage"),
    realEstateUnits: formData.get("realEstateUnits"),
    realEstatePropertyTaxes: formData.get("realEstatePropertyTaxes"),
    realEstateInsurance: formData.get("realEstateInsurance"),
    realEstateOtherExpenses: formData.get("realEstateOtherExpenses"),
    realEstateOperatingNotes: formData.get("realEstateOperatingNotes"),
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
    reportingCadence: formData.get("reportingCadence"),
    updateRequestStatus: formData.get("updateRequestStatus"),
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
  });

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
        reportUpdateMessage.textContent = formatReportSummaryError(error);
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
        reportUpdateMessage.textContent = formatReportSummaryError(error);
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
        reportUpdateMessage.textContent = formatReportSummaryError(error);
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
        reportUpdateMessage.textContent = formatReportSummaryError(error);
      }
    } finally {
      generateAllReportSummaryButton.disabled = false;
    }
  });
}

if (generateInvestmentHistorySummaryButton) {
  generateInvestmentHistorySummaryButton.addEventListener("click", async () => {
    const companyRecord = findCompanyRecord(selectedCompany, selectedCompanyEntity);
    const rows = companyRecord ? normalizeReportUpdateRows(companyRecord.reportUpdates) : [];
    if (!rows.length) {
      if (reportUpdateMessage) {
        reportUpdateMessage.textContent =
          "Add at least one update before generating an investment history summary.";
      }
      return;
    }

    generateInvestmentHistorySummaryButton.disabled = true;

    try {
      const result = await runReportAnalystPrompt("history-summary", rows, companyRecord, {
        loadingMessage: "Building investment history summary..."
      });
      renderReportInsight(
        "Investment history summary",
        result.answer,
        result.company,
        result.entity
      );
      if (reportUpdateMessage) {
        reportUpdateMessage.textContent = "Investment history summary is ready.";
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
        reportUpdateMessage.textContent = formatReportSummaryError(error);
      }
    } finally {
      generateInvestmentHistorySummaryButton.disabled = false;
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
  if (amountField) {
    updateStockValuePreview();
  }
});

addListener(capitalActivityList, "blur", (event) => {
  const amountField = event.target.closest('[data-capital-field="amount"]');
  if (amountField) {
    amountField.value = normalizeMoneyString(amountField.value);
  }
}, true);

addListener(capitalActivityList, "change", (event) => {
  const amountField = event.target.closest('[data-capital-field="amount"]');
  if (amountField) {
    amountField.value = normalizeMoneyString(amountField.value);
  }
});

attachFormattedInputHandlers();
applyFormInputFormatting();
updateStockDetailsVisibility();

["shareCount", "costBasisPerShare", "marketPrice", "marketPriceDate", "amount"].forEach((fieldName) => {
  if (form && form.elements && form.elements[fieldName]) {
    form.elements[fieldName].addEventListener("input", updateStockValuePreview);
    form.elements[fieldName].addEventListener("change", updateStockValuePreview);
  }
});

[cashAccountNameField, cashInstitutionField, cashAccountTypeField, cashBalanceField, cashBalanceDateField].forEach(
  (field) => {
    if (field) {
      field.addEventListener("input", syncCashFormFields);
      field.addEventListener("change", syncCashFormFields);
    }
  }
);

if (cashBalanceField) {
  cashBalanceField.addEventListener("blur", () => {
    cashBalanceField.value = normalizeMoneyString(cashBalanceField.value);
    syncCashFormFields();
  });
}

[
  bondIssuerField,
  bondDescriptionField,
  bondTypeField,
  bondEntityOwnerField,
  bondParValueField,
  bondPurchasePriceField,
  bondPurchaseDateField,
  bondCostBasisField,
  bondCouponRateField,
  bondCouponFrequencyField,
  bondMaturityDateField,
  bondCallDateField,
  bondCallPriceField,
  bondCurrentPriceField,
  bondMarketPriceDateField,
  bondYieldToMaturityField,
  bondYieldToCallField,
  bondCreditRatingField,
  bondInsurerField,
  bondTaxStatusField,
  bondAccruedInterestField
].forEach((field) => {
  if (field) {
    field.addEventListener("input", () => {
      syncBondFormFields();
      updateBondValuePreview();
    });
    field.addEventListener("change", () => {
      syncBondFormFields();
      updateBondValuePreview();
    });
  }
});

[
  bondParValueField,
  bondPurchasePriceField,
  bondCostBasisField,
  bondCallPriceField,
  bondCurrentPriceField,
  bondAccruedInterestField
].forEach((field) => {
  if (field) {
    field.addEventListener("blur", () => {
      field.value = normalizeMoneyString(field.value, moneyFieldDecimalPlaces[field.name] || 2);
      syncBondFormFields();
      updateBondValuePreview();
    });
  }
});

[
  realEstatePropertyNameField,
  realEstateAddressField,
  realEstateCityField,
  realEstateStateField,
  realEstateZipField,
  realEstateEntityOwnerField,
  realEstatePropertyTypeField,
  realEstateOwnershipPercentField,
  realEstateOwnershipNotesField,
  realEstateAcquisitionDateField,
  realEstatePurchasePriceField,
  realEstateCostBasisField,
  realEstateAppraisedValueField,
  realEstateAppraisalDateField,
  realEstateAppraiserField,
  realEstateAppraisalDocumentField,
  realEstateInternalValueOverrideField,
  realEstateInternalValueDateField,
  realEstateDebtField,
  realEstateLoanLenderField,
  realEstateDebtInterestRateField,
  realEstateDebtMaturityDateField,
  realEstateDebtServiceField,
  realEstateLoanNotesField,
  realEstateNoiField,
  realEstateRevenueField,
  realEstateOccupancyField,
  realEstateSquareFootageField,
  realEstateAcreageField,
  realEstateUnitsField,
  realEstatePropertyTaxesField,
  realEstateInsuranceField,
  realEstateOtherExpensesField,
  realEstateOperatingNotesField
].forEach((field) => {
  if (field) {
    field.addEventListener("input", () => {
      syncRealEstateFormFields();
      updateRealEstateValuePreview();
    });
    field.addEventListener("change", () => {
      syncRealEstateFormFields();
      updateRealEstateValuePreview();
    });
  }
});

[
  realEstatePurchasePriceField,
  realEstateCostBasisField,
  realEstateAppraisedValueField,
  realEstateInternalValueOverrideField,
  realEstateDebtField,
  realEstateDebtServiceField,
  realEstateNoiField,
  realEstateRevenueField,
  realEstateSquareFootageField,
  realEstateAcreageField,
  realEstatePropertyTaxesField,
  realEstateInsuranceField,
  realEstateOtherExpensesField
].forEach((field) => {
  if (field) {
    field.addEventListener("blur", () => {
      field.value = normalizeMoneyString(field.value);
      syncRealEstateFormFields();
      updateRealEstateValuePreview();
    });
  }
});

if (form && form.elements && form.elements.assetType) {
  form.elements.assetType.addEventListener("change", () => {
    updateStockDetailsVisibility({ clearHiddenFields: true });
  });
}

addListener(fetchStockQuoteButton, "click", () => {
  fetchStockQuoteForForm();
});

addListener(addPublicSquareStockButton, "click", () => {
  prefillStockPosition({
    company: "PublicSquare",
    ticker: "PSQH",
    exchange: "NYSE",
    notes: "PublicSquare public stock position."
  });
});

addListener(addSpaceXStockButton, "click", () => {
  prefillStockPosition({
    company: "SpaceX",
    ticker: "SPCX",
    exchange: "NASDAQ",
    notes: "SpaceX public stock position."
  });
});

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

addListener(requestLatestUpdateButton, "click", () => {
  openUpdateRequestModal();
});

addListener(updateRequestMaterialsList, "change", () => {
  refreshUpdateRequestBody();
});

addListener(closeUpdateRequestModalButton, "click", () => {
  closeUpdateRequestModal();
});

addListener(cancelUpdateRequestButton, "click", () => {
  closeUpdateRequestModal();
});

addListener(updateRequestModal, "click", (event) => {
  if (event.target === updateRequestModal) {
    closeUpdateRequestModal();
  }
});

addListener(updateRequestForm, "submit", async (event) => {
  event.preventDefault();
  if (!activeUpdateRequestInvestment) {
    updateRequestMessage.textContent = "Open an investment detail page first.";
    return;
  }

  const recipient = String(updateRequestRecipientField.value || "").trim();
  if (!recipient) {
    updateRequestMessage.textContent = "Add a primary contact email before sending.";
    return;
  }

  sendUpdateRequestButton.disabled = true;
  updateRequestMessage.textContent = "Sending update request...";

  try {
    const result = await fetchJson(
      `/api/investments/${activeUpdateRequestInvestment.id}/update-request`,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          recipient,
          subject: updateRequestSubjectField.value,
          body: updateRequestBodyField.value,
          materialsRequested: getSelectedUpdateRequestMaterials()
        })
      }
    );

    updateRequestMessage.textContent = result.message || "Update request email sent.";
    await loadUpdates();
    closeUpdateRequestModal();
    if (reportUpdateMessage) {
      reportUpdateMessage.textContent = "Update request email sent and logged to the timeline.";
    }
  } catch (error) {
    if (error.status === 401) {
      setSignedInState(null);
      updateRequestMessage.textContent = "Your session expired. Please sign in again.";
      return;
    }

    updateRequestMessage.textContent = error.message;
  } finally {
    sendUpdateRequestButton.disabled = false;
  }
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

addListener(xirrAuditEntitySelect, "change", () => {
  selectedXirrAuditEntity = xirrAuditEntitySelect.value;
  renderXirrAudit();
});

addListener(entityPerformanceCards, "click", (event) => {
  if (isDashboardViewer()) {
    return;
  }

  const holdingsCard = event.target.closest("[data-holdings-action]");
  if (holdingsCard) {
    event.preventDefault();
    event.stopPropagation();
    openAssetViewForEntity(holdingsCard.dataset.holdingsAction || "", holdingsCard.dataset.entity || "");
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

addListener(entityPerformanceCards, "keydown", (event) => {
  if (isDashboardViewer() || !["Enter", " "].includes(event.key)) {
    return;
  }

  const holdingsCard = event.target.closest("[data-holdings-action]");
  if (!holdingsCard) {
    return;
  }

  event.preventDefault();
  event.stopPropagation();
  openAssetViewForEntity(holdingsCard.dataset.holdingsAction || "", holdingsCard.dataset.entity || "");
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

  if (action === "public-stocks") {
    setAssetViewEntityFilter("public-stocks", "", "");
    showWorkspaceView("public-stocks");
    window.scrollTo({ top: 0, behavior: "smooth" });
    return;
  }

  if (action === "fixed-income") {
    setAssetViewEntityFilter("fixed-income", "", "");
    showWorkspaceView("fixed-income");
    window.scrollTo({ top: 0, behavior: "smooth" });
    return;
  }

  if (action === "real-estate") {
    setAssetViewEntityFilter("real-estate", "", "");
    showWorkspaceView("real-estate");
    window.scrollTo({ top: 0, behavior: "smooth" });
    return;
  }

  if (action === "cash") {
    setAssetViewEntityFilter("cash", "", "");
    showWorkspaceView("cash");
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
        reportUpdateMessage.textContent = formatReportSummaryError(error);
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
  const ticker = target.dataset.ticker || "";

  if (action === "add-stock" && ticker) {
    const preset = PUBLIC_STOCK_PRESETS.find(
      (stock) => String(stock.ticker || "").toUpperCase() === String(ticker).toUpperCase()
    );
    if (preset) {
      prefillStockPosition(preset);
    }
    return;
  }

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

addListener(publicStockList, "click", (event) => {
  const target = event.target.closest("[data-action], [data-company]");
  if (!target) {
    return;
  }

  const company = target.dataset.company || "";
  const entity = target.dataset.entity || "";
  const action = target.dataset.action || "";
  const investmentId = target.dataset.id || "";
  const ticker = target.dataset.ticker || "";

  if (action === "add-stock" && ticker) {
    const preset = PUBLIC_STOCK_PRESETS.find(
      (stock) => String(stock.ticker || "").toUpperCase() === String(ticker).toUpperCase()
    );
    if (preset) {
      prefillStockPosition(preset);
    }
    return;
  }

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
  }
});

addListener(publicStockEntityFilter, "change", () => {
  publicStockFilters.entity = publicStockEntityFilter.value;
  assetViewEntityFilterSource["public-stocks"] = publicStockFilters.entity ? "manual" : "";
  renderPublicStocks();
});

addListener(publicStockSearchFilter, "input", () => {
  publicStockFilters.search = publicStockSearchFilter.value;
  renderPublicStocks();
});

addListener(refreshPublicStockPricesButton, "click", () => {
  refreshPublicStockPrices({ force: true });
});

addListener(cashEntityFilter, "change", () => {
  cashFilters.entity = cashEntityFilter.value;
  assetViewEntityFilterSource.cash = cashFilters.entity ? "manual" : "";
  renderCash();
});

addListener(cashInstitutionFilter, "change", () => {
  cashFilters.institution = cashInstitutionFilter.value;
  renderCash();
});

addListener(fixedIncomeEntityFilter, "change", () => {
  fixedIncomeFilters.entity = fixedIncomeEntityFilter.value;
  assetViewEntityFilterSource["fixed-income"] = fixedIncomeFilters.entity ? "manual" : "";
  renderFixedIncome();
});

addListener(fixedIncomeTypeFilter, "change", () => {
  fixedIncomeFilters.type = fixedIncomeTypeFilter.value;
  renderFixedIncome();
});

addListener(fixedIncomeSearchFilter, "input", () => {
  fixedIncomeFilters.search = fixedIncomeSearchFilter.value;
  renderFixedIncome();
});

addListener(fixedIncomeMaturityFilter, "change", () => {
  fixedIncomeFilters.maturityYear = fixedIncomeMaturityFilter.value;
  renderFixedIncome();
});

addListener(realEstateEntityFilter, "change", () => {
  realEstateFilters.entity = realEstateEntityFilter.value;
  assetViewEntityFilterSource["real-estate"] = realEstateFilters.entity ? "manual" : "";
  renderRealEstate();
});

addListener(realEstateTypeFilter, "change", () => {
  realEstateFilters.type = realEstateTypeFilter.value;
  renderRealEstate();
});

addListener(realEstateSearchFilter, "input", () => {
  realEstateFilters.search = realEstateSearchFilter.value;
  renderRealEstate();
});

addListener(cashList, "click", (event) => {
  const target = event.target.closest("[data-action], [data-company]");
  if (!target) {
    return;
  }

  const action = target.dataset.action || "";
  const investmentId = target.dataset.id || "";
  const company = target.dataset.company || "";
  const entity = target.dataset.entity || "";

  if (action === "edit" && investmentId) {
    beginEditInvestment(investmentId);
    return;
  }

  if (company && canOpenCompanyDetails()) {
    selectedCompany = company;
    selectedCompanyEntity = entity;
    renderCompanyPanel();
    showWorkspaceView("portfolio");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }
});

addListener(realEstateList, "click", (event) => {
  const target = event.target.closest("[data-action], [data-company]");
  if (!target) {
    return;
  }

  const action = target.dataset.action || "";
  const investmentId = target.dataset.id || "";
  const company = target.dataset.company || "";
  const entity = target.dataset.entity || "";

  if (action === "edit" && investmentId) {
    beginEditInvestment(investmentId);
    return;
  }

  if (company && canOpenCompanyDetails()) {
    selectedCompany = company;
    selectedCompanyEntity = entity;
    renderCompanyPanel();
    showWorkspaceView("portfolio");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }
});

addListener(fixedIncomeList, "click", (event) => {
  const target = event.target.closest("[data-action], [data-company]");
  if (!target) {
    return;
  }

  const action = target.dataset.action || "";
  const investmentId = target.dataset.id || "";
  const company = target.dataset.company || "";
  const entity = target.dataset.entity || "";

  if (action === "edit" && investmentId) {
    beginEditInvestment(investmentId);
    return;
  }

  if (company && canOpenCompanyDetails()) {
    selectedCompany = company;
    selectedCompanyEntity = entity;
    renderCompanyPanel();
    showWorkspaceView("portfolio");
    window.scrollTo({ top: 0, behavior: "smooth" });
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

  if (
    editableTarget.dataset.moneyInput === "true" &&
    (event.type === "change" || event.type === "blur")
  ) {
    editableTarget.value = normalizeMoneyString(editableTarget.value);
  }

  const investmentId = editableTarget.dataset.id || "";
  if (investmentId) {
    setReconciliationRowDirty(investmentId, true);
  }
}

addListener(reconciliationList, "input", handleReconciliationFieldEdit);
addListener(reconciliationList, "change", handleReconciliationFieldEdit);
addListener(reconciliationList, "blur", handleReconciliationFieldEdit, true);

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

addListener(openAiUpdateAnalyzerButton, "click", () => {
  openAiUpdateAnalyzer();
});

addListener(cancelAiUpdateAnalysisButton, "click", () => {
  closeAiUpdateAnalyzer();
});

addListener(aiAnalysisInvestmentField, "change", () => {
  renderAiAnalysisReview();
});

addListener(aiAnalysisEntityField, "change", () => {
  renderAiAnalysisReview();
});

addListener(aiUpdateAnalysisForm, "submit", async (event) => {
  event.preventDefault();
  if (!aiUpdateAnalysisForm) {
    return;
  }

  const formData = new FormData(aiUpdateAnalysisForm);
  const payload = {
    sourceType: formData.get("sourceType"),
    sender: formData.get("sender"),
    subject: formData.get("subject"),
    sourceDate: formData.get("sourceDate"),
    sourceText: formData.get("sourceText"),
    investmentId: formData.get("investmentId"),
    entityId: formData.get("entityId")
  };

  if (aiUpdateInboxMessage) {
    aiUpdateInboxMessage.textContent = "Analyzing source material...";
  }
  if (runAiUpdateAnalysisButton) {
    runAiUpdateAnalysisButton.disabled = true;
  }

  try {
    latestAiUpdateAnalysis = await fetchJson("/api/ai-update-proposals/analyze", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload)
    });
    renderAiAnalysisReview();
    if (aiUpdateInboxMessage) {
      aiUpdateInboxMessage.textContent = "Analysis ready. Review it before creating a proposal.";
    }
    if (aiUpdateAnalysisReview) {
      aiUpdateAnalysisReview.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  } catch (error) {
    if (error.status === 401) {
      setSignedInState(null);
      return;
    }
    if (aiUpdateInboxMessage) {
      aiUpdateInboxMessage.textContent = error.message;
    }
  } finally {
    if (runAiUpdateAnalysisButton) {
      runAiUpdateAnalysisButton.disabled = false;
    }
  }
});

addListener(aiUpdateAnalysisReview, "click", async (event) => {
  const target = event.target.closest("[data-action]");
  if (!target) {
    return;
  }

  const action = target.dataset.action || "";
  if (action === "focus-ai-analysis-investment") {
    if (aiAnalysisInvestmentField) {
      aiAnalysisInvestmentField.focus();
    }
    return;
  }

  if (action === "focus-ai-analysis-entity") {
    if (aiAnalysisEntityField) {
      aiAnalysisEntityField.focus();
    }
    return;
  }

  if (action === "cancel-ai-analysis") {
    closeAiUpdateAnalyzer();
    return;
  }

  if (action !== "create-ai-analysis-proposal") {
    return;
  }

  target.disabled = true;
  if (aiUpdateInboxMessage) {
    aiUpdateInboxMessage.textContent = "Creating pending AI update proposal...";
  }

  try {
    const proposalPayload = buildProposalPayloadFromAnalysis();
    const result = await fetchJson("/api/ai-update-proposals", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(proposalPayload)
    });
    await loadAiUpdateProposals();
    selectedAiUpdateProposalId = result.proposal ? result.proposal.id : "";
    closeAiUpdateAnalyzer();
    renderAiUpdateInbox();
    if (aiUpdateInboxMessage) {
      aiUpdateInboxMessage.textContent = "Pending AI update proposal created.";
    }
  } catch (error) {
    if (error.status === 401) {
      setSignedInState(null);
      return;
    }
    if (aiUpdateInboxMessage) {
      aiUpdateInboxMessage.textContent = error.message;
    }
  } finally {
    target.disabled = false;
  }
});

addListener(aiUpdateInboxList, "click", (event) => {
  const target = event.target.closest("[data-action='view-ai-proposal']");
  if (!target) {
    return;
  }

  selectedAiUpdateProposalId = target.dataset.id || "";
  renderAiUpdateInbox();
  if (aiUpdateProposalDetail) {
    aiUpdateProposalDetail.scrollIntoView({ behavior: "smooth", block: "start" });
  }
});

addListener(aiUpdateProposalDetail, "click", async (event) => {
  const target = event.target.closest("[data-action]");
  if (!target) {
    return;
  }

  const action = target.dataset.action || "";
  const proposalId = target.dataset.id || selectedAiUpdateProposalId;

  if (action === "close-ai-proposal") {
    selectedAiUpdateProposalId = "";
    renderAiUpdateInbox();
    return;
  }

  if (action === "edit-ai-proposal-source") {
    const investmentId = target.dataset.investmentId || "";
    if (investmentId) {
      beginEditInvestment(investmentId);
    } else if (aiUpdateInboxMessage) {
      aiUpdateInboxMessage.textContent = "No matched live investment is attached to edit.";
    }
    return;
  }

  if (!["approve-ai-proposal", "reject-ai-proposal"].includes(action) || !proposalId) {
    return;
  }

  const endpointAction = action === "approve-ai-proposal" ? "approve" : "reject";
  target.disabled = true;
  if (aiUpdateInboxMessage) {
    aiUpdateInboxMessage.textContent =
      endpointAction === "approve" ? "Approving staged update..." : "Rejecting staged update...";
  }

  try {
    const result = await fetchJson(`/api/ai-update-proposals/${proposalId}/${endpointAction}`, {
      method: "POST"
    });
    await loadAiUpdateProposals();
    selectedAiUpdateProposalId = result.proposal ? result.proposal.id : selectedAiUpdateProposalId;
    renderAiUpdateInbox();
    if (aiUpdateInboxMessage) {
      aiUpdateInboxMessage.textContent =
        endpointAction === "approve"
          ? (result.applyResult && result.applyResult.message) || "AI update approved."
          : "AI update rejected.";
    }
  } catch (error) {
    if (error.status === 401) {
      setSignedInState(null);
      return;
    }
    if (aiUpdateInboxMessage) {
      aiUpdateInboxMessage.textContent = error.message;
    }
  } finally {
    target.disabled = false;
  }
});

(async function initializeApp() {
  await loadConfig();
  await loadUpdates();
  if (investmentsLoaded) {
    await loadTasks();
    await loadAiUpdateProposals();
  }
})().catch((error) => {
  if (error.status === 401) {
    setSignedInState(null);
    return;
  }

  loginMessage.textContent = error.message;
});

renderUploadedDocuments();
renderCapitalActivityRows([]);
