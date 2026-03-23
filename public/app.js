const loginPanel = document.getElementById("loginPanel");
const appPanel = document.getElementById("appPanel");
const loginForm = document.getElementById("loginForm");
const loginMessage = document.getElementById("loginMessage");
const loginCopy = document.getElementById("loginCopy");
const loginButton = document.getElementById("loginButton");
const logoutButton = document.getElementById("logoutButton");
const form = document.getElementById("investmentForm");
const formMessage = document.getElementById("formMessage");
const updatesList = document.getElementById("updatesList");
const authStatus = document.getElementById("authStatus");
const emailStatus = document.getElementById("emailStatus");
const recipientStatus = document.getElementById("recipientStatus");
const refreshButton = document.getElementById("refreshButton");
const submitButton = document.getElementById("submitButton");
const downloadCsvButton = document.getElementById("downloadCsvButton");
const downloadExcelButton = document.getElementById("downloadExcelButton");
const cancelEditButton = document.getElementById("cancelEditButton");
const editingInvestmentId = document.getElementById("editingInvestmentId");
const notesField = document.getElementById("notesField");
const deckFile = document.getElementById("deckFile");
const deckMessage = document.getElementById("deckMessage");
const summarizeDeckButton = document.getElementById("summarizeDeckButton");
const deckDropZone = document.getElementById("deckDropZone");
const deckFileName = document.getElementById("deckFileName");
const emailSummaryInput = document.getElementById("emailSummaryInput");
const summarizeEmailButton = document.getElementById("summarizeEmailButton");
const emailMessage = document.getElementById("emailMessage");
const dashboardCards = document.getElementById("dashboardCards");
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
const companyNextSteps = document.getElementById("companyNextSteps");
const companyTimeline = document.getElementById("companyTimeline");
const closeCompanyPanelButton = document.getElementById("closeCompanyPanelButton");

let currentUser = null;
let allInvestments = [];
let selectedCompany = "";

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
    ? `Signed in as ${user.email}`
    : "Please sign in to view updates";
}

function selectedDeckFile() {
  return deckFile.files && deckFile.files[0] ? deckFile.files[0] : null;
}

function updateDeckFileLabel(file) {
  deckFileName.textContent = file
    ? `Selected: ${file.name}`
    : "No file selected yet";
}

function currentFilters() {
  return {
    search: searchFilter.value.trim().toLowerCase(),
    status: statusFilter.value,
    stage: stageFilter.value,
    owner: ownerFilter.value
  };
}

function filterInvestments(investments) {
  const filters = currentFilters();
  return investments.filter((investment) => {
    const searchHaystack = [
      investment.company,
      investment.notes,
      investment.owner,
      investment.nextStep,
      investment.submittedBy
    ]
      .join(" ")
      .toLowerCase();

    const matchesSearch = !filters.search || searchHaystack.includes(filters.search);
    const matchesStatus = !filters.status || investment.status === filters.status;
    const matchesStage = !filters.stage || investment.stage === filters.stage;
    const matchesOwner = !filters.owner || investment.owner === filters.owner;

    return matchesSearch && matchesStatus && matchesStage && matchesOwner;
  });
}

function toNumber(value) {
  const amount = Number(value);
  return Number.isFinite(amount) ? amount : 0;
}

function formatMoney(value) {
  return `$${toNumber(value).toLocaleString()}`;
}

function buildDashboardCards(investments) {
  const uniqueCompanies = new Set(investments.map((investment) => investment.company).filter(Boolean));
  const openCount = investments.filter(
    (investment) => !["Passed", "Closed"].includes(investment.status)
  ).length;
  const approvedCount = investments.filter((investment) => investment.status === "Approved").length;
  const totalAmount = investments.reduce((sum, investment) => sum + toNumber(investment.amount), 0);

  return [
    { label: "Updates", value: String(investments.length) },
    { label: "Companies", value: String(uniqueCompanies.size) },
    { label: "Active pipeline", value: String(openCount) },
    { label: "Approved", value: String(approvedCount) },
    { label: "Reported amount", value: `$${totalAmount.toLocaleString()}` }
  ];
}

function renderDashboard(investments) {
  const cards = buildDashboardCards(investments);
  dashboardCards.innerHTML = cards
    .map(
      (card) => `
        <article class="dashboard-card">
          <p class="dashboard-label">${escapeHtml(card.label)}</p>
          <p class="dashboard-value">${escapeHtml(card.value)}</p>
        </article>
      `
    )
    .join("");
}

function renderFilterOptions() {
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

  assignOptions(statusFilter, "All statuses", statuses);
  assignOptions(stageFilter, "All stages", stages);
  assignOptions(ownerFilter, "All owners", owners);
}

function renderCompanySuggestions() {
  const companies = Array.from(
    new Set(allInvestments.map((item) => item.company).filter(Boolean))
  ).sort((left, right) => left.localeCompare(right));

  companySuggestions.innerHTML = companies
    .map((company) => `<option value="${escapeHtml(company)}"></option>`)
    .join("");
}

function beginEditInvestment(investmentId) {
  const investment = allInvestments.find((item) => item.id === investmentId);
  if (!investment) {
    return;
  }

  editingInvestmentId.value = investment.id;
  form.elements.company.value = investment.company || "";
  form.elements.amount.value = investment.amount || "";
  form.elements.currency.value = investment.currency || "USD";
  form.elements.stage.value = investment.stage || "";
  form.elements.status.value = investment.status || "";
  form.elements.owner.value = investment.owner || "";
  form.elements.nextStep.value = investment.nextStep || "";
  form.elements.recipients.value = Array.isArray(investment.recipients)
    ? investment.recipients.join(", ")
    : "";
  notesField.value = investment.notes || "";
  submitButton.textContent = "Save changes";
  cancelEditButton.classList.remove("hidden");
  formMessage.textContent = `Editing ${investment.company}.`;
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
    companyNextSteps.innerHTML = "";
    companyTimeline.innerHTML = "";
    return;
  }

  const companyUpdates = allInvestments
    .filter((investment) => companyKey(investment.company) === companyKey(selectedCompany))
    .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

  if (!companyUpdates.length) {
    companyPanel.classList.add("hidden");
    return;
  }

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
  companyPanel.classList.remove("hidden");
  companyPanelTitle.textContent = latest.company || selectedCompany;
  companyPanelCopy.textContent = `${companyUpdates.length} update${companyUpdates.length === 1 ? "" : "s"} saved for this company.`;
  companySummary.innerHTML = [
    { label: "Latest status", value: latest.status || "Not set" },
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
    { label: "Owners involved", value: uniqueOwners.length ? uniqueOwners.join(", ") : "Not set" },
    { label: "Statuses used", value: uniqueStatuses.length ? uniqueStatuses.join(", ") : "Not set" },
    { label: "Latest notes", value: latest.notes || "No notes provided." }
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

  companyNextSteps.innerHTML = nextSteps.length
    ? `<ul class="company-list">${nextSteps
        .slice(0, 6)
        .map((nextStep) => `<li>${escapeHtml(nextStep)}</li>`)
        .join("")}</ul>`
    : '<p class="update-meta">No next steps recorded yet.</p>';

  companyTimeline.innerHTML = companyUpdates
    .map(
      (investment) => `
        <article class="timeline-card">
          <div class="update-head">
            <h3>${escapeHtml(investment.status || "Update")}</h3>
            <span class="status-chip">${escapeHtml(investment.stage || "No stage")}</span>
          </div>
          <p class="update-meta">
            ${escapeHtml(investment.createdAt)} • Owner: ${escapeHtml(investment.owner || "Not set")}
          </p>
          <p class="update-meta">
            Amount: ${
              investment.amount
                ? `${escapeHtml(investment.currency)} ${escapeHtml(investment.amount)}`
                : "Amount not specified"
            } • Submitted by: ${escapeHtml(investment.submittedBy || "Unknown")}
          </p>
          <p class="update-meta">
            Next: ${escapeHtml(investment.nextStep || "No next step provided")}
          </p>
          <p class="update-notes">${escapeHtml(investment.notes || "No notes provided.")}</p>
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
            ${amount} • ${escapeHtml(investment.stage || "Stage not specified")}
          </p>
          <p class="update-meta">
            Owner: ${escapeHtml(investment.owner || "Not set")} • Submitted by:
            ${escapeHtml(investment.submittedBy || "Unknown")}
          </p>
          <p class="update-meta">
            Next: ${escapeHtml(investment.nextStep || "Not set")}
          </p>
          <p class="update-notes">${escapeHtml(investment.notes || "No notes provided.")}</p>
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

function renderAll() {
  renderCompanySuggestions();
  renderFilterOptions();
  const filteredInvestments = filterInvestments(allInvestments);
  renderDashboard(filteredInvestments);
  renderUpdates(filteredInvestments);
  renderCompanyPanel();
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

  loginCopy.textContent =
    config.authMode === "individual"
      ? `Use your email and your personal workspace password to sign in. ${config.teamUserCount} team login${config.teamUserCount === 1 ? "" : "s"} configured.`
      : "Use your email and the shared workspace password to unlock updates.";

  if (!config.aiConfigured) {
    deckMessage.textContent =
      "Add OPENAI_API_KEY in Render to turn on deck summarization.";
  }

  if (!config.authConfigured) {
    loginMessage.textContent =
      "The server still needs SESSION_SECRET plus TEAM_PASSWORD or TEAM_USERS in Render.";
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

async function loadUpdates() {
  try {
    const data = await fetchJson("/api/investments");
    allInvestments = data.investments;
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

    notesField.value = result.summary;
    deckMessage.textContent = "Deck summary added to notes.";
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
    amount: formData.get("amount"),
    currency: formData.get("currency"),
    stage: formData.get("stage"),
    status: formData.get("status"),
    owner: formData.get("owner"),
    nextStep: formData.get("nextStep"),
    recipients,
    notes: formData.get("notes")
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
  loadUpdates().catch((error) => {
    formMessage.textContent = error.message;
  });
});

downloadCsvButton.addEventListener("click", () => {
  window.location.href = "/api/investments.csv";
});

downloadExcelButton.addEventListener("click", () => {
  window.location.href = "/api/investments.xlsx";
});

cancelEditButton.addEventListener("click", () => {
  resetFormToCreateMode();
  formMessage.textContent = "Edit canceled.";
});

[searchFilter, statusFilter, stageFilter, ownerFilter].forEach((element) => {
  element.addEventListener("input", renderAll);
  element.addEventListener("change", renderAll);
});

closeCompanyPanelButton.addEventListener("click", () => {
  selectedCompany = "";
  renderCompanyPanel();
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
    companyPanel.scrollIntoView({ behavior: "smooth", block: "start" });
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

Promise.all([loadConfig(), loadUpdates()]).catch((error) => {
  if (error.status === 401) {
    setSignedInState(null);
    return;
  }

  loginMessage.textContent = error.message;
});
