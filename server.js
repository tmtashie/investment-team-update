const http = require("http");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

function loadEnvFile() {
  const envPath = path.join(__dirname, ".env");
  if (!fs.existsSync(envPath)) {
    return;
  }

  const lines = fs.readFileSync(envPath, "utf8").split(/\r?\n/);
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) {
      continue;
    }

    const separatorIndex = trimmed.indexOf("=");
    if (separatorIndex === -1) {
      continue;
    }

    const key = trimmed.slice(0, separatorIndex).trim();
    const value = trimmed.slice(separatorIndex + 1).trim();

    if (key && process.env[key] === undefined) {
      process.env[key] = value;
    }
  }
}

loadEnvFile();

const PORT = process.env.PORT || 3000;
const DATA_DIR = process.env.DATA_DIR
  ? path.resolve(process.env.DATA_DIR)
  : path.join(__dirname, "data");
const DATA_FILE = path.join(DATA_DIR, "investments.json");
const TASKS_FILE = path.join(DATA_DIR, "tasks.json");
const UPLOADS_DIR = path.join(DATA_DIR, "uploads");
const FAMILY_OFFICE_WORKBOOK_FILE = path.join(
  __dirname,
  "Family_Office_Private_Investment_Tracker.xlsx"
);
const PUBLIC_DIR = path.join(__dirname, "public");
const SESSION_COOKIE = "investment_team_session";
const SESSION_SECRET = process.env.SESSION_SECRET || "change-me-before-production";
const SESSION_DURATION_MS = 1000 * 60 * 60 * 24 * 7;
const MAX_BODY_SIZE_BYTES = 20 * 1024 * 1024;
const OPENAI_MODEL = process.env.OPENAI_MODEL || "gpt-4o-mini";
const INVESTMENT_ENTITIES = [
  "Beaman Ventures",
  "Lee Beaman",
  "Katherine Trust",
  "Natalie Trust"
];
const ENTITY_ALIASES = {
  "Beaman Ventures": "Beaman Ventures",
  "Lee Beaman": "Lee Beaman",
  "Kat Trust": "Katherine Trust",
  "Nat Trust": "Natalie Trust",
  "Katherine Trust": "Katherine Trust",
  "Natalie Trust": "Natalie Trust"
};

const DEFAULT_RECIPIENTS = splitCsv(process.env.TEAM_EMAILS || "");

function splitCsv(value) {
  return String(value)
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

function normalizeEntityName(value) {
  const raw = String(value || "").trim();
  return ENTITY_ALIASES[raw] || raw;
}

function parseTeamUsers(value) {
  return splitCsv(value).reduce((users, entry) => {
    const parts = entry.split(":").map((part) => part.trim());
    if (parts.length < 2) {
      return users;
    }

    const [emailRaw, passwordRaw, roleRaw] = parts;
    const email = emailRaw.toLowerCase();
    const password = passwordRaw;
    const role = roleRaw === "viewer" ? "viewer" : "editor";

    if (!email || !password) {
      return users;
    }

    users[email] = {
      email,
      password,
      role
    };
    return users;
  }, {});
}

const TEAM_USERS = parseTeamUsers(process.env.TEAM_USERS || "");
const ALLOWED_EMAILS = Array.from(
  new Set(
    splitCsv(process.env.TEAM_ALLOWED_EMAILS || "")
      .map((email) => email.toLowerCase())
      .concat(Object.keys(TEAM_USERS))
  )
);

function ensureDataFile() {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
  }

  if (!fs.existsSync(UPLOADS_DIR)) {
    fs.mkdirSync(UPLOADS_DIR, { recursive: true });
  }

  if (!fs.existsSync(DATA_FILE)) {
    fs.writeFileSync(DATA_FILE, "[]\n", "utf8");
  }

  if (!fs.existsSync(TASKS_FILE)) {
    fs.writeFileSync(TASKS_FILE, "[]\n", "utf8");
  }
}

function makeId() {
  return crypto.randomUUID();
}

function normalizeCompanyKey(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function normalizeDocuments(value) {
  if (!Array.isArray(value)) {
    return [];
  }

  return value
    .map((document) => ({
      id: String((document && document.id) || makeId()).trim(),
      name: String((document && document.name) || "").trim(),
      storedName: String((document && document.storedName) || "").trim(),
      url: String((document && document.url) || "").trim(),
      uploadedAt: String((document && document.uploadedAt) || new Date().toISOString()).trim()
    }))
    .filter((document) => document.name && document.url);
}

function safeFilename(filename) {
  const extension = path.extname(String(filename || "")).toLowerCase();
  const basename = path
    .basename(String(filename || ""), extension)
    .replace(/[^a-z0-9_-]+/gi, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 60);

  return `${basename || "document"}-${Date.now()}-${crypto.randomUUID().slice(0, 8)}${extension}`;
}

function getContentType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  return (
    {
      ".pdf": "application/pdf",
      ".xlsx":
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      ".xls": "application/vnd.ms-excel",
      ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
      ".doc": "application/msword",
      ".docx":
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      ".png": "image/png",
      ".jpg": "image/jpeg",
      ".jpeg": "image/jpeg",
      ".csv": "text/csv; charset=utf-8",
      ".txt": "text/plain; charset=utf-8"
    }[ext] || "application/octet-stream"
  );
}

function findLatestByCompanyKey(companyKey, investments) {
  if (!companyKey) {
    return null;
  }

  return investments.find((investment) => investment.companyKey === companyKey) || null;
}

function normalizeStructuredRow(row, fallback = {}) {
  return {
    date: String(row.date || fallback.date || "").trim(),
    type: String(row.type || fallback.type || "").trim(),
    title: String(row.title || fallback.title || "").trim(),
    amount: String(row.amount || fallback.amount || "").trim(),
    officialValue: String(row.officialValue || fallback.officialValue || "").trim(),
    internalValue: String(row.internalValue || fallback.internalValue || "").trim(),
    exitValue: String(row.exitValue || fallback.exitValue || "").trim(),
    totalPercent: String(row.totalPercent || fallback.totalPercent || "").trim(),
    entityPercent: String(row.entityPercent || fallback.entityPercent || "").trim(),
    notes: String(row.notes || fallback.notes || "").trim(),
    summary: String(row.summary || fallback.summary || "").trim(),
    sourceUpdateId: String(row.sourceUpdateId || fallback.sourceUpdateId || "").trim()
  };
}

function normalizeStructuredRows(rows, fallbackRows = []) {
  if (Array.isArray(rows) && rows.length) {
    return rows
      .map((row) => normalizeStructuredRow(row))
      .filter((row) => Object.values(row).some(Boolean));
  }

  return fallbackRows
    .map((row) => normalizeStructuredRow(row))
    .filter((row) => Object.values(row).some(Boolean));
}

function normalizeInvestment(entry) {
  const entity = String(entry.entity || "").trim();
  const investmentId = entry.id || makeId();
  const createdAt = String(entry.createdAt || new Date().toISOString());
  const notes = String(entry.notes || "").trim();
  const deckSummary = String(entry.deckSummary || "").trim();
  const capitalCallDate = String(entry.capitalCallDate || "").trim();
  const capitalCallAmount = String(entry.capitalCallAmount || "").trim();
  const distributionDate = String(entry.distributionDate || "").trim();
  const distributionAmount = String(entry.distributionAmount || "").trim();
  const valuationDate = String(entry.valuationDate || "").trim();
  const officialValue = String(entry.officialValue || "").trim();
  const internalValue = String(entry.internalValue || "").trim();
  const exitValue = String(entry.exitValue || "").trim();
  const ownershipPercent = String(entry.ownershipPercent || "").trim();
  const entityOwnershipPercent = String(entry.entityOwnershipPercent || "").trim();
  const ownershipNotes = String(entry.ownershipNotes || "").trim();
  const followOnCapitalAmount = String(entry.followOnCapitalAmount || "").trim();
  const followOnCapitalStatus = String(entry.followOnCapitalStatus || "").trim();
  const followOnCapitalNotes = String(entry.followOnCapitalNotes || "").trim();
  const decisionDate = String(entry.decisionDate || "").trim();
  const decisionType = String(entry.decisionType || "").trim();
  const decisionSummary = String(entry.decisionSummary || "").trim();
  const fallbackCapitalActivity = [];
  const fallbackResearchEntries = [];
  const fallbackValuationHistory = [];
  const fallbackOwnershipHistory = [];
  const fallbackFollowOnHistory = [];
  const fallbackDecisionLog = [];

  if (capitalCallAmount || capitalCallDate) {
    fallbackCapitalActivity.push({
      date: capitalCallDate || createdAt,
      type: "Capital Call",
      amount: capitalCallAmount,
      notes: "Captured from investment update",
      sourceUpdateId: investmentId
    });
  }

  if (distributionAmount || distributionDate) {
    fallbackCapitalActivity.push({
      date: distributionDate || createdAt,
      type: "Distribution",
      amount: distributionAmount,
      notes: "Captured from investment update",
      sourceUpdateId: investmentId
    });
  }

  if (notes) {
    fallbackResearchEntries.push({
      date: createdAt,
      type: "Operating Note",
      summary: notes,
      sourceUpdateId: investmentId
    });
  }

  if (deckSummary) {
    fallbackResearchEntries.push({
      date: createdAt,
      type: "Deck Summary",
      summary: deckSummary,
      sourceUpdateId: investmentId
    });
  }

  if (officialValue || internalValue || exitValue || valuationDate) {
    fallbackValuationHistory.push({
      date: valuationDate || createdAt,
      officialValue,
      internalValue,
      exitValue,
      notes: "Captured from investment update",
      sourceUpdateId: investmentId
    });
  }

  if (ownershipPercent || entityOwnershipPercent || ownershipNotes) {
    fallbackOwnershipHistory.push({
      date: createdAt,
      totalPercent: ownershipPercent,
      entityPercent: entityOwnershipPercent,
      notes: ownershipNotes,
      sourceUpdateId: investmentId
    });
  }

  if (followOnCapitalAmount || followOnCapitalStatus || followOnCapitalNotes) {
    fallbackFollowOnHistory.push({
      date: createdAt,
      amount: followOnCapitalAmount,
      type: followOnCapitalStatus,
      notes: followOnCapitalNotes,
      sourceUpdateId: investmentId
    });
  }

  if (decisionDate || decisionType || decisionSummary) {
    fallbackDecisionLog.push({
      date: decisionDate || createdAt,
      type: decisionType,
      summary: decisionSummary,
      sourceUpdateId: investmentId
    });
  }

  return {
    id: investmentId,
    company: String(entry.company || "").trim(),
    companyKey: entry.companyKey || normalizeCompanyKey(entry.company),
    entity,
    amount: String(entry.amount || "").trim(),
    currency: String(entry.currency || "USD").trim() || "USD",
    stage: String(entry.stage || "").trim(),
    status: String(entry.status || "").trim(),
    owner: String(entry.owner || "").trim(),
    nextStep: String(entry.nextStep || "").trim(),
    notes,
    deckSummary,
    capitalCallDate,
    capitalCallAmount,
    distributionDate,
    distributionAmount,
    valuationDate,
    officialValue,
    internalValue,
    exitValue,
    ownershipPercent,
    entityOwnershipPercent,
    ownershipNotes,
    followOnCapitalAmount,
    followOnCapitalStatus,
    followOnCapitalNotes,
    documentLinks: String(entry.documentLinks || "").trim(),
    documents: normalizeDocuments(entry.documents),
    decisionDate,
    decisionType,
    decisionSummary,
    researchEntries: normalizeStructuredRows(entry.researchEntries, fallbackResearchEntries),
    capitalActivity: normalizeStructuredRows(entry.capitalActivity, fallbackCapitalActivity),
    valuationHistory: normalizeStructuredRows(entry.valuationHistory, fallbackValuationHistory),
    ownershipHistory: normalizeStructuredRows(entry.ownershipHistory, fallbackOwnershipHistory),
    followOnHistory: normalizeStructuredRows(entry.followOnHistory, fallbackFollowOnHistory),
    decisionLog: normalizeStructuredRows(entry.decisionLog, fallbackDecisionLog),
    recipients: Array.isArray(entry.recipients)
      ? entry.recipients.map((value) => String(value).trim()).filter(Boolean)
      : [],
    submittedBy: String(entry.submittedBy || "").trim(),
    createdAt
  };
}

function normalizeTask(entry) {
  return {
    id: entry.id || makeId(),
    company: String(entry.company || "").trim(),
    companyKey: normalizeCompanyKey(entry.company),
    entity: String(entry.entity || "").trim(),
    title: String(entry.title || "").trim(),
    description: String(entry.description || "").trim(),
    dueDate: String(entry.dueDate || "").trim(),
    status: String(entry.status || "Open").trim() || "Open",
    priority: String(entry.priority || "Medium").trim() || "Medium",
    assignee: String(entry.assignee || "").trim(),
    category: String(entry.category || "").trim(),
    createdBy: String(entry.createdBy || "").trim(),
    createdAt: String(entry.createdAt || new Date().toISOString()),
    completedAt: String(entry.completedAt || "").trim()
  };
}

function writeInvestments(investments) {
  ensureDataFile();
  fs.writeFileSync(DATA_FILE, JSON.stringify(investments, null, 2) + "\n", "utf8");
}

function readInvestments() {
  ensureDataFile();
  try {
    const raw = fs.readFileSync(DATA_FILE, "utf8");
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      return [];
    }

    const normalized = parsed.map(normalizeInvestment);
    const changed = JSON.stringify(parsed) !== JSON.stringify(normalized);
    if (changed) {
      writeInvestments(normalized);
    }

    return normalized;
  } catch (error) {
    return [];
  }
}

function writeTasks(tasks) {
  ensureDataFile();
  fs.writeFileSync(TASKS_FILE, JSON.stringify(tasks, null, 2) + "\n", "utf8");
}

function readTasks() {
  ensureDataFile();
  try {
    const raw = fs.readFileSync(TASKS_FILE, "utf8");
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      return [];
    }

    const normalized = parsed.map(normalizeTask);
    const changed = JSON.stringify(parsed) !== JSON.stringify(normalized);
    if (changed) {
      writeTasks(normalized);
    }

    return normalized;
  } catch (error) {
    return [];
  }
}

function saveTask(entry) {
  const tasks = readTasks();
  const normalized = normalizeTask(entry);
  tasks.unshift(normalized);
  writeTasks(tasks);
  return normalized;
}

function updateTask(id, updates) {
  const tasks = readTasks();
  const index = tasks.findIndex((task) => task.id === id);

  if (index === -1) {
    return null;
  }

  const merged = normalizeTask({
    ...tasks[index],
    ...updates,
    id: tasks[index].id,
    createdAt: tasks[index].createdAt,
    createdBy: tasks[index].createdBy,
    completedAt:
      updates.status === "Completed" && !tasks[index].completedAt
        ? new Date().toISOString()
        : updates.status && updates.status !== "Completed"
          ? ""
          : tasks[index].completedAt
  });

  tasks[index] = merged;
  writeTasks(tasks);
  return merged;
}

function deleteTask(id) {
  const tasks = readTasks();
  const remaining = tasks.filter((task) => task.id !== id);

  if (remaining.length === tasks.length) {
    return false;
  }

  writeTasks(remaining);
  return true;
}

function saveInvestment(entry) {
  const investments = readInvestments();
  const normalizedEntry = normalizeInvestment(entry);
  const latestMatch = findLatestByCompanyKey(normalizedEntry.companyKey, investments);

  if (latestMatch && latestMatch.company) {
    normalizedEntry.company = latestMatch.company;
  }

  investments.unshift(normalizedEntry);
  writeInvestments(investments);
}

function updateInvestment(id, updates) {
  const investments = readInvestments();
  const index = investments.findIndex((investment) => investment.id === id);

  if (index === -1) {
    return null;
  }

  const merged = normalizeInvestment({
    ...investments[index],
    ...updates,
    id: investments[index].id,
    companyKey: normalizeCompanyKey(updates.company || investments[index].company),
    createdAt: investments[index].createdAt,
    submittedBy: investments[index].submittedBy
  });

  const latestMatch = findLatestByCompanyKey(
    merged.companyKey,
    investments.filter((investment) => investment.id !== id)
  );

  if (latestMatch && latestMatch.company) {
    merged.company = latestMatch.company;
  }

  investments[index] = merged;
  writeInvestments(investments);
  return merged;
}

function deleteInvestment(id) {
  const investments = readInvestments();
  const remaining = investments.filter((investment) => investment.id !== id);

  if (remaining.length === investments.length) {
    return false;
  }

  writeInvestments(remaining);
  return true;
}

function sortStructuredRows(rows) {
  return [...rows].sort((left, right) => {
    const rightTime = new Date(right.date || 0).getTime();
    const leftTime = new Date(left.date || 0).getTime();
    return rightTime - leftTime;
  });
}

function buildCompanyRecords(investments, tasks = []) {
  const grouped = new Map();

  investments.forEach((investment) => {
    if (!investment.companyKey) {
      return;
    }

    if (!grouped.has(investment.companyKey)) {
      grouped.set(investment.companyKey, []);
    }

    grouped.get(investment.companyKey).push(investment);
  });

  return Array.from(grouped.entries())
    .map(([key, companyInvestments]) => {
      const updates = [...companyInvestments].sort(
        (left, right) => new Date(right.createdAt).getTime() - new Date(left.createdAt).getTime()
      );
      const latest = updates[0];
      const companyTasks = tasks
        .filter((task) => task.companyKey === key)
        .sort(
          (left, right) =>
            new Date(left.dueDate || left.createdAt || 0).getTime() -
            new Date(right.dueDate || right.createdAt || 0).getTime()
        );
      const documents = updates
        .flatMap((investment) =>
          (investment.documents || []).map((document) => ({
            ...document,
            company: investment.company,
            entity: investment.entity,
            sourceUpdateId: investment.id
          }))
        )
        .sort(
          (left, right) =>
            new Date(right.uploadedAt || right.createdAt || 0).getTime() -
            new Date(left.uploadedAt || left.createdAt || 0).getTime()
        );

      return {
        company: latest.company,
        companyKey: key,
        profile: {
          entity: latest.entity,
          stage: latest.stage,
          status: latest.status,
          owner: latest.owner,
          nextStep: latest.nextStep,
          currency: latest.currency
        },
        updates,
        researchEntries: sortStructuredRows(updates.flatMap((investment) => investment.researchEntries || [])),
        capitalActivities: sortStructuredRows(
          updates.flatMap((investment) =>
            (investment.capitalActivity || []).map((activity) => ({
              ...activity,
              entity: investment.entity,
              currency: investment.currency,
              sourceUpdateId: investment.id
            }))
          )
        ),
        valuationHistory: sortStructuredRows(
          updates.flatMap((investment) =>
            (investment.valuationHistory || []).map((valuation) => ({
              ...valuation,
              entity: investment.entity,
              currency: investment.currency,
              sourceUpdateId: investment.id
            }))
          )
        ),
        ownershipHistory: sortStructuredRows(
          updates.flatMap((investment) =>
            (investment.ownershipHistory || []).map((ownership) => ({
              ...ownership,
              entity: investment.entity,
              sourceUpdateId: investment.id
            }))
          )
        ),
        followOnHistory: sortStructuredRows(
          updates.flatMap((investment) =>
            (investment.followOnHistory || []).map((followOn) => ({
              ...followOn,
              entity: investment.entity,
              currency: investment.currency,
              sourceUpdateId: investment.id
            }))
          )
        ),
        decisionLog: sortStructuredRows(
          updates.flatMap((investment) =>
            (investment.decisionLog || []).map((decision) => ({
              ...decision,
              entity: investment.entity,
              sourceUpdateId: investment.id
            }))
          )
        ),
        documents,
        tasks: companyTasks,
        entities: Array.from(new Set(updates.map((investment) => investment.entity).filter(Boolean)))
      };
    })
    .sort((left, right) => left.company.localeCompare(right.company));
}

function signSession(payload) {
  const encoded = Buffer.from(JSON.stringify(payload), "utf8").toString("base64url");
  const signature = crypto
    .createHmac("sha256", SESSION_SECRET)
    .update(encoded)
    .digest("base64url");
  return `${encoded}.${signature}`;
}

function verifySession(value) {
  if (!value || !value.includes(".")) {
    return null;
  }

  const [encoded, signature] = value.split(".");
  const expected = crypto
    .createHmac("sha256", SESSION_SECRET)
    .update(encoded)
    .digest("base64url");

  const signatureBuffer = Buffer.from(signature);
  const expectedBuffer = Buffer.from(expected);

  if (
    signatureBuffer.length !== expectedBuffer.length ||
    !crypto.timingSafeEqual(signatureBuffer, expectedBuffer)
  ) {
    return null;
  }

  try {
    const payload = JSON.parse(Buffer.from(encoded, "base64url").toString("utf8"));
    if (!payload.expiresAt || payload.expiresAt < Date.now()) {
      return null;
    }
    return payload;
  } catch (error) {
    return null;
  }
}

function parseCookies(request) {
  const header = request.headers.cookie || "";
  return header.split(";").reduce((cookies, chunk) => {
    const [name, ...rest] = chunk.trim().split("=");
    if (!name) {
      return cookies;
    }

    cookies[name] = decodeURIComponent(rest.join("="));
    return cookies;
  }, {});
}

function createSessionCookie(sessionValue, expiresAt) {
  const parts = [
    `${SESSION_COOKIE}=${encodeURIComponent(sessionValue)}`,
    "Path=/",
    "HttpOnly",
    "SameSite=Lax",
    `Max-Age=${Math.max(0, Math.floor((expiresAt - Date.now()) / 1000))}`
  ];

  if (process.env.NODE_ENV === "production") {
    parts.push("Secure");
  }

  return parts.join("; ");
}

function clearSessionCookie() {
  return `${SESSION_COOKIE}=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0`;
}

function getSessionUser(request) {
  const cookies = parseCookies(request);
  const session = verifySession(cookies[SESSION_COOKIE]);
  if (!session) {
    return null;
  }

  return {
    email: session.email,
    name: session.name || session.email,
    role: session.role || "editor"
  };
}

function canEdit(user) {
  return user && user.role !== "viewer";
}

function sendJson(response, statusCode, payload, headers = {}) {
  response.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    ...headers
  });
  response.end(JSON.stringify(payload));
}

function sendText(response, statusCode, body, headers = {}) {
  response.writeHead(statusCode, {
    "Content-Type": "text/plain; charset=utf-8",
    ...headers
  });
  response.end(body);
}

function parseNumericValue(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  const cleaned = String(value ?? "")
    .trim()
    .replace(/[$,()\s]/g, "")
    .replace(/[^\d.-]/g, "");

  if (!cleaned) {
    return 0;
  }

  const amount = Number(cleaned);
  return Number.isFinite(amount) ? amount : 0;
}

function parseWorkbookDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString();
  }

  const text = String(value || "").trim();
  if (!text) {
    return new Date().toISOString();
  }

  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) {
    return parsed.toISOString();
  }

  return new Date().toISOString();
}

function toPositiveNumber(value) {
  const amount = parseNumericValue(value);
  return amount > 0 ? amount : 0;
}

function yearFraction(fromDate, toDate) {
  return (toDate.getTime() - fromDate.getTime()) / (365 * 24 * 60 * 60 * 1000);
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

function normalizeHeaderKey(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function inferEntityFromContext(...values) {
  const haystack = values
    .map((value) => String(value || "").toLowerCase())
    .join(" ");

  if (haystack.includes("beaman ventures") || /\bbv\b/.test(haystack)) {
    return "Beaman Ventures";
  }

  if (haystack.includes("lee beaman") || /\blee\b/.test(haystack)) {
    return "Lee Beaman";
  }

  if (
    haystack.includes("kat trust") ||
    haystack.includes("katherine trust") ||
    /\bkat\b/.test(haystack)
  ) {
    return "Katherine Trust";
  }

  if (
    haystack.includes("nat trust") ||
    haystack.includes("natalie trust") ||
    /\bnat\b/.test(haystack)
  ) {
    return "Natalie Trust";
  }

  return "";
}

function extractImportTag(notes) {
  const match = String(notes || "").match(/^\[Workbook Import:[^\]]+\]/);
  return match ? match[0] : "";
}

function getExistingImportTags(investments) {
  return new Set(
    investments
      .map((investment) => extractImportTag(investment.notes))
      .filter(Boolean)
  );
}

function buildImportNote(tag, lines) {
  return [tag].concat(lines.filter(Boolean)).join("\n");
}

function csvEscape(value) {
  const text = String(value ?? "");
  if (/[",\n]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}

function buildInvestmentsCsv(investments) {
  const headers = [
    "Entity",
    "Company",
    "Amount",
    "Currency",
    "Stage",
    "Status",
    "Owner",
    "Next Step",
    "Notes",
    "Deck Summary",
    "Capital Call Date",
    "Capital Call Amount",
    "Distribution Date",
    "Distribution Amount",
    "Valuation Date",
    "Official Value",
    "Internal Value",
    "Exit Value",
    "Ownership Percent",
    "Entity Ownership Percent",
    "Ownership Notes",
    "Follow-On Capital Amount",
    "Follow-On Capital Status",
    "Follow-On Capital Notes",
    "Document Links",
    "Uploaded Documents",
    "Decision Date",
    "Decision Type",
    "Decision Summary",
    "Submitted By",
    "Recipients",
    "Created At"
  ];

  const rows = investments.map((investment) =>
    [
      investment.entity,
      investment.company,
      investment.amount,
      investment.currency,
      investment.stage,
      investment.status,
      investment.owner,
      investment.nextStep,
      investment.notes,
      investment.deckSummary,
      investment.capitalCallDate,
      investment.capitalCallAmount,
      investment.distributionDate,
      investment.distributionAmount,
      investment.valuationDate,
      investment.officialValue,
      investment.internalValue,
      investment.exitValue,
      investment.ownershipPercent,
      investment.entityOwnershipPercent,
      investment.ownershipNotes,
      investment.followOnCapitalAmount,
      investment.followOnCapitalStatus,
      investment.followOnCapitalNotes,
      investment.documentLinks,
      investment.documents.map((document) => `${document.name} (${document.url})`).join(" | "),
      investment.decisionDate,
      investment.decisionType,
      investment.decisionSummary,
      investment.submittedBy,
      Array.isArray(investment.recipients) ? investment.recipients.join(", ") : "",
      investment.createdAt
    ]
      .map(csvEscape)
      .join(",")
  );

  return [headers.map(csvEscape).join(","), ...rows].join("\n");
}

function buildInvestmentsWorkbookBuffer(investments) {
  const XLSX = require("xlsx");
  const rows = investments.map((investment) => ({
    Entity: investment.entity,
    Company: investment.company,
    Amount: investment.amount,
    Currency: investment.currency,
    Stage: investment.stage,
    Status: investment.status,
    Owner: investment.owner,
    "Next Step": investment.nextStep,
    Notes: investment.notes,
    "Deck Summary": investment.deckSummary,
    "Capital Call Date": investment.capitalCallDate,
    "Capital Call Amount": investment.capitalCallAmount,
    "Distribution Date": investment.distributionDate,
    "Distribution Amount": investment.distributionAmount,
    "Valuation Date": investment.valuationDate,
    "Official Value": investment.officialValue,
    "Internal Value": investment.internalValue,
    "Exit Value": investment.exitValue,
    "Ownership Percent": investment.ownershipPercent,
    "Entity Ownership Percent": investment.entityOwnershipPercent,
    "Ownership Notes": investment.ownershipNotes,
    "Follow-On Capital Amount": investment.followOnCapitalAmount,
    "Follow-On Capital Status": investment.followOnCapitalStatus,
    "Follow-On Capital Notes": investment.followOnCapitalNotes,
    "Document Links": investment.documentLinks,
    "Uploaded Documents": investment.documents
      .map((document) => `${document.name} (${document.url})`)
      .join(" | "),
    "Decision Date": investment.decisionDate,
    "Decision Type": investment.decisionType,
    "Decision Summary": investment.decisionSummary,
    "Submitted By": investment.submittedBy,
    Recipients: Array.isArray(investment.recipients) ? investment.recipients.join(", ") : "",
    "Created At": investment.createdAt
  }));

  const sheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "Investment Updates");
  return XLSX.write(workbook, { bookType: "xlsx", type: "buffer" });
}

function importWorkbookIntoInvestments(buffer, sessionUser, sourceName = "") {
  const XLSX = require("xlsx");
  const workbook = XLSX.read(buffer, { type: "buffer", cellDates: true });
  const investments = readInvestments();
  const existingTags = getExistingImportTags(investments);
  const imported = [];
  const usedSheets = [];

  function pushImported(entry, tag) {
    if (existingTags.has(tag)) {
      return false;
    }

    imported.push(normalizeInvestment(entry));
    existingTags.add(tag);
    return true;
  }

  const updateSheetNames = ["Investment Updates", "App Updates Export"];
  const updateSheetName = updateSheetNames.find((name) => workbook.Sheets[name]);
  if (updateSheetName) {
    usedSheets.push(updateSheetName);
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[updateSheetName], {
      defval: "",
      raw: false,
      cellDates: true
    });

    rows.forEach((row, index) => {
      const company = String(row.Company || row["Investment Name"] || "").trim();
      const entity = normalizeEntityName(row.Entity);

      if (!company || !entity) {
        return;
      }

      const tag = `[Workbook Import:${updateSheetName}:${index + 2}:${normalizeCompanyKey(company)}]`;
      const notes = buildImportNote(tag, [
        `Imported from workbook sheet "${updateSheetName}".`,
        String(row.Notes || "").trim()
      ]);

      pushImported(
        {
          company,
          entity,
          amount: String(row.Amount || "").trim(),
          currency: String(row.Currency || "USD").trim() || "USD",
          stage: String(row.Stage || "").trim(),
          status: String(row.Status || "Investment update").trim() || "Investment update",
          owner: String(row.Owner || "Workbook import").trim(),
          nextStep: String(row["Next Step"] || "").trim(),
          notes,
          deckSummary: String(row["Deck Summary"] || "").trim(),
          capitalCallDate: String(row["Capital Call Date"] || "").trim(),
          capitalCallAmount: String(row["Capital Call Amount"] || "").trim(),
          distributionDate: String(row["Distribution Date"] || "").trim(),
          distributionAmount: String(row["Distribution Amount"] || "").trim(),
          valuationDate: String(row["Valuation Date"] || "").trim(),
          officialValue: String(row["Official Value"] || "").trim(),
          internalValue: String(row["Internal Value"] || "").trim(),
          exitValue: String(row["Exit Value"] || "").trim(),
          ownershipPercent: String(row["Ownership Percent"] || "").trim(),
          entityOwnershipPercent: String(row["Entity Ownership Percent"] || "").trim(),
          ownershipNotes: String(row["Ownership Notes"] || "").trim(),
          followOnCapitalAmount: String(row["Follow-On Capital Amount"] || "").trim(),
          followOnCapitalStatus: String(row["Follow-On Capital Status"] || "").trim(),
          followOnCapitalNotes: String(row["Follow-On Capital Notes"] || "").trim(),
          documentLinks: String(row["Document Links"] || "").trim(),
          documents: [],
          decisionDate: String(row["Decision Date"] || "").trim(),
          decisionType: String(row["Decision Type"] || "").trim(),
          decisionSummary: String(row["Decision Summary"] || "").trim(),
          recipients: splitCsv(row.Recipients || ""),
          submittedBy: sessionUser.email,
          createdAt: parseWorkbookDate(row["Created At"] || row.Date)
        },
        tag
      );
    });
  }

  const cashSheetNames = ["Cash Flow Ledger", "Capital Call & Distribution Ledger"];
  const cashSheetName = cashSheetNames.find((name) => workbook.Sheets[name]);
  if (cashSheetName) {
    usedSheets.push(cashSheetName);
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[cashSheetName], {
      defval: "",
      raw: false,
      cellDates: true
    });

    rows.forEach((row, index) => {
      const company = String(
        row["Investment Name"] || row.Company || row.Investment || row["Investment ID"] || ""
      ).trim();
      const entity = normalizeEntityName(row.Entity);

      if (!company || !entity) {
        return;
      }

      const contributionAmount = parseNumericValue(row["Contribution Amount"]);
      const distributionAmount = parseNumericValue(row["Distribution Amount"]);
      const netCashFlow = parseNumericValue(row["Net Cash Flow"]);
      const amount = distributionAmount || contributionAmount || Math.abs(netCashFlow) || 0;
      const transactionType = String(row["Transaction Type"] || "Investment update").trim();
      const transactionDate = row["Transaction Date"] || row.Date;
      const transactionId =
        String(row["Transaction ID"] || "").trim() || `${cashSheetName}-${index + 2}`;
      const linkedValuationDate = String(row["Linked valuation date"] || row["Linked valuation date if relevant"] || "").trim();
      const notes = buildImportNote(`[Workbook Import:${transactionId}]`, [
        `Imported from workbook sheet "${cashSheetName}".`,
        transactionType ? `Transaction type: ${transactionType}` : "",
        transactionDate ? `Transaction date: ${String(transactionDate)}` : "",
        contributionAmount ? `Contribution amount: ${contributionAmount.toLocaleString()}` : "",
        distributionAmount ? `Distribution amount: ${distributionAmount.toLocaleString()}` : "",
        linkedValuationDate ? `Linked valuation date: ${linkedValuationDate}` : "",
        String(row.Description || "").trim(),
        String(row.Notes || "").trim()
      ]);

      const status =
        transactionType === "Full Exit"
          ? "Closed"
          : transactionType === "Write-off adjustment"
            ? "Passed"
            : "Investment update";

      pushImported(
        {
          company,
          entity,
          amount: amount ? String(amount) : "",
          currency: "USD",
          stage: String(row["Realized / Unrealized flag"] || "").trim(),
          status,
          owner: "Workbook import",
          nextStep: linkedValuationDate
            ? `Review valuation tied to ${linkedValuationDate}`
            : "",
          notes,
          capitalCallDate:
            contributionAmount || netCashFlow < 0 ? parseWorkbookDate(transactionDate) : "",
          capitalCallAmount:
            contributionAmount || netCashFlow < 0
              ? String(contributionAmount || Math.abs(netCashFlow) || "")
              : "",
          distributionDate:
            distributionAmount || netCashFlow > 0 ? parseWorkbookDate(transactionDate) : "",
          distributionAmount:
            distributionAmount || netCashFlow > 0
              ? String(distributionAmount || netCashFlow || "")
              : "",
          valuationDate: linkedValuationDate,
          officialValue: "",
          internalValue: "",
          exitValue: "",
          ownershipPercent: "",
          entityOwnershipPercent: "",
          ownershipNotes: "",
          recipients: [],
          documentLinks: "",
          documents: [],
          decisionDate: "",
          decisionType: "",
          decisionSummary: "",
          submittedBy: sessionUser.email,
          createdAt: parseWorkbookDate(transactionDate)
        },
        `[Workbook Import:${transactionId}]`
      );
    });
  }

  if (!imported.length) {
    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) {
        return;
      }

      const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: "",
        raw: false,
        cellDates: true
      });

      let headerRowIndex = -1;
      let headerMap = {};

      rows.forEach((row, index) => {
        if (headerRowIndex !== -1) {
          return;
        }

        const normalized = row.map(normalizeHeaderKey);
        if (
          normalized.includes("investment name") &&
          (normalized.includes("current valuation") ||
            normalized.includes("current valuation $") ||
            normalized.includes("active exited lost"))
        ) {
          headerRowIndex = index;
          headerMap = normalized.reduce((map, header, headerIndex) => {
            if (header) {
              map[header] = headerIndex;
            }
            return map;
          }, {});
        }
      });

      if (headerRowIndex === -1) {
        return;
      }

      usedSheets.push(sheetName);
      const inferredEntity = normalizeEntityName(
        inferEntityFromContext(
          sourceName,
          sheetName,
          workbook.Props && workbook.Props.Title,
          workbook.SheetNames.join(" ")
        )
      );

      rows.slice(headerRowIndex + 1).forEach((row, offset) => {
        const investmentName = String(row[headerMap["investment name"]] || "").trim();
        const category = String(row[headerMap["category"]] || "").trim();
        const statusText = String(row[headerMap["active exited lost"]] || "").trim();
        const notes = String(row[headerMap["notes"]] || "").trim();
        const leadContact = String(row[headerMap["lead contact"]] || "").trim();
        const dateCommitted = row[headerMap["date committed"]];
        const committedAmount = parseNumericValue(
          row[headerMap["committed $"]] || row[headerMap["committed"]]
        );
        const calledAmount = parseNumericValue(
          row[headerMap["called $"]] || row[headerMap["called"]]
        );
        const currentValuation = parseNumericValue(
          row[headerMap["current valuation $"]] || row[headerMap["current valuation"]]
        );
        const remainingAmount = parseNumericValue(
          row[headerMap["remaining $"]] || row[headerMap["remaining"]]
        );

        if (!investmentName) {
          return;
        }

        const looksLikeSectionRow =
          !category &&
          !statusText &&
          !notes &&
          !leadContact &&
          !committedAmount &&
          !calledAmount &&
          !currentValuation &&
          !remainingAmount;

        if (looksLikeSectionRow) {
          return;
        }

        const tag = `[Workbook Import:${sheetName}:${headerRowIndex + offset + 2}:${normalizeCompanyKey(
          investmentName
        )}]`;
        const importNotes = buildImportNote(tag, [
          `Imported from workbook sheet "${sheetName}".`,
          category ? `Category: ${category}` : "",
          leadContact ? `Lead contact: ${leadContact}` : "",
          committedAmount ? `Committed: ${committedAmount.toLocaleString()}` : "",
          calledAmount ? `Called: ${calledAmount.toLocaleString()}` : "",
          remainingAmount ? `Remaining: ${remainingAmount.toLocaleString()}` : "",
          currentValuation ? `Current valuation: ${currentValuation.toLocaleString()}` : "",
          notes
        ]);

        const normalizedStatus = (() => {
          const lowered = statusText.toLowerCase();
          if (lowered.includes("exit")) {
            return "Closed";
          }
          if (lowered.includes("lost") || lowered.includes("written")) {
            return "Passed";
          }
          if (lowered.includes("active")) {
            return "Investment update";
          }
          return statusText || "Investment update";
        })();

        pushImported(
          {
            company: investmentName,
            entity: inferredEntity || "Beaman Ventures",
            amount: String(calledAmount || committedAmount || currentValuation || ""),
            currency: "USD",
            stage: category,
            status: normalizedStatus,
            owner: leadContact || "Workbook import",
            nextStep:
              currentValuation && !statusText.toLowerCase().includes("exit")
                ? "Review valuation and update official/internal marks"
                : "",
            notes: importNotes,
            deckSummary: "",
            capitalCallDate: parseWorkbookDate(dateCommitted),
            capitalCallAmount: calledAmount ? String(calledAmount) : "",
            distributionDate: "",
            distributionAmount: "",
            valuationDate: parseWorkbookDate(dateCommitted),
            officialValue: currentValuation ? String(currentValuation) : "",
            internalValue: currentValuation ? String(currentValuation) : "",
            exitValue: "",
            ownershipPercent: "",
            entityOwnershipPercent: "",
            ownershipNotes: "",
            followOnCapitalAmount: "",
            followOnCapitalStatus: "",
            followOnCapitalNotes: "",
            documentLinks: "",
            documents: [],
            decisionDate: "",
            decisionType: "",
            decisionSummary: "",
            recipients: [],
            submittedBy: sessionUser.email,
            createdAt: parseWorkbookDate(dateCommitted)
          },
          tag
        );
      });
    });
  }

  if (!imported.length) {
    throw new Error(
      "No importable rows were found. Use a workbook with an Investment Updates sheet, a Cash Flow Ledger sheet, or a holdings sheet with columns like Investment Name and Current Valuation."
    );
  }

  const merged = imported
    .concat(investments)
    .sort((left, right) => new Date(right.createdAt) - new Date(left.createdAt));

  writeInvestments(merged);

  return {
    importedCount: imported.length,
    skippedCount: 0,
    sheets: usedSheets
  };
}

function serveStaticFile(response, filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const contentTypes = {
    ".html": "text/html; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".js": "application/javascript; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".png": "image/png",
    ".svg": "image/svg+xml",
    ".ico": "image/x-icon"
  };

  fs.readFile(filePath, (error, data) => {
    if (error) {
      sendText(response, 404, "Not found");
      return;
    }

    response.writeHead(200, {
      "Content-Type": contentTypes[ext] || "application/octet-stream"
    });
    response.end(data);
  });
}

function parseRequestBody(request) {
  return new Promise((resolve, reject) => {
    let body = "";

    request.on("data", (chunk) => {
      body += chunk.toString();
      if (body.length > MAX_BODY_SIZE_BYTES) {
        reject(new Error("Request body too large"));
        request.destroy();
      }
    });

    request.on("end", () => {
      if (!body) {
        resolve({});
        return;
      }

      try {
        resolve(JSON.parse(body));
      } catch (error) {
        reject(new Error("Invalid JSON body"));
      }
    });

    request.on("error", reject);
  });
}

function extractResponseText(payload) {
  if (typeof payload.output_text === "string" && payload.output_text.trim()) {
    return payload.output_text.trim();
  }

  if (!Array.isArray(payload.output)) {
    return "";
  }

  const fragments = [];
  for (const item of payload.output) {
    if (!Array.isArray(item.content)) {
      continue;
    }

    for (const contentItem of item.content) {
      if (contentItem.type === "output_text" && contentItem.text) {
        fragments.push(contentItem.text);
      }
    }
  }

  return fragments.join("\n").trim();
}

async function summarizeDeck({ filename, fileData, company, stage }) {
  const apiKey = process.env.OPENAI_API_KEY;

  if (!apiKey) {
    throw new Error("OpenAI summarization is not configured yet.");
  }

  const companyLine = company ? `Company: ${company}` : "Company: Not provided";
  const stageLine = stage ? `Stage: ${stage}` : "Stage: Not provided";
  const fileBytes = Buffer.from(fileData, "base64");
  const uploadForm = new FormData();
  const pdfBlob = new Blob([fileBytes], { type: "application/pdf" });

  uploadForm.append("purpose", "user_data");
  uploadForm.append("file", pdfBlob, filename);

  const uploadResponse = await fetch("https://api.openai.com/v1/files", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`
    },
    body: uploadForm
  });

  if (!uploadResponse.ok) {
    const errorText = await uploadResponse.text();
    throw new Error(`OpenAI file upload failed: ${errorText}`);
  }

  const uploadedFile = await uploadResponse.json();
  const fileId = uploadedFile.id;

  if (!fileId) {
    throw new Error("OpenAI file upload did not return a file ID.");
  }

  const response = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model: OPENAI_MODEL,
      input: [
        {
          role: "user",
          content: [
            {
              type: "input_text",
              text: [
                "Summarize this investment deck for an investment committee update.",
                "Return concise markdown using exactly these headings:",
                "## Thesis",
                "## Business",
                "## Market and traction",
                "## Financing",
                "## Risks",
                "## Recommendation",
                "## Questions for the team",
                "",
                companyLine,
                stageLine,
                "",
                "Under each heading, use short bullets.",
                "If a detail is missing, say that it was not clearly stated.",
                "Keep the tone direct and useful for a deal team."
              ].join("\n")
            },
            {
              type: "input_file",
              file_id: fileId
            }
          ]
        }
      ]
    })
  });

  if (!response.ok) {
    const errorText = await response.text();
    if (errorText.includes("insufficient_quota")) {
      throw new Error(
        "OpenAI summary failed: your OpenAI API account needs billing or more quota before deck summaries can run."
      );
    }
    throw new Error(`OpenAI summary failed: ${errorText}`);
  }

  const payload = await response.json();
  const summary = extractResponseText(payload);

  if (!summary) {
    throw new Error("OpenAI did not return a summary.");
  }

  return summary;
}

async function summarizeEmail({ emailText, company, stage }) {
  const apiKey = process.env.OPENAI_API_KEY;

  if (!apiKey) {
    throw new Error("OpenAI summarization is not configured yet.");
  }

  const companyLine = company ? `Company: ${company}` : "Company: Not provided";
  const stageLine = stage ? `Stage: ${stage}` : "Stage: Not provided";

  const response = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model: OPENAI_MODEL,
      input: [
        {
          role: "user",
          content: [
            {
              type: "input_text",
              text: [
                "Summarize this investment email or email thread for an investment team update.",
                "Return concise markdown using exactly these headings:",
                "## Thesis",
                "## Business",
                "## Traction or evidence",
                "## Financing or ask",
                "## Risks or open questions",
                "## Recommendation",
                "## Suggested next steps",
                "",
                companyLine,
                stageLine,
                "",
                "Under each heading, use short bullets.",
                "If the email does not mention a detail, say that it was not clearly stated.",
                "Focus on the parts that matter to a deal team.",
                "",
                "Email content:",
                emailText
              ].join("\n")
            }
          ]
        }
      ]
    })
  });

  if (!response.ok) {
    const errorText = await response.text();
    if (errorText.includes("insufficient_quota")) {
      throw new Error(
        "OpenAI summary failed: your OpenAI API account needs billing or more quota before email summaries can run."
      );
    }
    throw new Error(`OpenAI summary failed: ${errorText}`);
  }

  const payload = await response.json();
  const summary = extractResponseText(payload);

  if (!summary) {
    throw new Error("OpenAI did not return a summary.");
  }

  return summary;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function formatAmount(entry) {
  return entry.amount ? `${entry.currency} ${entry.amount}` : "Amount not specified";
}

function buildSummary(entry) {
  const entityLine = entry.entity || "Entity not specified";
  const companyLine = entry.company || "Unnamed investment";
  const amountLine = formatAmount(entry);
  const stageLine = entry.stage || "Stage not specified";
  const ownerLine = entry.owner || "No owner listed";
  const nextStepLine = entry.nextStep || "No next step provided";
  const noteLine = entry.notes || "No additional notes.";
  const deckSummaryLine = entry.deckSummary || "No deck summary attached.";
  const capitalCallLine = entry.capitalCallAmount
    ? `${entry.currency} ${entry.capitalCallAmount} on ${entry.capitalCallDate || "date not set"}`
    : "No capital call recorded";
  const distributionLine = entry.distributionAmount
    ? `${entry.currency} ${entry.distributionAmount} on ${entry.distributionDate || "date not set"}`
    : "No distribution recorded";
  const officialValueLine = entry.officialValue
    ? `${entry.currency} ${entry.officialValue}`
    : "No official value recorded";
  const internalValueLine = entry.internalValue
    ? `${entry.currency} ${entry.internalValue}`
    : "No internal value recorded";
  const exitValueLine = entry.exitValue
    ? `${entry.currency} ${entry.exitValue}`
    : "No exit value recorded";
  const ownershipPercentLine = entry.ownershipPercent
    ? `${entry.ownershipPercent}%`
    : "No total ownership recorded";
  const entityOwnershipPercentLine = entry.entityOwnershipPercent
    ? `${entry.entityOwnershipPercent}%`
    : "No entity ownership recorded";
  const ownershipNotesLine = entry.ownershipNotes || "No ownership notes provided.";
  const valuationDateLine = entry.valuationDate || "No valuation date recorded";
  const followOnAmountLine = entry.followOnCapitalAmount
    ? `${entry.currency} ${entry.followOnCapitalAmount}`
    : "No follow-on capital amount recorded";
  const followOnStatusLine = entry.followOnCapitalStatus || "No follow-on decision recorded";
  const followOnNotesLine = entry.followOnCapitalNotes || "No follow-on notes provided.";
  const documentLinksLine = entry.documentLinks || "No documents linked.";
  const uploadedDocumentsLine = Array.isArray(entry.documents) && entry.documents.length
    ? entry.documents.map((document) => `${document.name}: ${document.url}`).join("\n")
    : "No uploaded documents.";
  const decisionDateLine = entry.decisionDate || "No decision date recorded";
  const decisionTypeLine = entry.decisionType || "No decision recorded";
  const decisionSummaryLine = entry.decisionSummary || "No decision summary provided.";

  const subject = `${entry.status}: ${companyLine}`;

  const text = [
    `${entry.status} investment update`,
    "",
    `Company: ${companyLine}`,
    `Entity: ${entityLine}`,
    `Amount: ${amountLine}`,
    `Stage: ${stageLine}`,
    `Owner: ${ownerLine}`,
    `Submitted by: ${entry.submittedBy}`,
    `Next step: ${nextStepLine}`,
    `Capital call: ${capitalCallLine}`,
    `Distribution: ${distributionLine}`,
    `Valuation date: ${valuationDateLine}`,
    `Official value: ${officialValueLine}`,
    `Internal value: ${internalValueLine}`,
    `Exit value: ${exitValueLine}`,
    `Total ownership: ${ownershipPercentLine}`,
    `Entity ownership: ${entityOwnershipPercentLine}`,
    `Follow-on capital amount: ${followOnAmountLine}`,
    `Follow-on capital status: ${followOnStatusLine}`,
    `Decision date: ${decisionDateLine}`,
    `Decision type: ${decisionTypeLine}`,
    "",
    "Summary",
    noteLine,
    "",
    "Deck summary",
    deckSummaryLine,
    "",
    "Follow-on capital notes",
    followOnNotesLine,
    "",
    "Documents",
    documentLinksLine,
    "",
    "Uploaded documents",
    uploadedDocumentsLine,
    "",
    "Ownership notes",
    ownershipNotesLine,
    "",
    "Decision summary",
    decisionSummaryLine,
    "",
    `Submitted at: ${entry.createdAt}`
  ].join("\n");

  const statusColor = {
    "New lead": "#805500",
    "Under review": "#144a7c",
    Approved: "#13613e",
    Passed: "#7c1d1d",
    Closed: "#5c3b00"
  }[entry.status] || "#44403c";

  const html = `
    <div style="font-family: Arial, sans-serif; color: #1f2937; line-height: 1.6; max-width: 680px; margin: 0 auto;">
      <div style="padding: 24px; border-radius: 20px; background: #fffaf2; border: 1px solid #eadfd0;">
        <p style="margin: 0 0 12px; text-transform: uppercase; letter-spacing: 0.12em; color: #8f411f; font-size: 12px;">Investment Team Update</p>
        <h2 style="margin: 0 0 12px;">${escapeHtml(companyLine)}</h2>
        <p style="margin: 0 0 20px;">
          <span style="display: inline-block; padding: 6px 12px; border-radius: 999px; background: rgba(180, 87, 47, 0.10); color: ${statusColor}; font-weight: 700;">
            ${escapeHtml(entry.status)}
          </span>
        </p>
        <table style="border-collapse: collapse; width: 100%;">
          <tr><td style="padding: 8px 0; font-weight: bold; width: 140px;">Entity</td><td>${escapeHtml(entityLine)}</td></tr>
          <tr><td style="padding: 8px 0; font-weight: bold;">Amount</td><td>${escapeHtml(amountLine)}</td></tr>
          <tr><td style="padding: 8px 0; font-weight: bold;">Stage</td><td>${escapeHtml(stageLine)}</td></tr>
          <tr><td style="padding: 8px 0; font-weight: bold;">Owner</td><td>${escapeHtml(ownerLine)}</td></tr>
          <tr><td style="padding: 8px 0; font-weight: bold;">Submitted by</td><td>${escapeHtml(entry.submittedBy)}</td></tr>
          <tr><td style="padding: 8px 0; font-weight: bold;">Next step</td><td>${escapeHtml(nextStepLine)}</td></tr>
          <tr><td style="padding: 8px 0; font-weight: bold;">Capital call</td><td>${escapeHtml(capitalCallLine)}</td></tr>
          <tr><td style="padding: 8px 0; font-weight: bold;">Distribution</td><td>${escapeHtml(distributionLine)}</td></tr>
          <tr><td style="padding: 8px 0; font-weight: bold;">Valuation date</td><td>${escapeHtml(valuationDateLine)}</td></tr>
          <tr><td style="padding: 8px 0; font-weight: bold;">Official value</td><td>${escapeHtml(officialValueLine)}</td></tr>
          <tr><td style="padding: 8px 0; font-weight: bold;">Internal value</td><td>${escapeHtml(internalValueLine)}</td></tr>
          <tr><td style="padding: 8px 0; font-weight: bold;">Exit value</td><td>${escapeHtml(exitValueLine)}</td></tr>
          <tr><td style="padding: 8px 0; font-weight: bold;">Total ownership</td><td>${escapeHtml(ownershipPercentLine)}</td></tr>
          <tr><td style="padding: 8px 0; font-weight: bold;">Entity ownership</td><td>${escapeHtml(entityOwnershipPercentLine)}</td></tr>
          <tr><td style="padding: 8px 0; font-weight: bold;">Follow-on capital</td><td>${escapeHtml(followOnAmountLine)}</td></tr>
          <tr><td style="padding: 8px 0; font-weight: bold;">Follow-on status</td><td>${escapeHtml(followOnStatusLine)}</td></tr>
          <tr><td style="padding: 8px 0; font-weight: bold;">Decision date</td><td>${escapeHtml(decisionDateLine)}</td></tr>
          <tr><td style="padding: 8px 0; font-weight: bold;">Decision type</td><td>${escapeHtml(decisionTypeLine)}</td></tr>
        </table>
      </div>
      <div style="padding: 20px 8px 0;">
        <h3 style="margin-bottom: 8px;">Summary</h3>
        <p style="white-space: pre-wrap; margin-top: 0;">${escapeHtml(noteLine)}</p>
        <h3 style="margin-bottom: 8px;">Deck Summary</h3>
        <p style="white-space: pre-wrap; margin-top: 0;">${escapeHtml(deckSummaryLine)}</p>
        <h3 style="margin-bottom: 8px;">Follow-on Capital Notes</h3>
        <p style="white-space: pre-wrap; margin-top: 0;">${escapeHtml(followOnNotesLine)}</p>
        <h3 style="margin-bottom: 8px;">Documents</h3>
        <p style="white-space: pre-wrap; margin-top: 0;">${escapeHtml(documentLinksLine)}</p>
        <h3 style="margin-bottom: 8px;">Uploaded Documents</h3>
        <p style="white-space: pre-wrap; margin-top: 0;">${escapeHtml(uploadedDocumentsLine)}</p>
        <h3 style="margin-bottom: 8px;">Ownership Notes</h3>
        <p style="white-space: pre-wrap; margin-top: 0;">${escapeHtml(ownershipNotesLine)}</p>
        <h3 style="margin-bottom: 8px;">Decision Summary</h3>
        <p style="white-space: pre-wrap; margin-top: 0;">${escapeHtml(decisionSummaryLine)}</p>
        <p style="margin-top: 20px; color: #6b7280; font-size: 14px;">Submitted at: ${escapeHtml(entry.createdAt)}</p>
      </div>
    </div>
  `;

  return { subject, text, html };
}

async function sendEmail(summary, recipients) {
  const apiKey = process.env.RESEND_API_KEY;
  const fromEmail = process.env.FROM_EMAIL;

  if (!apiKey || !fromEmail || recipients.length === 0) {
    return {
      sent: false,
      reason: "Email not configured",
      preview: summary
    };
  }

  const resendResponse = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      from: fromEmail,
      to: recipients,
      subject: summary.subject,
      html: summary.html,
      text: summary.text
    })
  });

  if (!resendResponse.ok) {
    const errorText = await resendResponse.text();
    throw new Error(`Email send failed: ${errorText}`);
  }

  const result = await resendResponse.json();
  return {
    sent: true,
    id: result.id || null
  };
}

function validateLogin(payload) {
  const email = String(payload.email || "").trim().toLowerCase();
  const password = String(payload.password || "");
  const sharedPassword = String(process.env.TEAM_PASSWORD || "");
  const teamUser = TEAM_USERS[email];

  if (!email) {
    return { error: "Email is required." };
  }

  if (Object.keys(TEAM_USERS).length > 0) {
    if (!teamUser) {
      return { error: "That email is not set up for this workspace yet." };
    }

    if (password !== teamUser.password) {
      return { error: "Incorrect password." };
    }

    return { value: email, role: teamUser.role || "editor" };
  }

  if (!sharedPassword) {
    return { error: "Set TEAM_PASSWORD or TEAM_USERS on the server first." };
  }

  if (ALLOWED_EMAILS.length > 0 && !ALLOWED_EMAILS.includes(email)) {
    return { error: "That email is not allowed for this workspace." };
  }

  if (password !== sharedPassword) {
    return { error: "Incorrect password." };
  }

  return { value: email, role: "editor" };
}

function validateTaskSubmission(payload, sessionUser) {
  const clean = normalizeTask({
    company: payload.company,
    entity: payload.entity,
    title: payload.title,
    description: payload.description,
    dueDate: payload.dueDate,
    status: payload.status,
    priority: payload.priority,
    assignee: payload.assignee,
    category: payload.category,
    createdBy: sessionUser.email,
    createdAt: new Date().toISOString()
  });

  if (!clean.title) {
    return { error: "Task title is required." };
  }

  return { value: clean };
}

function validateTaskPatch(payload) {
  const clean = {
    company: String(payload.company || "").trim(),
    entity: String(payload.entity || "").trim(),
    title: String(payload.title || "").trim(),
    description: String(payload.description || "").trim(),
    dueDate: String(payload.dueDate || "").trim(),
    status: String(payload.status || "Open").trim() || "Open",
    priority: String(payload.priority || "Medium").trim() || "Medium",
    assignee: String(payload.assignee || "").trim(),
    category: String(payload.category || "").trim()
  };

  if (!clean.title) {
    return { error: "Task title is required." };
  }

  return { value: clean };
}

function validateSubmission(payload, sessionUser) {
  const clean = normalizeInvestment({
    company: payload.company,
    entity: payload.entity,
    amount: payload.amount,
    currency: payload.currency,
    stage: payload.stage,
    status: payload.status,
    owner: payload.owner,
    nextStep: payload.nextStep,
    notes: payload.notes,
    deckSummary: payload.deckSummary,
    capitalCallDate: payload.capitalCallDate,
    capitalCallAmount: payload.capitalCallAmount,
    distributionDate: payload.distributionDate,
    distributionAmount: payload.distributionAmount,
    valuationDate: payload.valuationDate,
    officialValue: payload.officialValue,
    internalValue: payload.internalValue,
    exitValue: payload.exitValue,
    ownershipPercent: payload.ownershipPercent,
    entityOwnershipPercent: payload.entityOwnershipPercent,
    ownershipNotes: payload.ownershipNotes,
    followOnCapitalAmount: payload.followOnCapitalAmount,
    followOnCapitalStatus: payload.followOnCapitalStatus,
    followOnCapitalNotes: payload.followOnCapitalNotes,
    documentLinks: payload.documentLinks,
    documents: payload.documents,
    decisionDate: payload.decisionDate,
    decisionType: payload.decisionType,
    decisionSummary: payload.decisionSummary,
    recipients: payload.recipients,
    submittedBy: sessionUser.email,
    createdAt: new Date().toISOString()
  });

  if (!clean.company) {
    return { error: "Company name is required." };
  }

  if (!clean.entity) {
    return { error: "Entity is required." };
  }

  if (!clean.status) {
    return { error: "Status is required." };
  }

  return { value: clean };
}

function validateInvestmentPatch(payload) {
  const clean = {
    company: String(payload.company || "").trim(),
    entity: String(payload.entity || "").trim(),
    amount: String(payload.amount || "").trim(),
    currency: String(payload.currency || "USD").trim() || "USD",
    stage: String(payload.stage || "").trim(),
    status: String(payload.status || "").trim(),
    owner: String(payload.owner || "").trim(),
    nextStep: String(payload.nextStep || "").trim(),
    notes: String(payload.notes || "").trim(),
    deckSummary: String(payload.deckSummary || "").trim(),
    capitalCallDate: String(payload.capitalCallDate || "").trim(),
    capitalCallAmount: String(payload.capitalCallAmount || "").trim(),
    distributionDate: String(payload.distributionDate || "").trim(),
    distributionAmount: String(payload.distributionAmount || "").trim(),
    valuationDate: String(payload.valuationDate || "").trim(),
    officialValue: String(payload.officialValue || "").trim(),
    internalValue: String(payload.internalValue || "").trim(),
    exitValue: String(payload.exitValue || "").trim(),
    ownershipPercent: String(payload.ownershipPercent || "").trim(),
    entityOwnershipPercent: String(payload.entityOwnershipPercent || "").trim(),
    ownershipNotes: String(payload.ownershipNotes || "").trim(),
    followOnCapitalAmount: String(payload.followOnCapitalAmount || "").trim(),
    followOnCapitalStatus: String(payload.followOnCapitalStatus || "").trim(),
    followOnCapitalNotes: String(payload.followOnCapitalNotes || "").trim(),
    documentLinks: String(payload.documentLinks || "").trim(),
    documents: normalizeDocuments(payload.documents),
    decisionDate: String(payload.decisionDate || "").trim(),
    decisionType: String(payload.decisionType || "").trim(),
    decisionSummary: String(payload.decisionSummary || "").trim(),
    recipients: Array.isArray(payload.recipients)
      ? payload.recipients.map((value) => String(value).trim()).filter(Boolean)
      : []
  };

  if (!clean.company) {
    return { error: "Company name is required." };
  }

  if (!clean.entity) {
    return { error: "Entity is required." };
  }

  if (!clean.status) {
    return { error: "Status is required." };
  }

  return { value: clean };
}

function requireAuth(request, response) {
  const user = getSessionUser(request);
  if (!user) {
    sendJson(response, 401, { error: "Please sign in first." });
    return null;
  }

  return user;
}

function requireEditor(request, response) {
  const user = requireAuth(request, response);
  if (!user) {
    return null;
  }

  if (!canEdit(user)) {
    sendJson(response, 403, { error: "Your account has view-only access." });
    return null;
  }

  return user;
}

process.on("uncaughtException", (error) => {
  console.error("Uncaught exception:", error);
});

process.on("unhandledRejection", (error) => {
  console.error("Unhandled rejection:", error);
});

const server = http.createServer(async (request, response) => {
  const rawPath = request.url || "/";

  if (request.method === "GET" && rawPath.startsWith("/api/health")) {
    sendJson(response, 200, { ok: true });
    return;
  }

  console.log(`[request] ${request.method} ${rawPath}`);

  const url = new URL(rawPath, `http://${request.headers.host || "localhost"}`);

  if (request.method === "GET" && url.pathname === "/api/health") {
    sendJson(response, 200, { ok: true });
    return;
  }

  if (request.method === "GET" && url.pathname === "/api/config") {
    const user = getSessionUser(request);
    sendJson(response, 200, {
      defaultRecipients: DEFAULT_RECIPIENTS,
      emailConfigured: Boolean(process.env.RESEND_API_KEY && process.env.FROM_EMAIL),
      aiConfigured: Boolean(process.env.OPENAI_API_KEY),
      entities: INVESTMENT_ENTITIES,
      familyOfficeWorkbookAvailable: fs.existsSync(FAMILY_OFFICE_WORKBOOK_FILE),
      authConfigured: Boolean(
        (process.env.TEAM_PASSWORD || Object.keys(TEAM_USERS).length > 0) && process.env.SESSION_SECRET
      ),
      authMode: Object.keys(TEAM_USERS).length > 0 ? "individual" : "shared",
      teamUserCount: Object.keys(TEAM_USERS).length,
      canEdit: canEdit(user),
      user
    });
    return;
  }

  if (request.method === "POST" && url.pathname === "/api/session") {
    try {
      const payload = await parseRequestBody(request);
      const validation = validateLogin(payload);

      if (validation.error) {
        sendJson(response, 401, { error: validation.error });
        return;
      }

      const expiresAt = Date.now() + SESSION_DURATION_MS;
      const sessionValue = signSession({
        email: validation.value,
        name: validation.value,
        role: validation.role || "editor",
        expiresAt
      });

      sendJson(
        response,
        200,
        {
          message: "Signed in.",
          user: {
            email: validation.value,
            name: validation.value,
            role: validation.role || "editor"
          }
        },
        {
          "Set-Cookie": createSessionCookie(sessionValue, expiresAt)
        }
      );
      return;
    } catch (error) {
      sendJson(response, 500, { error: error.message || "Unexpected server error." });
      return;
    }
  }

  if (request.method === "DELETE" && url.pathname === "/api/session") {
    sendJson(
      response,
      200,
      { message: "Signed out." },
      {
        "Set-Cookie": clearSessionCookie()
      }
    );
    return;
  }

  if (request.method === "GET" && url.pathname === "/api/investments") {
    const user = requireAuth(request, response);
    if (!user) {
      return;
    }

    const investments = readInvestments();
    const tasks = readTasks();
    sendJson(response, 200, {
      investments,
      companies: buildCompanyRecords(investments, tasks),
      user
    });
    return;
  }

  if (request.method === "GET" && url.pathname === "/api/tasks") {
    const user = requireAuth(request, response);
    if (!user) {
      return;
    }

    sendJson(response, 200, { tasks: readTasks(), user });
    return;
  }

  if (request.method === "GET" && url.pathname === "/api/investments.csv") {
    const user = requireAuth(request, response);
    if (!user) {
      return;
    }

    const csv = buildInvestmentsCsv(readInvestments());
    response.writeHead(200, {
      "Content-Type": "text/csv; charset=utf-8",
      "Content-Disposition": 'attachment; filename="investment-updates.csv"'
    });
    response.end(csv);
    return;
  }

  if (request.method === "GET" && url.pathname === "/api/investments.xlsx") {
    const user = requireAuth(request, response);
    if (!user) {
      return;
    }

    try {
      const workbookBuffer = buildInvestmentsWorkbookBuffer(readInvestments());
      response.writeHead(200, {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": 'attachment; filename="investment-updates.xlsx"'
      });
      response.end(workbookBuffer);
      return;
    } catch (error) {
      sendJson(response, 500, { error: error.message || "Excel export failed." });
      return;
    }
  }

  if (request.method === "GET" && url.pathname === "/api/family-office-workbook.xlsx") {
    const user = requireAuth(request, response);
    if (!user) {
      return;
    }

    if (!fs.existsSync(FAMILY_OFFICE_WORKBOOK_FILE)) {
      sendJson(response, 404, {
        error: "The family office workbook template is not available on the server yet."
      });
      return;
    }

    response.writeHead(200, {
      "Content-Type":
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition":
        'attachment; filename="family-office-private-investment-tracker.xlsx"'
    });
    fs.createReadStream(FAMILY_OFFICE_WORKBOOK_FILE).pipe(response);
    return;
  }

  if (request.method === "POST" && url.pathname === "/api/import-workbook") {
    const user = requireEditor(request, response);
    if (!user) {
      return;
    }

    try {
      const payload = await parseRequestBody(request);
      const filename = String(payload.filename || "").trim();
      const fileData = String(payload.fileData || "").trim();

      if (!filename.match(/\.(xlsx|xlsm|xls)$/i)) {
        sendJson(response, 400, {
          error: "Please upload an Excel workbook (.xlsx, .xlsm, or .xls)."
        });
        return;
      }

      if (!fileData) {
        sendJson(response, 400, { error: "Workbook file data is required." });
        return;
      }

      const buffer = Buffer.from(fileData, "base64");
      const result = importWorkbookIntoInvestments(buffer, user, filename);
      sendJson(response, 200, {
        message: `Imported ${result.importedCount} update${result.importedCount === 1 ? "" : "s"} from workbook.`,
        ...result
      });
      return;
    } catch (error) {
      sendJson(response, 500, { error: error.message || "Workbook import failed." });
      return;
    }
  }

  if (request.method === "POST" && url.pathname === "/api/upload-document") {
    const user = requireEditor(request, response);
    if (!user) {
      return;
    }

    try {
      const payload = await parseRequestBody(request);
      const filename = String(payload.filename || "").trim();
      const fileData = String(payload.fileData || "").trim();

      if (!filename) {
        sendJson(response, 400, { error: "Document filename is required." });
        return;
      }

      if (!fileData) {
        sendJson(response, 400, { error: "Document file data is required." });
        return;
      }

      ensureDataFile();
      const storedName = safeFilename(filename);
      const filePath = path.join(UPLOADS_DIR, storedName);
      fs.writeFileSync(filePath, Buffer.from(fileData, "base64"));

      sendJson(response, 201, {
        document: {
          id: makeId(),
          name: filename,
          storedName,
          url: `/uploads/${storedName}`,
          uploadedAt: new Date().toISOString(),
          uploadedBy: user.email
        }
      });
      return;
    } catch (error) {
      sendJson(response, 500, { error: error.message || "Document upload failed." });
      return;
    }
  }

  if (request.method === "POST" && url.pathname === "/api/investments") {
    const user = requireEditor(request, response);
    if (!user) {
      return;
    }

    try {
      const payload = await parseRequestBody(request);
      const validation = validateSubmission(payload, user);

      if (validation.error) {
        sendJson(response, 400, { error: validation.error });
        return;
      }

      const entry = validation.value;
      const recipients = entry.recipients.length > 0 ? entry.recipients : DEFAULT_RECIPIENTS;
      const completeEntry = {
        ...entry,
        recipients
      };
      const summary = buildSummary(completeEntry);

      saveInvestment(completeEntry);
      const emailResult = await sendEmail(summary, recipients);

      sendJson(response, 201, {
        message: "Investment update saved.",
        email: emailResult,
        summary
      });
      return;
    } catch (error) {
      sendJson(response, 500, {
        error: error.message || "Unexpected server error."
      });
      return;
    }
  }

  if (request.method === "PATCH" && url.pathname.startsWith("/api/investments/")) {
    const user = requireEditor(request, response);
    if (!user) {
      return;
    }

    try {
      const investmentId = url.pathname.split("/").pop();
      const payload = await parseRequestBody(request);
      const validation = validateInvestmentPatch(payload);

      if (validation.error) {
        sendJson(response, 400, { error: validation.error });
        return;
      }

      const updated = updateInvestment(investmentId, validation.value);
      if (!updated) {
        sendJson(response, 404, { error: "Investment update not found." });
        return;
      }

      sendJson(response, 200, { investment: updated });
      return;
    } catch (error) {
      sendJson(response, 500, { error: error.message || "Update failed." });
      return;
    }
  }

  if (request.method === "DELETE" && url.pathname.startsWith("/api/investments/")) {
    const user = requireEditor(request, response);
    if (!user) {
      return;
    }

    const investmentId = url.pathname.split("/").pop();
    const deleted = deleteInvestment(investmentId);

    if (!deleted) {
      sendJson(response, 404, { error: "Investment update not found." });
      return;
    }

    sendJson(response, 200, { message: "Investment update deleted." });
    return;
  }

  if (request.method === "POST" && url.pathname === "/api/summarize-deck") {
    const user = requireAuth(request, response);
    if (!user) {
      return;
    }

    try {
      const payload = await parseRequestBody(request);
      const filename = String(payload.filename || "").trim();
      const fileData = String(payload.fileData || "").trim();
      const company = String(payload.company || "").trim();
      const stage = String(payload.stage || "").trim();

      if (!filename.toLowerCase().endsWith(".pdf")) {
        sendJson(response, 400, { error: "Please upload a PDF deck." });
        return;
      }

      if (!fileData) {
        sendJson(response, 400, { error: "Deck file data is required." });
        return;
      }

      const summary = await summarizeDeck({ filename, fileData, company, stage });
      sendJson(response, 200, { summary });
      return;
    } catch (error) {
      sendJson(response, 500, { error: error.message || "Deck summarization failed." });
      return;
    }
  }

  if (request.method === "POST" && url.pathname === "/api/summarize-email") {
    const user = requireAuth(request, response);
    if (!user) {
      return;
    }

    try {
      const payload = await parseRequestBody(request);
      const emailText = String(payload.emailText || "").trim();
      const company = String(payload.company || "").trim();
      const stage = String(payload.stage || "").trim();

      if (!emailText) {
        sendJson(response, 400, { error: "Paste email content first." });
        return;
      }

      const summary = await summarizeEmail({ emailText, company, stage });
      sendJson(response, 200, { summary });
      return;
    } catch (error) {
      sendJson(response, 500, { error: error.message || "Email summarization failed." });
      return;
    }
  }

  if (request.method === "POST" && url.pathname === "/api/tasks") {
    const user = requireEditor(request, response);
    if (!user) {
      return;
    }

    try {
      const payload = await parseRequestBody(request);
      const validation = validateTaskSubmission(payload, user);

      if (validation.error) {
        sendJson(response, 400, { error: validation.error });
        return;
      }

      const task = saveTask(validation.value);
      sendJson(response, 201, { task });
      return;
    } catch (error) {
      sendJson(response, 500, { error: error.message || "Task creation failed." });
      return;
    }
  }

  if (request.method === "PATCH" && url.pathname.startsWith("/api/tasks/")) {
    const user = requireEditor(request, response);
    if (!user) {
      return;
    }

    try {
      const taskId = url.pathname.split("/").pop();
      const payload = await parseRequestBody(request);
      const validation = validateTaskPatch(payload);

      if (validation.error) {
        sendJson(response, 400, { error: validation.error });
        return;
      }

      const updated = updateTask(taskId, validation.value);
      if (!updated) {
        sendJson(response, 404, { error: "Task not found." });
        return;
      }

      sendJson(response, 200, { task: updated });
      return;
    } catch (error) {
      sendJson(response, 500, { error: error.message || "Task update failed." });
      return;
    }
  }

  if (request.method === "DELETE" && url.pathname.startsWith("/api/tasks/")) {
    const user = requireEditor(request, response);
    if (!user) {
      return;
    }

    const taskId = url.pathname.split("/").pop();
    const deleted = deleteTask(taskId);

    if (!deleted) {
      sendJson(response, 404, { error: "Task not found." });
      return;
    }

    sendJson(response, 200, { message: "Task deleted." });
    return;
  }

  if (request.method === "GET") {
    if (url.pathname.startsWith("/uploads/")) {
      const requestedUpload = path
        .normalize(url.pathname.replace(/^\/uploads\//, ""))
        .replace(/^(\.\.[/\\])+/, "");
      const filePath = path.join(UPLOADS_DIR, requestedUpload);

      if (!filePath.startsWith(UPLOADS_DIR)) {
        sendText(response, 403, "Forbidden");
        return;
      }

      fs.readFile(filePath, (error, data) => {
        if (error) {
          sendText(response, 404, "Not found");
          return;
        }

        response.writeHead(200, {
          "Content-Type": getContentType(filePath)
        });
        response.end(data);
      });
      return;
    }

    const requestedPath = url.pathname === "/" ? "/index.html" : url.pathname;
    const safePath = path
      .normalize(requestedPath)
      .replace(/^[/\\]+/, "")
      .replace(/^(\.\.[/\\])+/, "");
    const filePath = path.join(PUBLIC_DIR, safePath);

    if (!filePath.startsWith(PUBLIC_DIR)) {
      sendText(response, 403, "Forbidden");
      return;
    }

    serveStaticFile(response, filePath);
    return;
  }

  sendText(response, 405, "Method not allowed");
});

server.on("clientError", (error, socket) => {
  console.error("Client error:", error.message);
  socket.end("HTTP/1.1 400 Bad Request\r\n\r\n");
});

server.listen({ port: PORT, host: "::", ipv6Only: false }, () => {
  ensureDataFile();
  console.log(`Investment update app running at http://localhost:${PORT}`);
});
