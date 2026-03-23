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
const PUBLIC_DIR = path.join(__dirname, "public");
const SESSION_COOKIE = "investment_team_session";
const SESSION_SECRET = process.env.SESSION_SECRET || "change-me-before-production";
const SESSION_DURATION_MS = 1000 * 60 * 60 * 24 * 7;
const MAX_BODY_SIZE_BYTES = 20 * 1024 * 1024;
const OPENAI_MODEL = process.env.OPENAI_MODEL || "gpt-4o-mini";

const DEFAULT_RECIPIENTS = splitCsv(process.env.TEAM_EMAILS || "");

function splitCsv(value) {
  return String(value)
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

function parseTeamUsers(value) {
  return splitCsv(value).reduce((users, entry) => {
    const separatorIndex = entry.indexOf(":");
    if (separatorIndex === -1) {
      return users;
    }

    const email = entry.slice(0, separatorIndex).trim().toLowerCase();
    const password = entry.slice(separatorIndex + 1).trim();

    if (!email || !password) {
      return users;
    }

    users[email] = {
      email,
      password
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

  if (!fs.existsSync(DATA_FILE)) {
    fs.writeFileSync(DATA_FILE, "[]\n", "utf8");
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

function findLatestByCompanyKey(companyKey, investments) {
  if (!companyKey) {
    return null;
  }

  return investments.find((investment) => investment.companyKey === companyKey) || null;
}

function normalizeInvestment(entry) {
  return {
    id: entry.id || makeId(),
    company: String(entry.company || "").trim(),
    companyKey: entry.companyKey || normalizeCompanyKey(entry.company),
    amount: String(entry.amount || "").trim(),
    currency: String(entry.currency || "USD").trim() || "USD",
    stage: String(entry.stage || "").trim(),
    status: String(entry.status || "").trim(),
    owner: String(entry.owner || "").trim(),
    nextStep: String(entry.nextStep || "").trim(),
    notes: String(entry.notes || "").trim(),
    recipients: Array.isArray(entry.recipients)
      ? entry.recipients.map((value) => String(value).trim()).filter(Boolean)
      : [],
    submittedBy: String(entry.submittedBy || "").trim(),
    createdAt: String(entry.createdAt || new Date().toISOString())
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
    name: session.name || session.email
  };
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

function csvEscape(value) {
  const text = String(value ?? "");
  if (/[",\n]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}

function buildInvestmentsCsv(investments) {
  const headers = [
    "Company",
    "Amount",
    "Currency",
    "Stage",
    "Status",
    "Owner",
    "Next Step",
    "Notes",
    "Submitted By",
    "Recipients",
    "Created At"
  ];

  const rows = investments.map((investment) =>
    [
      investment.company,
      investment.amount,
      investment.currency,
      investment.stage,
      investment.status,
      investment.owner,
      investment.nextStep,
      investment.notes,
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
    Company: investment.company,
    Amount: investment.amount,
    Currency: investment.currency,
    Stage: investment.stage,
    Status: investment.status,
    Owner: investment.owner,
    "Next Step": investment.nextStep,
    Notes: investment.notes,
    "Submitted By": investment.submittedBy,
    Recipients: Array.isArray(investment.recipients) ? investment.recipients.join(", ") : "",
    "Created At": investment.createdAt
  }));

  const sheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "Investment Updates");
  return XLSX.write(workbook, { bookType: "xlsx", type: "buffer" });
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
  const companyLine = entry.company || "Unnamed investment";
  const amountLine = formatAmount(entry);
  const stageLine = entry.stage || "Stage not specified";
  const ownerLine = entry.owner || "No owner listed";
  const nextStepLine = entry.nextStep || "No next step provided";
  const noteLine = entry.notes || "No additional notes.";

  const subject = `${entry.status}: ${companyLine}`;

  const text = [
    `${entry.status} investment update`,
    "",
    `Company: ${companyLine}`,
    `Amount: ${amountLine}`,
    `Stage: ${stageLine}`,
    `Owner: ${ownerLine}`,
    `Submitted by: ${entry.submittedBy}`,
    `Next step: ${nextStepLine}`,
    "",
    "Summary",
    noteLine,
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
          <tr><td style="padding: 8px 0; font-weight: bold; width: 140px;">Amount</td><td>${escapeHtml(amountLine)}</td></tr>
          <tr><td style="padding: 8px 0; font-weight: bold;">Stage</td><td>${escapeHtml(stageLine)}</td></tr>
          <tr><td style="padding: 8px 0; font-weight: bold;">Owner</td><td>${escapeHtml(ownerLine)}</td></tr>
          <tr><td style="padding: 8px 0; font-weight: bold;">Submitted by</td><td>${escapeHtml(entry.submittedBy)}</td></tr>
          <tr><td style="padding: 8px 0; font-weight: bold;">Next step</td><td>${escapeHtml(nextStepLine)}</td></tr>
        </table>
      </div>
      <div style="padding: 20px 8px 0;">
        <h3 style="margin-bottom: 8px;">Summary</h3>
        <p style="white-space: pre-wrap; margin-top: 0;">${escapeHtml(noteLine)}</p>
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

    return { value: email };
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

  return { value: email };
}

function validateSubmission(payload, sessionUser) {
  const clean = normalizeInvestment({
    company: payload.company,
    amount: payload.amount,
    currency: payload.currency,
    stage: payload.stage,
    status: payload.status,
    owner: payload.owner,
    nextStep: payload.nextStep,
    notes: payload.notes,
    recipients: payload.recipients,
    submittedBy: sessionUser.email,
    createdAt: new Date().toISOString()
  });

  if (!clean.company) {
    return { error: "Company name is required." };
  }

  if (!clean.status) {
    return { error: "Status is required." };
  }

  return { value: clean };
}

function validateInvestmentPatch(payload) {
  const clean = {
    company: String(payload.company || "").trim(),
    amount: String(payload.amount || "").trim(),
    currency: String(payload.currency || "USD").trim() || "USD",
    stage: String(payload.stage || "").trim(),
    status: String(payload.status || "").trim(),
    owner: String(payload.owner || "").trim(),
    nextStep: String(payload.nextStep || "").trim(),
    notes: String(payload.notes || "").trim(),
    recipients: Array.isArray(payload.recipients)
      ? payload.recipients.map((value) => String(value).trim()).filter(Boolean)
      : []
  };

  if (!clean.company) {
    return { error: "Company name is required." };
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
      authConfigured: Boolean(
        (process.env.TEAM_PASSWORD || Object.keys(TEAM_USERS).length > 0) && process.env.SESSION_SECRET
      ),
      authMode: Object.keys(TEAM_USERS).length > 0 ? "individual" : "shared",
      teamUserCount: Object.keys(TEAM_USERS).length,
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
        expiresAt
      });

      sendJson(
        response,
        200,
        {
          message: "Signed in.",
          user: {
            email: validation.value,
            name: validation.value
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

    sendJson(response, 200, { investments: readInvestments(), user });
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

  if (request.method === "POST" && url.pathname === "/api/investments") {
    const user = requireAuth(request, response);
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
    const user = requireAuth(request, response);
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
    const user = requireAuth(request, response);
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

  if (request.method === "GET") {
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
