const test = require("node:test");
const assert = require("node:assert/strict");
const {
  MAX_ATTACHMENT_MESSAGE_BYTES,
  MAX_ATTACHMENT_PDF_BYTES,
  MAX_ATTACHMENT_RUN_BYTES,
  MAX_ATTACHMENTS_PER_MESSAGE,
  createMicrosoftGraphMailService
} = require("../services/microsoftGraphMailService");

function graphResponse(payload, status = 200, headers = {}) {
  return {
    ok: status >= 200 && status < 300,
    status,
    headers: {
      get: (name) => headers[String(name).toLowerCase()] || ""
    },
    text: async () => JSON.stringify(payload)
  };
}

function createFetchMock(responses) {
  const calls = [];
  const fetchImpl = async (url, options = {}) => {
    calls.push({ url: String(url), options });
    const next = responses.shift();
    if (!next) {
      throw new Error(`Unexpected fetch call: ${url}`);
    }
    return typeof next === "function" ? next(url, options) : next;
  };
  fetchImpl.calls = calls;
  return fetchImpl;
}

test("Microsoft Graph mail service uses client credentials and preserves supported non-inline attachments", async () => {
  const fetchImpl = createFetchMock([
    graphResponse({ access_token: "token-value" }),
    graphResponse({
      value: [
        { id: "folder-1", displayName: "AI Investment Updates" },
        { id: "other-folder", displayName: "Inbox" }
      ]
    }),
    graphResponse({
      value: [
        {
          id: "message-1",
          internetMessageId: "<message-1@example.test>",
          conversationId: "conversation-1",
          subject: "Healing Innovations update",
          from: { emailAddress: { address: "sender@example.test", name: "Sender" } },
          receivedDateTime: "2026-08-28T12:00:00Z",
          hasAttachments: true,
          body: { contentType: "html", content: "<p>Quarterly investment update.</p>" },
          bodyPreview: "Quarterly investment update."
        }
      ]
    }),
    graphResponse({
      value: [
        {
          id: "attachment-1",
          name: "Board Deck.pdf",
          contentType: "application/pdf",
          size: 123,
          isInline: false,
          contentBytes: Buffer.from("pdf").toString("base64")
        },
        {
          id: "attachment-2",
          name: "logo.png",
          contentType: "image/png",
          size: 5,
          isInline: true
        },
        {
          id: "attachment-3",
          name: "notes.docx",
          contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          size: 55,
          isInline: false
        }
      ]
    }),
    graphResponse({
      id: "attachment-1",
      name: "Board Deck.pdf",
      contentType: "application/pdf",
      size: 123,
      isInline: false,
      contentBytes: Buffer.from("pdf").toString("base64")
    }),
    graphResponse({
      id: "attachment-3",
      name: "notes.docx",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      size: 55,
      isInline: false,
      contentBytes: Buffer.from("docx").toString("base64")
    })
  ]);
  const service = createMicrosoftGraphMailService({
    tenantId: "tenant-id",
    clientId: "client-id",
    clientSecret: "client-secret",
    mailboxUser: "updates@example.test",
    folderName: "AI Investment Updates",
    fetchImpl,
    graphBaseUrl: "https://graph.test/v1.0",
    tokenBaseUrl: "https://login.test"
  });

  const result = await service.fetchIntakeMessages();

  assert.equal(result.mailbox, "updates@example.test");
  assert.equal(result.folder.displayName, "AI Investment Updates");
  assert.equal(result.messages.length, 1);
  assert.equal(result.messages[0].pdfAttachments.length, 1);
  assert.equal(result.messages[0].pdfAttachments[0].name, "Board Deck.pdf");
  assert.equal(result.messages[0].attachments.length, 2);
  assert.equal(result.messages[0].attachments[1].name, "notes.docx");
  assert.equal(result.messages[0].attachments[1].extractionStatus, "not-parsed");
  assert.equal(result.messages[0].skippedAttachments.length, 1);
  assert.equal(result.messages[0].unresolvedAttachments.length, 0);
  assert.match(fetchImpl.calls[0].url, /\/tenant-id\/oauth2\/v2\.0\/token$/);
  assert.match(fetchImpl.calls[0].options.body, /grant_type=client_credentials/);
  assert.match(fetchImpl.calls[0].options.body, /scope=https%3A%2F%2Fgraph\.microsoft\.com%2F\.default/);
  assert.equal(
    fetchImpl.calls.some((call) => /\/users\/updates%40example\.test\?\$select=id/.test(call.url)),
    false
  );
  assert.match(fetchImpl.calls[1].url, /\/users\/updates%40example\.test\/mailFolders\?\$top=100$/);
  assert.match(fetchImpl.calls[2].url, /mailFolders\/folder-1\/messages/);
  assert.match(fetchImpl.calls[3].url, /\/users\/updates%40example\.test\/messages\/message-1\/attachments\?\$top=21&\$select=/);
  assert.equal(fetchImpl.calls.every((call) => !call.options.method || call.options.method === "GET" || call.url.includes("oauth2")), true);
});

test("Microsoft Graph mail APIs address configured mailbox UPN directly without directory lookup", async () => {
  const fetchImpl = createFetchMock([
    graphResponse({ access_token: "token-value" }),
    graphResponse({
      value: [{ id: "folder-1", displayName: "AI Investment Updates" }]
    }),
    graphResponse({
      value: [
        {
          id: "message-1",
          internetMessageId: "<message-1@example.test>",
          subject: "Update",
          from: { emailAddress: { address: "sender@example.test" } },
          receivedDateTime: "2026-08-28T12:00:00Z",
          hasAttachments: true,
          body: { contentType: "text", content: "Update" }
        }
      ]
    }),
    graphResponse({
      value: [
        {
          id: "attachment-1",
          name: "Deck.pdf",
          contentType: "application/pdf",
          isInline: false,
          contentBytes: Buffer.from("pdf").toString("base64")
        }
      ]
    }),
    graphResponse({
      id: "attachment-1",
      name: "Deck.pdf",
      contentType: "application/pdf",
      isInline: false,
      contentBytes: Buffer.from("pdf").toString("base64")
    })
  ]);
  const service = createMicrosoftGraphMailService({
    tenantId: "tenant-id",
    clientId: "client-id",
    clientSecret: "client-secret",
    mailboxUser: "updates+ai@example.test",
    folderName: "AI Investment Updates",
    fetchImpl,
    graphBaseUrl: "https://graph.test/v1.0",
    tokenBaseUrl: "https://login.test"
  });

  await service.fetchIntakeMessages();

  const urls = fetchImpl.calls.map((call) => call.url);
  assert.equal(urls.some((url) => /\/users\/[^/]+\?\$select=id/.test(url)), false);
  assert.match(urls[1], /\/users\/updates%2Bai%40example\.test\/mailFolders\?\$top=100$/);
  assert.match(urls[2], /\/users\/updates%2Bai%40example\.test\/mailFolders\/folder-1\/messages/);
  assert.match(urls[3], /\/users\/updates%2Bai%40example\.test\/messages\/message-1\/attachments\?\$top=21&\$select=/);
});

test("Microsoft Graph safe config status does not expose client secret", () => {
  const service = createMicrosoftGraphMailService({
    tenantId: "tenant-id",
    clientId: "client-id",
    clientSecret: "client-secret",
    mailboxUser: "updates@example.test",
    folderName: "AI Investment Updates",
    fetchImpl: async () => graphResponse({})
  });

  const status = service.getSafeConfigStatus();

  assert.equal(status.configured, true);
  assert.equal(status.mailboxUser, "updates@example.test");
  assert.equal(status.folderName, "AI Investment Updates");
  assert.equal(Object.prototype.hasOwnProperty.call(status, "clientSecret"), false);
  assert.equal(JSON.stringify(status).includes("client-secret"), false);
});

test("read-only preview uses the production folder window and lists attachment metadata without downloading content", async () => {
  const fetchImpl = createFetchMock([
    graphResponse({ access_token: "token-value" }),
    graphResponse({ value: [{ id: "folder-1", displayName: "AI Investment Updates" }] }),
    graphResponse({ value: [{
      id: "message-1",
      internetMessageId: "<message-1@example.test>",
      subject: "Attainable Living update",
      from: { emailAddress: { address: "founder@attainableliving.example" } },
      receivedDateTime: "2026-09-18T12:00:00Z",
      hasAttachments: true,
      body: { contentType: "text", content: "Quarterly update" }
    }] }),
    graphResponse({ value: [
      { id: "deck", name: "Attainable Living.pdf", contentType: "application/pdf", size: 1024, isInline: false },
      { id: "large", name: "Large Model.xlsx", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", size: 11 * 1024 * 1024, isInline: false }
    ] })
  ]);
  const service = createMicrosoftGraphMailService({
    tenantId: "tenant",
    clientId: "client",
    clientSecret: "secret",
    mailboxUser: "updates@example.test",
    folderName: "AI Investment Updates",
    maxMessagesPerRun: 7,
    fetchImpl,
    graphBaseUrl: "https://graph.test/v1.0",
    tokenBaseUrl: "https://login.test"
  });

  const result = await service.fetchIntakePreviewMessages();

  assert.equal(result.maxMessagesPerRun, 7);
  assert.equal(result.messages[0].attachmentCount, 2);
  assert.equal(result.messages[0].attachmentPlan[0].disposition, "eligible");
  assert.equal(result.messages[0].attachmentPlan[1].disposition, "blocked");
  assert.match(result.messages[0].attachmentPlan[1].reason, /10 MB/);
  assert.equal(fetchImpl.calls.length, 4);
  assert.match(fetchImpl.calls[2].url, /\$top=7&\$orderby=receivedDateTime desc/);
  assert.equal(fetchImpl.calls.some((call) => /attachments\/deck$/.test(call.url)), false);
  assert.equal(JSON.stringify(result).includes("contentBytes"), false);
});

test("Microsoft Graph throttling returns a safe retry message", async () => {
  const fetchImpl = createFetchMock([
    graphResponse({ error: { message: "Too many requests" } }, 429, { "retry-after": "30" })
  ]);
  const service = createMicrosoftGraphMailService({
    tenantId: "tenant-id",
    clientId: "client-id",
    clientSecret: "client-secret",
    mailboxUser: "updates@example.test",
    fetchImpl,
    tokenBaseUrl: "https://login.test"
  });

  await assert.rejects(
    () => service.getAccessToken(),
    /Microsoft Graph request was throttled\. Retry after 30 seconds\./
  );
});

test("PDF attachment limits accept 25 MB while keeping conservative message, run, and count caps", async () => {
  const pdfBytes = Buffer.concat([
    Buffer.from("%PDF-", "latin1"),
    Buffer.alloc(MAX_ATTACHMENT_PDF_BYTES - 5)
  ]);
  const fetchImpl = createFetchMock([
    graphResponse({ access_token: "token-value" }),
    graphResponse({ value: [{ id: "folder-1", displayName: "AI Investment Updates" }] }),
    graphResponse({ value: [{
      id: "message-1", internetMessageId: "<message-1@example.test>", subject: "Opportunity",
      from: { emailAddress: { address: "sender@example.test" } }, receivedDateTime: "2026-09-16T12:00:00Z",
      hasAttachments: true, body: { contentType: "text", content: "Opportunity" }
    }] }),
    graphResponse({ value: [
      { id: "deck", name: "InvestorDeck.pdf", contentType: "application/pdf", size: MAX_ATTACHMENT_PDF_BYTES, isInline: false }
    ] }),
    graphResponse({
      id: "deck", name: "InvestorDeck.pdf", contentType: "application/pdf",
      size: MAX_ATTACHMENT_PDF_BYTES, isInline: false, contentBytes: pdfBytes.toString("base64")
    })
  ]);
  const service = createMicrosoftGraphMailService({
    tenantId: "tenant", clientId: "client", clientSecret: "secret", mailboxUser: "updates@example.test",
    fetchImpl, graphBaseUrl: "https://graph.test/v1.0", tokenBaseUrl: "https://login.test"
  });

  const result = await service.fetchIntakeMessages();

  assert.equal(result.messages[0].pdfAttachments.length, 1);
  assert.equal(result.messages[0].pdfAttachments[0].size, MAX_ATTACHMENT_PDF_BYTES);
  assert.equal(result.messages[0].unresolvedAttachments.length, 0);
  assert.equal(MAX_ATTACHMENT_MESSAGE_BYTES, 30 * 1024 * 1024);
  assert.equal(MAX_ATTACHMENT_RUN_BYTES, 50 * 1024 * 1024);
  assert.equal(MAX_ATTACHMENTS_PER_MESSAGE, 20);
});

test("oversized and unsupported attachments are visible as unresolved without download", async () => {
  const fetchImpl = createFetchMock([
    graphResponse({ access_token: "token-value" }),
    graphResponse({ value: [{ id: "folder-1", displayName: "AI Investment Updates" }] }),
    graphResponse({ value: [{
      id: "message-1", internetMessageId: "<message-1@example.test>", subject: "Opportunity",
      from: { emailAddress: { address: "sender@example.test" } }, receivedDateTime: "2026-09-16T12:00:00Z",
      hasAttachments: true, body: { contentType: "text", content: "Opportunity" }
    }] }),
    graphResponse({ value: [
      { id: "large", name: "LargeDeck.pdf", contentType: "application/pdf", size: 26 * 1024 * 1024, isInline: false },
      { id: "unsafe", name: "script.exe", contentType: "application/octet-stream", size: 100, isInline: false }
    ] })
  ]);
  const service = createMicrosoftGraphMailService({
    tenantId: "tenant", clientId: "client", clientSecret: "secret", mailboxUser: "updates@example.test",
    fetchImpl, graphBaseUrl: "https://graph.test/v1.0", tokenBaseUrl: "https://login.test"
  });
  const result = await service.fetchIntakeMessages();
  assert.equal(result.messages[0].attachments.length, 0);
  assert.equal(result.messages[0].unresolvedAttachments.length, 2);
  assert.match(result.messages[0].unresolvedAttachments[0].reason, /25 MB/);
  assert.match(result.messages[0].unresolvedAttachments[1].reason, /Unsupported/);
  assert.equal(fetchImpl.calls.length, 4);
});
