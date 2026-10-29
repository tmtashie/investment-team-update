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
          size: 3,
          isInline: false,
          contentBytes: Buffer.from("pdf").toString("base64")
        }
      ]
    }),
    graphResponse({
      id: "attachment-1",
      name: "Deck.pdf",
      contentType: "application/pdf",
      size: 3,
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

test("invalid attachment sizes fail closed without changing byte budgets or downloading content", async () => {
  const graphResponses = () => [
    graphResponse({ access_token: "token-value" }),
    graphResponse({ value: [{ id: "folder-1", displayName: "AI Investment Updates" }] }),
    graphResponse({ value: [{
      id: "message-1", internetMessageId: "<message-1@example.test>", subject: "Opportunity",
      from: { emailAddress: { address: "sender@example.test" } }, receivedDateTime: "2026-09-16T12:00:00Z",
      hasAttachments: true, body: { contentType: "text", content: "Opportunity" }
    }] }),
    graphResponse({ value: [
      { id: "negative", name: "Negative.pdf", contentType: "application/pdf", size: -1024, isInline: false },
      { id: "missing", name: "Missing.pdf", contentType: "application/pdf", isInline: false },
      { id: "fractional", name: "Fractional.pdf", contentType: "application/pdf", size: 1.5, isInline: false },
      { id: "string", name: "String.pdf", contentType: "application/pdf", size: "1024", isInline: false },
      { id: "unsafe", name: "Unsafe.pdf", contentType: "application/pdf", size: Number.MAX_SAFE_INTEGER + 1, isInline: false }
    ] })
  ];
  const createService = (fetchImpl) => createMicrosoftGraphMailService({
    tenantId: "tenant", clientId: "client", clientSecret: "secret", mailboxUser: "updates@example.test",
    fetchImpl, graphBaseUrl: "https://graph.test/v1.0", tokenBaseUrl: "https://login.test"
  });

  const previewFetch = createFetchMock(graphResponses());
  const preview = await createService(previewFetch).previewIntakeMessages();
  assert.deepEqual(
    preview.messages[0].attachments.map((attachment) => attachment.status),
    ["unresolved", "unresolved", "unresolved", "unresolved", "unresolved"]
  );
  assert.deepEqual(
    preview.messages[0].attachments.map((attachment) => attachment.projectedCountedBytes),
    [0, 0, 0, 0, 0]
  );
  assert.deepEqual(
    preview.messages[0].attachments.map((attachment) => attachment.size),
    [null, null, null, null, null]
  );
  assert.equal(preview.messages[0].projectedMessageAttachmentBytes, 0);
  assert.equal(preview.messages[0].projectedRunAttachmentBytesAfterMessage, 0);

  const intakeFetch = createFetchMock(graphResponses());
  const intake = await createService(intakeFetch).fetchIntakeMessages();
  assert.equal(intake.messages[0].attachments.length, 0);
  assert.equal(intake.messages[0].unresolvedAttachments.length, 5);
  assert.equal(
    intake.messages[0].unresolvedAttachments.every((attachment) => /size metadata/.test(attachment.reason)),
    true
  );
  assert.equal(intakeFetch.calls.length, 4);

  const reanalysisFetch = createFetchMock([
    graphResponse({ access_token: "token-value" }),
    graphResponse({ value: [{ id: "folder-1", displayName: "AI Investment Updates" }] }),
    graphResponse({
      id: "message-1", parentFolderId: "folder-1", internetMessageId: "<message-1@example.test>",
      subject: "Opportunity", from: { emailAddress: { address: "sender@example.test" } },
      receivedDateTime: "2026-09-16T12:00:00Z", hasAttachments: true,
      body: { contentType: "text", content: "Opportunity" }
    }),
    graphResponse({ value: [
      { id: "string", name: "String.pdf", contentType: "application/pdf", size: "1024", isInline: false }
    ] })
  ]);
  const reanalysis = await createService(reanalysisFetch).fetchIntakeMessageById("message-1");
  assert.equal(reanalysis.attachments.length, 0);
  assert.equal(reanalysis.unresolvedAttachments.length, 1);
  assert.match(reanalysis.unresolvedAttachments[0].reason, /size metadata/);
  assert.equal(reanalysisFetch.calls.length, 4);
});

test("intake preview reads message and attachment metadata without downloading attachment content", async () => {
  const fetchImpl = createFetchMock([
    graphResponse({ access_token: "token-value" }),
    graphResponse({ value: [{ id: "folder-1", displayName: "AI Investment Updates" }] }),
    graphResponse({ value: [{
      id: "attainable-graph-id",
      internetMessageId: "<attainable@example.test>",
      conversationId: "conversation-1",
      subject: "Attainable Living LLC",
      from: { emailAddress: { address: "cleseberg@thesignatry.com", name: "Colby Leseberg" } },
      receivedDateTime: "2026-09-18T16:00:00Z",
      hasAttachments: true,
      body: { contentType: "text", content: "Private source content" }
    }] }),
    graphResponse({ value: [
      { id: "deck", name: "Attainable Living Deck.pdf", contentType: "application/pdf", size: 24 * 1024 * 1024, isInline: false },
      { id: "logo", name: "logo.png", contentType: "image/png", size: 1000, isInline: true }
    ] })
  ]);
  const service = createMicrosoftGraphMailService({
    tenantId: "tenant", clientId: "client", clientSecret: "secret", mailboxUser: "updates@example.test",
    fetchImpl, graphBaseUrl: "https://graph.test/v1.0", tokenBaseUrl: "https://login.test", maxMessagesPerRun: 7
  });

  const result = await service.previewIntakeMessages();

  assert.equal(fetchImpl.calls.length, 4);
  assert.equal(fetchImpl.calls.some((call) => /attachments\/deck$/.test(call.url)), false);
  assert.equal(result.maxMessagesPerRun, 7);
  assert.equal(result.messages[0].graphMessageId, "attainable-graph-id");
  assert.equal(result.messages[0].attachmentCount, 2);
  assert.equal(result.messages[0].attachments[0].status, "eligible");
  assert.equal(result.messages[0].attachments[0].projectedCountedBytes, 24 * 1024 * 1024);
  assert.equal(result.messages[0].attachments[1].status, "skipped");
  assert.equal(Object.prototype.hasOwnProperty.call(result.messages[0], "body"), false);
  assert.equal(Object.prototype.hasOwnProperty.call(result.messages[0].attachments[0], "contentBytes"), false);
});

test("exact-message reanalysis fetches by Graph ID and requires the configured intake folder", async () => {
  const fetchImpl = createFetchMock([
    graphResponse({ access_token: "token-value" }),
    graphResponse({ value: [{ id: "folder-1", displayName: "AI Investment Updates" }] }),
    graphResponse({
      id: "bep-graph-id", parentFolderId: "folder-1", internetMessageId: "<bep@example.test>",
      subject: "BEP background and teasers", from: { emailAddress: { address: "sender@example.test" } },
      receivedDateTime: "2026-09-22T12:00:00Z", hasAttachments: true,
      body: { contentType: "text", content: "Current status evidence." }
    }),
    graphResponse({ value: [{ id: "deck", name: "BEP Core Fund VIII.pdf", contentType: "application/pdf", size: 3, isInline: false }] }),
    graphResponse({ id: "deck", name: "BEP Core Fund VIII.pdf", contentType: "application/pdf", size: 3, isInline: false, contentBytes: Buffer.from("pdf").toString("base64") })
  ]);
  const service = createMicrosoftGraphMailService({
    tenantId: "tenant", clientId: "client", clientSecret: "secret", mailboxUser: "updates@example.test",
    fetchImpl, graphBaseUrl: "https://graph.test/v1.0", tokenBaseUrl: "https://login.test"
  });
  const message = await service.fetchIntakeMessageById("bep-graph-id");
  assert.equal(message.internetMessageId, "<bep@example.test>");
  assert.equal(message.pdfAttachments.length, 1);
  assert.match(fetchImpl.calls[2].url, /messages\/bep-graph-id\?\$select=/);
  assert.equal(fetchImpl.calls.every((call) => !call.options.method || call.options.method === "GET" || call.url.includes("oauth2")), true);
});

test("fractional message limits are normalized before building the Graph query", async () => {
  const fetchImpl = createFetchMock([
    graphResponse({ access_token: "token-value" }),
    graphResponse({ value: [{ id: "folder-1", displayName: "AI Investment Updates" }] }),
    graphResponse({ value: [] })
  ]);
  const service = createMicrosoftGraphMailService({
    tenantId: "tenant",
    clientId: "client",
    clientSecret: "secret",
    mailboxUser: "updates@example.test",
    maxMessagesPerRun: "2.9",
    fetchImpl,
    graphBaseUrl: "https://graph.test/v1.0",
    tokenBaseUrl: "https://login.test"
  });

  const result = await service.previewIntakeMessages();

  assert.equal(result.maxMessagesPerRun, 2);
  assert.match(fetchImpl.calls[2].url, /\$top=2&/);
  assert.equal(fetchImpl.calls[2].url.includes("$top=2.9"), false);
});


test("message limit normalization preserves defaults and configured bounds", () => {
  const cases = [
    { value: undefined, expected: 10 },
    { value: "", expected: 10 },
    { value: "not-a-number", expected: 10 },
    { value: "0", expected: 10 },
    { value: "-4.8", expected: 1 },
    { value: "0.9", expected: 1 },
    { value: "50.9", expected: 50 },
    { value: "999", expected: 50 },
    { value: "Infinity", expected: 50 },
    { value: "-Infinity", expected: 1 }
  ];

  for (const { value, expected } of cases) {
    const service = createMicrosoftGraphMailService({ maxMessagesPerRun: value });
    assert.equal(
      service.getSafeConfigStatus().maxMessagesPerRun,
      expected,
      `expected ${String(value)} to normalize to ${expected}`
    );
  }
});

test("scheduled intake pages forward from a checkpoint and preserves a trusted continuation cursor", async () => {
  const nextLink = "https://graph.test/v1.0/users/updates%40example.test/mailFolders/folder-1/messages?%24skiptoken=next-page";
  const fetchImpl = createFetchMock([
    graphResponse({ access_token: "token-value" }),
    graphResponse({ value: [{ id: "folder-1", displayName: "AI Investment Updates" }] }),
    graphResponse({
      value: [{
        id: "message-1",
        internetMessageId: "<message-1@example.test>",
        subject: "Update",
        from: { emailAddress: { address: "sender@example.test" } },
        receivedDateTime: "2026-10-29T12:01:00Z",
        hasAttachments: false,
        body: { contentType: "text", content: "Investment update" }
      }],
      "@odata.nextLink": nextLink
    })
  ]);
  const service = createMicrosoftGraphMailService({
    tenantId: "tenant", clientId: "client", clientSecret: "secret", mailboxUser: "updates@example.test",
    fetchImpl, graphBaseUrl: "https://graph.test/v1.0", tokenBaseUrl: "https://login.test",
    maxMessagesPerRun: 50
  });

  const result = await service.fetchIntakeMessages({ since: "2026-10-29T12:00:00.000Z" });

  assert.equal(result.messages.length, 1);
  assert.equal(result.continuationCursor, nextLink);
  const pageUrl = new URL(fetchImpl.calls[2].url);
  assert.equal(pageUrl.searchParams.get("$orderby"), "receivedDateTime asc");
  assert.equal(pageUrl.searchParams.get("$filter"), "receivedDateTime ge 2026-10-29T12:00:00.000Z");
  assert.equal(pageUrl.searchParams.get("$top"), "50");
});

test("scheduled intake follows only Graph continuation cursors for the configured mailbox folder", async () => {
  const continuationCursor = "https://graph.test/v1.0/users/updates%40example.test/mailFolders/folder-1/messages?%24skiptoken=next-page";
  const fetchImpl = createFetchMock([
    graphResponse({ access_token: "token-value" }),
    graphResponse({ value: [{ id: "folder-1", displayName: "AI Investment Updates" }] }),
    graphResponse({ value: [] })
  ]);
  const service = createMicrosoftGraphMailService({
    tenantId: "tenant", clientId: "client", clientSecret: "secret", mailboxUser: "updates@example.test",
    fetchImpl, graphBaseUrl: "https://graph.test/v1.0", tokenBaseUrl: "https://login.test"
  });

  const result = await service.fetchIntakeMessages({
    since: "2026-10-29T12:00:00.000Z",
    continuationCursor
  });

  assert.equal(result.messages.length, 0);
  assert.equal(fetchImpl.calls[2].url, continuationCursor);
  const untrustedFetch = createFetchMock([
    graphResponse({ access_token: "token-value" }),
    graphResponse({ value: [{ id: "folder-1", displayName: "AI Investment Updates" }] })
  ]);
  const untrustedService = createMicrosoftGraphMailService({
    tenantId: "tenant", clientId: "client", clientSecret: "secret", mailboxUser: "updates@example.test",
    fetchImpl: untrustedFetch, graphBaseUrl: "https://graph.test/v1.0", tokenBaseUrl: "https://login.test"
  });
  await assert.rejects(
    untrustedService.fetchIntakeMessages({
      since: "2026-10-29T12:00:00.000Z",
      continuationCursor: "https://attacker.example/messages?skiptoken=stolen"
    }),
    /untrusted continuation cursor/
  );
  assert.equal(untrustedFetch.calls.length, 2);
});

test("exact-message reanalysis leaves failed attachment downloads unresolved", async () => {
  const fetchImpl = createFetchMock([
    graphResponse({ access_token: "token-value" }),
    graphResponse({ value: [{ id: "folder-1", displayName: "AI Investment Updates" }] }),
    graphResponse({
      id: "message-1", parentFolderId: "folder-1", internetMessageId: "<message-1@example.test>",
      subject: "Quarterly update", from: { emailAddress: { address: "sender@example.test" } },
      receivedDateTime: "2026-10-23T12:00:00Z", hasAttachments: true,
      body: { contentType: "text", content: "Current portfolio update evidence." }
    }),
    graphResponse({
      value: [
        { id: "deck", name: "Board Deck.pdf", contentType: "application/pdf", size: 3, isInline: false },
        { id: "notes", name: "Notes.pdf", contentType: "application/pdf", size: 3, isInline: false }
      ]
    }),
    graphResponse({ error: { message: "Attachment unavailable" } }, 503),
    graphResponse({
      id: "notes", name: "Notes.pdf", contentType: "application/pdf", size: 3,
      isInline: false, contentBytes: Buffer.from("pdf").toString("base64")
    })
  ]);
  const service = createMicrosoftGraphMailService({
    tenantId: "tenant", clientId: "client", clientSecret: "secret", mailboxUser: "updates@example.test",
    fetchImpl, graphBaseUrl: "https://graph.test/v1.0", tokenBaseUrl: "https://login.test"
  });

  const message = await service.fetchIntakeMessageById("message-1");

  assert.equal(message.attachments.length, 1);
  assert.equal(message.attachments[0].name, "Notes.pdf");
  assert.equal(message.pdfAttachments.length, 1);
  assert.equal(message.unresolvedAttachments.length, 1);
  assert.equal(message.unresolvedAttachments[0].name, "Board Deck.pdf");
  assert.equal(message.unresolvedAttachments[0].contentBytes, "");
  assert.equal(message.unresolvedAttachments[0].preservationStatus, "unresolved");
  assert.equal(message.unresolvedAttachments[0].reason, "Attachment unavailable");
});

