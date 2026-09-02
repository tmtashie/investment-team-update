const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const TOKEN_BASE_URL = "https://login.microsoftonline.com";

function cleanString(value, maxLength = 2000) {
  return String(value || "").trim().slice(0, maxLength);
}

function graphError(message, statusCode, details = {}) {
  const error = new Error(message);
  error.statusCode = statusCode || 500;
  error.details = details;
  return error;
}

function parseGraphCollection(payload) {
  return payload && Array.isArray(payload.value) ? payload.value : [];
}

function isPdfAttachment(attachment) {
  const name = cleanString(attachment && attachment.name, 500).toLowerCase();
  const contentType = cleanString(attachment && attachment.contentType, 200).toLowerCase();
  return name.endsWith(".pdf") || contentType === "application/pdf";
}

function createMicrosoftGraphMailService({
  tenantId,
  clientId,
  clientSecret,
  mailboxUser,
  folderName = "AI Investment Updates",
  maxMessagesPerRun = 10,
  fetchImpl = global.fetch,
  graphBaseUrl = GRAPH_BASE_URL,
  tokenBaseUrl = TOKEN_BASE_URL
} = {}) {
  const config = {
    tenantId: cleanString(tenantId, 200),
    clientId: cleanString(clientId, 200),
    clientSecret: cleanString(clientSecret, 2000),
    mailboxUser: cleanString(mailboxUser, 320),
    folderName: cleanString(folderName, 200) || "AI Investment Updates",
    maxMessagesPerRun: Math.max(1, Math.min(50, Number(maxMessagesPerRun) || 10))
  };

  function isConfigured() {
    return Boolean(
      config.tenantId &&
      config.clientId &&
      config.clientSecret &&
      config.mailboxUser &&
      config.folderName &&
      fetchImpl
    );
  }

  function getSafeConfigStatus() {
    return {
      configured: isConfigured(),
      mailboxUser: config.mailboxUser,
      folderName: config.folderName,
      maxMessagesPerRun: config.maxMessagesPerRun
    };
  }

  async function requestJson(url, options = {}) {
    const response = await fetchImpl(url, options);
    const text = await response.text();
    let payload = null;
    if (text) {
      try {
        payload = JSON.parse(text);
      } catch (error) {
        payload = { raw: text };
      }
    }
    if (!response.ok) {
      const retryAfter = response.headers && response.headers.get
        ? response.headers.get("retry-after")
        : "";
      throw graphError(
        retryAfter
          ? `Microsoft Graph request was throttled. Retry after ${retryAfter} seconds.`
          : (payload && payload.error && payload.error.message) || "Microsoft Graph request failed.",
        response.status,
        { retryAfter }
      );
    }
    return payload || {};
  }

  async function getAccessToken() {
    if (!isConfigured()) {
      throw graphError("Microsoft 365 email intake is not configured.", 400);
    }
    const params = new URLSearchParams();
    params.set("client_id", config.clientId);
    params.set("client_secret", config.clientSecret);
    params.set("scope", "https://graph.microsoft.com/.default");
    params.set("grant_type", "client_credentials");

    const payload = await requestJson(
      `${tokenBaseUrl}/${encodeURIComponent(config.tenantId)}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: params.toString()
      }
    );
    if (!payload.access_token) {
      throw graphError("Microsoft Graph token response did not include an access token.", 502);
    }
    return payload.access_token;
  }

  async function graphGet(path, token) {
    return requestJson(`${graphBaseUrl}${path}`, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json"
      }
    });
  }

  function mailboxMailApiPath() {
    return `/users/${encodeURIComponent(config.mailboxUser)}`;
  }

  async function resolveFolder(token) {
    const folders = parseGraphCollection(
      await graphGet(
        `${mailboxMailApiPath()}/mailFolders?$top=100`,
        token
      )
    );
    const folder = folders.find(
      (item) => cleanString(item.displayName, 200).toLowerCase() === config.folderName.toLowerCase()
    );
    if (!folder || !folder.id) {
      throw graphError(`Microsoft mail folder not found: ${config.folderName}`, 404);
    }
    return {
      id: cleanString(folder.id, 500),
      displayName: cleanString(folder.displayName, 200)
    };
  }

  async function listMessagesInFolder(token, folderId) {
    const select = [
      "id",
      "internetMessageId",
      "conversationId",
      "subject",
      "from",
      "sender",
      "receivedDateTime",
      "hasAttachments",
      "body",
      "bodyPreview"
    ].join(",");
    const path = `${mailboxMailApiPath()}/mailFolders/${encodeURIComponent(folderId)}/messages?$top=${config.maxMessagesPerRun}&$orderby=receivedDateTime desc&$select=${select}`;
    return parseGraphCollection(await graphGet(path, token)).map((message) => ({
      id: cleanString(message.id, 500),
      internetMessageId: cleanString(message.internetMessageId, 500),
      conversationId: cleanString(message.conversationId, 500),
      subject: cleanString(message.subject, 500),
      sender: cleanString(
        (message.from && message.from.emailAddress && message.from.emailAddress.address) ||
          (message.sender && message.sender.emailAddress && message.sender.emailAddress.address),
        320
      ),
      senderName: cleanString(
        (message.from && message.from.emailAddress && message.from.emailAddress.name) ||
          (message.sender && message.sender.emailAddress && message.sender.emailAddress.name),
        320
      ),
      receivedDateTime: cleanString(message.receivedDateTime, 80),
      hasAttachments: Boolean(message.hasAttachments),
      bodyContentType: cleanString(message.body && message.body.contentType, 40),
      body: cleanString(message.body && message.body.content, 60000),
      bodyPreview: cleanString(message.bodyPreview, 1000)
    }));
  }

  async function listAttachments(token, messageId) {
    const attachments = parseGraphCollection(
      await graphGet(
        `${mailboxMailApiPath()}/messages/${encodeURIComponent(messageId)}/attachments?$top=50`,
        token
      )
    );
    return attachments.map((attachment) => ({
      id: cleanString(attachment.id, 500),
      name: cleanString(attachment.name, 500),
      contentType: cleanString(attachment.contentType, 200),
      size: Number(attachment.size) || 0,
      isInline: Boolean(attachment.isInline),
      contentBytes: cleanString(attachment.contentBytes, 20 * 1024 * 1024),
      isPdf: isPdfAttachment(attachment)
    }));
  }

  async function fetchPdfAttachment(token, messageId, attachment) {
    let fullAttachment = attachment;
    if (!fullAttachment.contentBytes && attachment.id) {
      const payload = await graphGet(
        `${mailboxMailApiPath()}/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachment.id)}`,
        token
      );
      fullAttachment = {
        ...attachment,
        contentBytes: cleanString(payload.contentBytes, 20 * 1024 * 1024),
        contentType: cleanString(payload.contentType, 200),
        name: cleanString(payload.name, 500),
        size: Number(payload.size) || attachment.size || 0,
        isInline: Boolean(payload.isInline),
        isPdf: isPdfAttachment(payload)
      };
    }
    if (!fullAttachment.contentBytes) {
      throw graphError(`PDF attachment data was not available for ${attachment.name || "attachment"}.`, 502);
    }
    return {
      id: fullAttachment.id,
      name: fullAttachment.name,
      contentType: fullAttachment.contentType || "application/pdf",
      size: fullAttachment.size,
      contentBytes: fullAttachment.contentBytes,
      isInline: fullAttachment.isInline
    };
  }

  async function fetchIntakeMessages() {
    const token = await getAccessToken();
    const folder = await resolveFolder(token);
    const messages = await listMessagesInFolder(token, folder.id);
    const messagesWithAttachments = [];
    for (const message of messages) {
      const attachments = message.hasAttachments ? await listAttachments(token, message.id) : [];
      const pdfAttachments = [];
      const skippedAttachments = [];
      for (const attachment of attachments) {
        if (attachment.isInline) {
          skippedAttachments.push({ name: attachment.name, contentType: attachment.contentType, reason: "Inline attachment ignored." });
          continue;
        }
        if (!attachment.isPdf) {
          skippedAttachments.push({ name: attachment.name, contentType: attachment.contentType, reason: "Unsupported attachment type." });
          continue;
        }
        pdfAttachments.push(await fetchPdfAttachment(token, message.id, attachment));
      }
      messagesWithAttachments.push({
        ...message,
        mailbox: config.mailboxUser,
        folderId: folder.id,
        folderName: folder.displayName,
        pdfAttachments,
        skippedAttachments
      });
    }
    return {
      mailbox: config.mailboxUser,
      folder,
      messages: messagesWithAttachments
    };
  }

  return {
    fetchIntakeMessages,
    getAccessToken,
    getSafeConfigStatus,
    isConfigured,
    listAttachments,
    listMessagesInFolder,
    resolveFolder
  };
}

module.exports = {
  createMicrosoftGraphMailService
};
