const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const TOKEN_BASE_URL = "https://login.microsoftonline.com";
const MAX_ATTACHMENT_FILE_BYTES = 10 * 1024 * 1024;
const MAX_ATTACHMENT_PDF_BYTES = 25 * 1024 * 1024;
const MAX_ATTACHMENT_MESSAGE_BYTES = 30 * 1024 * 1024;
const MAX_ATTACHMENT_RUN_BYTES = 50 * 1024 * 1024;
const MAX_ATTACHMENTS_PER_MESSAGE = 20;
const SUPPORTED_ATTACHMENT_EXTENSIONS = new Set([
  ".pdf", ".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx", ".csv", ".txt", ".jpg", ".jpeg", ".png"
]);

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

function attachmentExtension(name) {
  const match = cleanString(name, 500).toLowerCase().match(/(\.[a-z0-9]+)$/);
  return match ? match[1] : "";
}

function isSupportedAttachment(attachment) {
  return SUPPORTED_ATTACHMENT_EXTENSIONS.has(attachmentExtension(attachment && attachment.name));
}

function maxAttachmentBytes(attachment) {
  return isPdfAttachment(attachment) ? MAX_ATTACHMENT_PDF_BYTES : MAX_ATTACHMENT_FILE_BYTES;
}

function attachmentLimitReason(attachment) {
  const limitMb = maxAttachmentBytes(attachment) / (1024 * 1024);
  return `Attachment exceeds the ${limitMb} MB per-file limit.`;
}

function decodedByteLength(contentBytes) {
  try {
    return Buffer.from(String(contentBytes || ""), "base64").length;
  } catch (error) {
    return 0;
  }
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
    const payload = await graphGet(
      `${mailboxMailApiPath()}/messages/${encodeURIComponent(messageId)}/attachments?$top=${MAX_ATTACHMENTS_PER_MESSAGE + 1}&$select=id,name,contentType,size,isInline`,
      token
    );
    const allAttachments = parseGraphCollection(payload);
    const attachments = allAttachments.slice(0, MAX_ATTACHMENTS_PER_MESSAGE).map((attachment) => ({
      id: cleanString(attachment.id, 500),
      name: cleanString(attachment.name, 500),
      contentType: cleanString(attachment.contentType, 200),
      size: Number(attachment.size) || 0,
      isInline: Boolean(attachment.isInline),
      contentBytes: "",
      attachmentType: cleanString(attachment["@odata.type"], 120),
      isPdf: isPdfAttachment(attachment),
      isSupported: isSupportedAttachment(attachment)
    }));
    attachments.truncated = Boolean(payload["@odata.nextLink"] || allAttachments.length > MAX_ATTACHMENTS_PER_MESSAGE);
    return attachments;
  }

  async function fetchAttachment(token, messageId, attachment) {
    let fullAttachment = attachment;
    if (!fullAttachment.contentBytes && attachment.id) {
      const payload = await graphGet(
        `${mailboxMailApiPath()}/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachment.id)}`,
        token
      );
      fullAttachment = {
        ...attachment,
        contentBytes: String(payload.contentBytes || ""),
        contentType: cleanString(payload.contentType, 200),
        name: cleanString(payload.name, 500),
        size: Number(payload.size) || attachment.size || 0,
        isInline: Boolean(payload.isInline),
        isPdf: isPdfAttachment(payload),
        isSupported: isSupportedAttachment(payload)
      };
    }
    if (!fullAttachment.contentBytes) {
      throw graphError(`Attachment data was not available for ${attachment.name || "attachment"}.`, 502);
    }
    return {
      id: fullAttachment.id,
      name: fullAttachment.name,
      contentType: fullAttachment.contentType || "application/octet-stream",
      size: fullAttachment.size,
      contentBytes: fullAttachment.contentBytes,
      isInline: fullAttachment.isInline,
      isPdf: Boolean(fullAttachment.isPdf),
      isSupported: Boolean(fullAttachment.isSupported)
    };
  }

  async function fetchIntakeMessages() {
    const token = await getAccessToken();
    const folder = await resolveFolder(token);
    const messages = await listMessagesInFolder(token, folder.id);
    const messagesWithAttachments = [];
    let runAttachmentBytes = 0;
    for (const message of messages) {
      const attachments = message.hasAttachments ? await listAttachments(token, message.id) : [];
      const preservedAttachments = [];
      const pdfAttachments = [];
      const skippedAttachments = [];
      const unresolvedAttachments = [];
      let messageAttachmentBytes = 0;
      if (attachments.truncated) {
        unresolvedAttachments.push({
          name: "Additional attachments",
          contentType: "",
          preservationStatus: "unresolved",
          reason: `Message attachment count exceeds the ${MAX_ATTACHMENTS_PER_MESSAGE}-file limit.`
        });
      }
      for (const attachment of attachments) {
        if (attachment.isInline) {
          skippedAttachments.push({ name: attachment.name, contentType: attachment.contentType, reason: "Inline attachment ignored." });
          continue;
        }
        if (!attachment.isSupported) {
          unresolvedAttachments.push({
            id: attachment.id,
            name: attachment.name,
            contentType: attachment.contentType,
            size: attachment.size,
            attachmentType: attachment.attachmentType,
            preservationStatus: "unresolved",
            reason: "Unsupported attachment type was not parsed or downloaded."
          });
          continue;
        }
        if (attachment.size > maxAttachmentBytes(attachment)) {
          unresolvedAttachments.push({ ...attachment, contentBytes: "", preservationStatus: "unresolved", reason: attachmentLimitReason(attachment) });
          continue;
        }
        if (messageAttachmentBytes + attachment.size > MAX_ATTACHMENT_MESSAGE_BYTES) {
          unresolvedAttachments.push({ ...attachment, contentBytes: "", preservationStatus: "unresolved", reason: "Attachment exceeds the 30 MB per-message limit." });
          continue;
        }
        if (runAttachmentBytes + attachment.size > MAX_ATTACHMENT_RUN_BYTES) {
          unresolvedAttachments.push({ ...attachment, contentBytes: "", preservationStatus: "unresolved", reason: "Attachment exceeds the 50 MB per-run limit." });
          continue;
        }
        try {
          const fetched = await fetchAttachment(token, message.id, attachment);
          const actualSize = decodedByteLength(fetched.contentBytes);
          if (!actualSize || actualSize > maxAttachmentBytes(fetched) || messageAttachmentBytes + actualSize > MAX_ATTACHMENT_MESSAGE_BYTES || runAttachmentBytes + actualSize > MAX_ATTACHMENT_RUN_BYTES) {
            unresolvedAttachments.push({ ...attachment, contentBytes: "", preservationStatus: "unresolved", reason: "Attachment data exceeded a configured preservation limit." });
            continue;
          }
          const preserved = { ...fetched, size: actualSize, preservationStatus: "preserved", extractionStatus: fetched.isPdf ? "pending" : "not-parsed" };
          messageAttachmentBytes += actualSize;
          runAttachmentBytes += actualSize;
          preservedAttachments.push(preserved);
          if (fetched.isPdf) pdfAttachments.push(preserved);
        } catch (error) {
          unresolvedAttachments.push({ ...attachment, contentBytes: "", preservationStatus: "unresolved", reason: error.message || "Attachment could not be preserved." });
        }
      }
      messagesWithAttachments.push({
        ...message,
        mailbox: config.mailboxUser,
        folderId: folder.id,
        folderName: folder.displayName,
        attachments: preservedAttachments,
        pdfAttachments,
        skippedAttachments,
        unresolvedAttachments
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
  MAX_ATTACHMENT_FILE_BYTES,
  MAX_ATTACHMENT_PDF_BYTES,
  MAX_ATTACHMENT_MESSAGE_BYTES,
  MAX_ATTACHMENT_RUN_BYTES,
  MAX_ATTACHMENTS_PER_MESSAGE,
  createMicrosoftGraphMailService
};
