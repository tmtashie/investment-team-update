function cleanString(value, maxLength = 2000) {
  return String(value || "").trim().slice(0, maxLength);
}

function normalizeAnalysisAudit(audit) {
  const source = audit && audit.source && typeof audit.source === "object" ? audit.source : {};
  const match = audit && audit.match && typeof audit.match === "object" ? audit.match : {};
  const counts = audit && audit.counts && typeof audit.counts === "object" ? audit.counts : {};
  return {
    sourceType: cleanString(audit && audit.sourceType, 80),
    subject: cleanString(audit && audit.subject, 500),
    sender: cleanString(audit && audit.sender, 320),
    receivedDateTime: cleanString(audit && audit.receivedDateTime, 80),
    filename: cleanString(audit && audit.filename, 500),
    graphMessageId: cleanString(audit && audit.graphMessageId, 500),
    internetMessageId: cleanString(audit && audit.internetMessageId, 500),
    finalMatchedInvestmentId: cleanString(match.finalMatchedInvestmentId || (audit && audit.finalMatchedInvestmentId), 120),
    finalMatchedInvestmentName: cleanString(match.finalMatchedInvestmentName || (audit && audit.finalMatchedInvestmentName), 200),
    matchedEntity: cleanString(match.matchedEntity || (audit && audit.matchedEntity), 120),
    matchConfidence: Number(match.matchConfidence || (audit && audit.matchConfidence)) || 0,
    matchReason: cleanString(match.matchReason || (audit && audit.matchReason), 1000),
    explicitMatch: Boolean(match.explicitMatch || (audit && audit.explicitMatch)),
    deterministicEvidenceTypes: Array.isArray(match.deterministicEvidenceTypes || (audit && audit.deterministicEvidenceTypes))
      ? (match.deterministicEvidenceTypes || audit.deterministicEvidenceTypes).map((item) => cleanString(item, 80)).filter(Boolean)
      : [],
    counts: {
      extractedFacts: Number(counts.extractedFacts) || 0,
      materialDevelopments: Number(counts.materialDevelopments) || 0,
      proposedChanges: Number(counts.proposedChanges) || 0,
      unverifiedClaims: Number(counts.unverifiedClaims) || 0,
      warnings: Number(counts.warnings) || 0
    },
    warnings: Array.isArray(audit && audit.warnings)
      ? audit.warnings.map((item) => cleanString(item, 500)).filter(Boolean).slice(0, 8)
      : [],
    shouldCreateProposal: Boolean(audit && audit.shouldCreateProposal),
    skipReason: cleanString(audit && audit.skipReason, 1000),
    analyzedAt: cleanString(audit && audit.analyzedAt, 80),
    source: {
      graphMessageId: cleanString(source.graphMessageId, 500),
      internetMessageId: cleanString(source.internetMessageId, 500)
    }
  };
}

function normalizeStateEntry(entry) {
  return {
    graphMessageId: cleanString(entry && entry.graphMessageId, 500),
    internetMessageId: cleanString(entry && entry.internetMessageId, 500),
    conversationId: cleanString(entry && entry.conversationId, 500),
    mailbox: cleanString(entry && entry.mailbox, 320),
    folderId: cleanString(entry && entry.folderId, 500),
    subject: cleanString(entry && entry.subject, 500),
    sender: cleanString(entry && entry.sender, 320),
    receivedDateTime: cleanString(entry && entry.receivedDateTime, 80),
    attachmentHashes: Array.isArray(entry && entry.attachmentHashes)
      ? entry.attachmentHashes.map((item) => cleanString(item, 128)).filter(Boolean)
      : [],
    processedAt: cleanString(entry && entry.processedAt, 80),
    proposalIds: Array.isArray(entry && entry.proposalIds)
      ? entry.proposalIds.map((item) => cleanString(item, 120)).filter(Boolean)
      : [],
    status: cleanString(entry && entry.status, 80) || "processed",
    error: cleanString(entry && entry.error, 1000),
    analysisAudits: Array.isArray(entry && entry.analysisAudits)
      ? entry.analysisAudits.map(normalizeAnalysisAudit)
      : []
  };
}

function messageDedupeKey(message) {
  return cleanString(message && message.internetMessageId, 500) ||
    cleanString(message && message.id, 500) ||
    cleanString(message && message.graphMessageId, 500);
}

function createAiEmailIntakeStateService({
  STATE_FILE,
  readJsonFile,
  writeJsonFile
}) {
  function readState() {
    const parsed = readJsonFile(STATE_FILE, []);
    return Array.isArray(parsed) ? parsed.map(normalizeStateEntry) : [];
  }

  function writeState(entries) {
    writeJsonFile(STATE_FILE, entries.map(normalizeStateEntry));
  }

  function findByMessage(message) {
    const key = messageDedupeKey(message);
    if (!key) {
      return null;
    }
    return readState().find((entry) =>
      entry.internetMessageId === key ||
      entry.graphMessageId === key ||
      (entry.internetMessageId && entry.internetMessageId === cleanString(message && message.internetMessageId, 500)) ||
      (entry.graphMessageId && entry.graphMessageId === cleanString(message && message.id, 500))
    ) || null;
  }

  function hasAttachmentHash(hash) {
    const cleanHash = cleanString(hash, 128);
    if (!cleanHash) {
      return false;
    }
    return readState().some((entry) => entry.attachmentHashes.includes(cleanHash));
  }

  function upsertEntry(entry) {
    const normalized = normalizeStateEntry(entry);
    const entries = readState();
    const key = normalized.internetMessageId || normalized.graphMessageId;
    const index = entries.findIndex((item) =>
      (key && (item.internetMessageId === key || item.graphMessageId === key)) ||
      (normalized.internetMessageId && item.internetMessageId === normalized.internetMessageId) ||
      (normalized.graphMessageId && item.graphMessageId === normalized.graphMessageId)
    );
    if (index === -1) {
      entries.unshift(normalized);
    } else {
      entries[index] = {
        ...entries[index],
        ...normalized,
        attachmentHashes: Array.from(new Set(entries[index].attachmentHashes.concat(normalized.attachmentHashes))),
        proposalIds: Array.from(new Set(entries[index].proposalIds.concat(normalized.proposalIds))),
        analysisAudits: entries[index].analysisAudits.concat(normalized.analysisAudits)
      };
    }
    writeState(entries);
    return normalized;
  }

  return {
    findByMessage,
    hasAttachmentHash,
    messageDedupeKey,
    readState,
    upsertEntry,
    writeState
  };
}

module.exports = {
  createAiEmailIntakeStateService,
  messageDedupeKey,
  normalizeStateEntry
};
