const { opportunityIdentity } = require("./newDealAnalysisService");

function createAiUpdateProposalService({
  AI_UPDATE_PROPOSALS_FILE,
  readJsonFile,
  writeJsonFile,
  writeMetadata,
  normalizeAiUpdateProposal,
  createBackupSnapshot,
  applyApprovedAiUpdateProposal
}) {
  function writeAiUpdateProposals(proposals) {
    writeJsonFile(AI_UPDATE_PROPOSALS_FILE, proposals);
  }

  function readAiUpdateProposals() {
    const parsed = readJsonFile(AI_UPDATE_PROPOSALS_FILE, []);
    if (!Array.isArray(parsed)) {
      return [];
    }

    const normalized = parsed.map(normalizeAiUpdateProposal);
    const changed = JSON.stringify(parsed) !== JSON.stringify(normalized);
    if (changed) {
      writeAiUpdateProposals(normalized);
      writeMetadata({ lastMigrationAt: new Date().toISOString() });
    }

    return normalized;
  }

  function sourceProposalSnapshot(sourceMessageKey) {
    return readAiUpdateProposals()
      .filter((proposal) => proposal.sourceMessageKey === sourceMessageKey)
      .map((proposal) => ({
        id: proposal.id,
        status: proposal.status,
        updatedAt: proposal.updatedAt,
        opportunityId: proposal.opportunityId,
        opportunityName: proposal.opportunityName,
        version: JSON.stringify(proposal)
      }))
      .sort((left, right) => left.id.localeCompare(right.id));
  }

  function proposalIdentityKeys(proposal) {
    const companyName = proposal && proposal.dealData && proposal.dealData.companyName;
    const companyValue = companyName && typeof companyName === "object" ? companyName.value : companyName;
    return new Set([
      proposal && proposal.opportunityId,
      ...(proposal && proposal.opportunityIdentityKeys || []),
      opportunityIdentity(proposal && proposal.opportunityName),
      opportunityIdentity(companyValue)
    ].filter(Boolean));
  }

  function proposalDocumentKeys(proposal) {
    return new Set((proposal && proposal.documents || [])
      .flatMap((document) => [document.hash, document.graphAttachmentId].filter(Boolean)));
  }

  function proposalsRepresentSameOpportunity(existing, incoming) {
    const incomingKeys = proposalIdentityKeys(incoming);
    if ([...proposalIdentityKeys(existing)].some((key) => incomingKeys.has(key))) return true;
    const incomingDocuments = proposalDocumentKeys(incoming);
    return [...proposalDocumentKeys(existing)].some((key) => incomingDocuments.has(key));
  }

  function reconcilePendingSourceOpportunities({
    sourceMessageKey,
    canonicalEntries,
    expectedSnapshot,
    reviewer,
    reconciledAt = new Date().toISOString()
  }) {
    const proposals = readAiUpdateProposals();
    const currentSnapshot = proposals
      .filter((proposal) => proposal.sourceMessageKey === sourceMessageKey)
      .map((proposal) => ({
        id: proposal.id,
        status: proposal.status,
        updatedAt: proposal.updatedAt,
        opportunityId: proposal.opportunityId,
        opportunityName: proposal.opportunityName,
        version: JSON.stringify(proposal)
      }))
      .sort((left, right) => left.id.localeCompare(right.id));
    if (JSON.stringify(currentSnapshot) !== JSON.stringify(expectedSnapshot || [])) {
      const error = new Error("Source proposals changed during reanalysis; no reconciliation was applied.");
      error.statusCode = 409;
      throw error;
    }
    const normalizedEntries = (Array.isArray(canonicalEntries) ? canonicalEntries : [])
      .map((entry) => normalizeAiUpdateProposal({ ...entry, sourceMessageKey, status: "pending", updatedAt: reconciledAt }));
    const canonicalIds = normalizedEntries.map((entry) => entry.opportunityId).filter(Boolean);
    if (!normalizedEntries.length || canonicalIds.length !== normalizedEntries.length || new Set(canonicalIds).size !== canonicalIds.length) {
      const error = new Error("Reanalysis did not produce a unique canonical opportunity set.");
      error.statusCode = 422;
      throw error;
    }

    const replacements = new Map();
    const supersededById = new Map();
    const canonicalProposals = [];
    for (const entry of normalizedEntries) {
      const matching = proposals.filter((proposal) =>
        proposal.sourceMessageKey === sourceMessageKey && proposal.opportunityId &&
          proposalsRepresentSameOpportunity(proposal, entry)
      );
      const pendingMatches = matching.filter((proposal) => proposal.status === "pending");
      const exactPending = pendingMatches.find((proposal) => proposal.opportunityId === entry.opportunityId);
      const namedPending = pendingMatches.find((proposal) =>
        opportunityIdentity(proposal.opportunityName) === entry.opportunityId
      );
      const selected = exactPending || namedPending || pendingMatches[0];
      if (!selected && matching.some((proposal) => ["approved", "rejected"].includes(proposal.status))) {
        const error = new Error("A reviewed source opportunity cannot be overwritten by reanalysis.");
        error.statusCode = 409;
        throw error;
      }
      const canonical = normalizeAiUpdateProposal({
        ...(selected || {}),
        ...entry,
        id: selected ? selected.id : entry.id,
        createdAt: selected ? selected.createdAt : entry.createdAt,
        status: "pending",
        reviewedBy: "",
        reviewedAt: "",
        reanalysisReconciledAt: reconciledAt,
        updatedAt: reconciledAt
      });
      canonicalProposals.push(canonical);
      if (selected) replacements.set(selected.id, canonical);
      pendingMatches.filter((proposal) => !selected || proposal.id !== selected.id).forEach((proposal) => {
        const existingTarget = supersededById.get(proposal.id);
        if (existingTarget && existingTarget !== canonical.id) {
          const error = new Error("A pending source alias matched more than one canonical opportunity.");
          error.statusCode = 409;
          throw error;
        }
        supersededById.set(proposal.id, canonical.id);
      });
    }

    const supersededProposals = [];
    const next = proposals.map((proposal) => {
      if (replacements.has(proposal.id)) return replacements.get(proposal.id);
      const canonicalId = supersededById.get(proposal.id);
      if (!canonicalId) return proposal;
      if (proposal.status !== "pending") {
        const error = new Error("Only unchanged pending aliases may be superseded by reanalysis.");
        error.statusCode = 409;
        throw error;
      }
      const superseded = normalizeAiUpdateProposal({
        ...proposal,
        status: "superseded",
        supersededByProposalIds: [canonicalId],
        supersededReason: "Explicit master-editor reanalysis reconciled this pending alias to its canonical source opportunity.",
        reviewedBy: reviewer,
        reviewedAt: reconciledAt,
        supersededAt: reconciledAt,
        reanalysisReconciledAt: reconciledAt,
        updatedAt: reconciledAt
      });
      supersededProposals.push(superseded);
      return superseded;
    });
    canonicalProposals.filter((proposal) => !proposals.some((existing) => existing.id === proposal.id))
      .forEach((proposal) => next.unshift(proposal));
    createBackupSnapshot("before-source-opportunity-reconciliation");
    writeAiUpdateProposals(next);
    return { proposals: canonicalProposals, supersededProposals };
  }

  function saveAiUpdateProposal(entry, { replacePendingSourceOpportunity = false } = {}) {
    const proposals = readAiUpdateProposals();
    const normalized = normalizeAiUpdateProposal({
      ...entry,
      updatedAt: new Date().toISOString()
    });
    let exactSourceOpportunity = normalized.sourceMessageKey && normalized.opportunityId && proposals.find(
      (proposal) => proposal.sourceMessageKey === normalized.sourceMessageKey &&
        proposal.opportunityId === normalized.opportunityId
    );
    if (!exactSourceOpportunity && replacePendingSourceOpportunity && normalized.sourceMessageKey) {
      const incomingKeys = new Set([
        normalized.opportunityId,
        ...(normalized.opportunityIdentityKeys || [])
      ].filter(Boolean));
      const documentKeys = new Set((normalized.documents || [])
        .flatMap((document) => [document.hash, document.graphAttachmentId].filter(Boolean)));
      exactSourceOpportunity = proposals.find((proposal) => {
        if (proposal.sourceMessageKey !== normalized.sourceMessageKey || proposal.status !== "pending") return false;
        const existingKeys = [proposal.opportunityId, ...(proposal.opportunityIdentityKeys || [])].filter(Boolean);
        if (existingKeys.some((key) => incomingKeys.has(key))) return true;
        return (proposal.documents || []).some((document) =>
          [document.hash, document.graphAttachmentId].filter(Boolean).some((key) => documentKeys.has(key))
        );
      });
    }
    if (exactSourceOpportunity) {
      if (replacePendingSourceOpportunity && exactSourceOpportunity.status === "pending") {
        return updateAiUpdateProposal(exactSourceOpportunity.id, {
          ...normalized,
          status: "pending",
          reviewedBy: "",
          reviewedAt: ""
        });
      }
      return exactSourceOpportunity;
    }
    if (normalized.proposalType === "new-deal") {
      const legacyExactSource = normalized.sourceMessageKey && !normalized.opportunityId && proposals.find(
        (proposal) => proposal.proposalType === "new-deal" &&
          proposal.sourceMessageKey === normalized.sourceMessageKey &&
          !proposal.opportunityId
      );
      if (legacyExactSource) return legacyExactSource;
      const sameOpportunity = normalized.opportunityFingerprint && proposals.find(
        (proposal) => proposal.proposalType === "new-deal" &&
          proposal.opportunityFingerprint === normalized.opportunityFingerprint &&
          proposal.status === "pending"
      );
      if (sameOpportunity) {
        const documentKey = (document) => document.hash || document.storedName || document.id || document.name;
        const documents = Array.from(
          new Map([...(sameOpportunity.documents || []), ...(normalized.documents || [])]
            .map((document) => [documentKey(document), document])).values()
        );
        return updateAiUpdateProposal(sameOpportunity.id, {
          documents,
          sourceMessageKeys: Array.from(new Set([
            ...(sameOpportunity.sourceMessageKeys || []),
            sameOpportunity.sourceMessageKey,
            normalized.sourceMessageKey
          ].filter(Boolean)))
        });
      }
    }
    createBackupSnapshot("before-ai-update-proposal-create");
    proposals.unshift(normalized);
    writeAiUpdateProposals(proposals);
    return normalized;
  }

  function updateAiUpdateProposal(id, updates) {
    const proposals = readAiUpdateProposals();
    const index = proposals.findIndex((proposal) => proposal.id === id);

    if (index === -1) {
      return null;
    }

    createBackupSnapshot("before-ai-update-proposal-update");
    const merged = normalizeAiUpdateProposal({
      ...proposals[index],
      ...updates,
      id: proposals[index].id,
      createdAt: proposals[index].createdAt,
      updatedAt: new Date().toISOString()
    });

    proposals[index] = merged;
    writeAiUpdateProposals(proposals);
    return merged;
  }

  function approveAiUpdateProposal(id, reviewer) {
    const reviewedAt = new Date().toISOString();
    const proposal = updateAiUpdateProposal(id, {
      status: "approved",
      reviewedBy: reviewer,
      reviewedAt
    });

    if (!proposal) {
      return null;
    }

    const applyResult = applyApprovedAiUpdateProposal(proposal);
    return {
      proposal,
      applyResult
    };
  }

  function rejectAiUpdateProposal(id, reviewer) {
    return updateAiUpdateProposal(id, {
      status: "rejected",
      reviewedBy: reviewer,
      reviewedAt: new Date().toISOString()
    });
  }

  return {
    readAiUpdateProposals,
    sourceProposalSnapshot,
    reconcilePendingSourceOpportunities,
    writeAiUpdateProposals,
    saveAiUpdateProposal,
    updateAiUpdateProposal,
    approveAiUpdateProposal,
    rejectAiUpdateProposal
  };
}

module.exports = {
  createAiUpdateProposalService
};
