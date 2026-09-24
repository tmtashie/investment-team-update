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

  function sourceProposalSnapshot(sourceMessageKey) {
    return readAiUpdateProposals()
      .filter((proposal) => proposal.sourceMessageKey === sourceMessageKey)
      .map((proposal) => ({
        id: proposal.id,
        status: proposal.status,
        updatedAt: proposal.updatedAt,
        opportunityId: proposal.opportunityId
      }))
      .sort((left, right) => left.id.localeCompare(right.id));
  }

  function assertSourceSnapshot(current, expected) {
    if (JSON.stringify(current) === JSON.stringify(expected)) return;
    const error = new Error("Source proposals changed during reanalysis; no proposals were reconciled.");
    error.statusCode = 409;
    throw error;
  }

  function proposalIdentityKeys(proposal) {
    return new Set([
      proposal.opportunityId,
      ...(proposal.opportunityIdentityKeys || []),
      opportunityIdentity(proposal.opportunityName),
      opportunityIdentity(proposal.dealData && proposal.dealData.companyName && proposal.dealData.companyName.value)
    ].filter(Boolean));
  }

  function proposalDocumentKeys(proposal) {
    return new Set((proposal.documents || []).flatMap((document) => [
      document.hash,
      document.graphAttachmentId
    ].filter(Boolean)));
  }

  function proposalsMatch(left, right) {
    const leftIdentities = proposalIdentityKeys(left);
    const rightIdentities = proposalIdentityKeys(right);
    if ([...leftIdentities].some((key) => rightIdentities.has(key))) return true;
    const leftDocuments = proposalDocumentKeys(left);
    const rightDocuments = proposalDocumentKeys(right);
    return [...leftDocuments].some((key) => rightDocuments.has(key));
  }

  function reconcileSourceProposals({
    sourceMessageKey,
    proposals: incoming,
    expectedSnapshot,
    reviewer,
    reconciledAt = new Date().toISOString()
  }) {
    const all = readAiUpdateProposals();
    const currentSnapshot = sourceProposalSnapshot(sourceMessageKey);
    assertSourceSnapshot(currentSnapshot, expectedSnapshot);
    const pendingSiblings = all.filter((proposal) =>
      proposal.sourceMessageKey === sourceMessageKey && proposal.status === "pending"
    );
    const reviewedSiblings = all.filter((proposal) =>
      proposal.sourceMessageKey === sourceMessageKey && ["approved", "rejected"].includes(proposal.status)
    );
    const normalizedIncoming = incoming.map((entry) => normalizeAiUpdateProposal({
      ...entry,
      sourceMessageKey,
      status: "pending",
      updatedAt: reconciledAt
    }));
    const incomingIds = normalizedIncoming.map((proposal) => proposal.opportunityId).filter(Boolean);
    if (normalizedIncoming.length === 0 || incomingIds.length !== normalizedIncoming.length || new Set(incomingIds).size !== incomingIds.length) {
      const error = new Error("Reanalysis must produce a non-empty set of unique canonical opportunities.");
      error.statusCode = 422;
      throw error;
    }

    const usedCanonicalIds = new Set();
    const refreshed = [];
    const superseded = [];
    for (const candidate of normalizedIncoming) {
      const matches = pendingSiblings.filter((proposal) =>
        !usedCanonicalIds.has(proposal.id) && proposalsMatch(proposal, candidate)
      );
      if (matches.length === 0 && reviewedSiblings.some((proposal) => proposalsMatch(proposal, candidate))) {
        const error = new Error("This source opportunity already has a reviewed proposal; reanalysis did not create a replacement.");
        error.statusCode = 409;
        throw error;
      }
      const canonical = matches.find((proposal) => proposal.opportunityId === candidate.opportunityId) ||
        matches.filter((proposal) => proposal.opportunityId).sort((left, right) =>
          String(left.createdAt || "").localeCompare(String(right.createdAt || "")) || left.id.localeCompare(right.id)
        )[0];
      const saved = normalizeAiUpdateProposal({
        ...candidate,
        id: canonical ? canonical.id : candidate.id,
        createdAt: canonical ? canonical.createdAt : candidate.createdAt,
        updatedAt: reconciledAt,
        status: "pending",
        reviewedBy: "",
        reviewedAt: "",
        supersededByProposalIds: [],
        supersededReason: ""
      });
      refreshed.push(saved);
      if (canonical) usedCanonicalIds.add(canonical.id);
      matches.filter((proposal) => !canonical || proposal.id !== canonical.id).forEach((alias) => {
        usedCanonicalIds.add(alias.id);
        superseded.push(normalizeAiUpdateProposal({
          ...alias,
          status: "superseded",
          supersededByProposalIds: [saved.id],
          supersededReason: "Explicit master-editor source reanalysis reconciled this pending alias to its canonical opportunity.",
          supersededAt: reconciledAt,
          reviewedBy: reviewer,
          reviewedAt: reconciledAt,
          updatedAt: reconciledAt
        }));
      });
    }

    const replacements = new Map(refreshed.concat(superseded).map((proposal) => [proposal.id, proposal]));
    const next = all.map((proposal) => replacements.get(proposal.id) || proposal);
    refreshed.filter((proposal) => !all.some((existing) => existing.id === proposal.id)).forEach((proposal) => next.unshift(proposal));
    createBackupSnapshot("before-ai-update-source-reconciliation");
    writeAiUpdateProposals(next);
    return {
      proposals: refreshed,
      refreshedProposalIds: refreshed.map((proposal) => proposal.id),
      supersededProposalIds: superseded.map((proposal) => proposal.id)
    };
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
    writeAiUpdateProposals,
    saveAiUpdateProposal,
    sourceProposalSnapshot,
    reconcileSourceProposals,
    updateAiUpdateProposal,
    approveAiUpdateProposal,
    rejectAiUpdateProposal
  };
}

module.exports = {
  createAiUpdateProposalService
};
