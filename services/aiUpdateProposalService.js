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
    updateAiUpdateProposal,
    approveAiUpdateProposal,
    rejectAiUpdateProposal
  };
}

module.exports = {
  createAiUpdateProposalService
};
