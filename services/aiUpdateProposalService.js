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

  function saveAiUpdateProposal(entry) {
    const proposals = readAiUpdateProposals();
    createBackupSnapshot("before-ai-update-proposal-create");
    const normalized = normalizeAiUpdateProposal({
      ...entry,
      updatedAt: new Date().toISOString()
    });
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
