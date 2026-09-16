function cleanString(value, maxLength = 2000) {
  return String(value || "").trim().slice(0, maxLength);
}

function claimValue(claim, { verifiedOnly = false } = {}) {
  if (!claim || typeof claim !== "object") return "";
  if (verifiedOnly && !["verified", "confirmed"].includes(claim.evidenceStatus)) return "";
  return cleanString(claim.authoritativeValue || claim.value, 2000);
}

function claimListValues(value, { verifiedOnly = false } = {}) {
  return (Array.isArray(value) ? value : [])
    .filter((claim) => !verifiedOnly || claim.evidenceStatus === "verified")
    .map((claim) => claimValue(claim))
    .filter(Boolean);
}

function createNewDealProposalApplyService({
  readInvestments,
  saveInvestment,
  readAiUpdateProposals,
  updateAiUpdateProposal,
  readCompanyDocuments,
  saveCompanyDocument,
  normalizeCompanyKey
}) {
  let mutationQueue = Promise.resolve();

  function runExclusive(operation) {
    const result = mutationQueue.then(operation, operation);
    mutationQueue = result.catch(() => {});
    return result;
  }

  function findCreatedInvestment(proposal, investments) {
    return investments.find((investment) =>
      (proposal.createdInvestmentId && investment.id === proposal.createdInvestmentId) ||
      investment.sourceProposalId === proposal.id
    ) || null;
  }

  function findDuplicateInvestment(proposal, companyName, investments) {
    const companyKey = normalizeCompanyKey(companyName);
    return investments.find((investment) => {
      if (investment.sourceProposalId === proposal.id) return false;
      if (proposal.opportunityFingerprint && investment.opportunityFingerprint === proposal.opportunityFingerprint) return true;
      const aliases = [investment.company].concat(investment.investmentAliases || []);
      return aliases.some((alias) => normalizeCompanyKey(alias) === companyKey);
    }) || null;
  }

  function validateApproval(proposal) {
    if (!proposal || proposal.proposalType !== "new-deal") throw Object.assign(new Error("Potential New Deal proposal not found."), { statusCode: 404 });
    if (proposal.createdInvestmentId) return;
    if (proposal.status !== "pending") throw Object.assign(new Error("Only pending Potential New Deal proposals can be approved."), { statusCode: 409 });
    if (!proposal.entityConfirmed || !cleanString(proposal.proposedEntity, 120)) {
      throw Object.assign(new Error("Confirm an investment entity before approval."), { statusCode: 409 });
    }
    const matchStatus = cleanString(proposal.matchResult && proposal.matchResult.status, 80) || "no-match";
    if (matchStatus !== "no-match" && !proposal.noExistingMatchConfirmed) {
      throw Object.assign(new Error("Resolve the existing-investment match before approval."), { statusCode: 409 });
    }
    const companyName = claimValue(proposal.dealData && proposal.dealData.companyName);
    if (!companyName) throw Object.assign(new Error("Confirm a company or deal name before approval."), { statusCode: 409 });
  }

  function associateDocuments(proposal, investment, reviewer) {
    const existingDocuments = readCompanyDocuments();
    const associatedDocuments = [];
    for (const document of proposal.documents || []) {
      if (document.preservationStatus !== "preserved" || !document.storedName || !document.url) continue;
      const existing = existingDocuments.find((item) =>
        (document.hash && item.hash === document.hash) ||
        (item.sourceProposalId === proposal.id && item.storedName === document.storedName)
      );
      if (existing) {
        associatedDocuments.push(existing);
        continue;
      }
      const saved = saveCompanyDocument({
        company: investment.company,
        entity: investment.entity,
        name: document.name,
        storedName: document.storedName,
        url: document.url,
        uploadedAt: document.uploadedAt,
        uploadedBy: reviewer,
        source: "microsoft-365-new-deal",
        notes: `Preserved from Potential New Deal proposal ${proposal.id}.`,
        hash: document.hash,
        sourceProposalId: proposal.id,
        sourceMessageKey: proposal.sourceMessageKey
      });
      existingDocuments.push(saved);
      associatedDocuments.push(saved);
    }
    return associatedDocuments;
  }

  async function approveNewDealProposal(id, reviewer) {
    return runExclusive(() => {
      const proposal = readAiUpdateProposals().find((item) => item.id === id);
      validateApproval(proposal);
      const investments = readInvestments();
      const previouslyCreated = findCreatedInvestment(proposal, investments);
      if (previouslyCreated) {
        const documents = associateDocuments(proposal, previouslyCreated, reviewer);
        const updatedProposal = proposal.createdInvestmentId
          ? proposal
          : updateAiUpdateProposal(proposal.id, { createdInvestmentId: previouslyCreated.id, status: "approved", reviewedBy: reviewer, reviewedAt: new Date().toISOString() });
        return { proposal: updatedProposal, investment: previouslyCreated, documents, idempotent: true };
      }
      const companyName = claimValue(proposal.dealData && proposal.dealData.companyName);
      const duplicate = findDuplicateInvestment(proposal, companyName, investments);
      if (duplicate) {
        throw Object.assign(new Error(`A matching investment already exists for ${duplicate.company}. Match this proposal to the existing investment.`), {
          statusCode: 409,
          duplicateInvestmentId: duplicate.id
        });
      }
      const checkClaim = proposal.dealData && proposal.dealData.proposedCheckSize;
      const amount = proposal.amountConfirmed ? claimValue(checkClaim, { verifiedOnly: true }) : "";
      const roundType = claimValue(proposal.dealData && proposal.dealData.roundType, { verifiedOnly: true });
      const summary = claimValue(proposal.dealData && proposal.dealData.dealSummary);
      const risks = claimListValues(proposal.dealData && proposal.dealData.keyRisks);
      const points = claimListValues(proposal.dealData && proposal.dealData.keyInvestmentPoints);
      const nextSteps = claimListValues(proposal.dealData && proposal.dealData.nextSteps);
      const deadlines = claimListValues(proposal.dealData && proposal.dealData.deadlines);
      const urls = claimListValues(proposal.dealData && proposal.dealData.relevantUrls, { verifiedOnly: true });
      const notes = [
        summary,
        points.length ? `Key investment points:\n- ${points.join("\n- ")}` : "",
        risks.length ? `Key risks:\n- ${risks.join("\n- ")}` : "",
        `Source proposal: ${proposal.id}`
      ].filter(Boolean).join("\n\n");
      const investment = saveInvestment({
        company: companyName,
        entity: proposal.proposedEntity,
        assetType: "Private Investment",
        status: "New Lead",
        stage: roundType,
        amount,
        currency: "USD",
        contactName: claimValue(proposal.dealData && proposal.dealData.contactName),
        contactEmail: claimValue(proposal.dealData && proposal.dealData.contactEmail),
        nextStep: nextSteps[0] || "",
        nextStepDueDate: deadlines[0] || "",
        notes,
        deckSummary: summary,
        documentLinks: urls.join("\n"),
        documents: [],
        capitalActivity: [],
        valuationHistory: [],
        ownershipHistory: [],
        followOnHistory: [],
        decisionLog: [],
        sourceProposalId: proposal.id,
        opportunityFingerprint: proposal.opportunityFingerprint,
        submittedBy: reviewer,
        createdAt: new Date().toISOString()
      });
      const associatedDocuments = associateDocuments(proposal, investment, reviewer);
      const reviewedAt = new Date().toISOString();
      const updatedProposal = updateAiUpdateProposal(proposal.id, {
        status: "approved",
        reviewedBy: reviewer,
        reviewedAt,
        createdInvestmentId: investment.id
      });
      return { proposal: updatedProposal, investment, documents: associatedDocuments, idempotent: false };
    });
  }

  return { approveNewDealProposal };
}

module.exports = { createNewDealProposalApplyService };
