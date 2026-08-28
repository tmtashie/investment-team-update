function cleanString(value, maxLength = 2000) {
  return String(value || "").trim().slice(0, maxLength);
}

function normalizeEvidenceDevelopment(item) {
  return {
    category: cleanString(item && item.category, 120),
    summary: cleanString(item && item.summary, 1000),
    sourceEvidence: cleanString(item && item.sourceEvidence, 1000),
    sourcePage: cleanString(item && item.sourcePage, 40),
    confidence: Number((item && item.confidence) || 0) || 0,
    riskLevel: cleanString(item && item.riskLevel, 40),
    importance: cleanString(item && item.importance, 40),
    evidenceStatus: cleanString(item && item.evidenceStatus, 40),
    verification: item && item.verification && typeof item.verification === "object"
      ? item.verification
      : {}
  };
}

function verifiedMaterialDevelopmentsFromProposal(proposal) {
  const extractedData = proposal && proposal.extractedData && typeof proposal.extractedData === "object"
    ? proposal.extractedData
    : {};
  return (Array.isArray(extractedData.materialDevelopments) ? extractedData.materialDevelopments : [])
    .filter((item) => cleanString(item && item.evidenceStatus).toLowerCase() === "verified")
    .map(normalizeEvidenceDevelopment)
    .filter((item) => item.summary && item.sourceEvidence);
}

function getProposalSourceFilename(proposal) {
  const extractedData = proposal && proposal.extractedData && typeof proposal.extractedData === "object"
    ? proposal.extractedData
    : {};
  const source = extractedData.source && typeof extractedData.source === "object"
    ? extractedData.source
    : {};
  const document = Array.isArray(proposal && proposal.documents) ? proposal.documents[0] : null;
  return cleanString(source.filename || (document && document.name) || proposal.sourceIdentifier, 500);
}

function getProposalAttachmentLink(proposal) {
  const document = Array.isArray(proposal && proposal.documents) ? proposal.documents[0] : null;
  return cleanString((document && (document.url || document.name)) || proposal.sourceIdentifier, 1000);
}

function buildMaterialDevelopmentNotes(developments, proposal, approvedAt, approver) {
  const header = [
    `AI update proposal: ${cleanString(proposal && proposal.id, 120)}`,
    `Approved by: ${cleanString(approver, 240) || "Unknown"}`,
    `Approved at: ${approvedAt}`,
    `Source: ${cleanString(proposal && proposal.sourceType, 80) || "Unknown"}${getProposalSourceFilename(proposal) ? ` - ${getProposalSourceFilename(proposal)}` : ""}`
  ].join("\n");
  const evidence = developments
    .map((item) => {
      const page = item.sourcePage ? `Page ${item.sourcePage}: ` : "";
      const confidence = item.confidence ? ` Confidence: ${item.confidence}%.` : "";
      return `- ${item.summary}\n  Evidence: ${page}${item.sourceEvidence}${confidence}`;
    })
    .join("\n");
  return `${header}\n\nVerified material developments:\n${evidence}`;
}

function buildApprovedMaterialDevelopmentsReportUpdate(proposal, approver, approvedAt) {
  const developments = verifiedMaterialDevelopmentsFromProposal(proposal);
  if (!developments.length) {
    return null;
  }
  const sourceFilename = getProposalSourceFilename(proposal);
  const date = cleanString(proposal && proposal.sourceDate, 80) || approvedAt;
  const summary = developments.map((item) => `• ${item.summary}`).join("\n");

  return {
    date,
    type: "AI Material Developments",
    title: sourceFilename ? `AI update from ${sourceFilename}` : "AI-approved material developments",
    reportPeriod: cleanString(proposal && proposal.sourceDate, 120),
    sourceType: cleanString(proposal && proposal.sourceType, 80),
    summary,
    originalNotes: buildMaterialDevelopmentNotes(developments, proposal, approvedAt, approver),
    aiSummary: summary,
    keyWins: developments
      .filter((item) => /customer|win|sales|progress|operational|operations|r&d|rd|financing|capital/i.test(item.category))
      .map((item) => `• ${item.summary}`)
      .join("\n"),
    keyRisks: developments
      .filter((item) => /risk|concern|delay|issue/i.test(item.category))
      .map((item) => `• ${item.summary}`)
      .join("\n"),
    keyMetrics: developments
      .filter((item) => /\d/.test(item.summary))
      .map((item) => `• ${item.summary}`)
      .join("\n"),
    attachmentLink: getProposalAttachmentLink(proposal),
    sourceUpdateId: cleanString(proposal && proposal.id, 120),
    aiProposalId: cleanString(proposal && proposal.id, 120),
    aiApprovedBy: cleanString(approver, 240),
    aiApprovedAt: approvedAt,
    aiSourceFilename: sourceFilename,
    aiSourceIdentifier: cleanString(proposal && proposal.sourceIdentifier, 1000),
    aiMaterialDevelopments: developments
  };
}

function applyApprovedAiUpdateProposalToInvestment({
  proposal,
  investment,
  approver,
  approvedAt,
  updateInvestment,
  normalizeStructuredRows
}) {
  const update = buildApprovedMaterialDevelopmentsReportUpdate(proposal, approver, approvedAt);
  if (!update) {
    return {
      applied: false,
      message: "Approval recorded. No verified material developments were available to append to the live investment.",
      appendedReportUpdate: null,
      ignoredProposedChanges: Array.isArray(proposal && proposal.proposedChanges)
        ? proposal.proposedChanges.length
        : 0
    };
  }
  if (!investment || !investment.id) {
    return {
      applied: false,
      message: "Approval recorded, but no matched live investment was found for the material development update.",
      appendedReportUpdate: null,
      ignoredProposedChanges: Array.isArray(proposal && proposal.proposedChanges)
        ? proposal.proposedChanges.length
        : 0
    };
  }

  const currentRows = typeof normalizeStructuredRows === "function"
    ? normalizeStructuredRows(investment.reportUpdates)
    : Array.isArray(investment.reportUpdates)
      ? investment.reportUpdates
      : [];
  const updated = updateInvestment(investment.id, {
    ...investment,
    reportUpdates: [update, ...currentRows]
  });

  return {
    applied: Boolean(updated),
    message: updated
      ? "Approval recorded. Verified material developments were appended to the live investment update history."
      : "Approval recorded, but the live investment update could not be written.",
    investmentId: investment.id,
    appendedReportUpdate: update,
    ignoredProposedChanges: Array.isArray(proposal && proposal.proposedChanges)
      ? proposal.proposedChanges.length
      : 0
  };
}

module.exports = {
  applyApprovedAiUpdateProposalToInvestment,
  buildApprovedMaterialDevelopmentsReportUpdate,
  verifiedMaterialDevelopmentsFromProposal
};
