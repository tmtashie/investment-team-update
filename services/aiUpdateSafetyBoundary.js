function parseAiNumericValue(value) {
  const text = String(value || "").trim().toLowerCase();
  const match = text.match(/(?<![a-z])-?\$?\s*([0-9]+(?:,[0-9]{3})*(?:\.[0-9]+)?|[0-9]*\.[0-9]+)\s*(m|mm|b|bn|k|thousand|million|billion)?(?![a-z])/i);
  if (!match) {
    return null;
  }
  const base = Number(match[1].replace(/,/g, ""));
  if (!Number.isFinite(base)) {
    return null;
  }
  const suffix = String(match[2] || "").toLowerCase();
  const multiplier =
    suffix === "k" || suffix === "thousand"
      ? 1000
      : suffix === "m" || suffix === "mm" || suffix === "million"
        ? 1000000
        : suffix === "b" || suffix === "bn" || suffix === "billion"
          ? 1000000000
          : 1;
  return base * multiplier;
}

function compactAiUpdateText(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

function getAiUpdateItemValue(item) {
  if (!item || typeof item !== "object") {
    return "";
  }
  if (item.proposedValue !== undefined) return item.proposedValue;
  if (item.proposed_value !== undefined) return item.proposed_value;
  if (item.value !== undefined) return item.value;
  if (item.summary !== undefined) return item.summary;
  return "";
}

function warningMessage(value) {
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return String(value).trim();
  }
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return "";
  }
  return String(
    value.message || value.reason || value.warning || value.detail || value.description || value.text || ""
  ).trim();
}

function getAiUpdateItemField(item) {
  return String(
    item && (item.field || item.actionType || item.action_type || item.category || "")
  ).trim();
}

function getClaimNumericText(item) {
  const text = String(getAiUpdateItemValue(item) || "");
  const match = text.match(/\$?\s*[0-9]+(?:,[0-9]{3})*(?:\.[0-9]+)?\s*(?:m|mm|b|bn|k|thousand|million|billion)?/i);
  return match ? match[0].replace(/\s+/g, " ").trim() : "";
}

function getUserFacingFieldLabel(item) {
  const fieldText = `${getAiUpdateItemField(item)} ${item && item.category || ""} ${item && item.summary || ""}`;
  const normalized = compactAiUpdateText(fieldText);
  if (/mrr|monthlyrecurringrevenue|recurringrevenue/.test(normalized)) return "Monthly Recurring Revenue";
  if (/revenue/.test(normalized)) return "Revenue";
  if (/lineofcredit|loc|facility/.test(normalized)) return "Line of Credit";
  if (/unitssold|unitsold/.test(normalized)) return "Units Sold";
  if (/unitsinstalled|installed/.test(normalized)) return "Units Installed";
  if (/training|trainings/.test(normalized)) return "Trainings";
  return getAiUpdateItemField(item) || String((item && item.category) || "Claim").trim();
}

function userFacingWarningForClaim(item) {
  const label = getUserFacingFieldLabel(item);
  const value = getClaimNumericText(item) || String(getAiUpdateItemValue(item) || "").trim();
  return `${label}${value ? `: ${value}` : ""} — Could not verify against source document.`;
}

function claimIdentity(item) {
  const label = compactAiUpdateText(getUserFacingFieldLabel(item));
  const numeric = parseAiNumericValue(getClaimNumericText(item));
  const value = numeric !== null ? String(numeric) : compactAiUpdateText(getAiUpdateItemValue(item));
  return `${label}:${value}`;
}

function buildUserFacingWarnings(analysis) {
  const input = analysis && typeof analysis === "object" ? analysis : {};
  const claims = Array.isArray(input.unverifiedClaims) ? input.unverifiedClaims : [];
  const seen = new Set();
  const lines = [];
  claims.forEach((claim) => {
    const key = claimIdentity(claim);
    if (!key || seen.has(key)) {
      return;
    }
    seen.add(key);
    lines.push(userFacingWarningForClaim(claim));
  });
  return lines;
}

function isHighRiskAiNumericItem(item) {
  const field = compactAiUpdateText(getAiUpdateItemField(item));
  const value = getAiUpdateItemValue(item);
  return (
    parseAiNumericValue(value) !== null &&
    /revenue|ebitda|cash|runway|valuation|ownership|capitalcall|distribution|commitment|debt|lineofcredit|loc|costbasis|nav|irr|moic|cashflow|recurringrevenue|mrr|arr|unit|units|kpi|pipeline/.test(field)
  );
}

function isVerifiedAiItem(item) {
  return String((item && item.evidenceStatus) || "").trim().toLowerCase() === "verified";
}

function buildRejectedAiClaim(item, reason, source) {
  return {
    category: String((item && item.category) || "").trim(),
    field: getAiUpdateItemField(item),
    value: String(getAiUpdateItemValue(item) || "").trim(),
    currentValue: String((item && (item.currentValue || item.current_value || item.current)) || "").trim(),
    sourcePage: item && (item.sourcePage || item.source_page || item.page || item.pageNumber) || "",
    sourceEvidence: String((item && item.sourceEvidence) || "").trim(),
    evidenceStatus: String((item && item.evidenceStatus) || "unresolved").trim() || "unresolved",
    confidence: Number((item && item.confidence) || 0) || 0,
    verification: item && item.verification || {},
    source,
    reason
  };
}

function rebuildWhatChanged(proposedChanges, materialDevelopments) {
  const lines = proposedChanges
    .filter(isVerifiedAiItem)
    .map((change) => {
      const field = getAiUpdateItemField(change) || "Value";
      const current = String(change.currentValue || "Not currently recorded").trim();
      const proposedValue = String(getAiUpdateItemValue(change) || "").trim();
      return proposedValue ? `${field} changed from ${current} to ${proposedValue}.` : "";
    })
    .concat(
      materialDevelopments
        .filter(isVerifiedAiItem)
        .map((development) => String(development.summary || "").trim())
    )
    .filter(Boolean);

  return lines.length
    ? Array.from(new Set(lines))
    : ["No verified portfolio changes identified from this document."];
}

function enforceAnalysisSafetyInvariant(analysis) {
  const input = analysis && typeof analysis === "object" ? analysis : {};
  const warnings = Array.isArray(input.warnings) ? input.warnings.map(warningMessage).filter(Boolean) : [];
  const unresolved = Array.isArray(input.unresolved) ? input.unresolved.map(warningMessage).filter(Boolean) : [];
  const unverifiedClaims = Array.isArray(input.unverifiedClaims) ? input.unverifiedClaims.slice() : [];
  const facts = Array.isArray(input.extractedFacts) ? input.extractedFacts : [];
  const developments = Array.isArray(input.materialDevelopments) ? input.materialDevelopments : [];
  const proposed = Array.isArray(input.proposedChanges) ? input.proposedChanges : [];

  const safeFacts = facts.filter((fact) => {
    if (isHighRiskAiNumericItem(fact) && !isVerifiedAiItem(fact)) {
      const reason = `Rejected unverified high-risk numeric fact: ${getAiUpdateItemField(fact) || "numeric claim"} (${getAiUpdateItemValue(fact)}).`;
      unverifiedClaims.push(buildRejectedAiClaim(fact, reason, "extractedFacts"));
      warnings.push(reason);
      unresolved.push(reason);
      return false;
    }
    return true;
  });

  const safeProposedChanges = proposed.filter((change) => {
    if (isHighRiskAiNumericItem(change) && !isVerifiedAiItem(change)) {
      const reason = `Rejected unverified high-risk numeric proposed change: ${getAiUpdateItemField(change) || "numeric field"} (${getAiUpdateItemValue(change)}).`;
      unverifiedClaims.push(buildRejectedAiClaim(change, reason, "proposedChanges"));
      warnings.push(reason);
      unresolved.push(reason);
      return false;
    }
    return true;
  });

  const safeDevelopments = developments.filter((development) => {
    if (!isVerifiedAiItem(development)) {
      const reason = `Could not verify material development with claim-level source evidence.`;
      unverifiedClaims.push(buildRejectedAiClaim(development, reason, "materialDevelopments"));
      return false;
    }
    return true;
  });

  return {
    ...input,
    extractedFacts: safeFacts,
    materialDevelopments: safeDevelopments,
    proposedChanges: safeProposedChanges,
    whatChanged: rebuildWhatChanged(safeProposedChanges, safeDevelopments),
    unverifiedClaims,
    userFacingWarnings: buildUserFacingWarnings({ ...input, unverifiedClaims }),
    warnings: Array.from(new Set(warnings.filter(Boolean))),
    unresolved: Array.from(new Set(unresolved.filter(Boolean)))
  };
}

function findUnsafeActionableItems(analysis) {
  const input = analysis && typeof analysis === "object" ? analysis : {};
  return []
    .concat(Array.isArray(input.extractedFacts) ? input.extractedFacts : [])
    .concat(Array.isArray(input.proposedChanges) ? input.proposedChanges : [])
    .filter((item) => isHighRiskAiNumericItem(item) && !isVerifiedAiItem(item));
}

function finalizeAnalysisForResponse(analysis) {
  let sanitized = enforceAnalysisSafetyInvariant(analysis);
  let unsafe = findUnsafeActionableItems(sanitized);
  if (unsafe.length) {
    sanitized = enforceAnalysisSafetyInvariant(sanitized);
    unsafe = findUnsafeActionableItems(sanitized);
  }
  if (unsafe.length) {
    return {
      ...sanitized,
      extractedFacts: [],
      proposedChanges: [],
      whatChanged: ["No verified portfolio changes identified from this document."],
      warnings: Array.from(
        new Set([...(sanitized.warnings || []), "Unsafe unverified numeric claims were removed before response."])
      ),
      unresolved: Array.from(
        new Set([...(sanitized.unresolved || []), "Unsafe unverified numeric claims were removed before response."])
      )
    };
  }
  return sanitized;
}

function getProposalCurrentValueFromInvestment(investment, field) {
  if (!investment) {
    return "Not currently recorded";
  }
  const normalizedField = compactAiUpdateText(field);
  const aliases = {
    nav: "internalValue",
    currentvalue: "internalValue",
    valuation: "officialValue",
    officialvalue: "officialValue",
    internalvalue: "internalValue",
    ownershippercentage: "ownershipPercent",
    ownershippercent: "ownershipPercent",
    entityownershippercentage: "entityOwnershipPercent",
    entityownershippercent: "entityOwnershipPercent",
    capitalcallamount: "capitalCallAmount",
    capitalcalldate: "capitalCallDate",
    capitalcallduedate: "capitalCallDate",
    distributionamount: "distributionAmount",
    distributiondate: "distributionDate"
  };
  const key = aliases[normalizedField];
  return key && investment[key] !== undefined && investment[key] !== ""
    ? investment[key]
    : "Not currently recorded";
}

function protectProposalCurrentValues(proposedChanges, investment) {
  return proposedChanges.map((change) => ({
    ...change,
    currentValue: getProposalCurrentValueFromInvestment(
      investment,
      getAiUpdateItemField(change)
    )
  }));
}

function enforceProposalSafetyInvariant(proposal, investment) {
  const extractedData = proposal && proposal.extractedData && typeof proposal.extractedData === "object"
    ? proposal.extractedData
    : {};
  const sanitizedAnalysis = finalizeAnalysisForResponse({
    extractedFacts: Array.isArray(extractedData.facts) ? extractedData.facts : [],
    materialDevelopments: Array.isArray(extractedData.materialDevelopments) ? extractedData.materialDevelopments : [],
    proposedChanges: Array.isArray(proposal && proposal.proposedChanges) ? proposal.proposedChanges : [],
    unverifiedClaims: Array.isArray(extractedData.unverifiedClaims) ? extractedData.unverifiedClaims : [],
    warnings: Array.isArray(extractedData.warnings) ? extractedData.warnings : [],
    unresolved: Array.isArray(extractedData.unresolved) ? extractedData.unresolved : []
  });
  const protectedProposedChanges = protectProposalCurrentValues(
    sanitizedAnalysis.proposedChanges,
    investment
  );
  const summaryLines = rebuildWhatChanged(protectedProposedChanges, sanitizedAnalysis.materialDevelopments);

  return {
    ...proposal,
    summary: summaryLines.map((item) => `• ${item}`).join("\n"),
    extractedData: {
      ...extractedData,
      facts: sanitizedAnalysis.extractedFacts,
      materialDevelopments: sanitizedAnalysis.materialDevelopments,
      unverifiedClaims: sanitizedAnalysis.unverifiedClaims,
      warnings: sanitizedAnalysis.warnings,
      unresolved: sanitizedAnalysis.unresolved
    },
    proposedChanges: protectedProposedChanges
  };
}

module.exports = {
  compactAiUpdateText,
  enforceAnalysisSafetyInvariant,
  enforceProposalSafetyInvariant,
  finalizeAnalysisForResponse,
  findUnsafeActionableItems,
  buildUserFacingWarnings,
  getAiUpdateItemValue,
  isHighRiskAiNumericItem,
  isVerifiedAiItem,
  parseAiNumericValue,
  rebuildWhatChanged
};
