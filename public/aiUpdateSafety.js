(function attachAiUpdateSafety(root) {
  function parseAiNumericValue(value) {
    const text = String(value || "").trim().toLowerCase();
    const match = text.match(/-?\$?\s*([0-9]+(?:,[0-9]{3})*(?:\.[0-9]+)?|[0-9]*\.[0-9]+)\s*(m|mm|b|bn|k|thousand|million|billion)?(?![a-z])/i);
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

  function compactText(value) {
    return String(value || "")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "");
  }

  function getItemValue(item) {
    if (!item || typeof item !== "object") {
      return "";
    }
    if (item.proposedValue !== undefined) return item.proposedValue;
    if (item.proposed_value !== undefined) return item.proposed_value;
    if (item.value !== undefined) return item.value;
    if (item.summary !== undefined) return item.summary;
    return "";
  }

  function getItemField(item) {
    return String(
      item && (item.field || item.actionType || item.action_type || item.category || "")
    ).trim();
  }

  function isVerified(item) {
    return String((item && item.evidenceStatus) || "").trim().toLowerCase() === "verified";
  }

  function isHighRiskNumeric(item) {
    const field = compactText(getItemField(item));
    return (
      parseAiNumericValue(getItemValue(item)) !== null &&
      /revenue|ebitda|cash|runway|valuation|ownership|capitalcall|distribution|commitment|debt|lineofcredit|loc|costbasis|nav|irr|moic|cashflow|recurringrevenue|mrr|arr|unit|units|kpi|pipeline/.test(field)
    );
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

  function getClaimNumericText(item) {
    const text = String(getItemValue(item) || "");
    const match = text.match(/\$?\s*[0-9]+(?:,[0-9]{3})*(?:\.[0-9]+)?\s*(?:m|mm|b|bn|k|thousand|million|billion)?/i);
    return match ? match[0].replace(/\s+/g, " ").trim() : "";
  }

  function getUserFacingFieldLabel(item) {
    const fieldText = `${getItemField(item)} ${(item && item.category) || ""} ${(item && item.summary) || ""}`;
    const normalized = compactText(fieldText);
    if (/mrr|monthlyrecurringrevenue|recurringrevenue/.test(normalized)) return "Monthly Recurring Revenue";
    if (/revenue/.test(normalized)) return "Revenue";
    if (/lineofcredit|loc|facility/.test(normalized)) return "Line of Credit";
    if (/unitssold|unitsold/.test(normalized)) return "Units Sold";
    if (/unitsinstalled|installed/.test(normalized)) return "Units Installed";
    if (/training|trainings/.test(normalized)) return "Trainings";
    return getItemField(item) || String((item && item.category) || "Claim").trim();
  }

  function claimIdentity(item) {
    const label = compactText(getUserFacingFieldLabel(item));
    const numeric = parseAiNumericValue(getClaimNumericText(item));
    const value = numeric !== null ? String(numeric) : compactText(getItemValue(item));
    return `${label}:${value}`;
  }

  function userFacingWarningForClaim(item) {
    const label = getUserFacingFieldLabel(item);
    const value = getClaimNumericText(item) || String(getItemValue(item) || "").trim();
    return `${label}${value ? `: ${value}` : ""} — Could not verify against source document.`;
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

  function isActionable(item) {
    return !(isHighRiskNumeric(item) && !isVerified(item));
  }

  function rejectedClaimFrom(item, source) {
    return {
      category: String((item && item.category) || "").trim(),
      field: getItemField(item),
      value: String(getItemValue(item) || "").trim(),
      currentValue: String((item && (item.currentValue || item.current_value || item.current)) || "").trim(),
      evidenceStatus: String((item && item.evidenceStatus) || "unresolved").trim() || "unresolved",
      sourceEvidence: String((item && item.sourceEvidence) || "").trim(),
      source
    };
  }

  function sanitizeForActionableView(analysis) {
    const input = analysis && typeof analysis === "object" ? analysis : {};
    const extractedFacts = Array.isArray(input.extractedFacts) ? input.extractedFacts : [];
    const proposedChanges = Array.isArray(input.proposedChanges) ? input.proposedChanges : [];
    const materialDevelopments = Array.isArray(input.materialDevelopments) ? input.materialDevelopments : [];
    const unverifiedClaims = Array.isArray(input.unverifiedClaims) ? input.unverifiedClaims.slice() : [];

    const safeFacts = extractedFacts.filter((fact) => {
      const safe = isActionable(fact);
      if (!safe) {
        unverifiedClaims.push(rejectedClaimFrom(fact, "extractedFacts"));
      }
      return safe;
    });
    const safeProposedChanges = proposedChanges.filter((change) => {
      const safe = isActionable(change);
      if (!safe) {
        unverifiedClaims.push(rejectedClaimFrom(change, "proposedChanges"));
      }
      return safe;
    });
    const safeDevelopments = materialDevelopments.filter((development) => {
      const safe = isVerified(development);
      if (!safe) {
        unverifiedClaims.push(rejectedClaimFrom(development, "materialDevelopments"));
      }
      return safe;
    });
    const rejectedCount =
      extractedFacts.length - safeFacts.length +
      proposedChanges.length - safeProposedChanges.length +
      materialDevelopments.length - safeDevelopments.length;

    return {
      ...input,
      extractedFacts: safeFacts,
      materialDevelopments: safeDevelopments,
      proposedChanges: safeProposedChanges,
      unverifiedClaims,
      userFacingWarnings: Array.isArray(input.userFacingWarnings) && input.userFacingWarnings.length
        ? input.userFacingWarnings
        : buildUserFacingWarnings({ ...input, unverifiedClaims }),
      whatChanged: rejectedCount
        ? ["No verified portfolio changes identified from this document."]
        : Array.isArray(input.whatChanged)
          ? input.whatChanged
          : []
    };
  }

  function getReportUpdatesEmptyMessage(totalRows, visibleRows) {
    return Number(totalRows || 0) > 0 && Number(visibleRows || 0) === 0
      ? "No reports match the current filters."
      : "No saved updates or reports yet. Add your first monthly report, quarterly letter, capital call, or call note above.";
  }

  function shouldRefreshInvestmentsAfterAiProposalAction(action, result) {
    return action === "approve" && Boolean(result && result.proposal);
  }

  const api = {
    buildUserFacingWarnings,
    getReportUpdatesEmptyMessage,
    getItemValue,
    isActionable,
    isHighRiskNumeric,
    isVerified,
    sanitizeForActionableView,
    shouldRefreshInvestmentsAfterAiProposalAction,
    warningMessage
  };

  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
  root.AiUpdateSafety = api;
})(typeof window !== "undefined" ? window : globalThis);
