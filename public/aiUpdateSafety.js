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

  function dealClaimLabel(item) {
    if (!item || typeof item !== "object") return "";
    return String(item.semanticLabel || "").trim();
  }

  function formatDealClaimItem(item) {
    if (!item || typeof item !== "object") return String(item || "");
    const value = item.value === undefined || item.value === null ? "" : String(item.value).trim();
    const label = dealClaimLabel(item);
    if (!label || !value || compactText(value).startsWith(compactText(label))) return value;
    return `${label}: ${value}`;
  }

  function claimItems(proposal, field) {
    const claim = proposal && proposal.dealData && proposal.dealData[field];
    return Array.isArray(claim) ? claim : claim && typeof claim === "object" ? [claim] : [];
  }

  function claimValue(proposal, field) {
    const item = claimItems(proposal, field)[0];
    return String((item && (item.authoritativeValue || item.value)) || "").trim();
  }

  function fundraisingIsClosed(proposal) {
    return /\bfundraising\s+(?:is\s+)?closed\b/i.test(claimValue(proposal, "stage"));
  }

  function parseMillions(value) {
    const match = String(value || "").match(/\$?\s*([0-9]+(?:,[0-9]{3})*(?:\.[0-9]+)?)\s*(B|BN|BILLION|M|MM|MILLION|K|THOUSAND)?\b/i);
    if (!match) return null;
    const amount = Number(match[1].replace(/,/g, ""));
    if (!Number.isFinite(amount)) return null;
    const unit = String(match[2] || "").toLowerCase();
    if (["b", "bn", "billion"].includes(unit)) return amount * 1000;
    if (["k", "thousand"].includes(unit)) return amount / 1000;
    return amount;
  }

  function formatMillions(value) {
    return `$${value.toLocaleString("en-US", { maximumFractionDigits: 3 })}MM`;
  }

  function deriveHistoricalTargetDifference(proposal) {
    if (!fundraisingIsClosed(proposal)) return null;
    const targetClaim = claimItems(proposal, "targetFundSize")[0];
    const committedClaim = claimItems(proposal, "amountCommitted")[0];
    if (!isVerified(targetClaim) || !isVerified(committedClaim)) return null;
    const target = parseMillions(targetClaim.authoritativeValue || targetClaim.value);
    const committed = parseMillions(committedClaim.authoritativeValue || committedClaim.value);
    if (target === null || committed === null || target <= committed) return null;
    const value = formatMillions(Math.round((target - committed) * 1000) / 1000);
    return {
      value,
      authoritativeValue: value,
      evidenceStatus: "verified",
      sourceEvidence: [targetClaim.sourceEvidence, committedClaim.sourceEvidence].filter(Boolean).join(" | "),
      sourceLocation: "Derived from verified target and historical commitments",
      semanticMeaning: "historical-unfunded-target-difference",
      currentAvailability: false,
      derivedFrom: [
        { field: "targetFundSize", value: targetClaim.value, sourceEvidence: targetClaim.sourceEvidence },
        { field: "amountCommitted", value: committedClaim.value, sourceEvidence: committedClaim.sourceEvidence }
      ]
    };
  }

  function sourceSupportsNextStep(claim) {
    if (String((claim && claim.evidenceStatus) || "").toLowerCase() === "confirmed") return true;
    if (!isVerified(claim)) return false;
    const evidence = String((claim && claim.sourceEvidence) || "");
    const hasAction = /\b(review|evaluate|diligence|contact|follow[- ]?up|schedule|meet|meeting|call|decide|respond|request|send|provide|discuss|consider|introduc(?:e|tion))\b/i.test(evidence);
    const isDirected = /\b(beaman(?: ventures)?|tyler|lee|we|our|you|your|please|should|will|agreed|scheduled|available|happy to|let me know|can connect)\b/i.test(evidence);
    return hasAction && isDirected;
  }

  function normalizeClosedFundraisingNarrative(value, targetFundSize) {
    const text = String(value || "").trim();
    if (!text) return "";
    const target = String(targetFundSize || "").trim();
    let changed = false;
    let normalized = text.replace(
      /\b(?:is\s+)?(?:currently\s+)?(?:raising|seeking\s+to\s+raise|seeks\s+to\s+raise)\s+\$?\s*[0-9]+(?:,[0-9]{3})*(?:\.[0-9]+)?\s*(?:b|bn|billion|m|mm|million|k|thousand)?\b/gi,
      () => {
        changed = true;
        return target ? `has a target fund size of ${target}` : "has a stated target fund size";
      }
    );
    normalized = normalized
      .replace(/\bactive fundraising\b/gi, () => { changed = true; return "fundraising closed"; })
      .replace(/\bopen for new commitments\b/gi, () => { changed = true; return "closed to new commitments"; });
    if (changed && !/\bfundraising\s+(?:is\s+)?closed\b/i.test(normalized)) {
      normalized = `${normalized.replace(/[.\s]+$/, "")}. Fundraising is closed.`;
    }
    return normalized;
  }

  function getDisplayDealClaimItems(proposal, field) {
    let items = claimItems(proposal, field);
    if (["amountRemaining", "proposedCheckSize"].includes(field)) {
      items = items.filter((item) => ["verified", "confirmed"].includes(String((item && item.evidenceStatus) || "").toLowerCase()));
    }
    if (field === "historicalTargetDifference" && !items.some((item) => item && item.value)) {
      const derived = deriveHistoricalTargetDifference(proposal);
      items = derived ? [derived] : items;
    }
    if (field === "nextSteps") items = items.filter(sourceSupportsNextStep);
    if (!fundraisingIsClosed(proposal) || field === "stage") return items;
    const target = claimValue(proposal, "targetFundSize");
    return items.map((item) => {
      const value = normalizeClosedFundraisingNarrative(item && item.value, target);
      return value === String((item && item.value) || "")
        ? item
        : { ...item, value, authoritativeValue: isVerified(item) ? value : "" };
    });
  }

  function formatDealClaimValue(proposal, field) {
    return getDisplayDealClaimItems(proposal, field).map(formatDealClaimItem).filter(Boolean).join("\n");
  }

  function formatAiUpdateProposalSummary(proposal) {
    const dealSummary = proposal && proposal.proposalType === "new-deal"
      ? formatDealClaimValue(proposal, "dealSummary")
      : "";
    if (dealSummary) return dealSummary;
    const summary = String((proposal && proposal.summary) || "").trim();
    return fundraisingIsClosed(proposal)
      ? normalizeClosedFundraisingNarrative(summary, claimValue(proposal, "targetFundSize"))
      : summary;
  }

  function normalizeAiUpdateProposalCounts(data) {
    const statuses = ["pending", "approved", "rejected", "superseded"];
    const proposals = data && Array.isArray(data.proposals) ? data.proposals : null;
    return statuses.reduce((counts, status) => {
      counts[status] = proposals
        ? proposals.filter((proposal) => proposal && proposal.status === status).length
        : Number((data && data.counts && data.counts[status]) || 0);
      return counts;
    }, {});
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
    dealClaimLabel,
    formatAiUpdateProposalSummary,
    formatDealClaimItem,
    formatDealClaimValue,
    getDisplayDealClaimItems,
    getReportUpdatesEmptyMessage,
    getItemValue,
    isActionable,
    isHighRiskNumeric,
    isVerified,
    normalizeAiUpdateProposalCounts,
    normalizeClosedFundraisingNarrative,
    sanitizeForActionableView,
    shouldRefreshInvestmentsAfterAiProposalAction,
    warningMessage
  };

  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
  root.AiUpdateSafety = api;
})(typeof window !== "undefined" ? window : globalThis);
