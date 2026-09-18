const SEMANTIC_ONLY_CONFIDENCE_CAP = 84;

function cleanString(value, maxLength = 2000) {
  return String(value || "").trim().slice(0, maxLength);
}

function normalizeMatchText(value) {
  return cleanString(value, 500)
    .toLowerCase()
    .replace(/\b(limited liability company|incorporated|corporation|company|partners|holdings)\b/g, " ")
    .replace(/\b(l\.?l\.?c\.?|inc\.?|corp\.?|co\.?|l\.?p\.?|llp|ltd\.?)\b/g, " ")
    .replace(/&/g, " and ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function compactMatchText(value) {
  return normalizeMatchText(value).replace(/\s+/g, "");
}

function hasExplicitPhrase(sourceText, phrase) {
  const normalizedPhrase = normalizeMatchText(phrase);
  if (!normalizedPhrase || normalizedPhrase.length < 4) {
    return false;
  }
  const normalizedSource = ` ${normalizeMatchText(sourceText)} `;
  if (normalizedSource.includes(` ${normalizedPhrase} `)) {
    return true;
  }
  const compactPhrase = compactMatchText(phrase);
  if (compactPhrase.length < 6) {
    return false;
  }
  const sourceTokens = normalizeMatchText(sourceText).split(" ").filter(Boolean);
  for (let start = 0; start < sourceTokens.length; start += 1) {
    let compactSpan = "";
    for (let end = start; end < sourceTokens.length && compactSpan.length < compactPhrase.length; end += 1) {
      compactSpan += compactMatchText(sourceTokens[end]);
      if (compactSpan === compactPhrase) {
        return true;
      }
    }
  }
  return false;
}

function uniqueValues(values) {
  return Array.from(new Set(values.map((value) => cleanString(value, 200)).filter(Boolean)));
}

function getInvestmentAliasValues(investment) {
  const aliases = [
    investment && investment.company,
    investment && investment.legalName,
    investment && investment.fundName,
    investment && investment.investmentName,
    investment && investment.companyKey
  ].concat(Array.isArray(investment && investment.aliases) ? investment.aliases : [])
    .concat(Array.isArray(investment && investment.investmentAliases) ? investment.investmentAliases : []);
  return uniqueValues(aliases);
}

function getRootDomain(sender) {
  const emailOrDomain = cleanString(sender, 240).toLowerCase();
  const domain = (emailOrDomain.match(/@([^>\s]+)/) || [])[1] || emailOrDomain;
  const cleanDomain = domain.replace(/^https?:\/\//, "").replace(/^www\./, "").split(/[/?#]/)[0];
  const parts = cleanDomain.split(".").filter(Boolean);
  return parts.length >= 2 ? parts.slice(-2, -1)[0] || "" : "";
}

function findMatchedAlias(sourceParts, aliases) {
  for (const alias of aliases) {
    if (hasExplicitPhrase(sourceParts.body, alias)) return { alias, location: "source body", weight: 100 };
    if (hasExplicitPhrase(sourceParts.subject, alias)) return { alias, location: "subject", weight: 92 };
    if (hasExplicitPhrase(sourceParts.filename, alias)) return { alias, location: "filename", weight: 72 };
  }
  return null;
}

function scoreDomainEvidence(sender, aliases) {
  const rootDomain = getRootDomain(sender);
  if (!rootDomain || rootDomain.length < 4) return null;
  const matchedAlias = aliases.find((alias) => {
    const compactAlias = compactMatchText(alias);
    return compactAlias.length >= 4 && compactAlias === compactMatchText(rootDomain);
  });
  return matchedAlias
    ? { alias: matchedAlias, domain: rootDomain, weight: 18, reason: `Sender domain '${rootDomain}' supports '${matchedAlias}'.` }
    : null;
}

function generateInvestmentMatchCandidates({ source, investments = [] }) {
  const sourceParts = {
    body: cleanString(source && source.sourceText, 60000),
    subject: cleanString(source && source.subject, 240),
    sender: cleanString(source && source.sender, 240),
    filename: cleanString(source && source.filename, 240)
  };
  const candidates = investments.map((investment) => {
    const aliases = getInvestmentAliasValues(investment);
    const aliasMatch = findMatchedAlias(sourceParts, aliases);
    const domainEvidence = scoreDomainEvidence(sourceParts.sender, aliases);
    const score = (aliasMatch ? aliasMatch.weight : 0) + (domainEvidence ? domainEvidence.weight : 0);
    const evidence = [];
    if (aliasMatch) evidence.push(`Exact ${aliasMatch.location} match for '${aliasMatch.alias}'.`);
    if (domainEvidence) evidence.push(domainEvidence.reason);
    const evidenceTypes = [];
    if (aliasMatch) {
      evidenceTypes.push(
        aliasMatch.location === "source body"
          ? "sourceBody"
          : aliasMatch.location === "subject"
            ? "subject"
            : "attachmentFilename"
      );
    }
    if (domainEvidence) evidenceTypes.push("senderDomain");
    return {
      investment,
      investmentId: cleanString(investment && investment.id, 200),
      investmentName: cleanString(investment && investment.company, 200),
      entityName: cleanString(investment && investment.entity, 120),
      score,
      hasExplicitNameEvidence: Boolean(aliasMatch),
      hasDomainEvidence: Boolean(domainEvidence),
      evidenceTypes,
      matchedAlias: aliasMatch ? aliasMatch.alias : "",
      reason: evidence.join(" ")
    };
  }).filter((candidate) => candidate.score > 0)
    .sort((left, right) => right.score - left.score || left.investmentName.localeCompare(right.investmentName));
  const best = candidates[0] || null;
  const runnerUp = candidates[1] || null;
  const hasCompetingCandidate = Boolean(
    best && runnerUp && (runnerUp.hasExplicitNameEvidence || runnerUp.score >= best.score - 12)
  );
  return {
    candidates,
    explicitCandidates: candidates.filter((candidate) => candidate.hasExplicitNameEvidence),
    best,
    hasCompetingCandidate,
    status: hasCompetingCandidate
      ? "ambiguous"
      : best && best.hasExplicitNameEvidence
        ? "existing-confident"
        : best
          ? "existing-possible"
          : "no-match"
  };
}

function confidenceFromDeterministicCandidate(candidate, hasCompetingCandidate) {
  if (!candidate) return 0;
  if (candidate.hasExplicitNameEvidence && /filename/i.test(candidate.reason || "")) {
    return hasCompetingCandidate ? 72 : candidate.hasDomainEvidence ? 84 : 78;
  }
  if (candidate.hasExplicitNameEvidence && !hasCompetingCandidate) return candidate.hasDomainEvidence ? 98 : 96;
  if (candidate.hasExplicitNameEvidence) return 88;
  if (candidate.hasDomainEvidence && !hasCompetingCandidate) return 78;
  return 62;
}

module.exports = {
  SEMANTIC_ONLY_CONFIDENCE_CAP,
  compactMatchText,
  confidenceFromDeterministicCandidate,
  generateInvestmentMatchCandidates,
  getInvestmentAliasValues,
  getRootDomain,
  hasExplicitPhrase,
  normalizeMatchText
};
