const crypto = require("crypto");

function cleanString(value, maxLength = 2000) {
  return String(value || "").trim().slice(0, maxLength);
}

function htmlToText(html) {
  return cleanString(html, 60000)
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&#39;/gi, "'")
    .replace(/&quot;/gi, '"')
    .replace(/[ \t]+/g, " ")
    .replace(/\n\s+/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function stripQuotedEmailText(text) {
  return cleanString(text, 60000)
    .split(/\n(?=From:|On .+ wrote:|-----Original Message-----)/i)[0]
    .split(/\n--\s*$/m)[0]
    .replace(/\n(sent from my iphone|sent from outlook for ios)[\s\S]*$/i, "")
    .trim();
}

function normalizeEmailBody(message) {
  const rawBody = cleanString(message && message.body, 60000);
  const contentType = cleanString(message && message.bodyContentType, 40).toLowerCase();
  const text = contentType === "html" || /<\/?[a-z][\s\S]*>/i.test(rawBody)
    ? htmlToText(rawBody)
    : rawBody;
  return stripQuotedEmailText(text || cleanString(message && message.bodyPreview, 2000));
}

function isMeaningfulBody(text) {
  const compact = cleanString(text, 60000)
    .replace(/\s+/g, " ")
    .trim();
  if (compact.length < 80) {
    return false;
  }
  if (/^(please\s+)?(see|review|find)\s+(the\s+)?attached/i.test(compact) && compact.length < 160) {
    return false;
  }
  return /revenue|ebitda|runway|cash|customer|pipeline|sales|units|update|quarter|investor|board|financial|development|milestone|risk|capital|valuation|mrr|arr/i.test(compact);
}

function attachmentHash(contentBytes) {
  return crypto
    .createHash("sha256")
    .update(Buffer.from(cleanString(contentBytes, 20 * 1024 * 1024), "base64"))
    .digest("hex");
}

function resultSkeleton(message) {
  return {
    graphMessageId: cleanString(message && message.id, 500),
    internetMessageId: cleanString(message && message.internetMessageId, 500),
    conversationId: cleanString(message && message.conversationId, 500),
    subject: cleanString(message && message.subject, 500),
    sender: cleanString(message && message.sender, 320),
    receivedDateTime: cleanString(message && message.receivedDateTime, 80),
    status: "skipped",
    reason: "",
    proposalIds: [],
    attachmentHashes: []
  };
}

function summarizeCounts(results) {
  return results.reduce(
    (counts, result) => {
      counts.checked += result.kind === "message" ? 1 : 0;
      if (result.status === "processed") counts.processed += 1;
      if (result.status === "skipped") counts.skipped += 1;
      if (result.status === "failed") counts.failed += 1;
      counts.proposalsCreated += result.status === "processed" && Array.isArray(result.proposalIds)
        ? result.proposalIds.length
        : 0;
      return counts;
    },
    { checked: 0, processed: 0, skipped: 0, failed: 0, proposalsCreated: 0 }
  );
}

function normalizeSenderList(value) {
  return cleanString(value, 4000)
    .split(",")
    .map((item) => item.trim().toLowerCase())
    .filter(Boolean);
}

function senderAllowed(sender, allowedSenders, allowedDomains) {
  const email = cleanString(sender, 320).toLowerCase();
  if (!email) {
    return false;
  }
  if (allowedSenders.length && !allowedSenders.includes(email)) {
    return false;
  }
  const domain = email.includes("@") ? email.split("@").pop() : "";
  if (allowedDomains.length && !allowedDomains.includes(domain)) {
    return false;
  }
  return true;
}

function createProposalPayload({ analysis, source, document }) {
  return {
    investmentId: analysis.investmentMatch && analysis.investmentMatch.investmentId,
    entityId: analysis.entityMatch && (analysis.entityMatch.entityId || analysis.entityMatch.entityName),
    sourceType: source.sourceType,
    sourceIdentifier: source.sourceIdentifier,
    sourceDate: source.sourceDate,
    sender: source.sender,
    subject: source.subject,
    confidenceScore: analysis.investmentMatch ? analysis.investmentMatch.confidence : 0,
    matchReason: [
      analysis.investmentMatch && analysis.investmentMatch.reason,
      analysis.entityMatch && analysis.entityMatch.reason
    ].filter(Boolean).join(" "),
    summary: Array.isArray(analysis.whatChanged) && analysis.whatChanged.length
      ? analysis.whatChanged.map((item) => `• ${item}`).join("\n")
      : "AI email intake analysis completed; review extracted facts and proposed changes.",
    extractedData: {
      facts: analysis.extractedFacts || [],
      warnings: analysis.warnings || [],
      unresolved: analysis.unresolved || [],
      candidates: analysis.candidates || [],
      materialDevelopments: analysis.materialDevelopments || [],
      unverifiedClaims: analysis.unverifiedClaims || [],
      source: {
        filename: source.filename || "",
        pageCount: source.pageCount || 0,
        diagnostics: source.diagnostics || {},
        graphMessageId: source.graphMessageId || "",
        internetMessageId: source.internetMessageId || ""
      },
      analyzedAt: source.analyzedAt
    },
    proposedChanges: analysis.proposedChanges || [],
    documents: document ? [document] : [],
    status: "pending"
  };
}

function shouldCreateProposal(analysis) {
  return Boolean(
    analysis &&
    analysis.investmentMatch &&
    analysis.investmentMatch.investmentId &&
    (
      (Array.isArray(analysis.proposedChanges) && analysis.proposedChanges.length) ||
      (Array.isArray(analysis.materialDevelopments) && analysis.materialDevelopments.length) ||
      (Array.isArray(analysis.extractedFacts) && analysis.extractedFacts.length)
    )
  );
}

function hasAutomatedExplicitInvestmentMatch(analysis) {
  const investmentMatch = analysis && analysis.investmentMatch ? analysis.investmentMatch : {};
  if (!investmentMatch.investmentId) {
    return false;
  }

  const warnings = []
    .concat(Array.isArray(analysis && analysis.warnings) ? analysis.warnings : [])
    .concat(Array.isArray(analysis && analysis.unresolved) ? analysis.unresolved : [])
    .map((item) => cleanString(item, 500).toLowerCase());
  if (warnings.some((warning) =>
    warning.includes("lacks explicit portfolio-name or alias evidence") ||
    warning.includes("lacks explicit name, alias, sender-domain, or subject evidence")
  )) {
    return false;
  }

  const reason = cleanString(investmentMatch.reason, 1000).toLowerCase();
  if (/exact .+ match for|sender domain .+ supports/.test(reason)) {
    return true;
  }

  return (Array.isArray(analysis.candidates) ? analysis.candidates : []).some((candidate) =>
    candidate &&
    candidate.investmentId === investmentMatch.investmentId &&
    /exact .+ match for|sender domain .+ supports/i.test(cleanString(candidate.reason, 1000))
  );
}

function getDeterministicEvidenceTypes(analysis) {
  const investmentId = analysis && analysis.investmentMatch && analysis.investmentMatch.investmentId;
  const reasons = [analysis && analysis.investmentMatch && analysis.investmentMatch.reason]
    .concat(
      (Array.isArray(analysis && analysis.candidates) ? analysis.candidates : [])
        .filter((candidate) => !investmentId || candidate.investmentId === investmentId)
        .map((candidate) => candidate.reason)
    )
    .map((reason) => cleanString(reason, 1000).toLowerCase())
    .filter(Boolean);
  const types = [];
  if (reasons.some((reason) => reason.includes("exact subject match"))) {
    types.push("subject");
  }
  if (reasons.some((reason) => reason.includes("sender domain"))) {
    types.push("senderDomain");
  }
  if (reasons.some((reason) => reason.includes("exact source body match"))) {
    types.push("sourceBody");
  }
  if (reasons.some((reason) => reason.includes("exact filename match"))) {
    types.push("attachmentFilename");
  }
  return types;
}

function countItems(value) {
  return Array.isArray(value) ? value.length : 0;
}

function safeWarningStrings(analysis) {
  return []
    .concat(Array.isArray(analysis && analysis.warnings) ? analysis.warnings : [])
    .concat(Array.isArray(analysis && analysis.unresolved) ? analysis.unresolved : [])
    .map((item) => {
      if (typeof item === "string") {
        return cleanString(item, 500);
      }
      if (item && typeof item === "object") {
        return cleanString(item.message || item.reason || item.warning, 500);
      }
      return "";
    })
    .filter(Boolean)
    .slice(0, 8);
}

function buildSkippedAnalysisAudit({ analysis, source, reason, shouldCreateProposalResult }) {
  const investmentMatch = analysis && analysis.investmentMatch ? analysis.investmentMatch : {};
  const entityMatch = analysis && analysis.entityMatch ? analysis.entityMatch : {};
  return {
    sourceType: cleanString(source && source.sourceType, 80),
    subject: cleanString(source && source.subject, 500),
    sender: cleanString(source && source.sender, 320),
    receivedDateTime: cleanString(source && source.sourceDate, 80),
    filename: cleanString(source && source.filename, 500),
    graphMessageId: cleanString(source && source.graphMessageId, 500),
    internetMessageId: cleanString(source && source.internetMessageId, 500),
    match: {
      finalMatchedInvestmentId: cleanString(investmentMatch.investmentId, 120),
      finalMatchedInvestmentName: cleanString(investmentMatch.investmentName, 200),
      matchedEntity: cleanString(entityMatch.entityName || entityMatch.entityId, 120),
      matchConfidence: Number(investmentMatch.confidence) || 0,
      matchReason: cleanString(investmentMatch.reason, 1000),
      explicitMatch: hasAutomatedExplicitInvestmentMatch(analysis),
      deterministicEvidenceTypes: getDeterministicEvidenceTypes(analysis)
    },
    counts: {
      extractedFacts: countItems(analysis && analysis.extractedFacts),
      materialDevelopments: countItems(analysis && analysis.materialDevelopments),
      proposedChanges: countItems(analysis && analysis.proposedChanges),
      unverifiedClaims: countItems(analysis && analysis.unverifiedClaims),
      warnings: countItems(analysis && analysis.warnings) + countItems(analysis && analysis.unresolved)
    },
    warnings: safeWarningStrings(analysis),
    shouldCreateProposal: Boolean(shouldCreateProposalResult),
    skipReason: cleanString(reason, 1000),
    analyzedAt: cleanString(source && source.analyzedAt, 80)
  };
}

function createAiEmailIntakeService({
  graphMailService,
  stateService,
  analyzeInvestmentUpdate,
  extractPdfTextFromUpload,
  finalizeAnalysisForResponse,
  enforceProposalSafetyInvariant,
  saveAiUpdateProposal,
  readInvestments,
  filterInvestmentsForUser,
  entities = [],
  canViewEntity = () => true,
  makeId = () => crypto.randomUUID(),
  saveUpload,
  allowedSenders = "",
  allowedDomains = ""
}) {
  const senderAllowlist = normalizeSenderList(allowedSenders);
  const domainAllowlist = normalizeSenderList(allowedDomains);

  async function analyzeSource({ source, investments, entitiesForUser, document }) {
    const result = await analyzeInvestmentUpdate({
      source,
      investments,
      entities: entitiesForUser
    });
    const analysis = finalizeAnalysisForResponse(result.analysis);
    const createProposalEligible = shouldCreateProposal(analysis);
    if (!createProposalEligible) {
      const reason = "No safely matched actionable analysis was produced.";
      return {
        proposal: null,
        analysis,
        reason,
        audit: buildSkippedAnalysisAudit({
          analysis,
          source,
          reason,
          shouldCreateProposalResult: createProposalEligible
        })
      };
    }
    if (!hasAutomatedExplicitInvestmentMatch(analysis)) {
      const reason = "Automated email intake requires explicit investment name, alias, sender-domain, subject, or filename evidence before creating a pending proposal.";
      return {
        proposal: null,
        analysis,
        reason,
        audit: buildSkippedAnalysisAudit({
          analysis,
          source,
          reason,
          shouldCreateProposalResult: createProposalEligible
        })
      };
    }
    const matchedInvestment = investments.find(
      (investment) => investment.id === (analysis.investmentMatch && analysis.investmentMatch.investmentId)
    ) || null;
    const proposal = enforceProposalSafetyInvariant(
      createProposalPayload({ analysis, source: { ...result.source, ...source }, document }),
      matchedInvestment
    );
    const saved = saveAiUpdateProposal(proposal);
    return { proposal: saved, analysis, reason: "" };
  }

  async function processBody({ message, bodyText, investments, entitiesForUser, analyzedAt }) {
    const source = {
      sourceType: "Email",
      sender: message.sender,
      subject: message.subject,
      sourceDate: message.receivedDateTime,
      sourceIdentifier: ["Microsoft 365", message.sender, message.subject, message.receivedDateTime]
        .map((item) => cleanString(item, 500))
        .filter(Boolean)
        .join(" | "),
      sourceText: [
        message.subject ? `Subject: ${message.subject}` : "",
        message.sender ? `From: ${message.sender}` : "",
        bodyText
      ].filter(Boolean).join("\n\n"),
      graphMessageId: message.id,
      internetMessageId: message.internetMessageId,
      analyzedAt
    };
    return analyzeSource({ source, investments, entitiesForUser });
  }

  async function processPdf({ message, attachment, investments, entitiesForUser, analyzedAt }) {
    const extracted = await extractPdfTextFromUpload({
      filename: attachment.name,
      mimeType: attachment.contentType || "application/pdf",
      fileData: attachment.contentBytes
    });
    const document = typeof saveUpload === "function"
      ? saveUpload({
          filename: extracted.filename,
          buffer: extracted.buffer,
          uploadedAt: analyzedAt,
          source: "microsoft-365-email"
        })
      : null;
    const sourceIdentifier = [
      "Microsoft 365 PDF",
      message.sender,
      message.subject,
      message.receivedDateTime,
      extracted.filename
    ].map((item) => cleanString(item, 500)).filter(Boolean).join(" | ");
    return analyzeSource({
      source: {
        sourceType: "PDF",
        sender: message.sender,
        subject: message.subject,
        sourceDate: message.receivedDateTime,
        sourceIdentifier,
        filename: extracted.filename,
        pageCount: extracted.pageCount,
        pages: extracted.pages,
        sourceText: extracted.combinedText,
        diagnostics: extracted.diagnostics || {},
        graphMessageId: message.id,
        internetMessageId: message.internetMessageId,
        analyzedAt
      },
      investments,
      entitiesForUser,
      document
    });
  }

  async function checkForNewEmails({ user }) {
    if (!graphMailService || !graphMailService.isConfigured()) {
      return {
        checked: 0,
        processed: 0,
        skipped: 0,
        failed: 1,
        proposalsCreated: 0,
        configured: false,
        error: "Microsoft 365 email intake is not configured.",
        results: []
      };
    }

    const investments = filterInvestmentsForUser(readInvestments(), user);
    const entitiesForUser = entities.filter((entity) => canViewEntity(user, entity));
    let run;
    try {
      run = await graphMailService.fetchIntakeMessages();
    } catch (error) {
      return {
        checked: 0,
        processed: 0,
        skipped: 0,
        failed: 1,
        proposalsCreated: 0,
        configured: true,
        error: error.message || "Microsoft 365 email intake failed.",
        results: []
      };
    }
    const results = [];

    for (const message of run.messages) {
      const base = { ...resultSkeleton(message), kind: "message" };
      try {
        if (!senderAllowed(message.sender, senderAllowlist, domainAllowlist)) {
          results.push({ ...base, status: "skipped", reason: "Sender is outside the configured intake allowlist." });
          continue;
        }
        const existing = stateService.findByMessage(message);
        if (existing && existing.status === "processed") {
          results.push({ ...base, status: "skipped", reason: "Duplicate message already processed.", proposalIds: existing.proposalIds });
          continue;
        }

        const analyzedAt = new Date().toISOString();
        const bodyText = normalizeEmailBody(message);
        const proposalIds = [];
        const attachmentHashes = [];
        const analysisAudits = [];
        const childResults = [];

        if (isMeaningfulBody(bodyText)) {
          try {
            const bodyResult = await processBody({ message, bodyText, investments, entitiesForUser, analyzedAt });
            if (bodyResult.proposal) {
              proposalIds.push(bodyResult.proposal.id);
              childResults.push({ type: "body", status: "processed", proposalId: bodyResult.proposal.id });
            } else {
              if (bodyResult.audit) {
                analysisAudits.push(bodyResult.audit);
              }
              childResults.push({ type: "body", status: "skipped", reason: bodyResult.reason });
            }
          } catch (error) {
            childResults.push({ type: "body", status: "failed", reason: error.message || "Email body analysis failed." });
          }
        } else {
          childResults.push({ type: "body", status: "skipped", reason: "Email body did not contain enough investment-update detail." });
        }

        for (const skipped of message.skippedAttachments || []) {
          childResults.push({ type: "attachment", status: "skipped", filename: skipped.name, reason: skipped.reason });
        }

        for (const attachment of message.pdfAttachments || []) {
          const hash = attachmentHash(attachment.contentBytes);
          attachmentHashes.push(hash);
          if (stateService.hasAttachmentHash(hash)) {
            childResults.push({ type: "attachment", status: "skipped", filename: attachment.name, reason: "Duplicate PDF attachment already processed." });
            continue;
          }
          try {
            const pdfResult = await processPdf({ message, attachment, investments, entitiesForUser, analyzedAt });
            if (pdfResult.proposal) {
              proposalIds.push(pdfResult.proposal.id);
              childResults.push({ type: "attachment", status: "processed", filename: attachment.name, proposalId: pdfResult.proposal.id });
            } else {
              if (pdfResult.audit) {
                analysisAudits.push(pdfResult.audit);
              }
              childResults.push({ type: "attachment", status: "skipped", filename: attachment.name, reason: pdfResult.reason });
            }
          } catch (error) {
            childResults.push({ type: "attachment", status: "failed", filename: attachment.name, reason: error.message || "PDF attachment analysis failed." });
          }
        }

        const failedChildren = childResults.filter((item) => item.status === "failed");
        const processedChildren = childResults.filter((item) => item.status === "processed");
        const status = processedChildren.length ? "processed" : failedChildren.length ? "failed" : "skipped";
        const reason = processedChildren.length
          ? ""
          : failedChildren.map((item) => item.reason).filter(Boolean).join("; ") ||
            childResults.map((item) => item.reason).filter(Boolean).join("; ") ||
            "No processable email content found.";

        stateService.upsertEntry({
          graphMessageId: message.id,
          internetMessageId: message.internetMessageId,
          conversationId: message.conversationId,
          mailbox: message.mailbox,
          folderId: message.folderId,
          subject: message.subject,
          sender: message.sender,
          receivedDateTime: message.receivedDateTime,
          attachmentHashes,
          processedAt: analyzedAt,
          proposalIds,
          status,
          error: status === "failed" ? reason : "",
          analysisAudits
        });
        results.push({ ...base, status, reason, proposalIds, attachmentHashes, children: childResults });
      } catch (error) {
        stateService.upsertEntry({
          graphMessageId: message.id,
          internetMessageId: message.internetMessageId,
          conversationId: message.conversationId,
          mailbox: message.mailbox,
          folderId: message.folderId,
          subject: message.subject,
          sender: message.sender,
          receivedDateTime: message.receivedDateTime,
          processedAt: new Date().toISOString(),
          status: "failed",
          error: error.message || "Message intake failed."
        });
        results.push({ ...base, status: "failed", reason: error.message || "Message intake failed." });
      }
    }

    return {
      ...summarizeCounts(results),
      configured: true,
      mailbox: run.mailbox,
      folderName: run.folder.displayName,
      results
    };
  }

  return {
    checkForNewEmails,
    isMeaningfulBody,
    normalizeEmailBody
  };
}

module.exports = {
  attachmentHash,
  buildSkippedAnalysisAudit,
  createAiEmailIntakeService,
  getDeterministicEvidenceTypes,
  hasAutomatedExplicitInvestmentMatch,
  htmlToText,
  isMeaningfulBody,
  normalizeEmailBody
};
