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
  return senderAllowlistOutcome(sender, allowedSenders, allowedDomains).allowed;
}

function senderAllowlistOutcome(sender, allowedSenders, allowedDomains) {
  const email = cleanString(sender, 320).toLowerCase();
  const domain = email.includes("@") ? email.split("@").pop() : "";
  const senderMatch = Boolean(email) && (!allowedSenders.length || allowedSenders.includes(email));
  const domainMatch = Boolean(email) && (!allowedDomains.length || allowedDomains.includes(domain));
  return {
    allowed: senderMatch && domainMatch,
    senderRuleConfigured: allowedSenders.length > 0,
    senderMatch,
    domainRuleConfigured: allowedDomains.length > 0,
    domainMatch
  };
}

function previewMessageEligibility(message, stateEntry, allowlist, now = new Date()) {
  if (!allowlist.allowed) {
    return { status: "skipped", reason: "Sender is outside the configured intake allowlist." };
  }
  if (!cleanString(message && (message.internetMessageId || message.graphMessageId || message.id), 500)) {
    return { status: "blocked", reason: "Message has no stable Internet Message ID or Graph ID." };
  }
  if (!stateEntry) {
    return { status: "eligible", reason: "No prior intake state exists for this source message." };
  }
  if (stateEntry.status === "processed") {
    return { status: "already-processed", reason: "Message has already been processed." };
  }
  if (stateEntry.status === "skipped") {
    return { status: "already-skipped", reason: "Message has terminal skipped intake state." };
  }
  if (stateEntry.status === "reserved") {
    const reservedAt = new Date(stateEntry.reservedAt || stateEntry.processedAt || 0).getTime();
    if (Number.isFinite(reservedAt) && now.getTime() - reservedAt < 15 * 60 * 1000) {
      return { status: "reserved", reason: "Message intake is already reserved and still within the 15-minute reservation window." };
    }
    return { status: "eligible", reason: "The previous reservation is stale and intake would reclaim this message." };
  }
  if (stateEntry.status === "failed") {
    return { status: "eligible", reason: "The previous intake attempt failed and this message is eligible for retry." };
  }
  return { status: "eligible", reason: `State '${stateEntry.status || "unknown"}' is not terminal and would be reclaimed.` };
}

function createProposalPayload({ analysis, source, document, documents }) {
  return {
    investmentId: analysis.investmentMatch && analysis.investmentMatch.investmentId,
    entityId: analysis.entityMatch && (analysis.entityMatch.entityId || analysis.entityMatch.entityName),
    sourceType: source.sourceType,
    sourceIdentifier: source.sourceIdentifier,
    sourceDate: source.sourceDate,
    sender: source.sender,
    subject: source.subject,
    opportunityName: source.opportunityName || "",
    opportunityId: source.opportunityId || "",
    opportunityIdentityKeys: Array.isArray(source.opportunityIdentityKeys) ? source.opportunityIdentityKeys : [],
    sourceMessageKey: source.sourceMessageKey || "",
    sourceMessageKeys: source.sourceMessageKey ? [source.sourceMessageKey] : [],
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
    documents: Array.isArray(documents) ? documents : document ? [document] : [],
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
  analyzePotentialNewDeal,
  analyzePotentialNewDeals,
  extractPdfTextFromUpload,
  finalizeAnalysisForResponse,
  enforceProposalSafetyInvariant,
  saveAiUpdateProposal,
  sourceProposalSnapshot,
  reconcileSourceProposals,
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

  async function prepareNewDealSource(message, bodyText) {
    const textParts = [
      message.subject ? `Subject: ${message.subject}` : "",
      message.senderName ? `Sender name: ${message.senderName}` : "",
      message.sender ? `Sender email: ${message.sender}` : "",
      bodyText
    ].filter(Boolean);
    const pdfExtractions = new Map();
    for (const attachment of message.pdfAttachments || []) {
      try {
        const extracted = await extractPdfTextFromUpload({
          filename: attachment.name,
          mimeType: attachment.contentType || "application/pdf",
          fileData: attachment.contentBytes
        });
        pdfExtractions.set(attachment.id || attachment.name, extracted);
      } catch (error) {
        pdfExtractions.set(attachment.id || attachment.name, { error: error.message || "PDF could not be parsed." });
      }
    }
    return {
      source: {
        sourceType: "Email",
        sender: message.sender,
        senderName: message.senderName,
        subject: message.subject,
        sourceDate: message.receivedDateTime,
        sourceIdentifier: ["Microsoft 365", message.sender, message.subject, message.receivedDateTime]
          .map((item) => cleanString(item, 500)).filter(Boolean).join(" | "),
        filename: (message.attachments || []).map((attachment) => attachment.name).filter(Boolean).join(" | "),
        sourceText: textParts.join("\n\n").slice(0, 60000),
        emailBodyText: bodyText,
        attachments: (message.pdfAttachments || []).map((attachment) => {
          const extraction = pdfExtractions.get(attachment.id || attachment.name);
          return {
            id: attachment.id || attachment.name,
            name: attachment.name,
            text: extraction && !extraction.error ? extraction.combinedText : ""
          };
        }).filter((attachment) => attachment.text),
        graphMessageId: message.id,
        internetMessageId: message.internetMessageId
      },
      pdfExtractions
    };
  }

  function preserveNewDealAttachments(message, analyzedAt, pdfExtractions) {
    const documents = [];
    const attachmentHashes = [];
    for (const attachment of message.attachments || []) {
      const hash = attachmentHash(attachment.contentBytes);
      if (hash) attachmentHashes.push(hash);
      const existing = hash && typeof stateService.findAttachmentByHash === "function"
        ? stateService.findAttachmentByHash(hash)
        : null;
      let document = existing;
      if (!document && typeof saveUpload === "function") {
        document = saveUpload({
          filename: attachment.name,
          buffer: Buffer.from(attachment.contentBytes, "base64"),
          uploadedAt: analyzedAt,
          source: "microsoft-365-email"
        });
      }
      const extraction = pdfExtractions.get(attachment.id || attachment.name);
      documents.push({
        ...(document || {}),
        name: attachment.name,
        contentType: attachment.contentType,
        size: attachment.size,
        hash,
        preservationStatus: document ? "preserved" : "unresolved",
        extractionStatus: attachment.isPdf
          ? extraction && !extraction.error ? "parsed" : "unresolved"
          : "not-parsed",
        reason: extraction && extraction.error ? extraction.error : "",
        graphAttachmentId: attachment.id,
        sourceMessageKey: stateService.messageDedupeKey(message)
      });
    }
    for (const attachment of message.unresolvedAttachments || []) {
      documents.push({
        id: attachment.id || makeId(),
        name: attachment.name || "Unresolved attachment",
        contentType: attachment.contentType || "",
        size: attachment.size || 0,
        preservationStatus: "unresolved",
        extractionStatus: "not-parsed",
        reason: attachment.reason || "Attachment could not be preserved.",
        graphAttachmentId: attachment.id || "",
        sourceMessageKey: stateService.messageDedupeKey(message)
      });
    }
    return { documents, attachmentHashes };
  }

  async function processPotentialNewDeals({
    message,
    bodyText,
    investments,
    entitiesForUser,
    analyzedAt,
    replacePendingSourceOpportunities = false,
    expectedSourceSnapshot = [],
    reviewer = ""
  }) {
    const prepared = await prepareNewDealSource(message, bodyText);
    const preserved = preserveNewDealAttachments(message, analyzedAt, prepared.pdfExtractions);
    const analyzeMany = typeof analyzePotentialNewDeals === "function"
      ? analyzePotentialNewDeals
      : async ({ source, investments: availableInvestments }) => [{
          opportunity: { name: source.subject || "Opportunity", opportunityId: "" },
          source,
          result: await analyzePotentialNewDeal({ source, investments: availableInvestments })
        }];
    const partitions = await analyzeMany({ source: prepared.source, investments });
    const proposals = [];
    const children = [];
    const sourceMessageKey = stateService.messageDedupeKey(message);
    for (const partition of partitions) {
      const opportunityDocuments = preserved.documents.filter((document) =>
        !partition.opportunity.attachmentIds || partition.opportunity.attachmentIds.includes(document.graphAttachmentId)
      );
      if (partition.result.route === "existing-investment") {
        const updateResult = await analyzeSource({
          source: {
            ...partition.source,
            analyzedAt,
            sourceMessageKey,
            opportunityName: partition.opportunity.name,
            opportunityId: partition.opportunity.opportunityId,
            opportunityIdentityKeys: partition.opportunity.opportunityIdentityKeys
          },
          investments,
          entitiesForUser,
          documents: opportunityDocuments,
          proposalSaveOptions: { replacePendingSourceOpportunity: replacePendingSourceOpportunities },
          stageOnly: replacePendingSourceOpportunities
        });
        if (updateResult.proposal) proposals.push(updateResult.proposal);
        children.push({
          type: "opportunity",
          opportunityId: partition.opportunity.opportunityId,
          opportunityName: partition.opportunity.name,
          route: "existing-investment",
          status: updateResult.proposal ? "processed" : "skipped",
          proposalId: updateResult.proposal && updateResult.proposal.id,
          reason: updateResult.reason
        });
        continue;
      }
      if (!partition.result.analysis || partition.result.route === "not-a-deal") {
        children.push({ type: "opportunity", opportunityId: partition.opportunity.opportunityId, opportunityName: partition.opportunity.name, route: partition.result.route, status: "skipped", reason: "Partition was not classified as an investment opportunity." });
        continue;
      }
      const analysis = partition.result.analysis;
      const proposalEntry = {
        proposalType: "new-deal",
        sourceType: "Email",
        sourceIdentifier: prepared.source.sourceIdentifier,
        sourceDate: message.receivedDateTime,
        sender: message.sender,
        subject: message.subject,
        confidenceScore: analysis.matchResult.confidence,
        matchReason: analysis.matchResult.reason,
        summary: analysis.dealData.dealSummary.value || analysis.classificationReason || "Potential new deal from Microsoft 365 email.",
        dealData: analysis.dealData,
        matchResult: analysis.matchResult,
        opportunityName: partition.opportunity.name,
        opportunityId: partition.opportunity.opportunityId,
        opportunityIdentityKeys: partition.opportunity.opportunityIdentityKeys,
        opportunityFingerprint: analysis.opportunityFingerprint,
        sourceMessageKey,
        sourceMessageKeys: [sourceMessageKey],
        proposedEntity: "Beaman Ventures",
        entityConfirmed: false,
        noExistingMatchConfirmed: analysis.matchResult.status === "no-match",
        amountConfirmed: false,
        documents: opportunityDocuments,
        status: "pending"
      };
      const proposal = replacePendingSourceOpportunities
        ? proposalEntry
        : saveAiUpdateProposal(proposalEntry);
      proposals.push(proposal);
      children.push({
        type: "opportunity",
        opportunityId: partition.opportunity.opportunityId,
        opportunityName: partition.opportunity.name,
        route: partition.result.route,
        status: "processed",
        proposalId: proposal.id
      });
    }
    if (replacePendingSourceOpportunities) {
      if (typeof reconcileSourceProposals !== "function") {
        const error = new Error("Transactional source-proposal reconciliation is not configured.");
        error.statusCode = 500;
        throw error;
      }
      const reconciliation = reconcileSourceProposals({
        sourceMessageKey,
        proposals,
        expectedSnapshot: expectedSourceSnapshot,
        reviewer,
        reconciledAt: analyzedAt
      });
      const byOpportunityId = new Map(reconciliation.proposals.map((proposal) => [proposal.opportunityId, proposal]));
      children.forEach((child) => {
        const reconciled = byOpportunityId.get(child.opportunityId);
        if (reconciled) child.proposalId = reconciled.id;
      });
      return { prepared, partitions, proposals: reconciliation.proposals, children, reconciliation, ...preserved };
    }
    return { prepared, partitions, proposals, children, ...preserved };
  }

  async function analyzeSource({ source, investments, entitiesForUser, document, documents, proposalSaveOptions, stageOnly = false }) {
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
      createProposalPayload({ analysis, source: { ...result.source, ...source }, document, documents }),
      matchedInvestment
    );
    const saved = stageOnly ? proposal : saveAiUpdateProposal(proposal, proposalSaveOptions);
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
        const reservation = typeof stateService.claimMessage === "function"
          ? stateService.claimMessage(message)
          : { claimed: !stateService.findByMessage(message), entry: stateService.findByMessage(message) };
        if (!reservation.claimed) {
          results.push({
            ...base,
            status: "skipped",
            reason: reservation.reason || "Duplicate message already processed.",
            proposalIds: reservation.entry && reservation.entry.proposalIds || []
          });
          continue;
        }

        const analyzedAt = new Date().toISOString();
        const bodyText = normalizeEmailBody(message);
        const proposalIds = [];
        const attachmentHashes = [];
        const analysisAudits = [];
        const childResults = [];

        if (typeof analyzePotentialNewDeals === "function" || typeof analyzePotentialNewDeal === "function") {
          const newDealResult = await processPotentialNewDeals({ message, bodyText, investments, entitiesForUser, analyzedAt });
          proposalIds.push(...newDealResult.proposals.map((proposal) => proposal.id));
          attachmentHashes.push(...newDealResult.attachmentHashes);
          childResults.push(...newDealResult.children);
          const failedOpportunity = newDealResult.children.some((child) => child.status === "failed");
          const processedOpportunity = newDealResult.children.some((child) => child.status === "processed");
          const status = processedOpportunity ? "processed" : failedOpportunity ? "failed" : "skipped";
          const reason = processedOpportunity ? "" : newDealResult.children.map((child) => child.reason).filter(Boolean).join("; ") || "Email was not classified as an investment opportunity.";
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
            attachments: newDealResult.documents || [],
            processedAt: analyzedAt,
            proposalIds,
            status,
            error: status === "failed" ? reason : ""
          });
          results.push({ ...base, status, reason, proposalIds, attachmentHashes, children: childResults });
          continue;
        }

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

  async function previewEmails({ now = new Date() } = {}) {
    if (!graphMailService || !graphMailService.isConfigured() || typeof graphMailService.previewIntakeMessages !== "function") {
      return {
        configured: false,
        error: "Microsoft 365 email intake preview is not configured.",
        messages: []
      };
    }
    let run;
    try {
      run = await graphMailService.previewIntakeMessages();
    } catch (error) {
      return {
        configured: true,
        error: error.message || "Microsoft 365 email intake preview failed.",
        messages: []
      };
    }
    const messages = run.messages.map((message) => {
      const lookupMessage = {
        id: message.graphMessageId,
        graphMessageId: message.graphMessageId,
        internetMessageId: message.internetMessageId
      };
      const stateEntry = stateService.findByMessage(lookupMessage);
      const allowlist = senderAllowlistOutcome(message.sender, senderAllowlist, domainAllowlist);
      const eligibility = previewMessageEligibility(lookupMessage, stateEntry, allowlist, now);
      return {
        ...message,
        allowlist,
        eligibilityStatus: eligibility.status,
        eligibilityReason: eligibility.reason,
        state: stateEntry
          ? {
              found: true,
              status: stateEntry.status,
              reservedAt: stateEntry.reservedAt,
              processedAt: stateEntry.processedAt,
              proposalIds: stateEntry.proposalIds,
              error: stateEntry.error
            }
          : { found: false, status: "", reservedAt: "", processedAt: "", proposalIds: [], error: "" }
      };
    });
    return {
      configured: true,
      mailbox: run.mailbox,
      folderName: run.folder.displayName,
      maxMessagesPerRun: run.maxMessagesPerRun,
      limits: run.limits,
      messages
    };
  }

  async function reanalyzeMessage({ user, graphMessageId, expectedInternetMessageId = "" }) {
    if (!graphMailService || !graphMailService.isConfigured() || typeof graphMailService.fetchIntakeMessageById !== "function") {
      const error = new Error("Microsoft 365 exact-message reanalysis is not configured.");
      error.statusCode = 400;
      throw error;
    }
    const message = await graphMailService.fetchIntakeMessageById(graphMessageId);
    if (expectedInternetMessageId && message.internetMessageId !== expectedInternetMessageId) {
      const error = new Error("The fetched source message does not match the preserved Internet Message ID.");
      error.statusCode = 409;
      throw error;
    }
    if (!senderAllowed(message.sender, senderAllowlist, domainAllowlist)) {
      const error = new Error("The source sender is outside the configured intake allowlist.");
      error.statusCode = 403;
      throw error;
    }
    const investments = filterInvestmentsForUser(readInvestments(), user);
    const entitiesForUser = entities.filter((entity) => canViewEntity(user, entity));
    const analyzedAt = new Date().toISOString();
    const bodyText = normalizeEmailBody(message);
    const sourceMessageKey = stateService.messageDedupeKey(message);
    const expectedSourceSnapshot = typeof sourceProposalSnapshot === "function"
      ? sourceProposalSnapshot(sourceMessageKey)
      : [];
    const processed = await processPotentialNewDeals({
      message,
      bodyText,
      investments,
      entitiesForUser,
      analyzedAt,
      replacePendingSourceOpportunities: true,
      expectedSourceSnapshot,
      reviewer: user && user.email
    });
    const proposalIds = processed.proposals.map((proposal) => proposal.id);
    if (!proposalIds.length) {
      const error = new Error("Reanalysis did not produce any separately reviewable proposals.");
      error.statusCode = 422;
      throw error;
    }
    stateService.upsertEntry({
      graphMessageId: message.id,
      internetMessageId: message.internetMessageId,
      conversationId: message.conversationId,
      mailbox: message.mailbox,
      folderId: message.folderId,
      subject: message.subject,
      sender: message.sender,
      receivedDateTime: message.receivedDateTime,
      attachmentHashes: processed.attachmentHashes,
      attachments: processed.documents,
      processedAt: analyzedAt,
      proposalIds,
      status: "processed",
      error: ""
    }, { replaceProposalIds: true });
    return {
      graphMessageId: message.id,
      internetMessageId: message.internetMessageId,
      subject: message.subject,
      proposalIds,
      proposalsCreated: proposalIds.length,
      refreshedProposalIds: processed.reconciliation.refreshedProposalIds,
      supersededProposalIds: processed.reconciliation.supersededProposalIds,
      children: processed.children
    };
  }

  return {
    checkForNewEmails,
    isMeaningfulBody,
    normalizeEmailBody,
    previewEmails,
    reanalyzeMessage
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
  normalizeEmailBody,
  previewMessageEligibility,
  senderAllowlistOutcome
};
