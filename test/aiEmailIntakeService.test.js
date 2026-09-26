const test = require("node:test");
const assert = require("node:assert/strict");
const {
  attachmentHash,
  createAiEmailIntakeService,
  hasAutomatedExplicitInvestmentMatch,
  isMeaningfulBody,
  normalizeEmailBody,
  previewMessageEligibility
} = require("../services/aiEmailIntakeService");
const { createAiUpdateAnalysisService } = require("../services/aiUpdateAnalysisService");
const { createAiEmailIntakeStateService } = require("../services/aiEmailIntakeStateService");
const { createAiUpdateProposalService } = require("../services/aiUpdateProposalService");
const { createNewDealAnalysisService } = require("../services/newDealAnalysisService");
const {
  formatAiUpdateProposalSummary,
  formatDealClaimValue,
  getDisplayDealClaimItems,
  normalizeAiUpdateProposalCounts
} = require("../public/aiUpdateSafety");
const bepFixture = require("./fixtures/bep-multi-opportunity.json");
const bepLegacyFiveProposals = require("./fixtures/bep-legacy-five-proposals.json");

const PDF_BYTES = Buffer.from("%PDF-1.4\nmock pdf\n").toString("base64");

function createMemoryStateService(initial = []) {
  let stored = initial;
  return {
    service: createAiEmailIntakeStateService({
      STATE_FILE: "state.json",
      readJsonFile: () => stored,
      writeJsonFile: (file, value) => {
        stored = value;
      }
    }),
    getStored: () => stored
  };
}

function createMessage(overrides = {}) {
  return {
    id: "graph-message-1",
    internetMessageId: "<graph-message-1@example.test>",
    conversationId: "conversation-1",
    mailbox: "updates@example.test",
    folderId: "folder-1",
    folderName: "AI Investment Updates",
    subject: "Healing Innovations quarterly update",
    sender: "trusted@example.test",
    receivedDateTime: "2026-08-28T12:00:00Z",
    bodyContentType: "html",
    body: "<p>Healing Innovations quarterly investor update with customer pipeline, units sold, cash runway, revenue commentary, product development milestones, operational risks, and board-level financial updates for Beaman Ventures.</p>",
    bodyPreview: "",
    attachments: [],
    pdfAttachments: [],
    skippedAttachments: [],
    unresolvedAttachments: [],
    ...overrides
  };
}

function createHarness({
  messages = [createMessage()],
  analysisFactory,
  analyzePotentialNewDeal,
  analyzePotentialNewDeals,
  exactMessage,
  extractPdfTextFromUpload,
  saveAiUpdateProposal,
  sourceProposalSnapshot,
  reconcilePendingSourceOpportunities,
  allowedSenders = "",
  allowedDomains = "",
  investments = [
    { id: "healing-id", company: "Healing Innovations", owningEntity: "Beaman Ventures" }
  ]
} = {}) {
  const savedProposals = [];
  const analyzeCalls = [];
  const finalizeCalls = [];
  const safetyCalls = [];
  const uploads = [];
  const { service: stateService, getStored } = createMemoryStateService();
  const graphMailService = {
    isConfigured: () => true,
    fetchIntakeMessages: async () => ({
      mailbox: "updates@example.test",
      folder: { id: "folder-1", displayName: "AI Investment Updates" },
      messages
    }),
    fetchIntakeMessageById: exactMessage ? async () => exactMessage : undefined
  };
  const saveProposal = saveAiUpdateProposal || ((proposal) => {
    const saved = { ...proposal, id: `proposal-${savedProposals.length + 1}` };
    savedProposals.push(saved);
    return saved;
  });
  const service = createAiEmailIntakeService({
    graphMailService,
    stateService,
    analyzeInvestmentUpdate: async ({ source, investments, entities }) => {
      analyzeCalls.push({ source, investments, entities });
      const analysis = analysisFactory
        ? analysisFactory({ source, investments, entities })
        : {
            investmentMatch: {
              investmentId: "healing-id",
              investmentName: "Healing Innovations",
              confidence: 96,
              reason: "Exact source body match for 'Healing Innovations'."
            },
            deterministicEvidence: {
              investmentId: "healing-id",
              types: ["sourceBody"]
            },
            entityMatch: {
              entityId: "beaman-ventures",
              entityName: "Beaman Ventures",
              confidence: 100,
              reason: "Source names Beaman Ventures."
            },
            extractedFacts: [],
            proposedChanges: [],
            materialDevelopments: [
              {
                summary: "Total Units Sold: 57 (+5).",
                evidenceStatus: "verified",
                sourceEvidence: "Total Units Sold: 57 (+5)",
                confidence: 95
              }
            ],
            unverifiedClaims: [],
            warnings: [],
            whatChanged: ["Total Units Sold: 57 (+5)."]
          };
      return { source, analysis };
    },
    analyzePotentialNewDeal,
    analyzePotentialNewDeals,
    extractPdfTextFromUpload: extractPdfTextFromUpload || (async ({ filename, fileData }) => ({
      filename,
      buffer: Buffer.from(fileData, "base64"),
      pageCount: 1,
      pages: [{ pageNumber: 1, text: "Healing Innovations Total Units Sold: 57 (+5)." }],
      combinedText: "Healing Innovations Total Units Sold: 57 (+5).",
      diagnostics: { extractionMethod: "mock" }
    })),
    finalizeAnalysisForResponse: (analysis) => {
      finalizeCalls.push(analysis);
      return { ...analysis, finalized: true };
    },
    enforceProposalSafetyInvariant: (proposal) => {
      safetyCalls.push(proposal);
      return { ...proposal, safetyChecked: true };
    },
    saveAiUpdateProposal: saveProposal,
    sourceProposalSnapshot: sourceProposalSnapshot || (() => []),
    reconcilePendingSourceOpportunities: reconcilePendingSourceOpportunities || (({ canonicalEntries }) => ({
      proposals: canonicalEntries.map((proposal) => saveProposal(proposal)),
      supersededProposals: []
    })),
    readInvestments: () => investments,
    filterInvestmentsForUser: (investments) => investments,
    entities: [{ id: "beaman-ventures", name: "Beaman Ventures" }],
    canViewEntity: () => true,
    saveUpload: ({ filename, buffer, uploadedAt }) => {
      const document = {
        id: `doc-${uploads.length + 1}`,
        name: filename,
        url: `/uploads/${filename}`,
        storedName: filename,
        uploadedAt,
        size: buffer.length
      };
      uploads.push(document);
      return document;
    },
    allowedSenders,
    allowedDomains
  });

  return {
    analyzeCalls,
    finalizeCalls,
    getStored,
    safetyCalls,
    savedProposals,
    service,
    uploads
  };
}

test("email body intake normalizes HTML and uses the existing AI analysis and proposal safety pipeline", async () => {
  const harness = createHarness();

  const result = await harness.service.checkForNewEmails({ user: { email: "editor@example.test" } });

  assert.equal(result.checked, 1);
  assert.equal(result.processed, 1);
  assert.equal(result.proposalsCreated, 1);
  assert.equal(harness.analyzeCalls.length, 1);
  assert.equal(harness.analyzeCalls[0].source.sourceType, "Email");
  assert.match(harness.analyzeCalls[0].source.sourceText, /Healing Innovations quarterly investor update/);
  assert.equal(harness.finalizeCalls.length, 1);
  assert.equal(harness.safetyCalls.length, 1);
  assert.equal(harness.savedProposals[0].status, "pending");
  assert.equal(harness.savedProposals[0].safetyChecked, true);
  assert.equal(harness.savedProposals[0].sourceType, "Email");
  assert.equal(harness.savedProposals[0].investmentId, "healing-id");
  assert.deepEqual(harness.getStored()[0].proposalIds, ["proposal-1"]);
});

test("PDF attachment intake extracts text and creates one pending proposal per PDF", async () => {
  const message = createMessage({
    body: "Please see the attached board materials.",
    pdfAttachments: [
      { id: "pdf-1", name: "Board Deck A.pdf", contentType: "application/pdf", contentBytes: PDF_BYTES },
      { id: "pdf-2", name: "Board Deck B.pdf", contentType: "application/pdf", contentBytes: Buffer.from("second pdf").toString("base64") }
    ]
  });
  const harness = createHarness({ messages: [message] });

  const result = await harness.service.checkForNewEmails({ user: { email: "editor@example.test" } });

  assert.equal(result.checked, 1);
  assert.equal(result.processed, 1);
  assert.equal(result.proposalsCreated, 2);
  assert.deepEqual(harness.analyzeCalls.map((call) => call.source.sourceType), ["PDF", "PDF"]);
  assert.deepEqual(harness.savedProposals.map((proposal) => proposal.sourceType), ["PDF", "PDF"]);
  assert.equal(harness.savedProposals[0].documents.length, 1);
  assert.equal(harness.uploads.length, 2);
  assert.equal(harness.getStored()[0].attachmentHashes.length, 2);
});

test("repeated checks do not duplicate an already processed message", async () => {
  const harness = createHarness();

  const first = await harness.service.checkForNewEmails({ user: { email: "editor@example.test" } });
  const second = await harness.service.checkForNewEmails({ user: { email: "editor@example.test" } });

  assert.equal(first.proposalsCreated, 1);
  assert.equal(second.proposalsCreated, 0);
  assert.equal(second.skipped, 1);
  assert.equal(second.results[0].reason, "Duplicate message already processed.");
  assert.equal(harness.savedProposals.length, 1);
});

test("duplicate PDF attachment hashes are skipped across messages", async () => {
  const messages = [
    createMessage({
      id: "message-1",
      internetMessageId: "<message-1@example.test>",
      body: "Please see attached.",
      pdfAttachments: [{ id: "pdf-1", name: "Deck.pdf", contentType: "application/pdf", contentBytes: PDF_BYTES }]
    }),
    createMessage({
      id: "message-2",
      internetMessageId: "<message-2@example.test>",
      body: "Please see attached.",
      pdfAttachments: [{ id: "pdf-2", name: "Deck Copy.pdf", contentType: "application/pdf", contentBytes: PDF_BYTES }]
    })
  ];
  const harness = createHarness({ messages });

  const result = await harness.service.checkForNewEmails({ user: { email: "editor@example.test" } });

  assert.equal(result.checked, 2);
  assert.equal(result.proposalsCreated, 1);
  assert.equal(result.results[1].status, "skipped");
  assert.match(result.results[1].reason, /Duplicate PDF attachment already processed/);
  assert.equal(harness.savedProposals.length, 1);
});

test("unsupported and inline attachments are skipped without blocking PDF processing", async () => {
  const message = createMessage({
    body: "Please see attached.",
    skippedAttachments: [
      { name: "logo.png", reason: "Inline attachment ignored." },
      { name: "notes.docx", reason: "Unsupported attachment type." }
    ],
    pdfAttachments: [{ id: "pdf-1", name: "Board Deck.pdf", contentType: "application/pdf", contentBytes: PDF_BYTES }]
  });
  const harness = createHarness({ messages: [message] });

  const result = await harness.service.checkForNewEmails({ user: { email: "editor@example.test" } });

  assert.equal(result.proposalsCreated, 1);
  assert.equal(result.results[0].children.filter((child) => child.status === "skipped").length, 3);
  assert.equal(harness.savedProposals.length, 1);
});

test("unmatched analysis is recorded as skipped instead of guessing an investment", async () => {
  const harness = createHarness({
    analysisFactory: () => ({
      investmentMatch: null,
      entityMatch: null,
      extractedFacts: [],
      proposedChanges: [],
      materialDevelopments: [],
      unverifiedClaims: [],
      warnings: ["No trusted investment match."],
      whatChanged: ["No verified portfolio changes identified from this document."]
    })
  });

  const result = await harness.service.checkForNewEmails({ user: { email: "editor@example.test" } });

  assert.equal(result.proposalsCreated, 0);
  assert.equal(result.skipped, 1);
  assert.match(result.results[0].reason, /No safely matched actionable analysis/);
  assert.equal(harness.savedProposals.length, 0);
  assert.equal(harness.getStored()[0].status, "skipped");
  assert.equal(harness.getStored()[0].analysisAudits.length, 1);
  assert.equal(harness.getStored()[0].analysisAudits[0].explicitMatch, false);
  assert.equal(harness.getStored()[0].analysisAudits[0].shouldCreateProposal, false);
});

test("automated explicit-match helper rejects warning-marked semantic matches", () => {
  assert.equal(
    hasAutomatedExplicitInvestmentMatch({
      investmentMatch: {
        investmentId: "healing-id",
        reason: "Model guessed from general update language."
      },
      warnings: ["Best match lacks explicit portfolio-name or alias evidence; confirm before creating a proposal."]
    }),
    false
  );
  assert.equal(
    hasAutomatedExplicitInvestmentMatch({
      investmentMatch: {
        investmentId: "finsync-id",
        reason: "Exact subject match for 'FINSYNC'. Sender domain 'finsync' supports 'FINSYNC'."
      },
      deterministicEvidence: {
        investmentId: "finsync-id",
        types: ["subject", "senderDomain"]
      },
      warnings: []
    }),
    true
  );
});

test("automated explicit-match helper does not trust model-authored reason text", () => {
  assert.equal(
    hasAutomatedExplicitInvestmentMatch({
      investmentMatch: {
        investmentId: "healing-id",
        reason: "Exact source body match for 'Healing Innovations'."
      },
      candidates: [
        {
          investmentId: "healing-id",
          reason: "Sender domain 'healing' supports 'Healing Innovations'."
        }
      ],
      warnings: []
    }),
    false
  );
});

test("automated email intake rejects semantic-only model matches even at confidence 84", async () => {
  const harness = createHarness({
    analysisFactory: () => ({
      investmentMatch: {
        investmentId: "healing-id",
        investmentName: "Healing Innovations",
        confidence: 84,
        reason: "Some content seemed related to this investment."
      },
      entityMatch: {
        entityId: "Beaman Ventures",
        entityName: "Beaman Ventures",
        confidence: 90,
        reason: "Matched from the selected investment's owning entity."
      },
      candidates: [],
      extractedFacts: [
        {
          field: "monthly recurring revenue",
          value: "$3.0 million",
          evidenceStatus: "verified"
        }
      ],
      proposedChanges: [],
      materialDevelopments: [],
      unverifiedClaims: [],
      warnings: ["Best match lacks explicit portfolio-name or alias evidence; confirm before creating a proposal."],
      whatChanged: ["Monthly recurring revenue was $3.0 million."]
    })
  });

  const result = await harness.service.checkForNewEmails({ user: { email: "editor@example.test" } });

  assert.equal(result.proposalsCreated, 0);
  assert.equal(result.skipped, 1);
  assert.match(result.results[0].reason, /requires explicit investment name/);
  assert.equal(harness.savedProposals.length, 0);
  assert.equal(harness.safetyCalls.length, 0);
  assert.equal(harness.getStored()[0].status, "skipped");
  assert.equal(harness.getStored()[0].analysisAudits.length, 1);
  assert.equal(harness.getStored()[0].analysisAudits[0].finalMatchedInvestmentId, "healing-id");
  assert.equal(harness.getStored()[0].analysisAudits[0].finalMatchedInvestmentName, "Healing Innovations");
  assert.equal(harness.getStored()[0].analysisAudits[0].explicitMatch, false);
  assert.equal(harness.getStored()[0].analysisAudits[0].shouldCreateProposal, true);
  assert.equal(harness.getStored()[0].analysisAudits[0].counts.extractedFacts, 1);
  assert.match(harness.getStored()[0].analysisAudits[0].skipReason, /requires explicit investment name/);
});

test("house-domain-only model match cannot pass automated proposal eligibility", async () => {
  const analysisService = createAiUpdateAnalysisService({
    normalizeEntityName: (value) => String(value || "").trim(),
    houseDomains: "beamanventures.com",
    callModel: async () => ({
      investmentMatch: {
        investmentId: "company-ventures",
        investmentName: "Company Ventures",
        confidence: 78,
        reason: "Sender domain 'beamanventures' supports 'Company Ventures'."
      },
      entityMatch: {},
      extractedFacts: [{ field: "company", value: "CHRP", sourceEvidence: "CHRP introduction" }],
      materialDevelopments: [],
      proposedChanges: [],
      warnings: [],
      unresolved: [],
      whatChanged: []
    })
  });

  const result = await analysisService.analyzeInvestmentUpdate({
    source: {
      sourceType: "Email",
      sender: "tyler@beamanventures.com",
      subject: "Test CHRP",
      sourceText: "CHRP introduction and investment opportunity."
    },
    investments: [{ id: "company-ventures", company: "Company Ventures", entity: "Beaman Ventures" }],
    entities: ["Beaman Ventures"]
  });

  assert.equal(result.analysis.candidates.length, 0);
  assert.equal(hasAutomatedExplicitInvestmentMatch(result.analysis), false);
  assert.equal(result.analysis.warnings.some((warning) => /lacks explicit/i.test(String(warning))), true);
});

test("valid deterministic FINSYNC match with no actionable content stores safe skipped audit", async () => {
  const message = createMessage({
    subject: "HIGHLY CONFIDENTIAL: FINSYNC August Update",
    sender: "tuckermathis@finsync.com",
    body: "FINSYNC August update discusses strategy and context but no verified actionable portfolio facts. client_secret=do-not-store raw confidential body marker."
  });
  const harness = createHarness({
    messages: [message],
    investments: [
      { id: "finsync-id", company: "FINSYNC", entity: "Beaman Ventures", investmentAliases: ["FINSYNC", "finsync"] }
    ],
    analysisFactory: () => ({
      investmentMatch: {
        investmentId: "finsync-id",
        investmentName: "FINSYNC",
        confidence: 98,
        reason: "Exact subject match for 'FINSYNC'. Sender domain 'finsync' supports 'FINSYNC'."
      },
      deterministicEvidence: {
        investmentId: "finsync-id",
        types: ["subject", "senderDomain"]
      },
      entityMatch: {
        entityId: "Beaman Ventures",
        entityName: "Beaman Ventures",
        confidence: 100,
        reason: "Matched from the selected investment's owning entity."
      },
      candidates: [
        {
          investmentId: "finsync-id",
          investmentName: "FINSYNC",
          confidence: 98,
          reason: "Exact subject match for 'FINSYNC'. Sender domain 'finsync' supports 'FINSYNC'."
        }
      ],
      extractedFacts: [],
      proposedChanges: [],
      materialDevelopments: [],
      unverifiedClaims: [],
      warnings: [],
      unresolved: [],
      whatChanged: ["No verified portfolio changes identified from this document."]
    })
  });

  const result = await harness.service.checkForNewEmails({ user: { email: "editor@example.test" } });
  const stored = harness.getStored()[0];
  const audit = stored.analysisAudits[0];
  const serializedState = JSON.stringify(stored);

  assert.equal(result.proposalsCreated, 0);
  assert.equal(result.skipped, 1);
  assert.equal(harness.savedProposals.length, 0);
  assert.equal(stored.status, "skipped");
  assert.equal(audit.finalMatchedInvestmentId, "finsync-id");
  assert.equal(audit.finalMatchedInvestmentName, "FINSYNC");
  assert.equal(audit.matchedEntity, "Beaman Ventures");
  assert.equal(audit.matchConfidence, 98);
  assert.equal(audit.explicitMatch, true);
  assert.deepEqual(audit.deterministicEvidenceTypes, ["subject", "senderDomain"]);
  assert.equal(audit.shouldCreateProposal, false);
  assert.equal(audit.counts.extractedFacts, 0);
  assert.equal(audit.counts.materialDevelopments, 0);
  assert.equal(audit.counts.proposedChanges, 0);
  assert.equal(audit.counts.unverifiedClaims, 0);
  assert.match(audit.skipReason, /No safely matched actionable analysis/);
  assert.doesNotMatch(serializedState, /do-not-store|raw confidential body marker|client_secret/);
});

test("automated FINSYNC email cannot create a Healing proposal when FINSYNC is absent", async () => {
  const message = createMessage({
    subject: "HIGHLY CONFIDENTIAL: FINSYNC August Update",
    sender: "tuckermathis@finsync.com",
    body: "FINSYNC August investor update includes customer growth, monthly recurring revenue, sales pipeline, capital planning, product milestones, and operational risks for the quarter."
  });
  const harness = createHarness({
    messages: [message],
    investments: [
      { id: "healing-id", company: "Healing Innovations", owningEntity: "Beaman Ventures" }
    ],
    analysisFactory: () => ({
      investmentMatch: {
        investmentId: "healing-id",
        investmentName: "Healing Innovations",
        confidence: 84,
        reason: "Some content pertains to potential strategic discussions."
      },
      entityMatch: { entityId: "Beaman Ventures", entityName: "Beaman Ventures", confidence: 90 },
      candidates: [],
      extractedFacts: [
        { field: "monthly recurring revenue", value: "$3.0 million", evidenceStatus: "verified" }
      ],
      proposedChanges: [],
      materialDevelopments: [],
      unverifiedClaims: [],
      warnings: ["Best match lacks explicit portfolio-name or alias evidence; confirm before creating a proposal."],
      whatChanged: ["No verified portfolio changes identified from this document."]
    })
  });

  const result = await harness.service.checkForNewEmails({ user: { email: "editor@example.test" } });

  assert.equal(result.proposalsCreated, 0);
  assert.equal(harness.savedProposals.some((proposal) => proposal.investmentId === "healing-id"), false);
  assert.equal(result.results[0].status, "skipped");
});

test("automated email intake accepts a valid deterministic explicit match", async () => {
  const harness = createHarness({
    analysisFactory: () => ({
      investmentMatch: {
        investmentId: "finsync-id",
        investmentName: "FINSYNC",
        confidence: 98,
        reason: "Exact subject match for 'FINSYNC'. Sender domain 'finsync' supports 'FINSYNC'."
      },
      deterministicEvidence: {
        investmentId: "finsync-id",
        types: ["subject", "senderDomain"]
      },
      entityMatch: {
        entityId: "Beaman Ventures",
        entityName: "Beaman Ventures",
        confidence: 100,
        reason: "Matched from the selected investment's owning entity."
      },
      candidates: [
        {
          investmentId: "finsync-id",
          investmentName: "FINSYNC",
          confidence: 98,
          reason: "Exact subject match for 'FINSYNC'. Sender domain 'finsync' supports 'FINSYNC'."
        }
      ],
      extractedFacts: [
        { field: "monthly recurring revenue", value: "$3.0 million", evidenceStatus: "verified" }
      ],
      proposedChanges: [],
      materialDevelopments: [],
      unverifiedClaims: [],
      warnings: [],
      whatChanged: ["Monthly recurring revenue was $3.0 million."]
    }),
    investments: [
      { id: "finsync-id", company: "FINSYNC", owningEntity: "Beaman Ventures", investmentAliases: ["finsync"] },
      { id: "healing-id", company: "Healing Innovations", owningEntity: "Beaman Ventures" }
    ]
  });

  const result = await harness.service.checkForNewEmails({ user: { email: "editor@example.test" } });

  assert.equal(result.proposalsCreated, 1);
  assert.equal(harness.savedProposals[0].investmentId, "finsync-id");
  assert.equal(harness.savedProposals[0].status, "pending");
  assert.deepEqual(harness.getStored()[0].analysisAudits, []);
});

test("automated email intake creates a pending proposal when email-body facts survive strict verification", async () => {
  const sourceText =
    "FINSYNC update. Based on results through July, we are now at approximately 62,824 customers and $3.0 million in monthly recurring revenue.";
  const message = createMessage({
    subject: "HIGHLY CONFIDENTIAL: FINSYNC August Update",
    sender: "tuckermathis@finsync.com",
    body: sourceText
  });
  const investmentRows = [
    { id: "finsync-id", company: "FINSYNC", entity: "Beaman Ventures", investmentAliases: ["FINSYNC", "finsync"] }
  ];
  const analysisService = createAiUpdateAnalysisService({
    normalizeEntityName: (value) => String(value || "").trim(),
    getNow: () => new Date("2026-09-02T12:00:00.000Z"),
    callModel: async () => ({
      investmentMatch: {
        investmentId: "finsync-id",
        investmentName: "FINSYNC",
        confidence: 84,
        reason: "Model did not extract the operating metrics."
      },
      entityMatch: {},
      extractedFacts: [
        {
          field: "revenue",
          value: "21.7M",
          sourceEvidence: "revenue was 21.7M",
          confidence: 97
        }
      ],
      materialDevelopments: [],
      proposedChanges: [
        {
          field: "revenue",
          proposedValue: "21.7M",
          sourceEvidence: "revenue was 21.7M",
          confidence: 97
        }
      ],
      warnings: [],
      unresolved: [],
      whatChanged: ["Revenue increased to 21.7M."]
    })
  });
  const beforeRows = JSON.stringify(investmentRows);
  const harness = createHarness({
    messages: [message],
    investments: investmentRows,
    analysisFactory: null
  });
  harness.service.checkForNewEmails = createAiEmailIntakeService({
    graphMailService: {
      isConfigured: () => true,
      fetchIntakeMessages: async () => ({
        mailbox: "updates@example.test",
        folder: { id: "folder-1", displayName: "AI Investment Updates" },
        messages: [message]
      })
    },
    stateService: createMemoryStateService().service,
    analyzeInvestmentUpdate: analysisService.analyzeInvestmentUpdate,
    extractPdfTextFromUpload: async () => {
      throw new Error("PDF should not be read");
    },
    finalizeAnalysisForResponse: (analysis) => analysis,
    enforceProposalSafetyInvariant: (proposal) => proposal,
    saveAiUpdateProposal: (proposal) => {
      const saved = { ...proposal, id: `proposal-${harness.savedProposals.length + 1}` };
      harness.savedProposals.push(saved);
      return saved;
    },
    readInvestments: () => investmentRows,
    filterInvestmentsForUser: (investments) => investments,
    entities: ["Beaman Ventures"],
    canViewEntity: () => true
  }).checkForNewEmails;

  const result = await harness.service.checkForNewEmails({ user: { email: "editor@example.test" } });
  const proposal = harness.savedProposals[0];

  assert.equal(result.proposalsCreated, 1);
  assert.equal(proposal.investmentId, "finsync-id");
  assert.equal(proposal.status, "pending");
  assert.equal(proposal.extractedData.facts.some((fact) => fact.field === "customerCount" && fact.evidenceStatus === "verified"), true);
  assert.equal(proposal.extractedData.facts.some((fact) => fact.field === "monthly recurring revenue" && fact.evidenceStatus === "verified"), true);
  assert.equal(proposal.extractedData.facts.some((fact) => fact.field === "revenue" && String(fact.value).includes("21.7")), false);
  assert.equal(proposal.proposedChanges.length, 0);
  assert.doesNotMatch(JSON.stringify(proposal), /Revenue increased to 21\.7M/);
  assert.equal(JSON.stringify(investmentRows), beforeRows);
});

test("one corrupt PDF does not stop later messages from processing", async () => {
  const messages = [
    createMessage({
      id: "bad-message",
      internetMessageId: "<bad-message@example.test>",
      body: "Please see attached.",
      pdfAttachments: [{ id: "bad-pdf", name: "Broken.pdf", contentType: "application/pdf", contentBytes: PDF_BYTES }]
    }),
    createMessage({
      id: "good-message",
      internetMessageId: "<good-message@example.test>"
    })
  ];
  const harness = createHarness({
    messages,
    extractPdfTextFromUpload: async ({ filename, fileData }) => {
      if (filename === "Broken.pdf") {
        throw new Error("PDF parse failed.");
      }
      return {
        filename,
        buffer: Buffer.from(fileData, "base64"),
        pageCount: 1,
        pages: [],
        combinedText: "Parsed PDF text",
        diagnostics: {}
      };
    }
  });

  const result = await harness.service.checkForNewEmails({ user: { email: "editor@example.test" } });

  assert.equal(result.checked, 2);
  assert.equal(result.failed, 1);
  assert.equal(result.processed, 1);
  assert.equal(result.proposalsCreated, 1);
});

test("sender allowlists skip unauthorized mail before analysis", async () => {
  const harness = createHarness({ allowedDomains: "approved.example" });

  const result = await harness.service.checkForNewEmails({ user: { email: "editor@example.test" } });

  assert.equal(result.checked, 1);
  assert.equal(result.skipped, 1);
  assert.equal(result.results[0].reason, "Sender is outside the configured intake allowlist.");
  assert.equal(harness.analyzeCalls.length, 0);
  assert.equal(harness.savedProposals.length, 0);
});

test("Graph run failures return a safe failed summary without tokens or secrets", async () => {
  const { service } = createHarness();
  service.checkForNewEmails = createAiEmailIntakeService({
    graphMailService: {
      isConfigured: () => true,
      fetchIntakeMessages: async () => {
        throw new Error("Microsoft Graph request failed.");
      }
    },
    stateService: createMemoryStateService().service,
    analyzeInvestmentUpdate: async () => {
      throw new Error("should not analyze");
    },
    extractPdfTextFromUpload: async () => {
      throw new Error("should not extract");
    },
    finalizeAnalysisForResponse: (analysis) => analysis,
    enforceProposalSafetyInvariant: (proposal) => proposal,
    saveAiUpdateProposal: (proposal) => proposal,
    readInvestments: () => [],
    filterInvestmentsForUser: (investments) => investments
  }).checkForNewEmails;

  const result = await service.checkForNewEmails({ user: { email: "editor@example.test" } });

  assert.equal(result.checked, 0);
  assert.equal(result.failed, 1);
  assert.equal(result.proposalsCreated, 0);
  assert.match(result.error, /Microsoft Graph request failed/);
  assert.doesNotMatch(JSON.stringify(result), /access_token|client_secret|token-value/);
});

test("email body normalization strips obvious markup, signatures, and quoted thread", () => {
  const text = normalizeEmailBody({
    bodyContentType: "html",
    body: "<div>Healing update &amp; customer pipeline.</div><div>-- </div><div>Signature</div><div>From: old thread</div>"
  });

  assert.equal(text, "Healing update & customer pipeline.");
  assert.equal(isMeaningfulBody("Please see attached."), false);
  assert.equal(
    isMeaningfulBody("Healing Innovations quarterly update includes revenue, customer pipeline, units sold, cash runway, investor update, milestones, and board-level financial risks."),
    true
  );
});

test("attachment hash is stable for dedupe", () => {
  assert.equal(attachmentHash(PDF_BYTES), attachmentHash(PDF_BYTES));
  assert.notEqual(attachmentHash(PDF_BYTES), attachmentHash(Buffer.from("different").toString("base64")));
});

test("potential new deal creates one pending proposal with preserved and unresolved attachments", async () => {
  const message = createMessage({
    subject: "NewCo seed opportunity",
    bodyContentType: "text",
    body: "NewCo is raising $5 million. Beaman Ventures proposed check size is $250,000.",
    attachments: [{
      id: "deck", name: "NewCo Deck.pdf", contentType: "application/pdf", size: 18,
      contentBytes: PDF_BYTES, isPdf: true, preservationStatus: "preserved"
    }],
    pdfAttachments: [{ id: "deck", name: "NewCo Deck.pdf", contentType: "application/pdf", contentBytes: PDF_BYTES }],
    unresolvedAttachments: [{ id: "model", name: "Model.pages", reason: "Unsupported attachment type was not parsed or downloaded." }]
  });
  const harness = createHarness({
    messages: [message],
    investments: [],
    analyzePotentialNewDeal: async () => ({
      route: "new-deal",
      analysis: {
        classificationReason: "Inbound financing opportunity.",
        opportunityFingerprint: "fingerprint-newco",
        matchResult: { status: "no-match", confidence: 0, reason: "No deterministic match.", candidates: [] },
        dealData: {
          companyName: { value: "NewCo", evidenceStatus: "verified", authoritativeValue: "NewCo" },
          dealSummary: { value: "Seed opportunity", evidenceStatus: "verified", authoritativeValue: "Seed opportunity" },
          proposedCheckSize: { value: "$250,000", evidenceStatus: "verified", authoritativeValue: "$250,000" }
        }
      }
    })
  });
  const first = await harness.service.checkForNewEmails({ user: { email: "editor@example.test" } });
  const second = await harness.service.checkForNewEmails({ user: { email: "editor@example.test" } });
  assert.equal(first.proposalsCreated, 1);
  assert.equal(second.proposalsCreated, 0);
  assert.equal(harness.savedProposals.length, 1);
  assert.equal(harness.savedProposals[0].proposalType, "new-deal");
  assert.equal(harness.savedProposals[0].documents.length, 2);
  assert.equal(harness.savedProposals[0].documents[0].storedName, "NewCo Deck.pdf");
  assert.equal(harness.savedProposals[0].documents[1].preservationStatus, "unresolved");
});

test("one sanitized source email creates three idempotent proposals with partitioned attachments", async () => {
  const names = ["BEP Core Fund VIII", "Project Pure", "Project Care"];
  const message = createMessage({
    subject: "BEP background and teasers",
    bodyContentType: "text",
    body: "BEP Core Fund VIII fundraising is closed. Project Pure is oversubscribed. Project Care is under LOI.",
    attachments: names.map((name, index) => ({ id: `attachment-${index + 1}`, name: `${name}.pdf`, contentType: "application/pdf", size: 18, contentBytes: Buffer.from(name).toString("base64"), isPdf: true })),
    pdfAttachments: names.map((name, index) => ({ id: `attachment-${index + 1}`, name: `${name}.pdf`, contentType: "application/pdf", contentBytes: Buffer.from(name).toString("base64") }))
  });
  const harness = createHarness({
    messages: [message],
    investments: [],
    extractPdfTextFromUpload: async ({ filename, fileData }) => ({ filename, buffer: Buffer.from(fileData, "base64"), pageCount: 1, pages: [], combinedText: filename, diagnostics: {} }),
    analyzePotentialNewDeals: async ({ source }) => source.attachments.map((attachment, index) => ({
      opportunity: { name: names[index], opportunityId: `opportunity-${index + 1}`, attachmentIds: [attachment.id] },
      source: { ...source, filename: attachment.name, sourceText: attachment.text },
      result: {
        route: "new-deal",
        analysis: {
          classificationReason: "Separately reviewable opportunity.",
          opportunityFingerprint: `fingerprint-${index + 1}`,
          matchResult: { status: "no-match", confidence: 0, reason: "No deterministic match.", candidates: [] },
          dealData: {
            companyName: { value: names[index], evidenceStatus: "verified", authoritativeValue: names[index] },
            dealSummary: { value: names[index], evidenceStatus: "verified", authoritativeValue: names[index] },
            proposedCheckSize: { value: "", evidenceStatus: "unresolved", authoritativeValue: "" }
          }
        }
      }
    }))
  });

  const first = await harness.service.checkForNewEmails({ user: { email: "editor@example.test" } });
  const second = await harness.service.checkForNewEmails({ user: { email: "editor@example.test" } });
  assert.equal(first.proposalsCreated, 3);
  assert.equal(second.proposalsCreated, 0);
  assert.deepEqual(harness.savedProposals.map((proposal) => proposal.opportunityName), names);
  assert.deepEqual(harness.savedProposals.map((proposal) => proposal.documents.map((document) => document.name)), names.map((name) => [`${name}.pdf`]));
  assert.equal(new Set(harness.savedProposals.map((proposal) => `${proposal.sourceMessageKey}:${proposal.opportunityId}`)).size, 3);
  assert.deepEqual(harness.getStored()[0].proposalIds, ["proposal-1", "proposal-2", "proposal-3"]);
});

test("explicit reanalysis fetches the exact preserved source and only stages pending proposals", async () => {
  const names = ["BEP Core Fund VIII", "Project Pure", "Project Care"];
  const exactMessage = createMessage({
    id: "bep-graph-id",
    internetMessageId: "<bep-source@example.test>",
    subject: "BEP background and teasers",
    bodyContentType: "text",
    body: "Core closed. Pure oversubscribed. Care under LOI.",
    attachments: names.map((name, index) => ({ id: `attachment-${index}`, name: `${name}.pdf`, contentType: "application/pdf", size: 10, contentBytes: Buffer.from(name).toString("base64"), isPdf: true })),
    pdfAttachments: names.map((name, index) => ({ id: `attachment-${index}`, name: `${name}.pdf`, contentType: "application/pdf", contentBytes: Buffer.from(name).toString("base64") }))
  });
  const harness = createHarness({
    messages: [],
    exactMessage,
    investments: [],
    extractPdfTextFromUpload: async ({ filename, fileData }) => ({ filename, buffer: Buffer.from(fileData, "base64"), pageCount: 1, pages: [], combinedText: filename, diagnostics: {} }),
    analyzePotentialNewDeals: async ({ source }) => source.attachments.map((attachment, index) => ({
      opportunity: { name: names[index], opportunityId: `opp-${index}`, attachmentIds: [attachment.id] },
      source: { ...source, sourceText: attachment.text },
      result: { route: "new-deal", analysis: { classificationReason: "Opportunity", opportunityFingerprint: `fp-${index}`, matchResult: { status: "no-match", confidence: 0, reason: "No match", candidates: [] }, dealData: { companyName: { value: names[index] }, dealSummary: { value: names[index] }, proposedCheckSize: { value: "", evidenceStatus: "unresolved", authoritativeValue: "" } } } }
    }))
  });
  const result = await harness.service.reanalyzeMessage({
    user: { email: "master@example.test" },
    graphMessageId: "bep-graph-id",
    expectedInternetMessageId: "<bep-source@example.test>"
  });
  assert.equal(result.proposalsCreated, 3);
  assert.deepEqual(harness.savedProposals.map((proposal) => proposal.status), ["pending", "pending", "pending"]);
  assert.deepEqual(harness.getStored()[0].proposalIds, ["proposal-1", "proposal-2", "proposal-3"]);
});

test("complete BEP reanalysis reconciles five legacy pending siblings to three canonical active proposals", async () => {
  const exactMessage = createMessage({
    id: "bep-graph-id",
    internetMessageId: "<bep-source@example.test>",
    subject: "BEP background and teasers",
    bodyContentType: "text",
    body: bepFixture.emailBody,
    attachments: bepFixture.attachments.map((attachment) => ({
      id: attachment.id,
      name: attachment.name,
      contentType: "application/pdf",
      size: attachment.text.length,
      contentBytes: Buffer.from(attachment.text).toString("base64"),
      isPdf: true
    })),
    pdfAttachments: bepFixture.attachments.map((attachment) => ({
      id: attachment.id,
      name: attachment.name,
      contentType: "application/pdf",
      contentBytes: Buffer.from(attachment.text).toString("base64")
    }))
  });
  const variedDecomposition = {
    opportunities: bepFixture.decomposition.opportunities.map((opportunity) => ({
      ...opportunity,
      name: opportunity.name === "Project Pure"
        ? "Project Pure Co-Investment"
        : opportunity.name === "Project Care" ? "Opportunity: Project Care Co-Invest" : opportunity.name,
      currentStatus: opportunity.name === "BEP Core Fund VIII"
        ? { value: "Active", sourceEvidence: "BEP Core Fund VIII fundraising memo is now close." }
        : opportunity.currentStatus,
      emailEvidence: opportunity.name === "BEP Core Fund VIII" ? [] : opportunity.emailEvidence
    }))
  };
  const modelResponses = [
    bepFixture.decomposition, ...bepFixture.analyses,
    variedDecomposition, ...bepFixture.analyses
  ];
  const analysisService = createNewDealAnalysisService({ callModel: async () => modelResponses.shift() });
  let stored = structuredClone(bepLegacyFiveProposals);
  let generatedId = 0;
  const investments = [];
  const proposalService = createAiUpdateProposalService({
    AI_UPDATE_PROPOSALS_FILE: "proposals.json",
    readJsonFile: () => stored,
    writeJsonFile: (_file, value) => { stored = value; },
    writeMetadata: () => {},
    normalizeAiUpdateProposal: (value) => ({
      id: value.id || `generated-proposal-${++generatedId}`,
      status: value.status || "pending",
      proposalType: value.proposalType || "new-deal",
      documents: value.documents || [],
      opportunityIdentityKeys: value.opportunityIdentityKeys || [],
      ...value
    }),
    createBackupSnapshot: () => {},
    applyApprovedAiUpdateProposal: () => ({ applied: false })
  });
  const harness = createHarness({
    messages: [],
    exactMessage,
    investments,
    analyzePotentialNewDeals: analysisService.analyzePotentialNewDeals,
    extractPdfTextFromUpload: async ({ filename, fileData }) => ({
      filename,
      buffer: Buffer.from(fileData, "base64"),
      pageCount: 1,
      pages: [],
      combinedText: bepFixture.attachments.find((attachment) => attachment.name === filename).text,
      diagnostics: {}
    }),
    saveAiUpdateProposal: proposalService.saveAiUpdateProposal,
    sourceProposalSnapshot: proposalService.sourceProposalSnapshot,
    reconcilePendingSourceOpportunities: proposalService.reconcilePendingSourceOpportunities
  });
  const first = await harness.service.reanalyzeMessage({
    user: { email: "master@example.test" },
    graphMessageId: "bep-graph-id",
    expectedInternetMessageId: "<bep-source@example.test>"
  });
  const second = await harness.service.reanalyzeMessage({
    user: { email: "master@example.test" },
    graphMessageId: "bep-graph-id",
    expectedInternetMessageId: "<bep-source@example.test>"
  });
  assert.equal(stored.length, 5);
  assert.deepEqual(second.proposalIds, first.proposalIds);
  assert.deepEqual(first.proposalIds.sort(), ["bep-core-canonical", "care-canonical", "pure-canonical"]);
  assert.deepEqual(first.supersededProposalIds.sort(), ["care-legacy-alias", "pure-legacy-alias"]);
  assert.deepEqual(second.supersededProposalIds, []);
  const pending = stored.filter((proposal) => proposal.status === "pending");
  const superseded = stored.filter((proposal) => proposal.status === "superseded");
  assert.deepEqual(pending.map((proposal) => proposal.opportunityName).sort(), [
    "BEP Core Fund VIII", "Project Care", "Project Pure"
  ]);
  assert.equal(superseded.length, 2);
  assert.deepEqual(
    superseded.map((proposal) => [proposal.id, proposal.supersededByProposalIds[0]]).sort(),
    [["care-legacy-alias", "care-canonical"], ["pure-legacy-alias", "pure-canonical"]]
  );
  const coreProposal = pending.find((proposal) => proposal.opportunityName === "BEP Core Fund VIII");
  const core = coreProposal.dealData;
  assert.equal(core.stage.authoritativeValue, "Fundraising closed");
  assert.equal(core.stage.sourceEvidence, "The fundraising memo is now close.");
  assert.equal(core.stage.supersededEvidence[0].value, "Active fundraising");
  assert.equal(core.targetFundSize.authoritativeValue, "$350MM");
  assert.equal(core.amountCommitted.authoritativeValue, "$94.5MM");
  assert.equal(core.minimumLpCommitment.authoritativeValue, "$5MM");
  assert.equal(core.historicalTargetDifference.authoritativeValue, "$255.5MM");
  assert.equal(core.historicalTargetDifference.currentAvailability, false);
  assert.deepEqual(core.historicalTargetDifference.derivedFrom.map((item) => item.field), ["targetFundSize", "amountCommitted"]);
  assert.equal(core.amountRemaining.value, "");
  assert.equal(core.amountRemaining.evidenceStatus, "unresolved");
  assert.equal(core.proposedCheckSize.authoritativeValue, "");
  assert.equal(core.proposedCheckSize.evidenceStatus, "unresolved");
  assert.deepEqual(core.deadlines, []);
  assert.deepEqual(core.financingTerms.map((claim) => claim.semanticLabel), [
    "Management fee", "Performance fee/carry", "Fund term"
  ]);
  assert.deepEqual(core.financingTerms.map((claim) => claim.authoritativeValue), [
    "1.75%", "17.5% carried interest above 8% preferred return", "10 years with two one-year extensions"
  ]);
  assert.equal(core.financingTerms.every((claim) => claim.sourceEvidence && claim.evidenceStatus === "verified"), true);
  const visibleCounts = normalizeAiUpdateProposalCounts({ proposals: stored });
  assert.deepEqual(visibleCounts, { pending: 3, approved: 0, rejected: 0, superseded: 2 });
  assert.equal(formatDealClaimValue(coreProposal, "historicalTargetDifference"), "$255.5MM");
  assert.equal(getDisplayDealClaimItems(coreProposal, "historicalTargetDifference")[0].currentAvailability, false);
  assert.equal(formatDealClaimValue(coreProposal, "amountRemaining"), "");
  assert.equal(formatDealClaimValue(coreProposal, "proposedCheckSize"), "");
  assert.match(formatDealClaimValue(coreProposal, "dealSummary"), /target fund size of \$350MM/i);
  assert.match(formatDealClaimValue(coreProposal, "dealSummary"), /fundraising is closed/i);
  assert.equal(formatAiUpdateProposalSummary(coreProposal), formatDealClaimValue(coreProposal, "dealSummary"));
  assert.doesNotMatch([
    "dealSummary", "whatCompanyDoes", "businessModel", "tractionRevenue", "keyInvestmentPoints"
  ].map((field) => formatDealClaimValue(coreProposal, field)).join("\n"), /currently raising|seeking to raise/i);
  assert.deepEqual(getDisplayDealClaimItems(coreProposal, "nextSteps"), []);
  assert.deepEqual(pending.map((proposal) => proposal.documents.map((document) => document.graphAttachmentId)).sort(), [
    ["core-fund-viii"], ["project-care"], ["project-pure"]
  ]);
  assert.equal(harness.uploads.length, 3);
  assert.equal(new Set(stored.map((proposal) => proposal.id)).size, 5);
  assert.deepEqual(harness.getStored()[0].proposalIds.sort(), ["bep-core-canonical", "care-canonical", "pure-canonical"]);
  assert.deepEqual(investments, []);
});

test("read-only intake preview reports allowlist and terminal state without invoking mutation or analysis", async () => {
  const calls = { claim: 0, upsert: 0, analysis: 0, proposals: 0 };
  const service = createAiEmailIntakeService({
    graphMailService: {
      isConfigured: () => true,
      previewIntakeMessages: async () => ({
        mailbox: "updates@example.test",
        folder: { id: "folder-1", displayName: "AI Investment Updates" },
        maxMessagesPerRun: 10,
        limits: { maxMessageBytes: 30, maxRunBytes: 50 },
        messages: [
          {
            graphMessageId: "attainable-graph-id", internetMessageId: "<attainable@example.test>",
            sender: "cleseberg@thesignatry.com", subject: "Attainable Living LLC", receivedDateTime: "2026-09-18T16:00:00Z",
            attachmentCount: 1, attachments: [], projectedMessageAttachmentBytes: 20, projectedRunAttachmentBytesAfterMessage: 20
          },
          {
            graphMessageId: "processed-graph-id", internetMessageId: "<processed@example.test>",
            sender: "founder@portfolio.example", subject: "Processed", receivedDateTime: "2026-09-18T15:00:00Z",
            attachmentCount: 0, attachments: [], projectedMessageAttachmentBytes: 0, projectedRunAttachmentBytesAfterMessage: 20
          }
        ]
      })
    },
    stateService: {
      findByMessage: (message) => message.graphMessageId === "processed-graph-id"
        ? { status: "processed", processedAt: "2026-09-18T15:05:00Z", proposalIds: ["proposal-1"], error: "" }
        : null,
      claimMessage: () => { calls.claim += 1; throw new Error("preview must not claim"); },
      upsertEntry: () => { calls.upsert += 1; throw new Error("preview must not write"); }
    },
    analyzeInvestmentUpdate: async () => { calls.analysis += 1; throw new Error("preview must not analyze"); },
    analyzePotentialNewDeal: async () => { calls.analysis += 1; throw new Error("preview must not analyze"); },
    saveAiUpdateProposal: () => { calls.proposals += 1; throw new Error("preview must not save proposals"); },
    readInvestments: () => [],
    filterInvestmentsForUser: (investments) => investments,
    allowedDomains: "thesignatry.com,portfolio.example"
  });

  const result = await service.previewEmails({ now: new Date("2026-09-18T16:10:00Z") });

  assert.equal(result.messages[0].eligibilityStatus, "eligible");
  assert.equal(result.messages[0].allowlist.domainMatch, true);
  assert.equal(result.messages[0].state.found, false);
  assert.equal(result.messages[1].eligibilityStatus, "already-processed");
  assert.deepEqual(calls, { claim: 0, upsert: 0, analysis: 0, proposals: 0 });
});

test("intake preview explains blocked, reserved, terminal, and retryable state", () => {
  const message = { graphMessageId: "graph-1", internetMessageId: "<message@example.test>" };
  const allowed = { allowed: true };
  const now = new Date("2026-09-18T16:10:00Z");

  assert.equal(previewMessageEligibility({}, null, allowed, now).status, "blocked");
  assert.equal(previewMessageEligibility(message, { status: "processed" }, allowed, now).status, "already-processed");
  assert.equal(previewMessageEligibility(message, { status: "skipped" }, allowed, now).status, "already-skipped");
  assert.equal(previewMessageEligibility(message, {
    status: "reserved", reservedAt: "2026-09-18T16:05:00Z"
  }, allowed, now).status, "reserved");
  assert.equal(previewMessageEligibility(message, {
    status: "reserved", reservedAt: "2026-09-18T15:30:00Z"
  }, allowed, now).status, "eligible");
  assert.equal(previewMessageEligibility(message, { status: "failed" }, allowed, now).status, "eligible");
  assert.equal(previewMessageEligibility(message, null, { allowed: false }, now).status, "skipped");
});
