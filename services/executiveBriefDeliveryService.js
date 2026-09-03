const crypto = require("crypto");
const { buildExecutiveBriefTemplate, validateRestrictedHtml } = require("./executiveBriefTemplate");

const INPUT_KEYS = new Set(["briefType", "subject", "html", "text"]);

class ExecutiveBriefError extends Error {
  constructor(code, message, statusCode = 400) {
    super(message);
    this.name = "ExecutiveBriefError";
    this.code = code;
    this.statusCode = statusCode;
  }
}

function centralDateParts(date) {
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone: "America/Chicago",
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).formatToParts(date).reduce((value, part) => {
    value[part.type] = part.value;
    return value;
  }, {});
  return {
    key: `${parts.year}-${parts.month}-${parts.day}`,
    display: new Intl.DateTimeFormat("en-US", {
      timeZone: "America/Chicago",
      year: "numeric",
      month: "long",
      day: "numeric"
    }).format(date)
  };
}

function expectedSubject(briefType, displayDate) {
  return `BEAMAN VENTURES | ${briefType.toUpperCase()} EXECUTIVE BRIEF | ${displayDate}`;
}

function validateExecutiveBriefInput(input, now) {
  if (!input || typeof input !== "object" || Array.isArray(input)) {
    throw new ExecutiveBriefError("INVALID_INPUT", "Executive brief input must be an object.");
  }
  const unknownKeys = Object.keys(input).filter((key) => !INPUT_KEYS.has(key));
  if (unknownKeys.length > 0) {
    throw new ExecutiveBriefError("INVALID_INPUT", "Executive brief input contains unsupported fields.");
  }
  if (!new Set(["morning", "evening"]).has(input.briefType)) {
    throw new ExecutiveBriefError("INVALID_BRIEF_TYPE", "Brief type must be morning or evening.");
  }
  const { key, display } = centralDateParts(now);
  const requiredSubject = expectedSubject(input.briefType, display);
  if (typeof input.subject !== "string" || input.subject !== requiredSubject) {
    throw new ExecutiveBriefError("INVALID_SUBJECT", `Subject must be exactly: ${requiredSubject}`);
  }
  if (typeof input.text !== "string" || !input.text.trim() || input.text.length > 100000 || input.text.includes("\0")) {
    throw new ExecutiveBriefError("INVALID_TEXT", "Brief text must contain between 1 and 100000 characters.");
  }
  if (typeof input.html !== "string") {
    throw new ExecutiveBriefError("INVALID_HTML", "Brief HTML is required.");
  }
  try {
    validateRestrictedHtml(input.html);
  } catch (error) {
    throw new ExecutiveBriefError("INVALID_HTML", error.message);
  }
  return { briefingDate: key, displayDate: display, requiredSubject };
}

function createExecutiveBriefDeliveryService({
  enabled = false,
  transport,
  stateService,
  now = () => new Date(),
  makeId = () => crypto.randomUUID(),
  audit = () => {}
}) {
  async function sendExecutiveBrief(input) {
    if (!enabled) {
      throw new ExecutiveBriefError("DELIVERY_DISABLED", "Executive brief delivery is disabled.", 503);
    }
    const currentTime = now();
    const validation = validateExecutiveBriefInput(input, currentTime);
    const idempotencyKey = `executive-brief/${validation.briefingDate}/${input.briefType}`;
    const contentSha256 = crypto.createHash("sha256")
      .update(`${input.subject}\n${input.text}\n${input.html}`)
      .digest("hex");
    const deliveryId = makeId();
    const attemptedAt = currentTime.toISOString();
    const reservation = stateService.reserve({
      deliveryId,
      idempotencyKey,
      briefType: input.briefType,
      briefingDate: validation.briefingDate,
      subject: validation.requiredSubject,
      contentSha256,
      attemptedAt
    });
    if (!reservation.reserved) {
      audit({ event: "executive_brief_blocked", briefType: input.briefType, briefingDate: validation.briefingDate, reason: reservation.reason });
      throw new ExecutiveBriefError("DUPLICATE_DELIVERY", "An executive brief delivery is already reserved for this type and Central date.", 409);
    }

    const rendered = buildExecutiveBriefTemplate({
      briefType: input.briefType,
      displayDate: validation.displayDate,
      html: input.html,
      text: input.text
    });
    audit({ event: "executive_brief_attempted", deliveryId, briefType: input.briefType, briefingDate: validation.briefingDate });

    try {
      const result = await transport.send({
        subject: validation.requiredSubject,
        html: rendered.html,
        text: rendered.text,
        idempotencyKey
      });
      const sentAt = now().toISOString();
      stateService.markSent(idempotencyKey, sentAt, result.providerMessageId);
      audit({ event: "executive_brief_sent", deliveryId, briefType: input.briefType, briefingDate: validation.briefingDate });
      return { status: "sent", deliveryId, briefType: input.briefType, briefingDate: validation.briefingDate };
    } catch (error) {
      stateService.markFailedClosed(idempotencyKey, now().toISOString());
      const safeErrorCodes = new Set([
        "DELIVERY_NOT_CONFIGURED",
        "DELIVERY_PROVIDER_FAILED",
        "DELIVERY_PROVIDER_REJECTED"
      ]);
      const errorCode = safeErrorCodes.has(error && error.code) ? error.code : "DELIVERY_FAILED";
      audit({ event: "executive_brief_failed_closed", deliveryId, briefType: input.briefType, briefingDate: validation.briefingDate, errorCode });
      throw new ExecutiveBriefError("DELIVERY_FAILED", "Executive brief delivery failed closed.", 502);
    }
  }

  return { sendExecutiveBrief };
}

module.exports = {
  ExecutiveBriefError,
  centralDateParts,
  createExecutiveBriefDeliveryService,
  expectedSubject,
  validateExecutiveBriefInput
};
