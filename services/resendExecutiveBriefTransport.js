const FIXED_EXECUTIVE_BRIEF_RECIPIENT = "tyler@beamanventures.com";

function createResendExecutiveBriefTransport({ apiKey, fromEmail, fetchImpl = global.fetch }) {
  async function send({ subject, html, text, idempotencyKey }) {
    if (!apiKey || !String(fromEmail || "").trim() || typeof fetchImpl !== "function") {
      const error = new Error("Executive brief delivery is not configured.");
      error.code = "DELIVERY_NOT_CONFIGURED";
      throw error;
    }

    let response;
    try {
      response = await fetchImpl("https://api.resend.com/emails", {
        method: "POST",
        headers: {
          Authorization: `Bearer ${apiKey}`,
          "Content-Type": "application/json",
          "Idempotency-Key": idempotencyKey
        },
        body: JSON.stringify({
          from: String(fromEmail).trim(),
          to: [FIXED_EXECUTIVE_BRIEF_RECIPIENT],
          subject,
          html,
          text
        })
      });
    } catch (cause) {
      const error = new Error("Executive brief provider request failed.");
      error.code = "DELIVERY_PROVIDER_FAILED";
      throw error;
    }

    if (!response.ok) {
      const error = new Error("Executive brief provider rejected the request.");
      error.code = "DELIVERY_PROVIDER_REJECTED";
      error.statusCode = Number(response.status) || 502;
      throw error;
    }

    const result = await response.json();
    return { providerMessageId: String(result && result.id || "").slice(0, 200) };
  }

  return { send };
}

module.exports = {
  FIXED_EXECUTIVE_BRIEF_RECIPIENT,
  createResendExecutiveBriefTransport
};
