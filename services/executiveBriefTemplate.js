const ALLOWED_HTML_TAGS = new Set([
  "p",
  "h2",
  "h3",
  "ul",
  "ol",
  "li",
  "strong",
  "em",
  "br"
]);

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function validateRestrictedHtml(value) {
  const html = String(value || "").trim();
  if (!html || html.length > 100000 || html.includes("\0")) {
    throw new Error("Brief HTML must contain between 1 and 100000 characters.");
  }

  const tokens = html.match(/<[^>]*>|[^<]+/g) || [];
  if (tokens.join("") !== html) {
    throw new Error("Brief HTML contains malformed markup.");
  }

  const stack = [];
  for (const token of tokens) {
    if (!token.startsWith("<")) {
      continue;
    }
    const match = token.match(/^<(\/)?([a-z0-9]+)\s*(\/?)>$/i);
    if (!match) {
      throw new Error("Brief HTML contains unsupported markup or attributes.");
    }
    const closing = Boolean(match[1]);
    const tag = match[2].toLowerCase();
    const selfClosing = Boolean(match[3]) || tag === "br";
    if (!ALLOWED_HTML_TAGS.has(tag)) {
      throw new Error("Brief HTML contains an unsupported tag.");
    }
    if (selfClosing) {
      if (closing || tag !== "br") {
        throw new Error("Brief HTML contains malformed markup.");
      }
      continue;
    }
    if (closing) {
      if (stack.pop() !== tag) {
        throw new Error("Brief HTML contains unbalanced markup.");
      }
    } else {
      stack.push(tag);
    }
  }
  if (stack.length > 0) {
    throw new Error("Brief HTML contains unbalanced markup.");
  }
  return html;
}

function buildExecutiveBriefTemplate({ briefType, displayDate, html, text }) {
  const title = briefType === "morning" ? "MORNING EXECUTIVE BRIEF" : "EVENING EXECUTIVE BRIEF";
  const bodyHtml = validateRestrictedHtml(html);
  const bodyText = String(text || "").trim();

  return {
    html: `<!doctype html>
<html>
  <body style="margin:0;padding:0;background:#f3f4f6;color:#172033;font-family:Arial,Helvetica,sans-serif;">
    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="background:#f3f4f6;">
      <tr>
        <td align="center" style="padding:24px 12px;">
          <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="max-width:600px;background:#ffffff;border:1px solid #d9dde5;">
            <tr>
              <td style="padding:30px 28px 24px;border-top:6px solid #14213d;text-align:center;">
                <div style="font-size:18px;font-weight:700;letter-spacing:1px;color:#14213d;">BEAMAN VENTURES</div>
                <div style="margin-top:7px;font-size:11px;letter-spacing:1px;color:#586174;">OFFICE OF THE CHIEF OF STAFF</div>
                <div style="width:72px;margin:20px auto;border-top:1px solid #a68b57;"></div>
                <div style="font-family:Georgia,Times New Roman,serif;font-size:24px;line-height:1.25;color:#111827;">${title}</div>
                <div style="margin-top:13px;font-size:13px;color:#586174;">Prepared for Tyler Tashie</div>
                <div style="margin-top:5px;font-size:13px;color:#586174;">${escapeHtml(displayDate)}</div>
              </td>
            </tr>
            <tr>
              <td style="padding:4px 28px 30px;font-size:15px;line-height:1.55;color:#172033;">${bodyHtml}</td>
            </tr>
            <tr>
              <td style="padding:20px 28px;background:#14213d;text-align:center;color:#ffffff;">
                <div style="font-size:11px;font-weight:700;letter-spacing:0.8px;">CONFIDENTIAL &mdash; BEAMAN VENTURES</div>
                <div style="margin-top:7px;font-size:11px;color:#d9dde5;">Prepared by the Office of the Chief of Staff</div>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </body>
</html>`,
    text: [
      "BEAMAN VENTURES",
      "OFFICE OF THE CHIEF OF STAFF",
      "",
      title,
      "Prepared for Tyler Tashie",
      displayDate,
      "",
      bodyText,
      "",
      "CONFIDENTIAL \u2014 BEAMAN VENTURES",
      "Prepared by the Office of the Chief of Staff"
    ].join("\n")
  };
}

module.exports = {
  ALLOWED_HTML_TAGS,
  buildExecutiveBriefTemplate,
  validateRestrictedHtml
};
