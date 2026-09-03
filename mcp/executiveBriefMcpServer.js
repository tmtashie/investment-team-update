const TOOL_NAME = "send_executive_brief";

const EXECUTIVE_BRIEF_TOOL = Object.freeze({
  name: TOOL_NAME,
  title: "Send executive brief",
  description: "Send one validated Morning or Evening Executive Brief to Tyler's fixed server-side address.",
  inputSchema: {
    type: "object",
    additionalProperties: false,
    required: ["briefType", "subject", "html", "text"],
    properties: {
      briefType: { type: "string", enum: ["morning", "evening"] },
      subject: { type: "string", minLength: 1, maxLength: 120 },
      html: { type: "string", minLength: 1, maxLength: 100000 },
      text: { type: "string", minLength: 1, maxLength: 100000 }
    }
  },
  outputSchema: {
    type: "object",
    additionalProperties: false,
    required: ["status", "deliveryId", "briefType", "briefingDate"],
    properties: {
      status: { type: "string", enum: ["sent"] },
      deliveryId: { type: "string" },
      briefType: { type: "string", enum: ["morning", "evening"] },
      briefingDate: { type: "string", pattern: "^\\d{4}-\\d{2}-\\d{2}$" }
    }
  },
  annotations: {
    readOnlyHint: false,
    destructiveHint: true,
    openWorldHint: true
  }
});

function jsonRpcError(id, code, message, data) {
  return { jsonrpc: "2.0", id: id ?? null, error: { code, message, ...(data ? { data } : {}) } };
}

function createExecutiveBriefMcpServer({ authService, deliveryService }) {
  async function handle({ body, headers = {} }) {
    const id = body && body.id;
    if (!body || body.jsonrpc !== "2.0" || typeof body.method !== "string") {
      return { statusCode: 400, body: jsonRpcError(id, -32600, "Invalid Request") };
    }

    try {
      await authService.authorize(headers);
    } catch (error) {
      return {
        statusCode: Number(error.statusCode) || 401,
        headers: error.statusCode === 401 ? { "WWW-Authenticate": "Bearer scope=\"executive_brief:send\"" } : {},
        body: jsonRpcError(id, -32001, error.message || "Authorization failed", { code: error.code || "AUTH_FAILED" })
      };
    }

    if (body.method === "notifications/initialized") {
      return { statusCode: 202, body: null };
    }
    if (body.method === "initialize") {
      return {
        statusCode: 200,
        body: {
          jsonrpc: "2.0",
          id,
          result: {
            protocolVersion: "2025-06-18",
            capabilities: { tools: { listChanged: false } },
            serverInfo: { name: "beaman-ventures-executive-brief", version: "1.0.0" },
            instructions: "This server exposes one write action for validated executive-brief delivery to a fixed recipient."
          }
        }
      };
    }
    if (body.method === "tools/list") {
      return { statusCode: 200, body: { jsonrpc: "2.0", id, result: { tools: [EXECUTIVE_BRIEF_TOOL] } } };
    }
    if (body.method !== "tools/call") {
      return { statusCode: 404, body: jsonRpcError(id, -32601, "Method not found") };
    }
    if (!body.params || body.params.name !== TOOL_NAME) {
      return { statusCode: 404, body: jsonRpcError(id, -32601, "Tool not found") };
    }

    try {
      const result = await deliveryService.sendExecutiveBrief(body.params.arguments);
      return {
        statusCode: 200,
        body: {
          jsonrpc: "2.0",
          id,
          result: {
            structuredContent: result,
            content: [{ type: "text", text: `${result.briefType} executive brief sent for ${result.briefingDate}.` }]
          }
        }
      };
    } catch (error) {
      return {
        statusCode: 200,
        body: {
          jsonrpc: "2.0",
          id,
          result: {
            isError: true,
            structuredContent: { code: error.code || "DELIVERY_FAILED" },
            content: [{ type: "text", text: error.message || "Executive brief delivery failed." }]
          }
        }
      };
    }
  }

  return { handle };
}

module.exports = { EXECUTIVE_BRIEF_TOOL, TOOL_NAME, createExecutiveBriefMcpServer };
