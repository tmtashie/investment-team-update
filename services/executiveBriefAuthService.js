class ExecutiveBriefAuthError extends Error {
  constructor(code, message, statusCode) {
    super(message);
    this.name = "ExecutiveBriefAuthError";
    this.code = code;
    this.statusCode = statusCode;
  }
}

function normalizeScopes(result) {
  if (Array.isArray(result && result.scopes)) {
    return result.scopes.map(String);
  }
  return String(result && (result.scope || result.scopes) || "").split(/\s+/).filter(Boolean);
}

function createExecutiveBriefAuthService({ tokenVerifier, requiredScope = "executive_brief:send" }) {
  if (typeof tokenVerifier !== "function") {
    throw new Error("An executive brief token verifier is required.");
  }

  async function authorize(headers = {}) {
    const authorization = String(headers.authorization || headers.Authorization || "");
    const match = authorization.match(/^Bearer ([^\s]+)$/);
    if (!match || match[1].length > 4096) {
      throw new ExecutiveBriefAuthError("AUTH_REQUIRED", "Executive brief authorization is required.", 401);
    }

    let result;
    try {
      result = await tokenVerifier(match[1]);
    } catch (error) {
      throw new ExecutiveBriefAuthError("AUTH_INVALID", "Executive brief authorization is invalid.", 401);
    }
    if (!result || result.active !== true) {
      throw new ExecutiveBriefAuthError("AUTH_INVALID", "Executive brief authorization is invalid.", 401);
    }
    if (!normalizeScopes(result).includes(requiredScope)) {
      throw new ExecutiveBriefAuthError("AUTH_SCOPE_REQUIRED", "Executive brief send scope is required.", 403);
    }
    return { authorized: true, principalId: String(result.subject || "authorized-principal").slice(0, 200) };
  }

  return { authorize, requiredScope };
}

function createUnconfiguredExecutiveBriefAuthService() {
  return {
    requiredScope: "executive_brief:send",
    authorize: async () => {
      throw new ExecutiveBriefAuthError(
        "AUTH_PROVIDER_NOT_CONFIGURED",
        "Executive brief authentication is not configured.",
        503
      );
    }
  };
}

module.exports = {
  ExecutiveBriefAuthError,
  createExecutiveBriefAuthService,
  createUnconfiguredExecutiveBriefAuthService
};
