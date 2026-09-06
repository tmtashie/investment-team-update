"use strict";

const crypto = require("node:crypto");

const TOKEN_PATTERN = /^[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$/;

function authorizationError(reason) {
  const error = new Error("Caller authorization failed.");
  error.code = "CALLER_UNAUTHORIZED";
  error.reason = reason;
  return error;
}

function decodePart(value) {
  try {
    return JSON.parse(Buffer.from(value, "base64url").toString("utf8"));
  } catch {
    throw authorizationError("malformed_token");
  }
}

function exactAudienceMatch(audience, expected) {
  return audience === expected || (Array.isArray(audience) && audience.includes(expected));
}

function parseScopes(scope) {
  if (typeof scope !== "string") return new Set();
  return new Set(scope.split(" ").filter(Boolean));
}

function createJwtAccessTokenVerifier({
  issuer,
  audience,
  allowedSubjects,
  publicKeys,
  requiredScope = "messages.read",
  maxTokenAgeSeconds = 30,
  maxTokenLifetimeSeconds = 60,
  clockSkewSeconds = 2,
  now = () => Math.floor(Date.now() / 1000),
  minimumIssuedAt = 0
}) {
  if (
    typeof issuer !== "string" || typeof audience !== "string" ||
    !(allowedSubjects instanceof Set) || allowedSubjects.size === 0 ||
    !(publicKeys instanceof Map) || publicKeys.size === 0
  ) {
    throw new TypeError("Invalid JWT verifier configuration.");
  }

  return Object.freeze({
    verify(token) {
      if (typeof token !== "string" || token.length > 8192 || !TOKEN_PATTERN.test(token)) {
        throw authorizationError("malformed_token");
      }
      const [encodedHeader, encodedPayload, encodedSignature] = token.split(".");
      const header = decodePart(encodedHeader);
      const claims = decodePart(encodedPayload);
      if (
        !header || header.alg !== "EdDSA" || typeof header.kid !== "string" ||
        (header.typ !== undefined && header.typ !== "JWT") ||
        Object.keys(header).some((key) => !new Set(["alg", "kid", "typ"]).has(key))
      ) {
        throw authorizationError("unsupported_token");
      }
      const publicKey = publicKeys.get(header.kid);
      if (!publicKey) throw authorizationError("unknown_key");
      let signature;
      try {
        signature = Buffer.from(encodedSignature, "base64url");
      } catch {
        throw authorizationError("malformed_token");
      }
      const signed = Buffer.from(`${encodedHeader}.${encodedPayload}`, "ascii");
      if (!crypto.verify(null, signed, publicKey, signature)) throw authorizationError("invalid_signature");

      const current = now();
      if (claims.iss !== issuer) throw authorizationError("wrong_issuer");
      if (!exactAudienceMatch(claims.aud, audience)) throw authorizationError("wrong_audience");
      if (typeof claims.sub !== "string" || !allowedSubjects.has(claims.sub)) {
        throw authorizationError("wrong_subject");
      }
      if (
        !Number.isInteger(claims.iat) || !Number.isInteger(claims.exp) ||
        claims.iat < minimumIssuedAt || claims.iat > current + clockSkewSeconds ||
        current - claims.iat > maxTokenAgeSeconds ||
        claims.exp <= current - clockSkewSeconds ||
        claims.exp <= claims.iat ||
        claims.exp - claims.iat > maxTokenLifetimeSeconds ||
        (claims.nbf !== undefined && (!Number.isInteger(claims.nbf) || claims.nbf > current + clockSkewSeconds))
      ) {
        throw authorizationError("stale_token");
      }
      if (typeof claims.jti !== "string" || claims.jti.length < 8 || claims.jti.length > 200) {
        throw authorizationError("missing_token_id");
      }
      if (!parseScopes(claims.scope).has(requiredScope)) throw authorizationError("missing_scope");
      return Object.freeze({ subject: claims.sub, tokenId: claims.jti, expiresAt: claims.exp });
    }
  });
}

module.exports = {
  authorizationError,
  createJwtAccessTokenVerifier
};
