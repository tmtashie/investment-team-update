# Phase 1B OAuth/DPoP discovery probe

## Purpose and boundary

This disposable probe tests whether ChatGPT custom MCP discovery can complete with an OAuth DPoP-bound access token through Secure MCP Tunnel. It is synthetic-only and has no Messages tools, database imports, attachments, persistence, indexing, or write capability.

The probe exposes OAuth protected-resource and authorization-server metadata, a synthetic authorization-code + PKCE flow, a DPoP-required token endpoint, and an authenticated MCP endpoint. MCP `tools/list` always returns an empty list. Every `tools/call` request returns JSON-RPC `-32601 Method not found`.

The authorization page offers exactly two fake principals: Synthetic Alice and Synthetic Tyler. These are test labels, not real OpenAI or workspace identities.

Secure MCP Tunnel forwards OAuth discovery and MCP authorization traffic to the private MCP resource server, but it does not tunnel the authorization server itself. The authorization-server endpoints in this process are therefore a local conformance harness only. A real ChatGPT platform run requires a separately approved disposable authorization server that is reachable by ChatGPT's OAuth flow and issues the same constrained synthetic tokens. This branch does not deploy or expose that server.

## Security model

- The authorization server signs short-lived JWT access tokens with an ephemeral Ed25519 key.
- Each token has a distinct synthetic `sub`, exact issuer, resource audience, client ID, `probe.discovery` scope, expiry, token ID, and `cnf.jkt` DPoP key thumbprint.
- Token and MCP requests require DPoP proofs with exact method and URI, a short proof age, a unique proof ID, and the current server nonce.
- MCP proofs must include the access-token hash and use the key bound into the token.
- Duplicate proof IDs, stale proofs, wrong nonces, wrong token hashes, wrong keys, expired tokens, and revoked principals fail closed.
- `SIGUSR2` rotates the in-memory freshness epoch and nonces. Proofs created before rotation fail afterward.
- `SIGUSR1` revokes Synthetic Alice only; Synthetic Tyler remains independently testable.
- Logs contain fixed event names, synthetic principal slots, booleans, and denial categories only. Tokens, authorization codes, subjects, proof IDs, request bodies, and query parameters are not logged.
- All authorization state is memory-only and disappears when the process stops.
- The HTTP server listens only on a mode-`0600` Unix-domain socket.

## Local start

The issuer and resource values must be the exact externally visible HTTPS identifiers that ChatGPT uses. Do not guess them or substitute the local Unix-socket URL. Running these endpoints on a Unix socket validates local behavior but does not make the authorization server reachable by ChatGPT.

```bash
export PHASE1B_SYNTHETIC_ONLY=1
export PHASE1B_PROBE_ISSUER='https://EXACT-SYNTHETIC-ISSUER'
export PHASE1B_PROBE_RESOURCE='https://EXACT-SYNTHETIC-RESOURCE'
export PHASE1B_PROBE_CLIENT_ID='https://EXACT-CHATGPT-CLIENT-ID-METADATA'
export PHASE1B_PROBE_REDIRECT_URI='EXACT-REDIRECT-URI-SHOWN-BY-CHATGPT'
export PHASE1B_PROBE_SOCKET="${TMPDIR:-/tmp}/phase1b-dpop-probe.sock"

node bin/phase1b-dpop-discovery-probe.js
```

The process prints only categorical JSON events to stderr. Expected startup metadata is `probe_started`, `transport: unix_socket`, and `tools: 0`.

The tunnel must use its documented Unix-socket HTTP configuration and a separate synthetic tunnel/draft. Do not reuse, edit, enable, or publish either Messages draft. Do not expose this probe directly as a public listener. The external disposable authorization server is a separate prerequisite and must not be improvised from the local MCP socket.

## Discovery-only platform procedure

1. Confirm the probe socket exists with mode `0600` and the tunnel is healthy.
2. Create a separate disabled/unpublished custom MCP draft using OAuth and the exact synthetic resource URL.
3. Complete authorization once as Synthetic Alice. Do not invoke a tool.
4. Confirm discovery reports zero tools and record only the categorical probe events.
5. Disconnect that synthetic authorization grant, then repeat with Synthetic Tyler in a separate draft or connection context. Do not reuse the Alice grant.
6. Compare only whether the principal slots are distinct and whether DPoP/token/proof validation succeeded. Never copy or log a token.
7. For the nonce check, allow the resource server's DPoP nonce challenge and observe whether ChatGPT retries with a valid proof.
8. For reconnect freshness, stop the tunnel after a valid discovery request, send `SIGUSR2` to the probe, restart the identical synthetic tunnel binding, and observe discovery only. A proof made before rotation must be rejected.
9. For revocation isolation, send `SIGUSR1` to the probe. Alice's existing access token must fail while a fresh Tyler discovery remains valid.
10. Stop before any tool invocation. Stop immediately if ChatGPT requires publishing/enabling, does not support DPoP, shares the two principal grants, or bypasses the nonce challenge.

## Evidence classification

Local tests establish only that the disposable authorization and resource server enforces the intended cryptographic policy. Platform claims require a real ChatGPT + Secure MCP Tunnel run:

- `not observed` until the tunnel delivers a DPoP proof to this HTTP resource server;
- `not observed` until two separate ChatGPT authorization contexts produce the two distinct synthetic subjects;
- `not observed` until a nonce challenge succeeds through the tunnel;
- `not observed` until reconnect and principal-revocation behavior are exercised through discovery.

Failure to complete any platform step is a compatibility finding. It is not permission to accept bearer tokens, remove nonce checks, extend proof age, share principals, or weaken the existing Messages bridge gates.
