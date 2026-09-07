# iMessage bridge Phase 1B synthetic security validation

## Status

This branch is synthetic-only and remains draft and unmerged. It does not grant Full Disk Access, open the live Messages database, create a public endpoint, connect an app to Alice or the Chief of Staff, install a background service, or provide any message mutation capability.

Secure MCP Tunnel is acceptable as a private outbound transport. Based on its documented contract, it is not sufficient as the local caller-authorization or request-freshness boundary. The production identity and freshness gates therefore remain fail-closed.

## Implemented prototype

```text
OpenAI-hosted tunnel endpoint
        ^ outbound HTTPS only
official tunnel-client (runtime key with Tunnels Read + Use)
        | synthetic validation used fixed local stdio MCP command
dedicated signable host supervisor prototype
        | synthetic-only guarded MCP process
existing three-tool Phase 1 bridge
        | SQLite readOnly + query_only + authorizer
temporary synthetic Messages fixture
```

The supervisor starts the official `tunnel-client` directly with `shell: false`, a fixed executable, a fixed synthetic MCP command, a minimal environment, one in-flight request, one concurrent MCP request, loopback-only health binding, and the fixed `https://api.openai.com` control plane. It never restarts a failed child and exposes no public listener. The runtime key is referenced by a mode-`0600`, current-user-owned, non-symlink file and is not placed in argv or logs.

The synthetic MCP child creates a temporary SQLite fixture, opens it through the unchanged Phase 1 read-only database layer, and deletes it on shutdown. The fixture has exactly one allowlisted synthetic conversation and adversarial messages. No real database path is accepted by this runtime.

The branch also contains a synthetic Streamable HTTP resource-server prototype over a mode-`0600` Unix-domain socket. Its JWT and replay checks are local architecture tests only. The macOS platform run did not prove that ChatGPT or Secure MCP Tunnel supplies those tokens, claims, or freshness inputs.

This Linux development environment cannot produce or verify a Developer ID signature, notarization ticket, Hardened Runtime bundle, or macOS TCC attribution. A separate macOS packaging and signing pass would be required before any Full Disk Access request or live test.

## Controlled macOS validation

At PR head `afc8ff86e1df229621516d600d8ff2dba9083664`:

- `tunnel-client` v0.0.14 arm64 ran successfully on macOS.
- Control-plane authentication and polling succeeded.
- Tunnel channel `main` remained healthy and continuously polling.
- Local `/healthz` returned `live`; `/readyz` returned `ready`.
- The PR #7 synthetic MCP was selected explicitly through `--mcp.command`; the embedded stub was not used.
- After the compatibility fix, `server/discover` returned JSON-RPC `-32601 Method not found`, legacy `initialize` negotiated `2025-11-25`, and action discovery succeeded.
- Replacement draft `iMessage Phase 1B Synthetic Test 2` discovered exactly `list_allowed_message_threads`, `read_recent_messages`, and `search_allowed_messages`, all marked READ.
- Stopping and restarting the tunnel client with the identical tunnel ID, channel, and MCP command restored the replacement draft connection. This proves normal restart durability for that draft.
- The original draft's persistent disconnection was draft-specific stale or invalid saved binding state; Support identified no documented repair or rebind-in-place flow.
- The disabled draft UI exposed no Test/Run/Invoke operation. No MCP tool was invoked.
- The attempted offline Refresh replay test was inconclusive because ChatGPT disabled Refresh before an offline request could be submitted. There was no evidence of replay, but replay rejection was not proven.
- No live Messages data, Full Disk Access, write action, attachment access, Alice connection, deployment, publication, or production permission was used.

## Documented platform-contract finding

Current Secure MCP Tunnel documentation does not specify cryptographically verifiable end-user, Workspace Agent, workspace, or app identity claims delivered to the local MCP resource server on each tunneled request. It also does not specify locally verifiable signed freshness metadata or a tunnel-layer replay-rejection guarantee.

Tunnel authentication proves that an authorized tunnel runtime reached the OpenAI control plane. It does not, from the documented local contract, prove which user or agent caused a particular MCP call. Tunnel readiness also does not prove that a delivered request is newly created rather than delayed.

The production stdio path consequently rejects every read tool call because caller identity and freshness cannot be verified. The allowlist, three-tool surface, read-only database controls, attachment prohibition, and metadata-only logging remain unchanged.

## Candidate authorization architectures

| Candidate | Principal authorization | Freshness and replay | Alice-specific authorization | Assessment |
| --- | --- | --- | --- | --- |
| OAuth bearer access token validated by the local resource server | Strong if a trusted authorization server issues a signed JWT access token or supports authenticated introspection and the resource server validates exact issuer, audience/resource, subject, client, scope, expiry, and revocation | Token expiry limits exposure, but a reusable bearer token and token `jti` do not prove per-request freshness; a first delayed delivery can remain valid | Yes only if the authorization-code grant authenticates Alice and the token carries a stable Alice-specific `sub`; a Tyler, workspace, shared-app, or client-credentials subject is insufficient | Necessary identity layer, insufficient alone |
| OAuth plus DPoP-bound access token and per-request proof | Adds sender constraint using a proof bound to HTTP method, target URI, access-token hash, key, issue time, and unique proof `jti`; the resource server can require its own nonce | Short proof lifetime, single-use `jti` cache, and rotating resource-server nonce can reject replay. The freshness epoch must rotate whenever the local host or tunnel becomes unavailable so pre-disconnect proofs fail after reconnect | The OAuth token can authorize Alice specifically; DPoP proves possession by the authorized client but does not itself identify Alice | Recommended, conditional on actual ChatGPT/tunnel support |
| OAuth with mutual-TLS certificate-bound tokens | Strong sender constraint for the OAuth client | Strong against stolen-token reuse, but does not by itself give a distinct signed application request timestamp or identify queued first delivery | Usually identifies a client/device certificate, not Alice; Alice still must be represented by the OAuth token | Poor fit because TLS terminates across the tunnel and the Unix-socket resource server cannot assume visibility of the original client certificate |
| Custom signed request envelope issued by a trusted authorization gateway | Can carry an explicit user or workload principal and exact resource/tool authorization | Can bind request-body digest, issue time, expiry, nonce, and unique request ID; local single-use cache rejects replay | Yes if the gateway first authenticates Alice or a dedicated Alice workload identity | Technically viable but adds a security-critical cloud gateway and is not currently supplied by the documented OpenAI path |

A bearer access token is not enough for the required freshness boundary. DPoP is preferable to a custom envelope because it is a standardized OAuth sender-constraining mechanism, but it only works if the ChatGPT MCP client creates a unique proof and Secure MCP Tunnel forwards it intact. The local tunnel client must not create a fresh proof after polling: that would convert queued work into apparently fresh work.

## Recommended trust boundary

```text
Alice or approved dedicated workload principal
        |
        | authentication and explicit consent
        v
Trusted authorization server
  - authorization-code + PKCE for a human principal
  - stable, non-shared Alice-specific subject
  - audience restricted to this MCP resource
  - messages.read only
        |
        | DPoP-bound access token
        v
ChatGPT MCP client                         UNPROVEN PLATFORM CAPABILITY
  - creates unique per-request DPoP proof
  - includes method, URI, token hash, iat, jti, host nonce
        |
        | OpenAI Secure MCP Tunnel
        | private outbound transport only
        v
Dedicated signed macOS host
        |
        | mode-0600 Unix-domain socket
        v
Local MCP resource server                 AUTHORIZATION BOUNDARY
  - validates issuer, audience, subject, client, scope, expiry
  - validates DPoP signature, key binding, method, URI, token hash
  - requires short proof age and current host nonce
  - rejects duplicate proof jti in bounded memory
  - rotates freshness epoch on host/tunnel disconnect
  - fails closed if any input is absent or unverifiable
        |
        | only after all checks pass
        v
Existing three read-only tools
        |
        | exact local per-thread allowlist
        v
Immutable/query-only Messages database access
```

Recommended architecture: retain Secure MCP Tunnel solely as transport; use the local Unix-domain-socket Streamable HTTP resource server as the enforcement point; require a trusted authorization server to issue an audience-restricted, Alice-specific delegated access token; and require DPoP or an equivalent server-nonce-bound per-request signature. If ChatGPT cannot produce and forward that proof, do not enable live reads.

Across disconnect/reconnect, the local supervisor must invalidate the current nonce/freshness epoch as soon as the tunnel becomes unhealthy. After reconnect, only proofs created in response to the new resource-server nonce may pass. The replay cache can remain bounded and memory-only; no message persistence or request-body logging is needed.

## Alice-specific authorization

OAuth can distinguish Alice from Tyler or another workspace user only when all of these are true:

1. The authorization server authenticates the actual calling principal rather than a shared workspace or application account.
2. It issues a stable, non-reassignable and non-shared `sub` for Alice, or a separately approved dedicated Alice workload identity.
3. The local resource server validates that exact issuer-and-subject pair, the intended MCP resource audience, client ID, and `messages.read` scope.
4. ChatGPT performs authorization in the intended caller's context and does not reuse another user's grant.

A token whose subject denotes Tyler, the workspace, the ChatGPT app, or a generic tunnel client cannot be treated as Alice authorization. DPoP does not repair an ambiguous subject; it only proves possession of the key bound to the token.

## Remaining unknowns and blockers

- Whether ChatGPT custom MCP OAuth can use a trusted authorization server that issues distinct delegated subjects for Alice and Tyler.
- The actual stable `sub`, `client_id`, issuer, audience/resource, scopes, token format, lifetime, refresh behavior, and revocation latency presented during synthetic discovery.
- Whether ChatGPT supports DPoP-bound access tokens for custom MCP resources.
- Whether a `DPoP` header, access-token `cnf.jkt`, and resource-server `DPoP-Nonce` challenge survive Secure MCP Tunnel unchanged.
- Which URI ChatGPT uses for the DPoP `htu` claim and how that maps to the tunnel's logical MCP resource URI.
- Whether the local host receives a reliable tunnel disconnect signal early enough to rotate its freshness epoch.
- Whether the control plane can deliver a pre-disconnect first request after reconnect; the draft UI could not exercise this.
- Whether Alice is a platform user principal or requires a separately managed workload identity.
- macOS Developer ID signing, notarization, Hardened Runtime packaging, TCC attribution, and least-privilege Full Disk Access remain unvalidated and out of scope.
- Prompt-injection isolation has only synthetic local evidence; no cloud-agent invocation occurred.

These are live-integration blockers. PR #7 is suitable as a synthetic architecture prototype, not for enabling a live one-thread integration.

## Single safest next validation

With separate approval, perform a discovery-only OAuth/DPoP compatibility test using a disposable authorization server, two synthetic principals, a separate disabled draft, and a synthetic identity-only MCP resource server. Do not expose Messages tools or data in that test.

The server should record only categorical validation outcomes, never tokens or identifiers. During OAuth-protected initialization and tool discovery, verify whether ChatGPT:

1. obtains an audience-restricted token for the selected synthetic principal;
2. produces distinguishable subjects for the two principals;
3. presents a DPoP-bound token and unique per-request proof;
4. answers a resource-server nonce challenge;
5. preserves proof fields through the tunnel; and
6. rejects or retries safely after token revocation and tunnel reconnect.

Stop if DPoP is unsupported or if principal identity is shared or ambiguous. Do not enable or publish the draft and do not invoke an MCP tool. This is the smallest test that can validate the missing contract without touching Messages or weakening the current production gates.

## Revocation and availability

- Host: stopping or revoking the supervisor terminates the tunnel client and disables restart for that process.
- Credential: deletion, unsafe file permissions, or server-side tunnel-key revocation prevents start or tunnel authentication.
- OAuth: the resource server must reject expired or revoked access tokens and rotate its accepted signing keys deliberately.
- Freshness: disconnect or host restart rotates the nonce/epoch; pre-rotation requests fail even if their access token remains valid.
- Workspace app: disabling the app prevents upstream use, but the local host must not rely on workspace state as caller identity.
- Offline: the local wrapper has no queue, retry store, cache, index, or backfill job. Any unverifiably fresh delivery fails closed.

## Prompt-injection boundary

Message text is untrusted data. The local bridge has only the three read-only Messages tools and receives no email, calendar, browser, shell, GitHub, attachment, or investment-data capability. Synthetic adversarial tests prove hostile message strings are returned inertly and never dispatched locally. This does not prove that a cloud agent with other tools will resist prompt injection. Before any live connection, Alice must use a separate extraction stage with no downstream write-capable tools, never follow instructions or links found in messages, and require Tyler's approval for any subsequent action.

## Later live-test permissions

A later one-thread test would require separate approval for:

1. Full Disk Access granted manually only to a Developer ID-signed, notarized, Hardened Runtime host bundle. Terminal should not retain Full Disk Access.
2. Outbound HTTPS from that host to the OpenAI control plane and the specifically approved authorization server.
3. A runtime key limited to Tunnels Read + Use, stored in macOS Keychain or a mode-`0600` file reference.
4. Workspace assignment only to the approved principal and intended workspace.
5. The existing local mode-`0600` one-thread allowlist.
6. A validated Alice-specific OAuth subject and DPoP/freshness contract.

No Automation, Accessibility, Apple Events, Contacts, Screen Recording, attachment, launch-at-login, inbound-network, or Messages mutation permission is required.

## References

- [OpenAI Secure MCP Tunnel](https://developers.openai.com/api/docs/guides/secure-mcp-tunnels)
- [MCP authorization specification 2025-11-25](https://modelcontextprotocol.io/specification/2025-11-25/basic/authorization)
- [OAuth JWT access token profile, RFC 9068](https://www.rfc-editor.org/rfc/rfc9068)
- [OAuth DPoP, RFC 9449](https://www.rfc-editor.org/rfc/rfc9449)
- [OAuth mTLS and certificate-bound access tokens, RFC 8705](https://www.rfc-editor.org/rfc/rfc8705)
- [OAuth token introspection, RFC 7662](https://www.rfc-editor.org/rfc/rfc7662)
- [Official tunnel-client connector behavior](https://github.com/openai/tunnel-client/blob/master/docs/connectors.md)
- [Official tunnel-client wire protocol](https://github.com/openai/tunnel-client/blob/master/docs/protocol.md)
- [Apple Full Disk Access](https://support.apple.com/guide/mac-help/change-privacy-security-settings-on-mac-mchl211c911f/mac)
- [Apple software notarization](https://developer.apple.com/documentation/security/notarizing-macos-software-before-distribution)
