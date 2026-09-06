# iMessage bridge Phase 1B synthetic security validation

## Status

This branch is synthetic-only. It does not grant Full Disk Access, open the live Messages database, create a public endpoint, connect an app to the Chief of Staff, install a background service, or provide any message mutation capability.

## Implemented architecture

```text
OpenAI-hosted tunnel endpoint
        ^ outbound HTTPS only
official tunnel-client (runtime key with Tunnels Read + Use)
        | local stdio, raw MCP JSON-RPC
dedicated signable host supervisor prototype
        | synthetic-only guarded MCP process
existing three-tool Phase 1 bridge
        | SQLite readOnly + query_only + authorizer
temporary synthetic Messages fixture
```

The supervisor starts the official `tunnel-client` directly with `shell: false`, a fixed executable, a fixed synthetic MCP command, a minimal environment, one in-flight request, one concurrent MCP request, loopback-only health binding, and the fixed `https://api.openai.com` control plane. It never restarts a failed child and exposes no listener. The runtime key is referenced by a mode-`0600`, current-user-owned, non-symlink file and is not placed in argv or logs.

This Linux development environment cannot produce or verify a Developer ID signature, notarization ticket, Hardened Runtime bundle, or macOS TCC attribution. The wrapper is structured for later bundling, but it is not represented as signed. A macOS packaging and signing pass is required before any Full Disk Access request or live test.

The synthetic MCP child creates a temporary SQLite fixture, opens it through the unchanged Phase 1 read-only database layer, and deletes it on shutdown. The fixture has exactly one allowlisted synthetic conversation and adversarial messages. No real database path is accepted by this runtime.

## Caller identity and replay finding

The Secure MCP Tunnel authenticates `tunnel-client` to the control plane and applies organization/workspace association. That proves the tunnel runtime, not the end-user identity to a local stdio child. The official protocol forwards raw JSON-RPC to stdio; HTTP authorization headers and caller deadlines are not part of that stream.

Therefore the local stdio host cannot independently verify that a request came from Alice or another specific workspace user. It also cannot reliably distinguish fresh work from a command queued while the Mac was offline. The production policy in this prototype fails every read tool call closed with `Caller identity cannot be verified by the local stdio host.` Synthetic tests can inject an out-of-band test identity and request ID to validate the intended authorization and memory-only replay guard, but that mechanism is deliberately unavailable to live stdio traffic.

This is a live-integration blocker, not an implementation detail to guess around. Resolving it requires separate approval to use a loopback/Unix-socket Streamable HTTP MCP boundary with OAuth or a signed freshness envelope supplied by a trustworthy upstream component. The three MCP tools and per-thread allowlist can remain unchanged in either design.

## Synthetic HTTP/OAuth iteration

The approved follow-up adds a local Streamable HTTP resource server over a mode-`0600` Unix-domain socket. The tunnel client uses a logical `http://localhost/mcp` URL while dialing that socket directly. No TCP MCP listener exists.

The exact identity primitive available on this path is the OAuth bearer access token issued by the configured authorization server after ChatGPT completes authorization code with PKCE. The tunnel forwards that `Authorization` header to the HTTP MCP server. It is not an OpenAI-signed Alice identity token. The authorization server determines the token subject.

The synthetic resource server validates an Ed25519-signed JWT access token with an exact issuer, exact audience/resource, allowlisted subject, `messages.read` scope, `iat`, `nbf`, `exp`, bounded lifetime, short maximum age, key ID, and signature. It rejects exact request replay using a bounded in-memory fingerprint cache and rejects tokens minted before the current host epoch.

This establishes the local validation design but leaves one platform compatibility assumption: standard OAuth bearer tokens are normally reused for multiple requests, and OpenAI does not document a per-request signed freshness proof. A short token age plus request fingerprint rejects stale and exact repeated work, but cannot prove that a first delivery was never briefly queued by the tunnel. A true no-queue assertion would require OpenAI to forward trustworthy enqueue/deadline metadata to the MCP server or support a per-request proof such as DPoP. Live access remains blocked until that behavior is manually validated or the requirement is revised.

### Unix socket versus loopback TCP

- Unix socket: recommended for macOS packaging. Filesystem ownership and mode protect connection access, no port is allocated, and the tunnel client documents direct Unix-socket dialing for a logical HTTP MCP URL.
- Loopback TCP: supported by the server only when explicitly bound to `127.0.0.1`. It is easier to inspect, but any local process can attempt the port and port-selection/state management adds avoidable surface. OAuth validation is still required in either design.

## Revocation and availability

- Host: the operator can stop or revoke the supervisor, which sends `SIGTERM` to the tunnel client and permanently disables restart for that process.
- Credential: deletion, unsafe file permissions, or server-side key revocation prevents start or causes tunnel authentication to fail. The wrapper never retries by launching a new client.
- Workspace app: disabling the app prevents the control plane from delivering calls. This is enforced upstream because stdio has no workspace-app identity signal.
- Offline: the local wrapper has no queue, retry store, cache, index, or backfill job. The tunnel service may queue work before the local client polls it; because freshness metadata is absent at stdio, live reads remain disabled.

## Prompt-injection boundary

Message text is returned as untrusted data. The local bridge has only the three read-only Messages tools and receives no email, calendar, browser, shell, GitHub, attachment, or investment-data capability. Synthetic adversarial tests prove hostile message strings are returned inertly and never dispatched locally. This does not prove that a cloud agent with other tools will resist prompt injection. Before any live connection, Alice must use a separate extraction stage with no downstream write-capable tools, must never follow instructions or links found in messages, and must require Tyler's approval for any subsequent action.

## Later live-test permissions

A later one-thread test would require separate approval for:

1. Full Disk Access granted manually only to a Developer ID-signed, notarized, Hardened Runtime host bundle. Terminal should not retain Full Disk Access.
2. Outbound HTTPS from that host to `api.openai.com:443`, or `mtls.api.openai.com:443` if control-plane mTLS is configured.
3. A runtime API key limited to Tunnels Read + Use, stored in macOS Keychain or exposed to the host through a mode-`0600` file reference. No Tunnels Manage/admin key belongs on the host.
4. Workspace administrator assignment of the app only to the approved user and association only with the intended workspace and Platform organization.
5. The existing local mode-`0600` one-thread allowlist.

No Automation, Accessibility, Apple Events, Contacts, Screen Recording, attachment, launch-at-login, inbound-network, or Messages mutation permission is required.

## References

- [OpenAI Secure MCP Tunnel](https://developers.openai.com/api/docs/guides/secure-mcp-tunnels)
- [Official tunnel-client connector behavior](https://github.com/openai/tunnel-client/blob/master/docs/connectors.md)
- [Official tunnel-client wire protocol](https://github.com/openai/tunnel-client/blob/master/docs/protocol.md)
- [Apple Full Disk Access](https://support.apple.com/guide/mac-help/change-privacy-security-settings-on-mac-mchl211c911f/mac)
- [Apple software notarization](https://developer.apple.com/documentation/security/notarizing-macos-software-before-distribution)
