# Read-only Apple Messages bridge: Phase 1

## Status and scope

This directory contains an isolated, local-only prototype. It is not connected to the investment application, the AI Chief of Staff, a network endpoint, or a production environment. It cannot send, reply, react, edit, delete, mark read or unread, retrieve attachments, or invoke AppleScript or UI automation.

The bridge requires Node.js 24.12 or newer because it uses the built-in `node:sqlite` module and SQLite defensive mode. This requirement applies only to the standalone bridge; it does not change the investment application's Node.js support.

## Recommended architecture

Run `bin/imessage-readonly-mcp.js` as a local stdio MCP child process on the Mac that owns the Messages library. The process:

1. Reads one fixed allowlist file at `~/.config/beaman-ventures/imessage-bridge/allowlist.json`.
2. Opens only `~/Library/Messages/chat.db` with SQLite `readOnly: true`, extension loading disabled, defensive mode enabled, `PRAGMA query_only = ON`, and an authorizer that denies write, schema, transaction, attach, detach, and non-query pragma operations.
3. Executes only source-controlled SQL with bound parameters.
4. Revalidates the current database participant set against the allowlist before each list, read, or search operation.
5. Writes newline-delimited MCP JSON-RPC responses to stdout. It creates no server socket and persists no message history.

Do not use SQLite `immutable=1` against the live Messages database. Immutable mode disables locking and change detection; SQLite warns that results can be incorrect or appear corrupt if the database changes. Messages is an active database, so ordinary read-only access is the safer choice.

## Allowlist

Copy the shape of `config/imessage-allowlist.example.json` to the fixed configuration path and replace every synthetic value locally. Never commit the local file. The file must be a regular, non-symlink file owned by the current user with mode `0600`; the bridge rejects it otherwise.

Each entry binds an opaque operator-chosen `threadId` alias to an exact private `chatGuid`, exact participant handles, and safe display labels. The raw GUID and handles are used only to query and validate membership; returned results contain the safe alias and configured display labels rather than phone numbers or Apple IDs. If group membership changes, access fails closed until Tyler reviews and updates the allowlist.

## MCP tools

All tools declare `readOnlyHint: true`, `destructiveHint: false`, `idempotentHint: true`, and `openWorldHint: false`. Every schema has `additionalProperties: false`. The stdio handler supports current `server/discover` negotiation for MCP 2026-07-28 and the 2025-11-25 legacy initialization handshake.

### `list_allowed_message_threads`

Input: an empty object.

Output: the configured safe thread alias, safe display name, and configured participant display names for allowlisted threads whose current membership still matches.

### `read_recent_messages`

Input:

```json
{
  "threadId": "non_sensitive_test_thread",
  "limit": 20
}
```

`limit` is optional and constrained to 1 through 50. Results are the most recent eligible messages, returned oldest to newest.

### `search_allowed_messages`

Input:

```json
{
  "threadId": "non_sensitive_test_thread",
  "query": "two or more characters",
  "limit": 20
}
```

Search is a bound, case-insensitive literal substring match over `message.text` within one allowed conversation. It is not SQL and cannot select multiple or unlisted conversations.

## Returned fields

Messages contain only:

- `threadId`
- `conversationDisplayName`
- configured participant display names
- `messageId`
- `replyToMessageId` when the database safely identifies a reply origin
- configured sender display name
- `direction` (`sent` or `received`)
- timestamp in ISO-like form with the applicable `America/Chicago` UTC offset
- text

The queries exclude reactions, system/group events, retracted messages, messages without plain `message.text`, and attachment data. The process does not log bodies, search text, handles, or conversation identifiers. Errors are replaced with fixed text and never include database values.

## macOS and Messages limitations

Apple does not publish or guarantee the `chat.db` schema. The prototype validates the columns it relies on and fails closed on incompatible schema changes. Current macOS data commonly uses `chat`, `handle`, `chat_handle_join`, `message`, and `chat_message_join`; Apple-epoch timestamps may be stored as seconds or nanoseconds.

On newer systems, some text exists only in the undocumented `attributedBody` typedstream when `message.text` is null. Phase 1 intentionally does not parse or return that blob, so some messages will be omitted. Attachments, audio, Digital Touch, rich app payloads, effects, reply contents, edit history, reactions, and contact-name resolution are also out of scope. Configured aliases are the only identity resolution source.

The database may change while Messages is running. Read-only SQLite locking and a one-second busy timeout favor a clean failure over stale or partial results. No live compatibility test has been performed by this development environment.

References: [Apple Full Disk Access guidance](https://support.apple.com/guide/mac-help/change-privacy-security-settings-on-mac-mchl211c911f/mac), [Node.js SQLite read-only and defensive options](https://nodejs.org/api/sqlite.html), [SQLite URI and immutable-mode cautions](https://sqlite.org/uri.html), and the [imessage-exporter schema diagnostics](https://github.com/ReagentX/imessage-exporter/blob/develop/docs/diagnostics.md). The Messages schema findings are reverse-engineered rather than an Apple-supported API contract.

## First one-thread local validation

Stop after any unexpected result; do not broaden the allowlist or inspect unrelated conversations.

1. Review this branch and its tests before granting access or using local data.
2. Install or select Node.js 24.12 or newer on Tyler's Mac. Do not install an unsigned helper or run the bridge as root.
3. In System Settings, open Privacy & Security, then Full Disk Access. Grant Full Disk Access only to the exact terminal application or signed local host process Tyler will use to start the bridge. Do not grant Automation, Accessibility, Contacts, Screen Recording, or other permissions. Restart that application if macOS requests it.
4. Do not open `chat.db` in a writable database browser. From the approved terminal, use the system SQLite client in read-only mode to identify one non-sensitive test thread by an exact known participant handle. Keep output metadata-only and bounded. Run the following command, then enter the statements at the `sqlite>` prompt so the handle is bound as data rather than interpolated into SQL:

   ```sh
   /usr/bin/sqlite3 -readonly "$HOME/Library/Messages/chat.db"
   ```

   ```sql
   .parameter init
   .parameter set @known_handle 'replace-with-the-known-test-handle'
   SELECT c.guid, COUNT(*) AS participant_count
   FROM chat AS c
   JOIN chat_handle_join AS chj ON chj.chat_id = c.ROWID
   JOIN handle AS h ON h.ROWID = chj.handle_id
   WHERE h.id = @known_handle
   GROUP BY c.ROWID, c.guid
   ORDER BY c.ROWID DESC
   LIMIT 5;
   .quit
   ```

   This step reveals identifiers, not message bodies. If the known handle can match more than one thread, Tyler must choose the intended non-sensitive thread from this bounded result rather than enumerate the library.
5. Query only that selected chat's participant metadata with another bound read-only query, then create one exact allowlist entry with a neutral `threadId` alias. Do not query `message.text` during discovery and do not expose the raw `chatGuid` to the MCP caller.

   ```sql
   .parameter init
   .parameter set @selected_guid 'replace-with-the-selected-guid'
   SELECT h.id
   FROM chat AS c
   JOIN chat_handle_join AS chj ON chj.chat_id = c.ROWID
   JOIN handle AS h ON h.ROWID = chj.handle_id
   WHERE c.guid = @selected_guid
   ORDER BY h.id
   LIMIT 32;
   .quit
   ```
6. Create `~/.config/beaman-ventures/imessage-bridge/allowlist.json` from the example, containing only that thread, then run `chmod 600` on it. Do not place the file in the repository.
7. Start `node bin/imessage-readonly-mcp.js` locally over stdio. Call `list_allowed_message_threads`, confirm it returns exactly one safe alias and no raw GUID or handle, then call `read_recent_messages` with that `threadId` and `limit: 5`. Do not search or request more history in the first validation.
8. Prove the connection is query-only in a separate disposable synthetic database using the automated test `node --test test/imessageBridge.test.js`. Do not attempt a write against Tyler's real `chat.db`; the code-level proof covers constructor read-only mode, query-only mode, the SQLite authorizer, and a rejected insert on a fixture.
9. Stop the MCP process. Revoke Full Disk Access if Tyler does not approve continued testing. Preserve no output transcript containing message bodies, and do not connect the process to the Chief of Staff.

## Approval gates after Phase 1

Tyler's explicit approval is required before the one-thread live test because it grants a process access to private Messages data. Separate approval is required for any parser of `attributedBody`, persistent cache or index, expanded allowlist, contact database access, network transport, packaging/signing, background service, deployment, or Chief of Staff connection. Sending and message mutation remain prohibited rather than merely approval-gated in this phase.
