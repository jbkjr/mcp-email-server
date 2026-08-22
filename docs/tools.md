# MCP Tools

> **Version scope:** The mail-only catalog on this page is Local Email App V2
> behavior. See [Version availability](getting-started.md#version-availability)
> before using this contract with a PyPI installation.

mcp-email-server exposes bounded account discovery, message, mailbox, and
composition operations as MCP tools. Tool schemas are generated from the running server, so the MCP client
can inspect each parameter and response type directly.

## Typical workflow

Most message workflows follow this sequence:

1. Call `list_available_accounts` to select an `account_name`.
2. Call `list_emails_metadata` to search a mailbox and obtain `email_id` values.
3. Pass those IDs to a read or mutation tool with the same mailbox name.
4. Call `get_emails_content` only for messages whose bodies are needed.

This separates lightweight metadata searches from potentially large body
retrievals.

MCP input schemas advertise the enforceable string and collection envelopes from
the centralized application limits, including `maxLength`, `minItems`, and
`maxItems` for account/mailbox names, UID collections, recipients, attachments,
and flags. JSON Schema counts characters while the application limits UTF-8
bytes, so every application service independently revalidates direct and MCP
callers; aggregate recipient and payload limits also remain application-owned.

## Account resource

The resource URI `email://{account_name}` returns the same stable non-secret
capability record used by account discovery. It does not return configuration or
masked credential objects.

## Account tools

### `list_available_accounts`

Lists all enabled accounts from the selected configuration mode as explicit
capability records. Each record contains `account_name`, `account_type`,
`description`, optional `email_address`, `can_receive`, and `can_send`. Account
descriptions are limited to 4 KiB of UTF-8 data and expose the same structural
bound in the output schema. In managed mode, disabled accounts are omitted before any credential lookup or provider
access. Use only an account with `can_receive=true` for mail reads and
`can_send=true` for `send_email`. Text content, structured content, and the output
schema describe the same fields.

If the result is empty, account setup is unavailable over MCP. The agent should
ask the user to run `mcp-email-server ui` or the documented interactive CLI in
their own terminal and must never request or relay credentials. The output schema
and application boundary allow at most 1,000 accounts; the canonical JSON must
also fit the shared 8 MiB response ceiling. Oversized authority data is rejected
with `limit_exceeded` rather than truncated.

MCP exposes no account, endpoint, policy, catalog, or credential mutation tool in
either mode. Use `mcp-email-server ui` or the user-operated `config` and
`account` CLI commands. This prevents an agent or chat transcript from becoming
a credential handoff surface. The complete tool names, descriptions, input and
output schemas, annotations, resource template, and visibility are static and
covered by an exact catalog contract test.

`add_email_account`, which exists in PyPI 0.16.0 and earlier, is intentionally
absent from Local Email App V2 rather than renamed. See
[Upgrading to Local Email App V2](getting-started.md#upgrading-to-local-email-app-v2)
for client discovery and configuration migration steps.

## Agent planning annotations

Every tool advertises reviewed MCP `readOnlyHint`, `destructiveHint`,
`idempotentHint`, and `openWorldHint` values:

| Tools                                                                                                                                     | Read-only | Destructive | Idempotent | Open world |
| ----------------------------------------------------------------------------------------------------------------------------------------- | --------- | ----------- | ---------- | ---------- |
| `list_available_accounts`, `list_allowed_recipients`, `list_allowed_senders`                                                              | yes       | no          | yes        | no         |
| `list_emails_metadata`, `list_mailboxes`, `list_labels`, `get_email_labels`                                                               | yes       | no          | yes        | yes        |
| `get_emails_content`                                                                                                                      | no        | no          | yes        | yes        |
| `send_email`, `forward_email`, `save_to_mailbox`, `copy_emails`, `create_folder`, `create_label`, `apply_label`                           | no        | no          | no         | yes        |
| `set_email_flags`, `mark_emails_as_read`                                                                                                  | no        | no          | yes        | yes        |
| `delete_emails`, `move_emails`, `archive_emails`, `remove_label`, `delete_label`, `download_attachment`, `delete_folder`, `rename_folder` | no        | yes         | no         | yes        |

<!-- Port slots. Tools are being ported onto this architecture in parallel; each cluster
     adds its tool names to the matching existing row above rather than appending a new
     row, so the table stays grouped by annotation set. Slot order everywhere: A, B1, C, B2.
     port-slot B2: create_label is non-destructive; delete_label is destructive; apply_label
     is non-destructive. -->

`get_emails_content` is conservatively non-read-only because
`mark_as_read=true` changes remote flags. Download is destructive because the
caller-selected destination may be replaced. Send, forward, and append create
externally meaningful effects but do not delete or replace an existing mailbox
item, so their destructive hint is false while their read-only and idempotent
hints are also false. `copy_emails`, `create_folder`, `create_label`, and
`apply_label` are additive for the same reason, while `delete_folder`,
`rename_folder`, and `delete_label` are destructive because they remove or
replace an existing mailbox together with everything it holds.

Annotations are advisory host/agent planning hints, not authorization or a
safe-retry guarantee. Tool descriptions, current policy, typed outcomes, and the
rule against replay after an ambiguous effect remain authoritative.

## Reading and searching

### `list_emails_metadata`

Searches one mailbox without downloading message bodies.

Important parameters include:

| Parameter                     | Default  | Description                                 |
| ----------------------------- | -------- | ------------------------------------------- |
| `account_name`                | Required | Configured account identifier.              |
| `page`                        | `1`      | One-based result page.                      |
| `page_size`                   | `10`     | Number of results per page, from 1 to 100.  |
| `mailbox`                     | `INBOX`  | Mailbox to search.                          |
| `before` / `since`            | None     | UTC datetime boundaries.                    |
| `subject`                     | None     | Subject filter.                             |
| `from_address` / `to_address` | None     | Address filters.                            |
| `seen`                        | None     | Filter by read status.                      |
| `flagged`                     | None     | Filter by flagged or starred status.        |
| `answered`                    | None     | Filter by replied status.                   |
| `body`                        | None     | Search message bodies with IMAP `BODY`.     |
| `text`                        | None     | Search headers and bodies with IMAP `TEXT`. |
| `has_attachment`              | None     | Apply a multipart attachment heuristic.     |
| `order`                       | `desc`   | Return ascending or descending results.     |

The response contains pagination metadata, a filtered `total`, and message
metadata including `email_id`, `message_id`, subject, sender, recipients, and
date. To and Cc are parsed as structured RFC 5322 address fields: a comma inside
a quoted display name is preserved, while addresses inside a group are returned
as individual recipient entries. Because this operation fetches headers only,
its `attachments` field is empty. `get_emails_content` populates attachment names
from the full message.

`has_attachment` uses a `multipart/mixed` heuristic. It can miss inline content
or report multipart messages that do not contain a conventional attachment.

When a sender allowlist is configured, blocked messages are removed before
pagination, so `total` and page sizes describe only visible messages.

The application keeps a rebuildable SQLite projection for unfiltered mailbox
pages. It uses that projection only after a small IMAP `STATUS` probe confirms
the same UIDVALIDITY, UIDNEXT, and message count and the projection covers the
whole mailbox. Text, date, address, flag, body, and attachment filters remain on
the bounded IMAP path so provider-specific search semantics and mutable flags
stay authoritative. ASCII filter values retain their exact text through IMAP
atom or quoted-string encoding. Non-ASCII filter values use synchronizing UTF-8
literals with `CHARSET UTF-8`; a provider that rejects that charset returns a
bounded search failure rather than receiving malformed raw UTF-8 command text.
Date criteria always use the protocol's English month tokens regardless of the
server process locale. A response normally omits `warnings`; if a validated IMAP
result was returned but its rebuildable projection could not be persisted, the
response includes `warnings: ["projection_write_failed"]`. It never includes the
local exception detail.

A refresh stores at most the 1,000 most recent UIDs and claims complete coverage
only when the whole mailbox fits that window and provider state is unchanged
across the refresh. Provider fallback accepts at most 10,000 unique canonical
single UIDs; ranges, sets, zero, duplicates, and values outside the IMAP UID
range are rejected before any UID FETCH. Metadata header requests use IMAP
partial fetches, with limits of 64 KiB per
message and 4 MiB total per metadata query or refresh. Each wire FETCH is also
sized below that aggregate ceiling. Missing, duplicate, or mismatched sender or
INTERNALDATE evidence is a bounded error because the server cannot otherwise
prove the exact total or ordering. Transport and protocol failures are mapped to
bounded categories without returning provider-controlled detail. If a work or
payload ceiling is exceeded,
the tool returns a bounded error instead of an inexact `total`, partial page, or
unbounded projection.

### `get_emails_content`

Fetches the body of one or more messages by `email_id`.

| Parameter         | Default  | Description                                                     |
| ----------------- | -------- | --------------------------------------------------------------- |
| `account_name`    | Required | Configured account identifier.                                  |
| `email_ids`       | Required | IDs returned by `list_emails_metadata`.                         |
| `mailbox`         | `INBOX`  | Mailbox containing the messages.                                |
| `mark_as_read`    | `false`  | Mark successfully retrieved messages as read.                   |
| `body_offset`     | `0`      | Character offset at which body output starts.                   |
| `max_body_length` | `20000`  | Maximum body characters returned per message, from 1 to 100000. |

If a body extends beyond the requested window, the returned body ends with
`...[TRUNCATED]`. Fetch the next chunk by increasing `body_offset` by
`max_body_length`.

Set `max_body_length` to `0` or `null` to disable truncation: the whole body from
`body_offset` onward is returned and the `...[TRUNCATED]` marker is never
appended. An untruncated body is still subject to the shared per-message and
aggregate body byte ceilings, so an oversized message returns a bounded limit
error for the request rather than a silently shortened body.

The batch response reports requested and retrieved counts and includes
`failed_ids` for messages that could not be fetched. A full-message literal from
a successful IMAP FETCH is parsed regardless of its byte length; protocol
metadata without a message literal is not treated as content. Each returned
email also includes nullable `in_reply_to` and `references` values from the
corresponding RFC headers. `references` is returned as one decoded, unfolded string with
folding spaces and tabs normalized. Missing and whitespace-only values become
`null`; if an invalid message repeats either header, the parser's first observed
value is returned. This is untrusted observational header data, not a validated
list of Message-IDs. Well-formed values can be passed back to the compose tools,
but malformed values containing other control characters can be returned and
will be rejected by compose validation. These fields are available only from
full-content reads: they are not part of `list_emails_metadata` and are not
persisted in the SQLite metadata projection.

MIME body extraction does not descend into attachment subtrees. In particular,
the body of an attached or forwarded `message/rfc822` message is never merged
into the containing message body, even when that part has no filename. If one
text part declares an unknown charset or contains invalid bytes, it falls back
to UTF-8 replacement decoding without hiding the other readable parts.

A message with no `text/plain` part has its HTML converted to plain text with the
standard library's HTML parser. `script`, `style`, and `head` content is dropped;
paragraphs, line breaks, headings, list markers, tab-separated table rows, and
blockquote markers are preserved; and a link's target is appended after its text
when it adds information. Anchor-only, `mailto:`, and `javascript:` targets are
never surfaced, and the scheme is checked after removing embedded control
characters. HTML email is never rendered or executed — the result is text.

A request accepts 1 to 500 canonical positive decimal ASCII IMAP UIDs; zero,
leading zero, non-ASCII digits, signs, ranges, sets, and values above the IMAP UID
limit are rejected before provider access. The provider adapter repeats this
validation before opening an IMAP connection as defense in depth. Raw messages
above 50 MiB are rejected before MIME parsing. The production provider also
counts each returned body's UTF-8 bytes before retaining it and stops immediately
if the batch would exceed the 50 MiB aggregate body budget; the application
validates the aggregate again at its provider boundary. Each returned thread
header is limited to 64 KiB of UTF-8 data and both count toward the 4 MiB
aggregate returned-header budget. The production provider enforces these header
budgets before retaining each parsed result, and the application independently
revalidates them at its provider boundary. Oversized values fail explicitly
rather than being truncated into an invalid thread chain.

When the complete valid batch exceeds the inline MCP response ceiling, the
server writes the canonical JSON response to a randomly named owner-only file in
a process-private temporary directory. The bounded response then has
`content_omitted=true`, an empty `emails` preview, and
`output_file_path`, `output_media_type`, `output_bytes`, `output_sha256`, and
`output_lifetime` fields. A local MCP host can inspect that exact path with its
own filesystem tool. The file is available only until the email-server process
exits; copy needed content before restarting. The server does not add a generic
file-download tool or remote URL. Spill requires the complete POSIX owner/no-follow
profile or the local fixed NTFS Windows DACL/reparse/identity profile. Windows
crash remnants are removed only after bounded prefix, type, owner, DACL, and
identity validation. Spill never falls back to a broadly accessible temporary
file. Without the required profile, bounded inline results remain available,
while a batch that requires spill returns a bounded error. The process-lifetime
notice still applies.

Body retrieval always uses IMAP PEEK. When requested, successfully retrieved IDs
are deduplicated and marked through the same application mutation workflow as
`mark_emails_as_read` in batches of at most 100. A known mark failure is logged
but does not discard successfully retrieved content. An unknown or
reconciliation-needed mark outcome stops later mark batches so the application
does not continue after ambiguous state.

## Composing messages

### `send_email`

Sends a message through the selected account's SMTP server. It supports:

- To, CC, and BCC recipients.
- Markdown bodies rendered to email-safe HTML, or pre-formatted raw HTML bodies.
- Attachments from file paths available to the server process. Relative paths use the process working directory; absolute paths are recommended.
- `Reply-To`, `In-Reply-To`, and `References` headers.

The tool is always present in the stable MCP catalog. The selected
`account_name` must itself be enabled and send-capable; an IMAP-only account is
rejected before SMTP access.

If a recipient allowlist is configured, every To, CC, and BCC address must be
allowed. SMTP delivery reports accepted, rejected, and unknown recipients
separately when the result is partial or ambiguous. Failed and unknown targets
include reviewed fixed diagnostics when available, for example
`smtp-mail-rejected`, `smtp-recipient-rejected`, `smtp-data-rejected`,
`smtp-data-unknown`, or `provider-timeout`. Unrecognized detail and raw provider
response text are omitted.

Internationalized addr-specs in the envelope or From, Sender, To, Cc, Bcc, or
Reply-To fields, and non-ASCII Message-ID, In-Reply-To, or References syntax,
require the provider's SMTPUTF8 extension. The server requests `SMTPUTF8` and
serializes the complete message with the matching policy. If the extension is
unavailable, every target fails with `smtp-utf8-unsupported` before `MAIL FROM`,
`RCPT TO`, or message data is sent. A non-ASCII display name with an ASCII
addr-spec is encoded as an ordinary RFC 5322 display name and does not by itself
require SMTPUTF8.

Saving the Sent copy is a second IMAP effect and is reported in its own
`sent-copy` section; a failed or unknown copy never changes an accepted delivery
into a failure. Do not retry the whole send to repair a Sent copy. Sent-copy
APPEND payloads use CRLF line endings for compatibility with strict IMAP
providers. An internationalized Sent copy additionally requires RFC 6855
`ENABLE` plus `UTF8=ACCEPT` or `UTF8=ONLY`; unsupported negotiation is reported
as `utf8-append-unsupported` without changing the successful SMTP outcome.

### `save_to_mailbox`

Composes a message and appends it to an IMAP mailbox instead of sending it. It
works without SMTP and is useful for drafts or templates. It shares recipient,
body, attachment, and threading fields with `send_email`, adds `mailbox` and
`flags`, and does not support `reply_to`.

The default mailbox is `Drafts`. When no explicit flags are supplied, the
message is saved with `\Draft` and `\Seen`. The response includes the RFC
message ID. It includes an assigned IMAP `email_id` only when the server returns
RFC 4315 `APPENDUID`; otherwise the value is `unknown`, and the target mailbox
must be searched before a later operation can address the saved message.

The same recipient allowlist used by `send_email` applies to this tool. The
complete MIME payload is serialized with CRLF line endings before IMAP APPEND for
compatibility with strict providers. Saved-message flags may be system flags or
provider keywords, but each must be one valid IMAP atom; legal values such as
`$Forwarded`, `project.name`, and `123flag` are accepted, while whitespace,
controls, and IMAP protocol specials are rejected.

The server refreshes capabilities before mailbox selection. A message with
internationalized address or thread-header syntax requires RFC 6855, and a
`UTF8=ONLY` server requires `ENABLE UTF8=ACCEPT` even for an ordinary ASCII-header
message. In an enabled session, LIST names retain their literal UTF-8 spelling
and internationalized mailbox arguments use escaped UTF-8 rather than Modified
UTF-7. Only a message whose headers require RFC 6532 uses the RFC 6855 UTF8
literal form. Missing capability or incomplete ENABLE evidence fails before
SELECT/APPEND with `utf8-append-unsupported`. A known APPEND success without
`APPENDUID` returns `email_id: unknown`. A lost APPEND result is instead tagged
`unknown`; the server does not replay it because that could create a duplicate
draft.

### `forward_email`

Forwards an existing message to new recipients through the selected account's
SMTP server. The server reads the source message over IMAP, composes a new
message below an optional note from the caller, and re-attaches the original's
attachments.

| Parameter             | Default  | Description                                   |
| --------------------- | -------- | --------------------------------------------- |
| `account_name`        | Required | Configured account identifier.                |
| `email_id`            | Required | UID of the source message to forward.         |
| `recipients`          | Required | Addresses that receive the forwarded message. |
| `source_mailbox`      | `INBOX`  | Mailbox that contains the source message.     |
| `body`                | `""`     | Note placed above the forwarded content.      |
| `cc`                  | None     | Additional CC recipients of the forward.      |
| `bcc`                 | None     | Additional BCC recipients of the forward.     |
| `include_attachments` | `true`   | Re-attach the source message's attachments.   |

The subject is derived from the source message as `Fwd: <original subject>`. A
source subject that already begins with `Fwd:` in any letter case is not
prefixed a second time.

The forwarded content is appended below the caller's note as a
`Forwarded message` block reporting the original's From, Recipients, Date, and
Subject. That block reports `Recipients:` rather than `To:` because the parsed
recipient list folds in Cc entries.

The note is caller-authored and is rendered as Markdown like any other body. The
forwarded block is quoted evidence read off another message, so its markup
characters are escaped before rendering: a source body containing `<b>` or a
`<script>` element reaches the recipient as the literal text it was, never as
live markup the account owner did not write.

The block is re-composed from the parsed plain-text body, so the original's HTML
formatting is not preserved in the quoted text. The forwarded content is never
silently truncated: the composed body, including any note you supply, is bounded
at 1 MiB and an oversized forward is rejected outright. Forward a message when
the recipient needs its attachments and substance; when byte-exact rendering
matters, save the parts with `download_attachment` and compose the message
explicitly with `send_email`.

Attachments carried into the forward keep the source part's MIME main type,
subtype, and parameters instead of being coerced into `application/*`. Set
`include_attachments=false` to forward only the text.

The tool is always present in the stable MCP catalog. The selected
`account_name` must itself be enabled and send-capable; an IMAP-only account is
rejected before SMTP access, exactly as for `send_email`.

A forward performs three independent provider effects: the IMAP read of the
source message, SMTP delivery, and the IMAP Sent copy. Current account authority
and policy are revalidated before each one. If the source message cannot be
read, the call fails before any SMTP session is opened, so a forward is never
delivered without the content and attachments it was supposed to carry. Delivery
and sent-copy outcomes are reported separately under the same rules as
`send_email`, and an ambiguous SMTP outcome is reported `unknown` and is never
replayed automatically.

Reading the source message is a mail read. When a sender allowlist is
configured, a message from a blocked sender is indistinguishable from a missing
message, so the forward fails without revealing that the message exists. The
recipient allowlist applies to the forward's To, CC, and BCC addresses exactly
as it does for `send_email`.

For a worked example, see
[Forward a message with its attachments](guides.md#forward-a-message-with-its-attachments).

### Markdown message bodies

Every composed message body — `send_email`, `forward_email`, and
`save_to_mailbox` alike — is written in Markdown and rendered to email-safe HTML
before the MIME container is built. Headings, bold and italic text, links,
bullet and numbered lists, tables, and fenced code blocks are all supported, and
single newlines become line breaks so ordinary prose keeps the shape you wrote
it in. The rendered document carries minimal inline styles rather than CSS
classes, because mail clients cannot be relied on to keep a stylesheet.

Set `html=true` on `send_email` or `save_to_mailbox` when the body is already
pre-formatted raw HTML. That suppresses rendering and sends the body exactly as
supplied; it is not a way to ask for a plain-text message. `forward_email` has
no `html` parameter: its note is always Markdown.

Rendering changes only the body part's subtype, never an address or threading
header, so it has no effect on whether a message requires SMTPUTF8.

### Quoted replies

When `send_email` is given `in_reply_to`, the server reads the message being
replied to over IMAP and appends it below your body as a collapsible quote
block, the way a mail client would. Set `quote_reply=false` to reply without
quoting.

The original is looked up by its Message-ID in INBOX first, then in the account's
Sent folder, so replying to your own message quotes it too. Both mailboxes are
searched inside a single IMAP session. If the original has an HTML body it is
quoted as markup with its document wrappers stripped; otherwise its text is
escaped, line-broken, and truncated at 5000 characters with an explicit
`[...quoted text truncated]` marker.

Reading the original is a read, not an effect, and it happens before any SMTP
session is opened. The two outcomes are deliberately different:

- The original is not in any searched mailbox. Nothing is wrong and there is
  nothing to quote, so the reply is sent unquoted.
- The original is there but cannot be read — a transport fault, an unparseable
  message, one that exceeds the raw message size limit. The call fails before
  SMTP, because a caller who asked for a quoted reply must never silently get an
  unquoted one.

When a sender allowlist is configured, a blocked sender's message is treated as
absent, so the allowlist cannot be used to probe which messages exist.

The merged body is revalidated against the 1 MiB body bound after the quote is
appended, so a large quote is rejected rather than silently truncated.

The quote's markup follows the account's mail service, because clients only
collapse a quote block they recognize. The service is detected from the IMAP
host: `imap.gmail.com` is Gmail, a loopback host with certificate verification
disabled is a ProtonMail Bridge, and anything else uses a generic blockquote.
Set `email_service` on the account (or `MCP_EMAIL_SERVER_EMAIL_SERVICE`) to
`protonmail`, `gmail`, or `generic` when detection guesses wrong — see
[Configuration](configuration.md).

<!-- port-slot C: send-path documentation goes here, above this marker — Markdown message
     bodies and quoted replies extend `send_email` rather than adding tools. -->

## Mailbox and mutation tools

### `list_mailboxes`

Lists IMAP mailboxes with their names, hierarchy delimiters, and flags. Call it
before moving or saving messages when provider-specific folder names are not
known.

`pattern` defaults to `*`, and `reference` defaults to an empty string. The
account name and both IMAP LIST values are validated before provider access;
`pattern` must be non-empty and pattern/reference values are each limited to
1,024 UTF-8 bytes. Literal mailbox names are reassembled at their declared byte
length, tagged LIST completion text is not returned as a mailbox, and malformed
literal framing fails the request. Special-use flags such as `\Sent` are matched
case-insensitively for folder discovery.

### `set_email_flags`

Adds or removes approved IMAP flags from one or more message IDs in the selected
mailbox. `operation` must be `add` or `remove`, and applies to every supplied
flag. The non-empty `flags` list accepts unique values from:

- `\Seen`
- `\Flagged`
- `\Answered`
- `\Draft`

The provider sends one UID-scoped `+FLAGS.SILENT` or `-FLAGS.SILENT` operation
per message so results retain caller order and per-ID evidence. The operation is
logically idempotent, but an `unknown` result is not retried automatically
because the mailbox UID epoch or current authority may have changed.

`\Deleted` is intentionally rejected and remains owned by `delete_emails`,
which applies target-scoped expunge safety. `\Recent` is server-controlled, and
provider-specific keywords are not part of the portable public contract. To
mark a message unread, remove `\Seen`.

### `mark_emails_as_read`

Marks one or more message IDs as read in the selected mailbox. This focused
common-workflow tool uses the same implementation as `set_email_flags` with
`operation="add"` and `flags=["\\Seen"]`.

### `move_emails`

Moves messages from `source_mailbox`, which defaults to `INBOX`, to a required
`destination_mailbox`. Native IMAP `MOVE` is preferred. The COPY-and-delete
fallback is available only when the server advertises `UIDPLUS`, allowing the
source to be removed with target-scoped `UID EXPUNGE`; otherwise the operation
fails before copying a message.

### `archive_emails`

Moves messages to the account's archive mailbox. The server first uses the RFC
6154 `\Archive` mailbox flag and then falls back to `Archive`, `Archives`, or
`[Gmail]/All Mail`. Archive uses the same native-MOVE or safe UIDPLUS fallback
rules as `move_emails`.

### `delete_emails`

Deletes one or more messages from the selected mailbox. The provider must
advertise `UIDPLUS`: the server flags and expunges only the requested UIDs with
`UID EXPUNGE` and never sends mailbox-wide `EXPUNGE`. Without `UIDPLUS`, the
operation fails before adding the `\Deleted` flag.

An all-known-success mutation keeps the existing success sentence. Partial or
ambiguous results use tagged `succeeded`, `failed`, and `unknown` sections in
input order. Unknown targets can include a fixed substep tag such as `store`,
`copy`, or `expunge-after-copy`. `unknown` means the provider effect may have
started but its final result was lost; the server does not replay it
automatically, and every result containing `unknown` includes a `reconciliation
needed` warning. The same warning also appears when a known provider effect may
be authoritative but the rebuildable local metadata projection could not be
invalidated.

Mutation requests accept 1 to 100 unique canonical positive decimal IMAP UIDs.
`set_email_flags` accepts one to four unique approved flags and exactly one
add/remove operation. Mailbox names are limited to 1,024 UTF-8 bytes. Compose
requests allow at most 100 total To/CC/BCC entries of at most 1,024 UTF-8 bytes
each; every entry must
contain exactly one address. Compose requests also allow a 64 KiB UTF-8 subject,
a 1 MiB UTF-8 body, and 20 attachments. Threading and Reply-To values
are limited to 64 KiB each. Each attachment path is limited to 4,096 bytes,
each existing attachment to 25 MiB, and their combined size to 50 MiB. Outbound
attachments preserve the inferred MIME main type and subtype (for example,
`image/png` remains `image/png`) instead of coercing every file to
`application/*`. Saved messages accept at most 100 flags of 128 bytes each before protocol syntax
validation. Mailbox names, recipient/header values, and subjects reject control
characters before provider access.

When a sender allowlist is active, blocked messages are never changed. See
[Sender allowlist](security.md#sender-allowlist) for the privacy behavior of
blocked IDs.

### `copy_emails`

Copies messages from `source_mailbox`, which defaults to `INBOX`, into a
required `destination_mailbox`, leaving the originals in place. This is the
`COPY` half of `move_emails` with no `\Deleted` flag and no expunge, so it needs
no `UIDPLUS` capability and cannot remove a message. Because a copy is additive
rather than a relocation, the source and destination may be the same mailbox.

Results use the same tagged per-ID format as the other mutation tools. Only the
destination mailbox's local metadata projection is invalidated: the source
mailbox is unchanged by a copy.

`copy_emails` is not gated by `enable_folder_management` — it moves message
copies between existing mailboxes rather than changing the folder layout.

### `create_folder`

Creates one IMAP mailbox by name. Call `list_mailboxes` first to learn the
server's hierarchy delimiter: a nested name is built with that delimiter (for
example `Archive/2026` or `Archive.2026`), and using the wrong one silently
creates a top-level folder.

### `delete_folder`

Deletes one IMAP mailbox by name. Most servers refuse to delete a mailbox that
still contains messages or child mailboxes, and report that refusal as a
`failed` result rather than an error.

### `rename_folder`

Renames one IMAP mailbox, carrying its messages and child mailboxes with it.
`old_name` and `new_name` must differ. Both spellings' local metadata
projections are invalidated because a rename changes where every message in the
mailbox lives.

`create_folder`, `delete_folder`, and `rename_folder` change the account's
folder layout, so they require `enable_folder_management=true`. See
[Folder management access](security.md#folder-management-access). The three
tools stay visible when the policy is off and refuse at call time, so a client
never has to re-read the catalog after a policy change.

A mailbox-shape result reports a single status rather than a per-ID batch. An
explicit server rejection is `failed` and means nothing changed; a lost or
cancelled response is `unknown`, is not retried automatically, and carries a
`reconciliation needed` warning because the mailbox may already have changed.

<!-- Port slots for new tool sections. Each cluster adds its `###` sections directly
     ABOVE its own marker so parallel ports produce non-overlapping hunks, and the
     final section order matches the fixed slot order A, B1, C, B2.
     port-slot B2: `create_label`, `delete_label`, `apply_label` — these belong in the
     "Label tools" section below rather than here, alongside the B1 label tools that
     already share its `Labels/` preamble. -->

## Label tools

A label is an ordinary IMAP mailbox whose name begins with the literal prefix
`Labels/`, which is how ProtonMail and ProtonMail Bridge expose labels over
IMAP. The prefix is part of the mailbox name rather than a server hierarchy
path, so it always uses `/` regardless of the delimiter the server advertises,
and an account that does not follow the convention simply has no labels. A
message carries a label because a copy of it lives in that label's mailbox;
the copy and the original share one `Message-ID`, which is how the label tools
link them.

Label names are limited to 1,017 UTF-8 bytes so that `Labels/<label_name>`
still fits inside the 1,024-byte mailbox bound, reject control characters, and
may not repeat the `Labels/` prefix themselves.

### `list_labels`

Lists the account's labels. Each entry reports the label `name` with the prefix
stripped, the `full_path` mailbox that stores it, the server's hierarchy
`delimiter`, and the mailbox `flags`. The bare `Labels/` container is a grouping
mailbox, not a label, and is never returned. The listing is bounded by the same
1,000-mailbox and result-size ceilings as `list_mailboxes`.

### `get_email_labels`

Reports which labels hold a copy of one message, as a list of label names.

The message's `Message-ID` is read once from `mailbox`, then every label mailbox
is probed for that identifier inside a single IMAP session, and the whole
workflow shares one provider deadline so fanning out across labels cannot extend
the budget one folder at a time. The `Message-ID` is sent as a quoted IMAP
search value with quoted-specials escaped; a `Message-ID` that is not printable
ASCII cannot be searched and fails the request instead of being interpolated raw.

The result is empty when the message cannot be read, carries no `Message-ID`, or
has no labels — a message hidden by the sender allowlist is indistinguishable
from a missing one. A label mailbox that cannot be selected or searched is
skipped, so one stale folder cannot hide the labels that did resolve.

### `remove_label`

Removes one label from one or more messages.

For each `email_id`, the message's `Message-ID` is read from `source_mailbox`,
its copy is located in `Labels/<label_name>` by a quoted `Message-ID` search, and
only that copy is marked `\Deleted` and removed with a target-scoped UID
EXPUNGE. Every effect is scoped to the label mailbox: the message named by the
caller is never modified, and neither is any other message in the label mailbox.
Removal requires the IMAP UIDPLUS capability, exactly as `delete_emails` does.

Per-ID results preserve caller order and distinguish success, failure, and
`unknown`. Unlike the other batch tools, a failure reports a reviewed reason
alongside the ID, because those reasons are actionable and non-sensitive:

| Detail                   | Meaning                                                                        |
| ------------------------ | ------------------------------------------------------------------------------ |
| `message-id-missing`     | The source message has no `Message-ID` to match a copy with.                   |
| `message-id-unsupported` | Its `Message-ID` is not printable ASCII and cannot be searched.                |
| `label-not-found`        | The label does not hold a copy of that message.                                |
| `label-unavailable`      | The label mailbox could not be selected.                                       |
| `label-search-failed`    | The label mailbox could not be searched.                                       |
| `uidplus-unavailable`    | The server cannot perform a target-scoped UID EXPUNGE.                         |
| `sender-policy`          | The sender allowlist blocked the message and `report_blocked_mutations` is on. |

When a sender allowlist is active and `report_blocked_mutations` is off, a
blocked ID is reported as a successful no-op and nothing is deleted. See
[Sender allowlist](security.md#sender-allowlist).

### `create_label`

Creates one label by creating the mailbox that stores it. Pass `label_name`
without the `Labels/` prefix; the tool adds it.

Creating a label adds a mailbox to the account, so it requires
`enable_folder_management=true` exactly as `create_folder` does. See
[Folder management access](security.md#folder-management-access). The tool stays
visible when the policy is off and refuses at call time.

### `delete_label`

Deletes one label by deleting the mailbox that stores it, discarding the label's
own copy of every message it holds. The messages themselves are unaffected: a
label copy is a separate message from the one in its own mailbox, so deleting a
label removes the labelling, not the mail. Most servers require the mailbox to
be empty first — remove the label from its messages, or expect a `failed`
result.

Deleting a label removes a mailbox, so it requires
`enable_folder_management=true` exactly as `delete_folder` does.

### `apply_label`

Applies one label to one or more messages by copying each message from
`source_mailbox` into `Labels/<label_name>`. The message in `source_mailbox` is
never moved or modified; labelling is purely additive, which is why this tool is
not gated by `enable_folder_management`.

The label mailbox must already exist — use `list_labels` to find one or
`create_label` to make one. Applying a label to a message that already carries
it adds a second copy rather than failing, because IMAP COPY is additive.

Per-ID results preserve caller order and distinguish success, failure, and
`unknown`, and the sender allowlist applies exactly as it does to `copy_emails`:
when the allowlist is active and `report_blocked_mutations` is off, a blocked ID
is reported as a successful no-op and nothing is copied.

`create_label`, `delete_label`, and `apply_label` name the label in their
results, never the `Labels/` mailbox they derive from it, so a caller works in
label names throughout.

A label name may itself contain `/`. `list_labels` reports `Labels/Work/2026` as
the label `Work/2026`, so the write tools accept the same spelling; on a
`/`-delimited server this nests the mailbox, and elsewhere it is a flat name
containing a slash.

## Attachments

### `download_attachment`

Downloads one named attachment from a message to the server host. By default,
the server creates a safe randomized filename under the current user's
`Downloads/mcp-email-server` directory. On Windows, it uses a valid Downloads
Known Folder registry value and otherwise falls back to the profile's `~/Downloads`; on
other platforms it uses `~/Downloads`. The returned
`saved_path` reports the resolved absolute destination.

`save_path` is optional. When supplied, it remains an exact destination: use an
absolute path when possible. A relative explicit path is resolved against the
server process's working directory.

The tool is registered even when downloading is disabled, but calling it then
raises a permission error. Enable it explicitly with:

```toml
enable_attachment_download = true
```

The application checks current account and feature policy, resolves and
preflights the local destination before provider construction, credential
resolution, download, or MIME decoding. For a default destination, it sanitizes
the attachment name, removes path/device syntax, adds a cryptographically random
suffix, and creates the application subdirectory with private permissions. It
checks authority again after fetch immediately before the write, so revocation
during a slow fetch discards the payload. Raw messages above 50 MiB and decoded
attachments above 25 MiB are rejected. The mail adapter returns bytes only and
never receives the resolved path.

The artifact adapter writes only the explicit or preflight-resolved destination.
It never falls back to the process working directory when default resolution
fails. POSIX uses pinned no-follow directory descriptors and owner-only files.
Windows supports only a local fixed NTFS drive-letter path and uses held
non-reparse handles, protected
DACLs, hard-link/identity checks, `FlushFileBuffers`, and same-volume
write-through replacement. Symlinked or junction parents, linked/permissive or
non-regular targets, UNC/network/device/alternate-stream/non-NTFS paths, and
replacement races fail closed. Existing private regular files may be replaced;
there is no weaker fallback.

Review [Attachment access](security.md#attachment-access) before enabling this
operation.

## Stable tool catalog

MCP initialization reports the installed `mcp-email-server` application version
in `serverInfo.version`, not the MCP SDK dependency version. The tool list is
static for the lifetime of a server process. `send_email`,
`list_allowed_recipients`, `list_allowed_senders`, and `download_attachment` are
always advertised. Account existence, enabled state, SMTP capability, and
current policies are enforced when each tool is called. The allowlist tools have
distinct empty semantics: an empty recipient list disables sending, while an
empty sender list does not restrict reading. Each list is limited to 1,000
entries, and the complete effective-configuration snapshot
is canonically serialized against the shared 8 MiB ceiling before either policy
result is returned; oversized authority data fails with `limit_exceeded`.

Account lifecycle changes therefore do not require a tools-list notification.
A bootstrap mode selection still requires a server restart because it changes
the selected configuration authority.

## Reply threading

To preserve conversation threading:

1. Fetch the original message with `get_emails_content`.
2. Use its RFC `message_id` as `in_reply_to`.
3. Build `references` from the returned `references` value followed by the
   original `message_id`, omitting missing values.
4. Send the reply with a suitable `Re:` subject.

For a complete example, see [Reply with proper threading](guides.md#reply-with-proper-threading).
