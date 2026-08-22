from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from datetime import datetime
from importlib.metadata import version
from typing import Annotated, Literal

import anyio
from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations
from pydantic import Field

from mcp_email_server.application.accounts import AvailableAccount, EffectiveConfiguration
from mcp_email_server.application.labels import MAXIMUM_LABEL_NAME_BYTES
from mcp_email_server.application.limits import APPLICATION_LIMITS
from mcp_email_server.application.metadata import ListEmailMetadataQuery
from mcp_email_server.application.mutations import (
    MUTABLE_EMAIL_FLAGS,
    AppendMutationOutcome,
    ApplyLabelCommand,
    ArchiveCommand,
    ArchiveMutationOutcome,
    BatchMutationOutcome,
    CopyCommand,
    CreateFolderCommand,
    CreateLabelCommand,
    DeleteCommand,
    DeleteFolderCommand,
    DeleteLabelCommand,
    DeleteMutationOutcome,
    FlagOperation,
    FolderMutationOutcome,
    ForwardCommand,
    MarkReadCommand,
    MoveCommand,
    MutableEmailFlag,
    RecipientPolicyDeniedError,
    RemoveLabelCommand,
    RenameFolderCommand,
    SaveToMailboxCommand,
    SendCommand,
    SendMutationOutcome,
    SetEmailFlagsCommand,
    TargetMutationOutcome,
)
from mcp_email_server.application.reads import (
    DownloadAttachmentCommand,
    GetEmailContentQuery,
    GetEmailLabelsQuery,
    ListLabelsQuery,
    ListMailboxesQuery,
)
from mcp_email_server.emails.models import (
    AttachmentDownloadResponse,
    EmailContentBatchResponse,
    EmailMetadataPageResponse,
    LabelInfo,
    MailboxInfo,
)
from mcp_email_server.runtime import close_application_runtime, get_application_runtime

MAX_IMAP_UID_CHARACTERS = len(str(APPLICATION_LIMITS.maximum_imap_uid))
UidInput = Annotated[str, Field(max_length=MAX_IMAP_UID_CHARACTERS, pattern=r"^[1-9][0-9]*$")]
AddressInput = Annotated[str, Field(max_length=APPLICATION_LIMITS.address_bytes)]
AttachmentPathInput = Annotated[str, Field(max_length=APPLICATION_LIMITS.attachment_path_bytes)]
FlagInput = Annotated[str, Field(max_length=APPLICATION_LIMITS.flag_bytes)]
AccountDiscoveryResult = Annotated[
    list[AvailableAccount],
    Field(max_length=APPLICATION_LIMITS.configured_accounts),
]
PolicyDiscoveryResult = Annotated[
    list[str],
    Field(max_length=APPLICATION_LIMITS.policy_entries),
]


def list_effective_accounts() -> list[AvailableAccount]:
    return get_application_runtime().accounts.discover()


async def list_email_metadata(query: ListEmailMetadataQuery) -> EmailMetadataPageResponse:
    return await get_application_runtime().metadata.execute(query)


async def send_email_command(command: SendCommand) -> SendMutationOutcome:
    return await get_application_runtime().mutations.send.execute(command)


async def forward_email_command(command: ForwardCommand) -> SendMutationOutcome:
    return await get_application_runtime().mutations.forward.execute(command)


async def save_to_mailbox_command(command: SaveToMailboxCommand) -> AppendMutationOutcome:
    return await get_application_runtime().mutations.save_to_mailbox.execute(command)


async def delete_emails_command(command: DeleteCommand) -> DeleteMutationOutcome:
    return await get_application_runtime().mutations.delete.execute(command)


async def set_email_flags_command(command: SetEmailFlagsCommand) -> BatchMutationOutcome:
    return await get_application_runtime().mutations.set_flags.execute(command)


async def mark_read_command(command: MarkReadCommand) -> BatchMutationOutcome:
    return await get_application_runtime().mutations.mark_read.execute(command)


async def move_emails_command(command: MoveCommand) -> BatchMutationOutcome:
    return await get_application_runtime().mutations.move.execute(command)


async def archive_emails_command(command: ArchiveCommand) -> ArchiveMutationOutcome:
    return await get_application_runtime().mutations.archive.execute(command)


async def get_email_content_query(query: GetEmailContentQuery) -> EmailContentBatchResponse:
    return await get_application_runtime().reads.content.execute(query)


async def list_mailboxes_query(query: ListMailboxesQuery) -> list[MailboxInfo]:
    return await get_application_runtime().reads.mailboxes.execute(query)


async def download_attachment_command(command: DownloadAttachmentCommand) -> AttachmentDownloadResponse:
    return await get_application_runtime().reads.attachments.execute(command)


def effective_configuration() -> EffectiveConfiguration:
    return get_application_runtime().configuration.execute()


_PUBLIC_SEND_DETAILS = frozenset({
    "not-attempted",
    "provider-timeout",
    "smtp-cancelled-before-data",
    "smtp-data-rejected",
    "smtp-data-unknown",
    "smtp-mail-cancelled",
    "smtp-mail-rejected",
    "smtp-mail-unavailable",
    "smtp-recipient-rejected",
    "smtp-session-lost-before-data",
    "smtp-8bitmime-required",
    "smtp-utf8-unsupported",
})
_PUBLIC_APPEND_DETAILS = frozenset({
    "append-unknown",
    "provider-timeout",
    "utf8-append-unsupported",
})


def _ordered_target_sections(
    outcomes: tuple[TargetMutationOutcome, ...],
    *,
    detail_allowlist: frozenset[str] | None = None,
    include_failed_detail: bool = False,
    include_unknown_detail: bool,
) -> list[str]:
    """Format contiguous statuses without reordering input-aligned outcomes."""
    sections: list[str] = []
    current_status: str | None = None
    current_targets: list[str] = []
    for item in outcomes:
        if item.status != current_status and current_targets:
            sections.append(f"{current_status}: {', '.join(current_targets)}")
            current_targets = []
        current_status = item.status
        target = item.target
        include_detail = (include_failed_detail and item.status == "failed") or (
            include_unknown_detail and item.status == "unknown"
        )
        detail_is_allowed = detail_allowlist is None or item.detail in detail_allowlist
        if include_detail and item.detail is not None and detail_is_allowed:
            target = f"{target} ({item.detail})"
        current_targets.append(target)
    if current_targets:
        sections.append(f"{current_status}: {', '.join(current_targets)}")
    return sections


def _tagged_batch_result(outcome: BatchMutationOutcome) -> str:
    sections = _ordered_target_sections(outcome.outcomes, include_unknown_detail=True)
    if outcome.reconciliation_needed:
        sections.append("warning: reconciliation needed")
    return "; ".join(sections)


def _tagged_send_result(outcome: SendMutationOutcome) -> str:
    sections = _ordered_target_sections(
        outcome.delivery,
        detail_allowlist=_PUBLIC_SEND_DETAILS,
        include_failed_detail=True,
        include_unknown_detail=True,
    )
    sent_copy = outcome.sent_copy.status
    sent_copy_context: list[str] = []
    if outcome.sent_copy.mailbox:
        sent_copy_context.append(outcome.sent_copy.mailbox)
    if outcome.sent_copy.detail in _PUBLIC_APPEND_DETAILS:
        sent_copy_context.append(outcome.sent_copy.detail)
    if sent_copy_context:
        sent_copy = f"{sent_copy} ({'; '.join(sent_copy_context)})"
    sections.append(f"sent-copy: {sent_copy}")
    if outcome.reconciliation_needed:
        sections.append("warning: reconciliation needed")
    return "; ".join(sections)


@asynccontextmanager
async def _application_lifespan(_server: FastMCP) -> AsyncIterator[dict[str, object]]:
    try:
        yield {}
    finally:
        with anyio.CancelScope(shield=True):
            await close_application_runtime()


MCP_SERVER_INSTRUCTIONS = (
    "When sending emails, the body supports Markdown formatting (bold, lists, headers, links, etc.) "
    "which is automatically converted to email-safe HTML. Use Markdown freely for well-formatted emails. "
    "Set html=True only if providing pre-formatted raw HTML."
)

mcp = FastMCP("email", instructions=MCP_SERVER_INSTRUCTIONS, lifespan=_application_lifespan)
# FastMCP 1.x does not expose its low-level server version in the constructor.
mcp._mcp_server.version = version("mcp-email-server")  # pyright: ignore[reportPrivateUsage]

_READ_ONLY_LOCAL = ToolAnnotations(
    readOnlyHint=True,
    destructiveHint=False,
    idempotentHint=True,
    openWorldHint=False,
)
_READ_ONLY_REMOTE = ToolAnnotations(
    readOnlyHint=True,
    destructiveHint=False,
    idempotentHint=True,
    openWorldHint=True,
)
_NONDESTRUCTIVE_REMOTE_MUTATION = ToolAnnotations(
    readOnlyHint=False,
    destructiveHint=False,
    idempotentHint=False,
    openWorldHint=True,
)
_IDEMPOTENT_REMOTE_MUTATION = ToolAnnotations(
    readOnlyHint=False,
    destructiveHint=False,
    idempotentHint=True,
    openWorldHint=True,
)
_DESTRUCTIVE_REMOTE_MUTATION = ToolAnnotations(
    readOnlyHint=False,
    destructiveHint=True,
    idempotentHint=False,
    openWorldHint=True,
)
_FILESYSTEM_WRITE = ToolAnnotations(
    readOnlyHint=False,
    destructiveHint=True,
    idempotentHint=False,
    openWorldHint=True,
)


@mcp.resource(
    "email://{account_name}",
    description="Return stable non-secret discovery capabilities for one configured account.",
)
async def get_account(account_name: str) -> AvailableAccount | None:
    return get_application_runtime().accounts.discover_one(account_name)


@mcp.tool(
    description=(
        "List configured accounts as stable non-secret capability records. Use only accounts with "
        "can_receive=true for mail reads and can_send=true for send_email. If the result is empty, ask the "
        "user to run `mcp-email-server ui` or the user-operated CLI; never ask for credentials in chat."
    ),
    annotations=_READ_ONLY_LOCAL,
)
async def list_available_accounts() -> AccountDiscoveryResult:
    return list_effective_accounts()


@mcp.tool(
    description="List email metadata (email_id, subject, sender, recipients, date) without body content. Returns email_id for use with get_emails_content.",
    annotations=_READ_ONLY_REMOTE,
)
async def list_emails_metadata(
    account_name: Annotated[
        str, Field(max_length=APPLICATION_LIMITS.account_name_bytes, description="The name of the email account.")
    ],
    page: Annotated[
        int,
        Field(default=1, ge=1, description="The page number to retrieve (starting from 1)."),
    ] = 1,
    page_size: Annotated[
        int,
        Field(default=10, ge=1, le=100, description="The number of emails to retrieve per page."),
    ] = 10,
    before: Annotated[
        datetime | None,
        Field(default=None, description="Retrieve emails before this datetime (UTC)."),
    ] = None,
    since: Annotated[
        datetime | None,
        Field(default=None, description="Retrieve emails since this datetime (UTC)."),
    ] = None,
    subject: Annotated[
        str | None,
        Field(
            default=None,
            max_length=APPLICATION_LIMITS.query_bytes,
            description="Filter emails by subject.",
        ),
    ] = None,
    from_address: Annotated[
        str | None,
        Field(
            default=None,
            max_length=APPLICATION_LIMITS.address_bytes,
            description="Filter emails by sender address.",
        ),
    ] = None,
    to_address: Annotated[
        str | None,
        Field(
            default=None,
            max_length=APPLICATION_LIMITS.address_bytes,
            description="Filter emails by recipient address.",
        ),
    ] = None,
    order: Annotated[
        Literal["asc", "desc"],
        Field(default=None, description="Sort matching emails by date: oldest first (`asc`) or newest first (`desc`)."),
    ] = "desc",
    mailbox: Annotated[
        str,
        Field(
            default="INBOX",
            max_length=APPLICATION_LIMITS.mailbox_bytes,
            description="The mailbox to search.",
        ),
    ] = "INBOX",
    seen: Annotated[
        bool | None,
        Field(default=None, description="Filter by read status: True=read, False=unread, None=all."),
    ] = None,
    flagged: Annotated[
        bool | None,
        Field(default=None, description="Filter by flagged/starred status: True=flagged, False=unflagged, None=all."),
    ] = None,
    answered: Annotated[
        bool | None,
        Field(default=None, description="Filter by replied status: True=replied, False=not replied, None=all."),
    ] = None,
    body: Annotated[
        str | None,
        Field(
            default=None,
            max_length=APPLICATION_LIMITS.query_bytes,
            description="Search for text in the email body (IMAP BODY).",
        ),
    ] = None,
    text: Annotated[
        str | None,
        Field(
            default=None,
            max_length=APPLICATION_LIMITS.query_bytes,
            description="Search for text in the entire message — headers and body (IMAP TEXT).",
        ),
    ] = None,
    has_attachment: Annotated[
        bool | None,
        Field(
            default=None,
            description="Filter by attachment presence: True=has attachment, False=none, None=all "
            "(multipart/mixed heuristic; may miss inline images or yield false positives).",
        ),
    ] = None,
) -> EmailMetadataPageResponse:
    return await list_email_metadata(
        ListEmailMetadataQuery(
            account_name=account_name,
            page=page,
            page_size=page_size,
            before=before,
            since=since,
            subject=subject,
            from_address=from_address,
            to_address=to_address,
            order=order,
            mailbox=mailbox,
            seen=seen,
            flagged=flagged,
            answered=answered,
            body=body,
            text=text,
            has_attachment=has_attachment,
        )
    )


@mcp.tool(
    description=(
        "Get the full content (including body and reply-thread headers) of one or more emails by their email_id. "
        "Use list_emails_metadata first. This tool is non-read-only because mark_as_read=true changes remote flags."
    ),
    annotations=_IDEMPOTENT_REMOTE_MUTATION,
)
async def get_emails_content(
    account_name: Annotated[
        str, Field(max_length=APPLICATION_LIMITS.account_name_bytes, description="The name of the email account.")
    ],
    email_ids: Annotated[
        list[UidInput],
        Field(
            min_length=1,
            max_length=APPLICATION_LIMITS.content_email_ids,
            description="One or more email_id values to retrieve, supplied as an array (obtained from list_emails_metadata).",
        ),
    ],
    mailbox: Annotated[
        str,
        Field(
            default="INBOX",
            max_length=APPLICATION_LIMITS.mailbox_bytes,
            description="The mailbox to retrieve emails from.",
        ),
    ] = "INBOX",
    mark_as_read: Annotated[
        bool,
        Field(
            default=False,
            description="If True, mark each successfully retrieved email as read. If marking fails, a warning is logged and retrieval still succeeds.",
        ),
    ] = False,
    body_offset: Annotated[
        int,
        Field(
            default=0,
            ge=0,
            description="Character offset into each email body to start reading from. Use together "
            "with max_body_length to page through long emails: if a returned body ends with the "
            "'...[TRUNCATED]' marker, fetch the next chunk with body_offset += max_body_length.",
        ),
    ] = 0,
    max_body_length: Annotated[
        int | None,
        Field(
            default=20000,
            ge=0,
            le=100000,
            description="Maximum number of body characters to return, counted from body_offset. "
            "If the body extends past this window, the '...[TRUNCATED]' marker is appended after "
            "the requested body window. Use 0 or null to return the whole body from body_offset "
            "onward with no truncation and no marker; explicit limits are 1 to 100000. An "
            "untruncated body is still subject to the shared per-message and aggregate body byte "
            "ceilings, which return a bounded limit error instead of a partial body.",
        ),
    ] = 20000,
) -> EmailContentBatchResponse:
    return await get_email_content_query(
        GetEmailContentQuery(
            account_name=account_name,
            email_ids=tuple(email_ids),
            mailbox=mailbox,
            mark_as_read=mark_as_read,
            body_offset=body_offset,
            max_body_length=max_body_length,
        )
    )


@mcp.tool(
    description=(
        "List the configured recipient allowlist — the addresses that send_email is permitted to "
        "send to and save_to_mailbox is permitted to address. Returns an empty list when unrestricted."
    ),
    annotations=_READ_ONLY_LOCAL,
)
async def list_allowed_recipients() -> PolicyDiscoveryResult:
    return list(effective_configuration().allowed_recipients)


@mcp.tool(
    description=(
        "List the configured inbound sender allowlist — the address patterns whose mail the server "
        "will read or act on. When configured, only these senders' mail is visible to the read tools "
        "(list_emails_metadata, get_emails_content, download_attachment) and eligible for the mutation "
        "tools (delete_emails, set_email_flags, mark_emails_as_read, move_emails, archive_emails, "
        "copy_emails, apply_label, remove_label). Returns an "
        "empty list "
        "when unrestricted."
    ),
    annotations=_READ_ONLY_LOCAL,
)
async def list_allowed_senders() -> PolicyDiscoveryResult:
    return list(effective_configuration().allowed_senders)


@mcp.tool(
    description=(
        "Send one email using the specified account. Supports reply threading. The body is written in Markdown "
        "and is rendered to email-safe HTML automatically; set html=true only when the body is already "
        "pre-formatted raw HTML. When in_reply_to is set, the original message is read over IMAP and appended "
        "as a collapsible quote block; if it cannot be read the call fails before any SMTP session is opened "
        "rather than sending an unquoted reply. Partial or ambiguous SMTP "
        "delivery reports per-recipient succeeded/failed/unknown status and reports the independent Sent-copy "
        "outcome separately; ambiguous effects are not retried automatically."
    ),
    annotations=_NONDESTRUCTIVE_REMOTE_MUTATION,
)
async def send_email(
    account_name: Annotated[
        str,
        Field(
            max_length=APPLICATION_LIMITS.account_name_bytes,
            description="The name of the email account to send from.",
        ),
    ],
    recipients: Annotated[
        list[AddressInput],
        Field(
            min_length=1, max_length=APPLICATION_LIMITS.recipients, description="A list of recipient email addresses."
        ),
    ],
    subject: Annotated[
        str,
        Field(max_length=APPLICATION_LIMITS.subject_bytes, description="The subject of the email."),
    ],
    body: Annotated[
        str,
        Field(
            max_length=APPLICATION_LIMITS.body_bytes,
            description="The body of the email, written in Markdown. It is rendered to email-safe HTML automatically.",
        ),
    ],
    cc: Annotated[
        list[AddressInput] | None,
        Field(default=None, max_length=APPLICATION_LIMITS.recipients, description="A list of CC email addresses."),
    ] = None,
    bcc: Annotated[
        list[AddressInput] | None,
        Field(default=None, max_length=APPLICATION_LIMITS.recipients, description="A list of BCC email addresses."),
    ] = None,
    html: Annotated[
        bool,
        Field(
            default=False,
            description="Set True only when body is already pre-formatted raw HTML, which suppresses Markdown rendering.",
        ),
    ] = False,
    attachments: Annotated[
        list[AttachmentPathInput] | None,
        Field(
            default=None,
            max_length=APPLICATION_LIMITS.attachments,
            description="A list of file paths to attach. Relative paths are resolved against the server process working directory; absolute paths are recommended.",
        ),
    ] = None,
    in_reply_to: Annotated[
        str | None,
        Field(
            default=None,
            max_length=APPLICATION_LIMITS.header_bytes,
            description="Message-ID of the email being replied to. Enables proper threading in email clients.",
        ),
    ] = None,
    references: Annotated[
        str | None,
        Field(
            default=None,
            max_length=APPLICATION_LIMITS.header_bytes,
            description="Space-separated Message-IDs for the thread chain. Usually includes in_reply_to plus ancestors.",
        ),
    ] = None,
    reply_to: Annotated[
        str | None,
        Field(
            default=None,
            max_length=APPLICATION_LIMITS.header_bytes,
            description="Email address to set as the Reply-To header. When set, email clients will reply to this address instead of the From address.",
        ),
    ] = None,
    quote_reply: Annotated[
        bool,
        Field(
            default=True,
            description="When in_reply_to is set, append the original message as a collapsible quote block. Set False to reply without quoting.",
        ),
    ] = True,
) -> str:
    try:
        outcome = await send_email_command(
            SendCommand(
                account_name=account_name,
                recipients=tuple(recipients),
                subject=subject,
                body=body,
                cc=tuple(cc or ()),
                bcc=tuple(bcc or ()),
                html=html,
                attachments=tuple(attachments or ()),
                in_reply_to=in_reply_to,
                references=references,
                reply_to=reply_to,
                quote_reply=quote_reply,
            )
        )
    except RecipientPolicyDeniedError as exc:
        raise ValueError("Recipient(s) not in allowlist") from exc
    if (
        all(item.status == "succeeded" for item in outcome.delivery)
        and outcome.sent_copy.status in ("succeeded", "skipped")
        and not outcome.reconciliation_needed
    ):
        recipient_str = ", ".join(recipients)
        attachment_info = f" with {len(attachments)} attachment(s)" if attachments else ""
        return f"Email sent successfully to {recipient_str}{attachment_info}"
    return f"Email delivery [{_tagged_send_result(outcome)}]"


@mcp.tool(
    description=(
        "Forward an existing message to new recipients using the specified account. The source message is read "
        "over IMAP first: if it cannot be read, the call fails before any SMTP session is opened, so a forward is "
        "never delivered without the content it was supposed to carry. The subject is derived from the source as "
        "'Fwd: <original subject>' without stacking a second prefix, the caller's Markdown note is placed above a "
        "forwarded block re-composed from the source's parsed text body and escaped so quoted content is delivered "
        "literally, and the source's attachments "
        "are re-attached with their original MIME types unless include_attachments is false. Partial or ambiguous "
        "SMTP delivery reports per-recipient succeeded/failed/unknown status and reports the independent "
        "Sent-copy outcome separately; ambiguous effects are not retried automatically."
    ),
    annotations=_NONDESTRUCTIVE_REMOTE_MUTATION,
)
async def forward_email(
    account_name: Annotated[
        str,
        Field(
            max_length=APPLICATION_LIMITS.account_name_bytes,
            description="The name of the email account to forward from.",
        ),
    ],
    email_id: Annotated[UidInput, Field(description="UID of the source message to forward.")],
    recipients: Annotated[
        list[AddressInput],
        Field(
            min_length=1,
            max_length=APPLICATION_LIMITS.recipients,
            description="A list of addresses that receive the forwarded message.",
        ),
    ],
    source_mailbox: Annotated[
        str,
        Field(
            default="INBOX",
            max_length=APPLICATION_LIMITS.mailbox_bytes,
            description="The mailbox that contains the source message.",
        ),
    ] = "INBOX",
    body: Annotated[
        str,
        Field(
            default="",
            max_length=APPLICATION_LIMITS.body_bytes,
            description="An optional Markdown note placed above the forwarded content.",
        ),
    ] = "",
    cc: Annotated[
        list[AddressInput] | None,
        Field(default=None, max_length=APPLICATION_LIMITS.recipients, description="A list of CC email addresses."),
    ] = None,
    bcc: Annotated[
        list[AddressInput] | None,
        Field(default=None, max_length=APPLICATION_LIMITS.recipients, description="A list of BCC email addresses."),
    ] = None,
    include_attachments: Annotated[
        bool,
        Field(default=True, description="Whether to re-attach the source message's attachments."),
    ] = True,
) -> str:
    try:
        outcome = await forward_email_command(
            ForwardCommand(
                account_name=account_name,
                recipients=tuple(recipients),
                subject="",
                body=body,
                cc=tuple(cc or ()),
                bcc=tuple(bcc or ()),
                source_email_id=email_id,
                source_mailbox=source_mailbox,
                include_attachments=include_attachments,
            )
        )
    except RecipientPolicyDeniedError as exc:
        raise ValueError("Recipient(s) not in allowlist") from exc
    if (
        all(item.status == "succeeded" for item in outcome.delivery)
        and outcome.sent_copy.status in ("succeeded", "skipped")
        and not outcome.reconciliation_needed
    ):
        return f"Email forwarded successfully to {', '.join(recipients)}"
    return f"Email forward [{_tagged_send_result(outcome)}]"


@mcp.tool(
    description="Compose an email and save it to an IMAP folder (e.g., Drafts). "
    "Shares recipient, body, attachment, and threading parameters with send_email; "
    "adds mailbox and flags, and does not support reply_to. "
    "Default folder is Drafts with \\Draft and \\Seen flags. "
    "Pure IMAP operation — works without SMTP configuration. An ambiguous APPEND is reported as unknown "
    "and is not retried automatically.",
    annotations=_NONDESTRUCTIVE_REMOTE_MUTATION,
)
async def save_to_mailbox(
    account_name: Annotated[
        str, Field(max_length=APPLICATION_LIMITS.account_name_bytes, description="The name of the email account.")
    ],
    recipients: Annotated[
        list[AddressInput],
        Field(
            min_length=1, max_length=APPLICATION_LIMITS.recipients, description="A list of recipient email addresses."
        ),
    ],
    subject: Annotated[
        str,
        Field(max_length=APPLICATION_LIMITS.subject_bytes, description="The subject of the email."),
    ],
    body: Annotated[
        str,
        Field(max_length=APPLICATION_LIMITS.body_bytes, description="The body of the email."),
    ],
    mailbox: Annotated[
        str,
        Field(
            default="Drafts",
            max_length=APPLICATION_LIMITS.mailbox_bytes,
            description="The IMAP folder to save to (e.g., 'Drafts', 'INBOX.Drafts', 'Templates').",
        ),
    ] = "Drafts",
    cc: Annotated[
        list[AddressInput] | None,
        Field(default=None, max_length=APPLICATION_LIMITS.recipients, description="A list of CC email addresses."),
    ] = None,
    bcc: Annotated[
        list[AddressInput] | None,
        Field(default=None, max_length=APPLICATION_LIMITS.recipients, description="A list of BCC email addresses."),
    ] = None,
    html: Annotated[
        bool,
        Field(default=False, description="Whether the email body is HTML (True) or plain text (False)."),
    ] = False,
    attachments: Annotated[
        list[AttachmentPathInput] | None,
        Field(
            default=None,
            max_length=APPLICATION_LIMITS.attachments,
            description="A list of file paths to attach. Relative paths are resolved against the server process working directory; absolute paths are recommended.",
        ),
    ] = None,
    in_reply_to: Annotated[
        str | None,
        Field(
            default=None,
            max_length=APPLICATION_LIMITS.header_bytes,
            description="Message-ID of the email being replied to. Enables proper threading in email clients.",
        ),
    ] = None,
    references: Annotated[
        str | None,
        Field(
            default=None,
            max_length=APPLICATION_LIMITS.header_bytes,
            description="Space-separated Message-IDs for the thread chain.",
        ),
    ] = None,
    flags: Annotated[
        list[FlagInput] | None,
        Field(
            default=None,
            max_length=APPLICATION_LIMITS.flags,
            description=r"IMAP flags to set on the message. Defaults to ['\Draft', '\Seen']. Common flags: '\Draft', '\Seen', '\Flagged'.",
        ),
    ] = None,
) -> str:
    try:
        outcome = await save_to_mailbox_command(
            SaveToMailboxCommand(
                account_name=account_name,
                recipients=tuple(recipients),
                subject=subject,
                body=body,
                mailbox=mailbox,
                cc=tuple(cc or ()),
                bcc=tuple(bcc or ()),
                html=html,
                attachments=tuple(attachments or ()),
                in_reply_to=in_reply_to,
                references=references,
                flags=tuple(flags) if flags is not None else None,
            )
        )
    except RecipientPolicyDeniedError as exc:
        raise ValueError("Recipient(s) not in allowlist") from exc
    if outcome.status == "succeeded" and not outcome.reconciliation_needed:
        email_id = outcome.uid or "unknown"
        return f"Email saved to '{mailbox}' successfully. Message-Id: {outcome.message_id}, email_id: {email_id}"
    detail = f" ({outcome.detail})" if outcome.detail in _PUBLIC_APPEND_DETAILS else ""
    warning = "; warning: reconciliation needed" if outcome.reconciliation_needed else ""
    return f"Email save [{outcome.status}{detail}: {mailbox}; Message-Id: {outcome.message_id}{warning}]"


@mcp.tool(
    description=(
        "Delete one or more emails by email_id. Deletion is recoverable where the account allows it: the emails "
        "are moved to the Trash mailbox, auto-detected via the RFC 6154 \\Trash flag (falling back to common names "
        "like Trash, Deleted Items, or [Gmail]/Trash). When the account has no Trash mailbox, and when the emails "
        "are already in Trash, they are instead removed permanently using target-scoped UID EXPUNGE. The result "
        "says which happened. Use list_emails_metadata first. Partial or ambiguous effects report per-ID "
        "succeeded/failed/unknown status and are not retried automatically."
    ),
    annotations=_DESTRUCTIVE_REMOTE_MUTATION,
)
async def delete_emails(
    account_name: Annotated[
        str, Field(max_length=APPLICATION_LIMITS.account_name_bytes, description="The name of the email account.")
    ],
    email_ids: Annotated[
        list[UidInput],
        Field(
            min_length=1,
            max_length=APPLICATION_LIMITS.mutation_uids,
            description="List of email_id to delete (obtained from list_emails_metadata).",
        ),
    ],
    mailbox: Annotated[
        str,
        Field(
            default="INBOX",
            max_length=APPLICATION_LIMITS.mailbox_bytes,
            description="The mailbox to delete emails from.",
        ),
    ] = "INBOX",
) -> str:
    outcome = await delete_emails_command(DeleteCommand(account_name, tuple(email_ids), mailbox))
    succeeded = outcome.batch.targets("succeeded")
    complete = len(succeeded) == len(email_ids) and not outcome.batch.reconciliation_needed
    if outcome.trash_mailbox is None:
        if complete:
            return f"Successfully deleted {len(succeeded)} email(s) permanently"
        return f"Delete result [{_tagged_batch_result(outcome.batch)}; permanent]"
    if complete:
        return f"Successfully deleted {len(succeeded)} email(s) by moving them to {outcome.trash_mailbox}"
    return f"Delete result [{_tagged_batch_result(outcome.batch)}; mailbox: {outcome.trash_mailbox}]"


@mcp.tool(
    description=(
        "Add or remove approved IMAP flags on one or more emails by email_id. Supported flags are \\Seen, "
        "\\Flagged, \\Answered, and \\Draft; \\Deleted and provider-specific keywords are not supported. "
        "Use list_emails_metadata first. Partial or ambiguous effects report per-ID succeeded/failed/unknown "
        "status and are not retried automatically."
    ),
    annotations=_IDEMPOTENT_REMOTE_MUTATION,
)
async def set_email_flags(
    account_name: Annotated[
        str, Field(max_length=APPLICATION_LIMITS.account_name_bytes, description="The name of the email account.")
    ],
    email_ids: Annotated[
        list[UidInput],
        Field(
            min_length=1,
            max_length=APPLICATION_LIMITS.mutation_uids,
            description="List of email_id values whose flags should be changed.",
        ),
    ],
    operation: Annotated[
        FlagOperation,
        Field(description="Whether to add or remove every supplied flag."),
    ],
    flags: Annotated[
        list[MutableEmailFlag],
        Field(
            min_length=1,
            max_length=len(MUTABLE_EMAIL_FLAGS),
            description="Unique approved flags to add or remove: \\Seen, \\Flagged, \\Answered, or \\Draft.",
        ),
    ],
    mailbox: Annotated[
        str,
        Field(
            default="INBOX",
            max_length=APPLICATION_LIMITS.mailbox_bytes,
            description="The mailbox containing the emails.",
        ),
    ] = "INBOX",
) -> str:
    outcome = await set_email_flags_command(
        SetEmailFlagsCommand(account_name, tuple(email_ids), operation, tuple(flags), mailbox)
    )
    succeeded = outcome.targets("succeeded")
    if len(succeeded) == len(email_ids) and not outcome.reconciliation_needed:
        verb, preposition = ("added", "to") if operation == "add" else ("removed", "from")
        return f"Successfully {verb} {', '.join(flags)} {preposition} {len(succeeded)} email(s)"
    return f"Set-flags result [{_tagged_batch_result(outcome)}]"


@mcp.tool(
    description=(
        "Mark one or more emails as read by email_id. This is the common-workflow equivalent of adding \\Seen "
        "with set_email_flags. Use list_emails_metadata first. Partial or ambiguous effects report per-ID "
        "succeeded/failed/unknown status and are not retried automatically."
    ),
    annotations=_IDEMPOTENT_REMOTE_MUTATION,
)
async def mark_emails_as_read(
    account_name: Annotated[
        str, Field(max_length=APPLICATION_LIMITS.account_name_bytes, description="The name of the email account.")
    ],
    email_ids: Annotated[
        list[UidInput],
        Field(
            min_length=1,
            max_length=APPLICATION_LIMITS.mutation_uids,
            description="List of email_id to mark as read (obtained from list_emails_metadata).",
        ),
    ],
    mailbox: Annotated[
        str,
        Field(
            default="INBOX",
            max_length=APPLICATION_LIMITS.mailbox_bytes,
            description="The mailbox containing the emails.",
        ),
    ] = "INBOX",
) -> str:
    outcome = await mark_read_command(MarkReadCommand(account_name, tuple(email_ids), mailbox))
    succeeded = outcome.targets("succeeded")
    if len(succeeded) == len(email_ids) and not outcome.reconciliation_needed:
        return f"Successfully marked {len(succeeded)} email(s) as read"
    return f"Mark-read result [{_tagged_batch_result(outcome)}]"


@mcp.tool(
    description=(
        "Move one or more emails between IMAP folders by email_id. Use list_emails_metadata and list_mailboxes "
        "first. Partial or ambiguous effects report per-ID succeeded/failed/unknown status and are not retried."
    ),
    annotations=_DESTRUCTIVE_REMOTE_MUTATION,
)
async def move_emails(
    account_name: Annotated[
        str, Field(max_length=APPLICATION_LIMITS.account_name_bytes, description="The name of the email account.")
    ],
    email_ids: Annotated[
        list[UidInput],
        Field(
            min_length=1,
            max_length=APPLICATION_LIMITS.mutation_uids,
            description="List of email_id to move (obtained from list_emails_metadata).",
        ),
    ],
    destination_mailbox: Annotated[
        str,
        Field(
            max_length=APPLICATION_LIMITS.mailbox_bytes,
            description="The destination mailbox/folder to move emails to.",
        ),
    ],
    source_mailbox: Annotated[
        str,
        Field(
            default="INBOX",
            max_length=APPLICATION_LIMITS.mailbox_bytes,
            description="The source mailbox containing the emails.",
        ),
    ] = "INBOX",
) -> str:
    outcome = await move_emails_command(
        MoveCommand(account_name, tuple(email_ids), source_mailbox, destination_mailbox)
    )
    succeeded = outcome.targets("succeeded")
    if len(succeeded) == len(email_ids) and not outcome.reconciliation_needed:
        return f"Successfully moved {len(succeeded)} email(s) to {destination_mailbox}"
    return f"Move result [{_tagged_batch_result(outcome)}]"


@mcp.tool(
    description="Archive one or more emails by moving them to the account's Archive folder, "
    "auto-detected via the RFC 6154 \\Archive flag (falling back to common names like Archive or "
    "[Gmail]/All Mail). Use list_emails_metadata first. Partial or ambiguous effects report per-ID "
    "succeeded/failed/unknown status and are not retried automatically.",
    annotations=_DESTRUCTIVE_REMOTE_MUTATION,
)
async def archive_emails(
    account_name: Annotated[
        str, Field(max_length=APPLICATION_LIMITS.account_name_bytes, description="The name of the email account.")
    ],
    email_ids: Annotated[
        list[UidInput],
        Field(
            min_length=1,
            max_length=APPLICATION_LIMITS.mutation_uids,
            description="List of email_id to archive (obtained from list_emails_metadata).",
        ),
    ],
    mailbox: Annotated[
        str,
        Field(
            default="INBOX",
            max_length=APPLICATION_LIMITS.mailbox_bytes,
            description="The source mailbox containing the emails.",
        ),
    ] = "INBOX",
) -> str:
    outcome = await archive_emails_command(ArchiveCommand(account_name, tuple(email_ids), mailbox))
    succeeded = outcome.batch.targets("succeeded")
    if len(succeeded) == len(email_ids) and not outcome.batch.reconciliation_needed:
        return f"Successfully archived {len(succeeded)} email(s) to {outcome.archive_mailbox}"
    return f"Archive result [{_tagged_batch_result(outcome.batch)}; mailbox: {outcome.archive_mailbox}]"


@mcp.tool(
    description="List available mailboxes/folders for an email account. Returns folder names, hierarchy delimiters, and flags. Useful for discovering folder names before moving emails.",
    annotations=_READ_ONLY_REMOTE,
)
async def list_mailboxes(
    account_name: Annotated[
        str, Field(max_length=APPLICATION_LIMITS.account_name_bytes, description="The name of the email account.")
    ],
    pattern: Annotated[
        str,
        Field(
            default="*",
            max_length=APPLICATION_LIMITS.mailbox_bytes,
            description="IMAP LIST pattern. Use '*' for all folders, 'INBOX.*' for INBOX children.",
        ),
    ] = "*",
    reference: Annotated[
        str,
        Field(
            default="",
            max_length=APPLICATION_LIMITS.mailbox_bytes,
            description="IMAP LIST reference name (namespace prefix). Usually empty.",
        ),
    ] = "",
) -> list[MailboxInfo]:
    return await list_mailboxes_query(
        ListMailboxesQuery(account_name=account_name, pattern=pattern, reference=reference)
    )


@mcp.tool(
    description="Download an email attachment. By default it is saved with a safe randomized name under the current user's Downloads/mcp-email-server directory; an explicit destination path remains supported. This feature must be explicitly enabled in settings (enable_attachment_download=true) due to security considerations.",
    annotations=_FILESYSTEM_WRITE,
)
async def download_attachment(
    account_name: Annotated[
        str, Field(max_length=APPLICATION_LIMITS.account_name_bytes, description="The name of the email account.")
    ],
    email_id: Annotated[
        str,
        Field(
            max_length=MAX_IMAP_UID_CHARACTERS,
            description="The email ID (obtained from list_emails_metadata or get_emails_content).",
        ),
    ],
    attachment_name: Annotated[
        str,
        Field(
            max_length=APPLICATION_LIMITS.attachment_path_bytes,
            description="The name of the attachment to download (as shown in the attachments list).",
        ),
    ],
    save_path: Annotated[
        str | None,
        Field(
            max_length=APPLICATION_LIMITS.attachment_path_bytes,
            description="Optional exact destination path. Omit it to use a safe randomized filename under the current user's Downloads/mcp-email-server directory. Relative explicit paths are resolved against the server process working directory.",
        ),
    ] = None,
    mailbox: Annotated[
        str,
        Field(
            max_length=APPLICATION_LIMITS.mailbox_bytes,
            description="The mailbox to search in (default: INBOX).",
        ),
    ] = "INBOX",
) -> AttachmentDownloadResponse:
    return await download_attachment_command(
        DownloadAttachmentCommand(
            account_name=account_name,
            email_id=email_id,
            attachment_name=attachment_name,
            save_path=save_path,
            mailbox=mailbox,
        )
    )


# Port slots for tools being ported onto this architecture in parallel. Each
# cluster registers its @mcp.tool functions directly ABOVE its own marker, so two
# clusters extending this tail produce non-overlapping hunks instead of a
# conflict. Slot order is fixed everywhere it appears: A, B1, C, B2.
_PUBLIC_FOLDER_DETAILS = frozenset({
    "create-rejected",
    "create-unknown",
    "delete-rejected",
    "delete-unknown",
    "provider-timeout",
    "rename-rejected",
    "rename-unknown",
})


async def copy_emails_command(command: CopyCommand) -> BatchMutationOutcome:
    return await get_application_runtime().mutations.copy.execute(command)


async def create_folder_command(command: CreateFolderCommand) -> FolderMutationOutcome:
    return await get_application_runtime().mutations.create_folder.execute(command)


async def delete_folder_command(command: DeleteFolderCommand) -> FolderMutationOutcome:
    return await get_application_runtime().mutations.delete_folder.execute(command)


async def rename_folder_command(command: RenameFolderCommand) -> FolderMutationOutcome:
    return await get_application_runtime().mutations.rename_folder.execute(command)


def _tagged_folder_result(outcome: FolderMutationOutcome) -> str:
    """Render one mailbox-shape outcome with only reviewed fixed detail tags."""
    status = outcome.status
    if outcome.detail in _PUBLIC_FOLDER_DETAILS:
        status = f"{status} ({outcome.detail})"
    sections = [status]
    if outcome.reconciliation_needed:
        sections.append("warning: reconciliation needed")
    return "; ".join(sections)


@mcp.tool(
    description=(
        "Copy one or more emails into another IMAP folder by email_id, leaving the originals in place. "
        "Use list_emails_metadata and list_mailboxes first. Partial or ambiguous effects report per-ID "
        "succeeded/failed/unknown status and are not retried automatically."
    ),
    annotations=_NONDESTRUCTIVE_REMOTE_MUTATION,
)
async def copy_emails(
    account_name: Annotated[
        str, Field(max_length=APPLICATION_LIMITS.account_name_bytes, description="The name of the email account.")
    ],
    email_ids: Annotated[
        list[UidInput],
        Field(
            min_length=1,
            max_length=APPLICATION_LIMITS.mutation_uids,
            description="List of email_id to copy (obtained from list_emails_metadata).",
        ),
    ],
    destination_mailbox: Annotated[
        str,
        Field(
            max_length=APPLICATION_LIMITS.mailbox_bytes,
            description="The destination mailbox/folder to copy emails into.",
        ),
    ],
    source_mailbox: Annotated[
        str,
        Field(
            default="INBOX",
            max_length=APPLICATION_LIMITS.mailbox_bytes,
            description="The source mailbox containing the emails.",
        ),
    ] = "INBOX",
) -> str:
    outcome = await copy_emails_command(
        CopyCommand(account_name, tuple(email_ids), source_mailbox, destination_mailbox)
    )
    succeeded = outcome.targets("succeeded")
    if len(succeeded) == len(email_ids) and not outcome.reconciliation_needed:
        return f"Successfully copied {len(succeeded)} email(s) to {destination_mailbox}"
    return f"Copy result [{_tagged_batch_result(outcome)}]"


@mcp.tool(
    description=(
        "Create a new IMAP folder/mailbox. Requires enable_folder_management=true in settings or "
        "MCP_EMAIL_SERVER_ENABLE_FOLDER_MANAGEMENT=true; managed-mode accounts never allow it."
    ),
    annotations=_NONDESTRUCTIVE_REMOTE_MUTATION,
)
async def create_folder(
    account_name: Annotated[
        str, Field(max_length=APPLICATION_LIMITS.account_name_bytes, description="The name of the email account.")
    ],
    folder_name: Annotated[
        str,
        Field(
            max_length=APPLICATION_LIMITS.mailbox_bytes,
            description="The name of the folder to create. Use list_mailboxes to learn the hierarchy delimiter.",
        ),
    ],
) -> str:
    outcome = await create_folder_command(CreateFolderCommand(account_name, folder_name))
    if outcome.status == "succeeded" and not outcome.reconciliation_needed:
        return f"Folder '{folder_name}' created"
    return f"Create-folder result [{_tagged_folder_result(outcome)}]"


@mcp.tool(
    description=(
        "Delete an IMAP folder/mailbox. Most servers require the folder to be empty and refuse to delete a "
        "folder that still has children. Requires enable_folder_management=true in settings or "
        "MCP_EMAIL_SERVER_ENABLE_FOLDER_MANAGEMENT=true; managed-mode accounts never allow it."
    ),
    annotations=_DESTRUCTIVE_REMOTE_MUTATION,
)
async def delete_folder(
    account_name: Annotated[
        str, Field(max_length=APPLICATION_LIMITS.account_name_bytes, description="The name of the email account.")
    ],
    folder_name: Annotated[
        str,
        Field(
            max_length=APPLICATION_LIMITS.mailbox_bytes,
            description="The name of the folder to delete.",
        ),
    ],
) -> str:
    outcome = await delete_folder_command(DeleteFolderCommand(account_name, folder_name))
    if outcome.status == "succeeded" and not outcome.reconciliation_needed:
        return f"Folder '{folder_name}' deleted"
    return f"Delete-folder result [{_tagged_folder_result(outcome)}]"


@mcp.tool(
    description=(
        "Rename an IMAP folder/mailbox, carrying its messages and child folders with it. Requires "
        "enable_folder_management=true in settings or MCP_EMAIL_SERVER_ENABLE_FOLDER_MANAGEMENT=true; "
        "managed-mode accounts never allow it."
    ),
    annotations=_DESTRUCTIVE_REMOTE_MUTATION,
)
async def rename_folder(
    account_name: Annotated[
        str, Field(max_length=APPLICATION_LIMITS.account_name_bytes, description="The name of the email account.")
    ],
    old_name: Annotated[
        str,
        Field(
            max_length=APPLICATION_LIMITS.mailbox_bytes,
            description="The current folder name.",
        ),
    ],
    new_name: Annotated[
        str,
        Field(
            max_length=APPLICATION_LIMITS.mailbox_bytes,
            description="The new folder name. It must differ from the current name.",
        ),
    ],
) -> str:
    outcome = await rename_folder_command(RenameFolderCommand(account_name, old_name, new_name))
    if outcome.status == "succeeded" and not outcome.reconciliation_needed:
        return f"Folder '{old_name}' renamed to '{new_name}'"
    return f"Rename-folder result [{_tagged_folder_result(outcome)}]"


# port-slot A: folder ops (copy_emails, create_folder, delete_folder, rename_folder)
LabelNameInput = Annotated[str, Field(max_length=MAXIMUM_LABEL_NAME_BYTES)]

# Batch results never reveal a `failed` detail (see `_tagged_batch_result`), but a
# label removal fails for reasons the caller can act on — the message carries no
# Message-ID, or the label was simply not applied. These reviewed fixed tags are
# the only detail text this tool may surface.
_PUBLIC_LABEL_DETAILS = frozenset({
    "expunge-not-attempted",
    "expunge-rejected",
    "expunge-unknown",
    "label-not-found",
    "label-search-failed",
    "label-unavailable",
    "message-id-missing",
    "message-id-unsupported",
    "not-attempted",
    "provider-timeout",
    "sender-policy",
    "store-rejected",
    "store-unknown",
    "uidplus-unavailable",
})


def _tagged_label_result(outcome: BatchMutationOutcome) -> str:
    sections = _ordered_target_sections(
        outcome.outcomes,
        detail_allowlist=_PUBLIC_LABEL_DETAILS,
        include_failed_detail=True,
        include_unknown_detail=True,
    )
    if outcome.reconciliation_needed:
        sections.append("warning: reconciliation needed")
    return "; ".join(sections)


async def list_labels_query(query: ListLabelsQuery) -> list[LabelInfo]:
    return await get_application_runtime().reads.labels.execute(query)


async def get_email_labels_query(query: GetEmailLabelsQuery) -> list[str]:
    return await get_application_runtime().reads.email_labels.execute(query)


async def remove_label_command(command: RemoveLabelCommand) -> BatchMutationOutcome:
    return await get_application_runtime().mutations.remove_label.execute(command)


@mcp.tool(
    description=(
        "List the account's labels. A label is an ordinary IMAP mailbox stored under the literal "
        "'Labels/' prefix, which is how ProtonMail and ProtonMail Bridge expose labels; accounts "
        "without that convention return an empty list. Each entry reports the label name, the "
        "mailbox that stores it, the hierarchy delimiter, and the mailbox flags."
    ),
    annotations=_READ_ONLY_REMOTE,
)
async def list_labels(
    account_name: Annotated[
        str, Field(max_length=APPLICATION_LIMITS.account_name_bytes, description="The name of the email account.")
    ],
) -> list[LabelInfo]:
    return await list_labels_query(ListLabelsQuery(account_name=account_name))


@mcp.tool(
    description=(
        "List the labels applied to one email. The message's Message-ID is read from its own mailbox "
        "and then looked for in every label mailbox within a single IMAP session. Returns an empty "
        "list when the message cannot be read, carries no Message-ID, or has no labels. A label "
        "mailbox that cannot be searched is skipped, so the result reports the labels that could be "
        "confirmed rather than failing outright."
    ),
    annotations=_READ_ONLY_REMOTE,
)
async def get_email_labels(
    account_name: Annotated[
        str, Field(max_length=APPLICATION_LIMITS.account_name_bytes, description="The name of the email account.")
    ],
    email_id: Annotated[UidInput, Field(description="The email_id to inspect (from list_emails_metadata).")],
    mailbox: Annotated[
        str,
        Field(
            default="INBOX",
            max_length=APPLICATION_LIMITS.mailbox_bytes,
            description="The mailbox that contains the email.",
        ),
    ] = "INBOX",
) -> list[str]:
    return await get_email_labels_query(
        GetEmailLabelsQuery(account_name=account_name, email_id=email_id, mailbox=mailbox)
    )


@mcp.tool(
    description=(
        "Remove one label from one or more emails. The label's own copy of each message is located "
        "in the 'Labels/<label_name>' mailbox by Message-ID and deleted there with a target-scoped "
        "UID EXPUNGE; the message in source_mailbox is never modified. Use list_labels and "
        "list_emails_metadata first. Partial or ambiguous effects report per-ID "
        "succeeded/failed/unknown status and are not retried automatically."
    ),
    annotations=_DESTRUCTIVE_REMOTE_MUTATION,
)
async def remove_label(
    account_name: Annotated[
        str, Field(max_length=APPLICATION_LIMITS.account_name_bytes, description="The name of the email account.")
    ],
    email_ids: Annotated[
        list[UidInput],
        Field(
            min_length=1,
            max_length=APPLICATION_LIMITS.mutation_uids,
            description="List of email_id to remove the label from (obtained from list_emails_metadata).",
        ),
    ],
    label_name: Annotated[
        LabelNameInput,
        Field(description="The label to remove, without the 'Labels/' prefix (obtained from list_labels)."),
    ],
    source_mailbox: Annotated[
        str,
        Field(
            default="INBOX",
            max_length=APPLICATION_LIMITS.mailbox_bytes,
            description="The mailbox that contains the emails.",
        ),
    ] = "INBOX",
) -> str:
    outcome = await remove_label_command(RemoveLabelCommand(account_name, tuple(email_ids), label_name, source_mailbox))
    succeeded = outcome.targets("succeeded")
    if len(succeeded) == len(email_ids) and not outcome.reconciliation_needed:
        return f"Successfully removed label '{label_name}' from {len(succeeded)} email(s)"
    return f"Remove-label result [{_tagged_label_result(outcome)}]"


# port-slot B1: label reads (list_labels, get_email_labels, remove_label)
# port-slot C: send path (no new tools; send_email gains Markdown and quoted replies)
async def create_label_command(command: CreateLabelCommand) -> FolderMutationOutcome:
    return await get_application_runtime().mutations.create_label.execute(command)


async def delete_label_command(command: DeleteLabelCommand) -> FolderMutationOutcome:
    return await get_application_runtime().mutations.delete_label.execute(command)


async def apply_label_command(command: ApplyLabelCommand) -> BatchMutationOutcome:
    return await get_application_runtime().mutations.apply_label.execute(command)


@mcp.tool(
    description=(
        "Create a new label. A label is an IMAP mailbox named 'Labels/<label_name>', so this creates "
        "that mailbox; pass the label name without the 'Labels/' prefix. Requires "
        "enable_folder_management=true in settings or MCP_EMAIL_SERVER_ENABLE_FOLDER_MANAGEMENT=true; "
        "managed-mode accounts never allow it."
    ),
    annotations=_NONDESTRUCTIVE_REMOTE_MUTATION,
)
async def create_label(
    account_name: Annotated[
        str, Field(max_length=APPLICATION_LIMITS.account_name_bytes, description="The name of the email account.")
    ],
    label_name: Annotated[
        LabelNameInput,
        Field(description="The label to create, without the 'Labels/' prefix."),
    ],
) -> str:
    outcome = await create_label_command(CreateLabelCommand(account_name, label_name))
    if outcome.status == "succeeded" and not outcome.reconciliation_needed:
        return f"Label '{label_name}' created"
    return f"Create-label result [{_tagged_folder_result(outcome)}]"


@mcp.tool(
    description=(
        "Delete a label and the label's own copies of every message it holds; the messages in their "
        "own mailboxes are not affected. Pass the label name without the 'Labels/' prefix. Requires "
        "enable_folder_management=true in settings or MCP_EMAIL_SERVER_ENABLE_FOLDER_MANAGEMENT=true; "
        "managed-mode accounts never allow it."
    ),
    annotations=_DESTRUCTIVE_REMOTE_MUTATION,
)
async def delete_label(
    account_name: Annotated[
        str, Field(max_length=APPLICATION_LIMITS.account_name_bytes, description="The name of the email account.")
    ],
    label_name: Annotated[
        LabelNameInput,
        Field(description="The label to delete, without the 'Labels/' prefix (obtained from list_labels)."),
    ],
) -> str:
    outcome = await delete_label_command(DeleteLabelCommand(account_name, label_name))
    if outcome.status == "succeeded" and not outcome.reconciliation_needed:
        return f"Label '{label_name}' deleted"
    return f"Delete-label result [{_tagged_folder_result(outcome)}]"


@mcp.tool(
    description=(
        "Apply one label to one or more emails by copying each message into the 'Labels/<label_name>' "
        "mailbox; the message in source_mailbox stays exactly where it is. The label mailbox must "
        "already exist — use list_labels or create_label first. Partial or ambiguous effects report "
        "per-ID succeeded/failed/unknown status and are not retried automatically."
    ),
    annotations=_NONDESTRUCTIVE_REMOTE_MUTATION,
)
async def apply_label(
    account_name: Annotated[
        str, Field(max_length=APPLICATION_LIMITS.account_name_bytes, description="The name of the email account.")
    ],
    email_ids: Annotated[
        list[UidInput],
        Field(
            min_length=1,
            max_length=APPLICATION_LIMITS.mutation_uids,
            description="List of email_id to apply the label to (obtained from list_emails_metadata).",
        ),
    ],
    label_name: Annotated[
        LabelNameInput,
        Field(description="The label to apply, without the 'Labels/' prefix (obtained from list_labels)."),
    ],
    source_mailbox: Annotated[
        str,
        Field(
            default="INBOX",
            max_length=APPLICATION_LIMITS.mailbox_bytes,
            description="The mailbox that contains the emails.",
        ),
    ] = "INBOX",
) -> str:
    outcome = await apply_label_command(ApplyLabelCommand(account_name, tuple(email_ids), label_name, source_mailbox))
    succeeded = outcome.targets("succeeded")
    if len(succeeded) == len(email_ids) and not outcome.reconciliation_needed:
        return f"Successfully applied label '{label_name}' to {len(succeeded)} email(s)"
    return f"Apply-label result [{_tagged_batch_result(outcome)}]"


# port-slot B2: label writes (create_label, delete_label, apply_label)
