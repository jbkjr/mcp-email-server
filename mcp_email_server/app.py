from datetime import datetime
from typing import Annotated, Literal

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from mcp_email_server.config import (
    AccountAttributes,
    EmailSettings,
    ProviderSettings,
    get_settings,
)
from mcp_email_server.emails.dispatcher import dispatch_handler
from mcp_email_server.emails.models import (
    AttachmentDownloadResponse,
    EmailContentBatchResponse,
    EmailDeleteResponse,
    EmailLabelsResponse,
    EmailMarkResponse,
    EmailMetadataPageResponse,
    EmailMoveResponse,
    EmailSendResponse,
    FolderListResponse,
    FolderOperationResponse,
    LabelListResponse,
)

mcp = FastMCP(
    "email",
    instructions="When sending emails, the body supports Markdown formatting (bold, lists, headers, links, etc.) which is automatically converted to email-safe HTML. Use Markdown freely for well-formatted emails. Set html=True only if providing pre-formatted raw HTML.",
)


@mcp.resource("email://{account_name}")
async def get_account(account_name: str) -> EmailSettings | ProviderSettings | None:
    settings = get_settings()
    return settings.get_account(account_name, masked=True)


@mcp.tool(description="List all configured email accounts with masked credentials.")
async def list_available_accounts() -> list[AccountAttributes]:
    settings = get_settings()
    return [account.masked() for account in settings.get_accounts()]


@mcp.tool(description="Add a new email account configuration to the settings.")
async def add_email_account(email: EmailSettings) -> str:
    settings = get_settings()
    settings.add_email(email)
    settings.store()
    return f"Successfully added email account '{email.account_name}'"


@mcp.tool(
    description="List email metadata (email_id, subject, sender, recipients, date) without body content. Returns email_id for use with get_emails_content."
)
async def list_emails_metadata(
    account_name: Annotated[str, Field(description="The name of the email account.")],
    page: Annotated[
        int,
        Field(default=1, description="The page number to retrieve (starting from 1)."),
    ] = 1,
    page_size: Annotated[int, Field(default=10, description="The number of emails to retrieve per page.")] = 10,
    before: Annotated[
        datetime | None,
        Field(default=None, description="Retrieve emails before this datetime (UTC)."),
    ] = None,
    since: Annotated[
        datetime | None,
        Field(default=None, description="Retrieve emails since this datetime (UTC)."),
    ] = None,
    subject: Annotated[str | None, Field(default=None, description="Filter emails by subject.")] = None,
    from_address: Annotated[str | None, Field(default=None, description="Filter emails by sender address.")] = None,
    to_address: Annotated[
        str | None,
        Field(default=None, description="Filter emails by recipient address."),
    ] = None,
    order: Annotated[
        Literal["asc", "desc"],
        Field(default=None, description="Order emails by field. `asc` or `desc`."),
    ] = "desc",
    mailbox: Annotated[
        str, Field(default="INBOX", description="IMAP folder path. Standard: INBOX, Sent, Drafts, Trash. Provider-specific: Gmail uses '[Gmail]/...' prefix (e.g., '[Gmail]/Sent Mail'); ProtonMail Bridge exposes folders as 'Folders/<name>' and labels as 'Labels/<name>'.")
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
        Field(default=None, description="Search for text in the email body (IMAP BODY search)."),
    ] = None,
    text: Annotated[
        str | None,
        Field(default=None, description="Search for text in the entire email including headers and body (IMAP TEXT search)."),
    ] = None,
    has_attachment: Annotated[
        bool | None,
        Field(default=None, description="Filter by attachment presence: True=has attachments, False=no attachments, None=all."),
    ] = None,
) -> EmailMetadataPageResponse:
    handler = dispatch_handler(account_name)

    return await handler.get_emails_metadata(
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


@mcp.tool(
    description="Get the full content (including body) of one or more emails by their email_id. Use list_emails_metadata first to get the email_id."
)
async def get_emails_content(
    account_name: Annotated[str, Field(description="The name of the email account.")],
    email_ids: Annotated[
        list[str],
        Field(
            description="List of email_id to retrieve (obtained from list_emails_metadata). Can be a single email_id or multiple email_ids."
        ),
    ],
    mailbox: Annotated[str, Field(default="INBOX", description="IMAP folder path. Standard: INBOX, Sent, Drafts, Trash. Provider-specific: Gmail uses '[Gmail]/...' prefix; ProtonMail Bridge uses 'Folders/<name>' and 'Labels/<name>'.")] = "INBOX",
    max_body_length: Annotated[
        int | None,
        Field(default=20000, description="Maximum body length in characters before truncation. Set to 0 or null for no limit. Default: 20000."),
    ] = 20000,
    mark_as_read: Annotated[
        bool,
        Field(default=False, description="Mark fetched emails as read. Default: False (emails remain unread)."),
    ] = False,
) -> EmailContentBatchResponse:
    handler = dispatch_handler(account_name)
    # Treat 0 as no limit
    effective_limit = max_body_length if max_body_length else None
    return await handler.get_emails_content(email_ids, mailbox, effective_limit, mark_as_read=mark_as_read)


@mcp.tool(
    description="Send an email using the specified account. Supports replying to emails with proper threading when in_reply_to is provided.",
)
async def send_email(
    account_name: Annotated[str, Field(description="The name of the email account to send from.")],
    recipients: Annotated[list[str], Field(description="A list of recipient email addresses.")],
    subject: Annotated[str, Field(description="The subject of the email.")],
    body: Annotated[str, Field(description="The email body. Supports Markdown formatting (bold, lists, headers, links, etc.) which is automatically converted to email-safe HTML.")],
    cc: Annotated[
        list[str] | None,
        Field(default=None, description="A list of CC email addresses."),
    ] = None,
    bcc: Annotated[
        list[str] | None,
        Field(default=None, description="A list of BCC email addresses."),
    ] = None,
    html: Annotated[
        bool,
        Field(default=False, description="Set to True only if the body contains pre-formatted raw HTML. When False (default), the body is automatically converted from Markdown/plain text to email-safe HTML."),
    ] = False,
    attachments: Annotated[
        list[str] | None,
        Field(
            default=None,
            description="A list of absolute file paths to attach to the email. Supports common file types (documents, images, archives, etc.).",
        ),
    ] = None,
    in_reply_to: Annotated[
        str | None,
        Field(
            default=None,
            description="Message-ID of the email being replied to. Enables proper threading in email clients.",
        ),
    ] = None,
    references: Annotated[
        str | None,
        Field(
            default=None,
            description="Space-separated Message-IDs for the thread chain. Usually includes in_reply_to plus ancestors.",
        ),
    ] = None,
    quote_reply: Annotated[
        bool,
        Field(
            default=True,
            description="When replying (in_reply_to is set), automatically fetch and append the quoted original message. Set to False if you've already included quoted text in the body.",
        ),
    ] = True,
) -> EmailSendResponse:
    handler = dispatch_handler(account_name)
    return await handler.send_email(
        recipients,
        subject,
        body,
        cc,
        bcc,
        html,
        attachments,
        in_reply_to,
        references,
        quote_reply,
    )


@mcp.tool(
    description="Forward an email to new recipients. The original message is included below any optional message you add, with original attachments forwarded automatically.",
)
async def forward_email(
    account_name: Annotated[str, Field(description="The name of the email account to send from.")],
    email_id: Annotated[str, Field(description="The email_id of the email to forward (obtained from list_emails_metadata).")],
    recipients: Annotated[list[str], Field(description="A list of recipient email addresses to forward to.")],
    mailbox: Annotated[
        str, Field(default="INBOX", description="IMAP folder path containing the email to forward.")
    ] = "INBOX",
    body: Annotated[
        str | None,
        Field(default=None, description="Optional message to prepend above the forwarded content. Supports Markdown formatting."),
    ] = None,
    cc: Annotated[
        list[str] | None,
        Field(default=None, description="A list of CC email addresses."),
    ] = None,
    bcc: Annotated[
        list[str] | None,
        Field(default=None, description="A list of BCC email addresses."),
    ] = None,
    html: Annotated[
        bool,
        Field(default=False, description="Set to True only if the body contains pre-formatted raw HTML. When False (default), the body is automatically converted from Markdown/plain text to email-safe HTML."),
    ] = False,
    attachments: Annotated[
        list[str] | None,
        Field(
            default=None,
            description="A list of additional absolute file paths to attach to the forwarded email.",
        ),
    ] = None,
) -> EmailSendResponse:
    handler = dispatch_handler(account_name)
    return await handler.forward_email(
        email_id,
        mailbox,
        recipients,
        body,
        cc,
        bcc,
        html,
        attachments,
    )


@mcp.tool(
    description="Delete one or more emails by their email_id. Moves to Trash when available (auto-detected via RFC 6154 flags), falls back to permanent deletion. Use list_emails_metadata first to get the email_id."
)
async def delete_emails(
    account_name: Annotated[str, Field(description="The name of the email account.")],
    email_ids: Annotated[
        list[str],
        Field(description="List of email_id to delete (obtained from list_emails_metadata)."),
    ],
    mailbox: Annotated[str, Field(default="INBOX", description="IMAP folder path. Standard: INBOX, Sent, Drafts, Trash. Provider-specific: Gmail uses '[Gmail]/...' prefix; ProtonMail Bridge uses 'Folders/<name>' and 'Labels/<name>'.")] = "INBOX",
) -> EmailDeleteResponse:
    handler = dispatch_handler(account_name)
    return await handler.delete_emails(email_ids, mailbox)


@mcp.tool(
    description="Archive one or more emails by moving them to the Archive folder (auto-detected via RFC 6154 flags or common folder names)."
)
async def archive_emails(
    account_name: Annotated[str, Field(description="The name of the email account.")],
    email_ids: Annotated[
        list[str],
        Field(description="List of email_id to archive (obtained from list_emails_metadata)."),
    ],
    mailbox: Annotated[str, Field(default="INBOX", description="IMAP folder path containing the emails to archive.")] = "INBOX",
) -> EmailMoveResponse:
    handler = dispatch_handler(account_name)
    return await handler.archive_emails(email_ids, mailbox)


@mcp.tool(
    description="Mark one or more emails as read or unread. Use list_emails_metadata first to get the email_id."
)
async def mark_emails(
    account_name: Annotated[str, Field(description="The name of the email account.")],
    email_ids: Annotated[
        list[str],
        Field(description="List of email_id to mark (obtained from list_emails_metadata)."),
    ],
    mark_as: Annotated[
        Literal["read", "unread"],
        Field(description="Mark emails as 'read' or 'unread'."),
    ],
    mailbox: Annotated[str, Field(default="INBOX", description="IMAP folder path. Standard: INBOX, Sent, Drafts, Trash. Provider-specific: Gmail uses '[Gmail]/...' prefix; ProtonMail Bridge uses 'Folders/<name>' and 'Labels/<name>'.")] = "INBOX",
) -> EmailMarkResponse:
    handler = dispatch_handler(account_name)
    return await handler.mark_emails(email_ids, mark_as, mailbox)


@mcp.tool(
    description="Download an email attachment and save it to the specified path. This feature must be explicitly enabled in settings (enable_attachment_download=true) due to security considerations.",
)
async def download_attachment(
    account_name: Annotated[str, Field(description="The name of the email account.")],
    email_id: Annotated[
        str, Field(description="The email ID (obtained from list_emails_metadata or get_emails_content).")
    ],
    attachment_name: Annotated[
        str, Field(description="The name of the attachment to download (as shown in the attachments list).")
    ],
    save_path: Annotated[str, Field(description="The absolute path where the attachment should be saved.")],
    mailbox: Annotated[str, Field(default="INBOX", description="IMAP folder path. Standard: INBOX, Sent, Drafts, Trash. Provider-specific: Gmail uses '[Gmail]/...' prefix; ProtonMail Bridge uses 'Folders/<name>' and 'Labels/<name>'.")] = "INBOX",
) -> AttachmentDownloadResponse:
    settings = get_settings()
    if not settings.enable_attachment_download:
        msg = (
            "Attachment download is disabled. Set 'enable_attachment_download=true' in settings to enable this feature."
        )
        raise PermissionError(msg)

    handler = dispatch_handler(account_name)
    return await handler.download_attachment(email_id, attachment_name, save_path, mailbox)


def _check_folder_management_enabled() -> None:
    """Check if folder management is enabled, raise PermissionError if not."""
    settings = get_settings()
    if not settings.enable_folder_management:
        msg = (
            "Folder management is disabled. Set 'enable_folder_management=true' in settings "
            "or 'MCP_EMAIL_SERVER_ENABLE_FOLDER_MANAGEMENT=true' environment variable to enable this feature."
        )
        raise PermissionError(msg)


@mcp.tool(
    description="List all folders/mailboxes for an email account. Returns folder names, hierarchy delimiters, and IMAP flags. Requires enable_folder_management=true.",
)
async def list_folders(
    account_name: Annotated[str, Field(description="The name of the email account.")],
) -> FolderListResponse:
    _check_folder_management_enabled()
    handler = dispatch_handler(account_name)
    return await handler.list_folders()


@mcp.tool(
    description="Move one or more emails to a different folder (removes from source). Use this to clear emails from INBOX. Uses IMAP MOVE command if supported, otherwise falls back to COPY + DELETE. Requires enable_folder_management=true.",
)
async def move_emails(
    account_name: Annotated[str, Field(description="The name of the email account.")],
    email_ids: Annotated[
        list[str],
        Field(description="List of email_id to move (obtained from list_emails_metadata)."),
    ],
    destination_folder: Annotated[str, Field(description="The destination folder name.")],
    source_mailbox: Annotated[
        str, Field(default="INBOX", description="The source mailbox to move emails from.")
    ] = "INBOX",
) -> EmailMoveResponse:
    _check_folder_management_enabled()
    handler = dispatch_handler(account_name)
    return await handler.move_emails(email_ids, destination_folder, source_mailbox)


@mcp.tool(
    description="Copy one or more emails to a different folder. The original emails remain in the source folder. Useful for applying labels in providers like Proton Mail. Requires enable_folder_management=true.",
)
async def copy_emails(
    account_name: Annotated[str, Field(description="The name of the email account.")],
    email_ids: Annotated[
        list[str],
        Field(description="List of email_id to copy (obtained from list_emails_metadata)."),
    ],
    destination_folder: Annotated[str, Field(description="The destination folder name.")],
    source_mailbox: Annotated[
        str, Field(default="INBOX", description="The source mailbox to copy emails from.")
    ] = "INBOX",
) -> EmailMoveResponse:
    _check_folder_management_enabled()
    handler = dispatch_handler(account_name)
    return await handler.copy_emails(email_ids, destination_folder, source_mailbox)


@mcp.tool(description="Create a new folder/mailbox. Requires enable_folder_management=true.")
async def create_folder(
    account_name: Annotated[str, Field(description="The name of the email account.")],
    folder_name: Annotated[str, Field(description="The name of the folder to create.")],
) -> FolderOperationResponse:
    _check_folder_management_enabled()
    handler = dispatch_handler(account_name)
    return await handler.create_folder(folder_name)


@mcp.tool(
    description="Delete a folder/mailbox. The folder must be empty on most IMAP servers. Requires enable_folder_management=true."
)
async def delete_folder(
    account_name: Annotated[str, Field(description="The name of the email account.")],
    folder_name: Annotated[str, Field(description="The name of the folder to delete.")],
) -> FolderOperationResponse:
    _check_folder_management_enabled()
    handler = dispatch_handler(account_name)
    return await handler.delete_folder(folder_name)


@mcp.tool(description="Rename a folder/mailbox. Requires enable_folder_management=true.")
async def rename_folder(
    account_name: Annotated[str, Field(description="The name of the email account.")],
    old_name: Annotated[str, Field(description="The current folder name.")],
    new_name: Annotated[str, Field(description="The new folder name.")],
) -> FolderOperationResponse:
    _check_folder_management_enabled()
    handler = dispatch_handler(account_name)
    return await handler.rename_folder(old_name, new_name)


@mcp.tool(
    description="List all labels for an email account (ProtonMail: folders under Labels/ prefix). Requires enable_folder_management=true."
)
async def list_labels(
    account_name: Annotated[str, Field(description="The name of the email account.")],
) -> LabelListResponse:
    _check_folder_management_enabled()
    handler = dispatch_handler(account_name)
    return await handler.list_labels()


@mcp.tool(
    description="Apply a label to one or more emails. NOTE: This only tags emails - originals stay in INBOX. To remove from INBOX, use move_emails instead. Requires enable_folder_management=true."
)
async def apply_label(
    account_name: Annotated[str, Field(description="The name of the email account.")],
    email_ids: Annotated[
        list[str],
        Field(description="List of email_id to label (obtained from list_emails_metadata)."),
    ],
    label_name: Annotated[str, Field(description="The label name (without Labels/ prefix).")],
    source_mailbox: Annotated[
        str, Field(default="INBOX", description="The source mailbox containing the emails.")
    ] = "INBOX",
) -> EmailMoveResponse:
    _check_folder_management_enabled()
    handler = dispatch_handler(account_name)
    return await handler.apply_label(email_ids, label_name, source_mailbox)


@mcp.tool(
    description="Remove a label from one or more emails. Deletes from label folder while preserving original emails. Requires enable_folder_management=true."
)
async def remove_label(
    account_name: Annotated[str, Field(description="The name of the email account.")],
    email_ids: Annotated[
        list[str],
        Field(description="List of email_id to unlabel (obtained from list_emails_metadata)."),
    ],
    label_name: Annotated[str, Field(description="The label name (without Labels/ prefix).")],
    source_mailbox: Annotated[
        str,
        Field(default="INBOX", description="The mailbox where the original emails reside (for Message-ID lookup)."),
    ] = "INBOX",
) -> EmailMoveResponse:
    _check_folder_management_enabled()
    handler = dispatch_handler(account_name)
    return await handler.remove_label(email_ids, label_name, source_mailbox)


@mcp.tool(description="Get all labels applied to a specific email. Requires enable_folder_management=true.")
async def get_email_labels(
    account_name: Annotated[str, Field(description="The name of the email account.")],
    email_id: Annotated[str, Field(description="The email_id to check (obtained from list_emails_metadata).")],
    source_mailbox: Annotated[
        str, Field(default="INBOX", description="The source mailbox containing the email.")
    ] = "INBOX",
) -> EmailLabelsResponse:
    _check_folder_management_enabled()
    handler = dispatch_handler(account_name)
    return await handler.get_email_labels(email_id, source_mailbox)


@mcp.tool(description="Create a new label (creates Labels/name folder). Requires enable_folder_management=true.")
async def create_label(
    account_name: Annotated[str, Field(description="The name of the email account.")],
    label_name: Annotated[str, Field(description="The label name to create (without Labels/ prefix).")],
) -> FolderOperationResponse:
    _check_folder_management_enabled()
    handler = dispatch_handler(account_name)
    return await handler.create_label(label_name)


@mcp.tool(
    description="Delete a label (deletes Labels/name folder). The label must be empty on most IMAP servers. Requires enable_folder_management=true."
)
async def delete_label(
    account_name: Annotated[str, Field(description="The name of the email account.")],
    label_name: Annotated[str, Field(description="The label name to delete (without Labels/ prefix).")],
) -> FolderOperationResponse:
    _check_folder_management_enabled()
    handler = dispatch_handler(account_name)
    return await handler.delete_label(label_name)
