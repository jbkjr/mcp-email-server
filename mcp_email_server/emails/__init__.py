import abc
from datetime import datetime
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from mcp_email_server.emails.models import (
        AttachmentDownloadResponse,
        EmailContentBatchResponse,
        EmailDeleteResponse,
        EmailLabelsResponse,
        EmailMarkResponse,
        EmailMetadataPageResponse,
        EmailMoveResponse,
        EmailSendResponse,
        FolderOperationResponse,
        LabelListResponse,
        MailboxInfo,
    )


class EmailHandler(abc.ABC):
    @abc.abstractmethod
    async def get_emails_metadata(
        self,
        page: int = 1,
        page_size: int = 10,
        before: datetime | None = None,
        since: datetime | None = None,
        subject: str | None = None,
        from_address: str | None = None,
        to_address: str | None = None,
        order: str = "desc",
        mailbox: str = "INBOX",
        seen: bool | None = None,
        flagged: bool | None = None,
        answered: bool | None = None,
        body: str | None = None,
        text: str | None = None,
        has_attachment: bool | None = None,
    ) -> "EmailMetadataPageResponse":
        """
        Get email metadata only (without body content) for better performance.

        Args:
            page: Page number (starting from 1).
            page_size: Number of emails per page.
            before: Filter emails before this datetime.
            since: Filter emails since this datetime.
            subject: Filter by subject (substring match).
            from_address: Filter by sender address.
            to_address: Filter by recipient address.
            order: Sort order ('asc' or 'desc').
            mailbox: Mailbox to search (e.g., 'INBOX', 'Labels/LabelName').
            seen: Filter by read status (True=read, False=unread, None=all).
            flagged: Filter by flagged/starred status (True=flagged, False=unflagged, None=all).
            answered: Filter by replied status (True=replied, False=not replied, None=all).
            body: Search for text in the email body (IMAP BODY).
            text: Search for text in the entire message, headers + body (IMAP TEXT).
            has_attachment: Filter by attachment presence (True/False/None) via a
                multipart/mixed Content-Type heuristic.
        """

    @abc.abstractmethod
    async def get_emails_content(
        self,
        email_ids: list[str],
        mailbox: str = "INBOX",
        mark_as_read: bool = False,
        body_offset: int = 0,
        max_body_length: int | None = 20000,
    ) -> "EmailContentBatchResponse":
        """
        Get full content (including body) of multiple emails by their email IDs (IMAP UIDs).

        Args:
            email_ids: List of email UIDs to retrieve.
            mailbox: Mailbox to search in.
            mark_as_read: If True, mark successfully fetched emails as read.
            body_offset: Character offset into each body to start reading from (for paging).
            max_body_length: Maximum number of body characters to return, counted from
                body_offset. 0 or None for no limit. If more remains, the body ends with
                the '...[TRUNCATED]' marker.
        """

    @abc.abstractmethod
    async def send_email(
        self,
        recipients: list[str],
        subject: str,
        body: str,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        html: bool = False,
        attachments: list[str] | None = None,
        in_reply_to: str | None = None,
        references: str | None = None,
        quote_reply: bool = True,
        reply_to: str | None = None,
    ) -> "EmailSendResponse":
        """
        Send email

        Args:
            recipients: List of recipient email addresses.
            subject: Email subject.
            body: Email body content (Markdown/plain text, auto-converted to HTML).
            cc: List of CC email addresses.
            bcc: List of BCC email addresses.
            html: If True, body is pre-formatted HTML (skip Markdown conversion).
            attachments: List of file paths to attach.
            in_reply_to: Message-ID of the email being replied to (for threading).
            references: Space-separated Message-IDs for the thread chain.
            quote_reply: When replying, auto-fetch and append the quoted original message.
            reply_to: Address to set as Reply-To header (overrides From for replies).
        """

    @abc.abstractmethod
    async def forward_email(
        self,
        email_id: str,
        mailbox: str,
        recipients: list[str],
        body: str | None = None,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        html: bool = False,
        attachments: list[str] | None = None,
    ) -> "EmailSendResponse":
        """
        Forward an email to new recipients.

        Args:
            email_id: The UID of the email to forward.
            mailbox: The mailbox containing the email.
            recipients: List of recipient email addresses.
            body: Optional message to prepend above the forwarded content (Markdown).
            cc: List of CC email addresses.
            bcc: List of BCC email addresses.
            html: If True, body is pre-formatted HTML (skip Markdown conversion).
            attachments: List of additional file paths to attach.
        """

    @abc.abstractmethod
    async def save_to_mailbox(
        self,
        recipients: list[str],
        subject: str,
        body: str,
        mailbox: str = "Drafts",
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        html: bool = False,
        attachments: list[str] | None = None,
        in_reply_to: str | None = None,
        references: str | None = None,
        flags: list[str] | None = None,
    ) -> str:
        """Compose an email and save it to the specified IMAP folder via APPEND."""

    @abc.abstractmethod
    async def delete_emails(self, email_ids: list[str], mailbox: str = "INBOX") -> "EmailDeleteResponse":
        """
        Delete emails by their IDs.
        """

    @abc.abstractmethod
    async def archive_emails(self, email_ids: list[str], mailbox: str = "INBOX") -> "EmailMoveResponse":
        """
        Archive emails by moving to the Archive folder (auto-detected via RFC 6154 \\Archive flag).

        Args:
            email_ids: List of email UIDs to archive.
            mailbox: The source mailbox (default: "INBOX").

        Returns:
            EmailMoveResponse with operation results.
        """

    @abc.abstractmethod
    async def move_emails(
        self, email_ids: list[str], source_mailbox: str, destination_mailbox: str
    ) -> tuple[list[str], list[str]]:
        """
        Move emails between mailboxes. Returns (moved_ids, failed_ids)

        Args:
            email_ids: List of email UIDs to move.
            source_mailbox: The mailbox to move emails from.
            destination_mailbox: The mailbox to move emails to.
        """

    @abc.abstractmethod
    async def list_mailboxes(self, pattern: str = "*", reference: str = "") -> list["MailboxInfo"]:
        """
        List available mailboxes/folders in the account.

        Args:
            pattern: IMAP LIST pattern (e.g., "*" for all, "INBOX.*" for INBOX children).
            reference: IMAP LIST reference name (namespace prefix).

        Returns:
            List of MailboxInfo with name, delimiter, and flags.
        """

    @abc.abstractmethod
    async def download_attachment(
        self,
        email_id: str,
        attachment_name: str,
        save_path: str | None = None,
        mailbox: str = "INBOX",
    ) -> "AttachmentDownloadResponse":
        """
        Download an email attachment and save it to disk.

        Args:
            email_id: The UID of the email containing the attachment.
            attachment_name: The filename of the attachment to download.
            save_path: Explicit destination path. When omitted, the attachment is
                saved with a sanitized randomized name under the current user's
                ``Downloads/mcp-email-server`` directory.
            mailbox: The mailbox to search in (default: "INBOX").

        Returns:
            AttachmentDownloadResponse with download result information.
        """

    @abc.abstractmethod
    async def copy_emails(
        self,
        email_ids: list[str],
        destination_folder: str,
        source_mailbox: str = "INBOX",
    ) -> "EmailMoveResponse":
        """
        Copy emails to a destination folder (preserves original).

        Args:
            email_ids: List of email UIDs to copy.
            destination_folder: The target folder name.
            source_mailbox: The source mailbox (default: "INBOX").

        Returns:
            EmailMoveResponse with operation results.
        """

    @abc.abstractmethod
    async def create_folder(self, folder_name: str) -> "FolderOperationResponse":
        """
        Create a new folder/mailbox.

        Args:
            folder_name: The name of the folder to create.

        Returns:
            FolderOperationResponse with operation result.
        """

    @abc.abstractmethod
    async def delete_folder(self, folder_name: str) -> "FolderOperationResponse":
        """
        Delete a folder/mailbox.

        Args:
            folder_name: The name of the folder to delete.

        Returns:
            FolderOperationResponse with operation result.
        """

    @abc.abstractmethod
    async def rename_folder(self, old_name: str, new_name: str) -> "FolderOperationResponse":
        """
        Rename a folder/mailbox.

        Args:
            old_name: The current folder name.
            new_name: The new folder name.

        Returns:
            FolderOperationResponse with operation result.
        """

    @abc.abstractmethod
    async def list_labels(self) -> "LabelListResponse":
        """
        List all labels (ProtonMail: folders under Labels/ prefix).

        Returns:
            LabelListResponse with list of labels.
        """

    @abc.abstractmethod
    async def apply_label(
        self,
        email_ids: list[str],
        label_name: str,
        source_mailbox: str = "INBOX",
    ) -> "EmailMoveResponse":
        """
        Apply a label to emails by copying to the label folder.

        Args:
            email_ids: List of email UIDs to label.
            label_name: The label name (without Labels/ prefix).
            source_mailbox: The source mailbox (default: "INBOX").

        Returns:
            EmailMoveResponse with operation results.
        """

    @abc.abstractmethod
    async def remove_label(
        self,
        email_ids: list[str],
        label_name: str,
        source_mailbox: str = "INBOX",
    ) -> "EmailMoveResponse":
        """
        Remove a label from emails by deleting from the label folder.

        Args:
            email_ids: List of email UIDs to unlabel.
            label_name: The label name (without Labels/ prefix).
            source_mailbox: Mailbox where the original emails reside (for Message-ID lookup).

        Returns:
            EmailMoveResponse with operation results.
        """

    @abc.abstractmethod
    async def get_email_labels(
        self,
        email_id: str,
        source_mailbox: str = "INBOX",
    ) -> "EmailLabelsResponse":
        """
        Get all labels applied to a specific email.

        Args:
            email_id: The email UID to check.
            source_mailbox: The source mailbox (default: "INBOX").

        Returns:
            EmailLabelsResponse with list of label names.
        """

    @abc.abstractmethod
    async def create_label(self, label_name: str) -> "FolderOperationResponse":
        """
        Create a new label (creates Labels/name folder).

        Args:
            label_name: The label name (without Labels/ prefix).

        Returns:
            FolderOperationResponse with operation result.
        """

    @abc.abstractmethod
    async def delete_label(self, label_name: str) -> "FolderOperationResponse":
        """
        Delete a label (deletes Labels/name folder).

        Args:
            label_name: The label name (without Labels/ prefix).

        Returns:
            FolderOperationResponse with operation result.
        """

    @abc.abstractmethod
    async def mark_emails(
        self,
        email_ids: list[str],
        mark_as: str,
        mailbox: str = "INBOX",
    ) -> "EmailMarkResponse":
        """
        Mark emails as read or unread.

        Args:
            email_ids: List of email UIDs to mark.
            mark_as: Either "read" or "unread".
            mailbox: The mailbox containing the emails (default: "INBOX").

        Returns:
            EmailMarkResponse with operation results.
        """
