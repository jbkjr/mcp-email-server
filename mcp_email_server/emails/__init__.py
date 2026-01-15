import abc
from datetime import datetime
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from mcp_email_server.emails.models import (
        AttachmentDownloadResponse,
        EmailContentBatchResponse,
        EmailLabelsResponse,
        EmailMetadataPageResponse,
        EmailMoveResponse,
        FolderListResponse,
        FolderOperationResponse,
        LabelListResponse,
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
    ) -> "EmailMetadataPageResponse":
        """
        Get email metadata only (without body content) for better performance
        """

    @abc.abstractmethod
    async def get_emails_content(self, email_ids: list[str], mailbox: str = "INBOX") -> "EmailContentBatchResponse":
        """
        Get full content (including body) of multiple emails by their email IDs (IMAP UIDs)
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
    ) -> None:
        """
        Send email
        """

    @abc.abstractmethod
    async def delete_emails(self, email_ids: list[str], mailbox: str = "INBOX") -> tuple[list[str], list[str]]:
        """
        Delete emails by their IDs. Returns (deleted_ids, failed_ids)
        """

    @abc.abstractmethod
    async def download_attachment(
        self,
        email_id: str,
        attachment_name: str,
        save_path: str,
        mailbox: str = "INBOX",
    ) -> "AttachmentDownloadResponse":
        """
        Download an email attachment and save it to the specified path.

        Args:
            email_id: The UID of the email containing the attachment.
            attachment_name: The filename of the attachment to download.
            save_path: The local path where the attachment will be saved.
            mailbox: The mailbox to search in (default: "INBOX").

        Returns:
            AttachmentDownloadResponse with download result information.
        """

    @abc.abstractmethod
    async def list_folders(self) -> "FolderListResponse":
        """
        List all folders/mailboxes for the account.

        Returns:
            FolderListResponse with list of folders and their metadata.
        """

    @abc.abstractmethod
    async def move_emails(
        self,
        email_ids: list[str],
        destination_folder: str,
        source_mailbox: str = "INBOX",
    ) -> "EmailMoveResponse":
        """
        Move emails to a destination folder.

        Args:
            email_ids: List of email UIDs to move.
            destination_folder: The target folder name.
            source_mailbox: The source mailbox (default: "INBOX").

        Returns:
            EmailMoveResponse with operation results.
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
    ) -> "EmailMoveResponse":
        """
        Remove a label from emails by deleting from the label folder.

        Args:
            email_ids: List of email UIDs to unlabel.
            label_name: The label name (without Labels/ prefix).

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
