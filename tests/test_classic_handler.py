from datetime import datetime, timezone
from email.message import EmailMessage
from unittest.mock import AsyncMock, patch

import pytest

from mcp_email_server.config import EmailServer, EmailSettings
from mcp_email_server.emails.classic import ClassicEmailHandler, EmailClient
from mcp_email_server.emails.models import (
    AttachmentDownloadResponse,
    EmailBodyResponse,
    EmailContentBatchResponse,
    EmailDeleteResponse,
    EmailMetadata,
    EmailMetadataPageResponse,
    EmailMoveResponse,
    MailboxInfo,
)


@pytest.fixture
def email_settings():
    return EmailSettings(
        account_name="test_account",
        full_name="Test User",
        email_address="test@example.com",
        incoming=EmailServer(
            user_name="test_user",
            password="test_password",
            host="imap.example.com",
            port=993,
            use_ssl=True,
        ),
        outgoing=EmailServer(
            user_name="test_user",
            password="test_password",
            host="smtp.example.com",
            port=465,
            use_ssl=True,
        ),
    )


@pytest.fixture
def classic_handler(email_settings):
    return ClassicEmailHandler(email_settings)


class TestClassicEmailHandler:
    def test_init(self, email_settings):
        """Test initialization of ClassicEmailHandler."""
        handler = ClassicEmailHandler(email_settings)

        assert handler.email_settings == email_settings
        assert isinstance(handler.incoming_client, EmailClient)
        assert isinstance(handler.outgoing_client, EmailClient)

        # Check that clients are initialized correctly
        assert handler.incoming_client.email_server == email_settings.incoming
        assert handler.outgoing_client.email_server == email_settings.outgoing
        assert handler.outgoing_client.sender == f"{email_settings.full_name} <{email_settings.email_address}>"

    def test_init_read_only_account(self):
        """Read-only accounts initialize without an outgoing SMTP client."""
        email_settings = EmailSettings(
            account_name="read_only",
            full_name="Read Only",
            email_address="read-only@example.com",
            incoming=EmailServer(
                user_name="reader",
                password="secret",
                host="imap.example.com",
                port=993,
                use_ssl=True,
            ),
        )

        handler = ClassicEmailHandler(email_settings)

        assert isinstance(handler.incoming_client, EmailClient)
        assert handler.outgoing_client is None

    @pytest.mark.asyncio
    async def test_get_emails(self, classic_handler):
        """Test get_emails method."""
        # Create test data
        now = datetime.now(timezone.utc)
        email_data = {
            "email_id": "123",
            "subject": "Test Subject",
            "from": "sender@example.com",
            "to": ["recipient@example.com"],
            "date": now,
            "attachments": [],
        }

        # Mock get_emails_metadata_page to return (list, total) in a single call
        mock_page = AsyncMock(return_value=([email_data], 1))

        with patch.object(classic_handler.incoming_client, "get_emails_metadata_page", mock_page):
            result = await classic_handler.get_emails_metadata(
                page=1,
                page_size=10,
                before=now,
                since=None,
                subject="Test",
                from_address="sender@example.com",
                to_address=None,
            )

            # Verify the result
            assert isinstance(result, EmailMetadataPageResponse)
            assert result.page == 1
            assert result.page_size == 10
            assert result.before == now
            assert result.since is None
            assert result.subject == "Test"
            assert len(result.emails) == 1
            assert isinstance(result.emails[0], EmailMetadata)
            assert result.emails[0].subject == "Test Subject"
            assert result.emails[0].sender == "sender@example.com"
            assert result.emails[0].date == now
            assert result.emails[0].attachments == []
            assert result.total == 1

            mock_page.assert_called_once_with(
                page=1,
                page_size=10,
                before=now,
                since=None,
                subject="Test",
                from_address="sender@example.com",
                to_address=None,
                order="desc",
                mailbox="INBOX",
                seen=None,
                flagged=None,
                answered=None,
                body=None,
                text=None,
                has_attachment=None,
                allowed_senders=[],
            )

    @pytest.mark.asyncio
    async def test_get_emails_with_mailbox(self, classic_handler):
        """Test get_emails method with custom mailbox."""
        now = datetime.now(timezone.utc)
        email_data = {
            "email_id": "456",
            "subject": "Sent Mail Subject",
            "from": "me@example.com",
            "to": ["recipient@example.com"],
            "date": now,
            "attachments": [],
        }

        mock_page = AsyncMock(return_value=([email_data], 1))

        with patch.object(classic_handler.incoming_client, "get_emails_metadata_page", mock_page):
            result = await classic_handler.get_emails_metadata(
                page=1,
                page_size=10,
                mailbox="Sent",
            )

            assert isinstance(result, EmailMetadataPageResponse)
            assert len(result.emails) == 1

            mock_page.assert_called_once_with(
                page=1,
                page_size=10,
                before=None,
                since=None,
                subject=None,
                from_address=None,
                to_address=None,
                order="desc",
                mailbox="Sent",
                seen=None,
                flagged=None,
                answered=None,
                body=None,
                text=None,
                has_attachment=None,
                allowed_senders=[],
            )

    @pytest.mark.asyncio
    async def test_send_email(self, classic_handler):
        """Test send_email method."""
        # Mock the outgoing_client.send_email method
        mock_send = AsyncMock()

        # Apply the mock
        with patch.object(classic_handler.outgoing_client, "send_email", mock_send):
            # Call the method
            await classic_handler.send_email(
                recipients=["recipient@example.com"],
                subject="Test Subject",
                body="Test Body",
                cc=["cc@example.com"],
                bcc=["bcc@example.com"],
            )

            # Verify the client method was called correctly
            mock_send.assert_called_once_with(
                ["recipient@example.com"],
                "Test Subject",
                "Test Body",
                ["cc@example.com"],
                ["bcc@example.com"],
                False,
                None,
                None,
                None,
                None,
            )

    @pytest.mark.asyncio
    async def test_send_email_with_attachments(self, classic_handler, tmp_path):
        """Test send_email method with attachments."""
        # Create a temporary test file
        test_file = tmp_path / "test_attachment.txt"
        test_file.write_text("This is a test attachment")

        # Mock the outgoing_client.send_email method
        mock_send = AsyncMock()

        # Apply the mock
        with patch.object(classic_handler.outgoing_client, "send_email", mock_send):
            # Call the method with attachments
            await classic_handler.send_email(
                recipients=["recipient@example.com"],
                subject="Test Subject",
                body="Test Body with attachment",
                attachments=[str(test_file)],
            )

            # Verify the client method was called correctly with attachments
            mock_send.assert_called_once_with(
                ["recipient@example.com"],
                "Test Subject",
                "Test Body with attachment",
                None,
                None,
                False,
                [str(test_file)],
                None,
                None,
                None,
            )

    @pytest.mark.asyncio
    async def test_read_only_account_rejects_send_email(self):
        """Read-only accounts cannot send email."""
        email_settings = EmailSettings(
            account_name="read_only",
            full_name="Read Only",
            email_address="read-only@example.com",
            incoming=EmailServer(
                user_name="reader",
                password="secret",
                host="imap.example.com",
                port=993,
                use_ssl=True,
            ),
        )
        handler = ClassicEmailHandler(email_settings)

        with pytest.raises(RuntimeError, match="SMTP is not configured"):
            await handler.send_email(
                recipients=["recipient@example.com"],
                subject="Test Subject",
                body="Test Body",
            )

    @pytest.mark.asyncio
    async def test_read_only_account_rejects_save_to_mailbox(self):
        """Read-only accounts cannot compose and save outbound drafts."""
        email_settings = EmailSettings(
            account_name="read_only",
            full_name="Read Only",
            email_address="read-only@example.com",
            incoming=EmailServer(
                user_name="reader",
                password="secret",
                host="imap.example.com",
                port=993,
                use_ssl=True,
            ),
        )
        handler = ClassicEmailHandler(email_settings)

        with pytest.raises(RuntimeError, match="SMTP is not configured"):
            await handler.save_to_mailbox(
                recipients=["recipient@example.com"],
                subject="Test Subject",
                body="Test Body",
            )

    @pytest.mark.asyncio
    async def test_delete_emails(self, classic_handler):
        """Test delete_emails permanently deletes when no Trash folder is detected."""
        mock_list = AsyncMock(return_value=[MailboxInfo(name="INBOX", delimiter="/", flags=[])])
        mock_delete = AsyncMock(return_value=(["123", "456"], []))

        with (
            patch.object(classic_handler.incoming_client, "list_mailboxes", mock_list),
            patch.object(classic_handler.incoming_client, "delete_emails", mock_delete),
        ):
            result = await classic_handler.delete_emails(
                email_ids=["123", "456"],
                mailbox="INBOX",
            )

            assert isinstance(result, EmailDeleteResponse)
            assert result.success is True
            assert result.deleted_ids == ["123", "456"]
            assert result.failed_ids == []
            assert result.mailbox == "INBOX"
            assert result.destination is None
            mock_delete.assert_called_once_with(["123", "456"], "INBOX")

    @pytest.mark.asyncio
    async def test_delete_emails_with_failures(self, classic_handler):
        """Test delete_emails permanently deletes from Trash with some failures."""
        mock_list = AsyncMock(return_value=[MailboxInfo(name="Trash", delimiter="/", flags=["\\Trash"])])
        mock_delete = AsyncMock(return_value=(["123"], ["456"]))

        with (
            patch.object(classic_handler.incoming_client, "list_mailboxes", mock_list),
            patch.object(classic_handler.incoming_client, "delete_emails", mock_delete),
        ):
            result = await classic_handler.delete_emails(
                email_ids=["123", "456"],
                mailbox="Trash",
            )

            assert isinstance(result, EmailDeleteResponse)
            assert result.success is False
            assert result.deleted_ids == ["123"]
            assert result.failed_ids == ["456"]
            assert result.destination is None
            mock_delete.assert_called_once_with(["123", "456"], "Trash")

    @pytest.mark.asyncio
    async def test_delete_emails_custom_mailbox(self, classic_handler):
        """Test delete_emails permanently deletes from custom mailbox when no Trash detected."""
        mock_list = AsyncMock(return_value=[MailboxInfo(name="INBOX", delimiter="/", flags=[])])
        mock_delete = AsyncMock(return_value=(["789"], []))

        with (
            patch.object(classic_handler.incoming_client, "list_mailboxes", mock_list),
            patch.object(classic_handler.incoming_client, "delete_emails", mock_delete),
        ):
            result = await classic_handler.delete_emails(
                email_ids=["789"],
                mailbox="Archive",
            )

            assert isinstance(result, EmailDeleteResponse)
            assert result.success is True
            assert result.deleted_ids == ["789"]
            assert result.failed_ids == []
            assert result.mailbox == "Archive"
            assert result.destination is None
            mock_delete.assert_called_once_with(["789"], "Archive")

    @pytest.mark.asyncio
    async def test_mark_emails(self, classic_handler):
        """Test mark_emails method."""
        mock_mark = AsyncMock(return_value=(["123", "456"], []))

        with patch.object(classic_handler.incoming_client, "mark_emails", mock_mark):
            result = await classic_handler.mark_emails(
                email_ids=["123", "456"],
                mark_as="read",
                mailbox="INBOX",
            )

            assert result.success is True
            assert result.marked_ids == ["123", "456"]
            assert result.failed_ids == []
            assert result.mailbox == "INBOX"
            assert result.marked_as == "read"
            mock_mark.assert_called_once_with(["123", "456"], "read", "INBOX", allowed_senders=[], report_blocked_mutations=False)

    @pytest.mark.asyncio
    async def test_mark_emails_with_failures(self, classic_handler):
        """Test mark_emails method with some failures."""
        mock_mark = AsyncMock(return_value=(["123"], ["456"]))

        with patch.object(classic_handler.incoming_client, "mark_emails", mock_mark):
            result = await classic_handler.mark_emails(
                email_ids=["123", "456"],
                mark_as="unread",
                mailbox="INBOX",
            )

            assert result.success is False
            assert result.marked_ids == ["123"]
            assert result.failed_ids == ["456"]
            assert result.marked_as == "unread"

    @pytest.mark.asyncio
    async def test_download_attachment(self, classic_handler, tmp_path):
        """Test download_attachment method."""
        save_path = str(tmp_path / "downloaded_attachment.pdf")

        mock_result = {
            "email_id": "123",
            "attachment_name": "document.pdf",
            "mime_type": "application/pdf",
            "size": 1024,
            "saved_path": save_path,
        }

        mock_download = AsyncMock(return_value=mock_result)

        with patch.object(classic_handler.incoming_client, "download_attachment", mock_download):
            result = await classic_handler.download_attachment(
                email_id="123",
                attachment_name="document.pdf",
                save_path=save_path,
            )

            assert isinstance(result, AttachmentDownloadResponse)
            assert result.email_id == "123"
            assert result.attachment_name == "document.pdf"
            assert result.mime_type == "application/pdf"
            assert result.size == 1024
            assert result.saved_path == save_path

            mock_download.assert_called_once_with("123", "document.pdf", save_path, "INBOX", allowed_senders=[])

    @pytest.mark.asyncio
    async def test_send_email_with_reply_headers(self, classic_handler):
        """Test sending email with reply headers."""
        mock_smtp = AsyncMock()
        mock_smtp.__aenter__.return_value = mock_smtp
        mock_smtp.__aexit__.return_value = None
        mock_smtp.login = AsyncMock()
        mock_smtp.send_message = AsyncMock()

        with patch("aiosmtplib.SMTP", return_value=mock_smtp):
            await classic_handler.send_email(
                recipients=["recipient@example.com"],
                subject="Re: Test",
                body="Reply body",
                in_reply_to="<original@example.com>",
                references="<original@example.com>",
            )

            call_args = mock_smtp.send_message.call_args
            msg = call_args[0][0]
            assert msg["In-Reply-To"] == "<original@example.com>"
            assert msg["References"] == "<original@example.com>"

    @pytest.mark.asyncio
    async def test_send_email_with_reply_to(self, classic_handler):
        """Test sending email with Reply-To header."""
        mock_smtp = AsyncMock()
        mock_smtp.__aenter__.return_value = mock_smtp
        mock_smtp.__aexit__.return_value = None
        mock_smtp.login = AsyncMock()
        mock_smtp.send_message = AsyncMock()

        with patch("aiosmtplib.SMTP", return_value=mock_smtp):
            await classic_handler.send_email(
                recipients=["recipient@example.com"],
                subject="Test",
                body="Body",
                reply_to="replyhere@example.com",
            )

            call_args = mock_smtp.send_message.call_args
            msg = call_args[0][0]
            assert msg["Reply-To"] == "replyhere@example.com"

    @pytest.mark.asyncio
    async def test_get_emails_content_includes_message_id(self, classic_handler):
        """Test that get_emails_content returns message_id from parsed email data."""
        now = datetime.now(timezone.utc)
        email_data = {
            "email_id": "123",
            "message_id": "<test-message-id@example.com>",
            "subject": "Test Subject",
            "from": "sender@example.com",
            "to": ["recipient@example.com"],
            "date": now,
            "body": "Test email body",
            "attachments": [],
        }

        # Mock the get_email_body_by_id method to return our test data
        mock_get_body = AsyncMock(return_value=email_data)

        with patch.object(classic_handler.incoming_client, "get_email_body_by_id", mock_get_body):
            result = await classic_handler.get_emails_content(
                email_ids=["123"],
                mailbox="INBOX",
            )

            # Verify the result
            assert isinstance(result, EmailContentBatchResponse)
            assert len(result.emails) == 1
            assert isinstance(result.emails[0], EmailBodyResponse)
            assert result.emails[0].email_id == "123"
            assert result.emails[0].message_id == "<test-message-id@example.com>"
            assert result.emails[0].subject == "Test Subject"
            assert result.emails[0].sender == "sender@example.com"
            assert result.emails[0].body == "Test email body"

            # Verify the client method was called correctly
            mock_get_body.assert_called_once_with("123", "INBOX", allowed_senders=[], body_offset=0, max_body_length=20000)

    @pytest.mark.asyncio
    async def test_get_emails_content_mark_as_read_true(self, classic_handler):
        """Test that mark_as_read=True calls mark_emails on successful fetches."""
        now = datetime.now(timezone.utc)
        email_data = {
            "email_id": "123",
            "message_id": "<test@example.com>",
            "subject": "Test",
            "from": "sender@example.com",
            "to": ["recipient@example.com"],
            "date": now,
            "body": "Test body",
            "attachments": [],
        }

        mock_get_body = AsyncMock(return_value=email_data)
        mock_mark = AsyncMock(return_value=(["123"], []))

        with (
            patch.object(classic_handler.incoming_client, "get_email_body_by_id", mock_get_body),
            patch.object(classic_handler.incoming_client, "mark_emails", mock_mark),
        ):
            result = await classic_handler.get_emails_content(
                email_ids=["123"],
                mailbox="INBOX",
                mark_as_read=True,
            )

            assert result.retrieved_count == 1
            mock_mark.assert_called_once_with(["123"], "read", "INBOX")

    @pytest.mark.asyncio
    async def test_get_emails_content_mark_as_read_false(self, classic_handler):
        """Test that mark_as_read=False (default) does not call mark_emails."""
        now = datetime.now(timezone.utc)
        email_data = {
            "email_id": "123",
            "message_id": "<test@example.com>",
            "subject": "Test",
            "from": "sender@example.com",
            "to": ["recipient@example.com"],
            "date": now,
            "body": "Test body",
            "attachments": [],
        }

        mock_get_body = AsyncMock(return_value=email_data)
        mock_mark = AsyncMock()

        with (
            patch.object(classic_handler.incoming_client, "get_email_body_by_id", mock_get_body),
            patch.object(classic_handler.incoming_client, "mark_emails", mock_mark),
        ):
            result = await classic_handler.get_emails_content(
                email_ids=["123"],
                mailbox="INBOX",
            )

            assert result.retrieved_count == 1
            mock_mark.assert_not_called()

    @pytest.mark.asyncio
    async def test_get_emails_content_returns_none(self, classic_handler):
        """Test get_emails_content handles None response (covers 1107-1108)."""
        # Mock the get_email_body_by_id method to return None
        mock_get_body = AsyncMock(return_value=None)

        with patch.object(classic_handler.incoming_client, "get_email_body_by_id", mock_get_body):
            result = await classic_handler.get_emails_content(
                email_ids=["123"],
                mailbox="INBOX",
            )

            # Verify the result
            assert isinstance(result, EmailContentBatchResponse)
            assert len(result.emails) == 0
            assert result.requested_count == 1
            assert result.retrieved_count == 0
            assert result.failed_ids == ["123"]

    @pytest.mark.asyncio
    async def test_get_emails_content_exception(self, classic_handler):
        """Test get_emails_content handles exception (covers 1109-1111)."""
        # Mock the get_email_body_by_id method to raise an exception
        mock_get_body = AsyncMock(side_effect=Exception("Connection error"))

        with patch.object(classic_handler.incoming_client, "get_email_body_by_id", mock_get_body):
            result = await classic_handler.get_emails_content(
                email_ids=["123", "456"],
                mailbox="INBOX",
            )

            # Verify the result - both emails should fail
            assert isinstance(result, EmailContentBatchResponse)
            assert len(result.emails) == 0
            assert result.requested_count == 2
            assert result.retrieved_count == 0
            assert result.failed_ids == ["123", "456"]


class TestEmailClientGetEmailBodyById:
    """Test EmailClient.get_email_body_by_id read-state behavior."""

    @pytest.fixture
    def email_client(self, email_settings):
        return EmailClient(email_settings.incoming)

    @staticmethod
    def _raw_email() -> bytes:
        msg = EmailMessage()
        msg["Subject"] = "Test Subject"
        msg["From"] = "sender@example.com"
        msg["To"] = "recipient@example.com"
        msg["Date"] = "Tue, 26 May 2026 04:30:00 +0000"
        msg["Message-ID"] = "<test@example.com>"
        msg.set_content("Test body")
        return msg.as_bytes()

    @pytest.mark.asyncio
    async def test_get_email_body_by_id_uses_peek_fetch_by_default(self, email_client, mock_imap):
        """Test default retrieval uses non-mutating PEEK fetch and does not STORE \\Seen."""
        mock_imap.uid = AsyncMock(return_value=("OK", [b"FETCH BODY[]", bytearray(self._raw_email())]))

        with patch.object(email_client, "_imap_connect", return_value=mock_imap):
            result = await email_client.get_email_body_by_id("123")

        assert result is not None
        assert result["email_id"] == "123"
        mock_imap.uid.assert_called_once_with("fetch", "123", "BODY.PEEK[]")

    @pytest.mark.asyncio
    async def test_get_email_body_by_id_marks_as_read_after_successful_parse(self, email_client, mock_imap):
        """Test mark_as_read=True stores \\Seen after a successful parse."""
        mock_imap.uid = AsyncMock(
            side_effect=[
                ("OK", [b"FETCH BODY[]", bytearray(self._raw_email())]),
                ("OK", [b"STORE +FLAGS (\\Seen)"]),
            ]
        )

        with patch.object(email_client, "_imap_connect", return_value=mock_imap):
            result = await email_client.get_email_body_by_id("123", mark_as_read=True)

        assert result is not None
        assert mock_imap.uid.call_args_list[0].args == ("fetch", "123", "BODY.PEEK[]")
        assert mock_imap.uid.call_args_list[1].args == ("store", "123", "+FLAGS", r"(\Seen)")

    @pytest.mark.asyncio
    async def test_get_email_body_by_id_does_not_mark_as_read_when_parse_fails(self, email_client, mock_imap):
        """Test failed parsing skips the \\Seen STORE side effect."""
        mock_imap.uid = AsyncMock(return_value=("OK", [b"FETCH BODY[]", bytearray(self._raw_email())]))

        with patch.object(email_client, "_imap_connect", return_value=mock_imap):
            with patch.object(email_client, "_parse_email_data", side_effect=ValueError("parse failed")):
                result = await email_client.get_email_body_by_id("123", mark_as_read=True)

        assert result is None
        mock_imap.uid.assert_called_once_with("fetch", "123", "BODY.PEEK[]")

    @pytest.mark.asyncio
    async def test_get_email_body_by_id_continues_when_mark_as_read_store_fails(self, email_client, mock_imap):
        """Test STORE failure is logged while retrieval still succeeds."""
        mock_imap.uid = AsyncMock(
            side_effect=[
                ("OK", [b"FETCH BODY[]", bytearray(self._raw_email())]),
                ("NO", [b"STORE failed"]),
            ]
        )

        with patch.object(email_client, "_imap_connect", return_value=mock_imap):
            result = await email_client.get_email_body_by_id("123", mark_as_read=True)

        assert result is not None
        assert result["email_id"] == "123"
        assert mock_imap.uid.call_count == 2



class TestEmailClientBatchMethods:
    """Test batch fetch methods for performance optimization."""

    @pytest.fixture
    def email_client(self, email_settings):
        return EmailClient(email_settings.incoming)

    def test_parse_headers(self, email_client):
        """Test _parse_headers method parses email headers correctly."""
        raw_headers = b"""From: sender@example.com
To: recipient@example.com
Cc: cc@example.com
Subject: Test Subject
Date: Mon, 20 Jan 2025 10:30:00 +0000

"""
        result = email_client._parse_headers("123", raw_headers)

        assert result is not None
        assert result["email_id"] == "123"
        assert result["subject"] == "Test Subject"
        assert result["from"] == "sender@example.com"
        assert "recipient@example.com" in result["to"]
        assert "cc@example.com" in result["to"]
        assert result["attachments"] == []

    def test_parse_headers_with_invalid_data(self, email_client):
        """Test _parse_headers handles malformed headers gracefully."""
        # Completely broken data that can't be parsed
        raw_headers = b"\xff\xfe\x00\x00"
        result = email_client._parse_headers("123", raw_headers)

        # Should return None or a valid dict with fallback values
        # The implementation catches exceptions and returns None
        assert result is None or isinstance(result, dict)

    def test_parse_headers_missing_date(self, email_client):
        """Test _parse_headers handles missing date with fallback."""
        raw_headers = b"""From: sender@example.com
To: recipient@example.com
Subject: No Date Email

"""
        result = email_client._parse_headers("123", raw_headers)

        assert result is not None
        assert result["email_id"] == "123"
        assert result["date"] is not None  # Should have fallback to now()

    @pytest.mark.asyncio
    async def test_batch_fetch_dates_empty_list(self, email_client):
        """Test _batch_fetch_dates with empty list returns empty dict."""
        mock_imap = AsyncMock()
        result = await email_client._batch_fetch_dates(mock_imap, [])

        assert result == {}
        mock_imap.uid.assert_not_called()

    @pytest.mark.asyncio
    async def test_batch_fetch_headers_empty_list(self, email_client):
        """Test _batch_fetch_headers with empty list returns empty dict."""
        mock_imap = AsyncMock()
        result = await email_client._batch_fetch_headers(mock_imap, [])

        assert result == {}
        mock_imap.uid.assert_not_called()

    @pytest.mark.asyncio
    async def test_batch_fetch_dates_parses_response(self, email_client):
        """Test _batch_fetch_dates correctly parses IMAP INTERNALDATE response."""
        mock_imap = AsyncMock()
        # Simulate IMAP response format for INTERNALDATE
        mock_imap.uid.return_value = (
            "OK",
            [
                b'1 FETCH (UID 100 INTERNALDATE "20-Jan-2025 10:30:00 +0000")',
                b'2 FETCH (UID 101 INTERNALDATE "21-Jan-2025 11:00:00 +0000")',
            ],
        )

        result = await email_client._batch_fetch_dates(mock_imap, [b"100", b"101"])

        assert "100" in result
        assert "101" in result
        assert result["100"].day == 20
        assert result["101"].day == 21

    @pytest.mark.asyncio
    async def test_batch_fetch_headers_parses_response(self, email_client):
        """Test _batch_fetch_headers correctly parses IMAP BODY[HEADER] response."""
        mock_imap = AsyncMock()
        # aioimaplib returns FETCH response in 3 parts:
        # - BODY[HEADER] line (no UID)
        # - header content as bytearray
        # - UID line
        mock_imap.uid.return_value = (
            "OK",
            [
                b"1 FETCH (BODY[HEADER] {100}",
                bytearray(b"From: sender@example.com\r\nTo: recipient@example.com\r\nSubject: Test\r\n\r\n"),
                b" UID 100)",
            ],
        )

        result = await email_client._batch_fetch_headers(mock_imap, ["100"])

        assert "100" in result
        assert result["100"]["subject"] == "Test"
        assert result["100"]["from"] == "sender@example.com"


class TestFindSpecialFolder:
    """Test _find_special_folder for RFC 6154 flag and fallback detection."""

    @pytest.mark.asyncio
    async def test_finds_folder_by_flag(self, classic_handler):
        """Test finding a folder by its RFC 6154 flag."""
        folders = [
            MailboxInfo(name="INBOX", delimiter="/", flags=["\\HasNoChildren"]),
            MailboxInfo(name="Bin", delimiter="/", flags=["\\Trash", "\\HasNoChildren"]),
        ]
        mock_list = AsyncMock(return_value=folders)

        with patch.object(classic_handler.incoming_client, "list_mailboxes", mock_list):
            result = await classic_handler._find_special_folder("\\Trash", ["Trash", "Deleted Items"])

        assert result == "Bin"

    @pytest.mark.asyncio
    async def test_falls_back_to_common_name(self, classic_handler):
        """Test fallback to common folder names when flag not found."""
        folders = [
            MailboxInfo(name="INBOX", delimiter="/", flags=["\\HasNoChildren"]),
            MailboxInfo(name="Trash", delimiter="/", flags=["\\HasNoChildren"]),
        ]
        mock_list = AsyncMock(return_value=folders)

        with patch.object(classic_handler.incoming_client, "list_mailboxes", mock_list):
            result = await classic_handler._find_special_folder("\\Trash", ["Trash", "Deleted Items"])

        assert result == "Trash"

    @pytest.mark.asyncio
    async def test_returns_none_when_not_found(self, classic_handler):
        """Test returns None when no matching folder exists."""
        folders = [
            MailboxInfo(name="INBOX", delimiter="/", flags=["\\HasNoChildren"]),
            MailboxInfo(name="Sent", delimiter="/", flags=["\\Sent"]),
        ]
        mock_list = AsyncMock(return_value=folders)

        with patch.object(classic_handler.incoming_client, "list_mailboxes", mock_list):
            result = await classic_handler._find_special_folder("\\Trash", ["Trash", "Deleted Items"])

        assert result is None

    @pytest.mark.asyncio
    async def test_caches_positive_result(self, classic_handler):
        """Test that positive results are cached."""
        folders = [MailboxInfo(name="Trash", delimiter="/", flags=["\\Trash"])]
        mock_list = AsyncMock(return_value=folders)

        with patch.object(classic_handler.incoming_client, "list_mailboxes", mock_list):
            result1 = await classic_handler._find_special_folder("\\Trash", ["Trash"])
            result2 = await classic_handler._find_special_folder("\\Trash", ["Trash"])

        assert result1 == "Trash"
        assert result2 == "Trash"
        mock_list.assert_called_once()  # Only one LIST call due to caching

    @pytest.mark.asyncio
    async def test_caches_negative_result(self, classic_handler):
        """Test that negative results are cached."""
        folders = [MailboxInfo(name="INBOX", delimiter="/", flags=[])]
        mock_list = AsyncMock(return_value=folders)

        with patch.object(classic_handler.incoming_client, "list_mailboxes", mock_list):
            result1 = await classic_handler._find_special_folder("\\Archive", ["Archive"])
            result2 = await classic_handler._find_special_folder("\\Archive", ["Archive"])

        assert result1 is None
        assert result2 is None
        mock_list.assert_called_once()


class TestDeleteEmailsSafeDelete:
    """Test delete_emails moves to Trash when available."""

    @pytest.mark.asyncio
    async def test_moves_to_trash_when_found(self, classic_handler):
        """Test delete_emails moves to Trash folder when it exists."""
        folders = [
            MailboxInfo(name="INBOX", delimiter="/", flags=[]),
            MailboxInfo(name="Trash", delimiter="/", flags=["\\Trash"]),
        ]
        mock_list = AsyncMock(return_value=folders)
        mock_move = AsyncMock(return_value=(["123", "456"], []))

        with (
            patch.object(classic_handler.incoming_client, "list_mailboxes", mock_list),
            patch.object(classic_handler.incoming_client, "move_emails", mock_move),
        ):
            result = await classic_handler.delete_emails(["123", "456"], "INBOX")

        assert isinstance(result, EmailDeleteResponse)
        assert result.success is True
        assert result.deleted_ids == ["123", "456"]
        assert result.failed_ids == []
        assert result.mailbox == "INBOX"
        assert result.destination == "Trash"
        mock_move.assert_called_once_with(["123", "456"], "INBOX", "Trash")

    @pytest.mark.asyncio
    async def test_permanent_delete_when_no_trash(self, classic_handler):
        """Test delete_emails permanently deletes when Trash folder not found."""
        folders = [MailboxInfo(name="INBOX", delimiter="/", flags=[])]
        mock_list = AsyncMock(return_value=folders)
        mock_delete = AsyncMock(return_value=(["123"], []))

        with (
            patch.object(classic_handler.incoming_client, "list_mailboxes", mock_list),
            patch.object(classic_handler.incoming_client, "delete_emails", mock_delete),
        ):
            result = await classic_handler.delete_emails(["123"], "INBOX")

        assert isinstance(result, EmailDeleteResponse)
        assert result.success is True
        assert result.deleted_ids == ["123"]
        assert result.destination is None
        mock_delete.assert_called_once_with(["123"], "INBOX")

    @pytest.mark.asyncio
    async def test_permanent_delete_when_already_in_trash(self, classic_handler):
        """Test delete_emails permanently deletes when already in Trash."""
        folders = [MailboxInfo(name="Trash", delimiter="/", flags=["\\Trash"])]
        mock_list = AsyncMock(return_value=folders)
        mock_delete = AsyncMock(return_value=(["123"], []))

        with (
            patch.object(classic_handler.incoming_client, "list_mailboxes", mock_list),
            patch.object(classic_handler.incoming_client, "delete_emails", mock_delete),
        ):
            result = await classic_handler.delete_emails(["123"], "Trash")

        assert isinstance(result, EmailDeleteResponse)
        assert result.success is True
        assert result.destination is None
        mock_delete.assert_called_once_with(["123"], "Trash")

    @pytest.mark.asyncio
    async def test_move_to_trash_with_failures(self, classic_handler):
        """Test delete_emails reports failures when moving to Trash."""
        folders = [MailboxInfo(name="Trash", delimiter="/", flags=["\\Trash"])]
        mock_list = AsyncMock(return_value=folders)
        mock_move = AsyncMock(return_value=(["123"], ["456"]))

        with (
            patch.object(classic_handler.incoming_client, "list_mailboxes", mock_list),
            patch.object(classic_handler.incoming_client, "move_emails", mock_move),
        ):
            result = await classic_handler.delete_emails(["123", "456"], "INBOX")

        assert result.success is False
        assert result.deleted_ids == ["123"]
        assert result.failed_ids == ["456"]
        assert result.destination == "Trash"


class TestArchiveEmails:
    """Test archive_emails moves to Archive folder."""

    @pytest.mark.asyncio
    async def test_archives_when_folder_found(self, classic_handler):
        """Test archive_emails moves to Archive folder when found."""
        folders = [
            MailboxInfo(name="INBOX", delimiter="/", flags=[]),
            MailboxInfo(name="Archive", delimiter="/", flags=["\\Archive"]),
        ]
        mock_list = AsyncMock(return_value=folders)
        mock_move = AsyncMock(return_value=(["123", "456"], []))

        with (
            patch.object(classic_handler.incoming_client, "list_mailboxes", mock_list),
            patch.object(classic_handler.incoming_client, "move_emails", mock_move),
        ):
            result = await classic_handler.archive_emails(["123", "456"], "INBOX")

        assert isinstance(result, EmailMoveResponse)
        assert result.success is True
        assert result.moved_ids == ["123", "456"]
        assert result.source_mailbox == "INBOX"
        assert result.destination_folder == "Archive"
        mock_move.assert_called_once_with(["123", "456"], "INBOX", "Archive")

    @pytest.mark.asyncio
    async def test_archives_with_gmail_all_mail(self, classic_handler):
        """Test archive_emails finds [Gmail]/All Mail by fallback name."""
        folders = [
            MailboxInfo(name="INBOX", delimiter="/", flags=[]),
            MailboxInfo(name="[Gmail]/All Mail", delimiter="/", flags=["\\All"]),
        ]
        mock_list = AsyncMock(return_value=folders)
        mock_move = AsyncMock(return_value=(["123"], []))

        with (
            patch.object(classic_handler.incoming_client, "list_mailboxes", mock_list),
            patch.object(classic_handler.incoming_client, "move_emails", mock_move),
        ):
            result = await classic_handler.archive_emails(["123"], "INBOX")

        assert result.success is True
        assert result.destination_folder == "[Gmail]/All Mail"

    @pytest.mark.asyncio
    async def test_raises_when_archive_not_found(self, classic_handler):
        """Test archive_emails raises ValueError when Archive folder not found."""
        folders = [MailboxInfo(name="INBOX", delimiter="/", flags=[])]
        mock_list = AsyncMock(return_value=folders)

        with patch.object(classic_handler.incoming_client, "list_mailboxes", mock_list):
            with pytest.raises(ValueError, match="Archive folder not found"):
                await classic_handler.archive_emails(["123"], "INBOX")

    @pytest.mark.asyncio
    async def test_archive_with_failures(self, classic_handler):
        """Test archive_emails reports partial failures."""
        folders = [MailboxInfo(name="Archive", delimiter="/", flags=["\\Archive"])]
        mock_list = AsyncMock(return_value=folders)
        mock_move = AsyncMock(return_value=(["123"], ["456"]))

        with (
            patch.object(classic_handler.incoming_client, "list_mailboxes", mock_list),
            patch.object(classic_handler.incoming_client, "move_emails", mock_move),
        ):
            result = await classic_handler.archive_emails(["123", "456"], "INBOX")

        assert result.success is False
        assert result.moved_ids == ["123"]
        assert result.failed_ids == ["456"]


def _make_header_entry(uid: str, subject: str = "", date: datetime | None = None) -> dict:
    """Create a minimal email header dict for testing."""
    return {
        "email_id": uid,
        "subject": subject or f"Email {uid}",
        "from": f"{uid}@test.com",
        "to": ["r@test.com"],
        "date": date or datetime.now(timezone.utc),
        "attachments": [],
    }


class TestSearchFallback:
    """Test UID ordering fallback when INTERNALDATE fetch fails."""

    @pytest.fixture
    def email_client(self, email_settings):
        return EmailClient(email_settings.incoming)

    @pytest.mark.asyncio
    async def test_uid_ordering_fallback_when_dates_empty(self, email_client):
        """When _batch_fetch_dates returns {}, UIDs should still reach header fetch via UID ordering."""
        mock_imap = AsyncMock()
        mock_imap.uid_search = AsyncMock(return_value=("OK", [b"100 200 300"]))

        header_data = {uid: _make_header_entry(uid) for uid in ("100", "200", "300")}

        with (
            patch.object(email_client, "_imap_connection") as mock_ctx,
            patch.object(email_client, "_batch_fetch_dates", return_value={}),
            patch.object(email_client, "_batch_fetch_headers", return_value=header_data) as mock_headers,
        ):
            mock_ctx.return_value.__aenter__ = AsyncMock(return_value=mock_imap)
            mock_ctx.return_value.__aexit__ = AsyncMock(return_value=False)

            results, total = await email_client.get_emails_metadata_page(
                page=1, page_size=10, order="asc",
            )

        assert total == 3
        assert len(results) == 3
        # ASC order: 100, 200, 300
        assert results[0]["email_id"] == "100"
        assert results[2]["email_id"] == "300"
        mock_headers.assert_called_once()

    @pytest.mark.asyncio
    async def test_uid_ordering_fallback_desc(self, email_client):
        """When dates are empty, desc ordering should use reverse UID sort."""
        mock_imap = AsyncMock()
        mock_imap.uid_search = AsyncMock(return_value=("OK", [b"100 200 300"]))

        header_data = {uid: _make_header_entry(uid) for uid in ("100", "200", "300")}

        with (
            patch.object(email_client, "_imap_connection") as mock_ctx,
            patch.object(email_client, "_batch_fetch_dates", return_value={}),
            patch.object(email_client, "_batch_fetch_headers", return_value=header_data),
        ):
            mock_ctx.return_value.__aenter__ = AsyncMock(return_value=mock_imap)
            mock_ctx.return_value.__aexit__ = AsyncMock(return_value=False)

            results, total = await email_client.get_emails_metadata_page(
                page=1, page_size=10, order="desc",
            )

        assert total == 3
        assert len(results) == 3
        # DESC order: 300, 200, 100
        assert results[0]["email_id"] == "300"
        assert results[2]["email_id"] == "100"

    @pytest.mark.asyncio
    async def test_partial_date_loss_still_returns_results(self, email_client):
        """When _batch_fetch_dates returns partial results, remaining UIDs (with dates) are returned normally."""
        mock_imap = AsyncMock()
        mock_imap.uid_search = AsyncMock(return_value=("OK", [b"100 200 300"]))

        date_100 = datetime(2025, 1, 20, tzinfo=timezone.utc)
        date_300 = datetime(2025, 1, 22, tzinfo=timezone.utc)

        # Only 2 out of 3 UIDs have dates
        partial_dates = {"100": date_100, "300": date_300}

        header_data = {
            "100": _make_header_entry("100", date=date_100),
            "300": _make_header_entry("300", date=date_300),
        }

        with (
            patch.object(email_client, "_imap_connection") as mock_ctx,
            patch.object(email_client, "_batch_fetch_dates", return_value=partial_dates),
            patch.object(email_client, "_batch_fetch_headers", return_value=header_data),
        ):
            mock_ctx.return_value.__aenter__ = AsyncMock(return_value=mock_imap)
            mock_ctx.return_value.__aexit__ = AsyncMock(return_value=False)

            results, total = await email_client.get_emails_metadata_page(
                page=1, page_size=10, order="desc",
            )

        assert total == 3
        # Only 2 UIDs had dates, so only 2 results (sorted by date desc: 300, 100)
        assert len(results) == 2
        assert results[0]["email_id"] == "300"
        assert results[1]["email_id"] == "100"

    @pytest.mark.asyncio
    async def test_header_fetch_loss_returns_partial_results(self, email_client):
        """When _batch_fetch_headers returns fewer results than requested, partial results are still returned."""
        mock_imap = AsyncMock()
        mock_imap.uid_search = AsyncMock(return_value=("OK", [b"100 200 300"]))

        # Only 2 out of 3 UIDs returned from header fetch
        header_data = {
            "100": _make_header_entry("100"),
            "300": _make_header_entry("300"),
        }

        with (
            patch.object(email_client, "_imap_connection") as mock_ctx,
            patch.object(email_client, "_batch_fetch_dates", return_value={}),
            patch.object(email_client, "_batch_fetch_headers", return_value=header_data),
        ):
            mock_ctx.return_value.__aenter__ = AsyncMock(return_value=mock_imap)
            mock_ctx.return_value.__aexit__ = AsyncMock(return_value=False)

            results, total = await email_client.get_emails_metadata_page(
                page=1, page_size=10, order="asc",
            )

        assert total == 3
        # Only 2 results because UID 200 was lost in header fetch
        assert len(results) == 2
        assert results[0]["email_id"] == "100"
        assert results[1]["email_id"] == "300"
