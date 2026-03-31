import asyncio
import email
import ssl
from datetime import datetime, timezone
from email.mime.text import MIMEText
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from mcp_email_server.config import EmailServer, EmailSettings
from mcp_email_server.emails.classic import (
    ClassicEmailHandler,
    EmailClient,
    _create_smtp_ssl_context,
    _detect_email_service,
    _format_forwarded_email_html,
    _format_quoted_reply_html,
)


@pytest.fixture
def email_server():
    return EmailServer(
        user_name="test_user",
        password="test_password",
        host="imap.example.com",
        port=993,
        use_ssl=True,
    )


@pytest.fixture
def email_client(email_server):
    return EmailClient(email_server, sender="Test User <test@example.com>")


class TestEmailClient:
    def test_init(self, email_server):
        """Test initialization of EmailClient."""
        client = EmailClient(email_server)
        assert client.email_server == email_server
        assert client.sender == email_server.user_name
        assert client.smtp_use_tls is True
        assert client.smtp_start_tls is False

        # Test with custom sender
        custom_sender = "Custom <custom@example.com>"
        client = EmailClient(email_server, sender=custom_sender)
        assert client.sender == custom_sender

    def test_parse_email_data_plain(self):
        """Test parsing plain text email."""
        # Create a simple plain text email
        msg = MIMEText("This is a test email body")
        msg["Subject"] = "Test Subject"
        msg["From"] = "sender@example.com"
        msg["To"] = "recipient@example.com"
        msg["Date"] = email.utils.formatdate()

        raw_email = msg.as_bytes()

        client = EmailClient(MagicMock())
        result = client._parse_email_data(raw_email)

        assert result["subject"] == "Test Subject"
        assert result["from"] == "sender@example.com"
        assert result["body"] == "This is a test email body"
        assert isinstance(result["date"], datetime)
        assert result["attachments"] == []

    def test_parse_email_data_with_attachments(self):
        """Test parsing email with attachments."""
        # This would require creating a multipart email with attachments
        # For simplicity, we'll mock the email parsing
        with patch("email.parser.BytesParser.parsebytes") as mock_parse:
            mock_email = MagicMock()
            mock_email.get.side_effect = lambda x, default=None: {
                "Subject": "Test Subject",
                "From": "sender@example.com",
                "Date": email.utils.formatdate(),
            }.get(x, default)
            mock_email.is_multipart.return_value = True

            # Mock parts
            text_part = MagicMock()
            text_part.get_content_type.return_value = "text/plain"
            text_part.get.return_value = ""  # Not an attachment
            text_part.get_payload.return_value = b"This is the email body"
            text_part.get_content_charset.return_value = "utf-8"

            attachment_part = MagicMock()
            attachment_part.get_content_type.return_value = "application/pdf"
            attachment_part.get.return_value = "attachment; filename=test.pdf"
            attachment_part.get_filename.return_value = "test.pdf"

            mock_email.walk.return_value = [text_part, attachment_part]
            mock_parse.return_value = mock_email

            client = EmailClient(MagicMock())
            result = client._parse_email_data(b"dummy email content")

            assert result["subject"] == "Test Subject"
            assert result["from"] == "sender@example.com"
            assert result["body"] == "This is the email body"
            assert isinstance(result["date"], datetime)
            assert result["attachments"] == ["test.pdf"]

    def test_build_search_criteria(self):
        """Test building search criteria for IMAP."""
        # Test with no criteria (should return ["ALL"])
        criteria = EmailClient._build_search_criteria()
        assert criteria == ["ALL"]

        # Test with before date
        before_date = datetime(2023, 1, 1, tzinfo=timezone.utc)
        criteria = EmailClient._build_search_criteria(before=before_date)
        assert criteria == ["BEFORE", "01-JAN-2023"]

        # Test with since date
        since_date = datetime(2023, 1, 1, tzinfo=timezone.utc)
        criteria = EmailClient._build_search_criteria(since=since_date)
        assert criteria == ["SINCE", "01-JAN-2023"]

        # Test with subject
        criteria = EmailClient._build_search_criteria(subject="Test")
        assert criteria == ["SUBJECT", "Test"]

        # Test with body
        criteria = EmailClient._build_search_criteria(body="Test")
        assert criteria == ["BODY", "Test"]

        # Test with text
        criteria = EmailClient._build_search_criteria(text="Test")
        assert criteria == ["TEXT", "Test"]

        # Test with from_address
        criteria = EmailClient._build_search_criteria(from_address="test@example.com")
        assert criteria == ["FROM", "test@example.com"]

        # Test with to_address
        criteria = EmailClient._build_search_criteria(to_address="test@example.com")
        assert criteria == ["TO", "test@example.com"]

        # Test with multiple criteria
        criteria = EmailClient._build_search_criteria(
            subject="Test", from_address="test@example.com", since=datetime(2023, 1, 1, tzinfo=timezone.utc)
        )
        assert criteria == ["SINCE", "01-JAN-2023", "SUBJECT", "Test", "FROM", "test@example.com"]

        # Test with seen=True (read emails)
        criteria = EmailClient._build_search_criteria(seen=True)
        assert criteria == ["SEEN"]

        # Test with seen=False (unread emails)
        criteria = EmailClient._build_search_criteria(seen=False)
        assert criteria == ["UNSEEN"]

        # Test with seen=None (all emails - no criteria added)
        criteria = EmailClient._build_search_criteria(seen=None)
        assert criteria == ["ALL"]

        # Test with flagged=True (starred emails)
        criteria = EmailClient._build_search_criteria(flagged=True)
        assert criteria == ["FLAGGED"]

        # Test with flagged=False (non-starred emails)
        criteria = EmailClient._build_search_criteria(flagged=False)
        assert criteria == ["UNFLAGGED"]

        # Test with answered=True (replied emails)
        criteria = EmailClient._build_search_criteria(answered=True)
        assert criteria == ["ANSWERED"]

        # Test with answered=False (not replied emails)
        criteria = EmailClient._build_search_criteria(answered=False)
        assert criteria == ["UNANSWERED"]

        # Test compound criteria: unread emails from a specific sender
        criteria = EmailClient._build_search_criteria(seen=False, from_address="sender@example.com")
        assert "UNSEEN" in criteria
        assert "FROM" in criteria
        assert "sender@example.com" in criteria

        # Test compound criteria: flagged and answered
        criteria = EmailClient._build_search_criteria(flagged=True, answered=True)
        assert "FLAGGED" in criteria
        assert "ANSWERED" in criteria

        # Test compound criteria: unread, flagged, from specific sender, with subject
        criteria = EmailClient._build_search_criteria(
            seen=False, flagged=True, from_address="test@example.com", subject="Important"
        )
        assert "UNSEEN" in criteria
        assert "FLAGGED" in criteria
        assert "FROM" in criteria
        assert "test@example.com" in criteria
        assert "SUBJECT" in criteria
        assert "Important" in criteria

    def test_build_search_criteria_multiword_subject(self):
        """Multi-word subjects must be quoted for IMAP."""
        criteria = EmailClient._build_search_criteria(subject="Meeting Notes")
        assert criteria == ["SUBJECT", '"Meeting Notes"']

    def test_build_search_criteria_multiword_from(self):
        """Multi-word from_address must be quoted for IMAP."""
        criteria = EmailClient._build_search_criteria(from_address="Alice Example")
        assert criteria == ["FROM", '"Alice Example"']

    def test_build_search_criteria_multiword_to(self):
        """Multi-word to_address must be quoted for IMAP."""
        criteria = EmailClient._build_search_criteria(to_address="Bob Smith")
        assert criteria == ["TO", '"Bob Smith"']

    def test_build_search_criteria_subject_with_embedded_quotes(self):
        """Embedded double quotes must be stripped (invalid in IMAP quoted strings)."""
        criteria = EmailClient._build_search_criteria(subject='He said "hello"')
        assert criteria == ["SUBJECT", '"He said hello"']

    @pytest.mark.asyncio
    async def test_get_emails_metadata_page(self, email_client):
        """Test getting emails page returns sorted, paginated results with total count."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.uid_search = AsyncMock(return_value=(None, [b"1 2 3"]))
        mock_imap.logout = AsyncMock()

        # Mock at the helper level - test behavior, not implementation
        mock_dates = {
            "1": datetime(2024, 1, 1, tzinfo=timezone.utc),
            "2": datetime(2024, 1, 2, tzinfo=timezone.utc),
            "3": datetime(2024, 1, 3, tzinfo=timezone.utc),
        }
        mock_metadata = {
            "1": {
                "email_id": "1",
                "subject": "Subject 1",
                "from": "a@test.com",
                "to": [],
                "date": datetime(2024, 1, 1, tzinfo=timezone.utc),
                "attachments": [],
            },
            "2": {
                "email_id": "2",
                "subject": "Subject 2",
                "from": "b@test.com",
                "to": [],
                "date": datetime(2024, 1, 2, tzinfo=timezone.utc),
                "attachments": [],
            },
            "3": {
                "email_id": "3",
                "subject": "Subject 3",
                "from": "c@test.com",
                "to": [],
                "date": datetime(2024, 1, 3, tzinfo=timezone.utc),
                "attachments": [],
            },
        }

        with patch.object(email_client, "_imap_connect", return_value=mock_imap):
            with patch.object(email_client, "_batch_fetch_dates", return_value=mock_dates) as mock_fetch_dates:
                with patch.object(
                    email_client, "_batch_fetch_headers", return_value=mock_metadata
                ) as mock_fetch_headers:
                    emails, total = await email_client.get_emails_metadata_page(page=1, page_size=10)

                    # Behavior: returns emails sorted by date desc (newest first)
                    assert len(emails) == 3
                    assert total == 3
                    assert emails[0]["subject"] == "Subject 3"
                    assert emails[1]["subject"] == "Subject 2"
                    assert emails[2]["subject"] == "Subject 1"

                    mock_imap.login.assert_called_once()
                    mock_imap.logout.assert_called_once()

                    # Verify helpers called with correct arguments
                    mock_fetch_dates.assert_called_once_with(mock_imap, [b"1", b"2", b"3"])
                    # Headers fetched for page UIDs in sorted order (desc by date)
                    mock_fetch_headers.assert_called_once_with(mock_imap, ["3", "2", "1"])

    @pytest.mark.asyncio
    async def test_get_email_count(self, email_client):
        """Test getting email count."""
        # Mock IMAP client
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.search = AsyncMock(return_value=(None, [b"1 2 3 4 5"]))
        mock_imap.uid_search = AsyncMock(return_value=(None, [b"1 2 3 4 5"]))
        mock_imap.logout = AsyncMock()

        # Mock IMAP class
        with patch.object(email_client, "_imap_connect", return_value=mock_imap):
            count = await email_client.get_email_count()

            assert count == 5

            # Verify IMAP methods were called correctly
            mock_imap.login.assert_called_once_with(
                email_client.email_server.user_name, email_client.email_server.password.get_secret_value()
            )
            mock_imap.select.assert_called_once_with('"INBOX"')
            mock_imap.uid_search.assert_called_once_with("ALL")
            mock_imap.logout.assert_called_once()

    @pytest.mark.asyncio
    async def test_send_email(self, email_client):
        """Test sending email."""
        # Mock SMTP client
        mock_smtp = AsyncMock()
        mock_smtp.__aenter__.return_value = mock_smtp
        mock_smtp.__aexit__.return_value = None
        mock_smtp.login = AsyncMock()
        mock_smtp.send_message = AsyncMock()

        with patch("aiosmtplib.SMTP", return_value=mock_smtp):
            await email_client.send_email(
                recipients=["recipient@example.com"],
                subject="Test Subject",
                body="Test Body",
                cc=["cc@example.com"],
                bcc=["bcc@example.com"],
            )

            # Verify SMTP methods were called correctly
            mock_smtp.login.assert_called_once_with(
                email_client.email_server.user_name, email_client.email_server.password.get_secret_value()
            )
            mock_smtp.send_message.assert_called_once()

            # Check that the message was constructed correctly
            call_args = mock_smtp.send_message.call_args
            msg = call_args[0][0]
            recipients = call_args[1]["recipients"]

            assert msg["Subject"] == "Test Subject"
            assert msg["From"] == email_client.sender
            assert msg["To"] == "recipient@example.com"
            assert msg["Cc"] == "cc@example.com"
            assert "Bcc" not in msg  # BCC should not be in headers

            # Check that all recipients are included in the SMTP call
            assert "recipient@example.com" in recipients
            assert "cc@example.com" in recipients
            assert "bcc@example.com" in recipients


class TestSendEmailMessageIdAndDate:
    @pytest.mark.asyncio
    async def test_send_email_sets_message_id_and_date(self, email_client):
        """Test that send_email sets Message-Id and Date headers."""
        mock_smtp = AsyncMock()
        mock_smtp.__aenter__.return_value = mock_smtp
        mock_smtp.__aexit__.return_value = None
        mock_smtp.login = AsyncMock()
        mock_smtp.send_message = AsyncMock()

        with patch("aiosmtplib.SMTP", return_value=mock_smtp):
            msg = await email_client.send_email(
                recipients=["recipient@example.com"],
                subject="Test Subject",
                body="Test Body",
            )

            assert msg["Message-Id"] is not None
            assert "@example.com>" in msg["Message-Id"]
            assert msg["Date"] is not None

    @pytest.mark.asyncio
    async def test_send_email_message_id_uses_sender_domain(self, email_server):
        """Test that Message-Id domain is extracted from the sender address."""
        client = EmailClient(email_server, sender="user@getsequel.app")
        mock_smtp = AsyncMock()
        mock_smtp.__aenter__.return_value = mock_smtp
        mock_smtp.__aexit__.return_value = None
        mock_smtp.login = AsyncMock()
        mock_smtp.send_message = AsyncMock()

        with patch("aiosmtplib.SMTP", return_value=mock_smtp):
            msg = await client.send_email(
                recipients=["recipient@example.com"],
                subject="Test",
                body="Body",
            )

            assert "@getsequel.app>" in msg["Message-Id"]

    @pytest.mark.asyncio
    async def test_send_email_same_message_on_smtp_and_return(self, email_client):
        """Test that the same msg object (with Message-Id) is sent and returned."""
        mock_smtp = AsyncMock()
        mock_smtp.__aenter__.return_value = mock_smtp
        mock_smtp.__aexit__.return_value = None
        mock_smtp.login = AsyncMock()
        mock_smtp.send_message = AsyncMock()

        with patch("aiosmtplib.SMTP", return_value=mock_smtp):
            returned_msg = await email_client.send_email(
                recipients=["recipient@example.com"],
                subject="Test",
                body="Body",
            )

            sent_msg = mock_smtp.send_message.call_args[0][0]
            assert sent_msg["Message-Id"] == returned_msg["Message-Id"]
            assert sent_msg["Date"] == returned_msg["Date"]


class TestParseEmailData:
    def test_parse_email_extracts_message_id(self, email_client):
        """Test that Message-ID header is extracted during parsing."""
        raw_email = b"""Message-ID: <test123@example.com>
From: sender@example.com
To: recipient@example.com
Subject: Test Subject
Date: Mon, 1 Jan 2024 12:00:00 +0000

Test body content
"""
        result = email_client._parse_email_data(raw_email, email_id="1")
        assert result["message_id"] == "<test123@example.com>"

    def test_parse_email_handles_missing_message_id(self, email_client):
        """Test graceful handling when Message-ID is missing."""
        raw_email = b"""From: sender@example.com
To: recipient@example.com
Subject: Test Subject
Date: Mon, 1 Jan 2024 12:00:00 +0000

Test body content
"""
        result = email_client._parse_email_data(raw_email, email_id="1")
        assert result["message_id"] is None


class TestSendEmailReplyHeaders:
    @pytest.mark.asyncio
    async def test_send_email_sets_in_reply_to_header(self, email_client):
        """Test that In-Reply-To header is set when provided."""
        mock_smtp = AsyncMock()
        mock_smtp.__aenter__.return_value = mock_smtp
        mock_smtp.__aexit__.return_value = None
        mock_smtp.login = AsyncMock()
        mock_smtp.send_message = AsyncMock()

        with patch("aiosmtplib.SMTP", return_value=mock_smtp):
            await email_client.send_email(
                recipients=["recipient@example.com"],
                subject="Re: Test",
                body="Reply body",
                in_reply_to="<original123@example.com>",
            )

            call_args = mock_smtp.send_message.call_args
            msg = call_args[0][0]
            assert msg["In-Reply-To"] == "<original123@example.com>"

    @pytest.mark.asyncio
    async def test_send_email_sets_references_header(self, email_client):
        """Test that References header is set when provided."""
        mock_smtp = AsyncMock()
        mock_smtp.__aenter__.return_value = mock_smtp
        mock_smtp.__aexit__.return_value = None
        mock_smtp.login = AsyncMock()
        mock_smtp.send_message = AsyncMock()

        with patch("aiosmtplib.SMTP", return_value=mock_smtp):
            await email_client.send_email(
                recipients=["recipient@example.com"],
                subject="Re: Test",
                body="Reply body",
                references="<first@example.com> <second@example.com>",
            )

            call_args = mock_smtp.send_message.call_args
            msg = call_args[0][0]
            assert msg["References"] == "<first@example.com> <second@example.com>"

    @pytest.mark.asyncio
    async def test_send_email_without_reply_headers(self, email_client):
        """Test that send works without reply headers (backward compatibility)."""
        mock_smtp = AsyncMock()
        mock_smtp.__aenter__.return_value = mock_smtp
        mock_smtp.__aexit__.return_value = None
        mock_smtp.login = AsyncMock()
        mock_smtp.send_message = AsyncMock()

        with patch("aiosmtplib.SMTP", return_value=mock_smtp):
            await email_client.send_email(
                recipients=["recipient@example.com"],
                subject="Test",
                body="Body",
            )

            call_args = mock_smtp.send_message.call_args
            msg = call_args[0][0]
            assert "In-Reply-To" not in msg
            assert "References" not in msg



class TestDeleteEmails:
    """Tests for delete_emails functionality."""

    @pytest.mark.asyncio
    async def test_delete_emails_success(self, email_client):
        """Test successful deletion of emails."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.uid = AsyncMock(return_value=(None, None))
        mock_imap.expunge = AsyncMock()
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "_imap_connect", return_value=mock_imap):
            deleted_ids, failed_ids = await email_client.delete_emails(["123", "456"])
            assert deleted_ids == ["123", "456"]
            assert failed_ids == []
            mock_imap.expunge.assert_called_once()

    @pytest.mark.asyncio
    async def test_delete_emails_partial_failure(self, email_client):
        """Test delete_emails with some failures."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.expunge = AsyncMock()
        mock_imap.logout = AsyncMock()

        call_count = [0]

        def uid_side_effect(*args):
            call_count[0] += 1
            if call_count[0] == 1:
                return (None, None)
            else:
                raise OSError("IMAP error")

        mock_imap.uid = AsyncMock(side_effect=uid_side_effect)

        with patch.object(email_client, "_imap_connect", return_value=mock_imap):
            deleted_ids, failed_ids = await email_client.delete_emails(["123", "456"])
            assert deleted_ids == ["123"]
            assert failed_ids == ["456"]

    @pytest.mark.asyncio
    async def test_delete_emails_logout_error(self, email_client):
        """Test delete_emails handles logout errors gracefully."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.uid = AsyncMock(return_value=(None, None))
        mock_imap.expunge = AsyncMock()
        mock_imap.logout = AsyncMock(side_effect=OSError("Connection closed"))

        with patch.object(email_client, "_imap_connect", return_value=mock_imap):
            deleted_ids, failed_ids = await email_client.delete_emails(["123"])
            assert deleted_ids == ["123"]
            assert failed_ids == []


class TestMarkEmails:
    """Tests for mark_emails method."""

    @pytest.mark.asyncio
    async def test_mark_emails_as_read_success(self, email_client):
        """Test marking emails as read successfully."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.uid = AsyncMock(return_value=(None, None))
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "_imap_connect", return_value=mock_imap):
            marked_ids, failed_ids = await email_client.mark_emails(
                email_ids=["123", "456"],
                mark_as="read",
                mailbox="INBOX",
            )

            assert marked_ids == ["123", "456"]
            assert failed_ids == []
            # Verify +FLAGS was used for marking as read
            calls = mock_imap.uid.call_args_list
            assert len(calls) == 2
            assert calls[0][0] == ("store", "123", "+FLAGS", r"(\Seen)")
            assert calls[1][0] == ("store", "456", "+FLAGS", r"(\Seen)")

    @pytest.mark.asyncio
    async def test_mark_emails_as_unread_success(self, email_client):
        """Test marking emails as unread successfully."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.uid = AsyncMock(return_value=(None, None))
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "_imap_connect", return_value=mock_imap):
            marked_ids, failed_ids = await email_client.mark_emails(
                email_ids=["123", "456"],
                mark_as="unread",
                mailbox="INBOX",
            )

            assert marked_ids == ["123", "456"]
            assert failed_ids == []
            # Verify -FLAGS was used for marking as unread
            calls = mock_imap.uid.call_args_list
            assert len(calls) == 2
            assert calls[0][0] == ("store", "123", "-FLAGS", r"(\Seen)")
            assert calls[1][0] == ("store", "456", "-FLAGS", r"(\Seen)")

    @pytest.mark.asyncio
    async def test_mark_emails_partial_failure(self, email_client):
        """Test marking emails with some failures."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        # First call succeeds, second raises exception
        mock_imap.uid = AsyncMock(side_effect=[None, Exception("Email not found")])
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "_imap_connect", return_value=mock_imap):
            marked_ids, failed_ids = await email_client.mark_emails(
                email_ids=["123", "456"],
                mark_as="read",
                mailbox="INBOX",
            )

            assert marked_ids == ["123"]
            assert failed_ids == ["456"]

    @pytest.mark.asyncio
    async def test_mark_emails_invalid_mark_as_value(self, email_client):
        """Test that invalid mark_as value raises ValueError."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "_imap_connect", return_value=mock_imap):
            with pytest.raises(ValueError) as exc_info:
                await email_client.mark_emails(
                    email_ids=["123"],
                    mark_as="invalid",
                    mailbox="INBOX",
                )
            assert "Invalid mark_as value" in str(exc_info.value)

    @pytest.mark.asyncio
    async def test_mark_emails_custom_mailbox(self, email_client):
        """Test marking emails in a custom mailbox."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.uid = AsyncMock(return_value=(None, None))
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "_imap_connect", return_value=mock_imap):
            await email_client.mark_emails(
                email_ids=["123"],
                mark_as="read",
                mailbox="[Gmail]/All Mail",
            )

            # Verify custom mailbox was selected (quoted)
            mock_imap.select.assert_called_once_with('"[Gmail]/All Mail"')

    @pytest.mark.asyncio
    async def test_mark_emails_logout_error(self, email_client):
        """Test mark_emails handles logout errors gracefully (covers logout exception handler)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.uid = AsyncMock(return_value=(None, None))
        mock_imap.logout = AsyncMock(side_effect=OSError("Connection closed"))

        with patch.object(email_client, "_imap_connect", return_value=mock_imap):
            # Should complete successfully despite logout error
            marked_ids, failed_ids = await email_client.mark_emails(
                email_ids=["123"],
                mark_as="read",
                mailbox="INBOX",
            )
            assert marked_ids == ["123"]
            assert failed_ids == []


class TestSmtpSslContext:
    """Tests for SMTP SSL context creation."""

    def test_create_smtp_ssl_context_with_verification(self):
        """When verify_ssl=True, should return None (use default verification)."""
        result = _create_smtp_ssl_context(verify_ssl=True)
        assert result is None

    def test_create_smtp_ssl_context_without_verification(self):
        """When verify_ssl=False, should return permissive SSL context."""
        result = _create_smtp_ssl_context(verify_ssl=False)

        assert result is not None
        assert isinstance(result, ssl.SSLContext)
        assert result.check_hostname is False
        assert result.verify_mode == ssl.CERT_NONE

    def test_email_client_get_smtp_ssl_context_default(self):
        """EmailClient should use verify_ssl from EmailServer (default True)."""
        server = EmailServer(
            user_name="test",
            password="test",
            host="smtp.example.com",
            port=587,
        )
        client = EmailClient(server)

        # Default verify_ssl is True, so should return None
        assert client.smtp_verify_ssl is True
        assert client._get_smtp_ssl_context() is None

    def test_email_client_get_smtp_ssl_context_disabled(self):
        """EmailClient should return permissive context when verify_ssl=False."""
        server = EmailServer(
            user_name="test",
            password="test",
            host="smtp.example.com",
            port=587,
            verify_ssl=False,
        )
        client = EmailClient(server)

        assert client.smtp_verify_ssl is False
        ctx = client._get_smtp_ssl_context()
        assert ctx is not None
        assert ctx.check_hostname is False
        assert ctx.verify_mode == ssl.CERT_NONE

    @pytest.mark.asyncio
    async def test_send_email_passes_tls_context(self):
        """send_email should pass tls_context to SMTP connection."""
        server = EmailServer(
            user_name="test",
            password="test",
            host="smtp.example.com",
            port=587,
            verify_ssl=False,
        )
        client = EmailClient(server, sender="test@example.com")

        mock_smtp = AsyncMock()
        mock_smtp.__aenter__.return_value = mock_smtp
        mock_smtp.__aexit__.return_value = None
        mock_smtp.login = AsyncMock()
        mock_smtp.send_message = AsyncMock()

        with patch("aiosmtplib.SMTP", return_value=mock_smtp) as mock_smtp_class:
            await client.send_email(
                recipients=["recipient@example.com"],
                subject="Test",
                body="Body",
            )

            # Verify SMTP was called with tls_context
            call_kwargs = mock_smtp_class.call_args.kwargs
            assert "tls_context" in call_kwargs
            ctx = call_kwargs["tls_context"]
            assert ctx is not None
            assert ctx.check_hostname is False
            assert ctx.verify_mode == ssl.CERT_NONE


class TestParseHeaders:
    def test_parse_headers_extracts_metadata(self, email_client):
        """Test that _parse_headers correctly extracts email metadata."""
        raw_headers = b"""From: sender@example.com
To: recipient@example.com, other@example.com
Cc: cc@example.com
Subject: Test Subject
Date: Mon, 1 Jan 2024 12:00:00 +0000

"""
        result = email_client._parse_headers("123", raw_headers)

        assert result["email_id"] == "123"
        assert result["subject"] == "Test Subject"
        assert result["from"] == "sender@example.com"
        assert "recipient@example.com" in result["to"]
        assert "other@example.com" in result["to"]
        assert "cc@example.com" in result["to"]

    def test_parse_headers_handles_missing_fields(self, email_client):
        """Test that _parse_headers handles emails with missing optional fields."""
        raw_headers = b"""From: sender@example.com
Subject: Minimal Email

"""
        result = email_client._parse_headers("456", raw_headers)

        assert result["email_id"] == "456"
        assert result["subject"] == "Minimal Email"
        assert result["to"] == []

    def test_parse_headers_returns_none_for_invalid(self, email_client):
        """Test that _parse_headers returns None for unparseable data."""
        result = email_client._parse_headers("789", b"\x00\x01\x02\x03")

        # Should return None or handle gracefully
        assert result is None or isinstance(result, dict)


class TestBatchFetchDates:
    @pytest.mark.asyncio
    async def test_batch_fetch_dates_parses_imap_response(self, email_client):
        """Test that _batch_fetch_dates correctly parses IMAP INTERNALDATE responses."""
        mock_imap = AsyncMock()
        mock_imap.uid = AsyncMock(
            return_value=(
                None,
                [
                    b'1 FETCH (UID 100 INTERNALDATE "01-Jan-2024 12:00:00 +0000")',
                    b'2 FETCH (UID 200 INTERNALDATE "02-Jan-2024 12:00:00 +0000")',
                    b"FETCH completed",
                ],
            )
        )

        result = await email_client._batch_fetch_dates(mock_imap, [b"100", b"200"])

        assert len(result) == 2
        assert "100" in result
        assert "200" in result
        assert result["100"].day == 1
        assert result["200"].day == 2

    @pytest.mark.asyncio
    async def test_batch_fetch_dates_empty_input(self, email_client):
        """Test that _batch_fetch_dates returns empty dict for empty input."""
        mock_imap = AsyncMock()
        result = await email_client._batch_fetch_dates(mock_imap, [])
        assert result == {}
        mock_imap.uid.assert_not_called()

    @pytest.mark.asyncio
    async def test_batch_fetch_dates_handles_fastmail_format(self, email_client):
        """Test that _batch_fetch_dates handles space-padded dates (Fastmail)."""
        mock_imap = AsyncMock()
        mock_imap.uid = AsyncMock(
            return_value=(
                None,
                [
                    b'1 FETCH (UID 100 INTERNALDATE " 1-Jan-2024 12:00:00 +0000")',
                    b"FETCH completed",
                ],
            )
        )

        result = await email_client._batch_fetch_dates(mock_imap, [b"100"])

        assert result["100"] == datetime(2024, 1, 1, 12, 0, 0, tzinfo=timezone.utc)


class TestBatchFetchHeaders:
    @pytest.mark.asyncio
    async def test_batch_fetch_headers_parses_imap_response(self, email_client):
        """Test that _batch_fetch_headers correctly parses IMAP header responses."""
        mock_imap = AsyncMock()
        # aioimaplib returns FETCH response in 3 parts:
        # - BODY[HEADER] line (no UID)
        # - header content as bytearray
        # - UID line
        mock_imap.uid = AsyncMock(
            return_value=(
                None,
                [
                    b"1 FETCH (BODY[HEADER] {50}",
                    bytearray(b"From: a@test.com\r\nSubject: Test\r\n\r\n"),
                    b" UID 100)",
                    b"FETCH completed",
                ],
            )
        )

        result = await email_client._batch_fetch_headers(mock_imap, ["100"])

        assert "100" in result
        assert result["100"]["subject"] == "Test"
        assert result["100"]["from"] == "a@test.com"

    @pytest.mark.asyncio
    async def test_batch_fetch_headers_empty_input(self, email_client):
        """Test that _batch_fetch_headers returns empty dict for empty input."""
        mock_imap = AsyncMock()
        result = await email_client._batch_fetch_headers(mock_imap, [])
        assert result == {}
        mock_imap.uid.assert_not_called()

    @pytest.mark.asyncio
    async def test_batch_fetch_headers_preserves_uid_mapping(self, email_client):
        """Test that _batch_fetch_headers returns dict keyed by UID."""
        mock_imap = AsyncMock()
        # aioimaplib returns each email as 3 items: BODY[HEADER] line, content, UID line
        mock_imap.uid = AsyncMock(
            return_value=(
                None,
                [
                    b"1 FETCH (BODY[HEADER] {50}",
                    bytearray(b"From: a@test.com\r\nSubject: First\r\n\r\n"),
                    b" UID 100)",
                    b"2 FETCH (BODY[HEADER] {50}",
                    bytearray(b"From: b@test.com\r\nSubject: Second\r\n\r\n"),
                    b" UID 200)",
                    b"FETCH completed",
                ],
            )
        )

        result = await email_client._batch_fetch_headers(mock_imap, ["100", "200"])

        assert len(result) == 2
        assert result["100"]["subject"] == "First"
        assert result["200"]["subject"] == "Second"

    @pytest.mark.asyncio
    async def test_batch_fetch_headers_skips_non_bytes_items(self, email_client):
        """Test that _batch_fetch_headers skips non-bytes items in response."""
        mock_imap = AsyncMock()
        mock_imap.uid = AsyncMock(
            return_value=(
                None,
                [
                    "not bytes",  # Should be skipped
                    b"1 FETCH (BODY[HEADER] {50}",
                    bytearray(b"From: a@test.com\r\nSubject: Test\r\n\r\n"),
                    b" UID 100)",
                ],
            )
        )

        result = await email_client._batch_fetch_headers(mock_imap, ["100"])

        assert "100" in result
        assert result["100"]["subject"] == "Test"

    @pytest.mark.asyncio
    async def test_batch_fetch_headers_skips_items_without_body_header(self, email_client):
        """Test that _batch_fetch_headers skips bytes without BODY[HEADER]."""
        mock_imap = AsyncMock()
        mock_imap.uid = AsyncMock(
            return_value=(
                None,
                [
                    b"some other data",  # No BODY[HEADER], should be skipped
                    b"1 FETCH (BODY[HEADER] {50}",
                    bytearray(b"From: a@test.com\r\nSubject: Test\r\n\r\n"),
                    b" UID 100)",
                ],
            )
        )

        result = await email_client._batch_fetch_headers(mock_imap, ["100"])

        assert "100" in result
        assert result["100"]["subject"] == "Test"

    @pytest.mark.asyncio
    async def test_batch_fetch_headers_skips_truncated_response(self, email_client):
        """Test that _batch_fetch_headers skips when i+2 >= len(data)."""
        mock_imap = AsyncMock()
        mock_imap.uid = AsyncMock(
            return_value=(
                None,
                [
                    b"1 FETCH (BODY[HEADER] {50}",
                    bytearray(b"From: a@test.com\r\nSubject: Test\r\n\r\n"),
                    # Missing UID line (i+2 doesn't exist)
                ],
            )
        )

        result = await email_client._batch_fetch_headers(mock_imap, ["100"])

        # Should return empty dict since response is truncated
        assert result == {}

    @pytest.mark.asyncio
    async def test_batch_fetch_headers_skips_non_bytearray_content(self, email_client):
        """Test that _batch_fetch_headers skips when data[i+1] is not bytearray."""
        mock_imap = AsyncMock()
        mock_imap.uid = AsyncMock(
            return_value=(
                None,
                [
                    b"1 FETCH (BODY[HEADER] {50}",
                    b"not a bytearray",  # Should be bytearray, not bytes
                    b" UID 100)",
                ],
            )
        )

        result = await email_client._batch_fetch_headers(mock_imap, ["100"])

        # Should return empty dict since content is not bytearray
        assert result == {}

    @pytest.mark.asyncio
    async def test_batch_fetch_headers_skips_non_bytes_uid_item(self, email_client):
        """Test that _batch_fetch_headers skips when data[i+2] is not bytes."""
        mock_imap = AsyncMock()
        mock_imap.uid = AsyncMock(
            return_value=(
                None,
                [
                    b"1 FETCH (BODY[HEADER] {50}",
                    bytearray(b"From: a@test.com\r\nSubject: Test\r\n\r\n"),
                    12345,  # Not bytes, should result in uid_item = None
                ],
            )
        )

        result = await email_client._batch_fetch_headers(mock_imap, ["100"])

        # Should return empty dict since UID item is not bytes
        assert result == {}

    @pytest.mark.asyncio
    async def test_batch_fetch_headers_skips_missing_uid_in_response(self, email_client):
        """Test that _batch_fetch_headers skips when UID regex doesn't match."""
        mock_imap = AsyncMock()
        mock_imap.uid = AsyncMock(
            return_value=(
                None,
                [
                    b"1 FETCH (BODY[HEADER] {50}",
                    bytearray(b"From: a@test.com\r\nSubject: Test\r\n\r\n"),
                    b" NO_UID_HERE)",  # No UID in this line
                ],
            )
        )

        result = await email_client._batch_fetch_headers(mock_imap, ["100"])

        # Should return empty dict since UID regex doesn't match
        assert result == {}

    @pytest.mark.asyncio
    async def test_batch_fetch_headers_handles_mixed_valid_invalid(self, email_client):
        """Test that _batch_fetch_headers processes valid items and skips invalid ones."""
        mock_imap = AsyncMock()
        mock_imap.uid = AsyncMock(
            return_value=(
                None,
                [
                    # Invalid: truncated (no UID line)
                    b"1 FETCH (BODY[HEADER] {50}",
                    bytearray(b"From: bad@test.com\r\nSubject: Bad\r\n\r\n"),
                    # Valid email
                    b"2 FETCH (BODY[HEADER] {50}",
                    bytearray(b"From: good@test.com\r\nSubject: Good\r\n\r\n"),
                    b" UID 200)",
                ],
            )
        )

        result = await email_client._batch_fetch_headers(mock_imap, ["100", "200"])

        # Only the valid email should be in results
        assert len(result) == 1
        assert "200" in result
        assert result["200"]["subject"] == "Good"


class TestParseEmailDataHtmlFallback:
    """Tests for HTML-to-text fallback in _parse_email_data."""

    def test_html_only_multipart_extracts_body(self, email_client):
        """HTML-only multipart email should have body extracted from HTML."""
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText

        msg = MIMEMultipart("alternative")
        msg["Subject"] = "HTML Newsletter"
        msg["From"] = "newsletter@example.com"
        msg["To"] = "user@example.com"
        msg["Date"] = email.utils.formatdate()

        html_part = MIMEText("<p>Hello from the newsletter!</p><p>Click <a href='https://example.com'>here</a>.</p>", "html")
        msg.attach(html_part)

        result = email_client._parse_email_data(msg.as_bytes())

        assert "Hello from the newsletter!" in result["body"]
        assert "(https://example.com)" in result["body"]
        assert "<p>" not in result["body"]

    def test_multipart_prefers_plain_text(self, email_client):
        """When both text/plain and text/html exist, prefer plain text."""
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText

        msg = MIMEMultipart("alternative")
        msg["Subject"] = "Both Formats"
        msg["From"] = "sender@example.com"
        msg["To"] = "user@example.com"
        msg["Date"] = email.utils.formatdate()

        plain_part = MIMEText("Plain text version", "plain")
        html_part = MIMEText("<p>HTML version</p>", "html")
        msg.attach(plain_part)
        msg.attach(html_part)

        result = email_client._parse_email_data(msg.as_bytes())

        assert result["body"] == "Plain text version"

    def test_non_multipart_html_converts_to_text(self, email_client):
        """Non-multipart HTML email should be converted to text."""
        from email.mime.text import MIMEText

        msg = MIMEText("<h1>Title</h1><p>Some content here.</p>", "html")
        msg["Subject"] = "HTML Only"
        msg["From"] = "sender@example.com"
        msg["To"] = "user@example.com"
        msg["Date"] = email.utils.formatdate()

        result = email_client._parse_email_data(msg.as_bytes())

        assert "Title" in result["body"]
        assert "Some content here." in result["body"]
        assert "<h1>" not in result["body"]
        assert "<p>" not in result["body"]

    def test_non_multipart_plain_text_unchanged(self, email_client):
        """Non-multipart plain text email should be returned as-is (regression test)."""
        msg = MIMEText("Just plain text, no HTML.", "plain")
        msg["Subject"] = "Plain"
        msg["From"] = "sender@example.com"
        msg["To"] = "user@example.com"
        msg["Date"] = email.utils.formatdate()

        result = email_client._parse_email_data(msg.as_bytes())

        assert result["body"] == "Just plain text, no HTML."

    def test_html_only_multipart_preserves_html_body(self, email_client):
        """HTML-only multipart email should preserve html_body in result."""
        from email.mime.multipart import MIMEMultipart

        msg = MIMEMultipart("alternative")
        msg["Subject"] = "HTML Newsletter"
        msg["From"] = "newsletter@example.com"
        msg["To"] = "user@example.com"
        msg["Date"] = email.utils.formatdate()

        html_part = MIMEText("<p>Hello from the newsletter!</p>", "html")
        msg.attach(html_part)

        result = email_client._parse_email_data(msg.as_bytes())

        assert "<p>Hello from the newsletter!</p>" in result["html_body"]

    def test_non_multipart_html_preserves_html_body(self, email_client):
        """Non-multipart HTML email should preserve html_body."""
        msg = MIMEText("<h1>Title</h1><p>Content</p>", "html")
        msg["Subject"] = "HTML Only"
        msg["From"] = "sender@example.com"
        msg["To"] = "user@example.com"
        msg["Date"] = email.utils.formatdate()

        result = email_client._parse_email_data(msg.as_bytes())

        assert "<h1>Title</h1>" in result["html_body"]
        assert "<p>Content</p>" in result["html_body"]

    def test_plain_text_has_empty_html_body(self, email_client):
        """Plain text email should have empty html_body."""
        msg = MIMEText("Just plain text", "plain")
        msg["Subject"] = "Plain"
        msg["From"] = "sender@example.com"
        msg["To"] = "user@example.com"
        msg["Date"] = email.utils.formatdate()

        result = email_client._parse_email_data(msg.as_bytes())

        assert result["html_body"] == ""

    def test_multipart_with_both_preserves_html_body(self, email_client):
        """Multipart with both text/plain and text/html should preserve html_body."""
        from email.mime.multipart import MIMEMultipart

        msg = MIMEMultipart("alternative")
        msg["Subject"] = "Both Formats"
        msg["From"] = "sender@example.com"
        msg["To"] = "user@example.com"
        msg["Date"] = email.utils.formatdate()

        plain_part = MIMEText("Plain text version", "plain")
        html_part = MIMEText("<p>HTML version</p>", "html")
        msg.attach(plain_part)
        msg.attach(html_part)

        result = email_client._parse_email_data(msg.as_bytes())

        # Body should be plain text (preferred)
        assert result["body"] == "Plain text version"
        # html_body should still have the HTML part
        assert "<p>HTML version</p>" in result["html_body"]


class TestFormatQuotedReplyHtml:
    """Tests for _format_quoted_reply_html helper."""

    def test_html_body_preserved_in_blockquote(self):
        """Test that original HTML body is preserved inside blockquote."""
        original = {
            "from": "Alice <alice@example.com>",
            "date": datetime(2024, 3, 15, 14, 30, tzinfo=timezone.utc),
            "body": "Plain text version",
            "html_body": "<p>Rich <strong>HTML</strong> content</p>",
        }
        result = _format_quoted_reply_html(original)

        assert '<blockquote type="cite"' in result
        assert "<p>Rich <strong>HTML</strong> content</p>" in result
        assert "Alice" in result
        assert "wrote:" in result

    def test_plain_text_fallback(self):
        """Test that plain text is used when html_body is empty."""
        original = {
            "from": "Bob <bob@example.com>",
            "date": datetime(2024, 3, 15, 14, 30, tzinfo=timezone.utc),
            "body": "Just plain text\nWith multiple lines",
            "html_body": "",
        }
        result = _format_quoted_reply_html(original)

        assert "<blockquote" in result
        assert "Just plain text" in result
        assert "With multiple lines" in result
        # Plain text should be HTML-escaped and use <br> for newlines
        assert "<br>" in result

    def test_strips_html_wrappers(self):
        """Test that DOCTYPE, html, head, body wrappers are stripped."""
        original = {
            "from": "sender@example.com",
            "date": datetime(2024, 1, 1, tzinfo=timezone.utc),
            "body": "",
            "html_body": '<!DOCTYPE html><html><head><style>h1{color:red}</style></head><body><p>Content</p></body></html>',
        }
        result = _format_quoted_reply_html(original)

        assert "<!DOCTYPE" not in result
        assert "<html>" not in result
        assert "<head>" not in result
        assert "<body>" not in result
        assert "<p>Content</p>" in result

    def test_escapes_sender_in_attribution(self):
        """Test that sender name is HTML-escaped in attribution line."""
        original = {
            "from": "Evil <script>alert('xss')</script>",
            "date": datetime(2024, 1, 1, tzinfo=timezone.utc),
            "body": "test",
            "html_body": "",
        }
        result = _format_quoted_reply_html(original)

        assert "<script>" not in result
        assert "&lt;script&gt;" in result

    def test_long_text_body_truncation(self):
        """Test that long plain text bodies are truncated in HTML quote."""
        long_body = "x" * 6000
        original = {
            "from": "sender@example.com",
            "date": datetime(2024, 1, 1, tzinfo=timezone.utc),
            "body": long_body,
            "html_body": "",
        }
        result = _format_quoted_reply_html(original)

        assert "[...quoted text truncated]" in result

    def test_missing_fields(self):
        """Test graceful handling of missing fields."""
        result = _format_quoted_reply_html({})

        assert "<blockquote" in result
        assert "Unknown" in result
        assert "Unknown date" in result

    def test_empty_body_and_html_body(self):
        """Test with both body and html_body empty."""
        original = {
            "from": "sender@example.com",
            "date": datetime(2024, 1, 1, tzinfo=timezone.utc),
            "body": "",
            "html_body": "",
        }
        result = _format_quoted_reply_html(original)

        assert "<blockquote" in result
        assert "</blockquote>" in result


@pytest.fixture
def email_settings():
    """Create test EmailSettings for ClassicEmailHandler tests."""
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


class TestAutoQuoteReply:
    """Tests for auto-quoting in ClassicEmailHandler.send_email."""

    @pytest.mark.asyncio
    async def test_quote_reply_appends_quoted_text(self, email_settings):
        """Test that quote_reply=True fetches and appends HTML blockquote."""
        handler = ClassicEmailHandler(email_settings)

        original_email = {
            "email_id": "42",
            "message_id": "<original@example.com>",
            "subject": "Original Subject",
            "from": "Alice <alice@example.com>",
            "to": ["test@example.com"],
            "date": datetime(2024, 3, 15, 14, 30, tzinfo=timezone.utc),
            "body": "Original message body",
            "html_body": "",
            "attachments": [],
        }

        with (
            patch.object(
                handler.incoming_client, "search_by_message_id", return_value="42"
            ) as mock_search,
            patch.object(
                handler.incoming_client, "get_email_body_by_id", return_value=original_email
            ) as mock_fetch,
            patch.object(handler.outgoing_client, "send_email", return_value=MagicMock()) as mock_send,
        ):
            await handler.send_email(
                recipients=["alice@example.com"],
                subject="Re: Original Subject",
                body="My reply text",
                in_reply_to="<original@example.com>",
                quote_reply=True,
            )

            mock_search.assert_called_once_with("<original@example.com>", "INBOX")
            mock_fetch.assert_called_once_with("42", "INBOX")

            # Default path uses HTML blockquote
            sent_body = mock_send.call_args[0][2]  # body is 3rd positional arg
            assert "My reply text" in sent_body
            assert "<blockquote" in sent_body
            assert "Original message body" in sent_body
            assert "Alice" in sent_body
            assert "wrote:" in sent_body
            # Should be sent as html=True (pre-built HTML document)
            assert mock_send.call_args[0][5] is True  # html

    @pytest.mark.asyncio
    async def test_quote_reply_false_skips_fetching(self, email_settings):
        """Test that quote_reply=False does not fetch the original."""
        handler = ClassicEmailHandler(email_settings)

        with (
            patch.object(handler.incoming_client, "search_by_message_id") as mock_search,
            patch.object(handler.outgoing_client, "send_email", return_value=MagicMock()),
        ):
            await handler.send_email(
                recipients=["alice@example.com"],
                subject="Re: Test",
                body="My reply",
                in_reply_to="<original@example.com>",
                quote_reply=False,
            )

            mock_search.assert_not_called()

    @pytest.mark.asyncio
    async def test_quote_reply_no_in_reply_to_skips(self, email_settings):
        """Test that without in_reply_to, no quoting happens."""
        handler = ClassicEmailHandler(email_settings)

        with (
            patch.object(handler.incoming_client, "search_by_message_id") as mock_search,
            patch.object(handler.outgoing_client, "send_email", return_value=MagicMock()),
        ):
            await handler.send_email(
                recipients=["alice@example.com"],
                subject="New email",
                body="Hello",
                quote_reply=True,
            )

            mock_search.assert_not_called()

    @pytest.mark.asyncio
    async def test_quote_reply_original_not_found_sends_without_quote(self, email_settings):
        """Test that if original email is not found, send proceeds without quote."""
        handler = ClassicEmailHandler(email_settings)

        with (
            patch.object(handler.incoming_client, "search_by_message_id", return_value=None),
            patch.object(handler.incoming_client, "list_folders", return_value=[]),
            patch.object(handler.outgoing_client, "send_email", return_value=MagicMock()) as mock_send,
        ):
            await handler.send_email(
                recipients=["alice@example.com"],
                subject="Re: Test",
                body="My reply",
                in_reply_to="<nonexistent@example.com>",
                quote_reply=True,
            )

            # Body should be unchanged (no quote appended)
            sent_body = mock_send.call_args[0][2]
            assert sent_body == "My reply"

    @pytest.mark.asyncio
    async def test_quote_reply_searches_sent_folder_as_fallback(self, email_settings):
        """Test that Sent folder is searched when original not found in INBOX."""
        handler = ClassicEmailHandler(email_settings)

        original_email = {
            "email_id": "99",
            "message_id": "<sent@example.com>",
            "subject": "Sent Subject",
            "from": "Test User <test@example.com>",
            "to": ["alice@example.com"],
            "date": datetime(2024, 3, 15, 14, 30, tzinfo=timezone.utc),
            "body": "I sent this originally",
            "html_body": "",
            "attachments": [],
        }

        from mcp_email_server.emails.models import Folder

        sent_folder = Folder(name="Sent", delimiter="/", flags=["\\Sent", "\\HasNoChildren"])

        # First search (INBOX) returns None, second search (Sent) returns UID
        search_side_effects = [None, "99"]

        with (
            patch.object(
                handler.incoming_client, "search_by_message_id", side_effect=search_side_effects
            ),
            patch.object(
                handler.incoming_client, "list_folders", return_value=[sent_folder]
            ),
            patch.object(
                handler.incoming_client, "get_email_body_by_id", return_value=original_email
            ) as mock_fetch,
            patch.object(handler.outgoing_client, "send_email", return_value=MagicMock()) as mock_send,
        ):
            await handler.send_email(
                recipients=["alice@example.com"],
                subject="Re: Sent Subject",
                body="Replying to my own email",
                in_reply_to="<sent@example.com>",
                quote_reply=True,
            )

            # Should have fetched from Sent folder
            mock_fetch.assert_called_once_with("99", "Sent")

            sent_body = mock_send.call_args[0][2]
            assert "<blockquote" in sent_body
            assert "I sent this originally" in sent_body

    @pytest.mark.asyncio
    async def test_html_blockquote_with_html_original(self, email_settings):
        """Default reply preserves original HTML in blockquote."""
        handler = ClassicEmailHandler(email_settings)

        original_email = {
            "email_id": "42",
            "message_id": "<original@example.com>",
            "subject": "Original Subject",
            "from": "Alice <alice@example.com>",
            "to": ["test@example.com"],
            "date": datetime(2024, 3, 15, 14, 30, tzinfo=timezone.utc),
            "body": "Plain text fallback",
            "html_body": "<p>Original <strong>HTML</strong> content</p>",
            "attachments": [],
        }

        with (
            patch.object(
                handler.incoming_client, "search_by_message_id", return_value="42"
            ),
            patch.object(
                handler.incoming_client, "get_email_body_by_id", return_value=original_email
            ),
            patch.object(handler.outgoing_client, "send_email", return_value=MagicMock()) as mock_send,
        ):
            await handler.send_email(
                recipients=["alice@example.com"],
                subject="Re: Original Subject",
                body="My reply text",
                in_reply_to="<original@example.com>",
                quote_reply=True,
            )

            sent_body = mock_send.call_args[0][2]
            # Full HTML document
            assert "<!DOCTYPE html>" in sent_body
            # User's reply converted from markdown
            assert "My reply text" in sent_body
            # Original HTML preserved in blockquote
            assert "<blockquote" in sent_body
            assert "Original <strong>HTML</strong> content" in sent_body
            # Attribution line
            assert "Alice" in sent_body
            assert "wrote:" in sent_body
            # Should be sent as html=True (pre-built document)
            assert mock_send.call_args[0][5] is True  # html

    @pytest.mark.asyncio
    async def test_html_blockquote_with_plain_text_original(self, email_settings):
        """Default reply with plain text original uses escaped text in blockquote."""
        handler = ClassicEmailHandler(email_settings)

        original_email = {
            "email_id": "42",
            "message_id": "<original@example.com>",
            "subject": "Original Subject",
            "from": "Bob <bob@example.com>",
            "to": ["test@example.com"],
            "date": datetime(2024, 3, 15, 14, 30, tzinfo=timezone.utc),
            "body": "Plain text original\nSecond line",
            "html_body": "",
            "attachments": [],
        }

        with (
            patch.object(
                handler.incoming_client, "search_by_message_id", return_value="42"
            ),
            patch.object(
                handler.incoming_client, "get_email_body_by_id", return_value=original_email
            ),
            patch.object(handler.outgoing_client, "send_email", return_value=MagicMock()) as mock_send,
        ):
            await handler.send_email(
                recipients=["bob@example.com"],
                subject="Re: Original Subject",
                body="My reply",
                in_reply_to="<original@example.com>",
                quote_reply=True,
            )

            sent_body = mock_send.call_args[0][2]
            assert "<blockquote" in sent_body
            assert "Plain text original" in sent_body
            assert "Second line" in sent_body
            assert mock_send.call_args[0][5] is True  # html

    @pytest.mark.asyncio
    async def test_html_true_reply_appends_blockquote(self, email_settings):
        """When html=True, blockquote is appended directly to raw HTML body."""
        handler = ClassicEmailHandler(email_settings)

        original_email = {
            "email_id": "42",
            "message_id": "<original@example.com>",
            "subject": "Original Subject",
            "from": "Alice <alice@example.com>",
            "to": ["test@example.com"],
            "date": datetime(2024, 3, 15, 14, 30, tzinfo=timezone.utc),
            "body": "Original message body",
            "html_body": "<p>Original HTML</p>",
            "attachments": [],
        }

        with (
            patch.object(
                handler.incoming_client, "search_by_message_id", return_value="42"
            ),
            patch.object(
                handler.incoming_client, "get_email_body_by_id", return_value=original_email
            ),
            patch.object(handler.outgoing_client, "send_email", return_value=MagicMock()) as mock_send,
        ):
            await handler.send_email(
                recipients=["alice@example.com"],
                subject="Re: Original Subject",
                body="<p>My raw HTML reply</p>",
                html=True,
                in_reply_to="<original@example.com>",
                quote_reply=True,
            )

            sent_body = mock_send.call_args[0][2]
            # Raw HTML body preserved
            assert "<p>My raw HTML reply</p>" in sent_body
            # Blockquote appended
            assert "<blockquote" in sent_body
            assert "Original HTML" in sent_body
            # Still sent as html=True
            assert mock_send.call_args[0][5] is True  # html


class TestFormatForwardedEmailHtml:
    """Tests for _format_forwarded_email_html helper."""

    def test_basic_formatting_with_all_fields(self):
        """Test forwarded message with all fields present."""
        original = {
            "from": "Alice <alice@example.com>",
            "date": datetime(2024, 3, 15, 14, 30, tzinfo=timezone.utc),
            "subject": "Original Subject",
            "to": ["bob@example.com"],
            "body": "Hello Bob",
            "html_body": "",
        }
        result = _format_forwarded_email_html(original)

        assert "---------- Forwarded message ---------" in result
        assert "From: Alice &lt;alice@example.com&gt;" in result
        assert "Subject: Original Subject" in result
        assert "To: bob@example.com" in result
        assert "Hello Bob" in result

    def test_html_body_preferred_over_text(self):
        """Test that HTML body is used when available."""
        original = {
            "from": "sender@example.com",
            "date": datetime(2024, 1, 1, tzinfo=timezone.utc),
            "subject": "Test",
            "to": ["recipient@example.com"],
            "body": "Plain text version",
            "html_body": "<p>Rich <strong>HTML</strong> content</p>",
        }
        result = _format_forwarded_email_html(original)

        assert "<p>Rich <strong>HTML</strong> content</p>" in result
        assert "Plain text version" not in result

    def test_text_fallback_with_escaping(self):
        """Test text fallback with HTML escaping."""
        original = {
            "from": "sender@example.com",
            "date": datetime(2024, 1, 1, tzinfo=timezone.utc),
            "subject": "Test",
            "to": [],
            "body": "Line with <html> & special chars\nSecond line",
            "html_body": "",
        }
        result = _format_forwarded_email_html(original)

        assert "&lt;html&gt;" in result
        assert "&amp; special chars" in result
        assert "<br>" in result

    def test_missing_fields(self):
        """Test graceful handling of missing fields."""
        result = _format_forwarded_email_html({})

        assert "---------- Forwarded message ---------" in result
        assert "From: Unknown" in result
        assert "Unknown date" in result

    def test_long_body_truncation(self):
        """Test that long text bodies are truncated."""
        long_body = "x" * 6000
        original = {
            "from": "sender@example.com",
            "date": datetime(2024, 1, 1, tzinfo=timezone.utc),
            "subject": "Test",
            "to": [],
            "body": long_body,
            "html_body": "",
        }
        result = _format_forwarded_email_html(original)

        assert "[...forwarded text truncated]" in result

    def test_multiple_recipients(self):
        """Test To field with multiple recipients."""
        original = {
            "from": "sender@example.com",
            "date": datetime(2024, 1, 1, tzinfo=timezone.utc),
            "subject": "Test",
            "to": ["alice@example.com", "bob@example.com"],
            "body": "test",
            "html_body": "",
        }
        result = _format_forwarded_email_html(original)

        assert "To: alice@example.com, bob@example.com" in result

    def test_strips_html_wrappers(self):
        """Test that HTML document wrappers are stripped from forwarded body."""
        original = {
            "from": "sender@example.com",
            "date": datetime(2024, 1, 1, tzinfo=timezone.utc),
            "subject": "Test",
            "to": [],
            "body": "",
            "html_body": '<!DOCTYPE html><html><head><style>h1{color:red}</style></head><body><p>Content</p></body></html>',
        }
        result = _format_forwarded_email_html(original)

        assert "<!DOCTYPE" not in result
        assert "<html>" not in result
        assert "<p>Content</p>" in result

    def test_escapes_sender_xss(self):
        """Test that sender is HTML-escaped to prevent XSS."""
        original = {
            "from": "Evil <script>alert('xss')</script>",
            "date": datetime(2024, 1, 1, tzinfo=timezone.utc),
            "subject": "Test",
            "to": [],
            "body": "test",
            "html_body": "",
        }
        result = _format_forwarded_email_html(original)

        assert "<script>" not in result
        assert "&lt;script&gt;" in result


class TestExtractAttachments:
    """Tests for EmailClient.extract_attachments method."""

    @pytest.mark.asyncio
    async def test_extracts_multiple_attachments(self, email_client):
        """Test extracting multiple attachments from an email."""
        from email.mime.application import MIMEApplication
        from email.mime.multipart import MIMEMultipart

        # Build a multipart email with attachments
        msg = MIMEMultipart()
        msg["Subject"] = "Test"
        msg["From"] = "sender@example.com"
        msg["To"] = "recipient@example.com"
        msg.attach(MIMEText("Body text", "plain"))

        pdf_data = b"%PDF-1.4 fake pdf data"
        pdf_part = MIMEApplication(pdf_data, _subtype="pdf")
        pdf_part.add_header("Content-Disposition", "attachment", filename="document.pdf")
        msg.attach(pdf_part)

        img_data = b"\x89PNG fake image data"
        img_part = MIMEApplication(img_data, _subtype="png")
        img_part.add_header("Content-Disposition", "attachment", filename="image.png")
        msg.attach(img_part)

        raw_email = msg.as_bytes()

        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.logout = AsyncMock()

        with (
            patch.object(email_client, "_imap_connect", return_value=mock_imap),
            patch.object(
                email_client,
                "_fetch_email_with_formats",
                return_value=[b"1 FETCH", bytearray(raw_email), b"UID 123)"],
            ),
            patch.object(email_client, "_extract_raw_email", return_value=raw_email),
        ):
            result = await email_client.extract_attachments("123", "INBOX")

            assert len(result) == 2
            assert result[0][0] == "document.pdf"
            assert "pdf" in result[0][1]
            assert result[0][2] == pdf_data
            assert result[1][0] == "image.png"
            assert "png" in result[1][1]
            assert result[1][2] == img_data

    @pytest.mark.asyncio
    async def test_returns_empty_for_no_attachments(self, email_client):
        """Test that emails without attachments return empty list."""
        msg = MIMEText("Just text, no attachments", "plain")
        msg["Subject"] = "Test"
        msg["From"] = "sender@example.com"
        msg["To"] = "recipient@example.com"
        raw_email = msg.as_bytes()

        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.logout = AsyncMock()

        with (
            patch.object(email_client, "_imap_connect", return_value=mock_imap),
            patch.object(
                email_client,
                "_fetch_email_with_formats",
                return_value=[b"1 FETCH", bytearray(raw_email), b"UID 123)"],
            ),
            patch.object(email_client, "_extract_raw_email", return_value=raw_email),
        ):
            result = await email_client.extract_attachments("123", "INBOX")

            assert result == []

    @pytest.mark.asyncio
    async def test_returns_empty_when_email_not_found(self, email_client):
        """Test that extract_attachments returns empty list when email not found."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.logout = AsyncMock()

        with (
            patch.object(email_client, "_imap_connect", return_value=mock_imap),
            patch.object(email_client, "_fetch_email_with_formats", return_value=None),
        ):
            result = await email_client.extract_attachments("999", "INBOX")

            assert result == []


class TestForwardEmail:
    """Tests for ClassicEmailHandler.forward_email."""

    @pytest.mark.asyncio
    async def test_forward_with_user_body(self, email_settings):
        """Test forwarding with a user message prepended."""
        handler = ClassicEmailHandler(email_settings)

        original_email = {
            "email_id": "42",
            "message_id": "<original@example.com>",
            "subject": "Original Subject",
            "from": "Alice <alice@example.com>",
            "to": ["test@example.com"],
            "date": datetime(2024, 3, 15, 14, 30, tzinfo=timezone.utc),
            "body": "Original message body",
            "html_body": "",
            "attachments": [],
        }

        with (
            patch.object(
                handler.incoming_client, "get_email_body_by_id", return_value=original_email
            ),
            patch.object(
                handler.incoming_client, "extract_attachments", return_value=[]
            ),
            patch.object(handler.outgoing_client, "send_email", return_value=MagicMock()) as mock_send,
        ):
            await handler.forward_email(
                email_id="42",
                mailbox="INBOX",
                recipients=["bob@example.com"],
                body="FYI, see below.",
            )

            mock_send.assert_called_once()
            sent_body = mock_send.call_args[0][2]  # body is 3rd positional arg
            assert "FYI, see below." in sent_body
            assert "---------- Forwarded message ---------" in sent_body
            assert "Original message body" in sent_body
            assert mock_send.call_args.kwargs.get("html") is True or mock_send.call_args[0][5] is True

    @pytest.mark.asyncio
    async def test_forward_without_user_body(self, email_settings):
        """Test forwarding without a user message."""
        handler = ClassicEmailHandler(email_settings)

        original_email = {
            "email_id": "42",
            "message_id": "<original@example.com>",
            "subject": "Original Subject",
            "from": "Alice <alice@example.com>",
            "to": ["test@example.com"],
            "date": datetime(2024, 3, 15, 14, 30, tzinfo=timezone.utc),
            "body": "Original message body",
            "html_body": "",
            "attachments": [],
        }

        with (
            patch.object(
                handler.incoming_client, "get_email_body_by_id", return_value=original_email
            ),
            patch.object(
                handler.incoming_client, "extract_attachments", return_value=[]
            ),
            patch.object(handler.outgoing_client, "send_email", return_value=MagicMock()) as mock_send,
        ):
            await handler.forward_email(
                email_id="42",
                mailbox="INBOX",
                recipients=["bob@example.com"],
            )

            mock_send.assert_called_once()
            sent_body = mock_send.call_args[0][2]
            assert "---------- Forwarded message ---------" in sent_body
            assert "Original message body" in sent_body

    @pytest.mark.asyncio
    async def test_forward_adds_fwd_prefix(self, email_settings):
        """Test that subject gets Fwd: prefix."""
        handler = ClassicEmailHandler(email_settings)

        original_email = {
            "email_id": "42",
            "subject": "Important News",
            "from": "alice@example.com",
            "to": ["test@example.com"],
            "date": datetime(2024, 1, 1, tzinfo=timezone.utc),
            "body": "content",
            "html_body": "",
            "attachments": [],
        }

        with (
            patch.object(handler.incoming_client, "get_email_body_by_id", return_value=original_email),
            patch.object(handler.incoming_client, "extract_attachments", return_value=[]),
            patch.object(handler.outgoing_client, "send_email", return_value=MagicMock()) as mock_send,
        ):
            await handler.forward_email(
                email_id="42", mailbox="INBOX", recipients=["bob@example.com"]
            )

            sent_subject = mock_send.call_args[0][1]
            assert sent_subject == "Fwd: Important News"

    @pytest.mark.asyncio
    async def test_forward_no_double_fwd_prefix(self, email_settings):
        """Test that subject already starting with Fwd: is not double-prefixed."""
        handler = ClassicEmailHandler(email_settings)

        original_email = {
            "email_id": "42",
            "subject": "Fwd: Already Forwarded",
            "from": "alice@example.com",
            "to": ["test@example.com"],
            "date": datetime(2024, 1, 1, tzinfo=timezone.utc),
            "body": "content",
            "html_body": "",
            "attachments": [],
        }

        with (
            patch.object(handler.incoming_client, "get_email_body_by_id", return_value=original_email),
            patch.object(handler.incoming_client, "extract_attachments", return_value=[]),
            patch.object(handler.outgoing_client, "send_email", return_value=MagicMock()) as mock_send,
        ):
            await handler.forward_email(
                email_id="42", mailbox="INBOX", recipients=["bob@example.com"]
            )

            sent_subject = mock_send.call_args[0][1]
            assert sent_subject == "Fwd: Already Forwarded"

    @pytest.mark.asyncio
    async def test_forward_original_not_found_raises(self, email_settings):
        """Test that forwarding a nonexistent email raises ValueError."""
        handler = ClassicEmailHandler(email_settings)

        with patch.object(handler.incoming_client, "get_email_body_by_id", return_value=None):
            with pytest.raises(ValueError, match="not found"):
                await handler.forward_email(
                    email_id="999", mailbox="INBOX", recipients=["bob@example.com"]
                )

    @pytest.mark.asyncio
    async def test_forward_html_body_passthrough(self, email_settings):
        """Test that html=True body is passed through without markdown conversion."""
        handler = ClassicEmailHandler(email_settings)

        original_email = {
            "email_id": "42",
            "subject": "Test",
            "from": "alice@example.com",
            "to": ["test@example.com"],
            "date": datetime(2024, 1, 1, tzinfo=timezone.utc),
            "body": "content",
            "html_body": "",
            "attachments": [],
        }

        with (
            patch.object(handler.incoming_client, "get_email_body_by_id", return_value=original_email),
            patch.object(handler.incoming_client, "extract_attachments", return_value=[]),
            patch.object(handler.outgoing_client, "send_email", return_value=MagicMock()) as mock_send,
        ):
            await handler.forward_email(
                email_id="42",
                mailbox="INBOX",
                recipients=["bob@example.com"],
                body="<p>My HTML message</p>",
                html=True,
            )

            sent_body = mock_send.call_args[0][2]
            assert "<p>My HTML message</p>" in sent_body
            assert "---------- Forwarded message ---------" in sent_body

    @pytest.mark.asyncio
    async def test_forward_includes_original_attachments(self, email_settings):
        """Test that original email attachments are forwarded."""
        handler = ClassicEmailHandler(email_settings)

        original_email = {
            "email_id": "42",
            "subject": "With Attachment",
            "from": "alice@example.com",
            "to": ["test@example.com"],
            "date": datetime(2024, 1, 1, tzinfo=timezone.utc),
            "body": "See attachment",
            "html_body": "",
            "attachments": ["report.pdf"],
        }

        original_attachments = [
            ("report.pdf", "application/pdf", b"%PDF-1.4 fake pdf"),
        ]

        with (
            patch.object(handler.incoming_client, "get_email_body_by_id", return_value=original_email),
            patch.object(handler.incoming_client, "extract_attachments", return_value=original_attachments),
            patch.object(handler.outgoing_client, "send_email", return_value=MagicMock()) as mock_send,
        ):
            await handler.forward_email(
                email_id="42", mailbox="INBOX", recipients=["bob@example.com"]
            )

            # Check extra_parts was passed with the attachment
            call_kwargs = mock_send.call_args.kwargs
            extra_parts = call_kwargs.get("extra_parts")
            assert extra_parts is not None
            assert len(extra_parts) == 1
            assert extra_parts[0].get_filename() == "report.pdf"

    @pytest.mark.asyncio
    async def test_forward_saves_to_sent(self, email_settings):
        """Test that forwarded email is saved to Sent folder when enabled."""
        email_settings.save_to_sent = True
        handler = ClassicEmailHandler(email_settings)

        original_email = {
            "email_id": "42",
            "subject": "Test",
            "from": "alice@example.com",
            "to": ["test@example.com"],
            "date": datetime(2024, 1, 1, tzinfo=timezone.utc),
            "body": "content",
            "html_body": "",
            "attachments": [],
        }

        mock_msg = MagicMock()

        with (
            patch.object(handler.incoming_client, "get_email_body_by_id", return_value=original_email),
            patch.object(handler.incoming_client, "extract_attachments", return_value=[]),
            patch.object(handler.outgoing_client, "send_email", return_value=mock_msg),
            patch.object(handler.outgoing_client, "append_to_sent", return_value=True) as mock_append,
        ):
            await handler.forward_email(
                email_id="42", mailbox="INBOX", recipients=["bob@example.com"]
            )

            mock_append.assert_called_once_with(
                mock_msg,
                email_settings.incoming,
                email_settings.sent_folder_name,
            )

    @pytest.mark.asyncio
    async def test_forward_no_threading_headers(self, email_settings):
        """Test that forwarded emails do not include threading headers."""
        handler = ClassicEmailHandler(email_settings)

        original_email = {
            "email_id": "42",
            "subject": "Test",
            "from": "alice@example.com",
            "to": ["test@example.com"],
            "date": datetime(2024, 1, 1, tzinfo=timezone.utc),
            "body": "content",
            "html_body": "",
            "attachments": [],
        }

        with (
            patch.object(handler.incoming_client, "get_email_body_by_id", return_value=original_email),
            patch.object(handler.incoming_client, "extract_attachments", return_value=[]),
            patch.object(handler.outgoing_client, "send_email", return_value=MagicMock()) as mock_send,
        ):
            await handler.forward_email(
                email_id="42", mailbox="INBOX", recipients=["bob@example.com"]
            )

            # Verify no in_reply_to or references were passed
            call_args = mock_send.call_args
            # in_reply_to and references should not be set (default None)
            assert call_args.kwargs.get("in_reply_to") is None
            assert call_args.kwargs.get("references") is None


class TestDetectEmailService:
    """Tests for _detect_email_service helper."""

    @staticmethod
    def _make_settings(
        imap_host: str = "imap.example.com",
        imap_port: int = 993,
        verify_ssl: bool = True,
        email_service: str | None = None,
    ) -> EmailSettings:
        """Create EmailSettings with only the fields relevant to service detection."""
        return EmailSettings(
            account_name="test",
            full_name="Test",
            email_address="test@example.com",
            incoming=EmailServer(
                user_name="test", password="test", host=imap_host, port=imap_port, verify_ssl=verify_ssl
            ),
            outgoing=EmailServer(
                user_name="test", password="test", host="smtp.example.com", port=465
            ),
            email_service=email_service,
        )

    def test_explicit_override(self):
        """Test that email_service config field takes priority."""
        settings = self._make_settings(imap_host="imap.gmail.com", email_service="protonmail")
        assert _detect_email_service(settings) == "protonmail"

    def test_gmail_host(self):
        """Test Gmail detection via IMAP host."""
        settings = self._make_settings(imap_host="imap.gmail.com")
        assert _detect_email_service(settings) == "gmail"

    def test_gmail_host_case_insensitive(self):
        """Test Gmail detection is case-insensitive."""
        settings = self._make_settings(imap_host="IMAP.GMAIL.COM")
        assert _detect_email_service(settings) == "gmail"

    def test_protonmail_localhost(self):
        """Test ProtonMail Bridge detection via localhost + verify_ssl=False."""
        settings = self._make_settings(imap_host="127.0.0.1", imap_port=1143, verify_ssl=False)
        assert _detect_email_service(settings) == "protonmail"

    def test_protonmail_localhost_name(self):
        """Test ProtonMail Bridge detection via 'localhost' hostname."""
        settings = self._make_settings(imap_host="localhost", imap_port=1143, verify_ssl=False)
        assert _detect_email_service(settings) == "protonmail"

    def test_localhost_with_verify_ssl_true_is_generic(self):
        """Test that localhost with verify_ssl=True does not trigger ProtonMail detection."""
        settings = self._make_settings(imap_host="localhost", verify_ssl=True)
        assert _detect_email_service(settings) == "generic"

    def test_generic_fallback(self):
        """Test generic fallback for unknown hosts."""
        settings = self._make_settings(imap_host="imap.example.com")
        assert _detect_email_service(settings) == "generic"

    def test_explicit_override_beats_gmail(self):
        """Test that explicit override wins over Gmail host detection."""
        settings = self._make_settings(imap_host="imap.gmail.com", email_service="generic")
        assert _detect_email_service(settings) == "generic"


class TestServiceAwareQuoteFormat:
    """Tests for service-specific quote formatting in _format_quoted_reply_html."""

    def _make_original(self):
        return {
            "from": "Alice <alice@example.com>",
            "date": datetime(2024, 3, 15, 14, 30, tzinfo=timezone.utc),
            "body": "Original message body",
            "html_body": "<p>Original <strong>HTML</strong> content</p>",
        }

    def test_protonmail_format(self):
        """Test ProtonMail-specific quote structure."""
        result = _format_quoted_reply_html(self._make_original(), service="protonmail")

        assert 'class="protonmail_quote"' in result
        assert '<div class="protonmail_quote">' in result
        assert '<blockquote class="protonmail_quote" type="cite">' in result
        assert "Alice" in result
        assert "wrote:" in result
        assert "Original <strong>HTML</strong> content" in result

    def test_gmail_format(self):
        """Test Gmail-specific quote structure."""
        result = _format_quoted_reply_html(self._make_original(), service="gmail")

        assert 'class="gmail_quote"' in result
        assert '<div class="gmail_quote">' in result
        assert '<div class="gmail_attr"' in result
        assert '<blockquote class="gmail_quote"' in result
        assert "border-inline-start:1px solid rgb(204,204,204)" in result
        assert "Alice" in result
        assert "wrote:" in result
        assert "Original <strong>HTML</strong> content" in result

    def test_generic_format(self):
        """Test generic quote structure (default)."""
        result = _format_quoted_reply_html(self._make_original(), service="generic")

        assert 'class="protonmail_quote"' not in result
        assert 'class="gmail_quote"' not in result
        assert '<blockquote type="cite"' in result
        assert "Alice" in result
        assert "wrote:" in result
        assert "Original <strong>HTML</strong> content" in result

    def test_default_service_is_generic(self):
        """Test that omitting service parameter gives generic format."""
        result = _format_quoted_reply_html(self._make_original())

        assert 'class="protonmail_quote"' not in result
        assert 'class="gmail_quote"' not in result
        assert '<blockquote type="cite"' in result

    def test_generic_attribution_inside_blockquote(self):
        """Test that generic format puts attribution inside blockquote."""
        result = _format_quoted_reply_html(self._make_original(), service="generic")

        # Attribution should be inside blockquote for generic format
        bq_start = result.index("<blockquote")
        bq_end = result.index("</blockquote>")
        attribution_pos = result.index("wrote:")
        assert bq_start < attribution_pos < bq_end

    def test_protonmail_attribution_outside_blockquote(self):
        """Test that ProtonMail format puts attribution outside blockquote."""
        result = _format_quoted_reply_html(self._make_original(), service="protonmail")

        bq_start = result.index('<blockquote class="protonmail_quote"')
        attribution_pos = result.index("wrote:")
        assert attribution_pos < bq_start

    def test_gmail_attribution_outside_blockquote(self):
        """Test that Gmail format puts attribution outside blockquote."""
        result = _format_quoted_reply_html(self._make_original(), service="gmail")

        bq_start = result.index('<blockquote class="gmail_quote"')
        attribution_pos = result.index("wrote:")
        assert attribution_pos < bq_start

    def test_plain_text_fallback_all_services(self):
        """Test that all service formats handle plain text fallback correctly."""
        original = {
            "from": "Bob <bob@example.com>",
            "date": datetime(2024, 1, 1, tzinfo=timezone.utc),
            "body": "Plain text body\nSecond line",
            "html_body": "",
        }
        for service in ("protonmail", "gmail", "generic"):
            result = _format_quoted_reply_html(original, service=service)
            assert "Plain text body" in result
            assert "Second line" in result
            assert "<br>" in result


class TestClassicEmailHandlerServiceDetection:
    """Tests for ClassicEmailHandler using detected email_service."""

    @staticmethod
    def _make_handler(
        imap_host: str,
        smtp_host: str,
        imap_port: int = 993,
        smtp_port: int = 465,
        verify_ssl: bool = True,
        email_address: str = "test@example.com",
    ) -> ClassicEmailHandler:
        """Create a ClassicEmailHandler with minimal boilerplate."""
        settings = EmailSettings(
            account_name="test",
            full_name="Test",
            email_address=email_address,
            incoming=EmailServer(
                user_name="test", password="test", host=imap_host, port=imap_port, verify_ssl=verify_ssl
            ),
            outgoing=EmailServer(
                user_name="test", password="test", host=smtp_host, port=smtp_port, verify_ssl=verify_ssl
            ),
        )
        return ClassicEmailHandler(settings)

    def test_handler_stores_detected_service(self):
        """Test that ClassicEmailHandler stores the detected service."""
        handler = self._make_handler(
            imap_host="127.0.0.1", smtp_host="127.0.0.1",
            imap_port=1143, smtp_port=1025, verify_ssl=False,
        )
        assert handler.email_service == "protonmail"

    def test_handler_gmail_service(self):
        """Test that ClassicEmailHandler detects Gmail."""
        handler = self._make_handler(imap_host="imap.gmail.com", smtp_host="smtp.gmail.com")
        assert handler.email_service == "gmail"

    @pytest.mark.asyncio
    async def test_send_reply_uses_service_format(self):
        """Test that send_email passes detected service to quote formatter."""
        handler = self._make_handler(
            imap_host="127.0.0.1", smtp_host="127.0.0.1",
            imap_port=1143, smtp_port=1025, verify_ssl=False,
        )

        original_email = {
            "email_id": "42",
            "message_id": "<original@example.com>",
            "subject": "Test",
            "from": "Alice <alice@example.com>",
            "to": ["test@proton.me"],
            "date": datetime(2024, 3, 15, 14, 30, tzinfo=timezone.utc),
            "body": "Original body",
            "html_body": "",
            "attachments": [],
        }

        with (
            patch.object(
                handler.incoming_client, "search_by_message_id", return_value="42"
            ),
            patch.object(
                handler.incoming_client, "get_email_body_by_id", return_value=original_email
            ),
            patch.object(handler.outgoing_client, "send_email", return_value=MagicMock()) as mock_send,
        ):
            await handler.send_email(
                recipients=["alice@example.com"],
                subject="Re: Test",
                body="My reply",
                in_reply_to="<original@example.com>",
                quote_reply=True,
            )

            sent_body = mock_send.call_args[0][2]
            assert 'class="protonmail_quote"' in sent_body
