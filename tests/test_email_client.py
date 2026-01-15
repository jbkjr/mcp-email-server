import asyncio
import email
from datetime import datetime, timezone
from email.mime.text import MIMEText
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from mcp_email_server.config import EmailServer
from mcp_email_server.emails.classic import EmailClient, _has_sort_capability


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

    @pytest.mark.asyncio
    async def test_get_emails_stream(self, email_client):
        """Test getting emails stream with batch fetch optimization."""
        # Mock IMAP client
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.search = AsyncMock(return_value=(None, [b"1 2 3"]))
        mock_imap.uid_search = AsyncMock(return_value=(None, [b"1 2 3"]))
        mock_imap.logout = AsyncMock()

        # Mock protocol.capabilities to not include SORT (test fallback path)
        mock_protocol = MagicMock()
        mock_protocol.capabilities = set()  # No SORT capability
        mock_imap.protocol = mock_protocol

        # Create batch fetch responses for dates and full headers
        # For batch date fetch: returns UID responses followed by date content
        date_response = [
            b"1 FETCH (UID 1 BODY[HEADER.FIELDS (DATE)] {30}",
            bytearray(b"Date: Mon, 1 Jan 2024 00:00:00 +0000\r\n"),
            b"2 FETCH (UID 2 BODY[HEADER.FIELDS (DATE)] {30}",
            bytearray(b"Date: Tue, 2 Jan 2024 00:00:00 +0000\r\n"),
            b"3 FETCH (UID 3 BODY[HEADER.FIELDS (DATE)] {30}",
            bytearray(b"Date: Wed, 3 Jan 2024 00:00:00 +0000\r\n"),
        ]

        # For batch header fetch: returns full headers for each email
        header_response = [
            b"1 FETCH (UID 1 BODY[HEADER] {100}",
            bytearray(
                b"From: sender@example.com\r\nTo: recipient@example.com\r\nSubject: Test Subject 1\r\nDate: Mon, 1 Jan 2024 00:00:00 +0000\r\n\r\n"
            ),
            b"2 FETCH (UID 2 BODY[HEADER] {100}",
            bytearray(
                b"From: sender@example.com\r\nTo: recipient@example.com\r\nSubject: Test Subject 2\r\nDate: Tue, 2 Jan 2024 00:00:00 +0000\r\n\r\n"
            ),
            b"3 FETCH (UID 3 BODY[HEADER] {100}",
            bytearray(
                b"From: sender@example.com\r\nTo: recipient@example.com\r\nSubject: Test Subject 3\r\nDate: Wed, 3 Jan 2024 00:00:00 +0000\r\n\r\n"
            ),
        ]

        # Mock uid to return different responses based on the fetch type
        def uid_side_effect(cmd, uid_list, fetch_type):
            if "HEADER.FIELDS" in fetch_type:
                return (None, date_response)
            else:
                return (None, header_response)

        mock_imap.uid = AsyncMock(side_effect=uid_side_effect)

        # Mock IMAP class
        with patch.object(email_client, "imap_class", return_value=mock_imap):
            emails = []
            async for email_data in email_client.get_emails_metadata_stream(page=1, page_size=10):
                emails.append(email_data)

            # We should get 3 emails (from the mocked search result "1 2 3")
            assert len(emails) == 3
            # With desc order (default), newest first (email 3)
            assert emails[0]["email_id"] == "3"
            assert emails[1]["email_id"] == "2"
            assert emails[2]["email_id"] == "1"

            # Verify IMAP methods were called correctly
            mock_imap.login.assert_called_once_with(
                email_client.email_server.user_name, email_client.email_server.password
            )
            mock_imap.select.assert_called_once_with('"INBOX"')
            mock_imap.uid_search.assert_called_once_with("ALL")
            # Batch fetch: 2 calls (dates + headers) instead of 3 individual calls
            assert mock_imap.uid.call_count == 2
            mock_imap.logout.assert_called_once()

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
        with patch.object(email_client, "imap_class", return_value=mock_imap):
            count = await email_client.get_email_count()

            assert count == 5

            # Verify IMAP methods were called correctly
            mock_imap.login.assert_called_once_with(
                email_client.email_server.user_name, email_client.email_server.password
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
                email_client.email_server.user_name, email_client.email_server.password
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


class TestSortCapability:
    """Tests for IMAP SORT capability detection."""

    def test_has_sort_capability_true(self):
        """Test detection when SORT is available."""
        mock_imap = MagicMock()
        mock_imap.protocol.capabilities = {"SORT", "IMAP4rev1", "IDLE"}
        assert _has_sort_capability(mock_imap) is True

    def test_has_sort_capability_false(self):
        """Test detection when SORT is not available."""
        mock_imap = MagicMock()
        mock_imap.protocol.capabilities = {"IMAP4rev1", "IDLE"}
        assert _has_sort_capability(mock_imap) is False

    def test_has_sort_capability_empty(self):
        """Test detection with empty capabilities."""
        mock_imap = MagicMock()
        mock_imap.protocol.capabilities = set()
        assert _has_sort_capability(mock_imap) is False

    def test_has_sort_capability_exception(self):
        """Test graceful handling when capabilities check fails."""
        mock_imap = MagicMock()
        mock_imap.protocol = MagicMock()
        type(mock_imap.protocol).capabilities = property(lambda self: (_ for _ in ()).throw(Exception("error")))
        assert _has_sort_capability(mock_imap) is False


class TestParseDateFromHeader:
    """Tests for date parsing helper."""

    def test_parse_valid_date(self, email_client):
        """Test parsing a valid RFC 2822 date."""
        date_str = "Mon, 1 Jan 2024 12:00:00 +0000"
        result = email_client._parse_date_from_header(date_str)
        assert result.year == 2024
        assert result.month == 1
        assert result.day == 1

    def test_parse_invalid_date(self, email_client):
        """Test parsing an invalid date returns current time."""
        date_str = "not a valid date"
        result = email_client._parse_date_from_header(date_str)
        # Should return a datetime close to now
        assert isinstance(result, datetime)
        assert result.tzinfo == timezone.utc

    def test_parse_empty_date(self, email_client):
        """Test parsing empty string returns current time."""
        result = email_client._parse_date_from_header("")
        assert isinstance(result, datetime)


class TestBatchFetchDates:
    """Tests for batch date fetching."""

    @pytest.mark.asyncio
    async def test_batch_fetch_dates_success(self, email_client):
        """Test successful batch fetching of dates."""
        mock_imap = AsyncMock()
        date_response = [
            b"1 FETCH (UID 1 BODY[HEADER.FIELDS (DATE)] {30}",
            bytearray(b"Date: Mon, 1 Jan 2024 00:00:00 +0000\r\n"),
            b"2 FETCH (UID 2 BODY[HEADER.FIELDS (DATE)] {30}",
            bytearray(b"Date: Tue, 2 Jan 2024 00:00:00 +0000\r\n"),
        ]
        mock_imap.uid = AsyncMock(return_value=(None, date_response))

        result = await email_client._batch_fetch_dates(mock_imap, [b"1", b"2"])

        assert len(result) == 2
        assert result[0][0] == "1"
        assert result[1][0] == "2"

    @pytest.mark.asyncio
    async def test_batch_fetch_dates_empty_list(self, email_client):
        """Test batch fetch with empty email list."""
        mock_imap = AsyncMock()
        result = await email_client._batch_fetch_dates(mock_imap, [])
        assert result == []

    @pytest.mark.asyncio
    async def test_batch_fetch_dates_error(self, email_client):
        """Test batch fetch handles errors gracefully."""
        mock_imap = AsyncMock()
        mock_imap.uid = AsyncMock(side_effect=Exception("Network error"))

        result = await email_client._batch_fetch_dates(mock_imap, [b"1", b"2"])
        assert result == []

    @pytest.mark.asyncio
    async def test_batch_fetch_dates_with_date_prefix(self, email_client):
        """Test parsing dates that include 'Date:' prefix."""
        mock_imap = AsyncMock()
        date_response = [
            b"1 FETCH (UID 1 BODY[HEADER.FIELDS (DATE)] {30}",
            bytearray(b"Date: Mon, 1 Jan 2024 00:00:00 +0000\r\n"),
        ]
        mock_imap.uid = AsyncMock(return_value=(None, date_response))

        result = await email_client._batch_fetch_dates(mock_imap, [b"1"])

        assert len(result) == 1
        assert result[0][1].year == 2024

    @pytest.mark.asyncio
    async def test_batch_fetch_dates_uid_after_data(self, email_client):
        """Test batch fetch dates when UID comes after data (Proton Bridge format)."""
        mock_imap = AsyncMock()
        # Proton Bridge returns date data first, then UID in a separate item
        date_response = [
            b"1 FETCH (BODY[HEADER.FIELDS (DATE)] {30}",
            bytearray(b"Date: Mon, 1 Jan 2024 00:00:00 +0000\r\n"),
            b" UID 1)",
            b"2 FETCH (BODY[HEADER.FIELDS (DATE)] {30}",
            bytearray(b"Tue, 2 Jan 2024 00:00:00 +0000\r\n"),  # Without "Date:" prefix
            b" UID 2)",
        ]
        mock_imap.uid = AsyncMock(return_value=(None, date_response))

        result = await email_client._batch_fetch_dates(mock_imap, [b"1", b"2"])

        assert len(result) == 2
        assert result[0][0] == "1"
        assert result[0][1].year == 2024
        assert result[1][0] == "2"


class TestBatchFetchHeaders:
    """Tests for batch header fetching."""

    @pytest.mark.asyncio
    async def test_batch_fetch_headers_success(self, email_client):
        """Test successful batch fetching of headers."""
        mock_imap = AsyncMock()
        header_response = [
            b"1 FETCH (UID 1 BODY[HEADER] {100}",
            bytearray(
                b"From: sender@example.com\r\nTo: recipient@example.com\r\nSubject: Test 1\r\nDate: Mon, 1 Jan 2024 00:00:00 +0000\r\n\r\n"
            ),
            b"2 FETCH (UID 2 BODY[HEADER] {100}",
            bytearray(
                b"From: sender@example.com\r\nTo: recipient@example.com\r\nSubject: Test 2\r\nDate: Tue, 2 Jan 2024 00:00:00 +0000\r\n\r\n"
            ),
        ]
        mock_imap.uid = AsyncMock(return_value=(None, header_response))

        result = await email_client._batch_fetch_headers(mock_imap, ["1", "2"])

        assert len(result) == 2
        assert result[0]["email_id"] == "1"
        assert result[0]["subject"] == "Test 1"
        assert result[1]["email_id"] == "2"
        assert result[1]["subject"] == "Test 2"

    @pytest.mark.asyncio
    async def test_batch_fetch_headers_empty_list(self, email_client):
        """Test batch fetch with empty email list."""
        mock_imap = AsyncMock()
        result = await email_client._batch_fetch_headers(mock_imap, [])
        assert result == []

    @pytest.mark.asyncio
    async def test_batch_fetch_headers_error(self, email_client):
        """Test batch fetch handles errors gracefully."""
        mock_imap = AsyncMock()
        mock_imap.uid = AsyncMock(side_effect=Exception("Network error"))

        result = await email_client._batch_fetch_headers(mock_imap, ["1", "2"])
        assert result == []

    @pytest.mark.asyncio
    async def test_batch_fetch_headers_uid_after_data(self, email_client):
        """Test batch fetch headers when UID comes after data (Proton Bridge format)."""
        mock_imap = AsyncMock()
        # Proton Bridge returns header data first, then UID in a separate item
        header_response = [
            b"1 FETCH (BODY[HEADER] {100}",
            bytearray(
                b"From: sender@example.com\r\nTo: recipient@example.com\r\nSubject: Test 1\r\nDate: Mon, 1 Jan 2024 00:00:00 +0000\r\n\r\n"
            ),
            b" UID 1)",
            b"2 FETCH (BODY[HEADER] {100}",
            bytearray(
                b"From: sender@example.com\r\nTo: recipient@example.com\r\nSubject: Test 2\r\nDate: Tue, 2 Jan 2024 00:00:00 +0000\r\n\r\n"
            ),
            b" UID 2)",
        ]
        mock_imap.uid = AsyncMock(return_value=(None, header_response))

        result = await email_client._batch_fetch_headers(mock_imap, ["1", "2"])

        assert len(result) == 2
        assert result[0]["email_id"] == "1"
        assert result[0]["subject"] == "Test 1"
        assert result[1]["email_id"] == "2"
        assert result[1]["subject"] == "Test 2"

    @pytest.mark.asyncio
    async def test_batch_fetch_headers_bytes_without_uid(self, email_client):
        """Test batch fetch headers ignores bytes items without UID match."""
        mock_imap = AsyncMock()
        # Include some bytes items that don't have UID (like status messages)
        header_response = [
            b"* OK Still here",  # No UID - should be skipped
            b"1 FETCH (UID 1 BODY[HEADER] {100}",
            bytearray(
                b"From: sender@example.com\r\nTo: recipient@example.com\r\nSubject: Test\r\nDate: Mon, 1 Jan 2024 00:00:00 +0000\r\n\r\n"
            ),
            b"FETCH completed",  # No UID - should be skipped
        ]
        mock_imap.uid = AsyncMock(return_value=(None, header_response))

        result = await email_client._batch_fetch_headers(mock_imap, ["1"])

        assert len(result) == 1
        assert result[0]["email_id"] == "1"


class TestParseHeaderToMetadata:
    """Tests for header parsing helper."""

    def test_parse_header_success(self, email_client):
        """Test successful header parsing."""
        raw_headers = b"From: sender@example.com\r\nTo: recipient@example.com\r\nSubject: Test\r\nDate: Mon, 1 Jan 2024 00:00:00 +0000\r\n\r\n"
        result = email_client._parse_header_to_metadata("123", raw_headers)

        assert result["email_id"] == "123"
        assert result["from"] == "sender@example.com"
        assert result["subject"] == "Test"
        assert result["to"] == ["recipient@example.com"]
        assert result["attachments"] == []

    def test_parse_header_with_cc(self, email_client):
        """Test parsing headers with CC recipients."""
        raw_headers = b"From: sender@example.com\r\nTo: to@example.com\r\nCc: cc1@example.com, cc2@example.com\r\nSubject: Test\r\nDate: Mon, 1 Jan 2024 00:00:00 +0000\r\n\r\n"
        result = email_client._parse_header_to_metadata("123", raw_headers)

        assert "to@example.com" in result["to"]
        assert "cc1@example.com" in result["to"]
        assert "cc2@example.com" in result["to"]

    def test_parse_header_invalid(self, email_client):
        """Test parsing invalid headers returns None."""
        # Create a scenario where parsing fails
        with patch("mcp_email_server.emails.classic.BytesParser") as mock_parser:
            mock_parser.return_value.parsebytes.side_effect = Exception("Parse error")
            result = email_client._parse_header_to_metadata("123", b"invalid")
            assert result is None


class TestGetEmailsStreamWithSort:
    """Tests for get_emails_metadata_stream with SORT capability."""

    @pytest.mark.asyncio
    async def test_get_emails_stream_with_sort(self, email_client):
        """Test getting emails using IMAP SORT when available."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.logout = AsyncMock()

        # Mock protocol.capabilities to include SORT
        mock_protocol = MagicMock()
        mock_protocol.capabilities = {"SORT", "IMAP4rev1"}
        mock_imap.protocol = mock_protocol

        # Mock SORT response (already sorted by date desc)
        sort_response = [b"3 2 1"]  # UIDs in sorted order

        # Mock header fetch for the page
        header_response = [
            b"3 FETCH (UID 3 BODY[HEADER] {100}",
            bytearray(
                b"From: sender@example.com\r\nTo: recipient@example.com\r\nSubject: Test 3\r\nDate: Wed, 3 Jan 2024 00:00:00 +0000\r\n\r\n"
            ),
            b"2 FETCH (UID 2 BODY[HEADER] {100}",
            bytearray(
                b"From: sender@example.com\r\nTo: recipient@example.com\r\nSubject: Test 2\r\nDate: Tue, 2 Jan 2024 00:00:00 +0000\r\n\r\n"
            ),
            b"1 FETCH (UID 1 BODY[HEADER] {100}",
            bytearray(
                b"From: sender@example.com\r\nTo: recipient@example.com\r\nSubject: Test 1\r\nDate: Mon, 1 Jan 2024 00:00:00 +0000\r\n\r\n"
            ),
        ]

        def uid_side_effect(cmd, *args):
            if cmd == "sort":
                return (None, sort_response)
            else:  # fetch
                return (None, header_response)

        mock_imap.uid = AsyncMock(side_effect=uid_side_effect)

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            emails = []
            async for email_data in email_client.get_emails_metadata_stream(page=1, page_size=10):
                emails.append(email_data)

            assert len(emails) == 3
            # Should be in SORT order (3, 2, 1 for desc)
            assert emails[0]["email_id"] == "3"
            assert emails[1]["email_id"] == "2"
            assert emails[2]["email_id"] == "1"

            # Verify SORT was called
            calls = mock_imap.uid.call_args_list
            assert calls[0][0][0] == "sort"

    @pytest.mark.asyncio
    async def test_get_emails_stream_sort_fallback_on_error(self, email_client):
        """Test fallback to batch fetch when SORT fails."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.uid_search = AsyncMock(return_value=(None, [b"1 2"]))
        mock_imap.logout = AsyncMock()

        # Mock protocol.capabilities to include SORT
        mock_protocol = MagicMock()
        mock_protocol.capabilities = {"SORT", "IMAP4rev1"}
        mock_imap.protocol = mock_protocol

        call_count = [0]

        # Mock responses - SORT fails, then fallback works
        date_response = [
            b"1 FETCH (UID 1 BODY[HEADER.FIELDS (DATE)] {30}",
            bytearray(b"Date: Mon, 1 Jan 2024 00:00:00 +0000\r\n"),
            b"2 FETCH (UID 2 BODY[HEADER.FIELDS (DATE)] {30}",
            bytearray(b"Date: Tue, 2 Jan 2024 00:00:00 +0000\r\n"),
        ]
        header_response = [
            b"2 FETCH (UID 2 BODY[HEADER] {100}",
            bytearray(
                b"From: sender@example.com\r\nTo: recipient@example.com\r\nSubject: Test 2\r\nDate: Tue, 2 Jan 2024 00:00:00 +0000\r\n\r\n"
            ),
            b"1 FETCH (UID 1 BODY[HEADER] {100}",
            bytearray(
                b"From: sender@example.com\r\nTo: recipient@example.com\r\nSubject: Test 1\r\nDate: Mon, 1 Jan 2024 00:00:00 +0000\r\n\r\n"
            ),
        ]

        def uid_side_effect(cmd, *args):
            call_count[0] += 1
            if cmd == "sort":
                raise RuntimeError("SORT not supported")
            elif "HEADER.FIELDS" in args[-1] if args else False:
                return (None, date_response)
            else:
                return (None, header_response)

        mock_imap.uid = AsyncMock(side_effect=uid_side_effect)

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            emails = []
            async for email_data in email_client.get_emails_metadata_stream(page=1, page_size=10):
                emails.append(email_data)

            # Should still get results via fallback
            assert len(emails) == 2

    @pytest.mark.asyncio
    async def test_get_emails_stream_empty_search(self, email_client):
        """Test handling of empty search results."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.uid_search = AsyncMock(return_value=(None, [b""]))
        mock_imap.logout = AsyncMock()

        mock_protocol = MagicMock()
        mock_protocol.capabilities = set()
        mock_imap.protocol = mock_protocol

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            emails = []
            async for email_data in email_client.get_emails_metadata_stream(page=1, page_size=10):
                emails.append(email_data)

            assert len(emails) == 0

    @pytest.mark.asyncio
    async def test_get_emails_stream_asc_order(self, email_client):
        """Test getting emails in ascending order."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.uid_search = AsyncMock(return_value=(None, [b"1 2 3"]))
        mock_imap.logout = AsyncMock()

        mock_protocol = MagicMock()
        mock_protocol.capabilities = set()
        mock_imap.protocol = mock_protocol

        date_response = [
            b"1 FETCH (UID 1 BODY[HEADER.FIELDS (DATE)] {30}",
            bytearray(b"Date: Mon, 1 Jan 2024 00:00:00 +0000\r\n"),
            b"2 FETCH (UID 2 BODY[HEADER.FIELDS (DATE)] {30}",
            bytearray(b"Date: Tue, 2 Jan 2024 00:00:00 +0000\r\n"),
            b"3 FETCH (UID 3 BODY[HEADER.FIELDS (DATE)] {30}",
            bytearray(b"Date: Wed, 3 Jan 2024 00:00:00 +0000\r\n"),
        ]
        header_response = [
            b"1 FETCH (UID 1 BODY[HEADER] {100}",
            bytearray(
                b"From: sender@example.com\r\nTo: recipient@example.com\r\nSubject: Test 1\r\nDate: Mon, 1 Jan 2024 00:00:00 +0000\r\n\r\n"
            ),
            b"2 FETCH (UID 2 BODY[HEADER] {100}",
            bytearray(
                b"From: sender@example.com\r\nTo: recipient@example.com\r\nSubject: Test 2\r\nDate: Tue, 2 Jan 2024 00:00:00 +0000\r\n\r\n"
            ),
            b"3 FETCH (UID 3 BODY[HEADER] {100}",
            bytearray(
                b"From: sender@example.com\r\nTo: recipient@example.com\r\nSubject: Test 3\r\nDate: Wed, 3 Jan 2024 00:00:00 +0000\r\n\r\n"
            ),
        ]

        def uid_side_effect(cmd, uid_list, fetch_type):
            if "HEADER.FIELDS" in fetch_type:
                return (None, date_response)
            else:
                return (None, header_response)

        mock_imap.uid = AsyncMock(side_effect=uid_side_effect)

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            emails = []
            async for email_data in email_client.get_emails_metadata_stream(page=1, page_size=10, order="asc"):
                emails.append(email_data)

            assert len(emails) == 3
            # Ascending order: oldest first
            assert emails[0]["email_id"] == "1"
            assert emails[1]["email_id"] == "2"
            assert emails[2]["email_id"] == "3"

    @pytest.mark.asyncio
    async def test_get_emails_stream_pagination(self, email_client):
        """Test pagination works correctly."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.uid_search = AsyncMock(return_value=(None, [b"1 2 3 4 5"]))
        mock_imap.logout = AsyncMock()

        mock_protocol = MagicMock()
        mock_protocol.capabilities = set()
        mock_imap.protocol = mock_protocol

        date_response = [
            b"1 FETCH (UID 1 BODY[HEADER.FIELDS (DATE)] {30}",
            bytearray(b"Date: Mon, 1 Jan 2024 00:00:00 +0000\r\n"),
            b"2 FETCH (UID 2 BODY[HEADER.FIELDS (DATE)] {30}",
            bytearray(b"Date: Tue, 2 Jan 2024 00:00:00 +0000\r\n"),
            b"3 FETCH (UID 3 BODY[HEADER.FIELDS (DATE)] {30}",
            bytearray(b"Date: Wed, 3 Jan 2024 00:00:00 +0000\r\n"),
            b"4 FETCH (UID 4 BODY[HEADER.FIELDS (DATE)] {30}",
            bytearray(b"Date: Thu, 4 Jan 2024 00:00:00 +0000\r\n"),
            b"5 FETCH (UID 5 BODY[HEADER.FIELDS (DATE)] {30}",
            bytearray(b"Date: Fri, 5 Jan 2024 00:00:00 +0000\r\n"),
        ]
        # Only return headers for page 2 (emails 3 and 2 in desc order)
        header_response = [
            b"3 FETCH (UID 3 BODY[HEADER] {100}",
            bytearray(
                b"From: sender@example.com\r\nTo: recipient@example.com\r\nSubject: Test 3\r\nDate: Wed, 3 Jan 2024 00:00:00 +0000\r\n\r\n"
            ),
            b"2 FETCH (UID 2 BODY[HEADER] {100}",
            bytearray(
                b"From: sender@example.com\r\nTo: recipient@example.com\r\nSubject: Test 2\r\nDate: Tue, 2 Jan 2024 00:00:00 +0000\r\n\r\n"
            ),
        ]

        def uid_side_effect(cmd, uid_list, fetch_type):
            if "HEADER.FIELDS" in fetch_type:
                return (None, date_response)
            else:
                return (None, header_response)

        mock_imap.uid = AsyncMock(side_effect=uid_side_effect)

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            emails = []
            # Page 2, page_size 2 with desc order means emails at positions 2-3 (0-indexed)
            # With 5 emails sorted desc: 5,4,3,2,1 -> page 2 gets 3,2
            async for email_data in email_client.get_emails_metadata_stream(page=2, page_size=2):
                emails.append(email_data)

            assert len(emails) == 2
            assert emails[0]["email_id"] == "3"
            assert emails[1]["email_id"] == "2"

    @pytest.mark.asyncio
    async def test_get_emails_stream_date_fetch_fallback(self, email_client):
        """Test fallback to full header fetch when date fetch fails."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.uid_search = AsyncMock(return_value=(None, [b"1 2"]))
        mock_imap.logout = AsyncMock()

        mock_protocol = MagicMock()
        mock_protocol.capabilities = set()
        mock_imap.protocol = mock_protocol

        header_response = [
            b"1 FETCH (UID 1 BODY[HEADER] {100}",
            bytearray(
                b"From: sender@example.com\r\nTo: recipient@example.com\r\nSubject: Test 1\r\nDate: Mon, 1 Jan 2024 00:00:00 +0000\r\n\r\n"
            ),
            b"2 FETCH (UID 2 BODY[HEADER] {100}",
            bytearray(
                b"From: sender@example.com\r\nTo: recipient@example.com\r\nSubject: Test 2\r\nDate: Tue, 2 Jan 2024 00:00:00 +0000\r\n\r\n"
            ),
        ]

        call_count = [0]

        def uid_side_effect(cmd, uid_list, fetch_type):
            call_count[0] += 1
            if "HEADER.FIELDS" in fetch_type:
                # Return empty to trigger fallback
                return (None, [])
            else:
                return (None, header_response)

        mock_imap.uid = AsyncMock(side_effect=uid_side_effect)

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            emails = []
            async for email_data in email_client.get_emails_metadata_stream(page=1, page_size=10):
                emails.append(email_data)

            # Should still get results via fallback
            assert len(emails) == 2


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

        with patch.object(email_client, "imap_class", return_value=mock_imap):
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

        with patch.object(email_client, "imap_class", return_value=mock_imap):
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

        with patch.object(email_client, "imap_class", return_value=mock_imap):
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

        with patch.object(email_client, "imap_class", return_value=mock_imap):
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

        with patch.object(email_client, "imap_class", return_value=mock_imap):
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

        with patch.object(email_client, "imap_class", return_value=mock_imap):
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

        with patch.object(email_client, "imap_class", return_value=mock_imap):
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

        with patch.object(email_client, "imap_class", return_value=mock_imap):
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

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            # Should complete successfully despite logout error
            marked_ids, failed_ids = await email_client.mark_emails(
                email_ids=["123"],
                mark_as="read",
                mailbox="INBOX",
            )
            assert marked_ids == ["123"]
            assert failed_ids == []
