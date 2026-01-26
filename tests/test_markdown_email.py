"""Test markdown email functionality."""

from unittest.mock import AsyncMock, patch

import pytest

from mcp_email_server.config import EmailServer
from mcp_email_server.emails.classic import EmailClient
from mcp_email_server.emails.markdown_utils import markdown_to_email_html


class TestMarkdownToEmailHtml:
    """Unit tests for the markdown_to_email_html utility function."""

    def test_basic_formatting(self):
        """Test basic markdown formatting (heading, bold, italic)."""
        text = "# Heading\n\nThis is **bold** and this is *italic*."
        result = markdown_to_email_html(text)

        assert "<h1>Heading</h1>" in result
        assert "<strong>bold</strong>" in result
        assert "<em>italic</em>" in result

    def test_newlines_convert_to_br(self):
        """Test that single newlines convert to <br> (nl2br extension)."""
        text = "Line 1\nLine 2\nLine 3"
        result = markdown_to_email_html(text)

        assert "<br" in result

    def test_unicode_preserved(self):
        """Test that unicode characters like em-dash are preserved."""
        text = "This is an em-dash: — and some accents: café résumé"
        result = markdown_to_email_html(text)

        assert "—" in result
        assert "café" in result
        assert "résumé" in result

    def test_tables_rendered(self):
        """Test that markdown tables are converted to HTML tables."""
        text = """| Header 1 | Header 2 |
| -------- | -------- |
| Cell 1   | Cell 2   |"""
        result = markdown_to_email_html(text)

        assert "<table>" in result
        assert "<th>" in result or "<td>" in result

    def test_fenced_code_blocks(self):
        """Test that fenced code blocks are converted."""
        text = """```python
def hello():
    print("Hello")
```"""
        result = markdown_to_email_html(text)

        assert "<code>" in result or "<pre>" in result

    def test_wrap_in_html_true(self):
        """Test that wrap_in_html=True wraps output in HTML document."""
        text = "Simple text"
        result = markdown_to_email_html(text, wrap_in_html=True)

        assert "<!DOCTYPE html>" in result
        assert "<html>" in result
        assert "<body" in result
        assert "font-family:" in result  # Inline styles present

    def test_wrap_in_html_false(self):
        """Test that wrap_in_html=False returns raw HTML content."""
        text = "Simple text"
        result = markdown_to_email_html(text, wrap_in_html=False)

        assert "<!DOCTYPE html>" not in result
        assert "<html>" not in result
        assert "<p>Simple text</p>" in result

    def test_empty_input(self):
        """Test that empty input produces valid output."""
        result = markdown_to_email_html("")

        assert "<!DOCTYPE html>" in result
        assert "<body" in result

    def test_links(self):
        """Test that markdown links are converted."""
        text = "Visit [Google](https://google.com) for more."
        result = markdown_to_email_html(text)

        assert '<a href="https://google.com">' in result
        assert "Google</a>" in result

    def test_lists(self):
        """Test that markdown lists are converted."""
        text = """- Item 1
- Item 2
- Item 3"""
        result = markdown_to_email_html(text)

        assert "<ul>" in result
        assert "<li>" in result


@pytest.fixture
def email_server():
    return EmailServer(
        user_name="test_user",
        password="test_password",
        host="smtp.example.com",
        port=465,
        use_ssl=True,
    )


@pytest.fixture
def email_client(email_server):
    return EmailClient(email_server, sender="Test User <test@example.com>")


class TestSendEmailWithMarkdown:
    """Integration tests for sending emails with markdown."""

    @pytest.mark.asyncio
    async def test_send_email_with_markdown(self, email_client):
        """Test sending email with markdown=True converts body to HTML."""
        mock_smtp = AsyncMock()
        mock_smtp.__aenter__ = AsyncMock(return_value=mock_smtp)
        mock_smtp.__aexit__ = AsyncMock()

        with patch("mcp_email_server.emails.classic.aiosmtplib.SMTP", return_value=mock_smtp):
            await email_client.send_email(
                recipients=["recipient@example.com"],
                subject="Test markdown email",
                body="This is **bold** and this is *italic*.",
                markdown=True,
            )

            mock_smtp.send_message.assert_called_once()
            message = mock_smtp.send_message.call_args[0][0]

            # Should not be multipart (no attachments)
            assert not message.is_multipart()

            # Check content type is HTML
            assert message.get_content_type() == "text/html"

            # Check body contains converted HTML
            payload = message.get_payload(decode=True).decode("utf-8")
            assert "<strong>bold</strong>" in payload
            assert "<em>italic</em>" in payload

    @pytest.mark.asyncio
    async def test_send_email_markdown_with_attachments(self, email_client, tmp_path):
        """Test sending markdown email with attachments."""
        test_file = tmp_path / "document.pdf"
        test_file.write_bytes(b"PDF content")

        mock_smtp = AsyncMock()
        mock_smtp.__aenter__ = AsyncMock(return_value=mock_smtp)
        mock_smtp.__aexit__ = AsyncMock()

        with patch("mcp_email_server.emails.classic.aiosmtplib.SMTP", return_value=mock_smtp):
            await email_client.send_email(
                recipients=["recipient@example.com"],
                subject="Markdown with attachment",
                body="# Report\n\nPlease see the **attached** document.",
                markdown=True,
                attachments=[str(test_file)],
            )

            mock_smtp.send_message.assert_called_once()
            message = mock_smtp.send_message.call_args[0][0]

            # Should be multipart (has attachments)
            assert message.is_multipart()

            # Get the text/html part (first part should be the body)
            html_part = None
            for part in message.walk():
                if part.get_content_type() == "text/html":
                    html_part = part
                    break

            assert html_part is not None, "No text/html part found"

            # Decode and check the HTML content
            html_content = html_part.get_payload(decode=True).decode("utf-8")
            assert "<h1>Report</h1>" in html_content
            assert "<strong>attached</strong>" in html_content

            # Check attachment is present
            message_str = str(message)
            assert "document.pdf" in message_str

    @pytest.mark.asyncio
    async def test_send_email_markdown_false_unchanged(self, email_client):
        """Test that markdown=False leaves body as plain text (backward compat)."""
        mock_smtp = AsyncMock()
        mock_smtp.__aenter__ = AsyncMock(return_value=mock_smtp)
        mock_smtp.__aexit__ = AsyncMock()

        with patch("mcp_email_server.emails.classic.aiosmtplib.SMTP", return_value=mock_smtp):
            await email_client.send_email(
                recipients=["recipient@example.com"],
                subject="Plain text email",
                body="This is **not** converted to HTML.",
                markdown=False,
            )

            mock_smtp.send_message.assert_called_once()
            message = mock_smtp.send_message.call_args[0][0]

            # Check content type is plain text
            assert message.get_content_type() == "text/plain"

            # Check body is unchanged
            payload = message.get_payload(decode=True).decode("utf-8")
            assert "**not**" in payload
            assert "<strong>" not in payload

    @pytest.mark.asyncio
    async def test_send_email_markdown_with_unicode(self, email_client):
        """Test markdown email preserves unicode characters."""
        mock_smtp = AsyncMock()
        mock_smtp.__aenter__ = AsyncMock(return_value=mock_smtp)
        mock_smtp.__aexit__ = AsyncMock()

        with patch("mcp_email_server.emails.classic.aiosmtplib.SMTP", return_value=mock_smtp):
            await email_client.send_email(
                recipients=["recipient@example.com"],
                subject="Unicode test",
                body="Em-dash: — and accents: café résumé naïve",
                markdown=True,
            )

            mock_smtp.send_message.assert_called_once()
            message = mock_smtp.send_message.call_args[0][0]

            payload = message.get_payload(decode=True).decode("utf-8")
            assert "—" in payload
            assert "café" in payload
            assert "résumé" in payload
            assert "naïve" in payload

    @pytest.mark.asyncio
    async def test_send_email_markdown_overrides_html(self, email_client):
        """Test that markdown=True correctly sets html=True internally."""
        mock_smtp = AsyncMock()
        mock_smtp.__aenter__ = AsyncMock(return_value=mock_smtp)
        mock_smtp.__aexit__ = AsyncMock()

        with patch("mcp_email_server.emails.classic.aiosmtplib.SMTP", return_value=mock_smtp):
            # Even with html=False explicitly, markdown=True should produce HTML
            await email_client.send_email(
                recipients=["recipient@example.com"],
                subject="Test",
                body="**Bold text**",
                html=False,
                markdown=True,
            )

            mock_smtp.send_message.assert_called_once()
            message = mock_smtp.send_message.call_args[0][0]

            # Content type should be HTML because markdown=True
            assert message.get_content_type() == "text/html"

    @pytest.mark.asyncio
    async def test_send_email_empty_body_with_markdown(self, email_client):
        """Test markdown email with empty body."""
        mock_smtp = AsyncMock()
        mock_smtp.__aenter__ = AsyncMock(return_value=mock_smtp)
        mock_smtp.__aexit__ = AsyncMock()

        with patch("mcp_email_server.emails.classic.aiosmtplib.SMTP", return_value=mock_smtp):
            await email_client.send_email(
                recipients=["recipient@example.com"],
                subject="Empty body test",
                body="",
                markdown=True,
            )

            mock_smtp.send_message.assert_called_once()
            message = mock_smtp.send_message.call_args[0][0]

            # Should still be HTML with wrapper
            assert message.get_content_type() == "text/html"
            payload = message.get_payload(decode=True).decode("utf-8")
            assert "<!DOCTYPE html>" in payload
