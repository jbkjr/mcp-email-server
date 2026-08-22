"""Markdown bodies are rendered to email-safe HTML on the way out."""

import pytest

from mcp_email_server import app as app_module
from mcp_email_server.application.mutations import _forwarded_body
from mcp_email_server.config import EmailServer
from mcp_email_server.emails.classic import EmailClient
from mcp_email_server.emails.markdown_utils import markdown_to_email_html, wrap_html_document


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

    def test_wrap_html_document_embeds_the_fragment_verbatim(self):
        """wrap_html_document is the document shell the quoted-reply path reuses."""
        wrapped = wrap_html_document("<p>fragment</p>")

        assert wrapped.startswith("<!DOCTYPE html>")
        assert "<p>fragment</p>" in wrapped
        assert wrapped.rstrip().endswith("</html>")


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


class TestComposeMessageRendersMarkdown:
    """compose_message owns the conversion, so every submission path inherits it."""

    def test_plain_body_becomes_an_html_document(self, email_client):
        message = email_client.compose_message(
            ["dest@example.com"],
            "Subject",
            "This is **bold**.",
        )

        assert message.get_content_type() == "text/html"
        payload = message.get_payload(decode=True).decode("utf-8")
        assert "<strong>bold</strong>" in payload
        assert "<!DOCTYPE html>" in payload

    def test_html_true_is_passed_through_untouched(self, email_client):
        raw = "<html><body><p>already **not** markdown</p></body></html>"
        message = email_client.compose_message(
            ["dest@example.com"],
            "Subject",
            raw,
            html=True,
        )

        assert message.get_content_type() == "text/html"
        assert message.get_payload(decode=True).decode("utf-8") == raw

    def test_conversion_keeps_the_utf8_charset(self, email_client):
        message = email_client.compose_message(["dest@example.com"], "Subject", "café — résumé")

        assert message.get_content_charset() == "utf-8"
        assert "café" in message.get_payload(decode=True).decode("utf-8")

    def test_multipart_body_part_is_rendered_html(self, email_client, tmp_path):
        upload = tmp_path / "note.txt"
        upload.write_text("local file")
        message = email_client.compose_message(
            ["dest@example.com"],
            "Subject",
            "# Heading",
            attachments=[str(upload)],
        )

        assert message.is_multipart()
        body_part = message.get_payload()[0]
        assert body_part.get_content_type() == "text/html"
        assert "<h1>Heading</h1>" in body_part.get_payload(decode=True).decode("utf-8")

    def test_markdown_body_does_not_force_smtputf8(self, email_client):
        """Rendering changes the subtype, never the address or threading headers."""
        message = email_client.compose_message(["dest@example.com"], "Subject", "plain ascii body")

        assert message.get_content_type() == "text/html"
        assert "utf-8" not in str(message["To"])


class TestForwardedBlockIsNotTreatedAsMarkdown:
    """A forwarded body is quoted evidence, so its markup must arrive literally."""

    def test_block_markup_is_escaped(self):
        merged = _forwarded_body("", "plain text with <b>tags</b> and <script>alert(1)</script>")

        assert "<b>" not in merged
        assert "&lt;b&gt;" in merged
        assert "&lt;script&gt;" in merged

    def test_caller_note_is_left_as_markdown(self):
        merged = _forwarded_body("see **this**", "quoted <i>body</i>")

        assert merged.startswith("see **this**\n\n")
        assert "&lt;i&gt;" in merged

    def test_escaped_block_survives_markdown_rendering_as_literal_text(self):
        rendered = markdown_to_email_html(_forwarded_body("", "danger <script>alert(1)</script>"))

        assert "<script>" not in rendered
        assert "&lt;script&gt;" in rendered


class TestServerInstructions:
    """Clients read the server instructions to learn that bodies are Markdown."""

    def test_instructions_describe_markdown_bodies(self):
        instructions = app_module.mcp.instructions

        assert instructions is not None
        assert "Markdown" in instructions
        assert "html=True" in instructions
