"""A reply quotes the message it answers, or fails rather than quietly dropping it."""

import asyncio
from datetime import UTC, datetime
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from mcp_email_server.adapters.mutations import ClassicMutationProvider
from mcp_email_server.application.mutations import (
    MutationAccountSnapshot,
    MutationProviderError,
    QuoteSource,
    SendCommand,
    SendService,
    _quoted_body,
)
from mcp_email_server.config import EmailServer, EmailSettings, Settings
from mcp_email_server.emails.classic import (
    MAX_QUOTED_BODY_LENGTH,
    MAX_RAW_EMAIL_BYTES,
    ClassicEmailHandler,
    EmailClient,
    _detect_email_service,
    _format_quoted_reply_html,
    _strip_html_wrappers,
)
from mcp_email_server.emails.markdown_utils import markdown_to_email_html


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
            user_name="test",
            password="test",
            host=imap_host,
            port=imap_port,
            verify_ssl=verify_ssl,
        ),
        outgoing=EmailServer(user_name="test", password="test", host="smtp.example.com", port=465),
        email_service=email_service,
    )


class TestDetectEmailService:
    """Tests for _detect_email_service helper."""

    def test_explicit_override(self):
        """Test that email_service config field takes priority."""
        assert _detect_email_service(_make_settings(imap_host="imap.gmail.com", email_service="protonmail")) == (
            "protonmail"
        )

    def test_gmail_host(self):
        """Test Gmail detection via IMAP host."""
        assert _detect_email_service(_make_settings(imap_host="imap.gmail.com")) == "gmail"

    def test_gmail_host_case_insensitive(self):
        """Test Gmail detection is case-insensitive."""
        assert _detect_email_service(_make_settings(imap_host="IMAP.GMAIL.COM")) == "gmail"

    def test_protonmail_localhost(self):
        """Test ProtonMail Bridge detection via localhost + verify_ssl=False."""
        settings = _make_settings(imap_host="127.0.0.1", imap_port=1143, verify_ssl=False)
        assert _detect_email_service(settings) == "protonmail"

    def test_protonmail_localhost_name(self):
        """Test ProtonMail Bridge detection via 'localhost' hostname."""
        settings = _make_settings(imap_host="localhost", imap_port=1143, verify_ssl=False)
        assert _detect_email_service(settings) == "protonmail"

    def test_localhost_with_verify_ssl_true_is_generic(self):
        """Test that localhost with verify_ssl=True does not trigger ProtonMail detection."""
        assert _detect_email_service(_make_settings(imap_host="localhost", verify_ssl=True)) == "generic"

    def test_generic_fallback(self):
        """Test generic fallback for unknown hosts."""
        assert _detect_email_service(_make_settings(imap_host="imap.example.com")) == "generic"

    def test_explicit_override_beats_gmail(self):
        """Test that explicit override wins over Gmail host detection."""
        settings = _make_settings(imap_host="imap.gmail.com", email_service="generic")
        assert _detect_email_service(settings) == "generic"


class TestFormatQuotedReplyHtml:
    """Tests for _format_quoted_reply_html helper."""

    def test_html_body_preserved_in_blockquote(self):
        """Test that original HTML body is preserved inside blockquote."""
        original = {
            "from": "Alice <alice@example.com>",
            "date": datetime(2024, 3, 15, 14, 30, tzinfo=UTC),
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
            "date": datetime(2024, 3, 15, 14, 30, tzinfo=UTC),
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
            "date": datetime(2024, 1, 1, tzinfo=UTC),
            "body": "",
            "html_body": "<!DOCTYPE html><html><head><style>h1{color:red}</style></head><body><p>C</p></body></html>",
        }
        result = _format_quoted_reply_html(original)

        assert "<!DOCTYPE" not in result
        assert "<html>" not in result
        assert "<head>" not in result
        assert "<body>" not in result
        assert "<p>C</p>" in result

    def test_escapes_sender_in_attribution(self):
        """Test that sender name is HTML-escaped in attribution line."""
        original = {
            "from": "Evil <script>alert('xss')</script>",
            "date": datetime(2024, 1, 1, tzinfo=UTC),
            "body": "test",
            "html_body": "",
        }
        result = _format_quoted_reply_html(original)

        assert "<script>" not in result
        assert "&lt;script&gt;" in result

    def test_long_text_body_truncation(self):
        """Test that long plain text bodies are truncated in HTML quote."""
        original = {
            "from": "sender@example.com",
            "date": datetime(2024, 1, 1, tzinfo=UTC),
            "body": "x" * (MAX_QUOTED_BODY_LENGTH + 1000),
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
            "date": datetime(2024, 1, 1, tzinfo=UTC),
            "body": "",
            "html_body": "",
        }
        result = _format_quoted_reply_html(original)

        assert "<blockquote" in result
        assert "</blockquote>" in result

    def test_strip_html_wrappers_keeps_inner_markup(self):
        stripped = _strip_html_wrappers("<!DOCTYPE html><HTML><BODY class='x'><p>keep</p></BODY></HTML>")

        assert stripped == "<p>keep</p>"


class TestServiceAwareQuoteFormat:
    """Tests for service-specific quote formatting in _format_quoted_reply_html."""

    def _make_original(self):
        return {
            "from": "Alice <alice@example.com>",
            "date": datetime(2024, 3, 15, 14, 30, tzinfo=UTC),
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
            "date": datetime(2024, 1, 1, tzinfo=UTC),
            "body": "Plain text body\nSecond line",
            "html_body": "",
        }
        for service in ("protonmail", "gmail", "generic"):
            result = _format_quoted_reply_html(original, service=service)
            assert "Plain text body" in result
            assert "Second line" in result
            assert "<br>" in result


class TestQuotedBodyMerge:
    """The merged body is what actually reaches composition."""

    def test_quote_follows_the_caller_body(self):
        merged = _quoted_body("my reply", "<div>quote</div>")

        assert merged == "my reply\n\n<div>quote</div>"

    def test_empty_body_leaves_the_quote_alone(self):
        assert _quoted_body("", "<div>quote</div>") == "<div>quote</div>"

    def test_quote_markup_survives_markdown_rendering(self):
        """Markdown passes block-level HTML through, so the quote arrives intact."""
        original = {
            "from": "Alice <alice@example.com>",
            "date": datetime(2024, 3, 15, 14, 30, tzinfo=UTC),
            "body": "Original body",
            "html_body": "",
        }
        quote = _format_quoted_reply_html(original, service="generic")
        rendered = markdown_to_email_html(_quoted_body("My **reply**", quote), wrap_in_html=False)

        assert "<strong>reply</strong>" in rendered
        assert '<blockquote type="cite"' in rendered
        assert "Original body" in rendered


def _snapshot(allowed_senders: tuple[str, ...] = ()) -> MutationAccountSnapshot:
    return MutationAccountSnapshot(
        account_name="test",
        mode="legacy",
        allowed_senders=allowed_senders,
        allowed_recipients=(),
        report_blocked_mutations=False,
        can_send=True,
    )


class _StubProvider:
    """Minimal MutationProvider double covering only the send/quote pair."""

    def __init__(self, quote: QuoteSource | None = None, error: Exception | None = None) -> None:
        self._quote = quote
        self._error = error
        self.quote_calls = 0
        self.sent: SendCommand | None = None

    async def fetch_quote_source(self, command: SendCommand, account: MutationAccountSnapshot):
        del command, account
        self.quote_calls += 1
        if self._error is not None:
            raise self._error
        return self._quote

    async def send(self, command: SendCommand, account: MutationAccountSnapshot):
        del account
        self.sent = command
        return MagicMock(outcomes=(), sent_message=None, has_accepted_recipient=False)


def _service_with(provider: _StubProvider) -> SendService:
    account = _snapshot()
    accounts = MagicMock()
    accounts.resolve.return_value = account
    providers = MagicMock()
    providers.open.return_value = MagicMock(account=account, provider=provider)
    return SendService(accounts, providers, MagicMock())


class TestSendServiceQuoting:
    """The workflow decides whether to quote, and never sends a silently unquoted reply."""

    @pytest.mark.asyncio
    async def test_reply_appends_the_quote_before_sending(self):
        provider = _StubProvider(QuoteSource(quote_html="<div>QUOTE</div>"))
        command = SendCommand(
            account_name="test",
            recipients=("alice@example.com",),
            subject="Re: hello",
            body="my reply",
            in_reply_to="<original@example.com>",
        )

        await _service_with(provider).execute(command)

        assert provider.sent is not None
        assert provider.sent.body == "my reply\n\n<div>QUOTE</div>"

    @pytest.mark.asyncio
    async def test_quote_reply_false_never_reads_the_original(self):
        provider = _StubProvider(QuoteSource(quote_html="<div>QUOTE</div>"))
        command = SendCommand(
            account_name="test",
            recipients=("alice@example.com",),
            subject="Re: hello",
            body="my reply",
            in_reply_to="<original@example.com>",
            quote_reply=False,
        )

        await _service_with(provider).execute(command)

        assert provider.quote_calls == 0
        assert provider.sent is not None
        assert provider.sent.body == "my reply"

    @pytest.mark.asyncio
    async def test_a_send_that_is_not_a_reply_never_reads_anything(self):
        provider = _StubProvider(QuoteSource(quote_html="<div>QUOTE</div>"))
        command = SendCommand(
            account_name="test",
            recipients=("alice@example.com",),
            subject="hello",
            body="not a reply",
        )

        await _service_with(provider).execute(command)

        assert provider.quote_calls == 0

    @pytest.mark.asyncio
    async def test_missing_original_degrades_to_an_unquoted_reply(self):
        provider = _StubProvider(quote=None)
        command = SendCommand(
            account_name="test",
            recipients=("alice@example.com",),
            subject="Re: hello",
            body="my reply",
            in_reply_to="<gone@example.com>",
        )

        await _service_with(provider).execute(command)

        assert provider.quote_calls == 1
        assert provider.sent is not None
        assert provider.sent.body == "my reply"

    @pytest.mark.asyncio
    async def test_unreadable_original_aborts_before_smtp(self):
        provider = _StubProvider(error=ValueError("Failed to fetch quote source with UID 42"))
        command = SendCommand(
            account_name="test",
            recipients=("alice@example.com",),
            subject="Re: hello",
            body="my reply",
            in_reply_to="<original@example.com>",
        )

        with pytest.raises(ValueError):
            await _service_with(provider).execute(command)

        assert provider.sent is None

    @pytest.mark.asyncio
    async def test_quote_read_timeout_aborts_before_smtp(self):
        provider = _StubProvider(error=TimeoutError())
        command = SendCommand(
            account_name="test",
            recipients=("alice@example.com",),
            subject="Re: hello",
            body="my reply",
            in_reply_to="<original@example.com>",
        )

        with pytest.raises(MutationProviderError):
            await _service_with(provider).execute(command)

        assert provider.sent is None

    @pytest.mark.asyncio
    async def test_merged_body_is_revalidated_against_the_body_bound(self):
        """A quote cannot push a reply past the limit unnoticed."""
        provider = _StubProvider(QuoteSource(quote_html="<div>" + "q" * 2_000_000 + "</div>"))
        command = SendCommand(
            account_name="test",
            recipients=("alice@example.com",),
            subject="Re: hello",
            body="my reply",
            in_reply_to="<original@example.com>",
        )

        with pytest.raises(ValueError):
            await _service_with(provider).execute(command)

        assert provider.sent is None


class TestQuoteSourceMailboxResolution:
    """INBOX first, then the Sent folder, and never fail a reply over the lookup."""

    @pytest.mark.asyncio
    async def test_inbox_then_sent(self):
        handler = ClassicEmailHandler(_make_settings())
        with patch.object(handler, "_find_special_folder", AsyncMock(return_value="Sent")):
            assert await handler.quote_source_mailboxes() == ("INBOX", "Sent")

    @pytest.mark.asyncio
    async def test_unresolvable_sent_folder_leaves_inbox_alone(self):
        handler = ClassicEmailHandler(_make_settings())
        with patch.object(handler, "_find_special_folder", AsyncMock(return_value=None)):
            assert await handler.quote_source_mailboxes() == ("INBOX",)

    @pytest.mark.asyncio
    async def test_failed_lookup_does_not_fail_the_reply(self):
        handler = ClassicEmailHandler(_make_settings())
        with patch.object(handler, "_find_special_folder", AsyncMock(side_effect=RuntimeError("LIST failed"))):
            assert await handler.quote_source_mailboxes() == ("INBOX",)

    @pytest.mark.asyncio
    async def test_sent_equal_to_inbox_is_not_searched_twice(self):
        handler = ClassicEmailHandler(_make_settings())
        with patch.object(handler, "_find_special_folder", AsyncMock(return_value="INBOX")):
            assert await handler.quote_source_mailboxes() == ("INBOX",)


class TestQuoteSourceProviderAdapter:
    """The adapter carries the account's detected service and sender allowlist through."""

    @pytest.mark.asyncio
    async def test_detected_service_and_allowlist_reach_the_client(self):
        settings = _make_settings(imap_host="imap.gmail.com")
        handler = ClassicEmailHandler(settings)
        fetch = AsyncMock(return_value="<div>QUOTE</div>")
        command = SendCommand(
            account_name="test",
            recipients=("alice@example.com",),
            subject="Re: hello",
            body="my reply",
            in_reply_to="<original@example.com>",
        )

        with (
            patch.object(handler, "quote_source_mailboxes", AsyncMock(return_value=("INBOX", "Sent"))),
            patch.object(handler.incoming_client, "fetch_quote_source", fetch),
        ):
            source = await ClassicMutationProvider(handler).fetch_quote_source(
                command, _snapshot(allowed_senders=("*@example.com",))
            )

        assert source == QuoteSource(quote_html="<div>QUOTE</div>")
        assert fetch.await_args.args == (
            "<original@example.com>",
            ("INBOX", "Sent"),
            ["*@example.com"],
            "gmail",
        )

    @pytest.mark.asyncio
    async def test_absent_original_is_reported_as_no_quote(self):
        handler = ClassicEmailHandler(_make_settings())
        command = SendCommand(
            account_name="test",
            recipients=("alice@example.com",),
            subject="Re: hello",
            body="my reply",
            in_reply_to="<original@example.com>",
        )

        with (
            patch.object(handler, "quote_source_mailboxes", AsyncMock(return_value=("INBOX",))),
            patch.object(handler.incoming_client, "fetch_quote_source", AsyncMock(return_value=None)),
        ):
            assert await ClassicMutationProvider(handler).fetch_quote_source(command, _snapshot()) is None

    @pytest.mark.asyncio
    async def test_client_failure_is_not_swallowed(self):
        handler = ClassicEmailHandler(_make_settings())
        command = SendCommand(
            account_name="test",
            recipients=("alice@example.com",),
            subject="Re: hello",
            body="my reply",
            in_reply_to="<original@example.com>",
        )

        with (
            patch.object(handler, "quote_source_mailboxes", AsyncMock(return_value=("INBOX",))),
            patch.object(
                handler.incoming_client,
                "fetch_quote_source",
                AsyncMock(side_effect=RuntimeError("SELECT failed")),
            ),
        ):
            with pytest.raises(MutationProviderError):
                await ClassicMutationProvider(handler).fetch_quote_source(command, _snapshot())


def _plain_source(body: str = "Original message body") -> bytes:
    return (
        b"From: Alice <alice@example.com>\r\n"
        b"To: test@example.com\r\n"
        b"Subject: Original\r\n"
        b"Message-ID: <original@example.com>\r\n"
        b"Date: Fri, 15 Mar 2024 14:30:00 +0000\r\n"
        b"Content-Type: text/plain; charset=utf-8\r\n"
        b"\r\n" + body.encode("utf-8")
    )


def _html_source() -> bytes:
    return (
        b"From: Alice <alice@example.com>\r\n"
        b"To: test@example.com\r\n"
        b"Subject: Original\r\n"
        b"Message-ID: <original@example.com>\r\n"
        b"Date: Fri, 15 Mar 2024 14:30:00 +0000\r\n"
        b"Content-Type: text/html; charset=utf-8\r\n"
        b"\r\n<html><body><p>Rich <strong>HTML</strong> body</p></body></html>"
    )


@pytest.fixture
def email_client():
    return EmailClient(
        EmailServer(user_name="test", password="test", host="imap.example.com", port=993, use_ssl=True),
        sender="Test User <test@example.com>",
    )


class TestFetchQuoteSourcePrimitive:
    """Reading the original is single-session, allowlisted, and never degrades silently."""

    @staticmethod
    def _mock_imap(search_results: dict[str, bytes]):
        """A mock IMAP whose SEARCH answer depends on the mailbox last SELECTed."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock(return_value=MagicMock(result="OK", lines=[]))
        mock_imap.logout = AsyncMock()
        state = {"mailbox": ""}

        async def _select(quoted_mailbox):
            state["mailbox"] = quoted_mailbox.strip('"')
            return ("OK", [b"1"])

        async def _uid_search(*_args, **_kwargs):
            return ("OK", [search_results.get(state["mailbox"], b"")])

        mock_imap.select = AsyncMock(side_effect=_select)
        mock_imap.uid_search = AsyncMock(side_effect=_uid_search)
        return mock_imap

    async def _run(
        self,
        email_client,
        raw_email: bytes | None,
        search_results: dict[str, bytes],
        mailboxes=("INBOX", "Sent"),
        **kwargs,
    ):
        async def _fake_fetch(_imap, _email_id):
            if raw_email is None:
                return None
            return [b"1 FETCH (BODY[] {%d}" % len(raw_email), bytearray(raw_email), b")"]

        mock_imap = self._mock_imap(search_results)
        with (
            patch.object(email_client, "_fetch_email_with_formats", side_effect=_fake_fetch),
            patch.object(email_client, "imap_class", return_value=mock_imap),
        ):
            result = await email_client.fetch_quote_source("<original@example.com>", mailboxes, **kwargs)
        return result, mock_imap

    @pytest.mark.asyncio
    async def test_quotes_a_plain_text_original_found_in_inbox(self, email_client):
        quote, mock_imap = await self._run(email_client, _plain_source(), {"INBOX": b"7"})

        assert quote is not None
        assert "Original message body" in quote
        assert "alice@example.com" in quote
        assert "wrote:" in quote
        # Sent is never selected once INBOX answers.
        assert [call.args[0] for call in mock_imap.select.await_args_list] == ['"INBOX"']
        mock_imap.logout.assert_awaited_once()

    @pytest.mark.asyncio
    async def test_falls_through_to_the_sent_folder(self, email_client):
        quote, mock_imap = await self._run(email_client, _plain_source(), {"Sent": b"12"})

        assert quote is not None
        assert "Original message body" in quote
        assert [call.args[0] for call in mock_imap.select.await_args_list] == ['"INBOX"', '"Sent"']

    @pytest.mark.asyncio
    async def test_absent_original_returns_none(self, email_client):
        quote, mock_imap = await self._run(email_client, _plain_source(), {})

        assert quote is None
        mock_imap.logout.assert_awaited_once()

    @pytest.mark.asyncio
    async def test_html_original_is_quoted_as_markup(self, email_client):
        quote, _ = await self._run(email_client, _html_source(), {"INBOX": b"7"})

        assert quote is not None
        assert "<p>Rich <strong>HTML</strong> body</p>" in quote
        assert "<html>" not in quote
        assert "<body>" not in quote

    @pytest.mark.asyncio
    async def test_service_selects_the_quote_shape(self, email_client):
        quote, _ = await self._run(email_client, _plain_source(), {"INBOX": b"7"}, service="protonmail")

        assert quote is not None
        assert 'class="protonmail_quote"' in quote

    @pytest.mark.asyncio
    async def test_blocked_sender_is_indistinguishable_from_absent(self, email_client):
        with patch.object(email_client, "_batch_fetch_senders", AsyncMock(return_value={"7": "spam@evil.test"})):
            quote, _ = await self._run(
                email_client,
                _plain_source(),
                {"INBOX": b"7"},
                mailboxes=("INBOX",),
                allowed_senders=["*@example.com"],
            )

        assert quote is None

    @pytest.mark.asyncio
    async def test_allowed_sender_is_quoted(self, email_client):
        with patch.object(email_client, "_batch_fetch_senders", AsyncMock(return_value={"7": "alice@example.com"})):
            quote, _ = await self._run(
                email_client,
                _plain_source(),
                {"INBOX": b"7"},
                mailboxes=("INBOX",),
                allowed_senders=["*@example.com"],
            )

        assert quote is not None
        assert "Original message body" in quote

    @pytest.mark.asyncio
    async def test_unfetchable_found_message_raises(self, email_client):
        """Found-but-unreadable must never degrade to an unquoted reply."""
        with pytest.raises(ValueError, match="Failed to fetch quote source"):
            await self._run(email_client, None, {"INBOX": b"7"})

    @pytest.mark.asyncio
    async def test_oversized_original_raises(self, email_client):
        oversized = _plain_source("x" * (MAX_RAW_EMAIL_BYTES + 1))

        with pytest.raises(ValueError, match="raw message size limit"):
            await self._run(email_client, oversized, {"INBOX": b"7"})

    @pytest.mark.asyncio
    async def test_newest_duplicate_wins(self, email_client):
        quote, _ = await self._run(email_client, _plain_source(), {"INBOX": b"3 9 5"})

        assert quote is not None

    @pytest.mark.asyncio
    async def test_select_failure_propagates(self, email_client):
        mock_imap = self._mock_imap({})
        mock_imap.select = AsyncMock(return_value=("NO", [b"nope"]))
        with (
            patch.object(email_client, "imap_class", return_value=mock_imap),
            pytest.raises(RuntimeError),
        ):
            await email_client.fetch_quote_source("<original@example.com>", ("INBOX",))
        mock_imap.logout.assert_awaited_once()


class TestEmailServiceConfiguration:
    """email_service is persisted, environment-overridable, and defaults to auto-detect."""

    def test_defaults_to_auto_detection(self):
        assert _make_settings().email_service is None

    @staticmethod
    def _store_account(config_path, monkeypatch, account: EmailSettings) -> Settings:
        """Persist one account to an isolated TOML file and read it back."""
        monkeypatch.setitem(Settings.model_config, "toml_file", str(config_path))
        settings = Settings()
        settings.emails.append(account)
        settings.store()
        return Settings()

    def test_survives_a_toml_round_trip(self, tmp_path, monkeypatch):
        config_path = tmp_path / "config.toml"

        reloaded = self._store_account(config_path, monkeypatch, _make_settings(email_service="protonmail"))

        assert config_path.read_text().count('email_service = "protonmail"') == 1
        assert [account.email_service for account in reloaded.emails] == ["protonmail"]

    def test_absent_field_round_trips_as_auto_detect(self, tmp_path, monkeypatch):
        config_path = tmp_path / "config.toml"

        reloaded = self._store_account(config_path, monkeypatch, _make_settings())

        assert "email_service" not in config_path.read_text()
        assert [account.email_service for account in reloaded.emails] == [None]

    def test_environment_override(self, monkeypatch):
        monkeypatch.setenv("MCP_EMAIL_SERVER_EMAIL_ADDRESS", "test@example.com")
        monkeypatch.setenv("MCP_EMAIL_SERVER_PASSWORD", "pass")
        monkeypatch.setenv("MCP_EMAIL_SERVER_IMAP_HOST", "imap.example.com")
        monkeypatch.setenv("MCP_EMAIL_SERVER_EMAIL_SERVICE", "gmail")

        from_env = EmailSettings.from_env()

        assert from_env is not None
        assert from_env.email_service == "gmail"
        assert _detect_email_service(from_env) == "gmail"

    def test_absent_environment_override_leaves_detection_alone(self, monkeypatch):
        monkeypatch.setenv("MCP_EMAIL_SERVER_EMAIL_ADDRESS", "test@example.com")
        monkeypatch.setenv("MCP_EMAIL_SERVER_PASSWORD", "pass")
        monkeypatch.setenv("MCP_EMAIL_SERVER_IMAP_HOST", "imap.gmail.com")
        monkeypatch.delenv("MCP_EMAIL_SERVER_EMAIL_SERVICE", raising=False)

        from_env = EmailSettings.from_env()

        assert from_env is not None
        assert from_env.email_service is None
        assert _detect_email_service(from_env) == "gmail"
