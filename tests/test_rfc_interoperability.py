"""RFC 5322 / RFC 3501 interoperability regressions.

Ported from upstream ai-zerolab/mcp-email-server PR #230 (commit ``a1668a9``),
adapted to this fork's pre-V2 ``classic.py``. Covers only the subset that is
reachable here: MIME charset robustness, ``message/rfc822`` isolation,
structural recipient parsing, IMAP LIST framing, the APPEND keyword atom
grammar, locale-independent SEARCH dates, RFC 6154 special-use case folding,
and the ASCII half of IMAP SEARCH astring encoding.
"""

from __future__ import annotations

import asyncio
import locale
from datetime import datetime, timezone
from email.message import EmailMessage
from email.policy import SMTP
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from mcp_email_server.emails.classic import (
    ClassicEmailHandler,
    EmailClient,
    _parse_list_responses,
    _validate_flags,
)
from mcp_email_server.emails.models import MailboxInfo


@pytest.fixture
def rfc_email_client(email_server):
    return EmailClient(email_server, sender="Test User <test@example.com>")


@pytest.fixture
def rfc_handler(email_settings):
    return ClassicEmailHandler(email_settings)


def _make_mock_imap(list_lines):
    """Build an AsyncMock IMAP client shaped like the fork's ``_imap_connection``."""
    imap = AsyncMock()
    imap._client_task = asyncio.Future()
    imap._client_task.set_result(None)
    imap.wait_hello_from_server = AsyncMock()
    imap.login = AsyncMock(return_value=MagicMock(result="OK", lines=[]))
    imap.select = AsyncMock(return_value=("OK", []))
    imap.logout = AsyncMock()
    imap.list = AsyncMock(return_value=("OK", list_lines))
    return imap


# ---------------------------------------------------------------------------
# Item 3: structural RFC 5322 recipient parsing
# ---------------------------------------------------------------------------


def test_recipient_headers_preserve_quoted_names_and_flatten_groups(rfc_email_client):
    """A comma inside a quoted display name is not an address separator."""
    raw_email = (
        b"From: sender@example.test\r\n"
        b'To: "Doe, John" <john@example.test>, Team: a@example.test, b@example.test;\r\n'
        b"Cc: cc@example.test\r\n"
        b"Subject: Address list\r\n"
        b"Date: Sat, 8 Aug 2026 00:00:00 +0000\r\n"
        b"\r\n"
        b"body"
    )

    parsed = rfc_email_client._parse_email_data(raw_email)
    metadata = rfc_email_client._parse_headers("1", raw_email.split(b"\r\n\r\n", 1)[0] + b"\r\n\r\n")

    expected = [
        '"Doe, John" <john@example.test>',
        "a@example.test",
        "b@example.test",
        "cc@example.test",
    ]
    assert parsed["to"] == expected
    assert metadata is not None
    assert metadata["to"] == expected


def test_multiple_to_header_instances_are_all_parsed(rfc_email_client):
    """``get_all`` covers repeated address fields, not just the first one."""
    raw_email = (
        b"From: sender@example.test\r\n"
        b"To: first@example.test\r\n"
        b"To: second@example.test\r\n"
        b"Subject: Repeated field\r\n"
        b"\r\n"
        b"body"
    )

    assert rfc_email_client._parse_email_data(raw_email)["to"] == [
        "first@example.test",
        "second@example.test",
    ]


# ---------------------------------------------------------------------------
# Item 1: unknown MIME charsets must not swallow the message
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("charset", ["x-no-such-codec", "unknown-8bit", "x-unknown", "iso-8859-8-i"])
def test_unknown_single_part_charset_falls_back_without_losing_message(rfc_email_client, charset):
    """``bytes.decode`` raises LookupError — not UnicodeDecodeError — here."""
    raw_email = (
        b"From: sender@example.test\r\n"
        b"To: recipient@example.test\r\n"
        b"Subject: Unknown charset\r\n"
        b"Content-Type: text/plain; charset=" + charset.encode("ascii") + b"\r\n"
        b"Content-Transfer-Encoding: 8bit\r\n"
        b"\r\n"
        b"hello \xff"
    )

    parsed = rfc_email_client._parse_email_data(raw_email)

    assert parsed["subject"] == "Unknown charset"
    assert parsed["body"] == "hello \ufffd"
    assert parsed["to"] == ["recipient@example.test"]


def test_unknown_multipart_charset_does_not_hide_valid_parts(rfc_email_client):
    message = EmailMessage()
    message["From"] = "sender@example.test"
    message["To"] = "recipient@example.test"
    message["Subject"] = "Mixed charsets"
    message.make_mixed()

    unknown = EmailMessage()
    unknown.set_type("text/plain")
    unknown.set_param("charset", "x-no-such-codec")
    unknown["Content-Transfer-Encoding"] = "8bit"
    unknown.set_payload(b"unknown \xff")
    message.attach(unknown)

    valid = EmailMessage()
    valid.set_content(" valid", charset="utf-8")
    message.attach(valid)

    parsed = rfc_email_client._parse_email_data(message.as_bytes(policy=SMTP))

    assert parsed["body"] == "unknown \ufffd valid\r\n"


def test_unknown_html_charset_still_reaches_the_text_fallback(rfc_email_client):
    """An HTML-only body with a bad charset must not vanish either."""
    raw_email = (
        b"From: sender@example.test\r\n"
        b"To: recipient@example.test\r\n"
        b"Subject: HTML unknown charset\r\n"
        b"Content-Type: text/html; charset=x-no-such-codec\r\n"
        b"Content-Transfer-Encoding: 8bit\r\n"
        b"\r\n"
        b"<p>hello \xff</p>"
    )

    parsed = rfc_email_client._parse_email_data(raw_email)

    assert "hello" in parsed["body"]
    assert parsed["html_body"] == "<p>hello \ufffd</p>"


# ---------------------------------------------------------------------------
# Item 2: message/rfc822 subtree isolation
# ---------------------------------------------------------------------------


def _attached_message(*, filename: str | None) -> bytes:
    disposition = b""
    parameters = b""
    if filename is not None:
        parameters = f'; name="{filename}"'.encode()
        disposition = f'Content-Disposition: attachment; filename="{filename}"\r\n'.encode()
    return (
        b"From: sender@example.test\r\n"
        b"To: recipient@example.test\r\n"
        b"Subject: Outer\r\n"
        b'MIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary="outer"\r\n'
        b"\r\n"
        b"--outer\r\nContent-Type: text/plain; charset=utf-8\r\n\r\nOUTER\r\n"
        b"--outer\r\nContent-Type: message/rfc822" + parameters + b"\r\n" + disposition + b"\r\n"
        b"From: nested@example.test\r\n"
        b"To: nested-recipient@example.test\r\n"
        b"Subject: Nested\r\n"
        b"Content-Type: text/plain; charset=utf-8\r\n"
        b"\r\n"
        b"INNER\r\n"
        b"--outer--\r\n"
    )


@pytest.mark.parametrize(
    ("filename", "expected_attachments"),
    [("forwarded.eml", ["forwarded.eml"]), (None, [])],
)
def test_attached_message_subtree_is_not_promoted_to_outer_body(
    rfc_email_client,
    filename,
    expected_attachments,
):
    """An encapsulated message is isolated even when it carries no filename."""
    parsed = rfc_email_client._parse_email_data(_attached_message(filename=filename))

    assert parsed["body"] == "OUTER"
    assert "INNER" not in parsed["body"]
    assert parsed["attachments"] == expected_attachments


def test_single_part_non_text_body_is_still_returned(rfc_email_client):
    """The fork keeps a fallback for a lone non-text body (e.g. text/calendar)."""
    raw_email = (
        b"From: sender@example.test\r\n"
        b"To: recipient@example.test\r\n"
        b"Subject: Invite\r\n"
        b"Content-Type: text/calendar; charset=utf-8\r\n"
        b"\r\n"
        b"BEGIN:VCALENDAR\r\nEND:VCALENDAR"
    )

    assert rfc_email_client._parse_email_data(raw_email)["body"] == "BEGIN:VCALENDAR\r\nEND:VCALENDAR"


def test_html_body_is_collected_even_when_plain_text_wins(rfc_email_client):
    """``quote_reply``/``forward_email`` consume ``html_body`` unconditionally."""
    message = EmailMessage()
    message["From"] = "sender@example.test"
    message["To"] = "recipient@example.test"
    message["Subject"] = "Alternative"
    message.set_content("plain text")
    message.add_alternative("<p>rich text</p>", subtype="html")

    parsed = rfc_email_client._parse_email_data(message.as_bytes(policy=SMTP))

    assert "plain text" in parsed["body"]
    assert "rich text" in parsed["html_body"]


# ---------------------------------------------------------------------------
# Item 5: APPEND keyword atom grammar
# ---------------------------------------------------------------------------


def test_imap_keyword_validation_accepts_complete_atom_grammar():
    assert _validate_flags(["$Forwarded", "$Junk", "project.name", "123flag"]) == (
        "($Forwarded $Junk project.name 123flag)"
    )
    assert _validate_flags(["$NotJunk", "$Phishing", "$MDNSent"]) == "($NotJunk $Phishing $MDNSent)"
    assert _validate_flags([r"\Seen", r"\Draft"]) == r"(\Seen \Draft)"


@pytest.mark.parametrize(
    "invalid",
    ["", "two words", "a b", "bad(flag", "bad{flag", "bad%flag", "bad*flag", 'bad"flag', "\\Seen)", "\\\\Seen"],
)
def test_imap_keyword_validation_rejects_atom_specials(invalid):
    with pytest.raises(ValueError, match="Invalid IMAP flag"):
        _validate_flags([invalid])


# ---------------------------------------------------------------------------
# Item 6: locale-independent SEARCH dates
# ---------------------------------------------------------------------------


def test_imap_dates_use_fixed_english_month_names():
    expected_months = ("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")

    for month, expected in enumerate(expected_months, start=1):
        value = datetime(2026, month, 5, tzinfo=timezone.utc)
        assert EmailClient._build_search_criteria(since=value) == ["SINCE", f"05-{expected}-2026"]
        assert EmailClient._build_search_criteria(before=value) == ["BEFORE", f"05-{expected}-2026"]


def test_imap_dates_ignore_the_process_locale():
    """``strftime('%b')`` would emit ``MÄRZ`` under de_DE and break every search."""
    original = locale.setlocale(locale.LC_TIME)
    try:
        try:
            locale.setlocale(locale.LC_TIME, "de_DE.UTF-8")
        except locale.Error:
            pytest.skip("de_DE.UTF-8 locale unavailable")
        criteria = EmailClient._build_search_criteria(since=datetime(2026, 3, 5, tzinfo=timezone.utc))
    finally:
        locale.setlocale(locale.LC_TIME, original)

    assert criteria == ["SINCE", "05-MAR-2026"]


# ---------------------------------------------------------------------------
# Item 4: IMAP LIST framing
# ---------------------------------------------------------------------------


def test_list_response_framing_ignores_completion_and_reassembles_literal():
    parsed = _parse_list_responses([
        b'(\\HasNoChildren) "/" "INBOX"',
        b'(\\sent) "/" {4}',
        bytearray(b"Sent"),
        b"List completed (0.001 + 0.000 secs).",
    ])

    assert [(mailbox.name, mailbox.delimiter, mailbox.flags) for mailbox in parsed] == [
        ("INBOX", "/", [r"\HasNoChildren"]),
        ("Sent", "/", [r"\sent"]),
    ]


@pytest.mark.parametrize(
    "lines",
    [
        [b'(\\Sent) "/" {4}'],
        [b'(\\Sent) "/" {5}', bytearray(b"Sent")],
    ],
)
def test_list_response_framing_rejects_invalid_literals(lines):
    with pytest.raises(ValueError, match="LIST literal"):
        _parse_list_responses(lines)


def test_list_literal_round_trips_spaces_quotes_and_backslashes():
    mailbox_name = b'A "quoted" \\ mailbox'
    parsed = _parse_list_responses([
        b'() "/" {' + str(len(mailbox_name)).encode("ascii") + b"}",
        bytearray(mailbox_name),
        b"LIST completed",
    ])

    assert len(parsed) == 1
    assert parsed[0].name == mailbox_name.decode("ascii")


@pytest.mark.asyncio
async def test_list_mailboxes_drops_completion_line_and_keeps_literal_name(rfc_email_client):
    """End-to-end: the tagged completion text is not a mailbox."""
    mock_imap = _make_mock_imap([
        b'(\\HasNoChildren) "/" "INBOX"',
        b'(\\Archive) "/" {7}',
        bytearray(b"Archive"),
        b"List completed (0.001 + 0.000 secs).",
    ])

    with patch.object(rfc_email_client, "imap_class", return_value=mock_imap):
        mailboxes = await rfc_email_client.list_mailboxes()

    assert [mailbox.name for mailbox in mailboxes] == ["INBOX", "Archive"]


@pytest.mark.asyncio
async def test_archive_folder_lookup_uses_the_literal_mailbox_name(rfc_handler):
    """A literal-form name must not reach COPY as ``{7}``."""
    mock_imap = _make_mock_imap([
        b'(\\HasNoChildren) "/" "INBOX"',
        b'(\\Archive) "/" {7}',
        bytearray(b"Archive"),
        b"List completed (0.001 + 0.000 secs).",
    ])
    mock_move = AsyncMock(return_value=(["123"], []))

    with (
        patch.object(rfc_handler.incoming_client, "imap_class", return_value=mock_imap),
        patch.object(rfc_handler.incoming_client, "move_emails", mock_move),
    ):
        result = await rfc_handler.archive_emails(["123"], "INBOX")

    assert result.destination_folder == "Archive"
    mock_move.assert_awaited_once_with(["123"], "INBOX", "Archive")


# ---------------------------------------------------------------------------
# Item 7: RFC 6154 special-use flag case
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_sent_folder_special_use_matching_is_case_insensitive(rfc_email_client):
    imap = AsyncMock()
    imap.list.return_value = (
        "OK",
        [b'(\\sent \\HasNoChildren) "/" "Gesendet"', b"List completed (0.001 + 0.000 secs)."],
    )

    assert await rfc_email_client._find_sent_folder_by_flag(imap) == "Gesendet"


@pytest.mark.asyncio
async def test_special_folder_matching_is_case_insensitive_but_exact(rfc_handler):
    """``\\archive`` matches; an unrelated attribute containing it does not."""
    folders = [
        MailboxInfo(name="Decoy", delimiter="/", flags=["\\NoArchive"]),
        MailboxInfo(name="Archiv", delimiter="/", flags=["\\archive"]),
    ]

    with patch.object(rfc_handler.incoming_client, "list_mailboxes", AsyncMock(return_value=folders)):
        assert await rfc_handler._find_special_folder("\\Archive", []) == "Archiv"


# ---------------------------------------------------------------------------
# Item 8: IMAP SEARCH astring encoding (ASCII half)
# ---------------------------------------------------------------------------


def test_search_astring_encoding_preserves_user_text(rfc_email_client):
    assert rfc_email_client._sanitize_imap_value("simple") == "simple"
    assert rfc_email_client._sanitize_imap_value("a]b") == "a]b"
    assert rfc_email_client._sanitize_imap_value("two words") == '"two words"'
    assert rfc_email_client._sanitize_imap_value("foo(bar)") == '"foo(bar)"'
    assert rfc_email_client._sanitize_imap_value('say"hi"') == '"say\\"hi\\""'
    assert rfc_email_client._sanitize_imap_value(r"C:\Path") == '"C:\\\\Path"'
    assert rfc_email_client._sanitize_imap_value("100%") == '"100%"'
    assert rfc_email_client._sanitize_imap_value("wild*card") == '"wild*card"'
    assert rfc_email_client._sanitize_imap_value("brace{1}") == '"brace{1}"'


def test_search_astring_keeps_user_quotes_instead_of_stripping_them(rfc_email_client):
    """The old sanitizer silently dropped the user's own quotes."""
    assert rfc_email_client._sanitize_imap_value('say "hi"') == '"say \\"hi\\""'


def test_search_non_ascii_stays_a_quoted_string(rfc_email_client):
    """This fork declares CHARSET utf-8 rather than sending a literal."""
    value = rfc_email_client._sanitize_imap_value("会议")
    assert isinstance(value, str)
    assert value == '"会议"'
    assert EmailClient._build_search_criteria(subject="会议") == ["SUBJECT", '"会议"']


def test_search_rejects_control_characters(rfc_email_client):
    with pytest.raises(ValueError, match="control characters"):
        rfc_email_client._sanitize_imap_value("bad\r\nvalue")
