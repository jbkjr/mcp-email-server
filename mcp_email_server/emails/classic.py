import asyncio
import base64
import binascii
import email.utils
import mimetypes
import re
import ssl
import time
import unicodedata
from collections.abc import AsyncGenerator, Callable
from contextlib import asynccontextmanager
from datetime import datetime, timezone
from email.header import Header
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.parser import BytesParser
from email.policy import default
from pathlib import Path
from typing import Any

import aioimaplib
import aiosmtplib

from mcp_email_server.config import EmailServer, EmailSettings, get_settings, sender_allowed
from mcp_email_server.emails import EmailHandler
from mcp_email_server.emails.html_utils import html_to_text
from mcp_email_server.emails.markdown_utils import markdown_to_email_html, wrap_html_document
from mcp_email_server.emails.models import (
    AttachmentDownloadResponse,
    EmailBodyResponse,
    EmailContentBatchResponse,
    EmailDeleteResponse,
    EmailLabelsResponse,
    EmailMarkResponse,
    EmailMetadata,
    EmailMetadataPageResponse,
    EmailMoveResponse,
    EmailSendResponse,
    FolderOperationResponse,
    Label,
    LabelListResponse,
    MailboxInfo,
)
from mcp_email_server.log import logger

# Maximum body length before truncation (characters)
MAX_BODY_LENGTH = 20000

# Maximum quoted body length before truncation (characters)
MAX_QUOTED_BODY_LENGTH = 5000

# ProtonMail Bridge labels prefix
LABELS_PREFIX = "Labels/"

# Default IMAP mailbox
DEFAULT_MAILBOX = "INBOX"

# IMAP flag strings
IMAP_FLAG_SEEN = r"(\Seen)"
IMAP_FLAG_DELETED = r"(\Deleted)"

# Common Sent folder names for auto-detection
SENT_FOLDER_CANDIDATES = [
    "Sent",
    "INBOX.Sent",
    "Sent Items",
    "Sent Mail",
    "[Gmail]/Sent Mail",
    "INBOX/Sent",
]

# Common Trash folder names for auto-detection
TRASH_FOLDER_CANDIDATES = [
    "Trash",
    "Deleted Items",
    "Deleted Messages",
    "[Gmail]/Trash",
    "INBOX.Trash",
]

# Common Archive folder names for auto-detection
ARCHIVE_FOLDER_CANDIDATES = [
    "Archive",
    "[Gmail]/All Mail",
    "INBOX.Archive",
]


def _detect_email_service(email_settings: EmailSettings) -> str:
    """Detect the email service from IMAP config for service-aware reply quoting.

    Detection order:
    1. Explicit override via email_service config field
    2. imap.gmail.com host → 'gmail' (includes Google Workspace custom domains)
    3. localhost/127.0.0.1 + verify_ssl=False → 'protonmail' (ProtonMail Bridge)
    4. Everything else → 'generic'

    Note: The localhost heuristic could mislabel other local IMAP servers with
    self-signed certs as ProtonMail. Use email_service config override for those cases.
    """
    if email_settings.email_service:
        return email_settings.email_service

    host = email_settings.incoming.host.lower()

    if host == "imap.gmail.com":
        return "gmail"

    if host in ("localhost", "127.0.0.1") and not email_settings.incoming.verify_ssl:
        return "protonmail"

    return "generic"


def _is_ok(result: Any) -> bool:
    """Check if an IMAP operation result indicates success."""
    status = result[0] if isinstance(result, tuple) else result
    return str(status).upper() == "OK"


# RFC 3501 system flags (except \Recent which is read-only) + custom keyword atoms
_VALID_IMAP_FLAG = re.compile(r"^\\[A-Za-z]+$|^[A-Za-z][A-Za-z0-9_-]*$")


def _validate_flags(flags: list[str]) -> str:
    """Validate and format IMAP flags into a parenthesised string.

    Accepts system flags (e.g. ``\\Draft``, ``\\Seen``) and custom keyword
    atoms.  Raises ``ValueError`` on anything that could inject IMAP protocol
    characters.
    """
    for flag in flags:
        if not _VALID_IMAP_FLAG.match(flag):
            msg = f"Invalid IMAP flag: {flag!r}"
            raise ValueError(msg)
    return "(" + " ".join(flags) + ")"


def encode_mailbox_name(mailbox: str) -> str:
    """Encode an IMAP mailbox name using RFC 3501 Modified UTF-7."""
    result: list[str] = []
    buffer: list[str] = []

    def flush_buffer() -> None:
        if not buffer:
            return
        text = "".join(buffer)
        encoded = base64.b64encode(text.encode("utf-16-be")).decode("ascii")
        result.append("&" + encoded.rstrip("=").replace("/", ",") + "-")
        buffer.clear()

    for char in mailbox:
        codepoint = ord(char)
        if char == "&":
            flush_buffer()
            result.append("&-")
        elif 0x20 <= codepoint <= 0x7E:
            flush_buffer()
            result.append(char)
        else:
            buffer.append(char)

    flush_buffer()
    return "".join(result)


def decode_mailbox_name(mailbox: str) -> str:
    """Decode an IMAP mailbox name from RFC 3501 Modified UTF-7."""
    result: list[str] = []
    index = 0

    while index < len(mailbox):
        char = mailbox[index]
        if char != "&":
            result.append(char)
            index += 1
            continue

        end = mailbox.find("-", index + 1)
        if end == -1:
            result.append(mailbox[index:])
            break
        if end == index + 1:
            result.append("&")
            index = end + 1
            continue

        encoded = mailbox[index + 1 : end].replace(",", "/")
        padding = "=" * (-len(encoded) % 4)
        try:
            decoded = base64.b64decode(encoded + padding, validate=True).decode("utf-16-be")
        except (binascii.Error, UnicodeDecodeError):
            result.append(mailbox[index : end + 1])
        else:
            result.append(decoded)
        index = end + 1

    return "".join(result)


def _skip_imap_whitespace(value: str, start: int) -> int:
    """Return the next non-whitespace index in an IMAP response line."""
    index = start
    while index < len(value) and value[index].isspace():
        index += 1
    return index


def _read_quoted_imap_token(value: str, start: int) -> tuple[str, int]:
    """Read a quoted IMAP token."""
    index = start + 1
    token: list[str] = []
    while index < len(value):
        char = value[index]
        if char == "\\" and index + 1 < len(value):
            token.append(value[index + 1])
            index += 2
            continue
        if char == '"':
            return "".join(token), index + 1
        token.append(char)
        index += 1
    return "".join(token), index


def _read_parenthesized_imap_token(value: str, start: int) -> tuple[str, int]:
    """Read a parenthesized IMAP token."""
    depth = 1
    index = start + 1
    token: list[str] = []
    while index < len(value):
        char = value[index]
        if char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
            if depth == 0:
                return "".join(token), index + 1
        token.append(char)
        index += 1
    return "".join(token), index


def _read_atom_imap_token(value: str, start: int) -> tuple[str, int]:
    """Read an atom IMAP token."""
    index = start
    while index < len(value) and not value[index].isspace():
        index += 1
    return value[start:index], index


def _read_imap_list_token(value: str, start: int) -> tuple[str | None, int]:
    """Read one token from an IMAP LIST response line."""
    index = _skip_imap_whitespace(value, start)
    if index >= len(value):
        return None, index
    if value[index] == '"':
        return _read_quoted_imap_token(value, index)
    if value[index] == "(":
        return _read_parenthesized_imap_token(value, index)
    return _read_atom_imap_token(value, index)


def _parse_list_response(item: bytes | str) -> MailboxInfo | None:
    """Parse one IMAP LIST response into a MailboxInfo object."""
    item_str = item.decode("utf-8", errors="replace") if isinstance(item, bytes) else str(item)
    item_str = item_str.strip()
    if not item_str:
        return None

    flags: list[str] = []
    position = 0
    if item_str.startswith("("):
        flags_token, position = _read_imap_list_token(item_str, position)
        flags = [flag.strip() for flag in (flags_token or "").split() if flag.strip()]

    delimiter_token, position = _read_imap_list_token(item_str, position)
    mailbox_token, _position = _read_imap_list_token(item_str, position)
    if delimiter_token is None or mailbox_token is None:
        return None

    delimiter = "" if delimiter_token.upper() == "NIL" else delimiter_token
    return MailboxInfo(name=decode_mailbox_name(mailbox_token), delimiter=delimiter, flags=flags)


def _quote_mailbox(mailbox: str) -> str:
    """Quote mailbox name for IMAP compatibility.

    Some IMAP servers (notably Proton Mail Bridge) require mailbox names
    to be quoted. This is valid per RFC 3501 and works with all IMAP servers.

    Per RFC 3501 Section 9 (Formal Syntax), quoted strings must escape
    backslashes and double-quote characters with a preceding backslash.
    Mailbox names with non-ASCII characters are encoded using Modified UTF-7
    as required by RFC 3501 Section 5.1.3.

    See: https://github.com/ai-zerolab/mcp-email-server/issues/87
    See: https://github.com/ai-zerolab/mcp-email-server/issues/172
    See: https://www.rfc-editor.org/rfc/rfc3501#section-9
    """
    encoded = encode_mailbox_name(mailbox)
    # Per RFC 3501, literal double-quote characters in a quoted string must
    # be escaped with a backslash. Backslashes themselves must also be escaped.
    escaped = encoded.replace("\\", "\\\\").replace('"', r"\"")
    return f'"{escaped}"'


def _uid_sort_key(uid: bytes | str) -> int:
    """Return a numeric sort key for IMAP UIDs."""
    value = uid.decode() if isinstance(uid, bytes) else uid
    return int(value)


def _imap_status(response: Any) -> str:
    """Return the normalized status from an aioimaplib response."""
    if hasattr(response, "result"):
        return str(response.result).upper()
    if isinstance(response, tuple) and response:
        return str(response[0]).upper()
    return str(response).upper()


def _format_imap_response_detail(response: Any) -> str:
    """Return a compact, readable IMAP response detail string."""
    status = _imap_status(response)
    lines = getattr(response, "lines", None)
    if lines is None and isinstance(response, tuple) and len(response) > 1:
        lines = response[1]

    detail_parts = []
    for line in lines or []:
        if isinstance(line, bytes):
            detail_parts.append(line.decode("utf-8", errors="replace"))
        else:
            detail_parts.append(str(line))

    detail = " ".join(part for part in detail_parts if part).strip()
    return f"{status} {detail}".strip()


def _raise_for_imap_error(response: Any, operation: str) -> None:
    """Raise when an IMAP command returns a non-OK status."""
    if _imap_status(response) != "OK":
        detail = _format_imap_response_detail(response)
        msg = f"{operation} failed" + (f": {detail}" if detail else "")
        raise RuntimeError(msg)


def _raise_for_imap_command_failure(response: Any, operation: str) -> None:
    """Raise only when an IMAP command reports an explicit failure status."""
    if _imap_status(response) in {"NO", "BAD"}:
        detail = _format_imap_response_detail(response)
        msg = f"{operation} failed" + (f": {detail}" if detail else "")
        raise RuntimeError(msg)


async def _send_imap_id(imap: aioimaplib.IMAP4 | aioimaplib.IMAP4_SSL) -> None:
    """Send IMAP ID command with fallback for strict servers like 163.com.

    aioimaplib's id() method sends ID command with spaces between parentheses
    and content (e.g., 'ID ( "name" "value" )'), which some strict IMAP servers
    like 163.com reject with 'BAD Parse command error'.

    This function first tries the standard id() method, and if it fails,
    falls back to sending a raw command with correct format.

    See: https://github.com/ai-zerolab/mcp-email-server/issues/85
    """
    try:
        response = await imap.id(name="mcp-email-server", version="1.0.0")
        if response.result != "OK":
            # Fallback for strict servers (e.g., 163.com)
            # Send raw command with correct parenthesis format
            new_tag = imap.protocol.new_tag()
            if hasattr(new_tag, "__await__"):
                new_tag = await new_tag
            await imap.protocol.execute(
                aioimaplib.Command(
                    "ID",
                    new_tag,
                    '("name" "mcp-email-server" "version" "1.0.0")',
                )
            )
    except Exception as e:
        logger.warning(f"IMAP ID command failed: {e!s}")


async def _imap_login(
    imap: aioimaplib.IMAP4 | aioimaplib.IMAP4_SSL,
    user_name: str,
    password: str,
) -> None:
    """Authenticate to IMAP and fail loudly when the server rejects credentials.

    aioimaplib's ``login()`` returns a Response with a ``.result`` of "OK",
    "NO", or "BAD". A "NO" response (e.g. wrong credentials, account locked,
    or a transient rate-limit cool-down on servers like Proton Mail Bridge)
    does NOT raise — and an unchecked caller will happily proceed to issue
    SELECT/FETCH on a NONAUTH connection, producing the misleading error
    ``command SELECT illegal in state NONAUTH``. Worse, each tool call then
    opens a fresh TCP connection and re-attempts ``LOGIN``, which amplifies
    rate-limits on servers that count failed-login attempts and locks the
    account out for tens of minutes.

    Raise immediately on a non-OK result so callers (and end users) see the
    real error and back off, and so a one-off auth failure does not cascade
    into a multi-minute lock-out.
    """
    response = await imap.login(user_name, password)
    if response.result == "OK":
        return
    detail = " ".join(
        line.decode("utf-8", errors="replace") if isinstance(line, bytes) else str(line)
        for line in (response.lines or [])
    ).strip()
    raise ConnectionError(
        f"IMAP login failed for {user_name!r}: {response.result}" + (f" ({detail})" if detail else "")
    )


def _create_ssl_context(verify_ssl: bool) -> ssl.SSLContext | None:
    """Create SSL context for SMTP/IMAP connections.

    Returns None for default verification, or permissive context
    for self-signed certificates when verify_ssl=False.
    """
    if verify_ssl:
        return None
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx


def _strip_html_wrappers(html_content: str) -> str:
    """Strip HTML document wrappers (DOCTYPE, html, head, body tags), keeping body content."""
    content = re.sub(r"<!DOCTYPE[^>]*>", "", html_content, flags=re.IGNORECASE)
    content = re.sub(r"<html[^>]*>", "", content, flags=re.IGNORECASE)
    content = re.sub(r"</html\s*>", "", content, flags=re.IGNORECASE)
    content = re.sub(r"<head[^>]*>.*?</head\s*>", "", content, flags=re.IGNORECASE | re.DOTALL)
    content = re.sub(r"<body[^>]*>", "", content, flags=re.IGNORECASE)
    content = re.sub(r"</body\s*>", "", content, flags=re.IGNORECASE)
    return content.strip()


def _format_quoted_reply_html(original_email: dict[str, Any], service: str = "generic") -> str:
    """Format an original email as an HTML blockquote for reply quoting.

    Uses service-specific HTML structure so email clients can properly collapse
    the quoted content (including the attribution line).

    Args:
        original_email: Parsed email dict with keys: from, date, body, html_body.
        service: Email service name ('protonmail', 'gmail', 'generic').

    Returns:
        HTML string containing an attribution line and styled blockquote.
    """
    from html import escape as html_escape

    sender = original_email.get("from", "Unknown")
    date = original_email.get("date")
    raw_html_body = original_email.get("html_body", "")
    text_body = original_email.get("body", "")

    # Format date
    if isinstance(date, datetime):
        date_str = date.strftime("%a, %b %d, %Y at %I:%M %p")
    else:
        date_str = str(date) if date else "Unknown date"

    # Build quoted content
    if raw_html_body:
        quoted_content = _strip_html_wrappers(raw_html_body)
    elif text_body:
        if len(text_body) > MAX_QUOTED_BODY_LENGTH:
            text_body = text_body[:MAX_QUOTED_BODY_LENGTH] + "\n[...quoted text truncated]"
        quoted_content = "<br>\n".join(html_escape(line) for line in text_body.splitlines())
    else:
        quoted_content = ""

    attribution = f"On {date_str}, {html_escape(sender)} wrote:"

    if service == "protonmail":
        return (
            f'<div class="protonmail_quote">'
            f"{attribution}<br>"
            f'<blockquote class="protonmail_quote" type="cite">'
            f"{quoted_content}"
            f"</blockquote><br>"
            f"</div>"
        )

    if service == "gmail":
        return (
            f'<div class="gmail_quote">'
            f'<div class="gmail_attr" dir="ltr">{attribution}<br></div>'
            f'<blockquote class="gmail_quote" style="'
            f"margin:0 0 0 .8ex;"
            f"border-inline-start:1px solid rgb(204,204,204);"
            f'padding-inline-start:1ex">'
            f"{quoted_content}"
            f"</blockquote>"
            f"</div>"
        )

    # generic: attribution inside blockquote for maximum compatibility
    return (
        f'<div style="margin-top: 1em;">'
        f'<blockquote type="cite" style="'
        f"margin: 0 0 0 0.5em; "
        f"padding: 0.5em 1em; "
        f"border-left: 3px solid #ccc; "
        f'color: #555;">'
        f'<p style="color: #666;">{attribution}</p>'
        f"{quoted_content}"
        f"</blockquote>"
        f"</div>"
    )


def _format_forwarded_email_html(original_email: dict[str, Any]) -> str:
    """Format an original email as an HTML forwarded message block.

    Uses Gmail-style forwarded message format with inline headers.

    Args:
        original_email: Parsed email dict with keys: from, date, subject, to, body, html_body.

    Returns:
        HTML string containing a forwarded message header and body.
    """
    from html import escape as html_escape

    sender = original_email.get("from", "Unknown")
    date = original_email.get("date")
    subject = original_email.get("subject", "")
    to_addresses = original_email.get("to", [])
    raw_html_body = original_email.get("html_body", "")
    text_body = original_email.get("body", "")

    # Format date
    if isinstance(date, datetime):
        date_str = date.strftime("%a, %b %d, %Y at %I:%M %p")
    else:
        date_str = str(date) if date else "Unknown date"

    # Format To field
    if isinstance(to_addresses, list):
        to_str = html_escape(", ".join(to_addresses))
    else:
        to_str = html_escape(str(to_addresses)) if to_addresses else ""

    # Build forwarded content
    if raw_html_body:
        forwarded_content = _strip_html_wrappers(raw_html_body)
    elif text_body:
        if len(text_body) > MAX_QUOTED_BODY_LENGTH:
            text_body = text_body[:MAX_QUOTED_BODY_LENGTH] + "\n[...forwarded text truncated]"
        forwarded_content = "<br>\n".join(html_escape(line) for line in text_body.splitlines())
    else:
        forwarded_content = ""

    return (
        f'<div style="margin-top: 1em;">'
        f'<p style="color: #666;">---------- Forwarded message ---------</p>'
        f'<p style="color: #666;">'
        f"From: {html_escape(sender)}<br>"
        f"Date: {html_escape(date_str)}<br>"
        f"Subject: {html_escape(subject)}<br>"
        f"To: {to_str}"
        f"</p>"
        f"<div>{forwarded_content}</div>"
        f"</div>"
    )


def _create_starttls_ssl_context(verify_ssl: bool) -> ssl.SSLContext:
    """Create a concrete SSL context for asyncio STARTTLS upgrades."""
    ctx = ssl.create_default_context(ssl.Purpose.SERVER_AUTH)
    if not verify_ssl:
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
    return ctx


def _imap_capabilities(imap: aioimaplib.IMAP4) -> set[str]:
    """Return normalized capabilities from an aioimaplib protocol."""
    return {
        capability.decode("utf-8", errors="replace").upper()
        if isinstance(capability, bytes)
        else str(capability).upper()
        for capability in getattr(imap.protocol, "capabilities", ())
    }


async def _imap_starttls(imap: aioimaplib.IMAP4, ssl_context: ssl.SSLContext, host: str) -> None:
    """Upgrade an IMAP connection to TLS via STARTTLS."""
    capabilities = _imap_capabilities(imap)
    if "STARTTLS" not in capabilities:
        await imap.protocol.capability()
        capabilities = _imap_capabilities(imap)
    if "STARTTLS" not in capabilities:
        raise OSError("IMAP server does not advertise STARTTLS capability")

    response = await imap.protocol.execute(
        aioimaplib.Command("STARTTLS", imap.protocol.new_tag(), loop=imap.protocol.loop)
    )
    status = _imap_status(response)
    if status != "OK":
        raise OSError(f"STARTTLS command failed: {status}")

    loop = asyncio.get_running_loop()
    tls_transport = await loop.start_tls(
        imap.protocol.transport,
        imap.protocol,
        ssl_context,
        server_hostname=host,
    )
    imap.protocol.transport = tls_transport
    await imap.protocol.capability()


# Backwards-compatible alias
_create_smtp_ssl_context = _create_ssl_context


class EmailClient:
    def __init__(self, email_server: EmailServer, sender: str | None = None):
        self.email_server = email_server
        self.sender = sender or email_server.user_name

        self.imap_class = aioimaplib.IMAP4_SSL if self.email_server.use_ssl else aioimaplib.IMAP4

        self.smtp_use_tls = self.email_server.use_ssl
        self.smtp_start_tls = self.email_server.start_ssl
        self.smtp_verify_ssl = self.email_server.verify_ssl

    def _imap_connect(self, server: "EmailServer | None" = None) -> aioimaplib.IMAP4_SSL | aioimaplib.IMAP4:
        """Create a new IMAP connection with the configured SSL context."""
        srv = server or self.email_server
        # Use self.imap_class (patchable in tests) for the primary server;
        # reconstruct for credential overrides that may differ in SSL mode.
        if srv is self.email_server:
            imap_cls = self.imap_class
        else:
            imap_cls = aioimaplib.IMAP4_SSL if srv.use_ssl else aioimaplib.IMAP4
        if srv.use_ssl:
            imap_ssl_context = _create_ssl_context(srv.verify_ssl)
            return imap_cls(srv.host, srv.port, ssl_context=imap_ssl_context)
        return imap_cls(srv.host, srv.port)

    @asynccontextmanager
    async def _imap_connection(
        self,
        mailbox: str | None = DEFAULT_MAILBOX,
        *,
        server: "EmailServer | None" = None,
    ) -> AsyncGenerator[aioimaplib.IMAP4 | aioimaplib.IMAP4_SSL, None]:
        """Async context manager for IMAP connections — the single connection idiom.

        Handles connect, greeting, optional STARTTLS upgrade, login (raising on a
        non-OK result), IMAP ID, optional mailbox selection (raising on a non-OK
        SELECT), and logout.

        Args:
            mailbox: Mailbox to select, or None to skip selection (e.g. for LIST).
            server: Override server credentials (e.g. the incoming server for
                append_to_sent / save_to_mailbox).
        """
        srv = server or self.email_server
        imap = self._imap_connect(server=srv)
        try:
            await imap._client_task
            await imap.wait_hello_from_server()
            if srv.start_ssl:
                ssl_context = _create_starttls_ssl_context(srv.verify_ssl)
                await _imap_starttls(imap, ssl_context, srv.host)
            await _imap_login(imap, srv.user_name, srv.password.get_secret_value())
            await _send_imap_id(imap)
            if mailbox is not None:
                select_response = await imap.select(_quote_mailbox(mailbox))
                _raise_for_imap_error(select_response, f"SELECT mailbox {mailbox}")
            yield imap
        finally:
            try:
                await imap.logout()
            except Exception as e:
                logger.info(f"Error during logout: {e}")

    def _get_smtp_ssl_context(self) -> ssl.SSLContext | None:
        """Get SSL context for SMTP connections based on verify_ssl setting."""
        return _create_ssl_context(self.smtp_verify_ssl)

    @staticmethod
    def _parse_recipients(email_message) -> list[str]:
        """Extract recipient addresses from To and Cc headers."""
        recipients = []
        to_header = email_message.get("To", "")
        if to_header:
            recipients = [addr.strip() for addr in to_header.split(",")]
        cc_header = email_message.get("Cc", "")
        if cc_header:
            recipients.extend([addr.strip() for addr in cc_header.split(",")])
        return recipients

    @staticmethod
    def _parse_date(date_str: str) -> datetime:
        """Parse email date string to datetime, with fallback to current time."""
        try:
            date_tuple = email.utils.parsedate_tz(date_str)
            if date_tuple:
                return datetime.fromtimestamp(email.utils.mktime_tz(date_tuple), tz=timezone.utc)
            return datetime.now(timezone.utc)
        except Exception:
            return datetime.now(timezone.utc)

    @staticmethod
    def _normalize_attachment_name(name: str) -> str:
        """Normalize attachment filenames for robust MIME round-trip matching."""
        return unicodedata.normalize("NFC", name)

    @staticmethod
    def _is_attachment_part(part) -> bool:
        """Determine whether a MIME part should be treated as an attachment.

        A strict check on ``Content-Disposition: attachment`` misses a common case:
        many clients (notably Apple Mail on iOS/macOS) send images, PDFs and other
        files with ``Content-Disposition: inline`` (or no disposition header at all)
        but with a filename parameter on the part. Those parts are real, user-facing
        attachments — the user uploaded a file and expects it to show up — even
        though they're inlined into the body via Content-ID references.

        Treat a part as an attachment when:
          - the disposition explicitly says ``attachment``, OR
          - the part carries a filename (works for ``inline`` or no disposition).

        Multipart container parts and bodyless text parts have no filename and an
        empty disposition, so they are correctly excluded.
        """
        content_disposition = str(part.get("Content-Disposition", "")).lower()
        if "attachment" in content_disposition:
            return True
        filename = part.get_filename()
        # Be defensive: only trust real string filenames. (Unconfigured MagicMock
        # instances in older tests return truthy MagicMock objects from
        # ``get_filename`` and would otherwise misclassify text parts.)
        return isinstance(filename, str) and bool(filename)

    def _parse_email_data(  # noqa: C901
        self,
        raw_email: bytes,
        email_id: str | None = None,
        body_offset: int = 0,
        max_body_length: int | None = MAX_BODY_LENGTH,
    ) -> dict[str, Any]:
        """Parse raw email data into a structured dictionary."""
        parser = BytesParser(policy=default)
        email_message = parser.parsebytes(raw_email)

        # Extract email parts
        subject = email_message.get("Subject", "")
        sender = email_message.get("From", "")
        date_str = email_message.get("Date", "")

        # Extract Message-ID for reply threading
        message_id = email_message.get("Message-ID")

        # Extract recipients and parse date
        to_addresses = self._parse_recipients(email_message)
        date = self._parse_date(date_str)

        # Get body content
        body = ""
        html_body = ""
        attachments = []

        if email_message.is_multipart():
            html_body = ""
            for part in email_message.walk():
                content_type = part.get_content_type()

                # Handle attachments — including inline-disposition parts with a
                # filename (Apple Mail commonly sends photos this way).
                if self._is_attachment_part(part):
                    filename = part.get_filename()
                    if filename:
                        attachments.append(filename)
                # Handle text parts - prefer text/plain
                elif content_type == "text/plain":
                    body_part = part.get_payload(decode=True)
                    if body_part:
                        charset = part.get_content_charset("utf-8")
                        try:
                            body += body_part.decode(charset)
                        except UnicodeDecodeError:
                            body += body_part.decode("utf-8", errors="replace")
                elif content_type == "text/html":
                    html_part = part.get_payload(decode=True)
                    if html_part:
                        charset = part.get_content_charset("utf-8")
                        try:
                            html_body += html_part.decode(charset)
                        except UnicodeDecodeError:
                            html_body += html_part.decode("utf-8", errors="replace")
            # Fall back to converted HTML if no plain text was found
            if not body and html_body:
                body = html_to_text(html_body)
        else:
            # Handle single-part emails
            content_type = email_message.get_content_type()
            payload = email_message.get_payload(decode=True)
            if payload:
                charset = email_message.get_content_charset("utf-8")
                try:
                    raw_body = payload.decode(charset)
                except UnicodeDecodeError:
                    raw_body = payload.decode("utf-8", errors="replace")
                if content_type == "text/html":
                    html_body = raw_body
                    body = html_to_text(raw_body)
                else:
                    body = raw_body
        if body_offset < 0:
            raise ValueError("body_offset must be >= 0")

        # Windowed body read. ``max_body_length`` of 0 or None means "no limit": return the full
        # body from ``body_offset`` onward. Otherwise return at most ``max_body_length`` characters
        # and append the ``...[TRUNCATED]`` marker when more of the body remains, so callers can
        # page through a long email by re-requesting with ``body_offset += max_body_length``.
        if body:
            if max_body_length:
                window = body[body_offset : body_offset + max_body_length]
                if body_offset + max_body_length < len(body):
                    window += "...[TRUNCATED]"
                body = window
            else:
                body = body[body_offset:]
        return {
            "email_id": email_id or "",
            "message_id": message_id,
            "subject": subject,
            "from": sender,
            "to": to_addresses,
            "body": body,
            "html_body": html_body,
            "date": date,
            "attachments": attachments,
        }

    @staticmethod
    def _sanitize_imap_value(value: str) -> str:
        """Sanitize a string value for IMAP search criteria.

        For multi-word values, strips embedded double quotes (invalid per RFC 3501
        Section 4.3) and wraps in double quotes. Single-word values pass through unchanged.
        """
        if " " not in value:
            return value
        sanitized = value.replace('"', "")
        return f'"{sanitized}"'

    @staticmethod
    def _build_search_criteria(
        before: datetime | None = None,
        since: datetime | None = None,
        subject: str | None = None,
        body: str | None = None,
        text: str | None = None,
        from_address: str | None = None,
        to_address: str | None = None,
        seen: bool | None = None,
        flagged: bool | None = None,
        answered: bool | None = None,
        has_attachment: bool | None = None,
    ) -> list[str]:
        search_criteria = []
        if before:
            search_criteria.extend(["BEFORE", before.strftime("%d-%b-%Y").upper()])
        if since:
            search_criteria.extend(["SINCE", since.strftime("%d-%b-%Y").upper()])
        # Substring-match fields (IMAP keyword, value)
        text_criteria = [
            ("SUBJECT", subject),
            ("BODY", body),
            ("TEXT", text),
            ("FROM", from_address),
            ("TO", to_address),
        ]
        for keyword, value in text_criteria:
            if value:
                search_criteria.extend([keyword, EmailClient._sanitize_imap_value(value)])

        # Attachment heuristic: most attachments are carried in multipart/mixed.
        # May miss some types (e.g. inline images) or yield false positives.
        if has_attachment is True:
            search_criteria.extend(["HEADER", "Content-Type", "multipart/mixed"])
        elif has_attachment is False:
            search_criteria.extend(["NOT", "HEADER", "Content-Type", "multipart/mixed"])

        # Flag-based criteria using mapping to reduce complexity
        flag_criteria = [
            (seen, {True: "SEEN", False: "UNSEEN"}),
            (flagged, {True: "FLAGGED", False: "UNFLAGGED"}),
            (answered, {True: "ANSWERED", False: "UNANSWERED"}),
        ]
        for flag_value, criteria_map in flag_criteria:
            if flag_value in criteria_map:
                search_criteria.append(criteria_map[flag_value])

        return search_criteria or ["ALL"]

    def _parse_headers(self, email_id: str, raw_headers: bytes) -> dict[str, Any] | None:
        """Parse raw email headers into metadata dictionary.

        Note: this parses only header data (BODY.PEEK[HEADER]) so it cannot
        populate the attachments list — that requires fetching BODYSTRUCTURE
        or the full body. The attachments list is intentionally returned
        empty here; ``_parse_email_data`` populates it from the full body.
        """
        try:
            parser = BytesParser(policy=default)
            email_message = parser.parsebytes(raw_headers)

            subject = email_message.get("Subject", "")
            sender = email_message.get("From", "")
            date_str = email_message.get("Date", "")
            # Expose Message-ID for reply threading and de-duplication on the client.
            message_id = email_message.get("Message-ID")

            to_addresses = self._parse_recipients(email_message)
            date = self._parse_date(date_str)

            return {
                "email_id": email_id,
                "message_id": message_id,
                "subject": subject,
                "from": sender,
                "to": to_addresses,
                "date": date,
                "attachments": [],
            }
        except Exception as e:
            logger.error(f"Error parsing email headers: {e!s}")
            return None

    async def _fetch_dates_chunk(
        self,
        imap: aioimaplib.IMAP4_SSL | aioimaplib.IMAP4,
        chunk: list[bytes],
        chunk_num: int,
        total_chunks: int,
        timeout: float = 30.0,
    ) -> dict[str, datetime]:
        """Fetch INTERNALDATE for a single chunk of UIDs."""
        uid_list = ",".join(uid.decode() for uid in chunk)
        chunk_start = time.perf_counter()
        _, data = await asyncio.wait_for(
            imap.uid("fetch", uid_list, "(INTERNALDATE)"),
            timeout=timeout,
        )
        chunk_elapsed = time.perf_counter() - chunk_start

        chunk_dates: dict[str, datetime] = {}
        for item in data:
            if not isinstance(item, bytes) or b"INTERNALDATE" not in item:
                continue
            uid_match = re.search(rb"UID (\d+)", item)
            date_match = re.search(rb'INTERNALDATE "([^"]+)"', item)
            if uid_match and date_match:
                uid = uid_match.group(1).decode()
                date_str = date_match.group(1).decode().strip()
                chunk_dates[uid] = datetime.strptime(date_str, "%d-%b-%Y %H:%M:%S %z")

        if total_chunks > 1:
            logger.info(f"Fetched dates chunk {chunk_num}/{total_chunks}: {len(chunk)} UIDs in {chunk_elapsed:.2f}s")

        return chunk_dates

    async def _batch_fetch_dates(
        self,
        imap: aioimaplib.IMAP4_SSL | aioimaplib.IMAP4,
        email_ids: list[bytes],
        chunk_size: int = 500,
    ) -> dict[str, datetime]:
        """Batch fetch INTERNALDATE for all UIDs in sequential chunks.

        Uses a conservative chunk_size (default 500) to avoid hitting
        Python's recursion limit in aioimaplib's recursive response parser
        (see: aioimaplib _handle_responses). IMAP connections are sequential
        by protocol, so chunks must be fetched serially — not in parallel.
        """
        if not email_ids:
            return {}

        # Split into chunks
        chunks = [email_ids[i : i + chunk_size] for i in range(0, len(email_ids), chunk_size)]
        total_chunks = len(chunks)

        # Fetch chunks sequentially (IMAP protocol is sequential on a single connection)
        uid_dates: dict[str, datetime] = {}
        for chunk_num, chunk in enumerate(chunks, 1):
            chunk_dates = await self._fetch_dates_chunk(imap, chunk, chunk_num, total_chunks)
            uid_dates.update(chunk_dates)

        return uid_dates

    async def _batch_fetch_headers(
        self,
        imap: aioimaplib.IMAP4_SSL | aioimaplib.IMAP4,
        email_ids: list[bytes] | list[str],
    ) -> dict[str, dict[str, Any]]:
        """Batch fetch headers for a list of UIDs."""
        if not email_ids:
            return {}

        # Normalize to list of strings
        str_ids = [uid.decode() if isinstance(uid, bytes) else uid for uid in email_ids]
        uid_list = ",".join(str_ids)
        _, data = await imap.uid("fetch", uid_list, "BODY.PEEK[HEADER]")

        results: dict[str, dict[str, Any]] = {}
        for i, item in enumerate(data):
            if not isinstance(item, bytes) or b"BODY[HEADER]" not in item:
                continue
            # First try to find UID in the same line (standard format)
            uid_match = re.search(rb"UID (\d+)", item)
            if uid_match and i + 1 < len(data) and isinstance(data[i + 1], bytearray):
                uid = uid_match.group(1).decode()
                raw_headers = bytes(data[i + 1])
                metadata = self._parse_headers(uid, raw_headers)
                if metadata:
                    results[uid] = metadata
            # Proton Bridge format: UID comes AFTER header data in a separate item
            # Format: [i]=b'N FETCH (BODY[HEADER] {size}', [i+1]=bytearray(headers), [i+2]=b' UID xxx)'
            elif i + 2 < len(data) and isinstance(data[i + 1], bytearray):
                uid_after_match = re.search(rb"UID (\d+)", data[i + 2]) if isinstance(data[i + 2], bytes) else None
                if uid_after_match:
                    uid = uid_after_match.group(1).decode()
                    raw_headers = bytes(data[i + 1])
                    metadata = self._parse_headers(uid, raw_headers)
                    if metadata:
                        results[uid] = metadata

        return results

    async def _batch_fetch_senders(
        self,
        imap: aioimaplib.IMAP4_SSL | aioimaplib.IMAP4,
        email_ids: list[bytes] | list[str],
        chunk_size: int = 500,
    ) -> dict[str, str]:
        """Batch fetch the From header for all UIDs (chunked, sequential), for allowlist filtering.

        Returns {uid: raw From header}. Fetches only HEADER.FIELDS (FROM) to stay light and reuses
        _parse_headers (which tolerates a From-only header block).
        """
        if not email_ids:
            return {}

        chunks = [email_ids[i : i + chunk_size] for i in range(0, len(email_ids), chunk_size)]
        senders: dict[str, str] = {}
        for chunk in chunks:
            str_ids = [uid.decode() if isinstance(uid, bytes) else uid for uid in chunk]
            uid_list = ",".join(str_ids)
            fetch_response = await imap.uid("fetch", uid_list, "BODY.PEEK[HEADER.FIELDS (FROM)]")
            _raise_for_imap_command_failure(fetch_response, f"FETCH From headers for UIDs {uid_list}")
            _, data = fetch_response
            for i, item in enumerate(data):
                if not isinstance(item, bytes) or b"BODY[HEADER" not in item:
                    continue
                uid_match = re.search(rb"UID (\d+)", item)
                if uid_match and i + 1 < len(data) and isinstance(data[i + 1], bytearray):
                    meta = self._parse_headers(uid_match.group(1).decode(), bytes(data[i + 1]))
                    if meta:
                        senders[meta["email_id"]] = meta["from"]
                elif i + 2 < len(data) and isinstance(data[i + 1], bytearray):
                    uid_after = re.search(rb"UID (\d+)", data[i + 2]) if isinstance(data[i + 2], bytes) else None
                    if uid_after:
                        meta = self._parse_headers(uid_after.group(1).decode(), bytes(data[i + 1]))
                        if meta:
                            senders[meta["email_id"]] = meta["from"]
        return senders

    async def _enforce_sender_allowlist(
        self,
        imap: aioimaplib.IMAP4_SSL | aioimaplib.IMAP4,
        email_id: str,
        allowed_senders: list[str] | None,
    ) -> None:
        """Raise ValueError (identical to not-found) when sender is not on the allowlist.

        No-op when ``allowed_senders`` is empty or None (backwards-compatible).
        """
        if allowed_senders:
            uid_senders = await self._batch_fetch_senders(imap, [email_id])
            if not sender_allowed(uid_senders.get(email_id, ""), allowed_senders):
                msg = f"Failed to fetch email with UID {email_id}"
                logger.error(msg)
                raise ValueError(msg)

    async def _blocked_uids(
        self,
        imap: aioimaplib.IMAP4_SSL | aioimaplib.IMAP4,
        email_ids: list[str],
        allowed_senders: list[str] | None,
    ) -> set[str]:
        """UIDs whose From is not on the allowlist. Empty set when no allowlist (no IMAP work)."""
        if not allowed_senders:
            return set()
        uid_senders = await self._batch_fetch_senders(imap, email_ids)
        return {uid for uid in email_ids if not sender_allowed(uid_senders.get(uid, ""), allowed_senders)}

    async def get_emails_metadata_page(
        self,
        page: int = 1,
        page_size: int = 10,
        before: datetime | None = None,
        since: datetime | None = None,
        subject: str | None = None,
        from_address: str | None = None,
        to_address: str | None = None,
        order: str = "desc",
        mailbox: str = DEFAULT_MAILBOX,
        seen: bool | None = None,
        flagged: bool | None = None,
        answered: bool | None = None,
        body: str | None = None,
        text: str | None = None,
        has_attachment: bool | None = None,
        allowed_senders: list[str] | None = None,
    ) -> tuple[list[dict[str, Any]], int]:
        """Fetch email metadata for a page, returning (metadata_list, total_count) in a single IMAP session."""
        async with self._imap_connection(mailbox) as imap:
            search_criteria = self._build_search_criteria(
                before,
                since,
                subject,
                body=body,
                text=text,
                from_address=from_address,
                to_address=to_address,
                seen=seen,
                flagged=flagged,
                answered=answered,
                has_attachment=has_attachment,
            )
            logger.info(f"Get metadata: Search criteria: {search_criteria}")

            # Search for messages - use UID SEARCH for better compatibility.
            # charset handling: aioimaplib defaults to "CHARSET utf-8", which
            # Microsoft Exchange rejects with `NO [BADCHARSET (US-ASCII)] The
            # specified charset is not supported.`, breaking all search/list
            # operations — so ASCII-only user searches omit the CHARSET token.
            # However, RFC 3501 requires a charset declaration for non-ASCII
            # search values: without it, servers such as Coremail interpret the
            # raw UTF-8 bytes as US-ASCII and silently return zero matches. Base
            # the decision only on user-supplied text fields, not on generated
            # criteria such as locale-dependent date strings.
            search_text_values = (subject, body, text, from_address, to_address)
            charset = "utf-8" if any(value and not value.isascii() for value in search_text_values) else None
            _, messages = await imap.uid_search(*search_criteria, charset=charset)

            # Handle empty or None responses
            if not messages or not messages[0]:
                logger.warning("No messages returned from search")
                return [], 0

            raw_ids = messages[0].split()

            # Sender allowlist: filter candidates BEFORE sorting/pagination so total + pages stay honest.
            if allowed_senders:
                uid_senders = await self._batch_fetch_senders(imap, raw_ids)
                raw_ids = [
                    uid for uid in raw_ids if sender_allowed(uid_senders.get(uid.decode(), ""), allowed_senders)
                ]
                logger.info(f"Sender allowlist active: {len(raw_ids)} of {len(uid_senders)} match")

            total = len(raw_ids)
            logger.info(f"Found {total} email IDs")

            # Normalize byte UIDs to strings once for use throughout
            str_ids = [uid.decode() if isinstance(uid, bytes) else uid for uid in raw_ids]

            if not raw_ids:
                return [], 0

            # Phase 1: Batch fetch INTERNALDATE for sorting (sequential chunks)
            fetch_dates_start = time.perf_counter()
            uid_dates = await self._batch_fetch_dates(imap, raw_ids)
            fetch_dates_elapsed = time.perf_counter() - fetch_dates_start

            if uid_dates and len(uid_dates) < total:
                logger.warning(
                    f"INTERNALDATE fetch missed {total - len(uid_dates)}/{total} UIDs"
                )

            if uid_dates:
                # Normal path: sort by INTERNALDATE
                sorted_uids = sorted(uid_dates.items(), key=lambda x: x[1], reverse=(order == "desc"))
                ordered_uids = [uid for uid, _ in sorted_uids]
            else:
                # Fallback: UID ordering (higher UID ~ newer email)
                logger.warning(
                    f"INTERNALDATE fetch returned no results for {total} UIDs, "
                    f"falling back to UID ordering"
                )
                ordered_uids = sorted(str_ids, key=int, reverse=(order == "desc"))

            # Paginate
            start = (page - 1) * page_size
            page_uids = ordered_uids[start : start + page_size]

            if not page_uids:
                logger.info(f"Phase 1 (dates): {len(uid_dates)} UIDs in {fetch_dates_elapsed:.2f}s, page {page} empty")
                return [], total

            # Phase 2: Batch fetch headers for requested page only
            fetch_headers_start = time.perf_counter()
            metadata_by_uid = await self._batch_fetch_headers(imap, page_uids)
            fetch_headers_elapsed = time.perf_counter() - fetch_headers_start

            logger.info(
                f"Fetched page {page}: {fetch_dates_elapsed:.2f}s dates ({len(uid_dates)} UIDs), "
                f"{fetch_headers_elapsed:.2f}s headers ({len(page_uids)} UIDs)"
            )

            # Collect in sorted order
            results = [metadata_by_uid[uid] for uid in page_uids if uid in metadata_by_uid]
            if len(results) < len(page_uids):
                logger.warning(
                    f"Header fetch missed {len(page_uids) - len(results)}/{len(page_uids)} UIDs on page {page}"
                )
            return results, total

    def _check_email_content(self, data: list) -> bool:
        """Check if the fetched data contains actual email content."""
        for item in data:
            if isinstance(item, bytes) and b"FETCH (" in item and b"RFC822" not in item and b"BODY" not in item:
                # This is just metadata, not actual content
                continue
            elif isinstance(item, bytes | bytearray) and len(item) > 100:
                # This looks like email content
                return True
        return False

    def _extract_raw_email(self, data: list) -> bytes | None:
        """Extract raw email bytes from IMAP response data."""
        # The email content is typically at index 1 as a bytearray
        if len(data) > 1 and isinstance(data[1], bytearray):
            return bytes(data[1])

        # Search through all items for email content
        for item in data:
            if isinstance(item, bytes | bytearray) and len(item) > 100:
                # Skip IMAP protocol responses
                if isinstance(item, bytes) and b"FETCH" in item:
                    continue
                # This is likely the email content
                return bytes(item) if isinstance(item, bytearray) else item
        return None

    async def _fetch_email_with_formats(self, imap, email_id: str) -> list | None:
        """Try non-mutating fetch formats to get email data."""
        fetch_formats = ["BODY.PEEK[]", "(BODY.PEEK[])"]

        for fetch_format in fetch_formats:
            try:
                response = await imap.uid("fetch", email_id, fetch_format)
                _raise_for_imap_error(response, f"FETCH email {email_id} with {fetch_format}")
                _, data = response

                if data and len(data) > 0 and self._check_email_content(data):
                    return data

            except Exception as e:
                logger.debug(f"Fetch format {fetch_format} failed: {e}")

        return None

    async def get_email_body_by_id(
        self,
        email_id: str,
        mailbox: str = DEFAULT_MAILBOX,
        mark_as_read: bool = False,
        allowed_senders: list[str] | None = None,
        body_offset: int = 0,
        max_body_length: int | None = MAX_BODY_LENGTH,
    ) -> dict[str, Any] | None:
        async with self._imap_connection(mailbox) as imap:
            # Sender allowlist: check the From header BEFORE reading the body, so a blocked
            # message is never fetched/parsed, never marked read, and is indistinguishable from
            # a missing/inaccessible one (caller sees None either way).
            if allowed_senders:
                uid_senders = await self._batch_fetch_senders(imap, [email_id])
                if not sender_allowed(uid_senders.get(email_id, ""), allowed_senders):
                    return None

            # Fetch the specific email by UID without implicitly marking it as read
            data = await self._fetch_email_with_formats(imap, email_id)
            if not data:
                logger.error(f"Failed to fetch UID {email_id} with any format")
                return None

            raw_email = self._extract_raw_email(data)
            if not raw_email:
                logger.error(f"Could not find email data in response for email ID: {email_id}")
                return None

            try:
                email_data = self._parse_email_data(
                    raw_email, email_id, body_offset=body_offset, max_body_length=max_body_length
                )
            except Exception as e:
                logger.error(f"Error parsing email: {e!s}")
                return None

            if mark_as_read:
                try:
                    store_response = await imap.uid("store", email_id, "+FLAGS", r"(\Seen)")
                    _raise_for_imap_error(store_response, f"STORE \\Seen for email {email_id}")
                except Exception as e:
                    logger.warning(f"Failed to mark email {email_id} as read: {e}")

            return email_data

    async def download_attachment(
        self,
        email_id: str,
        attachment_name: str,
        save_path: str,
        mailbox: str = DEFAULT_MAILBOX,
        allowed_senders: list[str] | None = None,
    ) -> dict[str, Any]:
        """Download a specific attachment from an email and save it to disk.

        Args:
            email_id: The UID of the email containing the attachment.
            attachment_name: The filename of the attachment to download.
            save_path: The local path where the attachment will be saved.
            mailbox: The mailbox to search in (default: "INBOX").
            allowed_senders: Optional sender allowlist; when set, a non-allowed sender's
                message is treated as not found and its body is never fetched.

        Returns:
            A dictionary with download result information.
        """
        async with self._imap_connection(mailbox) as imap:
            # Read-path allowlist: check the From header before fetching the body, so a
            # blocked sender's message is never read. Blocked fails identically to a missing
            # UID (same ValueError below), so it does not reveal whether the message exists.
            await self._enforce_sender_allowlist(imap, email_id, allowed_senders)

            data = await self._fetch_email_with_formats(imap, email_id)
            if not data:
                msg = f"Failed to fetch email with UID {email_id}"
                logger.error(msg)
                raise ValueError(msg)

            raw_email = self._extract_raw_email(data)
            if not raw_email:
                msg = f"Could not find email data for email ID: {email_id}"
                logger.error(msg)
                raise ValueError(msg)

            parser = BytesParser(policy=default)
            email_message = parser.parsebytes(raw_email)

            # Find the attachment
            attachment_data = None
            mime_type = None
            normalized_attachment_name = self._normalize_attachment_name(attachment_name)

            if email_message.is_multipart():
                for part in email_message.walk():
                    # Match attachments listed by ``_parse_email_data`` — this includes
                    # inline-disposition parts with a filename (e.g. iOS Mail photos).
                    if not self._is_attachment_part(part):
                        continue
                    filename = part.get_filename()
                    if not isinstance(filename, str):
                        continue
                    if self._normalize_attachment_name(filename) == normalized_attachment_name:
                        attachment_data = part.get_payload(decode=True)
                        mime_type = part.get_content_type()
                        break

            if attachment_data is None:
                msg = f"Attachment '{attachment_name}' not found in email {email_id}"
                logger.error(msg)
                raise ValueError(msg)

            # Save to disk
            save_file = Path(save_path)
            save_file.parent.mkdir(parents=True, exist_ok=True)
            save_file.write_bytes(attachment_data)

            logger.info(f"Attachment '{attachment_name}' saved to {save_path}")

            return {
                "email_id": email_id,
                "attachment_name": attachment_name,
                "mime_type": mime_type or "application/octet-stream",
                "size": len(attachment_data),
                "saved_path": str(save_file.resolve()),
            }

    async def extract_attachments(
        self,
        email_id: str,
        mailbox: str = DEFAULT_MAILBOX,
    ) -> list[tuple[str, str, bytes]]:
        """Extract all attachments from an email as in-memory data."""
        async with self._imap_connection(mailbox) as imap:
            data = await self._fetch_email_with_formats(imap, email_id)
            if not data:
                return []

            raw_email = self._extract_raw_email(data)
            if not raw_email:
                return []

            parser = BytesParser(policy=default)
            email_message = parser.parsebytes(raw_email)

            attachments = []
            if email_message.is_multipart():
                for part in email_message.walk():
                    content_disposition = str(part.get("Content-Disposition", ""))
                    if "attachment" in content_disposition:
                        filename = part.get_filename()
                        if filename:
                            attachment_data = part.get_payload(decode=True)
                            if attachment_data:
                                mime_type = part.get_content_type() or "application/octet-stream"
                                attachments.append((filename, mime_type, attachment_data))

            return attachments

    def _validate_attachment(self, file_path: str) -> Path:
        """Validate attachment file path."""
        path = Path(file_path)
        if not path.exists():
            msg = f"Attachment file not found: {file_path}"
            logger.error(msg)
            raise FileNotFoundError(msg)

        if not path.is_file():
            msg = f"Attachment path is not a file: {file_path}"
            logger.error(msg)
            raise ValueError(msg)

        return path

    def _create_attachment_part(self, path: Path) -> MIMEApplication:
        """Create MIME attachment part from file."""
        with open(path, "rb") as f:
            file_data = f.read()

        mime_type, _ = mimetypes.guess_type(str(path))
        if mime_type is None:
            mime_type = "application/octet-stream"

        attachment_part = MIMEApplication(file_data, _subtype=mime_type.split("/")[1])
        attachment_part.add_header(
            "Content-Disposition",
            "attachment",
            filename=path.name,
        )
        logger.info(f"Attached file: {path.name} ({mime_type})")
        return attachment_part

    def _create_message_with_attachments(self, body: str, html: bool, attachments: list[str]) -> MIMEMultipart:
        """Create multipart message with attachments."""
        msg = MIMEMultipart()
        content_type = "html" if html else "plain"
        text_part = MIMEText(body, content_type, "utf-8")
        msg.attach(text_part)

        for file_path in attachments:
            try:
                path = self._validate_attachment(file_path)
                attachment_part = self._create_attachment_part(path)
                msg.attach(attachment_part)
            except Exception as e:
                logger.error(f"Failed to attach file {file_path}: {e}")
                raise

        return msg

    def compose_message(
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
        extra_parts: list[MIMEApplication] | None = None,
        include_bcc_header: bool = False,
        reply_to: str | None = None,
    ) -> MIMEText | MIMEMultipart:
        """Compose an email message without sending it.

        Builds MIME structure, sets headers (Subject, From, To, Cc, Date,
        Message-Id, threading headers). Synchronous — no I/O.

        When ``include_bcc_header`` is True (used for local IMAP storage such
        as Drafts or Sent copies), the Bcc header is included so mail clients
        can display the BCC recipients.  When False (default, used for SMTP
        sending), the Bcc header is omitted — BCC recipients are delivered
        via the SMTP envelope only.
        """
        # Convert body to HTML unless it's already raw HTML
        if not html:
            body = markdown_to_email_html(body, wrap_in_html=True)
            html = True

        # Create message with or without attachments
        if attachments or extra_parts:
            msg = self._create_message_with_attachments(body, html, attachments or [])
            if extra_parts:
                for part in extra_parts:
                    msg.attach(part)
        else:
            content_type = "html" if html else "plain"
            msg = MIMEText(body, content_type, "utf-8")

        # Handle subject with special characters
        if any(ord(c) > 127 for c in subject):
            msg["Subject"] = Header(subject, "utf-8")
        else:
            msg["Subject"] = subject

        sender_name, sender_address = email.utils.parseaddr(self.sender)

        # Encode only the display name. RFC 2047 encoded-words must not cover the
        # addr-spec, otherwise strict clients may show the raw encoded blob.
        if sender_address:
            msg["From"] = email.utils.formataddr((str(Header(sender_name, "utf-8")), sender_address))
        elif any(ord(c) > 127 for c in self.sender):
            msg["From"] = Header(self.sender, "utf-8")
        else:
            msg["From"] = self.sender

        msg["To"] = ", ".join(recipients)

        # Add CC header if provided (visible to recipients)
        if cc:
            msg["Cc"] = ", ".join(cc)

        # Add BCC header when saving locally (drafts, sent copies)
        if bcc and include_bcc_header:
            msg["Bcc"] = ", ".join(bcc)

        # Set threading headers for replies
        if in_reply_to:
            msg["In-Reply-To"] = in_reply_to
        if references:
            msg["References"] = references
        if reply_to:
            msg["Reply-To"] = reply_to

        # Set Date and Message-Id headers
        msg["Date"] = email.utils.formatdate(localtime=True)
        sender_for_domain = sender_address or self.sender
        sender_domain = sender_for_domain.rsplit("@", 1)[-1].rstrip(">")
        msg["Message-Id"] = email.utils.make_msgid(domain=sender_domain)

        return msg

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
        reply_to: str | None = None,
        extra_parts: list[MIMEApplication] | None = None,
    ) -> MIMEText | MIMEMultipart:
        msg = self.compose_message(
            recipients,
            subject,
            body,
            cc,
            bcc,
            html,
            attachments,
            in_reply_to,
            references,
            extra_parts=extra_parts,
            reply_to=reply_to,
        )

        async with aiosmtplib.SMTP(
            hostname=self.email_server.host,
            port=self.email_server.port,
            start_tls=self.smtp_start_tls,
            use_tls=self.smtp_use_tls,
            tls_context=self._get_smtp_ssl_context(),
        ) as smtp:
            await smtp.login(self.email_server.user_name, self.email_server.password.get_secret_value())

            # Create a combined list of all recipients for delivery
            all_recipients = recipients.copy()
            if cc:
                all_recipients.extend(cc)
            if bcc:
                all_recipients.extend(bcc)

            await smtp.send_message(msg, recipients=all_recipients)

        # Return the message for potential saving to Sent folder
        return msg

    async def _find_sent_folder_by_flag(self, imap) -> str | None:
        """Find the Sent folder by searching for the \\Sent IMAP flag.

        Args:
            imap: Connected IMAP client

        Returns:
            The folder name with the \\Sent flag, or None if not found
        """
        try:
            # List all folders - aioimaplib requires reference_name and mailbox_pattern
            _, folders = await imap.list('""', "*")

            # Search for folder with \Sent flag
            for folder in folders:
                mailbox = _parse_list_response(folder)
                if mailbox and r"\Sent" in mailbox.flags:
                    logger.info(f"Found Sent folder by \\Sent flag: '{mailbox.name}'")
                    return mailbox.name
        except Exception as e:
            logger.debug(f"Error finding Sent folder by flag: {e}")

        return None

    async def append_to_sent(
        self,
        msg: MIMEText | MIMEMultipart,
        incoming_server: EmailServer,
        sent_folder_name: str | None = None,
    ) -> bool:
        """Append a sent message to the IMAP Sent folder.

        Args:
            msg: The email message that was sent
            incoming_server: IMAP server configuration for accessing Sent folder
            sent_folder_name: Override folder name, or None for auto-detection

        Returns:
            True if successfully saved, False otherwise
        """
        # Build candidate list: user override first, then common names
        sent_folder_candidates = [sent_folder_name, *SENT_FOLDER_CANDIDATES] if sent_folder_name else list(SENT_FOLDER_CANDIDATES)

        try:
            async with self._imap_connection(mailbox=None, server=incoming_server) as imap:
                # Try to find Sent folder by IMAP \Sent flag first
                flag_folder = await self._find_sent_folder_by_flag(imap)
                if flag_folder and flag_folder not in sent_folder_candidates:
                    sent_folder_candidates.insert(0, flag_folder)

                # Try to find and use the Sent folder
                for folder in sent_folder_candidates:
                    try:
                        logger.debug(f"Trying Sent folder: '{folder}'")
                        result = await imap.select(_quote_mailbox(folder))
                        logger.debug(f"Select result for '{folder}': {result}")

                        if _is_ok(result):
                            msg_bytes = msg.as_bytes()
                            logger.debug(f"Appending message to '{folder}'")
                            append_result = await imap.append(
                                msg_bytes,
                                mailbox=_quote_mailbox(folder),
                                flags=IMAP_FLAG_SEEN,
                            )
                            logger.debug(f"Append result: {append_result}")
                            if _is_ok(append_result):
                                logger.info(f"Saved sent email to '{folder}'")
                                return True
                            else:
                                logger.warning(f"Failed to append to '{folder}': {append_result}")
                        else:
                            logger.debug(f"Folder '{folder}' select returned: {result}")
                    except Exception as e:
                        logger.debug(f"Folder '{folder}' not available: {e}")
                        continue

                logger.warning("Could not find a valid Sent folder to save the message")
                return False

        except ConnectionError:
            raise
        except Exception as e:
            logger.error(f"Error saving to Sent folder: {e}")
            return False

    @staticmethod
    async def _batch_uid_operation(
        imap: Any,
        email_ids: list[str],
        operation_fn: Callable,
        op_name: str,
    ) -> tuple[list[str], list[str]]:
        """Execute an async IMAP operation on each email_id, returning (succeeded, failed).

        operation_fn(imap, email_id) should raise on failure.
        """
        succeeded = []
        failed = []
        for email_id in email_ids:
            try:
                await operation_fn(imap, email_id)
                succeeded.append(email_id)
            except Exception as e:
                logger.error(f"Failed to {op_name} email {email_id}: {e}")
                failed.append(email_id)
        return succeeded, failed

    async def append_to_mailbox(
        self,
        msg: MIMEText | MIMEMultipart,
        incoming_server: EmailServer,
        mailbox: str,
        flags: str = r"(\Draft \Seen)",
    ) -> str | None:
        """Append a message to the specified IMAP folder.

        Unlike append_to_sent, this targets a single user-specified mailbox
        without folder discovery. Returns the IMAP UID of the appended message
        (if the server supports APPENDUID / RFC 4315), or ``"unknown"`` on
        success without UID, or ``None`` on failure.
        """
        try:
            async with self._imap_connection(mailbox=None, server=incoming_server) as imap:
                result = await imap.select(_quote_mailbox(mailbox))
                status = result[0] if isinstance(result, tuple) else result
                if str(status).upper() != "OK":
                    logger.warning(f"Mailbox '{mailbox}' not found or not selectable: {status}")
                    return None

                msg_bytes = msg.as_bytes()
                append_result = await imap.append(
                    msg_bytes,
                    mailbox=_quote_mailbox(mailbox),
                    flags=flags,
                )
                append_status = append_result[0] if isinstance(append_result, tuple) else append_result
                if str(append_status).upper() == "OK":
                    # Try to extract UID from APPENDUID response (RFC 4315)
                    uid = None
                    if isinstance(append_result, tuple) and len(append_result) > 1:
                        for part in append_result[1]:
                            part_str = part.decode("utf-8") if isinstance(part, bytes) else str(part)
                            match = re.search(r"APPENDUID\s+\d+\s+(\d+)", part_str, re.IGNORECASE)
                            if match:
                                uid = match.group(1)
                                break
                    logger.info(f"Saved email to '{mailbox}'" + (f" (UID {uid})" if uid else ""))
                    return uid or "unknown"
                logger.warning(f"Failed to append to '{mailbox}': {append_status}")
                return None
        except ConnectionError:
            raise
        except Exception as e:
            logger.error(f"Error saving to mailbox '{mailbox}': {e}")
            return None

    async def delete_emails(
        self,
        email_ids: list[str],
        mailbox: str = DEFAULT_MAILBOX,
        allowed_senders: list[str] | None = None,
        report_blocked_mutations: bool = False,
    ) -> tuple[list[str], list[str]]:
        """Delete emails by their UIDs. Returns (deleted_ids, failed_ids).

        A blocked sender's UID is never flagged \\Deleted: by default it is reported as a
        no-op success (indistinguishable from a nonexistent UID); when
        report_blocked_mutations is True it is reported in failed_ids instead.
        """
        async with self._imap_connection(mailbox) as imap:
            blocked = await self._blocked_uids(imap, email_ids, allowed_senders)
            allowed_ids = [email_id for email_id in email_ids if email_id not in blocked]

            async def _op(imap: Any, email_id: str) -> None:
                await imap.uid("store", email_id, "+FLAGS", IMAP_FLAG_DELETED)

            deleted, failed = await self._batch_uid_operation(imap, allowed_ids, _op, "delete")
            if deleted:  # only expunge when messages were actually flagged \Deleted (never for no-op blocked UIDs)
                try:
                    expunge_response = await imap.expunge()
                    _raise_for_imap_error(expunge_response, "EXPUNGE deleted emails")
                except Exception as e:
                    logger.error(f"Failed to expunge deleted emails: {e}")
                    failed.extend(deleted)
                    deleted = []

        # Blocked UIDs are never flagged: reported as no-op success by default, or in
        # failed_ids when report_blocked_mutations is set.
        blocked_ids = [email_id for email_id in email_ids if email_id in blocked]
        if report_blocked_mutations:
            failed.extend(blocked_ids)
        else:
            deleted.extend(blocked_ids)
        return deleted, failed

    async def copy_emails(
        self,
        email_ids: list[str],
        destination_folder: str,
        source_mailbox: str = DEFAULT_MAILBOX,
    ) -> tuple[list[str], list[str]]:
        """Copy emails to a destination folder. Returns (copied_ids, failed_ids)."""
        async with self._imap_connection(source_mailbox) as imap:
            async def _op(imap: Any, email_id: str) -> None:
                result = await imap.uid("copy", email_id, _quote_mailbox(destination_folder))
                if not _is_ok(result):
                    raise RuntimeError(f"COPY failed: {result}")

            return await self._batch_uid_operation(imap, email_ids, _op, "copy")

    async def create_folder(self, folder_name: str) -> tuple[bool, str]:
        """Create a new folder. Returns (success, message)."""
        try:
            async with self._imap_connection(mailbox=None) as imap:
                result = await imap.create(_quote_mailbox(folder_name))
                if _is_ok(result):
                    logger.info(f"Created folder: {folder_name}")
                    return True, f"Folder '{folder_name}' created successfully"
                else:
                    logger.error(f"Failed to create folder {folder_name}: {result}")
                    return False, f"Failed to create folder: {result}"
        except Exception as e:
            logger.error(f"Error creating folder {folder_name}: {e}")
            return False, f"Error creating folder: {e}"

    async def delete_folder(self, folder_name: str) -> tuple[bool, str]:
        """Delete a folder. Returns (success, message)."""
        try:
            async with self._imap_connection(mailbox=None) as imap:
                result = await imap.delete(_quote_mailbox(folder_name))
                if _is_ok(result):
                    logger.info(f"Deleted folder: {folder_name}")
                    return True, f"Folder '{folder_name}' deleted successfully"
                else:
                    logger.error(f"Failed to delete folder {folder_name}: {result}")
                    return False, f"Failed to delete folder: {result}"
        except Exception as e:
            logger.error(f"Error deleting folder {folder_name}: {e}")
            return False, f"Error deleting folder: {e}"

    async def rename_folder(self, old_name: str, new_name: str) -> tuple[bool, str]:
        """Rename a folder. Returns (success, message)."""
        try:
            async with self._imap_connection(mailbox=None) as imap:
                result = await imap.rename(_quote_mailbox(old_name), _quote_mailbox(new_name))
                if _is_ok(result):
                    logger.info(f"Renamed folder '{old_name}' to '{new_name}'")
                    return True, f"Folder renamed from '{old_name}' to '{new_name}'"
                else:
                    logger.error(f"Failed to rename folder {old_name}: {result}")
                    return False, f"Failed to rename folder: {result}"
        except Exception as e:
            logger.error(f"Error renaming folder {old_name}: {e}")
            return False, f"Error renaming folder: {e}"

    async def list_labels(self) -> list[Label]:
        """List all labels (folders under Labels/ prefix)."""
        folders = await self.list_mailboxes()
        labels = []
        for folder in folders:
            if folder.name.startswith(LABELS_PREFIX):
                # Extract label name without prefix
                label_name = folder.name[len(LABELS_PREFIX):]
                if label_name:  # Skip if just "Labels/" with no name
                    labels.append(
                        Label(
                            name=label_name,
                            full_path=folder.name,
                            delimiter=folder.delimiter,
                            flags=folder.flags,
                        )
                    )
        return labels

    async def get_email_message_id(self, email_id: str, mailbox: str = DEFAULT_MAILBOX) -> str | None:
        """Get the Message-ID header for an email."""
        try:
            async with self._imap_connection(mailbox) as imap:
                _, data = await imap.uid("fetch", email_id, "BODY.PEEK[HEADER.FIELDS (MESSAGE-ID)]")

                for item in data:
                    if isinstance(item, bytearray):
                        header_str = bytes(item).decode("utf-8", errors="replace").strip()
                        if header_str.lower().startswith("message-id:"):
                            return header_str[11:].strip()

                return None
        except Exception as e:
            logger.error(f"Error getting Message-ID for email {email_id}: {e}")
            return None

    async def search_by_message_id(self, message_id: str, mailbox: str) -> str | None:
        """Search for an email by Message-ID in a specific mailbox. Returns email UID or None."""
        try:
            async with self._imap_connection(mailbox) as imap:
                return await self._search_message_id_in_mailbox(imap, message_id)
        except Exception as e:
            logger.debug(f"Error searching for Message-ID in {mailbox}: {e}")
            return None

    @staticmethod
    async def _search_message_id_in_mailbox(
        imap: aioimaplib.IMAP4 | aioimaplib.IMAP4_SSL,
        message_id: str,
    ) -> str | None:
        """Search for an email by Message-ID in the currently selected mailbox. Returns UID or None."""
        _, data = await imap.search(f'HEADER MESSAGE-ID "{message_id}"')

        if data and data[0]:
            seq_nums = data[0].decode("utf-8") if isinstance(data[0], bytes) else str(data[0])
            seq_list = seq_nums.split()
            if seq_list:
                _, fetch_data = await imap.fetch(seq_list[0], "(UID)")
                for item in fetch_data:
                    if isinstance(item, bytes):
                        item_str = item.decode("utf-8", errors="replace")
                        uid_match = re.search(r"UID\s+(\d+)", item_str)
                        if uid_match:
                            return uid_match.group(1)

        return None

    async def mark_emails(
        self,
        email_ids: list[str],
        mark_as: str,
        mailbox: str = DEFAULT_MAILBOX,
        allowed_senders: list[str] | None = None,
        report_blocked_mutations: bool = False,
    ) -> tuple[list[str], list[str]]:
        """Mark emails as read or unread. Returns (marked_ids, failed_ids).

        A blocked sender's UID is never flagged: by default it is reported as a no-op success
        (indistinguishable from a nonexistent UID); when report_blocked_mutations is True it is
        reported in failed_ids instead.
        """
        if mark_as == "read":
            flag_op = "+FLAGS"
        elif mark_as == "unread":
            flag_op = "-FLAGS"
        else:
            raise ValueError(f"Invalid mark_as value: {mark_as}. Must be 'read' or 'unread'.")

        async with self._imap_connection(mailbox) as imap:
            blocked = await self._blocked_uids(imap, email_ids, allowed_senders)
            allowed_ids = [email_id for email_id in email_ids if email_id not in blocked]

            async def _op(imap: Any, email_id: str) -> None:
                await imap.uid("store", email_id, flag_op, IMAP_FLAG_SEEN)

            marked, failed = await self._batch_uid_operation(imap, allowed_ids, _op, f"mark as {mark_as}")

        blocked_ids = [email_id for email_id in email_ids if email_id in blocked]
        if report_blocked_mutations:
            failed.extend(blocked_ids)
        else:
            marked.extend(blocked_ids)
        return marked, failed

    async def search_message_id_in_folders(
        self, message_id: str, folders: list[str]
    ) -> dict[str, str]:
        """Search for an email by Message-ID across multiple folders in a single IMAP session.

        Returns a dict mapping folder name to UID for folders where the email was found.
        """
        results: dict[str, str] = {}
        try:
            async with self._imap_connection(mailbox=None) as imap:
                for folder in folders:
                    try:
                        result = await imap.select(_quote_mailbox(folder))
                        if not _is_ok(result):
                            logger.debug(f"Failed to select folder {folder}: {result}")
                            continue
                        uid = await self._search_message_id_in_mailbox(imap, message_id)
                        if uid:
                            results[folder] = uid
                    except Exception as e:
                        logger.debug(f"Error searching for Message-ID in {folder}: {e}")
        except Exception as e:
            logger.debug(f"Error during multi-folder search: {e}")
        return results

    async def delete_from_folder(self, email_ids: list[str], folder: str) -> tuple[list[str], list[str]]:
        """Delete emails from a specific folder. Returns (deleted_ids, failed_ids)."""
        return await self.delete_emails(email_ids, folder)

    async def move_emails(
        self,
        email_ids: list[str],
        source_mailbox: str,
        destination_mailbox: str,
        allowed_senders: list[str] | None = None,
        report_blocked_mutations: bool = False,
    ) -> tuple[list[str], list[str]]:
        """Move emails to a different mailbox. Uses IMAP MOVE (RFC 6851) with COPY+DELETE fallback.

        A blocked sender's UID is never copied/moved: reported as a no-op success by default,
        or in failed_ids when report_blocked_mutations is True.
        """
        moved_ids = []
        failed_ids = []

        async with self._imap_connection(source_mailbox) as imap:
            capabilities = {str(capability).upper() for capability in getattr(imap, "capabilities", ())}
            has_move = hasattr(imap, "move") and "MOVE" in capabilities

            blocked = await self._blocked_uids(imap, email_ids, allowed_senders)
            copied: list[str] = []  # allowed UIDs that actually completed COPY+STORE on the fallback path
            for email_id in email_ids:
                if email_id in blocked:
                    (failed_ids if report_blocked_mutations else moved_ids).append(email_id)
                    continue
                try:
                    if has_move:
                        move_response = await imap.uid("move", email_id, _quote_mailbox(destination_mailbox))
                        _raise_for_imap_error(move_response, f"MOVE email {email_id}")
                    else:
                        copy_response = await imap.uid("copy", email_id, _quote_mailbox(destination_mailbox))
                        _raise_for_imap_error(copy_response, f"COPY email {email_id}")
                        store_response = await imap.uid("store", email_id, "+FLAGS", r"(\Deleted)")
                        _raise_for_imap_error(store_response, f"STORE \\Deleted for email {email_id}")
                        copied.append(email_id)
                    moved_ids.append(email_id)
                except Exception as e:
                    logger.error(f"Failed to move email {email_id}: {e}")
                    failed_ids.append(email_id)

            if copied:  # only expunge when real COPY+STORE happened — never for blocked/no-op UIDs
                try:
                    expunge_response = await imap.expunge()
                    _raise_for_imap_error(expunge_response, "EXPUNGE moved emails")
                except Exception as e:
                    logger.error(f"Failed to expunge moved emails: {e}")
                    failed_ids.extend(copied)
                    moved_ids = [uid for uid in moved_ids if uid not in set(copied)]

        return moved_ids, failed_ids

    async def list_mailboxes(self, pattern: str = "*", reference: str = "") -> list[MailboxInfo]:
        """List available IMAP mailboxes with flags and delimiter."""
        mailboxes = []

        async with self._imap_connection(mailbox=None) as imap:
            quoted_ref = _quote_mailbox(reference) if reference else '""'
            quoted_pattern = _quote_mailbox(pattern)
            response = await imap.list(quoted_ref, quoted_pattern)
            _raise_for_imap_error(response, f"LIST mailboxes with pattern {pattern}")
            _, data = response

            for item in data:
                mailbox = _parse_list_response(item)
                if mailbox:
                    mailboxes.append(mailbox)

        return mailboxes


class ClassicEmailHandler(EmailHandler):
    def __init__(self, email_settings: EmailSettings):
        self.email_settings = email_settings
        sender = f"{email_settings.full_name} <{email_settings.email_address}>"
        self.incoming_client = EmailClient(email_settings.incoming, sender=sender)
        self.outgoing_client = EmailClient(email_settings.outgoing, sender=sender) if email_settings.outgoing else None
        self.save_to_sent = email_settings.save_to_sent
        self.sent_folder_name = email_settings.sent_folder_name
        self.email_service = _detect_email_service(email_settings)

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
        mailbox: str = DEFAULT_MAILBOX,
        seen: bool | None = None,
        flagged: bool | None = None,
        answered: bool | None = None,
        body: str | None = None,
        text: str | None = None,
        has_attachment: bool | None = None,
    ) -> EmailMetadataPageResponse:
        email_data_list, total = await self.incoming_client.get_emails_metadata_page(
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
            allowed_senders=get_settings().allowed_senders,
        )
        emails = [EmailMetadata.from_email(data) for data in email_data_list]
        return EmailMetadataPageResponse(
            page=page,
            page_size=page_size,
            before=before,
            since=since,
            subject=subject,
            emails=emails,
            total=total,
        )

    async def get_emails_content(
        self,
        email_ids: list[str],
        mailbox: str = "INBOX",
        mark_as_read: bool = False,
        body_offset: int = 0,
        max_body_length: int | None = MAX_BODY_LENGTH,
    ) -> EmailContentBatchResponse:
        """Batch retrieve email body content, honoring the sender allowlist.

        The allowlist is enforced in the read path: get_email_body_by_id checks the From header
        before fetching the body, so a blocked message is never read or marked and returns None —
        indistinguishable from a missing/inaccessible one (both land in failed_ids).
        """
        allowed_senders = get_settings().allowed_senders
        emails = []
        failed_ids = []

        for email_id in email_ids:
            try:
                email_data = await self.incoming_client.get_email_body_by_id(
                    email_id,
                    mailbox,
                    allowed_senders=allowed_senders,
                    body_offset=body_offset,
                    max_body_length=max_body_length,
                )
                if not email_data:
                    failed_ids.append(email_id)
                    continue
                emails.append(
                    EmailBodyResponse(
                        email_id=email_data["email_id"],
                        message_id=email_data.get("message_id"),
                        subject=email_data["subject"],
                        sender=email_data["from"],
                        recipients=email_data["to"],
                        date=email_data["date"],
                        body=email_data["body"],
                        attachments=email_data["attachments"],
                    )
                )
            except Exception as e:
                logger.error(f"Failed to retrieve email {email_id}: {e}")
                failed_ids.append(email_id)

        if mark_as_read and emails:
            successful_ids = [e.email_id for e in emails]
            await self.incoming_client.mark_emails(successful_ids, "read", mailbox)

        return EmailContentBatchResponse(
            emails=emails,
            requested_count=len(email_ids),
            retrieved_count=len(emails),
            failed_ids=failed_ids,
        )

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
    ) -> EmailSendResponse:
        if self.outgoing_client is None:
            raise RuntimeError(f"SMTP is not configured for account '{self.email_settings.account_name}'")

        # Auto-quote the original message when replying
        if in_reply_to and quote_reply:
            original = await self._fetch_original_for_quote(in_reply_to)
            if original and not html:
                # Convert user's body to HTML, append HTML blockquote, send as raw HTML
                user_html = markdown_to_email_html(body, wrap_in_html=False)
                quote_html = _format_quoted_reply_html(original, service=self.email_service)
                body = wrap_html_document(user_html + quote_html)
                html = True  # Already converted, skip conversion in EmailClient
            elif original:
                # Raw HTML path: append HTML blockquote directly
                body += _format_quoted_reply_html(original, service=self.email_service)

        msg = await self.outgoing_client.send_email(
            recipients, subject, body, cc, bcc, html, attachments, in_reply_to, references, reply_to
        )

        # Save to Sent folder if enabled
        if self.save_to_sent and msg:
            # Add BCC header to the saved copy so users can see who was BCC'd.
            # This MUST happen after smtp.send_message() — that ordering is
            # load-bearing for security (BCC must not appear in sent headers).
            if bcc and msg["Bcc"] is None:
                msg["Bcc"] = ", ".join(bcc)
            try:
                await self.outgoing_client.append_to_sent(
                    msg,
                    self.email_settings.incoming,
                    self.sent_folder_name,
                )
            except Exception as e:
                logger.error(f"Failed to save email to Sent folder: {e}", exc_info=True)

        all_recipients = list(recipients)
        if cc:
            all_recipients.extend(cc)
        attachment_info = f" with {len(attachments)} attachment(s)" if attachments else ""
        return EmailSendResponse(
            success=True,
            recipients=all_recipients,
            subject=subject,
            message=f"Email sent successfully to {', '.join(recipients)}{attachment_info}",
        )

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
    ) -> EmailSendResponse:
        # Fetch original email
        original = await self.incoming_client.get_email_body_by_id(email_id, mailbox)
        if not original:
            msg = f"Email with ID {email_id} not found in {mailbox}"
            raise ValueError(msg)

        # Build subject with Fwd: prefix
        original_subject = original.get("subject", "")
        if original_subject.startswith("Fwd:") or original_subject.startswith("FWD:"):
            subject = original_subject
        else:
            subject = f"Fwd: {original_subject}"

        # Build forwarded content
        forwarded_html = _format_forwarded_email_html(original)

        # Build combined body
        if body:
            if html:
                user_html = body
            else:
                user_html = markdown_to_email_html(body, wrap_in_html=False)
            combined_html = wrap_html_document(user_html + forwarded_html)
        else:
            combined_html = wrap_html_document(forwarded_html)

        # Extract original attachments
        original_attachments = await self.incoming_client.extract_attachments(email_id, mailbox)
        extra_parts = []
        for filename, mime_type, data in original_attachments:
            subtype = mime_type.split("/")[1] if "/" in mime_type else "octet-stream"
            part = MIMEApplication(data, _subtype=subtype)
            part.add_header("Content-Disposition", "attachment", filename=filename)
            extra_parts.append(part)

        # Send the forwarded email (no threading headers)
        sent_msg = await self.outgoing_client.send_email(
            recipients,
            subject,
            combined_html,
            cc,
            bcc,
            html=True,
            attachments=attachments,
            extra_parts=extra_parts or None,
        )

        # Save to Sent folder if enabled
        if self.save_to_sent and sent_msg:
            try:
                await self.outgoing_client.append_to_sent(
                    sent_msg,
                    self.email_settings.incoming,
                    self.sent_folder_name,
                )
            except Exception as e:
                logger.error(f"Failed to save forwarded email to Sent folder: {e}", exc_info=True)

        return EmailSendResponse(
            success=True,
            recipients=list(recipients),
            subject=subject,
            message=f"Email forwarded successfully to {', '.join(recipients)}",
        )

    async def _fetch_original_for_quote(self, message_id: str) -> dict[str, Any] | None:
        """Fetch the original email by Message-ID for reply quoting.

        Searches INBOX first, then the Sent folder. Returns the parsed email
        dict or None if not found.
        """
        # Try INBOX first (most common - replying to received email)
        uid = await self.incoming_client.search_by_message_id(message_id, "INBOX")
        mailbox = "INBOX"

        # Try Sent folder if not found in INBOX
        if not uid:
            sent_folder = await self._find_sent_folder()
            if sent_folder:
                uid = await self.incoming_client.search_by_message_id(message_id, sent_folder)
                mailbox = sent_folder

        if not uid:
            logger.debug(f"Original email not found for quoting: {message_id}")
            return None

        original = await self.incoming_client.get_email_body_by_id(uid, mailbox)
        if not original:
            logger.debug(f"Could not fetch original email body for quoting: {message_id}")
            return None

        return original

    # Sentinel indicating a cached negative lookup (folder not found)
    _NOT_FOUND: str = "__not_found__"

    async def _find_special_folder(self, flag: str, fallback_names: list[str]) -> str | None:
        """Find a special folder by RFC 6154 flag, falling back to common names.

        Results are cached on the handler instance to avoid repeated LIST calls.
        """
        cache_key = f"_special_folder_{flag}"
        cached = getattr(self, cache_key, None)
        if cached is not None:
            return None if cached == self._NOT_FOUND else cached

        try:
            folders = await self.incoming_client.list_mailboxes()
            # Check for RFC 6154 flag first
            for folder in folders:
                if any(flag.lower() in f.lower() for f in folder.flags):
                    setattr(self, cache_key, folder.name)
                    return folder.name
            # Fall back to common names
            folder_names = {f.name for f in folders}
            for candidate in fallback_names:
                if candidate in folder_names:
                    setattr(self, cache_key, candidate)
                    return candidate
        except Exception as e:
            logger.debug(f"Error finding special folder with flag {flag}: {e}")

        setattr(self, cache_key, self._NOT_FOUND)
        return None

    async def _find_sent_folder(self) -> str | None:
        """Find the Sent folder name for the account."""
        return await self._find_special_folder("\\Sent", SENT_FOLDER_CANDIDATES)

    async def delete_emails(self, email_ids: list[str], mailbox: str = "INBOX") -> EmailDeleteResponse:
        """Delete emails by moving to Trash (auto-detected), falling back to permanent delete."""
        trash_folder = await self._find_special_folder("\\Trash", TRASH_FOLDER_CANDIDATES)
        already_in_trash = trash_folder and mailbox == trash_folder

        if trash_folder and not already_in_trash:
            # Move to Trash instead of permanent delete
            moved_ids, failed_ids = await self.incoming_client.move_emails(email_ids, mailbox, trash_folder)
            return EmailDeleteResponse(
                success=len(failed_ids) == 0,
                deleted_ids=moved_ids,
                failed_ids=failed_ids,
                mailbox=mailbox,
                destination=trash_folder,
            )

        # Already in Trash or no Trash folder detected -- permanently delete
        if not trash_folder:
            logger.warning("Trash folder not detected, permanently deleting emails")
        deleted_ids, failed_ids = await self.incoming_client.delete_emails(email_ids, mailbox)
        return EmailDeleteResponse(
            success=len(failed_ids) == 0,
            deleted_ids=deleted_ids,
            failed_ids=failed_ids,
            mailbox=mailbox,
        )

    async def archive_emails(self, email_ids: list[str], mailbox: str = "INBOX") -> EmailMoveResponse:
        """Archive emails by moving to the Archive folder (auto-detected via RFC 6154 \\Archive flag)."""
        archive_folder = await self._find_special_folder("\\Archive", ARCHIVE_FOLDER_CANDIDATES)
        if not archive_folder:
            raise ValueError(
                "Archive folder not found. No folder with \\Archive flag or common archive folder names "
                "(Archive, [Gmail]/All Mail) detected."
            )

        moved_ids, failed_ids = await self.incoming_client.move_emails(email_ids, mailbox, archive_folder)
        return EmailMoveResponse(
            success=len(failed_ids) == 0,
            moved_ids=moved_ids,
            failed_ids=failed_ids,
            source_mailbox=mailbox,
            destination_folder=archive_folder,
        )

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
        """Compose and save an email to the specified IMAP mailbox.

        BCC headers are preserved in the saved message so mail clients can
        display BCC recipients (unlike ``send_email``, where BCC is handled
        via the SMTP envelope only).

        Returns:
            A string in the format ``<message-id>|uid:<uid>``.

        Raises:
            ValueError: If any flag in *flags* is invalid per RFC 3501.
            RuntimeError: If the IMAP APPEND operation fails.
        """
        msg = self.incoming_client.compose_message(
            recipients,
            subject,
            body,
            cc,
            bcc,
            html,
            attachments,
            in_reply_to,
            references,
            include_bcc_header=True,
        )

        flags_str = r"(\Draft \Seen)" if flags is None else _validate_flags(flags)

        uid = await self.incoming_client.append_to_mailbox(msg, self.email_settings.incoming, mailbox, flags_str)

        if uid is None:
            raise RuntimeError(f"Failed to save email to mailbox '{mailbox}'")

        message_id = msg["Message-Id"] or "saved"
        return f"{message_id}|uid:{uid}"

    async def move_emails(
        self, email_ids: list[str], source_mailbox: str, destination_mailbox: str
    ) -> tuple[list[str], list[str]]:
        """Move emails between mailboxes. Returns (moved_ids, failed_ids)."""
        settings = get_settings()
        return await self.incoming_client.move_emails(
            email_ids,
            source_mailbox,
            destination_mailbox,
            allowed_senders=settings.allowed_senders,
            report_blocked_mutations=settings.report_blocked_mutations,
        )

    async def list_mailboxes(self, pattern: str = "*", reference: str = "") -> list[MailboxInfo]:
        """List available mailboxes with flags and delimiter."""
        return await self.incoming_client.list_mailboxes(pattern, reference)

    async def download_attachment(
        self,
        email_id: str,
        attachment_name: str,
        save_path: str,
        mailbox: str = "INBOX",
    ) -> AttachmentDownloadResponse:
        """Download an email attachment and save it to the specified path.

        Args:
            email_id: The UID of the email containing the attachment.
            attachment_name: The filename of the attachment to download.
            save_path: The local path where the attachment will be saved.
            mailbox: The mailbox to search in (default: "INBOX").

        Returns:
            AttachmentDownloadResponse with download result information.
        """
        allowed_senders = get_settings().allowed_senders
        result = await self.incoming_client.download_attachment(
            email_id, attachment_name, save_path, mailbox, allowed_senders=allowed_senders
        )
        return AttachmentDownloadResponse(
            email_id=result["email_id"],
            attachment_name=result["attachment_name"],
            mime_type=result["mime_type"],
            size=result["size"],
            saved_path=result["saved_path"],
        )

    async def copy_emails(
        self,
        email_ids: list[str],
        destination_folder: str,
        source_mailbox: str = "INBOX",
    ) -> EmailMoveResponse:
        """Copy emails to a destination folder (preserves original)."""
        copied_ids, failed_ids = await self.incoming_client.copy_emails(email_ids, destination_folder, source_mailbox)
        return EmailMoveResponse(
            success=len(failed_ids) == 0,
            moved_ids=copied_ids,
            failed_ids=failed_ids,
            source_mailbox=source_mailbox,
            destination_folder=destination_folder,
        )

    async def create_folder(self, folder_name: str) -> FolderOperationResponse:
        """Create a new folder/mailbox."""
        success, message = await self.incoming_client.create_folder(folder_name)
        return FolderOperationResponse(
            success=success,
            folder_name=folder_name,
            message=message,
        )

    async def delete_folder(self, folder_name: str) -> FolderOperationResponse:
        """Delete a folder/mailbox."""
        success, message = await self.incoming_client.delete_folder(folder_name)
        return FolderOperationResponse(
            success=success,
            folder_name=folder_name,
            message=message,
        )

    async def rename_folder(self, old_name: str, new_name: str) -> FolderOperationResponse:
        """Rename a folder/mailbox."""
        success, message = await self.incoming_client.rename_folder(old_name, new_name)
        return FolderOperationResponse(
            success=success,
            folder_name=new_name,
            message=message,
        )

    async def list_labels(self) -> LabelListResponse:
        """List all labels (ProtonMail: folders under Labels/ prefix)."""
        labels = await self.incoming_client.list_labels()
        return LabelListResponse(labels=labels, total=len(labels))

    async def apply_label(
        self,
        email_ids: list[str],
        label_name: str,
        source_mailbox: str = "INBOX",
    ) -> EmailMoveResponse:
        """Apply a label to emails by copying to the label folder."""
        label_folder = f"{LABELS_PREFIX}{label_name}"
        copied_ids, failed_ids = await self.incoming_client.copy_emails(email_ids, label_folder, source_mailbox)
        return EmailMoveResponse(
            success=len(failed_ids) == 0,
            moved_ids=copied_ids,
            failed_ids=failed_ids,
            source_mailbox=source_mailbox,
            destination_folder=label_folder,
        )

    async def remove_label(
        self,
        email_ids: list[str],
        label_name: str,
        source_mailbox: str = DEFAULT_MAILBOX,
    ) -> EmailMoveResponse:
        """Remove a label from emails by deleting from the label folder.

        This finds the emails in the label folder by their Message-ID and deletes them.
        The original emails in other folders are preserved.
        """
        label_folder = f"{LABELS_PREFIX}{label_name}"
        removed_ids = []
        failed_ids = []

        for email_id in email_ids:
            message_id = await self.incoming_client.get_email_message_id(email_id, source_mailbox)
            if not message_id:
                logger.warning(f"Could not get Message-ID for email {email_id}")
                failed_ids.append(email_id)
                continue

            # Find the email in the label folder
            label_uid = await self.incoming_client.search_by_message_id(message_id, label_folder)
            if not label_uid:
                logger.warning(f"Email {email_id} not found in label {label_name}")
                failed_ids.append(email_id)
                continue

            # Delete from label folder
            deleted, _failed = await self.incoming_client.delete_from_folder([label_uid], label_folder)
            if deleted:
                removed_ids.append(email_id)
            else:
                failed_ids.append(email_id)

        return EmailMoveResponse(
            success=len(failed_ids) == 0,
            moved_ids=removed_ids,
            failed_ids=failed_ids,
            source_mailbox=label_folder,
            destination_folder="",
        )

    async def get_email_labels(
        self,
        email_id: str,
        source_mailbox: str = "INBOX",
    ) -> EmailLabelsResponse:
        """Get all labels applied to a specific email."""
        # Get Message-ID from the source email
        message_id = await self.incoming_client.get_email_message_id(email_id, source_mailbox)
        if not message_id:
            return EmailLabelsResponse(email_id=email_id, labels=[])

        # Get all labels and search in a single session
        labels = await self.incoming_client.list_labels()
        if not labels:
            return EmailLabelsResponse(email_id=email_id, labels=[])

        folder_paths = [label.full_path for label in labels]
        found_folders = await self.incoming_client.search_message_id_in_folders(message_id, folder_paths)

        applied_labels = [label.name for label in labels if label.full_path in found_folders]
        return EmailLabelsResponse(email_id=email_id, labels=applied_labels)

    async def create_label(self, label_name: str) -> FolderOperationResponse:
        """Create a new label (creates Labels/name folder)."""
        label_folder = f"{LABELS_PREFIX}{label_name}"
        success, message = await self.incoming_client.create_folder(label_folder)
        return FolderOperationResponse(
            success=success,
            folder_name=label_name,
            message=message.replace(label_folder, label_name) if success else message,
        )

    async def delete_label(self, label_name: str) -> FolderOperationResponse:
        """Delete a label (deletes Labels/name folder)."""
        label_folder = f"{LABELS_PREFIX}{label_name}"
        success, message = await self.incoming_client.delete_folder(label_folder)
        return FolderOperationResponse(
            success=success,
            folder_name=label_name,
            message=message.replace(label_folder, label_name) if success else message,
        )

    async def mark_emails(
        self,
        email_ids: list[str],
        mark_as: str,
        mailbox: str = "INBOX",
    ) -> EmailMarkResponse:
        """Mark emails as read or unread."""
        settings = get_settings()
        marked_ids, failed_ids = await self.incoming_client.mark_emails(
            email_ids,
            mark_as,
            mailbox,
            allowed_senders=settings.allowed_senders,
            report_blocked_mutations=settings.report_blocked_mutations,
        )
        return EmailMarkResponse(
            success=len(failed_ids) == 0,
            marked_ids=marked_ids,
            failed_ids=failed_ids,
            mailbox=mailbox,
            marked_as=mark_as,
        )
