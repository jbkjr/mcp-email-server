"""Tests for the default attachment download directory (upstream PR #224 port).

Covers destination resolution, filename sanitization/traversal rejection, the
default directory being created on demand, explicit paths staying unchanged, and
the sender allowlist still blocking a disallowed sender when defaulting.
"""

import asyncio
import os
import re
import stat
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from mcp_email_server.config import EmailServer
from mcp_email_server.emails import attachment_paths
from mcp_email_server.emails.attachment_paths import (
    DEFAULT_ATTACHMENT_DIRECTORY,
    default_downloads_directory,
    prepare_private_directory,
    resolve_attachment_destination,
    safe_default_filename,
    write_private_file,
)
from mcp_email_server.emails.classic import EmailClient

FIXED_HEX = "0123456789abcdef0123456789abcdef"


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


@pytest.fixture
def fixed_hex(monkeypatch):
    """Make the randomized filename suffix deterministic."""
    monkeypatch.setattr(attachment_paths.secrets, "token_hex", lambda _length: FIXED_HEX)
    return FIXED_HEX


@pytest.fixture
def fake_downloads(monkeypatch, tmp_path):
    """Point the default download area at a temporary directory."""
    downloads = tmp_path / "Downloads"
    monkeypatch.setattr(attachment_paths, "default_downloads_directory", lambda: downloads)
    return downloads


def _build_email_with_attachment(filename: str, payload: bytes = b"%PDF-1.4 bytes") -> bytes:
    """Build a multipart/mixed email whose attachment part uses ``filename``."""
    msg = MIMEMultipart("mixed")
    msg["Subject"] = "Report"
    msg["From"] = "sender@example.com"
    msg["To"] = "recipient@example.com"
    msg["Date"] = "Fri, 8 May 2026 19:17:09 +0200"
    msg.attach(MIMEText("Please see attached report.", "plain", "utf-8"))
    pdf_part = MIMEApplication(payload, _subtype="pdf")
    pdf_part.add_header("Content-Disposition", "attachment", filename=filename)
    msg.attach(pdf_part)
    return msg.as_bytes()


def _mock_imap():
    mock_imap = AsyncMock()
    mock_imap._client_task = asyncio.Future()
    mock_imap._client_task.set_result(None)
    mock_imap.wait_hello_from_server = AsyncMock()
    mock_imap.login = AsyncMock(return_value=MagicMock(result="OK", lines=[]))
    mock_imap.select = AsyncMock(return_value=("OK", [b"1"]))
    mock_imap.logout = AsyncMock()
    return mock_imap


def _fetch_returning(raw_email: bytes):
    async def _fake_fetch(_imap, _email_id):
        return [b"1 FETCH (BODY[] {%d}" % len(raw_email), bytearray(raw_email), b")"]

    return _fake_fetch


class TestSafeDefaultFilename:
    """The generated basename must be sanitized, bounded, and randomized."""

    def test_bounded_and_preserves_short_extension(self, fixed_hex):
        filename = safe_default_filename(f"{'é' * 2000}.pdf")

        assert len(filename.encode("utf-8")) <= 217
        assert filename.endswith(f"-{FIXED_HEX}.pdf")

    def test_empty_after_sanitizing_falls_back_to_attachment(self, fixed_hex):
        assert safe_default_filename("...") == f"attachment-{FIXED_HEX}"
        assert safe_default_filename("") == f"attachment-{FIXED_HEX}"
        assert safe_default_filename("   ") == f"attachment-{FIXED_HEX}"

    def test_overlong_extension_is_not_treated_as_a_suffix(self, fixed_hex):
        long_suffix = "x" * 25

        assert safe_default_filename(f"report.{long_suffix}") == f"report.{long_suffix}-{FIXED_HEX}"

    def test_removes_unicode_format_controls(self, fixed_hex):
        """An RTL-override extension spoof cannot survive into the filename."""
        filename = safe_default_filename("invoice‮fdp.exe")

        assert "‮" not in filename
        assert filename == f"invoice_fdp-{FIXED_HEX}.exe"

    @pytest.mark.parametrize(
        "attachment_name",
        [
            "../../../etc/passwd",
            "..\\..\\windows\\system32\\config",
            "/etc/shadow",
            "sub/dir/evil.sh",
            "nul\x00byte.txt",
            "tab\tseparated.txt",
        ],
    )
    def test_traversal_and_separator_syntax_is_stripped(self, attachment_name, fixed_hex):
        filename = safe_default_filename(attachment_name)

        assert "/" not in filename
        assert "\\" not in filename
        assert "\x00" not in filename
        assert not filename.startswith(".")
        # Nothing in the result can re-introduce a path component.
        assert Path(filename).name == filename

    def test_suffix_is_random_per_call(self):
        first = safe_default_filename("document.pdf")
        second = safe_default_filename("document.pdf")

        assert first != second
        assert first.startswith("document-")
        assert first.endswith(".pdf")


class TestResolveAttachmentDestination:
    """Explicit paths are honored exactly; omitted paths land in the default area."""

    def test_explicit_absolute_path_is_unchanged(self, tmp_path):
        explicit = tmp_path / "downloads" / "report.pdf"

        destination, used_default = resolve_attachment_destination(str(explicit), "whatever.pdf")

        assert destination == explicit
        assert used_default is False

    def test_explicit_relative_path_resolves_against_cwd(self):
        destination, used_default = resolve_attachment_destination("relative/report.pdf", "report.pdf")

        assert destination == Path(os.path.abspath("relative/report.pdf"))
        assert used_default is False

    def test_explicit_tilde_path_is_expanded(self):
        destination, used_default = resolve_attachment_destination("~/report.pdf", "report.pdf")

        assert destination == Path.home() / "report.pdf"
        assert used_default is False

    def test_explicit_traversal_path_is_not_rewritten(self, tmp_path):
        """An explicit path keeps its exact prior meaning, traversal included."""
        explicit = tmp_path / "a" / ".." / "b.pdf"

        destination, used_default = resolve_attachment_destination(str(explicit), "b.pdf")

        assert destination == tmp_path / "b.pdf"
        assert used_default is False

    def test_omitted_path_uses_default_directory(self, fake_downloads, fixed_hex):
        destination, used_default = resolve_attachment_destination(None, "report.pdf")

        assert used_default is True
        assert destination == fake_downloads / DEFAULT_ATTACHMENT_DIRECTORY / f"report-{FIXED_HEX}.pdf"

    def test_omitted_path_confines_a_traversing_attachment_name(self, fake_downloads, fixed_hex):
        destination, used_default = resolve_attachment_destination(None, "../CON?.pdf")

        assert used_default is True
        assert destination.parent == fake_downloads / DEFAULT_ATTACHMENT_DIRECTORY
        assert destination.name == f"_CON_-{FIXED_HEX}.pdf"

    def test_default_downloads_directory_is_under_home(self):
        assert default_downloads_directory() == Path.home() / "Downloads"

    def test_resolution_failure_is_reported_as_permission_error(self, monkeypatch):
        def fail_resolution():
            raise RuntimeError("home directory unavailable")

        monkeypatch.setattr(attachment_paths, "default_downloads_directory", fail_resolution)

        with pytest.raises(PermissionError, match="could not be resolved"):
            resolve_attachment_destination(None, "document.pdf")


@pytest.mark.skipif(os.name != "posix", reason="POSIX permission model")
class TestPrivateDirectoryAndFile:
    """The default area is created owner-only and rejects unsafe reuse."""

    def test_directory_is_created_owner_only(self, tmp_path):
        target = tmp_path / "Downloads" / DEFAULT_ATTACHMENT_DIRECTORY

        prepare_private_directory(target)

        assert target.is_dir()
        assert stat.S_IMODE(target.stat().st_mode) & 0o077 == 0

    def test_existing_private_directory_is_reused(self, tmp_path):
        target = tmp_path / DEFAULT_ATTACHMENT_DIRECTORY
        target.mkdir(mode=0o700)

        prepare_private_directory(target)

        assert target.is_dir()

    def test_symlinked_directory_is_rejected(self, tmp_path):
        elsewhere = tmp_path / "elsewhere"
        elsewhere.mkdir(mode=0o700)
        target = tmp_path / DEFAULT_ATTACHMENT_DIRECTORY
        target.symlink_to(elsewhere, target_is_directory=True)

        with pytest.raises(PermissionError, match="unsafe"):
            prepare_private_directory(target)

    def test_world_writable_directory_is_rejected(self, tmp_path):
        target = tmp_path / DEFAULT_ATTACHMENT_DIRECTORY
        target.mkdir()
        os.chmod(target, 0o777)  # noqa: S103 - deliberately unsafe fixture

        with pytest.raises(PermissionError, match="permissions are unsafe"):
            prepare_private_directory(target)

    def test_uncreatable_directory_is_reported_as_permission_error(self, tmp_path):
        blocker = tmp_path / "blocker"
        blocker.write_bytes(b"not a directory")

        with pytest.raises(PermissionError, match="could not be created"):
            prepare_private_directory(blocker / "child")

    def test_file_is_written_owner_only(self, tmp_path):
        destination = tmp_path / "payload.bin"

        write_private_file(destination, b"secret")

        assert destination.read_bytes() == b"secret"
        assert stat.S_IMODE(destination.stat().st_mode) & 0o077 == 0

    def test_preplanted_symlink_target_is_rejected(self, tmp_path):
        victim = tmp_path / "victim.txt"
        victim.write_bytes(b"original")
        destination = tmp_path / "payload.bin"
        destination.symlink_to(victim)

        with pytest.raises(PermissionError, match="unsafe"):
            write_private_file(destination, b"secret")

        assert victim.read_bytes() == b"original"


class TestDownloadAttachmentDefaultPath:
    """End-to-end ``EmailClient.download_attachment`` behavior."""

    @pytest.mark.asyncio
    async def test_omitted_save_path_writes_into_default_directory(self, email_client, fake_downloads, fixed_hex):
        raw_email = _build_email_with_attachment("report.pdf", b"%PDF default area")
        expected = fake_downloads / DEFAULT_ATTACHMENT_DIRECTORY / f"report-{FIXED_HEX}.pdf"
        assert not fake_downloads.exists()

        with (
            patch.object(email_client, "_fetch_email_with_formats", side_effect=_fetch_returning(raw_email)),
            patch.object(email_client, "imap_class", return_value=_mock_imap()),
        ):
            result = await email_client.download_attachment(email_id="1", attachment_name="report.pdf")

        assert expected.read_bytes() == b"%PDF default area"
        assert Path(result["saved_path"]) == expected.resolve()
        assert result["attachment_name"] == "report.pdf"
        assert result["mime_type"] == "application/pdf"
        if os.name == "posix":
            # Neither the private directory nor the saved file is group/other readable.
            assert stat.S_IMODE(expected.parent.stat().st_mode) & 0o077 == 0
            assert stat.S_IMODE(expected.stat().st_mode) & 0o077 == 0

    @pytest.mark.asyncio
    async def test_traversing_attachment_name_cannot_escape_default_directory(
        self, email_client, fake_downloads, tmp_path, fixed_hex
    ):
        malicious = "../../../evil.pdf"
        raw_email = _build_email_with_attachment(malicious, b"payload")

        with (
            patch.object(email_client, "_fetch_email_with_formats", side_effect=_fetch_returning(raw_email)),
            patch.object(email_client, "imap_class", return_value=_mock_imap()),
        ):
            result = await email_client.download_attachment(email_id="1", attachment_name=malicious)

        saved = Path(result["saved_path"])
        assert saved.parent == (fake_downloads / DEFAULT_ATTACHMENT_DIRECTORY).resolve()
        assert saved.name.endswith(f"-{FIXED_HEX}.pdf")
        assert "/" not in saved.name and not saved.name.startswith(".")
        assert not (tmp_path / "evil.pdf").exists()
        assert not (tmp_path.parent / "evil.pdf").exists()

    @pytest.mark.asyncio
    async def test_explicit_save_path_is_honored_unchanged(self, email_client, fake_downloads, tmp_path):
        raw_email = _build_email_with_attachment("report.pdf", b"%PDF explicit")
        explicit = tmp_path / "nested" / "chosen-name.pdf"

        with (
            patch.object(email_client, "_fetch_email_with_formats", side_effect=_fetch_returning(raw_email)),
            patch.object(email_client, "imap_class", return_value=_mock_imap()),
        ):
            result = await email_client.download_attachment(
                email_id="1",
                attachment_name="report.pdf",
                save_path=str(explicit),
            )

        assert explicit.read_bytes() == b"%PDF explicit"
        assert Path(result["saved_path"]) == explicit.resolve()
        # The default area is never touched when an explicit path is supplied.
        assert not fake_downloads.exists()

    @pytest.mark.asyncio
    async def test_explicit_save_path_overwrites_as_before(self, email_client, tmp_path):
        """Explicit-path writes keep their pre-existing overwrite semantics."""
        raw_email = _build_email_with_attachment("report.pdf", b"new bytes")
        explicit = tmp_path / "report.pdf"
        explicit.write_bytes(b"old bytes")

        with (
            patch.object(email_client, "_fetch_email_with_formats", side_effect=_fetch_returning(raw_email)),
            patch.object(email_client, "imap_class", return_value=_mock_imap()),
        ):
            await email_client.download_attachment(
                email_id="1",
                attachment_name="report.pdf",
                save_path=str(explicit),
            )

        assert explicit.read_bytes() == b"new bytes"

    @pytest.mark.asyncio
    async def test_default_path_still_blocks_disallowed_sender(self, email_client, fake_downloads):
        """The sender allowlist is enforced before any body fetch, path or no path."""
        with (
            patch.object(
                email_client, "_batch_fetch_senders", AsyncMock(return_value={"1": "evil@blocked.com"})
            ) as mock_senders,
            patch.object(email_client, "_fetch_email_with_formats", AsyncMock()) as mock_fetch,
            patch.object(email_client, "imap_class", return_value=_mock_imap()),
        ):
            with pytest.raises(ValueError, match=re.escape("Failed to fetch email with UID 1")):
                await email_client.download_attachment(
                    email_id="1",
                    attachment_name="report.pdf",
                    allowed_senders=["*@allowed.com"],
                )

        mock_senders.assert_awaited_once()
        mock_fetch.assert_not_called()
        assert not fake_downloads.exists()

    @pytest.mark.asyncio
    async def test_default_path_allows_permitted_sender(self, email_client, fake_downloads, fixed_hex):
        raw_email = _build_email_with_attachment("report.pdf", b"%PDF allowed")

        with (
            patch.object(email_client, "_batch_fetch_senders", AsyncMock(return_value={"1": "ok@allowed.com"})),
            patch.object(email_client, "_fetch_email_with_formats", side_effect=_fetch_returning(raw_email)),
            patch.object(email_client, "imap_class", return_value=_mock_imap()),
        ):
            result = await email_client.download_attachment(
                email_id="1",
                attachment_name="report.pdf",
                allowed_senders=["*@allowed.com"],
            )

        expected = fake_downloads / DEFAULT_ATTACHMENT_DIRECTORY / f"report-{FIXED_HEX}.pdf"
        assert Path(result["saved_path"]) == expected.resolve()
        assert expected.read_bytes() == b"%PDF allowed"

    @pytest.mark.asyncio
    async def test_unresolvable_default_destination_fails_before_connecting(self, email_client, monkeypatch):
        """A destination that cannot be resolved is rejected before any IMAP work."""

        def fail_resolution():
            raise OSError("home directory unavailable")

        monkeypatch.setattr(attachment_paths, "default_downloads_directory", fail_resolution)
        connect = MagicMock(return_value=_mock_imap())

        with patch.object(email_client, "imap_class", connect):
            with pytest.raises(PermissionError, match="could not be resolved"):
                await email_client.download_attachment(email_id="1", attachment_name="report.pdf")

        connect.assert_not_called()
