from __future__ import annotations

import contextlib
import imaplib
import importlib.metadata
import os
import re
import smtplib
import sqlite3
import subprocess
import sys
import time
import uuid
from collections.abc import Iterator
from dataclasses import dataclass
from datetime import timedelta
from email import policy
from email.message import EmailMessage, Message
from email.parser import BytesParser
from email.utils import make_msgid
from pathlib import Path
from typing import Any

import pytest
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from mcp.types import TextContent

from mcp_email_server.bootstrap import read_bootstrap
from mcp_email_server.managed import SCHEMA_VERSION

pytestmark = pytest.mark.e2e

SMTP_HOST = "127.0.0.1"
SMTP_PORT = int(os.environ.get("MCP_EMAIL_SERVER_E2E_SMTP_PORT", "3025"))
IMAP_HOST = "127.0.0.1"
IMAP_PORT = int(os.environ.get("MCP_EMAIL_SERVER_E2E_IMAP_PORT", "3143"))
ALICE = ("alice@example.test", "alice-password")
BOB = ("bob@example.test", "bob-password")

CONFIG_TEMPLATE = f"""credential_storage = "plaintext"
enable_attachment_download = true
allowed_recipients = ["bob@example.test"]

[[emails]]
account_name = "alice"
full_name = "alice@example.test"
email_address = "alice@example.test"
save_to_sent = true
sent_folder_name = "Sent"

[emails.incoming]
user_name = "alice@example.test"
password = "alice-password"
host = "127.0.0.1"
port = {IMAP_PORT}
use_ssl = false
start_ssl = false
verify_ssl = true

[emails.outgoing]
user_name = "alice@example.test"
password = "alice-password"
host = "127.0.0.1"
port = {SMTP_PORT}
use_ssl = false
start_ssl = false
verify_ssl = true

[[emails]]
account_name = "bob"
full_name = "Bob Example"
email_address = "bob@example.test"
save_to_sent = false

[emails.incoming]
user_name = "bob@example.test"
password = "bob-password"
host = "127.0.0.1"
port = {IMAP_PORT}
use_ssl = false
start_ssl = false
verify_ssl = true
"""


@dataclass(frozen=True)
class ObservedMessage:
    uid: str
    message: Message
    flags: set[str]


@contextlib.contextmanager
def _imap_session(credentials: tuple[str, str]) -> Iterator[imaplib.IMAP4]:
    client = imaplib.IMAP4(IMAP_HOST, IMAP_PORT, timeout=5)
    try:
        status, _ = client.login(*credentials)
        assert status == "OK"
        yield client
    finally:
        with contextlib.suppress(Exception):
            client.logout()


def _wait_until_ready(timeout: float = 15) -> None:
    deadline = time.monotonic() + timeout
    last_error: Exception | None = None
    while time.monotonic() < deadline:
        try:
            with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=2) as smtp:
                smtp.login(*ALICE)
            with _imap_session(BOB):
                pass
            return
        except (OSError, smtplib.SMTPException, imaplib.IMAP4.error, AssertionError) as exc:
            last_error = exc
            time.sleep(0.25)
    pytest.fail(f"GreenMail is not ready on SMTP {SMTP_PORT}/IMAP {IMAP_PORT}: {last_error}")


def _ensure_empty_mailboxes(credentials: tuple[str, str], mailboxes: list[str]) -> None:
    with _imap_session(credentials) as client:
        for mailbox in mailboxes:
            if mailbox != "INBOX":
                status, _ = client.create(mailbox)
                assert status in {"OK", "NO"}
            status, _ = client.select(mailbox)
            assert status == "OK"
            status, data = client.uid("search", None, "ALL")
            assert status == "OK"
            for uid in (data[0] or b"").split():
                status, _ = client.uid("store", uid, "+FLAGS.SILENT", r"(\Deleted)")
                assert status == "OK"
            status, _ = client.expunge()
            assert status == "OK"


def _message_count(credentials: tuple[str, str], mailbox: str) -> int:
    with _imap_session(credentials) as client:
        status, data = client.select(mailbox, readonly=True)
        assert status == "OK"
        return int(data[0])


def _find_message(credentials: tuple[str, str], mailbox: str, subject: str) -> ObservedMessage | None:
    with _imap_session(credentials) as client:
        status, _ = client.select(mailbox, readonly=True)
        assert status == "OK"
        status, data = client.uid("search", None, "ALL")
        assert status == "OK"
        for uid in reversed((data[0] or b"").split()):
            status, fetched = client.uid("fetch", uid, "(BODY.PEEK[] FLAGS)")
            assert status == "OK"
            response = next((item for item in fetched if isinstance(item, tuple)), None)
            assert response is not None
            metadata, raw_message = response
            message = BytesParser(policy=policy.default).parsebytes(raw_message)
            if str(message.get("Subject", "")) != subject:
                continue
            flag_match = re.search(rb"FLAGS \(([^)]*)\)", metadata)
            flags = set(flag_match.group(1).decode().split()) if flag_match else set()
            return ObservedMessage(uid.decode(), message, flags)
    return None


def _wait_for_message(credentials: tuple[str, str], mailbox: str, subject: str, timeout: float = 5) -> ObservedMessage:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        observed = _find_message(credentials, mailbox, subject)
        if observed is not None:
            return observed
        time.sleep(0.1)
    pytest.fail(f"Message {subject!r} did not arrive in {mailbox!r}")


def _seed_message_as(
    sender: tuple[str, str],
    recipient: str,
    subject: str,
    body: str,
) -> None:
    message = EmailMessage()
    message["From"] = sender[0]
    message["To"] = recipient
    message["Subject"] = subject
    message["Message-ID"] = make_msgid(domain="example.test")
    message.set_content(body)
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=5) as smtp:
        smtp.login(*sender)
        smtp.send_message(message)


def _seed_message_with_attachment_as(
    sender: tuple[str, str],
    recipient: str,
    subject: str,
    body: str,
    *,
    filename: str,
    payload: bytes,
    maintype: str,
    subtype: str,
) -> None:
    """Plant a real multipart source message so a forward has parts to carry."""
    message = EmailMessage()
    message["From"] = sender[0]
    message["To"] = recipient
    message["Subject"] = subject
    message["Message-ID"] = make_msgid(domain="example.test")
    message.set_content(body)
    message.add_attachment(payload, maintype=maintype, subtype=subtype, filename=filename)
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=5) as smtp:
        smtp.login(*sender)
        smtp.send_message(message)


def _seed_message(subject: str, body: str) -> None:
    _seed_message_as(ALICE, BOB[0], subject, body)


def _mark_deleted_without_expunge(credentials: tuple[str, str], mailbox: str, uid: str) -> None:
    """Simulate another client leaving an unrelated message pending deletion."""
    with _imap_session(credentials) as client:
        status, _ = client.select(mailbox)
        assert status == "OK"
        status, _ = client.uid("store", uid, "+FLAGS.SILENT", r"(\Deleted)")
        assert status == "OK"


def _add_flags(credentials: tuple[str, str], mailbox: str, uid: str, flags: str) -> None:
    with _imap_session(credentials) as client:
        status, _ = client.select(mailbox)
        assert status == "OK"
        status, _ = client.uid("store", uid, "+FLAGS.SILENT", f"({flags})")
        assert status == "OK"


def _text_content(result: Any) -> str:
    return "\n".join(item.text for item in result.content if isinstance(item, TextContent))


async def _call_tool(session: ClientSession, name: str, arguments: dict[str, Any]) -> dict[str, Any]:
    result = await session.call_tool(name, arguments=arguments)
    assert result.isError is not True, f"{name} failed: {_text_content(result)}"
    assert result.structuredContent is not None, f"{name} returned no structured content"
    return result.structuredContent


async def _metadata_for_subject_in_mailbox(
    session: ClientSession, account_name: str, mailbox: str, subject: str
) -> dict[str, Any]:
    payload = await _call_tool(
        session,
        "list_emails_metadata",
        {"account_name": account_name, "mailbox": mailbox, "subject": subject, "page_size": 50},
    )
    matches = [email for email in payload["emails"] if email["subject"] == subject]
    assert len(matches) == 1, payload
    return matches[0]


async def _metadata_for_subject(session: ClientSession, account_name: str, subject: str) -> dict[str, Any]:
    return await _metadata_for_subject_in_mailbox(session, account_name, "INBOX", subject)


def _run_cli(console_script: Path, env: dict[str, str], arguments: list[str], *, stdin: str | None = None) -> str:
    completed = subprocess.run(  # noqa: S603 - fixed installed script with test-owned arguments
        [str(console_script), *arguments],
        cwd=Path.cwd(),
        env=env,
        input=stdin,
        capture_output=True,
        text=True,
        timeout=15,
        check=False,
    )
    assert completed.returncode == 0, completed.stderr or completed.stdout
    return completed.stdout


@pytest.mark.asyncio
async def test_managed_cli_setup_restart_and_stdio_list_mailboxes_against_greenmail(tmp_path: Path) -> None:
    """Prove CLI setup -> test -> restart -> live managed IMAP without catalog activation."""
    _wait_until_ready()
    _ensure_empty_mailboxes(ALICE, ["INBOX", "Drafts", "Archive"])
    subject = f"managed-index-{uuid.uuid4().hex}"
    _seed_message_as(BOB, ALICE[0], subject, "Managed indexed metadata")
    _wait_for_message(ALICE, "INBOX", subject)
    app_dir = tmp_path / "managed-app"
    app_dir.mkdir(mode=0o700)
    app_dir.chmod(0o700)
    config_path = app_dir / "config.toml"
    database = app_dir / "catalog.sqlite3"
    keyring_path = app_dir / "e2e-keyring.sqlite3"
    console_script = Path(sys.executable).with_name("mcp-email-server")
    assert console_script.is_file()
    server_env = {key: value for key, value in os.environ.items() if not key.startswith("MCP_EMAIL_SERVER_")}
    server_env.update({
        "MCP_EMAIL_SERVER_CONFIG_PATH": str(config_path),
        "MCP_EMAIL_SERVER_E2E_KEYRING_PATH": str(keyring_path),
        "MCP_EMAIL_SERVER_LOG_LEVEL": "WARNING",
        "PYTHON_KEYRING_BACKEND": "dev.greenmail.file_keyring.FileKeyring",
        "PYTHONPATH": str(Path.cwd()),
    })

    _run_cli(console_script, server_env, ["config", "init", "--database", str(database)])
    bootstrap = read_bootstrap(config_path)
    assert bootstrap.mode == "managed"
    assert bootstrap.db_path == database
    assert not config_path.exists()
    add_arguments = [
        "account",
        "add",
        "alice-managed",
        "--email",
        ALICE[0],
        "--full-name",
        "Alice Managed",
        "--imap-host",
        IMAP_HOST,
        "--imap-port",
        str(IMAP_PORT),
        "--imap-user",
        ALICE[0],
        "--no-imap-ssl",
        "--smtp-host",
        SMTP_HOST,
        "--smtp-port",
        str(SMTP_PORT),
        "--smtp-user",
        ALICE[0],
        "--no-smtp-ssl",
        "--password-stdin",
    ]
    assert ALICE[1] not in add_arguments
    add_output = _run_cli(console_script, server_env, add_arguments, stdin=f"{ALICE[1]}\n{ALICE[1]}\n")
    assert ALICE[1] not in add_output
    test_output = _run_cli(console_script, server_env, ["account", "test", "alice-managed"])
    assert "connectivity test passed" in test_output
    counts_before = (_message_count(ALICE, "INBOX"), _message_count(BOB, "INBOX"))
    outgoing_test_output = _run_cli(
        console_script,
        server_env,
        ["account", "test", "alice-managed", "outgoing"],
    )
    assert "Outgoing connectivity test passed" in outgoing_test_output
    assert (_message_count(ALICE, "INBOX"), _message_count(BOB, "INBOX")) == counts_before
    assert read_bootstrap(config_path).mode == "managed"
    assert not config_path.exists()

    server = StdioServerParameters(
        command=str(console_script),
        args=["stdio"],
        env=server_env,
        cwd=Path.cwd(),
    )
    async with stdio_client(server) as (read_stream, write_stream):
        async with ClientSession(
            read_stream,
            write_stream,
            read_timeout_seconds=timedelta(seconds=15),
        ) as session:
            await session.initialize()
            accounts = await _call_tool(session, "list_available_accounts", {})
            assert [account["account_name"] for account in accounts["result"]] == ["alice-managed"]
            assert ALICE[1] not in str(accounts)
            mailboxes = await _call_tool(session, "list_mailboxes", {"account_name": "alice-managed"})
            assert "INBOX" in {mailbox["name"] for mailbox in mailboxes["result"]}
            metadata = await _call_tool(
                session,
                "list_emails_metadata",
                {"account_name": "alice-managed", "page_size": 10},
            )
            assert metadata["total"] == 1
            assert metadata["emails"][0]["subject"] == subject
            with contextlib.closing(sqlite3.connect(database)) as connection:
                assert connection.execute("SELECT version FROM schema_metadata").fetchone()[0] == SCHEMA_VERSION
                assert connection.execute("SELECT completeness FROM index_coverage").fetchone()[0] == "COMPLETE"

            managed_uid = metadata["emails"][0]["email_id"]
            mark = await _call_tool(
                session,
                "mark_emails_as_read",
                {"account_name": "alice-managed", "email_ids": [managed_uid]},
            )
            assert mark["result"] == "Successfully marked 1 email(s) as read"
            assert r"\Seen" in _wait_for_message(ALICE, "INBOX", subject).flags
            unset_seen = await _call_tool(
                session,
                "set_email_flags",
                {
                    "account_name": "alice-managed",
                    "email_ids": [managed_uid],
                    "operation": "remove",
                    "flags": [r"\Seen"],
                },
            )
            assert unset_seen["result"] == r"Successfully removed \Seen from 1 email(s)"
            assert r"\Seen" not in _wait_for_message(ALICE, "INBOX", subject).flags
            set_flagged = await _call_tool(
                session,
                "set_email_flags",
                {
                    "account_name": "alice-managed",
                    "email_ids": [managed_uid],
                    "operation": "add",
                    "flags": [r"\Flagged"],
                },
            )
            assert set_flagged["result"] == r"Successfully added \Flagged to 1 email(s)"
            assert r"\Flagged" in _wait_for_message(ALICE, "INBOX", subject).flags
            with contextlib.closing(sqlite3.connect(database)) as connection:
                assert connection.execute("SELECT COUNT(*) FROM index_coverage").fetchone()[0] == 0

            refreshed = await _call_tool(
                session,
                "list_emails_metadata",
                {"account_name": "alice-managed", "page_size": 10},
            )
            assert refreshed["total"] == 1
            move = await _call_tool(
                session,
                "move_emails",
                {
                    "account_name": "alice-managed",
                    "email_ids": [managed_uid],
                    "source_mailbox": "INBOX",
                    "destination_mailbox": "Archive",
                },
            )
            assert move["result"] == "Successfully moved 1 email(s) to Archive"
            assert _find_message(ALICE, "INBOX", subject) is None
            _wait_for_message(ALICE, "Archive", subject)
            with contextlib.closing(sqlite3.connect(database)) as connection:
                assert connection.execute("SELECT COUNT(*) FROM index_coverage").fetchone()[0] == 0

            draft_subject = f"managed-draft-{uuid.uuid4().hex}"
            saved = await _call_tool(
                session,
                "save_to_mailbox",
                {
                    "account_name": "alice-managed",
                    "recipients": [BOB[0]],
                    "subject": draft_subject,
                    "body": "Managed draft body",
                    "mailbox": "Drafts",
                },
            )
            assert "Email saved to 'Drafts' successfully" in saved["result"]
            draft = _wait_for_message(ALICE, "Drafts", draft_subject)
            draft_page = await _call_tool(
                session,
                "list_emails_metadata",
                {"account_name": "alice-managed", "mailbox": "Drafts", "page_size": 10},
            )
            assert draft_page["total"] == 1
            deleted = await _call_tool(
                session,
                "delete_emails",
                {
                    "account_name": "alice-managed",
                    "email_ids": [draft.uid],
                    "mailbox": "Drafts",
                },
            )
            assert deleted["result"] == "Successfully deleted 1 email(s)"
            assert _find_message(ALICE, "Drafts", draft_subject) is None
            with contextlib.closing(sqlite3.connect(database)) as connection:
                remaining = connection.execute(
                    """SELECT COUNT(*) FROM index_coverage c
                       JOIN mailbox_projection m ON m.id = c.mailbox_id
                       WHERE m.remote_name = 'Drafts'"""
                ).fetchone()[0]
                assert remaining == 0

            invalid = await session.call_tool(
                "mark_emails_as_read",
                arguments={"account_name": "alice-managed", "email_ids": ["01"]},
            )
            assert invalid.isError is True
            validation_error = _text_content(invalid)
            assert "email_ids.0" in validation_error
            assert "pattern" in validation_error.lower()

            # Disablement commits in a separate management process and must be
            # observed before the next provider access in this same stdio session.
            _run_cli(
                console_script,
                server_env,
                ["account", "disable", "alice-managed", "--expected-revision", "3"],
            )
            denied = await session.call_tool("list_mailboxes", arguments={"account_name": "alice-managed"})
            assert denied.isError is True
            assert "not found" in _text_content(denied).lower()
            denied_metadata = await session.call_tool(
                "list_emails_metadata",
                arguments={"account_name": "alice-managed"},
            )
            assert denied_metadata.isError is True
            assert "not found" in _text_content(denied_metadata).lower()
            denied_mutation = await session.call_tool(
                "mark_emails_as_read",
                arguments={"account_name": "alice-managed", "email_ids": [managed_uid]},
            )
            assert denied_mutation.isError is True
            assert "not found" in _text_content(denied_mutation).lower()

            # Credential detachment, replacement, re-enable, active update, and
            # soft removal must all be observed by the already-running server.
            _run_cli(
                console_script,
                server_env,
                [
                    "account",
                    "remove-secret",
                    "alice-managed",
                    "incoming",
                    "--expected-revision",
                    "4",
                ],
            )
            _run_cli(
                console_script,
                server_env,
                ["account", "set-secret", "alice-managed", "incoming", "--password-stdin"],
                stdin=f"{ALICE[1]}\n",
            )
            _run_cli(
                console_script,
                server_env,
                ["account", "enable", "alice-managed", "--expected-revision", "6"],
            )
            restored = await _call_tool(session, "list_mailboxes", {"account_name": "alice-managed"})
            assert "INBOX" in {mailbox["name"] for mailbox in restored["result"]}
            _run_cli(
                console_script,
                server_env,
                [
                    "account",
                    "update",
                    "alice-managed",
                    "--expected-revision",
                    "7",
                    "--name",
                    "alice-managed-updated",
                ],
            )
            updated_accounts = await _call_tool(session, "list_available_accounts", {})
            assert [account["account_name"] for account in updated_accounts["result"]] == ["alice-managed-updated"]
            _run_cli(
                console_script,
                server_env,
                ["account", "disable", "alice-managed-updated", "--expected-revision", "8"],
            )
            _run_cli(
                console_script,
                server_env,
                [
                    "account",
                    "remove",
                    "alice-managed-updated",
                    "--expected-revision",
                    "9",
                    "--confirm",
                    "alice-managed-updated",
                ],
            )
            removed = await session.call_tool("list_mailboxes", arguments={"account_name": "alice-managed-updated"})
            assert removed.isError is True
            assert "not found" in _text_content(removed).lower()


@pytest.mark.asyncio
async def test_explicit_legacy_import_preview_apply_and_managed_stdio_against_greenmail(tmp_path: Path) -> None:
    """Prove effective legacy preview, confirmed import, automatic cutover, and stdio."""
    _wait_until_ready()
    app_dir = tmp_path / "managed-import"
    app_dir.mkdir(mode=0o700)
    app_dir.chmod(0o700)
    config_path = app_dir / "config.toml"
    config_path.write_text(CONFIG_TEMPLATE)
    config_path.chmod(0o600)
    database = app_dir / "catalog.sqlite3"
    keyring_path = app_dir / "e2e-keyring.sqlite3"
    console_script = Path(sys.executable).with_name("mcp-email-server")
    server_env = {key: value for key, value in os.environ.items() if not key.startswith("MCP_EMAIL_SERVER_")}
    server_env.update({
        "MCP_EMAIL_SERVER_CONFIG_PATH": str(config_path),
        "MCP_EMAIL_SERVER_E2E_KEYRING_PATH": str(keyring_path),
        "MCP_EMAIL_SERVER_LOG_LEVEL": "WARNING",
        "PYTHON_KEYRING_BACKEND": "dev.greenmail.file_keyring.FileKeyring",
        "PYTHONPATH": str(Path.cwd()),
        # A complete environment account proves effective legacy composition.
        "MCP_EMAIL_SERVER_ACCOUNT_NAME": "environment-only",
        "MCP_EMAIL_SERVER_EMAIL_ADDRESS": ALICE[0],
        "MCP_EMAIL_SERVER_PASSWORD": ALICE[1],
        "MCP_EMAIL_SERVER_IMAP_HOST": IMAP_HOST,
        "MCP_EMAIL_SERVER_IMAP_PORT": str(IMAP_PORT),
        "MCP_EMAIL_SERVER_IMAP_SSL": "false",
    })

    _run_cli(console_script, server_env, ["config", "init", "--database", str(database)])
    stored_source = config_path.read_bytes()
    preview = _run_cli(console_script, server_env, ["config", "import-legacy"])
    assert "account=alice action=create" in preview
    assert "account=bob action=create" in preview
    assert "account=environment-only action=create" in preview
    assert "secret_source=environment" in preview
    assert ALICE[1] not in preview
    with contextlib.closing(sqlite3.connect(database)) as connection:
        assert connection.execute("SELECT COUNT(*) FROM managed_account").fetchone()[0] == 0

    applied = _run_cli(
        console_script,
        server_env,
        ["config", "import-legacy", "--apply"],
        stdin="IMPORT\n",
    )
    assert "created=environment-only,alice,bob" in applied
    assert config_path.read_bytes() == stored_source
    bootstrap = read_bootstrap(config_path)
    assert bootstrap.mode == "managed"
    assert bootstrap.db_path == database
    assert bootstrap.revision == 2
    test_output = _run_cli(console_script, server_env, ["account", "test", "alice"])
    assert "connectivity test passed" in test_output
    environment_test = _run_cli(console_script, server_env, ["account", "test", "environment-only"])
    assert "connectivity test passed" in environment_test

    server = StdioServerParameters(
        command=str(console_script),
        args=["stdio"],
        env=server_env,
        cwd=Path.cwd(),
    )
    async with stdio_client(server) as (read_stream, write_stream):
        async with ClientSession(
            read_stream,
            write_stream,
            read_timeout_seconds=timedelta(seconds=15),
        ) as session:
            await session.initialize()
            accounts = await _call_tool(session, "list_available_accounts", {})
            assert {account["account_name"] for account in accounts["result"]} == {
                "alice",
                "bob",
                "environment-only",
            }
            mailboxes = await _call_tool(session, "list_mailboxes", {"account_name": "alice"})
            assert "INBOX" in {mailbox["name"] for mailbox in mailboxes["result"]}


@pytest.mark.asyncio
async def test_managed_stdio_missing_database_fails_closed_without_legacy_fallback(tmp_path: Path) -> None:
    """A selected managed catalog cannot silently fall back to preserved TOML rows."""
    app_dir = tmp_path / "managed-fail-closed"
    app_dir.mkdir(mode=0o700)
    app_dir.chmod(0o700)
    config_path = app_dir / "config.toml"
    database = app_dir / "catalog.sqlite3"
    keyring_path = app_dir / "e2e-keyring.sqlite3"
    console_script = Path(sys.executable).with_name("mcp-email-server")
    server_env = {key: value for key, value in os.environ.items() if not key.startswith("MCP_EMAIL_SERVER_")}
    server_env.update({
        "MCP_EMAIL_SERVER_CONFIG_PATH": str(config_path),
        "MCP_EMAIL_SERVER_E2E_KEYRING_PATH": str(keyring_path),
        "MCP_EMAIL_SERVER_LOG_LEVEL": "WARNING",
        "PYTHON_KEYRING_BACKEND": "dev.greenmail.file_keyring.FileKeyring",
        "PYTHONPATH": str(Path.cwd()),
    })
    _run_cli(console_script, server_env, ["config", "init", "--database", str(database)])
    _run_cli(
        console_script,
        server_env,
        [
            "account",
            "add",
            "managed-only",
            "--email",
            ALICE[0],
            "--full-name",
            "Managed Only",
            "--imap-host",
            IMAP_HOST,
            "--imap-port",
            str(IMAP_PORT),
            "--imap-user",
            ALICE[0],
            "--no-imap-ssl",
            "--password-stdin",
        ],
        stdin=f"{ALICE[1]}\n",
    )
    # Preserve a complete legacy account as a fallback tripwire. Managed startup
    # must ignore it even when the selected database disappears.
    with config_path.open("a") as destination:
        destination.write("\n" + CONFIG_TEMPLATE)
    config_path.chmod(0o600)
    missing_path = database.with_suffix(".missing")
    database.rename(missing_path)

    completed = subprocess.run(  # noqa: S603 - fixed installed script and literal stdio command
        [str(console_script), "stdio"],
        cwd=Path.cwd(),
        env=server_env,
        capture_output=True,
        text=True,
        timeout=15,
        check=False,
    )
    output = completed.stdout + completed.stderr
    assert completed.returncode == 1
    assert "missing" in output.lower()
    assert "alice-password" not in output
    assert "alice" not in output.lower()


@pytest.mark.asyncio
async def test_metadata_index_paging_fallback_and_restart_reuse_against_greenmail(tmp_path: Path) -> None:
    """Exercise population, qualified SQLite reuse, filters, bounds, and restart."""
    _wait_until_ready()
    _ensure_empty_mailboxes(BOB, ["INBOX"])
    run_id = uuid.uuid4().hex
    subjects = [f"metadata-index-{run_id}-{number}" for number in range(5)]
    for number, subject in enumerate(subjects):
        _seed_message(subject, f"indexed body {number}; unique needle {run_id}-{number}")
        _wait_for_message(BOB, "INBOX", subject)
    flagged = _wait_for_message(BOB, "INBOX", subjects[2])
    _add_flags(BOB, "INBOX", flagged.uid, r"\Seen \Flagged")

    config_path = tmp_path / "config.toml"
    config_path.write_text(CONFIG_TEMPLATE)
    config_path.chmod(0o600)
    database = tmp_path / "db.sqlite3"
    server_env = {key: value for key, value in os.environ.items() if not key.startswith("MCP_EMAIL_SERVER_")}
    server_env.update({
        "MCP_EMAIL_SERVER_CONFIG_PATH": str(config_path),
        "MCP_EMAIL_SERVER_CREDENTIAL_STORAGE": "plaintext",
        "MCP_EMAIL_SERVER_LOG_LEVEL": "WARNING",
    })
    console_script = Path(sys.executable).with_name("mcp-email-server")
    server = StdioServerParameters(
        command=str(console_script),
        args=["stdio"],
        env=server_env,
        cwd=Path.cwd(),
    )

    async def exercise_session(*, verify_filters: bool) -> tuple[str, dict[str, Any]]:
        async with stdio_client(server) as (read_stream, write_stream):
            async with ClientSession(
                read_stream,
                write_stream,
                read_timeout_seconds=timedelta(seconds=15),
            ) as session:
                await session.initialize()
                first = await _call_tool(
                    session,
                    "list_emails_metadata",
                    {"account_name": "bob", "page": 1, "page_size": 2},
                )
                assert first["total"] == 5
                assert len(first["emails"]) == 2
                assert [int(email["email_id"]) for email in first["emails"]] == sorted(
                    [int(email["email_id"]) for email in first["emails"]], reverse=True
                )
                with contextlib.closing(sqlite3.connect(database)) as connection:
                    coverage = connection.execute(
                        "SELECT completeness, message_count, observed_at FROM index_coverage"
                    ).fetchone()
                    rows = connection.execute("SELECT COUNT(*) FROM message_metadata_projection").fetchone()[0]
                assert coverage is not None
                assert coverage[0:2] == ("COMPLETE", 5)
                assert rows == 5

                second = await _call_tool(
                    session,
                    "list_emails_metadata",
                    {"account_name": "bob", "page": 2, "page_size": 2},
                )
                assert second["total"] == 5
                assert len(second["emails"]) == 2
                with contextlib.closing(sqlite3.connect(database)) as connection:
                    reused_at = connection.execute("SELECT observed_at FROM index_coverage").fetchone()[0]
                assert reused_at == coverage[2]

                if verify_filters:
                    filter_cases = [
                        ({"subject": subjects[1]}, 1),
                        ({"from_address": ALICE[0]}, 5),
                        ({"to_address": BOB[0]}, 5),
                        ({"seen": True}, 1),
                        ({"flagged": True}, 1),
                        ({"body": f"unique needle {run_id}-4"}, 1),
                        ({"text": subjects[3]}, 1),
                        ({"has_attachment": False}, 5),
                        ({"has_attachment": True}, 0),
                    ]
                    for filters, expected_total in filter_cases:
                        result = await _call_tool(
                            session,
                            "list_emails_metadata",
                            {"account_name": "bob", "page_size": 10, **filters},
                        )
                        assert result["total"] == expected_total, (filters, result)
                    invalid = await session.call_tool(
                        "list_emails_metadata",
                        arguments={"account_name": "bob", "page_size": 101},
                    )
                    assert invalid.isError is True
                return coverage[2], first

    first_observed_at, first_page = await exercise_session(verify_filters=True)
    restart_observed_at, restart_page = await exercise_session(verify_filters=False)
    assert restart_observed_at == first_observed_at
    assert restart_page == first_page


@pytest.mark.asyncio
async def test_current_stdio_server_against_greenmail(tmp_path: Path) -> None:
    """Exercise the current public MCP/CLI/config boundary against real mail sockets."""
    _wait_until_ready()
    _ensure_empty_mailboxes(ALICE, ["INBOX", "Sent", "Drafts", "Archive"])
    _ensure_empty_mailboxes(BOB, ["INBOX", "Drafts", "Archive"])

    run_id = uuid.uuid4().hex
    sent_subject = f"mcp-e2e-send-{run_id}"
    sent_body = f"Body produced through MCP stdio {run_id}"
    root_message_id = f"<root-{run_id}@example.test>"
    parent_message_id = f"<parent-{run_id}@example.test>"
    references = f"{root_message_id} {parent_message_id}"
    attachment_bytes = b"greenmail attachment roundtrip\x00\xff\n"
    attachment_source = tmp_path / "roundtrip.bin"
    attachment_source.write_bytes(attachment_bytes)
    attachment_download = tmp_path / "downloaded.bin"

    config_path = tmp_path / "config.toml"
    config_path.write_text(CONFIG_TEMPLATE)
    config_path.chmod(0o600)
    server_env = {key: value for key, value in os.environ.items() if not key.startswith("MCP_EMAIL_SERVER_")}
    server_env.update({
        "MCP_EMAIL_SERVER_CONFIG_PATH": str(config_path),
        "MCP_EMAIL_SERVER_CREDENTIAL_STORAGE": "plaintext",
        "MCP_EMAIL_SERVER_LOG_LEVEL": "WARNING",
    })
    console_script = Path(sys.executable).with_name("mcp-email-server")
    assert console_script.is_file(), f"Installed console script not found: {console_script}"
    server = StdioServerParameters(
        command=str(console_script),
        args=["stdio"],
        env=server_env,
        cwd=Path.cwd(),
    )

    async with stdio_client(server) as (read_stream, write_stream):
        async with ClientSession(
            read_stream,
            write_stream,
            read_timeout_seconds=timedelta(seconds=15),
        ) as session:
            initialized = await session.initialize()
            assert initialized.serverInfo.name == "email"
            assert initialized.serverInfo.version == importlib.metadata.version("mcp-email-server")

            tools = await session.list_tools()
            tool_names = {tool.name for tool in tools.tools}
            assert {
                "list_available_accounts",
                "list_emails_metadata",
                "get_emails_content",
                "send_email",
                "forward_email",
                "save_to_mailbox",
                "delete_emails",
                "set_email_flags",
                "mark_emails_as_read",
                "move_emails",
                "archive_emails",
                "list_mailboxes",
                "download_attachment",
            } <= tool_names

            accounts = await _call_tool(session, "list_available_accounts", {})
            assert {account["account_name"] for account in accounts["result"]} == {"alice", "bob"}

            send_result = await _call_tool(
                session,
                "send_email",
                {
                    "account_name": "alice",
                    "recipients": [BOB[0]],
                    "subject": sent_subject,
                    "body": sent_body,
                    "attachments": [str(attachment_source)],
                    "in_reply_to": parent_message_id,
                    "references": references,
                },
            )
            assert send_result["result"] == f"Email sent successfully to {BOB[0]} with 1 attachment(s)"

            delivered = _wait_for_message(BOB, "INBOX", sent_subject)
            assert sent_body in (delivered.message.get_body(preferencelist=("plain",)).get_content())
            delivered_from = delivered.message["From"]
            assert delivered_from is not None
            assert [(address.display_name, address.addr_spec) for address in delivered_from.addresses] == [
                ("alice@example.test", "alice@example.test")
            ]
            assert str(delivered.message["In-Reply-To"]) == parent_message_id
            assert str(delivered.message["References"]) == references
            delivered_attachments = list(delivered.message.iter_attachments())
            assert len(delivered_attachments) == 1
            assert delivered_attachments[0].get_filename() == attachment_source.name
            assert delivered_attachments[0].get_payload(decode=True) == attachment_bytes

            sent_copy = _wait_for_message(ALICE, "Sent", sent_subject)
            assert sent_body in sent_copy.message.get_body(preferencelist=("plain",)).get_content()
            sent_copy_from = sent_copy.message["From"]
            assert sent_copy_from is not None
            assert [(address.display_name, address.addr_spec) for address in sent_copy_from.addresses] == [
                ("alice@example.test", "alice@example.test")
            ]
            sent_copy_metadata = await _metadata_for_subject_in_mailbox(session, "alice", "Sent", sent_subject)
            sent_copy_content = await _call_tool(
                session,
                "get_emails_content",
                {
                    "account_name": "alice",
                    "mailbox": "Sent",
                    "email_ids": [sent_copy_metadata["email_id"]],
                },
            )
            assert sent_copy_content["emails"][0]["in_reply_to"] == parent_message_id
            assert sent_copy_content["emails"][0]["references"] == references

            denied_subject = f"mcp-e2e-denied-send-{run_id}"
            denied_recipient = f"missing-{run_id}@example.test"
            denied_send = await session.call_tool(
                "send_email",
                arguments={
                    "account_name": "alice",
                    "recipients": [BOB[0], denied_recipient],
                    "subject": denied_subject,
                    "body": "Recipient policy must reject before SMTP",
                },
            )
            assert denied_send.isError is True
            assert "not in allowlist" in _text_content(denied_send)
            assert _find_message(BOB, "INBOX", denied_subject) is None
            assert _find_message(ALICE, "Sent", denied_subject) is None

            forward_source_subject = f"mcp-e2e-forward-source-{run_id}"
            forward_source_body = f"Original content that must survive the forward {run_id}"
            forwarded_attachment_bytes = b"forwarded attachment bytes \x00\xfe\n"
            forwarded_attachment_name = "forwarded-report.pdf"
            _seed_message_with_attachment_as(
                BOB,
                ALICE[0],
                forward_source_subject,
                forward_source_body,
                filename=forwarded_attachment_name,
                payload=forwarded_attachment_bytes,
                maintype="application",
                subtype="pdf",
            )
            _wait_for_message(ALICE, "INBOX", forward_source_subject)
            forward_source_metadata = await _metadata_for_subject(session, "alice", forward_source_subject)
            forward_note = f"Please review this {run_id}"
            forward_result = await _call_tool(
                session,
                "forward_email",
                {
                    "account_name": "alice",
                    "email_id": forward_source_metadata["email_id"],
                    "recipients": [BOB[0]],
                    "body": forward_note,
                },
            )
            assert forward_result["result"] == f"Email forwarded successfully to {BOB[0]}"

            # Read the delivered forward back over plain imaplib rather than trusting
            # the server's own report of what it claims to have sent.
            forwarded_subject = f"Fwd: {forward_source_subject}"
            forwarded = _wait_for_message(BOB, "INBOX", forwarded_subject)
            forwarded_text = forwarded.message.get_body(preferencelist=("plain",)).get_content()
            assert forward_note in forwarded_text
            assert "---------- Forwarded message ----------" in forwarded_text
            assert f"From: {BOB[0]}" in forwarded_text
            assert f"Recipients: {ALICE[0]}" in forwarded_text
            assert f"Subject: {forward_source_subject}" in forwarded_text
            assert forward_source_body in forwarded_text
            forwarded_parts = list(forwarded.message.iter_attachments())
            assert len(forwarded_parts) == 1
            assert forwarded_parts[0].get_content_type() == "application/pdf"
            assert forwarded_parts[0].get_filename() == forwarded_attachment_name
            assert forwarded_parts[0].get_payload(decode=True) == forwarded_attachment_bytes

            forwarded_sent_copy = _wait_for_message(ALICE, "Sent", forwarded_subject)
            forwarded_sent_text = forwarded_sent_copy.message.get_body(preferencelist=("plain",)).get_content()
            assert forward_note in forwarded_sent_text
            assert forward_source_body in forwarded_sent_text
            assert [part.get_filename() for part in forwarded_sent_copy.message.iter_attachments()] == [
                forwarded_attachment_name
            ]

            denied_forward_subject = f"mcp-e2e-forward-denied-{run_id}"
            _seed_message_as(BOB, ALICE[0], denied_forward_subject, "This forward must never leave the process")
            _wait_for_message(ALICE, "INBOX", denied_forward_subject)
            denied_forward_metadata = await _metadata_for_subject(session, "alice", denied_forward_subject)
            denied_forward = await session.call_tool(
                "forward_email",
                arguments={
                    "account_name": "alice",
                    "email_id": denied_forward_metadata["email_id"],
                    "recipients": [denied_recipient],
                },
            )
            assert denied_forward.isError is True
            assert "not in allowlist" in _text_content(denied_forward)
            assert _find_message(BOB, "INBOX", f"Fwd: {denied_forward_subject}") is None
            assert _find_message(ALICE, "Sent", f"Fwd: {denied_forward_subject}") is None

            sent_metadata = await _metadata_for_subject(session, "bob", sent_subject)
            assert sent_metadata["sender"].endswith("<alice@example.test>") or sent_metadata["sender"] == ALICE[0]
            assert BOB[0] in sent_metadata["recipients"]
            # Metadata intentionally excludes thread headers and attachment names; the full-content path supplies them.
            assert sent_metadata["attachments"] == []
            assert "in_reply_to" not in sent_metadata
            assert "references" not in sent_metadata

            content = await _call_tool(
                session,
                "get_emails_content",
                {"account_name": "bob", "email_ids": [sent_metadata["email_id"]]},
            )
            assert content["requested_count"] == 1
            assert content["retrieved_count"] == 1
            assert content["failed_ids"] == []
            assert content["emails"][0]["in_reply_to"] == parent_message_id
            assert content["emails"][0]["references"] == references
            assert content["emails"][0]["attachments"] == [attachment_source.name]
            assert sent_body in content["emails"][0]["body"]

            mark_result = await _call_tool(
                session,
                "mark_emails_as_read",
                {"account_name": "bob", "email_ids": [sent_metadata["email_id"]]},
            )
            assert mark_result["result"] == "Successfully marked 1 email(s) as read"
            assert r"\Seen" in _wait_for_message(BOB, "INBOX", sent_subject).flags

            download = await _call_tool(
                session,
                "download_attachment",
                {
                    "account_name": "bob",
                    "email_id": sent_metadata["email_id"],
                    "attachment_name": attachment_source.name,
                    "save_path": str(attachment_download),
                },
            )
            assert download["attachment_name"] == attachment_source.name
            assert download["size"] == len(attachment_bytes)
            assert Path(download["saved_path"]) == attachment_download
            assert attachment_download.read_bytes() == attachment_bytes

            move_result = await _call_tool(
                session,
                "move_emails",
                {
                    "account_name": "bob",
                    "email_ids": [sent_metadata["email_id"]],
                    "source_mailbox": "INBOX",
                    "destination_mailbox": "Archive",
                },
            )
            assert move_result["result"] == "Successfully moved 1 email(s) to Archive"
            assert _find_message(BOB, "INBOX", sent_subject) is None
            _wait_for_message(BOB, "Archive", sent_subject)

            archive_subject = f"mcp-e2e-archive-{run_id}"
            _seed_message(archive_subject, "Archive this message")
            _wait_for_message(BOB, "INBOX", archive_subject)
            archive_metadata = await _metadata_for_subject(session, "bob", archive_subject)
            archive_result = await _call_tool(
                session,
                "archive_emails",
                {"account_name": "bob", "email_ids": [archive_metadata["email_id"]]},
            )
            assert archive_result["result"] == "Successfully archived 1 email(s) to Archive"
            assert _find_message(BOB, "INBOX", archive_subject) is None
            _wait_for_message(BOB, "Archive", archive_subject)

            draft_subject = f"mcp-e2e-draft-{run_id}"
            draft_body = "Draft body created through MCP"
            save_result = await _call_tool(
                session,
                "save_to_mailbox",
                {
                    "account_name": "alice",
                    "recipients": [BOB[0]],
                    "subject": draft_subject,
                    "body": draft_body,
                    "mailbox": "Drafts",
                },
            )
            assert "Email saved to 'Drafts' successfully" in save_result["result"]
            draft = _wait_for_message(ALICE, "Drafts", draft_subject)
            assert draft_body in draft.message.get_body(preferencelist=("plain",)).get_content()
            assert {r"\Draft", r"\Seen"} <= draft.flags

            draft_metadata = await _metadata_for_subject_in_mailbox(session, "alice", "Drafts", draft_subject)
            delete_draft = await _call_tool(
                session,
                "delete_emails",
                {
                    "account_name": "alice",
                    "email_ids": [draft_metadata["email_id"]],
                    "mailbox": "Drafts",
                },
            )
            assert delete_draft["result"] == "Successfully deleted 1 email(s)"
            assert _find_message(ALICE, "Drafts", draft_subject) is None

            # Another IMAP client may already have left an unrelated message with
            # \\Deleted set. A message-scoped MCP delete must expunge only its own
            # target rather than silently committing the other client's deletion.
            pending_subject = f"mcp-e2e-unrelated-pending-delete-{run_id}"
            delete_subject = f"mcp-e2e-scoped-delete-{run_id}"
            _seed_message(pending_subject, "Leave this message pending deletion")
            _seed_message(delete_subject, "Delete only this message")
            pending = _wait_for_message(BOB, "INBOX", pending_subject)
            _wait_for_message(BOB, "INBOX", delete_subject)
            delete_metadata = await _metadata_for_subject(session, "bob", delete_subject)
            _mark_deleted_without_expunge(BOB, "INBOX", pending.uid)
            assert r"\Deleted" in _wait_for_message(BOB, "INBOX", pending_subject).flags

            scoped_delete = await _call_tool(
                session,
                "delete_emails",
                {
                    "account_name": "bob",
                    "email_ids": [delete_metadata["email_id"]],
                    "mailbox": "INBOX",
                },
            )
            assert scoped_delete["result"] == "Successfully deleted 1 email(s)"
            assert _find_message(BOB, "INBOX", delete_subject) is None
            still_pending = _wait_for_message(BOB, "INBOX", pending_subject)
            assert r"\Deleted" in still_pending.flags

            # Native MOVE must preserve the same unrelated pending deletion too.
            move_pending_subject = f"mcp-e2e-unrelated-pending-move-{run_id}"
            move_target_subject = f"mcp-e2e-scoped-move-{run_id}"
            _seed_message(move_pending_subject, "Leave this message pending while another moves")
            _seed_message(move_target_subject, "Move only this message")
            move_pending = _wait_for_message(BOB, "INBOX", move_pending_subject)
            _wait_for_message(BOB, "INBOX", move_target_subject)
            move_target_metadata = await _metadata_for_subject(session, "bob", move_target_subject)
            _mark_deleted_without_expunge(BOB, "INBOX", move_pending.uid)

            scoped_move = await _call_tool(
                session,
                "move_emails",
                {
                    "account_name": "bob",
                    "email_ids": [move_target_metadata["email_id"]],
                    "source_mailbox": "INBOX",
                    "destination_mailbox": "Archive",
                },
            )
            assert scoped_move["result"] == "Successfully moved 1 email(s) to Archive"
            assert _find_message(BOB, "INBOX", move_target_subject) is None
            _wait_for_message(BOB, "Archive", move_target_subject)
            still_pending_after_move = _wait_for_message(BOB, "INBOX", move_pending_subject)
            assert r"\Deleted" in still_pending_after_move.flags

            mailboxes = await _call_tool(session, "list_mailboxes", {"account_name": "alice"})
            mailbox_names = {mailbox["name"] for mailbox in mailboxes["result"]}
            assert {"INBOX", "Sent", "Drafts", "Archive"} <= mailbox_names


@pytest.mark.asyncio
async def test_unlimited_body_length_returns_whole_body_and_rejects_oversized_window(tmp_path: Path) -> None:
    """max_body_length=0 returns an oversized body untruncated; 100001 is rejected."""
    _wait_until_ready()
    _ensure_empty_mailboxes(BOB, ["INBOX"])

    run_id = uuid.uuid4().hex
    subject = f"mcp-e2e-unlimited-body-{run_id}"
    # Longer than both the default window and the explicit maximum, so a truncating
    # request and an untruncated one cannot be confused for each other. Short unique
    # lines keep the message 7-bit clean; assertions never assume a line ending.
    line_count = 2_400
    body = "".join(f"line{index:05d}-{'x' * 40}\n" for index in range(line_count))
    assert len(body) > 100_001
    _seed_message(subject, body)
    _wait_for_message(BOB, "INBOX", subject)

    config_path = tmp_path / "config.toml"
    config_path.write_text(CONFIG_TEMPLATE)
    config_path.chmod(0o600)
    server_env = {key: value for key, value in os.environ.items() if not key.startswith("MCP_EMAIL_SERVER_")}
    server_env.update({
        "MCP_EMAIL_SERVER_CONFIG_PATH": str(config_path),
        "MCP_EMAIL_SERVER_CREDENTIAL_STORAGE": "plaintext",
        "MCP_EMAIL_SERVER_LOG_LEVEL": "WARNING",
    })
    console_script = Path(sys.executable).with_name("mcp-email-server")
    assert console_script.is_file(), f"Installed console script not found: {console_script}"
    server = StdioServerParameters(command=str(console_script), args=["stdio"], env=server_env, cwd=Path.cwd())

    async with stdio_client(server) as (read_stream, write_stream):
        async with ClientSession(read_stream, write_stream, read_timeout_seconds=timedelta(seconds=15)) as session:
            await session.initialize()
            metadata = await _metadata_for_subject(session, "bob", subject)

            truncated = await _call_tool(
                session,
                "get_emails_content",
                {"account_name": "bob", "email_ids": [metadata["email_id"]], "max_body_length": 1_000},
            )
            marker = "...[TRUNCATED]"
            truncated_body = truncated["emails"][0]["body"]
            assert truncated_body.endswith(marker)
            assert len(truncated_body) == 1_000 + len(marker)
            assert truncated_body.startswith("line00000-")
            assert f"line{line_count - 1:05d}-" not in truncated_body

            untruncated_bodies = []
            for unlimited in (0, None):
                whole = await _call_tool(
                    session,
                    "get_emails_content",
                    {"account_name": "bob", "email_ids": [metadata["email_id"]], "max_body_length": unlimited},
                )
                whole_body = whole["emails"][0]["body"]
                assert marker not in whole_body
                assert whole_body.startswith("line00000-")
                # The last seeded line proves nothing was dropped past the explicit ceiling.
                assert f"line{line_count - 1:05d}-" in whole_body
                assert len(whole_body) > 100_001
                untruncated_bodies.append(whole_body)
            assert untruncated_bodies[0] == untruncated_bodies[1]

            offset_window = await _call_tool(
                session,
                "get_emails_content",
                {
                    "account_name": "bob",
                    "email_ids": [metadata["email_id"]],
                    "body_offset": 50_000,
                    "max_body_length": 0,
                },
            )
            offset_body = offset_window["emails"][0]["body"]
            assert marker not in offset_body
            assert offset_body == untruncated_bodies[0][50_000:]

            rejected = await session.call_tool(
                "get_emails_content",
                arguments={
                    "account_name": "bob",
                    "email_ids": [metadata["email_id"]],
                    "max_body_length": 100_001,
                },
            )
            assert rejected.isError is True
            assert "max_body_length" in _text_content(rejected)


# Port slots for new per-cluster E2E tests. Each cluster appends its own test
# function directly ABOVE its own marker rather than extending
# `test_current_stdio_server_against_greenmail`, which is reserved for the
# trash-first delete change. Slot order is fixed everywhere: A, B1, C, B2.
def _hierarchy_delimiter(credentials: tuple[str, str]) -> str:
    """Read the server's own hierarchy delimiter instead of assuming one.

    GreenMail uses `.`, most other servers use `/`, and a nested folder name
    built from the wrong one silently becomes a top-level folder.
    """
    with _imap_session(credentials) as client:
        # imaplib splices arguments verbatim, so the empty reference must arrive quoted.
        status, rows = client.list('""', "INBOX")
        assert status == "OK", f"LIST INBOX failed: {status}"
        for row in rows:
            if isinstance(row, bytes) and b'"' in row:
                # `(flags) "<delimiter>" "INBOX"` — the delimiter is the first quoted token.
                return row.split(b'"')[1].decode()
    pytest.fail(f"No hierarchy delimiter reported for INBOX: {rows!r}")


def _mailbox_subjects(credentials: tuple[str, str], mailbox: str) -> set[str]:
    """Read a mailbox's subjects and always CLOSE before logging out.

    GreenMail keeps a mailbox busy when a session selects it and then
    disconnects without CLOSE: a later DELETE of that mailbox never answers at
    all. `_find_message` does not close, so a test that deletes the folder it
    just inspected must use this helper instead.
    """
    with _imap_session(credentials) as client:
        try:
            status, _ = client.select(mailbox, readonly=True)
            assert status == "OK", f"SELECT {mailbox} failed: {status}"
            status, rows = client.uid("search", None, "ALL")
            assert status == "OK", f"UID SEARCH in {mailbox} failed: {status}"
            subjects: set[str] = set()
            for uid in (rows[0] or b"").split():
                status, fetched = client.uid("fetch", uid, "(BODY.PEEK[HEADER])")
                assert status == "OK", f"UID FETCH {uid!r} in {mailbox} failed: {status}"
                response = next((item for item in fetched if isinstance(item, tuple)), None)
                assert response is not None
                message = BytesParser(policy=policy.default).parsebytes(response[1])
                subjects.add(str(message.get("Subject", "")))
            return subjects
        finally:
            with contextlib.suppress(Exception):
                client.close()


def _mailbox_exists(credentials: tuple[str, str], mailbox: str) -> bool:
    with _imap_session(credentials) as client:
        status, rows = client.list('""', f'"{mailbox}"')
        assert status == "OK", f"LIST {mailbox} failed: {status}"
        return any(isinstance(row, bytes) and row.strip() for row in rows)


def _drop_mailboxes(credentials: tuple[str, str], mailboxes: list[str]) -> None:
    with _imap_session(credentials) as client:
        with contextlib.suppress(Exception):
            client.close()
        for mailbox in mailboxes:
            with contextlib.suppress(Exception):
                client.delete(mailbox)


@pytest.mark.asyncio
async def test_folder_operations_stdio_round_trip_against_greenmail(tmp_path: Path) -> None:
    """create -> copy into -> rename -> delete through the MCP tools, plus the policy gate."""
    _wait_until_ready()
    _ensure_empty_mailboxes(BOB, ["INBOX"])

    run_id = uuid.uuid4().hex[:12]
    delimiter = _hierarchy_delimiter(BOB)
    created_folder = f"mcp-folder-{run_id}"
    # Built from the server's own delimiter: hard-coding `/` silently produces a
    # top-level folder on a `.`-delimited server such as GreenMail.
    child_folder = f"{created_folder}{delimiter}sub"
    renamed_folder = f"mcp-folder-{run_id}-renamed"
    copy_subject = f"mcp-e2e-copy-{run_id}"

    config_path = tmp_path / "config.toml"
    config_path.write_text(f"enable_folder_management = true\n{CONFIG_TEMPLATE}")
    config_path.chmod(0o600)
    base_env = {key: value for key, value in os.environ.items() if not key.startswith("MCP_EMAIL_SERVER_")}
    server_env = {
        **base_env,
        "MCP_EMAIL_SERVER_CONFIG_PATH": str(config_path),
        "MCP_EMAIL_SERVER_CREDENTIAL_STORAGE": "plaintext",
        "MCP_EMAIL_SERVER_LOG_LEVEL": "WARNING",
    }
    console_script = Path(sys.executable).with_name("mcp-email-server")
    assert console_script.is_file(), f"Installed console script not found: {console_script}"

    def _server(env: dict[str, str]) -> StdioServerParameters:
        return StdioServerParameters(command=str(console_script), args=["stdio"], env=env, cwd=Path.cwd())

    try:
        async with stdio_client(_server(server_env)) as (read_stream, write_stream):
            async with ClientSession(read_stream, write_stream, read_timeout_seconds=timedelta(seconds=15)) as session:
                await session.initialize()
                tool_names = {tool.name for tool in (await session.list_tools()).tools}
                assert {"copy_emails", "create_folder", "delete_folder", "rename_folder"} <= tool_names

                created = await _call_tool(
                    session, "create_folder", {"account_name": "bob", "folder_name": created_folder}
                )
                assert created["result"] == f"Folder '{created_folder}' created"
                assert _mailbox_exists(BOB, created_folder)

                child = await _call_tool(session, "create_folder", {"account_name": "bob", "folder_name": child_folder})
                assert child["result"] == f"Folder '{child_folder}' created"
                mailboxes = await _call_tool(session, "list_mailboxes", {"account_name": "bob"})
                listed = {mailbox["name"]: mailbox for mailbox in mailboxes["result"]}
                assert {created_folder, child_folder} <= listed.keys()
                assert listed[child_folder]["delimiter"] == delimiter

                _seed_message(copy_subject, f"Copy me into {created_folder}")
                _wait_for_message(BOB, "INBOX", copy_subject)
                metadata = await _metadata_for_subject(session, "bob", copy_subject)

                copied = await _call_tool(
                    session,
                    "copy_emails",
                    {
                        "account_name": "bob",
                        "email_ids": [metadata["email_id"]],
                        "source_mailbox": "INBOX",
                        "destination_mailbox": created_folder,
                    },
                )
                assert copied["result"] == f"Successfully copied 1 email(s) to {created_folder}"
                assert copy_subject in _mailbox_subjects(BOB, created_folder)
                # A copy is additive: the source message must still be in INBOX.
                assert _find_message(BOB, "INBOX", copy_subject) is not None

                # Remove the child first: most servers refuse to delete a parent
                # that still has children.
                dropped_child = await _call_tool(
                    session, "delete_folder", {"account_name": "bob", "folder_name": child_folder}
                )
                assert dropped_child["result"] == f"Folder '{child_folder}' deleted"
                assert not _mailbox_exists(BOB, child_folder)

                renamed = await _call_tool(
                    session,
                    "rename_folder",
                    {"account_name": "bob", "old_name": created_folder, "new_name": renamed_folder},
                )
                assert renamed["result"] == f"Folder '{created_folder}' renamed to '{renamed_folder}'"
                assert _mailbox_exists(BOB, renamed_folder)
                # The copied message travels with the mailbox.
                assert copy_subject in _mailbox_subjects(BOB, renamed_folder)

                deleted = await _call_tool(
                    session, "delete_folder", {"account_name": "bob", "folder_name": renamed_folder}
                )
                assert deleted["result"] == f"Folder '{renamed_folder}' deleted"
                assert not _mailbox_exists(BOB, renamed_folder)

        # A second server with the policy off must refuse the same shape mutation,
        # while the ungated copy tool stays available.
        gated_off_env = {**server_env, "MCP_EMAIL_SERVER_ENABLE_FOLDER_MANAGEMENT": "false"}
        async with stdio_client(_server(gated_off_env)) as (read_stream, write_stream):
            async with ClientSession(read_stream, write_stream, read_timeout_seconds=timedelta(seconds=15)) as session:
                await session.initialize()
                tool_names = {tool.name for tool in (await session.list_tools()).tools}
                # The tools stay visible; the denial is a runtime policy decision.
                assert {"create_folder", "delete_folder", "rename_folder"} <= tool_names

                denied = await session.call_tool(
                    "create_folder",
                    arguments={"account_name": "bob", "folder_name": f"{created_folder}-denied"},
                )
                assert denied.isError is True
                assert "Folder management is disabled" in _text_content(denied)
                assert not _mailbox_exists(BOB, f"{created_folder}-denied")
    finally:
        _drop_mailboxes(BOB, [child_folder, renamed_folder, created_folder, f"{created_folder}-denied"])


# port-slot A: folder ops (copy_emails, create_folder, delete_folder, rename_folder)
@pytest.mark.asyncio
async def test_label_reads_and_removal_against_greenmail(tmp_path: Path) -> None:
    """`Labels/` naming, Message-ID lookup, and label-scoped deletion over real IMAP.

    GreenMail's hierarchy delimiter is `.`, but the ProtonMail convention this port
    implements puts a literal `/` inside the mailbox *name*, not a hierarchy path.
    GreenMail accepts `CREATE "Labels/<name>"` verbatim and `LIST "" "Labels/*"`
    matches it, so this exercises the production naming path rather than a
    delimiter-adapted stand-in. Applying a label is another cluster's tool, so the
    label copy is seeded here with a raw APPEND of the message's own bytes.
    """
    _wait_until_ready()
    _ensure_empty_mailboxes(BOB, ["INBOX"])

    run_id = uuid.uuid4().hex[:8]
    label_name = f"E2E-{run_id}"
    label_mailbox = f"Labels/{label_name}"
    labelled_subject = f"mcp-e2e-labelled-{run_id}"
    plain_subject = f"mcp-e2e-unlabelled-{run_id}"

    def _raw_message(mailbox: str, uid: str) -> bytes:
        with _imap_session(BOB) as client:
            status, _ = client.select(mailbox, readonly=True)
            assert status == "OK"
            status, fetched = client.uid("fetch", uid, "(BODY.PEEK[])")
            assert status == "OK"
            response = next((item for item in fetched if isinstance(item, tuple)), None)
            assert response is not None
            return response[1]

    def _label_message_count() -> int:
        """Count the label mailbox, always releasing it with CLOSE.

        GreenMail wedges on DELETE of a mailbox that an earlier session selected
        and then dropped without CLOSE, so every inspection of a mailbox this
        test later deletes gives it back explicitly. The shared helpers do not
        CLOSE, which is why this one is local.
        """
        with _imap_session(BOB) as client:
            status, data = client.select(label_mailbox, readonly=True)
            assert status == "OK", f"SELECT {label_mailbox} failed: {status}"
            try:
                return int(data[0])
            finally:
                client.close()

    config_path = tmp_path / "config.toml"
    config_path.write_text(CONFIG_TEMPLATE)
    config_path.chmod(0o600)
    server_env = {key: value for key, value in os.environ.items() if not key.startswith("MCP_EMAIL_SERVER_")}
    server_env.update({
        "MCP_EMAIL_SERVER_CONFIG_PATH": str(config_path),
        "MCP_EMAIL_SERVER_CREDENTIAL_STORAGE": "plaintext",
        "MCP_EMAIL_SERVER_LOG_LEVEL": "WARNING",
    })
    console_script = Path(sys.executable).with_name("mcp-email-server")
    assert console_script.is_file(), f"Installed console script not found: {console_script}"
    server = StdioServerParameters(command=str(console_script), args=["stdio"], env=server_env, cwd=Path.cwd())

    try:
        with _imap_session(BOB) as client:
            status, _ = client.create(f'"{label_mailbox}"')
            assert status == "OK", f"CREATE {label_mailbox} failed: {status}"

        _seed_message(labelled_subject, "This message carries a label")
        _seed_message(plain_subject, "This message carries no label")
        labelled = _wait_for_message(BOB, "INBOX", labelled_subject)
        _wait_for_message(BOB, "INBOX", plain_subject)

        # Seed the label's own copy the way a provider would: the identical bytes,
        # so both copies share one Message-ID.
        with _imap_session(BOB) as client:
            status, _ = client.append(f'"{label_mailbox}"', None, None, _raw_message("INBOX", labelled.uid))
            assert status == "OK", f"APPEND to {label_mailbox} failed: {status}"
        assert _label_message_count() == 1

        async with stdio_client(server) as (read_stream, write_stream):
            async with ClientSession(
                read_stream,
                write_stream,
                read_timeout_seconds=timedelta(seconds=15),
            ) as session:
                await session.initialize()
                tools = await session.list_tools()
                assert {"list_labels", "get_email_labels", "remove_label"} <= {tool.name for tool in tools.tools}

                labels = await _call_tool(session, "list_labels", {"account_name": "bob"})
                seeded = [label for label in labels["result"] if label["name"] == label_name]
                assert len(seeded) == 1, labels
                assert seeded[0]["full_path"] == label_mailbox
                assert seeded[0]["delimiter"] == "."
                assert isinstance(seeded[0]["flags"], list)
                # The prefix is stripped and the bare container is never a label.
                assert not any(label["name"].startswith("Labels/") for label in labels["result"])
                assert all(label["name"] for label in labels["result"])

                labelled_metadata = await _metadata_for_subject(session, "bob", labelled_subject)
                plain_metadata = await _metadata_for_subject(session, "bob", plain_subject)

                applied = await _call_tool(
                    session,
                    "get_email_labels",
                    {"account_name": "bob", "email_id": labelled_metadata["email_id"], "mailbox": "INBOX"},
                )
                assert applied["result"] == [label_name]

                unapplied = await _call_tool(
                    session,
                    "get_email_labels",
                    {"account_name": "bob", "email_id": plain_metadata["email_id"], "mailbox": "INBOX"},
                )
                assert unapplied["result"] == []

                removed = await _call_tool(
                    session,
                    "remove_label",
                    {
                        "account_name": "bob",
                        "email_ids": [labelled_metadata["email_id"]],
                        "label_name": label_name,
                        "source_mailbox": "INBOX",
                    },
                )
                assert removed["result"] == f"Successfully removed label '{label_name}' from 1 email(s)"

                # Only the label's copy is gone: the message the caller named, and
                # every unrelated message, stay exactly where they were.
                assert _label_message_count() == 0
                assert _find_message(BOB, "INBOX", labelled_subject) is not None
                assert _find_message(BOB, "INBOX", plain_subject) is not None

                after = await _call_tool(
                    session,
                    "get_email_labels",
                    {"account_name": "bob", "email_id": labelled_metadata["email_id"], "mailbox": "INBOX"},
                )
                assert after["result"] == []

                repeated = await _call_tool(
                    session,
                    "remove_label",
                    {
                        "account_name": "bob",
                        "email_ids": [labelled_metadata["email_id"]],
                        "label_name": label_name,
                        "source_mailbox": "INBOX",
                    },
                )
                assert "label-not-found" in repeated["result"]
    finally:
        with contextlib.suppress(Exception), _imap_session(BOB) as client:
            # Release the mailbox from this session before removing it; a leftover
            # run-scoped label is harmless if GreenMail still refuses the DELETE.
            if client.select(f'"{label_mailbox}"')[0] == "OK":
                client.close()
            client.delete(f'"{label_mailbox}"')


# port-slot B1: label reads (list_labels, get_email_labels, remove_label)
# port-slot C: send path (Markdown rendering, quoted replies)@pytest.mark.asyncio
async def test_label_lifecycle_through_the_tools_against_greenmail(tmp_path: Path) -> None:
    """The whole label lifecycle driven through the MCP tools, not raw IMAP.

    This is the first coverage proving the label writes and the label reads agree
    on the `Labels/<name>` convention end to end. GreenMail's hierarchy delimiter
    is `.`, and the prefix stays a literal `/` inside the mailbox *name*, which is
    exactly what the ProtonMail convention requires.

    Two labels are used rather than one, because of a GreenMail defect that no
    ordering of a single label can dodge. GreenMail permanently breaks DELETE for
    any mailbox that some session ever SELECTed without CLOSE: the DELETE
    connection is dropped (EOF), a later SELECT+CLOSE does not repair it, and the
    damage outlives the session that caused it. `get_email_labels` SELECTs every
    label mailbox to search it and `remove_label` SELECTs the one it expunges
    from, neither of which issues CLOSE — CLOSE would expunge `\\Deleted` messages,
    which the scoped-expunge contract forbids. So a label those tools have touched
    can never be deleted here, and `delete_label` is exercised on a second label
    created after the last `get_email_labels` call and never selected by anything.
    A real IMAP server has no such restriction; SELECT+CLOSE and deleting a
    non-empty mailbox are both fine even in GreenMail.
    """
    _wait_until_ready()
    _ensure_empty_mailboxes(BOB, ["INBOX"])

    run_id = uuid.uuid4().hex[:8]
    # Read/remove half: the tools SELECT this one, so GreenMail will not delete it.
    flow_label = f"E2E-flow-{run_id}"
    flow_mailbox = f"Labels/{flow_label}"
    # Delete half: created last and never selected, so DELETE still works.
    shape_label = f"E2E-shape-{run_id}"
    shape_mailbox = f"Labels/{shape_label}"
    denied_label = f"E2E-denied-{run_id}"
    labelled_subject = f"mcp-e2e-apply-{run_id}"
    plain_subject = f"mcp-e2e-plain-{run_id}"

    def _label_subjects(mailbox: str) -> set[str]:
        """Read a label mailbox's subjects, always releasing it with CLOSE.

        Only the subjects are read: cluster C made message bodies HTML, and
        nothing here depends on the body.
        """
        with _imap_session(BOB) as client:
            status, _ = client.select(f'"{mailbox}"', readonly=True)
            assert status == "OK", f"SELECT {mailbox} failed: {status}"
            try:
                status, rows = client.uid("search", None, "ALL")
                assert status == "OK", f"UID SEARCH in {mailbox} failed: {status}"
                subjects: set[str] = set()
                for uid in (rows[0] or b"").split():
                    status, fetched = client.uid("fetch", uid, "(BODY.PEEK[HEADER])")
                    assert status == "OK", f"UID FETCH {uid!r} failed: {status}"
                    response = next((item for item in fetched if isinstance(item, tuple)), None)
                    assert response is not None
                    message = BytesParser(policy=policy.default).parsebytes(response[1])
                    subjects.add(str(message.get("Subject", "")))
                return subjects
            finally:
                client.close()

    config_path = tmp_path / "config.toml"
    config_path.write_text(f"enable_folder_management = true\n{CONFIG_TEMPLATE}")
    config_path.chmod(0o600)
    base_env = {key: value for key, value in os.environ.items() if not key.startswith("MCP_EMAIL_SERVER_")}
    server_env = {
        **base_env,
        "MCP_EMAIL_SERVER_CONFIG_PATH": str(config_path),
        "MCP_EMAIL_SERVER_CREDENTIAL_STORAGE": "plaintext",
        "MCP_EMAIL_SERVER_LOG_LEVEL": "WARNING",
    }
    console_script = Path(sys.executable).with_name("mcp-email-server")
    assert console_script.is_file(), f"Installed console script not found: {console_script}"

    def _server(env: dict[str, str]) -> StdioServerParameters:
        return StdioServerParameters(command=str(console_script), args=["stdio"], env=env, cwd=Path.cwd())

    try:
        _seed_message(labelled_subject, "This message gets a label")
        _seed_message(plain_subject, "This message stays unlabelled")
        _wait_for_message(BOB, "INBOX", labelled_subject)
        _wait_for_message(BOB, "INBOX", plain_subject)

        async with stdio_client(_server(server_env)) as (read_stream, write_stream):
            async with ClientSession(read_stream, write_stream, read_timeout_seconds=timedelta(seconds=15)) as session:
                await session.initialize()
                tool_names = {tool.name for tool in (await session.list_tools()).tools}
                assert {"create_label", "delete_label", "apply_label"} <= tool_names

                created = await _call_tool(session, "create_label", {"account_name": "bob", "label_name": flow_label})
                # The caller named a label, so the result names a label.
                assert created["result"] == f"Label '{flow_label}' created"
                assert "Labels/" not in created["result"]
                assert _mailbox_exists(BOB, flow_mailbox)

                # The read side recognizes the write side's mailbox as a label.
                labels = await _call_tool(session, "list_labels", {"account_name": "bob"})
                listed = [label for label in labels["result"] if label["name"] == flow_label]
                assert len(listed) == 1, labels
                assert listed[0]["full_path"] == flow_mailbox

                labelled_metadata = await _metadata_for_subject(session, "bob", labelled_subject)
                plain_metadata = await _metadata_for_subject(session, "bob", plain_subject)

                applied = await _call_tool(
                    session,
                    "apply_label",
                    {
                        "account_name": "bob",
                        "email_ids": [labelled_metadata["email_id"]],
                        "label_name": flow_label,
                        "source_mailbox": "INBOX",
                    },
                )
                assert applied["result"] == f"Successfully applied label '{flow_label}' to 1 email(s)"
                assert _label_subjects(flow_mailbox) == {labelled_subject}
                # Applying a label is additive: the original never moves.
                assert _find_message(BOB, "INBOX", labelled_subject) is not None

                confirmed = await _call_tool(
                    session,
                    "get_email_labels",
                    {"account_name": "bob", "email_id": labelled_metadata["email_id"], "mailbox": "INBOX"},
                )
                assert confirmed["result"] == [flow_label]

                untouched = await _call_tool(
                    session,
                    "get_email_labels",
                    {"account_name": "bob", "email_id": plain_metadata["email_id"], "mailbox": "INBOX"},
                )
                assert untouched["result"] == []

                removed = await _call_tool(
                    session,
                    "remove_label",
                    {
                        "account_name": "bob",
                        "email_ids": [labelled_metadata["email_id"]],
                        "label_name": flow_label,
                        "source_mailbox": "INBOX",
                    },
                )
                assert removed["result"] == f"Successfully removed label '{flow_label}' from 1 email(s)"
                assert _label_subjects(flow_mailbox) == set()
                assert _find_message(BOB, "INBOX", labelled_subject) is not None

                after_removal = await _call_tool(
                    session,
                    "get_email_labels",
                    {"account_name": "bob", "email_id": labelled_metadata["email_id"], "mailbox": "INBOX"},
                )
                assert after_removal["result"] == []

                # No `get_email_labels` call past this point: the delete target must
                # never be SELECTed. `list_labels` uses LIST and stays safe.
                shaped = await _call_tool(session, "create_label", {"account_name": "bob", "label_name": shape_label})
                assert shaped["result"] == f"Label '{shape_label}' created"
                assert _mailbox_exists(BOB, shape_mailbox)

                relabelled = await _call_tool(
                    session,
                    "apply_label",
                    {
                        "account_name": "bob",
                        "email_ids": [labelled_metadata["email_id"]],
                        "label_name": shape_label,
                    },
                )
                # source_mailbox defaults to INBOX.
                assert relabelled["result"] == f"Successfully applied label '{shape_label}' to 1 email(s)"
                assert _label_subjects(shape_mailbox) == {labelled_subject}

                dropped = await _call_tool(session, "delete_label", {"account_name": "bob", "label_name": shape_label})
                assert dropped["result"] == f"Label '{shape_label}' deleted"
                assert "Labels/" not in dropped["result"]
                assert not _mailbox_exists(BOB, shape_mailbox)

                # The label is gone from the read side too, and deleting a label
                # that held a copy never touched the messages themselves.
                remaining = await _call_tool(session, "list_labels", {"account_name": "bob"})
                assert not any(label["name"] == shape_label for label in remaining["result"])
                assert _find_message(BOB, "INBOX", labelled_subject) is not None
                assert _find_message(BOB, "INBOX", plain_subject) is not None

        # With the policy off the two shape tools refuse, while apply_label — which
        # only copies messages into an existing mailbox — keeps working.
        gated_off_env = {**server_env, "MCP_EMAIL_SERVER_ENABLE_FOLDER_MANAGEMENT": "false"}
        async with stdio_client(_server(gated_off_env)) as (read_stream, write_stream):
            async with ClientSession(read_stream, write_stream, read_timeout_seconds=timedelta(seconds=15)) as session:
                await session.initialize()
                tool_names = {tool.name for tool in (await session.list_tools()).tools}
                # The tools stay visible; the denial is a runtime policy decision.
                assert {"create_label", "delete_label", "apply_label"} <= tool_names

                for tool_name in ("create_label", "delete_label"):
                    denied = await session.call_tool(
                        tool_name,
                        arguments={"account_name": "bob", "label_name": denied_label},
                    )
                    assert denied.isError is True
                    assert "Folder management is disabled" in _text_content(denied)
                assert not _mailbox_exists(BOB, f"Labels/{denied_label}")

                still_labelled = await _metadata_for_subject(session, "bob", plain_subject)
                ungated = await _call_tool(
                    session,
                    "apply_label",
                    {
                        "account_name": "bob",
                        "email_ids": [still_labelled["email_id"]],
                        "label_name": flow_label,
                    },
                )
                assert ungated["result"] == f"Successfully applied label '{flow_label}' to 1 email(s)"
                assert _label_subjects(flow_mailbox) == {plain_subject}
    finally:
        for mailbox in (flow_mailbox, shape_mailbox, f"Labels/{denied_label}"):
            with contextlib.suppress(Exception), _imap_session(BOB) as client:
                # Release the mailbox before removing it. `flow_mailbox` was SELECTed
                # without CLOSE by the server's own read tools, so GreenMail may
                # refuse it outright; a leftover run-scoped label is harmless.
                if client.select(f'"{mailbox}"')[0] == "OK":
                    client.close()
                client.delete(f'"{mailbox}"')


# port-slot B2: label writes (create_label, delete_label, apply_label)
