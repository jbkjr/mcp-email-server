"""Prove GreenMail actually implements the mailbox verbs later E2E tests depend on.

CREATE, COPY, RENAME, and DELETE are mandatory in RFC 3501, but a test double is
only as conformant as its authors made it. Folder- and label-management coverage
is about to be written against these commands, so this probe fails loudly and
early rather than letting a missing verb surface as a confusing failure inside a
much larger test.

Ports are read inside the test body, never at import time: `dev/greenmail/run-e2e.sh`
exports the mapped ports into the pytest process, and collection of this module
must not depend on that having happened.
"""

from __future__ import annotations

import contextlib
import imaplib
import os
import uuid
from collections.abc import Iterator

import pytest

pytestmark = pytest.mark.e2e

IMAP_HOST = "127.0.0.1"
ALICE = ("alice@example.test", "alice-password")


def _imap_port() -> int:
    return int(os.environ.get("MCP_EMAIL_SERVER_E2E_IMAP_PORT", "3143"))


@contextlib.contextmanager
def _imap_session(credentials: tuple[str, str]) -> Iterator[imaplib.IMAP4]:
    client = imaplib.IMAP4(IMAP_HOST, _imap_port(), timeout=5)
    try:
        status, _ = client.login(*credentials)
        assert status == "OK"
        yield client
    finally:
        with contextlib.suppress(Exception):
            client.logout()


def _mailbox_names(client: imaplib.IMAP4) -> set[str]:
    status, rows = client.list()
    assert status == "OK", f"LIST failed: {status}"
    names: set[str] = set()
    for row in rows:
        if not isinstance(row, bytes):
            continue
        # LIST rows are `(flags) "delimiter" name`; the name is the final token and
        # is unquoted here because every probe mailbox is a plain ASCII atom.
        name = row.rsplit(b" ", 1)[-1].strip().strip(b'"').decode()
        if name:
            names.add(name)
    return names


def _hierarchy_delimiter(client: imaplib.IMAP4) -> str:
    # imaplib splices arguments in verbatim, so the empty reference name must
    # arrive already quoted or the server sees a missing argument.
    status, rows = client.list('""', "INBOX")
    assert status == "OK", f"LIST INBOX failed: {status}"
    for row in rows:
        if isinstance(row, bytes) and b'"' in row:
            # `(flags) "/" "INBOX"` — the delimiter is the first quoted token.
            return row.split(b'"')[1].decode()
    pytest.fail(f"No hierarchy delimiter reported for INBOX: {rows!r}")


def test_greenmail_supports_message_id_header_search() -> None:
    """Locating one message by `Message-ID` is how the label tools will find it across folders.

    GreenMail's SEARCH parser rejects an unquoted multi-word value with a blanket
    `BAD ... Search command not supported`, which reads like a missing feature but
    is only a quoting error. Assert the quoted form so a future regression is not
    mistaken for that.
    """
    message_id = f"<probe-{uuid.uuid4().hex}@example.test>"

    with _imap_session(ALICE) as client:
        status, _ = client.append(
            "INBOX",
            None,
            None,
            (
                f"Subject: message-id search probe\r\n"
                f"Message-ID: {message_id}\r\n"
                f"From: alice@example.test\r\n\r\nprobe\r\n"
            ).encode(),
        )
        assert status == "OK", f"APPEND failed: {status}"

        status, _ = client.select("INBOX")
        assert status == "OK", f"SELECT INBOX failed: {status}"

        status, rows = client.uid("search", None, "ALL")
        assert status == "OK", f"UID SEARCH ALL failed: {status}"

        status, rows = client.uid("search", None, "HEADER", "Message-ID", f'"{message_id}"')
        assert status == "OK", f"UID SEARCH HEADER Message-ID failed: {status}"
        assert len((rows[0] or b"").split()) == 1, f"Message-ID search matched {rows!r}, expected one message"


def test_greenmail_supports_create_copy_rename_and_delete() -> None:
    """One round trip through every mailbox verb the folder/label tools will use."""
    suffix = uuid.uuid4().hex[:8]
    probe_a = f"Probe-A-{suffix}"
    probe_b = f"Probe-B-{suffix}"

    with _imap_session(ALICE) as client:
        delimiter = _hierarchy_delimiter(client)
        print(f"GreenMail hierarchy delimiter: {delimiter!r}")

        try:
            status, _ = client.create(probe_a)
            assert status == "OK", f"CREATE {probe_a} failed: {status}"
            assert probe_a in _mailbox_names(client), f"CREATE {probe_a} did not appear in LIST"

            status, _ = client.select("INBOX")
            assert status == "OK", f"SELECT INBOX failed: {status}"
            status, _ = client.append(
                "INBOX",
                None,
                None,
                b"Subject: folder capability probe\r\nFrom: alice@example.test\r\n\r\nprobe\r\n",
            )
            assert status == "OK", f"APPEND to INBOX failed: {status}"

            status, _ = client.select("INBOX")
            assert status == "OK"
            status, rows = client.uid("search", None, "ALL")
            assert status == "OK", f"UID SEARCH ALL failed: {status}"
            uids = (rows[0] or b"").split()
            assert uids, "The probe message was not searchable in INBOX"
            # UIDs increase monotonically, so the newest is the message just appended
            # regardless of what else already sits in INBOX.
            uid = max(int(value) for value in uids)

            status, _ = client.uid("copy", str(uid), probe_a)
            assert status == "OK", f"UID COPY to {probe_a} failed: {status}"
            status, rows = client.select(probe_a)
            assert status == "OK", f"SELECT {probe_a} failed: {status}"
            assert int(rows[0]) == 1, f"COPY did not land exactly one message in {probe_a}: {rows!r}"

            # RENAME of the selected mailbox is legal but leaves the session in an
            # ambiguous state; close first so the failure below can only mean RENAME.
            status, _ = client.close()
            assert status == "OK", f"CLOSE failed: {status}"

            status, _ = client.rename(probe_a, probe_b)
            assert status == "OK", f"RENAME {probe_a} -> {probe_b} failed: {status}"
            names = _mailbox_names(client)
            assert probe_b in names, f"RENAME target {probe_b} is missing from LIST"
            assert probe_a not in names, f"RENAME left the source {probe_a} behind"

            status, rows = client.select(probe_b)
            assert status == "OK", f"SELECT {probe_b} failed: {status}"
            assert int(rows[0]) == 1, f"RENAME did not carry the message into {probe_b}: {rows!r}"
            status, _ = client.close()
            assert status == "OK"

            status, _ = client.delete(probe_b)
            assert status == "OK", f"DELETE {probe_b} failed: {status}"
            assert probe_b not in _mailbox_names(client), f"DELETE left {probe_b} in LIST"
        finally:
            with contextlib.suppress(Exception):
                client.close()
            for leftover in (probe_a, probe_b):
                with contextlib.suppress(Exception):
                    client.delete(leftover)
