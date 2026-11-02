"""Synthetic Gmail worker and optional real gateway-proxy session integration.

No sockets, account configuration, keyring, or provider network connections.
Set MCP_EMAIL_TEST_PROXY to an imap_proxy.py checkout to run the cross-repo gate.
"""

from __future__ import annotations

import asyncio
import importlib.util
import json
import os
import re
import sys
from email import message_from_bytes
from email.mime.text import MIMEText
from importlib.resources import files
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock

import pytest

from mcp_email_server.adapters.mutations import ClassicMutationProvider
from mcp_email_server.adapters.reads import ClassicReadProvider
from mcp_email_server.application.mutations import (
    ApplyLabelCommand,
    ApplyLabelService,
    ArchiveCommand,
    ArchiveService,
    CopyService,
    MoveService,
    MutationAccountSnapshot,
    MutationProviderAccess,
    RemoveLabelCommand,
    RemoveLabelService,
)
from mcp_email_server.application.reads import (
    EmailLabelService,
    GetEmailLabelsQuery,
    LabelDiscoveryService,
    ListLabelsQuery,
    MailboxDiscoveryService,
    ReadAccountSnapshot,
    ReadProviderAccess,
)
from mcp_email_server.config import EmailServer, EmailSettings
from mcp_email_server.emails.classic import EmailClient, _quote_mailbox, decode_mailbox_name
from mcp_email_server.emails.gmail import parse_metadata

DRAFTS = "[Gmail]/Brouillons é"
TRASH = "[Gmail]/Corbeille é"
MSGID = "18446744073709551614"
BOXES = {
    "INBOX": (),
    "work": (),
    "keep": (),
    "Labels/exact": (),
    'a "quote" (x)': (),
    DRAFTS: (r"\Drafts",),
    TRASH: (r"\Trash",),
    "All": (r"\All",),
    "Spam": (r"\Junk",),
}


class FakeIMAP:
    """In-memory mailbox views with different UIDs sharing one Gmail identity."""

    def __init__(self):
        self.protocol = SimpleNamespace(capabilities={"IMAP4rev1", "X-GM-EXT-1", "MOVE", "UIDPLUS"})
        self.protocol.capability = self.capability
        self.selected = None
        self.labels = {r"\Inbox", "work", "keep"}
        self.aliases = {"INBOX": "7", "work": "99", "keep": "8", "Labels/exact": "9", "All": "12", "Spam": "13"}
        self.commands = []
        self.validity = "42"
        self.identity = MSGID
        self.body = b"Message-ID: <fixture@example.invalid>\r\nSubject: fixture\r\nFrom: sender@example.invalid\r\nDate: Fri, 30 Oct 2026 10:00:00 +0000\r\n\r\nbody\r\n"
        self.draft = None
        self.draft_labels = {r"\Drafts"}
        self.fail_remove = False
        self.lose_store = False
        self.lose_append = False
        self.replace_identity = False
        self.auto_remove = True
        self.writes = []
        self.no_op_own_removals = 0
        self.ignore_removals = False
        self.all_mailbox = "All"

    async def login(self, *_):
        return "OK", []

    async def logout(self):
        self.commands.append(("LOGOUT",))
        return "OK", []

    async def capability(self):
        return "OK", [" ".join(self.protocol.capabilities).encode()]

    async def list(self, *_):
        self.commands.append(("LIST",))
        return "OK", [f'({" ".join(flags)}) "/" {_quote_mailbox(name)}'.encode() for name, flags in BOXES.items()]

    async def select(self, mailbox):
        self.selected = decode_mailbox_name(mailbox[1:-1].replace(r"\"", '"').replace("\\\\", "\\"))
        self.commands.append(("SELECT", mailbox))
        return "OK", [b"1 EXISTS", f"OK [UIDVALIDITY {self.validity}] selected".encode()]

    def metadata(self, uid):
        if self.selected == DRAFTS:
            if self.draft is None or uid != "11":
                return None
            identity, labels = "55", self.draft_labels
        elif self.aliases.get(self.selected) == uid:
            identity, labels = self.identity, self.labels
            system = {label.upper() for label in labels}
            if self.selected == "INBOX" and r"\INBOX" not in system:
                return None
            if self.selected == "Spam" and r"\SPAM" not in system:
                return None
            if self.selected == self.all_mailbox and system & {r"\SPAM", r"\TRASH", r"\DRAFTS"}:
                return None
            if self.selected not in {"INBOX", "Spam", self.all_mailbox} and self.selected not in labels:
                return None
            own_label = {"INBOX": r"\INBOX", "Spam": r"\SPAM"}.get(self.selected, self.selected)
            labels = {label for label in labels if (label.upper() if label.startswith("\\") else label) != own_label}
        else:
            return None
        label_list = " ".join(_quote_mailbox(label) for label in sorted(labels))
        return f"1 FETCH (UID {uid} X-GM-MSGID {identity} X-GM-LABELS ({label_list}))".encode()

    async def uid_search(self, *criteria, charset=None):
        self.commands.append(("SEARCH", *criteria))
        if "HEADER" in criteria and "Message-ID" in criteria:
            message = self.draft if self.selected == DRAFTS else self.body
            searched = str(criteria[-1]).strip('"')
            found = message is not None and searched in message.decode()
            return "OK", [
                b"11"
                if found and self.selected == DRAFTS
                else self.aliases.get(self.selected, "").encode()
                if found
                else b""
            ]
        return "OK", [self.aliases.get(self.selected, "").encode()]

    async def uid(self, operation, uid, *args):  # noqa: C901 - fake protocol dispatch
        self.commands.append((operation.upper(), uid, *args))
        if operation.lower() == "search":
            return await self.uid_search(uid, *args)
        if operation.lower() == "fetch":
            fields = args[0].upper()
            data = []
            for item in uid.split(","):
                meta = self.metadata(item)
                if meta is None:
                    continue
                if "X-GM-MSGID" in fields:
                    data.append(meta)
                elif "RFC822.SIZE" in fields:
                    data.append(f"1 FETCH (UID {item} RFC822.SIZE {len(self.body)})".encode())
                elif "INTERNALDATE" in fields:
                    data.append(f'1 FETCH (UID {item} INTERNALDATE "30-Oct-2026 10:00:00 +0000")'.encode())
                elif "BODY.PEEK" in fields:
                    raw = self.draft if self.selected == DRAFTS else self.body
                    section = "HEADER" if "HEADER" in fields else ""
                    data.extend([f"1 FETCH (UID {item} BODY[{section}] {{{len(raw)}}}".encode(), bytearray(raw), b")"])
            return "OK", data
        if operation.lower() == "store":
            self.writes.append((self.selected, uid, *args))
            label = args[1][1:-1]
            if label.startswith('"'):
                label = label[1:-1].replace(r"\"", '"').replace("\\\\", "\\")
            label = decode_mailbox_name(label)
            if self.fail_remove and args[0] == "-X-GM-LABELS":
                return "NO", [b"source retained"]
            if args[0] == "+X-GM-LABELS":
                self.labels.add(label)
                if self.selected == "Spam" and label == r"\Inbox" and self.auto_remove:
                    self.labels.discard(r"\Spam")
                if self.replace_identity:
                    self.identity = "56"
            else:
                own_label = {"INBOX": r"\Inbox", "Spam": r"\Spam"}.get(self.selected, self.selected)
                if label.upper() == own_label.upper() if label.startswith("\\") else label == own_label:
                    self.no_op_own_removals += 1
                elif not self.ignore_removals:
                    self.labels = {
                        item
                        for item in self.labels
                        if (item.upper() != label.upper() if label.startswith("\\") else item != label)
                    }
            if self.lose_store:
                raise ConnectionError("lost completion")
            return "OK", []
        if operation.lower() == "move":
            return "OK", []
        raise AssertionError(f"Unexpected Gmail command: {operation}")

    async def append(self, payload, mailbox, flags):
        self.commands.append(("APPEND", mailbox, flags, payload))
        self.draft = payload
        if self.lose_append:
            raise ConnectionError("lost completion")
        return "OK", [b"[APPENDUID 42 11] saved"]


class SpamRescueIMAP(FakeIMAP):
    """Model Gmail removing Spam membership as soon as Inbox is added."""

    def __init__(self, *, auto_remove=True, inbox_label=r"\INBOX"):
        super().__init__()
        self.labels = {r"\Spam", "keep"}
        self.auto_remove = auto_remove
        self.inbox_label = inbox_label
        self.inbox_size = 20_001

    async def uid_search(self, *criteria, charset=None):
        if criteria == ("ALL",) and self.selected == "INBOX":
            self.commands.append(("SEARCH", *criteria))
            return "OK", [b" ".join(str(uid).encode() for uid in range(1, self.inbox_size + 1))]
        return await super().uid_search(*criteria, charset=charset)

    def metadata(self, uid):
        labels = {label.upper() for label in self.labels}
        if (self.selected == "Spam" and r"\SPAM" not in labels) or (
            self.selected == "INBOX" and r"\INBOX" not in labels
        ):
            return None
        return super().metadata(uid)

    async def uid(self, operation, uid, *args):
        response = await super().uid(operation, uid, *args)
        if operation.lower() == "store" and args[0] == "+X-GM-LABELS" and r"\Inbox" in self.labels:
            self.labels.discard(r"\Inbox")
            self.labels.add(self.inbox_label)
            if self.auto_remove:
                self.labels.discard(r"\Spam")
        return response


def client_for(monkeypatch, imap, provider="gmail_imap"):
    server = EmailServer(user_name="synthetic", password="synthetic", host="imap.gmail.com", port=993, use_ssl=False)
    client = EmailClient(server, gateway_provider=provider)
    monkeypatch.setattr(client, "_connect_imap", AsyncMock(return_value=imap))
    monkeypatch.setattr(client, "_connect_imap_server", AsyncMock(return_value=imap))
    return client


@pytest.mark.asyncio
@pytest.mark.parametrize("inbox_label", [r"\Inbox", r"\INBOX"])
async def test_archive_incremental_and_reject_non_inbox(monkeypatch, inbox_label):
    fake = FakeIMAP()
    fake.labels = {inbox_label, "work", "keep"}
    client = client_for(monkeypatch, fake)
    outcome = await client._gmail_mutate(["7"], "INBOX", None, "archive")
    assert outcome.targets("succeeded") == ["7"]
    assert fake.writes == [("All", "12", "-X-GM-LABELS", r"(\Inbox)")]
    assert fake.no_op_own_removals == 0
    assert fake.labels == {"work", "keep"}
    with pytest.raises(ValueError, match="INBOX"):
        await client._gmail_mutate(["99"], "work", None, "archive")


@pytest.mark.asyncio
async def test_fake_models_hidden_selected_label_and_no_op_ok():
    fake = FakeIMAP()
    await fake.select('"INBOX"')
    assert r"\Inbox" not in parse_metadata([fake.metadata("7")])["7"]["gmail_labels"]
    assert (await fake.uid("store", "7", "-X-GM-LABELS", r"(\Inbox)"))[0] == "OK"
    assert r"\Inbox" in fake.labels
    await fake.select('"work"')
    assert "work" not in parse_metadata([fake.metadata("99")])["99"]["gmail_labels"]
    assert (await fake.uid("store", "99", "-X-GM-LABELS", '("work")'))[0] == "OK"
    assert "work" in fake.labels


@pytest.mark.asyncio
async def test_archive_without_user_labels_can_be_restored_from_all_mail(monkeypatch):
    fake = FakeIMAP()
    fake.labels = {r"\Inbox"}
    client = client_for(monkeypatch, fake)

    archived = await client._gmail_mutate(["7"], "INBOX", None, "archive")

    assert archived.targets("succeeded") == ["7"]
    assert fake.labels == set()
    assert fake.metadata("7") is None  # Source selection was restored, but its UID left.
    restored = await client.move_emails_with_outcome(["12"], "All", "INBOX")
    assert restored.targets("succeeded") == ["12"]
    assert fake.labels == {r"\Inbox"}
    assert fake.writes == [
        ("All", "12", "-X-GM-LABELS", r"(\Inbox)"),
        ("All", "12", "+X-GM-LABELS", r"(\Inbox)"),
    ]
    with pytest.raises(ValueError, match="restore to INBOX only"):
        await client.move_emails_with_outcome(["12"], "All", "keep")


@pytest.mark.asyncio
async def test_archive_resolves_localized_all_mail_from_list_flags(monkeypatch):
    localized = "[Gmail]/Tous les messages é"
    monkeypatch.delitem(BOXES, "All")
    monkeypatch.setitem(BOXES, localized, (r"\All",))
    fake = FakeIMAP()
    fake.all_mailbox = localized
    fake.aliases[localized] = "50192"
    client = client_for(monkeypatch, fake)

    result = await client._gmail_mutate(["7"], "INBOX", None, "archive")

    assert result.targets("succeeded") == ["7"]
    assert fake.writes == [(localized, "50192", "-X-GM-LABELS", r"(\Inbox)")]
    assert fake.labels == {"work", "keep"}


@pytest.mark.asyncio
@pytest.mark.parametrize("failure", ["no_op_ok", "lost_store", "post_fetch", "post_identity"])
async def test_archive_requires_post_condition_and_never_replays(monkeypatch, failure):
    fake = FakeIMAP()
    fake.ignore_removals = failure == "no_op_ok"
    fake.lose_store = failure == "lost_store"
    original = fake.metadata

    def metadata(uid):
        if fake.selected == "All" and fake.writes:
            if failure == "post_fetch":
                raise ConnectionError("lost confirmation")
            if failure == "post_identity":
                fake.identity = "56"
        return original(uid)

    fake.metadata = metadata
    client = client_for(monkeypatch, fake)

    result = await client._gmail_mutate(["7", "8"], "INBOX", None, "archive")

    assert result.targets("unknown") == ["7"]
    assert result.outcomes[1].detail == "not-attempted"
    assert result.reconciliation_needed
    assert fake.writes == [("All", "12", "-X-GM-LABELS", r"(\Inbox)")]
    if failure == "no_op_ok":
        assert result.outcomes[0].detail == "label-removal-unconfirmed"
        assert r"\Inbox" in fake.labels


@pytest.mark.asyncio
@pytest.mark.parametrize("after_write", [False, True])
async def test_archive_cancellation_stops_batch_without_more_writes(monkeypatch, after_write):
    fake = FakeIMAP()
    client = client_for(monkeypatch, fake)
    if after_write:
        original = fake.uid

        async def uid(operation, uid, *args):
            result = await original(operation, uid, *args)
            if operation == "store":
                raise asyncio.CancelledError
            return result

        fake.uid = uid
    else:
        monkeypatch.setattr(client, "_batch_fetch_message_ids", AsyncMock(side_effect=asyncio.CancelledError))

    result = await client._gmail_mutate(["7", "8"], "INBOX", None, "archive")

    assert result.outcomes[0].status == ("unknown" if after_write else "failed")
    assert result.outcomes[0].detail == "provider-cancelled"
    assert result.outcomes[1].detail == "not-attempted"
    assert len(fake.writes) == int(after_write)


@pytest.mark.asyncio
@pytest.mark.parametrize("failure", ["message_id", "all_identity", "no_hits", "too_many_hits", "no_all"])
async def test_archive_refuses_without_all_mail_identity_before_write(monkeypatch, failure):
    fake = FakeIMAP()
    if failure == "message_id":
        fake.body = b"Subject: fixture\r\n\r\nbody\r\n"
    elif failure == "all_identity":
        original = fake.metadata

        def metadata(uid):
            data = original(uid)
            return data.replace(MSGID.encode(), b"55") if fake.selected == "All" and data else data

        fake.metadata = metadata
    elif failure == "no_hits":
        fake.uid_search = AsyncMock(return_value=("OK", [b""]))
    elif failure == "too_many_hits":
        fake.uid_search = AsyncMock(return_value=("OK", [b"1 2 3 4 5 12"]))
    else:
        monkeypatch.setitem(BOXES, "All", ())
    client = client_for(monkeypatch, fake)

    result = await client._gmail_mutate(["7"], "INBOX", None, "archive")

    assert result.targets("failed") == ["7"]
    assert not fake.writes
    assert r"\Inbox" in fake.labels


@pytest.mark.asyncio
@pytest.mark.parametrize(
    "source,uid,destination,removed",
    [
        ("INBOX", "7", "keep", r"\Inbox"),
        ("work", "99", "keep", "work"),
        ("work", "99", "INBOX", "work"),
    ],
)
async def test_move_removes_source_in_all_mail_and_preserves_other_labels(
    monkeypatch, source, uid, destination, removed
):
    fake = FakeIMAP()
    client = client_for(monkeypatch, fake)

    result = await client.move_emails_with_outcome([uid], source, destination)

    assert result.targets("succeeded") == [uid]
    assert fake.labels == {r"\Inbox", "work", "keep"} - {removed}
    assert fake.writes[-1][0:3] == ("All", "12", "-X-GM-LABELS")
    assert fake.no_op_own_removals == 0


@pytest.mark.asyncio
@pytest.mark.parametrize("source,uid,label", [("work", "99", "work"), ("INBOX", "7", "keep"), ("All", "12", "work")])
async def test_remove_label_in_all_mail_confirms_absence(monkeypatch, source, uid, label):
    fake = FakeIMAP()
    client = client_for(monkeypatch, fake)

    result = await client.remove_label_with_outcome([uid], label, source)

    assert result.targets("succeeded") == [uid]
    assert fake.labels == {r"\Inbox", "work", "keep"} - {label}
    assert fake.writes == [("All", "12", "-X-GM-LABELS", f'("{label}")')]
    assert fake.no_op_own_removals == 0


@pytest.mark.asyncio
@pytest.mark.parametrize("operation", ["move", "remove"])
async def test_move_and_remove_never_accept_no_op_ok(monkeypatch, operation):
    fake = FakeIMAP()
    fake.ignore_removals = True
    client = client_for(monkeypatch, fake)

    result = await client._gmail_mutate(["99"], "work", "keep" if operation == "move" else "work", operation)

    assert result.targets("unknown") == ["99"]
    assert "work" in fake.labels
    assert len(fake.writes) == (2 if operation == "move" else 1)
    assert fake.writes[-1][0:3] == ("All", "12", "-X-GM-LABELS")


@pytest.mark.asyncio
async def test_move_never_removes_a_same_message_id_different_all_mail_identity(monkeypatch):
    fake = FakeIMAP()
    original = fake.metadata

    def metadata(uid):
        data = original(uid)
        return data.replace(MSGID.encode(), b"55") if fake.selected == "All" and data else data

    fake.metadata = metadata
    client = client_for(monkeypatch, fake)

    result = await client.move_emails_with_outcome(["99"], "work", "INBOX")

    assert result.targets("unknown") == ["99"]
    assert fake.writes == [("work", "99", "+X-GM-LABELS", r"(\Inbox)")]
    assert "work" in fake.labels


@pytest.mark.asyncio
async def test_all_mail_restore_requires_inbox_post_condition(monkeypatch):
    fake = FakeIMAP()
    fake.labels = {"keep"}
    original = fake.metadata

    def metadata(uid):
        if fake.writes:
            fake.labels.discard(r"\Inbox")
        return original(uid)

    fake.metadata = metadata
    client = client_for(monkeypatch, fake)

    result = await client.move_emails_with_outcome(["12"], "All", "INBOX")

    assert result.targets("unknown") == ["12"]
    assert fake.writes == [("All", "12", "+X-GM-LABELS", r"(\Inbox)")]


@pytest.mark.asyncio
async def test_next_uid_is_read_in_restored_source_view(monkeypatch):
    fake = FakeIMAP()
    client = client_for(monkeypatch, fake)

    result = await client.move_emails_with_outcome(["99", "12"], "work", "INBOX")

    assert result.targets("succeeded") == ["99"]
    assert result.targets("failed") == ["12"]  # All Mail's UID is not a source UID.
    assert len(fake.writes) == 2


@pytest.mark.asyncio
async def test_literal_labels_prefix_move_uid_one_through_application(monkeypatch):
    import mcp_email_server.app as app

    label = "Labels/gateway-test-b"
    monkeypatch.setitem(BOXES, label, ())
    fake = FakeIMAP()
    fake.aliases[label] = "1"
    fake.labels = {label, "keep"}
    client = client_for(monkeypatch, fake)
    provider = ClassicMutationProvider(SimpleNamespace(incoming_client=client))
    account = MutationAccountSnapshot("test", "legacy", (), (), True, False, gateway_provider="gmail_imap")
    authority = SimpleNamespace(resolve=lambda *args, **kwargs: account)
    factory = SimpleNamespace(open=lambda *args, **kwargs: MutationProviderAccess(account, provider))
    projection = SimpleNamespace(invalidate=AsyncMock())
    service = MoveService(authority, factory, SimpleNamespace(open=AsyncMock(return_value=projection)))
    monkeypatch.setattr(app, "move_emails_command", service.execute)

    summary = await app.move_emails("test", ["1"], "INBOX", label)

    assert summary == "Successfully moved 1 email(s) to INBOX"
    assert fake.labels == {r"\Inbox", "keep"}
    assert fake.writes[-1] == ("All", "12", "-X-GM-LABELS", '("Labels/gateway-test-b")')
    projection.invalidate.assert_awaited_once_with((label, "INBOX"))


@pytest.mark.asyncio
async def test_move_confirmation_partial_and_identity_change(monkeypatch):
    for failure in (None, "fail_remove", "replace_identity", "lose_store"):
        fake = FakeIMAP()
        if failure:
            setattr(fake, failure, True)
        client = client_for(monkeypatch, fake)
        result = await client.move_emails_with_outcome(["7"], "INBOX", "keep")
        assert result.outcomes[0].status == ("succeeded" if failure is None else "unknown")
        writes = [c for c in fake.commands if c[0] == "STORE"]
        assert writes[0] == ("STORE", "7", "+X-GM-LABELS", '("keep")')
        assert len(writes) == (1 if failure in {"replace_identity", "lose_store"} else 2)
        assert "work" in fake.labels and "keep" in fake.labels
        if failure == "fail_remove":
            assert result.outcomes[0].detail == "destination-added-source-retained"
            assert r"\Inbox" in fake.labels


@pytest.mark.asyncio
async def test_alias_uid_is_mailbox_local_and_protected_overlap(monkeypatch):
    fake = FakeIMAP()
    client = client_for(monkeypatch, fake)
    assert (await client._gmail_mutate(["99"], "work", "keep", "remove")).targets("succeeded") == ["99"]
    assert r"\Inbox" in fake.labels
    assert fake.writes == [("All", "12", "-X-GM-LABELS", '("keep")')]
    fake.labels.add(r"\Trash")
    result = await client.move_emails_with_outcome(["7"], "INBOX", "keep")
    assert result.outcomes[0].detail == "protected-label-overlap"
    fake.aliases["INBOX"] = "17"
    result = await client.move_emails_with_outcome(["7"], "INBOX", "keep")
    assert result.outcomes[0].status == "failed"


@pytest.mark.asyncio
async def test_exact_labels_and_spam_rescue(monkeypatch):
    fake = FakeIMAP()
    client = client_for(monkeypatch, fake)
    for label in ["keep", "Labels/exact", 'a "quote" (x)']:
        assert (await client.copy_emails_with_outcome(["7"], "INBOX", label)).targets("succeeded") == ["7"]
        assert label in fake.labels
    for label in [r"\Trash", "unknown", TRASH]:
        with pytest.raises(ValueError):
            await client.copy_emails_with_outcome(["7"], "INBOX", label)
    fake.labels = {r"\Spam", "keep"}
    assert (await client.move_emails_with_outcome(["13"], "Spam", "INBOX")).targets("succeeded") == ["13"]
    assert fake.labels == {r"\Inbox", "keep"}
    with pytest.raises(ValueError):
        await client.move_emails_with_outcome(["13"], "Spam", "keep")


@pytest.mark.asyncio
@pytest.mark.parametrize("auto_remove", [True, False])
@pytest.mark.parametrize("inbox_label", [r"\Inbox", r"\INBOX"])
async def test_spam_rescue_confirms_destination_identity_and_restores_selection(monkeypatch, auto_remove, inbox_label):
    fake = SpamRescueIMAP(auto_remove=auto_remove, inbox_label=inbox_label)
    client = client_for(monkeypatch, fake)

    result = await client.move_emails_with_outcome(["13"], "Spam", "INBOX")

    assert result.targets("succeeded" if auto_remove else "unknown") == ["13"]
    assert result.reconciliation_needed is not auto_remove
    assert fake.labels == {inbox_label, "keep"} | (set() if auto_remove else {r"\Spam"})
    assert ("FETCH", "7", "(UID X-GM-MSGID X-GM-LABELS)") in fake.commands
    assert fake.selected == "Spam"
    assert ("SEARCH", "HEADER", "Message-ID", '"<fixture@example.invalid>"') in fake.commands
    assert ("SEARCH", "ALL") not in fake.commands
    header_fetch = next(c for c in fake.commands if c[0] == "FETCH" and "HEADER.FIELDS (MESSAGE-ID)" in c[2])
    assert header_fetch[1] == "13"
    assert fake.commands.index(header_fetch) < next(i for i, c in enumerate(fake.commands) if c[0] == "STORE")
    writes = [c for c in fake.commands if c[0] == "STORE"]
    assert writes == [("STORE", "13", "+X-GM-LABELS", r"(\Inbox)")]


@pytest.mark.asyncio
async def test_spam_rescue_checks_only_bounded_message_id_hits(monkeypatch):
    fake = SpamRescueIMAP()
    fake.uid_search = AsyncMock(return_value=("OK", [b"1 2 3 4 7"]))
    original = fake.metadata

    def metadata(uid):
        if fake.selected == "INBOX" and uid in {"1", "2", "3", "4"}:
            return f"1 FETCH (UID {uid} X-GM-MSGID 55 X-GM-LABELS (\\INBOX))".encode()
        return original(uid)

    fake.metadata = metadata
    client = client_for(monkeypatch, fake)

    result = await client.move_emails_with_outcome(["13"], "Spam", "INBOX")

    assert result.targets("succeeded") == ["13"]
    fetches = [c[1] for c in fake.commands if c[0] == "FETCH"]
    assert fetches == ["13", "13", "1,2,3,4,7"]
    fake.uid_search.assert_awaited_once_with("HEADER", "Message-ID", '"<fixture@example.invalid>"', charset=None)


@pytest.mark.asyncio
@pytest.mark.parametrize(
    "failure",
    ["same_message_id_different_gmail_id", "membership", "search", "no_hits", "oversized", "lost_store"],
)
async def test_spam_rescue_requires_real_evidence_and_never_replays(monkeypatch, failure):
    fake = SpamRescueIMAP()
    if failure == "same_message_id_different_gmail_id":
        fake.replace_identity = True
    elif failure == "membership":
        fake.inbox_label = "unrelated"
    elif failure == "search":
        fake.uid_search = AsyncMock(return_value=("NO", [b"rejected"]))
    elif failure == "no_hits":
        fake.uid_search = AsyncMock(return_value=("OK", [b""]))
    elif failure == "oversized":
        fake.uid_search = AsyncMock(return_value=("OK", [b"1 2 3 4 5 7"]))
    elif failure == "lost_store":
        fake.lose_store = True
    client = client_for(monkeypatch, fake)

    result = await client.move_emails_with_outcome(["13", "14"], "Spam", "INBOX")

    assert result.targets("unknown") == ["13"]
    assert result.outcomes[1].detail == "not-attempted"
    assert result.reconciliation_needed
    assert len([c for c in fake.commands if c[0] == "STORE"]) == 1
    assert ("SEARCH", "ALL") not in fake.commands
    if failure == "same_message_id_different_gmail_id":
        assert ("SEARCH", "HEADER", "Message-ID", '"<fixture@example.invalid>"') in fake.commands
    if failure == "oversized":
        assert not any(c[0] == "FETCH" and c[1] != "13" for c in fake.commands)


@pytest.mark.asyncio
@pytest.mark.parametrize("header", [b"", b"Message-ID: \r\n", b"Message-ID: <non-ascii-\xc3\xa9@example.invalid>\r\n"])
async def test_spam_rescue_without_usable_message_id_stays_unknown_without_search(monkeypatch, header):
    fake = SpamRescueIMAP()
    fake.body = header + b"Subject: fixture\r\n\r\nbody\r\n"
    client = client_for(monkeypatch, fake)

    result = await client.move_emails_with_outcome(["13"], "Spam", "INBOX")

    assert result.targets("unknown") == ["13"]
    assert result.reconciliation_needed
    assert not any(c[0] == "SEARCH" for c in fake.commands)
    assert [c for c in fake.commands if c[0] == "STORE"] == [("STORE", "13", "+X-GM-LABELS", r"(\Inbox)")]


@pytest.mark.asyncio
@pytest.mark.parametrize("failure", ["fetch", "restore", "source_identity"])
async def test_spam_rescue_read_failure_or_source_replacement_stops_batch(monkeypatch, failure):
    fake = SpamRescueIMAP(auto_remove=failure != "source_identity")
    if failure == "fetch":
        original = fake.metadata

        def metadata(uid):
            if fake.selected == "INBOX":
                raise ConnectionError("lost read")
            return original(uid)

        fake.metadata = metadata
    else:
        original = fake.select

        async def select(mailbox):
            restoring = fake.selected == "INBOX"
            if restoring and failure == "restore":
                return "NO", [b"rejected"]
            if restoring:
                fake.identity = "56"
            return await original(mailbox)

        fake.select = select
    client = client_for(monkeypatch, fake)

    result = await client.move_emails_with_outcome(["13", "14"], "Spam", "INBOX")

    assert result.targets("unknown") == ["13"]
    assert result.outcomes[1].detail == "not-attempted"
    assert result.reconciliation_needed
    assert len([c for c in fake.commands if c[0] == "STORE"]) == 1


@pytest.mark.asyncio
async def test_user_label_confirmation_remains_case_sensitive(monkeypatch):
    fake = FakeIMAP()
    client = client_for(monkeypatch, fake)
    original = fake.metadata

    def metadata(uid):
        if any(c[0] == "STORE" for c in fake.commands):
            fake.labels.discard("keep")
            fake.labels.add("KEEP")
        return original(uid)

    fake.metadata = metadata
    result = await client.move_emails_with_outcome(["7"], "INBOX", "keep")

    assert result.targets("unknown") == ["7"]
    assert result.outcomes[0].detail == "destination-added-identity-unconfirmed"
    assert len([c for c in fake.commands if c[0] == "STORE"]) == 1


@pytest.mark.asyncio
async def test_append_exact_once_localized_and_discard_no_fallback(monkeypatch):
    for lost in (False, True):
        fake = FakeIMAP()
        fake.lose_append = lost
        client = client_for(monkeypatch, fake)
        msg = MIMEText("héllo", "plain", "utf-8")
        msg["Message-ID"] = "<old@example.invalid>"
        msg["Message-ID"] = "<duplicate@example.invalid>"
        result = await client.append_to_mailbox_with_outcome(msg, client.email_server, DRAFTS)
        assert result.status == ("unknown" if lost else "succeeded")
        assert result.uid == (None if lost else "11")
        assert result.uidvalidity == (None if lost else "42")
        appends = [c for c in fake.commands if c[0] == "APPEND"]
        assert len(appends) == 1
        assert appends[0][1:3] == (_quote_mailbox(DRAFTS), r"(\Draft \Seen)")
        parsed = message_from_bytes(appends[0][3])
        assert parsed.get_all("Message-ID") == [result.message_id]
        assert result.message_id.isascii() and "old@" not in result.message_id
        fake.lose_append = False
        for missing in ("MOVE", "UIDPLUS"):
            fake.protocol.capabilities.discard(missing)
            discarded = await client.move_emails_with_outcome(["11"], DRAFTS, TRASH)
            assert discarded.outcomes[0].detail == "move-uidplus-unavailable"
            fake.protocol.capabilities.add(missing)
        assert (await client.move_emails_with_outcome(["11"], DRAFTS, TRASH)).targets("succeeded") == ["11"]
        assert not any(c[0] in {"COPY", "EXPUNGE", "CLOSE"} for c in fake.commands)


@pytest.mark.asyncio
async def test_reads_identity_chunks_and_byte_limit(monkeypatch):
    fake = FakeIMAP()
    client = client_for(monkeypatch, fake)
    data = await client.get_email_body_by_id("7")
    assert data["gmail_message_id"] == MSGID and data["email_id"] == "7"
    assert set(data["gmail_labels"]) == fake.labels - {r"\Inbox"}
    monkeypatch.setenv("MCP_EMAIL_SERVER_MAX_FETCH_BYTES", "1")
    fake.commands.clear()
    assert await client.get_email_body_by_id("7") is None
    assert not any("BODY.PEEK" in str(c) for c in fake.commands)
    fake.commands.clear()
    await client._gmail_metadata(fake, ["7"] * 0)  # empty read does no work
    with pytest.raises(ValueError):
        await client.move_emails_with_outcome(["7"] * 26, "INBOX", "keep")
    assert not fake.commands


@pytest.mark.asyncio
async def test_fetch_chunks_and_standard_search(monkeypatch):
    fake = FakeIMAP()
    client = client_for(monkeypatch, fake)
    fake.metadata = lambda uid: f'1 FETCH (UID {uid} X-GM-MSGID {MSGID} X-GM-LABELS ("keep"))'.encode()
    uids = [str(n) for n in range(1, 13)]
    await client._batch_fetch_dates(fake, uids)
    await client._batch_fetch_headers(fake, uids)
    await client._batch_fetch_senders(fake, uids)
    fetches = [c for c in fake.commands if c[0] == "FETCH"]
    assert all(len(c[1].split(",")) <= 5 and re.fullmatch(r"[0-9,]+", c[1]) for c in fetches)
    assert all("BODY[" not in c[2] for c in fetches)
    fake.commands.clear()
    total, emails = await client.get_emails_metadata(text="bounded", has_attachment=False)
    assert total == 1 and emails[0]["gmail_message_id"] == MSGID
    search = next(c for c in fake.commands if c[0] == "SEARCH")
    assert search[1:] == ("TEXT", "bounded", "NOT", "HEADER", "Content-Type", "multipart/mixed")
    with pytest.raises(ValueError, match="BODY"):
        await client.get_emails_metadata(body="unsupported")


@pytest.mark.asyncio
async def test_label_application_services_and_discovery(monkeypatch):
    fake = FakeIMAP()
    client = client_for(monkeypatch, fake)
    handler = SimpleNamespace(incoming_client=client, list_mailboxes=client.list_mailboxes)
    mutation_provider = ClassicMutationProvider(handler)
    account = MutationAccountSnapshot("test", "legacy", (), (), True, False, gateway_provider="gmail_imap")
    authority = SimpleNamespace(resolve=lambda *args, **kwargs: account)
    factory = SimpleNamespace(open=lambda *args, **kwargs: MutationProviderAccess(account, mutation_provider))
    projection = SimpleNamespace(invalidate=AsyncMock())
    projections = SimpleNamespace(open_projection=AsyncMock(return_value=projection))
    copy = CopyService(authority, factory, projections)
    result = await ApplyLabelService(copy).execute(ApplyLabelCommand("test", ("7",), "Labels/exact"))
    assert result.targets("succeeded") == ["7"] and "Labels/exact" in fake.labels
    remove = RemoveLabelService(authority, factory, projections)
    result = await remove.execute(RemoveLabelCommand("test", ("7",), "Labels/exact"))
    assert result.targets("succeeded") == ["7"] and "Labels/exact" not in fake.labels
    read_account = ReadAccountSnapshot("test", "legacy", (), False, gateway_provider="gmail_imap")
    reader = ClassicReadProvider(handler)
    read_authority = SimpleNamespace(resolve=lambda *args, **kwargs: read_account)
    read_factory = SimpleNamespace(open=lambda *args, **kwargs: ReadProviderAccess(read_account, reader))
    labels = await LabelDiscoveryService(MailboxDiscoveryService(read_authority, read_factory)).execute(
        ListLabelsQuery("test")
    )
    assert {label.name for label in labels} == {"work", "keep", "Labels/exact", 'a "quote" (x)'}
    applied = await EmailLabelService(read_authority, read_factory).execute(GetEmailLabelsQuery("test", "7"))
    assert set(applied) == {"work", "keep"}


def test_config_env_and_toml_round_trip(monkeypatch, tmp_path):
    import tomllib

    import tomli_w

    monkeypatch.setenv("MCP_EMAIL_SERVER_EMAIL_ADDRESS", "synthetic@example.invalid")
    monkeypatch.setenv("MCP_EMAIL_SERVER_PASSWORD", "synthetic")
    monkeypatch.setenv("MCP_EMAIL_SERVER_IMAP_HOST", "127.0.0.1")
    monkeypatch.setenv("MCP_EMAIL_SERVER_GATEWAY_PROVIDER", "gmail_imap")
    account = EmailSettings.from_env()
    assert account.gateway_provider == "gmail_imap" and not account.can_send
    path = tmp_path / "synthetic.toml"
    path.write_text(tomli_w.dumps(account.model_dump(mode="json", exclude_none=True)))
    assert EmailSettings.model_validate(tomllib.loads(path.read_text())).gateway_provider == "gmail_imap"
    monkeypatch.setenv("MCP_EMAIL_SERVER_GATEWAY_PROVIDER", "invalid")
    assert EmailSettings.from_env() is None


@pytest.mark.asyncio
async def test_default_provider_dispatch_is_unchanged(monkeypatch):
    for provider in (None, "proton_bridge"):
        fake = FakeIMAP()
        client = client_for(monkeypatch, fake, provider)
        result = await client.move_emails_with_outcome(["7"], "INBOX", "keep")
        assert result.targets("succeeded") == ["7"]
        assert any(c[0] == "MOVE" for c in fake.commands)
        assert not any("X-GM-LABELS" in str(c) for c in fake.commands)


def test_config_is_explicit_and_identity_is_decimal():
    settings = EmailSettings(
        account_name="local",
        full_name="Test",
        email_address="x@example.invalid",
        incoming=EmailServer(user_name="x", password="x", host="imap.gmail.com", port=993),
        email_service="gmail",
    )
    assert settings.gateway_provider is None
    assert (
        EmailSettings.model_validate({**settings.model_dump(), "gateway_provider": "gmail_imap"}).gateway_provider
        == "gmail_imap"
    )
    assert (
        parse_metadata([f'1 FETCH (UID 99 X-GM-MSGID {MSGID} X-GM-LABELS ("work" "keep"))'.encode()])["99"][
            "gmail_message_id"
        ]
        == MSGID
    )


def test_packaged_contract_and_worktree_import():
    import mcp_email_server.emails.classic as classic

    assert Path(classic.__file__).resolve().parents[2] == Path(__file__).resolve().parents[1]
    contract = json.loads(files("mcp_email_server").joinpath("gateway_gmail_contract.json").read_text())
    assert contract == {"version": 1, "provider": "gmail_imap", "mutation_semantics": "incremental-labels-no-expunge"}


@pytest.mark.asyncio
async def test_partial_gmail_summary_is_honest(monkeypatch):
    import mcp_email_server.app as app

    fake = FakeIMAP()
    fake.fail_remove = True
    client = client_for(monkeypatch, fake)
    outcome = await client.move_emails_with_outcome(["7"], "INBOX", "keep")
    monkeypatch.setattr(app, "move_emails_command", AsyncMock(return_value=outcome))
    summary = await app.move_emails("test", ["7"], "keep")
    assert not summary.startswith("Successfully")
    assert "destination-added-source-retained" in summary


@pytest.mark.asyncio
async def test_archive_adapter_uses_archive_dispatch(monkeypatch):
    fake = FakeIMAP()
    client = client_for(monkeypatch, fake)
    provider = ClassicMutationProvider(SimpleNamespace(incoming_client=client))
    account = MutationAccountSnapshot("test", "legacy", (), (), True, False, gateway_provider="gmail_imap")
    result = await provider.archive_gmail(ArchiveCommand("test", ("7",), "INBOX"), account)
    assert result.targets("succeeded") == ["7"]


@pytest.mark.asyncio
@pytest.mark.parametrize("projection_fails", [False, True])
async def test_archive_service_dispatches_gmail_and_invalidates_once(monkeypatch, projection_fails):
    fake = FakeIMAP()
    client = client_for(monkeypatch, fake)
    provider = ClassicMutationProvider(
        SimpleNamespace(incoming_client=client, email_settings=SimpleNamespace(gateway_provider="gmail_imap"))
    )
    account = MutationAccountSnapshot("test", "legacy", (), (), True, False, gateway_provider="gmail_imap")
    authority = SimpleNamespace(resolve=lambda *args, **kwargs: account)
    factory = SimpleNamespace(open=lambda *args, **kwargs: MutationProviderAccess(account, provider))
    projection = SimpleNamespace(
        invalidate=AsyncMock(side_effect=RuntimeError("unavailable") if projection_fails else None)
    )
    projections = SimpleNamespace(open=AsyncMock(return_value=projection))

    result = await ArchiveService(authority, factory, projections).execute(ArchiveCommand("test", ("7",), "INBOX"))

    assert result.archive_mailbox == "INBOX"
    assert result.gmail_archive is True
    assert result.batch.targets("succeeded") == ["7"]
    assert result.batch.reconciliation_needed is projection_fails
    assert fake.labels == {"work", "keep"}
    assert fake.writes == [("All", "12", "-X-GM-LABELS", r"(\Inbox)")]
    projection.invalidate.assert_awaited_once_with(("INBOX",))


@pytest.mark.asyncio
async def test_gmail_archive_summary_names_inbox_label_removal(monkeypatch):
    import mcp_email_server.app as app
    from mcp_email_server.application.mutations import (
        ArchiveMutationOutcome,
        BatchMutationOutcome,
        TargetMutationOutcome,
    )

    outcome = ArchiveMutationOutcome(
        BatchMutationOutcome((TargetMutationOutcome("7", "succeeded"),)), "INBOX", gmail_archive=True
    )
    monkeypatch.setattr(app, "archive_emails_command", AsyncMock(return_value=outcome))

    assert await app.archive_emails("test", ["7"]) == "Successfully archived 1 email(s): removed the Inbox label"


def load_proxy():
    path = Path(os.environ.get("MCP_EMAIL_TEST_PROXY", "/Users/jbkjr/dev/mcp-gateway/imap-proxy/imap_proxy.py"))
    if not path.is_file():
        pytest.skip("Set MCP_EMAIL_TEST_PROXY to run real gateway policy integration")
    spec = importlib.util.spec_from_file_location("gmail_test_imap_proxy", path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


class ProxyIMAP:
    """Pass every worker command through the real proxy Session and fake upstream."""

    def __init__(self, ip, monkeypatch):
        self.ip = ip
        self.fake = FakeIMAP()
        self.protocol = self.fake.protocol
        self.protocol.capability = self.capability
        self.wire = []
        self.capabilities = self.protocol.capabilities
        self.last_completion = b""
        self.captured = None
        self.response = None
        cfg = ip.Config(
            proxy_username="synthetic",
            proxy_password_sha256=b"x" * 32,
            upstream_username="synthetic@example.invalid",
            upstream_password="synthetic",
            provider="gmail_imap",
            account_name="synthetic",
            transport="gmail_implicit_tls_ca",
            upstream_port=993,
            upstream_starttls=False,
            snapshot_file="/synthetic",
            upstream_ca_file="/synthetic",
            drafts_mailbox=DRAFTS,
            trash_mailbox=TRASH,
            discard_enabled=True,
        )
        self.server = ip.ProxyServer(cfg)
        self.output = []
        writer = SimpleNamespace(write=self.output.append, drain=AsyncMock())
        self.session = ip.Session(cfg, asyncio.StreamReader(), writer, 1, self.server)
        self.session.upstream = self
        self.session.authenticated = True
        monkeypatch.setattr(ip, "read_gmail_snapshot", lambda _: {})

    def command(self, wire, literal=None):
        tokens = []
        self.ip.tokenize_line(wire, [tokens], False)
        args = tokens[2:]
        parts = [(wire, None)]
        if literal is not None:
            args.append(self.ip.String(literal, literal=True))
            parts = [(wire + b" ", self.ip.LiteralPart(literal, nonsync=False, literal8=False))]
        return self.ip.Command("T", tokens[1].upper(), args, parts)

    async def _run(self, cmd):  # noqa: C901 - fake upstream command dispatch
        ip = self.ip
        args = cmd.args
        name = cmd.name
        wire = cmd.parts[0][0].decode()
        if name == "LIST":
            return await self.fake.list()
        if name in {"SELECT", "EXAMINE"}:
            return await self.fake.select(_quote_mailbox(decode_mailbox_name(ip._astring(args[0]).decode())))
        if name == "CAPABILITY":
            return await self.fake.capability()
        if name == "LOGOUT":
            return await self.fake.logout()
        if name == "APPEND":
            return await self.fake.append(args[2].data, _quote_mailbox(DRAFTS), r"(\Draft \Seen)")
        assert name == "UID", wire
        operation = args[0].upper()
        if operation == "SEARCH":
            values = [ip._astring(a).decode() for a in args[1:]]
            return await self.fake.uid_search(*values)
        uid = args[1].text
        if operation == "FETCH":
            response = await self.fake.uid("fetch", uid, wire.split(" FETCH ", 1)[1].split(" ", 1)[1])
            if "BODY.PEEK[]" in wire and "X-GM-MSGID" in wire:
                # The proxy's private post-APPEND/provenance read includes body.
                meta = response[1][0][:-1]
                body = self.fake.draft if self.fake.selected == DRAFTS else self.fake.body
                response = "OK", [meta + f" BODY[] {{{len(body)}}}".encode(), bytearray(body), b")"]
            return response
        if operation == "STORE":
            label = ip._astring(args[3][0]).decode()
            return await self.fake.uid("store", uid, args[2].upper(), f"({ip._quote(label.encode()).decode()})")
        if operation == "MOVE":
            return await self.fake.uid("move", uid, ip._quote(ip._astring(args[2])).decode())
        raise AssertionError(wire)

    def _wire_lines(self, response, cmd):
        lines = []
        for item in response[1]:
            if isinstance(item, bytearray):
                lines.append(bytes(item))
            else:
                prefix = (
                    b"* LIST "
                    if cmd.name == "LIST"
                    else b"* SEARCH "
                    if cmd.name == "UID" and cmd.args[0].upper() == "SEARCH"
                    else b"* "
                )
                lines.append((item if item == b")" else (prefix + item).rstrip()) + b"\r\n")
        return lines

    async def internal(self, tag, parts):
        cmd = self.command(parts[0][0])
        response = await self._run(cmd)
        return response[0].encode(), self._wire_lines(response, cmd)

    async def execute(self, tag, parts, sink):
        cmd = self.pending
        response = await self._run(cmd)
        self.response = response
        if self.captured is not None:
            self.captured.extend(self._wire_lines(response, cmd))
        self.last_completion = (
            b"T " + response[0].encode() + b" " + (response[1][0] if cmd.name == "APPEND" else b"done")
        )
        return response[0].encode()

    async def invoke(self, text, literal=None):
        cmd = self.command(text.encode(), literal)
        self.pending = cmd
        self.wire.append((text, literal))
        self.response = None
        self.output.clear()
        await self.session._authenticated(cmd)
        return self.response if self.response is not None else ("NO", list(self.output))

    async def login(self, *_):
        return "OK", []

    async def logout(self):
        return await self.invoke("T LOGOUT")

    async def list(self, reference='""', pattern='"*"'):
        return await self.invoke(f"T LIST {reference} {pattern}")

    async def select(self, mailbox):
        return await self.invoke(f"T SELECT {mailbox}")

    async def capability(self):
        return await self.invoke("T CAPABILITY")

    async def uid(self, operation, uid, *args):
        return await self.invoke(f"T UID {operation.upper()} {uid} {' '.join(args)}")

    async def uid_search(self, *criteria, charset=None):
        return await self.invoke("T UID SEARCH " + " ".join(criteria))

    async def append(self, payload, mailbox, flags):
        return await self.invoke(f"T APPEND {mailbox} {flags}", payload)


@pytest.mark.asyncio
async def test_real_proxy_session_worker_contract(monkeypatch):
    ip = load_proxy()
    proxy = ProxyIMAP(ip, monkeypatch)
    client = client_for(monkeypatch, proxy)
    body = await client.get_email_body_by_id("7")
    assert body["gmail_message_id"] == MSGID
    total, rows = await client.get_emails_metadata(text="bounded", has_attachment=False)
    assert total == 1 and rows[0]["gmail_message_id"] == MSGID
    assert (await client.get_emails_metadata(text="a[query]"))[0] == 1
    assert (await client.move_emails_with_outcome(["7"], "INBOX", "keep")).targets("succeeded") == ["7"]
    assert proxy.fake.labels == {"work", "keep"}
    proxy.fake.labels.add(r"\Inbox")
    assert (await client._gmail_mutate(["7"], "INBOX", None, "archive")).targets("succeeded") == ["7"]
    assert (await client.remove_label_with_outcome(["99"], "work", "work")).targets("succeeded") == ["99"]
    proxy.fake.labels.update({"work", r"\Trash"})
    assert (await client.move_emails_with_outcome(["7"], "INBOX", "keep")).outcomes[0].status == "failed"
    proxy.fake.labels.discard(r"\Trash")
    proxy.fake.labels = {r"\Spam", "keep"}
    assert (await client.move_emails_with_outcome(["13"], "Spam", "INBOX")).targets("succeeded") == ["13"]
    proxy.fake.fail_remove = True
    partial = await client.move_emails_with_outcome(["7"], "INBOX", "work")
    assert partial.outcomes[0].detail == "destination-added-source-retained"
    proxy.fake.fail_remove = False
    msg = MIMEText("synthetic draft", "plain", "utf-8")
    result = await client.append_to_mailbox_with_outcome(msg, client.email_server, DRAFTS)
    assert result.status == "succeeded" and (result.uidvalidity, result.uid) == ("42", "11")
    assert proxy.server.appended_drafts.contains(42, 11)
    # APPEND invalidates selection; the next worker call must select Drafts.
    assert proxy.session.selected is ip.UNKNOWN_SELECTION
    assert (await client.move_emails_with_outcome(["11"], DRAFTS, TRASH)).targets("succeeded") == ["11"]
    for missing in ("MOVE", "UIDPLUS"):
        proxy.protocol.capabilities.discard(missing)
        assert (await client.move_emails_with_outcome(["11"], DRAFTS, TRASH)).outcomes[0].status == "failed"
        proxy.protocol.capabilities.add(missing)
    proxy.fake.draft_labels.add("keep")
    assert (await client.move_emails_with_outcome(["11"], DRAFTS, TRASH)).outcomes[0].status == "failed"
    proxy.fake.draft_labels.discard("keep")
    proxy.fake.draft += b"edited"
    assert (await client.move_emails_with_outcome(["11"], DRAFTS, TRASH)).outcomes[0].status == "failed"
    proxy.fake.validity = "43"
    assert (await client.move_emails_with_outcome(["11"], DRAFTS, TRASH)).outcomes[0].status == "failed"
    assert not any(re.search(r"COPY|EXPUNGE|CLOSE|X-GM-RAW|ENABLE|\\Deleted", wire) for wire, _ in proxy.wire)
    proxy.server.appended_drafts.close()


@pytest.mark.asyncio
async def test_real_proxy_session_spam_rescue_after_source_disappears(monkeypatch):
    proxy = ProxyIMAP(load_proxy(), monkeypatch)
    proxy.fake = SpamRescueIMAP()
    client = client_for(monkeypatch, proxy)

    result = await client.move_emails_with_outcome(["13"], "Spam", "INBOX")

    assert result.targets("succeeded") == ["13"]
    assert proxy.fake.labels == {r"\INBOX", "keep"}
    writes = [wire for wire, _ in proxy.wire if " STORE " in wire]
    assert writes == [r"T UID STORE 13 +X-GM-LABELS (\Inbox)"]
    assert ('T UID SEARCH HEADER Message-ID "<fixture@example.invalid>"', None) in proxy.wire
    assert not any("SEARCH ALL" in wire for wire, _ in proxy.wire)
    assert any("BODY.PEEK[HEADER.FIELDS (MESSAGE-ID)]" in wire for wire, _ in proxy.wire)
    assert not any(re.search(r"COPY|EXPUNGE|CLOSE|X-GM-RAW|ENABLE|\\Deleted", wire) for wire, _ in proxy.wire)
    proxy.server.appended_drafts.close()


@pytest.mark.asyncio
@pytest.mark.parametrize("no_op_ok", [False, True])
async def test_real_proxy_all_mail_archive_and_restore_contract(monkeypatch, no_op_ok):
    proxy = ProxyIMAP(load_proxy(), monkeypatch)
    proxy.fake.labels = {r"\Inbox"}
    proxy.fake.ignore_removals = no_op_ok
    client = client_for(monkeypatch, proxy)
    try:
        result = await client._gmail_mutate(["7"], "INBOX", None, "archive")

        assert result.targets("unknown" if no_op_ok else "succeeded") == ["7"]
        assert len(proxy.fake.writes) == 1
        assert proxy.fake.writes[0][:3] == ("All", "12", "-X-GM-LABELS")
        assert (r"T UID STORE 12 -X-GM-LABELS (\Inbox)", None) in proxy.wire
        if not no_op_ok:
            assert proxy.fake.labels == set()
            result = await client.move_emails_with_outcome(["12"], "All", "INBOX")
            assert result.targets("succeeded") == ["12"]
            assert proxy.fake.labels == {r"\Inbox"}
            proxy.fake.labels.update({"work", "keep"})
            result = await client.move_emails_with_outcome(["99"], "work", "INBOX")
            assert result.targets("succeeded") == ["99"]
            assert proxy.fake.labels == {r"\Inbox", "keep"}
            assert proxy.fake.writes[-1] == ("All", "12", "-X-GM-LABELS", '("work")')
        assert not any(
            "SEARCH ALL" in wire or re.search(r"COPY|EXPUNGE|CLOSE|\\Deleted", wire) for wire, _ in proxy.wire
        )
    finally:
        proxy.server.appended_drafts.close()
