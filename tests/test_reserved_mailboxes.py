"""Reserved mailboxes: an account owner's mailboxes that no MCP mutation may fill or reshape.

An account can list mailboxes in ``reserved_mailboxes``. The owner uses such a
mailbox as a signal of their own (for example, a label whose application means
"I approve this"), so the server must never put a message into one, nor create,
rename, or delete one or a mailbox that contains one. Reading, and taking a
message *out* of a reserved mailbox, stay allowed.

The layers covered are the matcher, configuration loading, the application
services (refusal before any provider is opened), the classic provider adapter
(refusal before any IMAP command, using fresh authority), the MCP tool surface,
and the raw protocol executor.
"""

from __future__ import annotations

import asyncio
from dataclasses import replace
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from mcp_email_server import app as app_module
from mcp_email_server import config as config_module
from mcp_email_server.adapters import protocol as protocol_adapter
from mcp_email_server.adapters.authority import LocalAccountResolution
from mcp_email_server.adapters.mutations import ClassicMutationProvider, LocalMutationBackend
from mcp_email_server.application.labels import label_mailbox
from mcp_email_server.application.limits import APPLICATION_LIMITS
from mcp_email_server.application.mutations import (
    AppendMutationOutcome,
    ApplyLabelCommand,
    ArchiveCommand,
    BatchMutationOutcome,
    CopyCommand,
    CreateFolderCommand,
    CreateLabelCommand,
    CreateMailboxCommand,
    DeleteCommand,
    DeleteFolderCommand,
    DeleteLabelCommand,
    FolderMutationOutcome,
    MoveCommand,
    MutationAccountSnapshot,
    MutationProviderAccess,
    MutationServices,
    RemoveLabelCommand,
    RenameFolderCommand,
    ReplyCommand,
    ReplySource,
    SaveDraftCommand,
    SaveToMailboxCommand,
    TargetMutationOutcome,
)
from mcp_email_server.application.protocol import ProtocolCommandService, ProtocolFailure, ProtocolRequest
from mcp_email_server.application.reserved_mailboxes import (
    ReservedMailboxError,
    normalize_reserved_mailboxes,
    require_unreserved,
    reserved_entry,
)
from mcp_email_server.config import EmailSettings, Settings
from mcp_email_server.emails.classic import EmailClient

RESERVED = ("Labels/Send",)
REFUSED = r"reserved mailbox 'Labels/Send'"
# Built at runtime: the formatter would otherwise turn escapes into look-alike literals.
_FULLWIDTH_SEND = "".join(chr(ord(character) + 0xFEE0) for character in "Send")
_ZERO_WIDTH_SEND = "Se" + chr(0x200B) + "nd"


def _account(**changes: object) -> MutationAccountSnapshot:
    account = MutationAccountSnapshot(
        account_name="primary",
        mode="legacy",
        allowed_senders=(),
        allowed_recipients=(),
        report_blocked_mutations=False,
        can_send=True,
        enable_folder_management=True,
        reserved_mailboxes=RESERVED,
    )
    return replace(account, **changes)


def _provider() -> MagicMock:
    """A provider whose every effect succeeds, so only policy can stop a call."""
    provider = MagicMock()
    batch = BatchMutationOutcome((TargetMutationOutcome("1", "succeeded"),))
    folder = FolderMutationOutcome("succeeded")
    for name in ("copy", "move", "delete", "remove_label", "set_flags"):
        setattr(provider, name, AsyncMock(return_value=batch))
    for name in ("create_folder", "delete_folder", "rename_folder"):
        setattr(provider, name, AsyncMock(return_value=folder))
    provider.create_mailbox = AsyncMock(return_value="created")
    provider.save_to_mailbox = AsyncMock(return_value=AppendMutationOutcome("succeeded", "<id@example.test>", "1"))
    provider.find_trash_mailbox = AsyncMock(return_value="Trash")
    provider.find_archive_mailbox = AsyncMock(return_value="Archive")
    provider.find_junk_mailbox = AsyncMock(return_value="Spam")
    provider.find_drafts_mailbox = AsyncMock(return_value="Drafts")
    provider.fetch_reply_source = AsyncMock(
        return_value=ReplySource("<parent@example.test>", None, None, "Hello", "sender@example.test")
    )
    return provider


_EFFECTS = (
    "copy",
    "move",
    "delete",
    "remove_label",
    "set_flags",
    "create_folder",
    "delete_folder",
    "rename_folder",
    "create_mailbox",
    "save_to_mailbox",
)


def _assert_no_effect(provider: MagicMock) -> None:
    for name in _EFFECTS:
        getattr(provider, name).assert_not_called()


def _services(
    *,
    account: MutationAccountSnapshot | None = None,
    opened_account: MutationAccountSnapshot | None = None,
    provider: MagicMock | None = None,
) -> tuple[MutationServices, MagicMock, MagicMock]:
    resolved = account if account is not None else _account()
    authority = MagicMock()
    authority.resolve.return_value = resolved
    selected = provider if provider is not None else _provider()
    factory = MagicMock()
    factory.open.return_value = MutationProviderAccess(
        opened_account if opened_account is not None else resolved, selected
    )
    projections = MagicMock()
    projections.open = AsyncMock(return_value=MagicMock(invalidate=AsyncMock()))
    return MutationServices.compose(authority, factory, projections), factory, selected


# --------------------------------------------------------------------------- #
# Matching
# --------------------------------------------------------------------------- #


@pytest.mark.parametrize(
    "mailbox",
    [
        "Labels/Send",
        label_mailbox("Send"),
        "labels/send",
        "LABELS/SEND",
        "Labels/SEND",
        "  Labels/Send  ",
        "Labels/ Send ",
        "Labels//Send",
        "Labels/Send/",
        # Hierarchy-delimiter variants: a `.`-delimited spelling is the same path.
        "Labels.Send",
        "labels.SEND",
        # Modified UTF-7 spellings that decode to the reserved name.
        "Labels/&AFM-end",
        "&AEw-abels/Send",
        # Compatibility forms and invisible format characters.
        "Labels/" + _FULLWIDTH_SEND,
        "Labels/" + _ZERO_WIDTH_SEND,
        # Children of a reserved mailbox are reserved too.
        "Labels/Send/Later",
        "Labels.Send.Later",
        "labels/send/a/b",
    ],
)
def test_destination_variants_resolve_to_the_reserved_entry(mailbox: str) -> None:
    assert reserved_entry(mailbox, RESERVED) == "Labels/Send"


@pytest.mark.parametrize(
    "mailbox",
    [
        "Labels/Sending",
        "Labels/Send2",
        "Labels/Sent",
        "Labels/Later",
        "Send",
        "Folders/Send",
        "Folders/Labels/Send",
        "INBOX",
        "Drafts",
        # A parent is not a destination inside the reserved mailbox.
        "Labels",
        # An `&` the server receives escaped is a literal, not this mailbox.
        "Labels/&-AFM-end",
    ],
)
def test_unrelated_destinations_are_not_reserved(mailbox: str) -> None:
    assert reserved_entry(mailbox, RESERVED) is None


@pytest.mark.parametrize("mailbox", ["Labels", "labels", " Labels/ ", "LABELS."])
def test_a_parent_counts_for_mailbox_shape_changes_only(mailbox: str) -> None:
    assert reserved_entry(mailbox, RESERVED) is None
    assert reserved_entry(mailbox, RESERVED, include_ancestors=True) == "Labels/Send"


def test_a_configured_entry_matches_however_it_was_spelled() -> None:
    for configured in (" labels/SEND ", "Labels.Send", "Labels/&AFM-end"):
        assert reserved_entry("Labels/Send", (configured,)) == configured


@pytest.mark.parametrize("mailbox", ["Labels/&AFM", "Labels/&!!!-end", "Labels/&AFMA-nd"])
def test_malformed_modified_utf7_is_compared_literally(mailbox: str) -> None:
    assert reserved_entry(mailbox, RESERVED) is None


def test_an_entry_naming_no_segment_is_ignored_rather_than_reserving_everything() -> None:
    assert reserved_entry("INBOX", ("/",)) is None
    assert reserved_entry("Labels/Send", ("/", "Labels/Send")) == "Labels/Send"


def test_an_empty_reservation_never_matches() -> None:
    assert reserved_entry("Labels/Send", ()) is None
    assert reserved_entry("Labels", (), include_ancestors=True) is None
    require_unreserved("primary", (), destinations=("Labels/Send",), shapes=("Labels",))


def test_refusal_is_a_permission_error_naming_the_mailbox_and_its_owner() -> None:
    with pytest.raises(ReservedMailboxError) as caught:
        require_unreserved("primary", RESERVED, destinations=("INBOX", "labels/send"))

    assert isinstance(caught.value, PermissionError)
    message = str(caught.value)
    assert "reserved mailbox 'Labels/Send'" in message
    assert "account owner" in message
    assert "'primary'" in message
    assert "Nothing was changed" in message


def test_shape_refusal_covers_the_parent_while_destination_does_not() -> None:
    require_unreserved("primary", RESERVED, destinations=("Labels",))
    with pytest.raises(ReservedMailboxError, match=REFUSED):
        require_unreserved("primary", RESERVED, shapes=("Labels",))


def test_normalize_strips_and_deduplicates_entries_in_order() -> None:
    assert normalize_reserved_mailboxes([" Labels/Send ", "Labels/Send", "Archive/Hold"]) == [
        "Labels/Send",
        "Archive/Hold",
    ]
    assert normalize_reserved_mailboxes([]) == []


@pytest.mark.parametrize(
    "values",
    [
        [""],
        ["   "],
        ["/"],
        [". /"],
        ["Labels/\x00Send"],
        [7],
        ["x" * (APPLICATION_LIMITS.mailbox_bytes + 1)],
        [f"Box{index}" for index in range(APPLICATION_LIMITS.policy_entries + 1)],
    ],
)
def test_normalize_rejects_entries_that_name_no_mailbox(values: list[object]) -> None:
    with pytest.raises(ValueError, match="reserved_mailboxes"):
        normalize_reserved_mailboxes(values)


# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #

_TOML = """
credential_storage = "plaintext"

[[emails]]
account_name = "owner"
full_name = "Owner"
email_address = "owner@example.test"
reserved_mailboxes = ["Labels/Send"]

[emails.incoming]
user_name = "owner"
password = "synthetic"
host = "127.0.0.1"
port = 1143
use_ssl = false

[[emails]]
account_name = "other"
full_name = "Other"
email_address = "other@example.test"

[emails.incoming]
user_name = "other"
password = "synthetic"
host = "127.0.0.1"
port = 1143
use_ssl = false
"""


@pytest.fixture
def legacy_config():
    config_module.CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    config_module.CONFIG_PATH.write_text(_TOML, encoding="utf-8")
    config_module.clear_settings_cache()
    yield
    config_module.clear_settings_cache()


def test_reserved_mailboxes_default_empty(email_settings: EmailSettings) -> None:
    assert email_settings.reserved_mailboxes == []


def test_account_field_is_normalized_and_validated(email_settings: EmailSettings) -> None:
    changed = EmailSettings.model_validate({
        **email_settings.model_dump(),
        "reserved_mailboxes": [" Labels/Send ", "Labels/Send"],
    })
    assert changed.reserved_mailboxes == ["Labels/Send"]
    with pytest.raises(ValueError, match="reserved_mailboxes"):
        EmailSettings.model_validate({**email_settings.model_dump(), "reserved_mailboxes": ["/"]})


def test_toml_reservation_reaches_only_its_own_account_snapshot(legacy_config) -> None:
    settings = Settings()
    owner = settings.get_account("owner")
    other = settings.get_account("other")
    assert isinstance(owner, EmailSettings) and owner.reserved_mailboxes == ["Labels/Send"]
    assert isinstance(other, EmailSettings) and other.reserved_mailboxes == []

    backend = LocalMutationBackend()
    assert backend.resolve("owner").reserved_mailboxes == ("Labels/Send",)
    assert backend.resolve("other").reserved_mailboxes == ()


def test_toml_round_trip_preserves_the_reservation(legacy_config) -> None:
    settings = Settings()
    settings.store()
    reloaded = Settings()
    owner = reloaded.get_account("owner")
    assert isinstance(owner, EmailSettings) and owner.reserved_mailboxes == ["Labels/Send"]


def _environment_account(monkeypatch: pytest.MonkeyPatch, name: str = "env") -> None:
    monkeypatch.setenv("MCP_EMAIL_SERVER_ACCOUNT_NAME", name)
    monkeypatch.setenv("MCP_EMAIL_SERVER_EMAIL_ADDRESS", "env@example.test")
    monkeypatch.setenv("MCP_EMAIL_SERVER_PASSWORD", "synthetic")
    monkeypatch.setenv("MCP_EMAIL_SERVER_IMAP_HOST", "127.0.0.1")


def test_environment_account_reads_a_comma_separated_reservation(monkeypatch: pytest.MonkeyPatch) -> None:
    _environment_account(monkeypatch)
    monkeypatch.setenv("MCP_EMAIL_SERVER_RESERVED_MAILBOXES", " Labels/Send , ,Archive/Hold")

    account = EmailSettings.from_env()

    assert account is not None
    assert account.reserved_mailboxes == ["Labels/Send", "Archive/Hold"]


def test_environment_account_defaults_to_no_reservation(monkeypatch: pytest.MonkeyPatch) -> None:
    _environment_account(monkeypatch)
    account = EmailSettings.from_env()
    assert account is not None
    assert account.reserved_mailboxes == []


def test_environment_override_of_a_stored_account_keeps_its_reservation_when_unset(
    monkeypatch: pytest.MonkeyPatch, legacy_config
) -> None:
    """Replacing an account's endpoints from the environment must not drop its reservation."""
    _environment_account(monkeypatch, "owner")
    owner = Settings().get_account("owner")
    assert isinstance(owner, EmailSettings)
    assert owner.incoming.user_name == "env@example.test"
    assert owner.reserved_mailboxes == ["Labels/Send"]


def test_environment_override_of_a_stored_account_can_replace_its_reservation(
    monkeypatch: pytest.MonkeyPatch, legacy_config
) -> None:
    _environment_account(monkeypatch, "owner")
    monkeypatch.setenv("MCP_EMAIL_SERVER_RESERVED_MAILBOXES", "Archive/Hold")
    owner = Settings().get_account("owner")
    assert isinstance(owner, EmailSettings) and owner.reserved_mailboxes == ["Archive/Hold"]


# --------------------------------------------------------------------------- #
# Application services: refused before any provider is opened
# --------------------------------------------------------------------------- #


@pytest.mark.parametrize(
    ("service", "command"),
    [
        ("apply_label", ApplyLabelCommand("primary", ("1",), "Send", "Drafts")),
        ("apply_label", ApplyLabelCommand("primary", ("1",), "send", "Drafts")),
        ("apply_label", ApplyLabelCommand("primary", ("1",), "SEND")),
        ("apply_label", ApplyLabelCommand("primary", ("1",), "Send/Later")),
        ("copy", CopyCommand("primary", ("1",), "Drafts", "Labels/Send")),
        ("copy", CopyCommand("primary", ("1",), "Drafts", " labels/send ")),
        ("copy", CopyCommand("primary", ("1",), "Drafts", "Labels.Send")),
        ("copy", CopyCommand("primary", ("1",), "Drafts", "Labels/&AFM-end")),
        ("move", MoveCommand("primary", ("1",), "Drafts", "Labels/Send")),
        ("move", MoveCommand("primary", ("1",), "INBOX", "LABELS/SEND/child")),
        ("save_to_mailbox", SaveToMailboxCommand("primary", ("a@example.test",), "s", "b", mailbox="Labels/Send")),
        ("save_to_mailbox", SaveToMailboxCommand("primary", ("a@example.test",), "s", "b", mailbox="labels/SEND")),
        ("create_mailbox", CreateMailboxCommand("primary", "Labels/Send")),
        ("create_mailbox", CreateMailboxCommand("primary", "Labels/Send/Child")),
        ("create_mailbox", CreateMailboxCommand("primary", "Labels")),
        ("create_folder", CreateFolderCommand("primary", "labels/send")),
        ("create_label", CreateLabelCommand("primary", "Send")),
        ("delete_folder", DeleteFolderCommand("primary", "Labels/Send")),
        ("delete_folder", DeleteFolderCommand("primary", "Labels")),
        ("delete_label", DeleteLabelCommand("primary", "send")),
        ("rename_folder", RenameFolderCommand("primary", "Labels/Send", "Labels/Elsewhere")),
        ("rename_folder", RenameFolderCommand("primary", "Labels/Elsewhere", "Labels/Send")),
        ("rename_folder", RenameFolderCommand("primary", "Folders/Outbox", "Labels/Send/Outbox")),
        ("rename_folder", RenameFolderCommand("primary", "Labels", "Tags")),
        ("rename_folder", RenameFolderCommand("primary", "Labels/Send/Child", "Folders/Child")),
    ],
)
async def test_every_route_into_or_reshaping_a_reserved_mailbox_is_refused_before_provider_access(
    service: str, command: object
) -> None:
    services, factory, provider = _services()

    with pytest.raises(ReservedMailboxError, match=REFUSED):
        await getattr(services, service).execute(command)

    factory.open.assert_not_called()
    _assert_no_effect(provider)


async def test_reply_draft_into_a_reserved_mailbox_is_refused_before_the_source_is_read() -> None:
    services, factory, provider = _services()
    command = ReplyCommand(
        account_name="primary",
        recipients=("sender@example.test",),
        subject="",
        body="Thanks",
        source_email_id="7",
        source_mailbox="INBOX",
        reply_mode="reply",
        draft_mailbox="Labels/Send",
    )

    with pytest.raises(ReservedMailboxError, match=REFUSED):
        await services.reply.execute(command)

    factory.open.assert_not_called()
    provider.fetch_reply_source.assert_not_called()
    _assert_no_effect(provider)


@pytest.mark.parametrize("drafts_mailbox", ["Labels/Send", None])
async def test_save_draft_is_refused_when_the_drafts_mailbox_resolves_to_a_reserved_one(
    drafts_mailbox: str | None,
) -> None:
    provider = _provider()
    provider.find_drafts_mailbox = AsyncMock(return_value="Labels/Send")
    account = _account(drafts_mailbox=drafts_mailbox)
    services, _factory, _ = _services(account=account, provider=provider)

    with pytest.raises(ReservedMailboxError, match=REFUSED):
        await services.save_draft.execute(SaveDraftCommand("primary", (), "s", "b"))

    _assert_no_effect(provider)


async def test_a_discovered_junk_destination_that_is_reserved_is_refused() -> None:
    provider = _provider()
    provider.find_junk_mailbox = AsyncMock(return_value="Labels/Send")
    services, _factory, _ = _services(provider=provider)

    with pytest.raises(ReservedMailboxError, match=REFUSED):
        await services.move.execute(MoveCommand("primary", ("1",), "INBOX", destination_role="junk"))

    _assert_no_effect(provider)


async def test_a_discovered_archive_destination_that_is_reserved_is_refused() -> None:
    provider = _provider()
    provider.find_archive_mailbox = AsyncMock(return_value="Labels/Send")
    services, _factory, _ = _services(provider=provider)

    with pytest.raises(ReservedMailboxError, match=REFUSED):
        await services.archive.execute(ArchiveCommand("primary", ("1",), "INBOX"))

    _assert_no_effect(provider)


async def test_a_discovered_trash_destination_that_is_reserved_is_refused() -> None:
    provider = _provider()
    provider.find_trash_mailbox = AsyncMock(return_value="Labels/Send")
    services, _factory, _ = _services(provider=provider)

    with pytest.raises(ReservedMailboxError, match=REFUSED):
        await services.delete.execute(DeleteCommand("primary", ("1",), "INBOX"))

    _assert_no_effect(provider)


@pytest.mark.parametrize(
    ("service", "command"),
    [
        ("apply_label", ApplyLabelCommand("primary", ("1",), "Send", "Drafts")),
        ("copy", CopyCommand("primary", ("1",), "Drafts", "Labels/Send")),
        ("move", MoveCommand("primary", ("1",), "Drafts", "Labels/Send")),
        ("save_to_mailbox", SaveToMailboxCommand("primary", ("a@example.test",), "s", "b", mailbox="Labels/Send")),
        ("create_mailbox", CreateMailboxCommand("primary", "Labels/Send")),
        ("create_label", CreateLabelCommand("primary", "Send")),
        ("delete_label", DeleteLabelCommand("primary", "Send")),
        ("rename_folder", RenameFolderCommand("primary", "Labels/Send", "Labels/Elsewhere")),
    ],
)
async def test_a_reservation_added_before_the_provider_opens_still_refuses(service: str, command: object) -> None:
    """The opened account's fresh authority is checked again before the effect."""
    services, factory, provider = _services(
        account=_account(reserved_mailboxes=()), opened_account=_account(reserved_mailboxes=RESERVED)
    )

    with pytest.raises(ReservedMailboxError, match=REFUSED):
        await getattr(services, service).execute(command)

    factory.open.assert_called()
    _assert_no_effect(provider)


@pytest.mark.parametrize(
    ("service", "command", "effect"),
    [
        # Taking a message out of a reserved mailbox only withdraws the owner's signal.
        ("remove_label", RemoveLabelCommand("primary", ("1",), "Send", "Drafts"), "remove_label"),
        ("move", MoveCommand("primary", ("1",), "Labels/Send", "Drafts"), "move"),
        ("copy", CopyCommand("primary", ("1",), "Labels/Send", "Archive"), "copy"),
        ("delete", DeleteCommand("primary", ("1",), "Labels/Send"), "move"),
        ("archive", ArchiveCommand("primary", ("1",), "Labels/Send"), "move"),
        # Neighbouring mailboxes are unaffected.
        ("apply_label", ApplyLabelCommand("primary", ("1",), "Sending", "Drafts"), "copy"),
        ("copy", CopyCommand("primary", ("1",), "Drafts", "Labels/Later"), "copy"),
        ("move", MoveCommand("primary", ("1",), "INBOX", "Send"), "move"),
        ("save_to_mailbox", SaveToMailboxCommand("primary", ("a@example.test",), "s", "b"), "save_to_mailbox"),
        ("save_draft", SaveDraftCommand("primary", (), "s", "b"), "save_to_mailbox"),
        ("create_label", CreateLabelCommand("primary", "Later"), "create_folder"),
        ("rename_folder", RenameFolderCommand("primary", "Labels/Later", "Labels/Someday"), "rename_folder"),
        ("create_mailbox", CreateMailboxCommand("primary", "Folders/Send"), "create_mailbox"),
    ],
)
async def test_withdrawing_and_unrelated_mutations_still_reach_the_provider(
    service: str, command: object, effect: str
) -> None:
    services, _factory, provider = _services()

    await getattr(services, service).execute(command)

    getattr(provider, effect).assert_awaited_once()


async def test_an_account_without_a_reservation_behaves_exactly_as_before() -> None:
    services, _factory, provider = _services(account=_account(reserved_mailboxes=()))

    await services.apply_label.execute(ApplyLabelCommand("primary", ("1",), "Send", "Drafts"))
    await services.create_label.execute(CreateLabelCommand("primary", "Send"))

    assert provider.copy.await_args.args[0].destination_mailbox == "Labels/Send"
    assert provider.create_folder.await_args.args[0].folder_name == "Labels/Send"


async def test_only_the_reserving_account_is_restricted(legacy_config) -> None:
    """Two real legacy accounts: the same request is refused on one and allowed on the other."""
    backend = LocalMutationBackend()
    provider = _provider()
    factory = MagicMock()
    factory.open.side_effect = lambda name, **_kwargs: MutationProviderAccess(backend.resolve(name), provider)
    projections = MagicMock()
    projections.open = AsyncMock(return_value=MagicMock(invalidate=AsyncMock()))
    services = MutationServices.compose(backend, factory, projections)

    with pytest.raises(ReservedMailboxError, match=REFUSED):
        await services.apply_label.execute(ApplyLabelCommand("owner", ("1",), "Send", "Drafts"))
    provider.copy.assert_not_called()

    await services.apply_label.execute(ApplyLabelCommand("other", ("1",), "Send", "Drafts"))
    provider.copy.assert_awaited_once()


# --------------------------------------------------------------------------- #
# Real services over the classic adapter and IMAP client: no IMAP command at all
# --------------------------------------------------------------------------- #


def _mock_imap() -> AsyncMock:
    imap = AsyncMock()
    imap._client_task = asyncio.Future()
    imap._client_task.set_result(None)
    imap.wait_hello_from_server = AsyncMock()
    imap.login = AsyncMock(return_value=MagicMock(result="OK", lines=[]))
    imap.id = AsyncMock(return_value=MagicMock(result="OK"))
    imap.select = AsyncMock(return_value=("OK", []))
    imap.uid = AsyncMock(return_value=("OK", []))
    imap.create = AsyncMock(return_value=("OK", []))
    imap.delete = AsyncMock(return_value=("OK", []))
    imap.expunge = AsyncMock(return_value=("OK", []))
    imap.logout = AsyncMock()
    imap.protocol = MagicMock(capabilities=("IMAP4rev1", "UIDPLUS", "MOVE"))
    imap.protocol.capability = AsyncMock()
    return imap


def _classic(
    client: EmailClient,
    *,
    resolved: MutationAccountSnapshot,
    opened: MutationAccountSnapshot | None = None,
    fresh: MutationAccountSnapshot | None = None,
) -> MutationServices:
    handler = MagicMock()
    handler.incoming_client = client
    handler.outgoing_client = None
    handler.email_settings = SimpleNamespace(incoming=client.email_server)
    authority = MagicMock()
    authority.resolve.return_value = resolved
    provider = ClassicMutationProvider(handler, (lambda: fresh) if fresh is not None else None)
    factory = MagicMock()
    factory.open.return_value = MutationProviderAccess(opened or resolved, provider)
    projections = MagicMock()
    projections.open = AsyncMock(return_value=MagicMock(invalidate=AsyncMock()))
    return MutationServices.compose(authority, factory, projections)


_CLASSIC_ROUTES = [
    ("apply_label", ApplyLabelCommand("primary", ("1",), "Send", "Drafts")),
    ("copy", CopyCommand("primary", ("1",), "Drafts", "labels/send")),
    ("move", MoveCommand("primary", ("1",), "Drafts", "Labels/Send")),
    ("save_to_mailbox", SaveToMailboxCommand("primary", ("a@example.test",), "s", "b", mailbox="Labels/Send")),
    ("create_mailbox", CreateMailboxCommand("primary", "Labels/Send")),
    ("create_label", CreateLabelCommand("primary", "Send")),
    ("delete_folder", DeleteFolderCommand("primary", "Labels")),
    ("rename_folder", RenameFolderCommand("primary", "Labels/Elsewhere", "Labels/Send")),
]


@pytest.mark.parametrize(("service", "command"), _CLASSIC_ROUTES)
async def test_refusal_opens_no_imap_session(email_server, service: str, command: object) -> None:
    client = EmailClient(email_server)
    imap = _mock_imap()
    services = _classic(client, resolved=_account())
    with patch.object(client, "imap_class", return_value=imap) as connect:
        with pytest.raises(ReservedMailboxError, match=REFUSED):
            await getattr(services, service).execute(command)

    connect.assert_not_called()
    assert imap.mock_calls == []


@pytest.mark.parametrize(("service", "command"), _CLASSIC_ROUTES)
async def test_the_provider_rechecks_fresh_authority_before_any_imap_command(
    email_server, service: str, command: object
) -> None:
    """A reservation the application layer could not see still stops the wire command."""
    client = EmailClient(email_server)
    imap = _mock_imap()
    unreserved = _account(reserved_mailboxes=())
    services = _classic(client, resolved=unreserved, opened=unreserved, fresh=_account())
    with patch.object(client, "imap_class", return_value=imap) as connect:
        with pytest.raises(ReservedMailboxError, match=REFUSED):
            await getattr(services, service).execute(command)

    connect.assert_not_called()
    assert imap.mock_calls == []


async def test_the_installed_mutation_guard_refuses_a_reservation_added_mid_effect(email_server) -> None:
    """The guard the provider leaves on the IMAP client re-reads authority each time it runs."""
    client = EmailClient(email_server)
    handler = MagicMock(incoming_client=client, outgoing_client=None)
    current = {"account": _account(reserved_mailboxes=())}
    provider = ClassicMutationProvider(handler, lambda: current["account"])
    provider._guard(current["account"], "organize", destinations=("Labels/Send",))
    assert client.mutation_guard is not None

    current["account"] = _account()
    with pytest.raises(ReservedMailboxError, match=REFUSED):
        client.mutation_guard()


async def test_unreserved_copy_still_reaches_imap(email_server) -> None:
    client = EmailClient(email_server)
    imap = _mock_imap()
    services = _classic(client, resolved=_account())
    with patch.object(client, "imap_class", return_value=imap):
        result = await services.copy.execute(CopyCommand("primary", ("1",), "Drafts", "Labels/Later"))

    assert result.targets("succeeded") == ["1"]
    assert [call.args for call in imap.uid.await_args_list] == [("copy", "1", '"Labels/Later"')]


# --------------------------------------------------------------------------- #
# MCP tools
# --------------------------------------------------------------------------- #


@pytest.mark.parametrize(
    ("tool", "arguments"),
    [
        (
            "apply_label",
            {"account_name": "primary", "email_ids": ["1"], "label_name": "Send", "source_mailbox": "Drafts"},
        ),
        ("copy_emails", {"account_name": "primary", "email_ids": ["1"], "destination_mailbox": "Labels/Send"}),
        ("move_emails", {"account_name": "primary", "email_ids": ["1"], "destination_mailbox": "labels/send"}),
        (
            "save_to_mailbox",
            {
                "account_name": "primary",
                "recipients": ["a@example.test"],
                "subject": "s",
                "body": "b",
                "mailbox": "Labels/Send",
            },
        ),
        ("create_mailbox", {"account_name": "primary", "mailbox": "Labels/Send"}),
        ("create_label", {"account_name": "primary", "label_name": "Send"}),
        ("delete_label", {"account_name": "primary", "label_name": "Send"}),
        ("create_folder", {"account_name": "primary", "folder_name": "Labels/Send"}),
        ("delete_folder", {"account_name": "primary", "folder_name": "Labels"}),
        ("rename_folder", {"account_name": "primary", "old_name": "Labels/Send", "new_name": "Labels/Old"}),
    ],
)
async def test_mcp_tools_surface_the_refusal_without_any_effect(
    monkeypatch: pytest.MonkeyPatch, tool: str, arguments: dict[str, object]
) -> None:
    services, factory, provider = _services()
    monkeypatch.setattr(app_module, "get_application_runtime", lambda: SimpleNamespace(mutations=services))

    with pytest.raises(Exception, match=REFUSED):
        await app_module.mcp.call_tool(tool, arguments)

    factory.open.assert_not_called()
    _assert_no_effect(provider)


async def test_mcp_remove_label_still_withdraws_from_a_reserved_mailbox(monkeypatch: pytest.MonkeyPatch) -> None:
    services, _factory, provider = _services()
    monkeypatch.setattr(app_module, "get_application_runtime", lambda: SimpleNamespace(mutations=services))

    result = await app_module.mcp.call_tool(
        "remove_label",
        {"account_name": "primary", "email_ids": ["1"], "label_name": "Send", "source_mailbox": "Drafts"},
    )

    assert "Successfully removed label 'Send'" in str(result)
    provider.remove_label.assert_awaited_once()


# --------------------------------------------------------------------------- #
# Raw protocol execution
# --------------------------------------------------------------------------- #


@pytest.fixture
def resolution(email_settings: EmailSettings, tmp_path) -> LocalAccountResolution:
    account = email_settings.model_copy(update={"allow_protocol_commands": True, "reserved_mailboxes": ["Labels/Send"]})
    settings = Settings.model_construct(
        emails=[account],
        providers=[],
        db_location=str(tmp_path / "index.sqlite3"),
        enable_attachment_content=True,
        allowed_senders=[],
        allowed_recipients=["*"],
        allowed_mutations=["draft", "organize", "delete", "send", "append"],
    )
    return LocalAccountResolution("legacy", settings, account)


def test_raw_imap_is_refused_for_an_account_with_reserved_mailboxes(resolution: LocalAccountResolution) -> None:
    with pytest.raises(ProtocolFailure, match="reserved_mailboxes"):
        protocol_adapter.require_protocol_access(resolution, "imap")


def test_raw_imap_admission_is_unchanged_without_a_reservation(resolution: LocalAccountResolution) -> None:
    unreserved = replace(resolution, account=resolution.account.model_copy(update={"reserved_mailboxes": []}))
    protocol_adapter.require_protocol_access(unreserved, "imap")


def test_raw_smtp_admission_is_not_affected_by_a_reservation(resolution: LocalAccountResolution) -> None:
    """SMTP never writes a mailbox; send policy is unchanged."""
    protocol_adapter.require_protocol_access(resolution, "smtp")


def test_a_reservation_added_mid_session_stops_the_next_raw_command(
    monkeypatch: pytest.MonkeyPatch, resolution: LocalAccountResolution
) -> None:
    unreserved = replace(resolution, account=resolution.account.model_copy(update={"reserved_mailboxes": []}))
    session = protocol_adapter.LocalProtocolSession(unreserved, "imap")
    monkeypatch.setattr(protocol_adapter, "resolve_local_account", lambda *args, **kwargs: resolution)

    with pytest.raises(ProtocolFailure, match="reserved_mailboxes"):
        session.guard()


async def test_raw_imap_refusal_happens_before_any_connection(
    monkeypatch: pytest.MonkeyPatch, resolution: LocalAccountResolution
) -> None:
    monkeypatch.setattr(protocol_adapter, "resolve_local_account", lambda *args, **kwargs: resolution)
    opened = MagicMock(side_effect=AssertionError("no session may open"))
    monkeypatch.setattr(protocol_adapter.LocalProtocolSession, "open", opened)
    service = ProtocolCommandService(protocol_adapter.LocalProtocolBackend())

    result = await service.execute(
        ProtocolRequest.model_validate({
            "account_name": "test_account",
            "commands": [{"command": 'SELECT "Drafts"'}, {"command": 'UID COPY 1 "Labels/Send"'}],
        })
    )

    assert result.status == "stopped"
    assert result.error is not None and "reserved_mailboxes" in result.error
    assert [step.status for step in result.results] == ["not_attempted", "not_attempted"]
    opened.assert_not_called()
