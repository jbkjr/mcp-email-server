"""Label reads and label removal across the application, adapter, and provider layers.

Labels are ordinary IMAP mailboxes under a literal ``Labels/`` prefix, so these
tests care about three things the mailbox tools do not: the prefix projection,
the Message-ID lookup that links a message to its label copies, and the promise
that removing a label never touches the message the caller named.
"""

from __future__ import annotations

import asyncio
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock, Mock, patch

import pytest

from mcp_email_server.adapters.mutations import ClassicMutationProvider
from mcp_email_server.adapters.reads import ClassicReadProvider
from mcp_email_server.app import get_email_labels, list_labels, remove_label
from mcp_email_server.application import limits as limits_module
from mcp_email_server.application import reads as reads_module
from mcp_email_server.application.labels import (
    LABEL_LIST_PATTERN,
    LABEL_MAILBOX_PREFIX,
    MAXIMUM_LABEL_NAME_BYTES,
    label_mailbox,
    label_name_from_mailbox,
    validate_label_name,
)
from mcp_email_server.application.limits import APPLICATION_LIMITS
from mcp_email_server.application.mutations import (
    BatchMutationOutcome,
    MutationAccountSnapshot,
    MutationProviderAccess,
    RemoveLabelCommand,
    RemoveLabelService,
    TargetMutationOutcome,
)
from mcp_email_server.application.reads import (
    EmailLabelService,
    GetEmailLabelsQuery,
    LabelDiscoveryService,
    ListLabelsQuery,
    MailboxDiscoveryService,
    ReadAccountSnapshot,
    ReadProviderAccess,
    ReadProviderError,
)
from mcp_email_server.emails.classic import EmailClient, _quoted_imap_message_id
from mcp_email_server.emails.models import MailboxInfo

MESSAGE_ID = "<abc-123@example.test>"


# --------------------------------------------------------------------------
# Shared fakes
# --------------------------------------------------------------------------


def _read_account() -> ReadAccountSnapshot:
    return ReadAccountSnapshot(
        account_name="work",
        mode="managed",
        allowed_senders=("allowed@example.test",),
        enable_attachment_download=False,
    )


def _read_ports(account: ReadAccountSnapshot | None = None):
    resolved = account or _read_account()
    authority = Mock()
    authority.resolve.return_value = resolved
    provider = Mock()
    factory = Mock()
    factory.open.return_value = ReadProviderAccess(resolved, provider)
    return authority, factory, provider


def _mailbox(name: str, *, delimiter: str = "/", flags: list[str] | None = None) -> MailboxInfo:
    return MailboxInfo(name=name, delimiter=delimiter, flags=flags or [r"\HasNoChildren"])


def _mutation_account() -> MutationAccountSnapshot:
    return MutationAccountSnapshot(
        account_name="work",
        mode="managed",
        allowed_senders=("allowed@example.test",),
        allowed_recipients=(),
        report_blocked_mutations=False,
        can_send=True,
    )


def _mutation_ports(outcome: BatchMutationOutcome | Exception):
    account = _mutation_account()
    authority = Mock()
    authority.resolve.return_value = account
    provider = Mock()
    if isinstance(outcome, Exception):
        provider.remove_label = AsyncMock(side_effect=outcome)
    else:
        provider.remove_label = AsyncMock(return_value=outcome)
    factory = Mock()
    factory.open.return_value = MutationProviderAccess(account, provider)
    projection = Mock()
    projection.invalidate = AsyncMock()
    projections = Mock()
    projections.open = AsyncMock(return_value=projection)
    return authority, factory, projections, provider, projection


# --------------------------------------------------------------------------
# The `Labels/` naming convention
# --------------------------------------------------------------------------


def test_label_mailbox_prefixes_and_round_trips() -> None:
    assert label_mailbox("Work") == "Labels/Work"
    assert label_name_from_mailbox("Labels/Work") == "Work"
    # Nested names keep their internal separators; only the prefix is stripped.
    assert label_name_from_mailbox("Labels/Work/2026") == "Work/2026"


def test_label_name_from_mailbox_rejects_non_labels() -> None:
    assert label_name_from_mailbox("INBOX") is None
    assert label_name_from_mailbox("Labels") is None
    # The bare container is a grouping mailbox, not a label.
    assert label_name_from_mailbox(LABEL_MAILBOX_PREFIX) is None


@pytest.mark.parametrize(
    ("value", "message"),
    [
        ("", "must not be empty"),
        ("   ", "must not be empty"),
        ("Work\nInjected", "control characters"),
        ("Labels/Work", "must not repeat"),
        ("x" * (MAXIMUM_LABEL_NAME_BYTES + 1), "exceeds"),
    ],
)
def test_validate_label_name_rejects_unusable_values(value: str, message: str) -> None:
    with pytest.raises(ValueError, match=message):
        validate_label_name(value)


def test_label_name_bound_leaves_room_for_the_prefix() -> None:
    longest = "x" * MAXIMUM_LABEL_NAME_BYTES
    assert len(label_mailbox(longest).encode("utf-8")) == APPLICATION_LIMITS.mailbox_bytes


# --------------------------------------------------------------------------
# list_labels (application)
# --------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_label_discovery_projects_prefixed_mailboxes_and_skips_the_container() -> None:
    authority, factory, provider = _read_ports()
    provider.list_mailboxes = AsyncMock(
        return_value=[
            _mailbox("INBOX"),
            _mailbox("Labels"),
            _mailbox(LABEL_MAILBOX_PREFIX),
            _mailbox("Labels/Work", flags=[r"\HasChildren"]),
            _mailbox("Labels/Work/2026"),
        ]
    )
    service = LabelDiscoveryService(MailboxDiscoveryService(authority, factory))

    labels = await service.execute(ListLabelsQuery(account_name="work"))

    assert [(label.name, label.full_path) for label in labels] == [
        ("Work", "Labels/Work"),
        ("Work/2026", "Labels/Work/2026"),
    ]
    assert labels[0].flags == [r"\HasChildren"]
    assert labels[0].delimiter == "/"


@pytest.mark.asyncio
async def test_label_discovery_asks_the_provider_only_for_label_mailboxes() -> None:
    authority, factory, provider = _read_ports()
    provider.list_mailboxes = AsyncMock(return_value=[])
    service = LabelDiscoveryService(MailboxDiscoveryService(authority, factory))

    assert await service.execute(ListLabelsQuery(account_name="work")) == []

    query = provider.list_mailboxes.await_args.args[0]
    assert query.pattern == LABEL_LIST_PATTERN == "Labels/*"
    assert query.account_name == "work"


@pytest.mark.asyncio
async def test_label_discovery_inherits_the_mailbox_count_ceiling() -> None:
    authority, factory, provider = _read_ports()
    provider.list_mailboxes = AsyncMock(
        return_value=[_mailbox(f"Labels/label-{index}") for index in range(APPLICATION_LIMITS.mailboxes + 1)]
    )
    service = LabelDiscoveryService(MailboxDiscoveryService(authority, factory))

    with pytest.raises(ReadProviderError, match="mailbox count exceeds"):
        await service.execute(ListLabelsQuery(account_name="work"))


def test_list_labels_query_validates_the_account_name() -> None:
    with pytest.raises(ValueError, match="account_name"):
        ListLabelsQuery(account_name="").validate()


# --------------------------------------------------------------------------
# get_email_labels (application)
# --------------------------------------------------------------------------


def _email_label_provider(*, message_id: str | None, mailboxes: list[MailboxInfo], found: tuple[str, ...]):
    authority, factory, provider = _read_ports()
    provider.fetch_message_id = AsyncMock(return_value=message_id)
    provider.list_mailboxes = AsyncMock(return_value=mailboxes)
    provider.search_message_id_in_mailboxes = AsyncMock(return_value=found)
    return EmailLabelService(authority, factory), provider


@pytest.mark.asyncio
async def test_email_labels_reports_only_labels_that_hold_the_message() -> None:
    service, provider = _email_label_provider(
        message_id=MESSAGE_ID,
        mailboxes=[_mailbox("Labels/Work"), _mailbox("Labels/Personal"), _mailbox(LABEL_MAILBOX_PREFIX)],
        found=("Labels/Personal",),
    )

    assert await service.execute(GetEmailLabelsQuery(account_name="work", email_id="42")) == ["Personal"]

    message_id, probed = provider.search_message_id_in_mailboxes.await_args.args
    assert message_id == MESSAGE_ID
    assert probed == ("Labels/Work", "Labels/Personal")


@pytest.mark.asyncio
async def test_email_labels_returns_empty_when_the_message_id_is_unavailable() -> None:
    """A blocked sender and a missing message are the same observable outcome."""
    service, provider = _email_label_provider(message_id=None, mailboxes=[_mailbox("Labels/Work")], found=())

    assert await service.execute(GetEmailLabelsQuery(account_name="work", email_id="42")) == []

    provider.search_message_id_in_mailboxes.assert_not_awaited()
    provider.list_mailboxes.assert_not_awaited()


@pytest.mark.asyncio
async def test_email_labels_returns_empty_without_any_labels() -> None:
    service, provider = _email_label_provider(message_id=MESSAGE_ID, mailboxes=[], found=())

    assert await service.execute(GetEmailLabelsQuery(account_name="work", email_id="42")) == []

    provider.search_message_id_in_mailboxes.assert_not_awaited()


@pytest.mark.asyncio
async def test_email_labels_caps_the_number_of_probed_folders() -> None:
    service, provider = _email_label_provider(
        message_id=MESSAGE_ID,
        mailboxes=[_mailbox(f"Labels/label-{index}") for index in range(APPLICATION_LIMITS.mailboxes + 1)],
        found=(),
    )

    with pytest.raises(ReadProviderError, match="mailbox count exceeds"):
        await service.execute(GetEmailLabelsQuery(account_name="work", email_id="42"))

    provider.search_message_id_in_mailboxes.assert_not_awaited()


@pytest.mark.asyncio
async def test_email_labels_ignores_mailboxes_the_provider_reports_but_did_not_list() -> None:
    service, _ = _email_label_provider(
        message_id=MESSAGE_ID,
        mailboxes=[_mailbox("Labels/Work")],
        found=("Labels/Work", "Archive"),
    )

    assert await service.execute(GetEmailLabelsQuery(account_name="work", email_id="42")) == ["Work"]


@pytest.mark.asyncio
async def test_email_labels_bounds_the_whole_workflow_by_one_provider_deadline(monkeypatch) -> None:
    """The multi-folder fan-out must not buy itself a fresh budget per call."""
    monkeypatch.setattr(
        reads_module,
        "APPLICATION_LIMITS",
        limits_module.ApplicationLimits(provider_timeout_seconds=0.05),
    )
    authority, factory, provider = _read_ports()

    async def slow_fetch(*_args, **_kwargs):
        await asyncio.sleep(0.02)
        return MESSAGE_ID

    async def slow_list(*_args, **_kwargs):
        await asyncio.sleep(0.02)
        return [_mailbox("Labels/Work")]

    async def slow_search(*_args, **_kwargs):
        await asyncio.sleep(0.05)
        return ("Labels/Work",)

    provider.fetch_message_id = AsyncMock(side_effect=slow_fetch)
    provider.list_mailboxes = AsyncMock(side_effect=slow_list)
    provider.search_message_id_in_mailboxes = AsyncMock(side_effect=slow_search)
    service = EmailLabelService(authority, factory)

    with pytest.raises(ReadProviderError, match="timed out"):
        await service.execute(GetEmailLabelsQuery(account_name="work", email_id="42"))


@pytest.mark.parametrize(
    ("query", "message"),
    [
        (GetEmailLabelsQuery(account_name="", email_id="1"), "account_name"),
        (GetEmailLabelsQuery(account_name="work", email_id="0"), "email_id"),
        (GetEmailLabelsQuery(account_name="work", email_id="not-a-uid"), "email_id"),
        (GetEmailLabelsQuery(account_name="work", email_id="1", mailbox=""), "mailbox"),
    ],
)
def test_email_labels_query_rejects_invalid_input(query: GetEmailLabelsQuery, message: str) -> None:
    with pytest.raises(ValueError, match=message):
        query.validate()


# --------------------------------------------------------------------------
# remove_label (application)
# --------------------------------------------------------------------------


def _remove_command(**overrides) -> RemoveLabelCommand:
    return RemoveLabelCommand(
        account_name=overrides.pop("account_name", "work"),
        email_ids=overrides.pop("email_ids", ("42",)),
        label_name=overrides.pop("label_name", "Work"),
        source_mailbox=overrides.pop("source_mailbox", "INBOX"),
    )


def test_remove_label_command_targets_only_the_label_mailbox() -> None:
    assert _remove_command().label_mailbox == "Labels/Work"


@pytest.mark.parametrize(
    ("command", "message"),
    [
        (_remove_command(account_name=""), "account_name"),
        (_remove_command(email_ids=()), "email_ids"),
        (_remove_command(email_ids=("1", "1")), "duplicates"),
        (_remove_command(email_ids=tuple(str(i) for i in range(1, APPLICATION_LIMITS.mutation_uids + 2))), "at most"),
        (_remove_command(label_name=""), "label_name"),
        (_remove_command(label_name="Labels/Work"), "must not repeat"),
        (_remove_command(label_name="Work\r\nX"), "control characters"),
        (_remove_command(source_mailbox=""), "mailbox"),
    ],
)
def test_remove_label_command_rejects_invalid_input(command: RemoveLabelCommand, message: str) -> None:
    with pytest.raises(ValueError, match=message):
        command.validate()


@pytest.mark.asyncio
async def test_remove_label_service_invalidates_only_the_label_mailbox() -> None:
    authority, factory, projections, provider, projection = _mutation_ports(
        BatchMutationOutcome((TargetMutationOutcome("42", "succeeded"),))
    )
    service = RemoveLabelService(authority, factory, projections)

    outcome = await service.execute(_remove_command())

    assert outcome.targets("succeeded") == ["42"]
    # The source mailbox is untouched by the effect, so invalidating it would be
    # a false claim that INBOX changed.
    projection.invalidate.assert_awaited_once_with(("Labels/Work",))
    assert provider.remove_label.await_args.args[0].label_mailbox == "Labels/Work"


@pytest.mark.asyncio
async def test_remove_label_service_skips_invalidation_when_nothing_changed() -> None:
    authority, factory, projections, _provider, projection = _mutation_ports(
        BatchMutationOutcome((TargetMutationOutcome("42", "failed", "label-not-found"),))
    )
    service = RemoveLabelService(authority, factory, projections)

    outcome = await service.execute(_remove_command())

    assert outcome.outcomes[0].detail == "label-not-found"
    assert not outcome.reconciliation_needed
    projection.invalidate.assert_not_awaited()


@pytest.mark.asyncio
async def test_remove_label_service_reports_a_timeout_as_unknown() -> None:
    authority, factory, projections, _provider, _projection = _mutation_ports(TimeoutError())
    service = RemoveLabelService(authority, factory, projections)

    outcome = await service.execute(_remove_command(email_ids=("42", "43")))

    assert [item.status for item in outcome.outcomes] == ["unknown", "unknown"]
    assert {item.detail for item in outcome.outcomes} == {"provider-timeout"}
    assert outcome.reconciliation_needed


@pytest.mark.asyncio
async def test_remove_label_service_flags_reconciliation_when_projection_fails() -> None:
    authority, factory, projections, _provider, projection = _mutation_ports(
        BatchMutationOutcome((TargetMutationOutcome("42", "succeeded"),))
    )
    projection.invalidate = AsyncMock(side_effect=RuntimeError("projection down"))
    service = RemoveLabelService(authority, factory, projections)

    outcome = await service.execute(_remove_command())

    assert outcome.targets("succeeded") == ["42"]
    assert outcome.reconciliation_needed


# --------------------------------------------------------------------------
# Adapters
# --------------------------------------------------------------------------


def _read_adapter(client: Mock) -> ClassicReadProvider:
    return ClassicReadProvider(SimpleNamespace(incoming_client=client))  # pyright: ignore[reportArgumentType]


@pytest.mark.asyncio
async def test_read_adapter_forwards_mailbox_and_allowlist_for_message_id_lookup() -> None:
    client = Mock()
    client.fetch_message_id = AsyncMock(return_value=MESSAGE_ID)
    provider = _read_adapter(client)

    result = await provider.fetch_message_id(
        GetEmailLabelsQuery(account_name="work", email_id="42", mailbox="Archive"),
        _read_account(),
    )

    assert result == MESSAGE_ID
    client.fetch_message_id.assert_awaited_once_with("42", "Archive", allowed_senders=["allowed@example.test"])


@pytest.mark.asyncio
async def test_read_adapter_sanitizes_message_id_lookup_failures() -> None:
    client = Mock()
    client.fetch_message_id = AsyncMock(side_effect=RuntimeError("IMAP said NO to alice@example.test"))
    provider = _read_adapter(client)

    with pytest.raises(ReadProviderError) as excinfo:
        await provider.fetch_message_id(GetEmailLabelsQuery(account_name="work", email_id="42"), _read_account())

    assert "alice@example.test" not in str(excinfo.value)


@pytest.mark.asyncio
async def test_read_adapter_keeps_unusable_message_ids_distinguishable() -> None:
    client = Mock()
    client.search_message_id_in_mailboxes = AsyncMock(side_effect=ValueError("Message-ID must be printable ASCII"))
    provider = _read_adapter(client)

    with pytest.raises(ValueError, match="printable ASCII"):
        await provider.search_message_id_in_mailboxes("\x01", ("Labels/Work",))


@pytest.mark.asyncio
async def test_read_adapter_sanitizes_label_membership_failures() -> None:
    client = Mock()
    client.search_message_id_in_mailboxes = AsyncMock(side_effect=RuntimeError("SELECT Labels/Secret failed"))
    provider = _read_adapter(client)

    with pytest.raises(ReadProviderError) as excinfo:
        await provider.search_message_id_in_mailboxes(MESSAGE_ID, ("Labels/Work",))

    assert "Secret" not in str(excinfo.value)


@pytest.mark.asyncio
async def test_mutation_adapter_forwards_the_derived_label_mailbox_and_policy() -> None:
    client = Mock()
    client.remove_label_with_outcome = AsyncMock(
        return_value=BatchMutationOutcome((TargetMutationOutcome("42", "succeeded"),))
    )
    provider = ClassicMutationProvider(SimpleNamespace(incoming_client=client))  # pyright: ignore[reportArgumentType]
    account = MutationAccountSnapshot(
        account_name="work",
        mode="managed",
        allowed_senders=("allowed@example.test",),
        allowed_recipients=(),
        report_blocked_mutations=True,
        can_send=True,
    )

    await provider.remove_label(_remove_command(source_mailbox="Archive"), account)

    client.remove_label_with_outcome.assert_awaited_once_with(
        ["42"],
        "Labels/Work",
        "Archive",
        ["allowed@example.test"],
        True,
    )


# --------------------------------------------------------------------------
# Provider: Message-ID search value construction
# --------------------------------------------------------------------------


def test_quoted_message_id_is_always_quoted() -> None:
    assert _quoted_imap_message_id(MESSAGE_ID) == f'"{MESSAGE_ID}"'
    # Surrounding whitespace from an unfolded header never reaches the wire.
    assert _quoted_imap_message_id(f"  {MESSAGE_ID}\t") == f'"{MESSAGE_ID}"'


def test_quoted_message_id_escapes_rather_than_interpolates_quoted_specials() -> None:
    hostile = '<a"@b.test> ALL OR TEXT "x'
    assert _quoted_imap_message_id(hostile) == '"<a\\"@b.test> ALL OR TEXT \\"x"'
    assert _quoted_imap_message_id(r"<a\b@c.test>") == r'"<a\\b@c.test>"'


@pytest.mark.parametrize("value", ["", "   ", "<a\rb@c.test>", "<a\nb@c.test>", "<a\x7fb@c.test>", "<ünicode@c.test>"])
def test_quoted_message_id_rejects_values_that_cannot_be_quoted(value: str) -> None:
    with pytest.raises(ValueError, match="Message-ID must"):
        _quoted_imap_message_id(value)


# --------------------------------------------------------------------------
# Provider: EmailClient primitives
# --------------------------------------------------------------------------


@pytest.fixture
def email_client(email_server):  # email_server comes from conftest.py
    return EmailClient(email_server)


def _header_blocks(blocks: dict[str, bytes]) -> tuple[str, list]:
    """Build a FETCH response carrying one header literal per UID."""
    data: list = []
    for uid, raw_headers in blocks.items():
        data.append(f"1 FETCH (UID {uid} BODY[HEADER.FIELDS (X)] {{{len(raw_headers)}}}".encode())
        data.append(bytearray(raw_headers))
        data.append(b")")
    return "OK", data


def _fetch_response(uid: str, raw_headers: bytes) -> tuple[str, list]:
    return _header_blocks({uid: raw_headers})


def _make_mock_imap(**overrides):
    capabilities = overrides.pop("capabilities", ("IMAP4rev1", "UIDPLUS"))
    mock = AsyncMock()
    mock._client_task = asyncio.Future()
    mock._client_task.set_result(None)
    mock.wait_hello_from_server = AsyncMock()
    mock.login = AsyncMock(return_value=MagicMock(result="OK", lines=[]))
    mock.id = AsyncMock(return_value=MagicMock(result="OK"))
    mock.select = AsyncMock(return_value=("OK", []))
    mock.uid = AsyncMock(return_value=("OK", []))
    mock.uid_search = AsyncMock(return_value=("OK", [b""]))
    mock.logout = AsyncMock()
    mock.protocol = MagicMock(capabilities=capabilities)
    mock.protocol.capability = AsyncMock()
    for key, value in overrides.items():
        setattr(mock, key, value)
    return mock


@pytest.mark.asyncio
async def test_fetch_message_id_returns_the_header_value(email_client) -> None:
    mock_imap = _make_mock_imap()
    mock_imap.uid = AsyncMock(return_value=_fetch_response("42", f"Message-ID: {MESSAGE_ID}\r\n\r\n".encode()))

    with patch.object(email_client, "imap_class", return_value=mock_imap):
        assert await email_client.fetch_message_id("42", "INBOX") == MESSAGE_ID


@pytest.mark.asyncio
async def test_fetch_message_id_returns_none_without_the_header(email_client) -> None:
    mock_imap = _make_mock_imap()
    mock_imap.uid = AsyncMock(return_value=_fetch_response("42", b"Subject: no identifier\r\n\r\n"))

    with patch.object(email_client, "imap_class", return_value=mock_imap):
        assert await email_client.fetch_message_id("42", "INBOX") is None


@pytest.mark.asyncio
async def test_fetch_message_id_hides_blocked_senders_behind_the_missing_case(email_client) -> None:
    """A blocked sender yields None *and* never reaches the Message-ID fetch."""
    fetched: list[str] = []

    async def uid_side_effect(command, *args):
        if command == "fetch":
            fetched.append(args[1])
            return _fetch_response("42", b"From: blocked@denied.test\r\n\r\n")
        return "OK", []

    mock_imap = _make_mock_imap()
    mock_imap.uid = AsyncMock(side_effect=uid_side_effect)

    with patch.object(email_client, "imap_class", return_value=mock_imap):
        result = await email_client.fetch_message_id("42", "INBOX", allowed_senders=["*@allowed.test"])

    assert result is None
    assert all("MESSAGE-ID" not in item for item in fetched)


@pytest.mark.asyncio
async def test_fetch_message_id_rejects_a_non_canonical_uid(email_client) -> None:
    with pytest.raises(ValueError, match="IMAP UID"):
        await email_client.fetch_message_id("007", "INBOX")


@pytest.mark.asyncio
async def test_search_message_id_in_mailboxes_uses_the_quoted_search_value(email_client) -> None:
    mock_imap = _make_mock_imap()
    mock_imap.uid_search = AsyncMock(return_value=("OK", [b"7"]))

    with patch.object(email_client, "imap_class", return_value=mock_imap):
        found = await email_client.search_message_id_in_mailboxes(MESSAGE_ID, ("Labels/Work",))

    assert found == ("Labels/Work",)
    assert mock_imap.uid_search.await_args.args == ("HEADER", "Message-ID", f'"{MESSAGE_ID}"')


@pytest.mark.asyncio
async def test_search_message_id_in_mailboxes_skips_folders_it_cannot_use(email_client) -> None:
    async def select_side_effect(mailbox):
        return ("NO", []) if "Broken" in mailbox else ("OK", [])

    async def search_side_effect(*args, **_kwargs):
        del args
        return "OK", [b"7"] if selected[-1].strip('"') == "Labels/Work" else [b""]

    selected: list[str] = []

    async def tracking_select(mailbox):
        selected.append(mailbox)
        return await select_side_effect(mailbox)

    mock_imap = _make_mock_imap()
    mock_imap.select = AsyncMock(side_effect=tracking_select)
    mock_imap.uid_search = AsyncMock(side_effect=search_side_effect)

    with patch.object(email_client, "imap_class", return_value=mock_imap):
        found = await email_client.search_message_id_in_mailboxes(
            MESSAGE_ID, ("Labels/Broken", "Labels/Work", "Labels/Other")
        )

    # The unselectable folder is skipped instead of failing the whole lookup.
    assert found == ("Labels/Work",)


@pytest.mark.asyncio
async def test_search_message_id_in_mailboxes_skips_a_failing_search(email_client) -> None:
    mock_imap = _make_mock_imap()
    mock_imap.uid_search = AsyncMock(side_effect=[RuntimeError("SEARCH exploded"), ("OK", [b"7"])])

    with patch.object(email_client, "imap_class", return_value=mock_imap):
        found = await email_client.search_message_id_in_mailboxes(MESSAGE_ID, ("Labels/Work", "Labels/Personal"))

    assert found == ("Labels/Personal",)


@pytest.mark.asyncio
async def test_search_message_id_in_mailboxes_rejects_an_unusable_id_before_connecting(email_client) -> None:
    connected = Mock()
    with patch.object(email_client, "imap_class", connected):
        with pytest.raises(ValueError, match="Message-ID must"):
            await email_client.search_message_id_in_mailboxes("<a\rb@c.test>", ("Labels/Work",))
    connected.assert_not_called()


@pytest.mark.asyncio
async def test_search_message_id_in_mailboxes_short_circuits_on_no_folders(email_client) -> None:
    connected = Mock()
    with patch.object(email_client, "imap_class", connected):
        assert await email_client.search_message_id_in_mailboxes(MESSAGE_ID, ()) == ()
    connected.assert_not_called()


# --------------------------------------------------------------------------
# Provider: remove_label_with_outcome
# --------------------------------------------------------------------------


class _LabelMailboxFake:
    """Two-mailbox IMAP fake that makes label-scoped effects observable."""

    def __init__(
        self,
        *,
        message_ids: dict[str, str] | None = None,
        label_uids: dict[str, list[str]] | None = None,
        capabilities: tuple[str, ...] = ("IMAP4rev1", "UIDPLUS"),
        selectable: tuple[str, ...] = ("INBOX", "Labels/Work"),
        senders: dict[str, str] | None = None,
    ) -> None:
        self._client_task = asyncio.Future()
        self._client_task.set_result(None)
        self.protocol = SimpleNamespace(capabilities=capabilities, capability=AsyncMock())
        self._message_ids = {"42": MESSAGE_ID} if message_ids is None else message_ids
        self._label_uids = {MESSAGE_ID: ["900"]} if label_uids is None else label_uids
        self._selectable = selectable
        self._senders = senders or {}
        self.selected: str | None = None
        self.uid_calls: list[tuple[str, str, str | None]] = []
        self.mailbox_wide_expunge_called = False

    async def wait_hello_from_server(self):
        return None

    async def login(self, _username: str, _password: str):
        return SimpleNamespace(result="OK", lines=[])

    async def id(self, **_kwargs):
        return SimpleNamespace(result="OK")

    async def select(self, mailbox: str):
        name = mailbox.strip('"')
        if name not in self._selectable:
            return "NO", []
        self.selected = name
        return "OK", []

    async def uid_search(self, *criteria: str, charset: str | None = None):
        del charset
        assert criteria[0:2] == ("HEADER", "Message-ID")
        quoted = criteria[2]
        assert quoted.startswith('"') and quoted.endswith('"'), quoted
        message_id = quoted[1:-1].replace('\\"', '"').replace("\\\\", "\\")
        uids = self._label_uids.get(message_id, []) if self.selected != "INBOX" else []
        return "OK", [" ".join(uids).encode()]

    async def uid(self, command: str, *args: str):
        self.uid_calls.append((command, args[0], self.selected))
        if command == "fetch":
            requested = args[0].split(",")
            if "MESSAGE-ID" in args[1]:
                return _header_blocks({
                    uid: f"Message-ID: {self._message_ids[uid]}\r\n\r\n".encode()
                    for uid in requested
                    if uid in self._message_ids
                })
            return _header_blocks({
                uid: f"From: {self._senders.get(uid, 'allowed@example.test')}\r\n\r\n".encode() for uid in requested
            })
        return "OK", []

    async def expunge(self):
        self.mailbox_wide_expunge_called = True
        return "OK", []

    async def logout(self):
        return "OK", []


def _uid_calls(fake: _LabelMailboxFake, command: str) -> list[tuple[str, str | None]]:
    return [(uids, mailbox) for name, uids, mailbox in fake.uid_calls if name == command]


@pytest.mark.asyncio
async def test_remove_label_deletes_only_the_label_copy(email_client) -> None:
    fake = _LabelMailboxFake()

    with patch.object(email_client, "imap_class", return_value=fake):
        outcome = await email_client.remove_label_with_outcome(["42"], "Labels/Work", "INBOX")

    assert outcome.targets("succeeded") == ["42"]
    # Every effect names the label's own UID inside the label mailbox; the UID the
    # caller passed is never flagged or expunged.
    assert _uid_calls(fake, "store") == [("900", "Labels/Work")]
    assert _uid_calls(fake, "expunge") == [("900", "Labels/Work")]
    assert not fake.mailbox_wide_expunge_called


@pytest.mark.asyncio
async def test_remove_label_expunges_every_copy_of_one_message(email_client) -> None:
    fake = _LabelMailboxFake(label_uids={MESSAGE_ID: ["900", "901"]})

    with patch.object(email_client, "imap_class", return_value=fake):
        outcome = await email_client.remove_label_with_outcome(["42"], "Labels/Work", "INBOX")

    assert outcome.targets("succeeded") == ["42"]
    assert _uid_calls(fake, "store") == [("900,901", "Labels/Work")]
    assert _uid_calls(fake, "expunge") == [("900,901", "Labels/Work")]


@pytest.mark.asyncio
async def test_remove_label_reports_a_message_without_an_identifier(email_client) -> None:
    fake = _LabelMailboxFake(message_ids={})

    with patch.object(email_client, "imap_class", return_value=fake):
        outcome = await email_client.remove_label_with_outcome(["42"], "Labels/Work", "INBOX")

    assert outcome.outcomes == (TargetMutationOutcome("42", "failed", "message-id-missing"),)
    assert _uid_calls(fake, "store") == []


@pytest.mark.asyncio
async def test_remove_label_reports_a_message_the_label_does_not_hold(email_client) -> None:
    fake = _LabelMailboxFake(label_uids={})

    with patch.object(email_client, "imap_class", return_value=fake):
        outcome = await email_client.remove_label_with_outcome(["42"], "Labels/Work", "INBOX")

    assert outcome.outcomes == (TargetMutationOutcome("42", "failed", "label-not-found"),)
    assert _uid_calls(fake, "store") == []


@pytest.mark.asyncio
async def test_remove_label_reports_an_unavailable_label_mailbox(email_client) -> None:
    fake = _LabelMailboxFake(selectable=("INBOX",))

    with patch.object(email_client, "imap_class", return_value=fake):
        outcome = await email_client.remove_label_with_outcome(["42"], "Labels/Work", "INBOX")

    assert outcome.outcomes == (TargetMutationOutcome("42", "failed", "label-unavailable"),)
    assert _uid_calls(fake, "store") == []


@pytest.mark.asyncio
async def test_remove_label_refuses_to_delete_without_uidplus(email_client) -> None:
    fake = _LabelMailboxFake(capabilities=("IMAP4rev1",))

    with patch.object(email_client, "imap_class", return_value=fake):
        outcome = await email_client.remove_label_with_outcome(["42"], "Labels/Work", "INBOX")

    assert outcome.outcomes == (TargetMutationOutcome("42", "failed", "uidplus-unavailable"),)
    assert _uid_calls(fake, "store") == []
    assert not fake.mailbox_wide_expunge_called


@pytest.mark.asyncio
async def test_remove_label_treats_a_blocked_sender_as_a_silent_no_op(email_client) -> None:
    fake = _LabelMailboxFake(senders={"42": "blocked@denied.test"})

    with patch.object(email_client, "imap_class", return_value=fake):
        outcome = await email_client.remove_label_with_outcome(
            ["42"], "Labels/Work", "INBOX", ["*@allowed.test"], False
        )

    assert outcome.outcomes == (TargetMutationOutcome("42", "succeeded", None),)
    assert _uid_calls(fake, "store") == []


@pytest.mark.asyncio
async def test_remove_label_reports_a_blocked_sender_when_configured(email_client) -> None:
    fake = _LabelMailboxFake(senders={"42": "blocked@denied.test"})

    with patch.object(email_client, "imap_class", return_value=fake):
        outcome = await email_client.remove_label_with_outcome(["42"], "Labels/Work", "INBOX", ["*@allowed.test"], True)

    assert outcome.outcomes == (TargetMutationOutcome("42", "failed", "sender-policy"),)
    assert _uid_calls(fake, "store") == []


@pytest.mark.asyncio
async def test_remove_label_rejects_a_non_canonical_uid(email_client) -> None:
    with pytest.raises(ValueError, match="IMAP UID"):
        await email_client.remove_label_with_outcome(["1 OR 2"], "Labels/Work", "INBOX")


@pytest.mark.asyncio
async def test_remove_label_preserves_caller_order_across_mixed_outcomes(email_client) -> None:
    fake = _LabelMailboxFake(
        message_ids={"42": MESSAGE_ID, "44": "<other@example.test>"},
        label_uids={MESSAGE_ID: ["900"]},
    )

    with patch.object(email_client, "imap_class", return_value=fake):
        outcome = await email_client.remove_label_with_outcome(["42", "43", "44"], "Labels/Work", "INBOX")

    assert [item.target for item in outcome.outcomes] == ["42", "43", "44"]
    assert [item.status for item in outcome.outcomes] == ["succeeded", "failed", "failed"]
    assert [item.detail for item in outcome.outcomes] == [None, "message-id-missing", "label-not-found"]


# --------------------------------------------------------------------------
# MCP tools
# --------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_list_labels_tool_maps_the_query() -> None:
    handler = AsyncMock(return_value=[])
    with patch("mcp_email_server.app.list_labels_query", handler):
        assert await list_labels(account_name="work") == []
    assert handler.await_args.args[0] == ListLabelsQuery(account_name="work")


@pytest.mark.asyncio
async def test_get_email_labels_tool_maps_the_query() -> None:
    handler = AsyncMock(return_value=["Work"])
    with patch("mcp_email_server.app.get_email_labels_query", handler):
        assert await get_email_labels(account_name="work", email_id="42", mailbox="Archive") == ["Work"]
    assert handler.await_args.args[0] == GetEmailLabelsQuery(account_name="work", email_id="42", mailbox="Archive")


@pytest.mark.asyncio
async def test_remove_label_tool_reports_a_clean_success() -> None:
    handler = AsyncMock(
        return_value=BatchMutationOutcome((
            TargetMutationOutcome("42", "succeeded"),
            TargetMutationOutcome("43", "succeeded"),
        ))
    )
    with patch("mcp_email_server.app.remove_label_command", handler):
        result = await remove_label(account_name="work", email_ids=["42", "43"], label_name="Work")

    assert result == "Successfully removed label 'Work' from 2 email(s)"
    command = handler.await_args.args[0]
    assert (command.label_name, command.source_mailbox) == ("Work", "INBOX")


@pytest.mark.asyncio
async def test_remove_label_tool_surfaces_reviewed_failure_details() -> None:
    """Unlike the generic batch tools, a label removal explains why an ID failed."""
    handler = AsyncMock(
        return_value=BatchMutationOutcome((
            TargetMutationOutcome("42", "succeeded"),
            TargetMutationOutcome("43", "failed", "label-not-found"),
            TargetMutationOutcome("44", "failed", "message-id-missing"),
        ))
    )
    with patch("mcp_email_server.app.remove_label_command", handler):
        result = await remove_label(account_name="work", email_ids=["42", "43", "44"], label_name="Work")

    assert result == ("Remove-label result [succeeded: 42; failed: 43 (label-not-found), 44 (message-id-missing)]")


@pytest.mark.asyncio
async def test_remove_label_tool_omits_detail_outside_the_reviewed_set() -> None:
    handler = AsyncMock(
        return_value=BatchMutationOutcome((TargetMutationOutcome("42", "failed", "NO [AUTHENTICATIONFAILED] alice"),))
    )
    with patch("mcp_email_server.app.remove_label_command", handler):
        result = await remove_label(account_name="work", email_ids=["42"], label_name="Work")

    assert result == "Remove-label result [failed: 42]"


@pytest.mark.asyncio
async def test_remove_label_tool_reports_reconciliation() -> None:
    handler = AsyncMock(return_value=BatchMutationOutcome((TargetMutationOutcome("42", "unknown", "expunge-unknown"),)))
    with patch("mcp_email_server.app.remove_label_command", handler):
        result = await remove_label(account_name="work", email_ids=["42"], label_name="Work")

    assert result == ("Remove-label result [unknown: 42 (expunge-unknown); warning: reconciliation needed]")
