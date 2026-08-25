"""Trash-first deletion: discovery, placement, and the permanent fallback.

Deleting is the one mutation whose destination is decided by the server rather
than the caller, so these tests pin three separate things: that discovery
resolves the Trash mailbox the way the RFC 6154 special-use rules say it should,
that an ambiguous discovery never resolves in favour of the irreversible branch,
and that the caller can always tell which branch ran.
"""

from __future__ import annotations

import asyncio
from dataclasses import replace
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from mcp_email_server.adapters.mutations import ClassicMutationProvider
from mcp_email_server.application import mutations as mutations_module
from mcp_email_server.application.limits import APPLICATION_LIMITS
from mcp_email_server.application.mutations import (
    BatchMutationOutcome,
    DeleteCommand,
    MutationAccountSnapshot,
    MutationProviderAccess,
    MutationProviderError,
    MutationServices,
    TargetMutationOutcome,
)
from mcp_email_server.config import EmailServer, EmailSettings
from mcp_email_server.emails.classic import ClassicEmailHandler
from mcp_email_server.emails.models import MailboxInfo


@pytest.fixture
def email_settings() -> EmailSettings:
    return EmailSettings(
        account_name="test_account",
        full_name="Test User",
        email_address="test@example.com",
        incoming=EmailServer(
            user_name="test_user",
            password="test_password",
            host="imap.example.com",
            port=993,
            use_ssl=True,
        ),
        outgoing=EmailServer(
            user_name="test_user",
            password="test_password",
            host="smtp.example.com",
            port=465,
            use_ssl=True,
        ),
    )


@pytest.fixture
def classic_handler(email_settings: EmailSettings) -> ClassicEmailHandler:
    return ClassicEmailHandler(email_settings)


def _account(**changes: object) -> MutationAccountSnapshot:
    account = MutationAccountSnapshot(
        account_name="primary",
        mode="managed",
        allowed_senders=(),
        allowed_recipients=(),
        report_blocked_mutations=False,
        can_send=True,
    )
    return replace(account, **changes)


def _batch(*outcomes: TargetMutationOutcome) -> BatchMutationOutcome:
    return BatchMutationOutcome(outcomes)


def _provider(trash_mailbox: str | None = "Trash") -> MagicMock:
    """A provider that resolves ``trash_mailbox`` and succeeds at whatever it is asked."""
    provider = MagicMock()
    provider.find_trash_mailbox = AsyncMock(return_value=trash_mailbox)
    provider.move = AsyncMock(return_value=_batch(TargetMutationOutcome("7", "succeeded")))
    provider.delete = AsyncMock(return_value=_batch(TargetMutationOutcome("7", "succeeded")))
    return provider


def _services(
    provider: MagicMock,
    *,
    account: MutationAccountSnapshot | None = None,
) -> tuple[MutationServices, MagicMock, MagicMock]:
    current = account if account is not None else _account()
    authority = MagicMock()
    authority.resolve.return_value = current
    factory = MagicMock()
    factory.open.return_value = MutationProviderAccess(current, provider)
    projection = MagicMock()
    projection.invalidate = AsyncMock()
    projections = MagicMock()
    projections.open = AsyncMock(return_value=projection)
    return MutationServices.compose(authority, factory, projections), factory, projection


# ---------------------------------------------------------------------------
# Provider discovery
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_the_trash_flag_wins_over_any_common_name(classic_handler: ClassicEmailHandler) -> None:
    """RFC 6154 is authoritative: the server's own name for its trash is used verbatim."""
    mailboxes = [
        MailboxInfo(name="INBOX", delimiter="/", flags=[]),
        MailboxInfo(name="Trash", delimiter="/", flags=[]),
        MailboxInfo(name="Papierkorb", delimiter="/", flags=["\\Trash"]),
    ]

    with patch.object(classic_handler.incoming_client, "list_mailboxes", AsyncMock(return_value=mailboxes)):
        assert await classic_handler._find_trash_folder() == "Papierkorb"


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("mailbox_name", "expected"),
    [
        ("Trash", "Trash"),
        ("trash", "trash"),
        ("TRASH", "TRASH"),
        ("Deleted Items", "Deleted Items"),
        ("deleted messages", "deleted messages"),
        ("[Gmail]/Trash", "[Gmail]/Trash"),
        ("INBOX.Trash", "INBOX.Trash"),
    ],
)
async def test_common_trash_names_match_case_insensitively_and_keep_server_spelling(
    classic_handler: ClassicEmailHandler,
    mailbox_name: str,
    expected: str,
) -> None:
    """The candidate list is matched case-insensitively, but SELECT gets the server's spelling."""
    mailboxes = [
        MailboxInfo(name="INBOX", delimiter="/", flags=[]),
        MailboxInfo(name=mailbox_name, delimiter="/", flags=[]),
    ]

    with patch.object(classic_handler.incoming_client, "list_mailboxes", AsyncMock(return_value=mailboxes)):
        assert await classic_handler._find_trash_folder() == expected


@pytest.mark.asyncio
async def test_an_unrelated_mailbox_is_never_mistaken_for_trash(classic_handler: ClassicEmailHandler) -> None:
    """Substring lookalikes stay unmatched; nothing here should reach a destructive branch by accident."""
    mailboxes = [
        MailboxInfo(name="INBOX", delimiter="/", flags=[]),
        MailboxInfo(name="Trashed Drafts", delimiter="/", flags=[]),
        MailboxInfo(name="Not Deleted Items", delimiter="/", flags=["\\NoTrashHere"]),
    ]

    with patch.object(classic_handler.incoming_client, "list_mailboxes", AsyncMock(return_value=mailboxes)):
        assert await classic_handler._find_trash_folder() is None


@pytest.mark.asyncio
async def test_trash_and_archive_resolutions_share_one_list(classic_handler: ClassicEmailHandler) -> None:
    """Registering a third special-use kind must not cost a third LIST per lookup."""
    mailboxes = [
        MailboxInfo(name="INBOX", delimiter="/", flags=[]),
        MailboxInfo(name="Bin", delimiter="/", flags=["\\Trash"]),
        MailboxInfo(name="Archive", delimiter="/", flags=[]),
    ]
    mock_list = AsyncMock(return_value=mailboxes)

    with patch.object(classic_handler.incoming_client, "list_mailboxes", mock_list):
        assert await classic_handler._find_trash_folder() == "Bin"
        assert await classic_handler._find_trash_folder() == "Bin"
        assert await classic_handler._find_archive_folder() == "Archive"

    # One LIST per kind, cached thereafter; the kinds do not share a cache entry.
    assert mock_list.await_count == 2


@pytest.mark.asyncio
async def test_a_failed_trash_lookup_propagates_rather_than_reporting_no_trash(
    classic_handler: ClassicEmailHandler,
) -> None:
    """A broken LIST must not be readable as "this account has no Trash"."""
    with patch.object(
        classic_handler.incoming_client,
        "list_mailboxes",
        AsyncMock(side_effect=ConnectionError("LIST failed")),
    ):
        with pytest.raises(ConnectionError):
            await classic_handler._find_trash_folder()


# ---------------------------------------------------------------------------
# Adapter
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_adapter_reports_no_trash_when_none_exists() -> None:
    handler = MagicMock()
    handler._find_trash_folder = AsyncMock(return_value=None)

    assert await ClassicMutationProvider(handler).find_trash_mailbox("INBOX") is None


@pytest.mark.asyncio
@pytest.mark.parametrize("source_mailbox", ["Trash", "trash", "TRASH"])
async def test_adapter_reports_no_trash_when_deleting_from_trash(source_mailbox: str) -> None:
    """Emptying the trash must actually empty it, whatever case the caller spelled it in."""
    handler = MagicMock()
    handler._find_trash_folder = AsyncMock(return_value="Trash")

    assert await ClassicMutationProvider(handler).find_trash_mailbox(source_mailbox) is None


@pytest.mark.asyncio
async def test_adapter_returns_a_distinct_trash_mailbox() -> None:
    handler = MagicMock()
    handler._find_trash_folder = AsyncMock(return_value="Deleted Items")

    assert await ClassicMutationProvider(handler).find_trash_mailbox("INBOX") == "Deleted Items"


@pytest.mark.asyncio
async def test_adapter_sanitizes_a_failed_lookup_into_a_provider_error() -> None:
    """The service can only fail closed on a discovery error if the error survives the adapter."""
    handler = MagicMock()
    handler._find_trash_folder = AsyncMock(side_effect=ConnectionError("LIST failed"))

    with pytest.raises(MutationProviderError):
        await ClassicMutationProvider(handler).find_trash_mailbox("INBOX")


# ---------------------------------------------------------------------------
# Application workflow
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_delete_moves_to_the_discovered_trash_mailbox() -> None:
    provider = _provider("Trash")
    services, _, projection = _services(provider)

    result = await services.delete.execute(DeleteCommand("primary", ("7",)))

    provider.move.assert_awaited_once()
    provider.delete.assert_not_awaited()
    move = provider.move.await_args.args[0]
    assert (move.source_mailbox, move.destination_mailbox) == ("INBOX", "Trash")
    assert result.trash_mailbox == "Trash"
    assert result.batch.targets("succeeded") == ["7"]
    # A move empties one mailbox and fills another, so both projections are stale.
    projection.invalidate.assert_awaited_once_with(("INBOX", "Trash"))


@pytest.mark.asyncio
async def test_delete_falls_back_to_permanent_removal_without_a_trash_mailbox() -> None:
    provider = _provider(None)
    services, factory, projection = _services(provider)

    result = await services.delete.execute(DeleteCommand("primary", ("7",)))

    provider.delete.assert_awaited_once()
    provider.move.assert_not_awaited()
    assert result.trash_mailbox is None
    assert result.batch.targets("succeeded") == ["7"]
    projection.invalidate.assert_awaited_once_with(("INBOX",))
    # The irreversible branch gets the same re-resolved authority as the move branch.
    assert factory.open.call_count == 2


@pytest.mark.asyncio
async def test_deleting_from_trash_is_permanent() -> None:
    """The provider answers "nowhere further to move to", and the workflow honours it."""
    provider = _provider(None)
    services, _, _ = _services(provider)

    result = await services.delete.execute(DeleteCommand("primary", ("7",), "Trash"))

    provider.find_trash_mailbox.assert_awaited_once_with("Trash")
    provider.delete.assert_awaited_once()
    assert result.trash_mailbox is None


@pytest.mark.asyncio
async def test_authority_is_reresolved_between_discovery_and_the_delete_effect() -> None:
    """Discovery and the effect are separate provider accesses, as every mutation requires."""
    provider = _provider("Trash")
    services, factory, _ = _services(provider)

    await services.delete.execute(DeleteCommand("primary", ("7",)))

    assert factory.open.call_count == 2
    assert factory.open.call_args_list[0] == factory.open.call_args_list[1]
    assert factory.open.call_args_list[-1].kwargs == {"expected_mode": "managed", "purpose": "incoming"}


@pytest.mark.asyncio
async def test_the_move_carries_the_accounts_sender_allowlist_policy() -> None:
    """Trash-first must not become a way around the allowlist the permanent path honours."""
    account = _account(allowed_senders=("*@example.test",), report_blocked_mutations=True)
    provider = _provider("Trash")
    services, _, _ = _services(provider, account=account)

    await services.delete.execute(DeleteCommand("primary", ("7",)))

    assert provider.move.await_args.args[1] is account


@pytest.mark.asyncio
async def test_an_ambiguous_discovery_never_falls_through_to_a_permanent_delete(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A lookup that times out is not evidence of "no Trash"; nothing may be deleted."""
    monkeypatch.setattr(
        mutations_module,
        "APPLICATION_LIMITS",
        replace(APPLICATION_LIMITS, provider_timeout_seconds=0.01),
    )
    provider = _provider("Trash")

    async def _hang(*_args: object, **_kwargs: object) -> str:
        await asyncio.Event().wait()
        raise AssertionError("unreachable")

    provider.find_trash_mailbox = AsyncMock(side_effect=_hang)
    services, _, projection = _services(provider)

    with pytest.raises(MutationProviderError, match="trash mailbox discovery timed out"):
        await services.delete.execute(DeleteCommand("primary", ("7",)))

    provider.delete.assert_not_awaited()
    provider.move.assert_not_awaited()
    projection.invalidate.assert_not_awaited()


@pytest.mark.asyncio
async def test_a_failed_discovery_never_falls_through_to_a_permanent_delete() -> None:
    """Same rule for an outright provider failure: fail closed, delete nothing."""
    provider = _provider("Trash")
    provider.find_trash_mailbox = AsyncMock(side_effect=MutationProviderError("provider_failure"))
    services, _, projection = _services(provider)

    with pytest.raises(MutationProviderError):
        await services.delete.execute(DeleteCommand("primary", ("7",)))

    provider.delete.assert_not_awaited()
    provider.move.assert_not_awaited()
    projection.invalidate.assert_not_awaited()


@pytest.mark.asyncio
async def test_a_denied_authority_between_discovery_and_the_move_deletes_nothing() -> None:
    """Authority revoked after discovery aborts before the effect, not halfway through it."""
    provider = _provider("Trash")
    services, factory, _ = _services(provider)
    factory.open.side_effect = [factory.open.return_value, PermissionError("account is read-only")]

    with pytest.raises(PermissionError):
        await services.delete.execute(DeleteCommand("primary", ("7",)))

    provider.move.assert_not_awaited()
    provider.delete.assert_not_awaited()


@pytest.mark.asyncio
async def test_a_timed_out_move_is_unknown_rather_than_reported_as_a_permanent_delete(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        mutations_module,
        "APPLICATION_LIMITS",
        replace(APPLICATION_LIMITS, provider_timeout_seconds=0.01),
    )

    async def _hang(*_args: object, **_kwargs: object) -> BatchMutationOutcome:
        await asyncio.Event().wait()
        raise AssertionError("unreachable")

    provider = _provider("Trash")
    provider.move = AsyncMock(side_effect=_hang)
    services, _, projection = _services(provider)

    result = await services.delete.execute(DeleteCommand("primary", ("7",)))

    assert result.trash_mailbox == "Trash"
    assert result.batch.targets("unknown") == ["7"]
    assert result.batch.reconciliation_needed is True
    # An unknown move may already have landed, so both ends are stale.
    projection.invalidate.assert_awaited_once_with(("INBOX", "Trash"))


@pytest.mark.asyncio
async def test_a_known_move_failure_leaves_both_projections_alone() -> None:
    provider = _provider("Trash")
    provider.move = AsyncMock(return_value=_batch(TargetMutationOutcome("7", "failed", "copy-rejected")))
    services, _, projection = _services(provider)

    result = await services.delete.execute(DeleteCommand("primary", ("7",)))

    assert result.batch.targets("failed") == ["7"]
    assert result.trash_mailbox == "Trash"
    projection.invalidate.assert_not_awaited()


@pytest.mark.asyncio
async def test_an_uninvalidatable_projection_warns_without_erasing_the_move() -> None:
    provider = _provider("Trash")
    services, _, projection = _services(provider)
    projection.invalidate = AsyncMock(side_effect=RuntimeError("projection unavailable"))

    result = await services.delete.execute(DeleteCommand("primary", ("7",)))

    assert result.batch.targets("succeeded") == ["7"]
    assert result.batch.reconciliation_needed is True
    assert result.trash_mailbox == "Trash"


@pytest.mark.asyncio
async def test_an_invalid_command_is_rejected_before_any_discovery() -> None:
    """Validation still precedes every provider access, discovery included."""
    provider = _provider("Trash")
    services, factory, _ = _services(provider)

    with pytest.raises(ValueError, match="mailbox"):
        await services.delete.execute(DeleteCommand("primary", ("7",), "INBOX\r\nEXPUNGE"))

    factory.open.assert_not_called()
    provider.find_trash_mailbox.assert_not_awaited()
