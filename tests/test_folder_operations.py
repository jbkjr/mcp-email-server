"""Folder-shape mutations: copy_emails and create/delete/rename_folder.

Covers the four layers the tools pass through: the application services
(authority, gating, projection invalidation, timeout mapping), the classic
provider adapter (policy threading, error sanitization, special-use cache
invalidation), the IMAP provider itself (quoting, allowlist gating, outcome
classification), and the MCP tool dispatch/result strings.
"""

from __future__ import annotations

import asyncio
from dataclasses import replace
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from mcp_email_server.adapters.mutations import ClassicMutationProvider
from mcp_email_server.app import copy_emails, create_folder, delete_folder, rename_folder
from mcp_email_server.application.limits import APPLICATION_LIMITS
from mcp_email_server.application.mutations import (
    BatchMutationOutcome,
    CopyCommand,
    CreateFolderCommand,
    DeleteFolderCommand,
    FolderMutationOutcome,
    MutationAccountSnapshot,
    MutationProjectionError,
    MutationProviderAccess,
    MutationProviderError,
    MutationServices,
    RenameFolderCommand,
    TargetMutationOutcome,
)
from mcp_email_server.emails.classic import EmailClient

FOLDER_MANAGEMENT_DENIED = r"Folder management is disabled"


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


def _services(
    *,
    account: MutationAccountSnapshot | None = None,
    opened_account: MutationAccountSnapshot | None = None,
    provider: MagicMock | None = None,
    projection: MagicMock | None = None,
) -> tuple[MutationServices, MagicMock, MagicMock, MagicMock]:
    """Compose services over mocks, allowing the opened account to differ.

    A separate ``opened_account`` is what makes the "policy revoked between
    resolution and provider construction" case expressible at all.
    """
    resolved = account if account is not None else _account()
    authority = MagicMock()
    authority.resolve.return_value = resolved
    selected_provider = provider if provider is not None else MagicMock()
    factory = MagicMock()
    factory.open.return_value = MutationProviderAccess(
        opened_account if opened_account is not None else resolved,
        selected_provider,
    )
    selected_projection = projection if projection is not None else MagicMock()
    if projection is None:
        selected_projection.invalidate = AsyncMock()
    projections = MagicMock()
    projections.open = AsyncMock(return_value=selected_projection)
    return (
        MutationServices.compose(authority, factory, projections),
        authority,
        factory,
        selected_projection,
    )


def _enabled() -> MutationAccountSnapshot:
    return _account(mode="legacy", enable_folder_management=True)


# --------------------------------------------------------------------------- #
# Command contracts
# --------------------------------------------------------------------------- #


def test_copy_command_allows_the_same_source_and_destination() -> None:
    """A copy is additive, so unlike a move it may target its own mailbox."""
    CopyCommand("primary", ("1",), "INBOX", "INBOX").validate()


@pytest.mark.parametrize(
    ("command", "message"),
    [
        (CopyCommand("", ("1",), "INBOX", "Archive"), "account_name"),
        (CopyCommand("primary", (), "INBOX", "Archive"), "email_ids must not be empty"),
        (CopyCommand("primary", ("1", "1"), "INBOX", "Archive"), "must not contain duplicates"),
        (CopyCommand("primary", ("0",), "INBOX", "Archive"), "email_ids item"),
        (CopyCommand("primary", ("1",), "IN\nBOX", "Archive"), "mailbox"),
        (CopyCommand("primary", ("1",), "INBOX", "Arch\x00ive"), "mailbox"),
        (CreateFolderCommand("primary", ""), "mailbox"),
        (CreateFolderCommand("primary", "Bad\nName"), "mailbox"),
        (DeleteFolderCommand("primary", "Bad\x00Name"), "mailbox"),
        (RenameFolderCommand("primary", "Same", "Same"), "old_name and new_name must differ"),
        (RenameFolderCommand("primary", "Old", "New\nName"), "mailbox"),
    ],
)
def test_folder_commands_reject_unsupported_contracts(command: object, message: str) -> None:
    with pytest.raises(ValueError, match=message):
        command.validate()  # type: ignore[attr-defined]


def test_folder_commands_reject_oversized_names() -> None:
    oversized = "a" * (APPLICATION_LIMITS.mailbox_bytes + 1)

    with pytest.raises(ValueError, match="mailbox"):
        CreateFolderCommand("primary", oversized).validate()
    with pytest.raises(ValueError, match="mailbox"):
        CopyCommand("primary", ("1",), "INBOX", oversized).validate()


def test_copy_command_accepts_the_full_supported_batch() -> None:
    email_ids = tuple(str(value) for value in range(1, APPLICATION_LIMITS.mutation_uids + 1))

    CopyCommand("primary", email_ids, "INBOX", "Archive").validate()

    with pytest.raises(ValueError, match="at most"):
        CopyCommand("primary", (*email_ids, "999999"), "INBOX", "Archive").validate()


# --------------------------------------------------------------------------- #
# CopyService
# --------------------------------------------------------------------------- #


@pytest.mark.asyncio
async def test_copy_invalidates_only_the_destination_mailbox() -> None:
    """COPY leaves the source untouched, so its projection is still accurate."""
    provider = MagicMock()
    provider.copy = AsyncMock(return_value=BatchMutationOutcome((TargetMutationOutcome("11", "succeeded"),)))
    services, _, factory, projection = _services(provider=provider)

    result = await services.copy.execute(CopyCommand("primary", ("11",), "INBOX", "Archive"))

    assert result.targets("succeeded") == ["11"]
    assert result.reconciliation_needed is False
    factory.open.assert_called_once_with("primary", expected_mode="managed", purpose="incoming")
    projection.invalidate.assert_awaited_once_with(("Archive",))


@pytest.mark.asyncio
async def test_copy_timeout_maps_every_uid_to_unknown_without_replay() -> None:
    provider = MagicMock()
    provider.copy = AsyncMock(side_effect=TimeoutError())
    services, _, _, projection = _services(provider=provider)

    result = await services.copy.execute(CopyCommand("primary", ("11", "12"), "INBOX", "Archive"))

    assert result.targets("unknown") == ["11", "12"]
    assert [item.detail for item in result.outcomes] == ["provider-timeout", "provider-timeout"]
    assert result.reconciliation_needed is True
    provider.copy.assert_awaited_once()
    projection.invalidate.assert_awaited_once_with(("Archive",))


@pytest.mark.asyncio
async def test_copy_known_failure_does_not_invalidate_the_projection() -> None:
    provider = MagicMock()
    provider.copy = AsyncMock(
        return_value=BatchMutationOutcome((TargetMutationOutcome("11", "failed", "copy-rejected"),))
    )
    services, _, _, projection = _services(provider=provider)

    result = await services.copy.execute(CopyCommand("primary", ("11",), "INBOX", "Archive"))

    assert result.targets("failed") == ["11"]
    projection.invalidate.assert_not_awaited()


@pytest.mark.asyncio
async def test_copy_success_survives_projection_failure_with_a_warning() -> None:
    provider = MagicMock()
    provider.copy = AsyncMock(return_value=BatchMutationOutcome((TargetMutationOutcome("11", "succeeded"),)))
    projection = MagicMock()
    projection.invalidate = AsyncMock(side_effect=MutationProjectionError("unavailable"))
    services, _, _, _ = _services(provider=provider, projection=projection)

    result = await services.copy.execute(CopyCommand("primary", ("11",), "INBOX", "Archive"))

    assert result.targets("succeeded") == ["11"]
    assert result.reconciliation_needed is True


@pytest.mark.asyncio
async def test_copy_is_not_gated_by_folder_management() -> None:
    """Copying messages does not change the folder layout, so it stays ungated."""
    provider = MagicMock()
    provider.copy = AsyncMock(return_value=BatchMutationOutcome((TargetMutationOutcome("11", "succeeded"),)))
    services, _, _, _ = _services(account=_account(enable_folder_management=False), provider=provider)

    result = await services.copy.execute(CopyCommand("primary", ("11",), "INBOX", "Archive"))

    assert result.targets("succeeded") == ["11"]


# --------------------------------------------------------------------------- #
# Folder-shape services: authority gating
# --------------------------------------------------------------------------- #


@pytest.mark.parametrize(
    ("service_name", "command"),
    [
        ("create_folder", CreateFolderCommand("primary", "Projects")),
        ("delete_folder", DeleteFolderCommand("primary", "Projects")),
        ("rename_folder", RenameFolderCommand("primary", "Projects", "Archive/Projects")),
    ],
)
@pytest.mark.asyncio
async def test_folder_mutations_are_denied_before_any_provider_is_opened(
    service_name: str,
    command: object,
) -> None:
    provider = MagicMock()
    services, authority, factory, projection = _services(
        account=_account(mode="legacy", enable_folder_management=False),
        provider=provider,
    )

    with pytest.raises(PermissionError, match=FOLDER_MANAGEMENT_DENIED):
        await getattr(services, service_name).execute(command)

    authority.resolve.assert_called_once_with("primary")
    factory.open.assert_not_called()
    projection.invalidate.assert_not_awaited()


@pytest.mark.parametrize(
    ("service_name", "command", "provider_method"),
    [
        ("create_folder", CreateFolderCommand("primary", "Projects"), "create_folder"),
        ("delete_folder", DeleteFolderCommand("primary", "Projects"), "delete_folder"),
        ("rename_folder", RenameFolderCommand("primary", "Projects", "Renamed"), "rename_folder"),
    ],
)
@pytest.mark.asyncio
async def test_folder_mutations_are_denied_when_the_opened_account_revokes_the_policy(
    service_name: str,
    command: object,
    provider_method: str,
) -> None:
    """The pre-open snapshot may be stale; the opened account decides the effect."""
    provider = MagicMock()
    setattr(provider, provider_method, AsyncMock(return_value=FolderMutationOutcome("succeeded")))
    services, _, factory, _ = _services(
        account=_enabled(),
        opened_account=_account(mode="legacy", enable_folder_management=False),
        provider=provider,
    )

    with pytest.raises(PermissionError, match=FOLDER_MANAGEMENT_DENIED):
        await getattr(services, service_name).execute(command)

    factory.open.assert_called_once()
    getattr(provider, provider_method).assert_not_awaited()


def test_folder_management_denial_names_both_configuration_surfaces() -> None:
    services, _, _, _ = _services(account=_account(mode="legacy", enable_folder_management=False))

    with pytest.raises(PermissionError) as caught:
        asyncio.run(services.create_folder.execute(CreateFolderCommand("primary", "Projects")))

    message = str(caught.value)
    assert "enable_folder_management=true" in message
    assert "MCP_EMAIL_SERVER_ENABLE_FOLDER_MANAGEMENT=true" in message


@pytest.mark.asyncio
async def test_folder_mutations_validate_the_command_before_resolving_authority() -> None:
    services, authority, factory, _ = _services(account=_enabled())

    with pytest.raises(ValueError, match="mailbox"):
        await services.create_folder.execute(CreateFolderCommand("primary", "Bad\nName"))

    authority.resolve.assert_not_called()
    factory.open.assert_not_called()


# --------------------------------------------------------------------------- #
# Folder-shape services: effects and projection
# --------------------------------------------------------------------------- #


@pytest.mark.asyncio
async def test_create_folder_reresolves_authority_and_invalidates_its_mailbox() -> None:
    provider = MagicMock()
    provider.create_folder = AsyncMock(return_value=FolderMutationOutcome("succeeded"))
    services, authority, factory, projection = _services(account=_enabled(), provider=provider)

    result = await services.create_folder.execute(CreateFolderCommand("primary", "Projects"))

    assert result.status == "succeeded"
    assert result.reconciliation_needed is False
    authority.resolve.assert_called_once_with("primary")
    factory.open.assert_called_once_with("primary", expected_mode="legacy", purpose="incoming")
    projection.invalidate.assert_awaited_once_with(("Projects",))


@pytest.mark.asyncio
async def test_delete_folder_success_invalidates_its_mailbox() -> None:
    provider = MagicMock()
    provider.delete_folder = AsyncMock(return_value=FolderMutationOutcome("succeeded"))
    services, _, _, projection = _services(account=_enabled(), provider=provider)

    result = await services.delete_folder.execute(DeleteFolderCommand("primary", "Projects"))

    assert result.status == "succeeded"
    projection.invalidate.assert_awaited_once_with(("Projects",))


@pytest.mark.asyncio
async def test_rename_folder_invalidates_both_the_old_and_the_new_mailbox() -> None:
    """A rename moves every message between two spellings in one effect."""
    provider = MagicMock()
    provider.rename_folder = AsyncMock(return_value=FolderMutationOutcome("succeeded"))
    services, _, _, projection = _services(account=_enabled(), provider=provider)

    result = await services.rename_folder.execute(RenameFolderCommand("primary", "Projects", "Work/Projects"))

    assert result.status == "succeeded"
    projection.invalidate.assert_awaited_once_with(("Projects", "Work/Projects"))


@pytest.mark.parametrize(
    ("service_name", "command", "provider_method"),
    [
        ("create_folder", CreateFolderCommand("primary", "Projects"), "create_folder"),
        ("delete_folder", DeleteFolderCommand("primary", "Projects"), "delete_folder"),
        ("rename_folder", RenameFolderCommand("primary", "Projects", "Renamed"), "rename_folder"),
    ],
)
@pytest.mark.asyncio
async def test_folder_mutation_timeout_is_unknown_and_requires_reconciliation(
    service_name: str,
    command: object,
    provider_method: str,
) -> None:
    provider = MagicMock()
    setattr(provider, provider_method, AsyncMock(side_effect=TimeoutError()))
    services, _, _, projection = _services(account=_enabled(), provider=provider)

    result = await getattr(services, service_name).execute(command)

    assert result.status == "unknown"
    assert result.detail == "provider-timeout"
    assert result.reconciliation_needed is True
    # The mailbox may already exist, be gone, or be renamed, so the projection is stale.
    projection.invalidate.assert_awaited_once()
    getattr(provider, provider_method).assert_awaited_once()


@pytest.mark.asyncio
async def test_folder_mutation_known_rejection_does_not_invalidate_the_projection() -> None:
    provider = MagicMock()
    provider.create_folder = AsyncMock(return_value=FolderMutationOutcome("failed", "create-rejected"))
    services, _, _, projection = _services(account=_enabled(), provider=provider)

    result = await services.create_folder.execute(CreateFolderCommand("primary", "Projects"))

    assert result.status == "failed"
    assert result.reconciliation_needed is False
    projection.invalidate.assert_not_awaited()


@pytest.mark.asyncio
async def test_folder_mutation_success_survives_projection_failure_with_a_warning() -> None:
    provider = MagicMock()
    provider.rename_folder = AsyncMock(return_value=FolderMutationOutcome("succeeded"))
    projection = MagicMock()
    projection.invalidate = AsyncMock(side_effect=MutationProjectionError("unavailable"))
    services, _, _, _ = _services(account=_enabled(), provider=provider, projection=projection)

    result = await services.rename_folder.execute(RenameFolderCommand("primary", "Projects", "Renamed"))

    assert result.status == "succeeded"
    assert result.reconciliation_needed is True


def test_folder_outcome_unknown_always_requires_reconciliation() -> None:
    assert FolderMutationOutcome("unknown", "create-unknown").reconciliation_needed is True
    assert FolderMutationOutcome("failed", "create-rejected").reconciliation_needed is False


# --------------------------------------------------------------------------- #
# ClassicMutationProvider adapter
# --------------------------------------------------------------------------- #


@pytest.mark.asyncio
async def test_copy_adapter_threads_the_sender_policy_into_the_provider() -> None:
    outcome = BatchMutationOutcome((TargetMutationOutcome("1", "succeeded"),))
    handler = MagicMock()
    handler.incoming_client.copy_emails_with_outcome = AsyncMock(return_value=outcome)
    account = MutationAccountSnapshot("primary", "managed", ("*@allowed.test",), (), True, can_send=True)

    result = await ClassicMutationProvider(handler).copy(
        CopyCommand("primary", ("1", "2"), "INBOX", "Archive"), account
    )

    assert result is outcome
    handler.incoming_client.copy_emails_with_outcome.assert_awaited_once_with(
        ["1", "2"], "INBOX", "Archive", ["*@allowed.test"], True
    )


@pytest.mark.asyncio
@pytest.mark.parametrize(
    "error",
    [asyncio.CancelledError(), ValueError("invalid mutation"), PermissionError("mutation denied")],
)
async def test_copy_adapter_preserves_control_and_policy_exceptions(error: BaseException) -> None:
    handler = MagicMock()
    handler.incoming_client.copy_emails_with_outcome = AsyncMock(side_effect=error)

    with pytest.raises(type(error)) as caught:
        await ClassicMutationProvider(handler).copy(CopyCommand("primary", ("1",), "INBOX", "Archive"), _account())

    assert caught.value is error


@pytest.mark.asyncio
async def test_copy_adapter_sanitizes_unexpected_provider_failure() -> None:
    provider_detail = "provider-controlled secret detail"
    handler = MagicMock()
    handler.incoming_client.copy_emails_with_outcome = AsyncMock(side_effect=RuntimeError(provider_detail))

    with pytest.raises(MutationProviderError, match=r"^provider_failure: mutation provider request failed$") as caught:
        await ClassicMutationProvider(handler).copy(CopyCommand("primary", ("1",), "INBOX", "Archive"), _account())

    assert provider_detail not in str(caught.value)
    assert caught.value.__cause__ is None


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("adapter_method", "command", "client_method", "expected_arguments"),
    [
        (
            "create_folder",
            CreateFolderCommand("primary", "Projects"),
            "create_mailbox_with_outcome",
            ("Projects",),
        ),
        (
            "delete_folder",
            DeleteFolderCommand("primary", "Projects"),
            "delete_mailbox_with_outcome",
            ("Projects",),
        ),
        (
            "rename_folder",
            RenameFolderCommand("primary", "Projects", "Renamed"),
            "rename_mailbox_with_outcome",
            ("Projects", "Renamed"),
        ),
    ],
)
async def test_folder_adapter_forwards_names_and_drops_stale_special_use_cache(
    adapter_method: str,
    command: object,
    client_method: str,
    expected_arguments: tuple[str, ...],
) -> None:
    outcome = FolderMutationOutcome("succeeded")
    handler = MagicMock()
    setattr(handler.incoming_client, client_method, AsyncMock(return_value=outcome))

    result = await getattr(ClassicMutationProvider(handler), adapter_method)(command, _enabled())

    assert result is outcome
    getattr(handler.incoming_client, client_method).assert_awaited_once_with(*expected_arguments)
    handler._invalidate_special_folder_cache.assert_called_once_with()


@pytest.mark.asyncio
async def test_folder_adapter_drops_the_special_use_cache_even_when_the_effect_fails() -> None:
    """A failed CREATE can still have raced a concurrent change, and unknown certainly can."""
    handler = MagicMock()
    handler.incoming_client.create_mailbox_with_outcome = AsyncMock(side_effect=RuntimeError("boom"))

    with pytest.raises(MutationProviderError):
        await ClassicMutationProvider(handler).create_folder(CreateFolderCommand("primary", "Projects"), _enabled())

    handler._invalidate_special_folder_cache.assert_called_once_with()


# --------------------------------------------------------------------------- #
# EmailClient IMAP provider
# --------------------------------------------------------------------------- #


@pytest.fixture
def email_client(email_server):
    return EmailClient(email_server)


def _mock_imap(**overrides: object) -> AsyncMock:
    mock = AsyncMock()
    mock._client_task = asyncio.Future()
    mock._client_task.set_result(None)
    mock.wait_hello_from_server = AsyncMock()
    mock.login = AsyncMock(return_value=MagicMock(result="OK", lines=[]))
    mock.id = AsyncMock(return_value=MagicMock(result="OK"))
    mock.select = AsyncMock(return_value=("OK", []))
    mock.uid = AsyncMock(return_value=("OK", []))
    mock.create = AsyncMock(return_value=("OK", []))
    mock.delete = AsyncMock(return_value=("OK", []))
    mock.rename = AsyncMock(return_value=("OK", []))
    mock.expunge = AsyncMock(return_value=("OK", []))
    mock.logout = AsyncMock()
    mock.protocol = MagicMock(capabilities=("IMAP4rev1", "UIDPLUS"))
    mock.protocol.capability = AsyncMock()
    for key, value in overrides.items():
        setattr(mock, key, value)
    return mock


def _uid_targets(imap: AsyncMock, operation: str) -> list[str]:
    return [call.args[1] for call in imap.uid.call_args_list if call.args and call.args[0] == operation]


@pytest.mark.asyncio
async def test_copy_quotes_the_destination_and_leaves_the_source_intact(email_client) -> None:
    imap = _mock_imap()
    with patch.object(email_client, "imap_class", return_value=imap):
        outcome = await email_client.copy_emails_with_outcome(["11", "12"], "INBOX", "All Mail")

    assert [item.status for item in outcome.outcomes] == ["succeeded", "succeeded"]
    assert imap.uid.await_args_list[0].args == ("copy", "11", '"All Mail"')
    assert _uid_targets(imap, "copy") == ["11", "12"]
    # The whole point of copy: no \Deleted flag and no expunge of any kind.
    assert _uid_targets(imap, "store") == []
    assert _uid_targets(imap, "expunge") == []
    imap.expunge.assert_not_called()
    imap.select.assert_awaited_once_with('"INBOX"')


@pytest.mark.asyncio
async def test_copy_encodes_a_non_ascii_destination_as_modified_utf7(email_client) -> None:
    imap = _mock_imap()
    with patch.object(email_client, "imap_class", return_value=imap):
        await email_client.copy_emails_with_outcome(["11"], "INBOX", "Entwürfe")

    assert imap.uid.await_args_list[0].args == ("copy", "11", '"Entw&APw-rfe"')


@pytest.mark.asyncio
async def test_copy_rejection_is_failed_and_transport_loss_is_unknown(email_client) -> None:
    responses = {"11": ("NO", []), "12": ("BYE", []), "13": ("OK", [])}
    imap = _mock_imap(uid=AsyncMock(side_effect=lambda _op, uid, *_rest: responses[uid]))
    with patch.object(email_client, "imap_class", return_value=imap):
        outcome = await email_client.copy_emails_with_outcome(["11", "12", "13"], "INBOX", "Archive")

    assert [(item.target, item.status, item.detail) for item in outcome.outcomes] == [
        ("11", "failed", "copy-rejected"),
        ("12", "unknown", "copy-unknown"),
        ("13", "succeeded", None),
    ]
    assert outcome.reconciliation_needed is True


@pytest.mark.asyncio
async def test_copy_cancellation_stops_the_batch_and_marks_the_rest_not_attempted(email_client) -> None:
    async def uid(_operation: str, email_id: str, *_rest: str):
        if email_id == "12":
            raise asyncio.CancelledError
        return "OK", []

    imap = _mock_imap(uid=AsyncMock(side_effect=uid))
    with patch.object(email_client, "imap_class", return_value=imap):
        outcome = await email_client.copy_emails_with_outcome(["11", "12", "13"], "INBOX", "Archive")

    assert [(item.target, item.status, item.detail) for item in outcome.outcomes] == [
        ("11", "succeeded", None),
        ("12", "unknown", "copy-unknown"),
        ("13", "failed", "not-attempted"),
    ]


@pytest.mark.asyncio
async def test_copy_blocked_sender_is_indistinguishable_from_a_missing_uid(email_client) -> None:
    senders = {"1": "ok@allowed.test", "2": "evil@blocked.test"}
    imap = _mock_imap()
    with (
        patch.object(email_client, "_batch_fetch_senders", AsyncMock(return_value=senders)),
        patch.object(email_client, "imap_class", return_value=imap),
    ):
        outcome = await email_client.copy_emails_with_outcome(
            ["1", "2"], "INBOX", "Archive", allowed_senders=["*@allowed.test"]
        )

    assert [(item.target, item.status, item.detail) for item in outcome.outcomes] == [
        ("1", "succeeded", None),
        ("2", "succeeded", None),
    ]
    assert _uid_targets(imap, "copy") == ["1"]


@pytest.mark.asyncio
async def test_copy_blocked_sender_is_reported_when_configured(email_client) -> None:
    senders = {"1": "ok@allowed.test", "2": "evil@blocked.test"}
    imap = _mock_imap()
    with (
        patch.object(email_client, "_batch_fetch_senders", AsyncMock(return_value=senders)),
        patch.object(email_client, "imap_class", return_value=imap),
    ):
        outcome = await email_client.copy_emails_with_outcome(
            ["1", "2"],
            "INBOX",
            "Archive",
            allowed_senders=["*@allowed.test"],
            report_blocked_mutations=True,
        )

    assert [(item.target, item.status, item.detail) for item in outcome.outcomes] == [
        ("1", "succeeded", None),
        ("2", "failed", "sender-policy"),
    ]
    assert _uid_targets(imap, "copy") == ["1"]


@pytest.mark.asyncio
async def test_copy_with_every_uid_blocked_issues_no_provider_effect(email_client) -> None:
    senders = {"1": "evil@blocked.test", "2": "spam@blocked.test"}
    imap = _mock_imap()
    with (
        patch.object(email_client, "_batch_fetch_senders", AsyncMock(return_value=senders)),
        patch.object(email_client, "imap_class", return_value=imap),
    ):
        outcome = await email_client.copy_emails_with_outcome(
            ["1", "2"], "INBOX", "Archive", allowed_senders=["*@allowed.test"]
        )

    assert [item.status for item in outcome.outcomes] == ["succeeded", "succeeded"]
    assert _uid_targets(imap, "copy") == []


@pytest.mark.asyncio
async def test_copy_rejects_a_malformed_uid_before_connecting(email_client) -> None:
    with patch.object(email_client, "imap_class", side_effect=AssertionError("must not connect")):
        with pytest.raises(ValueError, match="UID"):
            await email_client.copy_emails_with_outcome(["not-a-uid"], "INBOX", "Archive")


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("client_method", "arguments", "imap_method", "expected_arguments"),
    [
        ("create_mailbox_with_outcome", ("Entwürfe",), "create", ('"Entw&APw-rfe"',)),
        ("delete_mailbox_with_outcome", ("All Mail",), "delete", ('"All Mail"',)),
        ("rename_mailbox_with_outcome", ("Old", 'New"Name'), "rename", ('"Old"', r'"New\"Name"')),
    ],
)
async def test_mailbox_shape_commands_quote_names_for_a_select_shape_session(
    email_client,
    client_method: str,
    arguments: tuple[str, ...],
    imap_method: str,
    expected_arguments: tuple[str, ...],
) -> None:
    imap = _mock_imap()
    with patch.object(email_client, "imap_class", return_value=imap):
        outcome = await getattr(email_client, client_method)(*arguments)

    assert outcome.status == "succeeded"
    assert outcome.detail is None
    getattr(imap, imap_method).assert_awaited_once_with(*expected_arguments)
    # Mailbox-shape commands never select a mailbox and never negotiate UTF8=ACCEPT.
    imap.select.assert_not_awaited()
    imap.logout.assert_awaited_once()


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("client_method", "imap_method", "operation"),
    [
        ("create_mailbox_with_outcome", "create", "create"),
        ("delete_mailbox_with_outcome", "delete", "delete"),
    ],
)
@pytest.mark.parametrize(
    ("response", "status", "suffix"),
    [
        (("NO", []), "failed", "rejected"),
        (("BAD", []), "failed", "rejected"),
        (("BYE", []), "unknown", "unknown"),
        (SimpleNamespace(result="???"), "unknown", "unknown"),
    ],
)
async def test_mailbox_shape_commands_classify_provider_evidence(
    email_client,
    client_method: str,
    imap_method: str,
    operation: str,
    response: object,
    status: str,
    suffix: str,
) -> None:
    imap = _mock_imap(**{imap_method: AsyncMock(return_value=response)})
    with patch.object(email_client, "imap_class", return_value=imap):
        outcome = await getattr(email_client, client_method)("Projects")

    assert outcome.status == status
    assert outcome.detail == f"{operation}-{suffix}"
    assert outcome.reconciliation_needed is (status == "unknown")


@pytest.mark.asyncio
@pytest.mark.parametrize("error", [asyncio.CancelledError(), RuntimeError("transport lost")])
async def test_mailbox_shape_command_failure_is_unknown_not_rejection(email_client, error: BaseException) -> None:
    """A lost response cannot prove the mailbox was left unchanged."""
    imap = _mock_imap(rename=AsyncMock(side_effect=error))
    with patch.object(email_client, "imap_class", return_value=imap):
        outcome = await email_client.rename_mailbox_with_outcome("Old", "New")

    assert outcome.status == "unknown"
    assert outcome.detail == "rename-unknown"
    assert outcome.reconciliation_needed is True
    imap.logout.assert_awaited_once()


# --------------------------------------------------------------------------- #
# MCP tool dispatch
# --------------------------------------------------------------------------- #


def _batch(*outcomes: tuple[str, str, str | None]) -> BatchMutationOutcome:
    return BatchMutationOutcome(tuple(TargetMutationOutcome(*outcome) for outcome in outcomes))


@pytest.mark.asyncio
async def test_copy_emails_tool_reports_success_and_defaults_to_inbox() -> None:
    command_handler = AsyncMock(return_value=_batch(("1", "succeeded", None), ("2", "succeeded", None)))
    with patch("mcp_email_server.app.copy_emails_command", command_handler):
        result = await copy_emails("test_account", ["1", "2"], "Archive")

    assert result == "Successfully copied 2 email(s) to Archive"
    command = command_handler.await_args.args[0]
    assert isinstance(command, CopyCommand)
    assert command.source_mailbox == "INBOX"
    assert command.destination_mailbox == "Archive"
    assert command.email_ids == ("1", "2")


@pytest.mark.asyncio
async def test_copy_emails_tool_reports_partial_results_in_caller_order() -> None:
    command_handler = AsyncMock(
        return_value=_batch(
            ("1", "succeeded", None), ("2", "failed", "copy-rejected"), ("3", "unknown", "copy-unknown")
        )
    )
    with patch("mcp_email_server.app.copy_emails_command", command_handler):
        result = await copy_emails("test_account", ["1", "2", "3"], "Archive", "Sent")

    assert result == (
        "Copy result [succeeded: 1; failed: 2; unknown: 3 (copy-unknown); warning: reconciliation needed]"
    )
    assert command_handler.await_args.args[0].source_mailbox == "Sent"


@pytest.mark.asyncio
async def test_create_folder_tool_reports_the_created_name() -> None:
    command_handler = AsyncMock(return_value=FolderMutationOutcome("succeeded"))
    with patch("mcp_email_server.app.create_folder_command", command_handler):
        result = await create_folder("test_account", "Projects")

    assert result == "Folder 'Projects' created"
    assert command_handler.await_args.args[0] == CreateFolderCommand("test_account", "Projects")


@pytest.mark.asyncio
async def test_delete_folder_tool_reports_the_deleted_name() -> None:
    command_handler = AsyncMock(return_value=FolderMutationOutcome("succeeded"))
    with patch("mcp_email_server.app.delete_folder_command", command_handler):
        result = await delete_folder("test_account", "Projects")

    assert result == "Folder 'Projects' deleted"
    assert command_handler.await_args.args[0] == DeleteFolderCommand("test_account", "Projects")


@pytest.mark.asyncio
async def test_rename_folder_tool_reports_both_names() -> None:
    command_handler = AsyncMock(return_value=FolderMutationOutcome("succeeded"))
    with patch("mcp_email_server.app.rename_folder_command", command_handler):
        result = await rename_folder("test_account", "Projects", "Work/Projects")

    assert result == "Folder 'Projects' renamed to 'Work/Projects'"
    assert command_handler.await_args.args[0] == RenameFolderCommand("test_account", "Projects", "Work/Projects")


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("outcome", "expected"),
    [
        (FolderMutationOutcome("failed", "create-rejected"), "Create-folder result [failed (create-rejected)]"),
        (
            FolderMutationOutcome("unknown", "provider-timeout"),
            "Create-folder result [unknown (provider-timeout); warning: reconciliation needed]",
        ),
        (
            FolderMutationOutcome("succeeded", reconciliation_needed=True),
            "Create-folder result [succeeded; warning: reconciliation needed]",
        ),
        # Unreviewed provider detail is dropped rather than surfaced to the caller.
        (FolderMutationOutcome("failed", "raw provider text"), "Create-folder result [failed]"),
    ],
)
async def test_create_folder_tool_renders_only_reviewed_detail(outcome: FolderMutationOutcome, expected: str) -> None:
    command_handler = AsyncMock(return_value=outcome)
    with patch("mcp_email_server.app.create_folder_command", command_handler):
        result = await create_folder("test_account", "Projects")

    assert result == expected


@pytest.mark.asyncio
async def test_folder_tools_surface_the_policy_denial_to_the_caller() -> None:
    command_handler = AsyncMock(side_effect=PermissionError("Folder management is disabled."))
    with patch("mcp_email_server.app.delete_folder_command", command_handler):
        with pytest.raises(PermissionError, match=FOLDER_MANAGEMENT_DENIED):
            await delete_folder("test_account", "Projects")
