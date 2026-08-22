"""Label writes: create_label, delete_label, and apply_label.

These three tools own no provider primitive. A label is an ordinary IMAP
mailbox named ``Labels/<label_name>``, so creating one is ``create_folder``,
deleting one is ``delete_folder``, and applying one is ``copy_emails`` — each
with the mailbox name derived from the label. The tests therefore concentrate on
the translation and on proving the delegation really is the same effect: the
same ``enable_folder_management`` gate, the same sender allowlist, the same
projection invalidation, and the same timeout mapping, reached through the same
service instances rather than a parallel implementation.

The layers covered are the command contracts, the application services, the
classic provider adapter down to the IMAP client, and the MCP tool dispatch.
"""

from __future__ import annotations

import asyncio
from dataclasses import replace
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from mcp_email_server.adapters.mutations import ClassicMutationProvider
from mcp_email_server.app import apply_label, create_label, delete_label
from mcp_email_server.application.labels import (
    LABEL_MAILBOX_PREFIX,
    MAXIMUM_LABEL_NAME_BYTES,
    label_name_from_mailbox,
)
from mcp_email_server.application.limits import APPLICATION_LIMITS
from mcp_email_server.application.mutations import (
    ApplyLabelCommand,
    BatchMutationOutcome,
    CopyCommand,
    CreateFolderCommand,
    CreateLabelCommand,
    DeleteFolderCommand,
    DeleteLabelCommand,
    FolderMutationOutcome,
    MutationAccountSnapshot,
    MutationProjectionError,
    MutationProviderAccess,
    MutationServices,
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
    )
    return replace(account, **changes)


def _enabled(**changes: object) -> MutationAccountSnapshot:
    return _account(mode="legacy", enable_folder_management=True, **changes)


def _services(
    *,
    account: MutationAccountSnapshot | None = None,
    opened_account: MutationAccountSnapshot | None = None,
    provider: object | None = None,
    projection: MagicMock | None = None,
) -> tuple[MutationServices, MagicMock, MagicMock, MagicMock]:
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


# --------------------------------------------------------------------------- #
# Command contracts
# --------------------------------------------------------------------------- #


@pytest.mark.parametrize(
    ("command", "mailbox"),
    [
        (CreateLabelCommand("primary", "Work"), "Labels/Work"),
        (DeleteLabelCommand("primary", "Work"), "Labels/Work"),
        (ApplyLabelCommand("primary", ("1",), "Work"), "Labels/Work"),
        # A label name is only ever prefixed; the server's hierarchy delimiter
        # never enters into it, so a `.`-delimited server sees the same name.
        (CreateLabelCommand("primary", "Работа"), "Labels/Работа"),
        (CreateLabelCommand("primary", "with space"), "Labels/with space"),
    ],
)
def test_label_commands_map_the_name_onto_its_mailbox(command: object, mailbox: str) -> None:
    command.validate()  # type: ignore[attr-defined]
    assert command.label_mailbox == mailbox  # type: ignore[attr-defined]


def test_a_label_name_may_contain_a_slash_and_round_trips_unchanged() -> None:
    """Pinned decision: `/` inside a label name is accepted, not rejected.

    ``list_labels`` already surfaces ``Labels/Work/2026`` as the label
    ``Work/2026`` because ``label_name_from_mailbox`` strips only the prefix.
    Rejecting `/` here would make a listable label impossible to create or
    delete, so the write tools accept exactly what the read tools report. On a
    `/`-delimited server this nests the mailbox; on any other server it is a flat
    name that happens to contain a slash. Either way the name round-trips.
    """
    command = CreateLabelCommand("primary", "Work/2026")
    command.validate()

    assert command.label_mailbox == "Labels/Work/2026"
    assert label_name_from_mailbox(command.label_mailbox) == "Work/2026"


@pytest.mark.parametrize(
    ("command", "message"),
    [
        (CreateLabelCommand("", "Work"), "account_name"),
        (CreateLabelCommand("primary", ""), "label_name must not be empty"),
        (CreateLabelCommand("primary", "   "), "label_name must not be empty"),
        (CreateLabelCommand("primary", "Bad\nName"), "label_name must not contain control characters"),
        (CreateLabelCommand("primary", "Bad\x00Name"), "label_name must not contain control characters"),
        (CreateLabelCommand("primary", "Labels/Work"), "must not repeat"),
        (DeleteLabelCommand("primary", ""), "label_name must not be empty"),
        (DeleteLabelCommand("primary", "Bad\x7fName"), "label_name must not contain control characters"),
        (DeleteLabelCommand("primary", "Labels/Work"), "must not repeat"),
        (ApplyLabelCommand("primary", (), "Work"), "email_ids must not be empty"),
        (ApplyLabelCommand("primary", ("1", "1"), "Work"), "must not contain duplicates"),
        (ApplyLabelCommand("primary", ("0",), "Work"), "email_ids item"),
        (ApplyLabelCommand("primary", ("1",), ""), "label_name must not be empty"),
        (ApplyLabelCommand("primary", ("1",), "Labels/Work"), "must not repeat"),
        (ApplyLabelCommand("primary", ("1",), "Work", "IN\nBOX"), "mailbox"),
    ],
)
def test_label_write_commands_reject_unsupported_contracts(command: object, message: str) -> None:
    with pytest.raises(ValueError, match=message):
        command.validate()  # type: ignore[attr-defined]


@pytest.mark.parametrize(
    "command_type",
    [CreateLabelCommand, DeleteLabelCommand],
)
def test_label_names_are_bounded_so_the_mailbox_still_fits(command_type: type) -> None:
    """The bound is the mailbox bound minus the prefix, not the mailbox bound."""
    longest = command_type("primary", "a" * MAXIMUM_LABEL_NAME_BYTES)
    longest.validate()
    assert len(longest.label_mailbox.encode("utf-8")) == APPLICATION_LIMITS.mailbox_bytes

    with pytest.raises(ValueError, match=f"label_name exceeds {MAXIMUM_LABEL_NAME_BYTES} bytes"):
        command_type("primary", "a" * (MAXIMUM_LABEL_NAME_BYTES + 1)).validate()


def test_apply_label_accepts_the_full_supported_batch() -> None:
    ApplyLabelCommand(
        "primary",
        tuple(str(index) for index in range(1, APPLICATION_LIMITS.mutation_uids + 1)),
        "Work",
    ).validate()

    with pytest.raises(ValueError, match="email_ids"):
        ApplyLabelCommand(
            "primary",
            tuple(str(index) for index in range(1, APPLICATION_LIMITS.mutation_uids + 2)),
            "Work",
        ).validate()


# --------------------------------------------------------------------------- #
# Application services: delegation
# --------------------------------------------------------------------------- #


def test_label_writes_reuse_the_folder_and_copy_services_themselves() -> None:
    """Not "an equivalent service" — the same instances the folder tools use.

    This is what makes every gate, policy, and invalidation rule shared by
    construction rather than by a second implementation kept in step by hand.
    """
    services, _, _, _ = _services()

    assert services.create_label._create_folder is services.create_folder
    assert services.delete_label._delete_folder is services.delete_folder
    assert services.apply_label._copy is services.copy


@pytest.mark.asyncio
async def test_create_label_creates_the_prefixed_mailbox_and_invalidates_it() -> None:
    provider = MagicMock()
    provider.create_folder = AsyncMock(return_value=FolderMutationOutcome("succeeded"))
    services, _, factory, projection = _services(account=_enabled(), provider=provider)

    result = await services.create_label.execute(CreateLabelCommand("primary", "Work"))

    assert result.status == "succeeded"
    assert result.reconciliation_needed is False
    command = provider.create_folder.await_args.args[0]
    assert isinstance(command, CreateFolderCommand)
    assert command.folder_name == "Labels/Work"
    factory.open.assert_called_once()
    projection.invalidate.assert_awaited_once_with(("Labels/Work",))


@pytest.mark.asyncio
async def test_delete_label_deletes_the_prefixed_mailbox_and_invalidates_it() -> None:
    provider = MagicMock()
    provider.delete_folder = AsyncMock(return_value=FolderMutationOutcome("succeeded"))
    services, _, _, projection = _services(account=_enabled(), provider=provider)

    result = await services.delete_label.execute(DeleteLabelCommand("primary", "Work"))

    assert result.status == "succeeded"
    command = provider.delete_folder.await_args.args[0]
    assert isinstance(command, DeleteFolderCommand)
    assert command.folder_name == "Labels/Work"
    projection.invalidate.assert_awaited_once_with(("Labels/Work",))


@pytest.mark.asyncio
async def test_apply_label_copies_into_the_label_mailbox_and_leaves_the_source_alone() -> None:
    provider = MagicMock()
    provider.copy = AsyncMock(
        return_value=BatchMutationOutcome((
            TargetMutationOutcome("11", "succeeded"),
            TargetMutationOutcome("12", "succeeded"),
        ))
    )
    services, _, _, projection = _services(provider=provider)

    result = await services.apply_label.execute(ApplyLabelCommand("primary", ("11", "12"), "Work", "Archive"))

    assert result.targets("succeeded") == ["11", "12"]
    command = provider.copy.await_args.args[0]
    assert isinstance(command, CopyCommand)
    assert (command.source_mailbox, command.destination_mailbox) == ("Archive", "Labels/Work")
    assert command.email_ids == ("11", "12")
    # A label copy is additive, so only the label mailbox can have gone stale.
    projection.invalidate.assert_awaited_once_with(("Labels/Work",))
    provider.move.assert_not_called()


@pytest.mark.asyncio
async def test_apply_label_defaults_its_source_to_the_inbox() -> None:
    provider = MagicMock()
    provider.copy = AsyncMock(return_value=BatchMutationOutcome((TargetMutationOutcome("11", "succeeded"),)))
    services, _, _, _ = _services(provider=provider)

    await services.apply_label.execute(ApplyLabelCommand("primary", ("11",), "Work"))

    assert provider.copy.await_args.args[0].source_mailbox == "INBOX"


# --------------------------------------------------------------------------- #
# Application services: policy gate
# --------------------------------------------------------------------------- #


@pytest.mark.parametrize(
    ("service_name", "command", "provider_method"),
    [
        ("create_label", CreateLabelCommand("primary", "Work"), "create_folder"),
        ("delete_label", DeleteLabelCommand("primary", "Work"), "delete_folder"),
    ],
)
@pytest.mark.asyncio
async def test_label_shape_writes_are_denied_before_any_provider_is_opened(
    service_name: str,
    command: object,
    provider_method: str,
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
    getattr(provider, provider_method).assert_not_called()
    projection.invalidate.assert_not_awaited()


@pytest.mark.parametrize(
    ("service_name", "command", "provider_method"),
    [
        ("create_label", CreateLabelCommand("primary", "Work"), "create_folder"),
        ("delete_label", DeleteLabelCommand("primary", "Work"), "delete_folder"),
    ],
)
@pytest.mark.asyncio
async def test_label_shape_writes_are_denied_when_the_opened_account_revokes_the_policy(
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


def test_label_shape_denial_names_both_configuration_surfaces() -> None:
    services, _, _, _ = _services(account=_account(mode="legacy", enable_folder_management=False))

    with pytest.raises(PermissionError) as caught:
        asyncio.run(services.create_label.execute(CreateLabelCommand("primary", "Work")))

    message = str(caught.value)
    assert "enable_folder_management=true" in message
    assert "MCP_EMAIL_SERVER_ENABLE_FOLDER_MANAGEMENT=true" in message


@pytest.mark.asyncio
async def test_apply_label_is_not_gated_by_folder_management() -> None:
    """Labelling a message changes a mailbox's contents, not the folder layout."""
    provider = MagicMock()
    provider.copy = AsyncMock(return_value=BatchMutationOutcome((TargetMutationOutcome("11", "succeeded"),)))
    services, _, _, _ = _services(account=_account(enable_folder_management=False), provider=provider)

    result = await services.apply_label.execute(ApplyLabelCommand("primary", ("11",), "Work"))

    assert result.targets("succeeded") == ["11"]
    provider.copy.assert_awaited_once()


@pytest.mark.parametrize(
    ("service_name", "command"),
    [
        ("create_label", CreateLabelCommand("primary", "Bad\nName")),
        ("delete_label", DeleteLabelCommand("primary", "Labels/Work")),
        ("apply_label", ApplyLabelCommand("primary", ("1",), "Bad\x00Name")),
    ],
)
@pytest.mark.asyncio
async def test_label_writes_validate_the_label_name_before_resolving_authority(
    service_name: str,
    command: object,
) -> None:
    """The error names `label_name`, not the mailbox the caller never supplied."""
    services, authority, factory, _ = _services(account=_enabled())

    with pytest.raises(ValueError, match="label_name"):
        await getattr(services, service_name).execute(command)

    authority.resolve.assert_not_called()
    factory.open.assert_not_called()


# --------------------------------------------------------------------------- #
# Application services: ambiguity and projection
# --------------------------------------------------------------------------- #


@pytest.mark.parametrize(
    ("service_name", "command", "provider_method"),
    [
        ("create_label", CreateLabelCommand("primary", "Work"), "create_folder"),
        ("delete_label", DeleteLabelCommand("primary", "Work"), "delete_folder"),
    ],
)
@pytest.mark.asyncio
async def test_label_shape_timeout_is_unknown_and_requires_reconciliation(
    service_name: str,
    command: object,
    provider_method: str,
) -> None:
    provider = MagicMock()
    setattr(provider, provider_method, AsyncMock(side_effect=TimeoutError()))
    services, _, _, projection = _services(account=_enabled(), provider=provider)

    result = await getattr(services, service_name).execute(command)

    assert (result.status, result.detail) == ("unknown", "provider-timeout")
    assert result.reconciliation_needed is True
    # A lost response may still have changed the mailbox, so it is invalidated.
    projection.invalidate.assert_awaited_once_with(("Labels/Work",))


@pytest.mark.asyncio
async def test_apply_label_timeout_maps_every_uid_to_unknown_without_replay() -> None:
    provider = MagicMock()
    provider.copy = AsyncMock(side_effect=TimeoutError())
    services, _, _, _ = _services(provider=provider)

    result = await services.apply_label.execute(ApplyLabelCommand("primary", ("11", "12"), "Work"))

    assert result.targets("unknown") == ["11", "12"]
    assert [item.detail for item in result.outcomes] == ["provider-timeout", "provider-timeout"]
    assert result.reconciliation_needed is True
    provider.copy.assert_awaited_once()


@pytest.mark.asyncio
async def test_create_label_known_rejection_does_not_invalidate_the_projection() -> None:
    provider = MagicMock()
    provider.create_folder = AsyncMock(return_value=FolderMutationOutcome("failed", "create-rejected"))
    services, _, _, projection = _services(account=_enabled(), provider=provider)

    result = await services.create_label.execute(CreateLabelCommand("primary", "Work"))

    assert (result.status, result.detail) == ("failed", "create-rejected")
    projection.invalidate.assert_not_awaited()


@pytest.mark.asyncio
async def test_label_write_success_survives_projection_failure_with_a_warning() -> None:
    provider = MagicMock()
    provider.create_folder = AsyncMock(return_value=FolderMutationOutcome("succeeded"))
    projection = MagicMock()
    projection.invalidate = AsyncMock(side_effect=MutationProjectionError("unavailable"))
    services, _, _, _ = _services(account=_enabled(), provider=provider, projection=projection)

    result = await services.create_label.execute(CreateLabelCommand("primary", "Work"))

    assert result.status == "succeeded"
    assert result.reconciliation_needed is True


# --------------------------------------------------------------------------- #
# Provider adapter and IMAP client
# --------------------------------------------------------------------------- #


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
    mock.expunge = AsyncMock(return_value=("OK", []))
    mock.logout = AsyncMock()
    mock.protocol = MagicMock(capabilities=("IMAP4rev1", "UIDPLUS"))
    mock.protocol.capability = AsyncMock()
    for key, value in overrides.items():
        setattr(mock, key, value)
    return mock


def _classic_services(
    client: EmailClient,
    account: MutationAccountSnapshot,
) -> MutationServices:
    """Compose the real services over the real adapter and IMAP client."""
    handler = MagicMock()
    handler.incoming_client = client
    authority = MagicMock()
    authority.resolve.return_value = account
    factory = MagicMock()
    factory.open.return_value = MutationProviderAccess(account, ClassicMutationProvider(handler))
    projections = MagicMock()
    projections.open = AsyncMock(return_value=MagicMock(invalidate=AsyncMock()))
    return MutationServices.compose(authority, factory, projections)


@pytest.fixture
def email_client(email_server):
    return EmailClient(email_server)


@pytest.mark.parametrize(
    ("service_name", "command", "imap_method"),
    [
        ("create_label", CreateLabelCommand("primary", "Work"), "create"),
        ("delete_label", DeleteLabelCommand("primary", "Work"), "delete"),
    ],
)
@pytest.mark.asyncio
async def test_label_shape_writes_reach_imap_with_the_quoted_prefixed_mailbox(
    email_client,
    service_name: str,
    command: object,
    imap_method: str,
) -> None:
    """The literal `Labels/` prefix reaches the server verbatim inside the name."""
    imap = _mock_imap()
    services = _classic_services(email_client, _enabled())
    with patch.object(email_client, "imap_class", return_value=imap):
        result = await getattr(services, service_name).execute(command)

    assert result.status == "succeeded"
    getattr(imap, imap_method).assert_awaited_once_with('"Labels/Work"')


@pytest.mark.asyncio
async def test_a_non_ascii_label_reaches_imap_as_modified_utf7(email_client) -> None:
    imap = _mock_imap()
    services = _classic_services(email_client, _enabled())
    with patch.object(email_client, "imap_class", return_value=imap):
        await services.create_label.execute(CreateLabelCommand("primary", "Entwürfe"))

    imap.create.assert_awaited_once_with('"Labels/Entw&APw-rfe"')


@pytest.mark.asyncio
async def test_apply_label_copies_to_the_label_mailbox_without_touching_the_source(email_client) -> None:
    imap = _mock_imap()
    services = _classic_services(email_client, _account())
    with patch.object(email_client, "imap_class", return_value=imap):
        result = await services.apply_label.execute(ApplyLabelCommand("primary", ("11", "12"), "Work"))

    assert result.targets("succeeded") == ["11", "12"]
    assert [call.args for call in imap.uid.await_args_list] == [
        ("copy", "11", '"Labels/Work"'),
        ("copy", "12", '"Labels/Work"'),
    ]
    # Applying a label never removes the original: no \Deleted, no expunge.
    imap.expunge.assert_not_called()
    imap.select.assert_awaited_once_with('"INBOX"')


@pytest.mark.asyncio
async def test_apply_label_hides_a_blocked_sender_exactly_as_copy_does(email_client) -> None:
    """The sender allowlist rides along because there is only one COPY path."""
    senders = {"1": "ok@allowed.test", "2": "evil@blocked.test"}
    imap = _mock_imap()
    services = _classic_services(email_client, _account(allowed_senders=("*@allowed.test",)))
    with (
        patch.object(email_client, "_batch_fetch_senders", AsyncMock(return_value=senders)),
        patch.object(email_client, "imap_class", return_value=imap),
    ):
        result = await services.apply_label.execute(ApplyLabelCommand("primary", ("1", "2"), "Work"))

    # A blocked message is indistinguishable from one that was never there.
    assert [(item.target, item.status, item.detail) for item in result.outcomes] == [
        ("1", "succeeded", None),
        ("2", "succeeded", None),
    ]
    assert [call.args[1] for call in imap.uid.await_args_list] == ["1"]


@pytest.mark.asyncio
async def test_apply_label_reports_a_blocked_sender_when_configured(email_client) -> None:
    senders = {"1": "ok@allowed.test", "2": "evil@blocked.test"}
    imap = _mock_imap()
    services = _classic_services(
        email_client,
        _account(allowed_senders=("*@allowed.test",), report_blocked_mutations=True),
    )
    with (
        patch.object(email_client, "_batch_fetch_senders", AsyncMock(return_value=senders)),
        patch.object(email_client, "imap_class", return_value=imap),
    ):
        result = await services.apply_label.execute(ApplyLabelCommand("primary", ("1", "2"), "Work"))

    assert [(item.target, item.status, item.detail) for item in result.outcomes] == [
        ("1", "succeeded", None),
        ("2", "failed", "sender-policy"),
    ]
    assert [call.args[1] for call in imap.uid.await_args_list] == ["1"]


# --------------------------------------------------------------------------- #
# MCP tool dispatch
# --------------------------------------------------------------------------- #


def _batch(*outcomes: tuple[str, str, str | None]) -> BatchMutationOutcome:
    return BatchMutationOutcome(tuple(TargetMutationOutcome(*outcome) for outcome in outcomes))


@pytest.mark.asyncio
async def test_create_label_tool_reports_the_label_not_the_mailbox() -> None:
    command_handler = AsyncMock(return_value=FolderMutationOutcome("succeeded"))
    with patch("mcp_email_server.app.create_label_command", command_handler):
        result = await create_label("test_account", "Work")

    assert result == "Label 'Work' created"
    assert LABEL_MAILBOX_PREFIX not in result
    assert command_handler.await_args.args[0] == CreateLabelCommand("test_account", "Work")


@pytest.mark.asyncio
async def test_delete_label_tool_reports_the_label_not_the_mailbox() -> None:
    command_handler = AsyncMock(return_value=FolderMutationOutcome("succeeded"))
    with patch("mcp_email_server.app.delete_label_command", command_handler):
        result = await delete_label("test_account", "Work")

    assert result == "Label 'Work' deleted"
    assert LABEL_MAILBOX_PREFIX not in result
    assert command_handler.await_args.args[0] == DeleteLabelCommand("test_account", "Work")


@pytest.mark.asyncio
async def test_apply_label_tool_reports_success_and_defaults_to_inbox() -> None:
    command_handler = AsyncMock(return_value=_batch(("1", "succeeded", None), ("2", "succeeded", None)))
    with patch("mcp_email_server.app.apply_label_command", command_handler):
        result = await apply_label("test_account", ["1", "2"], "Work")

    assert result == "Successfully applied label 'Work' to 2 email(s)"
    command = command_handler.await_args.args[0]
    assert isinstance(command, ApplyLabelCommand)
    assert command.source_mailbox == "INBOX"
    assert command.email_ids == ("1", "2")
    assert command.label_name == "Work"


@pytest.mark.asyncio
async def test_apply_label_tool_reports_partial_results_in_caller_order() -> None:
    command_handler = AsyncMock(
        return_value=_batch(
            ("1", "succeeded", None), ("2", "failed", "copy-rejected"), ("3", "unknown", "copy-unknown")
        )
    )
    with patch("mcp_email_server.app.apply_label_command", command_handler):
        result = await apply_label("test_account", ["1", "2", "3"], "Work", "Archive")

    assert result == (
        "Apply-label result [succeeded: 1; failed: 2; unknown: 3 (copy-unknown); warning: reconciliation needed]"
    )
    assert command_handler.await_args.args[0].source_mailbox == "Archive"


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("outcome", "expected"),
    [
        (FolderMutationOutcome("failed", "create-rejected"), "Create-label result [failed (create-rejected)]"),
        (
            FolderMutationOutcome("unknown", "provider-timeout"),
            "Create-label result [unknown (provider-timeout); warning: reconciliation needed]",
        ),
        (
            FolderMutationOutcome("succeeded", reconciliation_needed=True),
            "Create-label result [succeeded; warning: reconciliation needed]",
        ),
        # Unreviewed provider detail is dropped rather than surfaced to the caller.
        (FolderMutationOutcome("failed", "raw provider text"), "Create-label result [failed]"),
    ],
)
async def test_create_label_tool_renders_only_reviewed_detail(outcome: FolderMutationOutcome, expected: str) -> None:
    command_handler = AsyncMock(return_value=outcome)
    with patch("mcp_email_server.app.create_label_command", command_handler):
        result = await create_label("test_account", "Work")

    assert result == expected


@pytest.mark.asyncio
async def test_delete_label_tool_renders_a_reviewed_rejection() -> None:
    command_handler = AsyncMock(return_value=FolderMutationOutcome("failed", "delete-rejected"))
    with patch("mcp_email_server.app.delete_label_command", command_handler):
        result = await delete_label("test_account", "Work")

    assert result == "Delete-label result [failed (delete-rejected)]"


@pytest.mark.parametrize("tool_name", ["create_label", "delete_label"])
@pytest.mark.asyncio
async def test_label_shape_tools_surface_the_policy_denial_to_the_caller(tool_name: str) -> None:
    command_handler = AsyncMock(side_effect=PermissionError("Folder management is disabled."))
    tool = {"create_label": create_label, "delete_label": delete_label}[tool_name]
    with patch(f"mcp_email_server.app.{tool_name}_command", command_handler):
        with pytest.raises(PermissionError, match=FOLDER_MANAGEMENT_DENIED):
            await tool("test_account", "Work")
