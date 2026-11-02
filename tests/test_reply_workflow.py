"""Regression coverage for environment Sent policy and server-derived reply threading."""

from dataclasses import replace
from email.parser import BytesParser
from email.policy import default
from unittest.mock import AsyncMock, MagicMock, patch

import pytest
import tomli_w

from mcp_email_server import app as app_module
from mcp_email_server.adapters.management import LocalManagementBackend
from mcp_email_server.adapters.mutations import LocalMutationBackend
from mcp_email_server.application.mutations import (
    AppendMutationOutcome,
    DeliveryMutationOutcome,
    MutationAccountSnapshot,
    MutationProviderAccess,
    MutationProviderError,
    MutationServices,
    RecipientPolicyDeniedError,
    ReplyCommand,
    ReplySource,
    SendCommand,
    SendMutationOutcome,
    SentCopyMutationOutcome,
    TargetMutationOutcome,
)
from mcp_email_server.config import CONFIG_PATH, Settings
from mcp_email_server.emails.classic import EmailClient, _serialize_message_for_imap_append


def environment_account(monkeypatch, name):
    for key, value in {
        "ACCOUNT_NAME": name,
        "EMAIL_ADDRESS": "sender@example.test",
        "PASSWORD": "synthetic",
        "IMAP_HOST": "imap.example.test",
        "SMTP_HOST": "smtp.example.test",
    }.items():
        monkeypatch.setenv("MCP_EMAIL_SERVER_" + key, value)
    monkeypatch.delenv("MCP_EMAIL_SERVER_SAVE_TO_SENT", raising=False)
    monkeypatch.delenv("MCP_EMAIL_SERVER_SENT_FOLDER_NAME", raising=False)


@pytest.mark.parametrize("override,expected", [(None, False), ("false", False), ("true", True)])
def test_environment_preserves_sent_policy_in_runtime_and_import(monkeypatch, email_settings, override, expected):
    stored = email_settings.model_copy(update={"save_to_sent": False, "sent_folder_name": "Special Sent"})
    environment_account(monkeypatch, stored.account_name)
    if override is not None:
        monkeypatch.setenv("MCP_EMAIL_SERVER_SAVE_TO_SENT", override)
    CONFIG_PATH.write_text(tomli_w.dumps({"emails": [stored.model_dump(exclude_none=True)]}))
    runtime = Settings()
    assert runtime.emails[0].save_to_sent is expected
    assert runtime.emails[0].sent_folder_name == "Special Sent"
    assert runtime.emails[0].incoming.host == "imap.example.test"
    preview = LocalManagementBackend()._effective_legacy_accounts({"emails": [stored.model_dump(exclude_none=True)]})
    assert preview[0].save_to_sent is expected
    assert preview[0].sent_folder_name == "Special Sent"
    assert stored.save_to_sent is False


def test_environment_only_account_still_saves_sent(monkeypatch):
    environment_account(monkeypatch, "new")
    assert Settings().get_account("new").save_to_sent is True


def test_current_config_wins_over_legacy_with_environment_credentials(monkeypatch, tmp_path, email_settings):
    from mcp_email_server import config as config_module

    current = tmp_path / "current.toml"
    legacy = tmp_path / "legacy.toml"
    for path, enabled in ((current, False), (legacy, True)):
        account = email_settings.model_copy(update={"save_to_sent": enabled})
        path.write_text(tomli_w.dumps({"emails": [account.model_dump(exclude_none=True)]}))
    environment_account(monkeypatch, email_settings.account_name)
    monkeypatch.delenv("MCP_EMAIL_SERVER_CONFIG_PATH", raising=False)
    monkeypatch.setattr(config_module, "DEFAULT_CONFIG_PATH", str(current))
    monkeypatch.setattr(config_module, "LEGACY_CONFIG_PATH", str(legacy))
    selected = config_module._resolve_config_path()
    assert selected == current
    monkeypatch.setitem(Settings.model_config, "toml_file", selected)
    assert Settings().get_account(email_settings.account_name).save_to_sent is False
    monkeypatch.setenv("MCP_EMAIL_SERVER_CONFIG_PATH", str(legacy))
    selected = config_module._resolve_config_path()
    assert selected == legacy
    monkeypatch.setitem(Settings.model_config, "toml_file", selected)
    assert Settings().get_account(email_settings.account_name).save_to_sent is True


def smtp_session():
    smtp = AsyncMock()
    smtp.__aenter__.return_value = smtp
    smtp.__aexit__.return_value = False
    smtp.supports_extension = MagicMock(return_value=True)
    return smtp


@pytest.mark.parametrize("length", [65, 78, 79, 90, 250, 900])
@pytest.mark.parametrize("stem", ["x", "CAJ+a_b-c=Z=w.q", "é"])
async def test_long_thread_ids_stay_literal_on_smtp_and_append(email_server, length, stem):
    """Long IDs in both headers must reach every wire path as whole, unencoded tokens."""
    local_part = (stem * length).encode()[: length - 16].decode(errors="ignore")
    ids = [f"<{local_part}{index}@example.test>" for index in range(4)]
    client = EmailClient(email_server, sender="sender@example.test")
    arguments = (["recipient@example.test"], "Re: Topic", "body")
    smtp = smtp_session()
    with patch("mcp_email_server.emails.classic.aiosmtplib.SMTP", return_value=smtp):
        delivery = await client.send_email_with_outcome(*arguments, in_reply_to=ids[-1], references=" ".join(ids))
    smtp.data.assert_awaited_once()
    draft = client.compose_message(*arguments, in_reply_to=ids[-1], references=" ".join(ids), include_bcc_header=True)
    for raw in (
        smtp.data.await_args.args[0],
        _serialize_message_for_imap_append(delivery.sent_message),
        _serialize_message_for_imap_append(draft),
    ):
        headers = raw.split(b"\r\n\r\n", 1)[0]
        assert b"=?" not in headers
        assert all(message_id.encode() in headers for message_id in ids)
        assert all(len(line) <= 998 for line in headers.split(b"\r\n"))
        parsed = BytesParser(policy=default).parsebytes(raw)
        assert str(parsed["In-Reply-To"]).split() == [ids[-1]]
        assert str(parsed["References"]).split() == ids


@pytest.mark.parametrize("save_to_sent", [False, True])
@pytest.mark.parametrize("dedicated_reply", [False, True])
async def test_active_service_environment_to_wire_and_sent_copy(
    monkeypatch, email_settings, save_to_sent, dedicated_reply, tmp_path
):
    attachment = tmp_path / "synthetic.txt"
    attachment.write_text("Synthetic attachment for regression testing")
    stored = email_settings.model_copy(update={"save_to_sent": save_to_sent})
    environment_account(monkeypatch, stored.account_name)
    CONFIG_PATH.write_text(tomli_w.dumps({"emails": [stored.model_dump(exclude_none=True)]}))
    settings = Settings()
    monkeypatch.setattr("mcp_email_server.adapters.authority.get_settings", lambda: settings)
    backend = LocalMutationBackend()
    projections = MagicMock()
    projections.open = AsyncMock(return_value=MagicMock(invalidate=AsyncMock()))
    services = MutationServices.compose(backend, backend, projections)
    smtp = smtp_session()
    append = AsyncMock(return_value=SentCopyMutationOutcome("succeeded", mailbox="Sent"))
    quote = AsyncMock(return_value=None)
    with (
        patch("mcp_email_server.emails.classic.aiosmtplib.SMTP", return_value=smtp),
        patch.object(EmailClient, "append_to_sent_with_outcome", append),
        patch.object(
            EmailClient,
            "fetch_reply_source",
            AsyncMock(return_value=ReplySource("parent@example.test", None, None, "Topic")),
        ),
        patch("mcp_email_server.emails.classic.ClassicEmailHandler.fetch_quote_source", quote),
    ):
        if dedicated_reply:
            result = await services.reply.execute(
                reply(account_name=stored.account_name, attachments=(str(attachment),))
            )
        else:
            result = await services.send.execute(
                SendCommand(
                    stored.account_name,
                    ("recipient@example.test",),
                    "Re: Topic",
                    "Body",
                    in_reply_to="parent@example.test",
                    references="parent@example.test",
                    quote_reply=False,
                    attachments=(str(attachment),),
                )
            )
    smtp.data.assert_awaited_once()
    quote.assert_not_awaited()
    wire = BytesParser(policy=default).parsebytes(smtp.data.await_args.args[0])
    assert [part.get_filename() for part in wire.iter_attachments()] == ["synthetic.txt"]
    assert str(wire["Subject"]) == "Re: Topic"
    assert str(wire["In-Reply-To"]) == "<parent@example.test>"
    assert str(wire["References"]) == "<parent@example.test>"
    assert result.sent_copy.status == ("succeeded" if save_to_sent else "skipped")
    assert append.await_count == int(save_to_sent)
    if save_to_sent:
        saved = append.await_args.args[0]
        assert saved["Message-ID"] == wire["Message-ID"]
        assert saved["In-Reply-To"] == wire["In-Reply-To"]
        assert saved["References"] == wire["References"]


def workflow(source, **changes):
    account = MutationAccountSnapshot("test", "legacy", (), (), False, can_send=True)
    account = replace(account, **changes)
    provider = MagicMock()
    provider.fetch_reply_source = AsyncMock(return_value=source)
    provider.send = AsyncMock()
    provider.save_to_mailbox = AsyncMock()
    authority = MagicMock(resolve=MagicMock(return_value=account))
    factory = MagicMock(open=MagicMock(return_value=MutationProviderAccess(account, provider)))
    projections = MagicMock(open=AsyncMock(return_value=MagicMock(invalidate=AsyncMock())))
    return MutationServices.compose(authority, factory, projections), provider


def reply(**changes):
    command = ReplyCommand(
        account_name="test",
        recipients=("recipient@example.test",),
        subject="",
        body="body",
        source_mailbox="Archive",
        source_email_id="42",
        reply_mode="reply",
    )
    return replace(command, **changes)


@pytest.mark.parametrize(
    "references,parent_reply,expected",
    [
        (None, None, "<parent@example.test>"),
        (
            "<root@example.test> <mid@example.test>",
            None,
            "<root@example.test> <mid@example.test> <parent@example.test>",
        ),
        (
            "<root@example.test>\r\n <mid@example.test>",
            "<ignored@example.test>",
            "<root@example.test> <mid@example.test> <parent@example.test>",
        ),
        (None, "<root@example.test>", "<root@example.test> <parent@example.test>"),
    ],
)
async def test_reply_and_draft_construct_headers(references, parent_reply, expected):
    services, provider = workflow(ReplySource("parent@example.test", references, parent_reply, "Topic"))
    provider.save_to_mailbox.return_value = AppendMutationOutcome("succeeded", "<draft@example.test>", mailbox="Drafts")
    result = await services.reply.execute(reply(draft_mailbox="Drafts"))
    assert result.status == "succeeded"
    draft = provider.save_to_mailbox.await_args.args[0]
    assert draft.mailbox == "Drafts"
    assert draft.in_reply_to == "<parent@example.test>"
    assert draft.references == expected
    assert draft.subject == "Re: Topic"
    assert draft.recipients == ("recipient@example.test",)
    assert draft.body == "body"
    provider.send.assert_not_called()


@pytest.mark.parametrize(
    "subject,expected",
    [("Topic", "Re: Topic"), ("Re: Topic", "Re: Topic"), ("RE: Topic", "RE: Topic"), ("", "Re: ")],
)
async def test_reply_send_derives_subject_and_never_quotes(subject, expected):
    services, provider = workflow(ReplySource("<parent@example.test>", None, None, subject))
    provider.send.return_value = DeliveryMutationOutcome(
        (TargetMutationOutcome("recipient@example.test", "succeeded"),), None
    )
    await services.reply.execute(reply())
    sent = provider.send.await_args.args[0]
    assert sent.subject == expected
    assert sent.quote_reply is False
    assert sent.body == "body"
    assert sent.bcc == ()
    provider.fetch_quote_source.assert_not_called()


@pytest.mark.parametrize(
    "source",
    [
        ReplySource("", None, None, "Topic"),
        ReplySource("<bad>", None, None, "Topic"),
        ReplySource("<one@example.test> <two@example.test>", None, None, "Topic"),
        ReplySource("<one@example.test>", "malformed chain", None, "Topic"),
        ReplySource("<one@example.test>", "", None, "Topic"),
        ReplySource("<one@example.test>", None, "<two@example.test> <three@example.test>", "Topic"),
        ReplySource("<one@example.test>", "<x@example.test> " * 10000, None, "Topic"),
        ReplySource("<" + "x" * 999 + "@example.test>", None, None, "Topic"),
        ReplySource("<one@example.test>", None, None, "Topic\r\nBcc: other@example.test"),
    ],
)
@pytest.mark.parametrize("draft_mailbox", [None, "Drafts"])
async def test_reply_invalid_headers_fail_before_effect(source, draft_mailbox):
    services, provider = workflow(source)
    with pytest.raises(ValueError):
        await services.reply.execute(reply(draft_mailbox=draft_mailbox))
    provider.send.assert_not_called()
    provider.save_to_mailbox.assert_not_called()


@pytest.mark.parametrize("draft_mailbox", [None, "Drafts"])
async def test_reply_unreadable_source_fails_before_effect(draft_mailbox):
    services, provider = workflow(None)
    provider.fetch_reply_source.side_effect = ValueError("Reply source could not be resolved")
    with pytest.raises(ValueError, match="could not be resolved"):
        await services.reply.execute(reply(draft_mailbox=draft_mailbox))
    provider.send.assert_not_called()
    provider.save_to_mailbox.assert_not_called()


@pytest.mark.parametrize(
    "command,changes,error",
    [
        (None, {"allowed_recipients": ("other@example.test",)}, RecipientPolicyDeniedError),
        (None, {"can_send": False}, MutationProviderError),
        (None, {"allowed_mutations": ("append",)}, PermissionError),
        ({"draft_mailbox": "Drafts"}, {"allowed_mutations": ("send",)}, PermissionError),
    ],
)
async def test_reply_policy_runs_before_header_read(command, changes, error):
    services, provider = workflow(ReplySource("<one@example.test>", None, None, "Topic"), **changes)
    with pytest.raises(error):
        await services.reply.execute(reply(**(command or {})))
    provider.fetch_reply_source.assert_not_called()


@pytest.mark.parametrize(
    "changes",
    [
        {"reply_mode": ""},
        {"reply_mode": "all"},
        {"recipients": ("one@example.test", "two@example.test")},
        {"cc": ("another@example.test",)},
        {"subject": "Caller subject"},
        {"bcc": ("hidden@example.test",)},
        {"in_reply_to": "<x@example.test>"},
        {"references": "<x@example.test>"},
        {"source_email_id": ""},
        {"source_mailbox": ""},
    ],
)
def test_reply_command_rejects_inferred_or_derived_inputs(changes):
    with pytest.raises(ValueError):
        reply(**changes).validate()


def test_reply_all_requires_explicit_choice():
    with pytest.raises(ValueError, match="choose reply-all"):
        reply(cc=("another@example.test",)).validate()
    reply(cc=("another@example.test",), reply_mode="reply-all").validate()
    reply(recipients=("one@example.test", "two@example.test"), reply_mode="reply-all").validate()


def imap_session(*data):
    imap = AsyncMock()
    imap.login.return_value = MagicMock(result="OK")
    imap.select.return_value = ("OK", [])
    imap.uid.return_value = ("OK", list(data))
    return imap


@pytest.mark.parametrize("uid_after", [False, True])
async def test_reply_source_reads_only_raw_headers_by_uid(email_server, uid_after):
    client = EmailClient(email_server)
    raw = b"Message-ID: <original@example.test>\r\nReferences: <root@example.test>\r\nSubject: Topic\r\n\r\n"
    opening = b"1 FETCH (BODY[HEADER]<0> {100}" if uid_after else b"1 FETCH (UID 42 BODY[HEADER]<0> {100}"
    closing = b" UID 42)" if uid_after else b")"
    imap = imap_session(opening, bytearray(raw), closing)
    with patch.object(client, "_connect_imap", AsyncMock(return_value=imap)):
        result = await client.fetch_reply_source("42", "Archive")
    assert result == ReplySource("<original@example.test>", "<root@example.test>", None, "Topic")
    imap.select.assert_awaited_once_with('"Archive"')
    imap.uid.assert_awaited_once()
    assert imap.uid.await_args.args[:2] == ("fetch", "42")
    assert imap.uid.await_args.args[2].startswith("BODY.PEEK[HEADER]<0.")
    imap.logout.assert_awaited_once()


async def test_reply_source_unfolds_a_folded_chain(email_server):
    client = EmailClient(email_server)
    raw = (
        b"Message-ID:\r\n <original@example.test>\r\n"
        b"References: <root@example.test>\r\n <mid@example.test>\r\nSubject: =?utf-8?q?T=C3=B3pico?=\r\n\r\n"
    )
    imap = imap_session(b"1 FETCH (UID 42 BODY[HEADER]<0> {100}", bytearray(raw), b")")
    with patch.object(client, "_connect_imap", AsyncMock(return_value=imap)):
        result = await client.fetch_reply_source("42", "Archive")
    assert result.message_id.strip() == "<original@example.test>"
    assert result.references.split() == ["<root@example.test>", "<mid@example.test>"]
    assert result.subject == "Tópico"


@pytest.mark.parametrize(
    "raw",
    [
        b"Message-ID: <x@example.test>",
        b"Message-ID: <x@example.test>\r\nMessage-ID: <y@example.test>\r\n\r\n",
        b"Message-ID: <x@example.test>\r\nReferences: <a@example.test>\r\nReferences: <b@example.test>\r\n\r\n",
        b"Subject: one\r\nSubject: two\r\n\r\n",
        b"not a header\r\n\r\n",
        b"X-Large: " + b"x" * 100000 + b"\r\n\r\n",
    ],
)
async def test_raw_reply_source_rejects_incomplete_ambiguous_or_large_headers(email_server, raw):
    client = EmailClient(email_server)
    imap = imap_session(b"1 FETCH (UID 42 BODY[HEADER] {100}", bytearray(raw), b")")
    with patch.object(client, "_connect_imap", AsyncMock(return_value=imap)), pytest.raises(ValueError):
        await client.fetch_reply_source("42", "Archive")
    imap.logout.assert_awaited_once()


@pytest.mark.parametrize("data", [[], [b"1 FETCH (UID 7 BODY[HEADER] {2}", bytearray(b"\r\n"), b")"]])
async def test_reply_source_missing_or_other_uid_is_not_resolved(email_server, data):
    client = EmailClient(email_server)
    imap = imap_session(*data)
    with patch.object(client, "_connect_imap", AsyncMock(return_value=imap)), pytest.raises(ValueError):
        await client.fetch_reply_source("42", "Archive")


async def test_reply_source_allowlist_runs_before_header_read(email_server):
    client = EmailClient(email_server)
    imap = imap_session()
    with (
        patch.object(client, "_connect_imap", AsyncMock(return_value=imap)),
        patch.object(client, "_enforce_sender_allowlist", AsyncMock(side_effect=ValueError("not found"))),
        pytest.raises(ValueError),
    ):
        await client.fetch_reply_source("42", "Archive", ["allowed@example.test"])
    imap.uid.assert_not_awaited()
    imap.logout.assert_awaited_once()


async def test_tool_reports_like_send_email_and_passes_explicit_choice():
    handler = AsyncMock(
        return_value=SendMutationOutcome(
            (TargetMutationOutcome("recipient@example.test", "succeeded"),),
            SentCopyMutationOutcome("skipped"),
            message_id="<new@example.test>",
        )
    )
    with patch("mcp_email_server.app.reply_email_command", handler):
        result = await app_module.reply_email(
            account_name="test",
            source_mailbox="Archive",
            original_email_id="42",
            recipients=["recipient@example.test"],
            body="body",
            reply_mode="reply",
        )
    assert result == "Email reply sent successfully to recipient@example.test. Message-Id: <new@example.test>"
    command = handler.await_args.args[0]
    assert (command.source_mailbox, command.source_email_id, command.reply_mode) == ("Archive", "42", "reply")
    assert command.subject == ""
    assert command.in_reply_to is None


async def test_tool_keeps_partial_delivery_and_sent_copy_evidence():
    handler = AsyncMock(
        return_value=SendMutationOutcome(
            (TargetMutationOutcome("recipient@example.test", "succeeded"),),
            SentCopyMutationOutcome("failed", "Sent", "append"),
        )
    )
    with patch("mcp_email_server.app.reply_email_command", handler):
        result = await app_module.reply_email("test", "Archive", "42", ["recipient@example.test"], "body", "reply")
    assert result.startswith("Email reply [")
    assert "sent-copy: failed" in result


async def test_tool_reports_draft_like_save_to_mailbox_and_denied_recipients():
    handler = AsyncMock(return_value=AppendMutationOutcome("succeeded", "<d@example.test>", uid="9", mailbox="Drafts"))
    with patch("mcp_email_server.app.reply_email_command", handler):
        result = await app_module.reply_email(
            "test", "Archive", "42", ["recipient@example.test"], "body", "reply", draft_mailbox="Drafts"
        )
    assert result == "Email saved to 'Drafts' successfully. Message-Id: <d@example.test>, email_id: 9"
    denied = AsyncMock(side_effect=RecipientPolicyDeniedError("denied"))
    with patch("mcp_email_server.app.reply_email_command", denied), pytest.raises(ValueError, match="not in allowlist"):
        await app_module.reply_email("test", "Archive", "42", ["other@example.test"], "body", "reply")


async def test_reply_annotations_match_send_email():
    tools = {tool.name: tool for tool in await app_module.mcp.list_tools()}
    assert tools["reply_email"].annotations == tools["send_email"].annotations
