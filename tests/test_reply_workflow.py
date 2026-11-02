"""Regression coverage for environment Sent policy and server-derived reply threading."""

from email.parser import BytesParser
from email.policy import default
from unittest.mock import AsyncMock, MagicMock, patch

import pytest
import tomli_w

from mcp_email_server.adapters.management import LocalManagementBackend
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
