"""Outgoing sender-software identification is configurable, bounded, and omittable."""

import pytest
from pydantic import ValidationError

from mcp_email_server.config import (
    DEFAULT_SMTP_USER_AGENT,
    DEFAULT_SMTP_X_MAILER,
    EmailServer,
    EmailSettings,
    Settings,
)
from mcp_email_server.emails.classic import EmailClient


def _server(**overrides) -> EmailServer:
    return EmailServer(
        user_name="test",
        password="test",
        host="smtp.example.com",
        port=465,
        use_ssl=True,
        **overrides,
    )


def _compose(server: EmailServer):
    client = EmailClient(server, sender="Test User <test@example.com>")
    return client.compose_message(["dest@example.com"], "Subject", "body")


class TestIdentificationHeaderDefaults:
    def test_defaults_match_the_project_name(self):
        server = _server()

        assert server.smtp_user_agent == DEFAULT_SMTP_USER_AGENT
        assert server.smtp_x_mailer == DEFAULT_SMTP_X_MAILER
        assert DEFAULT_SMTP_USER_AGENT == "mcp-email-server"
        assert DEFAULT_SMTP_X_MAILER == "mcp-email-server"

    def test_defaults_reach_the_composed_message(self):
        message = _compose(_server())

        assert message["User-Agent"] == "mcp-email-server"
        assert message["X-Mailer"] == "mcp-email-server"


class TestIdentificationHeaderOverride:
    def test_configured_values_reach_the_composed_message(self):
        message = _compose(_server(smtp_user_agent="Custom Agent 1.0", smtp_x_mailer="Custom Mailer 2.0"))

        assert message["User-Agent"] == "Custom Agent 1.0"
        assert message["X-Mailer"] == "Custom Mailer 2.0"

    def test_each_header_is_configured_independently(self):
        message = _compose(_server(smtp_user_agent="Only Agent"))

        assert message["User-Agent"] == "Only Agent"
        assert message["X-Mailer"] == DEFAULT_SMTP_X_MAILER

    def test_surrounding_whitespace_is_trimmed(self):
        assert _server(smtp_user_agent="  Padded Agent  ").smtp_user_agent == "Padded Agent"


class TestIdentificationHeaderOmission:
    def test_empty_string_omits_the_header(self):
        message = _compose(_server(smtp_user_agent="", smtp_x_mailer=""))

        assert message["User-Agent"] is None
        assert message["X-Mailer"] is None

    def test_whitespace_only_value_omits_the_header(self):
        message = _compose(_server(smtp_user_agent="   "))

        assert message["User-Agent"] is None
        assert message["X-Mailer"] == DEFAULT_SMTP_X_MAILER

    def test_omitting_one_header_leaves_the_other(self):
        message = _compose(_server(smtp_x_mailer=""))

        assert message["User-Agent"] == DEFAULT_SMTP_USER_AGENT
        assert message["X-Mailer"] is None


class TestIdentificationHeaderValidation:
    """A configured value must not be able to smuggle a second header in."""

    @pytest.mark.parametrize("injected", ["Agent\r\nBcc: attacker@example.test", "Agent\nX-Evil: 1", "Agent\x00"])
    def test_control_characters_are_rejected(self, injected):
        with pytest.raises(ValidationError, match="control characters"):
            _server(smtp_user_agent=injected)

    def test_x_mailer_is_validated_too(self):
        with pytest.raises(ValidationError, match="control characters"):
            _server(smtp_x_mailer="Mailer\rInjected: yes")

    def test_delete_character_is_rejected(self):
        with pytest.raises(ValidationError, match="control characters"):
            _server(smtp_user_agent="Agent\x7f")

    def test_ordinary_punctuation_is_accepted(self):
        assert _server(smtp_user_agent="Agent/1.0 (+https://example.test)").smtp_user_agent == (
            "Agent/1.0 (+https://example.test)"
        )


class TestIdentificationHeaderConfiguration:
    def test_init_applies_the_values_to_both_servers(self):
        account = EmailSettings.init(
            account_name="test",
            full_name="Test",
            email_address="test@example.com",
            user_name="test",
            password="pass",
            imap_host="imap.example.com",
            smtp_host="smtp.example.com",
            smtp_user_agent="Shared Agent",
            smtp_x_mailer="",
        )

        assert account.incoming.smtp_user_agent == "Shared Agent"
        assert account.outgoing is not None
        assert account.outgoing.smtp_user_agent == "Shared Agent"
        assert account.outgoing.smtp_x_mailer == ""

    def test_environment_override(self, monkeypatch):
        monkeypatch.setenv("MCP_EMAIL_SERVER_EMAIL_ADDRESS", "test@example.com")
        monkeypatch.setenv("MCP_EMAIL_SERVER_PASSWORD", "pass")
        monkeypatch.setenv("MCP_EMAIL_SERVER_IMAP_HOST", "imap.example.com")
        monkeypatch.setenv("MCP_EMAIL_SERVER_SMTP_HOST", "smtp.example.com")
        monkeypatch.setenv("MCP_EMAIL_SERVER_SMTP_USER_AGENT", "Env Agent")
        monkeypatch.setenv("MCP_EMAIL_SERVER_SMTP_X_MAILER", "")

        account = EmailSettings.from_env()

        assert account is not None
        assert account.outgoing is not None
        assert account.outgoing.smtp_user_agent == "Env Agent"
        assert account.outgoing.smtp_x_mailer == ""

    def test_absent_environment_variables_keep_the_defaults(self, monkeypatch):
        monkeypatch.setenv("MCP_EMAIL_SERVER_EMAIL_ADDRESS", "test@example.com")
        monkeypatch.setenv("MCP_EMAIL_SERVER_PASSWORD", "pass")
        monkeypatch.setenv("MCP_EMAIL_SERVER_IMAP_HOST", "imap.example.com")
        monkeypatch.delenv("MCP_EMAIL_SERVER_SMTP_USER_AGENT", raising=False)
        monkeypatch.delenv("MCP_EMAIL_SERVER_SMTP_X_MAILER", raising=False)

        account = EmailSettings.from_env()

        assert account is not None
        assert account.incoming.smtp_user_agent == DEFAULT_SMTP_USER_AGENT
        assert account.incoming.smtp_x_mailer == DEFAULT_SMTP_X_MAILER

    def test_survives_a_toml_round_trip(self, tmp_path, monkeypatch):
        config_path = tmp_path / "config.toml"
        monkeypatch.setitem(Settings.model_config, "toml_file", str(config_path))
        account = EmailSettings.init(
            account_name="test",
            full_name="Test",
            email_address="test@example.com",
            user_name="test",
            password="pass",
            imap_host="imap.example.com",
            smtp_host="smtp.example.com",
            smtp_user_agent="Persisted Agent",
            smtp_x_mailer="",
        )
        settings = Settings()
        settings.emails.append(account)
        settings.store()

        reloaded = Settings()

        assert reloaded.emails[0].outgoing is not None
        assert reloaded.emails[0].outgoing.smtp_user_agent == "Persisted Agent"
        assert reloaded.emails[0].outgoing.smtp_x_mailer == ""
