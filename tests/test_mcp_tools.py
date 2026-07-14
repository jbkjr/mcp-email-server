from datetime import datetime, timezone
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from mcp_email_server import app as app_module
from mcp_email_server.app import (
    add_email_account,
    archive_emails,
    delete_emails,
    download_attachment,
    forward_email,
    get_emails_content,
    list_allowed_recipients,
    list_available_accounts,
    list_emails_metadata,
    list_mailboxes,
    mark_emails,
    move_emails,
    save_to_mailbox,
    send_email,
)
from mcp_email_server.config import EmailServer, EmailSettings, ProviderSettings
from mcp_email_server.emails.models import (
    AttachmentDownloadResponse,
    EmailBodyResponse,
    EmailContentBatchResponse,
    EmailDeleteResponse,
    EmailMarkResponse,
    EmailMetadata,
    EmailMetadataPageResponse,
    EmailMoveResponse,
    EmailSendResponse,
    MailboxInfo,
)


class TestMcpTools:
    @pytest.mark.asyncio
    async def test_list_available_accounts(self):
        """Test list_available_accounts MCP tool."""
        # Create test accounts
        email_settings = EmailSettings(
            account_name="test_email",
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

        provider_settings = ProviderSettings(
            account_name="test_provider",
            provider_name="test",
            api_key="test_key",
        )

        # Mock the get_settings function
        mock_settings = MagicMock()
        mock_settings.get_accounts.return_value = [email_settings, provider_settings]

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            # Call the function
            result = await list_available_accounts()

            # Verify the result
            assert len(result) == 2
            assert result[0].account_name == "test_email"
            assert result[1].account_name == "test_provider"

            # Verify get_accounts was called correctly
            mock_settings.get_accounts.assert_called_once()

    @pytest.mark.asyncio
    async def test_add_email_account(self):
        """Test add_email_account MCP tool."""
        # Create test email settings
        email_settings = EmailSettings(
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

        # Mock the get_settings function
        mock_settings = MagicMock()

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            # Call the function
            result = await add_email_account(email_settings)

            # Verify the return value
            assert result == "Successfully added email account 'test_account'"

            # Verify add_email and store were called correctly
            mock_settings.add_email.assert_called_once_with(email_settings)
            mock_settings.store.assert_called_once()

    @pytest.mark.asyncio
    async def test_add_email_account_duplicate_name_does_not_corrupt_cache(self, tmp_path, monkeypatch):
        """A rejected duplicate-name add must not leave a mutated account list cached.

        Settings.add_email() reassigns settings.emails, and pydantic's
        validate_assignment applies the new value BEFORE the
        check_unique_account_names after-validator raises — so a naive
        `settings.add_email(email); settings.store()` sequence leaves the cached
        Settings instance corrupted even though the raise happened before store().
        """
        import mcp_email_server.config as config_module
        from mcp_email_server.config import Settings

        cfg = tmp_path / "config.toml"
        monkeypatch.setitem(Settings.model_config, "toml_file", cfg)
        config_module._settings = None

        email_settings = EmailSettings(
            account_name="dup",
            full_name="Test User",
            email_address="test@example.com",
            incoming=EmailServer(user_name="test_user", password="test_password", host="imap.example.com", port=993),
        )
        duplicate = EmailSettings(
            account_name="dup",
            full_name="Someone Else",
            email_address="other@example.com",
            incoming=EmailServer(user_name="other_user", password="other_password", host="imap.other.com", port=993),
        )

        try:
            await add_email_account(email_settings)

            with pytest.raises(Exception, match="Duplicate account name"):
                await add_email_account(duplicate)

            # The cache must have been discarded, not left holding two "dup" entries.
            reloaded = config_module.get_settings()
            assert [e.account_name for e in reloaded.emails] == ["dup"]

            # A subsequent, non-conflicting add must succeed rather than tripping
            # over a corrupted duplicate that was never actually persisted.
            new_account = EmailSettings(
                account_name="brand_new",
                full_name="New User",
                email_address="new@example.com",
                incoming=EmailServer(user_name="new_user", password="new_password", host="imap.new.com", port=993),
            )
            result = await add_email_account(new_account)
            assert result == "Successfully added email account 'brand_new'"
        finally:
            config_module._settings = None

    @pytest.mark.asyncio
    async def test_list_emails_metadata(self):
        """Test list_emails_metadata MCP tool."""
        # Create test data
        now = datetime.now(timezone.utc)
        email_metadata = EmailMetadata(
            email_id="12345",
            subject="Test Subject",
            sender="sender@example.com",
            recipients=["recipient@example.com"],
            date=now,
            attachments=[],
        )

        email_metadata_page = EmailMetadataPageResponse(
            page=1,
            page_size=10,
            before=now,
            since=None,
            subject="Test",
            emails=[email_metadata],
            total=1,
        )

        # Mock the dispatch_handler function
        mock_handler = AsyncMock()
        mock_handler.get_emails_metadata.return_value = email_metadata_page

        with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
            # Call the function
            result = await list_emails_metadata(
                account_name="test_account",
                page=1,
                page_size=10,
                before=now,
                since=None,
                subject="Test",
                from_address="sender@example.com",
                to_address=None,
            )

            # Verify the result
            assert result == email_metadata_page
            assert result.page == 1
            assert result.page_size == 10
            assert result.before == now
            assert result.subject == "Test"
            assert len(result.emails) == 1
            assert result.emails[0].subject == "Test Subject"
            assert result.emails[0].email_id == "12345"

            # Verify dispatch_handler and get_emails_metadata were called correctly
            mock_handler.get_emails_metadata.assert_called_once_with(
                page=1,
                page_size=10,
                before=now,
                since=None,
                subject="Test",
                from_address="sender@example.com",
                to_address=None,
                order="desc",
                mailbox="INBOX",
                seen=None,
                flagged=None,
                answered=None,
                body=None,
                text=None,
                has_attachment=None,
            )

    @pytest.mark.asyncio
    async def test_list_emails_metadata_with_mailbox(self):
        """Test list_emails_metadata MCP tool with custom mailbox."""
        now = datetime.now(timezone.utc)
        email_metadata = EmailMetadata(
            email_id="12345",
            subject="Sent Subject",
            sender="me@example.com",
            recipients=["recipient@example.com"],
            date=now,
            attachments=[],
        )

        email_metadata_page = EmailMetadataPageResponse(
            page=1,
            page_size=10,
            before=None,
            since=None,
            subject=None,
            emails=[email_metadata],
            total=1,
        )

        mock_handler = AsyncMock()
        mock_handler.get_emails_metadata.return_value = email_metadata_page

        with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
            result = await list_emails_metadata(
                account_name="test_account",
                mailbox="Sent",
            )

            assert result == email_metadata_page
            mock_handler.get_emails_metadata.assert_called_once_with(
                page=1,
                page_size=10,
                before=None,
                since=None,
                subject=None,
                from_address=None,
                to_address=None,
                order="desc",
                mailbox="Sent",
                seen=None,
                flagged=None,
                answered=None,
                body=None,
                text=None,
                has_attachment=None,
            )

    @pytest.mark.asyncio
    async def test_get_emails_content_single(self):
        """Test get_emails_content MCP tool with single email."""
        # Create test data
        now = datetime.now(timezone.utc)
        email_body = EmailBodyResponse(
            email_id="12345",
            subject="Test Subject",
            sender="sender@example.com",
            recipients=["recipient@example.com"],
            date=now,
            body="This is the test email body content.",
            attachments=["attachment1.pdf"],
        )

        batch_response = EmailContentBatchResponse(
            emails=[email_body],
            requested_count=1,
            retrieved_count=1,
            failed_ids=[],
        )

        # Mock the dispatch_handler function
        mock_handler = AsyncMock()
        mock_handler.get_emails_content.return_value = batch_response

        with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
            # Call the function
            result = await get_emails_content(
                account_name="test_account",
                email_ids=["12345"],
            )

            # Verify the result
            assert result == batch_response
            assert result.requested_count == 1
            assert result.retrieved_count == 1
            assert len(result.failed_ids) == 0
            assert len(result.emails) == 1
            assert result.emails[0].email_id == "12345"
            assert result.emails[0].subject == "Test Subject"

            # Verify dispatch_handler and get_emails_content were called correctly
            mock_handler.get_emails_content.assert_called_once_with(["12345"], "INBOX", mark_as_read=False, body_offset=0, max_body_length=20000)

    @pytest.mark.asyncio
    async def test_get_emails_content_batch(self):
        """Test get_emails_content MCP tool with multiple emails."""
        # Create test data
        now = datetime.now(timezone.utc)
        email1 = EmailBodyResponse(
            email_id="12345",
            subject="Test Subject 1",
            sender="sender1@example.com",
            recipients=["recipient@example.com"],
            date=now,
            body="This is the first test email body content.",
            attachments=[],
        )

        email2 = EmailBodyResponse(
            email_id="12346",
            subject="Test Subject 2",
            sender="sender2@example.com",
            recipients=["recipient@example.com"],
            date=now,
            body="This is the second test email body content.",
            attachments=["attachment1.pdf"],
        )

        batch_response = EmailContentBatchResponse(
            emails=[email1, email2],
            requested_count=3,
            retrieved_count=2,
            failed_ids=["12347"],
        )

        # Mock the dispatch_handler function
        mock_handler = AsyncMock()
        mock_handler.get_emails_content.return_value = batch_response

        with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
            # Call the function
            result = await get_emails_content(
                account_name="test_account",
                email_ids=["12345", "12346", "12347"],
            )

            # Verify the result
            assert result == batch_response
            assert result.requested_count == 3
            assert result.retrieved_count == 2
            assert len(result.failed_ids) == 1
            assert result.failed_ids[0] == "12347"
            assert len(result.emails) == 2
            assert result.emails[0].email_id == "12345"
            assert result.emails[1].email_id == "12346"

            # Verify dispatch_handler and get_emails_content were called correctly
            mock_handler.get_emails_content.assert_called_once_with(["12345", "12346", "12347"], "INBOX", mark_as_read=False, body_offset=0, max_body_length=20000)

    @pytest.mark.asyncio
    async def test_get_emails_content_with_mailbox(self):
        """Test get_emails_content MCP tool with custom mailbox."""
        now = datetime.now(timezone.utc)
        email_body = EmailBodyResponse(
            email_id="12345",
            subject="Sent Subject",
            sender="me@example.com",
            recipients=["recipient@example.com"],
            date=now,
            body="This is a sent email.",
            attachments=[],
        )

        batch_response = EmailContentBatchResponse(
            emails=[email_body],
            requested_count=1,
            retrieved_count=1,
            failed_ids=[],
        )

        mock_handler = AsyncMock()
        mock_handler.get_emails_content.return_value = batch_response

        with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
            result = await get_emails_content(
                account_name="test_account",
                email_ids=["12345"],
                mailbox="Sent",
            )

            assert result == batch_response
            mock_handler.get_emails_content.assert_called_once_with(["12345"], "Sent", mark_as_read=False, body_offset=0, max_body_length=20000)

    @pytest.mark.asyncio
    async def test_tool_visibility_hides_outbound_tools_for_read_only_accounts(self):
        """Read-only deployments hide send_email but keep IMAP-only tools like save_to_mailbox."""
        read_only_account = EmailSettings(
            account_name="read_only",
            full_name="Read Only",
            email_address="read-only@example.com",
            incoming=EmailServer(
                user_name="reader",
                password="secret",
                host="imap.example.com",
                port=993,
                use_ssl=True,
            ),
        )
        mock_settings = MagicMock()
        mock_settings.get_accounts.return_value = [read_only_account]

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            tool_names = {tool.name for tool in await app_module.mcp.list_tools()}

        assert "send_email" not in tool_names
        assert "save_to_mailbox" in tool_names
        assert "list_emails_metadata" in tool_names
        assert "get_emails_content" in tool_names

    @pytest.mark.asyncio
    async def test_tool_visibility_shows_outbound_tools_for_send_capable_accounts(self):
        """SMTP-configured deployments should expose outbound tools."""
        send_capable_account = EmailSettings(
            account_name="send_capable",
            full_name="Send Capable",
            email_address="sender@example.com",
            incoming=EmailServer(
                user_name="reader",
                password="secret",
                host="imap.example.com",
                port=993,
                use_ssl=True,
            ),
            outgoing=EmailServer(
                user_name="sender",
                password="secret",
                host="smtp.example.com",
                port=465,
                use_ssl=True,
            ),
        )
        mock_settings = MagicMock()
        mock_settings.get_accounts.return_value = [send_capable_account]

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            tool_names = {tool.name for tool in await app_module.mcp.list_tools()}

        assert "send_email" in tool_names
        assert "save_to_mailbox" in tool_names

    @pytest.mark.asyncio
    async def test_send_email(self):
        """Test send_email MCP tool returns EmailSendResponse."""
        mock_response = EmailSendResponse(
            success=True,
            recipients=["recipient@example.com", "cc@example.com"],
            subject="Test Subject",
            message="Email sent successfully to recipient@example.com",
        )
        mock_handler = AsyncMock()
        mock_handler.send_email.return_value = mock_response

        mock_settings = MagicMock()
        mock_settings.allowed_recipients = []
        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
                result = await send_email(
                    account_name="test_account",
                    recipients=["recipient@example.com"],
                    subject="Test Subject",
                    body="Test Body",
                    cc=["cc@example.com"],
                    bcc=["bcc@example.com"],
                )

                assert isinstance(result, EmailSendResponse)
                assert result.success is True
                assert "recipient@example.com" in result.recipients

                mock_handler.send_email.assert_called_once_with(
                    ["recipient@example.com"],
                    "Test Subject",
                    "Test Body",
                    ["cc@example.com"],
                    ["bcc@example.com"],
                    False,
                    None,
                    None,  # in_reply_to
                    None,  # references
                    True,  # quote_reply (default is True)
                    None,  # reply_to
                )

    @pytest.mark.asyncio
    async def test_delete_emails(self):
        """Test delete_emails MCP tool returns EmailDeleteResponse."""
        mock_response = EmailDeleteResponse(
            success=True, deleted_ids=["12345", "12346"], failed_ids=[], mailbox="INBOX"
        )
        mock_handler = AsyncMock()
        mock_handler.delete_emails.return_value = mock_response

        with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
            result = await delete_emails(
                account_name="test_account",
                email_ids=["12345", "12346"],
            )

            assert isinstance(result, EmailDeleteResponse)
            assert result.success is True
            assert result.deleted_ids == ["12345", "12346"]
            mock_handler.delete_emails.assert_called_once_with(["12345", "12346"], "INBOX")

    @pytest.mark.asyncio
    async def test_delete_emails_with_failures(self):
        """Test delete_emails MCP tool with some failures."""
        mock_response = EmailDeleteResponse(
            success=False, deleted_ids=["12345"], failed_ids=["12346", "12347"], mailbox="INBOX"
        )
        mock_handler = AsyncMock()
        mock_handler.delete_emails.return_value = mock_response

        with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
            result = await delete_emails(
                account_name="test_account",
                email_ids=["12345", "12346", "12347"],
            )

            assert isinstance(result, EmailDeleteResponse)
            assert result.success is False
            assert result.failed_ids == ["12346", "12347"]
            mock_handler.delete_emails.assert_called_once_with(["12345", "12346", "12347"], "INBOX")

    @pytest.mark.asyncio
    async def test_delete_emails_with_mailbox(self):
        """Test delete_emails MCP tool with custom mailbox."""
        mock_response = EmailDeleteResponse(
            success=True, deleted_ids=["12345"], failed_ids=[], mailbox="Trash"
        )
        mock_handler = AsyncMock()
        mock_handler.delete_emails.return_value = mock_response

        with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
            result = await delete_emails(
                account_name="test_account",
                email_ids=["12345"],
                mailbox="Trash",
            )

            assert isinstance(result, EmailDeleteResponse)
            assert result.mailbox == "Trash"
            mock_handler.delete_emails.assert_called_once_with(["12345"], "Trash")

    @pytest.mark.asyncio
    async def test_mark_emails_read(self):
        """Test mark_emails MCP tool for marking as read."""
        mock_response = EmailMarkResponse(
            success=True,
            marked_ids=["12345", "12346"],
            failed_ids=[],
            mailbox="INBOX",
            marked_as="read",
        )
        mock_handler = AsyncMock()
        mock_handler.mark_emails.return_value = mock_response

        with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
            result = await mark_emails(
                account_name="test_account",
                email_ids=["12345", "12346"],
                mark_as="read",
            )

            assert result.success is True
            assert result.marked_ids == ["12345", "12346"]
            assert result.marked_as == "read"
            mock_handler.mark_emails.assert_called_once_with(["12345", "12346"], "read", "INBOX")

    @pytest.mark.asyncio
    async def test_mark_emails_as_unread(self):
        """Test mark_emails MCP tool for marking as unread."""
        mock_response = EmailMarkResponse(
            success=True,
            marked_ids=["12345"],
            failed_ids=[],
            mailbox="INBOX",
            marked_as="unread",
        )
        mock_handler = AsyncMock()
        mock_handler.mark_emails.return_value = mock_response

        with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
            result = await mark_emails(
                account_name="test_account",
                email_ids=["12345"],
                mark_as="unread",
            )

            assert result.success is True
            assert result.marked_as == "unread"

    @pytest.mark.asyncio
    async def test_mark_emails_with_failures(self):
        """Test mark_emails MCP tool with some failures."""
        mock_response = EmailMarkResponse(
            success=False,
            marked_ids=["12345"],
            failed_ids=["12346"],
            mailbox="INBOX",
            marked_as="read",
        )
        mock_handler = AsyncMock()
        mock_handler.mark_emails.return_value = mock_response

        with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
            result = await mark_emails(
                account_name="test_account",
                email_ids=["12345", "12346"],
                mark_as="read",
            )

            assert result.success is False
            assert result.marked_ids == ["12345"]
            assert result.failed_ids == ["12346"]

    @pytest.mark.asyncio
    async def test_mark_emails_custom_mailbox(self):
        """Test mark_emails MCP tool with custom mailbox."""
        mock_response = EmailMarkResponse(
            success=True,
            marked_ids=["12345"],
            failed_ids=[],
            mailbox="[Gmail]/All Mail",
            marked_as="read",
        )
        mock_handler = AsyncMock()
        mock_handler.mark_emails.return_value = mock_response

        with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
            result = await mark_emails(
                account_name="test_account",
                email_ids=["12345"],
                mark_as="read",
                mailbox="[Gmail]/All Mail",
            )

            assert result.mailbox == "[Gmail]/All Mail"
            mock_handler.mark_emails.assert_called_once_with(["12345"], "read", "[Gmail]/All Mail")

    @pytest.mark.asyncio
    async def test_download_attachment_disabled(self):
        """Test download_attachment MCP tool when feature is disabled."""
        mock_settings = MagicMock()
        mock_settings.enable_attachment_download = False

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with pytest.raises(PermissionError) as exc_info:
                await download_attachment(
                    account_name="test_account",
                    email_id="12345",
                    attachment_name="document.pdf",
                    save_path="/var/downloads/document.pdf",
                )

            assert "Attachment download is disabled" in str(exc_info.value)

    @pytest.mark.asyncio
    async def test_download_attachment_enabled(self):
        """Test download_attachment MCP tool when feature is enabled."""
        attachment_response = AttachmentDownloadResponse(
            email_id="12345",
            attachment_name="document.pdf",
            mime_type="application/pdf",
            size=1024,
            saved_path="/var/downloads/document.pdf",
        )

        mock_settings = MagicMock()
        mock_settings.enable_attachment_download = True

        mock_handler = AsyncMock()
        mock_handler.download_attachment.return_value = attachment_response

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
                result = await download_attachment(
                    account_name="test_account",
                    email_id="12345",
                    attachment_name="document.pdf",
                    save_path="/var/downloads/document.pdf",
                )

                assert result == attachment_response
                assert result.email_id == "12345"
                assert result.attachment_name == "document.pdf"
                assert result.mime_type == "application/pdf"
                assert result.size == 1024

                mock_handler.download_attachment.assert_called_once_with(
                    "12345", "document.pdf", "/var/downloads/document.pdf", "INBOX"
                )

    @pytest.mark.asyncio
    async def test_send_email_with_reply_headers(self):
        """Test send_email MCP tool with reply headers."""
        mock_response = EmailSendResponse(
            success=True,
            recipients=["recipient@example.com"],
            subject="Re: Test",
            message="Email sent successfully to recipient@example.com",
        )
        mock_handler = AsyncMock()
        mock_handler.send_email.return_value = mock_response

        mock_settings = MagicMock()
        mock_settings.allowed_recipients = []
        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
                result = await send_email(
                    account_name="test",
                    recipients=["recipient@example.com"],
                    subject="Re: Test",
                    body="Reply body",
                    in_reply_to="<original@example.com>",
                    references="<original@example.com>",
                )

                mock_handler.send_email.assert_called_once()
                call_args = mock_handler.send_email.call_args
                assert "<original@example.com>" in str(call_args)
                assert isinstance(result, EmailSendResponse)
                assert "recipient@example.com" in result.recipients

    @pytest.mark.asyncio
    async def test_get_emails_content_includes_message_id(self):
        """Test that get_emails_content returns message_id."""
        from datetime import datetime, timezone

        mock_handler = AsyncMock()
        mock_handler.get_emails_content = AsyncMock(
            return_value=EmailContentBatchResponse(
                emails=[
                    EmailBodyResponse(
                        email_id="123",
                        message_id="<test@example.com>",
                        subject="Test",
                        sender="sender@example.com",
                        recipients=["recipient@example.com"],
                        date=datetime.now(timezone.utc),
                        body="Test body",
                        attachments=[],
                    )
                ],
                requested_count=1,
                retrieved_count=1,
                failed_ids=[],
            )
        )

        with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
            result = await get_emails_content(
                account_name="test",
                email_ids=["123"],
            )

            assert result.emails[0].message_id == "<test@example.com>"

    @pytest.mark.asyncio
    async def test_get_emails_content_mark_as_read(self):
        """Test that mark_as_read parameter is passed through to handler."""
        from datetime import datetime, timezone

        mock_handler = AsyncMock()
        mock_handler.get_emails_content = AsyncMock(
            return_value=EmailContentBatchResponse(
                emails=[
                    EmailBodyResponse(
                        email_id="123",
                        subject="Test",
                        sender="sender@example.com",
                        recipients=["recipient@example.com"],
                        date=datetime.now(timezone.utc),
                        body="Test body",
                        attachments=[],
                    )
                ],
                requested_count=1,
                retrieved_count=1,
                failed_ids=[],
            )
        )

        with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
            await get_emails_content(
                account_name="test",
                email_ids=["123"],
                mark_as_read=True,
            )

            mock_handler.get_emails_content.assert_called_once_with(["123"], "INBOX", mark_as_read=True, body_offset=0, max_body_length=20000)

    @pytest.mark.asyncio
    async def test_forward_email(self):
        """Test forward_email MCP tool returns EmailSendResponse."""
        mock_response = EmailSendResponse(
            success=True,
            recipients=["bob@example.com"],
            subject="Fwd: Original",
            message="Email forwarded successfully to bob@example.com",
        )
        mock_handler = AsyncMock()
        mock_handler.forward_email.return_value = mock_response

        with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
            result = await forward_email(
                account_name="test_account",
                email_id="12345",
                recipients=["bob@example.com"],
                mailbox="INBOX",
                body="FYI",
            )

            assert isinstance(result, EmailSendResponse)
            assert result.success is True
            assert result.recipients == ["bob@example.com"]
            mock_handler.forward_email.assert_called_once_with(
                "12345",
                "INBOX",
                ["bob@example.com"],
                "FYI",
                None,  # cc
                None,  # bcc
                False,  # html
                None,  # attachments
            )

    @pytest.mark.asyncio
    async def test_forward_email_multiple_recipients(self):
        """Test forward_email MCP tool with multiple recipients."""
        mock_response = EmailSendResponse(
            success=True,
            recipients=["alice@example.com", "bob@example.com"],
            subject="Fwd: Original",
            message="Email forwarded successfully to alice@example.com, bob@example.com",
        )
        mock_handler = AsyncMock()
        mock_handler.forward_email.return_value = mock_response

        with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
            result = await forward_email(
                account_name="test_account",
                email_id="12345",
                recipients=["alice@example.com", "bob@example.com"],
            )

            assert isinstance(result, EmailSendResponse)
            assert result.recipients == ["alice@example.com", "bob@example.com"]

    @pytest.mark.asyncio
    async def test_archive_emails(self):
        """Test archive_emails MCP tool returns EmailMoveResponse."""
        mock_response = EmailMoveResponse(
            success=True,
            moved_ids=["12345", "12346"],
            failed_ids=[],
            source_mailbox="INBOX",
            destination_folder="Archive",
        )
        mock_handler = AsyncMock()
        mock_handler.archive_emails.return_value = mock_response

        with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
            result = await archive_emails(
                account_name="test_account",
                email_ids=["12345", "12346"],
            )

            assert isinstance(result, EmailMoveResponse)
            assert result.success is True
            assert result.moved_ids == ["12345", "12346"]
            assert result.destination_folder == "Archive"
            mock_handler.archive_emails.assert_called_once_with(["12345", "12346"], "INBOX")

    @pytest.mark.asyncio
    async def test_archive_emails_custom_mailbox(self):
        """Test archive_emails MCP tool with custom mailbox."""
        mock_response = EmailMoveResponse(
            success=True,
            moved_ids=["12345"],
            failed_ids=[],
            source_mailbox="Sent",
            destination_folder="Archive",
        )
        mock_handler = AsyncMock()
        mock_handler.archive_emails.return_value = mock_response

        with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
            result = await archive_emails(
                account_name="test_account",
                email_ids=["12345"],
                mailbox="Sent",
            )

            assert result.source_mailbox == "Sent"
            mock_handler.archive_emails.assert_called_once_with(["12345"], "Sent")

    @pytest.mark.asyncio
    async def test_move_emails(self):
        """Test move_emails MCP tool."""
        mock_handler = AsyncMock()
        mock_handler.move_emails.return_value = (["12345", "12346"], [])

        with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
            result = await move_emails(
                account_name="test_account",
                email_ids=["12345", "12346"],
                destination_mailbox="Archive",
            )

            assert result == "Successfully moved 2 email(s) to Archive"
            mock_handler.move_emails.assert_called_once_with(["12345", "12346"], "INBOX", "Archive")

    @pytest.mark.asyncio
    async def test_move_emails_with_source_mailbox(self):
        """Test move_emails MCP tool with custom source mailbox."""
        mock_handler = AsyncMock()
        mock_handler.move_emails.return_value = (["12345"], [])

        with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
            result = await move_emails(
                account_name="test_account",
                email_ids=["12345"],
                source_mailbox="Trash",
                destination_mailbox="INBOX",
            )

            assert result == "Successfully moved 1 email(s) to INBOX"
            mock_handler.move_emails.assert_called_once_with(["12345"], "Trash", "INBOX")

    @pytest.mark.asyncio
    async def test_move_emails_with_failures(self):
        """Test move_emails MCP tool with some failures."""
        mock_handler = AsyncMock()
        mock_handler.move_emails.return_value = (["12345"], ["12346", "12347"])

        with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
            result = await move_emails(
                account_name="test_account",
                email_ids=["12345", "12346", "12347"],
                destination_mailbox="Archive",
            )

            assert result == "Successfully moved 1 email(s) to Archive, failed to move 2 email(s): 12346, 12347"
            mock_handler.move_emails.assert_called_once_with(["12345", "12346", "12347"], "INBOX", "Archive")

    @pytest.mark.asyncio
    async def test_list_mailboxes(self):
        """Test list_mailboxes MCP tool."""
        mock_handler = AsyncMock()
        mock_handler.list_mailboxes.return_value = [
            MailboxInfo(name="INBOX", delimiter="/", flags=["\\HasChildren"]),
            MailboxInfo(name="Sent", delimiter="/", flags=["\\Sent", "\\HasNoChildren"]),
            MailboxInfo(name="Drafts", delimiter="/", flags=["\\Drafts", "\\HasNoChildren"]),
            MailboxInfo(name="Trash", delimiter="/", flags=["\\Trash", "\\HasNoChildren"]),
            MailboxInfo(name="Archive", delimiter="/", flags=["\\HasNoChildren"]),
        ]

        with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
            result = await list_mailboxes(account_name="test_account")

            assert len(result) == 5
            assert result[0].name == "INBOX"
            assert result[0].delimiter == "/"
            assert "\\HasChildren" in result[0].flags
            assert result[1].name == "Sent"
            assert "\\Sent" in result[1].flags
            mock_handler.list_mailboxes.assert_called_once_with("*", "")

    @pytest.mark.asyncio
    async def test_list_mailboxes_with_pattern(self):
        """Test list_mailboxes MCP tool with custom pattern."""
        mock_handler = AsyncMock()
        mock_handler.list_mailboxes.return_value = [
            MailboxInfo(name="INBOX.Clients", delimiter=".", flags=["\\HasNoChildren"]),
            MailboxInfo(name="INBOX.Projects", delimiter=".", flags=["\\HasNoChildren"]),
        ]

        with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
            result = await list_mailboxes(account_name="test_account", pattern="INBOX.*")

            assert len(result) == 2
            assert result[0].delimiter == "."
            mock_handler.list_mailboxes.assert_called_once_with("INBOX.*", "")

    @pytest.mark.asyncio
    async def test_list_allowed_recipients_hidden_when_unconfigured(self):
        mock_settings = MagicMock()
        mock_settings.allowed_recipients = []
        mock_settings.get_accounts.return_value = []
        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            tool_names = {tool.name for tool in await app_module.mcp.list_tools()}
        assert "list_allowed_recipients" not in tool_names

    @pytest.mark.asyncio
    async def test_list_allowed_recipients_visible_when_configured(self):
        mock_settings = MagicMock()
        mock_settings.allowed_recipients = ["alice@example.com"]
        mock_settings.get_accounts.return_value = []
        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            tool_names = {tool.name for tool in await app_module.mcp.list_tools()}
        assert "list_allowed_recipients" in tool_names

    @pytest.mark.asyncio
    async def test_list_allowed_recipients_returns_list(self):
        mock_settings = MagicMock()
        mock_settings.allowed_recipients = ["alice@example.com", "bob@example.com"]
        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            result = await list_allowed_recipients()
        assert result == ["alice@example.com", "bob@example.com"]

    @pytest.mark.asyncio
    async def test_send_email_no_allowlist_allows_any_recipient(self):
        mock_settings = MagicMock()
        mock_settings.allowed_recipients = []
        mock_handler = AsyncMock()
        mock_handler.send_email.return_value = EmailSendResponse(
            success=True, recipients=["anyone@example.com"], subject="S", message="sent"
        )
        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
                result = await send_email(account_name="test", recipients=["anyone@example.com"], subject="S", body="B")
        assert "anyone@example.com" in result.recipients
        mock_handler.send_email.assert_called_once()

    @pytest.mark.asyncio
    async def test_send_email_blocks_unlisted_recipient(self):
        mock_settings = MagicMock()
        mock_settings.allowed_recipients = ["alice@example.com"]
        mock_handler = AsyncMock()
        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
                with pytest.raises(ValueError, match="not in allowlist"):
                    await send_email(account_name="test", recipients=["mallory@evil.com"], subject="S", body="B")
        mock_handler.send_email.assert_not_called()

    @pytest.mark.asyncio
    async def test_send_email_blocks_unlisted_bcc(self):
        mock_settings = MagicMock()
        mock_settings.allowed_recipients = ["alice@example.com"]
        mock_handler = AsyncMock()
        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
                with pytest.raises(ValueError, match=r"mallory@evil\.com"):
                    await send_email(
                        account_name="test",
                        recipients=["alice@example.com"],
                        subject="S",
                        body="B",
                        bcc=["mallory@evil.com"],
                    )
        mock_handler.send_email.assert_not_called()

    @pytest.mark.asyncio
    async def test_send_email_allows_listed_recipient_with_display_name(self):
        mock_settings = MagicMock()
        mock_settings.allowed_recipients = ["alice@example.com"]
        mock_handler = AsyncMock()
        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
                await send_email(
                    account_name="test",
                    recipients=["Alice <Alice@Example.com>"],
                    subject="S",
                    body="B",
                )
        mock_handler.send_email.assert_called_once()

    @pytest.mark.asyncio
    async def test_save_to_mailbox_blocks_unlisted_recipient(self):
        mock_settings = MagicMock()
        mock_settings.allowed_recipients = ["alice@example.com"]
        mock_handler = AsyncMock()
        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
                with pytest.raises(ValueError, match="not in allowlist"):
                    await save_to_mailbox(account_name="test", recipients=["mallory@evil.com"], subject="S", body="B")
        mock_handler.save_to_mailbox.assert_not_called()

    @pytest.mark.asyncio
    async def test_save_to_mailbox_allows_listed_recipient(self):
        mock_settings = MagicMock()
        mock_settings.allowed_recipients = ["alice@example.com"]
        mock_handler = AsyncMock()
        mock_handler.save_to_mailbox.return_value = "<mid@example.com>|uid:42"
        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
                result = await save_to_mailbox(
                    account_name="test", recipients=["alice@example.com"], subject="S", body="B"
                )
        mock_handler.save_to_mailbox.assert_called_once()
        assert "saved" in result.lower()

    @pytest.mark.asyncio
    async def test_send_email_blocks_packed_multi_address_recipient(self):
        mock_settings = MagicMock()
        mock_settings.allowed_recipients = ["alice@example.com"]
        mock_handler = AsyncMock()
        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
                with pytest.raises(ValueError, match=r"mallory@evil\.com"):
                    await send_email(
                        account_name="test",
                        recipients=["alice@example.com, mallory@evil.com"],
                        subject="S",
                        body="B",
                    )
        mock_handler.send_email.assert_not_called()

    @pytest.mark.asyncio
    async def test_save_to_mailbox_blocks_packed_multi_address_recipient(self):
        mock_settings = MagicMock()
        mock_settings.allowed_recipients = ["alice@example.com"]
        mock_handler = AsyncMock()
        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
                with pytest.raises(ValueError, match=r"mallory@evil\.com"):
                    await save_to_mailbox(
                        account_name="test",
                        recipients=["alice@example.com, mallory@evil.com"],
                        subject="S",
                        body="B",
                    )
        mock_handler.save_to_mailbox.assert_not_called()
