"""Tests for folder management functionality."""

import asyncio
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from mcp_email_server.app import (
    copy_emails,
    create_folder,
    delete_folder,
    list_folders,
    move_emails,
    rename_folder,
)
from mcp_email_server.config import EmailServer, EmailSettings
from mcp_email_server.emails.classic import ClassicEmailHandler, EmailClient
from mcp_email_server.emails.models import (
    EmailMoveResponse,
    Folder,
    FolderListResponse,
    FolderOperationResponse,
)

# ============================================================================
# MCP Tool Tests - Permission Checks
# ============================================================================


class TestFolderManagementDisabled:
    """Test that folder management tools raise PermissionError when disabled."""

    @pytest.mark.asyncio
    async def test_list_folders_disabled(self):
        """Test list_folders raises PermissionError when disabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = False

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with pytest.raises(PermissionError) as exc_info:
                await list_folders(account_name="test_account")

            assert "Folder management is disabled" in str(exc_info.value)

    @pytest.mark.asyncio
    async def test_move_emails_disabled(self):
        """Test move_emails raises PermissionError when disabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = False

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with pytest.raises(PermissionError) as exc_info:
                await move_emails(
                    account_name="test_account",
                    email_ids=["123"],
                    destination_folder="Archive",
                )

            assert "Folder management is disabled" in str(exc_info.value)

    @pytest.mark.asyncio
    async def test_copy_emails_disabled(self):
        """Test copy_emails raises PermissionError when disabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = False

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with pytest.raises(PermissionError) as exc_info:
                await copy_emails(
                    account_name="test_account",
                    email_ids=["123"],
                    destination_folder="Archive",
                )

            assert "Folder management is disabled" in str(exc_info.value)

    @pytest.mark.asyncio
    async def test_create_folder_disabled(self):
        """Test create_folder raises PermissionError when disabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = False

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with pytest.raises(PermissionError) as exc_info:
                await create_folder(account_name="test_account", folder_name="NewFolder")

            assert "Folder management is disabled" in str(exc_info.value)

    @pytest.mark.asyncio
    async def test_delete_folder_disabled(self):
        """Test delete_folder raises PermissionError when disabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = False

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with pytest.raises(PermissionError) as exc_info:
                await delete_folder(account_name="test_account", folder_name="OldFolder")

            assert "Folder management is disabled" in str(exc_info.value)

    @pytest.mark.asyncio
    async def test_rename_folder_disabled(self):
        """Test rename_folder raises PermissionError when disabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = False

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with pytest.raises(PermissionError) as exc_info:
                await rename_folder(
                    account_name="test_account",
                    old_name="OldName",
                    new_name="NewName",
                )

            assert "Folder management is disabled" in str(exc_info.value)


class TestFolderManagementEnabled:
    """Test that folder management tools work when enabled."""

    @pytest.mark.asyncio
    async def test_list_folders_enabled(self):
        """Test list_folders works when enabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = True

        folder_response = FolderListResponse(
            folders=[
                Folder(name="INBOX", delimiter="/", flags=["\\HasNoChildren"]),
                Folder(name="Sent", delimiter="/", flags=["\\HasNoChildren", "\\Sent"]),
                Folder(name="Archive", delimiter="/", flags=["\\HasNoChildren"]),
            ],
            total=3,
        )

        mock_handler = AsyncMock()
        mock_handler.list_folders.return_value = folder_response

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
                result = await list_folders(account_name="test_account")

                assert result == folder_response
                assert len(result.folders) == 3
                assert result.folders[0].name == "INBOX"
                mock_handler.list_folders.assert_called_once()

    @pytest.mark.asyncio
    async def test_move_emails_enabled(self):
        """Test move_emails works when enabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = True

        move_response = EmailMoveResponse(
            success=True,
            moved_ids=["123", "456"],
            failed_ids=[],
            source_mailbox="INBOX",
            destination_folder="Archive",
        )

        mock_handler = AsyncMock()
        mock_handler.move_emails.return_value = move_response

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
                result = await move_emails(
                    account_name="test_account",
                    email_ids=["123", "456"],
                    destination_folder="Archive",
                )

                assert result == move_response
                assert result.success is True
                assert result.moved_ids == ["123", "456"]
                mock_handler.move_emails.assert_called_once_with(["123", "456"], "Archive", "INBOX")

    @pytest.mark.asyncio
    async def test_copy_emails_enabled(self):
        """Test copy_emails works when enabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = True

        copy_response = EmailMoveResponse(
            success=True,
            moved_ids=["123"],
            failed_ids=[],
            source_mailbox="INBOX",
            destination_folder="Labels/Important",
        )

        mock_handler = AsyncMock()
        mock_handler.copy_emails.return_value = copy_response

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
                result = await copy_emails(
                    account_name="test_account",
                    email_ids=["123"],
                    destination_folder="Labels/Important",
                )

                assert result == copy_response
                assert result.success is True
                mock_handler.copy_emails.assert_called_once_with(["123"], "Labels/Important", "INBOX")

    @pytest.mark.asyncio
    async def test_create_folder_enabled(self):
        """Test create_folder works when enabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = True

        create_response = FolderOperationResponse(
            success=True,
            folder_name="Projects/2024",
            message="Folder 'Projects/2024' created successfully",
        )

        mock_handler = AsyncMock()
        mock_handler.create_folder.return_value = create_response

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
                result = await create_folder(
                    account_name="test_account",
                    folder_name="Projects/2024",
                )

                assert result == create_response
                assert result.success is True
                mock_handler.create_folder.assert_called_once_with("Projects/2024")

    @pytest.mark.asyncio
    async def test_delete_folder_enabled(self):
        """Test delete_folder works when enabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = True

        delete_response = FolderOperationResponse(
            success=True,
            folder_name="OldFolder",
            message="Folder 'OldFolder' deleted successfully",
        )

        mock_handler = AsyncMock()
        mock_handler.delete_folder.return_value = delete_response

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
                result = await delete_folder(
                    account_name="test_account",
                    folder_name="OldFolder",
                )

                assert result == delete_response
                assert result.success is True
                mock_handler.delete_folder.assert_called_once_with("OldFolder")

    @pytest.mark.asyncio
    async def test_rename_folder_enabled(self):
        """Test rename_folder works when enabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = True

        rename_response = FolderOperationResponse(
            success=True,
            folder_name="NewName",
            message="Folder renamed from 'OldName' to 'NewName'",
        )

        mock_handler = AsyncMock()
        mock_handler.rename_folder.return_value = rename_response

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
                result = await rename_folder(
                    account_name="test_account",
                    old_name="OldName",
                    new_name="NewName",
                )

                assert result == rename_response
                assert result.success is True
                mock_handler.rename_folder.assert_called_once_with("OldName", "NewName")


# ============================================================================
# Handler Tests
# ============================================================================


@pytest.fixture
def email_settings():
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
def classic_handler(email_settings):
    return ClassicEmailHandler(email_settings)


class TestClassicEmailHandlerFolders:
    """Test ClassicEmailHandler folder operations."""

    @pytest.mark.asyncio
    async def test_list_folders(self, classic_handler):
        """Test list_folders handler method."""
        # EmailClient.list_folders returns list[Folder], not FolderListResponse
        mock_folders = [
            Folder(name="INBOX", delimiter="/", flags=["\\HasNoChildren"]),
            Folder(name="Sent", delimiter="/", flags=["\\Sent"]),
        ]

        mock_list = AsyncMock(return_value=mock_folders)

        with patch.object(classic_handler.incoming_client, "list_folders", mock_list):
            result = await classic_handler.list_folders()

            assert isinstance(result, FolderListResponse)
            assert len(result.folders) == 2
            assert result.total == 2
            mock_list.assert_called_once()

    @pytest.mark.asyncio
    async def test_move_emails(self, classic_handler):
        """Test move_emails handler method."""
        # EmailClient.move_emails returns (moved_ids, failed_ids) tuple
        mock_move = AsyncMock(return_value=(["123"], []))

        with patch.object(classic_handler.incoming_client, "move_emails", mock_move):
            result = await classic_handler.move_emails(
                email_ids=["123"],
                destination_folder="Archive",
                source_mailbox="INBOX",
            )

            assert isinstance(result, EmailMoveResponse)
            assert result.success is True
            assert result.moved_ids == ["123"]
            assert result.failed_ids == []
            mock_move.assert_called_once_with(["123"], "Archive", "INBOX")

    @pytest.mark.asyncio
    async def test_copy_emails(self, classic_handler):
        """Test copy_emails handler method."""
        # EmailClient.copy_emails returns (copied_ids, failed_ids) tuple
        mock_copy = AsyncMock(return_value=(["123"], []))

        with patch.object(classic_handler.incoming_client, "copy_emails", mock_copy):
            result = await classic_handler.copy_emails(
                email_ids=["123"],
                destination_folder="Backup",
                source_mailbox="INBOX",
            )

            assert isinstance(result, EmailMoveResponse)
            assert result.success is True
            assert result.moved_ids == ["123"]
            mock_copy.assert_called_once_with(["123"], "Backup", "INBOX")

    @pytest.mark.asyncio
    async def test_create_folder(self, classic_handler):
        """Test create_folder handler method."""
        # EmailClient.create_folder returns (success, message) tuple
        mock_create = AsyncMock(return_value=(True, "Folder created"))

        with patch.object(classic_handler.incoming_client, "create_folder", mock_create):
            result = await classic_handler.create_folder("NewFolder")

            assert isinstance(result, FolderOperationResponse)
            assert result.success is True
            assert result.folder_name == "NewFolder"
            mock_create.assert_called_once_with("NewFolder")

    @pytest.mark.asyncio
    async def test_delete_folder(self, classic_handler):
        """Test delete_folder handler method."""
        # EmailClient.delete_folder returns (success, message) tuple
        mock_delete = AsyncMock(return_value=(True, "Folder deleted"))

        with patch.object(classic_handler.incoming_client, "delete_folder", mock_delete):
            result = await classic_handler.delete_folder("OldFolder")

            assert isinstance(result, FolderOperationResponse)
            assert result.success is True
            assert result.folder_name == "OldFolder"
            mock_delete.assert_called_once_with("OldFolder")

    @pytest.mark.asyncio
    async def test_rename_folder(self, classic_handler):
        """Test rename_folder handler method."""
        # EmailClient.rename_folder returns (success, message) tuple
        mock_rename = AsyncMock(return_value=(True, "Folder renamed"))

        with patch.object(classic_handler.incoming_client, "rename_folder", mock_rename):
            result = await classic_handler.rename_folder("OldName", "NewName")

            assert isinstance(result, FolderOperationResponse)
            assert result.success is True
            assert result.folder_name == "NewName"
            mock_rename.assert_called_once_with("OldName", "NewName")


# ============================================================================
# EmailClient Tests
# ============================================================================


@pytest.fixture
def email_server():
    return EmailServer(
        user_name="test_user",
        password="test_password",
        host="imap.example.com",
        port=993,
        use_ssl=True,
    )


@pytest.fixture
def email_client(email_server):
    return EmailClient(email_server, sender="Test User <test@example.com>")


class TestEmailClientFolders:
    """Test EmailClient folder operations."""

    @pytest.mark.asyncio
    async def test_list_folders(self, email_client):
        """Test list_folders IMAP operation."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.list = AsyncMock(
            return_value=(
                "OK",
                [
                    b'(\\HasNoChildren) "/" "INBOX"',
                    b'(\\HasNoChildren \\Sent) "/" "Sent"',
                    b'(\\HasChildren) "/" "Folders"',
                ],
            )
        )
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            result = await email_client.list_folders()

            # EmailClient.list_folders returns list[Folder]
            assert isinstance(result, list)
            assert len(result) == 3
            assert result[0].name == "INBOX"
            assert result[1].name == "Sent"
            assert "\\Sent" in result[1].flags
            mock_imap.list.assert_called_once_with('""', "*")

    @pytest.mark.asyncio
    async def test_copy_emails(self, email_client):
        """Test copy_emails IMAP operation."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.uid = AsyncMock(return_value=("OK", [b"[COPYUID 1234 1:2 100:101]"]))
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            copied_ids, failed_ids = await email_client.copy_emails(["123", "456"], "Archive", "INBOX")

            # EmailClient.copy_emails returns (copied_ids, failed_ids) tuple
            assert copied_ids == ["123", "456"]
            assert failed_ids == []
            mock_imap.select.assert_called_once_with('"INBOX"')

    @pytest.mark.asyncio
    async def test_move_emails_with_move_command(self, email_client):
        """Test move_emails using MOVE command."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        # First call (MOVE) succeeds
        mock_imap.uid = AsyncMock(return_value=("OK", [b"OK"]))
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            moved_ids, failed_ids = await email_client.move_emails(["123"], "Archive", "INBOX")

            # EmailClient.move_emails returns (moved_ids, failed_ids) tuple
            assert moved_ids == ["123"]
            assert failed_ids == []

    @pytest.mark.asyncio
    async def test_create_folder(self, email_client):
        """Test create_folder IMAP operation."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.create = AsyncMock(return_value=("OK", [b"CREATE completed"]))
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            success, message = await email_client.create_folder("NewFolder")

            # EmailClient.create_folder returns (success, message) tuple
            assert success is True
            assert "NewFolder" in message
            mock_imap.create.assert_called_once_with('"NewFolder"')

    @pytest.mark.asyncio
    async def test_delete_folder(self, email_client):
        """Test delete_folder IMAP operation."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.delete = AsyncMock(return_value=("OK", [b"DELETE completed"]))
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            success, message = await email_client.delete_folder("OldFolder")

            # EmailClient.delete_folder returns (success, message) tuple
            assert success is True
            assert "OldFolder" in message
            mock_imap.delete.assert_called_once_with('"OldFolder"')

    @pytest.mark.asyncio
    async def test_rename_folder(self, email_client):
        """Test rename_folder IMAP operation."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.rename = AsyncMock(return_value=("OK", [b"RENAME completed"]))
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            success, _message = await email_client.rename_folder("OldName", "NewName")

            # EmailClient.rename_folder returns (success, message) tuple
            assert success is True
            mock_imap.rename.assert_called_once_with('"OldName"', '"NewName"')


class TestEmailClientFolderEdgeCases:
    """Test edge cases for EmailClient folder operations."""

    @pytest.mark.asyncio
    async def test_copy_emails_partial_failure(self, email_client):
        """Test copy_emails with partial failures."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        # Simulate failure for one email
        mock_imap.uid = AsyncMock(side_effect=[("OK", []), ("NO", [b"Message not found"])])
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            copied_ids, failed_ids = await email_client.copy_emails(["123", "456"], "Archive", "INBOX")

            # First email succeeds, second fails with NO status
            assert copied_ids == ["123"]
            assert failed_ids == ["456"]

    @pytest.mark.asyncio
    async def test_create_folder_failure(self, email_client):
        """Test create_folder when it fails."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.create = AsyncMock(return_value=("NO", [b"Folder already exists"]))
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            success, _message = await email_client.create_folder("ExistingFolder")

            # EmailClient.create_folder returns (success, message) tuple
            assert success is False

    @pytest.mark.asyncio
    async def test_list_folders_special_characters(self, email_client):
        """Test list_folders with special characters in folder names."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.list = AsyncMock(
            return_value=(
                "OK",
                [
                    b'(\\HasNoChildren) "/" "INBOX"',
                    b'(\\HasNoChildren) "/" "Folders/My Folder"',
                    b'(\\HasNoChildren) "/" "[Gmail]/Sent Mail"',
                ],
            )
        )
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            result = await email_client.list_folders()

            # EmailClient.list_folders returns list[Folder]
            assert len(result) == 3
            assert result[1].name == "Folders/My Folder"
            assert result[2].name == "[Gmail]/Sent Mail"

    @pytest.mark.asyncio
    async def test_move_emails_move_returns_non_ok(self, email_client):
        """Test move_emails fallback when MOVE returns non-OK status (covers 927->935)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.expunge = AsyncMock()
        mock_imap.logout = AsyncMock()

        # MOVE returns non-OK status (not exception), then COPY succeeds
        mock_imap.uid = AsyncMock(
            side_effect=[
                ("NO", [b"MOVE not supported"]),  # MOVE returns NO status
                ("OK", [b"[COPYUID 1234 1 100]"]),  # COPY succeeds
                ("OK", []),  # STORE \\Deleted succeeds
            ]
        )

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            moved_ids, failed_ids = await email_client.move_emails(["123"], "Archive", "INBOX")

            assert moved_ids == ["123"]
            assert failed_ids == []
            # Should have called uid 3 times: move (NO), copy, store
            assert mock_imap.uid.call_count == 3
            mock_imap.expunge.assert_called_once()

    @pytest.mark.asyncio
    async def test_move_emails_fallback_to_copy_delete(self, email_client):
        """Test move_emails falls back to COPY+DELETE when MOVE fails (covers 931-940)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.expunge = AsyncMock()
        mock_imap.logout = AsyncMock()

        # First call (MOVE) fails with exception, second call (COPY) succeeds
        mock_imap.uid = AsyncMock(
            side_effect=[
                Exception("MOVE not supported"),  # MOVE fails
                ("OK", [b"[COPYUID 1234 1 100]"]),  # COPY succeeds
                ("OK", []),  # STORE \\Deleted succeeds
            ]
        )

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            moved_ids, failed_ids = await email_client.move_emails(["123"], "Archive", "INBOX")

            assert moved_ids == ["123"]
            assert failed_ids == []
            # Should have called uid 3 times: move, copy, store
            assert mock_imap.uid.call_count == 3
            mock_imap.expunge.assert_called_once()

    @pytest.mark.asyncio
    async def test_move_emails_copy_failure_after_move_failure(self, email_client):
        """Test move_emails when both MOVE and COPY fail (covers 941-946)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.expunge = AsyncMock()
        mock_imap.logout = AsyncMock()

        # MOVE fails, then COPY also fails
        mock_imap.uid = AsyncMock(
            side_effect=[
                Exception("MOVE not supported"),  # MOVE fails
                ("NO", [b"COPY failed"]),  # COPY fails with NO status
            ]
        )

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            moved_ids, failed_ids = await email_client.move_emails(["123"], "Archive", "INBOX")

            assert moved_ids == []
            assert failed_ids == ["123"]

    @pytest.mark.asyncio
    async def test_move_emails_exception_during_operation(self, email_client):
        """Test move_emails when an exception occurs during processing (covers 944-946)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.expunge = AsyncMock()
        mock_imap.logout = AsyncMock()

        # MOVE fails, then COPY raises unexpected exception
        mock_imap.uid = AsyncMock(
            side_effect=[
                Exception("MOVE not supported"),  # MOVE fails
                Exception("Connection lost during COPY"),  # COPY fails with exception
            ]
        )

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            moved_ids, failed_ids = await email_client.move_emails(["123"], "Archive", "INBOX")

            assert moved_ids == []
            assert failed_ids == ["123"]

    @pytest.mark.asyncio
    async def test_move_emails_logout_error(self, email_client):
        """Test move_emails handles logout error gracefully (covers 955-956)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.uid = AsyncMock(return_value=("OK", []))
        mock_imap.expunge = AsyncMock()
        mock_imap.logout = AsyncMock(side_effect=Exception("Logout error"))

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            moved_ids, failed_ids = await email_client.move_emails(["123"], "Archive", "INBOX")

            # Should complete despite logout error
            assert moved_ids == ["123"]
            assert failed_ids == []

    @pytest.mark.asyncio
    async def test_create_folder_exception(self, email_client):
        """Test create_folder handles exceptions (covers 979-981)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.create = AsyncMock(side_effect=Exception("Connection lost"))
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            success, message = await email_client.create_folder("NewFolder")

            assert success is False
            assert "Error creating folder" in message

    @pytest.mark.asyncio
    async def test_create_folder_logout_error(self, email_client):
        """Test create_folder handles logout error (covers 985-986)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.create = AsyncMock(return_value=("OK", []))
        mock_imap.logout = AsyncMock(side_effect=Exception("Logout error"))

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            success, _message = await email_client.create_folder("NewFolder")

            # Should complete despite logout error
            assert success is True

    @pytest.mark.asyncio
    async def test_delete_folder_failure(self, email_client):
        """Test delete_folder with NO status (covers 1004-1005)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.delete = AsyncMock(return_value=("NO", [b"Folder not empty"]))
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            success, message = await email_client.delete_folder("NonEmptyFolder")

            assert success is False
            assert "Failed to delete folder" in message

    @pytest.mark.asyncio
    async def test_delete_folder_exception(self, email_client):
        """Test delete_folder handles exceptions (covers 1007-1009)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.delete = AsyncMock(side_effect=Exception("Connection lost"))
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            success, message = await email_client.delete_folder("SomeFolder")

            assert success is False
            assert "Error deleting folder" in message

    @pytest.mark.asyncio
    async def test_delete_folder_logout_error(self, email_client):
        """Test delete_folder handles logout error (covers 1013-1014)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.delete = AsyncMock(return_value=("OK", []))
        mock_imap.logout = AsyncMock(side_effect=Exception("Logout error"))

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            success, _message = await email_client.delete_folder("SomeFolder")

            # Should complete despite logout error
            assert success is True

    @pytest.mark.asyncio
    async def test_rename_folder_failure(self, email_client):
        """Test rename_folder with NO status (covers 1032-1033)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.rename = AsyncMock(return_value=("NO", [b"Folder does not exist"]))
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            success, message = await email_client.rename_folder("OldName", "NewName")

            assert success is False
            assert "Failed to rename folder" in message

    @pytest.mark.asyncio
    async def test_rename_folder_exception(self, email_client):
        """Test rename_folder handles exceptions (covers 1035-1037)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.rename = AsyncMock(side_effect=Exception("Connection lost"))
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            success, message = await email_client.rename_folder("OldName", "NewName")

            assert success is False
            assert "Error renaming folder" in message

    @pytest.mark.asyncio
    async def test_rename_folder_logout_error(self, email_client):
        """Test rename_folder handles logout error (covers 1041-1042)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.rename = AsyncMock(return_value=("OK", []))
        mock_imap.logout = AsyncMock(side_effect=Exception("Logout error"))

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            success, _message = await email_client.rename_folder("OldName", "NewName")

            # Should complete despite logout error
            assert success is True

    def test_parse_list_response_empty_string(self, email_client):
        """Test _parse_list_response with empty string (covers 804-805)."""
        result = email_client._parse_list_response("")
        assert result is None

    def test_parse_list_response_list_completed(self, email_client):
        """Test _parse_list_response with 'LIST completed.' (covers 804-805)."""
        result = email_client._parse_list_response("LIST completed.")
        assert result is None

    def test_parse_list_response_no_parentheses(self, email_client):
        """Test _parse_list_response with no parentheses (covers 811-812)."""
        result = email_client._parse_list_response("invalid folder response")
        assert result is None

    def test_parse_list_response_malformed_parts(self, email_client):
        """Test _parse_list_response with malformed parts (covers 827-830)."""
        # Has parentheses but not enough quoted parts
        result = email_client._parse_list_response(b'(\\HasNoChildren) "/"')
        assert result is None

    def test_parse_list_response_empty_flags(self, email_client):
        """Test _parse_list_response with empty flags."""
        # Empty flags but valid format
        result = email_client._parse_list_response(b'() "/" "test"')
        assert result is not None
        assert result.name == "test"
        assert result.flags == []

    def test_parse_list_response_folder_constructor_exception(self, email_client):
        """Test _parse_list_response handles Folder constructor exception (covers 827-828)."""

        # Mock Folder to raise an exception
        with patch("mcp_email_server.emails.classic.Folder", side_effect=ValueError("Test error")):
            result = email_client._parse_list_response(b'(\\HasNoChildren) "/" "INBOX"')
            assert result is None

    @pytest.mark.asyncio
    async def test_list_folders_logout_error(self, email_client):
        """Test list_folders handles logout error (covers 857-858)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.list = AsyncMock(
            return_value=(
                "OK",
                [b'(\\HasNoChildren) "/" "INBOX"'],
            )
        )
        mock_imap.logout = AsyncMock(side_effect=Exception("Logout error"))

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            result = await email_client.list_folders()

            # Should complete despite logout error
            assert len(result) == 1
            assert result[0].name == "INBOX"

    @pytest.mark.asyncio
    async def test_list_folders_with_invalid_items(self, email_client):
        """Test list_folders filters out invalid/unparseable items (covers 848->846)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.list = AsyncMock(
            return_value=(
                "OK",
                [
                    b'(\\HasNoChildren) "/" "INBOX"',
                    b"",  # Empty string - should be filtered out
                    b"LIST completed.",  # Invalid - should be filtered out
                    b'(\\HasNoChildren) "/" "Sent"',
                    b"invalid folder response",  # No parentheses - should be filtered out
                ],
            )
        )
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            result = await email_client.list_folders()

            # Only valid folders should be returned
            assert len(result) == 2
            assert result[0].name == "INBOX"
            assert result[1].name == "Sent"

    @pytest.mark.asyncio
    async def test_copy_emails_exception(self, email_client):
        """Test copy_emails handles exception during copy (covers 888-890)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.uid = AsyncMock(side_effect=Exception("Connection lost"))
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            copied_ids, failed_ids = await email_client.copy_emails(["123"], "Archive", "INBOX")

            assert copied_ids == []
            assert failed_ids == ["123"]

    @pytest.mark.asyncio
    async def test_copy_emails_logout_error(self, email_client):
        """Test copy_emails handles logout error (covers 895-896)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.uid = AsyncMock(return_value=("OK", []))
        mock_imap.logout = AsyncMock(side_effect=Exception("Logout error"))

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            copied_ids, failed_ids = await email_client.copy_emails(["123"], "Archive", "INBOX")

            # Should complete despite logout error
            assert copied_ids == ["123"]
            assert failed_ids == []

    @pytest.mark.asyncio
    async def test_copy_emails_non_tuple_result(self, email_client):
        """Test copy_emails when IMAP returns non-tuple result (covers ternary else branch)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        # Return a string instead of tuple
        mock_imap.uid = AsyncMock(return_value="OK")
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            copied_ids, failed_ids = await email_client.copy_emails(["123"], "Archive", "INBOX")
            assert copied_ids == ["123"]
            assert failed_ids == []

    @pytest.mark.asyncio
    async def test_move_emails_non_tuple_move_result(self, email_client):
        """Test move_emails when MOVE returns non-tuple result (covers ternary else branch)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.expunge = AsyncMock()
        # Return a string instead of tuple for MOVE
        mock_imap.uid = AsyncMock(return_value="OK")
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            moved_ids, failed_ids = await email_client.move_emails(["123"], "Archive", "INBOX")
            assert moved_ids == ["123"]
            assert failed_ids == []

    @pytest.mark.asyncio
    async def test_move_emails_non_tuple_copy_result(self, email_client):
        """Test move_emails when COPY fallback returns non-tuple result (covers ternary else branch)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.expunge = AsyncMock()
        mock_imap.logout = AsyncMock()

        # MOVE fails, then COPY returns non-tuple string
        mock_imap.uid = AsyncMock(
            side_effect=[
                Exception("MOVE not supported"),  # MOVE fails
                "OK",  # COPY returns string instead of tuple
                "OK",  # STORE returns string
            ]
        )

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            moved_ids, failed_ids = await email_client.move_emails(["123"], "Archive", "INBOX")
            assert moved_ids == ["123"]
            assert failed_ids == []

    @pytest.mark.asyncio
    async def test_create_folder_non_tuple_result(self, email_client):
        """Test create_folder when IMAP returns non-tuple result (covers ternary else branch)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        # Return a string instead of tuple
        mock_imap.create = AsyncMock(return_value="OK")
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            success, message = await email_client.create_folder("NewFolder")
            assert success is True
            assert "NewFolder" in message

    @pytest.mark.asyncio
    async def test_delete_folder_non_tuple_result(self, email_client):
        """Test delete_folder when IMAP returns non-tuple result (covers ternary else branch)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        # Return a string instead of tuple
        mock_imap.delete = AsyncMock(return_value="OK")
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            success, message = await email_client.delete_folder("OldFolder")
            assert success is True
            assert "OldFolder" in message

    @pytest.mark.asyncio
    async def test_rename_folder_non_tuple_result(self, email_client):
        """Test rename_folder when IMAP returns non-tuple result (covers ternary else branch)."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        # Return a string instead of tuple
        mock_imap.rename = AsyncMock(return_value="OK")
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            success, _message = await email_client.rename_folder("OldName", "NewName")
            assert success is True


# ============================================================================
# Config Tests for enable_folder_management
# ============================================================================


class TestFolderManagementConfig:
    """Test configuration for folder management."""

    def test_default_disabled(self):
        """Test that folder management is disabled by default."""
        from mcp_email_server.config import Settings

        with patch("mcp_email_server.config.CONFIG_PATH") as mock_path:
            mock_path.exists.return_value = False
            mock_path.with_name.return_value = mock_path

            # Create settings without any config
            with patch.object(Settings, "settings_customise_sources", return_value=()):
                settings = Settings()
                assert settings.enable_folder_management is False

    def test_env_variable_enables(self):
        """Test that environment variable can enable folder management."""
        import os

        from mcp_email_server.config import Settings

        with patch.dict(os.environ, {"MCP_EMAIL_SERVER_ENABLE_FOLDER_MANAGEMENT": "true"}):
            with patch("mcp_email_server.config.CONFIG_PATH") as mock_path:
                mock_path.exists.return_value = False
                mock_path.with_name.return_value = mock_path

                with patch.object(Settings, "settings_customise_sources", return_value=()):
                    settings = Settings()
                    assert settings.enable_folder_management is True

    def test_env_variable_disables(self):
        """Test that environment variable can explicitly disable folder management."""
        import os

        from mcp_email_server.config import Settings

        with patch.dict(os.environ, {"MCP_EMAIL_SERVER_ENABLE_FOLDER_MANAGEMENT": "false"}):
            with patch("mcp_email_server.config.CONFIG_PATH") as mock_path:
                mock_path.exists.return_value = False
                mock_path.with_name.return_value = mock_path

                with patch.object(Settings, "settings_customise_sources", return_value=()):
                    settings = Settings()
                    assert settings.enable_folder_management is False
