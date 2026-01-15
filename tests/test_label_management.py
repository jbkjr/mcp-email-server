"""Tests for ProtonMail label management functionality."""

import asyncio
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from mcp_email_server.app import (
    apply_label,
    create_label,
    delete_label,
    get_email_labels,
    list_labels,
    remove_label,
)
from mcp_email_server.config import EmailServer, EmailSettings
from mcp_email_server.emails.classic import ClassicEmailHandler, EmailClient
from mcp_email_server.emails.models import (
    EmailLabelsResponse,
    EmailMoveResponse,
    Folder,
    FolderOperationResponse,
    Label,
    LabelListResponse,
)

# ============================================================================
# MCP Tool Tests - Permission Checks
# ============================================================================


class TestLabelManagementDisabled:
    """Test that label management tools raise PermissionError when folder management is disabled."""

    @pytest.mark.asyncio
    async def test_list_labels_disabled(self):
        """Test list_labels raises PermissionError when folder management is disabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = False

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with pytest.raises(PermissionError) as exc_info:
                await list_labels(account_name="test_account")

            assert "Folder management is disabled" in str(exc_info.value)

    @pytest.mark.asyncio
    async def test_apply_label_disabled(self):
        """Test apply_label raises PermissionError when folder management is disabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = False

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with pytest.raises(PermissionError) as exc_info:
                await apply_label(
                    account_name="test_account",
                    email_ids=["123"],
                    label_name="Important",
                )

            assert "Folder management is disabled" in str(exc_info.value)

    @pytest.mark.asyncio
    async def test_remove_label_disabled(self):
        """Test remove_label raises PermissionError when folder management is disabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = False

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with pytest.raises(PermissionError) as exc_info:
                await remove_label(
                    account_name="test_account",
                    email_ids=["123"],
                    label_name="Important",
                )

            assert "Folder management is disabled" in str(exc_info.value)

    @pytest.mark.asyncio
    async def test_get_email_labels_disabled(self):
        """Test get_email_labels raises PermissionError when folder management is disabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = False

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with pytest.raises(PermissionError) as exc_info:
                await get_email_labels(
                    account_name="test_account",
                    email_id="123",
                )

            assert "Folder management is disabled" in str(exc_info.value)

    @pytest.mark.asyncio
    async def test_create_label_disabled(self):
        """Test create_label raises PermissionError when folder management is disabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = False

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with pytest.raises(PermissionError) as exc_info:
                await create_label(
                    account_name="test_account",
                    label_name="NewLabel",
                )

            assert "Folder management is disabled" in str(exc_info.value)

    @pytest.mark.asyncio
    async def test_delete_label_disabled(self):
        """Test delete_label raises PermissionError when folder management is disabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = False

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with pytest.raises(PermissionError) as exc_info:
                await delete_label(
                    account_name="test_account",
                    label_name="OldLabel",
                )

            assert "Folder management is disabled" in str(exc_info.value)


class TestLabelManagementEnabled:
    """Test that label management tools work when folder management is enabled."""

    @pytest.mark.asyncio
    async def test_list_labels_enabled(self):
        """Test list_labels works when enabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = True

        label_response = LabelListResponse(
            labels=[
                Label(name="Important", full_path="Labels/Important", delimiter="/", flags=[]),
                Label(name="Work", full_path="Labels/Work", delimiter="/", flags=[]),
            ],
            total=2,
        )

        mock_handler = AsyncMock()
        mock_handler.list_labels.return_value = label_response

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
                result = await list_labels(account_name="test_account")

                assert result == label_response
                assert len(result.labels) == 2
                assert result.labels[0].name == "Important"
                mock_handler.list_labels.assert_called_once()

    @pytest.mark.asyncio
    async def test_apply_label_enabled(self):
        """Test apply_label works when enabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = True

        apply_response = EmailMoveResponse(
            success=True,
            moved_ids=["123"],
            failed_ids=[],
            source_mailbox="INBOX",
            destination_folder="Labels/Important",
        )

        mock_handler = AsyncMock()
        mock_handler.apply_label.return_value = apply_response

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
                result = await apply_label(
                    account_name="test_account",
                    email_ids=["123"],
                    label_name="Important",
                )

                assert result == apply_response
                assert result.success is True
                mock_handler.apply_label.assert_called_once_with(["123"], "Important", "INBOX")

    @pytest.mark.asyncio
    async def test_remove_label_enabled(self):
        """Test remove_label works when enabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = True

        remove_response = EmailMoveResponse(
            success=True,
            moved_ids=["123"],
            failed_ids=[],
            source_mailbox="Labels/Important",
            destination_folder="",
        )

        mock_handler = AsyncMock()
        mock_handler.remove_label.return_value = remove_response

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
                result = await remove_label(
                    account_name="test_account",
                    email_ids=["123"],
                    label_name="Important",
                )

                assert result == remove_response
                assert result.success is True
                mock_handler.remove_label.assert_called_once_with(["123"], "Important")

    @pytest.mark.asyncio
    async def test_get_email_labels_enabled(self):
        """Test get_email_labels works when enabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = True

        labels_response = EmailLabelsResponse(
            email_id="123",
            labels=["Important", "Work"],
        )

        mock_handler = AsyncMock()
        mock_handler.get_email_labels.return_value = labels_response

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
                result = await get_email_labels(
                    account_name="test_account",
                    email_id="123",
                )

                assert result == labels_response
                assert result.labels == ["Important", "Work"]
                mock_handler.get_email_labels.assert_called_once_with("123", "INBOX")

    @pytest.mark.asyncio
    async def test_create_label_enabled(self):
        """Test create_label works when enabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = True

        create_response = FolderOperationResponse(
            success=True,
            folder_name="NewLabel",
            message="Label 'NewLabel' created successfully",
        )

        mock_handler = AsyncMock()
        mock_handler.create_label.return_value = create_response

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
                result = await create_label(
                    account_name="test_account",
                    label_name="NewLabel",
                )

                assert result == create_response
                assert result.success is True
                mock_handler.create_label.assert_called_once_with("NewLabel")

    @pytest.mark.asyncio
    async def test_delete_label_enabled(self):
        """Test delete_label works when enabled."""
        mock_settings = MagicMock()
        mock_settings.enable_folder_management = True

        delete_response = FolderOperationResponse(
            success=True,
            folder_name="OldLabel",
            message="Label 'OldLabel' deleted successfully",
        )

        mock_handler = AsyncMock()
        mock_handler.delete_label.return_value = delete_response

        with patch("mcp_email_server.app.get_settings", return_value=mock_settings):
            with patch("mcp_email_server.app.dispatch_handler", return_value=mock_handler):
                result = await delete_label(
                    account_name="test_account",
                    label_name="OldLabel",
                )

                assert result == delete_response
                assert result.success is True
                mock_handler.delete_label.assert_called_once_with("OldLabel")


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


class TestClassicEmailHandlerLabels:
    """Test ClassicEmailHandler label operations."""

    @pytest.mark.asyncio
    async def test_list_labels(self, classic_handler):
        """Test list_labels handler method."""
        mock_labels = [
            Label(name="Important", full_path="Labels/Important", delimiter="/", flags=[]),
            Label(name="Work", full_path="Labels/Work", delimiter="/", flags=[]),
        ]

        mock_list = AsyncMock(return_value=mock_labels)

        with patch.object(classic_handler.incoming_client, "list_labels", mock_list):
            result = await classic_handler.list_labels()

            assert isinstance(result, LabelListResponse)
            assert len(result.labels) == 2
            assert result.total == 2
            mock_list.assert_called_once()

    @pytest.mark.asyncio
    async def test_apply_label(self, classic_handler):
        """Test apply_label handler method."""
        mock_copy = AsyncMock(return_value=(["123"], []))

        with patch.object(classic_handler.incoming_client, "copy_emails", mock_copy):
            result = await classic_handler.apply_label(
                email_ids=["123"],
                label_name="Important",
                source_mailbox="INBOX",
            )

            assert isinstance(result, EmailMoveResponse)
            assert result.success is True
            assert result.moved_ids == ["123"]
            assert result.destination_folder == "Labels/Important"
            mock_copy.assert_called_once_with(["123"], "Labels/Important", "INBOX")

    @pytest.mark.asyncio
    async def test_remove_label(self, classic_handler):
        """Test remove_label handler method."""
        mock_get_message_id = AsyncMock(return_value="<msg123@example.com>")
        mock_search = AsyncMock(return_value="456")
        mock_delete = AsyncMock(return_value=(["456"], []))

        with patch.object(classic_handler.incoming_client, "get_email_message_id", mock_get_message_id):
            with patch.object(classic_handler.incoming_client, "search_by_message_id", mock_search):
                with patch.object(classic_handler.incoming_client, "delete_from_folder", mock_delete):
                    result = await classic_handler.remove_label(
                        email_ids=["123"],
                        label_name="Important",
                    )

                    assert isinstance(result, EmailMoveResponse)
                    assert result.success is True
                    assert result.moved_ids == ["123"]
                    mock_get_message_id.assert_called_once_with("123", "INBOX")
                    mock_search.assert_called_once_with("<msg123@example.com>", "Labels/Important")
                    mock_delete.assert_called_once_with(["456"], "Labels/Important")

    @pytest.mark.asyncio
    async def test_get_email_labels(self, classic_handler):
        """Test get_email_labels handler method."""
        mock_labels = [
            Label(name="Important", full_path="Labels/Important", delimiter="/", flags=[]),
            Label(name="Work", full_path="Labels/Work", delimiter="/", flags=[]),
        ]
        mock_get_message_id = AsyncMock(return_value="<msg123@example.com>")
        mock_list_labels = AsyncMock(return_value=mock_labels)
        # Email found in Important but not Work
        mock_search = AsyncMock(side_effect=["789", None])

        with patch.object(classic_handler.incoming_client, "get_email_message_id", mock_get_message_id):
            with patch.object(classic_handler.incoming_client, "list_labels", mock_list_labels):
                with patch.object(classic_handler.incoming_client, "search_by_message_id", mock_search):
                    result = await classic_handler.get_email_labels(
                        email_id="123",
                        source_mailbox="INBOX",
                    )

                    assert isinstance(result, EmailLabelsResponse)
                    assert result.email_id == "123"
                    assert result.labels == ["Important"]

    @pytest.mark.asyncio
    async def test_create_label(self, classic_handler):
        """Test create_label handler method."""
        mock_create = AsyncMock(return_value=(True, "Folder 'Labels/NewLabel' created successfully"))

        with patch.object(classic_handler.incoming_client, "create_folder", mock_create):
            result = await classic_handler.create_label("NewLabel")

            assert isinstance(result, FolderOperationResponse)
            assert result.success is True
            assert result.folder_name == "NewLabel"
            mock_create.assert_called_once_with("Labels/NewLabel")

    @pytest.mark.asyncio
    async def test_delete_label(self, classic_handler):
        """Test delete_label handler method."""
        mock_delete = AsyncMock(return_value=(True, "Folder 'Labels/OldLabel' deleted successfully"))

        with patch.object(classic_handler.incoming_client, "delete_folder", mock_delete):
            result = await classic_handler.delete_label("OldLabel")

            assert isinstance(result, FolderOperationResponse)
            assert result.success is True
            assert result.folder_name == "OldLabel"
            mock_delete.assert_called_once_with("Labels/OldLabel")


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


class TestEmailClientLabels:
    """Test EmailClient label operations."""

    @pytest.mark.asyncio
    async def test_list_labels_filters_labels_prefix(self, email_client):
        """Test list_labels filters only Labels/ prefix folders."""
        mock_folders = [
            Folder(name="INBOX", delimiter="/", flags=[]),
            Folder(name="Sent", delimiter="/", flags=[]),
            Folder(name="Labels/Important", delimiter="/", flags=[]),
            Folder(name="Labels/Work", delimiter="/", flags=[]),
            Folder(name="Folders/Archive", delimiter="/", flags=[]),
        ]

        mock_list = AsyncMock(return_value=mock_folders)

        with patch.object(email_client, "list_folders", mock_list):
            result = await email_client.list_labels()

            assert len(result) == 2
            assert result[0].name == "Important"
            assert result[0].full_path == "Labels/Important"
            assert result[1].name == "Work"
            assert result[1].full_path == "Labels/Work"

    @pytest.mark.asyncio
    async def test_list_labels_empty(self, email_client):
        """Test list_labels returns empty list when no labels exist."""
        mock_folders = [
            Folder(name="INBOX", delimiter="/", flags=[]),
            Folder(name="Sent", delimiter="/", flags=[]),
        ]

        mock_list = AsyncMock(return_value=mock_folders)

        with patch.object(email_client, "list_folders", mock_list):
            result = await email_client.list_labels()

            assert len(result) == 0

    @pytest.mark.asyncio
    async def test_get_email_message_id(self, email_client):
        """Test get_email_message_id IMAP operation."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        mock_imap.uid = AsyncMock(
            return_value=(
                "OK",
                [bytearray(b"Message-ID: <unique123@example.com>")],
            )
        )
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            result = await email_client.get_email_message_id("123", "INBOX")

            assert result == "<unique123@example.com>"

    @pytest.mark.asyncio
    async def test_search_by_message_id(self, email_client):
        """Test search_by_message_id IMAP operation."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        # search returns sequence numbers
        mock_imap.search = AsyncMock(return_value=("OK", [b"1"]))
        # fetch returns UID for the sequence number
        mock_imap.fetch = AsyncMock(return_value=("OK", [b"1 FETCH (UID 456)"]))
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            result = await email_client.search_by_message_id("<unique123@example.com>", "Labels/Important")

            assert result == "456"
            mock_imap.select.assert_called_once_with('"Labels/Important"')

    @pytest.mark.asyncio
    async def test_search_by_message_id_not_found(self, email_client):
        """Test search_by_message_id returns None when email not found."""
        mock_imap = AsyncMock()
        mock_imap._client_task = asyncio.Future()
        mock_imap._client_task.set_result(None)
        mock_imap.wait_hello_from_server = AsyncMock()
        mock_imap.login = AsyncMock()
        mock_imap.select = AsyncMock()
        # search returns empty when not found
        mock_imap.search = AsyncMock(return_value=("OK", [b""]))
        mock_imap.logout = AsyncMock()

        with patch.object(email_client, "imap_class", return_value=mock_imap):
            result = await email_client.search_by_message_id("<notfound@example.com>", "Labels/Work")

            assert result is None


class TestEmailClientLabelEdgeCases:
    """Test edge cases for EmailClient label operations."""

    @pytest.mark.asyncio
    async def test_list_labels_skips_labels_folder_itself(self, email_client):
        """Test that list_labels skips the 'Labels' folder itself if it exists."""
        mock_folders = [
            Folder(name="Labels", delimiter="/", flags=["\\HasChildren"]),
            Folder(name="Labels/Important", delimiter="/", flags=[]),
        ]

        mock_list = AsyncMock(return_value=mock_folders)

        with patch.object(email_client, "list_folders", mock_list):
            result = await email_client.list_labels()

            # Should only include "Important", not the "Labels" folder itself
            assert len(result) == 1
            assert result[0].name == "Important"

    @pytest.mark.asyncio
    async def test_remove_label_email_not_in_label(self, classic_handler):
        """Test remove_label when email is not in the label folder."""
        mock_get_message_id = AsyncMock(return_value="<msg123@example.com>")
        mock_search = AsyncMock(return_value=None)  # Email not found in label

        with patch.object(classic_handler.incoming_client, "get_email_message_id", mock_get_message_id):
            with patch.object(classic_handler.incoming_client, "search_by_message_id", mock_search):
                result = await classic_handler.remove_label(
                    email_ids=["123"],
                    label_name="Important",
                )

                assert isinstance(result, EmailMoveResponse)
                assert result.success is False
                assert result.failed_ids == ["123"]
                assert result.moved_ids == []

    @pytest.mark.asyncio
    async def test_get_email_labels_no_message_id(self, classic_handler):
        """Test get_email_labels when email has no Message-ID."""
        mock_get_message_id = AsyncMock(return_value=None)

        with patch.object(classic_handler.incoming_client, "get_email_message_id", mock_get_message_id):
            result = await classic_handler.get_email_labels(
                email_id="123",
                source_mailbox="INBOX",
            )

            assert isinstance(result, EmailLabelsResponse)
            assert result.email_id == "123"
            assert result.labels == []
