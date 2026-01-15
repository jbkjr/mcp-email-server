import email.utils
import mimetypes
from collections.abc import AsyncGenerator
from datetime import datetime, timezone
from email.header import Header
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.parser import BytesParser
from email.policy import default
from pathlib import Path
from typing import Any

import aioimaplib
import aiosmtplib

from mcp_email_server.config import EmailServer, EmailSettings
from mcp_email_server.emails import EmailHandler
from mcp_email_server.emails.models import (
    AttachmentDownloadResponse,
    EmailBodyResponse,
    EmailContentBatchResponse,
    EmailLabelsResponse,
    EmailMetadata,
    EmailMetadataPageResponse,
    EmailMoveResponse,
    Folder,
    FolderListResponse,
    FolderOperationResponse,
    Label,
    LabelListResponse,
)
from mcp_email_server.log import logger


def _quote_mailbox(mailbox: str) -> str:
    """Quote mailbox name for IMAP compatibility.

    Some IMAP servers (notably Proton Mail Bridge) require mailbox names
    to be quoted. This is valid per RFC 3501 and works with all IMAP servers.

    Per RFC 3501 Section 9 (Formal Syntax), quoted strings must escape
    backslashes and double-quote characters with a preceding backslash.

    See: https://github.com/ai-zerolab/mcp-email-server/issues/87
    See: https://www.rfc-editor.org/rfc/rfc3501#section-9
    """
    # Per RFC 3501, literal double-quote characters in a quoted string must
    # be escaped with a backslash. Backslashes themselves must also be escaped.
    escaped = mailbox.replace("\\", "\\\\").replace('"', r"\"")
    return f'"{escaped}"'


async def _send_imap_id(imap: aioimaplib.IMAP4 | aioimaplib.IMAP4_SSL) -> None:
    """Send IMAP ID command with fallback for strict servers like 163.com.

    aioimaplib's id() method sends ID command with spaces between parentheses
    and content (e.g., 'ID ( "name" "value" )'), which some strict IMAP servers
    like 163.com reject with 'BAD Parse command error'.

    This function first tries the standard id() method, and if it fails,
    falls back to sending a raw command with correct format.

    See: https://github.com/ai-zerolab/mcp-email-server/issues/85
    """
    try:
        response = await imap.id(name="mcp-email-server", version="1.0.0")
        if response.result != "OK":
            # Fallback for strict servers (e.g., 163.com)
            # Send raw command with correct parenthesis format
            await imap.protocol.execute(
                aioimaplib.Command(
                    "ID",
                    imap.protocol.new_tag(),
                    '("name" "mcp-email-server" "version" "1.0.0")',
                )
            )
    except Exception as e:
        logger.warning(f"IMAP ID command failed: {e!s}")


def _has_sort_capability(imap: aioimaplib.IMAP4 | aioimaplib.IMAP4_SSL) -> bool:
    """Check if the IMAP server supports the SORT extension (RFC 5256).

    The SORT capability allows server-side sorting of emails, which is much
    more efficient than fetching all emails and sorting client-side.
    """
    try:
        capabilities = imap.protocol.capabilities
        return "SORT" in capabilities
    except Exception:
        return False


class EmailClient:
    def __init__(self, email_server: EmailServer, sender: str | None = None):
        self.email_server = email_server
        self.sender = sender or email_server.user_name

        self.imap_class = aioimaplib.IMAP4_SSL if self.email_server.use_ssl else aioimaplib.IMAP4

        self.smtp_use_tls = self.email_server.use_ssl
        self.smtp_start_tls = self.email_server.start_ssl

    def _parse_email_data(self, raw_email: bytes, email_id: str | None = None) -> dict[str, Any]:  # noqa: C901
        """Parse raw email data into a structured dictionary."""
        parser = BytesParser(policy=default)
        email_message = parser.parsebytes(raw_email)

        # Extract email parts
        subject = email_message.get("Subject", "")
        sender = email_message.get("From", "")
        date_str = email_message.get("Date", "")

        # Extract Message-ID for reply threading
        message_id = email_message.get("Message-ID")

        # Extract recipients
        to_addresses = []
        to_header = email_message.get("To", "")
        if to_header:
            # Simple parsing - split by comma and strip whitespace
            to_addresses = [addr.strip() for addr in to_header.split(",")]

        # Also check CC recipients
        cc_header = email_message.get("Cc", "")
        if cc_header:
            to_addresses.extend([addr.strip() for addr in cc_header.split(",")])

        # Parse date
        try:
            date_tuple = email.utils.parsedate_tz(date_str)
            date = (
                datetime.fromtimestamp(email.utils.mktime_tz(date_tuple), tz=timezone.utc)
                if date_tuple
                else datetime.now(timezone.utc)
            )
        except Exception:
            date = datetime.now(timezone.utc)

        # Get body content
        body = ""
        attachments = []

        if email_message.is_multipart():
            for part in email_message.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition", ""))

                # Handle attachments
                if "attachment" in content_disposition:
                    filename = part.get_filename()
                    if filename:
                        attachments.append(filename)
                # Handle text parts
                elif content_type == "text/plain":
                    body_part = part.get_payload(decode=True)
                    if body_part:
                        charset = part.get_content_charset("utf-8")
                        try:
                            body += body_part.decode(charset)
                        except UnicodeDecodeError:
                            body += body_part.decode("utf-8", errors="replace")
        else:
            # Handle plain text emails
            payload = email_message.get_payload(decode=True)
            if payload:
                charset = email_message.get_content_charset("utf-8")
                try:
                    body = payload.decode(charset)
                except UnicodeDecodeError:
                    body = payload.decode("utf-8", errors="replace")
        # TODO: Allow retrieving full email body
        if body and len(body) > 20000:
            body = body[:20000] + "...[TRUNCATED]"
        return {
            "email_id": email_id or "",
            "message_id": message_id,
            "subject": subject,
            "from": sender,
            "to": to_addresses,
            "body": body,
            "date": date,
            "attachments": attachments,
        }

    @staticmethod
    def _build_search_criteria(
        before: datetime | None = None,
        since: datetime | None = None,
        subject: str | None = None,
        body: str | None = None,
        text: str | None = None,
        from_address: str | None = None,
        to_address: str | None = None,
    ):
        search_criteria = []
        if before:
            search_criteria.extend(["BEFORE", before.strftime("%d-%b-%Y").upper()])
        if since:
            search_criteria.extend(["SINCE", since.strftime("%d-%b-%Y").upper()])
        if subject:
            search_criteria.extend(["SUBJECT", subject])
        if body:
            search_criteria.extend(["BODY", body])
        if text:
            search_criteria.extend(["TEXT", text])
        if from_address:
            search_criteria.extend(["FROM", from_address])
        if to_address:
            search_criteria.extend(["TO", to_address])

        # If no specific criteria, search for ALL
        if not search_criteria:
            search_criteria = ["ALL"]

        return search_criteria

    @staticmethod
    def _parse_date_from_header(date_str: str) -> datetime:
        """Parse a date string from an email header into a datetime object."""
        try:
            date_tuple = email.utils.parsedate_tz(date_str)
            if date_tuple:
                return datetime.fromtimestamp(email.utils.mktime_tz(date_tuple), tz=timezone.utc)
        except Exception as e:
            logger.debug(f"Failed to parse date '{date_str}': {e}")
        return datetime.now(timezone.utc)

    async def _batch_fetch_dates(
        self,
        imap: aioimaplib.IMAP4 | aioimaplib.IMAP4_SSL,
        email_ids: list[bytes],
    ) -> list[tuple[str, datetime]]:
        """Batch fetch only Date headers for a list of email UIDs.

        Returns a list of (email_id, date) tuples.
        This is much more efficient than fetching full headers when we only need dates.

        Note: Different IMAP servers return different response formats:
        - Some include UID in the FETCH line: b'1 FETCH (UID 1 BODY[...]'
        - Others (like Proton Bridge) return UID separately: b' UID 1)'
        This method handles both formats.
        """
        if not email_ids:
            return []

        # Join UIDs for batch fetch: "1,2,3,4,5"
        uid_list = ",".join(uid.decode("utf-8") for uid in email_ids)

        try:
            import re

            # Fetch only the Date header field - much smaller than full headers
            _, data = await imap.uid("fetch", uid_list, "BODY.PEEK[HEADER.FIELDS (DATE)]")

            results: list[tuple[str, datetime]] = []
            pending_uid: str | None = None
            pending_date: datetime | None = None

            for item in data:
                if isinstance(item, bytearray):
                    # This is the date header content
                    date_str = bytes(item).decode("utf-8", errors="replace").strip()
                    # Remove "Date: " prefix if present
                    if date_str.lower().startswith("date:"):
                        date_str = date_str[5:].strip()
                    pending_date = self._parse_date_from_header(date_str)
                    # If we already have a UID (from before the data), emit result
                    if pending_uid is not None:
                        results.append((pending_uid, pending_date))
                        pending_uid = None
                        pending_date = None
                elif isinstance(item, bytes):
                    item_str = item.decode("utf-8", errors="replace")
                    # Look for UID in any bytes item
                    uid_match = re.search(r"UID\s+(\d+)", item_str)
                    if uid_match:
                        if pending_date is not None:
                            # UID came after the data
                            results.append((uid_match.group(1), pending_date))
                            pending_date = None
                        else:
                            # UID came before the data - save it
                            pending_uid = uid_match.group(1)

            return results
        except Exception as e:
            logger.error(f"Error in batch fetch dates: {e}")
            return []

    async def _batch_fetch_headers(
        self,
        imap: aioimaplib.IMAP4 | aioimaplib.IMAP4_SSL,
        email_ids: list[str],
    ) -> list[dict[str, Any]]:
        """Batch fetch full headers for a list of email UIDs.

        Returns a list of metadata dictionaries.

        Note: Different IMAP servers return different response formats:
        - Some include UID in the FETCH line: b'1 FETCH (UID 1 BODY[...]'
        - Others (like Proton Bridge) return UID separately: b' UID 1)'
        This method handles both formats.
        """
        if not email_ids:
            return []

        # Join UIDs for batch fetch: "1,2,3,4,5"
        uid_list = ",".join(email_ids)

        try:
            import re

            _, data = await imap.uid("fetch", uid_list, "BODY.PEEK[HEADER]")

            results: list[dict[str, Any]] = []
            pending_uid: str | None = None
            pending_headers: bytes | None = None

            for item in data:
                if isinstance(item, bytearray):
                    # Store headers; emit if we already have a UID
                    pending_headers = bytes(item)
                    if pending_uid is not None:
                        self._append_header_metadata(results, pending_uid, pending_headers)
                        pending_uid, pending_headers = None, None
                elif isinstance(item, bytes):
                    uid_match = re.search(r"UID\s+(\d+)", item.decode("utf-8", errors="replace"))
                    if not uid_match:
                        continue
                    if pending_headers is not None:
                        # UID came after the data
                        self._append_header_metadata(results, uid_match.group(1), pending_headers)
                        pending_headers = None
                    else:
                        # UID came before the data - save it
                        pending_uid = uid_match.group(1)

            return results
        except Exception as e:
            logger.error(f"Error in batch fetch headers: {e}")
            return []

    def _append_header_metadata(self, results: list[dict[str, Any]], uid: str, headers: bytes) -> None:
        """Parse headers and append to results if successful."""
        metadata = self._parse_header_to_metadata(uid, headers)
        if metadata:
            results.append(metadata)

    def _parse_header_to_metadata(self, email_id: str, raw_headers: bytes) -> dict[str, Any] | None:
        """Parse raw email headers into a metadata dictionary."""
        try:
            parser = BytesParser(policy=default)
            email_message = parser.parsebytes(raw_headers)

            subject = email_message.get("Subject", "")
            sender = email_message.get("From", "")
            date_str = email_message.get("Date", "")

            to_addresses = []
            to_header = email_message.get("To", "")
            if to_header:
                to_addresses = [addr.strip() for addr in to_header.split(",")]

            cc_header = email_message.get("Cc", "")
            if cc_header:
                to_addresses.extend([addr.strip() for addr in cc_header.split(",")])

            date = self._parse_date_from_header(date_str)

            return {
                "email_id": email_id,
                "subject": subject,
                "from": sender,
                "to": to_addresses,
                "date": date,
                "attachments": [],
            }
        except Exception as e:
            logger.error(f"Error parsing header metadata: {e}")
            return None

    async def get_email_count(
        self,
        before: datetime | None = None,
        since: datetime | None = None,
        subject: str | None = None,
        from_address: str | None = None,
        to_address: str | None = None,
        mailbox: str = "INBOX",
    ) -> int:
        imap = self.imap_class(self.email_server.host, self.email_server.port)
        try:
            # Wait for the connection to be established
            await imap._client_task
            await imap.wait_hello_from_server()

            # Login and select inbox
            await imap.login(self.email_server.user_name, self.email_server.password)
            await _send_imap_id(imap)
            await imap.select(_quote_mailbox(mailbox))
            search_criteria = self._build_search_criteria(
                before, since, subject, from_address=from_address, to_address=to_address
            )
            logger.info(f"Count: Search criteria: {search_criteria}")
            # Search for messages and count them - use UID SEARCH for consistency
            _, messages = await imap.uid_search(*search_criteria)
            return len(messages[0].split())
        finally:
            # Ensure we logout properly
            try:
                await imap.logout()
            except Exception as e:
                logger.info(f"Error during logout: {e}")

    async def get_emails_metadata_stream(  # noqa: C901
        self,
        page: int = 1,
        page_size: int = 10,
        before: datetime | None = None,
        since: datetime | None = None,
        subject: str | None = None,
        from_address: str | None = None,
        to_address: str | None = None,
        order: str = "desc",
        mailbox: str = "INBOX",
    ) -> AsyncGenerator[dict[str, Any], None]:
        imap = self.imap_class(self.email_server.host, self.email_server.port)
        try:
            # Wait for the connection to be established
            await imap._client_task
            await imap.wait_hello_from_server()

            # Login and select inbox
            await imap.login(self.email_server.user_name, self.email_server.password)
            await _send_imap_id(imap)
            await imap.select(_quote_mailbox(mailbox))

            search_criteria = self._build_search_criteria(
                before, since, subject, from_address=from_address, to_address=to_address
            )
            logger.info(f"Get metadata: Search criteria: {search_criteria}")

            # Calculate pagination offsets
            start = (page - 1) * page_size
            end = start + page_size

            # Check if server supports SORT extension (RFC 5256)
            if _has_sort_capability(imap):
                # Use server-side sorting - much more efficient
                sort_order = "(REVERSE DATE)" if order == "desc" else "(DATE)"
                logger.info(f"Using IMAP SORT with {sort_order}")

                try:
                    _, sort_response = await imap.uid("sort", sort_order, "UTF-8", *search_criteria)

                    if not sort_response or not sort_response[0]:
                        logger.warning("No messages returned from SORT")
                        return

                    # Parse sorted UIDs
                    sorted_uids = sort_response[0].split()
                    logger.info(f"SORT returned {len(sorted_uids)} UIDs")

                    # Paginate the sorted UIDs
                    page_uids = [uid.decode("utf-8") for uid in sorted_uids[start:end]]

                    if not page_uids:
                        return

                    # Batch fetch full headers for just the page
                    metadata_list = await self._batch_fetch_headers(imap, page_uids)

                    # Sort the results to match the SORT order (batch fetch may return unordered)
                    uid_order = {uid: i for i, uid in enumerate(page_uids)}
                    metadata_list.sort(key=lambda m: uid_order.get(m["email_id"], 999999))

                    for metadata in metadata_list:
                        yield metadata
                    return

                except Exception as e:
                    logger.warning(f"SORT command failed, falling back to batch fetch: {e}")
                    # Fall through to batch fetch fallback

            # Fallback: Batch fetch approach (for servers without SORT)
            # This is still much faster than the old N individual fetches
            logger.info("Using batch fetch fallback (server doesn't support SORT)")

            # Search for messages
            _, messages = await imap.uid_search(*search_criteria)

            if not messages or not messages[0]:
                logger.warning("No messages returned from search")
                return

            email_ids = messages[0].split()
            logger.info(f"Found {len(email_ids)} email IDs")

            if not email_ids:
                return

            # Batch fetch just the Date headers for all emails (much smaller than full headers)
            date_tuples = await self._batch_fetch_dates(imap, email_ids)

            if not date_tuples:
                # Fallback: if batch date fetch failed, try with full headers
                logger.warning("Batch date fetch returned no results, using full header fetch")
                all_uids = [uid.decode("utf-8") for uid in email_ids]
                all_metadata = await self._batch_fetch_headers(imap, all_uids)
                all_metadata.sort(key=lambda x: x["date"], reverse=(order == "desc"))
                for metadata in all_metadata[start:end]:
                    yield metadata
                return

            # Sort by date
            date_tuples.sort(key=lambda x: x[1], reverse=(order == "desc"))

            # Paginate
            page_tuples = date_tuples[start:end]
            page_uids = [uid for uid, _ in page_tuples]

            if not page_uids:
                return

            # Batch fetch full headers for just the page
            metadata_list = await self._batch_fetch_headers(imap, page_uids)

            # Sort results to match the date order
            uid_order = {uid: i for i, uid in enumerate(page_uids)}
            metadata_list.sort(key=lambda m: uid_order.get(m["email_id"], 999999))

            for metadata in metadata_list:
                yield metadata

        finally:
            # Ensure we logout properly
            try:
                await imap.logout()
            except Exception as e:
                logger.info(f"Error during logout: {e}")

    def _check_email_content(self, data: list) -> bool:
        """Check if the fetched data contains actual email content."""
        for item in data:
            if isinstance(item, bytes) and b"FETCH (" in item and b"RFC822" not in item and b"BODY" not in item:
                # This is just metadata, not actual content
                continue
            elif isinstance(item, bytes | bytearray) and len(item) > 100:
                # This looks like email content
                return True
        return False

    def _extract_raw_email(self, data: list) -> bytes | None:
        """Extract raw email bytes from IMAP response data."""
        # The email content is typically at index 1 as a bytearray
        if len(data) > 1 and isinstance(data[1], bytearray):
            return bytes(data[1])

        # Search through all items for email content
        for item in data:
            if isinstance(item, bytes | bytearray) and len(item) > 100:
                # Skip IMAP protocol responses
                if isinstance(item, bytes) and b"FETCH" in item:
                    continue
                # This is likely the email content
                return bytes(item) if isinstance(item, bytearray) else item
        return None

    async def _fetch_email_with_formats(self, imap, email_id: str) -> list | None:
        """Try different fetch formats to get email data."""
        fetch_formats = ["RFC822", "BODY[]", "BODY.PEEK[]", "(BODY.PEEK[])"]

        for fetch_format in fetch_formats:
            try:
                _, data = await imap.uid("fetch", email_id, fetch_format)

                if data and len(data) > 0 and self._check_email_content(data):
                    return data

            except Exception as e:
                logger.debug(f"Fetch format {fetch_format} failed: {e}")

        return None

    async def get_email_body_by_id(self, email_id: str, mailbox: str = "INBOX") -> dict[str, Any] | None:
        imap = self.imap_class(self.email_server.host, self.email_server.port)
        try:
            # Wait for the connection to be established
            await imap._client_task
            await imap.wait_hello_from_server()

            # Login and select inbox
            await imap.login(self.email_server.user_name, self.email_server.password)
            await _send_imap_id(imap)
            await imap.select(_quote_mailbox(mailbox))

            # Fetch the specific email by UID
            data = await self._fetch_email_with_formats(imap, email_id)
            if not data:
                logger.error(f"Failed to fetch UID {email_id} with any format")
                return None

            # Extract raw email data
            raw_email = self._extract_raw_email(data)
            if not raw_email:
                logger.error(f"Could not find email data in response for email ID: {email_id}")
                return None

            # Parse the email
            try:
                return self._parse_email_data(raw_email, email_id)
            except Exception as e:
                logger.error(f"Error parsing email: {e!s}")
                return None

        finally:
            # Ensure we logout properly
            try:
                await imap.logout()
            except Exception as e:
                logger.info(f"Error during logout: {e}")

    async def download_attachment(
        self,
        email_id: str,
        attachment_name: str,
        save_path: str,
        mailbox: str = "INBOX",
    ) -> dict[str, Any]:
        """Download a specific attachment from an email and save it to disk.

        Args:
            email_id: The UID of the email containing the attachment.
            attachment_name: The filename of the attachment to download.
            save_path: The local path where the attachment will be saved.
            mailbox: The mailbox to search in (default: "INBOX").

        Returns:
            A dictionary with download result information.
        """
        imap = self.imap_class(self.email_server.host, self.email_server.port)
        try:
            await imap._client_task
            await imap.wait_hello_from_server()

            await imap.login(self.email_server.user_name, self.email_server.password)
            await _send_imap_id(imap)
            await imap.select(_quote_mailbox(mailbox))

            data = await self._fetch_email_with_formats(imap, email_id)
            if not data:
                msg = f"Failed to fetch email with UID {email_id}"
                logger.error(msg)
                raise ValueError(msg)

            raw_email = self._extract_raw_email(data)
            if not raw_email:
                msg = f"Could not find email data for email ID: {email_id}"
                logger.error(msg)
                raise ValueError(msg)

            parser = BytesParser(policy=default)
            email_message = parser.parsebytes(raw_email)

            # Find the attachment
            attachment_data = None
            mime_type = None

            if email_message.is_multipart():
                for part in email_message.walk():
                    content_disposition = str(part.get("Content-Disposition", ""))
                    if "attachment" in content_disposition:
                        filename = part.get_filename()
                        if filename == attachment_name:
                            attachment_data = part.get_payload(decode=True)
                            mime_type = part.get_content_type()
                            break

            if attachment_data is None:
                msg = f"Attachment '{attachment_name}' not found in email {email_id}"
                logger.error(msg)
                raise ValueError(msg)

            # Save to disk
            save_file = Path(save_path)
            save_file.parent.mkdir(parents=True, exist_ok=True)
            save_file.write_bytes(attachment_data)

            logger.info(f"Attachment '{attachment_name}' saved to {save_path}")

            return {
                "email_id": email_id,
                "attachment_name": attachment_name,
                "mime_type": mime_type or "application/octet-stream",
                "size": len(attachment_data),
                "saved_path": str(save_file.resolve()),
            }

        finally:
            try:
                await imap.logout()
            except Exception as e:
                logger.info(f"Error during logout: {e}")

    def _validate_attachment(self, file_path: str) -> Path:
        """Validate attachment file path."""
        path = Path(file_path)
        if not path.exists():
            msg = f"Attachment file not found: {file_path}"
            logger.error(msg)
            raise FileNotFoundError(msg)

        if not path.is_file():
            msg = f"Attachment path is not a file: {file_path}"
            logger.error(msg)
            raise ValueError(msg)

        return path

    def _create_attachment_part(self, path: Path) -> MIMEApplication:
        """Create MIME attachment part from file."""
        with open(path, "rb") as f:
            file_data = f.read()

        mime_type, _ = mimetypes.guess_type(str(path))
        if mime_type is None:
            mime_type = "application/octet-stream"

        attachment_part = MIMEApplication(file_data, _subtype=mime_type.split("/")[1])
        attachment_part.add_header(
            "Content-Disposition",
            "attachment",
            filename=path.name,
        )
        logger.info(f"Attached file: {path.name} ({mime_type})")
        return attachment_part

    def _create_message_with_attachments(self, body: str, html: bool, attachments: list[str]) -> MIMEMultipart:
        """Create multipart message with attachments."""
        msg = MIMEMultipart()
        content_type = "html" if html else "plain"
        text_part = MIMEText(body, content_type, "utf-8")
        msg.attach(text_part)

        for file_path in attachments:
            try:
                path = self._validate_attachment(file_path)
                attachment_part = self._create_attachment_part(path)
                msg.attach(attachment_part)
            except Exception as e:
                logger.error(f"Failed to attach file {file_path}: {e}")
                raise

        return msg

    async def send_email(
        self,
        recipients: list[str],
        subject: str,
        body: str,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        html: bool = False,
        attachments: list[str] | None = None,
        in_reply_to: str | None = None,
        references: str | None = None,
    ):
        # Create message with or without attachments
        if attachments:
            msg = self._create_message_with_attachments(body, html, attachments)
        else:
            content_type = "html" if html else "plain"
            msg = MIMEText(body, content_type, "utf-8")

        # Handle subject with special characters
        if any(ord(c) > 127 for c in subject):
            msg["Subject"] = Header(subject, "utf-8")
        else:
            msg["Subject"] = subject

        # Handle sender name with special characters
        if any(ord(c) > 127 for c in self.sender):
            msg["From"] = Header(self.sender, "utf-8")
        else:
            msg["From"] = self.sender

        msg["To"] = ", ".join(recipients)

        # Add CC header if provided (visible to recipients)
        if cc:
            msg["Cc"] = ", ".join(cc)

        # Set threading headers for replies
        if in_reply_to:
            msg["In-Reply-To"] = in_reply_to
        if references:
            msg["References"] = references

        # Note: BCC recipients are not added to headers (they remain hidden)
        # but will be included in the actual recipients for SMTP delivery

        async with aiosmtplib.SMTP(
            hostname=self.email_server.host,
            port=self.email_server.port,
            start_tls=self.smtp_start_tls,
            use_tls=self.smtp_use_tls,
        ) as smtp:
            await smtp.login(self.email_server.user_name, self.email_server.password)

            # Create a combined list of all recipients for delivery
            all_recipients = recipients.copy()
            if cc:
                all_recipients.extend(cc)
            if bcc:
                all_recipients.extend(bcc)

            await smtp.send_message(msg, recipients=all_recipients)

        # Return the message for potential saving to Sent folder
        return msg

    async def _find_sent_folder_by_flag(self, imap) -> str | None:
        """Find the Sent folder by searching for the \\Sent IMAP flag.

        Args:
            imap: Connected IMAP client

        Returns:
            The folder name with the \\Sent flag, or None if not found
        """
        try:
            # List all folders - aioimaplib requires reference_name and mailbox_pattern
            _, folders = await imap.list('""', "*")

            # Search for folder with \Sent flag
            for folder in folders:
                folder_str = folder.decode("utf-8") if isinstance(folder, bytes) else str(folder)
                # IMAP LIST response format: (flags) "delimiter" "name"
                # Example: (\Sent \HasNoChildren) "/" "Gesendete Objekte"
                if r"\Sent" in folder_str or "\\Sent" in folder_str:
                    # Extract folder name from the response
                    # Split by quotes and get the last quoted part
                    parts = folder_str.split('"')
                    if len(parts) >= 3:
                        folder_name = parts[-2]  # The folder name is the second-to-last quoted part
                        logger.info(f"Found Sent folder by \\Sent flag: '{folder_name}'")
                        return folder_name
        except Exception as e:
            logger.debug(f"Error finding Sent folder by flag: {e}")

        return None

    async def append_to_sent(
        self,
        msg: MIMEText | MIMEMultipart,
        incoming_server: EmailServer,
        sent_folder_name: str | None = None,
    ) -> bool:
        """Append a sent message to the IMAP Sent folder.

        Args:
            msg: The email message that was sent
            incoming_server: IMAP server configuration for accessing Sent folder
            sent_folder_name: Override folder name, or None for auto-detection

        Returns:
            True if successfully saved, False otherwise
        """
        imap_class = aioimaplib.IMAP4_SSL if incoming_server.use_ssl else aioimaplib.IMAP4
        imap = imap_class(incoming_server.host, incoming_server.port)

        # Common Sent folder names across different providers
        sent_folder_candidates = [
            sent_folder_name,  # User-specified override (if provided)
            "Sent",
            "INBOX.Sent",
            "Sent Items",
            "Sent Mail",
            "[Gmail]/Sent Mail",
            "INBOX/Sent",
        ]
        # Filter out None values
        sent_folder_candidates = [f for f in sent_folder_candidates if f]

        try:
            await imap._client_task
            await imap.wait_hello_from_server()
            await imap.login(incoming_server.user_name, incoming_server.password)
            await _send_imap_id(imap)

            # Try to find Sent folder by IMAP \Sent flag first
            flag_folder = await self._find_sent_folder_by_flag(imap)
            if flag_folder and flag_folder not in sent_folder_candidates:
                # Add it at the beginning (high priority)
                sent_folder_candidates.insert(0, flag_folder)

            # Try to find and use the Sent folder
            for folder in sent_folder_candidates:
                try:
                    logger.debug(f"Trying Sent folder: '{folder}'")
                    # Try to select the folder to verify it exists
                    result = await imap.select(_quote_mailbox(folder))
                    logger.debug(f"Select result for '{folder}': {result}")

                    # aioimaplib returns (status, data) where status is a string like 'OK' or 'NO'
                    status = result[0] if isinstance(result, tuple) else result
                    if str(status).upper() == "OK":
                        # Folder exists, append the message
                        msg_bytes = msg.as_bytes()
                        logger.debug(f"Appending message to '{folder}'")
                        # aioimaplib.append signature: (message_bytes, mailbox, flags, date)
                        append_result = await imap.append(
                            msg_bytes,
                            mailbox=_quote_mailbox(folder),
                            flags=r"(\Seen)",
                        )
                        logger.debug(f"Append result: {append_result}")
                        append_status = append_result[0] if isinstance(append_result, tuple) else append_result
                        if str(append_status).upper() == "OK":
                            logger.info(f"Saved sent email to '{folder}'")
                            return True
                        else:
                            logger.warning(f"Failed to append to '{folder}': {append_status}")
                    else:
                        logger.debug(f"Folder '{folder}' select returned: {status}")
                except Exception as e:
                    logger.debug(f"Folder '{folder}' not available: {e}")
                    continue

            logger.warning("Could not find a valid Sent folder to save the message")
            return False

        except Exception as e:
            logger.error(f"Error saving to Sent folder: {e}")
            return False
        finally:
            try:
                await imap.logout()
            except Exception as e:
                logger.debug(f"Error during logout: {e}")

    async def delete_emails(self, email_ids: list[str], mailbox: str = "INBOX") -> tuple[list[str], list[str]]:
        """Delete emails by their UIDs. Returns (deleted_ids, failed_ids)."""
        imap = self.imap_class(self.email_server.host, self.email_server.port)
        deleted_ids = []
        failed_ids = []

        try:
            await imap._client_task
            await imap.wait_hello_from_server()
            await imap.login(self.email_server.user_name, self.email_server.password)
            await _send_imap_id(imap)
            await imap.select(_quote_mailbox(mailbox))

            for email_id in email_ids:
                try:
                    await imap.uid("store", email_id, "+FLAGS", r"(\Deleted)")
                    deleted_ids.append(email_id)
                except Exception as e:
                    logger.error(f"Failed to delete email {email_id}: {e}")
                    failed_ids.append(email_id)

            await imap.expunge()
        finally:
            try:
                await imap.logout()
            except Exception as e:
                logger.info(f"Error during logout: {e}")

        return deleted_ids, failed_ids

    def _parse_list_response(self, folder_data: bytes | str) -> Folder | None:
        """Parse a single IMAP LIST response line into a Folder object.

        IMAP LIST response format: (flags) "delimiter" "name"
        Example: (\\HasNoChildren \\Sent) "/" "Sent"
        """
        folder_str = folder_data.decode("utf-8") if isinstance(folder_data, bytes) else str(folder_data)

        # Skip empty or invalid responses
        if not folder_str or folder_str == "LIST completed.":
            return None

        try:
            # Extract flags (content between first set of parentheses)
            flags_start = folder_str.find("(")
            flags_end = folder_str.find(")")
            if flags_start == -1 or flags_end == -1:
                return None

            flags_str = folder_str[flags_start + 1 : flags_end]
            flags = [f.strip() for f in flags_str.split() if f.strip()]

            # Extract delimiter and name from the rest
            # Format after flags: "delimiter" "name"
            rest = folder_str[flags_end + 1 :].strip()
            parts = rest.split('"')
            # parts should be like: ['', '/', ' ', 'INBOX', '']
            if len(parts) >= 4:
                delimiter = parts[1]
                folder_name = parts[3]
                return Folder(name=folder_name, delimiter=delimiter, flags=flags)

        except Exception as e:
            logger.debug(f"Error parsing folder response '{folder_str}': {e}")

        return None

    async def list_folders(self) -> list[Folder]:
        """List all folders/mailboxes."""
        imap = self.imap_class(self.email_server.host, self.email_server.port)
        folders = []

        try:
            await imap._client_task
            await imap.wait_hello_from_server()
            await imap.login(self.email_server.user_name, self.email_server.password)
            await _send_imap_id(imap)

            # List all folders
            _, folder_data = await imap.list('""', "*")

            for item in folder_data:
                folder = self._parse_list_response(item)
                if folder:
                    folders.append(folder)

            logger.info(f"Found {len(folders)} folders")
            return folders

        finally:
            try:
                await imap.logout()
            except Exception as e:
                logger.info(f"Error during logout: {e}")

    async def copy_emails(
        self,
        email_ids: list[str],
        destination_folder: str,
        source_mailbox: str = "INBOX",
    ) -> tuple[list[str], list[str]]:
        """Copy emails to a destination folder. Returns (copied_ids, failed_ids)."""
        imap = self.imap_class(self.email_server.host, self.email_server.port)
        copied_ids = []
        failed_ids = []

        try:
            await imap._client_task
            await imap.wait_hello_from_server()
            await imap.login(self.email_server.user_name, self.email_server.password)
            await _send_imap_id(imap)
            await imap.select(_quote_mailbox(source_mailbox))

            for email_id in email_ids:
                try:
                    result = await imap.uid("copy", email_id, _quote_mailbox(destination_folder))
                    status = result[0] if isinstance(result, tuple) else result
                    if str(status).upper() == "OK":
                        copied_ids.append(email_id)
                        logger.debug(f"Copied email {email_id} to {destination_folder}")
                    else:
                        logger.error(f"Failed to copy email {email_id}: {status}")
                        failed_ids.append(email_id)
                except Exception as e:
                    logger.error(f"Failed to copy email {email_id}: {e}")
                    failed_ids.append(email_id)

        finally:
            try:
                await imap.logout()
            except Exception as e:
                logger.info(f"Error during logout: {e}")

        return copied_ids, failed_ids

    async def move_emails(
        self,
        email_ids: list[str],
        destination_folder: str,
        source_mailbox: str = "INBOX",
    ) -> tuple[list[str], list[str]]:
        """Move emails to a destination folder. Returns (moved_ids, failed_ids).

        Attempts to use MOVE command first (RFC 6851), falls back to COPY + DELETE.
        """
        imap = self.imap_class(self.email_server.host, self.email_server.port)
        moved_ids = []
        failed_ids = []

        try:
            await imap._client_task
            await imap.wait_hello_from_server()
            await imap.login(self.email_server.user_name, self.email_server.password)
            await _send_imap_id(imap)
            await imap.select(_quote_mailbox(source_mailbox))

            for email_id in email_ids:
                try:
                    # Try MOVE command first (RFC 6851)
                    try:
                        result = await imap.uid("move", email_id, _quote_mailbox(destination_folder))
                        status = result[0] if isinstance(result, tuple) else result
                        if str(status).upper() == "OK":
                            moved_ids.append(email_id)
                            logger.debug(f"Moved email {email_id} to {destination_folder} using MOVE")
                            continue
                    except Exception as move_error:
                        logger.debug(f"MOVE command failed, falling back to COPY+DELETE: {move_error}")

                    # Fallback: COPY + mark as deleted
                    copy_result = await imap.uid("copy", email_id, _quote_mailbox(destination_folder))
                    copy_status = copy_result[0] if isinstance(copy_result, tuple) else copy_result
                    if str(copy_status).upper() == "OK":
                        await imap.uid("store", email_id, "+FLAGS", r"(\Deleted)")
                        moved_ids.append(email_id)
                        logger.debug(f"Moved email {email_id} to {destination_folder} using COPY+DELETE")
                    else:
                        logger.error(f"Failed to copy email {email_id}: {copy_status}")
                        failed_ids.append(email_id)
                except Exception as e:
                    logger.error(f"Failed to move email {email_id}: {e}")
                    failed_ids.append(email_id)

            # Expunge deleted messages
            if moved_ids:
                await imap.expunge()

        finally:
            try:
                await imap.logout()
            except Exception as e:
                logger.info(f"Error during logout: {e}")

        return moved_ids, failed_ids

    async def create_folder(self, folder_name: str) -> tuple[bool, str]:
        """Create a new folder. Returns (success, message)."""
        imap = self.imap_class(self.email_server.host, self.email_server.port)

        try:
            await imap._client_task
            await imap.wait_hello_from_server()
            await imap.login(self.email_server.user_name, self.email_server.password)
            await _send_imap_id(imap)

            result = await imap.create(_quote_mailbox(folder_name))
            status = result[0] if isinstance(result, tuple) else result
            if str(status).upper() == "OK":
                logger.info(f"Created folder: {folder_name}")
                return True, f"Folder '{folder_name}' created successfully"
            else:
                logger.error(f"Failed to create folder {folder_name}: {status}")
                return False, f"Failed to create folder: {status}"

        except Exception as e:
            logger.error(f"Error creating folder {folder_name}: {e}")
            return False, f"Error creating folder: {e}"
        finally:
            try:
                await imap.logout()
            except Exception as e:
                logger.info(f"Error during logout: {e}")

    async def delete_folder(self, folder_name: str) -> tuple[bool, str]:
        """Delete a folder. Returns (success, message)."""
        imap = self.imap_class(self.email_server.host, self.email_server.port)

        try:
            await imap._client_task
            await imap.wait_hello_from_server()
            await imap.login(self.email_server.user_name, self.email_server.password)
            await _send_imap_id(imap)

            result = await imap.delete(_quote_mailbox(folder_name))
            status = result[0] if isinstance(result, tuple) else result
            if str(status).upper() == "OK":
                logger.info(f"Deleted folder: {folder_name}")
                return True, f"Folder '{folder_name}' deleted successfully"
            else:
                logger.error(f"Failed to delete folder {folder_name}: {status}")
                return False, f"Failed to delete folder: {status}"

        except Exception as e:
            logger.error(f"Error deleting folder {folder_name}: {e}")
            return False, f"Error deleting folder: {e}"
        finally:
            try:
                await imap.logout()
            except Exception as e:
                logger.info(f"Error during logout: {e}")

    async def rename_folder(self, old_name: str, new_name: str) -> tuple[bool, str]:
        """Rename a folder. Returns (success, message)."""
        imap = self.imap_class(self.email_server.host, self.email_server.port)

        try:
            await imap._client_task
            await imap.wait_hello_from_server()
            await imap.login(self.email_server.user_name, self.email_server.password)
            await _send_imap_id(imap)

            result = await imap.rename(_quote_mailbox(old_name), _quote_mailbox(new_name))
            status = result[0] if isinstance(result, tuple) else result
            if str(status).upper() == "OK":
                logger.info(f"Renamed folder '{old_name}' to '{new_name}'")
                return True, f"Folder renamed from '{old_name}' to '{new_name}'"
            else:
                logger.error(f"Failed to rename folder {old_name}: {status}")
                return False, f"Failed to rename folder: {status}"

        except Exception as e:
            logger.error(f"Error renaming folder {old_name}: {e}")
            return False, f"Error renaming folder: {e}"
        finally:
            try:
                await imap.logout()
            except Exception as e:
                logger.info(f"Error during logout: {e}")

    async def list_labels(self) -> list[Label]:
        """List all labels (folders under Labels/ prefix)."""
        folders = await self.list_folders()
        labels = []
        for folder in folders:
            if folder.name.startswith("Labels/"):
                # Extract label name without prefix
                label_name = folder.name[7:]  # Remove "Labels/" prefix
                if label_name:  # Skip if just "Labels/" with no name
                    labels.append(
                        Label(
                            name=label_name,
                            full_path=folder.name,
                            delimiter=folder.delimiter,
                            flags=folder.flags,
                        )
                    )
        return labels

    async def get_email_message_id(self, email_id: str, mailbox: str = "INBOX") -> str | None:
        """Get the Message-ID header for an email."""
        imap = self.imap_class(self.email_server.host, self.email_server.port)

        try:
            await imap._client_task
            await imap.wait_hello_from_server()
            await imap.login(self.email_server.user_name, self.email_server.password)
            await _send_imap_id(imap)
            await imap.select(_quote_mailbox(mailbox))

            _, data = await imap.uid("fetch", email_id, "BODY.PEEK[HEADER.FIELDS (MESSAGE-ID)]")

            for item in data:
                if isinstance(item, bytearray):
                    header_str = bytes(item).decode("utf-8", errors="replace").strip()
                    if header_str.lower().startswith("message-id:"):
                        return header_str[11:].strip()

            return None

        except Exception as e:
            logger.error(f"Error getting Message-ID for email {email_id}: {e}")
            return None
        finally:
            try:
                await imap.logout()
            except Exception as e:
                logger.info(f"Error during logout: {e}")

    async def search_by_message_id(self, message_id: str, mailbox: str) -> str | None:
        """Search for an email by Message-ID in a specific mailbox. Returns email UID or None."""
        import re

        imap = self.imap_class(self.email_server.host, self.email_server.port)

        try:
            await imap._client_task
            await imap.wait_hello_from_server()
            await imap.login(self.email_server.user_name, self.email_server.password)
            await _send_imap_id(imap)
            await imap.select(_quote_mailbox(mailbox))

            # Search by Message-ID header (returns sequence numbers, not UIDs)
            _, data = await imap.search(f'HEADER MESSAGE-ID "{message_id}"')

            # data[0] contains space-separated sequence numbers
            if data and data[0]:
                seq_nums = data[0].decode("utf-8") if isinstance(data[0], bytes) else str(data[0])
                seq_list = seq_nums.split()
                if seq_list:
                    # Fetch the UID for this sequence number
                    _, fetch_data = await imap.fetch(seq_list[0], "(UID)")
                    for item in fetch_data:
                        if isinstance(item, bytes):
                            item_str = item.decode("utf-8", errors="replace")
                            uid_match = re.search(r"UID\s+(\d+)", item_str)
                            if uid_match:
                                return uid_match.group(1)

            return None

        except Exception as e:
            logger.debug(f"Error searching for Message-ID in {mailbox}: {e}")
            return None
        finally:
            try:
                await imap.logout()
            except Exception as e:
                logger.info(f"Error during logout: {e}")

    async def delete_from_folder(self, email_ids: list[str], folder: str) -> tuple[list[str], list[str]]:
        """Delete emails from a specific folder. Returns (deleted_ids, failed_ids)."""
        return await self.delete_emails(email_ids, folder)


class ClassicEmailHandler(EmailHandler):
    def __init__(self, email_settings: EmailSettings):
        self.email_settings = email_settings
        self.incoming_client = EmailClient(email_settings.incoming)
        self.outgoing_client = EmailClient(
            email_settings.outgoing,
            sender=f"{email_settings.full_name} <{email_settings.email_address}>",
        )
        self.save_to_sent = email_settings.save_to_sent
        self.sent_folder_name = email_settings.sent_folder_name

    async def get_emails_metadata(
        self,
        page: int = 1,
        page_size: int = 10,
        before: datetime | None = None,
        since: datetime | None = None,
        subject: str | None = None,
        from_address: str | None = None,
        to_address: str | None = None,
        order: str = "desc",
        mailbox: str = "INBOX",
    ) -> EmailMetadataPageResponse:
        emails = []
        async for email_data in self.incoming_client.get_emails_metadata_stream(
            page, page_size, before, since, subject, from_address, to_address, order, mailbox
        ):
            emails.append(EmailMetadata.from_email(email_data))
        total = await self.incoming_client.get_email_count(
            before, since, subject, from_address=from_address, to_address=to_address, mailbox=mailbox
        )
        return EmailMetadataPageResponse(
            page=page,
            page_size=page_size,
            before=before,
            since=since,
            subject=subject,
            emails=emails,
            total=total,
        )

    async def get_emails_content(self, email_ids: list[str], mailbox: str = "INBOX") -> EmailContentBatchResponse:
        """Batch retrieve email body content"""
        emails = []
        failed_ids = []

        for email_id in email_ids:
            try:
                email_data = await self.incoming_client.get_email_body_by_id(email_id, mailbox)
                if email_data:
                    emails.append(
                        EmailBodyResponse(
                            email_id=email_data["email_id"],
                            message_id=email_data.get("message_id"),
                            subject=email_data["subject"],
                            sender=email_data["from"],
                            recipients=email_data["to"],
                            date=email_data["date"],
                            body=email_data["body"],
                            attachments=email_data["attachments"],
                        )
                    )
                else:
                    failed_ids.append(email_id)
            except Exception as e:
                logger.error(f"Failed to retrieve email {email_id}: {e}")
                failed_ids.append(email_id)

        return EmailContentBatchResponse(
            emails=emails,
            requested_count=len(email_ids),
            retrieved_count=len(emails),
            failed_ids=failed_ids,
        )

    async def send_email(
        self,
        recipients: list[str],
        subject: str,
        body: str,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        html: bool = False,
        attachments: list[str] | None = None,
        in_reply_to: str | None = None,
        references: str | None = None,
    ) -> None:
        msg = await self.outgoing_client.send_email(
            recipients, subject, body, cc, bcc, html, attachments, in_reply_to, references
        )

        # Save to Sent folder if enabled
        if self.save_to_sent and msg:
            try:
                await self.outgoing_client.append_to_sent(
                    msg,
                    self.email_settings.incoming,
                    self.sent_folder_name,
                )
            except Exception as e:
                logger.error(f"Failed to save email to Sent folder: {e}", exc_info=True)

    async def delete_emails(self, email_ids: list[str], mailbox: str = "INBOX") -> tuple[list[str], list[str]]:
        """Delete emails by their UIDs. Returns (deleted_ids, failed_ids)."""
        return await self.incoming_client.delete_emails(email_ids, mailbox)

    async def download_attachment(
        self,
        email_id: str,
        attachment_name: str,
        save_path: str,
        mailbox: str = "INBOX",
    ) -> AttachmentDownloadResponse:
        """Download an email attachment and save it to the specified path.

        Args:
            email_id: The UID of the email containing the attachment.
            attachment_name: The filename of the attachment to download.
            save_path: The local path where the attachment will be saved.
            mailbox: The mailbox to search in (default: "INBOX").

        Returns:
            AttachmentDownloadResponse with download result information.
        """
        result = await self.incoming_client.download_attachment(email_id, attachment_name, save_path, mailbox)
        return AttachmentDownloadResponse(
            email_id=result["email_id"],
            attachment_name=result["attachment_name"],
            mime_type=result["mime_type"],
            size=result["size"],
            saved_path=result["saved_path"],
        )

    async def list_folders(self) -> FolderListResponse:
        """List all folders/mailboxes for the account."""
        folders = await self.incoming_client.list_folders()
        return FolderListResponse(folders=folders, total=len(folders))

    async def move_emails(
        self,
        email_ids: list[str],
        destination_folder: str,
        source_mailbox: str = "INBOX",
    ) -> EmailMoveResponse:
        """Move emails to a destination folder."""
        moved_ids, failed_ids = await self.incoming_client.move_emails(email_ids, destination_folder, source_mailbox)
        return EmailMoveResponse(
            success=len(failed_ids) == 0,
            moved_ids=moved_ids,
            failed_ids=failed_ids,
            source_mailbox=source_mailbox,
            destination_folder=destination_folder,
        )

    async def copy_emails(
        self,
        email_ids: list[str],
        destination_folder: str,
        source_mailbox: str = "INBOX",
    ) -> EmailMoveResponse:
        """Copy emails to a destination folder (preserves original)."""
        copied_ids, failed_ids = await self.incoming_client.copy_emails(email_ids, destination_folder, source_mailbox)
        return EmailMoveResponse(
            success=len(failed_ids) == 0,
            moved_ids=copied_ids,
            failed_ids=failed_ids,
            source_mailbox=source_mailbox,
            destination_folder=destination_folder,
        )

    async def create_folder(self, folder_name: str) -> FolderOperationResponse:
        """Create a new folder/mailbox."""
        success, message = await self.incoming_client.create_folder(folder_name)
        return FolderOperationResponse(
            success=success,
            folder_name=folder_name,
            message=message,
        )

    async def delete_folder(self, folder_name: str) -> FolderOperationResponse:
        """Delete a folder/mailbox."""
        success, message = await self.incoming_client.delete_folder(folder_name)
        return FolderOperationResponse(
            success=success,
            folder_name=folder_name,
            message=message,
        )

    async def rename_folder(self, old_name: str, new_name: str) -> FolderOperationResponse:
        """Rename a folder/mailbox."""
        success, message = await self.incoming_client.rename_folder(old_name, new_name)
        return FolderOperationResponse(
            success=success,
            folder_name=new_name,
            message=message,
        )

    async def list_labels(self) -> LabelListResponse:
        """List all labels (ProtonMail: folders under Labels/ prefix)."""
        labels = await self.incoming_client.list_labels()
        return LabelListResponse(labels=labels, total=len(labels))

    async def apply_label(
        self,
        email_ids: list[str],
        label_name: str,
        source_mailbox: str = "INBOX",
    ) -> EmailMoveResponse:
        """Apply a label to emails by copying to the label folder."""
        label_folder = f"Labels/{label_name}"
        copied_ids, failed_ids = await self.incoming_client.copy_emails(email_ids, label_folder, source_mailbox)
        return EmailMoveResponse(
            success=len(failed_ids) == 0,
            moved_ids=copied_ids,
            failed_ids=failed_ids,
            source_mailbox=source_mailbox,
            destination_folder=label_folder,
        )

    async def remove_label(
        self,
        email_ids: list[str],
        label_name: str,
    ) -> EmailMoveResponse:
        """Remove a label from emails by deleting from the label folder.

        This finds the emails in the label folder by their Message-ID and deletes them.
        The original emails in other folders are preserved.
        """
        label_folder = f"Labels/{label_name}"
        removed_ids = []
        failed_ids = []

        for email_id in email_ids:
            # Get the Message-ID from the source email
            # Note: We need to find this email in the label folder
            # The email_id provided is from the source mailbox, not the label folder
            # We need to search by Message-ID to find the copy in the label folder
            message_id = await self.incoming_client.get_email_message_id(email_id, "INBOX")
            if not message_id:
                logger.warning(f"Could not get Message-ID for email {email_id}")
                failed_ids.append(email_id)
                continue

            # Find the email in the label folder
            label_uid = await self.incoming_client.search_by_message_id(message_id, label_folder)
            if not label_uid:
                logger.warning(f"Email {email_id} not found in label {label_name}")
                failed_ids.append(email_id)
                continue

            # Delete from label folder
            deleted, _failed = await self.incoming_client.delete_from_folder([label_uid], label_folder)
            if deleted:
                removed_ids.append(email_id)
            else:
                failed_ids.append(email_id)

        return EmailMoveResponse(
            success=len(failed_ids) == 0,
            moved_ids=removed_ids,
            failed_ids=failed_ids,
            source_mailbox=label_folder,
            destination_folder="",
        )

    async def get_email_labels(
        self,
        email_id: str,
        source_mailbox: str = "INBOX",
    ) -> EmailLabelsResponse:
        """Get all labels applied to a specific email."""
        # Get Message-ID from the source email
        message_id = await self.incoming_client.get_email_message_id(email_id, source_mailbox)
        if not message_id:
            return EmailLabelsResponse(email_id=email_id, labels=[])

        # Get all labels
        labels = await self.incoming_client.list_labels()
        applied_labels = []

        # Search each label folder for this email
        for label in labels:
            found_uid = await self.incoming_client.search_by_message_id(message_id, label.full_path)
            if found_uid:
                applied_labels.append(label.name)

        return EmailLabelsResponse(email_id=email_id, labels=applied_labels)

    async def create_label(self, label_name: str) -> FolderOperationResponse:
        """Create a new label (creates Labels/name folder)."""
        label_folder = f"Labels/{label_name}"
        success, message = await self.incoming_client.create_folder(label_folder)
        return FolderOperationResponse(
            success=success,
            folder_name=label_name,
            message=message.replace(label_folder, label_name) if success else message,
        )

    async def delete_label(self, label_name: str) -> FolderOperationResponse:
        """Delete a label (deletes Labels/name folder)."""
        label_folder = f"Labels/{label_name}"
        success, message = await self.incoming_client.delete_folder(label_folder)
        return FolderOperationResponse(
            success=success,
            folder_name=label_name,
            message=message.replace(label_folder, label_name) if success else message,
        )
