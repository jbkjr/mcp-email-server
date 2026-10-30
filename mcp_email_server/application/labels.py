"""Label naming and LIST roles shared by label tools.

ProtonMail (and ProtonMail Bridge) exposes labels as ordinary IMAP mailboxes
under a literal ``Labels/`` prefix. The prefix is part of the mailbox *name*, not
a server hierarchy path, so it is spelled with ``/`` regardless of the hierarchy
delimiter the server advertises. Keeping the convention in one module lets the
read tools, the mutation commands, and the label-write tools agree on it without
importing each other.
"""

from __future__ import annotations

from mcp_email_server.application.limits import APPLICATION_LIMITS, validate_controlled_string
from mcp_email_server.emails.models import MailboxInfo

LABEL_MAILBOX_PREFIX = "Labels/"
LABEL_LIST_PATTERN = f"{LABEL_MAILBOX_PREFIX}*"
MAXIMUM_LABEL_NAME_BYTES = APPLICATION_LIMITS.mailbox_bytes - len(LABEL_MAILBOX_PREFIX.encode("utf-8"))


def validate_label_name(
    label_name: object, *, field_name: str = "label_name", gateway_provider: str | None = None
) -> str:
    """Validate one label name so its mailbox stays inside the mailbox bound."""

    name = validate_controlled_string(
        label_name,
        field_name=field_name,
        maximum_bytes=APPLICATION_LIMITS.mailbox_bytes
        if gateway_provider == "gmail_imap"
        else MAXIMUM_LABEL_NAME_BYTES,
    )
    if gateway_provider != "gmail_imap" and name.startswith(LABEL_MAILBOX_PREFIX):
        raise ValueError(f"{field_name} must not repeat the '{LABEL_MAILBOX_PREFIX}' prefix")
    return name


def label_mailbox(label_name: str, *, gateway_provider: str | None = None) -> str:
    """Return an exact Gmail label or the historical Proton label mailbox."""

    if gateway_provider == "gmail_imap":
        return validate_controlled_string(
            label_name, field_name="label_name", maximum_bytes=APPLICATION_LIMITS.mailbox_bytes
        )
    return f"{LABEL_MAILBOX_PREFIX}{validate_label_name(label_name)}"


def label_name_from_mailbox(mailbox: str) -> str | None:
    """Return the label a mailbox represents, or ``None`` when it is not one.

    The bare ``Labels/`` container is a grouping mailbox rather than a label, so
    it yields ``None`` alongside every mailbox outside the prefix.
    """

    if not mailbox.startswith(LABEL_MAILBOX_PREFIX):
        return None
    return mailbox[len(LABEL_MAILBOX_PREFIX) :] or None


_SPECIAL = {r"\All", r"\Drafts", r"\Trash", r"\Junk", r"\Sent", r"\Flagged", r"\Important", r"\Archive"}


def mailbox_role(name: str, mailboxes: list[MailboxInfo]) -> str:
    """Resolve unique special-use flags; ordinary labels retain exact spelling."""
    matches = [m for m in mailboxes if m.name == name or name.upper() == m.name.upper() == "INBOX"]
    if len(matches) != 1:
        return "unknown"
    flags = {f.upper() for f in matches[0].flags}
    if flags & {r"\NOSELECT", r"\NONEXISTENT"}:
        return "unknown"
    roles = flags & {f.upper() for f in _SPECIAL}
    if roles:
        if len(roles) != 1:
            return "unknown"
        flag = next(iter(roles))
        if sum(flag in {f.upper() for f in m.flags} for m in mailboxes) != 1:
            return "unknown"
        return flag[1:].lower()
    if name.upper() == "INBOX":
        return "inbox"
    if name.startswith(("[Gmail]", "[GoogleMail]", "[Google Mail]", "\\")) or flags - {
        r"\HASNOCHILDREN",
        r"\HASCHILDREN",
        r"\MARKED",
        r"\UNMARKED",
    }:
        return "unknown"
    return "label"


def special_mailbox(role: str, mailboxes: list[MailboxInfo]) -> str:
    names = [m.name for m in mailboxes if mailbox_role(m.name, mailboxes) == role]
    if len(names) != 1:
        raise ValueError(f"Gmail requires one LIST-discovered {role} mailbox")
    return names[0]
