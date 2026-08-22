"""The label naming convention shared by every label tool.

ProtonMail (and ProtonMail Bridge) exposes labels as ordinary IMAP mailboxes
under a literal ``Labels/`` prefix. The prefix is part of the mailbox *name*, not
a server hierarchy path, so it is spelled with ``/`` regardless of the hierarchy
delimiter the server advertises. Keeping the convention in one module lets the
read tools, the mutation commands, and the label-write tools agree on it without
importing each other.
"""

from __future__ import annotations

from mcp_email_server.application.limits import APPLICATION_LIMITS, validate_controlled_string

LABEL_MAILBOX_PREFIX = "Labels/"
LABEL_LIST_PATTERN = f"{LABEL_MAILBOX_PREFIX}*"
MAXIMUM_LABEL_NAME_BYTES = APPLICATION_LIMITS.mailbox_bytes - len(LABEL_MAILBOX_PREFIX.encode("utf-8"))


def validate_label_name(label_name: object, *, field_name: str = "label_name") -> str:
    """Validate one label name so its mailbox stays inside the mailbox bound."""

    name = validate_controlled_string(
        label_name,
        field_name=field_name,
        maximum_bytes=MAXIMUM_LABEL_NAME_BYTES,
    )
    if name.startswith(LABEL_MAILBOX_PREFIX):
        raise ValueError(f"{field_name} must not repeat the '{LABEL_MAILBOX_PREFIX}' prefix")
    return name


def label_mailbox(label_name: str) -> str:
    """Return the mailbox that stores one label."""

    return f"{LABEL_MAILBOX_PREFIX}{validate_label_name(label_name)}"


def label_name_from_mailbox(mailbox: str) -> str | None:
    """Return the label a mailbox represents, or ``None`` when it is not one.

    The bare ``Labels/`` container is a grouping mailbox rather than a label, so
    it yields ``None`` alongside every mailbox outside the prefix.
    """

    if not mailbox.startswith(LABEL_MAILBOX_PREFIX):
        return None
    return mailbox[len(LABEL_MAILBOX_PREFIX) :] or None
