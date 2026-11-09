"""Mailboxes an account reserves for its owner's own use.

Some owners use a mailbox as a signal: applying a label, for example, can be
their approval for another process to act on the message. Such a mailbox is
listed in the account's ``reserved_mailboxes``. This server then refuses every
mutation that would place a message in a reserved mailbox (COPY, MOVE, APPEND,
label application) and every mailbox-shape change (CREATE, DELETE, RENAME) of a
reserved mailbox or of a mailbox above or below one. Taking a message *out* of a
reserved mailbox and reading one stay allowed, because neither can produce the
owner's signal.

Matching is deliberately broader than the server's own name resolution, so a
spelling the server would treat as the same mailbox cannot slip through. Both
the configured entry and the requested mailbox are compared as a sequence of
hierarchy segments after:

* decoding RFC 3501 modified UTF-7 (the raw spelling is compared as well);
* NFKC normalization, removal of invisible format characters, and case folding;
* treating both ``/`` and ``.`` as hierarchy delimiters, stripping whitespace
  around each segment, and dropping empty segments.

A requested mailbox matches when its segments equal a reserved entry's or start
with them (a child of a reserved mailbox is reserved too). Mailbox-shape changes
also match a mailbox whose segments are a prefix of a reserved entry's, because
deleting or renaming a parent takes the reserved mailbox with it. Over-matching
only refuses a request; it never allows one.
"""

from __future__ import annotations

import base64
import binascii
import unicodedata
from collections.abc import Iterable, Sequence

from mcp_email_server.application.limits import APPLICATION_LIMITS, validate_controlled_string

_DELIMITERS = ("/", ".")


class ReservedMailboxError(PermissionError):
    """A mutation would place a message in, or reshape, a reserved mailbox."""


def _decode_modified_utf7(name: str) -> str:
    """Leniently decode RFC 3501 modified UTF-7; undecodable runs stay literal."""

    result: list[str] = []
    index = 0
    while index < len(name):
        start = name.find("&", index)
        if start == -1:
            result.append(name[index:])
            break
        result.append(name[index:start])
        end = name.find("-", start + 1)
        if end == -1:
            result.append(name[start:])
            break
        encoded = name[start + 1 : end]
        if not encoded:
            result.append("&")
        else:
            padded = encoded.replace(",", "/") + "=" * (-len(encoded) % 4)
            try:
                result.append(base64.b64decode(padded, validate=True).decode("utf-16-be"))
            except (binascii.Error, UnicodeDecodeError):
                result.append(name[start : end + 1])
        index = end + 1
    return "".join(result)


def _segments(name: str) -> tuple[str, ...]:
    text = unicodedata.normalize("NFKC", name)
    text = "".join(character for character in text if unicodedata.category(character) != "Cf")
    text = unicodedata.normalize("NFKC", text.casefold())
    for delimiter in _DELIMITERS[1:]:
        text = text.replace(delimiter, _DELIMITERS[0])
    return tuple(segment.strip() for segment in text.split(_DELIMITERS[0]) if segment.strip())


def _keys(name: str) -> frozenset[tuple[str, ...]]:
    """Every segment spelling this name could mean: as given, and decoded."""

    return frozenset({_segments(name), _segments(_decode_modified_utf7(name))})


def normalize_reserved_mailboxes(values: Iterable[object]) -> list[str]:
    """Validate configured entries, strip surrounding whitespace, de-duplicate in order."""

    entries: list[str] = []
    for value in values:
        if not isinstance(value, str):
            raise ValueError("reserved_mailboxes entries must be strings")  # noqa: TRY004 - validation contract
        entry = validate_controlled_string(
            value,
            field_name="reserved_mailboxes entry",
            maximum_bytes=APPLICATION_LIMITS.mailbox_bytes,
        ).strip()
        if any(not key for key in _keys(entry)):
            raise ValueError(f"reserved_mailboxes entry {entry!r} does not name a mailbox")
        if entry not in entries:
            entries.append(entry)
    if len(entries) > APPLICATION_LIMITS.policy_entries:
        raise ValueError(f"reserved_mailboxes exceeds the limit of {APPLICATION_LIMITS.policy_entries} entries")
    return entries


def reserved_entry(
    mailbox: str,
    reserved: Sequence[str],
    *,
    include_ancestors: bool = False,
) -> str | None:
    """Return the configured entry ``mailbox`` falls under, or ``None``.

    ``include_ancestors`` also matches a mailbox that contains a reserved one,
    for mailbox-shape changes.
    """

    if not reserved:
        return None
    candidates = _keys(mailbox)
    for entry in reserved:
        for key in _keys(entry):
            if not key:
                continue
            for candidate in candidates:
                if candidate[: len(key)] == key:
                    return entry
                if include_ancestors and candidate and key[: len(candidate)] == candidate:
                    return entry
    return None


def require_unreserved(
    account_name: str,
    reserved: Sequence[str],
    *,
    destinations: Iterable[str] = (),
    shapes: Iterable[str] = (),
) -> None:
    """Refuse before any effect when a destination or shape change touches a reserved mailbox.

    ``destinations`` are mailboxes a message would be placed in; ``shapes`` are
    mailboxes that would be created, deleted, or renamed (either name).
    """

    if not reserved:
        return
    checks = [(mailbox, False) for mailbox in destinations] + [(mailbox, True) for mailbox in shapes]
    for mailbox, include_ancestors in checks:
        entry = reserved_entry(mailbox, reserved, include_ancestors=include_ancestors)
        if entry is not None:
            raise ReservedMailboxError(
                f"Refused: '{mailbox}' is, is inside, or contains the reserved mailbox '{entry}' on account "
                f"'{account_name}'. Reserved mailboxes are kept for the account owner's own use: this "
                "server never places messages in them, and never creates, renames, or deletes them or a "
                "mailbox that contains them. Nothing was changed."
            )
