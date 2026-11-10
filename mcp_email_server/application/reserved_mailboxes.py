r"""Mailboxes an account reserves for its owner's own use.

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
* treating ``/``, ``.``, and the server's own hierarchy delimiter (when the
  caller supplies it) as hierarchy delimiters, stripping whitespace around each
  segment, and dropping empty segments.

The provider adapter discovers that delimiter with ``LIST <entry> ""`` for each
reserved entry before checking a mutation, so a server whose delimiter is, for
example, ``|`` cannot hide a child or parent of a reserved mailbox; when it
cannot be discovered, the mutation is refused. Application-level checks run
before any IMAP session and use ``/`` and ``.`` alone. Proton's literal
``Labels/`` prefix is a naming convention applied before matching, not a
hierarchy rule, and is unaffected.

A requested mailbox matches when its segments equal a reserved entry's or start
with them (a child of a reserved mailbox is reserved too). Mailbox-shape changes
also match a mailbox whose segments are a prefix of a reserved entry's, because
deleting or renaming a parent takes the reserved mailbox with it. Over-matching
only refuses a request; it never allows one.

Flags and keywords are part of the signal too: another process may act on a
message only when it also carries ``\Draft``. So, for an account with any
reservation, adding ``\Draft`` is refused in every mailbox (servers may share one
message's flags across all the mailboxes it appears in, so the mailbox named in
the request does not show where the flag lands), and on a message whose source
mailbox is reserved only read state (``\Seen``) may change. Removing a label and
moving a message out stay allowed.
"""

from __future__ import annotations

import base64
import binascii
import unicodedata
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass

from mcp_email_server.application.limits import APPLICATION_LIMITS, validate_controlled_string

_DELIMITERS = ("/", ".")


class ReservedMailboxError(PermissionError):
    """A mutation would place a message in, or reshape, a reserved mailbox."""


@dataclass(frozen=True)
class FlagChange:
    """One requested flag (``keywords=False``) or keyword (``keywords=True``) STORE."""

    mailbox: str
    operation: str
    flags: tuple[str, ...] = ()
    keywords: bool = False


# Read state carries no approval signal, so it may change inside a reserved mailbox.
_SOURCE_SAFE_FLAGS = frozenset({r"\seen"})
_DRAFT_FLAG = r"\draft"


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


def _segments(name: str, delimiters: Iterable[str] = ()) -> tuple[str, ...]:
    text = unicodedata.normalize("NFKC", name)
    text = "".join(character for character in text if unicodedata.category(character) != "Cf")
    text = unicodedata.normalize("NFKC", text.casefold())
    # A server delimiter is normalized like the name so the two compare alike.
    extra = (unicodedata.normalize("NFKC", delimiter.casefold()) for delimiter in delimiters if delimiter)
    for delimiter in (*_DELIMITERS[1:], *extra):
        if delimiter:
            text = text.replace(delimiter, _DELIMITERS[0])
    return tuple(segment.strip() for segment in text.split(_DELIMITERS[0]) if segment.strip())


def _keys(name: str, delimiters: Iterable[str] = ()) -> frozenset[tuple[str, ...]]:
    """Every segment spelling this name could mean: as given, and decoded."""

    delimiters = tuple(delimiters)
    return frozenset({_segments(name, delimiters), _segments(_decode_modified_utf7(name), delimiters)})


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
    delimiters: Iterable[str] = (),
) -> str | None:
    """Return the configured entry ``mailbox`` falls under, or ``None``.

    ``include_ancestors`` also matches a mailbox that contains a reserved one,
    for mailbox-shape changes. ``delimiters`` are the server's own hierarchy
    delimiters, matched in addition to ``/`` and ``.``; an empty one (a flat,
    ``NIL`` hierarchy) adds nothing.
    """

    if not reserved:
        return None
    delimiters = tuple(delimiters)
    candidates = _keys(mailbox, delimiters)
    for entry in reserved:
        for key in _keys(entry, delimiters):
            if not key:
                continue
            for candidate in candidates:
                if candidate[: len(key)] == key:
                    return entry
                if include_ancestors and candidate and key[: len(candidate)] == candidate:
                    return entry
    return None


def _require_flag_change_allowed(
    account_name: str,
    reserved: Sequence[str],
    change: FlagChange,
    delimiters: tuple[str, ...],
) -> None:
    if not change.keywords and change.operation == "add" and any(f.casefold() == _DRAFT_FLAG for f in change.flags):
        raise ReservedMailboxError(
            f"Refused: account '{account_name}' has reserved mailboxes, so this server never adds \\Draft to a "
            "message there or anywhere else: a server may share one message's flags across every mailbox it is "
            "in, so the flag could mark a message in a reserved mailbox as a draft. Removing \\Draft and "
            "changing other flags outside a reserved mailbox stay allowed. Nothing was changed."
        )
    entry = reserved_entry(change.mailbox, reserved, delimiters=delimiters)
    if entry is None:
        return
    if not change.keywords and all(f.casefold() in _SOURCE_SAFE_FLAGS for f in change.flags):
        return
    raise ReservedMailboxError(
        f"Refused: '{change.mailbox}' is or is inside the reserved mailbox '{entry}' on account "
        f"'{account_name}'. Only read state (\\Seen) may change on messages there; no other flag or tag. "
        "Moving a message out, removing the label, and reading stay allowed. Nothing was changed."
    )


def require_unreserved(
    account_name: str,
    reserved: Sequence[str],
    *,
    destinations: Iterable[str] = (),
    shapes: Iterable[str] = (),
    flag_change: FlagChange | None = None,
    delimiters: Iterable[str] = (),
) -> None:
    """Refuse before any effect when a mutation touches a reserved mailbox.

    ``destinations`` are mailboxes a message would be placed in; ``shapes`` are
    mailboxes that would be created, deleted, or renamed (either name);
    ``flag_change`` is a flag or keyword STORE and the mailbox it runs in;
    ``delimiters`` are the server's discovered hierarchy delimiters.
    """

    if not reserved:
        return
    delimiters = tuple(delimiters)
    if flag_change is not None:
        _require_flag_change_allowed(account_name, reserved, flag_change, delimiters)
    checks = [(mailbox, False) for mailbox in destinations] + [(mailbox, True) for mailbox in shapes]
    for mailbox, include_ancestors in checks:
        entry = reserved_entry(mailbox, reserved, include_ancestors=include_ancestors, delimiters=delimiters)
        if entry is not None:
            raise ReservedMailboxError(
                f"Refused: '{mailbox}' is, is inside, or contains the reserved mailbox '{entry}' on account "
                f"'{account_name}'. Reserved mailboxes are kept for the account owner's own use: this "
                "server never places messages in them, and never creates, renames, or deletes them or a "
                "mailbox that contains them. Nothing was changed."
            )


def require_known_hierarchy(account_name: str, reserved: Sequence[str], discovered: Mapping[str, str] | None) -> None:
    """Refuse when the server's hierarchy delimiter is unknown for any reserved entry.

    ``discovered`` maps each reserved entry (and ``""`` for the default
    namespace) to the delimiter the server reported. Without it a child or
    parent of a reserved mailbox could hide behind a delimiter the static rules
    do not know, so the mutation is refused instead of guessed.
    """

    if not reserved:
        return
    if discovered is not None and "" in discovered and all(entry in discovered for entry in reserved):
        return
    raise ReservedMailboxError(
        f"Refused: account '{account_name}' has reserved mailboxes, and the server's hierarchy delimiter for "
        "them could not be determined, so this request cannot be checked against them. Nothing was changed; "
        "retry."
    )


def hierarchy_delimiters(discovered: Mapping[str, str]) -> tuple[str, ...]:
    """The distinct non-empty delimiters the server reported."""

    return tuple(sorted({delimiter for delimiter in discovered.values() if delimiter}))
