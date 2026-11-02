"""Small, explicit Gmail gateway policy helpers (never hostname detection)."""

import re

from mcp_email_server.application.labels import mailbox_role, special_mailbox  # noqa: F401


def label_present(labels: list[str], label: str) -> bool:
    """Compare system labels without case and user labels exactly."""
    return any(item.casefold() == label.casefold() if label.startswith("\\") else item == label for item in labels)


def validate_uids(uids: list[str]) -> None:
    if not 1 <= len(uids) <= 25 or len(set(uids)) != len(uids):
        raise ValueError("Gmail requires at most 25 explicit unique UIDs")
    if any(re.fullmatch(r"[1-9][0-9]{0,9}", uid) is None or int(uid) >= 2**32 for uid in uids):
        raise ValueError("Gmail requires explicit numeric UIDs")


def parse_metadata(data: list) -> dict[str, dict]:  # noqa: C901 - strict metadata-only grammar
    """Parse metadata-only FETCH records; never scan body/header literals."""
    from mcp_email_server.emails.classic import decode_mailbox_name

    result = {}
    grammar = re.compile(r'"(?:[^"\\\r\n]|\\["\\])*"|[()]|[^\s()"\r\n]+')
    for item in data:
        if not isinstance(item, bytes) or re.match(rb"\d+ FETCH \(", item, re.I) is None:
            continue
        text = item.decode("ascii")
        record = re.fullmatch(r"\d+ FETCH \((.*)\)", text, re.I)
        if record is None:
            raise ValueError("Incomplete Gmail metadata")
        tokens = []
        end = 0
        for match in grammar.finditer(record[1]):
            if record[1][end : match.start()].strip():
                raise ValueError("Invalid Gmail metadata token")
            tokens.append(match[0])
            end = match.end()
        if record[1][end:].strip():
            raise ValueError("Invalid Gmail metadata suffix")
        fields: dict[str, str | list[str]] = {}
        index = 0
        while index < len(tokens):
            key = tokens[index].upper()
            index += 1
            if key in fields or index == len(tokens):
                raise ValueError("Duplicate or incomplete Gmail metadata")
            if key in {"UID", "X-GM-MSGID"}:
                fields[key] = tokens[index]
                index += 1
            elif key == "X-GM-LABELS" and tokens[index] == "(":
                index += 1
                labels = []
                while index < len(tokens) and tokens[index] != ")":
                    if tokens[index] == "(":
                        raise ValueError("Invalid Gmail label list")
                    labels.append(tokens[index])
                    index += 1
                if index == len(tokens):
                    raise ValueError("Incomplete Gmail label list")
                fields[key] = labels
                index += 1
            else:
                raise ValueError("Unexpected Gmail metadata attribute")
        uid, msgid, labels = fields.get("UID"), fields.get("X-GM-MSGID"), fields.get("X-GM-LABELS")
        if (
            not isinstance(uid, str)
            or re.fullmatch(r"[1-9][0-9]{0,9}", uid) is None
            or int(uid) >= 2**32
            or not isinstance(msgid, str)
            or re.fullmatch(r"[1-9][0-9]{0,19}", msgid) is None
            or int(msgid) >= 2**64
            or not isinstance(labels, list)
            or uid in result
        ):
            raise ValueError("Invalid Gmail UID/identity metadata")
        parsed = []
        for value in labels:
            if value.startswith('"'):
                value = re.sub(r'\\(["\\])', r"\1", value[1:-1])
            parsed.append(decode_mailbox_name(value))
        result[uid] = {"gmail_message_id": msgid, "gmail_labels": parsed}
    return result
