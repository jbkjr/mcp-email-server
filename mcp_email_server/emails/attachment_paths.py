"""Destination resolution for downloaded email attachments.

An explicit ``save_path`` is honored exactly as given (expanded and made
absolute, nothing else). When no path is supplied, the destination is resolved
under a per-user ``Downloads/mcp-email-server`` directory using a sanitized,
bounded basename plus a cryptographically random suffix, so a provider-supplied
attachment name can never steer the write outside that directory.

Ported from upstream ai-zerolab/mcp-email-server PR #224. The Windows Known
Folder registry lookup from that PR is intentionally omitted: this fork does not
carry upstream's Windows filesystem-security support.
"""

import os
import re
import secrets
import stat
import unicodedata
from pathlib import Path

# Private subdirectory created inside the user's Downloads directory.
DEFAULT_ATTACHMENT_DIRECTORY = "mcp-email-server"

# Characters that must never appear in a generated basename. Beyond the POSIX
# path separator, this also covers the Windows-reserved set and C0/C1-adjacent
# controls, so a name generated here is safe on any host that reads the file.
_INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f\x7f]')

# Byte budgets keeping the generated name well inside the common 255-byte
# per-component filesystem limit once the random suffix is appended.
_MAX_STEM_BYTES = 160
_MAX_SUFFIX_BYTES = 24


def _truncate_utf8(value: str, maximum_bytes: int) -> str:
    """Truncate ``value`` to at most ``maximum_bytes`` UTF-8 bytes."""
    encoded = value.encode("utf-8")
    if len(encoded) <= maximum_bytes:
        return value
    return encoded[:maximum_bytes].decode("utf-8", errors="ignore")


def safe_default_filename(attachment_name: str) -> str:
    """Build a bounded, sanitized, randomized basename from an attachment name.

    Path separators, traversal segments, device-ish syntax, and Unicode control
    or format characters (including RTL-override extension spoofing) are all
    reduced to ``_``. A random suffix is appended so two downloads of the same
    attachment never collide or overwrite each other.
    """
    normalized = unicodedata.normalize("NFC", attachment_name)
    without_controls = "".join(
        "_" if unicodedata.category(character) in {"Cc", "Cf", "Cs"} else character for character in normalized
    )
    cleaned = _INVALID_FILENAME_CHARS.sub("_", without_controls).strip(" .")
    if not cleaned:
        cleaned = "attachment"
    suffix = Path(cleaned).suffix
    if len(suffix.encode("utf-8")) > _MAX_SUFFIX_BYTES:
        suffix = ""
    stem = cleaned[: -len(suffix)] if suffix else cleaned
    stem = _truncate_utf8(stem.rstrip(" ."), _MAX_STEM_BYTES).rstrip(" .") or "attachment"
    return f"{stem}-{secrets.token_hex(16)}{suffix}"


def default_downloads_directory() -> Path:
    """Return the current user's Downloads directory."""
    return Path.home() / "Downloads"


def resolve_attachment_destination(save_path: str | None, attachment_name: str) -> tuple[Path, bool]:
    """Resolve the absolute destination for an attachment download.

    Args:
        save_path: Explicit destination path, or ``None`` to use the default area.
        attachment_name: The attachment filename, used only when defaulting.

    Returns:
        ``(destination, used_default)`` where ``destination`` is absolute and
        ``used_default`` reports whether the default download area was used.

    Raises:
        PermissionError: If the destination cannot be resolved at all.
    """
    try:
        if save_path is not None:
            return Path(os.path.abspath(Path(save_path).expanduser())), False
        requested = (
            default_downloads_directory() / DEFAULT_ATTACHMENT_DIRECTORY / safe_default_filename(attachment_name)
        )
        return Path(os.path.abspath(requested)), True
    except (OSError, RuntimeError, ValueError) as exc:
        msg = "Attachment destination could not be resolved"
        raise PermissionError(msg) from exc


def prepare_private_directory(directory: Path) -> None:
    """Create ``directory`` owner-only, validating it when it already exists.

    Only the application subdirectory is treated as the sensitive parent; the
    Downloads directory above it keeps whatever permissions the user gave it.
    """
    try:
        directory.parent.mkdir(parents=True, exist_ok=True)
        directory.mkdir(mode=0o700, exist_ok=True)
    except OSError as exc:
        msg = "Attachment destination directory could not be created"
        raise PermissionError(msg) from exc

    if os.name != "posix":
        return

    try:
        # lstat, not stat: a symlink planted at the application directory must be
        # rejected rather than followed.
        metadata = directory.lstat()
    except OSError as exc:
        msg = "Attachment destination directory is unsafe"
        raise PermissionError(msg) from exc

    if not stat.S_ISDIR(metadata.st_mode):
        msg = "Attachment destination directory is unsafe"
        raise PermissionError(msg)
    if stat.S_IMODE(metadata.st_mode) & 0o022 and not metadata.st_mode & stat.S_ISVTX:
        msg = "Attachment destination directory permissions are unsafe"
        raise PermissionError(msg)
    if metadata.st_uid not in {0, os.geteuid()}:
        msg = "Attachment destination directory ownership is unsafe"
        raise PermissionError(msg)


def write_private_file(destination: Path, payload: bytes) -> None:
    """Write ``payload`` to a newly created owner-only file.

    ``O_EXCL`` means an attacker-planted file (or symlink) at the randomized
    destination fails the write instead of being followed or truncated.
    """
    flags = os.O_WRONLY | os.O_CREAT | os.O_EXCL | getattr(os, "O_NOFOLLOW", 0)
    try:
        descriptor = os.open(destination, flags, 0o600)
    except OSError as exc:
        msg = "Attachment destination is unsafe"
        raise PermissionError(msg) from exc
    with os.fdopen(descriptor, "wb") as handle:
        handle.write(payload)
