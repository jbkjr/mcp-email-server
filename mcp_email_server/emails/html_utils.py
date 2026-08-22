"""HTML-to-text conversion for email body extraction.

Uses only the standard library (``html.parser``) so that reading mail never
depends on a third-party HTML tree builder. The converter keeps the structure a
reader needs — paragraph breaks, list markers, table cells, blockquote markers —
and preserves useful link targets instead of discarding them.
"""

from __future__ import annotations

import re
from html.parser import HTMLParser

# Hrefs are attacker-controlled. Compare schemes only after removing the ASCII
# whitespace and C0 controls that HTML permits inside an attribute value, so
# "java&#10;script:" cannot slip past the scheme check.
_HREF_CONTROL_CHARACTERS = re.compile(r"[\x00-\x20]+")
_UNSAFE_HREF_SCHEMES = ("mailto:", "javascript:")

_HEADING_TAGS = frozenset({"h1", "h2", "h3", "h4", "h5", "h6"})

# Tags whose open or close emits fixed text and needs no other state.
_START_TAG_TEXT = {
    "br": "\n",
    "hr": "\n---\n",
    "blockquote": "\n> ",
    "pre": "\n",
    "ul": "\n",
    "table": "\n",
}
_END_TAG_TEXT = {
    "blockquote": "\n",
    "pre": "\n",
}


def _is_useful_href(href: str) -> bool:
    """Return whether a link target is worth surfacing in extracted text."""

    if not href or href.startswith("#"):
        return False
    normalized = _HREF_CONTROL_CHARACTERS.sub("", href).lower()
    return not normalized.startswith(_UNSAFE_HREF_SCHEMES)


class _HTMLToTextParser(HTMLParser):
    """HTMLParser subclass that converts HTML to readable plain text."""

    # Tags whose content should be suppressed entirely
    SKIP_TAGS = frozenset({"style", "script", "head"})

    # Block-level tags that produce paragraph breaks
    BLOCK_TAGS = frozenset({"p", "div", "article", "section", "main", "header", "footer", "nav", "aside"})

    def __init__(self) -> None:
        super().__init__()
        self._result: list[str] = []
        self._skip_depth = 0  # Nesting depth inside SKIP_TAGS
        self._ol_counter: list[int] = []  # Stack of ordered-list counters
        self._table_row: list[str] = []
        self._current_link_href = ""
        self._current_link_start = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in self.SKIP_TAGS:
            self._skip_depth += 1
        elif self._skip_depth == 0:
            self._open_tag(tag, attrs)

    def handle_endtag(self, tag: str) -> None:
        if tag in self.SKIP_TAGS:
            self._skip_depth = max(0, self._skip_depth - 1)
        elif self._skip_depth == 0:
            self._close_tag(tag)

    def handle_data(self, data: str) -> None:
        # ``convert_charrefs`` defaults to True, so character references outside
        # script/style arrive here already decoded; references inside a skipped
        # element are dropped with the rest of that element's content.
        if self._skip_depth > 0:
            return
        self._result.append(data)

    def _open_tag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in self.BLOCK_TAGS or tag in _HEADING_TAGS:
            self._result.append("\n\n")
        elif tag in _START_TAG_TEXT:
            self._result.append(_START_TAG_TEXT[tag])
        elif tag == "ol":
            self._ol_counter.append(0)
            self._result.append("\n")
        elif tag == "li":
            self._open_list_item()
        elif tag == "a":
            self._open_link(attrs)
        elif tag == "tr":
            self._table_row = []
        elif tag == "img":
            alt = dict(attrs).get("alt") or ""
            if alt:
                self._result.append(f"[{alt}]")

    def _close_tag(self, tag: str) -> None:
        if tag in self.BLOCK_TAGS or tag in _HEADING_TAGS:
            self._result.append("\n\n")
        elif tag in _END_TAG_TEXT:
            self._result.append(_END_TAG_TEXT[tag])
        elif tag == "ol":
            if self._ol_counter:
                self._ol_counter.pop()
        elif tag == "a":
            self._close_link()
        elif tag == "tr":
            self._close_table_row()
        elif tag in ("td", "th"):
            self._table_row.append(self._extract_last_cell())

    def _open_list_item(self) -> None:
        """Start a list item, numbering it inside the innermost ordered list."""

        if self._ol_counter:
            self._ol_counter[-1] += 1
            self._result.append(f"\n{self._ol_counter[-1]}. ")
        else:
            self._result.append("\n- ")

    def _open_link(self, attrs: list[tuple[str, str | None]]) -> None:
        """Remember the target and where its text begins, for ``_close_link``."""

        self._current_link_href = dict(attrs).get("href") or ""
        self._current_link_start = len(self._result)

    def _close_link(self) -> None:
        """Append the link target when it adds information the link text does not."""

        href = self._current_link_href
        self._current_link_href = ""
        if not _is_useful_href(href):
            return
        link_text = "".join(self._result[self._current_link_start :]).strip()
        if not link_text:
            self._result.append(href)
        elif link_text != href:
            self._result.append(f" ({href})")

    def _close_table_row(self) -> None:
        if self._table_row:
            self._result.append("\t".join(self._table_row) + "\n")
            self._table_row = []

    def _extract_last_cell(self) -> str:
        """Extract text for the most recent table cell."""
        # Walk backwards through result to find cell content
        # (content added since the last td/th start or tr start)
        text_parts: list[str] = []
        for index in range(len(self._result) - 1, -1, -1):
            part = self._result[index]
            if part == "\n" or part.endswith("\n"):
                break
            text_parts.insert(0, part)
        # Remove collected parts from result to avoid duplication
        if text_parts:
            self._result = self._result[: len(self._result) - len(text_parts)]
        return "".join(text_parts).strip()

    def get_text(self) -> str:
        return "".join(self._result)


def html_to_text(html: str) -> str:
    """Convert HTML to readable plain text.

    Args:
        html: HTML string to convert.

    Returns:
        Plain text representation of the HTML content.
    """
    if not html:
        return ""

    parser = _HTMLToTextParser()
    parser.feed(html)
    text = parser.get_text()

    # Collapse runs of 3+ newlines into 2
    text = re.sub(r"\n{3,}", "\n\n", text)

    return text.strip()
