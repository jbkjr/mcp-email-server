"""HTML-to-text converter for email body parsing.

Uses only stdlib (html.parser) to convert HTML emails to readable plain text.
"""

import re
from html.parser import HTMLParser


class _HTMLToTextParser(HTMLParser):
    """HTMLParser subclass that converts HTML to readable plain text."""

    # Tags whose content should be suppressed entirely
    SKIP_TAGS = {"style", "script", "head"}

    # Block-level tags that produce paragraph breaks
    BLOCK_TAGS = {"p", "div", "article", "section", "main", "header", "footer", "nav", "aside"}

    def __init__(self):
        super().__init__()
        self._result: list[str] = []
        self._skip_depth = 0  # Nesting depth inside SKIP_TAGS
        self._in_pre = False
        self._ol_counter: list[int] = []  # Stack of ordered-list counters
        self._in_ol = False
        self._in_table = False
        self._table_row: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in self.SKIP_TAGS:
            self._skip_depth += 1
            return
        if self._skip_depth > 0:
            return

        if tag == "br":
            self._result.append("\n")
        elif tag == "hr":
            self._result.append("\n---\n")
        elif tag in self.BLOCK_TAGS:
            self._result.append("\n\n")
        elif tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            self._result.append("\n\n")
        elif tag == "blockquote":
            self._result.append("\n> ")
        elif tag == "pre":
            self._in_pre = True
            self._result.append("\n")
        elif tag == "ul":
            self._result.append("\n")
        elif tag == "ol":
            self._ol_counter.append(0)
            self._in_ol = True
            self._result.append("\n")
        elif tag == "li":
            if self._ol_counter:
                self._ol_counter[-1] += 1
                self._result.append(f"\n{self._ol_counter[-1]}. ")
            else:
                self._result.append("\n- ")
        elif tag == "a":
            # Store href for later use in handle_endtag
            href = dict(attrs).get("href", "")
            self._current_link_href = href
        elif tag == "table":
            self._in_table = True
            self._result.append("\n")
        elif tag == "tr":
            self._table_row = []
        elif tag in ("td", "th"):
            pass  # Content will be captured in handle_data
        elif tag == "img":
            alt = dict(attrs).get("alt", "")
            if alt:
                self._result.append(f"[{alt}]")

    def handle_endtag(self, tag: str) -> None:
        if tag in self.SKIP_TAGS:
            self._skip_depth = max(0, self._skip_depth - 1)
            return
        if self._skip_depth > 0:
            return

        if tag in self.BLOCK_TAGS:
            self._result.append("\n\n")
        elif tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            self._result.append("\n\n")
        elif tag == "blockquote":
            self._result.append("\n")
        elif tag == "pre":
            self._in_pre = False
            self._result.append("\n")
        elif tag == "ol":
            if self._ol_counter:
                self._ol_counter.pop()
            self._in_ol = bool(self._ol_counter)
        elif tag == "a":
            href = getattr(self, "_current_link_href", "")
            if href and not href.startswith(("#", "mailto:", "javascript:")):
                self._result.append(f" ({href})")
            self._current_link_href = ""
        elif tag == "tr":
            if self._table_row:
                self._result.append("\t".join(self._table_row) + "\n")
                self._table_row = []
        elif tag in ("td", "th"):
            # Capture the last text segment as a cell value
            # Get text since last cell boundary
            cell_text = self._extract_last_cell()
            self._table_row.append(cell_text)
        elif tag == "table":
            self._in_table = False

    def handle_data(self, data: str) -> None:
        if self._skip_depth > 0:
            return
        if self._in_pre:
            self._result.append(data)
        else:
            self._result.append(data)

    def handle_entityref(self, name: str) -> None:
        if self._skip_depth > 0:
            return
        entity_map = {"nbsp": " ", "amp": "&", "lt": "<", "gt": ">", "quot": '"'}
        self._result.append(entity_map.get(name, f"&{name};"))

    def handle_charref(self, name: str) -> None:
        if self._skip_depth > 0:
            return
        try:
            if name.startswith("x"):
                char = chr(int(name[1:], 16))
            else:
                char = chr(int(name))
            self._result.append(char)
        except (ValueError, OverflowError):
            self._result.append(f"&#{name};")

    def _extract_last_cell(self) -> str:
        """Extract text for the most recent table cell."""
        # Walk backwards through result to find cell content
        # (content added since the last td/th start or tr start)
        text_parts = []
        for i in range(len(self._result) - 1, -1, -1):
            part = self._result[i]
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

    # Strip leading/trailing whitespace
    text = text.strip()

    return text
