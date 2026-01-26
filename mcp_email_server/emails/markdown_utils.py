"""Markdown to email-safe HTML conversion utilities."""

import markdown as md

# Extensions safe for email clients (no CSS class dependencies)
EMAIL_SAFE_EXTENSIONS = ["tables", "fenced_code", "nl2br"]

# Minimal inline styles for consistent email rendering
EMAIL_BODY_STYLE = (
    "font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, "
    "'Helvetica Neue', Arial, sans-serif; "
    "line-height: 1.6; "
    "color: #333;"
)


def markdown_to_email_html(text: str, wrap_in_html: bool = True) -> str:
    """Convert markdown text to email-safe HTML.

    Args:
        text: Markdown-formatted text
        wrap_in_html: If True, wrap output in minimal HTML document with inline styles

    Returns:
        HTML string suitable for email clients
    """
    html_content = md.markdown(
        text,
        extensions=EMAIL_SAFE_EXTENSIONS,
        extension_configs={
            "fenced_code": {"lang_prefix": ""},  # No CSS classes
        },
    )

    if wrap_in_html:
        return f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="{EMAIL_BODY_STYLE}">
{html_content}
</body>
</html>"""

    return html_content
