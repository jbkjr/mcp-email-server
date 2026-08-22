"""Tests for the stdlib html_to_text converter used for HTML body extraction."""

from mcp_email_server.emails.html_utils import html_to_text


class TestHtmlToTextBasics:
    def test_empty_input(self):
        assert html_to_text("") == ""

    def test_plain_text_passthrough(self):
        assert html_to_text("Hello world") == "Hello world"

    def test_single_paragraph(self):
        result = html_to_text("<p>Hello world</p>")
        assert "Hello world" in result

    def test_multiple_paragraphs(self):
        result = html_to_text("<p>First paragraph</p><p>Second paragraph</p>")
        assert "First paragraph" in result
        assert "Second paragraph" in result
        # Paragraphs should be separated by blank lines
        assert "\n\n" in result

    def test_br_tags(self):
        result = html_to_text("Line one<br>Line two<br/>Line three")
        assert "Line one\nLine two\nLine three" in result

    def test_hr_tag(self):
        result = html_to_text("Above<hr>Below")
        assert "---" in result


class TestHtmlToTextHeadings:
    def test_heading_tags(self):
        result = html_to_text("<h1>Title</h1><p>Content</p>")
        assert "Title" in result
        assert "Content" in result

    def test_multiple_heading_levels(self):
        result = html_to_text("<h1>H1</h1><h2>H2</h2><h3>H3</h3>")
        assert "H1" in result
        assert "H2" in result
        assert "H3" in result


class TestHtmlToTextLists:
    def test_unordered_list(self):
        result = html_to_text("<ul><li>Apple</li><li>Banana</li><li>Cherry</li></ul>")
        assert "- Apple" in result
        assert "- Banana" in result
        assert "- Cherry" in result

    def test_ordered_list(self):
        result = html_to_text("<ol><li>First</li><li>Second</li><li>Third</li></ol>")
        assert "1. First" in result
        assert "2. Second" in result
        assert "3. Third" in result

    def test_nested_ordered_lists_restart_and_resume_numbering(self):
        result = html_to_text("<ol><li>One<ol><li>Inner</li></ol></li><li>Two</li></ol>")
        assert "1. One" in result
        assert "1. Inner" in result
        assert "2. Two" in result


class TestHtmlToTextLinks:
    def test_link_shows_url(self):
        result = html_to_text('<a href="https://example.com">Click here</a>')
        assert "Click here" in result
        assert "(https://example.com)" in result

    def test_link_skips_anchor_href(self):
        result = html_to_text('<a href="#section">Jump</a>')
        assert "Jump" in result
        assert "(#section)" not in result

    def test_link_skips_mailto_href(self):
        result = html_to_text('<a href="mailto:test@example.com">Email</a>')
        assert "Email" in result
        assert "(mailto:" not in result

    def test_link_skips_javascript_href(self):
        result = html_to_text("<a href=\"javascript:alert('x')\">Click</a>")
        assert "Click" in result
        assert "javascript:" not in result
        assert "alert" not in result

    def test_link_skips_control_character_obfuscated_scheme(self):
        """Control characters inside the href cannot smuggle an unsafe scheme through."""
        result = html_to_text("<a href=\"java&#10;script:alert('x')\">obfuscated</a>")
        assert "obfuscated" in result
        assert "script:" not in result
        assert "alert" not in result

    def test_link_skips_empty_href(self):
        result = html_to_text('<a href="">empty</a>')
        assert result == "empty"

    def test_link_without_href_attribute(self):
        result = html_to_text("<a>anchor</a>")
        assert result == "anchor"

    def test_link_text_equal_to_url_is_not_repeated(self):
        result = html_to_text('<a href="https://example.com/help">https://example.com/help</a>')
        assert result == "https://example.com/help"

    def test_textless_link_still_surfaces_the_url(self):
        result = html_to_text('<a href="https://example.com/textless"></a>')
        assert result == "https://example.com/textless"


class TestHtmlToTextStripping:
    def test_style_content_stripped(self):
        result = html_to_text("<style>.foo { color: red; }</style><p>Visible text</p>")
        assert "color: red" not in result
        assert "Visible text" in result

    def test_script_content_stripped(self):
        result = html_to_text("<script>alert('xss')</script><p>Safe text</p>")
        assert "alert" not in result
        assert "Safe text" in result

    def test_head_content_stripped(self):
        result = html_to_text("<head><title>Page Title</title></head><body><p>Body text</p></body>")
        assert "Page Title" not in result
        assert "Body text" in result

    def test_character_references_inside_skipped_tags_are_dropped(self):
        result = html_to_text("<script>if (a &lt; b) alert(&#65;)</script><p>Safe</p>")
        assert result == "Safe"


class TestHtmlToTextFullDocument:
    def test_full_html_document(self):
        html = """<!DOCTYPE html>
<html>
<head>
    <title>Test Email</title>
    <style>body { font-family: Arial; }</style>
</head>
<body>
    <h1>Welcome</h1>
    <p>This is a <strong>test</strong> email.</p>
    <p>It has <a href="https://example.com">a link</a> and multiple paragraphs.</p>
    <ul>
        <li>Item one</li>
        <li>Item two</li>
    </ul>
</body>
</html>"""
        result = html_to_text(html)

        assert "Welcome" in result
        assert "test" in result
        assert "email" in result
        assert "(https://example.com)" in result
        assert "- Item one" in result
        assert "- Item two" in result
        # Style and head content should be stripped
        assert "font-family" not in result
        assert "Test Email" not in result


class TestHtmlToTextTables:
    def test_table_extracts_text(self):
        html = """<table>
<tr><th>Name</th><th>Age</th></tr>
<tr><td>Alice</td><td>30</td></tr>
<tr><td>Bob</td><td>25</td></tr>
</table>"""
        result = html_to_text(html)
        assert "Name" in result
        assert "Age" in result
        assert "Alice" in result
        assert "Bob" in result

    def test_table_row_cells_are_tab_separated(self):
        result = html_to_text("<table><tr><td>Alice</td><td>30</td></tr></table>")
        assert "Alice\t30" in result


class TestHtmlToTextEdgeCases:
    def test_excessive_newlines_collapsed(self):
        result = html_to_text("<p>A</p><p></p><p></p><p></p><p>B</p>")
        # Should not have more than 2 consecutive newlines
        assert "\n\n\n" not in result
        assert "A" in result
        assert "B" in result

    def test_html_entities(self):
        result = html_to_text("&amp; &lt; &gt; &quot; &nbsp;")
        assert "&" in result
        assert "<" in result
        assert ">" in result

    def test_nested_skip_tags(self):
        result = html_to_text("<style><style>nested</style></style><p>visible</p>")
        assert "nested" not in result
        assert "visible" in result

    def test_blockquote(self):
        result = html_to_text("<blockquote>Quoted text</blockquote>")
        assert "> Quoted text" in result

    def test_img_alt_text(self):
        result = html_to_text('<img src="photo.jpg" alt="A nice photo">')
        assert "[A nice photo]" in result

    def test_img_without_alt_is_dropped(self):
        result = html_to_text('<p>Before</p><img src="photo.jpg"><p>After</p>')
        assert "photo.jpg" not in result
        assert "Before" in result
        assert "After" in result

    def test_preformatted_block_keeps_its_text(self):
        result = html_to_text("<pre>line one\nline two</pre>")
        assert "line one\nline two" in result
