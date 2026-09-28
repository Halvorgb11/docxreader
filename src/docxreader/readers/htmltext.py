"""HTML -> markdown-lignende tekst, med `html.parser` fra standardbiblioteket.

Brukes for e-poster som bare har HTML-tekst (og kan brukes av en HTML-leser
senere). Målet er lesbar tekst, ikke perfekt markdown:
- <p>, <div>, <br> -> linjeskift / tomme linjer
- <h1>–<h6>        -> #-overskrifter
- <li>             -> "- punkt"
- <table>          -> markdown-tabell
- <script>, <style>, <head> hoppes over
"""

import re
from html.parser import HTMLParser

from docxreader.blocks import rows_to_markdown

BLOCK_TAGS = {"p", "div", "ul", "ol", "blockquote", "hr", "section", "article", "header", "footer"}
SKIP_TAGS = {"script", "style", "head", "title"}


class _HTMLToText(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)  # &amp; -> &, &nbsp; -> mellomrom osv.
        self.out: list[str] = []
        self.skip = 0
        # En stabel med tabeller (tabeller kan ligge inni tabeller). Hver tabell
        # er en liste med rader; hver rad en liste med celletekster.
        self.tables: list[list[list[str]]] = []

    def _write(self, text: str) -> None:
        if self.tables and self.tables[-1] and self.tables[-1][-1]:
            self.tables[-1][-1][-1] += text  # inni en tabellcelle
        elif not self.tables:
            self.out.append(text)

    def handle_starttag(self, tag, attrs):
        if tag in SKIP_TAGS:
            self.skip += 1
        elif tag == "br":
            self._write(" " if self.tables else "\n")
        elif tag in BLOCK_TAGS:
            self._write("\n\n")
        elif re.fullmatch(r"h[1-6]", tag):
            self._write("\n\n" + "#" * int(tag[1]) + " ")
        elif tag == "li":
            self._write("\n- ")
        elif tag == "table":
            self.tables.append([])
        elif tag == "tr" and self.tables:
            self.tables[-1].append([])
        elif tag in ("td", "th") and self.tables:
            if not self.tables[-1]:
                self.tables[-1].append([])  # celle uten <tr> først
            self.tables[-1][-1].append("")

    def handle_endtag(self, tag):
        if tag in SKIP_TAGS:
            self.skip = max(0, self.skip - 1)
        elif tag in BLOCK_TAGS or re.fullmatch(r"h[1-6]", tag):
            self._write("\n\n")
        elif tag == "table" and self.tables:
            rows = [row for row in self.tables.pop() if any(cell.strip() for cell in row)]
            text = rows_to_markdown([[" ".join(c.split()) for c in row] for row in rows]) if rows else ""
            if self.tables:  # tabell inni en tabell: legg teksten i ytre celle
                self._write(" ".join(text.split()))
            else:
                self.out.append("\n\n" + text + "\n\n")

    def handle_data(self, data):
        if not self.skip:
            self._write(re.sub(r"\s+", " ", data))  # HTML: all whitespace = ett mellomrom


def html_to_text(html: str) -> str:
    parser = _HTMLToText()
    parser.feed(html)
    parser.close()
    text = "".join(parser.out)
    lines = [line.strip() for line in text.splitlines()]
    return re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()
