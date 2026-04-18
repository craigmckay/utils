#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pdf_bookmark.py

Add bookmarks to a PDF based on every BOLD + UPPERCASE paragraph, and
configure the PDF's open-view preferences:
    - Page mode:   Bookmarks + content  (/UseOutlines)
    - Page layout: One page at a time   (/SinglePage)
    - Zoom:        Fit to page          (/Fit)


python pdf_bookmark.py "C:\Users\craig\Dropbox\Christian\Ministry\AWiiS\230\May 26 - A5F.pdf" "C:\Users\craig\Dropbox\Christian\Ministry\AWiiS\230\awiis_230.pdf" --title "A Word in its Season 230 - May 2026" --author "Various"

"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import List, Tuple

import pdfplumber
from pypdf import PdfReader, PdfWriter
from pypdf.generic import Fit


# ---------------------------------------------------------------------------
# Heading detection
# ---------------------------------------------------------------------------

def _is_bold(fontname: str | None) -> bool:
    """A font is considered bold if 'Bold' appears anywhere in its name."""
    return bool(fontname) and "Bold" in fontname


def _line_is_bold_upper(line_chars) -> bool:
    """
    A line counts as a bold-uppercase heading line if:
      * every non-whitespace character is rendered in a bold font, AND
      * every alphabetic character is uppercase, AND
      * the line contains at least one alphabetic character.
    Punctuation, digits, quotation marks etc. are ignored for the upper check.
    """
    non_space = [c for c in line_chars if c["text"].strip()]
    if not non_space:
        return False
    if not all(_is_bold(c.get("fontname", "")) for c in non_space):
        return False
    alpha = [c["text"] for c in non_space if c["text"].isalpha()]
    if not alpha:
        return False
    return all(ch.isupper() for ch in alpha)


def find_bold_upper_paragraphs(pdf_path: Path) -> List[Tuple[int, str]]:
    """
    Walk every page and return a list of (page_index_0based, paragraph_text)
    for each BOLD + UPPERCASE paragraph found.

    Adjacent bold-uppercase lines on the same page are merged into a single
    paragraph (so a multi-line title becomes one bookmark).
    """
    results: List[Tuple[int, str]] = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_index, page in enumerate(pdf.pages):
            # Page 1 is the cover page -- never carries real article headings.
            if page_index == 0:
                continue
            chars = page.chars
            if not chars:
                continue

            # Group chars into lines keyed by rounded top-y position.
            lines_by_top: dict[int, list] = {}
            for ch in chars:
                lines_by_top.setdefault(round(ch["top"]), []).append(ch)

            # Sort lines top-to-bottom; build (top, text, is_heading) list.
            line_records = []
            for top in sorted(lines_by_top.keys()):
                line_chars = sorted(lines_by_top[top], key=lambda c: c["x0"])
                text = "".join(c["text"] for c in line_chars).strip()
                if not text:
                    continue
                line_records.append(
                    (top, text, _line_is_bold_upper(line_chars))
                )

            # Merge consecutive heading lines into single paragraphs.
            i = 0
            while i < len(line_records):
                top, text, is_head = line_records[i]
                if not is_head:
                    i += 1
                    continue

                parts = [text]
                j = i + 1
                while j < len(line_records) and line_records[j][2]:
                    parts.append(line_records[j][1])
                    j += 1

                paragraph = " ".join(parts)
                # Ignore one-character stray headings, keep everything else.
                if len(paragraph) >= 2:
                    results.append((page_index, paragraph))
                i = j

    return results


# ---------------------------------------------------------------------------
# PDF writing
# ---------------------------------------------------------------------------

def build_output_pdf(
    input_path: Path,
    output_path: Path,
    bookmarks: List[Tuple[int, str]],
    pdf_title: str | None = None,
    pdf_author: str | None = None,
) -> None:
    """Clone the input PDF, add bookmarks, and set the open-view options."""
    reader = PdfReader(str(input_path))
    writer = PdfWriter(clone_from=reader)

    # Remove any pre-existing outline so only our bookmarks appear.
    # pypdf exposes the outline through an internal root list; the safest
    # way to clear it is to drop /Outlines from the document catalog.
    from pypdf.generic import NameObject
    if NameObject("/Outlines") in writer._root_object:
        del writer._root_object[NameObject("/Outlines")]
    writer._outline = []  # reset pypdf's internal outline tracking

    # --- Bookmarks (outline entries) -------------------------------------
    # Each bookmark uses a /Fit destination so it lands at "Fit page" zoom
    # on the target page.
    for page_index, bm_title in bookmarks:
        if 0 <= page_index < len(writer.pages):
            writer.add_outline_item(
                title=bm_title,
                page_number=page_index,
                fit=Fit.fit(),  # /Fit -- whole page fits in the window
            )

    # --- Viewer preferences ----------------------------------------------
    # Show the bookmarks pane alongside the page content.
    writer.page_mode = "/UseOutlines"
    # One full page shown at a time (no continuous scroll).
    writer.page_layout = "/SinglePage"

    # Also set an explicit OpenAction on the first page at /Fit zoom, so a
    # viewer that honours it opens the document already fit-to-page.
    if len(writer.pages) > 0:
        from pypdf.generic import ArrayObject
        first_page_ref = writer.pages[0].indirect_reference
        # OpenAction array: [ page /Fit ] -- "fit the whole page in the window"
        open_action = ArrayObject([first_page_ref, NameObject("/Fit")])
        writer._root_object[NameObject("/OpenAction")] = open_action

    # --- Document metadata -----------------------------------------------
    if pdf_title is not None or pdf_author is not None:
        from pypdf.generic import create_string_object

        # Update the /Info dictionary (traditional PDF metadata).
        info = writer._info.get_object()
        if pdf_title is not None:
            info[NameObject("/Title")] = create_string_object(pdf_title)
            print("Title:  " + pdf_title)
        if pdf_author is not None:
            info[NameObject("/Author")] = create_string_object(pdf_author)
            print("Author: " + pdf_author)

        # PDF viewers (Acrobat, etc.) prefer XMP over /Info when both exist.
        # Remove any embedded XMP stream so our /Info values are used.
        if NameObject("/Metadata") in writer._root_object:
            del writer._root_object[NameObject("/Metadata")]

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("wb") as fh:
        writer.write(fh)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Add bookmarks to a PDF for every BOLD + UPPERCASE paragraph, "
            "and set the file to open with the bookmarks pane visible, "
            "fit-to-page zoom, and one-page-at-a-time layout."
        )
    )
    parser.add_argument("input", type=Path, help="Input PDF file")
    parser.add_argument(
        "output",
        type=Path,
        nargs="?",
        help="Output PDF file (default: <input>_bookmarked.pdf)",
    )
    parser.add_argument(
        "--title",
        metavar="TEXT",
        help="Set the PDF Title metadata property",
    )
    parser.add_argument(
        "--author",
        metavar="TEXT",
        help="Set the PDF Author metadata property",
    )
    args = parser.parse_args(argv)

    input_path: Path = args.input
    if not input_path.is_file():
        print(f"error: input file not found: {input_path}", file=sys.stderr)
        return 1

    output_path: Path = (
        args.output
        if args.output is not None
        else input_path.with_name(f"{input_path.stem}_bookmarked.pdf")
    )

    print(f"Scanning {input_path.name} for bold uppercase paragraphs...")
    bookmarks = find_bold_upper_paragraphs(input_path)

    if not bookmarks:
        print("No bold uppercase paragraphs found. Writing file with "
              "view options set but no bookmarks.")
    else:
        print(f"Found {len(bookmarks)} bookmark(s):")
        for page_index, bm_title in bookmarks:
            print(f"  p.{page_index + 1}: {bm_title}")

    build_output_pdf(
        input_path,
        output_path,
        bookmarks,
        pdf_title=args.title,
        pdf_author=args.author,
    )
    print(f"\nWrote: {output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())