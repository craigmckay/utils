#\!/usr/bin/env python3
# -*- coding: utf-8 -*-
#
# calibre_repdf.py
#
# Finds all books in Calibre tagged with both 'awiis' and 're-pdf',
# strips the booklet imposition with Ghostscript, then adds bookmarks
# using pdf_bookmark.py — overwriting the original PDF in the library.
#
# Usage:
#   python calibre_repdf.py
#
# Prerequisites: gswin64c on PATH, pdfplumber + pypdf installed.
# Close Calibre before running (SQLite locking).

import sqlite3
import subprocess
import sys
import tempfile
from pathlib import Path

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

CALIBRE_DB  = Path(r"C:\Users\craig\Calibre Library\metadata.db")
CALIBRE_LIB = Path(r"C:\Users\craig\Calibre Library")
BOOKMARK_PY = Path(r"c:\dev\utils\ministry\pdf_bookmark.py")
GSWIN       = "gswin64c"   # must be on PATH


# ---------------------------------------------------------------------------
# Database
# ---------------------------------------------------------------------------

def get_books():
    """Return all books tagged 'awiis' AND 're-pdf' with their PDF paths."""
    conn = sqlite3.connect(str(CALIBRE_DB))
    conn.row_factory = sqlite3.Row

    rows = conn.execute("""
        SELECT DISTINCT b.id, b.title, b.path, d.name AS filename
        FROM   books b
        JOIN   books_tags_link bt1 ON b.id = bt1.book
        JOIN   tags            t1  ON bt1.tag = t1.id AND LOWER(t1.name) = 'awiis'
        JOIN   books_tags_link bt2 ON b.id = bt2.book
        JOIN   tags            t2  ON bt2.tag = t2.id AND LOWER(t2.name) = 're-pdf'
        JOIN   data            d   ON b.id = d.book  AND d.format = 'PDF'
        ORDER  BY b.sort
    """).fetchall()

    books = []
    for row in rows:
        authors = conn.execute("""
            SELECT a.name
            FROM   authors           a
            JOIN   books_authors_link ba ON a.id = ba.author
            WHERE  ba.book = ?
            ORDER  BY ba.id
        """, (row["id"],)).fetchall()

        author_str = " & ".join(a["name"] for a in authors)
        pdf_path   = CALIBRE_LIB / row["path"] / (row["filename"] + ".pdf")

        books.append({
            "id":       row["id"],
            "title":    row["title"],
            "author":   author_str,
            "pdf_path": pdf_path,
        })

    conn.close()
    return books


# ---------------------------------------------------------------------------
# Processing
# ---------------------------------------------------------------------------

def process_book(book):
    pdf_path = book["pdf_path"]
    title    = book["title"]
    author   = book["author"]

    print(f"\n{'='*60}")
    print(f"Book:   {pdf_path.name}")
    print(f"Title:  {title}")
    print(f"Author: {author}")

    if not pdf_path.exists():
        print(f"ERROR: PDF not found at {pdf_path}")
        return False

    # Create a temp file in the same folder (avoids cross-drive moves later)
    tmp_fd, tmp_name = tempfile.mkstemp(suffix=".pdf", dir=pdf_path.parent)
    tmp_path = Path(tmp_name)
    import os; os.close(tmp_fd)

    # Step 1: Ghostscript — strip booklet imposition, compress
    print("\n[1/2] Ghostscript: stripping booklet content...")
    gs = subprocess.run(
        [
            GSWIN,
            "-dBATCH", "-dNOPAUSE", "-dSAFER",
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.7",
            "-dPDFSETTINGS=/ebook",
            "-dCompressFonts=true",
            "-dSubsetFonts=true",
            f"-sOutputFile={tmp_path}",
            str(pdf_path),
        ],
        capture_output=True, text=True
    )

    if gs.returncode != 0 or not tmp_path.exists() or tmp_path.stat().st_size == 0:
        print("ERROR: Ghostscript failed.")
        if gs.stderr:
            print(gs.stderr[-600:])
        tmp_path.unlink(missing_ok=True)
        return False

    # Step 2: pdf_bookmark.py — detect headings and overwrite the original
    print("[2/2] Adding bookmarks (overwriting original)...")
    bm = subprocess.run(
        [
            sys.executable,
            str(BOOKMARK_PY),
            str(tmp_path),
            str(pdf_path),
            "--title",  title,
            "--author", author,
        ],
        capture_output=True, text=True
    )

    tmp_path.unlink(missing_ok=True)

    if bm.returncode != 0:
        print("ERROR: pdf_bookmark.py failed.")
        if bm.stderr:
            print(bm.stderr[-600:])
        return False

    print(bm.stdout.strip())
    print("Done.")
    return True


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    print(f"Calibre library : {CALIBRE_LIB}")
    print(f"Bookmark script : {BOOKMARK_PY}")
    print(f"Ghostscript     : {GSWIN}")

    # Sanity checks
    if not CALIBRE_DB.exists():
        sys.exit(f"ERROR: Calibre database not found: {CALIBRE_DB}")
    if not BOOKMARK_PY.exists():
        sys.exit(f"ERROR: pdf_bookmark.py not found: {BOOKMARK_PY}")

    books = get_books()
    if not books:
        print("\nNo books found with tags 'awiis' and 're-pdf'. Nothing to do.")
        return

    print(f"\nFound {len(books)} book(s) to process:")
    for b in books:
        print(f"  • {b['title']}  ({b['author']})")

    ok = fail = 0
    for book in books:
        if process_book(book):
            ok += 1
        else:
            fail += 1

    print(f"\n{'='*60}")
    print(f"Complete: {ok} succeeded, {fail} failed.")


if __name__ == "__main__":
    main()