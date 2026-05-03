import fitz  # PyMuPDF
import os
import sqlite3
import shutil
import subprocess

# Config
CALIBRE_LIBRARY = r"C:\Users\craig\Calibre Library"
DB_PATH = os.path.join(CALIBRE_LIBRARY, "metadata.db")
COVER_FILENAME = "cover.jpg"

def title_sort(title):
    # Calibre's title_sort moves leading articles to the end
    # e.g. "The Cat" -> "Cat, The"
    articles = ['the ', 'a ', 'an ']
    title_lower = title.lower()
    for article in articles:
        if title_lower.startswith(article):
            return title[len(article):] + ', ' + title[:len(article)].rstrip()
    return title

def convert_to_epub(docx_path, epub_path, cover_path):
    result = subprocess.run([
        r"C:\Program Files\Calibre2\ebook-convert",
        docx_path,
        epub_path,
        "--cover", cover_path,
        "--level1-toc", "//h:h1",
        "--level2-toc", "//h:h2",
        "--level3-toc", "//h:h3",
        "--epub-version", "3"
    ], capture_output=True, text=True)
    print(f"  stdout: {result.stdout}")
    print(f"  stderr: {result.stderr}")

def extract_first_page_as_image(pdf_path, output_path, dpi=150):
    doc = fitz.open(pdf_path)
    page = doc[0]
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = page.get_pixmap(matrix=mat)
    pix.save(output_path)
    doc.close()

def update_covers():
    conn = sqlite3.connect(DB_PATH)
    conn.create_function("title_sort", 1, title_sort)
    cursor = conn.cursor()

    cursor.execute("""
        SELECT 
            b.id,
            b.title,
            b.path,
            'C:\\Users\\craig\\Calibre Library\\' || replace(b.path, '/', '\\') || '\\' || d.name || '.pdf' AS pdf_full_path
        FROM books b
        JOIN books_series_link bsl ON b.id = bsl.book
        JOIN series s ON bsl.series = s.id
        LEFT JOIN books_publishers_link bpl ON b.id = bpl.book
        LEFT JOIN publishers p ON bpl.publisher = p.id
        JOIN data d ON b.id = d.book AND d.format = 'PDF'
        WHERE s.name = 'Notes of Ministry'
        AND (p.name NOT LIKE 'Brown%' OR p.name IS NULL)
        AND b.series_index BETWEEN 422 AND 433  
        ORDER BY b.series_index
    """)

    rows = cursor.fetchall()

    for row in rows:
        book_id, title, book_path, pdf_path = row
        book_folder = os.path.join(CALIBRE_LIBRARY, book_path.replace('/', '\\'))
        cover_path = os.path.join(book_folder, COVER_FILENAME)

        if not os.path.exists(pdf_path):
            print(f"PDF not found: {pdf_path}")
            continue

        try:
            extract_first_page_as_image(pdf_path, cover_path)

            # Update has_cover flag in Calibre db
            cursor.execute("UPDATE books SET has_cover = 1 WHERE id = ?", (book_id,))
            conn.commit()

            print(f"Updated cover: {title}")

        except Exception as e:
            print(f"Error processing {title}: {e}")

        docx_path = pdf_path.replace('.pdf', '.docx')
        epub_path = pdf_path.replace('.pdf', '.epub')
        convert_to_epub(docx_path, epub_path, cover_path)

    conn.close()
    print("Done!")

if __name__ == "__main__":
    update_covers()
