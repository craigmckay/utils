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


def update_covers():
    conn = sqlite3.connect(DB_PATH)
    conn.create_function("title_sort", 1, title_sort)
    cursor = conn.cursor()

    cursor.execute("UPDATE books SET pubdate='1976-01-01 00:00:00+00:00' WHERE id=2040")
    cursor.execute("UPDATE books SET pubdate='1976-02-01 00:00:00+00:00' WHERE id=2425")
    cursor.execute("UPDATE books SET pubdate='1976-03-01 00:00:00+00:00' WHERE id=2422")
    cursor.execute("UPDATE books SET pubdate='1976-04-01 00:00:00+00:00' WHERE id=2419")
    cursor.execute("UPDATE books SET pubdate='1976-05-01 00:00:00+00:00' WHERE id=2420")
    cursor.execute("UPDATE books SET pubdate='1976-06-01 00:00:00+00:00' WHERE id=2421")
    cursor.execute("UPDATE books SET pubdate='1976-07-01 00:00:00+00:00' WHERE id=2423")
    cursor.execute("UPDATE books SET pubdate='1976-08-01 00:00:00+00:00' WHERE id=2424")
    cursor.execute("UPDATE books SET pubdate='1976-09-01 00:00:00+00:00' WHERE id=2418")
    cursor.execute("UPDATE books SET pubdate='1976-10-01 00:00:00+00:00' WHERE id=2427")
    cursor.execute("UPDATE books SET pubdate='1976-11-01 00:00:00+00:00' WHERE id=2428")
    cursor.execute("UPDATE books SET pubdate='1976-12-01 00:00:00+00:00' WHERE id=2426")
    
    conn.commit()
    conn.close()
    print("Done!")

if __name__ == "__main__":
    update_covers()
