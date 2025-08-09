import pyodbc
import json
import os
import re
from collections import defaultdict
from utils.book_map import book_map

# === 1. Setup Paths for Multiple Versions ===
#write down the name as is from the file!
version_files = [
    ("kkjvdb", "kkjvdb.mdb"),
    ("nivdb", "nivdb.mdb"),
    ("ngayok", "ngayok.mdb")
]

json_path = os.path.join("bible_data", "bible_combined.json")

# === 2. Regex to Extract Verses ===
verse_pattern = re.compile(r'(\d+:\d+)\s+(.*?)(?=\d+:\d+|\Z)', re.DOTALL)

# === 3. Nested defaultdict Structure ===
combined = defaultdict(lambda: defaultdict(lambda: defaultdict(dict)))

# === 4. Read Each Bible Version ===
version_key_map = {
    "kkjvdb": "kkjv",
    "nivdb": "niv",
    "ngayok": "ngayok"
}

for version, filename in version_files:
    db_path = os.path.join("bible_data", filename)
    print(f"🔄 Processing version: {version} from {filename}")

    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        rf'DBQ={db_path};'
    )
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    table_name = version.upper()
    cursor.execute(f"SELECT BOOK, TCHP, CONTENT FROM {table_name}")
    rows = cursor.fetchall()

    lang_key = version_key_map[version]

    for row in rows:
        book = row.BOOK.strip()
        chapter = str(row.TCHP)
        content = row.CONTENT.strip()

        for ref, text in verse_pattern.findall(content):
            chap, verse = ref.split(":")
            combined[book][chap][verse][lang_key] = text.strip()

    conn.close()


# === 5. Convert nested defaultdict to normal dict ===
def recursive_default_to_dict(d):
    if isinstance(d, defaultdict):
        return {k: recursive_default_to_dict(v) for k, v in d.items()}
    return d

final_output = recursive_default_to_dict(combined)

# === 6. Import the English Books!
for book_kor, chapters in final_output.items():
    eng = book_map.get(book_kor, "")
    for chapter in chapters.values():
        for verse in chapter.values():
            verse["book_eng"] = eng

with open("bible_data/bible_combined.json", "w", encoding="utf-8") as f:
    json.dump(final_output, f, ensure_ascii=False, indent=2)

# === 7. Save to JSON ===
os.makedirs(os.path.dirname(json_path), exist_ok=True)
with open(json_path, "w", encoding="utf-8") as f:
    json.dump(final_output, f, ensure_ascii=False, indent=2)

print(f"✅ Exported Bible to {json_path} with {len(final_output)} books")

