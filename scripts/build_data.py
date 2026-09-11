from pathlib import Path
import json
import zipfile
import re
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
DATA_DIR.mkdir(exist_ok=True)
OUTPUT = DATA_DIR / "checky-data.json"

BOOKS = [
    ("history", "建筑历史", "01-《秋季必背册子·建筑历史》.docx"),
    ("urban", "城市设计", "02-《秋季必背册子·城市设计》.docx"),
    ("physics", "建筑物理", "03-《秋季必背册子·建筑物理》.docx"),
    ("construction", "构造技术", "04-《秋季必背册子·构造技术》.docx"),
]

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def read_docx(path):
    with zipfile.ZipFile(path, "r") as z:
        xml = z.read("word/document.xml")

    root = ET.fromstring(xml)
    paragraphs = []

    for p in root.findall(".//w:p", NS):
        text = "".join(t.text or "" for t in p.findall(".//w:t", NS))
        text = re.sub(r"\s+", " ", text).strip()

        if text:
            paragraphs.append(text)

    return paragraphs


def is_heading(text):
    patterns = [
        r"^第[一二三四五六七八九十百千万0-9]+[章节篇]",
        r"^[一二三四五六七八九十]+、",
        r"^\d+(\.\d+)*[、.]?\s*",
        r"^第[一二三四五六七八九十]+部分",
        r"^第[一二三四五六七八九十]+篇",
        r"^(绪论|概述|目录|附录|总结)$",
    ]

    return any(re.match(p, text) for p in patterns)


def make_questions(subject_id, chapter_id, chapter_name, paragraphs):
    questions = []

    for i, text in enumerate(paragraphs, 1):
        if len(text) < 8:
            continue

        qid = f"{chapter_id}-q{i}"

        keyword = text[:80]

        questions.append({
            "id": qid,
            "subject": subject_id,
            "chapterId": chapter_id,
            "chapter": chapter_name,
            "type": "recall",
            "question": f"请主动回忆并说明：{text}",
            "answer": text,
            "content": text,
            "keyword": keyword
        })

    return questions


def build_book(book_id, subject_name, filename):
    path = ROOT / filename

    if not path.exists():
        print(f"WARNING: file not found: {filename}")
        return {
            "id": book_id,
            "name": subject_name,
            "filename": filename,
            "chapters": []
        }, []

    paragraphs = read_docx(path)

    chapters = []
    current_name = None
    current_paragraphs = []

    for text in paragraphs:
        if is_heading(text):
            if current_name and current_paragraphs:
                chapters.append(
                    (current_name, current_paragraphs)
                )

            current_name = text
            current_paragraphs = []
        else:
            if current_name is None:
                current_name = "全文内容"

            current_paragraphs.append(text)

    if current_name and current_paragraphs:
        chapters.append(
            (current_name, current_paragraphs)
        )

    if not chapters:
        chapters = [("全文内容", paragraphs)]

    book_chapters = []
    all_questions = []

    for index, (chapter_name, chapter_paragraphs) in enumerate(chapters, 1):
        chapter_id = f"{book_id}-ch{index}"

        questions = make_questions(
            book_id,
            chapter_id,
            chapter_name,
            chapter_paragraphs
        )

        question_ids = [q["id"] for q in questions]

        chapter = {
            "id": chapter_id,
            "bookId": book_id,
            "subject": book_id,
            "name": chapter_name,
            "title": chapter_name,
            "count": len(questions),
            "questionCount": len(questions),
            "questionIds": question_ids,
            "questions": questions
        }

        book_chapters.append(chapter)
        all_questions.extend(questions)

    book = {
        "id": book_id,
        "name": subject_name,
        "title": subject_name,
        "filename": filename,
        "chapters": book_chapters,
        "chapterCount": len(book_chapters),
        "questionCount": len(all_questions)
    }

    return book, all_questions


def main():
    books = []
    chapters = []
    questions = []

    for book_id, subject_name, filename in BOOKS:
        book, book_questions = build_book(
            book_id,
            subject_name,
            filename
        )

        books.append(book)
        questions.extend(book_questions)

        for chapter in book["chapters"]:
            chapters.append(chapter)

        print(
            subject_name,
            "chapters=",
            len(book["chapters"]),
            "questions=",
            len(book_questions)
        )

    data = {
        "version": "2.0",
        "books": books,
        "subjects": books,
        "chapters": chapters,
        "questions": questions,
        "stats": {
            "books": len(books),
            "chapters": len(chapters),
            "questions": len(questions)
        }
    }

    OUTPUT.write_text(
        json.dumps(
            data,
            ensure_ascii=False,
            indent=2
        ),
        encoding="utf-8"
    )

    print("Generated:", OUTPUT)
    print("Books:", len(books))
    print("Chapters:", len(chapters))
    print("Questions:", len(questions))


if __name__ == "__main__":
    main()
