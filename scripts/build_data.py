from pathlib import Path
import json
import re
import zipfile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
DATA_DIR.mkdir(exist_ok=True)

OUTPUT = DATA_DIR / "checky-data.json"

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
}


BOOKS = [
    {
        "id": "history",
        "name": "建筑历史",
        "files": [
            "01-《秋季必背册子·建筑历史》.docx",
            "01-《秋季必背册子·建筑历史》（朴是内部资料，请勿转发）.pdf",
        ]
    },
    {
        "id": "urban",
        "name": "城市设计",
        "files": [
            "02-《秋季必背册子·城市设计》.docx",
            "02-《秋季必背册子·城市设计》（朴是内部资料，请勿转发）.pdf",
        ]
    },
    {
        "id": "construction",
        "name": "建筑构造",
        "files": [
            "03-《秋季必背册子·建筑构造》.docx",
            "03-《秋季必背册子·建筑构造》（朴是内部资料，请勿转发）.pdf",
            "04-《秋季必背册子·构造技术》.docx",
        ]
    }
]


def clean_text(text):
    text = text.replace("\u3000", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def read_docx(path):
    with zipfile.ZipFile(path, "r") as z:
        xml = z.read("word/document.xml")

    root = ET.fromstring(xml)
    paragraphs = []

    for p in root.findall(".//w:p", NS):
        text = "".join(
            t.text or ""
            for t in p.findall(".//w:t", NS)
        )

        text = clean_text(text)

        if text:
            paragraphs.append(text)

    return paragraphs


def read_pdf(path):
    try:
        from pypdf import PdfReader
    except ImportError:
        print("pypdf not installed")
        return []

    reader = PdfReader(str(path))
    paragraphs = []

    for page in reader.pages:
        try:
            text = page.extract_text() or ""
        except Exception:
            text = ""

        for line in text.splitlines():
            line = clean_text(line)

            if line:
                paragraphs.append(line)

    return paragraphs


def read_file(path):
    suffix = path.suffix.lower()

    if suffix == ".docx":
        return read_docx(path)

    if suffix == ".pdf":
        return read_pdf(path)

    return []


def find_source(book):
    for filename in book["files"]:
        path = ROOT / filename

        if path.exists():
            print("FOUND:", filename)
            return path

    print("NOT FOUND:", book["name"])
    return None


def is_heading(text):
    patterns = [
        r"^第[一二三四五六七八九十百千万0-9]+[章节篇]",
        r"^第[一二三四五六七八九十]+部分",
        r"^第[一二三四五六七八九十]+篇",
        r"^[一二三四五六七八九十]+、",
        r"^\d+(\.\d+)*[、.]",
        r"^(绪论|概述|目录|附录|总结|结语)$",
    ]

    return any(
        re.match(pattern, text)
        for pattern in patterns
    )


def split_chapters(paragraphs):
    chapters = []

    current_title = "全文内容"
    current_content = []

    for text in paragraphs:
        if is_heading(text):
            if current_content:
                chapters.append(
                    (current_title, current_content)
                )

            current_title = text
            current_content = []

        else:
            current_content.append(text)

    if current_content:
        chapters.append(
            (current_title, current_content)
        )

    if not chapters:
        chapters.append(
            ("全文内容", paragraphs)
        )

    return chapters


def make_keywords(text):
    text = clean_text(text)

    # 去掉常见的连接词，寻找比较适合主动回忆的关键词
    stop_words = [
        "的", "是", "和", "与", "及", "以及",
        "在", "对", "于", "为", "了", "中",
        "其", "等", "通过", "可以", "主要",
        "进行", "具有", "一个", "这种"
    ]

    parts = re.split(
        r"[，。；、：:（）()！？!? \t]+",
        text
    )

    result = []

    for part in parts:
        part = part.strip()

        if len(part) < 2:
            continue

        if part in stop_words:
            continue

        if part not in result:
            result.append(part)

        if len(result) >= 6:
            break

    if not result:
        result = [text[:12]]

    return " · ".join(result)


def make_question_title(text, index):
    text = clean_text(text)

    # 绝不把原文直接放进题目标题。
    # 标题只作为主动回忆提示。
    first = text[:12]

    return f"知识点 {index}：请回忆“{first}……”的完整内容"


def make_questions(book_id, chapter_id, chapter_name, paragraphs):
    questions = []

    for index, text in enumerate(paragraphs, 1):
        text = clean_text(text)

        if len(text) < 10:
            continue

        qid = f"{chapter_id}-q{len(questions)+1}"

        keyword = make_keywords(text)

        question = make_question_title(
            text,
            len(questions) + 1
        )

        questions.append({
            "id": qid,
            "subject": book_id,
            "chapterId": chapter_id,
            "chapter": chapter_name,
            "type": "recall",

            "shortTitle": f"主动回忆第 {len(questions)+1} 个知识点",

            "question": question,

            "keyword": keyword,

            "answer": text,

            "content": text
        })

    return questions


def build_book(book):
    source = find_source(book)

    if source is None:
        return {
            "id": book["id"],
            "name": book["name"],
            "title": book["name"],
            "chapters": [],
            "chapterCount": 0,
            "questionCount": 0
        }

    paragraphs = read_file(source)

    print(
        book["name"],
        "paragraphs=",
        len(paragraphs)
    )

    chapter_data = split_chapters(paragraphs)

    chapters = []
    total_questions = 0

    for index, (chapter_name, content) in enumerate(
        chapter_data,
        1
    ):
        chapter_id = f"{book['id']}-ch{index}"

        questions = make_questions(
            book["id"],
            chapter_id,
            chapter_name,
            content
        )

        question_ids = [
            q["id"]
            for q in questions
        ]

        chapter = {
            "id": chapter_id,
            "bookId": book["id"],
            "subject": book["id"],
            "name": chapter_name,
            "title": chapter_name,

            "questionCount": len(questions),
            "count": len(questions),

            "questionIds": question_ids,
            "questions": questions
        }

        chapters.append(chapter)
        total_questions += len(questions)

    return {
        "id": book["id"],
        "name": book["name"],
        "title": book["name"],

        "chapters": chapters,

        "chapterCount": len(chapters),
        "questionCount": total_questions
    }


def main():
    books = []
    all_chapters = []
    all_questions = []

    for book in BOOKS:
        result = build_book(book)

        books.append(result)

        for chapter in result["chapters"]:
            all_chapters.append(chapter)

            for question in chapter["questions"]:
                all_questions.append(question)

        print(
            "RESULT:",
            result["name"],
            "chapters=",
            result["chapterCount"],
            "questions=",
            result["questionCount"]
        )

    data = {
        "version": "3.0",

        "books": books,

        "subjects": books,

        "chapters": all_chapters,

        "questions": all_questions,

        "stats": {
            "books": len(books),
            "chapters": len(all_chapters),
            "questions": len(all_questions)
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

    print()
    print("================================")
    print("CHECKY DATA BUILD SUCCESS")
    print("Books:", len(books))
    print("Chapters:", len(all_chapters))
    print("Questions:", len(all_questions))
    print("Output:", OUTPUT)
    print("================================")


if __name__ == "__main__":
    main()
