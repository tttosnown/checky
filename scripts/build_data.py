from pathlib import Path
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from html import unescape

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
DATA_DIR.mkdir(exist_ok=True)

OUTPUT = DATA_DIR / "checky-data.json"

BOOKS = [
    ("history", "建筑历史", "01-《秋季必背册子·建筑历史》.docx"),
    ("urban", "城市设计", "02-《秋季必背册子·城市设计》.docx"),
    ("physics", "建筑物理", "03-《秋季必背册子·建筑物理》.docx"),
    ("structure", "构造技术", "04-《秋季必背册子·构造技术》.docx"),
]

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


# =========================================================
# 基础文字处理
# =========================================================

def clean_text(text):
    if not text:
        return ""

    text = unescape(text)
    text = text.replace("\u00a0", " ")
    text = text.replace("\u200b", "")
    text = text.replace("\ufeff", "")
    text = text.replace("\r\n", "\n")
    text = text.replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)

    lines = []

    for line in text.split("\n"):
        line = line.strip()
        if line:
            lines.append(line)

    return "\n".join(lines).strip()


def normalize_line(line):
    line = clean_text(line)

    line = re.sub(
        r"^[•·▪●○◆◇■□★☆]+\s*",
        "",
        line
    )

    line = re.sub(r"\s+", " ", line)

    return line.strip()


# =========================================================
# 只读取 DOCX 的 document.xml
#
# 重要：
# 不读取 word/media/*
# 因此不会触发 image21.png CRC 错误
# =========================================================

def read_docx_paragraphs(doc_path):

    doc_path = Path(doc_path)

    if not doc_path.exists():
        raise FileNotFoundError(
            f"找不到文件：{doc_path}"
        )

    with zipfile.ZipFile(doc_path, "r") as z:

        names = set(z.namelist())

        if "word/document.xml" not in names:
            raise RuntimeError(
                f"{doc_path.name} 中没有 word/document.xml"
            )

        # 这里只读取文字 XML
        xml_bytes = z.read("word/document.xml")

    root = ET.fromstring(xml_bytes)

    paragraphs = []

    for paragraph in root.findall(".//w:p", NS):

        parts = []

        for node in paragraph.iter():

            tag = node.tag

            if tag == f"{{{W_NS}}}t":

                if node.text:
                    parts.append(node.text)

            elif tag == f"{{{W_NS}}}tab":

                parts.append(" ")

            elif tag == f"{{{W_NS}}}br":

                parts.append("\n")

            elif tag == f"{{{W_NS}}}cr":

                parts.append("\n")

        text = "".join(parts)
        text = normalize_line(text)

        if text:
            paragraphs.append(text)

    return paragraphs


# =========================================================
# 判断是不是章节标题
# =========================================================

def is_chapter_heading(text):

    if not text:
        return False

    patterns = [

        # 第一章 / 第二章
        r"^第[一二三四五六七八九十百千万0-9]+[章节篇部分]",

        # 一、 二、 三、
        r"^[一二三四五六七八九十百千万]+[、.．]",

        # 1、 2、 3、
        r"^[0-9]+[、.．]",

        # 1.1 / 2.3
        r"^[0-9]+\.[0-9]+",

        # 1. / 2.
        r"^[0-9]+\.",

    ]

    for pattern in patterns:

        if re.match(pattern, text):
            return True

    # 短标题也视为章节
    if (
        len(text) <= 35
        and not re.search(r"[。！？；，]", text)
    ):
        return True

    return False


# =========================================================
# 判断明显不是正文的东西
# =========================================================

def is_noise(text):

    if not text:
        return True

    if re.fullmatch(r"\d+", text):
        return True

    if re.fullmatch(
        r"[-—_~·•●○◆◇■□*]+",
        text
    ):
        return True

    return False


# =========================================================
# 长段落拆分
# =========================================================

def split_text(text, max_length=180):

    text = clean_text(text)

    if not text:
        return []

    if len(text) <= max_length:
        return [text]

    pieces = re.split(
        r"(?<=[。！？；])",
        text
    )

    result = []
    current = ""

    for piece in pieces:

        piece = piece.strip()

        if not piece:
            continue

        if len(current) + len(piece) <= max_length:

            current += piece

        else:

            if current:
                result.append(current)

            if len(piece) <= max_length:

                current = piece

            else:

                for i in range(
                    0,
                    len(piece),
                    max_length
                ):

                    result.append(
                        piece[i:i + max_length]
                    )

                current = ""

    if current:
        result.append(current)

    return result


# =========================================================
# 生成题目
# =========================================================

def make_question(
    question_id,
    subject_id,
    subject_name,
    chapter_id,
    chapter_name,
    text
):

    return {
        "id": question_id,

        "subject": subject_id,
        "subjectName": subject_name,

        "chapter": chapter_id,
        "chapterName": chapter_name,

        "type": "recall",

        "question": text,
        "answer": text,

        "content": text,

        "mastery": 0,
        "difficulty": 1,

    }


# =========================================================
# 解析一本书
# =========================================================

def parse_book(
    book_id,
    subject_name,
    filename
):

    doc_path = ROOT / filename

    print("")
    print("=" * 70)
    print(f"处理：{subject_name}")
    print(f"文件：{filename}")
    print("=" * 70)

    paragraphs = read_docx_paragraphs(
        doc_path
    )

    paragraphs = [
        normalize_line(x)
        for x in paragraphs
        if not is_noise(x)
    ]

    print(
        f"读取文字段落：{len(paragraphs)}"
    )

    chapters = []
    questions = []

    current_chapter = None
    chapter_index = 0
    question_index = 0

    for paragraph in paragraphs:

        # -------------------------------------------------
        # 发现章节
        # -------------------------------------------------

        if is_chapter_heading(paragraph):

            chapter_index += 1

            chapter_id = (
                f"{book_id}-chapter-{chapter_index}"
            )

            current_chapter = {
                "id": chapter_id,
                "name": paragraph,
                "title": paragraph,
                "subject": book_id,
                "subjectName": subject_name,
                "questions": [],
                "questionIds": [],
            }

            chapters.append(
                current_chapter
            )

            continue

        # -------------------------------------------------
        # 如果正文还没有章节
        # -------------------------------------------------

        if current_chapter is None:

            chapter_index += 1

            chapter_id = (
                f"{book_id}-chapter-{chapter_index}"
            )

            current_chapter = {
                "id": chapter_id,
                "name": "基础内容",
                "title": "基础内容",
                "subject": book_id,
                "subjectName": subject_name,
                "questions": [],
                "questionIds": [],
            }

            chapters.append(
                current_chapter
            )

        # -------------------------------------------------
        # 正文生成题目
        # -------------------------------------------------

        pieces = split_text(
            paragraph
        )

        for piece in pieces:

            if not piece:
                continue

            question_index += 1

            question_id = (
                f"{book_id}-q-{question_index}"
            )

            question = make_question(
                question_id,
                book_id,
                subject_name,
                current_chapter["id"],
                current_chapter["name"],
                piece
            )

            questions.append(
                question
            )

            current_chapter[
                "questions"
            ].append(question)

            current_chapter[
                "questionIds"
            ].append(question_id)

    # -----------------------------------------------------
    # 章节数量统计
    # -----------------------------------------------------

    for chapter in chapters:

        chapter["count"] = len(
            chapter["questions"]
        )

        chapter["questionCount"] = len(
            chapter["questions"]
        )

    print(
        f"章节：{len(chapters)}"
    )

    print(
        f"题目：{len(questions)}"
    )

    return {
        "id": book_id,

        "subject": book_id,
        "subjectName": subject_name,

        "name": subject_name,
        "title": subject_name,

        "source": filename,

        "chapters": chapters,

        "questions": questions,

        "count": len(questions),
        "questionCount": len(questions),

    }


# =========================================================
# 去重
# =========================================================

def deduplicate_questions(
    questions
):

    result = []
    seen = set()

    for q in questions:

        key = (
            q.get("subject", ""),
            q.get("chapter", ""),
            q.get("question", "")
        )

        if key in seen:
            continue

        seen.add(key)

        result.append(q)

    return result


# =========================================================
# 主程序
# =========================================================

def main():

    print("")
    print("=" * 70)
    print("CHECKY DATA BUILDER")
    print("=" * 70)
    print(
        "DOCX读取模式：仅读取 word/document.xml"
    )
    print(
        "不会读取 word/media 图片"
    )
    print("=" * 70)

    books = []

    all_questions = []

    all_chapters = []

    total_subjects = 0

    for book in BOOKS:

        try:

            parsed = parse_book(
                *book
            )

            parsed["questions"] = (
                deduplicate_questions(
                    parsed["questions"]
                )
            )

            books.append(
                parsed
            )

            all_questions.extend(
                parsed["questions"]
            )

            all_chapters.extend(
                parsed["chapters"]
            )

            total_subjects += 1

        except Exception as e:

            print("")
            print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
            print(
                f"{book[1]} 处理失败"
            )
            print(
                str(e)
            )
            print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")

            books.append({

                "id": book[0],

                "subject": book[0],
                "subjectName": book[1],

                "name": book[1],
                "title": book[1],

                "source": book[2],

                "chapters": [],
                "questions": [],

                "count": 0,
                "questionCount": 0,

                "error": str(e)

            })

    # -----------------------------------------------------
    # 最终去重
    # -----------------------------------------------------

    all_questions = (
        deduplicate_questions(
            all_questions
        )
    )

    # -----------------------------------------------------
    # 构建最终数据
    # -----------------------------------------------------

    data = {

        "version": "2.0",

        "generated": True,

        "subjects": books,

        "books": books,

        "chapters": all_chapters,

        "questions": all_questions,

        "stats": {

            "subjectCount":
                total_subjects,

            "bookCount":
                len(books),

            "chapterCount":
                len(all_chapters),

            "questionCount":
                len(all_questions),

        }

    }

    # -----------------------------------------------------
    # 写入 JSON
    # -----------------------------------------------------

    OUTPUT.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    with OUTPUT.open(
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            data,
            f,
            ensure_ascii=False,
            indent=2
        )

    print("")
    print("=" * 70)
    print("BUILD SUCCESS")
    print("=" * 70)

    print(
        f"科目：{total_subjects}"
    )

    print(
        f"章节：{len(all_chapters)}"
    )

    print(
        f"题目：{len(all_questions)}"
    )

    print(
        f"输出：{OUTPUT}"
    )

    print("=" * 70)


if __name__ == "__main__":
    main()
