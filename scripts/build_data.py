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


# ============================================================
# 书籍配置
# ============================================================

BOOKS = [
    ("history", "建筑历史", "01-《秋季必背册子·建筑历史》.docx"),
    ("urban", "城市设计", "02-《秋季必背册子·城市设计》.docx"),
    ("physics", "建筑物理", "03-《秋季必背册子·建筑物理》.docx"),
    ("structure", "构造技术", "04-《秋季必背册子·构造技术》.docx"),
]


# ============================================================
# DOCX 文字读取
#
# 重要：
# 不使用 python-docx。
# 不读取 word/media/*
# 只读取 word/document.xml。
#
# 这样即使 DOCX 里面某一张图片损坏，
# 也不会因为 image21.png 的 CRC 错误导致整个构建失败。
# ============================================================

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

NS = {
    "w": W_NS
}


def clean_text(text):
    """清理 Word 文本中的多余空格和特殊字符。"""

    if not text:
        return ""

    text = unescape(text)

    # 常见 Word 特殊空白
    text = text.replace("\u00a0", " ")
    text = text.replace("\u200b", "")
    text = text.replace("\ufeff", "")

    # 统一换行
    text = text.replace("\r\n", "\n")
    text = text.replace("\r", "\n")

    # 连续空格压缩
    text = re.sub(r"[ \t]+", " ", text)

    # 清理每行首尾空格
    lines = []
    for line in text.split("\n"):
        line = line.strip()
        if line:
            lines.append(line)

    return "\n".join(lines).strip()


def read_docx_text(doc_path):
    """
    直接从 DOCX ZIP 中读取 word/document.xml。

    注意：
    DOCX 本质上是 ZIP。
    我们只读取 document.xml，不主动读取 word/media 下的图片。

    如果图片 CRC 损坏，只要 document.xml 本身正常，
    就不会影响文字提取。
    """

    doc_path = Path(doc_path)

    if not doc_path.exists():
        raise FileNotFoundError(
            f"找不到 DOCX 文件：{doc_path}"
        )

    try:
        with zipfile.ZipFile(doc_path, "r") as z:
            names = set(z.namelist())

            if "word/document.xml" not in names:
                raise ValueError(
                    f"{doc_path.name} 不是有效的 Word DOCX 文件："
                    "缺少 word/document.xml"
                )

            # 这里只读取 document.xml
            xml_bytes = z.read("word/document.xml")

    except zipfile.BadZipFile as e:
        raise RuntimeError(
            f"DOCX 文件本身不是有效 ZIP：{doc_path.name}"
        ) from e

    except KeyError as e:
        raise RuntimeError(
            f"DOCX 中找不到 word/document.xml：{doc_path.name}"
        ) from e

    except Exception as e:
        raise RuntimeError(
            f"读取 DOCX 失败：{doc_path.name}\n{e}"
        ) from e

    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError as e:
        raise RuntimeError(
            f"word/document.xml 无法解析：{doc_path.name}"
        ) from e

    paragraphs = []

    # 按 Word 段落读取
    for paragraph in root.findall(".//w:p", NS):

        parts = []

        for node in paragraph.iter():

            tag = node.tag

            if tag == f"{{{W_NS}}}t":
                if node.text:
                    parts.append(node.text)

            elif tag == f"{{{W_NS}}}tab":
                parts.append("\t")

            elif tag == f"{{{W_NS}}}br":
                parts.append("\n")

            elif tag == f"{{{W_NS}}}cr":
                parts.append("\n")

        paragraph_text = "".join(parts)
        paragraph_text = clean_text(paragraph_text)

        if paragraph_text:
            paragraphs.append(paragraph_text)

    return paragraphs


# ============================================================
# 文本处理
# ============================================================

def normalize_line(line):
    """进一步清理单行文字。"""

    line = line.strip()

    # 去除常见 Word 自动编号残留
    line = re.sub(r"^[•·▪●○◆◇■□]+\s*", "", line)

    # 清理连续空格
    line = re.sub(r"\s+", " ", line)

    return line.strip()


def is_noise_line(line):
    """判断是否是明显无用的噪声。"""

    if not line:
        return True

    # 页码
    if re.fullmatch(r"\d+", line):
        return True

    # 单独的装饰符号
    if re.fullmatch(r"[-—_~·•●○◆◇■□*]+", line):
        return True

    return False


def clean_paragraphs(paragraphs):
    result = []

    for paragraph in paragraphs:

        paragraph = normalize_line(paragraph)

        if is_noise_line(paragraph):
            continue

        result.append(paragraph)

    return result


# ============================================================
# 标题识别
# ============================================================

def is_chapter_heading(text):
    """
    判断一段文字是否可能是章节标题。
    """

    if not text:
        return False

    patterns = [
        r"^第[一二三四五六七八九十百千万0-9]+[章节篇部分]",
        r"^[一二三四五六七八九十]+[、.．]",
        r"^[0-9]+[、.．]",
        r"^[0-9]+\.[0-9]+",
        r"^[A-Z]\.[A-Z]?\s",
    ]

    for pattern in patterns:
        if re.match(pattern, text):
            return True

    # 长度较短且没有句号，通常可能是标题
    if len(text) <= 30 and not re.search(r"[。！？；]", text):
        return True

    return False


# ============================================================
# 问题生成
# ============================================================

def make_question(text, subject, chapter=""):
    """
    根据原始文字生成一个基础主动回忆题。

    这里保持数据结构简单稳定，
    前端可以继续根据 type / answer 等字段使用。
    """

    text = clean_text(text)

    if not text:
        return None

    question = {
        "question": text,
        "answer": text,
        "subject": subject,
        "chapter": chapter,
        "type": "recall",
    }

    return question


def split_long_text(text, max_length=180):
    """
    对特别长的段落进行适度切分。
    """

    text = clean_text(text)

    if len(text) <= max_length:
        return [text]

    # 优先按中文句号、分号切
    pieces = re.split(r"(?<=[。！？；])", text)

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
                # 太长时强制切割
                for i in range(0, len(piece), max_length):
                    chunk = piece[i:i + max_length]
                    if chunk:
                        result.append(chunk)

                current = ""

    if current:
        result.append(current)

    return result


# ============================================================
# 单本书解析
# ============================================================

def parse_book(book_id, subject, filename):

    doc_path = ROOT / filename

    print("=" * 60)
    print(f"正在处理：{subject}")
    print(f"文件：{filename}")
    print("=" * 60)

    if not doc_path.exists():
        print(f"WARNING：文件不存在，跳过：{doc_path}")
        return {
            "id": book_id,
            "subject": subject,
            "source": filename,
            "chapters": [],
            "questions": [],
            "paragraphs": [],
        }

    # 关键：这里不再使用 Document(doc_path)
    paragraphs = read_docx_text(doc_path)

    paragraphs = clean_paragraphs(paragraphs)

    print(f"成功读取文字段落：{len(paragraphs)}")

    chapters = []
    questions = []

    current_chapter = ""

    for paragraph in paragraphs:

        # 判断章节标题
        if is_chapter_heading(paragraph):

            current_chapter = paragraph

            if paragraph not in chapters:
                chapters.append(paragraph)

            continue

        # 普通正文
        pieces = split_long_text(paragraph)

        for piece in pieces:

            question = make_question(
                piece,
                subject,
                current_chapter
            )

            if question:
                questions.append(question)

    print(f"识别章节：{len(chapters)}")
    print(f"生成题目：{len(questions)}")

    return {
        "id": book_id,
        "subject": subject,
        "source": filename,
        "chapters": chapters,
        "questions": questions,
        "paragraphs": paragraphs,
    }


# ============================================================
# 去重
# ============================================================

def deduplicate_questions(questions):

    result = []
    seen = set()

    for q in questions:

        key = (
            q.get("subject", ""),
            q.get("chapter", ""),
            q.get("question", ""),
        )

        if key in seen:
            continue

        seen.add(key)
        result.append(q)

    return result


# ============================================================
# 主程序
# ============================================================

def main():

    print("")
    print("=" * 70)
    print("CHECKY DATA BUILDER")
    print("=" * 70)
    print("模式：直接读取 DOCX 的 document.xml")
    print("不会读取 word/media 图片")
    print("")

    books = []
    all_questions = []

    for book in BOOKS:

        try:

            parsed = parse_book(*book)

            parsed["questions"] = deduplicate_questions(
                parsed.get("questions", [])
            )

            books.append(parsed)

            all_questions.extend(parsed["questions"])

        except Exception as e:

            print("")
            print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
            print(f"处理 {book[1]} 时出现错误：")
            print(e)
            print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
            print("")

            # 这里不让一本书的错误导致整个程序直接崩溃
            books.append({
                "id": book[0],
                "subject": book[1],
                "source": book[2],
                "chapters": [],
                "questions": [],
                "paragraphs": [],
                "error": str(e),
            })

    all_questions = deduplicate_questions(all_questions)

    data = {
        "version": "2.0",
        "generated": True,
        "books": books,
        "questions": all_questions,
        "stats": {
            "book_count": len(books),
            "question_count": len(all_questions),
        },
    }

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)

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
    print(f"输出文件：{OUTPUT}")
    print(f"书籍数量：{len(books)}")
    print(f"题目数量：{len(all_questions)}")
    print("=" * 70)


if __name__ == "__main__":
    main()
