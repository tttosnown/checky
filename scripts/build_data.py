from pathlib import Path
import json
import re
import zipfile
import shutil
import tempfile
from docx import Document

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
DATA_DIR.mkdir(exist_ok=True)
OUTPUT = DATA_DIR / "checky-data.json"

BOOKS = [
    ("history", "建筑历史", "01-《秋季必背册子·建筑历史》.docx"),
    ("urban", "城市设计", "02-《秋季必背册子·城市设计》.docx"),
    ("construction", "建筑构造", "03-《秋季必背册子·建筑构造》.docx"),
]

NUMBERED = re.compile(
    r"^(?:第[一二三四五六七八九十百千万0-9]+[章节篇部]|"
    r"[一二三四五六七八九十百千万]+[、.．]|"
    r"\d+[、.．]|"
    r"[（(][一二三四五六七八九十百千万0-9]+[）)])"
)


def clean(text):
    return re.sub(r"\s+", " ", text or "").strip()


def is_heading_style(style_name):
    if not style_name:
        return False
    s = style_name.lower().replace(" ", "")
    return s.startswith("heading") or s.startswith("标题")


def heading_level(style_name):
    if not style_name:
        return None

    m = re.search(r"(\d+)", style_name)

    if m and (
        style_name.lower().startswith("heading")
        or "标题" in style_name
    ):
        return int(m.group(1))

    return None


def looks_like_heading(text):
    if not text or len(text) > 80:
        return False

    if NUMBERED.match(text):
        return True

    if len(text) <= 28 and not re.search(
        r"[。！？；：,:，]$",
        text
    ):
        return True

    return False


def make_id(*parts):
    raw = "-".join(parts)
    return re.sub(
        r"[^a-zA-Z0-9_-]+",
        "-",
        raw
    ).strip("-").lower()


def extract_keywords(title, content):
    text = f"{title} {content}"
    cues = []

    candidates = [
        "背景", "概念", "定义", "目的", "意义", "特征", "特点", "原则",
        "类型", "分类", "组成", "构成", "方法", "策略", "过程", "影响",
        "问题", "作用", "优点", "缺点", "人物", "作品", "案例", "材料",
        "结构", "构造", "设计", "空间", "形式", "功能", "发展", "思想"
    ]

    for c in candidates:
        if c in text and c not in cues:
            cues.append(c)

        if len(cues) >= 6:
            break

    if not cues:
        cues = [title]

    return cues


def repair_docx(path):
    """
    Try to repair a DOCX whose internal ZIP has a bad CRC.

    The document text/XML is preserved.
    Corrupted image files are skipped if they cannot be read.
    python-docx can then open the repaired document.
    """

    try:
        test_zip = zipfile.ZipFile(path, "r")
        test_zip.testzip()
        test_zip.close()

        # ZIP is healthy.
        return path

    except (zipfile.BadZipFile, RuntimeError, OSError) as e:
        print(f"ZIP problem detected in {path.name}: {e}")
        print("Attempting to repair DOCX...")

    temp_dir = Path(tempfile.mkdtemp(prefix="checky_docx_"))
    repaired_path = temp_dir / path.name

    try:
        # First extract entries manually.
        with zipfile.ZipFile(path, "r") as zin:
            with zipfile.ZipFile(
                repaired_path,
                "w",
                compression=zipfile.ZIP_DEFLATED
            ) as zout:

                for info in zin.infolist():
                    try:
                        data = zin.read(info.filename)

                    except (zipfile.BadZipFile, RuntimeError, OSError) as e:
                        # A corrupted image does not affect the text
                        # extraction pipeline, so omit it.
                        if (
                            info.filename.startswith("word/media/")
                            or info.filename.startswith("media/")
                        ):
                            print(
                                f"Skipping corrupted media: "
                                f"{info.filename}"
                            )
                            continue

                        # If a non-media file is corrupted,
                        # we should not silently continue.
                        raise RuntimeError(
                            f"Cannot repair non-media DOCX file: "
                            f"{info.filename}"
                        ) from e

                    zout.writestr(info, data)

        # Verify repaired file.
        with zipfile.ZipFile(repaired_path, "r") as z:
            bad = z.testzip()

            if bad:
                raise RuntimeError(
                    f"Repaired DOCX still has CRC error: {bad}"
                )

        print(f"Repaired: {path.name}")
        return repaired_path

    except Exception:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise


def parse_book(book_id, title, filename):
    path = ROOT / filename

    if not path.exists():
        return {
            "id": book_id,
            "title": title,
            "source": filename,
            "chapters": [],
            "error": f"File not found: {filename}",
        }

    # Repair DOCX if its internal ZIP has CRC problems.
    doc_path = repair_docx(path)

    doc = Document(doc_path)

    paragraphs = []

    for p in doc.paragraphs:
        text = clean(p.text)

        if not text:
            continue

        style = p.style.name if p.style else ""

        paragraphs.append((text, style))

    chapters = []
    current_chapter = None
    current_item = None

    def ensure_chapter(ch_title, idx):
        nonlocal current_chapter

        current_chapter = {
            "id": make_id(
                book_id,
                "chapter",
                str(idx),
                ch_title
            ),
            "title": ch_title,
            "items": [],
        }

        chapters.append(current_chapter)

        return current_chapter

    for idx, (text, style) in enumerate(
        paragraphs,
        start=1
    ):
        lvl = heading_level(style)
        numbered = bool(NUMBERED.match(text))

        if (
            lvl == 1
            or re.match(
                r"^第[一二三四五六七八九十百千万0-9]+[章节篇部]",
                text
            )
        ):
            ensure_chapter(
                text,
                len(chapters) + 1
            )

            current_item = None
            continue

        if current_chapter is None:
            ensure_chapter("未分类", 1)

        if (
            lvl in (2, 3)
            or (numbered and len(text) <= 80)
            or (
                looks_like_heading(text)
                and len(text) <= 32
            )
        ):
            current_item = {
                "id": make_id(
                    book_id,
                    "item",
                    str(len(current_chapter["items"]) + 1),
                    text
                ),
                "title": text,
                "keywords": [],
                "content": "",
            }

            current_chapter["items"].append(
                current_item
            )

            continue

        if current_item is None:
            current_item = {
                "id": make_id(
                    book_id,
                    "item",
                    str(len(current_chapter["items"]) + 1),
                    current_chapter["title"]
                ),
                "title": current_chapter["title"],
                "keywords": [],
                "content": "",
            }

            current_chapter["items"].append(
                current_item
            )

        current_item["content"] = (
            current_item["content"]
            + "\n"
            + text
        ).strip()

    for chapter in chapters:
        for item in chapter["items"]:
            item["keywords"] = extract_keywords(
                item["title"],
                item["content"]
            )

    return {
        "id": book_id,
        "title": title,
        "source": filename,
        "chapters": chapters,
    }


def main():
    books = [
        parse_book(*b)
        for b in BOOKS
    ]

    payload = {
        "version": 1,
        "generatedBy": "CHECKY Word pipeline",
        "books": books,
    }

    OUTPUT.write_text(
        json.dumps(
            payload,
            ensure_ascii=False,
            indent=2
        ),
        encoding="utf-8",
    )

    print(f"Generated {OUTPUT}")


if __name__ == "__main__":
    main()
