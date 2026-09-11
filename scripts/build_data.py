from pathlib import Path
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from html import unescape

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT = DATA_DIR / "checky-data.json"

BOOKS = [
    ("history", "建筑历史", "01-《秋季必背册子·建筑历史》.docx"),
    ("urban", "城市设计", "02-《秋季必背册子·城市设计》.docx"),
    ("physics", "建筑物理", "03-《秋季必背册子·建筑物理》.docx"),
    ("structure", "构造技术", "04-《秋季必背册子·构造技术》.docx"),
]

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

def clean_text(s):
    s = unescape(s or "")
    s = s.replace("\u00a0"," ").replace("\u200b","").replace("\ufeff","")
    s = s.replace("\r\n","\n").replace("\r","\n")
    s = re.sub(r"[ \t]+"," ",s)
    return "\n".join(x.strip() for x in s.split("\n") if x.strip()).strip()

def read_docx_text(path):
    with zipfile.ZipFile(path, "r") as z:
        if "word/document.xml" not in z.namelist():
            raise RuntimeError(f"{path.name} 缺少 word/document.xml")
        # 关键：只读取文字 XML，不碰 word/media/*，绕过损坏图片 CRC
        xml = z.read("word/document.xml")
    root = ET.fromstring(xml)
    out = []
    for p in root.findall(".//w:p", NS):
        parts = []
        for n in p.iter():
            if n.tag == f"{{{W_NS}}}t" and n.text:
                parts.append(n.text)
            elif n.tag == f"{{{W_NS}}}tab":
                parts.append(" ")
            elif n.tag in (f"{{{W_NS}}}br", f"{{{W_NS}}}cr"):
                parts.append("\n")
        s = clean_text("".join(parts))
        if s:
            out.append(s)
    return out

def is_noise(s):
    return not s or re.fullmatch(r"\d+", s) or re.fullmatch(r"[-—_~·•●○◆◇■□*]+", s)

def is_heading(s):
    if not s or len(s) > 80:
        return False
    patterns = [
        r"^第[一二三四五六七八九十百千万0-9]+[章节篇部分]",
        r"^[一二三四五六七八九十百千万]+[、.．]",
        r"^[0-9]+[、.．]",
        r"^[0-9]+\.[0-9]+",
        r"^[（(][一二三四五六七八九十百千万]+[）)]",
    ]
    return any(re.match(p, s) for p in patterns)

def split_text(s, limit=220):
    s = clean_text(s)
    if len(s) <= limit:
        return [s] if s else []
    pieces = re.split(r"(?<=[。！？；])", s)
    result, cur = [], ""
    for p in pieces:
        p = p.strip()
        if not p:
            continue
        if len(cur) + len(p) <= limit:
            cur += p
        else:
            if cur:
                result.append(cur)
            if len(p) <= limit:
                cur = p
            else:
                for i in range(0, len(p), limit):
                    result.append(p[i:i+limit])
                cur = ""
    if cur:
        result.append(cur)
    return result

def make_question(book_id, subject, chapter_id, chapter_name, n, text):
    return {
        "id": f"{book_id}-q-{n}",
        "subject": book_id,
        "subjectName": subject,
        "chapter": chapter_id,
        "chapterId": chapter_id,
        "chapterName": chapter_name,
        "type": "recall",
        "question": text,
        "answer": text,
        "content": text,
    }

def parse_book(book_id, subject, filename):
    path = ROOT / filename
    if not path.exists():
        raise FileNotFoundError(f"找不到：{filename}")

    raw = [x for x in read_docx_text(path) if not is_noise(x)]
    numbered_headings = [x for x in raw if is_heading(x)]

    # 如果一本书没有可靠章节标题，也绝不生成“0题目”。
    # 自动建立一个“全文内容”章节，把所有正文纳入题库。
    chapters, questions = [], []
    current = None
    qn = 0
    cn = 0

    for line in raw:
        if is_heading(line):
            cn += 1
            current = {
                "id": f"{book_id}-chapter-{cn}",
                "name": line,
                "title": line,
                "subject": book_id,
                "subjectName": subject,
                "questions": [],
                "questionIds": [],
            }
            chapters.append(current)
            continue

        if current is None:
            cn += 1
            current = {
                "id": f"{book_id}-chapter-{cn}",
                "name": "全文内容",
                "title": "全文内容",
                "subject": book_id,
                "subjectName": subject,
                "questions": [],
                "questionIds": [],
            }
            chapters.append(current)

        for piece in split_text(line):
            qn += 1
            q = make_question(
                book_id, subject, current["id"], current["name"], qn, piece
            )
            current["questions"].append(q)
            current["questionIds"].append(q["id"])
            questions.append(q)

    # 极端情况下仍然保证至少有一个章节
    if not questions and raw:
        cn += 1
        current = {
            "id": f"{book_id}-chapter-fallback",
            "name": "全文内容",
            "title": "全文内容",
            "subject": book_id,
            "subjectName": subject,
            "questions": [],
            "questionIds": [],
        }
        chapters = [current]
        for line in raw:
            for piece in split_text(line):
                qn += 1
                q = make_question(
                    book_id, subject, current["id"], current["name"], qn, piece
                )
                current["questions"].append(q)
                current["questionIds"].append(q["id"])
                questions.append(q)

    for c in chapters:
        c["count"] = len(c["questions"])
        c["questionCount"] = len(c["questions"])

    print(f"{subject}: 段落={len(raw)} 章节={len(chapters)} 题目={len(questions)}")
    return {
        "id": book_id,
        "subject": book_id,
        "subjectName": subject,
        "name": subject,
        "title": subject,
        "source": filename,
        "chapters": chapters,
        "questions": questions,
        "count": len(questions),
        "questionCount": len(questions),
    }

def dedupe(items):
    seen, out = set(), []
    for q in items:
        k = (q.get("subject"), q.get("chapter"), q.get("question"))
        if k not in seen:
            seen.add(k)
            out.append(q)
    return out

def main():
    books, all_questions, all_chapters = [], [], []

    for spec in BOOKS:
        try:
            book = parse_book(*spec)
        except Exception as e:
            print(f"[ERROR] {spec[1]}: {e}")
            book = {
                "id": spec[0],
                "subject": spec[0],
                "subjectName": spec[1],
                "name": spec[1],
                "title": spec[1],
                "source": spec[2],
                "chapters": [],
                "questions": [],
                "count": 0,
                "questionCount": 0,
                "error": str(e),
            }
        books.append(book)
        all_chapters.extend(book["chapters"])
        all_questions.extend(book["questions"])

    all_questions = dedupe(all_questions)

    # 重新建立每个章节的题目引用，保证前端点击章节一定能找到题目
    qmap = {q["id"]: q for q in all_questions}
    for book in books:
        for c in book["chapters"]:
            c["questionIds"] = [i for i in c.get("questionIds", []) if i in qmap]
            c["questions"] = [qmap[i] for i in c["questionIds"]]
            c["count"] = len(c["questions"])
            c["questionCount"] = len(c["questions"])

    data = {
        "version": "3.0",
        "generated": True,
        "books": books,
        "subjects": books,
        "chapters": all_chapters,
        "questions": all_questions,
        "stats": {
            "bookCount": len(books),
            "chapterCount": len(all_chapters),
            "questionCount": len(all_questions),
        },
    }

    with OUTPUT.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print("=" * 60)
    print("CHECKY BUILD SUCCESS")
    print(f"章节：{len(all_chapters)}")
    print(f"题目：{len(all_questions)}")
    print(f"输出：{OUTPUT}")
    print("=" * 60)

if __name__ == "__main__":
    main()
