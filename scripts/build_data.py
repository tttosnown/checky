from pathlib import Path
import json
import zipfile
import re
import xml.etree.ElementTree as ET


# ============================================================
# 基础路径
# ============================================================

ROOT = Path(__file__).resolve().parents[1]

DATA_DIR = ROOT / "data"
DATA_DIR.mkdir(exist_ok=True)

OUTPUT = DATA_DIR / "checky-data.json"


# ============================================================
# 四本资料
# 文件名必须与仓库中的 DOCX 完全一致
# ============================================================

BOOKS = [
    (
        "history",
        "建筑历史",
        "01-《秋季必背册子·建筑历史》.docx"
    ),
    (
        "urban",
        "城市设计",
        "02-《秋季必背册子·城市设计》.docx"
    ),
    (
        "physics",
        "建筑物理",
        "03-《秋季必背册子·建筑物理》.docx"
    ),
    (
        "construction",
        "构造技术",
        "04-《秋季必背册子·构造技术》.docx"
    ),
]


# ============================================================
# 读取 DOCX
#
# 重要：
# DOCX 本质上是 ZIP。
#
# 之前的问题是：
# word/media/image21.png
# 出现 CRC 错误。
#
# 所以这里完全不读取图片。
# 只读取：
#
# word/document.xml
#
# 这样不会再因为图片损坏导致整个构建失败。
# ============================================================

def read_docx_paragraphs(path):

    with zipfile.ZipFile(path, "r") as archive:

        xml_data = archive.read(
            "word/document.xml"
        )

    root = ET.fromstring(xml_data)

    namespace = {
        "w":
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    }

    paragraphs = []

    for paragraph in root.findall(
        ".//w:p",
        namespace
    ):

        texts = []

        for text_node in paragraph.findall(
            ".//w:t",
            namespace
        ):

            if text_node.text:
                texts.append(
                    text_node.text
                )

        text = "".join(texts)

        # 清理多余空格
        text = re.sub(
            r"\s+",
            " ",
            text
        ).strip()

        if text:
            paragraphs.append(text)

    return paragraphs


# ============================================================
# 判断是不是章节标题
# ============================================================

def is_heading(text):

    text = text.strip()

    if len(text) < 2:
        return False

    if len(text) > 100:
        return False


    patterns = [

        # 第 一章 / 第1章 / 第一节
        r"^第[一二三四五六七八九十百千万0-9]+[章节篇]",

        # 一、二、三、
        r"^[一二三四五六七八九十百千万]+[、.．]",

        # 1. / 1、 / 1.2 / 1.2.3
        r"^\d+([.、．]\d+)*[、.．]?\s*\S+",

        # 常见章节标题
        r"^(绪论|概述|目录|附录|总结)",

        # 第一部分
        r"^第[一二三四五六七八九十]+部分",

        # 第一篇
        r"^第[一二三四五六七八九十]+篇",

    ]


    for pattern in patterns:

        if re.match(
            pattern,
            text
        ):
            return True


    return False


# ============================================================
# 从正文生成主动回忆题
#
# 这里不是简单显示正文。
#
# 每一段内容都会成为一个可点击的主动回忆题。
# ============================================================

def make_questions(
    lines,
    subject_name,
    chapter_name,
    chapter_id
):

    questions = []


    for line in lines:

        line = line.strip()

        if len(line) < 10:
            continue


        # 太长的段落拆开
        if len(line) > 450:

            pieces = re.split(
                r"[。！？；]",
                line
            )

            pieces = [
                p.strip()
                for p in pieces
                if len(p.strip()) >= 18
            ]

        else:

            pieces = [line]


        for piece in pieces:

            piece = piece.strip()

            if len(piece) < 10:
                continue


            question_number = (
                len(questions) + 1
            )


            question_id = (
                f"{chapter_id}-q{question_number}"
            )


            # 提取一些中文关键词
            keyword_list = re.findall(
                r"[\u4e00-\u9fff]{2,8}",
                piece
            )


            keywords = "；".join(
                keyword_list[:10]
            )


            question = {

                "id":
                    question_id,

                "subject":
                    subject_name,

                "subjectName":
                    subject_name,

                "chapter":
                    chapter_name,

                "chapterId":
                    chapter_id,

                "chapterName":
                    chapter_name,

                "type":
                    "recall",

                "question":
                    "请主动回忆并说明：" +
                    piece[:120],

                "answer":
                    piece,

                "content":
                    piece,

                "keyword":
                    keywords

            }


            questions.append(
                question
            )


            # 防止极端情况下生成过多数据
            if len(questions) >= 3000:

                return questions


    return questions


# ============================================================
# 主程序
# ============================================================

all_books = []

all_chapters = []

all_questions = []


for (
    book_id,
    book_name,
    filename
) in BOOKS:


    file_path =
        ROOT / filename


    if not file_path.exists():

        print(
            "找不到文件：",
            filename
        )

        continue


    print()
    print(
        "正在读取：",
        book_name
    )


    # --------------------------------------------------------
    # 读取正文
    # --------------------------------------------------------

    paragraphs = read_docx_paragraphs(
        file_path
    )


    print(
        "读取到段落：",
        len(paragraphs)
    )


    # --------------------------------------------------------
    # 根据标题切分章节
    # --------------------------------------------------------

    chapter_groups = []

    current_chapter = None


    for paragraph in paragraphs:


        if is_heading(
            paragraph
        ):

            current_chapter = {

                "name":
                    paragraph,

                "lines":
                    []

            }

            chapter_groups.append(
                current_chapter
            )


        else:

            if current_chapter is not None:

                current_chapter[
                    "lines"
                ].append(
                    paragraph
                )


    # --------------------------------------------------------
    # 如果没有识别出任何章节
    #
    # 不允许变成 0。
    #
    # 整本书至少建立一个“全文内容”章节。
    # --------------------------------------------------------

    if not chapter_groups:

        chapter_groups = [

            {
                "name":
                    "全文内容",

                "lines":
                    paragraphs
            }

        ]


    # --------------------------------------------------------
    # 建立书籍对象
    # --------------------------------------------------------

    book = {

        "id":
            book_id,

        "name":
            book_name,

        "title":
            book_name,

        "chapters":
            []

    }


    # --------------------------------------------------------
    # 建立章节
    # --------------------------------------------------------

    for index, group in enumerate(
        chapter_groups,
        start=1
    ):


        chapter_id = (
            f"{book_id}-c{index}"
        )


        chapter_name = (
            group["name"]
        )


        # ----------------------------------------------------
        # 生成题目
        # ----------------------------------------------------

        questions = make_questions(

            group["lines"],

            book_name,

            chapter_name,

            chapter_id

        )


        # ----------------------------------------------------
        # 如果章节正文太少
        #
        # 也保留章节。
        # ----------------------------------------------------

        chapter = {

            "id":
                chapter_id,

            "bookId":
                book_id,

            "subjectId":
                book_id,

            "name":
                chapter_name,

            "title":
                chapter_name,

            "questionIds":
                [
                    q["id"]
                    for q in questions
                ],

            "questions":
                questions,

            "count":
                len(questions),

            "questionCount":
                len(questions)

        }


        book["chapters"].append(
            chapter
        )

        all_chapters.append(
            chapter
        )

        all_questions.extend(
            questions
        )


    all_books.append(
        book
    )


    print(
        book_name,
        "章节：",
        len(book["chapters"]),
        "题目：",
        sum(
            len(c["questions"])
            for c in book["chapters"]
        )
    )


# ============================================================
# 最终 JSON
# ============================================================

result = {

    "version":
        "2.0",

    "books":
        all_books,

    "subjects":
        all_books,

    "chapters":
        all_chapters,

    "questions":
        all_questions,

    "stats": {

        "books":
            len(all_books),

        "chapters":
            len(all_chapters),

        "questions":
            len(all_questions)

    }

}


# ============================================================
# 写入文件
# ============================================================

OUTPUT.write_text(

    json.dumps(
        result,
        ensure_ascii=False,
        indent=2
    ),

    encoding="utf-8"

)


print()
print(
    "======================================"
)

print(
    "CHECKY 数据生成完成"
)

print(
    "书籍：",
    len(all_books)
)

print(
    "章节：",
    len(all_chapters)
)

print(
    "题目：",
    len(all_questions)
)

print(
    "输出：",
    OUTPUT
)

print(
    "======================================"
)
