# CHECKY 完整替换包

这次不是只替换一个脚本。

需要替换：

- `index.html`
- `scripts/build_data.py`
- `.github/workflows/build-data.yml`

不要删除或重新上传现有 DOCX。

## 这版解决的问题

1. DOCX 中损坏的 `word/media/image*.png` 不再影响读取文字。
2. 建筑物理、构造技术如果没有被可靠识别出章节，不会再生成 0 题目，而会自动建立“全文内容”章节。
3. 前端不再只显示“多少个章节”，点击章节会真正打开该章节。
4. 章节内可以逐条进入主动回忆。
5. 前端直接读取 `data/checky-data.json`。
6. GitHub Actions 自动生成并提交 JSON。

## 替换后

提交这三个文件后：

GitHub → Actions → Build CHECKY data

等待绿色。

然后打开 GitHub Pages。

如果浏览器仍显示旧页面，先强制刷新一次。
