# CHECKY

CHECKY 是基于三个 Word 册子的主动回忆学习页面。

## 使用方法

1. 将三个 `.docx` 文件放在仓库根目录。
2. 将 `.github/workflows/build-data.yml` 和 `scripts/build_data.py` 放入对应目录。
3. 将 `index.html` 放在仓库根目录。
4. GitHub Actions 会自动读取 Word 并生成 `data/checky-data.json`。
5. GitHub Pages 页面读取该 JSON。

## 注意

当前版本主要用于验证 Word 的章节/知识点结构解析。
关键词为保守的第一版提示，后续可根据实际识别结果继续优化。

评分历史保存在浏览器 localStorage，不写入生成的 JSON，因此重新生成题库不会覆盖你的评分记录。
