# PDF 转 Word 工具

将扫描版 PDF 文件转换为保持原始排版的 Word 文档，支持 OCR 文字识别和印章图片保留。

## 功能特性

- OCR 文字识别（RapidOCR）- 识别扫描版 PDF 中的文字
- 智能段落合并 - 根据缩进、字号、对齐方式合并 OCR 识别的碎片文本
- 印章图片保留 - 提取 PDF 中的印章图像并精确定位到 Word 中
- 精确排版还原 - 映射 PDF 坐标到 Word 页面，保持原始布局
- 署名和日期定位 - 准确还原签名人、日期的缩进和垂直位置

## 依赖

- PyMuPDF (fitz) - PDF 解析和图像提取
- python-docx - Word 文档生成
- RapidOCR - OCR 文字识别
- Pillow / NumPy - 图像处理

## 使用方法

```bash
python pdf2word.py
```

将待转换的 PDF 文件放在同一目录下，运行脚本即可生成对应的 Word 文档。

## 技术要点

- 印章提取：从 PDF 中提取 ImageMask 类型的印章图像，自动检测颜色反转并还原
- 印章定位：使用页面绝对坐标（relativeFrom: page）精确定位印章位置
- 段落合并：基于缩进变化、字号一致性、对齐方式等特征合并 OCR 碎片段落
- 垂直间距：根据 PDF 中段落的 Y 坐标差计算 Word 中的段前/段后间距
