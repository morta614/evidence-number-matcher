# 投诉取证单号匹配工具

## 📖 简介

这是一套用于自动化处理投诉取证截图的工具链，可以：
1. 解压取证压缩包并预览截图
2. 批量 OCR 识别投诉单号和日期
3. 自动匹配 Excel 总表并标记来源

## 🛠️ 工具组成

### 1. 取证预览工具（`evidence_viewer.py`）
- **功能**：Web 界面预览压缩包中的所有截图
- **用途**：勾选需要识别的图片
- **输出**：`selected_images.csv`

### 2. OCR 识别工具（`ocr_extractor.py`）
- **功能**：批量 OCR 识别截图
- **提取内容**：截图名称、反馈产品、作品名称、投诉单号、日期
- **输出**：`complaint_records.xlsx`

### 3. Excel 匹配工具（`excel_matcher.py`）
- **功能**：匹配总表和 OCR 结果
- **输出**：带【取证来源】列的总表

## 🚀 快速开始

### 安装依赖

```bash
# Python 依赖
pip install streamlit pillow pandas openpyxl pytesseract

# 系统依赖（CentOS/RHEL）
yum install -y tesseract tesseract-langpack-chi_sim

# 系统依赖（Ubuntu/Debian）
apt install -y tesseract-ocr tesseract-ocr-chi-sim

# Windows：下载安装 Tesseract OCR
# 1. 访问 https://github.com/UB-Mannheim/tesseract/wiki
# 2. 下载 tesseract-ocr-w64-setup-5.x.x.exe
# 3. 安装时勾选中文语言包
# 4. 设置环境变量：TESSDATA_PREFIX=安装路径\tessdata
```

### 使用流程

```bash
# 1. 启动 Web 预览工具
streamlit run 匹配单号工具/evidence_viewer.py

# 2. 在浏览器中上传压缩包、勾选截图、导出 CSV

# 3. OCR 识别
python 匹配单号工具/ocr_extractor.py --input selected_images.csv --output complaint_records.xlsx

# 4. 匹配总表（可选）
python 匹配单号工具/excel_matcher.py --master 总表.xlsx --ocr complaint_records.xlsx
```

## 📋 详细说明

请查看 `使用说明.md` 获取完整的操作指南和故障排除。

## ⚠️ 注意事项

1. 支持 .zip 格式的压缩包
2. OCR 识别准确率取决于截图清晰度
3. 确保总表中的日期格式为 `YYYY-MM-DD`
4. 建议分批处理大量图片（>1000 张）

## 📞 问题反馈

如有问题，请提交 Issue 或联系开发者。
