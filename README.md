# 取证单号匹配工具

## 📋 简介

一套用于自动化处理投诉取证截图的工具，支持批量 OCR 识别和 Excel 匹配。

## 🎯 功能

### 三大核心功能

1. **取证预览**：上传压缩包，预览并勾选需要识别的截图
2. **OCR 识别**：批量识别截图中的投诉单号、日期、反馈产品、作品名称
3. **Excel 匹配**：匹配总表和 OCR 结果，标记取证来源

### 特点

- ✅ Web 界面，操作简单
- ✅ 支持批量处理
- ✅ 仅按单号匹配，忽略日期差异
- ✅ 输出取证来源和识别日期两列
- ✅ 一键安装，自动配置

## 🚀 快速开始

### 方式一：一键安装包（推荐）

1. 下载 `取证单号匹配工具-一键安装包.zip`
2. 解压缩到任意目录
3. 双击运行 `一键安装.bat`
4. 安装完成后，双击运行 `启动工具.bat`
5. 浏览器自动打开 http://localhost:8501

### 方式二：手动安装

#### 1. 安装 Python

- 访问：https://www.python.org/downloads/
- 下载 Python 3.7 或更高版本
- 安装时**勾选 "Add Python to PATH"**
- 安装完成后**重启电脑**

#### 2. 安装依赖

```bash
pip install streamlit pillow pandas openpyxl pytesseract tqdm
```

或使用清华镜像（国内更快）：

```bash
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

#### 3. 安装 Tesseract OCR

- 下载：https://github.com/UB-Mannheim/tesseract/wiki
- 下载文件：`tesseract-ocr-w64-setup-5.4.0.exe`
- 双击运行安装程序
- **重要**：勾选 "Chinese (Simplified)" 语言包
- 安装到：`C:\Program Files\Tesseract-OCR`
- 设置环境变量：
  - 系统变量 - 新建：
    - 变量名：`TESSDATA_PREFIX`
    - 变量值：`C:\Program Files\Tesseract-OCR\tessdata`
  - 系统变量 - Path - 编辑 - 新建：
    - `C:\Program Files\Tesseract-OCR`
- **重启电脑**

#### 4. 启动工具

```bash
streamlit run evidence_viewer.py
```

## 📖 使用教程

### 标签页 1：取证预览

1. 上传压缩包（.zip 格式）
2. 自动解压并显示所有图片
3. 点击按钮选择需要的图片（✓ = 已选）
4. 点击「导出选中图片」

### 标签页 2：OCR 识别

1. 自动读取选中的图片
2. 点击「开始 OCR 识别」
3. 等待识别完成（进度条显示）
4. 下载识别结果（Excel）

### 标签页 3：Excel 匹配

1. 上传投诉总表（Excel）
2. 选择投诉单号列
3. 点击「开始匹配」
4. 下载匹配结果

## 📊 匹配结果

匹配完成后，总表会新增 **2 列**：

| 列名 | 内容 | 示例 |
|------|------|------|
| **取证来源** | 来源压缩包 - 截图名称 | 20251010投诉侵权记录.zip - 20251010151435010疑似诈骗.png |
| **识别日期** | OCR 识别出的投诉日期 | 2025-10-10 |

## 🛠️ 技术栈

- **Python**：3.7+
- **Web 框架**：Streamlit
- **OCR 引擎**：Tesseract OCR 5.4.0
- **数据处理**：Pandas + OpenPyXL
- **图像处理**：Pillow

## 🎯 系统要求

| 项目 | 要求 |
|------|------|
| 操作系统 | Windows 10/11 (64位）|
| Python | 3.7 或更高版本 |
| 内存 | 4GB 或更多（推荐 8GB+）|
| 磁盘空间 | 2GB 或更多 |

## 🔧 常见问题

### Q1: 安装时提示 "未检测到 Python"

**A:** 请先安装 Python 3.7 或更高版本，并确保勾选 "Add Python to PATH"

### Q2: Tesseract 安装后还是提示未安装

**A:** 运行 `设置环境变量.bat`，然后**重启电脑**

### Q3: OCR 识别不出来中文

**A:** 确保 Tesseract 安装时勾选了 "Chinese (Simplified)" 语言包

### Q4: 浏览器打不开

**A:** 手动在浏览器中访问 http://localhost:8501

### Q5: 匹配不上单号

**A:** 检查总表中的单号格式和 OCR 识别的单号格式是否一致

## 📄 文件说明

| 文件 | 说明 |
|------|------|
| `evidence_viewer.py` | 主程序，包含三个标签页的功能 |
| `requirements.txt` | Python 依赖包列表 |
| `.gitignore` | Git 忽略文件配置 |

## 🚩 更新日志

### v1.0.0 (2026-03-05)
- 初始版本发布
- 支持取证预览、OCR 识别、Excel 匹配三大功能
- 支持单号匹配（忽略日期）
- 匹配后输出取证来源和识别日期两列
- 提供一键安装包

## 📞 技术支持

如有问题，请检查：
1. Python 版本是否 >= 3.7
2. Tesseract 是否正确安装并设置环境变量
3. 端口 8501 是否被占用

## 📜 许可证

本项目仅供个人学习使用。

## ⚠️ 注意事项

1. **仅支持 Windows 系统**
2. 需要管理员权限运行安装脚本
3. 首次使用建议测试小批量图片
4. 大批量处理建议分批进行

---

**开发团队**：墨斗 (Modou)
**版本**：v1.0.0
**最后更新**：2026-03-05
