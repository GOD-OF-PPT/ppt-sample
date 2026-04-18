# 试卷 PDF → PPTX 转换器

将试卷 PDF 自动转换为 PowerPoint 文件，每道题独占一张 Slide，多页题目并排展示。

## 效果

- 自动识别题目编号（支持 `1.` `1、` `第1题` 等格式）
- 每道题一张 Slide，跨页题目按页并排截图
- 自动裁剪四周空白，内容填满 Slide
- 支持文字层 PDF 和扫描版 PDF（OCR 自动检测题号）

---

## 环境要求

- Python 3.9+
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract)（处理扫描版 PDF 时需要）

---

## 安装步骤

### 1. 安装 Tesseract

**macOS**
```bash
brew install tesseract
```

**Ubuntu / Debian**
```bash
sudo apt install tesseract-ocr
```

**Windows**

下载安装包：https://github.com/UB-Mannheim/tesseract/wiki
安装后将 Tesseract 目录（如 `C:\Program Files\Tesseract-OCR`）添加到系统 PATH。

---

### 2. 克隆项目

```bash
git clone https://github.com/GOD-OF-PPT/ppt-sample.git
cd ppt-sample
```

### 3. 安装 Python 依赖

```bash
pip install -r requirements.txt
```

---

## 启动服务

```bash
python app.py
```

服务默认运行在 `http://localhost:5000`，浏览器打开即可使用。

---

## 使用方法

1. 打开 `http://localhost:5000`
2. 点击上传区域或将 PDF 文件拖入
3. 点击 **开始转换**
4. 转换完成后点击 **下载 PPTX 文件**

---

## 题号识别规则

| 格式 | 示例 |
|------|------|
| 数字 + 句点 | `1.` `2.` |
| 数字 + 顿号 | `1、` `2、` |
| 中文题号 | `第1题` `第 2 题` |

若 PDF 为扫描件（无文字层），程序自动切换为 OCR 识别模式。识别不到题号时，回退为每页一张 Slide。
