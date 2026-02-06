# DOC2MD：Word 文档与 Markdown 格式转换器

一个本地运行的文档格式转换工具，支持 Word 文档 (DOC/DOCX) 与 MD 格式的双向转换。

## ✨ 功能特性

- 📝 **DOCX → Markdown**：将 Word 文档转换为 Markdown 格式
- 📄 **Markdown → DOCX**：将 Markdown 内容转换为 Word 文档
- 🎨 现代化深色主题界面
- 📁 支持拖拽上传文件
- 🔒 完全本地运行，数据安全

## 🚀 快速开始

### 安装

```bash
# 进入项目目录
cd doc2md

# 安装依赖
npm install
```

### 运行

```bash
npm start
```

启动后访问 http://localhost:3000

## 📖 使用说明

### DOCX 转 Markdown

1. 打开"DOCX → Markdown"标签页
2. 拖拽或点击选择 `.docx` 文件
3. 点击"转换为 Markdown"按钮
4. 复制结果或下载为 `.md` 文件

### Markdown 转 DOCX

1. 切换到"Markdown → DOCX"标签页
2. 输入或上传 Markdown 内容
3. 点击"转换为 DOCX"按钮
4. 自动下载 Word 文档

## 🛠 技术栈

- **后端**: Node.js + Express
- **DOCX 解析**: mammoth.js
- **Markdown 转换**: turndown / docx

## 📝 支持的格式

- 标题 (H1-H4)
- 加粗、斜体
- 有序/无序列表
- 代码块和行内代码
- 中文内容

## 📄 许可证

MIT License
