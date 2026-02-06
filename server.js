const express = require('express');
const multer = require('multer');
const mammoth = require('mammoth');
const TurndownService = require('turndown');
const { marked } = require('marked');
const { Document, Packer, Paragraph, TextRun, HeadingLevel } = require('docx');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = 3000;

// Configure multer for file uploads
const storage = multer.memoryStorage();
const upload = multer({ storage });

// Serve static files
app.use(express.static('public'));
app.use(express.json());

// Initialize Turndown for HTML to Markdown conversion
const turndownService = new TurndownService({
    headingStyle: 'atx',
    codeBlockStyle: 'fenced'
});

// API: Convert DOCX to Markdown
app.post('/api/docx-to-md', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: '请上传文件' });
        }

        // Convert docx to HTML using mammoth
        const result = await mammoth.convertToHtml({ buffer: req.file.buffer });
        const html = result.value;

        // Convert HTML to Markdown using Turndown
        const markdown = turndownService.turndown(html);

        res.json({ 
            success: true, 
            markdown,
            messages: result.messages
        });
    } catch (error) {
        console.error('Conversion error:', error);
        res.status(500).json({ error: '转换失败: ' + error.message });
    }
});

// Parse markdown and create docx paragraphs
function parseMarkdownToDocxElements(markdown) {
    const lines = markdown.split('\n');
    const elements = [];
    let inCodeBlock = false;
    let codeBlockContent = [];

    for (let line of lines) {
        // Handle code blocks
        if (line.startsWith('```')) {
            if (inCodeBlock) {
                elements.push(new Paragraph({
                    children: [new TextRun({ text: codeBlockContent.join('\n'), font: 'Consolas' })],
                    spacing: { before: 100, after: 100 }
                }));
                codeBlockContent = [];
            }
            inCodeBlock = !inCodeBlock;
            continue;
        }

        if (inCodeBlock) {
            codeBlockContent.push(line);
            continue;
        }

        // Skip empty lines but add spacing
        if (line.trim() === '') {
            elements.push(new Paragraph({ text: '' }));
            continue;
        }

        // Handle headings
        if (line.startsWith('# ')) {
            elements.push(new Paragraph({
                text: line.substring(2),
                heading: HeadingLevel.HEADING_1
            }));
        } else if (line.startsWith('## ')) {
            elements.push(new Paragraph({
                text: line.substring(3),
                heading: HeadingLevel.HEADING_2
            }));
        } else if (line.startsWith('### ')) {
            elements.push(new Paragraph({
                text: line.substring(4),
                heading: HeadingLevel.HEADING_3
            }));
        } else if (line.startsWith('#### ')) {
            elements.push(new Paragraph({
                text: line.substring(5),
                heading: HeadingLevel.HEADING_4
            }));
        } else if (line.startsWith('- ') || line.startsWith('* ')) {
            // Handle list items
            elements.push(new Paragraph({
                children: [new TextRun({ text: '• ' + line.substring(2) })]
            }));
        } else if (/^\d+\. /.test(line)) {
            // Handle numbered lists
            elements.push(new Paragraph({
                children: [new TextRun({ text: line })]
            }));
        } else {
            // Handle regular paragraphs with inline formatting
            const children = parseInlineFormatting(line);
            elements.push(new Paragraph({ children }));
        }
    }

    return elements;
}

// Parse inline formatting (bold, italic, code)
function parseInlineFormatting(text) {
    const runs = [];
    let currentText = '';
    let i = 0;

    while (i < text.length) {
        // Bold (**text**)
        if (text.substring(i, i + 2) === '**') {
            if (currentText) {
                runs.push(new TextRun({ text: currentText }));
                currentText = '';
            }
            const endIndex = text.indexOf('**', i + 2);
            if (endIndex !== -1) {
                runs.push(new TextRun({ text: text.substring(i + 2, endIndex), bold: true }));
                i = endIndex + 2;
                continue;
            }
        }
        
        // Italic (*text* or _text_)
        if ((text[i] === '*' || text[i] === '_') && text[i - 1] !== '*' && text[i + 1] !== '*') {
            const marker = text[i];
            if (currentText) {
                runs.push(new TextRun({ text: currentText }));
                currentText = '';
            }
            const endIndex = text.indexOf(marker, i + 1);
            if (endIndex !== -1 && text[endIndex - 1] !== marker) {
                runs.push(new TextRun({ text: text.substring(i + 1, endIndex), italics: true }));
                i = endIndex + 1;
                continue;
            }
        }

        // Inline code (`code`)
        if (text[i] === '`') {
            if (currentText) {
                runs.push(new TextRun({ text: currentText }));
                currentText = '';
            }
            const endIndex = text.indexOf('`', i + 1);
            if (endIndex !== -1) {
                runs.push(new TextRun({ text: text.substring(i + 1, endIndex), font: 'Consolas' }));
                i = endIndex + 1;
                continue;
            }
        }

        currentText += text[i];
        i++;
    }

    if (currentText) {
        runs.push(new TextRun({ text: currentText }));
    }

    return runs.length > 0 ? runs : [new TextRun({ text })];
}

// API: Convert Markdown to DOCX
app.post('/api/md-to-docx', upload.single('file'), async (req, res) => {
    try {
        let markdown;
        
        if (req.file) {
            markdown = req.file.buffer.toString('utf-8');
        } else if (req.body.markdown) {
            markdown = req.body.markdown;
        } else {
            return res.status(400).json({ error: '请上传文件或提供Markdown内容' });
        }

        const elements = parseMarkdownToDocxElements(markdown);

        const doc = new Document({
            sections: [{
                properties: {},
                children: elements
            }]
        });

        const buffer = await Packer.toBuffer(doc);

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', 'attachment; filename="converted.docx"');
        res.send(buffer);
    } catch (error) {
        console.error('Conversion error:', error);
        res.status(500).json({ error: '转换失败: ' + error.message });
    }
});

app.listen(PORT, () => {
    console.log(`🚀 服务器运行在 http://localhost:${PORT}`);
});
