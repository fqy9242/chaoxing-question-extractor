// ==UserScript==
// @name         超星学习通作业题目提取工具（背题模式-自动转字母版）
// @namespace    http://tampermonkey.net/
// @version      1.7
// @description  提取超星学习通作业题目，支持“题目(答案)”格式的Word导出，自动将文字答案转换为选项字母(A/B)，自动去除原题括号
// @author       Assistant
// @match        *://*.chaoxing.com/mooc-ans/mooc2/work/view*
// @match        *://*.chaoxing.com/exam-ans*
// @match        *://mooc1.chaoxing.com/exam-ans*
// @grant        none
// @require      https://unpkg.com/docx@7.8.2/build/index.js
// @require      https://unpkg.com/file-saver@2.0.5/dist/FileSaver.min.js
// ==/UserScript==

(function() {
    'use strict';

    // === 基础工具函数 ===
    function waitForElement(selector, timeout = 10000) {
        return new Promise((resolve, reject) => {
            const startTime = Date.now();
            function check() {
                const element = document.querySelector(selector);
                if (element) resolve(element);
                else if (Date.now() - startTime > timeout) reject(new Error(`元素 ${selector} 未找到`));
                else setTimeout(check, 100);
            }
            check();
        });
    }

    // === 数据结构 ===
    class ImageInfo {
        constructor(src, alt = '', width = '', height = '') {
            this.src = src; this.alt = alt; this.width = width; this.height = height;
        }
    }

    class Question {
        constructor() {
            this.type = ''; this.number = ''; this.content = ''; this.contentImages = [];
            this.options = []; this.optionImages = [];
            this.myAnswer = ''; this.correctAnswer = ''; this.score = '';
            this.isCorrect = false; this.analysis = ''; this.analysisImages = [];
        }
    }

    // === 题目解析器 ===
    class QuestionParser {
        constructor() { this.questions = []; }

        parseAllQuestions() {
            this.questions = [];
            document.querySelectorAll('.questionLi').forEach((container, index) => {
                try {
                    const q = this.parseQuestion(container, index + 1);
                    if (q) this.questions.push(q);
                } catch (e) { console.error(e); }
            });
            return this.questions;
        }

        extractImages(element) {
            const images = [];
            if (!element) return images;
            element.querySelectorAll('img').forEach(img => {
                const src = img.getAttribute('src') || img.getAttribute('data-original');
                if (src) images.push(new ImageInfo(src, img.alt, img.width, img.height));
            });
            return images;
        }

        getTextContent(element) {
            if (!element) return '';
            const cloned = element.cloneNode(true);
            cloned.querySelectorAll('img').forEach((img, i) => {
                const span = document.createElement('span');
                span.textContent = `[图片${i + 1}]`;
                img.parentNode.replaceChild(span, img);
            });
            // 移除多余空白
            return cloned.textContent.replace(/\s+/g, ' ').trim();
        }

        parseQuestion(container, index) {
            const q = new Question();

            // 1. 编号与类型
            const titleEl = container.querySelector('.mark_name');
            if (titleEl) {
                const text = titleEl.textContent.trim();
                const typeMatch = text.match(/\((.*?题)\)/);
                const numMatch = text.match(/^(\d+)\./);
                q.type = typeMatch ? typeMatch[1] : '未知题型';
                q.number = numMatch ? numMatch[1] : index.toString();
            }

            // 2. 内容
            const contentEl = container.querySelector('.qtContent');
            if (contentEl) {
                q.content = this.getTextContent(contentEl);
                q.contentImages = this.extractImages(contentEl);
            }

            // 3. 选项
            container.querySelectorAll('.mark_letter li, .qtDetail li').forEach((opt, i) => {
                const text = this.getTextContent(opt);
                // 移除选项前的 "A. " 或 "A "，只保留内容以便后续比对
                // 但为了Word显示，我们这里存原始文本，比对时再处理
                if (text) {
                    q.options.push(text);
                    const imgs = this.extractImages(opt);
                    if (imgs.length) q.optionImages.push({ optionIndex: i, images: imgs });
                }
            });

            // 4. 答案处理
            const getAns = (sel) => {
                const arr = [];
                container.querySelectorAll(sel).forEach(el => {
                    const t = el.textContent.replace(/\s+/g, ' ').trim();
                    if(t) arr.push(t);
                });
                return arr.join(',');
            };

            q.myAnswer = getAns('.stuAnswerContent');
            q.correctAnswer = getAns('.rightAnswerContent');

            // 清理答案前缀
            const cleanAns = (ans) => ans.replace(/正确答案[:：]\s*/g, '').replace(/我的答案[:：]\s*/g, '').trim();
            q.myAnswer = cleanAns(q.myAnswer);
            q.correctAnswer = cleanAns(q.correctAnswer);

            // 5. 解析与分数
            const analysisEl = container.querySelector('.qtAnalysis');
            if (analysisEl) {
                q.analysis = this.getTextContent(analysisEl);
                q.analysisImages = this.extractImages(analysisEl);
            }
            const scoreEl = container.querySelector('.totalScore i');
            if (scoreEl) q.score = scoreEl.textContent.trim();
            q.isCorrect = !!container.querySelector('.marking_dui');

            return q;
        }

        getWorkTitle() {
            const el = document.querySelector('.mark_title');
            return el ? el.textContent.replace(/\s+/g, ' ').trim() : '作业题目';
        }

        getStatistics() {
            const stats = { totalQuestions: this.questions.length, correctCount: 0, totalScore: '0', maxScore: '0', totalImages: 0 };
            stats.correctCount = this.questions.filter(q => q.isCorrect).length;
            this.questions.forEach(q => stats.totalImages += (q.contentImages.length + q.analysisImages.length));
            return stats;
        }
    }

    // === Word 生成器 ===
    class WordGenerator {
        constructor() {
            this.docx = window.docx;
            this.saveAs = window.saveAs || saveAs;
            this.imageCache = new Map();
        }

        async downloadImage(url) {
            if (this.imageCache.has(url)) return this.imageCache.get(url);
            try {
                const res = await fetch(url);
                const blob = await res.blob();
                this.imageCache.set(url, blob);
                return blob;
            } catch (e) { return null; }
        }

        // === 核心逻辑：智能答案转换 ===
        // 将 "对" 转为 "A"，将 "文字选项" 转为 "A/B/C"
        convertAnswerToLetter(answer, options, type) {
            if (!answer) return " ";
            answer = answer.trim();

            // 1. 如果已经是 A, B, C, D 或者 A,B 这种格式，直接返回
            // 简单判断：全是大写字母、逗号、空格组成，且长度较短
            if (/^[A-Z\s,]+$/.test(answer) && answer.length < 10) {
                return answer;
            }

            // 2. 判断题处理
            if (type.includes('判断')) {
                if (['对', '正确', '√', 'True'].some(k => answer.includes(k))) return 'A';
                if (['错', '错误', '×', 'False'].some(k => answer.includes(k))) return 'B';
            }

            // 3. 文本匹配处理 (针对选择题提取出是文字的情况)
            // 尝试在选项中寻找答案文本
            if (options && options.length > 0) {
                // 有些多选题答案是 "文字A, 文字B"
                const ansParts = answer.split(/[,，;；]/);
                const matchedLetters = [];

                for (let part of ansParts) {
                    part = part.trim();
                    if (!part) continue;

                    let foundIndex = -1;
                    for (let i = 0; i < options.length; i++) {
                        // 选项通常是 "A. 内容" 或 "A 内容" 或直接 "内容"
                        // 我们去掉开头的字母和点，进行纯内容比对
                        let optContent = options[i].replace(/^[A-Z][\.\s、]\s*/, '').trim();

                        // 全等 或者 包含 (防止标点符号差异)
                        if (optContent === part || (optContent.length > 2 && optContent.includes(part)) || (part.length > 2 && part.includes(optContent))) {
                            foundIndex = i;
                            break;
                        }
                    }

                    if (foundIndex !== -1) {
                        matchedLetters.push(String.fromCharCode(65 + foundIndex));
                    } else {
                        // 没找到匹配，就保留原文，防止丢失信息
                        matchedLetters.push(part);
                    }
                }

                if (matchedLetters.length > 0) {
                    // 如果匹配到了字母，排序并返回 (如 B,A -> A,B)
                    if (matchedLetters.every(l => /^[A-Z]$/.test(l))) {
                        return matchedLetters.sort().join('');
                    }
                    return matchedLetters.join(',');
                }
            }

            // 兜底：如果实在转换不了，返回原样
            return answer;
        }

        // === 核心逻辑：背题模式生成 ===
        async generateImportWord(questions, title) {
            const children = [];

            children.push(new this.docx.Paragraph({
                children: [new this.docx.TextRun({ text: title, bold: true, size: 32 })],
                alignment: this.docx.AlignmentType.CENTER,
                spacing: { after: 400 }
            }));

            for (const q of questions) {
                // 1. 获取原始答案
                let rawAnswer = q.correctAnswer || q.myAnswer || " ";

                // 2. 转换为字母 (A/B/C/D)
                let finalAnswer = this.convertAnswerToLetter(rawAnswer, q.options, q.type);

                // 3. 处理题目内容：去除原有的括号 ( ) （ ）
                // 逻辑：移除中文或英文括号，且括号内为空或仅含空格
                // 或者直接移除末尾的括号区域，因为通常填空在末尾
                let cleanContent = q.content;
                // 替换所有空括号
                cleanContent = cleanContent.replace(/(\s*[（(]\s*[)）]\s*)+/g, ' ');
                // 去除首尾空白
                cleanContent = cleanContent.trim();

                // 4. 拼接标题行: 9. 题目内容(A)
                children.push(new this.docx.Paragraph({
                    children: [
                        new this.docx.TextRun({
                            text: `${q.number}. ${cleanContent}`,
                            bold: true
                        }),
                        new this.docx.TextRun({
                            text: `(${finalAnswer})`,
                            bold: true,
                            color: "FF0000" // 红色答案
                        })
                    ],
                    spacing: { before: 100, after: 60 }
                }));

                // 5. 图片
                if (q.contentImages.length > 0) {
                    const imgs = await this.createImagesParagraphs(q.contentImages, "", true);
                    children.push(...imgs);
                }

                // 6. 选项
                if (q.options.length > 0) {
                    for (const opt of q.options) {
                        children.push(new this.docx.Paragraph({
                            children: [ new this.docx.TextRun({ text: opt }) ],
                            indent: { left: 0 },
                            spacing: { after: 0, line: 240 }
                        }));
                    }
                }

                // 7. 分隔空行
                children.push(new this.docx.Paragraph({ text: "" }));
            }

            return new this.docx.Document({
                sections: [{ properties: {}, children: children }]
            });
        }

        // 详细版生成逻辑 (保持功能完整性)
        async generateWord(questions, title, stats, options) {
            const children = [];
            children.push(new this.docx.Paragraph({
                children: [new this.docx.TextRun({ text: title + " (解析版)", bold: true, size: 32 })],
                alignment: this.docx.AlignmentType.CENTER,
                spacing: { after: 400 }
            }));

            for (const q of questions) {
                children.push(new this.docx.Paragraph({
                    children: [new this.docx.TextRun({ text: `${q.number}. [${q.type}] ${q.content}`, bold: true })],
                    spacing: { before: 200, after: 100 }
                }));

                q.options.forEach(opt => {
                    children.push(new this.docx.Paragraph({
                        children: [new this.docx.TextRun({ text: opt })],
                        indent: { left: 400 },
                        spacing: { after: 60 }
                    }));
                });

                const ansChildren = [];
                if (options.includeCorrectAnswer) ansChildren.push(new this.docx.TextRun({ text: `正确答案: ${q.correctAnswer}  `, color: "009900", bold: true }));
                if (options.includeMyAnswer) ansChildren.push(new this.docx.TextRun({ text: `我的答案: ${q.myAnswer}  `, color: "000000" }));

                if (ansChildren.length > 0) children.push(new this.docx.Paragraph({ children: ansChildren, spacing: { before: 100 } }));

                if (options.includeAnalysis && q.analysis) {
                    children.push(new this.docx.Paragraph({
                        children: [new this.docx.TextRun({ text: `解析: ${q.analysis}`, color: "666666" })],
                        spacing: { after: 200 }
                    }));
                }

                if (options.includeSeparator) children.push(new this.docx.Paragraph({ border: { bottom: { color: "CCCCCC", space: 1, value: "single", size: 6 } } }));
            }

            return new this.docx.Document({ sections: [{ children: children }] });
        }

        async createImagesParagraphs(images, prefix, embed) {
            const paragraphs = [];
            for (const img of images) {
                try {
                    const blob = await this.downloadImage(img.src);
                    if (blob) {
                        paragraphs.push(new this.docx.Paragraph({
                            children: [new this.docx.ImageRun({ data: blob, transformation: { width: 300, height: 200 } })],
                            alignment: this.docx.AlignmentType.CENTER
                        }));
                    }
                } catch (e) {}
            }
            return paragraphs;
        }

        async downloadWord(doc, filename) {
            try {
                const blob = await this.docx.Packer.toBlob(doc);
                saveAs(blob, `${filename}.docx`);
                return true;
            } catch (e) { return false; }
        }
    }

    // === UI 界面 ===
    class ExtractorUI {
        constructor() {
            this.parser = new QuestionParser();
            this.wordGenerator = new WordGenerator();
        }

        createUI() {
            const div = document.createElement('div');
            div.id = 'cx-tool-ui';
            div.innerHTML = `
                <style>
                    #cx-tool-ui { position: fixed; top: 100px; right: 20px; width: 300px; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.15); z-index: 99999; display: none; font-family: sans-serif; }
                    #cx-tool-ui.show { display: block; }
                    .cx-btn { display: block; width: 100%; padding: 10px; margin: 10px 0; border: none; border-radius: 4px; color: white; cursor: pointer; font-size: 14px; }
                    .btn-blue { background: #409EFF; } .btn-blue:hover { background: #66b1ff; }
                    .btn-green { background: #67C23A; } .btn-green:hover { background: #85ce61; }
                    .btn-purple { background: #6f42c1; } .btn-purple:hover { background: #8359d1; }
                    .cx-close { position: absolute; top: 10px; right: 10px; cursor: pointer; color: #999; }
                    .cx-opt { font-size: 12px; color: #666; margin-bottom: 5px; }
                    .cx-tip { font-size: 12px; color: #999; margin-top: 5px; }
                </style>
                <div class="cx-close" onclick="document.getElementById('cx-tool-ui').classList.remove('show')">✕</div>
                <h3 style="margin-top:0">题目提取 v1.7</h3>
                <div id="cx-msg" style="color:#67C23A;font-size:12px;height:20px;"></div>

                <button class="cx-btn btn-blue" onclick="window.cxTool.parse()">1. 解析题目</button>

                <div style="border-top:1px solid #eee; margin:15px 0;"></div>

                <button class="cx-btn btn-purple" id="btn-dl-import" disabled onclick="window.cxTool.downloadImport()">
                    2. 下载 背题模式(Word)
                </button>
                <div class="cx-tip">✨ 自动去括号，文字答案自动转字母(A/B)</div>

                <button class="cx-btn btn-green" id="btn-dl-detail" disabled onclick="window.cxTool.downloadDetail()">
                    3. 下载 详细解析版(Word)
                </button>
            `;
            document.body.appendChild(div);

            // 悬浮球
            const floatBtn = document.createElement('div');
            floatBtn.innerHTML = '📝';
            floatBtn.style.cssText = 'position:fixed;bottom:100px;right:20px;width:50px;height:50px;background:#6f42c1;color:white;border-radius:50%;text-align:center;line-height:50px;font-size:24px;cursor:pointer;z-index:99998;box-shadow:0 2px 10px rgba(0,0,0,0.2);';
            floatBtn.onclick = () => document.getElementById('cx-tool-ui').classList.add('show');
            document.body.appendChild(floatBtn);
        }

        async parse() {
            this.msg('正在解析...');
            try {
                this.questions = this.parser.parseAllQuestions();
                if(!this.questions.length) throw new Error('未找到题目');
                this.msg(`解析成功: ${this.questions.length}题`);
                document.getElementById('btn-dl-import').disabled = false;
                document.getElementById('btn-dl-detail').disabled = false;
            } catch(e) {
                this.msg('解析失败: ' + e.message, 'red');
            }
        }

        async downloadImport() {
            if(!this.questions) return;
            this.msg('正在生成背题文档...');
            const title = this.parser.getWorkTitle();
            const doc = await this.wordGenerator.generateImportWord(this.questions, title);
            await this.wordGenerator.downloadWord(doc, `${title}_背题模式`);
            this.msg('下载成功!');
        }

        async downloadDetail() {
            if(!this.questions) return;
            this.msg('正在生成详细文档...');
            const title = this.parser.getWorkTitle();
            const options = {
                includeCorrectAnswer: true, includeMyAnswer: true,
                includeAnalysis: true, includeSeparator: true
            };
            const doc = await this.wordGenerator.generateWord(this.questions, title, {}, options);
            await this.wordGenerator.downloadWord(doc, `${title}_详细版`);
            this.msg('下载成功!');
        }

        msg(txt, color='#67C23A') {
            const el = document.getElementById('cx-msg');
            el.style.color = color;
            el.textContent = txt;
        }

        init() {
            const check = setInterval(() => {
                if(window.docx && window.saveAs) {
                    clearInterval(check);
                    this.createUI();
                    window.cxTool = this;
                    console.log('Chaoxing Tool Loaded');
                }
            }, 500);
        }
    }

    new ExtractorUI().init();
})();