# EquationRescue

**在浏览器本地，把现有 DOCX 里的显式 LaTeX 转成 Word 原生公式标记。**

选择文件，逐条查看源码与近似预览，勾选需要转换的公式，下载新副本。不支持的写法保留原文，并说明原因。无需账号、后端、API Key 或模型调用；文档内容不会上传。

**v0.1 为实验版。** 面向简单讲义与笔记，不是通用 LaTeX 转换器。请保留原文件，并在 Word 中检查输出的排版和公式可编辑性。

[在线试用](https://mozzie49.github.io/equation-rescue/) · [CI 与线上检查](https://github.com/mozzie49/equation-rescue/actions/workflows/ci.yml) · [English](README.md) · [支持范围](docs/supported-syntax.md) · [验证记录](docs/validation.md)

## 快速开始

需要 Node 22.12+ 或 Node 24。

```sh
npm ci
npm run dev
```

打开 Vite 显示的本地地址，点击“试用教学示例”，或选择自己的 `.docx`。界面支持中英文切换。下载的 JSON 报告包含文件名和公式源码，请像原文档一样保管。

```sh
npm run check
npm run build
npm run preview
```

`dist/` 可部署到普通静态文件服务器。依赖和字体都在构建产物中，无 CDN、分析统计或上传接口。首次载入页面和示例会向本站请求静态资源；上传文件的处理在浏览器中完成。这不是 PWA，也未提供离线安装包。

## 小而明确的范围

- 识别正文普通段落和表格单元格中的 `\(...\)`、`\[...\]`、`$$...$$`
- 支持分数、平方根、上下标、希腊字母、基本运算符、求和与连乘
- 处理被 Word 拆成多个简单文本 run 的公式，保留周围文字的格式
- 逐条转换或跳过，不认识的语法不会被猜测性改写
- 仅修改主文档 XML；其他 ZIP 成员的局部文件头和压缩数据保持原样
- 未选择任何公式时，返回与输入逐字节一致的副本

显示公式必须独占一个段落。单个 `$` 不参与识别，以免把价格当公式。矩阵、宏、LaTeX 包、`\text`、`\left`/`\right`、带次数的根式等暂不支持。

含已有公式、书签、域、链接、对象、XML 注释或复杂结构的段落会原样保留。页眉、页脚、脚注、文本框等不修复。含隐藏文字、带修订、启用修订、受保护、数字签名、宏、加密、旧 DOC、扫描件、ZIP64 和严格 OOXML 文档不在范围内。

转换后的公式可能改变行高和分页；保留格式并不等于页面像素完全一致。不要把此工具当作恶意文件扫描器：原有嵌入内容和外部关系仍然保留。

## 真正做过的验证

原创教学讲义包含 8 个候选项：7 个可转换公式和 1 个故意不支持的矩阵。源码、转换后的 DOCX、报告和真实 LibreOffice 渲染图都在仓库中。结构测试与 LibreOffice 渲染**不能证明所有 Microsoft Word 版本均正确显示或编辑**。

```sh
npm test
npm run build
npx playwright install chromium
npm run test:e2e
```

GitHub Actions 在部署前运行 62 项单元测试、生产构建、生产依赖审计和 9 项 Chromium 浏览器测试；部署后再访问真实 Pages 地址，检查 DOCX 与报告下载内容。实际执行结果详见[验证记录](docs/validation.md)。

## 来源与许可证

本项目没有发明 LaTeX 转 OMML 或 DOCX 修复算法。[AfterMath](https://github.com/axobase001/aftermath) 已经提供 MIT 许可的 Python 命令行方案。本项目尝试的是“现有文件、浏览器本地、逐项确认”的工作流。

采用 fflate、saxes、xmlchars 和 KaTeX，许可证见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。窄范围解析器和讲义为本项目原创，开发使用了 AI 辅助。软件采用 MIT 许可证。
