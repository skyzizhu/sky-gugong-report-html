# Sky Gugong Report HTML

这个目录当前使用的核心能力是 `sky-gugong-report-html`：

- 将 Word `.docx` 日报/报告转换为故宫风格的响应式 HTML 页面
- 输出完整站点目录：
  - `index.html`
  - `css/styles.css`
  - `js/main.js`
  - `images/`
- 保留 Word 中的标题、正文、表格和嵌入图片
- 适配移动端展示

## 使用方式

运行下面的命令，把 Word 文件转换成 HTML 站点目录：

```bash
python3 /Users/fushan/.codex/skills/sky-gugong-report-html/scripts/build_gugong_report.py 输入文件.docx 输出目录
```

例如：

```bash
python3 /Users/fushan/.codex/skills/sky-gugong-report-html/scripts/build_gugong_report.py /Users/fushan/Desktop/gugong/20260421日报1206期.docx /Users/fushan/Desktop/gugong/20260421日报1206期-html
```

生成完成后，输出目录中会包含：

- `index.html`
- `css/styles.css`
- `js/main.js`
- `images/`

可以直接本地打开 `index.html`，也可以放到 Web 服务目录下通过 `localhost` 访问。

## 内容规则

生成 HTML 时遵循以下规则：

- 必须按照 Word 大纲和原始排版顺序输出
- 不得遗漏 Word 中的正文、标题、表格、图片
- 不得添加 Word 原文之外的报告内容
- 图片必须提取到单独目录，并通过相对路径链接到 HTML
- 移动端必须可读、可见，不依赖不稳定的滚动触发才能显示正文

## 图片提取与压缩规则

当 Word `.docx` 中包含嵌入图片时，处理流程如下：

1. 从 Word 文件中识别并提取所有嵌入图片。
2. 判断每张图片的原始文件大小。
3. 如果图片小于 `400KB`：
   - 原样保留，不做压缩。
4. 如果图片大于等于 `400KB`：
   - 执行网页友好优化；
   - 仅当最长边超过 `2000px` 时，按比例缩小；
   - 不允许裁切图片，不允许改变长宽比例；
   - 无透明通道的图片优先保存为优化后的 JPEG，质量为 `86`；
   - 带透明通道的图片优先保留为 PNG 并进行优化；
   - 如果优化后文件反而更大，则回退为原图。
5. 所有最终图片统一输出到生成目录下的 `images/` 文件夹。
6. `index.html` 中通过相对路径引用这些图片。

## 当前示例

当前目录中的示例输入/输出：

- 输入 Word：
  [20260421日报1206期.docx](/Users/fushan/Desktop/gugong/20260421日报1206期.docx)
- 输出 HTML：
  [20260421日报1206期-html/index.html](/Users/fushan/Desktop/gugong/20260421日报1206期-html/index.html)

