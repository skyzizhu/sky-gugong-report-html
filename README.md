# Sky Gugong Report HTML

`sky-gugong-report-html` 用来把 Word `.docx` 日报/报告转换成故宫风格的响应式 HTML 站点。

它会：
- 保留 Word 中的标题、正文、表格和嵌入图片
- 输出完整站点目录，而不是单独一个 HTML 文件
- 适配移动端阅读
- 使用固定的故宫视觉模板生成页面
- 自动为页面写入分享卡片所需的 `meta` 标签
- 自动附带一张默认分享预览图

## 输出目录结构

每次生成的结果都是一个完整目录，结构固定为：

```text
report-folder/
  index.html
  css/
    styles.css
  js/
    main.js
  images/
    ...
```

## 最终目录命名规则

默认情况下，不传第二个参数时，脚本会在 Word 源文件所在目录下自动创建一个以当天日期命名的结果目录：

```text
YYYYMMDD_report
```

例如在 `2026-05-09` 生成时，目录会是：

```text
20260509_report
```

如果同名目录已经存在，不会覆盖旧结果，而是自动顺延：

```text
20260509_report_2
20260509_report_3
...
```

也就是说：
- Word 文件在哪个目录
- 默认生成的网页结果目录就在哪个目录下

## 如何使用

### 1. 默认方式：自动按日期命名输出目录

```bash
python3 /Users/fushan/.codex/skills/sky-gugong-report-html/scripts/build_gugong_report.py 输入文件.docx
```

例如：

```bash
python3 /Users/fushan/.codex/skills/sky-gugong-report-html/scripts/build_gugong_report.py /Users/fushan/Desktop/模板参考（12.26）.docx
```

这会在 `/Users/fushan/Desktop/` 下生成类似：

```text
20260509_report/
```

### 2. 手动指定输出目录

如果你想自己指定目录名，也可以传第二个参数：

```bash
python3 /Users/fushan/.codex/skills/sky-gugong-report-html/scripts/build_gugong_report.py 输入文件.docx 输出目录
```

例如：

```bash
python3 /Users/fushan/.codex/skills/sky-gugong-report-html/scripts/build_gugong_report.py /Users/fushan/Desktop/gugong/20260421日报1206期.docx /Users/fushan/Desktop/gugong/20260421日报1206期-html
```

如果你手动指定的目录已经存在，脚本同样不会覆盖，而是自动生成：

```text
输出目录_2
输出目录_3
...
```

## 生成后怎么用

生成完成后，结果目录中会包含：
- `index.html`
- `css/styles.css`
- `js/main.js`
- `images/`

可以直接本地打开 `index.html`，也可以把整个结果目录放到 Web 服务目录下，通过 `localhost` 访问。

注意：
- 转换完成后的反馈必须同时显示生成后的完整目录和本地入口文件 `index.html`
- 不能只单独拿走 `index.html`
- 必须把整个结果目录一起保留，因为 HTML 会通过相对路径引用 `css/`、`js/` 和 `images/`

## 分享卡片规则

生成 HTML 时，脚本会自动在 `head` 中写入分享卡片需要的元标签，包括：
- `description`
- `og:title`
- `og:description`
- `og:url`
- `og:image`
- `twitter:title`
- `twitter:description`
- `twitter:image`

当前规则如下：
- 分享标题：使用 Word 文档标题
- 分享描述：使用目录概览中的一级目录标题，去掉 `一、二、三…` 这类编号后，用 `、` 拼接
- 分享图片：使用 skill 自带的默认分享图 `assets/share-default.png`

每次生成时，这张默认图片会自动复制到输出目录：

```text
images/share-default.png
```

这样在上传到公网后，别人分享链接时就可以优先显示统一的故宫预览图。

## 上传到阿里云 OSS

仓库提供 `scripts/upload_to_oss.py`，用于把生成后的完整目录上传到阿里云 OSS，并输出可在浏览器直接打开的 `index.html` 链接。

重要流程：
- Word 转 HTML 后，先检查本地生成目录和 `index.html`
- 向用户显示生成后的完整目录和本地入口文件
- 不要自动上传
- 必须先让用户确认是否上传到云端服务器
- 用户确认后，才运行 OSS 上传命令
- 最终交付使用 `final_url`

真实密钥配置放在本地文件中，不提交到 Git：

```text
config/oss_config.json
```

可以从模板复制：

```bash
cp config/oss_config.example.json config/oss_config.json
```

配置字段：

```json
{
  "access_key_id": "YOUR_ACCESS_KEY_ID",
  "access_key_secret": "YOUR_ACCESS_KEY_SECRET",
  "endpoint": "https://oss-cn-beijing.aliyuncs.com",
  "bucket": "gugong-report",
  "public_base_url": "http://your-report-domain.example.com",
  "remote_prefix": "",
  "public_read": true,
  "signed_url_expires_days": 3
}
```

上传示例：

```bash
python3 scripts/upload_to_oss.py /Users/fushan/Desktop/20260513_report
```

如果 `remote_prefix` 为空，脚本会使用本地目录名作为 OSS 前缀。如果配置了 `public_base_url`，最终链接会使用该域名，例如：

```text
http://your-report-domain.example.com/20260513_report/index.html?Expires=...
```

同时，分享卡片里的 `og:url` 和 `og:image` 也会基于 `public_base_url` 生成绝对公网地址。要让微信、企业微信、飞书等平台稳定抓取到预览图，`public_base_url` 需要配置成实际可访问的公网域名，不能只用本地路径或相对路径。

脚本会输出 `public_url` 和 `final_url`。最终交付必须使用 `final_url`，它会带 `response-content-disposition=inline` 签名参数，确保浏览器直接浏览页面，而不是下载 `index.html`。默认有效期为 `signed_url_expires_days` 配置的天数。当前建议配置为 `3`，表示提交上传后三天内可访问，三天后签名过期。

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
原样保留，不做压缩。
4. 如果图片大于等于 `400KB`：
执行网页友好优化。
5. 大图优化时遵循以下规则：
仅当最长边超过 `2000px` 时，按比例缩小。
6. 图片处理限制如下：
不允许裁切图片，不允许改变长宽比例。
7. 无透明通道的图片：
优先保存为优化后的 JPEG，质量为 `86`。
8. 带透明通道的图片：
优先保留为 PNG 并进行优化。
9. 如果优化后文件反而更大：
回退为原图。
10. 所有最终图片统一输出到生成目录下的 `images/` 文件夹。
11. `index.html` 中通过相对路径引用这些图片。

## 当前示例

当前目录中的示例输入/输出：
- 输入 Word：
[20260421日报1206期.docx](/Users/fushan/Desktop/gugong/20260421日报1206期.docx)
- 输出 HTML：
[20260421日报1206期-html/index.html](/Users/fushan/Desktop/gugong/20260421日报1206期-html/index.html)
