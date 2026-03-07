# Word Format App

按 `/Users/mao/Desktop/苏水储〔2018〕1号(规范文件材料格式通知).pdf` 中的要求，对 `.docx` 文档执行基础公文格式整理。

## 当前能力

- 上传 `.docx` 并生成新的整理后文档
- 返回格式检查报告，列出段落识别结果、处理动作和待人工复核项
- 统一 A4 页面和页边距
- 统一标题、一级标题、二级标题、正文、发文字号的字体与字号
- 统一正文行距为 29.5 磅
- “附件：...” 自动按左空 2 字并在正文下空一行处理
- 统一数字和英文为 `Times New Roman`
- 自动写入 4 号宋体 `- 页码 -` 样式页脚，当前默认靠右

## 启动

```bash
cd /Users/mao/Documents/Codex/word-format-app
python3 -m pip install -r requirements.txt
uvicorn main:app --reload
```

浏览器打开 [http://127.0.0.1:8000](http://127.0.0.1:8000)。

## 最简单的公网部署

推荐方案：

- 前端静态页面部署到 Netlify
- Python/FastAPI 后端部署到 Render

### 1. 部署后端到 Render

- 把 `word-format-app` 目录上传到 GitHub
- 在 Render 新建 `Web Service`
- 选择这个仓库
- Build Command 填：`pip install -r requirements.txt`
- Start Command 填：`uvicorn main:app --host 0.0.0.0 --port $PORT`

也可以直接使用仓库里的 [render.yaml](/Users/mao/Documents/Codex/word-format-app/render.yaml)。

部署完成后，记下 Render 给你的后端地址，例如：

```text
https://word-format-app-api.onrender.com
```

### 2. 部署前端到 Netlify

- 打开 [config.js](/Users/mao/Documents/Codex/word-format-app/static/config.js)
- 把 `API_BASE_URL` 改成你的 Render 地址

例如：

```js
window.APP_CONFIG = {
  API_BASE_URL: "https://word-format-app-api.onrender.com",
};
```

- 在 Netlify 里部署这个项目时，把发布目录设为 `static`

仓库里已经提供了 [netlify.toml](/Users/mao/Documents/Codex/word-format-app/netlify.toml)，可直接使用。

## 说明

- 当前规则主要基于段落内容识别标题层级，适合常见通知、公文、方案、报告。
- 结构层次序数按现有文本识别和保留，不会自动把错误编号改写成正确编号。
- 当前接口先返回 JSON 报告和下载地址，再单独下载整理后的 `.docx`。
- 盖章避让、装订、双面印刷偏差等要求不属于 `.docx` 内部版式对象，当前版本不自动处理。
- 如果本机未安装“方正小标宋_GBK / 方正仿宋_GBK / 方正黑体_GBK / 方正楷体_GBK”，Word 会回退到其他字体显示。
