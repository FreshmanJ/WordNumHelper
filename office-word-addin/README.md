# WordNumHelper Office Word Add-in

这个子项目把仓库现有的 DOCX 图表编号整理能力搬进了 Word Office 加载项任务窗格。

## 目录说明

- `taskpane.html`: Word 任务窗格页面
- `app.js`: 复用原项目的扫描、规则校验、重编号和导出逻辑
- `office-bridge.js`: 连接 Word 当前文档读取 / 写回能力
- `manifest.xml`: Office 加载项清单
- `commands.html` / `commands.js`: Ribbon 按钮入口页

## 当前行为

- 桌面版 Word:
  - 可以直接读取当前打开的 Word 文档
  - 扫描后可将修正结果直接写回当前文档
  - 也保留上传 DOCX 和下载修正文件两种方式
- Word 网页版 / 普通浏览器:
  - 保留上传 DOCX、本地预览、下载修正文件
  - 不直接读取当前 Word 文档

## 使用方式

1. 先把 `office-word-addin` 整个目录部署到一个 HTTPS 地址，或者本地用 HTTPS 静态服务托管。
2. 把 [manifest.xml](/E:/FreshmanJ/PersonalProjects/WordNumHelper/office-word-addin/manifest.xml) 里的 `https://localhost:3000` 全部替换成真实地址。
3. 在 Word 中侧载这个 `manifest.xml`。
4. 打开插件后：
   - 桌面版 Word 可点击“读取当前 Word 文档”
   - 或直接上传 `.docx`
   - 设置规则后点击“扫描并预览”
   - 选择“下载修正后的 DOCX”或“写回当前文档”

## 本地校验

```powershell
npm.cmd --prefix office-word-addin run check
```
