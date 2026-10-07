# 军乐团团员档案生成器

静态 GitHub Pages 网站。填写资料后，浏览器会在本地生成一个 Obsidian 档案 ZIP 包；资料与照片不会上传到服务器或保存到浏览器存储。

ZIP 根目录内含：

- 以“姓名”命名的 Markdown 档案；
- Word 档案表；
- Excel 台账。
- 以“所在声部”命名、内容为该声部名称的 TXT 文件。

接收者将 ZIP 解压后，管理员把 Markdown 文件移至 Obsidian 库对应的声部文件夹即可。

## 本地运行

```bash
npm install
npm run dev
```

## 发布到 GitHub Pages

1. 将此目录推送到 GitHub 仓库的 `main` 分支。
2. 在仓库 **Settings → Pages** 中，将 Source 设为 **GitHub Actions**。
3. 推送后等待 `Deploy GitHub Pages` 工作流完成，页面地址会显示在工作流结果中。

模板 PDF 和中文字体位于 `public/assets/`，必须保留，PDF 生成功能才能正常运行。
