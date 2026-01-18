# OneNote to Markdown Exporter

这是一个 Python 脚本，用于将 Microsoft OneNote 笔记本导出为 Markdown 格式，方便迁移到 Obsidian、Notion 等笔记工具。

## 功能特点

- 📚 **批量导出**：自动遍历所有笔记本、分区和页面。
- 🖼 **图片/附件下载**：自动下载页面中的图片和附件，并将其转换为本地相对路径引用。
- 📝 **Markdown 转换**：将 OneNote 的 HTML 内容转换为标准的 Markdown 格式。
- 🔄 **增量更新**：自动跳过已存在的完整笔记，支持断点续传。
- 🚀 **并发控制**：内置重试机制和限流处理（429 Too Many Requests），防止 API 请求过于频繁。

## 依赖环境

- Python 3.6+
- 依赖库：`requests`, `msal`, `beautifulsoup4`, `markdownify`

## 安装与使用

1. **克隆或下载本仓库**

2. **安装依赖**

   ```bash
   pip install -r requirements.txt
   ```

3. **运行脚本**

   ```bash
   python onenote_export.py
   ```

4. **认证登录**
   
   脚本运行后会显示一个 Microsoft 登录代码和网址。
   - 打开浏览器访问提示的 URL (https://microsoft.com/devicelogin)。
   - 输入终端显示的 `user_code`。
   - 登录你的 Microsoft 账号并授权应用访问 OneNote。

5. **等待导出**
   
   登录成功后，脚本会自动开始下载笔记。导出内容将保存在当前目录下的 `OneNote_Export` 文件夹中。

## 目录结构

导出后的结构如下：
```
OneNote_Export/
├── 笔记本名称/
│   ├── 分区名称/
│   │   ├── 页面标题.md
│   │   └── assets/
│   │       ├── image1.png
│   │       └── attachment.pdf
```

## 注意事项

- **API 限制**：Microsoft Graph API 有速率限制。脚本已内置自动等待机制，如果遇到 429 错误会自动暂停并重试，请耐心等待。
- **内容完整性**：脚本会尝试处理大部分 OneNote 格式，但复杂的布局（如复杂的表格嵌套）可能在转换 Markdown 时会有所损耗。
- **隐私安全**：脚本仅在本地运行，Token 仅用于访问你的 OneNote 数据，不会上传任何数据到第三方服务器。

## License

MIT License
