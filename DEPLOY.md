# Hugging Face Spaces 部署说明

## 为什么不继续用 GitHub Pages

新版正式入口需要运行 Flask 后端，处理 `/api/upload`，并调用 `excel_image_extractor.py` 从 Excel 文件中提取内嵌照片和企业微信二维码。GitHub Pages 只能托管静态 HTML/CSS/JS，不能运行 Python、不能提供 Flask API，也不能调用后端图片提取逻辑。

因此 GitHub Pages 可以保留为旧版静态页或跳转页，但不适合作为新版正式入口。

## 为什么使用 Hugging Face Spaces

Hugging Face Spaces 支持免费的 Docker 部署，可以运行 Flask、gunicorn、Chromium、poppler-utils 和中文字体。这样线上版本可以和本地 Flask 版本保持一致，并支持后端解析 Excel 内嵌图片。

免费版 Spaces 会休眠。长时间无人访问后，首次打开可能需要等待容器重新启动。

## 如何部署

1. 在 Hugging Face 创建新的 Space。
2. 选择 SDK 为 Docker。
3. 将本仓库内容同步到 Space 仓库。
4. 确认 `README.md` 顶部包含 Spaces 配置，`sdk: docker`，`app_port: 7860`。
5. Space 会使用 `Dockerfile` 构建镜像，并通过以下命令启动：

```bash
gunicorn app:app --bind 0.0.0.0:${PORT:-7860} --timeout 120 --workers 1
```

当前 `ARTIFACTS` 是进程内存字典，所以 workers 必须保持为 1。否则上传后生成的预览和下载内容可能落在不同 worker，导致找不到 artifact。

## 如何测试上传 Excel

1. 打开 Space 页面。
2. 点击“上传 Excel 模板”或拖拽 `.xlsx` 文件到上传区。
3. 使用包含“姓名、职务、工号、照片、企业微信二维码”列的 Excel。
4. 确认页面出现工牌预览。
5. 检查照片和企业微信二维码是否被正确填入工牌。
6. 测试打印、PDF 下载和 PNG ZIP 下载。

如果 PDF 生成失败，优先检查 Docker 构建日志中 Chromium 是否安装成功，以及 `GOOGLE_CHROME_BIN=/usr/bin/chromium` 是否生效。如果 PNG ZIP 为空，检查 `poppler-utils` 是否安装成功。

## 域名处理

`work-card.mole.today` 可以先作为跳转页保留，跳转到 Hugging Face Space 地址。等新版稳定后，再决定是否将正式入口完全迁移到 Space 或其他 Python Web 服务。
