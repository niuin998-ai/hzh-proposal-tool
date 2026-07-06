# HZH Proposal Tool Deployment Guide

## 项目简介

HZH Proposal Tool 是一个 Python Streamlit 内部工具，用于生成医旅客户方案。当前功能包括：

- 中文后台界面
- POI 点位管理、筛选、搜索、编辑、软删除
- POI CSV / Excel 导入
- POI 列表导出
- POI 图片上传与管理
- 根据客户需求生成英文线路初稿
- 导出英文 PowerPoint 方案
- 生成可编辑草稿后导出单页客户宣传页 PDF

## 技术栈

- 应用框架：Streamlit
- 语言：Python
- 数据处理：pandas / openpyxl
- PPT 生成：python-pptx
- PDF 生成：ReportLab
- 图片处理：Pillow
- 数据存储：本地 CSV 文件，默认 `data/poi_database.csv`
- 图片存储：本地文件，默认 `assets/poi_images/`
- 前端框架：无独立 React / Next.js / Vite 前端
- 后端服务：Streamlit 内置服务
- 数据库：当前没有外部数据库

## 本地运行

```bash
cd /Users/wuqiong/Documents/Codex/2026-07-01/hzh-proposal-tool
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

本地访问：

```text
http://localhost:8501
```

## 环境变量

当前项目没有 API Key、数据库密码等必须配置的敏感信息。

可选环境变量用于生产环境持久化存储：

| 变量名 | 说明 | 示例 |
|---|---|---|
| `HZH_STORAGE_DIR` | 统一存储目录。设置后数据、导出模板和上传图片默认写到这个目录下 | `/var/data/hzh-proposal-tool` |
| `HZH_DATA_DIR` | POI CSV 数据目录，优先级高于 `HZH_STORAGE_DIR` | `/var/data/hzh-proposal-tool/data` |
| `HZH_OUTPUTS_DIR` | 导出模板目录 | `/var/data/hzh-proposal-tool/outputs` |
| `HZH_POI_IMAGE_DIR` | POI 上传图片目录 | `/var/data/hzh-proposal-tool/assets/poi_images` |

本地可参考 `.env.example`，不要提交真实 `.env` 或 `.streamlit/secrets.toml`。

## 推荐部署方案

### 方案 A：Streamlit Community Cloud，适合快速试用

优点：部署最快，直接连接 GitHub。

限制：本地文件写入不是长期可靠的生产数据库。POI 编辑、图片上传等运行时写入内容，可能在应用重启、重部署或多用户同时操作时丢失或互相覆盖。

适合：内部演示、小范围试用、先让同事访问。

部署步骤：

1. 确保代码已提交并推送到 GitHub。
2. 打开 Streamlit Community Cloud。
3. 选择 GitHub 仓库。
4. Branch 选择 `main`。
5. Main file path 填 `app.py`。
6. Requirements file 使用项目根目录的 `requirements.txt`。
7. 部署完成后打开平台提供的公网 URL。

Build / Start 设置通常不需要手动填写。Streamlit 会自动安装 `requirements.txt` 并运行 `app.py`。

### 方案 B：Render / Railway，推荐用于正式内部使用

优点：可以配置持久化磁盘，更适合 POI CSV 和上传图片这类本地文件存储。

推荐环境变量：

```text
HZH_STORAGE_DIR=/var/data/hzh-proposal-tool
```

Start Command：

```bash
streamlit run app.py --server.port $PORT --server.address 0.0.0.0
```

Build Command：

```bash
pip install -r requirements.txt
```

如果平台不自动提供 `$PORT`，可用：

```bash
streamlit run app.py --server.port 8501 --server.address 0.0.0.0
```

### 方案 C：Docker / 云服务器，适合长期正式化

适合需要更高控制权、公司内网、权限管理、外部数据库、对象存储的版本。

建议后续升级：

- POI 数据迁移到 PostgreSQL / Supabase
- 图片迁移到 S3 / Cloudflare R2 / Supabase Storage
- 增加登录权限
- 增加多用户并发编辑保护

## GitHub 准备清单

需要提交：

- `app.py`
- `customer_pdf_generator.py`
- `poi_importer.py`
- `requirements.txt`
- `README.md`
- `DEPLOYMENT.md`
- `.streamlit/config.toml`
- `.env.example`
- `data/poi_database.csv`
- `assets/logo/`
- 初始需要在线可用的 POI 图片

不要提交：

- `.venv/`
- `.env`
- `.streamlit/secrets.toml`
- `__pycache__/`
- `.DS_Store`
- 临时输出文件、测试 PPT、测试 PDF

## 部署前检查

在项目根目录运行：

```bash
python -m py_compile app.py customer_pdf_generator.py poi_importer.py
streamlit run app.py
```

确认页面可打开后，再测试导出。

## 部署后测试清单

1. 首页可以打开。
2. 顶部 tab 可以在「方案生成」和「POI 点位管理」之间切换。
3. POI 数据可以正常读取。
4. POI 可以搜索、筛选、编辑、软删除。
5. 点击 POI 编辑不会跳到方案生成页。
6. POI 图片可以预览；若生产环境配置了持久磁盘，上传后重启仍应存在。
7. 填写基础信息后可以生成方案初稿。
8. 每日 POI 自动推荐可以正常显示。
9. 英文 PPT 可以下载。
10. 单页客户宣传页 PDF 可以先生成草稿、编辑，再下载最终 PDF。
11. 下载文件名正常。
12. 刷新页面不出现运行错误。

## 文件存储说明

当前应用仍使用本地 CSV 和本地图片文件。

- Streamlit Community Cloud：适合试用，但运行时写入不建议作为长期正式数据源。
- Render / Railway：建议开启持久化磁盘，并设置 `HZH_STORAGE_DIR`。
- 长期正式版：建议把 POI 数据和图片迁移到数据库 + 对象存储。

## 常见问题

### 部署后找不到 POI 数据

确认 `data/poi_database.csv` 已提交到 GitHub。若使用持久化磁盘，首次启动时应用会从项目内置 `data/poi_database.csv` 复制初始数据。

### 上传图片后重启丢失

如果部署在 Streamlit Community Cloud，这是预期风险。请改用带持久磁盘的平台，或迁移到云存储。

### PDF 或 PPT 生成失败

确认依赖安装成功：

- `python-pptx`
- `reportlab`
- `Pillow`
- `pandas`
- `openpyxl`

这些已写入 `requirements.txt`。

### 部署平台要求端口

使用：

```bash
streamlit run app.py --server.port $PORT --server.address 0.0.0.0
```

## 回滚方式

如果上线后出问题：

1. 在 GitHub 找到上一个稳定 commit。
2. 在部署平台选择 Redeploy previous deploy，或回退 Git commit 后重新部署。
3. 如果使用持久化磁盘，回滚代码前建议备份 `data/poi_database.csv` 和上传图片目录。

本地 Git 回滚示例：

```bash
git log --oneline
git revert <commit_sha>
git push origin main
```
