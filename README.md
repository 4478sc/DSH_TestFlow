# TestFlow

多项目测试记录网站：把一份「测试计划 Excel」变成可检索、可记录、可导出的暗色工作台。

[English](#english) · [功能](#功能) · [快速开始](#快速开始) · [表格格式](#表格格式) · [配置](#配置) · [API](#api) · [测试](#测试)

![tests](https://github.com/4478sc/DSH_TestFlow/actions/workflows/test.yml/badge.svg)

---

## 功能

**多项目**

- 上传 `.xlsx` / `.xlsm` 测试计划表即可创建项目，按表头识别列，列顺序可不同。
- 顶栏下拉切换项目，`?project=<id>` 可直接分享某个项目；删除项目会二次确认（连同其记录与附件）。

**测试项**

- 每个项目下维护测试项：编号、优先级、模块、名称、测试指导（测试目的 / 测试连接·激励 / 需记录的指标 / 建议测试条件 / 计算·处理方法 / 论文参考·建议判据 / 仪器）。
- 新建测试项（编号按现有前缀/位宽自动建议）、删除测试项、编辑测试说明。
- 复制 / 剪切 / 粘贴测试项：支持跨模块、跨项目。复制不带走记录；剪切会把测试项连同记录和附件一起移动。

**测试记录**

- 每个测试项可登记多条记录：仿真 / 实测、结果、单位、Pass·Fail、测试日期、备注。
- 每条记录可上传多个 Excel（xlsx / xls / xlsm）与多个 JPG/JPEG，支持图片预览、下载、删除和失败重试。
- 按模块、优先级、关键词筛选测试项；按判定和模式筛选记录。

**数据进出**

- 导出当前项目全部记录为 XLSX（需要 openpyxl）。
- 备份 ZIP：SQLite 快照 + 全部附件。

**安全边界**

- 无账号体系，默认只监听 `127.0.0.1`，适合单机 / 内网使用。
- Host / Origin 白名单（回环 + 本机网卡地址 + `AFE_ALLOWED_HOSTS`），阻断 DNS rebinding 与跨站写入。
- 附件按扩展名 + 文件头校验，单文件 50 MiB、单请求 100 MiB；静态文件路径穿越防护。

---

## 快速开始

要求：Python 3.10+，以及 `openpyxl`（解析上传的表格、导出 XLSX 都需要）。

```bash
python -m pip install openpyxl
python server.py --port 3080
```

打开 <http://127.0.0.1:3080>。首次启动没有任何项目，点击顶栏的 **+**，上传一份测试计划表格即可创建第一个项目。

数据保存在 `data/`（`afe-testflow.db` 与 `uploads/`），随服务重启保留。

可选参数：

```bash
python server.py --host 0.0.0.0 --port 3080 --data-dir /path/to/data
```

---

## 表格格式

上传的表格需要一张「测试计划」工作表，表头包含（名称匹配，顺序可变，多余列忽略）：

```
编号 | 优先级 | 模块 | 测试项目 | 测试目的 | 测试连接/激励 | 需记录的指标 |
建议测试条件 | 计算/处理方法 | 论文参考/建议判据 | 仪器 | 实测结果 | 单位 | Pass/Fail | 测试日期 | 备注
```

- 至少要能识别出「测试项目」列；缺少时会明确报错。
- 可选工作表：**指标汇总**（模块 / 建议优先记录的核心指标 / 结果 / 单位 / 目标·判据 / 备注）、**填写说明**（项目 / 说明）。
- `.xls` 请先在 Excel 中另存为 `.xlsx`。
- 表格只读解析，不会回写原文件。

---

## 配置

| 参数 | 环境变量 | 说明 |
|---|---|---|
| `--host` | `AFE_HOST` | 监听地址，默认 `127.0.0.1` |
| `--port` | `AFE_PORT` | 监听端口，默认 `3080` |
| `--data-dir` | — | 数据目录，默认 `<项目>/data` |
| `--seed` | — | 初始项目种子文件，默认 `seed.json`（可选） |
| — | `AFE_ALLOWED_HOSTS` | 额外的 Host/Origin 白名单，逗号 / 空格分隔 |
| — | `AFE_QUIET` | 设为 1 关闭请求日志 |

---

## API

| 方法 | 路径 | 说明 |
|---|---|---|
| GET | `/api/health` | 健康检查 |
| GET | `/api/projects` | 项目列表 |
| POST | `/api/projects` | 上传表格创建项目（multipart：`file`、`name`） |
| DELETE | `/api/projects/:id` | 删除项目（级联记录与附件） |
| GET | `/api/projects/:id/bootstrap` | 项目 + 测试项 + 指标汇总 + 填写说明 |
| POST | `/api/projects/:id/tests` | 新建测试项 |
| PUT | `/api/projects/:id/tests/:tid` | 编辑测试指导字段 |
| DELETE | `/api/projects/:id/tests/:tid` | 删除测试项 |
| GET / POST | `/api/projects/:id/tests/:tid/records` | 记录列表 / 新建记录 |
| PUT | `/api/records/:id` | 编辑记录 |
| POST | `/api/records/:id/attachments` | 上传附件（multipart：`kind=data|image`） |
| DELETE | `/api/attachments/:id` | 删除附件 |
| GET | `/api/attachments/:id/file` | 下载 / 预览附件 |
| GET / POST / DELETE | `/api/clipboard` | 查看 / 复制剪切 / 清空剪贴板 |
| POST | `/api/clipboard/paste` | 粘贴到目标项目与模块 |
| GET | `/api/projects/:id/export.xlsx` | 导出该项目记录为 XLSX |
| GET | `/api/backup.zip` | 备份数据库与附件 |

错误统一返回 `{"error": "..."}`，写操作需要合法的 `Origin`。

---

## 目录结构

```
server.py            HTTP 服务、SQLite 存储、项目 / 测试项 / 记录 / 附件 / 剪贴板 API
workbook_import.py   测试计划表格解析（按表头识别列）
test_server.py       后端单元与接口测试
public/
  index.html         页面骨架
  app.js             前端逻辑（原生 JS，无构建步骤）
  styles.css         暗色主题样式
data/                运行时数据（SQLite + 附件，已在 .gitignore 中忽略）
```

前端不依赖任何外部 CDN 或在线服务，全部资源随仓库提供。

---

## 测试

```bash
python -m unittest -v test_server.py
```

覆盖：项目创建/删除、表格解析、测试项增删改、记录校验、附件上传下载删除、剪贴板复制剪切粘贴（含跨项目搬移记录）、XLSX 导出、ZIP 备份、Host/Origin 校验、路径穿越与大小限制、重启持久化、空库首次使用。

---

## 隐私与安全

- 仓库不含任何真实测试数据、数据库、附件或个人信息；`data/` 与表格文件均在忽略列表中。
- 网站没有身份认证，请只在本机或受控内网使用；对外监听时配合 `AFE_ALLOWED_HOSTS` 与防火墙。

---

## English

TestFlow is a self-hosted, multi-project test-record web app. Import a test-plan workbook (.xlsx/.xlsm) to
create a project, maintain test items (add, edit guidance, delete, copy/cut/paste across modules and projects),
log simulation or measurement runs with Excel/JPG attachments, and export records to XLSX or back up to ZIP.

Python standard library + SQLite on the server; vanilla JS/CSS on the client with no build step and no CDN.
Run `python -m pip install openpyxl && python server.py --port 3080` and open <http://127.0.0.1:3080>.

---

## License

MIT — see [LICENSE](LICENSE).
