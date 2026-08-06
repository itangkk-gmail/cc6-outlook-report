# Outlook CC6 Daily Report 自动汇总

通过 Microsoft Graph API 连接 Outlook 邮箱，自动下载主题同时包含 **CC6** 与 **report** 的邮件附件，并把日报数据合并到一张总表。

## 功能

- 筛选主题含 `CC6` + `report` 的邮件（不区分大小写）
- 附件自动保存到指定目录（不删除旧文件；重名自动加后缀）
- 从文件名解析日期（如 `2026-6-4` → `2026-06-04`），写入总表 `Date` 列
- 主题含 `update`：按该日期 **整批覆盖** 总表当天数据
- 非 update：向总表 **追加**
- `processed.json` 记录已处理邮件，避免重复合并
- 每次运行写入成功/失败状态与日志（窗口不会一闪而过）
- 支持本地文件合并模式（无需 Outlook，用于测试或手动补录）

---

## 目录

- [环境要求](#环境要求)
- [第一次使用：注册 Azure 应用](#第一次使用注册-azure-应用)
- [安装](#安装)
- [配置说明](#配置说明)
- [运行方式](#运行方式)
- [登录认证](#登录认证)
- [本地文件合并（无需 Outlook）](#本地文件合并不需要-outlook)
- [任务计划（每天自动运行）](#任务计划每天自动运行)
- [成功/失败查看](#成功失败查看)
- [数据说明](#数据说明)
- [常见问题](#常见问题)
- [文件结构](#文件结构)

---

## 环境要求

1. **Windows** 操作系统
2. **Python 3.10+**（安装时勾选 *Add Python to PATH*）
3. **Outlook 邮箱**（任意支持 Microsoft Graph 的 Microsoft 账户，个人/企业均可）

> 注意：本工具通过 Microsoft Graph API 访问邮箱，**不需要安装桌面版 Outlook**，也不需要 Outlook 客户端保持运行。只要有能收邮件的 Microsoft 账户即可。

---

## 第一次使用：注册 Azure 应用

首次使用时，需要在 Azure 门户注册一个应用，以获取访问邮箱的权限。**每个用户只需操作一次**。

### 步骤 1：进入 Azure 应用注册页面

打开浏览器访问：  
[https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/CreateApplicationBlade](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/CreateApplicationBlade)

用你的 Microsoft 账户登录（就是你收 CC6 邮件那个账户）。

### 步骤 2：注册新应用

填写以下信息：

| 字段 | 值 |
|------|-----|
| **名称** | `OutlookCC6Report`（或任意你喜欢的名字） |
| **支持的帐户类型** | 选择 **"仅限个人 Microsoft 帐户"**（如果是个人邮箱如 @outlook.com/@hotmail.com） 或 **"任何组织目录中的帐户和个人 Microsoft 帐户"**（如果是企业邮箱） |
| **重定向 URI** | 不用填，后续步骤会配置 |

点击 **"注册"**。

### 步骤 3：获取应用程序（客户端）ID

注册成功后，页面会跳转到应用概览页。复制 **"应用程序(客户端) ID"**（一串类似 `756985aa-4502-44b5-8820-6634859e76ea` 的字符串）。

### 步骤 4：配置重定向 URI

1. 在左侧菜单点击 **"身份验证"**
2. 点击 **"添加平台"** → 选择 **"移动和桌面应用程序"**
3. 在自定义重定向 URI 中输入：`http://localhost:8400`
4. 点击 **"配置"**
5. 保存页面

### 步骤 5：设置 API 权限

1. 在左侧菜单点击 **"API 权限"**
2. 点击 **"添加权限"** → 选择 **"Microsoft Graph"** → **"委托的权限"**
3. 搜索并勾选 `Mail.Read`（在 "邮件" 分类下，权限读取用户邮件）
4. 点击 **"添加权限"**
5. 如果显示需要管理员同意，个人账户忽略即可（无需管理员同意）

### 步骤 6：填入配置

将步骤 3 中复制的 **"客户端 ID"** 填入项目根目录的 `config.json`：

```json
"graph": {
    "client_id": "你的客户端ID",
    "tenant": "consumers",
    "token_cache_file": "./logs/graph_token_cache.json"
}
```

| 字段 | 说明 |
|------|------|
| `client_id` | 从 Azure 复制的应用程序（客户端）ID |
| `tenant` | 个人邮箱用 `consumers`；企业邮箱用 `organizations` 或公司 tenant ID |
| `token_cache_file` | 登录凭据缓存文件路径，免重复登录 |

> Azure 应用注册是一次性的。注册完、填好 client_id 后，后续所有用户只需要在首次运行时登录一次即可。

---

## 安装

```bat
cd /d "D:\AiWorkSpace\outlook report"
python -m pip install -r requirements.txt
```

`requirements.txt` 依赖：

| 包 | 用途 |
|----|------|
| `openpyxl` | 读写 Excel 文件 |
| `pandas` | 数据处理与合并 |
| `msal` | Microsoft 身份认证 |
| `requests` | 调用 Microsoft Graph API |

---

## 配置说明

编辑 `config.json`，各字段说明如下：

```json
{
  "download_dir": "./downloads",
  "master_file": "./output/CC6_master.xlsx",
  "master_sheet": "Master",
  "processed_file": "./processed.json",
  "log_file": "./logs/run.log",
  "status_file": "./logs/last_status.txt",
  "status_json": "./logs/last_status.json",

  "outlook": {
    "folder": "Inbox",
    "lookback_days": 30,
    "subject_keywords": ["CC6", "report"],
    "update_keyword": "update",
    "attachment_extensions": [".xlsx", ".xls"]
  },

  "graph": {
    "client_id": "你的客户端ID",
    "tenant": "consumers",
    "token_cache_file": "./logs/graph_token_cache.json"
  },

  "excel": {
    "sheet_name": "DAILY REPORT",
    "header_keyword": "ID /",
    "data_start_col": 8,
    "data_end_col": 23
  }
}
```

### 路径类配置

| 字段 | 默认值 | 说明 |
|------|--------|------|
| `download_dir` | `./downloads` | 邮件附件下载保存目录 |
| `master_file` | `./output/CC6_master.xlsx` | 合并后的总表路径 |
| `master_sheet` | `Master` | 总表的工作表名称 |
| `processed_file` | `./processed.json` | 已处理邮件记录（自动维护，勿手动修改） |
| `log_file` | `./logs/run.log` | 历史追加日志 |
| `status_file` | `./logs/last_status.txt` | 最近一次运行状态（文本格式） |
| `status_json` | `./logs/last_status.json` | 最近一次运行状态（JSON格式） |

### Outlook 筛选配置

| 字段 | 默认值 | 说明 |
|------|--------|------|
| `folder` | `Inbox` | 要扫描的邮件文件夹名称（大小写不敏感；如找不到则回退到收件箱） |
| `lookback_days` | `30` | 扫描最近多少天内的邮件 |
| `subject_keywords` | `["CC6", "report"]` | 邮件主题必须**同时包含**的关键词列表（不区分大小写） |
| `update_keyword` | `update` | 主题中包含此关键词代表是**覆盖更新**而非追加 |
| `attachment_extensions` | `[".xlsx", ".xls"]` | 允许下载的附件类型 |

### Graph API 配置

| 字段 | 说明 |
|------|------|
| `client_id` | Azure AD 应用程序（客户端）ID，**必填** |
| `tenant` | `consumers`（个人账户）、`organizations`（企业账户）、或具体 tenant ID |
| `token_cache_file` | 登录凭据缓存文件，存于本地免重复登录 |

### Excel 解析配置

| 字段 | 默认值 | 说明 |
|------|--------|------|
| `sheet_name` | `DAILY REPORT` | 日报文件中要读取的工作表名称 |
| `header_keyword` | `ID /` | 通过该关键字定位表头行（在 `data_start_col` 列中查找） |
| `data_start_col` | `8` | 数据起始列号（H 列 = 8） |
| `data_end_col` | `23` | 数据结束列号（W 列 = 23），共读取 16 列 |

---

## 运行方式

| 方式 | 命令/操作 | 适合场景 |
|------|-----------|----------|
| **BAT（双击）** | 双击 `run.bat` | 最简单，日常使用 |
| **BAT（计划任务）** | `run.bat --nopause` | 配合任务计划程序自动运行 |
| **PowerShell** | `.\run.ps1` 或右键"使用 PowerShell 运行" | 颜色区分成功/失败，更直观 |
| **直接 Python** | `python main.py` | 调试排错 |
| **仅合并（无网络）** | `python main.py --merge-only --file "xxx.xlsx"` | 手动补录或测试 |

### 手动运行

```bat
run.bat
```

或：

```powershell
.\run.ps1
```

### 计划任务模式（不等待按键）

```bat
run.bat --nopause
```

```powershell
.\run.ps1 -NoPause
```

### 使用自定义配置

```bat
run.bat --config "other_config.json"
```

```powershell
.\run.ps1 -Config "other_config.json"
```

---

## 登录认证

### 首次登录

第一次运行 `run.bat` 或 `python main.py` 时，程序会：

1. 在终端打印一个 Microsoft 登录链接
2. **自动打开默认浏览器**跳转到该链接
3. 用你的 Microsoft 账户登录并授权（勾选"代表你读取邮件"）
4. 登录成功后浏览器显示 **"登录成功"** 页面，可关闭浏览器
5. 凭据自动缓存到 `logs/graph_token_cache.json`，**之后运行无需再次登录**

### 自动运行模式

- 首次运行会自动弹出浏览器要求登录
- 登录成功后会缓存 token，后续运行自动使用缓存的凭据
- 如果运行时桌面无浏览器（如远程桌面/服务器），需要提前登录一次以生成 token 缓存

### 重新登录

如果 token 过期或需要换账户，删除 `logs/graph_token_cache.json` 后重新运行即可触发新的登录流程：

```bat
del logs\graph_token_cache.json
run.bat
```

---

## 本地文件合并（不需要 Outlook）

当你需要手动补录某个日报文件到总表、或不方便连接网络时，可以使用 `--merge-only` 模式：

```bat
:: 追加模式（默认）
python main.py --merge-only --file "sample\No.165Daily Report  2026-6-4.xlsx"

:: 覆盖更新模式（替换该日期在总表中的所有数据）
python main.py --merge-only --file "sample\No.165Daily Report  2026-6-4.xlsx" --update

:: 批量合并多个文件
python main.py --merge-only --file "file1.xlsx" --file "file2.xlsx" --file "file3.xlsx"
```

> 注意：`--merge-only` 模式不会写入 `processed.json`，不会阻止之后从 Outlook 重新下载同名文件。

---

## 任务计划（每天自动运行）

### 设置步骤

1. 按 `Win + R`，输入 `taskschd.msc`，打开「任务计划程序」
2. 右侧点击 **"创建基本任务"**
3. **名称**：`CC6 Outlook Report`
4. **触发器**：选择"每天"，设置运行时间（建议上午 9:00 之后，确保前一天的邮件已到达）
5. **操作** → "启动程序"：
   - **程序**：`D:\AiWorkSpace\outlook report\run.bat`
   - **参数**：`--nopause`
   - **起始于**：`D:\AiWorkSpace\outlook report`
6. 勾选 **"只有用户登录时才运行"**（Graph API 需要用户会话完成首次认证）
7. 完成创建

### 高级设置（推荐）

创建后右键任务 → 属性：

- **常规** 标签 → 勾选 **"不管用户是否登录都运行"** 并勾选 **"使用最高权限运行"**（如适用）
- **触发器** 标签 → 可添加多个触发器（如每天 9:00 和 17:00 各运行一次）
- **设置** 标签 → 勾选 **"如果任务失败，按以下频率重新启动"**（如每分钟，最多 3 次）

### 验证自动运行

任务计划跑完后，打开 `logs\last_status.txt`，确认内容为：

```
status=SUCCESS
...
```

如为 `FAILED`，查看对应的 `logs\run_YYYYMMDD_HHMMSS.log` 了解具体原因。

### 首次自动运行的准备

运行任务计划的机器（通常是你的办公电脑）需要满足：

1. Python 已安装且可通过命令行 `python` 访问
2. 至少手动运行过一次 `run.bat`，完成浏览器登录，生成 `graph_token_cache.json`
3. 运行时间段电脑处于开机并解锁状态（锁屏不影响，但必须在用户会话中）

---

## 成功/失败查看

### 运行时窗口

- 双击 `run.bat` 运行结束后，窗口会暂停显示结果：
  - `RESULT: SUCCESS` — 运行成功
  - `RESULT: FAILED` — 运行失败，同时显示退出码和日志路径

### 日志文件

| 文件 | 内容 |
|------|------|
| `logs/last_status.txt` | 最近一次运行的摘要状态 |
| `logs/last_status.json` | 同上（JSON 格式，方便程序读取） |
| `logs/run_YYYYMMDD_HHMMSS.log` | **本次**完整运行日志 |
| `logs/run.log` | 历史追加日志（所有运行的汇总） |

### 运行结果解读

**成功示例（`last_status.json`）**：

```json
{
  "status": "SUCCESS",
  "ok": true,
  "exit_code": 0,
  "message": "Completed: merged=2, skipped=0",
  "details": {
    "merged": 2,
    "skipped": 0,
    "errors": 0,
    "attachments": 2
  }
}
```

**失败示例**：

```json
{
  "status": "FAILED",
  "ok": false,
  "exit_code": 1,
  "message": "Token expired – delete logs/graph_token_cache.json and re-run"
}
```

| 字段 | 含义 |
|------|------|
| `merged` | 成功合并到总表的文件数 |
| `skipped` | 已处理过的文件数（`processed.json` 中已有记录） |
| `errors` | 合并失败的文件数 |
| `attachments` | 本次找到的符合条件的附件总数 |
| `no_date` | 从文件名无法解析日期的文件数（附件已保存但未合并） |

---

## 数据说明

### 日报源表结构

日报 Excel 文件的工作表 `DAILY REPORT` 中，程序读取 **H–W 列（第 8–23 列，共 16 列）** 的人力/进度数据：

- 通过关键字 `ID /` 定位表头行
- 从表头下一行开始，读取所有 ID 列为数字的数据行
- 遇到第一个非数字 ID 且已有数据行后，停止读取

### 总表结构

总表 `CC6_master.xlsx` 存储在 `output/` 目录中：

- **第一列** `Date`：数据来源日期（从文件名解析）
- **后续 16 列**：与源表 H–W 列对应

### 合并逻辑

| 邮件主题条件 | 处理方式 |
|-------------|----------|
| 不含 `update` | **追加**：新行添加到总表末尾 |
| 含 `update`（如 `updated`、`Update`） | **覆盖更新**：删除总表中该日期的所有旧行，再写入新数据 |

### 去重机制

- `processed.json` 以 `消息ID|文件名` 为键记录已处理的附件
- 同一封邮件的同一附件不会重复下载和合并
- 如需强制重新处理某个附件，在 `processed.json` 中删除对应条目后重新运行

---

## 常见问题

### Q: 运行时提示 "请先在 config.json -> graph -> client_id 中填入你的 Azure AD 应用 ID"

**A:** 还没有注册 Azure 应用。请按照 [第一次使用：注册 Azure 应用](#第一次使用注册-azure-应用) 完成配置。

### Q: 浏览器登录后提示 "AADSTS50011: The reply URL specified does not match..."

**A:** 重定向 URI 配置不正确。进入 Azure 门户 → 应用 → 身份验证 → 移动和桌面应用程序 → 确保 `http://localhost:8400` 已添加。

### Q: 运行时提示 "Graph API token expired"

**A:** 登录凭据已失效。删除缓存文件后重新运行即可触发登录：

```bat
del logs\graph_token_cache.json
run.bat
```

### Q: 没有匹配到任何邮件附件

**A:** 检查以下几点：
1. `config.json` 中 `lookback_days` 是否覆盖了目标日期
2. `folder` 文件夹名称是否与 Outlook 中一致
3. 邮件主题是否同时包含 `subject_keywords` 中的**所有**关键词
4. 附件后缀是否在 `attachment_extensions` 列表中

### Q: 文件已下载但未合并，日志提示 "cannot parse date from filename"

**A:** 文件名中缺少日期。程序要求文件名包含格式为 `YYYY-M-D` 或 `YYYY-MM-DD` 的日期（如 `2026-6-4`、`2026-08-04`）。可将文件改名后，用 `--merge-only` 模式手动合并。

### Q: 提示 "Python not found"

**A:** Python 未安装或未加入 PATH：
1. 重新安装 Python，**务必勾选** "Add Python to PATH"
2. 或把 Python 安装目录手动加入系统环境变量 PATH

### Q: 用了 `--merge-only` 之后又从 Outlook 重新下载了相同文件

**A:** `--merge-only` 模式不写入 `processed.json`，设计如此。如果需要阻止重复下载，需手动在 `processed.json` 中补充对应条目。

### Q: 只有网页版 Outlook（如 QQ邮箱代收），没有 Microsoft 账户

**A:** 本工具依赖 Microsoft Graph API，必须通过 Microsoft 账户（@outlook.com / @hotmail.com / Microsoft 365 企业邮箱）访问。如果邮件是由 QQ 邮箱等转发到 Outlook 的，确保运行程序的 Microsoft 账户能收到这些转发邮件即可。

### Q: 需要管理员权限吗？

**A:** 不需要。使用个人账户的 "仅限个人 Microsoft 帐户" 模式注册应用无需任何管理员审批。

---

## 文件结构

```
outlook report/
├── config.json              # 主配置文件（请修改 client_id）
├── main.py                  # 程序入口
├── outlook_client.py        # Microsoft Graph API 客户端（认证 + 邮件/附件下载）
├── excel_merge.py           # Excel 读写与合并逻辑
├── requirements.txt         # Python 依赖
├── run.bat                  # Windows 批处理启动器
├── run.ps1                  # PowerShell 启动器
├── processed.json           # 已处理附件记录（自动维护）
├── downloads/               # 邮件附件下载目录
├── output/                  # 总表输出目录
│   └── CC6_master.xlsx      # 合并后的日报总表
├── logs/                    # 日志目录
│   ├── last_status.txt      # 最近运行状态（文本）
│   ├── last_status.json     # 最近运行状态（JSON）
│   ├── graph_token_cache.json # 登录凭据缓存
│   ├── run.log              # 历史追加日志
│   └── run_YYYYMMDD_*.log   # 每次运行的独立日志
└── sample/                  # 示例日报文件（用于 --merge-only 测试）
```
