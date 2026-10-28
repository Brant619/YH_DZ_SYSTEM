# 永辉短信对账系统

这是 `D:\YH\永辉短信对账` 的 Git 化程序仓库。仓库只保存代码，不保存账单、发票、导出文件、浏览器登录态或账号密码。

## 目录边界

| 内容 | 保存位置 | 是否进 Git |
| --- | --- | --- |
| Node/PowerShell 程序 | 本仓库 `src/`、`mcp/`、`scripts/` | 是 |
| 机器相关配置、海星账号 | `config/local.json` | 否 |
| 月度账单、PDF、Excel、ZIP | `YH_DATA_DIR` 指定的外部目录 | 否 |
| Edge 登录态 | `config/local.json` 中 `browserProfileDirectory` | 否 |

账务数据继续用 U 盘统一备份。换电脑时，代码通过 Git 恢复，数据通过 U 盘恢复。

## 当前电脑首次使用

当前电脑的 Git 程序目录为：

```text
D:\YH\永辉短信对账\ai
```

在项目目录打开 PowerShell：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\setup.ps1 `
  -DataDirectory "D:\YH\永辉短信对账"

powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\start-service.ps1 -Open
```

`setup.ps1` 会优先读取旧配置 `D:\YH\永辉短信对账\ai\config.json`，生成本机的 `config/local.json`。该文件已被 `.gitignore` 忽略。

启动后访问：

```text
http://127.0.0.1:8099/
```

需要开机自启时：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install-startup.ps1
```

## 换电脑恢复

### 1. 安装基础软件

- Git
- Node.js 20 或更高版本
- Microsoft Excel
- Microsoft Word
- Microsoft Edge

当前部分 PDF 识别会优先使用 `pdftotext`，也可通过 `PDFTOTEXT_PATH` 指定该程序。OCR 脚本仍要求 Windows PowerShell。

### 2. 克隆代码

```powershell
git clone <你的远程仓库地址> D:\YH\永辉短信对账\ai
cd D:\YH\永辉短信对账\ai
```

### 3. 从 U 盘恢复账务数据

将 U 盘中的完整账务目录复制到新电脑，例如：

```text
D:\YH\永辉短信对账
```

至少需要保留：

```text
2026年5月\
2026年6月\
...
最新明细导出\
internal-reconciliation-groups.json
```

如果 U 盘中包含了 `ai\edge-profile`，可以继续复用海星登录态；如果不包含，首次导出时重新登录即可。

### 4. 指定数据目录

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\setup.ps1 `
  -DataDirectory "D:\YH\永辉短信对账"
```

如果数据盘符不同，只需改 `-DataDirectory`。程序不再依赖固定的 Git 仓库位置。

### 5. 启动

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\start-service.ps1 -Open
```

## DeepSeek Harness / MCP

仓库自带无第三方依赖的 MCP stdio 服务：

```powershell
npm run mcp
```

安装到当前用户的 DeepSeek Harness 配置：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install-dsh-mcp.ps1
```

脚本会将 `mcp-yh-reconciliation` 写入：

```text
%USERPROFILE%\.dsh\profiles\desktop\cordis.patch.yml
```

重启 DeepSeek Harness 后可使用这些工具：

- `yh_health`
- `yh_list_month_files`
- `yh_export_summary_bill`
- `yh_organize_summary_bill`
- `yh_generate_one_click_bill`
- `yh_check_internal_reconciliation_accounts`
- `yh_generate_internal_reconciliation`
- `yh_get_internal_reconciliation_groups`

MCP 服务会按需启动本地 `8099` 服务；海星登录失效时，仍可能需要人工在弹出的 Edge 窗口完成登录。

## 验证

```powershell
npm run check
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\start-service.ps1 -Port 8099
(Invoke-RestMethod http://127.0.0.1:8099/api/health) | ConvertTo-Json
```

## 不应提交的内容

- 月度账单和发票
- `config/local.json`
- 海星用户名和密码
- `edge-profile`
- `logs`
- 临时导出的 Excel、CSV、PDF、ZIP

远程仓库建议使用私有仓库。
