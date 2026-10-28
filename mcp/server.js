#!/usr/bin/env node

const fs = require("fs");
const path = require("path");
const { execFile } = require("child_process");
const { promisify } = require("util");

const execFileAsync = promisify(execFile);
const projectRoot = path.resolve(process.env.YH_PROJECT_ROOT || path.join(__dirname, ".."));
const configPath = process.env.YH_CONFIG_PATH || path.join(projectRoot, "config", "local.json");
const startScript = path.join(projectRoot, "scripts", "start-service.ps1");
const port = Number(process.env.PORT || 8099);
const apiBase = `http://127.0.0.1:${port}`;
const serverInfo = { name: "yh-sms-reconciliation", version: "1.0.0" };

const monthSchema = {
  type: "object",
  properties: {
    month: {
      type: "string",
      description: "账期，格式 YYYY-MM，例如 2026-09"
    }
  },
  required: ["month"],
  additionalProperties: false
};

const tools = [
  {
    name: "yh_health",
    description: "检查永辉短信对账服务、数据目录和运行状态。",
    inputSchema: { type: "object", properties: {}, additionalProperties: false }
  },
  {
    name: "yh_list_month_files",
    description: "列出某个账期目录中的文件，不读取 Excel 内容。",
    inputSchema: monthSchema
  },
  {
    name: "yh_export_summary_bill",
    description: "从海星系统导出指定账期的汇总账单。浏览器登录失效时可能需要人工完成登录。",
    inputSchema: monthSchema
  },
  {
    name: "yh_organize_summary_bill",
    description: "整理已导出的汇总账单 CSV，生成汇总账单 Excel。",
    inputSchema: monthSchema
  },
  {
    name: "yh_generate_one_click_bill",
    description: "执行一键流程：导出汇总账单、整理 Excel，并在文件存在时追加港澳台和 AI 挂短账单。",
    inputSchema: monthSchema
  },
  {
    name: "yh_check_internal_reconciliation_accounts",
    description: "校验指定账期已合并账单中的账号组覆盖情况。",
    inputSchema: monthSchema
  },
  {
    name: "yh_generate_internal_reconciliation",
    description: "生成飞书内部对账 Markdown、截图和 Word 文件。",
    inputSchema: {
      type: "object",
      properties: {
        month: {
          type: "string",
          description: "账期，格式 YYYY-MM，例如 2026-09"
        },
        cutoff: {
          type: "string",
          description: "反馈截止时间，必须是可解析的 ISO 日期时间"
        }
      },
      required: ["month", "cutoff"],
      additionalProperties: false
    }
  },
  {
    name: "yh_get_internal_reconciliation_groups",
    description: "读取保存在外部数据目录中的内部对账账号组配置。",
    inputSchema: { type: "object", properties: {}, additionalProperties: false }
  }
];

function readConfig() {
  try {
    return JSON.parse(fs.readFileSync(configPath, "utf8").replace(/^\uFEFF/, ""));
  } catch {
    return {};
  }
}

function validateMonth(value) {
  const match = /^(\d{4})-(\d{2})$/.exec(String(value || ""));
  if (!match || Number(match[2]) < 1 || Number(match[2]) > 12) {
    throw new Error("month 必须是 YYYY-MM 格式。");
  }
  return {
    compact: `${match[1]}${match[2]}`,
    folderName: `${match[1]}年${Number(match[2])}月`
  };
}

async function getHealth() {
  const response = await fetch(`${apiBase}/api/health`, { signal: AbortSignal.timeout(2500) });
  if (!response.ok) throw new Error(`Health check returned HTTP ${response.status}.`);
  return response.json();
}

async function ensureService() {
  try {
    return await getHealth();
  } catch {
    if (process.platform !== "win32") {
      throw new Error("该工作流依赖 Windows、Office 和 PowerShell。");
    }
    await execFileAsync("powershell.exe", [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      startScript,
      "-Port",
      String(port)
    ], {
      cwd: projectRoot,
      windowsHide: true,
      timeout: 45000,
      maxBuffer: 2 * 1024 * 1024
    });
    for (let attempt = 0; attempt < 30; attempt += 1) {
      await new Promise((resolve) => setTimeout(resolve, 500));
      try {
        return await getHealth();
      } catch {
        // The service can need a few seconds when Office automation starts cold.
      }
    }
    throw new Error("本地服务未能在 15 秒内启动。");
  }
}

async function callApi(route, body) {
  await ensureService();
  const response = await fetch(`${apiBase}${route}`, {
    method: body === undefined ? "GET" : "POST",
    headers: body === undefined ? undefined : { "Content-Type": "application/json" },
    body: body === undefined ? undefined : JSON.stringify(body),
    signal: AbortSignal.timeout(15 * 60 * 1000)
  });
  const text = await response.text();
  let payload;
  try {
    payload = JSON.parse(text);
  } catch {
    payload = { raw: text };
  }
  if (!response.ok) {
    throw new Error(payload.message || payload.reason || `HTTP ${response.status}`);
  }
  return payload;
}

async function callTool(name, args) {
  if (name === "yh_health") {
    const health = await ensureService();
    return {
      ...health,
      configPath,
      configExists: fs.existsSync(configPath)
    };
  }

  if (name === "yh_list_month_files") {
    const month = validateMonth(args.month);
    const health = await ensureService();
    const folder = path.join(health.dataDirectory, month.folderName);
    if (!fs.existsSync(folder)) return { month: args.month, folder, files: [] };
    const files = fs.readdirSync(folder, { withFileTypes: true })
      .filter((entry) => entry.isFile())
      .map((entry) => {
        const filePath = path.join(folder, entry.name);
        const stat = fs.statSync(filePath);
        return {
          name: entry.name,
          size: stat.size,
          modifiedAt: stat.mtime.toISOString(),
          filePath
        };
      })
      .sort((left, right) => right.modifiedAt.localeCompare(left.modifiedAt));
    return { month: args.month, folder, files };
  }

  if (name === "yh_export_summary_bill") {
    return callApi("/api/export-summary-bill", { month: args.month });
  }
  if (name === "yh_organize_summary_bill") {
    return callApi("/api/organize-summary-bill", { month: args.month });
  }
  if (name === "yh_generate_one_click_bill") {
    return callApi("/api/generate-one-click-bill", { month: args.month });
  }
  if (name === "yh_check_internal_reconciliation_accounts") {
    return callApi("/api/check-internal-reconciliation-accounts", { month: args.month });
  }
  if (name === "yh_generate_internal_reconciliation") {
    return callApi("/api/generate-internal-reconciliation", {
      month: args.month,
      cutoff: args.cutoff
    });
  }
  if (name === "yh_get_internal_reconciliation_groups") {
    return callApi("/api/internal-reconciliation-groups");
  }

  throw new Error(`Unknown tool: ${name}`);
}

function send(message) {
  process.stdout.write(`${JSON.stringify(message)}\n`);
}

function sendResult(id, result) {
  send({ jsonrpc: "2.0", id, result });
}

function sendError(id, code, message) {
  send({ jsonrpc: "2.0", id, error: { code, message } });
}

async function handleMessage(message) {
  if (!message || message.jsonrpc !== "2.0") return;
  if (message.id === undefined) return;

  const { id, method, params = {} } = message;
  try {
    if (method === "initialize") {
      sendResult(id, {
        protocolVersion: params.protocolVersion || "2024-11-05",
        capabilities: { tools: {} },
        serverInfo
      });
      return;
    }
    if (method === "ping") {
      sendResult(id, {});
      return;
    }
    if (method === "tools/list") {
      sendResult(id, { tools });
      return;
    }
    if (method === "tools/call") {
      const result = await callTool(params.name, params.arguments || {});
      sendResult(id, {
        content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
        isError: false
      });
      return;
    }
    sendError(id, -32601, `Method not found: ${method}`);
  } catch (error) {
    sendResult(id, {
      content: [{ type: "text", text: error.message || String(error) }],
      isError: true
    });
  }
}

let inputBuffer = "";
process.stdin.setEncoding("utf8");
process.stdin.on("data", (chunk) => {
  inputBuffer += chunk;
  let newline;
  while ((newline = inputBuffer.indexOf("\n")) >= 0) {
    const line = inputBuffer.slice(0, newline).trim();
    inputBuffer = inputBuffer.slice(newline + 1);
    if (!line) continue;
    try {
      const message = JSON.parse(line);
      Promise.resolve(handleMessage(message)).catch((error) => {
        if (message.id !== undefined) sendError(message.id, -32603, error.message || String(error));
      });
    } catch (error) {
      process.stderr.write(`Invalid MCP JSON: ${error.message}\n`);
    }
  }
});
