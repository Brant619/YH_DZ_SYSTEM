const http = require("http");
const fs = require("fs");
const path = require("path");
const { execFile } = require("child_process");

const root = __dirname;
const configPath = process.env.YH_CONFIG_PATH || path.join(root, "..", "config", "local.json");
let localConfig = {};
try {
  localConfig = JSON.parse(fs.readFileSync(configPath, "utf8").replace(/^\uFEFF/, ""));
} catch {
  localConfig = {};
}
const dataRoot = path.resolve(process.env.YH_DATA_DIR || localConfig.dataDirectory || path.join(root, "..", "data"));
const port = Number(process.env.PORT || 8099);
const host = "127.0.0.1";
const maxJsonBodyBytes = 1024 * 1024;
const maxUploadBytes = 200 * 1024 * 1024;
const maxDownloadBytes = 200 * 1024 * 1024;
const latestDetailUrlPrefix = "https://taiyi.yonghui.cn/s/it-service-sms-alert/v1/report/smsData/";
let activeExport = null;
const activeTasks = new Map();
const internalGroupsFile = path.join(dataRoot, "internal-reconciliation-groups.json");
const defaultInternalGroups = [
  { id: 1, accounts: ["s_scyh", "s_scyh_hxzx"], contacts: ["刘香君"] },
  { id: 2, accounts: ["s_jsyh", "s_ahyh", "s_zjyh"], contacts: ["章玉丹"] },
  { id: 3, accounts: ["s_fjyh", "s_fjfzdq_dy"], contacts: ["宋文辉"] },
  { id: 4, accounts: ["s_zlglb_qa"], contacts: ["杨琳"] },
  { id: 5, accounts: ["s_ywtj", "s_cqxs"], contacts: ["张木森"] },
  { id: 6, accounts: ["s_1233_sms"], contacts: ["陈文杰"] },
  { id: 7, accounts: ["sn_kfyy", "s_gdxt"], contacts: ["邵力", "何晓莹"] },
  { id: 8, accounts: ["sn_mss_p"], contacts: ["李雪"] },
  { id: 9, accounts: ["s_xsly"], contacts: ["刘德渊"] },
  { id: 10, accounts: ["s_ptdc_qg"], contacts: ["张婷"] },
  { id: 11, accounts: ["s_cq_yddq", "s_cq_cqdq", "s_cq_yndq"], contacts: ["陈语", "刘婕"] },
  { id: 12, accounts: ["s_xian_dj"], contacts: ["彭勇"] }
];

const text = {
  year: "\u5e74",
  month: "\u6708",
  billPrefix: "\u5bf9\u8d26\u5355_",
  organizePrefix: "\u6c47\u603b\u8d26\u5355\u6574\u7406_",
  hmtPrefix: "\u6e2f\u6fb3\u53f0\u8d26\u5355_",
  aiShortPrefix: "AI\u6302\u77ed\u8d26\u5355_",
  latestDetailPrefix: "\u6700\u65b0\u660e\u7ec6_",
  expenseComparePrefix: "\u8d39\u7528\u5bf9\u6bd4_",
  filterPrefix: "\u7b5b\u9009\u8d26\u5355_",
  combinedPrefix: "\u5bf9\u8d26\u5408\u5e76_",
  invalidJson: "\u8bf7\u6c42\u53c2\u6570\u4e0d\u662f\u6709\u6548 JSON\u3002",
  invalidMonth: "\u8bf7\u9009\u62e9\u6b63\u786e\u7684\u8d26\u671f\uff0c\u683c\u5f0f\u5e94\u4e3a YYYY-MM\u3002",
  exportSuccess: "\u5bfc\u51fa\u6210\u529f\u3002",
  noBill: "\u8be5\u8d26\u671f\u4e0d\u5b58\u5728\u8d26\u5355\uff0c\u672a\u5bfc\u51fa\u6587\u4ef6\u3002",
  organizeSuccess: "\u6574\u7406\u5b8c\u6210\u3002",
  downloadSuccess: "\u4e0b\u8f7d\u6210\u529f\u3002",
  appendSuccess: "\u8ffd\u52a0\u5b8c\u6210\u3002"
};

const types = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".csv": "text/csv; charset=utf-8",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
};

const filterBillTypes = {
  jiangxiJumeng: {
    title: "\u6c5f\u897f\u805a\u68a6",
    columnName: "\u901a\u9053\u540d",
    matchValue: "\u6c5f\u897f\u805a\u68a6"
  },
  beijingLingdu: {
    title: "\u5317\u4eac\u51cc\u6e21",
    columnName: "\u901a\u9053\u540d",
    matchValue: "\u5317\u4eac\u51cc\u6e21"
  },
  jingdong: {
    title: "\u4eac\u4e1c",
    columnName: "\u901a\u9053\u540d",
    matchValue: "\u4eac\u4e1c"
  },
  tengyu: {
    title: "腾域",
    columnName: "通道名",
    matchValue: "腾域"
  },
  techBill: {
    title: "\u79d1\u6280\u8d26\u5355",
    columnName: "\u5bf9\u63a5\u4eba",
    matchValue: "\u987e\u4f73\u96ef"
  }
};

const invoiceProviders = {
  jiangxiJumeng: {
    title: "江西聚梦"
  },
  beijingLingdu: {
    title: "北京凌渡"
  },
  jingdong: {
    title: "京东"
  },
  tengyu: { title: "\u817e\u57df" }
};

function formatMoney(value) {
  const number = Number(value || 0);
  return Math.round(number * 100) / 100;
}

function moneyText(value) {
  return formatMoney(value).toFixed(2);
}

function filenameFromUrl(url, contentType) {
  const pathname = decodeURIComponent(new URL(url).pathname);
  const base = path.basename(pathname).replace(/[<>:"/\\|?*\x00-\x1F]/g, "_");
  if (base && base.includes(".")) return base;
  if ((contentType || "").includes("spreadsheetml")) return `${text.hmtPrefix}source.xlsx`;
  if ((contentType || "").includes("excel")) return `${text.hmtPrefix}source.xls`;
  if ((contentType || "").includes("csv")) return `${text.hmtPrefix}source.csv`;
  return `${text.hmtPrefix}source.xlsx`;
}

function extensionFromContentType(contentType) {
  const value = String(contentType || "").toLowerCase();
  if (value.includes("spreadsheetml")) return ".xlsx";
  if (value.includes("ms-excel") || value.includes("excel")) return ".xls";
  if (value.includes("csv")) return ".csv";
  if (value.includes("zip")) return ".zip";
  return ".xlsx";
}

function filenameFromContentDisposition(header) {
  const value = String(header || "");
  const encoded = /filename\*\s*=\s*UTF-8''("?)([^";]+)\1/i.exec(value);
  if (encoded) {
    try {
      return safeFilename(decodeURIComponent(encoded[2]), "");
    } catch {
      return safeFilename(encoded[2], "");
    }
  }
  const plain = /filename\s*=\s*(?:"([^"]+)"|([^;]+))/i.exec(value);
  return safeFilename(plain?.[1] || plain?.[2] || "", "");
}

function validateLatestDetailUrl(value) {
  let parsed;
  try {
    parsed = new URL(String(value || "").trim());
  } catch {
    throw new Error("请输入有效的最新明细下载 URL。");
  }
  if (
    parsed.protocol !== "https:"
    || parsed.hostname !== "taiyi.yonghui.cn"
    || !parsed.pathname.startsWith("/s/it-service-sms-alert/v1/report/smsData/")
  ) {
    throw new Error(`仅支持 ${latestDetailUrlPrefix} 开头的下载 URL。`);
  }
  const token = parsed.pathname.slice("/s/it-service-sms-alert/v1/report/smsData/".length);
  if (!token || token.includes("/") || token.length > 256) {
    throw new Error("最新明细下载 URL 的末段标识无效。");
  }
  return parsed;
}

function timestampText(date = new Date()) {
  return [
    date.getFullYear(),
    String(date.getMonth() + 1).padStart(2, "0"),
    String(date.getDate()).padStart(2, "0"),
    "_",
    String(date.getHours()).padStart(2, "0"),
    String(date.getMinutes()).padStart(2, "0"),
    String(date.getSeconds()).padStart(2, "0")
  ].join("");
}

function sendJson(res, statusCode, payload) {
  res.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET,POST,PUT,OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  });
  res.end(JSON.stringify(payload));
}

function runEndpoint(handler, req, res, label) {
  if (typeof handler !== "function") {
    sendJson(res, 501, { message: `${label}功能尚未加载，请检查后端接口实现。` });
    return;
  }
  Promise.resolve(handler(req, res)).catch((error) => {
    sendJson(res, 500, { message: error.message || `${label}失败。` });
  });
}

function readBody(req, limitBytes = maxJsonBodyBytes) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let total = 0;
    req.on("data", (chunk) => {
      total += chunk.length;
      if (total > limitBytes) {
        reject(new Error("请求内容过大。"));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.on("error", reject);
  });
}

function readBodyBuffer(req, limitBytes = maxUploadBytes) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let total = 0;
    req.on("data", (chunk) => {
      total += chunk.length;
      if (total > limitBytes) {
        reject(new Error("上传文件过大，请控制在 200MB 以内。"));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

function safeFilename(name, fallback) {
  const clean = path.basename(name || "").replace(/[<>:"/\\|?*\x00-\x1F]/g, "_");
  return clean || fallback;
}

function parseMultipartUpload(body, contentType) {
  const boundaryMatch = /boundary=(?:"([^"]+)"|([^;]+))/i.exec(contentType || "");
  const boundary = boundaryMatch && (boundaryMatch[1] || boundaryMatch[2]);
  if (!boundary) return null;

  const marker = Buffer.from(`--${boundary}`);
  const parts = [];
  let position = body.indexOf(marker);

  while (position !== -1) {
    position += marker.length;
    if (body[position] === 45 && body[position + 1] === 45) break;
    if (body[position] === 13 && body[position + 1] === 10) position += 2;

    const headerEnd = body.indexOf(Buffer.from("\r\n\r\n"), position);
    if (headerEnd === -1) break;

    const nextMarker = body.indexOf(marker, headerEnd + 4);
    if (nextMarker === -1) break;

    const headers = body.slice(position, headerEnd).toString("utf8");
    let dataEnd = nextMarker;
    if (body[dataEnd - 2] === 13 && body[dataEnd - 1] === 10) dataEnd -= 2;
    const data = body.slice(headerEnd + 4, dataEnd);
    parts.push({ headers, data });
    position = nextMarker;
  }

  const fields = {};
  const files = {};
  let file = null;

  for (const part of parts) {
    const disposition = /content-disposition:[^\r\n]+/i.exec(part.headers)?.[0] || "";
    const name = /name="([^"]+)"/i.exec(disposition)?.[1];
    const filename = /filename="([^"]*)"/i.exec(disposition)?.[1];
    if (!name) continue;
    if (filename !== undefined) {
      file = { field: name, filename, data: part.data };
      files[name] = file;
    } else {
      fields[name] = part.data.toString("utf8");
    }
  }

  return { fields, file, files };
}

function execFileText(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    execFile(command, args, {
      windowsHide: true,
      maxBuffer: 1024 * 1024 * 10,
      ...options
    }, (error, stdout, stderr) => {
      if (error) {
        const detail = String(stderr || stdout || error.message || "").trim();
        reject(new Error(detail || "外部程序执行失败。"));
        return;
      }
      resolve(String(stdout || ""));
    });
  });
}

function uniqueFileSystemPath(basePath) {
  if (!fs.existsSync(basePath)) return basePath;
  const extension = path.extname(basePath);
  const stem = extension ? basePath.slice(0, -extension.length) : basePath;
  for (let index = 2; index < 1000; index += 1) {
    const candidate = `${stem} (${index})${extension}`;
    if (!fs.existsSync(candidate)) return candidate;
  }
  throw new Error("同名文件过多，无法生成新的保存路径。");
}

function isPathInside(basePath, candidatePath) {
  const relative = path.relative(path.resolve(basePath), path.resolve(candidatePath));
  return relative === "" || (!relative.startsWith("..") && !path.isAbsolute(relative));
}

function validateZipEntries(output) {
  const entries = String(output || "")
    .split(/\r?\n/)
    .map((entry) => entry.trim().replace(/\\/g, "/"))
    .filter(Boolean);
  if (!entries.length) throw new Error("压缩包为空或无法读取。");
  for (const entry of entries) {
    const segments = entry.split("/").filter(Boolean);
    if (
      entry.startsWith("/")
      || /^[a-zA-Z]:/.test(entry)
      || segments.includes("..")
    ) {
      throw new Error(`压缩包包含不安全路径：${entry}`);
    }
  }
}

function listPdfFiles(folder) {
  const files = [];
  const visit = (current) => {
    for (const entry of fs.readdirSync(current, { withFileTypes: true })) {
      const fullPath = path.join(current, entry.name);
      if (entry.isDirectory()) {
        visit(fullPath);
      } else if (entry.isFile() && path.extname(entry.name).toLowerCase() === ".pdf") {
        files.push(fullPath);
      }
    }
  };
  visit(folder);
  return files;
}

function documentCategoryFromContent(content) {
  const compact = String(content || "").replace(/\s+/g, "");
  let invoiceScore = 0;
  let billScore = 0;

  // 只看票面内容，不使用压缩包目录或原文件名。发票字段通常成组出现，
  // 账单则以账号、结算周期和业务量明细为主。
  if (/(?:电子发票|增值税专用发票|增值税普通发票)/.test(compact)) invoiceScore += 8;
  if (/发票号码/.test(compact)) invoiceScore += 5;
  if (/开票日期/.test(compact)) invoiceScore += 3;
  if (/(?:纳税人识别号|统一社会信用代码)/.test(compact)) invoiceScore += 3;
  if (/价税合计/.test(compact)) invoiceScore += 3;
  if (/(?:税率\/征收率|税额)/.test(compact)) invoiceScore += 2;

  if (/(?:结算单|对账单|账单明细)/.test(compact)) billScore += 7;
  if (/账号编码/.test(compact)) billScore += 5;
  if (/账号名/.test(compact)) billScore += 3;
  if (/(?:成功条数|发送条数|计费条数)/.test(compact)) billScore += 3;
  if (/(?:结算周期|计费周期|账期)/.test(compact)) billScore += 3;
  if (/(?:结算金额|应付金额)/.test(compact)) billScore += 2;

  // 业务只有两类。内容无法提取或没有发票特征时，按结算凭据/账单处理。
  return invoiceScore > billScore && invoiceScore >= 5 ? "发票" : "账单";
}

async function partitionInvoicePdfs(sourceFolder) {
  const pdfFiles = listPdfFiles(sourceFolder);
  const billFiles = [];
  const invoiceFiles = [];
  const contentByFile = new Map();

  for (const filePath of pdfFiles) {
    let content = "";
    try {
      content = await extractPdfText(filePath);
    } catch {
      // 无文本层且 OCR 失败时，仍保留该 PDF，并按非发票结算凭据处理。
      content = "";
    }
    contentByFile.set(filePath, content);
    const category = documentCategoryFromContent(content);
    if (category === "账单") {
      billFiles.push(filePath);
    } else {
      invoiceFiles.push(filePath);
    }
  }

  return { pdfFiles, billFiles, invoiceFiles, contentByFile };
}

function findPdfTextTool() {
  const candidates = [
    process.env.PDFTOTEXT_PATH,
    "C:\\Program Files\\Git\\mingw64\\bin\\pdftotext.exe",
    "C:\\Program Files\\Git\\usr\\bin\\pdftotext.exe",
    "pdftotext.exe",
    "pdftotext"
  ].filter(Boolean);
  for (const candidate of candidates) {
    if (!path.isAbsolute(candidate) || fs.existsSync(candidate)) return candidate;
  }
  throw new Error("未找到 pdftotext，无法读取 PDF 内容。");
}

async function extractPdfText(filePath, forceOcr = false) {
  let content = "";
  try {
    content = await execFileText(findPdfTextTool(), ["-table", "-enc", "UTF-8", filePath, "-"], {
      cwd: path.dirname(filePath),
      timeout: 30000
    });
  } catch {
    content = "";
  }
  if (!forceOcr && content.replace(/\s+/g, "").length >= 20) return content;

  const ocrScript = path.join(root, "pdf-ocr.ps1");
  if (!fs.existsSync(ocrScript)) return content;
  return execFileText("powershell.exe", [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    ocrScript,
    filePath
  ], {
    cwd: root,
    timeout: 180000
  });
}

function normalizeMoney(value) {
  let normalized = String(value || "")
    .replace(/[．。]/g, ".")
    .replace(/[¥￥\s]/g, "");
  if (/^\d+[,，、]\d{1,2}$/.test(normalized)) {
    normalized = normalized.replace(/[,，、]/, ".");
  }
  normalized = normalized.replace(/[,，、]/g, "");
  const number = Number(normalized);
  return Number.isFinite(number) ? number.toFixed(2) : "";
}

function parseBillDetails(content) {
  const lines = String(content || "").replace(/\f/g, "").split(/\r?\n/);
  const header = lines.find((line) => line.includes("账号编码") && line.includes("账号名"));
  const accounts = [];
  const accountNames = [];
  let total = "";

  if (header) {
    const accountStart = header.indexOf("账号编码");
    const accountEnd = header.indexOf("账号名", accountStart + 1);
    const accountNameStart = accountEnd;
    const accountNameEnd = [
      "短信类型",
      "签名",
      "成功条数",
      "对接人",
      "单价",
      "金额",
      "合计"
    ]
      .map((label) => header.indexOf(label, accountNameStart + 1))
      .filter((index) => index > accountNameStart)
      .sort((left, right) => left - right)[0] ?? header.length;
    const totalStart = header.lastIndexOf("合计");

    for (const line of lines) {
      if (line === header) continue;
      if (accountStart >= 0 && accountEnd > accountStart && line.length > accountStart) {
        const cell = line.slice(accountStart, accountEnd).trim().replace(/\s+/g, "");
        if (/^(?:s|sn)_[a-zA-Z0-9_]+$/i.test(cell)) {
          accounts.push(cell);
          const accountName = line
            .slice(accountNameStart, accountNameEnd)
            .trim()
            .replace(/\s+/g, "");
          accountNames.push(accountName);
        } else if (/^[a-zA-Z0-9_]+$/.test(cell) && accounts.length) {
          accounts[accounts.length - 1] += cell;
        }
      }
      if (totalStart >= 0 && line.length > totalStart) {
        const totalCell = line.slice(totalStart).trim();
        const match = /^[¥￥]?\s*([\d,]+\.\d{1,2})$/.exec(totalCell);
        if (match) total = normalizeMoney(match[1]);
      }
    }
  }

  if (!total) {
    const standalone = [...String(content || "").matchAll(/^\s*([\d,，]+[.．][\d]{2})\s*$/gm)];
    if (standalone.length) total = normalizeMoney(standalone[standalone.length - 1][1]);
  }

  if (!accounts.length) {
    for (const line of lines) {
      const compact = line.replace(/\s+/g, "");
      const matches = compact.match(/(?:s|sn)_[a-zA-Z0-9_]+/gi) || [];
      accounts.push(...matches);
    }
  }

  if (!accountNames.filter(Boolean).length) {
    const compact = String(content || "").replace(/\s+/g, "");
    for (const match of compact.matchAll(/账号编码((?:s|sn)_[a-zA-Z0-9_]+)账号名(.+?)短信类型/gi)) {
      const accountIndex = accounts.findIndex((account) => account.toLowerCase() === match[1].toLowerCase());
      if (accountIndex >= 0) accountNames[accountIndex] = match[2];
    }
  }

  if (!total) {
    const compact = String(content || "").replace(/\s+/g, "");
    const labeledAmounts = [
      ...compact.matchAll(/(?:结算金额|应付金额|总金额|价税合计|合计)[:：]?[¥￥´`']?([\d,，、]+[.．,，、][\d]{1,2})/g)
    ];
    if (labeledAmounts.length) {
      total = normalizeMoney(labeledAmounts[labeledAmounts.length - 1][1]);
    }
  }

  return {
    accounts: [...new Set(accounts)],
    accountNames: [...new Set(accountNames.filter(Boolean))],
    amount: total
  };
}

function parseCsvRows(content) {
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;
  const source = String(content || "").replace(/^\uFEFF/, "");

  for (let index = 0; index < source.length; index += 1) {
    const character = source[index];
    if (quoted) {
      if (character === '"' && source[index + 1] === '"') {
        field += '"';
        index += 1;
      } else if (character === '"') {
        quoted = false;
      } else {
        field += character;
      }
    } else if (character === '"') {
      quoted = true;
    } else if (character === ",") {
      row.push(field);
      field = "";
    } else if (character === "\n") {
      row.push(field.replace(/\r$/, ""));
      rows.push(row);
      row = [];
      field = "";
    } else {
      field += character;
    }
  }

  if (field || row.length) {
    row.push(field.replace(/\r$/, ""));
    rows.push(row);
  }
  return rows;
}

function loadAccountNameMap(targetDir, month) {
  const csvPath = path.join(targetDir, `${text.billPrefix}${month.compact}.csv`);
  if (!fs.existsSync(csvPath)) return new Map();
  const rows = parseCsvRows(fs.readFileSync(csvPath, "utf8"));
  const headers = rows.shift() || [];
  const accountIndex = headers.indexOf("账号编码");
  const accountNameIndex = headers.indexOf("账号名");
  if (accountIndex < 0 || accountNameIndex < 0) return new Map();

  const mapping = new Map();
  for (const row of rows) {
    const account = String(row[accountIndex] || "").trim().toLowerCase();
    const accountName = String(row[accountNameIndex] || "").trim();
    if (account && accountName && !mapping.has(account)) mapping.set(account, accountName);
  }
  return mapping;
}

function decodeXmlEntities(value) {
  return String(value || "")
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, number) => String.fromCodePoint(Number(number)))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function xmlText(value) {
  return [...String(value || "").matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)]
    .map((match) => decodeXmlEntities(match[1]))
    .join("");
}

function cellColumnIndex(reference) {
  const letters = /^([A-Z]+)/i.exec(String(reference || ""))?.[1]?.toUpperCase();
  if (!letters) return -1;
  let index = 0;
  for (const character of letters) {
    index = index * 26 + character.charCodeAt(0) - 64;
  }
  return index - 1;
}

async function readFirstXlsxSheetRows(filePath) {
  const tempFolder = path.join(
    path.dirname(filePath),
    `.billing-xlsx-${process.pid}-${Date.now()}`
  );
  fs.mkdirSync(tempFolder, { recursive: true });
  try {
    await execFileText("tar.exe", ["-xf", filePath, "-C", tempFolder], {
      timeout: 60000
    });
    const sharedStringsPath = path.join(tempFolder, "xl", "sharedStrings.xml");
    const sharedStrings = [];
    if (fs.existsSync(sharedStringsPath)) {
      const sharedXml = fs.readFileSync(sharedStringsPath, "utf8");
      for (const match of sharedXml.matchAll(/<si(?:\s[^>]*)?>([\s\S]*?)<\/si>/g)) {
        sharedStrings.push(xmlText(match[1]));
      }
    }

    const sheetPath = path.join(tempFolder, "xl", "worksheets", "sheet1.xml");
    if (!fs.existsSync(sheetPath)) return [];
    const sheetXml = fs.readFileSync(sheetPath, "utf8");
    const rows = [];
    for (const rowMatch of sheetXml.matchAll(/<row(?:\s[^>]*)?>([\s\S]*?)<\/row>/g)) {
      const values = [];
      for (const cellMatch of rowMatch[1].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
        const attributes = cellMatch[1];
        const cellContent = cellMatch[2];
        const reference = /\br="([A-Z]+\d+)"/i.exec(attributes)?.[1];
        const column = cellColumnIndex(reference);
        if (column < 0) continue;
        const cellType = /\bt="([^"]+)"/i.exec(attributes)?.[1] || "";
        const rawValue = /<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/i.exec(cellContent)?.[1] || "";
        let value = cellType === "inlineStr"
          ? xmlText(cellContent)
          : decodeXmlEntities(rawValue);
        if (cellType === "s") {
          value = sharedStrings[Number(value)] || "";
        }
        values[column] = value;
      }
      rows.push(values);
    }
    return rows;
  } finally {
    fs.rmSync(tempFolder, { recursive: true, force: true });
  }
}

async function loadBillingRecordsFromWorkbook(targetDir, month, providerTitle) {
  const workbookPath = path.join(
    targetDir,
    `${text.filterPrefix}${providerTitle}_${month.compact}.xlsx`
  );
  if (!fs.existsSync(workbookPath)) return [];

  const rows = await readFirstXlsxSheetRows(workbookPath);
  if (!rows.length) return [];
  const headers = rows.shift().map((value) => String(value || "").trim());
  const indexes = {
    provider: headers.indexOf("通道名"),
    account: headers.indexOf("账号编码"),
    accountName: headers.indexOf("账号名"),
    contact: headers.indexOf("对接人"),
    amount: headers.indexOf("金额"),
    invoiceTitle: headers.indexOf("开票抬头")
  };
  if (Object.values(indexes).some((index) => index < 0)) return [];

  const grouped = new Map();
  for (const row of rows) {
    if (String(row[indexes.provider] || "").trim() !== providerTitle) continue;
    const contact = String(row[indexes.contact] || "").trim();
    const invoiceTitle = String(row[indexes.invoiceTitle] || "").trim();
    const account = String(row[indexes.account] || "").trim();
    const accountName = String(row[indexes.accountName] || "").trim();
    const amount = Number(String(row[indexes.amount] || "").replace(/[,，\s]/g, ""));
    if (!contact || !account || !Number.isFinite(amount) || amount <= 0) continue;

    const key = `${providerTitle}\n${invoiceTitle}\n${contact}`;
    const record = grouped.get(key) || {
      invoiceTitle,
      contact,
      amountValue: 0,
      accounts: [],
      accountNames: []
    };
    record.amountValue += amount;
    if (!record.accounts.some((item) => item.toLowerCase() === account.toLowerCase())) {
      record.accounts.push(account);
    }
    if (accountName && !record.accountNames.includes(accountName)) {
      record.accountNames.push(accountName);
    }
    grouped.set(key, record);
  }

  return [...grouped.values()]
    .map((record) => ({
      invoiceTitle: record.invoiceTitle,
      contact: record.contact,
      amount: moneyText(record.amountValue),
      accounts: record.accounts,
      accountNames: record.accountNames
    }))
    .filter((record) => record.accounts.length && Number(record.amount) > 0);
}

async function loadBillingRecords(targetDir, month, providerTitle) {
  const workbookRecords = await loadBillingRecordsFromWorkbook(targetDir, month, providerTitle);
  if (workbookRecords.length) return workbookRecords;

  const csvPath = path.join(targetDir, `${text.billPrefix}${month.compact}.csv`);
  if (!fs.existsSync(csvPath)) return [];
  const rows = parseCsvRows(fs.readFileSync(csvPath, "utf8"));
  const headers = rows.shift() || [];
  const indexes = {
    provider: headers.indexOf("通道名"),
    account: headers.indexOf("账号编码"),
    accountName: headers.indexOf("账号名"),
    contact: headers.indexOf("对接人"),
    amount: headers.indexOf("合计"),
    invoiceTitle: headers.indexOf("开票抬头")
  };
  if (Object.values(indexes).some((index) => index < 0)) return [];

  const grouped = new Map();
  for (const row of rows) {
    if (String(row[indexes.provider] || "").trim() !== providerTitle) continue;
    const amount = normalizeMoney(row[indexes.amount]);
    const contact = String(row[indexes.contact] || "").trim();
    const invoiceTitle = String(row[indexes.invoiceTitle] || "").trim();
    const account = String(row[indexes.account] || "").trim();
    const accountName = String(row[indexes.accountName] || "").trim();
    if (!amount || !contact || !account) continue;
    const key = `${invoiceTitle}\n${contact}\n${amount}`;
    const record = grouped.get(key) || {
      invoiceTitle,
      contact,
      amount,
      accounts: [],
      accountNames: []
    };
    if (!record.accounts.some((item) => item.toLowerCase() === account.toLowerCase())) {
      record.accounts.push(account);
    }
    if (accountName && !record.accountNames.includes(accountName)) {
      record.accountNames.push(accountName);
    }
    grouped.set(key, record);
  }
  return [...grouped.values()];
}

function naturalDocumentNameCompare(left, right) {
  return path.basename(left).localeCompare(path.basename(right), "zh-CN", {
    numeric: true,
    sensitivity: "base"
  });
}

function pairBillsAndInvoicesByArchiveOrder(bills, invoices) {
  const unresolvedBills = bills
    .filter((bill) => !bill.amount)
    .sort((left, right) => naturalDocumentNameCompare(left.filePath, right.filePath));
  const availableInvoices = invoices
    .filter((invoice) => invoice.amount)
    .sort((left, right) => naturalDocumentNameCompare(left.filePath, right.filePath));
  if (!unresolvedBills.length || unresolvedBills.length !== availableInvoices.length) return 0;

  const billFolders = new Set(unresolvedBills.map((bill) => path.basename(path.dirname(bill.filePath))));
  const invoiceFolders = new Set(availableInvoices.map((invoice) => path.basename(path.dirname(invoice.filePath))));
  const hasSeparateFolders = [...billFolders].some((name) => name.includes("账单"))
    && [...invoiceFolders].some((name) => name.includes("发票"))
    && ![...billFolders].some((name) => invoiceFolders.has(name));

  const sortedBills = unresolvedBills;
  const sortedInvoices = availableInvoices.sort((left, right) => {
    if (!hasSeparateFolders) return naturalDocumentNameCompare(left.filePath, right.filePath);
    return naturalDocumentNameCompare(left.filePath, right.filePath);
  });

  sortedBills.forEach((bill, index) => {
    const invoice = sortedInvoices[index];
    bill.amount = invoice.amount;
    bill.inferredFromInvoiceOrder = true;
  });
  return sortedBills.length;
}

function billingIdentityFromFilename(filePath) {
  const stem = path.basename(filePath, path.extname(filePath)).replace(/-\d+$/, "");
  const separator = stem.lastIndexOf("-");
  if (separator < 0) return { invoiceTitle: "", contact: "" };
  return {
    invoiceTitle: stem.slice(0, separator).trim(),
    contact: stem.slice(separator + 1).trim()
  };
}

function amountFromFilename(filePath) {
  const stem = path.basename(filePath, path.extname(filePath));
  const matches = [...stem.matchAll(/(\d[\d,，]*[.．]\d{1,2})/g)];
  return normalizeMoney(matches[matches.length - 1]?.[1]);
}

function findBillingRecord(filePath, amount, records) {
  const identity = billingIdentityFromFilename(filePath);
  const exactIdentity = records.filter(
    (record) => record.invoiceTitle === identity.invoiceTitle && record.contact === identity.contact
  );
  if (amount) {
    const exactAmount = exactIdentity.filter((record) => record.amount === amount);
    if (exactAmount.length === 1) return exactAmount[0];
  }
  if (!amount && exactIdentity.length === 1) return exactIdentity[0];
  if (amount && identity.contact) {
    const contactAndAmount = records.filter(
      (record) => record.contact === identity.contact && record.amount === amount
    );
    if (contactAndAmount.length === 1) return contactAndAmount[0];
  }
  if (amount && !identity.contact) {
    const amountOnly = records.filter((record) => record.amount === amount);
    if (amountOnly.length === 1) return amountOnly[0];
  }
  return null;
}

function normalizedOcrEvidence(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^\p{Script=Han}a-z0-9_]+/gu, "");
}

function ocrTokenScore(content, value) {
  const evidence = normalizedOcrEvidence(content);
  const token = normalizedOcrEvidence(value);
  if (!evidence || !token) return 0;
  if (evidence.includes(token)) return 100 + token.length;

  const characters = [...token];
  if (characters.length < 2) return evidence.includes(token) ? 10 : 0;
  let matched = 0;
  for (let index = 0; index < characters.length - 1; index += 1) {
    if (evidence.includes(characters.slice(index, index + 2).join(""))) matched += 1;
  }
  return (matched / (characters.length - 1)) * 30;
}

function amountNearTotalScore(content, amount) {
  const compact = String(content || "").replace(/\s+/g, "");
  const labels = ["合计", "台计", "合汁"];
  const labelIndex = labels
    .map((label) => compact.indexOf(label))
    .filter((index) => index >= 0)
    .sort((left, right) => left - right)[0];
  if (labelIndex === undefined) return 0;

  const window = compact.slice(Math.max(0, labelIndex - 35), labelIndex + 45);
  const digits = window.replace(/\D/g, "");
  const target = String(amount || "").replace(/\D/g, "");
  if (!target || !digits) return 0;
  if (digits.includes(target)) return 250;

  for (let length = target.length - 1; length >= 3; length -= 1) {
    for (let start = 0; start + length <= target.length; start += 1) {
      if (digits.includes(target.slice(start, start + length))) {
        return (length / target.length) * 180;
      }
    }
  }
  return 0;
}

function billingRecordEvidenceScore(content, record) {
  const invoiceTitleScore = ocrTokenScore(content, record.invoiceTitle);
  const accountNameScore = Math.max(0, ...record.accountNames.map((name) => ocrTokenScore(content, name)));
  const accountScore = Math.max(0, ...record.accounts.map((account) => ocrTokenScore(content, account)));
  return invoiceTitleScore
    + accountNameScore * 1.5
    + accountScore * 2
    + amountNearTotalScore(content, record.amount);
}

function assignBillingRecordsFromInvoiceSet(bills, invoices, billingRecords) {
  const invoiceAmounts = new Set(invoices.map((invoice) => invoice.amount).filter(Boolean));
  const usedAmounts = new Set(
    bills
      .filter((bill) => bill.accounts?.length && bill.amount && (bill.isTechnologyDepartment || bill.accountNames?.length))
      .map((bill) => bill.amount)
  );
  const unresolvedBills = bills.filter(
    (bill) => !bill.accounts?.length || !bill.amount || (!bill.isTechnologyDepartment && !bill.accountNames?.length)
  );
  const candidates = billingRecords.filter(
    (record) => invoiceAmounts.has(record.amount) && !usedAmounts.has(record.amount)
  );

  // Only infer a one-to-one pairing when the invoice set and unresolved bill set
  // account for exactly the same settlement totals.
  if (!unresolvedBills.length || candidates.length !== unresolvedBills.length) return;

  const scores = unresolvedBills.map((bill) =>
    candidates.map((record) => billingRecordEvidenceScore(bill.content, record))
  );
  let assignment = [];

  if (candidates.length <= 16) {
    let states = new Map([[0, { score: 0, records: [] }]]);
    for (let billIndex = 0; billIndex < unresolvedBills.length; billIndex += 1) {
      const nextStates = new Map();
      for (const [mask, state] of states) {
        for (let recordIndex = 0; recordIndex < candidates.length; recordIndex += 1) {
          const bit = 1 << recordIndex;
          if (mask & bit) continue;
          const nextMask = mask | bit;
          const nextScore = state.score + scores[billIndex][recordIndex];
          const previous = nextStates.get(nextMask);
          if (!previous || nextScore > previous.score) {
            nextStates.set(nextMask, {
              score: nextScore,
              records: [...state.records, recordIndex]
            });
          }
        }
      }
      states = nextStates;
    }
    assignment = states.get((1 << candidates.length) - 1)?.records || [];
  } else {
    const pairs = [];
    scores.forEach((row, billIndex) => {
      row.forEach((score, recordIndex) => pairs.push({ billIndex, recordIndex, score }));
    });
    pairs.sort((left, right) => right.score - left.score);
    const assignedBills = new Set();
    const assignedRecords = new Set();
    assignment = new Array(unresolvedBills.length);
    for (const pair of pairs) {
      if (assignedBills.has(pair.billIndex) || assignedRecords.has(pair.recordIndex)) continue;
      assignment[pair.billIndex] = pair.recordIndex;
      assignedBills.add(pair.billIndex);
      assignedRecords.add(pair.recordIndex);
    }
  }

  assignment.forEach((recordIndex, billIndex) => {
    const bill = unresolvedBills[billIndex];
    const record = candidates[recordIndex];
    if (record.contact === "顾佳雯") {
      bill.accounts = ["科技部"];
      bill.accountNames = [];
      bill.isTechnologyDepartment = true;
    } else {
      bill.accounts = [...record.accounts];
      bill.accountNames = [...record.accountNames];
    }
    bill.amount = record.amount;
    bill.inferredFromInvoiceSet = true;
  });
}

function parseInvoiceDetails(content) {
  const compact = String(content || "").replace(/\s+/g, "");
  const match = /[（(]小写[）)][¥￥´`']?([\d,，]+(?:[.．][\d]{1,2})?)/.exec(compact)
    || /价税合计[^\d]{0,40}[¥￥´`']?([\d,，、]+[.．,，、][\d]{1,2})/.exec(compact);
  return { amount: normalizeMoney(match?.[1]) };
}

function documentPairKey(filePath) {
  return path.basename(filePath, path.extname(filePath))
    .replace(/^(?:账单|发票)[-_]/, "")
    .replace(/-(?:账单|发票)(?:\s*\(\d+\))?$/, "")
    .replace(/-(?:账单|发票)-/g, "-")
    .replace(/-+/g, "-")
    .trim()
    .toLowerCase();
}

function cleanFilenamePart(value) {
  return String(value || "")
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "_")
    .replace(/\s+/g, "")
    .replace(/-+/g, "-")
    .replace(/^[-_.]+|[-_.]+$/g, "");
}

function renamedPdfPath(filePath, month, accounts, amount, category, accountNames = []) {
  const accountText = cleanFilenamePart(accounts.join("+"));
  const accountNameText = cleanFilenamePart(accountNames.join("+"));
  const suffix = accountNameText ? `-${accountNameText}` : "";
  const fileName = `${month.compact}-${accountText}-${amount}-${category}${suffix}.pdf`;
  const target = path.join(path.dirname(filePath), fileName);
  if (path.resolve(target) === path.resolve(filePath)) return filePath;
  return uniqueFileSystemPath(target);
}

function findLatestInvoiceFolder(targetDir, provider) {
  if (!fs.existsSync(targetDir)) return null;
  return fs.readdirSync(targetDir, { withFileTypes: true })
    .filter((entry) => entry.isDirectory() && !entry.name.startsWith(".") && entry.name.includes(provider.title))
    .map((entry) => {
      const folder = path.join(targetDir, entry.name);
      let hasPdfs = false;
      try {
        hasPdfs = listPdfFiles(folder).length > 0;
      } catch {
        hasPdfs = false;
      }
      return { folder, hasPdfs, mtimeMs: fs.statSync(folder).mtimeMs };
    })
    .filter((item) => item.hasPdfs)
    .sort((left, right) => right.mtimeMs - left.mtimeMs)[0]?.folder || null;
}

async function archiveInvoicePackage(req, res, forcedProviderKey = "") {
  try {
    const body = await readBodyBuffer(req);
    const upload = parseMultipartUpload(body, req.headers["content-type"]);
    const month = parseMonth(upload?.fields?.month);
    const providerKey = forcedProviderKey || String(upload?.fields?.provider || "");
    const provider = invoiceProviders[providerKey];
    const archive = upload?.files?.archive || upload?.file;
    if (!provider) {
      sendJson(res, 400, { message: "供应商类型无效。" });
      return;
    }
    if (!month) {
      sendJson(res, 400, { message: text.invalidMonth });
      return;
    }
    if (!archive?.data?.length) {
      sendJson(res, 400, { message: `请选择${provider.title}账单发票 ZIP 压缩包。` });
      return;
    }

    const originalName = safeFilename(archive.filename, `${provider.title}账单发票_${month.compact}.zip`);
    if (path.extname(originalName).toLowerCase() !== ".zip") {
      sendJson(res, 400, { message: "当前仅支持 ZIP 压缩包。" });
      return;
    }

    const targetDir = path.resolve(dataRoot, month.folderName);
    fs.mkdirSync(targetDir, { recursive: true });
    const archivePath = uniqueFileSystemPath(path.join(targetDir, originalName));
    const folderBase = `${month.compact}-${provider.title}-账单发票`;
    const sourceFolder = uniqueFileSystemPath(path.join(targetDir, folderBase));
    fs.writeFileSync(archivePath, archive.data);
    fs.mkdirSync(sourceFolder, { recursive: true });

    try {
      const entries = await execFileText("tar.exe", ["-tf", archivePath], { timeout: 30000 });
      validateZipEntries(entries);
      await execFileText("tar.exe", ["-xf", archivePath, "-C", sourceFolder], { timeout: 120000 });
    } catch (error) {
      fs.rmSync(sourceFolder, { recursive: true, force: true });
      fs.rmSync(archivePath, { force: true });
      throw error;
    }

    const pdfFiles = listPdfFiles(sourceFolder);
    sendJson(res, 200, {
      message: "压缩包已归档并解压。",
      provider: providerKey,
      providerTitle: provider.title,
      folder: targetDir,
      archivePath,
      sourceFolder,
      pdfCount: pdfFiles.length
    });
  } catch (error) {
    sendJson(res, 500, { message: error.message || "压缩包归档失败。" });
  }
}

async function organizeInvoicePdfs(req, res, forcedProviderKey = "") {
  try {
    const payload = JSON.parse(await readBody(req) || "{}");
    const month = parseMonth(payload.month);
    const providerKey = forcedProviderKey || String(payload.provider || "");
    const provider = invoiceProviders[providerKey];
    if (!provider) {
      sendJson(res, 400, { message: "供应商类型无效。" });
      return;
    }
    if (!month) {
      sendJson(res, 400, { message: text.invalidMonth });
      return;
    }

    const targetDir = path.resolve(dataRoot, month.folderName);
    const sourceFolder = payload.sourceFolder
      ? path.resolve(String(payload.sourceFolder))
      : findLatestInvoiceFolder(targetDir, provider);
    if (!sourceFolder || !isPathInside(targetDir, sourceFolder)) {
      sendJson(res, 400, { message: `未找到该账期已解压的${provider.title}账单发票目录，请先上传并解压压缩包。` });
      return;
    }
    if (!fs.existsSync(sourceFolder) || !fs.statSync(sourceFolder).isDirectory()) {
      sendJson(res, 404, { message: "已解压目录不存在，请重新上传压缩包。" });
      return;
    }

    const { billFiles, invoiceFiles, contentByFile } = await partitionInvoicePdfs(sourceFolder);
    if (!billFiles.length && !invoiceFiles.length) {
      sendJson(res, 400, { message: "解压目录中没有找到 PDF 文件。" });
      return;
    }

    const results = [];
    const bills = [];
    const accountNameMap = loadAccountNameMap(targetDir, month);
    const billingRecords = await loadBillingRecords(targetDir, month, provider.title);
    const technologyPattern = providerKey === "jingdong" ? /科技部/ : /顾佳雯|科技部/;
    const invoices = [];
    const invoiceByKey = new Map();
    for (const filePath of invoiceFiles) {
      try {
        const content = contentByFile.get(filePath) || "";
        const details = {
          filePath,
          key: documentPairKey(filePath),
          content,
          isTechnologyDepartment: technologyPattern.test(`${path.basename(filePath)}\n${content}`),
          ...parseInvoiceDetails(content)
        };
        invoices.push(details);
        if (details.key) invoiceByKey.set(details.key, details);
      } catch (error) {
        invoices.push({
          filePath,
          key: documentPairKey(filePath),
          content: "",
          amount: "",
          isTechnologyDepartment: technologyPattern.test(path.basename(filePath)),
          error: error.message
        });
      }
    }

    for (const filePath of billFiles) {
      try {
        const key = documentPairKey(filePath);
        let content = contentByFile.get(filePath) || "";
        let details = parseBillDetails(content);
        if (!details.accounts.length || !details.accountNames.length || !details.amount) {
          content = await extractPdfText(filePath, true);
          details = parseBillDetails(content);
        }
        if (!details.amount) {
          details.amount = invoiceByKey.get(key)?.amount || amountFromFilename(filePath);
        }
        const isTechnologyDepartment = technologyPattern.test(`${path.basename(filePath)}\n${content}`);
        let billingRecord = findBillingRecord(filePath, details.amount, billingRecords);
        const pairedInvoiceAmount = invoiceByKey.get(key)?.amount;
        if (pairedInvoiceAmount) {
          let invoiceBillingRecord = findBillingRecord(filePath, pairedInvoiceAmount, billingRecords);
          if (!invoiceBillingRecord) {
            const amountRecords = billingRecords.filter((record) => record.amount === pairedInvoiceAmount);
            if (amountRecords.length === 1) invoiceBillingRecord = amountRecords[0];
          }
          if (invoiceBillingRecord) billingRecord = invoiceBillingRecord;
        }
        if (isTechnologyDepartment) {
          const technologyRecords = billingRecords.filter((record) => record.contact === "顾佳雯");
          if (technologyRecords.length === 1) billingRecord = technologyRecords[0];
        }
        if (billingRecord) {
          details.accounts = billingRecord.accounts;
          details.accountNames = billingRecord.accountNames;
          details.amount = billingRecord.amount;
        }
        if (details.accounts.length && details.accountNames.length < details.accounts.length) {
          details.accountNames = [
            ...new Set([
              ...details.accountNames,
              ...details.accounts
                .map((account) => accountNameMap.get(account.toLowerCase()))
                .filter(Boolean)
            ])
          ];
        }
        if (isTechnologyDepartment) {
          details.accounts = ["科技部"];
          details.accountNames = [];
          details.isTechnologyDepartment = true;
        }
        bills.push({ filePath, key, content, ...details });
      } catch (error) {
        bills.push({
          filePath,
          key: documentPairKey(filePath),
          accounts: [],
          accountNames: [],
          amount: "",
          error: error.message
        });
      }
    }

    const pairedByArchiveOrder = pairBillsAndInvoicesByArchiveOrder(bills, invoices);
    if (!pairedByArchiveOrder) {
      assignBillingRecordsFromInvoiceSet(bills, invoices, billingRecords);
    }
    for (const bill of bills) {
      if (bill.accounts?.length || !bill.amount) continue;
      const amountRecords = billingRecords.filter((record) => record.amount === bill.amount);
      if (amountRecords.length !== 1) continue;
      bill.accounts = [...amountRecords[0].accounts];
      bill.accountNames = [...amountRecords[0].accountNames];
    }

    const billByKey = new Map();
    const billsByAmount = new Map();
    bills.forEach((bill) => {
      if (bill.key) billByKey.set(bill.key, bill);
      if (bill.amount) {
        const list = billsByAmount.get(bill.amount) || [];
        list.push(bill);
        billsByAmount.set(bill.amount, list);
      }
    });

    for (const bill of bills) {
      const oldPath = bill.filePath;
      if (!bill.accounts.length || !bill.amount || (!bill.isTechnologyDepartment && !bill.accountNames.length)) {
        results.push({
          category: "账单",
          oldName: path.basename(oldPath),
          newName: "",
          accounts: bill.accounts,
          accountNames: bill.accountNames,
          amount: bill.amount,
          ok: false,
          warning: bill.error || "未识别出完整账号编码、账号名或合计金额，已保留原文件名。"
        });
        continue;
      }
      const newPath = renamedPdfPath(
        oldPath,
        month,
        bill.accounts,
        bill.amount,
        "账单",
        bill.accountNames
      );
      if (newPath !== oldPath) fs.renameSync(oldPath, newPath);
      bill.renamedPath = newPath;
      results.push({
        category: "账单",
        oldName: path.basename(oldPath),
        newName: path.basename(newPath),
        accounts: bill.accounts,
        accountNames: bill.accountNames,
        amount: bill.amount,
        ok: true,
        warning: ""
      });
    }

    for (const invoice of invoices) {
      const filePath = invoice.filePath;
      try {
        if (invoice.error) throw new Error(invoice.error);
        const key = invoice.key;
        let pairedBill = billByKey.get(key);
        if (
          invoice.amount
          && billsByAmount.get(invoice.amount)?.length === 1
          && (!pairedBill || pairedBill.amount !== invoice.amount)
        ) {
          pairedBill = billsByAmount.get(invoice.amount)[0];
        }
        const isTechnologyDepartment = invoice.isTechnologyDepartment || pairedBill?.isTechnologyDepartment;
        const invoiceRecord = invoice.amount
          ? billingRecords.filter((record) => record.amount === invoice.amount)
          : [];
        const fallbackRecord = invoiceRecord.length === 1 ? invoiceRecord[0] : null;
        const namingAccounts = isTechnologyDepartment
          ? ["科技部"]
          : pairedBill?.accounts || fallbackRecord?.accounts;
        const namingAccountNames = isTechnologyDepartment
          ? []
          : pairedBill?.accountNames || fallbackRecord?.accountNames;
        if (!invoice.amount || !namingAccounts?.length || (!isTechnologyDepartment && !namingAccountNames?.length)) {
          results.push({
            category: "发票",
            oldName: path.basename(filePath),
            newName: "",
            accounts: namingAccounts || [],
            accountNames: namingAccountNames || [],
            amount: invoice.amount,
            ok: false,
            warning: !invoice.amount
              ? "未识别出发票价税合计，已保留原文件名。"
              : "未找到对应账单的账号编码或账号名，已保留原文件名。"
          });
          continue;
        }
        const newPath = renamedPdfPath(
          filePath,
          month,
          namingAccounts,
          invoice.amount,
          "发票",
          namingAccountNames
        );
        if (newPath !== filePath) fs.renameSync(filePath, newPath);
        results.push({
          category: "发票",
          oldName: path.basename(filePath),
          newName: path.basename(newPath),
          accounts: namingAccounts,
          accountNames: namingAccountNames,
          amount: invoice.amount,
          ok: true,
          warning: pairedBill?.amount && pairedBill.amount !== invoice.amount
            ? `对应账单金额为 ${pairedBill.amount}，与发票金额不一致。`
            : ""
        });
      } catch (error) {
        results.push({
          category: "发票",
          oldName: path.basename(filePath),
          newName: "",
          accounts: [],
          accountNames: [],
          amount: "",
          ok: false,
          warning: error.message || "PDF 读取失败，已保留原文件名。"
        });
      }
    }

    const renamedCount = results.filter((item) => item.ok).length;
    const warningCount = results.filter((item) => !item.ok || item.warning).length;
    sendJson(res, 200, {
      message: `${provider.title}账单与发票文件名整理完成。`,
      provider: providerKey,
      providerTitle: provider.title,
      folder: targetDir,
      sourceFolder,
      renamedCount,
      warningCount,
      results
    });
  } catch (error) {
    sendJson(res, 500, { message: error.message || "PDF 文件名整理失败。" });
  }
}

function parseOrganizedInvoicePdf(filePath, month) {
  const name = path.basename(filePath);
  const escapedMonth = month.compact.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = new RegExp(`^${escapedMonth}-(.+?)-(\\d+(?:\\.\\d+)?)-(账单|发票)(?:-.+)?\\.pdf$`, "i").exec(name);
  if (!match) return null;

  const accounts = match[1]
    .split("+")
    .map((account) => cleanFilenamePart(account))
    .filter(Boolean);
  if (!accounts.length) return null;

  return {
    accounts: [...new Set(accounts)],
    category: match[3]
  };
}

function mergedInvoiceTargetPath(sourcePath, targetFolder, providerTitle) {
  const originalTarget = path.join(targetFolder, path.basename(sourcePath));
  if (!fs.existsSync(originalTarget)) {
    return originalTarget;
  }

  const extension = path.extname(sourcePath);
  const stem = path.basename(sourcePath, extension);
  const providerTarget = path.join(targetFolder, `${stem}-${providerTitle}${extension}`);
  return uniqueFileSystemPath(providerTarget);
}

function buildInternalGroupAccountIndex() {
  const groupByAccount = new Map();
  for (const group of readInternalGroups()) {
    for (const account of group.accounts) {
      const key = account.toLowerCase();
      if (groupByAccount.has(key)) {
        throw new Error(`内部对账账号组存在重复账号编码：${account}。请先在“飞书推送（内部对账）”中调整账号组。`);
      }
      groupByAccount.set(key, group);
    }
  }
  return groupByAccount;
}

function mergedInvoiceFolderDefinition(month, account, groupByAccount) {
  const group = groupByAccount.get(account.toLowerCase());
  if (group) {
    return {
      key: `group:${group.id}`,
      name: `${month.compact}-${cleanFilenamePart(group.accounts.join("+"))}`,
      isGroup: true
    };
  }
  return {
    key: `account:${account.toLowerCase()}`,
    name: `${month.compact}-${cleanFilenamePart(account)}`,
    isGroup: false
  };
}

async function mergeInvoicePdfsByAccount(req, res) {
  try {
    const payload = JSON.parse(await readBody(req) || "{}");
    const month = parseMonth(payload.month);
    if (!month) {
      sendJson(res, 400, { message: text.invalidMonth });
      return;
    }

    const targetDir = path.resolve(dataRoot, month.folderName);
    const sourceFolders = Object.entries(invoiceProviders)
      .map(([providerKey, provider]) => ({
        providerKey,
        providerTitle: provider.title,
        sourceFolder: findLatestInvoiceFolder(targetDir, provider)
      }))
      .filter((item) => item.sourceFolder);
    if (!sourceFolders.length) {
      sendJson(res, 400, { message: "未找到该账期已完成命名的渠道商账单发票目录，请先整理三个渠道商的 PDF 文件名。" });
      return;
    }

    const documents = [];
    const sourceSummaries = [];

    for (const source of sourceFolders) {
      let providerMatchedCount = 0;
      for (const filePath of listPdfFiles(source.sourceFolder)) {
        const details = parseOrganizedInvoicePdf(filePath, month);
        if (!details) continue;

        providerMatchedCount += 1;
        documents.push({ filePath, accounts: details.accounts, providerTitle: source.providerTitle });
      }
      sourceSummaries.push({
        provider: source.providerTitle,
        sourceFolder: source.sourceFolder,
        matchedPdfCount: providerMatchedCount
      });
    }

    if (!documents.length) {
      sendJson(res, 400, {
        message: "已找到渠道商目录，但未发现当前账期按“月份-账号编码-金额-类型”命名的 PDF，请先完成 PDF 文件名整理。",
        sourceFolders
      });
      return;
    }

    const outputFolder = path.join(targetDir, "渠道商合并后的账单和发票");
    const groupByAccount = buildInternalGroupAccountIndex();
    fs.rmSync(outputFolder, { recursive: true, force: true });
    fs.mkdirSync(outputFolder, { recursive: true });

    let copiedCount = 0;
    const folders = new Map();
    for (const document of documents) {
      for (const account of document.accounts) {
        const definition = mergedInvoiceFolderDefinition(month, account, groupByAccount);
        const accountFolder = path.join(outputFolder, definition.name);
        if (!folders.has(definition.key)) {
          fs.mkdirSync(accountFolder, { recursive: true });
          folders.set(definition.key, { ...definition, folder: accountFolder });
        }
        const targetPath = mergedInvoiceTargetPath(document.filePath, accountFolder, document.providerTitle);
        fs.copyFileSync(document.filePath, targetPath);
        copiedCount += 1;
      }
    }

    sendJson(res, 200, {
      message: "渠道商账单和发票已按账号编码合并。",
      outputFolder,
      copiedCount,
      accountFolderCount: folders.size,
      groupFolderCount: [...folders.values()].filter((item) => item.isGroup).length,
      singleAccountFolderCount: [...folders.values()].filter((item) => !item.isGroup).length,
      matchedPdfCount: documents.length,
      sourceSummaries
    });
  } catch (error) {
    sendJson(res, 500, { message: error.message || "渠道商账单和发票合并失败。" });
  }
}

function parseMonth(value) {
  const match = /^(\d{4})-(\d{2})$/.exec(value || "");
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  if (month < 1 || month > 12) return null;
  return {
    year,
    month,
    value: `${year}-${String(month).padStart(2, "0")}`,
    compact: `${year}${String(month).padStart(2, "0")}`,
    folderName: `${year}${text.year}${month}${text.month}`
  };
}

function findBillFile(folder, compactMonth, options = {}) {
  if (!fs.existsSync(folder)) return null;
  const { extension, extensions, includePrefix, excludePrefix } = options;
  const allowedExtensions = extensions || (extension ? [extension] : null);
  const files = fs.readdirSync(folder)
    .filter((name) => name.includes(compactMonth) && !name.endsWith(".crdownload"))
    .filter((name) => !allowedExtensions || allowedExtensions.includes(path.extname(name).toLowerCase()))
    .filter((name) => !includePrefix || name.startsWith(includePrefix))
    .filter((name) => !excludePrefix || !name.startsWith(excludePrefix))
    .map((name) => {
      const filePath = path.join(folder, name);
      const stat = fs.statSync(filePath);
      return { filePath, mtimeMs: stat.mtimeMs, size: stat.size };
    })
    .filter((file) => file.size > 0)
    .sort((a, b) => b.mtimeMs - a.mtimeMs);
  return files[0]?.filePath || null;
}

function parseScriptResult(stdout, fallbackReason) {
  if (!(stdout || "").trim()) return {};
  try {
    return JSON.parse(stdout);
  } catch {
    return { ok: false, reason: fallbackReason, stdout };
  }
}

function scriptErrorMessage(error, result, fallback) {
  if (error?.cancelled) return "\u5df2\u505c\u6b62\u751f\u6210\u3002";
  if (result?.reason) return result.reason;
  if (error?.killed && error?.signal === "SIGTERM") return "操作超时，请稍后重试。";
  if (error?.code === "ENOENT") return "执行环境缺失，请确认 Node/PowerShell/Excel 已安装。";
  return error?.message || fallback;
}

function execFileJson(command, args, options, fallbackReason) {
  return new Promise((resolve) => {
    execFile(command, args, options, (error, stdout, stderr) => {
      const result = parseScriptResult(stdout, fallbackReason);
      resolve({ error, stdout, stderr, result });
    });
  });
}

function taskStoppedResponse(res, folder, extra = {}) {
  sendJson(res, 200, {
    stopped: true,
    message: "已停止生成。",
    folder,
    ...extra
  });
}

async function exportSummaryBill(req, res) {
  let payload;
  try {
    payload = JSON.parse(await readBody(req) || "{}");
  } catch {
    sendJson(res, 400, { message: text.invalidJson });
    return;
  }

  const month = parseMonth(payload.month);
  if (!month) {
    sendJson(res, 400, { message: text.invalidMonth });
    return;
  }

  const taskId = normalizeTaskId(payload.taskId);
  const targetDir = path.resolve(dataRoot, month.folderName);
  fs.mkdirSync(targetDir, { recursive: true });

  const { error, stderr, result, cancelled } = await execFileJsonTask(
    taskId,
    "导出汇总账单",
    process.execPath,
    [path.join(root, "seastar-export.js"), month.value],
    {
      cwd: root,
      timeout: 10 * 60 * 1000,
      maxBuffer: 1024 * 1024 * 5,
      windowsHide: true
    },
    "汇总账单导出脚本返回内容无法解析。"
  );

  if (cancelled) {
    taskStoppedResponse(res, targetDir);
    return;
  }

  if (error || result.ok === false) {
    sendJson(res, 502, {
      message: scriptErrorMessage(error, result, "汇总账单导出失败。"),
      folder: targetDir,
      detail: result,
      stderr
    });
    return;
  }

  if (result.hasBill === false) {
    sendJson(res, 200, {
      message: text.noBill,
      exported: false,
      folder: targetDir,
      detail: result
    });
    return;
  }

  const sourcePath = result.filePath || findBillFile(targetDir, month.compact, {
    extensions: [".csv", ".xlsx", ".xls"],
    excludePrefix: text.organizePrefix
  });
  if (!sourcePath || !fs.existsSync(sourcePath)) {
    sendJson(res, 502, {
      message: "汇总账单已触发导出，但未找到下载文件。",
      exported: false,
      folder: targetDir,
      detail: result
    });
    return;
  }

  const extension = path.extname(sourcePath).toLowerCase() || ".csv";
  const outputFile = path.join(targetDir, `${text.billPrefix}${month.compact}${extension}`);
  if (path.resolve(sourcePath) !== path.resolve(outputFile)) {
    if (fs.existsSync(outputFile)) fs.unlinkSync(outputFile);
    fs.renameSync(sourcePath, outputFile);
  }

  sendJson(res, 200, {
    message: text.exportSuccess,
    exported: true,
    folder: targetDir,
    filePath: outputFile,
    sourceFile: sourcePath,
    detail: result
  });
}

function powershellScriptTask(taskId, label, scriptName, args, fallbackReason, timeout = 180000) {
  return execFileJsonTask(taskId, label, "powershell.exe", [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    path.join(root, scriptName),
    ...args
  ], {
    cwd: root,
    timeout,
    maxBuffer: 1024 * 1024 * 10,
    windowsHide: true
  }, fallbackReason);
}

function sendScriptOutcome(res, status, fallbackMessage) {
  const { error, stderr, result, cancelled, folder } = status;
  if (cancelled) {
    taskStoppedResponse(res, folder || "");
    return false;
  }
  if (error || result.ok === false) {
    sendJson(res, 502, {
      message: scriptErrorMessage(error, result, fallbackMessage),
      folder,
      detail: result,
      stderr
    });
    return false;
  }
  return true;
}

async function organizeSummaryBill(req, res) {
  let payload;
  try {
    payload = JSON.parse(await readBody(req) || "{}");
  } catch {
    sendJson(res, 400, { message: text.invalidJson });
    return;
  }
  const month = parseMonth(payload.month);
  if (!month) {
    sendJson(res, 400, { message: text.invalidMonth });
    return;
  }
  const taskId = normalizeTaskId(payload.taskId);
  const targetDir = path.resolve(dataRoot, month.folderName);
  const sourceFile = findBillFile(targetDir, month.compact, {
    extensions: [".csv"],
    includePrefix: text.billPrefix
  }) || path.join(targetDir, `${text.billPrefix}${month.compact}.csv`);
  const outputFile = path.join(targetDir, `${text.organizePrefix}${month.compact}.xlsx`);
  if (!fs.existsSync(sourceFile)) {
    sendJson(res, 404, { message: "未找到汇总账单 CSV，请先导出汇总账单。", folder: targetDir, sourceFile });
    return;
  }
  const status = await powershellScriptTask(taskId, "整理汇总账单", "excel-organize.ps1", [
    sourceFile,
    outputFile
  ], "汇总账单整理脚本返回内容无法解析。", 300000);
  status.folder = targetDir;
  if (!sendScriptOutcome(res, status, "汇总账单整理失败。")) return;
  sendJson(res, 200, { message: text.organizeSuccess, folder: targetDir, ...status.result });
}

async function importHmtBill(req, res) {
  let payload;
  try {
    payload = JSON.parse(await readBody(req) || "{}");
  } catch {
    sendJson(res, 400, { message: text.invalidJson });
    return;
  }
  const month = parseMonth(payload.month);
  const url = String(payload.url || "").trim();
  if (!month) {
    sendJson(res, 400, { message: text.invalidMonth });
    return;
  }
  if (!/^https?:\/\//i.test(url)) {
    sendJson(res, 400, { message: "请输入有效的港澳台账单下载 URL。" });
    return;
  }
  try {
    const targetDir = path.resolve(dataRoot, month.folderName);
    fs.mkdirSync(targetDir, { recursive: true });
    const response = await fetch(url);
    if (!response.ok) throw new Error(`下载失败：HTTP ${response.status}`);
    const buffer = Buffer.from(await response.arrayBuffer());
    const originalName = safeFilename(filenameFromUrl(url, response.headers.get("content-type")), `${text.hmtPrefix}source.xlsx`);
    const extension = path.extname(originalName).toLowerCase() || ".xlsx";
    const outputFile = path.join(targetDir, `${text.hmtPrefix}${month.compact}${extension}`);
    const archivedFile = uniqueFileSystemPath(path.join(targetDir, `${text.hmtPrefix}${month.compact}_${originalName}`));
    fs.writeFileSync(outputFile, buffer);
    if (path.resolve(archivedFile) !== path.resolve(outputFile)) fs.writeFileSync(archivedFile, buffer);
    sendJson(res, 200, { message: text.downloadSuccess, folder: targetDir, filePath: outputFile, archivedFile });
  } catch (error) {
    sendJson(res, 502, { message: error.message || "港澳台账单下载失败。" });
  }
}

async function exportLatestDetail(req, res) {
  let payload;
  try {
    payload = JSON.parse(await readBody(req) || "{}");
  } catch {
    sendJson(res, 400, { message: text.invalidJson });
    return;
  }

  let url;
  try {
    url = validateLatestDetailUrl(payload.url);
  } catch (error) {
    sendJson(res, 400, { message: error.message });
    return;
  }

  try {
    const response = await fetch(url, {
      redirect: "follow",
      headers: {
        "Accept": "*/*",
        "User-Agent": "Mozilla/5.0 YH-Reconciliation/1.0"
      }
    });
    if (!response.ok) throw new Error(`下载失败：HTTP ${response.status}`);

    const contentLength = Number(response.headers.get("content-length") || 0);
    if (contentLength > maxDownloadBytes) throw new Error("最新明细文件过大，请控制在 200MB 以内。");

    const buffer = Buffer.from(await response.arrayBuffer());
    if (!buffer.length) throw new Error("下载结果为空，未保存文件。");
    if (buffer.length > maxDownloadBytes) throw new Error("最新明细文件过大，请控制在 200MB 以内。");

    const contentType = response.headers.get("content-type") || "application/octet-stream";
    const normalizedContentType = contentType.toLowerCase();
    if (normalizedContentType.includes("json") || normalizedContentType.includes("html")) {
      const detail = buffer.toString("utf8").slice(0, 300);
      throw new Error(`下载地址未返回账单文件：${detail || "响应内容无效"}`);
    }

    const fallbackExtension = extensionFromContentType(contentType);
    const dispositionName = filenameFromContentDisposition(response.headers.get("content-disposition"));
    const fallbackName = `最新明细${fallbackExtension}`;
    const originalName = safeFilename(dispositionName || fallbackName, fallbackName);
    const extension = path.extname(originalName).toLowerCase() || fallbackExtension;
    const targetDir = path.join(dataRoot, "最新明细导出");
    fs.mkdirSync(targetDir, { recursive: true });

    const outputFile = path.join(targetDir, `最新明细${extension}`);
    const archivedFile = uniqueFileSystemPath(
      path.join(targetDir, `${text.latestDetailPrefix}${timestampText()}_${originalName}`)
    );
    fs.writeFileSync(outputFile, buffer);
    fs.writeFileSync(archivedFile, buffer);

    sendJson(res, 200, {
      message: text.downloadSuccess,
      folder: targetDir,
      filePath: outputFile,
      archivedFile,
      originalName,
      contentType,
      size: buffer.length
    });
  } catch (error) {
    sendJson(res, 502, { message: error.message || "最新明细导出失败。" });
  }
}

async function importAiShortBill(req, res) {
  try {
    const body = await readBodyBuffer(req);
    const upload = parseMultipartUpload(body, req.headers["content-type"]);
    const month = parseMonth(upload?.fields?.month);
    const file = upload?.file;
    if (!month) {
      sendJson(res, 400, { message: text.invalidMonth });
      return;
    }
    if (!file || !file.data?.length) {
      sendJson(res, 400, { message: "请选择需要导入的 AI挂短账单文件。" });
      return;
    }
    const targetDir = path.resolve(dataRoot, month.folderName);
    fs.mkdirSync(targetDir, { recursive: true });
    const originalName = safeFilename(file.filename, `${text.aiShortPrefix}${month.compact}.xlsx`);
    const extension = path.extname(originalName).toLowerCase() || ".xlsx";
    const outputFile = path.join(targetDir, `${text.aiShortPrefix}${month.compact}${extension}`);
    const archivedFile = uniqueFileSystemPath(path.join(targetDir, `${text.aiShortPrefix}${month.compact}_${originalName}`));
    fs.writeFileSync(outputFile, file.data);
    if (path.resolve(archivedFile) !== path.resolve(outputFile)) fs.writeFileSync(archivedFile, file.data);
    sendJson(res, 200, { message: text.downloadSuccess, folder: targetDir, filePath: outputFile, archivedFile });
  } catch (error) {
    sendJson(res, 500, { message: error.message || "AI挂短账单导入失败。" });
  }
}

function latestWorkbook(targetDir, month, options) {
  return findBillFile(targetDir, month.compact, {
    extensions: [".xlsx", ".xls"],
    ...options
  });
}

async function appendHmtBill(req, res) {
  let payload;
  try {
    payload = JSON.parse(await readBody(req) || "{}");
  } catch {
    sendJson(res, 400, { message: text.invalidJson });
    return;
  }
  const month = parseMonth(payload.month);
  if (!month) {
    sendJson(res, 400, { message: text.invalidMonth });
    return;
  }
  const targetDir = path.resolve(dataRoot, month.folderName);
  const hmtFile = latestWorkbook(targetDir, month, { includePrefix: text.hmtPrefix });
  const targetFile = path.join(targetDir, `${text.organizePrefix}${month.compact}.xlsx`);
  if (!fs.existsSync(targetFile)) {
    sendJson(res, 404, { message: "未找到整理后的汇总账单，请先整理汇总账单。", targetFile });
    return;
  }
  if (!hmtFile) {
    sendJson(res, 404, { message: "未找到港澳台账单，请先下载或导入港澳台账单。", folder: targetDir });
    return;
  }
  const status = await powershellScriptTask(normalizeTaskId(payload.taskId), "追加港澳台账单", "hmt-append.ps1", [
    hmtFile,
    targetFile
  ], "港澳台追加脚本返回内容无法解析。", 300000);
  status.folder = targetDir;
  if (!sendScriptOutcome(res, status, "港澳台账单追加失败。")) return;
  sendJson(res, 200, { message: text.appendSuccess, folder: targetDir, ...status.result });
}

async function appendAiShortBill(req, res) {
  let payload;
  try {
    payload = JSON.parse(await readBody(req) || "{}");
  } catch {
    sendJson(res, 400, { message: text.invalidJson });
    return;
  }
  const month = parseMonth(payload.month);
  if (!month) {
    sendJson(res, 400, { message: text.invalidMonth });
    return;
  }
  const targetDir = path.resolve(dataRoot, month.folderName);
  const aiShortFile = latestWorkbook(targetDir, month, { includePrefix: text.aiShortPrefix });
  const targetFile = path.join(targetDir, `${text.organizePrefix}${month.compact}.xlsx`);
  if (!fs.existsSync(targetFile)) {
    sendJson(res, 404, { message: "未找到整理后的汇总账单，请先整理汇总账单。", targetFile });
    return;
  }
  if (!aiShortFile) {
    sendJson(res, 404, { message: "未找到 AI挂短账单，请先导入 AI挂短账单。", folder: targetDir });
    return;
  }
  const status = await powershellScriptTask(normalizeTaskId(payload.taskId), "追加AI挂短账单", "ai-short-append.ps1", [
    aiShortFile,
    targetFile
  ], "AI挂短追加脚本返回内容无法解析。", 300000);
  status.folder = targetDir;
  if (!sendScriptOutcome(res, status, "AI挂短账单追加失败。")) return;
  sendJson(res, 200, { message: text.appendSuccess, folder: targetDir, ...status.result });
}

async function filterYhBill(req, res) {
  let payload;
  try {
    payload = JSON.parse(await readBody(req) || "{}");
  } catch {
    sendJson(res, 400, { message: text.invalidJson });
    return;
  }
  const month = parseMonth(payload.month);
  const config = filterBillTypes[payload.type];
  if (!month) {
    sendJson(res, 400, { message: text.invalidMonth });
    return;
  }
  if (!config) {
    sendJson(res, 400, { message: "筛选账单类型无效。" });
    return;
  }
  const targetDir = path.resolve(dataRoot, month.folderName);
  const sourceFile = path.join(targetDir, `${text.combinedPrefix}${month.compact}.xlsx`);
  const outputFile = path.join(targetDir, `${text.filterPrefix}${config.title}_${month.compact}.xlsx`);
  if (!fs.existsSync(sourceFile)) {
    sendJson(res, 404, { message: "未找到合并账单，请先执行账单合并。", sourceFile });
    return;
  }
  const status = await powershellScriptTask(normalizeTaskId(payload.taskId), "筛选账单", "filter-yh-bill.ps1", [
    sourceFile,
    outputFile,
    config.columnName,
    config.matchValue
  ], "筛选账单脚本返回内容无法解析。", 300000);
  status.folder = targetDir;
  if (!sendScriptOutcome(res, status, "筛选账单失败。")) return;
  sendJson(res, 200, { message: text.organizeSuccess, folder: targetDir, ...status.result });
}

async function combineWorkbooks(req, res) {
  let payload;
  try {
    payload = JSON.parse(await readBody(req) || "{}");
  } catch {
    sendJson(res, 400, { message: text.invalidJson });
    return;
  }
  const month = parseMonth(payload.month);
  if (!month) {
    sendJson(res, 400, { message: text.invalidMonth });
    return;
  }
  const targetDir = path.resolve(dataRoot, month.folderName);
  const sources = [
    { sheetName: "汇总账单", filePath: path.join(targetDir, `${text.organizePrefix}${month.compact}.xlsx`) },
    { sheetName: "费用对比", filePath: path.join(targetDir, `${text.expenseComparePrefix}${month.compact}.xlsx`) }
  ].filter((item) => item.filePath && fs.existsSync(item.filePath));
  if (!sources.length) {
    sendJson(res, 404, { message: "未找到可合并的 Excel 文件。", folder: targetDir });
    return;
  }
  const outputFile = path.join(targetDir, `${text.combinedPrefix}${month.compact}.xlsx`);
  const status = await powershellScriptTask(normalizeTaskId(payload.taskId), "合并账单", "combine-workbook.ps1", [
    outputFile,
    JSON.stringify(sources)
  ], "合并账单脚本返回内容无法解析。", 300000);
  status.folder = targetDir;
  if (!sendScriptOutcome(res, status, "合并账单失败。")) return;
  sendJson(res, 200, { message: text.organizeSuccess, folder: targetDir, ...status.result });
}

async function generateFeishuThirdPartyScript(req, res) {
  let payload;
  try {
    payload = JSON.parse(await readBody(req) || "{}");
  } catch {
    sendJson(res, 400, { message: text.invalidJson });
    return;
  }

  const month = parseMonth(payload.month);
  if (!month) {
    sendJson(res, 400, { message: text.invalidMonth });
    return;
  }

  const targetDir = path.resolve(dataRoot, month.folderName);
  const sourceFile = path.join(targetDir, `${text.expenseComparePrefix}${month.compact}.xlsx`);
  const outputFile = path.join(targetDir, `飞书沟通（三方群）_${month.compact}.txt`);
  if (!fs.existsSync(sourceFile)) {
    sendJson(res, 404, { message: "未找到费用对比表，请先生成费用对比。", sourceFile });
    return;
  }

  const status = await powershellScriptTask(
    normalizeTaskId(payload.taskId),
    "生成飞书沟通（三方群）话术",
    "feishu-third-party-script.ps1",
    [sourceFile, outputFile, String(month.year), String(month.month)],
    "三方群话术生成脚本返回内容无法解析。",
    180000
  );
  status.folder = targetDir;
  if (!sendScriptOutcome(res, status, "三方群话术生成失败。")) return;
  sendJson(res, 200, {
    message: "话术生成完成。",
    folder: targetDir,
    outputFile,
    messages: status.result.messages || {},
    ...status.result
  });
}

async function generateExpenseCompare(req, res) {
  const body = await readBodyBuffer(req);
  const upload = parseMultipartUpload(body, req.headers["content-type"]);
  if (!upload) {
    sendJson(res, 400, { message: "费用对比请求格式无效，请重新选择账单文件后再试。" });
    return;
  }

  const month = parseMonth(upload.fields.month);
  if (!month) {
    sendJson(res, 400, { message: text.invalidMonth });
    return;
  }

  const requiredUploads = {
    yh: { field: "yhFile", title: "永辉整理账单" },
    jm: { field: "jmFile", title: "江西聚梦短信/港澳台账单" },
    jmAi: { field: "jmAiFile", title: "江西聚梦AI语音挂短账单" },
    beijing: { field: "beijingFile", title: "北京凌渡账单" },
    jd: { field: "jdFile", title: "京东账单" },
    tengyu: { field: "tengyuFile", title: "腾域账单" }
  };
  const missingUploads = Object.values(requiredUploads)
    .filter(({ field }) => !upload.files[field]?.data?.length)
    .map(({ title }) => title);
  if (missingUploads.length) {
    sendJson(res, 400, { message: `请选择${missingUploads.join("、")}。` });
    return;
  }

  const targetDir = path.resolve(dataRoot, month.folderName);
  fs.mkdirSync(targetDir, { recursive: true });
  const outputFile = path.join(targetDir, `${text.expenseComparePrefix}${month.compact}.xlsx`);
  const uploadDir = fs.mkdtempSync(path.join(targetDir, ".expense-compare-"));
  let status;
  try {
    const required = {};
    for (const [key, { field, title }] of Object.entries(requiredUploads)) {
      const file = upload.files[field];
      const originalName = safeFilename(file.filename, `${title}.xlsx`);
      const extension = path.extname(originalName).toLowerCase();
      if (![".xlsx", ".xls"].includes(extension)) {
        sendJson(res, 400, { message: `${title}仅支持 Excel 文件（.xlsx 或 .xls）。` });
        return;
      }
      required[key] = path.join(uploadDir, `${field}${extension}`);
      fs.writeFileSync(required[key], file.data);
    }

    status = await powershellScriptTask(normalizeTaskId(upload.fields.taskId), "生成费用对比", "expense-compare.ps1", [
      required.yh,
      required.jm,
      required.jmAi,
      required.beijing,
      required.jd,
      required.tengyu,
      outputFile
    ], "费用对比脚本返回内容无法解析。", 300000);
  } finally {
    fs.rmSync(uploadDir, { recursive: true, force: true });
  }

  status.folder = targetDir;
  if (!sendScriptOutcome(res, status, "费用对比生成失败。")) return;
  sendJson(res, 200, {
    message: text.organizeSuccess,
    folder: targetDir,
    ...status.result,
    totals: status.result
  });
}

async function generateOneClickBill(req, res) {
  let payload;
  try {
    payload = JSON.parse(await readBody(req) || "{}");
  } catch {
    sendJson(res, 400, { message: text.invalidJson });
    return;
  }
  const month = parseMonth(payload.month);
  if (!month) {
    sendJson(res, 400, { message: text.invalidMonth });
    return;
  }
  const taskId = normalizeTaskId(payload.taskId);
  const targetDir = path.resolve(dataRoot, month.folderName);
  fs.mkdirSync(targetDir, { recursive: true });
  const steps = [];
  const pushStep = (name, ok, message, extra = {}) => steps.push({ name, ok, skipped: false, message, ...extra });

  const exportStatus = await execFileJsonTask(taskId, "一键生成-导出汇总账单", process.execPath, [
    path.join(root, "seastar-export.js"),
    month.value
  ], {
    cwd: root,
    timeout: 10 * 60 * 1000,
    maxBuffer: 1024 * 1024 * 5,
    windowsHide: true
  }, "汇总账单导出脚本返回内容无法解析。");
  exportStatus.folder = targetDir;
  if (!sendScriptOutcome(res, exportStatus, "汇总账单导出失败。")) return;
  if (exportStatus.result.hasBill === false) {
    sendJson(res, 200, { message: text.noBill, exported: false, folder: targetDir, steps });
    return;
  }
  const exportedSource = exportStatus.result.filePath || findBillFile(targetDir, month.compact, { extensions: [".csv", ".xlsx", ".xls"], excludePrefix: text.organizePrefix });
  const csvFile = path.join(targetDir, `${text.billPrefix}${month.compact}${path.extname(exportedSource || ".csv").toLowerCase() || ".csv"}`);
  if (exportedSource && fs.existsSync(exportedSource) && path.resolve(exportedSource) !== path.resolve(csvFile)) {
    if (fs.existsSync(csvFile)) fs.unlinkSync(csvFile);
    fs.renameSync(exportedSource, csvFile);
  }
  pushStep("导出汇总账单", true, csvFile);

  const organizeFile = path.join(targetDir, `${text.organizePrefix}${month.compact}.xlsx`);
  const organizeStatus = await powershellScriptTask(taskId, "一键生成-整理汇总账单", "excel-organize.ps1", [
    csvFile,
    organizeFile
  ], "汇总账单整理脚本返回内容无法解析。", 300000);
  organizeStatus.folder = targetDir;
  if (!sendScriptOutcome(res, organizeStatus, "汇总账单整理失败。")) return;
  pushStep("汇总账单整理", true, organizeStatus.result.outputFile || organizeFile);

  const hmtFile = latestWorkbook(targetDir, month, { includePrefix: text.hmtPrefix });
  if (hmtFile) {
    const hmtStatus = await powershellScriptTask(taskId, "一键生成-追加港澳台账单", "hmt-append.ps1", [
      hmtFile,
      organizeFile
    ], "港澳台追加脚本返回内容无法解析。", 300000);
    hmtStatus.folder = targetDir;
    if (!sendScriptOutcome(res, hmtStatus, "港澳台账单追加失败。")) return;
    pushStep("港澳台追加", true, `${hmtStatus.result.appendedRows || 0} 行`);
  } else {
    steps.push({ name: "港澳台追加", ok: true, skipped: true, message: "未找到港澳台账单，已跳过。" });
  }

  const aiShortFile = latestWorkbook(targetDir, month, { includePrefix: text.aiShortPrefix });
  if (aiShortFile) {
    const aiStatus = await powershellScriptTask(taskId, "一键生成-追加AI挂短账单", "ai-short-append.ps1", [
      aiShortFile,
      organizeFile
    ], "AI挂短追加脚本返回内容无法解析。", 300000);
    aiStatus.folder = targetDir;
    if (!sendScriptOutcome(res, aiStatus, "AI挂短账单追加失败。")) return;
    pushStep("AI挂短追加", true, `${aiStatus.result.appendedRows || 0} 行`);
  } else {
    steps.push({ name: "AI挂短追加", ok: true, skipped: true, message: "未找到 AI挂短账单，已跳过。" });
  }

  sendJson(res, 200, {
    message: "一键生成完成。",
    folder: targetDir,
    outputFile: organizeFile,
    steps
  });
}

function normalizeInternalGroups(groups) {
  if (!Array.isArray(groups)) throw new Error("账号组数据格式不正确。");
  return groups.map((group, index) => {
    const id = Number(group.id) || index + 1;
    const accounts = [...new Set((Array.isArray(group.accounts) ? group.accounts : [])
      .map((value) => String(value || "").trim())
      .filter(Boolean))];
    const contacts = [...new Set((Array.isArray(group.contacts) ? group.contacts : [])
      .map((value) => String(value || "").trim())
      .filter(Boolean))];
    if (!accounts.length) throw new Error(`第 ${index + 1} 个账号组至少需要一个账号编码。`);
    if (!contacts.length) throw new Error(`第 ${index + 1} 个账号组至少需要一个对接人。`);
    return { id, accounts, contacts };
  }).map((group, index) => ({ ...group, id: index + 1 }));
}

function readInternalGroups() {
  if (!fs.existsSync(internalGroupsFile)) {
    fs.writeFileSync(internalGroupsFile, `${JSON.stringify(defaultInternalGroups, null, 2)}\n`, "utf8");
    return defaultInternalGroups;
  }
  try {
    return normalizeInternalGroups(JSON.parse(fs.readFileSync(internalGroupsFile, "utf8")));
  } catch {
    return defaultInternalGroups;
  }
}

function writeInternalGroups(groups) {
  const normalized = normalizeInternalGroups(groups);
  fs.writeFileSync(internalGroupsFile, `${JSON.stringify(normalized, null, 2)}\n`, "utf8");
  return normalized;
}

function withSystemInternalGroups(groups) {
  const hasTechnologyGroup = groups.some((group) =>
    group.name === "科技部"
    || group.rule === "contact"
    || group.contacts.includes("顾佳雯")
  );
  if (hasTechnologyGroup) return groups;

  const nextId = groups.reduce((max, group) => Math.max(max, Number(group.id) || 0), 0) + 1;
  return [
    ...groups,
    {
      id: nextId,
      name: "科技部",
      accounts: [],
      contacts: ["顾佳雯"],
      rule: "contact",
      matchValue: "顾佳雯",
      isSystem: true
    }
  ];
}

async function getInternalReconciliationGroups(req, res) {
  sendJson(res, 200, { groups: readInternalGroups() });
}

async function saveInternalReconciliationGroups(req, res) {
  try {
    const payload = JSON.parse(await readBody(req) || "{}");
    const groups = writeInternalGroups(payload.groups);
    sendJson(res, 200, { message: "账号组已保存。", groups });
  } catch (error) {
    sendJson(res, 400, { message: error.message || "账号组保存失败。" });
  }
}

async function checkInternalReconciliationAccounts(req, res) {
  let payload;
  try {
    payload = JSON.parse(await readBody(req) || "{}");
  } catch {
    sendJson(res, 400, { message: text.invalidJson });
    return;
  }

  const month = parseMonth(payload.month);
  if (!month) {
    sendJson(res, 400, { message: text.invalidMonth });
    return;
  }

  const taskId = normalizeTaskId(payload.taskId);
  const targetDir = path.resolve(dataRoot, month.folderName);
  const sourceFile = path.join(targetDir, `${text.combinedPrefix}${month.compact}.xlsx`);
  if (!fs.existsSync(sourceFile)) {
    sendJson(res, 404, {
      message: "未找到账单合并文件，请先生成对应账期的账单合并 Excel。",
      folder: targetDir,
      sourceFile
    });
    return;
  }

  const groups = withSystemInternalGroups(readInternalGroups());
  const groupsBase64 = Buffer.from(JSON.stringify(groups), "utf8").toString("base64");
  const status = await powershellScriptTask(
    taskId,
    "校验内部对账账号完整性",
    "internal-reconciliation-account-check.ps1",
    [sourceFile, groupsBase64],
    "账号完整性校验脚本返回内容无法解析。",
    120000
  );
  const { error, stderr, result, cancelled } = status;

  if (cancelled) {
    taskStoppedResponse(res, targetDir, { sourceFile });
    return;
  }
  if (error || result.ok === false) {
    sendJson(res, 502, {
      message: scriptErrorMessage(error, result, "账号完整性校验失败。"),
      folder: targetDir,
      sourceFile,
      detail: result,
      stderr
    });
    return;
  }

  const uncoveredCount = Number(result.uncoveredAccountCount || 0);
  const duplicateCount = Array.isArray(result.duplicateAccounts) ? result.duplicateAccounts.length : 0;
  const message = result.complete
    ? `校验完成：本月账单涉及的 ${result.billAccountCount || 0} 个账号编码已全部包含在已保存的账号组中。`
    : `校验完成：发现 ${uncoveredCount} 个本月账单账号未包含，${duplicateCount} 个账号被重复配置。`;

  sendJson(res, 200, {
    message,
    folder: targetDir,
    sourceFile,
    ...result
  });
}

function formatInternalCutoff(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  const weekdays = ["日", "一", "二", "三", "四", "五", "六"];
  return `${date.getFullYear()}.${date.getMonth() + 1}.${date.getDate()}/周${weekdays[date.getDay()]} ${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")}`;
}

function buildInternalMessage(month, group, cutoff) {
  const accounts = group.accounts.join("、");
  const accountSubject = group.name ? `${group.name}（${accounts}）` : accounts;
  const contacts = group.contacts.map((contact) => `@${contact}`).join(" ");
  const cutoffText = formatInternalCutoff(cutoff) || "账期截止时间";
  return `${contacts} 你好：

针对${month.year}年${month.month}月对${accountSubject}短信费用账单请核对确认，截止${cutoffText}之前，若无回复，默认为无异议，我们将进行供应商结算开票处理，请知悉。

（如果有开票不是这个抬头，请及时提出来）`;
}

async function generateInternalReconciliation(req, res) {
  let payload;
  try {
    payload = JSON.parse(await readBody(req) || "{}");
  } catch {
    sendJson(res, 400, { message: text.invalidJson });
    return;
  }

  const month = parseMonth(payload.month);
  if (!month) {
    sendJson(res, 400, { message: text.invalidMonth });
    return;
  }

  const cutoff = String(payload.cutoff || "");
  if (!formatInternalCutoff(cutoff)) {
    sendJson(res, 400, { message: "请选择有效的截止时间。" });
    return;
  }

  const taskId = normalizeTaskId(payload.taskId);
  const targetDir = path.resolve(dataRoot, month.folderName);
  const sourceFile = path.join(targetDir, `${text.combinedPrefix}${month.compact}.xlsx`);
  if (!fs.existsSync(sourceFile)) {
    sendJson(res, 404, {
      message: "未找到账单合并文件，请先生成对应账期的账单合并 Excel。",
      folder: targetDir,
      sourceFile
    });
    return;
  }

  const groups = withSystemInternalGroups(readInternalGroups());
  const tempDir = path.join(targetDir, `.飞书推送（内部对账）_${month.compact}_临时`);
  const assetDirName = `飞书推送（内部对账）_${month.compact}_图片`;
  const assetDir = path.join(targetDir, assetDirName);
  const outputMd = path.join(targetDir, `飞书推送（内部对账）_${month.compact}.md`);
  const outputDocx = path.join(targetDir, `飞书推送（内部对账）_${month.compact}.docx`);

  fs.rmSync(tempDir, { recursive: true, force: true });
  fs.rmSync(assetDir, { recursive: true, force: true });

  const groupsBase64 = Buffer.from(JSON.stringify(groups), "utf8").toString("base64");
  const child = execFile("powershell.exe", [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    path.join(root, "internal-reconciliation.ps1"),
    sourceFile,
    tempDir,
    groupsBase64
  ], {
    cwd: root,
    timeout: 180000,
    maxBuffer: 1024 * 1024 * 5,
    windowsHide: true
  }, (error, stdout, stderr) => {
    const cancelled = isTaskCancelled(taskId);
    finishTask(taskId);
    if (cancelled) {
      taskStoppedResponse(res, targetDir, { sourceFile, outputMd, outputDocx });
      return;
    }

    const result = parseScriptResult(stdout, "内部对账截图脚本返回内容无法解析。");
    if (error || result.ok === false) {
      sendJson(res, 502, {
        message: scriptErrorMessage(error, result, "内部对账截图生成失败。"),
        folder: targetDir,
        sourceFile,
        outputMd,
        outputDocx,
        detail: result,
        stderr
      });
      return;
    }

    const resultById = new Map((result.groups || []).map((item) => [Number(item.id), item]));
    const messages = groups.map((group) => {
      const generated = resultById.get(Number(group.id)) || {};
      return {
        ...group,
        ...generated,
        hasImage: Boolean(generated.imageFile && fs.existsSync(generated.imageFile)),
        message: buildInternalMessage(month, group, cutoff),
        warning: generated.missingAccounts?.length
          ? `未匹配账号：${generated.missingAccounts.join("、")}`
          : generated.matchedRows
            ? ""
            : "该账号组在当前账期没有匹配数据。"
      };
    });

    const markdownParts = [];
    for (const item of messages) {
      markdownParts.push(item.name
        ? `${item.name}（${item.accounts.join("、")}）：`
        : `${item.accounts.join("、")}：`);
      markdownParts.push("");
      if (item.hasImage) {
        const imageName = `内部对账_${item.id}.png`;
        const sourceImage = item.imageFile || path.join(tempDir, imageName);
        const destImage = path.join(assetDir, imageName);
        fs.mkdirSync(assetDir, { recursive: true });
        if (fs.existsSync(sourceImage)) {
          fs.copyFileSync(sourceImage, destImage);
        }
        item.assetFile = destImage;
        markdownParts.push(`![image-internal-reconciliation-${item.id}](${assetDirName}/${imageName})`);
        markdownParts.push("");
      } else {
        markdownParts.push(`> 截图：账号编码 ${item.accounts.join("、")} 在当前账期没有匹配数据。`);
        markdownParts.push("");
      }
      markdownParts.push(item.message);
      markdownParts.push("");
      markdownParts.push("");
    }

    fs.writeFileSync(outputMd, `${markdownParts.join("\n")}\n`, "utf8");

    execFile("powershell.exe", [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      path.join(root, "internal-reconciliation-docx.ps1"),
      outputDocx,
      JSON.stringify(messages)
    ], {
      cwd: root,
      timeout: 180000,
      maxBuffer: 1024 * 1024 * 5,
      windowsHide: true
    }, (docxError, docxStdout, docxStderr) => {
      if (docxError) {
        sendJson(res, 502, {
          message: scriptErrorMessage(docxError, parseScriptResult(docxStdout, "Word 文档脚本返回内容无法解析。"), "Word 版内部对账生成失败。"),
          folder: targetDir,
          markdownFile: outputMd,
          docxFile: outputDocx,
          sourceFile,
          detail: docxStdout,
          stderr: docxStderr
        });
        return;
      }

      sendJson(res, 200, {
        message: "飞书推送（内部对账）生成完成。",
        folder: targetDir,
        outputFile: outputDocx,
        markdownFile: outputMd,
        docxFile: outputDocx,
        sourceFile,
        groups: messages
      });
    });
  });
  registerTask(taskId, child, "生成飞书推送（内部对账）");
}

function normalizeTaskId(value) {
  const id = String(value || "").trim();
  return /^[a-zA-Z0-9_-]{8,80}$/.test(id) ? id : null;
}

function isTaskCancelled(taskId) {
  return Boolean(taskId && activeTasks.get(taskId)?.cancelled);
}

function registerTask(taskId, child, label) {
  if (!taskId) return;
  activeTasks.set(taskId, { child, label, cancelled: false });
}

function finishTask(taskId) {
  if (taskId) activeTasks.delete(taskId);
}

function stopChildProcess(child) {
  if (!child || !child.pid) return;
  if (process.platform === "win32") {
    execFile("taskkill", ["/PID", String(child.pid), "/T", "/F"], { windowsHide: true }, () => {});
    return;
  }
  child.kill("SIGTERM");
}

function execFileJsonTask(taskId, label, command, args, options, fallbackReason) {
  return new Promise((resolve) => {
    const child = execFile(command, args, options, (error, stdout, stderr) => {
      const cancelled = isTaskCancelled(taskId);
      finishTask(taskId);
      const result = parseScriptResult(stdout, fallbackReason);
      resolve({
        error: cancelled ? { cancelled: true, message: "\u5df2\u505c\u6b62\u751f\u6210\u3002" } : error,
        stdout,
        stderr,
        result,
        cancelled
      });
    });
    registerTask(taskId, child, label);
  });
}

async function stopTask(req, res) {
  let payload;
  try {
    payload = JSON.parse(await readBody(req) || "{}");
  } catch {
    sendJson(res, 400, { message: text.invalidJson });
    return;
  }

  const taskId = normalizeTaskId(payload.taskId);
  if (!taskId) {
    sendJson(res, 400, { message: "\u4efb\u52a1\u6807\u8bc6\u65e0\u6548\u3002" });
    return;
  }

  const task = activeTasks.get(taskId);
  if (!task) {
    sendJson(res, 404, { message: "\u6ca1\u6709\u627e\u5230\u6b63\u5728\u8fd0\u884c\u7684\u751f\u6210\u4efb\u52a1\u3002" });
    return;
  }

  task.cancelled = true;
  stopChildProcess(task.child);
  sendJson(res, 200, { message: "已发送停止指令。" });
}

function serveGeneratedImage(req, res) {
  let filePath;
  try {
    const url = new URL(req.url, `http://${host}:${port}`);
    filePath = path.resolve(url.searchParams.get("file") || "");
  } catch {
    res.writeHead(400, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("图片路径无效。");
    return;
  }

  const relativePath = path.relative(dataRoot, filePath);
  if (!filePath || relativePath.startsWith("..") || path.isAbsolute(relativePath) || path.extname(filePath).toLowerCase() !== ".png") {
    res.writeHead(403, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("禁止访问该图片。");
    return;
  }

  fs.readFile(filePath, (error, content) => {
    if (error) {
      res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      res.end("图片不存在。");
      return;
    }
    res.writeHead(200, {
      "Content-Type": "image/png",
      "Access-Control-Allow-Origin": "*",
      "Cache-Control": "no-store"
    });
    res.end(content);
  });
}

async function openGeneratedFile(req, res) {
  let payload;
  try {
    payload = JSON.parse(await readBody(req) || "{}");
  } catch {
    sendJson(res, 400, { message: text.invalidJson });
    return;
  }

  let filePath = path.resolve(String(payload.filePath || ""));
  const relativePath = path.relative(dataRoot, filePath);
  if (!filePath || relativePath.startsWith("..") || path.isAbsolute(relativePath)) {
    sendJson(res, 403, { message: "禁止打开该文件。" });
    return;
  }
  if (!fs.existsSync(filePath) && fs.existsSync(path.dirname(filePath))) {
    filePath = path.dirname(filePath);
  }
  if (!fs.existsSync(filePath)) {
    sendJson(res, 404, { message: "文件不存在。" });
    return;
  }

  const openScript = path.join(root, "open-generated-file.ps1");
  const targetPathFile = path.join(
    process.env.TEMP || process.env.TMP || root,
    `yh-open-target-${process.pid}-${Date.now()}.txt`
  );
  try {
    fs.writeFileSync(targetPathFile, filePath, "utf8");
    await execFileText("powershell.exe", [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      openScript,
      "-TargetPathFile",
      targetPathFile
    ], {
      cwd: root,
      timeout: 15000
    });
    sendJson(res, 200, { message: "已打开文件。", filePath });
  } catch (error) {
    sendJson(res, 500, { message: `打开文件失败：${error.message}` });
  } finally {
    fs.rmSync(targetPathFile, { force: true });
  }
}

function serveStatic(req, res) {
  const url = new URL(req.url, `http://${host}:${port}`);
  const requested = decodeURIComponent(url.pathname === "/" ? "/index.html" : url.pathname);
  const filePath = path.normalize(path.join(root, requested));
  const relativePath = path.relative(root, filePath);

  if (relativePath.startsWith("..") || path.isAbsolute(relativePath)) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }

  fs.readFile(filePath, (error, content) => {
    if (error) {
      res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      res.end("\u9875\u9762\u4e0d\u5b58\u5728");
      return;
    }
    res.writeHead(200, { "Content-Type": types[path.extname(filePath)] || "application/octet-stream" });
    res.end(content);
  });
}

const server = http.createServer((req, res) => {
  if (req.method === "OPTIONS") {
    res.writeHead(204, {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET,POST,PUT,OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type"
    });
    res.end();
    return;
  }
  if (req.method === "GET" && req.url === "/api/health") {
    sendJson(res, 200, {
      ok: true,
      service: "yh-sms-reconciliation",
      port,
      pid: process.pid,
      dataDirectory: dataRoot
    });
    return;
  }
  if (req.method === "GET" && req.url.startsWith("/api/generated-image?")) {
    serveGeneratedImage(req, res);
    return;
  }
  if (req.method === "GET" && req.url === "/api/internal-reconciliation-groups") {
    getInternalReconciliationGroups(req, res);
    return;
  }
  if (req.method === "PUT" && req.url === "/api/internal-reconciliation-groups") {
    saveInternalReconciliationGroups(req, res);
    return;
  }
  if (req.method === "POST" && req.url === "/api/check-internal-reconciliation-accounts") {
    checkInternalReconciliationAccounts(req, res);
    return;
  }
  if (req.method === "POST" && req.url === "/api/generate-internal-reconciliation") {
    generateInternalReconciliation(req, res);
    return;
  }
  if (req.method === "POST" && req.url === "/api/archive-jiangxi-invoice-package") {
    archiveInvoicePackage(req, res, "jiangxiJumeng");
    return;
  }
  if (req.method === "POST" && req.url === "/api/organize-jiangxi-invoice-pdfs") {
    organizeInvoicePdfs(req, res, "jiangxiJumeng");
    return;
  }
  if (req.method === "POST" && req.url === "/api/archive-invoice-package") {
    archiveInvoicePackage(req, res);
    return;
  }
  if (req.method === "POST" && req.url === "/api/organize-invoice-pdfs") {
    organizeInvoicePdfs(req, res);
    return;
  }
  if (req.method === "POST" && req.url === "/api/merge-invoice-pdfs-by-account") {
    mergeInvoicePdfsByAccount(req, res);
    return;
  }
  if (req.method === "POST" && req.url === "/api/export-summary-bill") {
    runEndpoint(typeof exportSummaryBill === "function" ? exportSummaryBill : null, req, res, "导出汇总账单");
    return;
  }
  if (req.method === "POST" && req.url === "/api/organize-summary-bill") {
    runEndpoint(typeof organizeSummaryBill === "function" ? organizeSummaryBill : null, req, res, "整理汇总账单");
    return;
  }
  if (req.method === "POST" && req.url === "/api/import-hmt-bill") {
    runEndpoint(typeof importHmtBill === "function" ? importHmtBill : null, req, res, "导入港澳台账单");
    return;
  }
  if (req.method === "POST" && req.url === "/api/export-latest-detail") {
    runEndpoint(typeof exportLatestDetail === "function" ? exportLatestDetail : null, req, res, "最新明细导出");
    return;
  }
  if (req.method === "POST" && req.url === "/api/append-hmt-bill") {
    runEndpoint(typeof appendHmtBill === "function" ? appendHmtBill : null, req, res, "追加港澳台账单");
    return;
  }
  if (req.method === "POST" && req.url === "/api/import-ai-short-bill") {
    runEndpoint(typeof importAiShortBill === "function" ? importAiShortBill : null, req, res, "导入AI挂短账单");
    return;
  }
  if (req.method === "POST" && req.url === "/api/append-ai-short-bill") {
    runEndpoint(typeof appendAiShortBill === "function" ? appendAiShortBill : null, req, res, "追加AI挂短账单");
    return;
  }
  if (req.method === "POST" && req.url === "/api/generate-one-click-bill") {
    runEndpoint(typeof generateOneClickBill === "function" ? generateOneClickBill : null, req, res, "一键生成汇总账单");
    return;
  }
  if (req.method === "POST" && req.url === "/api/generate-expense-compare") {
    runEndpoint(typeof generateExpenseCompare === "function" ? generateExpenseCompare : null, req, res, "生成费用对比");
    return;
  }
  if (req.method === "POST" && req.url === "/api/filter-yh-bill") {
    runEndpoint(typeof filterYhBill === "function" ? filterYhBill : null, req, res, "筛选账单");
    return;
  }
  if (req.method === "POST" && req.url === "/api/combine-workbooks") {
    runEndpoint(typeof combineWorkbooks === "function" ? combineWorkbooks : null, req, res, "合并账单");
    return;
  }
  if (req.method === "POST" && req.url === "/api/generate-feishu-third-party-script") {
    runEndpoint(typeof generateFeishuThirdPartyScript === "function" ? generateFeishuThirdPartyScript : null, req, res, "生成三方群话术");
    return;
  }
  if (req.method === "POST" && req.url === "/api/open-generated-file") {
    openGeneratedFile(req, res);
    return;
  }
  if (req.method === "POST" && req.url === "/api/stop-task") {
    stopTask(req, res);
    return;
  }
  serveStatic(req, res);
});

server.listen(port, host, () => {
  console.log(`YH SMS reconciliation started: http://${host}:${port}/`);
});
