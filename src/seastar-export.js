const fs = require("fs");
const path = require("path");
const { execFileSync, spawn } = require("child_process");

const ROOT = __dirname;
const CONFIG_PATH = process.env.YH_CONFIG_PATH || path.join(ROOT, "..", "config", "local.json");

function readConfig() {
  try {
    return JSON.parse(fs.readFileSync(CONFIG_PATH, "utf8").replace(/^\uFEFF/, ""));
  } catch {
    return {};
  }
}

const config = readConfig();
const seaStarConfig = config.seastar || {};
const EDGE_PATH = process.env.EDGE_PATH
  || config.edgePath
  || "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe";
const DEBUG_PORT = Number(process.env.SEASTAR_DEBUG_PORT || 9333);
const DATA_ROOT = path.resolve(process.env.YH_DATA_DIR || config.dataDirectory || path.join(ROOT, "..", "data"));
const PROFILE_DIR = path.resolve(
  process.env.YH_EDGE_PROFILE_DIR
  || seaStarConfig.browserProfileDirectory
  || path.join(DATA_ROOT, ".edge-profile")
);
const SEASTAR_HOME_URL = "https://seastar.yonghuivip.com/";
const SMS_BILL_URL = "https://seastar.yonghuivip.com/itservice/sms/sms_bill";
const HEADLESS = process.env.SEASTAR_HEADLESS !== "0";
const SHOW_BROWSER = process.env.SEASTAR_SHOW_BROWSER === "1";
const LOGIN_WAIT_MS = Number(process.env.SEASTAR_LOGIN_WAIT_MS || 5 * 60 * 1000);
const monthArg = process.argv[2];

const TEXT = {
  basic: "\u57fa\u7840\u670d\u52a1",
  monthly: "\u6708\u5ea6\u5bf9\u8d26",
  business: "\u4e1a\u52a1\u65b9",
  query: "\u67e5\u8be2",
  export: "\u5bfc\u51fa",
  noData: ["\u6682\u65e0\u6570\u636e", "\u65e0\u6570\u636e", "\u6ca1\u6709\u6570\u636e"],
  months: [
    "\u4e00\u6708", "\u4e8c\u6708", "\u4e09\u6708", "\u56db\u6708",
    "\u4e94\u6708", "\u516d\u6708", "\u4e03\u6708", "\u516b\u6708",
    "\u4e5d\u6708", "\u5341\u6708", "\u5341\u4e00\u6708", "\u5341\u4e8c\u6708"
  ]
};

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
    folderName: `${year}\u5e74${month}\u6708`
  };
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function getSeastarCredentials() {
  const username = process.env.SEASTAR_USERNAME || seaStarConfig.username || seaStarConfig.account || "";
  const password = process.env.SEASTAR_PASSWORD || seaStarConfig.password || "";
  return { username, password, hasCredentials: Boolean(username && password) };
}

function runPowerShell(command, options = {}) {
  const encoded = Buffer.from(command, "utf16le").toString("base64");
  return execFileSync("powershell.exe", ["-NoProfile", "-EncodedCommand", encoded], {
    stdio: "ignore",
    windowsHide: true,
    ...options
  });
}

async function requestJson(url, options) {
  const response = await fetch(url, options);
  if (!response.ok) throw new Error(`${url} returned ${response.status}`);
  return response.json();
}

function stopManagedEdge() {
  try {
    const escapedProfile = PROFILE_DIR.replace(/\\/g, "\\\\");
    const command = [
      "$items = Get-CimInstance Win32_Process | Where-Object {",
      `$_.Name -match 'msedge' -and $_.CommandLine -match '--remote-debugging-port=${DEBUG_PORT}'`,
      `-and $_.CommandLine -match '${escapedProfile}'`,
      "};",
      "foreach($p in $items){",
      "try { Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue } catch {}",
      "try { taskkill /PID $p.ProcessId /T /F 2>$null | Out-Null } catch {}",
      "}"
    ].join(" ");
    runPowerShell(command);
  } catch {}
}

function startEdge(headless = HEADLESS) {
  if (!fs.existsSync(EDGE_PATH)) throw new Error(`Edge not found: ${EDGE_PATH}`);
  fs.mkdirSync(PROFILE_DIR, { recursive: true });
  const args = [
    `--remote-debugging-port=${DEBUG_PORT}`,
    `--user-data-dir=${PROFILE_DIR}`,
    "--no-first-run",
    "--disable-popup-blocking",
    "--disable-gpu",
    "--disable-background-timer-throttling",
    "--new-window",
    "--window-size=1440,1000",
    SEASTAR_HOME_URL
  ];
  if (headless) {
    args.splice(2, 0, "--headless=new");
    spawn(EDGE_PATH, args, { detached: true, stdio: "ignore", windowsHide: true }).unref();
  } else {
    args.splice(args.length - 1, 0, SHOW_BROWSER ? "--window-position=80,80" : "--window-position=-32000,-32000");
    const psQuote = (value) => `'${String(value).replace(/'/g, "''")}'`;
    const command = `Start-Process -FilePath ${psQuote(EDGE_PATH)} -ArgumentList ${args.map(psQuote).join(",")}`;
    runPowerShell(command, { timeout: 10000 });
  }
}

async function waitForDebugPort() {
  const started = Date.now();
  while (Date.now() - started < 25000) {
    try {
      await requestJson(`http://127.0.0.1:${DEBUG_PORT}/json/version`);
      return;
    } catch {
      await sleep(250);
    }
  }
  throw new Error("Edge automation start timeout.");
}

async function startFreshEdge() {
  stopManagedEdge();
  await sleep(500);
  startEdge(HEADLESS);
  try {
    await waitForDebugPort();
  } catch (error) {
    if (!HEADLESS) throw error;
    stopManagedEdge();
    await sleep(800);
    startEdge(false);
    await waitForDebugPort();
  }
}

async function getPageSocketUrl() {
  const targets = await requestJson(`http://127.0.0.1:${DEBUG_PORT}/json`);
  const pages = targets.filter((item) => item.type === "page" && item.webSocketDebuggerUrl);
  const bill = pages.find((item) => (item.url || "").includes("/itservice/sms/sms_bill"));
  const page = bill || pages[0];
  if (!page) throw new Error("No controllable Edge page.");
  return page.webSocketDebuggerUrl;
}

class CdpClient {
  constructor(socketUrl) {
    this.socket = new WebSocket(socketUrl);
    this.nextId = 1;
    this.pending = new Map();
  }

  async open() {
    await new Promise((resolve, reject) => {
      this.socket.addEventListener("open", resolve, { once: true });
      this.socket.addEventListener("error", reject, { once: true });
      this.socket.addEventListener("message", (event) => {
        const data = JSON.parse(event.data);
        if (!data.id || !this.pending.has(data.id)) return;
        const item = this.pending.get(data.id);
        clearTimeout(item.timer);
        this.pending.delete(data.id);
        if (data.error) item.reject(new Error(data.error.message));
        else item.resolve(data.result);
      });
    });
  }

  send(method, params = {}, timeoutMs = 20000) {
    const id = this.nextId++;
    this.socket.send(JSON.stringify({ id, method, params }));
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        if (!this.pending.has(id)) return;
        this.pending.delete(id);
        reject(new Error(`${method} timeout`));
      }, timeoutMs);
      this.pending.set(id, { resolve, reject, timer });
    });
  }

  close() {
    try {
      this.socket.close();
    } catch {}
  }
}

async function waitForExecutionContext(cdp, timeoutMs = 15000) {
  const started = Date.now();
  let lastError = null;
  while (Date.now() - started < timeoutMs) {
    try {
      const result = await cdp.send("Runtime.evaluate", {
        expression: "document.readyState",
        returnByValue: true
      }, 3000);
      if (!result.exceptionDetails && result.result?.value) return;
    } catch (error) {
      lastError = error;
    }
    await sleep(250);
  }
  throw new Error(`Edge page execution context not ready: ${lastError?.message || "timeout"}`);
}

async function evaluate(cdp, expression, timeoutMs = 20000) {
  const result = await cdp.send("Runtime.evaluate", {
    expression,
    awaitPromise: true,
    returnByValue: true
  }, timeoutMs);
  if (result.exceptionDetails) throw new Error(result.exceptionDetails.text || "Page evaluation failed.");
  return result.result.value;
}

async function createCdp(downloadPath) {
  const cdp = new CdpClient(await getPageSocketUrl());
  await cdp.open();
  await cdp.send("Page.enable");
  await cdp.send("Runtime.enable");
  await cdp.send("Browser.setDownloadBehavior", {
    behavior: "allow",
    downloadPath,
    eventsEnabled: true
  });
  await waitForExecutionContext(cdp);
  return cdp;
}

async function showBrowserWindow(cdp) {
  try {
    const { windowId } = await cdp.send("Browser.getWindowForTarget", {});
    await cdp.send("Browser.setWindowBounds", {
      windowId,
      bounds: { windowState: "normal", left: 80, top: 80, width: 1440, height: 1000 }
    });
  } catch {}
}

async function runActionWithRetry(cdpRef, downloadPath, expression, timeoutMs = 30000) {
  for (let attempt = 1; attempt <= 5; attempt += 1) {
    try {
      return await evaluate(cdpRef.current, expression, timeoutMs);
    } catch (error) {
      const transient = /Execution context was destroyed|Cannot find context|Cannot find default execution context|navigated|closed|detached|target/i.test(error.message);
      if (!transient || attempt === 5) throw error;
      try {
        cdpRef.current.close();
      } catch {}
      await sleep(Math.min(1000 + attempt * 500, 2500));
      cdpRef.current = await createCdp(downloadPath);
    }
  }
  throw new Error("Action retry failed.");
}

function listFiles(dir) {
  if (!fs.existsSync(dir)) return new Map();
  return new Map(fs.readdirSync(dir).map((name) => {
    const full = path.join(dir, name);
    const stat = fs.statSync(full);
    return [name, { full, mtimeMs: stat.mtimeMs, size: stat.size }];
  }));
}

function findDownloadedFile(dir, before) {
  const current = listFiles(dir);
  const files = Array.from(current.values()).filter((item) => {
    const name = path.basename(item.full);
    return !name.endsWith(".crdownload")
      && !name.endsWith(".tmp")
      && (!before.has(name) || item.mtimeMs > before.get(name).mtimeMs || item.size !== before.get(name).size)
      && item.size > 0;
  });
  files.sort((a, b) => b.mtimeMs - a.mtimeMs);
  return files[0]?.full || null;
}

async function waitForDownloadedFile(dir, before, timeoutMs = 25000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    const file = findDownloadedFile(dir, before);
    if (file) {
      const firstSize = fs.statSync(file).size;
      await sleep(500);
      if (fs.existsSync(file) && fs.statSync(file).size === firstSize) return file;
    }
    await sleep(300);
  }
  return null;
}

function actionScriptV2(month, options = {}) {
  const monthName = TEXT.months[month.month - 1];
  const loginWaitMs = Number(options.loginWaitMs || LOGIN_WAIT_MS);
  const credentials = options.credentials || {};
  return `
    (async () => {
      const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
      const visible = (el) => {
        if (!el) return false;
        const r = el.getBoundingClientRect();
        const s = getComputedStyle(el);
        return r.width > 0 && r.height > 0 && r.x > -50 && r.y > -50 && s.visibility !== "hidden" && s.display !== "none";
      };
      const norm = (text) => String(text || "").replace(/\\s+/g, "").trim();
      const bodyText = () => document.body.innerText || "";
      const pageHas = (...items) => items.every((item) => bodyText().includes(item));
      const clickElement = async (el) => {
        if (!el) return false;
        el.scrollIntoView({ block: "center", inline: "center" });
        await sleep(120);
        const r = el.getBoundingClientRect();
        const x = Math.min(Math.max(r.left + r.width / 2, 1), window.innerWidth - 1);
        const y = Math.min(Math.max(r.top + r.height / 2, 1), window.innerHeight - 1);
        const hit = document.elementFromPoint(x, y);
        const target = hit?.closest("button,a,[role='button']") || el.closest("button,a,[role='button'],li") || hit || el;
        ["pointerdown", "mousedown", "pointerup", "mouseup", "click"].forEach((type) => {
          target.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window, clientX: x, clientY: y }));
        });
        await sleep(350);
        return true;
      };
      const waitFor = async (test, timeout = 8000, interval = 180) => {
        const started = Date.now();
        while (Date.now() - started < timeout) {
          const value = test();
          if (value) return value;
          await sleep(interval);
        }
        return null;
      };
      const clickText = async (text, exact = false) => {
        const expected = norm(text);
        const targets = [...document.querySelectorAll("button,a,[role='button'],li,span,div")]
          .filter(visible)
          .filter((el) => {
            const current = norm(el.innerText || el.title || el.textContent);
            return exact ? current === expected : current.includes(expected);
          });
        targets.sort((a, b) => {
          const ar = a.getBoundingClientRect();
          const br = b.getBoundingClientRect();
          const aclickable = a.matches("button,a,[role='button']") ? 0 : 1;
          const bclickable = b.matches("button,a,[role='button']") ? 0 : 1;
          return aclickable - bclickable || (ar.width * ar.height) - (br.width * br.height);
        });
        const target = targets[0];
        if (!target) return false;
        return await clickElement(target);
      };
      const setNativeValue = (el, value) => {
        const proto = Object.getPrototypeOf(el);
        const setter = Object.getOwnPropertyDescriptor(proto, "value")?.set;
        setter ? setter.call(el, value) : el.value = value;
        el.dispatchEvent(new Event("input", { bubbles: true }));
        el.dispatchEvent(new Event("change", { bubbles: true }));
      };
      const fillLoginForm = async () => {
        const username = ${JSON.stringify(credentials.username || "")};
        const password = ${JSON.stringify(credentials.password || "")};
        if (!username || !password) return false;
        const inputs = [...document.querySelectorAll("input")].filter(visible);
        const userInput = inputs.find((el) => /text|email|tel|search|^$/i.test(el.type || "") && !/password/i.test(el.type || ""));
        const passInput = inputs.find((el) => /password/i.test(el.type || ""));
        if (!userInput || !passInput) return false;
        userInput.focus();
        setNativeValue(userInput, username);
        passInput.focus();
        setNativeValue(passInput, password);
        await sleep(200);
        return await clickText("\\u767b\\u5f55") || await clickText("\\u786e\\u5b9a") || await clickText("\\u63d0\\u4ea4");
      };
      const waitForIdle = async () => {
        await sleep(500);
        await waitFor(() => {
          const loading = [...document.querySelectorAll(".ant-spin-spinning, .ant-btn-loading, [aria-busy='true']")]
            .some(visible);
          return loading ? null : true;
        }, 15000);
        await sleep(300);
        return true;
      };
      const clickLoginIfNeeded = async () => {
        if (!pageHas("\\u767b\\u5f55")) return true;
        await clickText("\\u4f01\\u4e1a\\u5fae\\u4fe1\\u767b\\u5f55") || await clickText("\\u767b\\u5f55");
        await sleep(1000);
        const fillStarted = Date.now();
        while (Date.now() - fillStarted < 20000 && pageHas("\\u767b\\u5f55")) {
          if (await fillLoginForm()) break;
          await sleep(1000);
        }
        return Boolean(await waitFor(() => !pageHas("\\u767b\\u5f55") || pageHas("\\u4e2d\\u95f4\\u4ef6\\u4ea7\\u54c1\\u5217\\u8868") || pageHas(${JSON.stringify(TEXT.basic)}), ${loginWaitMs}));
      };
      const clickBasicServiceCard = async () => {
        const enters = [...document.querySelectorAll("button,a,[role='button'],span,div")]
          .filter(visible)
          .filter((el) => norm(el.innerText || el.title || el.textContent).includes(norm("\\u70b9\\u51fb\\u8fdb\\u5165")));
        const matches = [];
        for (const enter of enters) {
          let node = enter;
          for (let depth = 0; node && depth < 8; depth += 1, node = node.parentElement) {
            const text = norm(node.innerText || node.textContent);
            if (text.includes(norm(${JSON.stringify(TEXT.basic)})) && text.includes(norm("\\u70b9\\u51fb\\u8fdb\\u5165"))) {
              const r = node.getBoundingClientRect();
              matches.push({ enter, card: node, area: r.width * r.height });
              break;
            }
          }
        }
        matches.sort((a, b) => a.area - b.area);
        const match = matches[0];
        if (!match) return false;
        await clickElement(match.enter);
        await waitForIdle();
        return true;
      };
      const onMonthlyPage = () => pageHas(${JSON.stringify(TEXT.query)}, ${JSON.stringify(TEXT.export)}, "\\u8d26\\u671f", "\\u4f9b\\u5e94\\u5546");
      const ensureMonthlyPage = async () => {
        if (onMonthlyPage()) return true;
        if (!location.href.includes("/itservice/sms/sms_bill")) {
          location.href = ${JSON.stringify(SMS_BILL_URL)};
          await waitFor(() => onMonthlyPage() || pageHas("\\u4e2d\\u95f4\\u4ef6\\u4ea7\\u54c1\\u5217\\u8868") || pageHas("\\u70b9\\u51fb\\u8fdb\\u5165"), 15000);
          await waitForIdle();
          if (onMonthlyPage()) return true;
        }
        if (pageHas("\\u4e2d\\u95f4\\u4ef6\\u4ea7\\u54c1\\u5217\\u8868") || pageHas("\\u70b9\\u51fb\\u8fdb\\u5165")) {
          if (!await clickBasicServiceCard()) return false;
          await waitFor(() => pageHas(${JSON.stringify(TEXT.basic)}), 15000);
        }
        if (!onMonthlyPage()) {
          await clickText(${JSON.stringify(TEXT.monthly)}, true) || await clickText(${JSON.stringify(TEXT.monthly)});
        }
        return Boolean(await waitFor(onMonthlyPage, 15000));
      };
      const ensureBusinessTab = async () => {
        const ready = () => {
          const text = bodyText();
          return text.includes(${JSON.stringify(TEXT.business)})
            && text.includes("\\u8d26\\u671f")
            && text.includes(${JSON.stringify(TEXT.query)})
            && text.includes(${JSON.stringify(TEXT.export)});
        };
        // The tab is rendered differently across portal versions (Ant tabs,
        // buttons, or links).  Restrict the search to likely tab controls so
        // a table cell/label named “业务方” is not clicked accidentally.
        const expected = norm(${JSON.stringify(TEXT.business)});
        const tabs = [...document.querySelectorAll("[role='tab'], .ant-tabs-tab, .ant-tabs-tab-btn, button, a")]
          .filter(visible)
          .filter((el) => norm(el.innerText || el.textContent || el.title) === expected);
        const activeTab = tabs.find((el) => el.getAttribute("aria-selected") === "true"
          || el.classList.contains("ant-tabs-tab-active")
          || el.parentElement?.classList.contains("ant-tabs-tab-active"));
        if (activeTab && ready()) return true;
        if (tabs[0]) {
          // Native click is the most reliable path for Ant Design's tab
          // button, and avoids accidentally matching a similarly named cell.
          try { tabs[0].click(); } catch { await clickElement(tabs[0]); }
          await sleep(350);
        }
        else await clickText(${JSON.stringify(TEXT.business)}, true) || await clickText(${JSON.stringify(TEXT.business)});
        await waitForIdle();
        return Boolean(await waitFor(ready, 15000));
      };
      const setMonth = async () => {
        const input = await waitFor(() => {
          const inputs = [...document.querySelectorAll("input")]
            .filter((el) => visible(el))
            .filter((el) => {
              const text = [el.placeholder, el.id, el.name, el.closest(".ant-form-item")?.innerText, el.parentElement?.innerText].join(" ");
              return /Select month|\\u8d26\\u671f|\\u6708\\u4efd|month/i.test(text);
            });
          return inputs[0];
        }, 6000);
        if (!input) return { ok: false, reason: "month input not found" };
        input.click();
        await sleep(200);

        for (let i = 0; i < 10; i += 1) {
          const yearEl = document.querySelector(".ant-calendar-month-panel-year-select-content");
          const currentYear = Number(yearEl?.innerText?.replace(/\\D/g, ""));
          if (!currentYear || currentYear === ${month.year}) break;
          const selector = currentYear > ${month.year} ? ".ant-calendar-month-panel-prev-year-btn" : ".ant-calendar-month-panel-next-year-btn";
          const yearButton = document.querySelector(selector);
          if (!yearButton) return { ok: false, reason: "year switch button not found" };
          yearButton.click();
          await sleep(120);
        }

        const monthCell = await waitFor(() => [...document.querySelectorAll(".ant-calendar-month-panel-cell")]
          .find((el) => visible(el) && (el.getAttribute("title") === ${JSON.stringify(monthName)} || norm(el.innerText) === norm(${JSON.stringify(monthName)}))), 5000);
        if (!monthCell) return { ok: false, reason: "month cell not found" };
        monthCell.click();
        const selected = await waitFor(() => input.value === ${JSON.stringify(month.value)} ? input.value : null, 3000);
        return selected ? { ok: true, value: selected } : { ok: false, reason: "month was not selected", value: input.value };
      };
      const getRows = () => [...document.querySelectorAll("tbody tr, .ant-table-row")]
        .filter(visible)
        .filter((row) => {
          const text = norm(row.innerText);
          const cellCount = row.querySelectorAll("td").length;
          return cellCount >= 5 && text && !${JSON.stringify(TEXT.noData)}.some((item) => text.includes(norm(item)));
        });

      if (!await clickLoginIfNeeded()) return { ok: false, reason: ${JSON.stringify(credentials.hasCredentials ? "海星自动登录未完成，可能需要验证码/短信/二次确认，请在打开的 Edge 窗口完成登录后重新导出。" : "海星登录未完成，且未配置账号密码，请在打开的 Edge 窗口完成登录后重新导出。")}, body: bodyText().slice(0, 1000) };
      if (!await ensureMonthlyPage()) return { ok: false, reason: "未能进入 基础服务 -> 月度对账。", body: bodyText().slice(0, 1000) };
      if (!await ensureBusinessTab()) return { ok: false, reason: "未能切换到 月度对账 -> 业务方。", body: bodyText().slice(0, 1000) };

      const monthResult = await setMonth();
      if (!monthResult.ok) return monthResult;

      if (!await clickText(${JSON.stringify(TEXT.query)})) return { ok: false, reason: "query button not found" };
      await waitForIdle();
      const status = await waitFor(() => {
        const rows = getRows();
        if (rows.length) return { hasBill: true, rows: rows.length };
        const noData = [...document.querySelectorAll(".ant-empty, .ant-table-placeholder, .ant-empty-description, td, div")]
          .filter(visible)
          .some((el) => ${JSON.stringify(TEXT.noData)}.some((item) => norm(el.innerText).includes(norm(item))));
        return noData ? { hasBill: false, rows: 0 } : null;
      }, 12000);
      if (!status || !status.hasBill) return { ok: true, hasBill: false, reason: "no bill", selectedMonth: monthResult.value, body: bodyText().slice(0, 1000) };

      if (!await clickText(${JSON.stringify(TEXT.export)})) return { ok: false, reason: "export button not found" };
      return { ok: true, hasBill: true, queried: true, rows: status.rows || 0, selectedMonth: monthResult.value };
    })()
  `;
}

async function main() {
  const month = parseMonth(monthArg);
  if (!month) throw new Error("Invalid month. Use YYYY-MM.");

  const targetDir = path.resolve(DATA_ROOT, month.folderName);
  fs.mkdirSync(targetDir, { recursive: true });
  const before = listFiles(targetDir);
  const credentials = getSeastarCredentials();

  await startFreshEdge();
  const cdpRef = { current: await createCdp(targetDir) };

  try {
    if (SHOW_BROWSER) await showBrowserWindow(cdpRef.current);
    await sleep(1200);

    const result = await runActionWithRetry(
      cdpRef,
      targetDir,
      actionScriptV2(month, { loginWaitMs: LOGIN_WAIT_MS, credentials }),
      LOGIN_WAIT_MS + 30000
    );
    if (!result.ok || result.hasBill === false) {
      console.log(JSON.stringify({ ...result, folder: targetDir }, null, 2));
      return;
    }

    const filePath = await waitForDownloadedFile(targetDir, before);
    console.log(JSON.stringify({
      ok: Boolean(filePath),
      hasBill: true,
      folder: targetDir,
      filePath,
      reason: filePath ? "exported" : "download file not detected"
    }, null, 2));
  } finally {
    cdpRef.current.close();
    stopManagedEdge();
  }
}

main().catch((error) => {
  console.log(JSON.stringify({ ok: false, reason: error.message }, null, 2));
  process.exitCode = 1;
});
