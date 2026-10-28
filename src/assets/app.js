const now = new Date();
const defaultMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
const apiBase = window.location.protocol === "file:" ? "http://127.0.0.1:8099" : "";
const latestDetailFixedUrl = "https://taiyi.yonghui.cn/s/it-service-sms-alert/v1/report/smsData/";

const pageTitle = document.querySelector("#page-title");
const navLinks = document.querySelectorAll("[data-view-link]");
const navToggles = document.querySelectorAll("[data-nav-toggle]");
const views = document.querySelectorAll("[data-view]");

const billMonth = document.querySelector("#bill-month");
const organizeMonth = document.querySelector("#organize-month");
const hmtMonth = document.querySelector("#hmt-month");
const aiShortMonth = document.querySelector("#ai-short-month");
const oneClickMonth = document.querySelector("#one-click-month");
const expenseMonth = document.querySelector("#expense-month");
const combinedWorkbookMonth = document.querySelector("#combined-workbook-month");
const feishuThirdPartyMonth = document.querySelector("#feishu-third-party-month");
const feishuInternalMonth = document.querySelector("#feishu-internal-month");
const feishuInternalCutoff = document.querySelector("#feishu-internal-cutoff");
const filterBillsMonth = document.querySelector("#filter-bills-month");
const invoiceJmMonth = document.querySelector("#invoice-jm-month");
const invoiceBeijingMonth = document.querySelector("#invoice-beijing-month");
const invoiceJingdongMonth = document.querySelector("#invoice-jingdong-month");
const invoiceTengyuMonth = document.querySelector("#invoice-tengyu-month");

const hmtUrl = document.querySelector("#hmt-url");
const latestDetailUrl = document.querySelector("#latest-detail-url");
const fillLatestDetailUrlButton = document.querySelector("#fill-latest-detail-url");
const aiShortFile = document.querySelector("#ai-short-file");
const oneClickHmtUrl = document.querySelector("#one-click-hmt-url");
const oneClickAiShortFile = document.querySelector("#one-click-ai-short-file");
const expenseYhFile = document.querySelector("#expense-yh-file");
const expenseJmFile = document.querySelector("#expense-jm-file");
const expenseJmAiFile = document.querySelector("#expense-jm-ai-file");
const expenseBeijingFile = document.querySelector("#expense-beijing-file");
const expenseJdFile = document.querySelector("#expense-jd-file");
const expenseTengyuFile = document.querySelector("#expense-tengyu-file");
const invoiceJmArchive = document.querySelector("#invoice-jm-archive");

const exportButton = document.querySelector("#export-summary-bill");
const organizeButton = document.querySelector("#organize-summary-bill");
const hmtButton = document.querySelector("#import-hmt-bill");
const exportLatestDetailButton = document.querySelector("#export-latest-detail");
const appendHmtButton = document.querySelector("#append-hmt-bill");
const aiShortButton = document.querySelector("#import-ai-short-bill");
const appendAiShortButton = document.querySelector("#append-ai-short-bill");
const oneClickButton = document.querySelector("#generate-one-click-bill");
const expenseCompareButton = document.querySelector("#generate-expense-compare");
const combinedWorkbookButton = document.querySelector("#combine-workbooks");
const feishuThirdPartyButton = document.querySelector("#generate-feishu-third-party-script");
const feishuInternalButton = document.querySelector("#generate-feishu-internal");
const checkInternalAccountsButton = document.querySelector("#check-internal-accounts");
const addInternalGroupButton = document.querySelector("#add-internal-group");
const saveInternalGroupsButton = document.querySelector("#save-internal-groups");
const jiangxiJumengButton = document.querySelector("#filter-jiangxi-jumeng");
const beijingLingduButton = document.querySelector("#filter-beijing-lingdu");
const jingdongButton = document.querySelector("#filter-jingdong");
const tengyuButton = document.querySelector("#filter-tengyu");











const techBillButton = document.querySelector("#filter-tech-bill");
const archiveInvoiceJmButton = document.querySelector("#archive-invoice-jm");
const organizeInvoiceJmButton = document.querySelector("#organize-invoice-jm");
const invoiceMergeMonth = document.querySelector("#invoice-merge-month");
const mergeInvoicePdfsByAccountButton = document.querySelector("#merge-invoice-pdfs-by-account");
const invoiceMergeResult = document.querySelector("#invoice-merge-result");

const exportResult = document.querySelector("#export-result");
const organizeResult = document.querySelector("#organize-result");
const hmtResult = document.querySelector("#hmt-result");
const latestDetailResult = document.querySelector("#latest-detail-result");
const aiShortResult = document.querySelector("#ai-short-result");
const oneClickResult = document.querySelector("#one-click-result");
const expenseCompareResult = document.querySelector("#expense-compare-result");
const combinedWorkbookResult = document.querySelector("#combined-workbook-result");
const feishuThirdPartyResult = document.querySelector("#feishu-third-party-result");
const jiangxiJumengResult = document.querySelector("#jiangxi-jumeng-result");
const beijingLingduResult = document.querySelector("#beijing-lingdu-result");
const jingdongResult = document.querySelector("#jingdong-result");
const tengyuResult = document.querySelector("#tengyu-result");











const techBillResult = document.querySelector("#tech-bill-result");
const oneClickReport = document.querySelector("#one-click-report");
const expenseCompareReport = document.querySelector("#expense-compare-report");
const combinedWorkbookReport = document.querySelector("#combined-workbook-report");
const feishuThirdPartyReport = document.querySelector("#feishu-third-party-report");
const internalAccountCheckResult = document.querySelector("#internal-account-check-result");
const feishuInternalResult = document.querySelector("#feishu-internal-result");
const feishuInternalReport = document.querySelector("#feishu-internal-report");
const internalGroupsEditor = document.querySelector("#internal-groups-editor");
const internalGroupsResult = document.querySelector("#internal-groups-result");
const invoiceJmResult = document.querySelector("#invoice-jm-result");
const invoiceJmReport = document.querySelector("#invoice-jm-report");
const invoiceJmSummary = document.querySelector("#invoice-jm-summary");
const invoiceJmReportBody = document.querySelector("#invoice-jm-report-body");
let internalGroups = [];
let invoiceJmSourceFolder = "";
const organizeReport = document.querySelector("#organize-report");
const organizeReportBody = document.querySelector("#organize-report-body");

const fileInputLabels = {
  "ai-short-file": "AI语音挂短账单文件",
  "one-click-ai-short-file": "AI语音挂短账单文件",
  "expense-yh-file": "永辉整理账单",
  "expense-jm-file": "江西聚梦短信/港澳台账单",
  "expense-jm-ai-file": "江西聚梦AI语音挂短账单",
  "expense-beijing-file": "北京凌渡账单",
  "expense-jd-file": "京东账单",
  "invoice-jm-archive": "供应商发票 ZIP 压缩包",
  "invoice-beijing-archive": "北京凌渡账单发票 ZIP 压缩包",
  "invoice-jingdong-archive": "京东账单发票 ZIP 压缩包" ,
  "invoice-tengyu-archive": "腾域账单发票 ZIP 压缩包" ,
};

[billMonth, organizeMonth, hmtMonth, aiShortMonth, oneClickMonth, expenseMonth, combinedWorkbookMonth, feishuThirdPartyMonth, filterBillsMonth, invoiceJmMonth, invoiceBeijingMonth, invoiceJingdongMonth, invoiceTengyuMonth, invoiceMergeMonth].forEach((input) => {
  if (input) input.value = defaultMonth;
});
if (feishuInternalMonth) feishuInternalMonth.value = defaultMonth;
if (feishuInternalCutoff) {
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() + 3);
  cutoff.setHours(18, 0, 0, 0);
  feishuInternalCutoff.value = `${cutoff.getFullYear()}-${String(cutoff.getMonth() + 1).padStart(2, "0")}-${String(cutoff.getDate()).padStart(2, "0")}T18:00`;
}
setResult(aiShortResult, "请选择 AI挂短 Excel 账单后导入。");

function fileNameText(input) {
  const files = Array.from(input.files || []);
  if (!files.length) return "点击选择文件，或拖动文件到这里";
  if (files.length === 1) return files[0].name;
  return `${files.length} 个文件已选择`;
}

function setDropFile(input, files) {
  if (!files || !files.length) return;
  const transfer = new DataTransfer();
  transfer.items.add(files[0]);
  input.files = transfer.files;
  input.dispatchEvent(new Event("change", { bubbles: true }));
}

function enhanceFileInputs() {
  document.querySelectorAll("input[type='file'].file-input").forEach((input) => {
    if (input.dataset.dragReady === "true") return;
    input.dataset.dragReady = "true";

    const parent = input.parentElement;
    const parentTitle = parent?.querySelector(":scope > span")?.textContent?.trim();
    const label = fileInputLabels[input.id] || parentTitle || "账单文件";
    const drop = document.createElement(parent?.tagName === "LABEL" ? "span" : "label");
    drop.className = "drag-upload";
    if (drop.tagName === "LABEL") drop.htmlFor = input.id;
    drop.tabIndex = 0;

    const icon = document.createElement("span");
    icon.className = "drag-upload-icon";
    icon.textContent = "+";

    const title = document.createElement("span");
    title.className = "drag-upload-title";
    title.textContent = label;

    const hint = document.createElement("span");
    hint.className = "drag-upload-hint";
    hint.textContent = fileNameText(input);

    if (parent?.classList.contains("file-field")) {
      const oldTitle = parent.querySelector(":scope > span");
      if (oldTitle) oldTitle.remove();
    } else {
      const oldLabel = document.querySelector(`label[for="${input.id}"]`);
      if (oldLabel) oldLabel.remove();
    }

    input.classList.add("visually-hidden-file");
    input.insertAdjacentElement("beforebegin", drop);
    drop.append(icon, title, hint, input);

    input.addEventListener("change", () => {
      hint.textContent = fileNameText(input);
      drop.classList.toggle("has-file", Boolean(input.files && input.files.length));
    });

    ["dragenter", "dragover"].forEach((eventName) => {
      drop.addEventListener(eventName, (event) => {
        event.preventDefault();
        drop.classList.add("drag-over");
      });
    });

    ["dragleave", "drop"].forEach((eventName) => {
      drop.addEventListener(eventName, (event) => {
        event.preventDefault();
        drop.classList.remove("drag-over");
      });
    });

    drop.addEventListener("drop", (event) => {
      setDropFile(input, event.dataTransfer.files);
    });

    drop.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        input.click();
      }
    });
  });
}

enhanceFileInputs();

function setResult(target, message, type = "info") {
  if (!target) return;
  target.textContent = message;
  target.dataset.type = type;
}

function showStoppedResult(result, target) {
  if (!result?.stopped) return false;
  setResult(target, result.message || "\u5df2\u505c\u6b62\u751f\u6210\u3002", "info");
  return true;
}

function switchView(name) {
  navLinks.forEach((link) => link.classList.toggle("active", link.dataset.viewLink === name));
  views.forEach((view) => view.classList.toggle("active", view.dataset.view === name));

  const activeGroup = document.querySelector("[data-view-link].active")?.closest("[data-nav-group]");
  activeGroup?.classList.add("open");
  activeGroup?.querySelector("[data-nav-toggle]")?.setAttribute("aria-expanded", "true");

  const titles = {
    combinedWorkbook: "账单合并",
    feishuCommunicationThirdParty: "飞书沟通（三方群）",
    feishuInternalReconciliation: "飞书推送（内部对账）",
    invoiceReconciliation: "发票核对整理",
    filterBills: "筛选账单",
    latestDetailExport: "最新明细导出",
    export: "导出汇总账单",
    organize: "汇总账单整理",
    hmt: "港澳台账单导入",
    aiShort: "AI挂短账单导入",
    oneClick: "汇总账单一键生成",
    expenseCompare: "费用对比生成"
  };
  pageTitle.textContent = titles[name] || titles.oneClick;
}

function setButton(button, disabled, icon, label) {
  button.disabled = disabled;
  button.innerHTML = `<span class="button-icon" aria-hidden="true">${icon}</span><span>${label}</span>`;
  if (!disabled && button._stopButton) {
    endGenerationTask(button);
  }
}

function createTaskId(prefix) {
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}

function stopButtonFor(button) {
  if (!button) return null;
  if (button._stopButton) return button._stopButton;

  const stopButton = document.createElement("button");
  stopButton.type = "button";
  stopButton.className = "primary-export stop-generation-button";
  stopButton.hidden = true;
  stopButton.innerHTML = `<span class="button-icon" aria-hidden="true">×</span><span>${"\u505c\u6b62\u751f\u6210"}</span>`;
  button.insertAdjacentElement("afterend", stopButton);
  button._stopButton = stopButton;
  return stopButton;
}

function beginGenerationTask(button, resultTarget, prefix) {
  const taskId = createTaskId(prefix);
  const stopButton = stopButtonFor(button);
  if (!stopButton) return taskId;

  stopButton.hidden = false;
  stopButton.disabled = false;
  stopButton.dataset.taskId = taskId;
  stopButton.onclick = async () => {
    stopButton.disabled = true;
    setResult(resultTarget, "\u6b63\u5728\u505c\u6b62\u751f\u6210...");
    try {
      await postJson("/api/stop-task", { taskId });
      setResult(resultTarget, "\u5df2\u53d1\u9001\u505c\u6b62\u6307\u4ee4\uff0c\u8bf7\u7a0d\u5019\u3002");
    } catch (error) {
      setResult(resultTarget, `\u505c\u6b62\u751f\u6210\u5931\u8d25\uff1a${error.message}`, "error");
      stopButton.disabled = false;
    }
  };
  return taskId;
}

function endGenerationTask(button) {
  const stopButton = stopButtonFor(button);
  if (!stopButton) return;
  stopButton.hidden = true;
  stopButton.disabled = false;
  stopButton.onclick = null;
  delete stopButton.dataset.taskId;
}

navToggles.forEach((toggle) => {
  toggle.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    const group = toggle.closest("[data-nav-group]");
    const open = !group.classList.contains("open");
    group.classList.toggle("open", open);
    toggle.setAttribute("aria-expanded", String(open));
  });
});

navLinks.forEach((link) => {
  link.addEventListener("click", (event) => {
    event.preventDefault();
    switchView(link.dataset.viewLink);
  });
});

function renderOneClickReport(result) {
  const steps = result.steps || [];
  oneClickReport.innerHTML = `
    <div class="step-list">
      ${steps.map((step) => `
        <div class="step-row" data-status="${step.skipped ? "skipped" : step.ok ? "ok" : "error"}">
          <strong>${step.name}</strong>
          <span>${step.message || "-"}</span>
        </div>
      `).join("")}
    </div>
    ${result.outputFile ? `<p class="output-path">整理账单：${result.outputFile}</p>` : ""}
  `;
}

function renderExpenseCompareReport(result) {
  const totals = result.totals || {};
  const rows = [
    ["江西聚梦 YH短信", totals.yhJmSms],
    ["江西聚梦账单短信", totals.jmSms],
    ["江西聚梦TTS", totals.jmTts],
    ["江西聚梦港澳台", totals.jmHmt],
    ["AI挂短", totals.jmShort],
    ["AI语音", totals.jmVoiceAi],
    ["京东 YH/账单", `${totals.yhJd ?? "-"} / ${totals.jd ?? "-"}`],
    ["北京凌渡 YH/账单", `${totals.yhBeijing ?? "-"} / ${totals.beijing ?? "-"}`]
  ];
  expenseCompareReport.innerHTML = `
    <div class="step-list">
      ${rows.map(([name, value]) => `
        <div class="step-row" data-status="ok">
          <strong>${name}</strong>
          <span>${value ?? "-"}</span>
        </div>
      `).join("")}
    </div>
    ${result.outputFile ? `<p class="output-path">费用对比表：${result.outputFile}</p>` : ""}
  `;
}

function renderScriptReport(target, messages) {
  if (!target) return;
  const rows = [
    ["京东", messages?.jingdong],
    ["江西聚梦", messages?.jiangxiJumeng],
    ["北京凌渡", messages?.beijingLingdu],
    ["腾域", messages?.tengyu]
  ].filter(([, text]) => text);

  target.innerHTML = `
    <div class="script-list">
      ${rows.map(([name, text]) => `
        <div class="script-item">
          <div class="script-item-head">
            <strong>${name}</strong>
            <button class="copy-script-button" type="button" data-copy-script="${name}">复制</button>
          </div>
          <pre>${text}</pre>
        </div>
      `).join("")}
    </div>
  `;

  target.querySelectorAll("[data-copy-script]").forEach((button) => {
    button.addEventListener("click", async () => {
      const item = button.closest(".script-item");
      const text = item?.querySelector("pre")?.textContent || "";
      try {
        await navigator.clipboard.writeText(text);
        button.textContent = "已复制";
        setTimeout(() => {
          button.textContent = "复制";
        }, 1200);
      } catch {
        button.textContent = "复制失败";
      }
    });
  });
}

function renderIssues(result) {
  organizeReport.classList.remove("hidden");
  const issues = result.issues || [];
  const notes = result.notes || [];
  const issueItems = issues.length ? issues.map((item) => `<li>${item}</li>`).join("") : "<li>未发现明显异常。</li>";
  const noteItems = notes.length ? notes.map((item) => `<li>${item}</li>`).join("") : "<li>已完成整理。</li>";

  organizeReportBody.innerHTML = `
    <p><strong>源文件：</strong>${result.sourceFile || "-"}</p>
    <p><strong>整理文件：</strong>${result.outputFile || "-"}</p>
    <p><strong>合计整理：</strong>${result.mergedGroups || 0} 组连续相同合计已合并居中。</p>
    <h3>异常检查</h3>
    <ul>${issueItems}</ul>
    <h3>说明</h3>
    <ul>${noteItems}</ul>
  `;
}

async function readJsonResponse(response) {
  const body = await response.text();
  try {
    return body ? JSON.parse(body) : {};
  } catch {
    return { message: body || `服务返回异常：HTTP ${response.status}` };
  }
}

async function postJson(url, payload) {
  let response;
  try {
    response = await fetch(`${apiBase}${url}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });
  } catch {
    throw new Error("无法连接本地服务，请确认系统已启动并使用 http://127.0.0.1:8099/ 打开。");
  }
  const result = await readJsonResponse(response);
  if (!response.ok) throw new Error(result.message || "请求失败。");
  return result;
}

async function postForm(url, formData) {
  let response;
  try {
    response = await fetch(`${apiBase}${url}`, {
      method: "POST",
      body: formData
    });
  } catch {
    throw new Error("无法连接本地服务，请确认系统已启动并使用 http://127.0.0.1:8099/ 打开。");
  }
  const result = await readJsonResponse(response);
  if (!response.ok) throw new Error(result.message || "请求失败。");
  return result;
}

function addOpenFileButton(target, filePath) {
  if (!target || !filePath) return;

  const actions = document.createElement("div");
  actions.className = "open-file-actions";

  const button = document.createElement("button");
  button.type = "button";
  button.className = "primary-export open-file-button";
  button.innerHTML = `<span class="button-icon" aria-hidden="true">↗</span><span>${"\u6253\u5f00\u751f\u6210\u6587\u4ef6"}</span>`;

  button.addEventListener("click", async () => {
    button.disabled = true;
    const originalContent = button.innerHTML;
    button.innerHTML = `<span class="button-icon" aria-hidden="true">…</span><span>正在打开</span>`;
    try {
      await postJson("/api/open-generated-file", { filePath });
      button.innerHTML = `<span class="button-icon" aria-hidden="true">✓</span><span>已打开</span>`;
      await new Promise((resolve) => window.setTimeout(resolve, 1200));
    } catch (error) {
      setResult(target, `${"\u6253\u5f00\u6587\u4ef6\u5931\u8d25\uff1a"}${error.message}`, "error");
    } finally {
      button.innerHTML = originalContent;
      button.disabled = false;
    }
  });

  actions.appendChild(button);
  target.appendChild(actions);
}

function escapeTableText(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function renderInvoiceJmReport(result) {
  if (!invoiceJmReport || !invoiceJmReportBody || !invoiceJmSummary) return;
  const rows = result.results || [];
  invoiceJmReport.classList.remove("hidden");
  invoiceJmSummary.textContent = `成功 ${result.renamedCount || 0} 个，需检查 ${result.warningCount || 0} 个`;
  invoiceJmReportBody.innerHTML = rows.map((item) => {
    const detail = item.ok
      ? `${escapeTableText(item.newName || "-")}${item.warning ? `<small>${escapeTableText(item.warning)}</small>` : ""}`
      : `<span class="invoice-warning">${escapeTableText(item.warning || "未处理")}</span>`;
    return `
      <tr data-status="${item.ok && !item.warning ? "ok" : "warning"}">
        <td><span class="invoice-type">${escapeTableText(item.category)}</span></td>
        <td class="invoice-filename">${escapeTableText(item.oldName)}</td>
        <td>${escapeTableText((item.accounts || []).join("、") || "-")}</td>
        <td>${escapeTableText(item.amount || "-")}</td>
        <td class="invoice-filename">${detail}</td>
      </tr>
    `;
  }).join("");
}

invoiceJmMonth?.addEventListener("change", () => {
  invoiceJmSourceFolder = "";
  invoiceJmReport?.classList.add("hidden");
  setResult(invoiceJmResult, "账期已切换，请上传对应月份的供应商压缩包。");
});

archiveInvoiceJmButton?.addEventListener("click", async () => {
  const archive = invoiceJmArchive?.files?.[0];
  if (!archive) {
    setResult(invoiceJmResult, "请选择供应商发票 ZIP 压缩包。", "error");
    return;
  }
  const formData = new FormData();
  formData.append("month", invoiceJmMonth.value);
  formData.append("archive", archive);
  setButton(archiveInvoiceJmButton, true, "…", "正在解压");
  invoiceJmReport?.classList.add("hidden");
  setResult(invoiceJmResult, "正在将原压缩包保存到账期目录并安全解压。");
  try {
    const result = await postForm("/api/archive-jiangxi-invoice-package", formData);
    invoiceJmSourceFolder = result.sourceFolder || "";
    setResult(
      invoiceJmResult,
      `归档完成：共发现 ${result.pdfCount || 0} 个 PDF。解压目录：${result.sourceFolder}`,
      "success"
    );
    addOpenFileButton(invoiceJmResult, result.sourceFolder);
  } catch (error) {
    setResult(invoiceJmResult, `归档解压失败：${error.message}`, "error");
  } finally {
    setButton(archiveInvoiceJmButton, false, "↓", "归档并解压");
  }
});

organizeInvoiceJmButton?.addEventListener("click", async () => {
  setButton(organizeInvoiceJmButton, true, "…", "正在整理");
  invoiceJmReport?.classList.add("hidden");
  setResult(invoiceJmResult, "正在读取账单与发票 PDF，识别账号编码和金额。");
  try {
    const result = await postJson("/api/organize-jiangxi-invoice-pdfs", {
      month: invoiceJmMonth.value,
      sourceFolder: invoiceJmSourceFolder
    });
    invoiceJmSourceFolder = result.sourceFolder || invoiceJmSourceFolder;
    setResult(
      invoiceJmResult,
      `整理完成：成功命名 ${result.renamedCount || 0} 个文件，${result.warningCount || 0} 个文件需要检查。`,
      result.warningCount ? "info" : "success"
    );
    addOpenFileButton(invoiceJmResult, result.sourceFolder);
    renderInvoiceJmReport(result);
  } catch (error) {
    setResult(invoiceJmResult, `PDF 文件名整理失败：${error.message}`, "error");
  } finally {
    setButton(organizeInvoiceJmButton, false, "✓", "整理 PDF 文件名");
  }
});

function setupInvoiceProvider(config) {
  const monthInput = document.querySelector(`#invoice-${config.id}-month`);
  const archiveInput = document.querySelector(`#invoice-${config.id}-archive`);
  const archiveButton = document.querySelector(`#archive-invoice-${config.id}`);
  const organizeButton = document.querySelector(`#organize-invoice-${config.id}`);
  const resultTarget = document.querySelector(`#invoice-${config.id}-result`);
  const report = document.querySelector(`#invoice-${config.id}-report`);
  const summary = document.querySelector(`#invoice-${config.id}-summary`);
  const reportBody = document.querySelector(`#invoice-${config.id}-report-body`);
  let sourceFolder = "";

  if (!monthInput || !archiveInput || !archiveButton || !organizeButton) return;

  const renderReport = (result) => {
    const rows = result.results || [];
    report.classList.remove("hidden");
    summary.textContent = `成功 ${result.renamedCount || 0} 个，需检查 ${result.warningCount || 0} 个`;
    reportBody.innerHTML = rows.map((item) => {
      const detail = item.ok
        ? `${escapeTableText(item.newName || "-")}${item.warning ? `<small>${escapeTableText(item.warning)}</small>` : ""}`
        : `<span class="invoice-warning">${escapeTableText(item.warning || "未处理")}</span>`;
      return `
        <tr data-status="${item.ok && !item.warning ? "ok" : "warning"}">
          <td><span class="invoice-type">${escapeTableText(item.category)}</span></td>
          <td class="invoice-filename">${escapeTableText(item.oldName)}</td>
          <td>${escapeTableText((item.accounts || []).join("、") || "-")}</td>
          <td>${escapeTableText(item.amount || "-")}</td>
          <td class="invoice-filename">${detail}</td>
        </tr>
      `;
    }).join("");
  };

  monthInput.addEventListener("change", () => {
    sourceFolder = "";
    report.classList.add("hidden");
    setResult(resultTarget, "账期已切换，请上传对应月份的供应商压缩包。");
  });

  archiveButton.addEventListener("click", async () => {
    const archive = archiveInput.files?.[0];
    if (!archive) {
      setResult(resultTarget, `请选择${config.title}账单发票 ZIP 压缩包。`, "error");
      return;
    }
    const formData = new FormData();
    formData.append("month", monthInput.value);
    formData.append("provider", config.provider);
    formData.append("archive", archive);
    setButton(archiveButton, true, "…", "正在解压");
    report.classList.add("hidden");
    setResult(resultTarget, "正在将原压缩包保存到账期目录并安全解压。");
    try {
      const result = await postForm("/api/archive-invoice-package", formData);
      sourceFolder = result.sourceFolder || "";
      setResult(
        resultTarget,
        `归档完成：共发现 ${result.pdfCount || 0} 个 PDF。解压目录：${result.sourceFolder}`,
        "success"
      );
      addOpenFileButton(resultTarget, result.sourceFolder);
    } catch (error) {
      setResult(resultTarget, `归档解压失败：${error.message}`, "error");
    } finally {
      setButton(archiveButton, false, "↓", "归档并解压");
    }
  });

  organizeButton.addEventListener("click", async () => {
    setButton(organizeButton, true, "…", "正在整理");
    report.classList.add("hidden");
    setResult(resultTarget, "正在读取账单与发票 PDF，识别账号编码和金额。");
    try {
      const result = await postJson("/api/organize-invoice-pdfs", {
        month: monthInput.value,
        provider: config.provider,
        sourceFolder
      });
      sourceFolder = result.sourceFolder || sourceFolder;
      setResult(
        resultTarget,
        `整理完成：成功命名 ${result.renamedCount || 0} 个文件，${result.warningCount || 0} 个文件需要检查。`,
        result.warningCount ? "info" : "success"
      );
      addOpenFileButton(resultTarget, result.sourceFolder);
      renderReport(result);
    } catch (error) {
      setResult(resultTarget, `PDF 文件名整理失败：${error.message}`, "error");
    } finally {
      setButton(organizeButton, false, "✓", "整理 PDF 文件名");
    }
  });
}

setupInvoiceProvider({
  id: "beijing",
  provider: "beijingLingdu",
  title: "北京凌渡"
});

setupInvoiceProvider({
  id: "jingdong",
  provider: "jingdong",
  title: "京东"
});

setupInvoiceProvider({ id: "tengyu", provider: "tengyu", title: "腾域" });














mergeInvoicePdfsByAccountButton?.addEventListener("click", async () => {
  setButton(mergeInvoicePdfsByAccountButton, true, "…", "正在合并");
  setResult(invoiceMergeResult, "正在按内部对账账号组重建渠道商账单和发票目录。");
  try {
    const result = await postJson("/api/merge-invoice-pdfs-by-account", {
      month: invoiceMergeMonth.value
    });
    setResult(
      invoiceMergeResult,
      `合并完成：生成 ${result.accountFolderCount || 0} 个子文件夹，其中 ${result.groupFolderCount || 0} 个内部账号组目录，复制 ${result.copiedCount || 0} 个文件。`,
      "success"
    );
    addOpenFileButton(invoiceMergeResult, result.outputFolder);
  } catch (error) {
    setResult(invoiceMergeResult, `渠道商账单和发票合并失败：${error.message}`, "error");
  } finally {
    setButton(mergeInvoicePdfsByAccountButton, false, "⊕", "按内部账号组合并");
  }
});

exportButton.addEventListener("click", async () => {
  const taskId = beginGenerationTask(exportButton, exportResult, "export");
  setButton(exportButton, true, "…", "正在导出");
  setResult(exportResult, "正在后台导出汇总账单。");
  try {
    const result = await postJson("/api/export-summary-bill", { month: billMonth.value, taskId });
    if (showStoppedResult(result, exportResult)) return;
    setResult(exportResult, result.exported ? `导出成功：${result.filePath}` : result.message, result.exported ? "success" : "info");
  } catch (error) {
    setResult(exportResult, `导出请求失败：${error.message}`, "error");
  } finally {
    setButton(exportButton, false, "→", "导出汇总账单");
  }
});

organizeButton.addEventListener("click", async () => {
  const taskId = beginGenerationTask(organizeButton, organizeResult, "organize");
  setButton(organizeButton, true, "…", "正在整理");
  organizeReport.classList.add("hidden");
  setResult(organizeResult, "正在检查账单并生成整理版文件。");
  try {
    const result = await postJson("/api/organize-summary-bill", { month: organizeMonth.value, taskId });
    if (showStoppedResult(result, organizeResult)) return;
    setResult(organizeResult, `整理完成：${result.outputFile}`, "success");
    renderIssues(result);
  } catch (error) {
    setResult(organizeResult, `整理请求失败：${error.message}`, "error");
  } finally {
    setButton(organizeButton, false, "✓", "开始整理");
  }
});

hmtButton.addEventListener("click", async () => {
  if (!hmtUrl.value.trim()) {
    setResult(hmtResult, "请输入港澳台账单下载 URL。", "error");
    return;
  }
  setButton(hmtButton, true, "…", "正在下载");
  setResult(hmtResult, "正在下载港澳台账单。");
  try {
    const result = await postJson("/api/import-hmt-bill", { month: hmtMonth.value, url: hmtUrl.value.trim() });
    setResult(hmtResult, `下载成功：${result.filePath}`, "success");
  } catch (error) {
    setResult(hmtResult, `下载请求失败：${error.message}`, "error");
  } finally {
    setButton(hmtButton, false, "→", "下载账单");
  }
});

exportLatestDetailButton?.addEventListener("click", async () => {
  const url = latestDetailUrl?.value.trim() || "";
  if (!url) {
    setResult(latestDetailResult, "请输入最新明细下载 URL。", "error");
    return;
  }

  setButton(exportLatestDetailButton, true, "…", "正在导出");
  setResult(latestDetailResult, "正在下载最新明细账单。");
  try {
    const result = await postJson("/api/export-latest-detail", { url });
    setResult(
      latestDetailResult,
      `导出成功：${result.filePath}${result.archivedFile ? `；归档：${result.archivedFile}` : ""}`,
      "success"
    );
    addOpenFileButton(latestDetailResult, result.filePath);
  } catch (error) {
    setResult(latestDetailResult, `最新明细导出失败：${error.message}`, "error");
  } finally {
    setButton(exportLatestDetailButton, false, "↓", "导出最新明细");
  }
});

fillLatestDetailUrlButton?.addEventListener("click", () => {
  if (!latestDetailUrl) return;
  latestDetailUrl.value = latestDetailFixedUrl;
  latestDetailUrl.focus();
  latestDetailUrl.setSelectionRange(latestDetailFixedUrl.length, latestDetailFixedUrl.length);
  setResult(latestDetailResult, "固定地址已填入，请在末尾粘贴会变化的标识。");
});

appendHmtButton.addEventListener("click", async () => {
  setButton(appendHmtButton, true, "…", "正在追加");
  setResult(hmtResult, "正在筛选江西聚梦港澳台账单并追加到整理账单。");
  try {
    const result = await postJson("/api/append-hmt-bill", { month: hmtMonth.value });
    setResult(hmtResult, `追加完成：${result.appendedRows} 行，已重新合并 ${result.mergedGroups || 0} 组合计，文件：${result.targetFile}`, "success");
  } catch (error) {
    setResult(hmtResult, `追加失败：${error.message}`, "error");
  } finally {
    setButton(appendHmtButton, false, "+", "追加到整理账单");
  }
});

aiShortButton.addEventListener("click", async () => {
  const file = aiShortFile.files && aiShortFile.files[0];
  if (!file) {
    setResult(aiShortResult, "请选择需要导入的 AI挂短账单文件。", "error");
    return;
  }
  const formData = new FormData();
  formData.append("month", aiShortMonth.value);
  formData.append("file", file);
  setButton(aiShortButton, true, "…", "正在导入");
  setResult(aiShortResult, "正在导入 AI挂短账单。");
  try {
    const result = await postForm("/api/import-ai-short-bill", formData);
    setResult(aiShortResult, `导入成功：${result.filePath}`, "success");
  } catch (error) {
    setResult(aiShortResult, `导入失败：${error.message}`, "error");
  } finally {
    setButton(aiShortButton, false, "+", "导入账单");
  }
});

appendAiShortButton.addEventListener("click", async () => {
  setButton(appendAiShortButton, true, "…", "正在追加");
  setResult(aiShortResult, "正在提取永辉语音AI的语音费用和挂短费用，并追加到整理账单。");
  try {
    const result = await postJson("/api/append-ai-short-bill", { month: aiShortMonth.value });
    setResult(aiShortResult, `追加完成：${result.appendedRows} 行，语音费用 ${result.voiceFee || "-"}，挂短费用 ${result.shortFee || "-"}，合计 ${result.totalFee || "-"}，文件：${result.targetFile}`, "success");
  } catch (error) {
    setResult(aiShortResult, `追加失败：${error.message}`, "error");
  } finally {
    setButton(appendAiShortButton, false, "+", "追加到整理账单");
  }
});

oneClickButton.addEventListener("click", async () => {
  const taskId = beginGenerationTask(oneClickButton, oneClickResult, "oneclick");
  setButton(oneClickButton, true, "…", "正在生成");
  oneClickReport.innerHTML = "";
  setResult(oneClickResult, "正在准备账单并按流程生成汇总账单。");
  try {
    const setupSteps = [];
    const hmtUrlValue = oneClickHmtUrl.value.trim();
    if (hmtUrlValue) {
      setResult(oneClickResult, "正在下载港澳台账单。");
      const hmtImport = await postJson("/api/import-hmt-bill", { month: oneClickMonth.value, url: hmtUrlValue });
      setupSteps.push({ name: "港澳台账单下载", ok: true, skipped: false, message: `下载成功：${hmtImport.filePath}` });
    } else {
      setupSteps.push({ name: "港澳台账单下载", ok: true, skipped: true, message: "未填写下载 URL，使用已导入文件或跳过。" });
    }

    const aiFile = oneClickAiShortFile.files && oneClickAiShortFile.files[0];
    if (aiFile) {
      setResult(oneClickResult, "正在导入 AI挂短账单。");
      const formData = new FormData();
      formData.append("month", oneClickMonth.value);
      formData.append("file", aiFile);
      const aiImport = await postForm("/api/import-ai-short-bill", formData);
      setupSteps.push({ name: "AI挂短账单导入", ok: true, skipped: false, message: `导入成功：${aiImport.filePath}` });
    } else {
      setupSteps.push({ name: "AI挂短账单导入", ok: true, skipped: true, message: "未选择文件，使用已导入文件或跳过。" });
    }

    setResult(oneClickResult, "正在执行导出、整理和追加。");
    const result = await postJson("/api/generate-one-click-bill", { month: oneClickMonth.value, taskId });
    if (showStoppedResult(result, oneClickResult)) return;
    result.steps = [...setupSteps, ...(result.steps || [])];
    setResult(oneClickResult, `${result.message} 文件：${result.outputFile || "-"}`, "success");
    addOpenFileButton(oneClickResult, result.outputFile);
    renderOneClickReport(result);
  } catch (error) {
    setResult(oneClickResult, `一键生成失败：${error.message}`, "error");
  } finally {
    setButton(oneClickButton, false, "✓", "一键生成汇总账单");
  }
});

async function runFilterBill(config) {
  const taskId = beginGenerationTask(config.button, config.result, `filter-${config.type}`);
  setButton(config.button, true, "…", "正在筛选");
  setResult(config.result, `正在筛选${config.title}账单。`);
  try {
    const result = await postJson("/api/filter-yh-bill", {
      month: config.month.value,
      type: config.type,
      taskId
    });
    if (showStoppedResult(result, config.result)) return;
    setResult(
      config.result,
      `${config.title}筛选完成：${result.matchedRows} 行，金额合计 ${result.amountTotal ?? 0}，文件：${result.outputFile}`,
      "success"
    );
    addOpenFileButton(config.result, result.outputFile);
  } catch (error) {
    setResult(config.result, `${config.title}筛选失败：${error.message}`, "error");
  } finally {
    setButton(config.button, false, "✓", "生成筛选账单");
  }
}

[
  {
    type: "jiangxiJumeng",
    title: "江西聚梦",
    month: filterBillsMonth,
    button: jiangxiJumengButton,
    result: jiangxiJumengResult
  },
  {
    type: "beijingLingdu",
    title: "北京凌渡",
    month: filterBillsMonth,
    button: beijingLingduButton,
    result: beijingLingduResult
  },
  {
    type: "jingdong",
    title: "京东",
    month: filterBillsMonth,
    button: jingdongButton,
    result: jingdongResult
  },
  {
    type: "tengyu", title: "腾域", month: filterBillsMonth, button: tengyuButton, result: tengyuResult
  },
{
  },
  {
  },
  {
  },
  {
  },
  {
    type: "techBill",
    title: "科技账单",
    month: filterBillsMonth,
    button: techBillButton,
    result: techBillResult
  }
].forEach((config) => {
  config.button?.addEventListener("click", () => runFilterBill(config));
});

combinedWorkbookButton?.addEventListener("click", async () => {
  const taskId = beginGenerationTask(combinedWorkbookButton, combinedWorkbookResult, "combine");
  setButton(combinedWorkbookButton, true, "…", "正在合并");
  combinedWorkbookReport.innerHTML = "";
  setResult(combinedWorkbookResult, "正在将各模块账单合并到同一个 Excel。");
  try {
    const result = await postJson("/api/combine-workbooks", { month: combinedWorkbookMonth.value, taskId });
    if (showStoppedResult(result, combinedWorkbookResult)) return;
    setResult(combinedWorkbookResult, `合并完成：${result.outputFile}`, "success");
    addOpenFileButton(combinedWorkbookResult, result.outputFile);
    const sheets = result.sheets || [];
    combinedWorkbookReport.innerHTML = `
      <div class="step-list">
        ${sheets.map((sheet) => `
          <div class="step-row" data-status="ok">
            <strong>${sheet}</strong>
            <span>已加入合并 Excel</span>
          </div>
        `).join("")}
      </div>
      ${result.outputFile ? `<p class="output-path">合并账单：${result.outputFile}</p>` : ""}
    `;
  } catch (error) {
    setResult(combinedWorkbookResult, `合并失败：${error.message}`, "error");
  } finally {
    setButton(combinedWorkbookButton, false, "✓", "生成合并 Excel");
  }
});

feishuThirdPartyButton?.addEventListener("click", async () => {
  const taskId = beginGenerationTask(feishuThirdPartyButton, feishuThirdPartyResult, "feishu3p");
  setButton(feishuThirdPartyButton, true, "鈥?", "正在生成");
  feishuThirdPartyReport.innerHTML = "";
  setResult(feishuThirdPartyResult, "正在读取费用对比表并生成三方群话术。");
  try {
    const result = await postJson("/api/generate-feishu-third-party-script", {
      month: feishuThirdPartyMonth.value,
      taskId
    });
    if (showStoppedResult(result, feishuThirdPartyResult)) return;
    setResult(feishuThirdPartyResult, `话术生成完成：${result.outputFile}`, "success");
    addOpenFileButton(feishuThirdPartyResult, result.outputFile);
    renderScriptReport(feishuThirdPartyReport, result.messages);
  } catch (error) {
    setResult(feishuThirdPartyResult, `话术生成失败：${error.message}`, "error");
  } finally {
    setButton(feishuThirdPartyButton, false, "✓", "生成话术");
  }
});

function renderInternalGroups() {
  if (!internalGroupsEditor) return;
  internalGroupsEditor.innerHTML = "";
  internalGroups.forEach((group, index) => {
    const row = document.createElement("div");
    row.className = "internal-group-row";
    row.dataset.index = String(index);
    row.innerHTML = `
      <div class="internal-group-number">${index + 1}</div>
      <label>
        <span>账号编码</span>
        <textarea data-field="accounts" rows="2" placeholder="s_scyh, s_scyh_hxzx">${group.accounts.join(", ")}</textarea>
      </label>
      <label>
        <span>对接人</span>
        <input data-field="contacts" type="text" value="${group.contacts.join(", ")}" placeholder="刘香君">
      </label>
      <button class="danger-action" type="button" data-remove-group>删除</button>
    `;
    row.querySelector("[data-remove-group]").addEventListener("click", () => {
      internalGroups.splice(index, 1);
      renderInternalGroups();
    });
    internalGroupsEditor.appendChild(row);
  });
}

function readInternalGroupsFromEditor() {
  return Array.from(internalGroupsEditor?.querySelectorAll(".internal-group-row") || []).map((row, index) => {
    const accounts = row.querySelector("[data-field='accounts']").value
      .split(/[,\n，、]/)
      .map((value) => value.trim())
      .filter(Boolean);
    const contacts = row.querySelector("[data-field='contacts']").value
      .split(/[,\n，、]/)
      .map((value) => value.trim())
      .filter(Boolean);
    return { id: index + 1, accounts, contacts };
  });
}

async function loadInternalGroups() {
  if (!internalGroupsEditor) return;
  try {
    const response = await fetch(`${apiBase}/api/internal-reconciliation-groups`);
    const result = await readJsonResponse(response);
    if (!response.ok) throw new Error(result.message || "账号组读取失败");
    internalGroups = result.groups || [];
    renderInternalGroups();
  } catch (error) {
    setResult(internalGroupsResult, `账号组读取失败：${error.message}`, "error");
  }
}

addInternalGroupButton?.addEventListener("click", () => {
  internalGroups = readInternalGroupsFromEditor();
  internalGroups.push({ id: internalGroups.length + 1, accounts: [""], contacts: [""] });
  renderInternalGroups();
});

saveInternalGroupsButton?.addEventListener("click", async () => {
  internalGroups = readInternalGroupsFromEditor();
  try {
    const result = await fetch(`${apiBase}/api/internal-reconciliation-groups`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ groups: internalGroups })
    });
    const payload = await readJsonResponse(result);
    if (!result.ok) throw new Error(payload.message || "账号组保存失败");
    internalGroups = payload.groups || internalGroups;
    renderInternalGroups();
    setResult(internalGroupsResult, "账号组已保存。", "success");
  } catch (error) {
    setResult(internalGroupsResult, `账号组保存失败：${error.message}`, "error");
  }
});

function accountGroupLabel(groupIds, groupNameById = new Map()) {
  const rawIds = Array.isArray(groupIds) ? groupIds : (groupIds == null ? [] : [groupIds]);
  const ids = rawIds.map((id) => Number(id)).filter((id) => Number.isFinite(id));
  return ids.map((id) => groupNameById.get(String(id)) || `第 ${id} 组`).join("、");
}

function renderAccountChips(accounts, emptyText, kind, groupById = new Map(), groupNameById = new Map()) {
  if (!accounts.length) {
    return `<p class="account-check-empty">${escapeClipboardHtml(emptyText)}</p>`;
  }
  return `
    <div class="account-chip-list">
      ${accounts.map((account) => {
        const label = accountGroupLabel(groupById.get(String(account)), groupNameById);
        return `
          <span class="account-chip ${kind}">
            <strong>${escapeClipboardHtml(account)}</strong>
            ${label ? `<small>${escapeClipboardHtml(label)}</small>` : ""}
          </span>
        `;
      }).join("")}
    </div>
  `;
}

function renderInternalAccountCheck(result) {
  if (!internalAccountCheckResult) return;
  const coveredAccounts = result.coveredAccounts || [];
  const uncoveredAccounts = result.uncoveredAccounts || [];
  const unusedAccounts = result.unusedAccounts || [];
  const duplicateAccounts = result.duplicateAccounts || [];
  const groupNameById = new Map(
    (result.groups || [])
      .filter((group) => group.name)
      .map((group) => [String(group.id), group.name])
  );
  const coverageByAccount = new Map(
    (result.coverage || []).map((item) => [String(item.account), item.groupIds || []])
  );
  const unusedByAccount = new Map(
    unusedAccounts.map((item) => [String(item.account), item.groupIds || []])
  );
  const complete = Boolean(result.complete);
  const statusText = complete ? "完整性校验通过" : "发现未完整覆盖";

  internalAccountCheckResult.innerHTML = `
    <div class="account-check-card ${complete ? "is-complete" : "has-warning"}">
      <div class="account-check-head">
        <div>
          <span class="account-check-kicker">账号编码完整性</span>
          <strong>${escapeClipboardHtml(result.message || statusText)}</strong>
        </div>
        <span class="account-check-status">${statusText}</span>
      </div>
      <div class="account-check-metrics">
        <span><strong>${Number(result.billAccountCount || 0)}</strong> 本月账单账号</span>
        <span><strong>${Number(result.coveredAccountCount || 0)}</strong> 已包含</span>
        <span class="${uncoveredAccounts.length ? "is-warning" : ""}"><strong>${Number(result.uncoveredAccountCount || 0)}</strong> 未包含</span>
        <span><strong>${Number(result.unusedAccountCount || 0)}</strong> 本月未出现</span>
      </div>
      <div class="account-check-grid">
        <section>
          <h3>已包含的账号</h3>
          ${renderAccountChips(
            coveredAccounts,
            "没有已验证包含的账号。",
            "is-covered",
            coverageByAccount,
            groupNameById
          )}
        </section>
        <section class="${uncoveredAccounts.length ? "account-check-attention" : ""}">
          <h3>未包含的账号</h3>
          ${renderAccountChips(uncoveredAccounts, "本月账单账号已全部包含。", "is-missing")}
        </section>
        <section>
          <h3>已配置但本月未出现</h3>
          ${renderAccountChips(
            unusedAccounts.map((item) => item.account),
            "没有闲置的账号编码。",
            "is-unused",
            unusedByAccount,
            groupNameById
          )}
        </section>
      </div>
      ${duplicateAccounts.length ? `
        <div class="account-check-duplicates">
          <strong>重复配置：</strong>
          ${duplicateAccounts.map((item) => `${escapeClipboardHtml(item.account)}（${escapeClipboardHtml(accountGroupLabel(item.groupIds, groupNameById))}）`).join("、")}
        </div>
      ` : ""}
    </div>
  `;
}

checkInternalAccountsButton?.addEventListener("click", async () => {
  const taskId = beginGenerationTask(
    checkInternalAccountsButton,
    internalAccountCheckResult,
    "check-internal-accounts"
  );
  setButton(checkInternalAccountsButton, true, "…", "正在校验");
  internalAccountCheckResult.innerHTML = "";
  setResult(internalAccountCheckResult, "正在读取本月账单并比对已保存账号组。");
  try {
    const result = await postJson("/api/check-internal-reconciliation-accounts", {
      month: feishuInternalMonth.value,
      taskId
    });
    if (showStoppedResult(result, internalAccountCheckResult)) return;
    renderInternalAccountCheck(result);
  } catch (error) {
    setResult(internalAccountCheckResult, `账号完整性校验失败：${error.message}`, "error");
  } finally {
    setButton(checkInternalAccountsButton, false, "✓", "校验账号完整性");
  }
});

function escapeClipboardHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

async function fetchInternalImage(group) {
  if (!group?.assetFile) throw new Error("该账号组没有可复制的截图。");
  const response = await fetch(`${apiBase}/api/generated-image?file=${encodeURIComponent(group.assetFile)}`);
  if (!response.ok) throw new Error("截图读取失败。");
  return response.blob();
}

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error || new Error("截图转换失败。"));
    reader.readAsDataURL(blob);
  });
}

function loadImageFromDataUrl(dataUrl) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("截图加载失败。"));
    image.src = dataUrl;
  });
}

function wrapCanvasText(context, text, maxWidth) {
  const lines = [];
  String(text || "").split(/\r?\n/).forEach((paragraph) => {
    if (!paragraph) {
      lines.push("");
      return;
    }
    let line = "";
    Array.from(paragraph).forEach((char) => {
      const testLine = line + char;
      if (context.measureText(testLine).width > maxWidth && line) {
        lines.push(line);
        line = char;
      } else {
        line = testLine;
      }
    });
    lines.push(line);
  });
  return lines;
}

function canvasToPngBlob(canvas) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (blob) resolve(blob);
      else reject(new Error("合成图片失败。"));
    }, "image/png");
  });
}

async function buildInternalGroupImage(group) {
  const imageBlob = await fetchInternalImage(group);
  const imageDataUrl = await blobToDataUrl(imageBlob);
  const image = await loadImageFromDataUrl(imageDataUrl);

  const padding = 28;
  const gap = 24;
  const imageWidth = image.naturalWidth || image.width;
  const imageHeight = image.naturalHeight || image.height;
  const canvasWidth = Math.max(imageWidth + padding * 2, 1100);
  const textWidth = canvasWidth - padding * 2;
  const measureCanvas = document.createElement("canvas");
  const measureContext = measureCanvas.getContext("2d");
  measureContext.font = "28px Microsoft YaHei, Segoe UI, Arial, sans-serif";
  const lines = wrapCanvasText(measureContext, group.message, textWidth);
  const lineHeight = 46;
  const textHeight = lines.length * lineHeight;
  const canvasHeight = padding + imageHeight + gap + textHeight + padding;
  const canvas = document.createElement("canvas");
  canvas.width = canvasWidth;
  canvas.height = canvasHeight;
  const context = canvas.getContext("2d");

  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  context.drawImage(image, padding, padding);
  context.fillStyle = "#111827";
  context.font = "28px Microsoft YaHei, Segoe UI, Arial, sans-serif";
  context.textBaseline = "top";

  let y = padding + imageHeight + gap;
  lines.forEach((line) => {
    context.fillText(line, padding, y);
    y += lineHeight;
  });

  return canvasToPngBlob(canvas);
}

async function copyInternalGroup(group) {
  if (!window.ClipboardItem || !navigator.clipboard?.write) {
    throw new Error("当前浏览器不支持富文本复制，请分别复制图片和话术。");
  }
  const imageBlob = await fetchInternalImage(group);
  const imageDataUrl = await blobToDataUrl(imageBlob);
  const plainText = group.message;
  const html = `
    <div>
      <p><img src="${imageDataUrl}" alt="内部对账截图"></p>
      <div style="white-space:pre-wrap">${escapeClipboardHtml(group.message)}</div>
    </div>
  `;
  await navigator.clipboard.write([
    new ClipboardItem({
      "image/png": imageBlob,
      "text/html": new Blob([html], { type: "text/html" }),
      "text/plain": new Blob([plainText], { type: "text/plain" })
    })
  ]);
}

function renderInternalReconciliationReport(result) {
  if (!feishuInternalReport) return;
  const groups = result.groups || [];
  const groupsById = new Map(groups.map((group) => [String(group.id), group]));
  feishuInternalReport.innerHTML = `
    <div class="script-list">
      ${groups.map((group) => `
        <div class="script-item internal-script-item">
          <div class="script-item-head">
            <strong>${group.name ? `${escapeClipboardHtml(group.name)} · ` : ""}第${group.id}组：${group.contacts.join("、")} · ${group.accounts.join("、")}</strong>
            <div class="script-item-actions">
              ${group.assetFile ? `<button class="copy-script-button" type="button" data-copy-group="${group.id}">复制整组</button>` : ""}
              ${group.assetFile ? `<button class="copy-script-button" type="button" data-copy-image="${group.id}">复制图片</button>` : ""}
              <button class="copy-script-button" type="button" data-copy-internal="${group.id}">复制话术</button>
            </div>
          </div>
          ${group.warning ? `<p class="internal-warning">${group.warning}</p>` : ""}
          ${group.hasImage ? `<p class="output-path">匹配 ${group.matchedRows || 0} 行，可直接复制到飞书</p>` : `<p class="output-path">当前账期没有匹配数据</p>`}
          <pre>${group.message}</pre>
        </div>
      `).join("")}
    </div>
  `;

  feishuInternalReport.querySelectorAll("[data-copy-group]").forEach((button) => {
    button.addEventListener("click", async () => {
      const group = groupsById.get(button.dataset.copyGroup);
      try {
        await copyInternalGroup(group);
        button.textContent = "已复制整组";
        setTimeout(() => { button.textContent = "复制整组"; }, 1500);
      } catch (error) {
        setResult(feishuInternalResult, `复制整组失败：${error.message}`, "error");
      }
    });
  });

  feishuInternalReport.querySelectorAll("[data-copy-image]").forEach((button) => {
    button.addEventListener("click", async () => {
      const group = groupsById.get(button.dataset.copyImage);
      try {
        const imageBlob = await fetchInternalImage(group);
        await navigator.clipboard.write([new ClipboardItem({ "image/png": imageBlob })]);
        button.textContent = "图片已复制";
        setTimeout(() => { button.textContent = "复制图片"; }, 1500);
      } catch (error) {
        setResult(feishuInternalResult, `复制图片失败：${error.message}`, "error");
      }
    });
  });

  feishuInternalReport.querySelectorAll("[data-copy-internal]").forEach((button) => {
    button.addEventListener("click", async () => {
      const group = groupsById.get(button.dataset.copyInternal);
      const text = `${group.accounts.join("、")}：\n\n${group.message}`;
      try {
        await navigator.clipboard.writeText(text);
        button.textContent = "话术已复制";
        setTimeout(() => { button.textContent = "复制话术"; }, 1500);
      } catch {
        button.textContent = "复制失败";
      }
    });
  });
}

feishuInternalButton?.addEventListener("click", async () => {
  const taskId = beginGenerationTask(feishuInternalButton, feishuInternalResult, "feishu-internal");
  setButton(feishuInternalButton, true, "…", "正在生成");
  feishuInternalReport.innerHTML = "";
  setResult(feishuInternalResult, "正在从账单合并中筛选账号并生成截图和话术。");
  try {
    const result = await postJson("/api/generate-internal-reconciliation", {
      month: feishuInternalMonth.value,
      cutoff: feishuInternalCutoff.value,
      taskId
    });
    if (showStoppedResult(result, feishuInternalResult)) return;
    setResult(feishuInternalResult, `生成完成：${result.outputFile}`, "success");
    addOpenFileButton(feishuInternalResult, result.outputFile);
    renderInternalReconciliationReport(result);
  } catch (error) {
    setResult(feishuInternalResult, `生成失败：${error.message}`, "error");
  } finally {
    setButton(feishuInternalButton, false, "✓", "生成截图和话术");
  }
});

loadInternalGroups();

expenseCompareButton.addEventListener("click", async () => {
  const requiredFiles = [
    [expenseYhFile, "永辉整理账单"],
    [expenseJmFile, "江西聚梦短信/港澳台账单"],
    [expenseJmAiFile, "江西聚梦AI语音挂短账单"],
    [expenseBeijingFile, "北京凌渡账单"],
    [expenseJdFile, "京东账单"]
  ];
  const missing = requiredFiles.find(([input]) => !input.files || !input.files[0]);
  if (missing) {
    setResult(expenseCompareResult, `请选择${missing[1]}。`, "error");
    return;
  }

  const taskId = beginGenerationTask(expenseCompareButton, expenseCompareResult, "expense");
  const formData = new FormData();
  formData.append("month", expenseMonth.value);
  formData.append("taskId", taskId);
  formData.append("yhFile", expenseYhFile.files[0]);
  formData.append("jmFile", expenseJmFile.files[0]);
  formData.append("jmAiFile", expenseJmAiFile.files[0]);
  formData.append("beijingFile", expenseBeijingFile.files[0]);
  formData.append("jdFile", expenseJdFile.files[0]);
  formData.append("tengyuFile", expenseTengyuFile.files[0]);

  setButton(expenseCompareButton, true, "…", "正在生成");
  expenseCompareReport.innerHTML = "";
  setResult(expenseCompareResult, "正在生成费用对比 Excel。");
  try {
    const result = await postForm("/api/generate-expense-compare", formData);
    if (showStoppedResult(result, expenseCompareResult)) return;
    setResult(expenseCompareResult, `费用对比生成完成：${result.outputFile}`, "success");
    addOpenFileButton(expenseCompareResult, result.outputFile);
    renderExpenseCompareReport(result);
  } catch (error) {
    setResult(expenseCompareResult, `费用对比生成失败：${error.message}`, "error");
  } finally {
    setButton(expenseCompareButton, false, "✓", "一键生成费用对比");
  }
});
