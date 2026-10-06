const SPREADSHEET_ID = "1qCgTkXKG0N7snFfWMX-6jxF75bHE_mU_M6qKu2A1238";

const CRAWLER = {
  PROJECTS_SHEET: "Search_Projects",
  QUEUE_SHEET: "Crawl_Queue",
  RESULTS_SHEET: "Crawl_Results",
  LOG_SHEET: "Sheet1",
  MAX_SEEDS: 20,
  MAX_PAGES_PER_SEED: 8,
  MAX_RESULTS_PER_PAGE: 4,
  REQUEST_TIMEOUT_MS: 20000,
  USER_AGENT: "KampusReachBot/1.0 (+public-campus-contact-research)"
};

const PROJECT_HEADERS = [
  "Project ID", "Created At", "Project", "Focus", "Region", "Category",
  "Role Focus", "Mode", "Source URLs", "PIC Name", "PIC Contact", "PIC Unit",
  "Status", "Total URLs", "Processed URLs", "Result Count", "Last Update", "Error"
];
const QUEUE_HEADERS = [
  "Queue ID", "Project ID", "URL", "Status", "Attempts", "Added At",
  "Started At", "Finished At", "Error"
];
const RESULT_HEADERS = [
  "Result ID", "Project ID", "Crawled At", "Organization", "Category", "Role",
  "Email", "Phone", "Address", "Region", "Website", "Source URL", "Evidence",
  "Confidence", "Readiness", "Review Status"
];

/** Jalankan sekali dari editor Apps Script untuk membuat tab dan trigger crawler. */
function setupCrawler() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureSheet_(ss, CRAWLER.PROJECTS_SHEET, PROJECT_HEADERS, "#0f6b68");
  ensureSheet_(ss, CRAWLER.QUEUE_SHEET, QUEUE_HEADERS, "#d58b31");
  ensureSheet_(ss, CRAWLER.RESULTS_SHEET, RESULT_HEADERS, "#377a9b");
  ensureSheet_(ss, CRAWLER.LOG_SHEET, ["Timestamp", "Project", "Region", "Category", "Keywords", "PIC Name", "PIC Contact", "PIC Role", "Status"], "#71818b");

  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === "processCrawlQueue")
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));
  ScriptApp.newTrigger("processCrawlQueue").timeBased().everyMinutes(5).create();
  return { ok: true, message: "Crawler siap", spreadsheetId: SPREADSHEET_ID };
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  let response;
  try {
    const action = params.action || "ping";
    if (action === "status" || action === "results") {
      response = getProjectBundle_(String(params.projectId || ""));
    } else if (action === "process") {
      response = processCrawlQueue_(1, 18000);
    } else {
      response = { ok: true, service: "KampusReach Google Sheets crawler", version: "2.0" };
    }
  } catch (error) {
    response = { ok: false, error: String(error && error.message || error) };
  }
  return output_(response, params.callback);
}

function doPost(e) {
  let payload = {};
  try {
    payload = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    ensureCrawlerSheets_();
    if (payload.action === "rerun_search") {
      const response = requeueSearch_(payload);
      if (response.ok && response.queued > 0) processCrawlQueue_(2, 22000);
      return output_(response);
    }
    if (payload.action === "create_search" || payload.type === "search") {
      const response = createSearch_(payload);
      if (response.ok && response.queued > 0) processCrawlQueue_(2, 22000);
      return output_(response);
    }
    return output_({ ok: false, error: "Action tidak dikenal" });
  } catch (error) {
    return output_({ ok: false, error: String(error && error.message || error) });
  }
}

function requeueSearch_(payload) {
  const projectId = clean_(payload.projectId, 80);
  if (!projectId) return { ok: false, error: "projectId wajib diisi" };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const project = getProjectRow_(ss, projectId);
  const urls = unique_(String(payload.sourceUrls || project.sourceUrls || "").split(/[\n,;]+/)
    .map(value => normalizeUrl_(value.trim()))
    .filter(url => url && isSafePublicUrl_(url)))
    .slice(0, CRAWLER.MAX_SEEDS);
  if (!urls.length) return { ok: false, error: "Project tidak memiliki URL publik" };
  const queue = ss.getSheetByName(CRAWLER.QUEUE_SHEET);
  const now = new Date();
  const rows = urls.map(url => [Utilities.getUuid(), projectId, url, "Queued", 0, now, "", "", ""]);
  queue.getRange(queue.getLastRow() + 1, 1, rows.length, QUEUE_HEADERS.length).setValues(rows);
  ss.getSheetByName(CRAWLER.PROJECTS_SHEET).getRange(project.row, 13, 1, 6).setValues([["Sedang mencari", urls.length, 0, project.resultCount, now, ""]]);
  return { ok: true, projectId: projectId, queued: urls.length, status: "Sedang mencari" };
}

function createSearch_(payload) {
  const projectId = clean_(payload.projectId, 80) || Utilities.getUuid();
  const projectName = clean_(payload.project, 180);
  if (!projectName) return { ok: false, error: "Nama project wajib diisi" };

  const urls = unique_(String(payload.sourceUrls || "").split(/[\n,;]+/)
    .map(value => value.trim())
    .filter(Boolean)
    .map(normalizeUrl_)
    .filter(url => url && isSafePublicUrl_(url)))
    .slice(0, CRAWLER.MAX_SEEDS);
  if (!urls.length) return { ok: false, error: "Tambahkan minimal satu URL publik yang valid" };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const projects = ss.getSheetByName(CRAWLER.PROJECTS_SHEET);
  const queue = ss.getSheetByName(CRAWLER.QUEUE_SHEET);
  const log = ss.getSheetByName(CRAWLER.LOG_SHEET);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const now = new Date();
    projects.appendRow([
      projectId, now, projectName, clean_(payload.focus || payload.keywords, 500),
      clean_(payload.region, 120), clean_(payload.category, 100), clean_(payload.roleFocus, 250),
      clean_(payload.mode, 30) || "Deep", urls.join("\n"), clean_(payload.picName, 150),
      clean_(payload.picContact, 200), clean_(payload.picRole || payload.picUnit, 150),
      "Sedang mencari", urls.length, 0, 0, now, ""
    ]);
    const rows = urls.map(url => [Utilities.getUuid(), projectId, url, "Queued", 0, now, "", "", ""]);
    queue.getRange(queue.getLastRow() + 1, 1, rows.length, QUEUE_HEADERS.length).setValues(rows);
    log.appendRow([
      now, projectName, clean_(payload.region, 120), clean_(payload.category, 100),
      clean_(payload.focus || payload.keywords, 500), clean_(payload.picName, 150),
      clean_(payload.picContact, 200), clean_(payload.picRole || payload.picUnit, 150), "Sedang mencari"
    ]);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, projectId: projectId, queued: urls.length, status: "Sedang mencari" };
}

function processCrawlQueue() {
  return processCrawlQueue_(5, 280000);
}

/** Jalankan sekali bila Google belum meminta izin untuk mengakses website publik. */
function authorizeCrawlerNetwork() {
  const response = UrlFetchApp.fetch("https://example.com/", {
    muteHttpExceptions: true,
    headers: { "User-Agent": CRAWLER.USER_AGENT }
  });
  return { ok: response.getResponseCode() < 500, status: response.getResponseCode() };
}

function processCrawlQueue_(limit, maxRuntimeMs) {
  ensureCrawlerSheets_();
  const started = Date.now();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const queueSheet = ss.getSheetByName(CRAWLER.QUEUE_SHEET);
  const resultSheet = ss.getSheetByName(CRAWLER.RESULTS_SHEET);
  if (queueSheet.getLastRow() < 2) return { ok: true, processed: 0 };

  const rows = queueSheet.getRange(2, 1, queueSheet.getLastRow() - 1, QUEUE_HEADERS.length).getValues();
  let processed = 0;
  const touchedProjects = {};

  for (let index = 0; index < rows.length; index++) {
    if (processed >= limit || Date.now() - started > maxRuntimeMs) break;
    const row = rows[index];
    const status = String(row[3] || "");
    const attempts = Number(row[4] || 0);
    if (!["Queued", "Retry"].includes(status) || attempts >= 3) continue;

    const sheetRow = index + 2;
    const projectId = String(row[1]);
    const url = String(row[2]);
    touchedProjects[projectId] = true;
    queueSheet.getRange(sheetRow, 4, 1, 4).setValues([["Running", attempts + 1, row[5] || new Date(), new Date()]]);
    SpreadsheetApp.flush();

    try {
      const project = getProjectRow_(ss, projectId);
      const resultRows = crawlSeed_(url, project, maxRuntimeMs - (Date.now() - started));
      appendUniqueResults_(resultSheet, resultRows);
      queueSheet.getRange(sheetRow, 4).setValue("Done");
      queueSheet.getRange(sheetRow, 8, 1, 2).setValues([[new Date(), ""]]);
    } catch (error) {
      const nextStatus = attempts + 1 >= 3 ? "Error" : "Retry";
      queueSheet.getRange(sheetRow, 4).setValue(nextStatus);
      queueSheet.getRange(sheetRow, 8, 1, 2).setValues([[new Date(), clean_(error.message || error, 500)]]);
    }
    processed++;
  }

  Object.keys(touchedProjects).forEach(projectId => updateProjectSummary_(ss, projectId));
  return { ok: true, processed: processed };
}

function crawlSeed_(seedUrl, project, runtimeLeftMs) {
  if (!isSafePublicUrl_(seedUrl)) throw new Error("URL tidak diizinkan");
  if (!isAllowedByRobots_(seedUrl)) throw new Error("robots.txt melarang crawling");

  const maxPages = project.mode === "Quick" ? 3 : project.mode === "Strict" ? 5 : CRAWLER.MAX_PAGES_PER_SEED;
  const urls = [seedUrl];
  const visited = {};
  const output = [];
  const started = Date.now();

  while (urls.length && Object.keys(visited).length < maxPages && Date.now() - started < Math.max(6000, runtimeLeftMs - 2500)) {
    const url = urls.shift();
    if (visited[url] || !isSafePublicUrl_(url)) continue;
    visited[url] = true;
    const page = fetchHtml_(url);
    const extracted = extractContacts_(page.html, page.finalUrl, project);
    extracted.rows.forEach(row => output.push(row));
    if (Object.keys(visited).length === 1) {
      extractUsefulLinks_(page.html, page.finalUrl).forEach(link => {
        if (!visited[link] && sameHost_(link, seedUrl)) urls.push(link);
      });
    }
  }
  return output;
}

function fetchHtml_(url) {
  const response = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    validateHttpsCertificates: true,
    headers: { "User-Agent": CRAWLER.USER_AGENT, "Accept": "text/html,application/xhtml+xml" }
  });
  const code = response.getResponseCode();
  if (code < 200 || code >= 400) throw new Error("HTTP " + code + " untuk " + url);
  const headers = response.getHeaders();
  const contentType = String(headers["Content-Type"] || headers["content-type"] || "");
  if (contentType && !/text\/html|application\/xhtml\+xml/i.test(contentType)) throw new Error("Konten bukan HTML");
  const html = response.getContentText().slice(0, 1500000);
  return { html: html, finalUrl: url };
}

function extractContacts_(html, sourceUrl, project) {
  const titleMatch = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  const title = cleanText_(titleMatch ? titleMatch[1] : "");
  const text = cleanText_(html
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<!--([\s\S]*?)-->/g, " ")
    .replace(/<[^>]+>/g, " "));

  const emails = unique_((text.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi) || [])
    .map(value => value.toLowerCase())
    .filter(value => !/\.(png|jpg|jpeg|gif|svg|webp)$/i.test(value)))
    .slice(0, CRAWLER.MAX_RESULTS_PER_PAGE);
  const phones = unique_((text.match(/(?:\+62|62|0)[\s().-]?(?:\d[\s().-]?){8,13}\d/g) || [])
    .map(normalizePhone_)
    .filter(value => value.length >= 9 && value.length <= 16))
    .slice(0, CRAWLER.MAX_RESULTS_PER_PAGE);
  if (!emails.length && !phones.length) return { rows: [] };

  const organization = clean_(title.replace(/\s*[|–—-]\s*(Beranda|Home|Official|Website).*$/i, ""), 180) || hostOf_(sourceUrl);
  const role = inferRole_(text, project.roleFocus);
  const category = inferCategory_(text, project.category);
  const address = extractAddress_(text);
  const resultCount = Math.max(emails.length, phones.length, 1);
  const rows = [];
  for (let i = 0; i < resultCount; i++) {
    const email = emails[i] || emails[0] || "";
    const phone = phones[i] || phones[0] || "";
    let confidence = 42;
    if (/^https:\/\//i.test(sourceUrl)) confidence += 8;
    if (email) confidence += 18;
    if (email && email.split("@")[1] && hostOf_(sourceUrl).includes(email.split("@")[1].replace(/^www\./, ""))) confidence += 10;
    if (phone) confidence += 10;
    if (address) confidence += 6;
    if (role !== "Kontak umum") confidence += 6;
    confidence = Math.min(98, confidence);
    const evidence = [email && "email publik", phone && "telepon publik", address && "alamat", role !== "Kontak umum" && "indikasi unit/PIC"].filter(Boolean).join(" · ");
    rows.push([
      Utilities.getUuid(), project.projectId, new Date(), organization, category, role, email, phone,
      address, project.region || "Indonesia", rootUrl_(sourceUrl), sourceUrl, evidence,
      confidence, confidence >= 75 && email ? "Siap Dihubungi" : "Perlu Verifikasi",
      confidence >= 65 ? "Auto-stored" : "Perlu Review"
    ]);
  }
  return { rows: rows };
}

function extractUsefulLinks_(html, baseUrl) {
  const links = [];
  const regex = /<a\b[^>]*href\s*=\s*["']([^"'#]+)["'][^>]*>([\s\S]*?)<\/a>/gi;
  let match;
  while ((match = regex.exec(html)) !== null) {
    const label = cleanText_(match[2]);
    const href = match[1];
    if (!/(kontak|contact|hubungi|humas|kerja.?sama|cooperation|akademik|profil|about|direktori|directory|pimpinan|staff)/i.test(label + " " + href)) continue;
    const absolute = resolveUrl_(href, baseUrl);
    if (absolute && isSafePublicUrl_(absolute)) links.push(absolute);
  }
  return unique_(links).slice(0, CRAWLER.MAX_PAGES_PER_SEED - 1);
}

function appendUniqueResults_(sheet, rows) {
  if (!rows.length) return 0;
  const existing = {};
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, RESULT_HEADERS.length).getValues().forEach(row => {
      existing[[row[1], row[6], row[7], row[11]].join("|").toLowerCase()] = true;
    });
  }
  const fresh = rows.filter(row => {
    const key = [row[1], row[6], row[7], row[11]].join("|").toLowerCase();
    if (existing[key]) return false;
    existing[key] = true;
    return true;
  });
  if (fresh.length) sheet.getRange(sheet.getLastRow() + 1, 1, fresh.length, RESULT_HEADERS.length).setValues(fresh);
  return fresh.length;
}

function getProjectBundle_(projectId) {
  if (!projectId) return { ok: false, error: "projectId wajib diisi" };
  ensureCrawlerSheets_();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  updateProjectSummary_(ss, projectId);
  const project = getProjectRow_(ss, projectId);
  const resultsSheet = ss.getSheetByName(CRAWLER.RESULTS_SHEET);
  const results = [];
  if (resultsSheet.getLastRow() > 1) {
    resultsSheet.getRange(2, 1, resultsSheet.getLastRow() - 1, RESULT_HEADERS.length).getValues().forEach(row => {
      if (String(row[1]) !== projectId) return;
      results.push({
        id: String(row[0]), institution: String(row[3] || ""), category: String(row[4] || "Perguruan Tinggi"),
        role: String(row[5] || "Kontak umum"), email: String(row[6] || "-"), phone: String(row[7] || "-"),
        address: String(row[8] || "-"), region: String(row[9] || "Indonesia"), website: String(row[10] || ""),
        sourceUrl: String(row[11] || ""), evidence: String(row[12] || "Crawl publik"), confidence: Number(row[13] || 0),
        readiness: String(row[14] || "Perlu Verifikasi"), status: String(row[15] || "Perlu Review"), source: "Google Sheets crawler"
      });
    });
  }
  return { ok: true, project: project, results: results };
}

function getProjectRow_(ss, projectId) {
  const sheet = ss.getSheetByName(CRAWLER.PROJECTS_SHEET);
  if (sheet.getLastRow() < 2) throw new Error("Project tidak ditemukan");
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, PROJECT_HEADERS.length).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === projectId) return {
      row: i + 2, projectId: String(rows[i][0]), name: String(rows[i][2]), focus: String(rows[i][3]),
      region: String(rows[i][4]), category: String(rows[i][5]), roleFocus: String(rows[i][6]), mode: String(rows[i][7]),
      sourceUrls: String(rows[i][8]), status: String(rows[i][12]), totalUrls: Number(rows[i][13] || 0),
      processedUrls: Number(rows[i][14] || 0), resultCount: Number(rows[i][15] || 0), error: String(rows[i][17] || "")
    };
  }
  throw new Error("Project tidak ditemukan: " + projectId);
}

function updateProjectSummary_(ss, projectId) {
  let project;
  try { project = getProjectRow_(ss, projectId); } catch (error) { return; }
  const queueSheet = ss.getSheetByName(CRAWLER.QUEUE_SHEET);
  const resultSheet = ss.getSheetByName(CRAWLER.RESULTS_SHEET);
  let total = 0, processed = 0, pending = 0, errors = 0, lastError = "";
  if (queueSheet.getLastRow() > 1) {
    queueSheet.getRange(2, 1, queueSheet.getLastRow() - 1, QUEUE_HEADERS.length).getValues().forEach(row => {
      if (String(row[1]) !== projectId) return;
      total++;
      const status = String(row[3]);
      if (status === "Done" || status === "Error") processed++;
      if (["Queued", "Retry", "Running"].includes(status)) pending++;
      if (status === "Error") { errors++; lastError = String(row[8] || ""); }
    });
  }
  let resultCount = 0;
  if (resultSheet.getLastRow() > 1) {
    resultSheet.getRange(2, 2, resultSheet.getLastRow() - 1, 1).getValues().forEach(row => { if (String(row[0]) === projectId) resultCount++; });
  }
  const status = pending > 0 ? "Sedang mencari" : resultCount > 0 ? "Data tersedia" : errors === total && total > 0 ? "Error" : "No Results";
  const sheet = ss.getSheetByName(CRAWLER.PROJECTS_SHEET);
  sheet.getRange(project.row, 13, 1, 6).setValues([[status, total, processed, resultCount, new Date(), lastError]]);
}

function ensureCrawlerSheets_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureSheet_(ss, CRAWLER.PROJECTS_SHEET, PROJECT_HEADERS, "#0f6b68");
  ensureSheet_(ss, CRAWLER.QUEUE_SHEET, QUEUE_HEADERS, "#d58b31");
  ensureSheet_(ss, CRAWLER.RESULTS_SHEET, RESULT_HEADERS, "#377a9b");
  ensureSheet_(ss, CRAWLER.LOG_SHEET, ["Timestamp", "Project", "Region", "Category", "Keywords", "PIC Name", "PIC Contact", "PIC Role", "Status"], "#71818b");
}

function ensureSheet_(ss, name, headers, color) {
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground(color).setFontColor("#ffffff").setFontWeight("bold").setWrap(true);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

function isAllowedByRobots_(url) {
  const root = rootUrl_(url);
  const cache = CacheService.getScriptCache();
  let robots = cache.get("robots:" + root);
  if (robots === null) {
    try {
      const response = UrlFetchApp.fetch(root + "/robots.txt", { muteHttpExceptions: true, headers: { "User-Agent": CRAWLER.USER_AGENT } });
      robots = response.getResponseCode() === 200 ? response.getContentText().slice(0, 50000) : "";
    } catch (error) { robots = ""; }
    cache.put("robots:" + root, robots, 21600);
  }
  if (!robots) return true;
  const path = pathOf_(url);
  let applies = false;
  const lines = robots.split(/\r?\n/);
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].replace(/#.*/, "").trim();
    if (/^user-agent\s*:/i.test(line)) applies = /^user-agent\s*:\s*\*/i.test(line);
    else if (applies && /^disallow\s*:/i.test(line)) {
      const denied = line.replace(/^disallow\s*:\s*/i, "").trim();
      if (denied && path.indexOf(denied) === 0) return false;
    }
  }
  return true;
}

function normalizeUrl_(value) {
  let url = clean_(value, 2000);
  if (!url) return "";
  if (!/^https?:\/\//i.test(url)) url = "https://" + url;
  return url.replace(/#.*$/, "");
}
function isSafePublicUrl_(url) {
  if (!/^https?:\/\//i.test(url) || url.length > 2000 || /@/.test(url.replace(/^https?:\/\//i, "").split("/")[0])) return false;
  const host = hostOf_(url);
  if (!host || host === "localhost" || /\.local$/i.test(host) || /^\[/.test(host)) return false;
  if (/^(10\.|127\.|169\.254\.|192\.168\.|0\.)/.test(host)) return false;
  const private172 = host.match(/^172\.(\d+)\./);
  if (private172 && Number(private172[1]) >= 16 && Number(private172[1]) <= 31) return false;
  return true;
}
function resolveUrl_(href, base) {
  if (/^https?:\/\//i.test(href)) return href.replace(/#.*$/, "");
  if (/^(mailto:|tel:|javascript:|data:)/i.test(href)) return "";
  const root = rootUrl_(base);
  if (href.charAt(0) === "/") return root + href.replace(/#.*$/, "");
  const directory = base.replace(/[?#].*$/, "").replace(/\/[^/]*$/, "/");
  return (directory + href).replace(/#.*$/, "");
}
function sameHost_(a, b) { return hostOf_(a) === hostOf_(b); }
function hostOf_(url) { const match = String(url).match(/^https?:\/\/([^/:?#]+)/i); return match ? match[1].toLowerCase().replace(/^www\./, "") : ""; }
function rootUrl_(url) { const match = String(url).match(/^(https?):\/\/([^/:?#]+)/i); return match ? match[1].toLowerCase() + "://" + match[2] : ""; }
function pathOf_(url) { const match = String(url).match(/^https?:\/\/[^/]+(\/[^?#]*)?/i); return match && match[1] ? match[1] : "/"; }
function normalizePhone_(value) { return String(value).replace(/[^\d+]/g, "").replace(/^620/, "62"); }
function unique_(values) { const seen = {}; return values.filter(value => { const key = String(value).toLowerCase(); if (!key || seen[key]) return false; seen[key] = true; return true; }); }
function clean_(value, max) { return String(value == null ? "" : value).replace(/[\u0000-\u001f\u007f]/g, " ").replace(/\s+/g, " ").trim().slice(0, max || 500); }
function cleanText_(value) { return clean_(String(value || "").replace(/&nbsp;|&#160;/gi, " ").replace(/&amp;/gi, "&").replace(/&quot;/gi, '"').replace(/&#39;|&apos;/gi, "'"), 200000); }
function inferRole_(text, requested) {
  const candidates = ["Kerja Sama", "Humas", "Hubungan Masyarakat", "Akademik", "Penerimaan Mahasiswa", "Administrasi", "Rektorat"];
  for (let i = 0; i < candidates.length; i++) if (new RegExp(candidates[i].replace(" ", "\\s*"), "i").test(text)) return candidates[i];
  return clean_(requested, 120) || "Kontak umum";
}
function inferCategory_(text, requested) {
  if (requested && requested !== "Semua") return requested;
  if (/politeknik/i.test(text)) return "Politeknik";
  if (/institut/i.test(text)) return "Institut";
  if (/sekolah tinggi|stikes|stkip/i.test(text)) return "Sekolah Tinggi";
  if (/universitas|university/i.test(text)) return "Universitas";
  return "Perguruan Tinggi";
}
function extractAddress_(text) {
  const match = text.match(/(?:alamat|address)\s*[:\-]?\s*([^|]{15,180}?)(?=(?:telepon|telp|phone|email|fax|$))/i);
  return match ? clean_(match[1], 220) : "";
}
function output_(data, callback) {
  const json = JSON.stringify(data);
  const safeCallback = /^[A-Za-z_$][\w$]*(?:\.[A-Za-z_$][\w$]*)*$/.test(String(callback || "")) ? String(callback) : "";
  if (safeCallback) return ContentService.createTextOutput(safeCallback + "(" + json + ");").setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
