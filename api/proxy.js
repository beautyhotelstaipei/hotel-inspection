const https = require('https');
const http = require('http');

// ============================================================
// 設定
// ============================================================

// 身分標示：不帶這個的請求容易被 Google 防濫用機制判定為機器人，
// 直接回傳攔截頁（HTML）而不是資料。帶上正常瀏覽器的身分可降低被攔機率。
const USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36';

// 唯讀操作：重試不會造成重複寫入，任何失敗都可以安全重試
const READ_ONLY_ACTIONS = [
  'login', 'getRefData', 'getDashboard',
  'getInspections', 'getAnomalies', 'getStatistics', 'getArchive'
];

const MAX_ATTEMPTS = 3;        // 最多嘗試次數（含第一次）
const TOTAL_BUDGET_MS = 20000; // 所有嘗試加起來的時間上限，避免超過 Vercel 執行上限

const sleep = (ms) => new Promise(r => setTimeout(r, ms));

// ============================================================
// 主流程
// ============================================================
module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  const GAS_URL = process.env.GAS_URL;
  if (!GAS_URL) {
    res.json({ success: false, message: '未設定 GAS_URL 環境變數' });
    return;
  }

  const body = req.body || {};
  const action = body.action || '';
  const isReadOnly = READ_ONLY_ACTIONS.includes(action);

  const started = Date.now();
  let lastErr = null;

  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    try {
      let result;
      // 照片上傳用 POST（base64 太大無法放 URL）
      // 其他請求用 GET（避免 GAS POST 302 跳轉資料丟失的問題）
      if (action === 'uploadPhoto') {
        result = await httpPost(GAS_URL, JSON.stringify(body), 0);
      } else {
        const encodedData = encodeURIComponent(JSON.stringify(body));
        result = await httpGet(GAS_URL + '?data=' + encodedData, 0);
      }
      if (attempt > 1) {
        console.log('[proxy] action=' + action + ' 第 ' + attempt + ' 次嘗試成功');
      }
      res.json(result);
      return;
    } catch (err) {
      lastErr = err;

      // 重要：能不能重試，取決於 GAS 是否可能已經執行過。
      // 唯讀操作 → 永遠可以重試。
      // 寫入操作 → 只有在「確定 GAS 還沒跑到」時才能重試，
      //            否則重試會造成重複的巡房紀錄 / 重複的照片。
      const canRetry = isReadOnly || !err.gasMayHaveRun;

      console.error('[proxy] action=' + action + ' 第 ' + attempt + ' 次失敗'
        + '（可重試=' + canRetry + '）：' + err.message);

      if (!canRetry) break;
      if (attempt >= MAX_ATTEMPTS) break;
      if (Date.now() - started > TOTAL_BUDGET_MS) break;

      await sleep(300 * attempt); // 300ms、600ms 遞增等待
    }
  }

  // 全部嘗試都失敗：回友善訊息給使用者，技術細節另外放 detail 供除錯
  res.json({
    success: false,
    message: '連線不穩定，請稍後再試一次',
    detail: lastErr ? lastErr.message : '未知錯誤',
    // 前端據此判斷：true = GAS 可能已寫入，不可盲目重送
    mayHaveWritten: !!(lastErr && lastErr.gasMayHaveRun)
  });
};

// ============================================================
// GET 請求（跟隨重定向）
// ============================================================
function httpGet(url, depth) {
  return new Promise((resolve, reject) => {
    if (depth > 5) { reject(new Error('Too many redirects')); return; }
    const urlObj = new URL(url);
    const lib = url.startsWith('https') ? https : http;
    const options = {
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      method: 'GET',
      headers: {
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'zh-TW,zh;q=0.9,en;q=0.8',
        'User-Agent': USER_AGENT
      }
    };
    const req = lib.request(options, (response) => {
      if ([301,302,303,307,308].includes(response.statusCode) && response.headers.location) {
        response.resume();
        // 已經跟著跳轉 = GAS 那邊的程式已經跑過了，
        // 之後若再失敗，寫入類操作就不能重試
        httpGet(response.headers.location, depth + 1)
          .then(resolve)
          .catch(e => { e.gasMayHaveRun = true; reject(e); });
        return;
      }
      let data = '';
      response.on('data', chunk => { data += chunk; });
      response.on('end', () => {
        const trimmed = data.trim();
        if (trimmed.startsWith('{') || trimmed.startsWith('[')) {
          try { resolve(JSON.parse(trimmed)); }
          catch (e) { reject(new Error('JSON解析失敗：' + trimmed.substring(0, 200))); }
        } else {
          reject(new Error('GAS回傳非JSON（HTTP ' + response.statusCode + '）：'
            + trimmed.substring(0, 200)));
        }
      });
    });
    req.on('error', reject);
    req.setTimeout(15000, () => {
      req.destroy();
      // 逾時代表請求已送出，GAS 可能正在執行中
      const e = new Error('請求逾時');
      e.gasMayHaveRun = true;
      reject(e);
    });
    req.end();
  });
}

// ============================================================
// POST 請求（照片上傳用，跟隨重定向並保留 body）
// ============================================================
function httpPost(url, body, depth) {
  return new Promise((resolve, reject) => {
    if (depth > 5) { reject(new Error('Too many redirects')); return; }
    const urlObj = new URL(url);
    const lib = url.startsWith('https') ? https : http;
    const options = {
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(body),
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'zh-TW,zh;q=0.9,en;q=0.8',
        'User-Agent': USER_AGENT
      }
    };
    const req = lib.request(options, (response) => {
      if ([301,302,303,307,308].includes(response.statusCode) && response.headers.location) {
        response.resume();
        httpPost(response.headers.location, body, depth + 1)
          .then(resolve)
          .catch(e => { e.gasMayHaveRun = true; reject(e); });
        return;
      }
      let data = '';
      response.on('data', chunk => { data += chunk; });
      response.on('end', () => {
        const trimmed = data.trim();
        if (trimmed.startsWith('{') || trimmed.startsWith('[')) {
          try { resolve(JSON.parse(trimmed)); }
          catch (e) { reject(new Error('JSON解析失敗：' + trimmed.substring(0, 200))); }
        } else {
          reject(new Error('GAS回傳非JSON（HTTP ' + response.statusCode + '）：'
            + trimmed.substring(0, 200)));
        }
      });
    });
    req.on('error', reject);
    req.setTimeout(60000, () => {
      req.destroy();
      const e = new Error('照片上傳逾時');
      e.gasMayHaveRun = true;
      reject(e);
    });
    req.write(body);
    req.end();
  });
}
