const https = require('https');
const http = require('http');

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

  try {
    const body = req.body || {};
    const action = body.action || '';

    let result;
    // 照片上傳用 POST（base64 太大無法放 URL）
    // 其他請求用 GET（避免 GAS POST 302 跳轉資料丟失的問題）
    if (action === 'uploadPhoto') {
      result = await httpPost(GAS_URL, JSON.stringify(body), 0);
    } else {
      const encodedData = encodeURIComponent(JSON.stringify(body));
      const url = GAS_URL + '?data=' + encodedData;
      result = await httpGet(url, 0);
    }

    res.json(result);
  } catch (err) {
    res.json({ success: false, message: err.message });
  }
};

// GET 請求（跟隨重定向）
function httpGet(url, depth) {
  return new Promise((resolve, reject) => {
    if (depth > 5) { reject(new Error('Too many redirects')); return; }
    const urlObj = new URL(url);
    const lib = url.startsWith('https') ? https : http;
    const options = {
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      method: 'GET',
      headers: { 'Accept': 'application/json' }
    };
    const req = lib.request(options, (response) => {
      if ([301,302,303,307,308].includes(response.statusCode) && response.headers.location) {
        response.resume();
        httpGet(response.headers.location, depth + 1).then(resolve).catch(reject);
        return;
      }
      let data = '';
      response.on('data', chunk => { data += chunk; });
      response.on('end', () => {
        const trimmed = data.trim();
        if (trimmed.startsWith('{') || trimmed.startsWith('[')) {
          try { resolve(JSON.parse(trimmed)); }
          catch (e) { reject(new Error('JSON解析失敗：' + trimmed.substring(0, 100))); }
        } else {
          reject(new Error('GAS回傳非JSON：' + trimmed.substring(0, 100)));
        }
      });
    });
    req.on('error', reject);
    req.setTimeout(30000, () => { req.destroy(); reject(new Error('請求逾時')); });
    req.end();
  });
}

// POST 請求（跟隨重定向，保留 body）
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
        'Accept': 'application/json'
      }
    };
    const req = lib.request(options, (response) => {
      if ([301,302,303,307,308].includes(response.statusCode) && response.headers.location) {
        response.resume();
        httpPost(response.headers.location, body, depth + 1).then(resolve).catch(reject);
        return;
      }
      let data = '';
      response.on('data', chunk => { data += chunk; });
      response.on('end', () => {
        const trimmed = data.trim();
        if (trimmed.startsWith('{') || trimmed.startsWith('[')) {
          try { resolve(JSON.parse(trimmed)); }
          catch (e) { reject(new Error('JSON解析失敗：' + trimmed.substring(0, 100))); }
        } else {
          reject(new Error('GAS回傳非JSON：' + trimmed.substring(0, 100)));
        }
      });
    });
    req.on('error', reject);
    req.setTimeout(60000, () => { req.destroy(); reject(new Error('照片上傳逾時')); });
    req.write(body);
    req.end();
  });
}
