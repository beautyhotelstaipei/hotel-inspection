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
    const body = JSON.stringify(req.body || {});
    const encodedData = encodeURIComponent(body);
    const url = GAS_URL + '?data=' + encodedData;

    const result = await httpGet(url, 0);
    res.json(result);
  } catch (err) {
    res.json({ success: false, message: err.message });
  }
};

function httpGet(url, depth) {
  return new Promise((resolve, reject) => {
    if (depth > 5) { reject(new Error('Too many redirects')); return; }

    const isHttps = url.startsWith('https');
    const lib = isHttps ? https : http;
    const urlObj = new URL(url);

    const options = {
      hostname: urlObj.hostname,
      path: urlObj.pathname + urlObj.search,
      method: 'GET',
      headers: { 'Accept': 'application/json' }
    };

    const req = lib.request(options, (response) => {
      const status = response.statusCode;

      if ([301, 302, 303, 307, 308].includes(status) && response.headers.location) {
        response.resume();
        httpGet(response.headers.location, depth + 1).then(resolve).catch(reject);
        return;
      }

      let data = '';
      response.on('data', chunk => { data += chunk; });
      response.on('end', () => {
        const trimmed = data.trim();
        if (trimmed.startsWith('{') || trimmed.startsWith('[')) {
          try {
            resolve(JSON.parse(trimmed));
          } catch (e) {
            reject(new Error('JSON解析失敗：' + trimmed.substring(0, 100)));
          }
        } else {
          reject(new Error('GAS回傳非JSON：' + trimmed.substring(0, 100)));
        }
      });
    });

    req.on('error', reject);
    req.setTimeout(30000, () => {
      req.destroy();
      reject(new Error('請求逾時'));
    });
    req.end();
  });
}
