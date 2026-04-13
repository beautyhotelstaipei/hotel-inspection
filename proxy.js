const https = require('https');
const http = require('http');
const { URL } = require('url');

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  const GAS_URL = process.env.GAS_URL;
  if (!GAS_URL) {
    res.status(500).json({ success: false, message: '未設定 GAS_URL 環境變數' });
    return;
  }

  try {
    const body = JSON.stringify(req.body || {});
    const result = await fetchWithRedirect(GAS_URL, body);
    res.status(200).json(result);
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
};

function fetchWithRedirect(url, body, redirectCount = 0) {
  return new Promise((resolve, reject) => {
    if (redirectCount > 5) {
      reject(new Error('Too many redirects'));
      return;
    }

    const parsedUrl = new URL(url);
    const options = {
      hostname: parsedUrl.hostname,
      path: parsedUrl.pathname + parsedUrl.search,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(body)
      }
    };

    const lib = parsedUrl.protocol === 'https:' ? https : http;
    const reqObj = lib.request(options, (response) => {
      if (response.statusCode >= 300 && response.statusCode < 400 && response.headers.location) {
        fetchWithRedirect(response.headers.location, body, redirectCount + 1)
          .then(resolve).catch(reject);
        return;
      }

      let data = '';
      response.on('data', chunk => { data += chunk; });
      response.on('end', () => {
        try {
          resolve(JSON.parse(data));
        } catch (e) {
          reject(new Error('Invalid JSON response: ' + data.substring(0, 200)));
        }
      });
    });

    reqObj.on('error', reject);
    reqObj.write(body);
    reqObj.end();
  });
}
