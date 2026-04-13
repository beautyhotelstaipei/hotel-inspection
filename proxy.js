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
    res.status(500).json({ success: false, message: '未設定 GAS_URL 環境變數' });
    return;
  }

  try {
    const body = req.body || {};
    // 用 GET + data 參數呼叫 GAS，避免 POST 302 跳轉後資料丟失
    const encodedData = encodeURIComponent(JSON.stringify(body));
    const url = GAS_URL + '?data=' + encodedData;

    const response = await fetch(url);
    const text = await response.text();

    try {
      const result = JSON.parse(text);
      res.status(200).json(result);
    } catch (parseErr) {
      res.status(500).json({
        success: false,
        message: 'GAS 回傳格式錯誤：' + text.substring(0, 150)
      });
    }
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
};
