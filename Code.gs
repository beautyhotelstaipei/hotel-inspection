// ============================================================
// 飯店巡房系統 v2 - 後端 API Code.gs
// ============================================================

const SPREADSHEET_ID = '1IvbprKoesY2eiINISgHfxGy-BrIBp4b7GworYuUWLvM';

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ============================================================
// API 入口
// ============================================================
function doPost(e) {
  let params = {};
  try {
    params = JSON.parse(e.postData.contents);
  } catch (err) {
    params = e.parameter || {};
  }
  return handleRequest(params);
}

function doGet(e) {
  return handleRequest(e.parameter || {});
}

function handleRequest(params) {
  const action = params.action || '';
  let result;
  try {
    switch (action) {
      case 'login': result = login(params.account, params.password); break;
      case 'getRefData': result = getReferenceData(params.scope); break;
      case 'getDashboard': result = getDashboardData(params.scope); break;
      case 'submitInspection': result = submitInspection(typeof params.data === 'string' ? JSON.parse(params.data) : params.data); break;
      case 'getInspections': result = getInspectionRecords(typeof params.filters === 'string' ? JSON.parse(params.filters) : params.filters); break;
      case 'getAnomalies': result = getAnomalyRecords(typeof params.filters === 'string' ? JSON.parse(params.filters) : params.filters); break;
      case 'updateAnomalyStatus': result = updateAnomalyStatus(params.anomalyId, params.newStatus, params.updaterName); break;
      case 'updateInspection': result = updateInspectionRecord(params.recordId, typeof params.data === 'string' ? JSON.parse(params.data) : params.data, params.updaterName); break;
      case 'getStatistics': result = getStatistics(typeof params.filters === 'string' ? JSON.parse(params.filters) : params.filters); break;
      case 'getArchive': result = getArchiveData(typeof params.filters === 'string' ? JSON.parse(params.filters) : params.filters); break;
      case 'uploadPhoto': result = { success: false, message: '請在系統設定填入 Google Drive 資料夾 ID 後啟用此功能' }; break;
      default: result = { success: false, message: '未知的操作：' + action };
    }
  } catch (err) {
    result = { success: false, message: err.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// 系統設定
// ============================================================
function getSystemSettings() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('系統設定');
    const defaults = {
      '系統名稱': '飯店巡房系統',
      '主題顏色': '#1a1a2e',
      '字體大小': '16',
      '每頁顯示筆數': '20',
      'Line Notify Token': '',
      '通知Email': '',
      '寄送頻率': '',
      '照片資料夾ID': '',
      '封存試算表ID': '',
      '顯示時效統計': '否'
    };
    if (!sheet) return defaults;
    const data = sheet.getDataRange().getValues();
    const settings = Object.assign({}, defaults);
    data.slice(1).forEach(row => {
      if (row[0] && row[1] !== undefined && row[1] !== '') {
        settings[String(row[0]).trim()] = String(row[1]).trim();
      }
    });
    return settings;
  } catch (e) {
    return { '系統名稱': '飯店巡房系統', '主題顏色': '#1a1a2e', '字體大小': '16', '每頁顯示筆數': '20' };
  }
}

// ============================================================
// 自訂欄位
// ============================================================
function getCustomFields() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('自訂欄位');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    return data.slice(1)
      .filter(row => row[0] && String(row[5]).trim() === '是')
      .map(row => ({
        name: String(row[0]).trim(),
        type: String(row[1]).trim(),
        required: String(row[2]).trim() === '是',
        positions: String(row[3]).trim().split(',').map(s => s.trim()),
        options: row[4] ? String(row[4]).trim().split(',').map(s => s.trim()) : []
      }));
  } catch (e) {
    return [];
  }
}

// ============================================================
// 登入
// ============================================================
function login(account, password) {
  try {
    const ss = getSpreadsheet();
    const data = ss.getSheetByName('人員清單').getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[1]).toLowerCase().trim() === String(account).toLowerCase().trim() &&
        String(row[2]).trim() === String(password).trim() &&
        String(row[5]).trim() === '是') {
        return {
          success: true,
          user: {
            name: row[0], account: row[1], title: row[3],
            scope: String(row[4]).trim(),
            visibility: String(row[6] || '一般').trim()
          }
        };
      }
    }
    return { success: false, message: '帳號或密碼錯誤' };
  } catch (e) {
    return { success: false, message: '系統錯誤：' + e.toString() };
  }
}

// ============================================================
// 權限
// ============================================================
function hasAccess(scope, area, locationName) {
  if (!scope) return false;
  scope = String(scope).trim();
  if (scope === '全部') return true;
  const scopes = scope.split(',').map(s => s.trim());
  if (area && scopes.includes(String(area).trim())) return true;
  if (locationName && scopes.includes(String(locationName).trim())) return true;
  return false;
}

function canSeeInspector(viewerScope, inspectorVisibility) {
  if (viewerScope === '全部') return true;
  if (String(inspectorVisibility).trim() === '限高層') return false;
  return true;
}

function getVisibilityMap() {
  const data = getSpreadsheet().getSheetByName('人員清單').getDataRange().getValues();
  const map = {};
  data.slice(1).forEach(row => {
    if (row[0]) map[String(row[0]).trim()] = String(row[6] || '一般').trim();
  });
  return map;
}

function getFilteredStaff(viewerScope) {
  const data = getSpreadsheet().getSheetByName('人員清單').getDataRange().getValues();
  return data.slice(1).filter(row => {
    if (String(row[5]).trim() !== '是') return false;
    if (!canSeeInspector(viewerScope, String(row[6] || '一般').trim())) return false;
    if (viewerScope === '全部') return true;
    const staffScope = String(row[4]).trim();
    if (staffScope === '全部') return false;
    const vs = viewerScope.split(',').map(s => s.trim());
    const ss = staffScope.split(',').map(s => s.trim());
    return vs.some(v => ss.some(s => s === v));
  }).map(row => ({ name: row[0], account: row[1], title: row[3], scope: row[4] }));
}

// ============================================================
// 結案判斷
// ============================================================
function isResolved(type, status) {
  if (!status) return false;
  const s = String(status).split('（')[0].trim();
  const resolved = {
    '維修': ['巡查人員處理完成', '現場其他人員處理完成', '已完修'],
    '清潔': ['巡查人員處理完成', '房務員處理完成', '已完成'],
    '備品': ['巡查人員補齊', '已補齊'],
    '其他': ['巡查人員處理完成', '已完成']
  };
  return (resolved[type] || []).some(r => s.startsWith(r));
}

// ============================================================
// 取得參考資料
// ============================================================
function getReferenceData(userScope) {
  try {
    const ss = getSpreadsheet();
    const result = {};

    result['據點清單'] = ss.getSheetByName('據點清單').getDataRange().getValues().slice(1)
      .filter(row => String(row[4]).trim() === '是' && hasAccess(userScope, row[0], row[2]));

    ['巡視區域', '嚴重度', '房間狀態', '問題類型'].forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (sheet) result[name] = sheet.getDataRange().getValues().slice(1)
        .filter(row => String(row[row.length - 1]).trim() === '是');
    });

    ['維修狀態', '清潔狀態', '備品狀態', '其他狀態'].forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (sheet) {
        result[name] = sheet.getDataRange().getValues().slice(1)
          .filter(row => String(row[1]).trim() === '是')
          .map(row => ({
            name: String(row[0]).trim(),
            type: String(row[2] || '').trim(),
            color: String(row[3] || '').trim()
          }));
      } else { result[name] = []; }
    });

    result['人員清單'] = getFilteredStaff(userScope);
    result['系統設定'] = getSystemSettings();
    result['自訂欄位'] = getCustomFields();
    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// 今日儀表板
// ============================================================
function getDashboardData(userScope) {
  try {
    const ss = getSpreadsheet();
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const inspData = ss.getSheetByName('巡房紀錄').getDataRange().getValues();
    const anomalyData = ss.getSheetByName('異常紀錄').getDataRange().getValues();
    const visMap = getVisibilityMap();

    const todayInsp = inspData.slice(1).filter(row => {
      if (!row[0]) return false;
      const d = new Date(row[7]); d.setHours(0, 0, 0, 0);
      return d.getTime() === today.getTime()
        && hasAccess(userScope, row[2], row[4])
        && canSeeInspector(userScope, visMap[String(row[1]).trim()]);
    });

    const todayAnomalies = anomalyData.slice(1).filter(row => {
      if (!row[0]) return false;
      const d = new Date(row[11]); d.setHours(0, 0, 0, 0);
      return d.getTime() === today.getTime()
        && hasAccess(userScope, null, row[3])
        && canSeeInspector(userScope, visMap[String(row[2]).trim()]);
    });

    const pendingCount = anomalyData.slice(1).filter(row =>
      row[0] && !isResolved(row[7], row[9])
      && hasAccess(userScope, null, row[3])
      && canSeeInspector(userScope, visMap[String(row[2]).trim()])
    ).length;

    // 排名
    const rankMap = {};
    todayInsp.forEach(row => {
      const inspector = String(row[1]).trim();
      if (!rankMap[inspector]) rankMap[inspector] = new Set();
      rankMap[inspector].add(String(row[5]).toLowerCase());
    });
    const ranking = Object.entries(rankMap)
      .map(([name, rooms]) => ({ name, count: rooms.size }))
      .sort((a, b) => b.count - a.count).slice(0, 5);

    const locData = ss.getSheetByName('據點清單').getDataRange().getValues();
    const locationStats = locData.slice(1)
      .filter(loc => String(loc[4]).trim() === '是' && hasAccess(userScope, loc[0], loc[2]))
      .map(loc => {
        const locInsp = todayInsp.filter(r => r[4] === loc[2]);
        return {
          area: loc[0], type: loc[1], name: loc[2], totalRooms: loc[3],
          inspCount: locInsp.length,
          uniqueRooms: [...new Set(locInsp.map(r => String(r[5]).toLowerCase()))].length,
          anomalyCount: todayAnomalies.filter(r => r[3] === loc[2]).length
        };
      });

    return {
      success: true,
      data: {
        totalInspections: todayInsp.length,
        totalUniqueRooms: [...new Set(todayInsp.map(r => r[4] + '-' + String(r[5]).toLowerCase()))].length,
        totalAnomalies: todayAnomalies.length,
        pendingCount, ranking, locationStats
      }
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// 送出巡房紀錄
// ============================================================
function submitInspection(data) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('巡房紀錄');
    const anomalySheet = ss.getSheetByName('異常紀錄');
    const now = new Date();
    const startTime = data.startTime ? new Date(data.startTime) : now;
    const endTime = data.endTime ? new Date(data.endTime) : now;
    const duration = Math.max(0, Math.round((endTime - startTime) / 60000));
    const recordId = 'R' + Utilities.formatDate(now, 'Asia/Taipei', 'yyyyMMddHHmmss');

    // 基本欄位 + 自訂欄位
    const baseRow = [recordId, data.inspector, data.area, data.type, data.location,
      data.roomNumber, data.roomStatus, startTime, endTime, duration + '分鐘',
      data.generalNote || '', now, data.fillType || '巡房中填寫'];

    // 加入自訂欄位值
    const customFields = getCustomFields().filter(f => f.positions.includes('巡房'));
    customFields.forEach(f => {
      baseRow.push(data.customFields ? (data.customFields[f.name] || '') : '');
    });

    sheet.appendRow(baseRow);

    // 一般異常
    if (data.anomalies && data.anomalies.length > 0) {
      data.anomalies.forEach((anomaly, idx) => {
        const anomalyId = 'A' + Utilities.formatDate(now, 'Asia/Taipei', 'yyyyMMddHHmmss') + String(idx).padStart(2, '0');
        anomalySheet.appendRow([anomalyId, recordId, data.inspector, data.location,
          data.roomNumber, anomaly.area, anomaly.severity, anomaly.type,
          anomaly.note || '', anomaly.status || '', anomaly.status || '', now, '', '']);
      });
    }

    // 整體備註追蹤異常
    if (data.noteAnomalies && data.noteAnomalies.length > 0) {
      data.noteAnomalies.forEach((anomaly, idx) => {
        const anomalyId = 'AN' + Utilities.formatDate(now, 'Asia/Taipei', 'yyyyMMddHHmmss') + String(idx).padStart(2, '0');
        anomalySheet.appendRow([anomalyId, recordId, data.inspector, data.location,
          data.roomNumber, '整體備註', anomaly.severity, anomaly.type,
          anomaly.note || data.generalNote || '', anomaly.status || '', anomaly.status || '', now, '', '']);
      });
    }

    // Line 通知
    const settings = getSystemSettings();
    if (settings['Line Notify Token'] && data.anomalies && data.anomalies.some(a => a.severity === '嚴重')) {
      sendLineNotify(settings['Line Notify Token'], data, data.anomalies.filter(a => a.severity === '嚴重'));
    }

    return { success: true, recordId };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// Line 通知
// ============================================================
function sendLineNotify(token, inspData, anomalies) {
  try {
    const msg = '\n🚨 嚴重異常通知\n'
      + '館名：' + inspData.location + '\n'
      + '房號：' + inspData.roomNumber + '\n'
      + '巡房者：' + inspData.inspector + '\n'
      + '異常：' + anomalies.map(a => a.inspArea + '(' + a.type + ')').join('、');

    UrlFetchApp.fetch('https://notify-api.line.me/api/notify', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: { message: msg }
    });
  } catch (e) {
    Logger.log('Line通知失敗：' + e.toString());
  }
}

// ============================================================
// 查詢巡房紀錄
// ============================================================
function getInspectionRecords(filters) {
  try {
    const ss = getSpreadsheet();
    const inspData = ss.getSheetByName('巡房紀錄').getDataRange().getValues();
    const anomalyData = ss.getSheetByName('異常紀錄').getDataRange().getValues();
    const modData = ss.getSheetByName('修改紀錄').getDataRange().getValues();
    const visMap = getVisibilityMap();
    const pageSize = parseInt(filters.pageSize) || 20;
    const customFields = getCustomFields().filter(f => f.positions.includes('巡房'));

    const results = inspData.slice(1).filter(row => {
      if (!row[0]) return false;
      if (!hasAccess(filters.userScope, row[2], row[4])) return false;
      if (!canSeeInspector(filters.userScope, visMap[String(row[1]).trim()])) return false;
      const startDate = row[7] ? new Date(row[7]) : null;
      if (startDate) {
        if (filters.startDate) { const s = new Date(filters.startDate); s.setHours(0, 0, 0, 0); if (startDate < s) return false; }
        if (filters.endDate) { const e = new Date(filters.endDate); e.setHours(23, 59, 59, 999); if (startDate > e) return false; }
      }
      if (filters.location && filters.location !== '全部' && row[4] !== filters.location) return false;
      if (filters.roomNumber && String(row[5]).toLowerCase() !== String(filters.roomNumber).toLowerCase()) return false;
      if (filters.inspector && filters.inspector !== '全部' && row[1] !== filters.inspector) return false;
      // 自訂欄位篩選
      if (filters.customFilters) {
        for (const [fieldName, value] of Object.entries(filters.customFilters)) {
          if (!value) continue;
          const fieldIdx = customFields.findIndex(f => f.name === fieldName);
          if (fieldIdx >= 0 && String(row[13 + fieldIdx]).toLowerCase().indexOf(String(value).toLowerCase()) < 0) return false;
        }
      }
      const anomalies = anomalyData.slice(1).filter(a => a[1] === row[0]);
      if (filters.onlyAnomalies && anomalies.length === 0) return false;
      return true;
    }).map(row => {
      const anomalies = anomalyData.slice(1).filter(a => a[1] === row[0]).map(a => ({
        anomalyId: a[0], inspArea: a[5], severity: a[6], type: a[7], note: a[8], status: a[9]
      }));
      const modifications = modData.slice(1).filter(m => m[1] === row[0]).map(m => ({
        modifier: m[2],
        modTime: m[3] ? Utilities.formatDate(new Date(m[3]), 'Asia/Taipei', 'yyyy-MM-dd HH:mm') : '',
        modNote: m[4]
      }));
      // 自訂欄位值
      const customFieldValues = {};
      customFields.forEach((f, idx) => { customFieldValues[f.name] = row[13 + idx] || ''; });

      return {
        recordId: row[0], inspector: row[1], area: row[2], type: row[3],
        location: row[4], roomNumber: row[5], roomStatus: row[6],
        startTime: row[7] ? Utilities.formatDate(new Date(row[7]), 'Asia/Taipei', 'yyyy-MM-dd HH:mm') : '',
        endTime: row[8] ? Utilities.formatDate(new Date(row[8]), 'Asia/Taipei', 'yyyy-MM-dd HH:mm') : '',
        duration: row[9], generalNote: row[10],
        createdAt: row[11] ? Utilities.formatDate(new Date(row[11]), 'Asia/Taipei', 'yyyy-MM-dd HH:mm') : '',
        fillType: row[12] || '巡房中填寫',
        customFields: customFieldValues,
        anomalies, modifications
      };
    }).reverse();

    return { success: true, data: results.slice(0, pageSize), total: results.length };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// 查詢異常追蹤
// ============================================================
function getAnomalyRecords(filters) {
  try {
    const ss = getSpreadsheet();
    const data = ss.getSheetByName('異常紀錄').getDataRange().getValues();
    const locData = ss.getSheetByName('據點清單').getDataRange().getValues();
    const visMap = getVisibilityMap();
    const locAreaMap = {};
    locData.slice(1).forEach(row => { locAreaMap[row[2]] = row[0]; });
    const pageSize = parseInt(filters.pageSize) || 20;
    const settings = getSystemSettings();

    const results = data.slice(1).filter(row => {
      if (!row[0]) return false;
      const locName = row[3]; const locArea = locAreaMap[locName] || '';
      if (!hasAccess(filters.userScope, locArea, locName)) return false;
      if (!canSeeInspector(filters.userScope, visMap[String(row[2]).trim()])) return false;
      const d = row[11] ? new Date(row[11]) : null;
      if (d) {
        if (filters.startDate) { const s = new Date(filters.startDate); s.setHours(0, 0, 0, 0); if (d < s) return false; }
        if (filters.endDate) { const e = new Date(filters.endDate); e.setHours(23, 59, 59, 999); if (d > e) return false; }
      }
      if (filters.location && filters.location !== '全部' && locName !== filters.location) return false;
      if (filters.type && filters.type !== '全部' && row[7] !== filters.type) return false;
      if (filters.severity && filters.severity !== '全部' && row[6] !== filters.severity) return false;
      if (filters.roomNumber && String(row[4]).toLowerCase() !== String(filters.roomNumber).toLowerCase()) return false;
      if (filters.onlyPending && isResolved(row[7], row[9])) return false;
      return true;
    }).map(row => {
      const createdAt = row[11] ? new Date(row[11]) : null;
      const updatedAt = row[12] ? new Date(row[12]) : null;
      let timeEfficiency = '';
      if (settings['顯示時效統計'] === '是' && createdAt) {
        const endTime = isResolved(row[7], row[9]) && updatedAt ? updatedAt : new Date();
        const hours = Math.round((endTime - createdAt) / 3600000);
        timeEfficiency = hours < 24 ? hours + '小時' : Math.round(hours / 24) + '天';
      }
      return {
        anomalyId: row[0], recordId: row[1], inspector: row[2],
        location: row[3], roomNumber: row[4], inspArea: row[5],
        severity: row[6], type: row[7], note: row[8], status: row[9],
        createdAt: createdAt ? Utilities.formatDate(createdAt, 'Asia/Taipei', 'yyyy-MM-dd HH:mm') : '',
        updatedAt: updatedAt ? Utilities.formatDate(updatedAt, 'Asia/Taipei', 'yyyy-MM-dd HH:mm') : '',
        updatedBy: row[13], resolved: isResolved(row[7], row[9]),
        timeEfficiency
      };
    }).reverse();

    return { success: true, data: results.slice(0, pageSize), total: results.length };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// 更新異常狀態
// ============================================================
function updateAnomalyStatus(anomalyId, newStatus, updaterName) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('異常紀錄');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === anomalyId) {
        sheet.getRange(i + 1, 10).setValue(newStatus);
        sheet.getRange(i + 1, 13).setValue(new Date());
        sheet.getRange(i + 1, 14).setValue(updaterName);
        return { success: true };
      }
    }
    return { success: false, message: '找不到此紀錄' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// 修改巡房紀錄
// ============================================================
function updateInspectionRecord(recordId, updatedData, updaterName) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('巡房紀錄');
    const modSheet = ss.getSheetByName('修改紀錄');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== recordId) continue;
      if (String(data[i][1]) !== updaterName) return { success: false, message: '只能修改自己的紀錄' };
      if (updatedData.roomNumber !== undefined) sheet.getRange(i + 1, 6).setValue(updatedData.roomNumber);
      if (updatedData.roomStatus !== undefined) sheet.getRange(i + 1, 7).setValue(updatedData.roomStatus);
      const modId = 'M' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMddHHmmss');
      modSheet.appendRow([modId, recordId, updaterName, new Date(), updatedData.modNote || '',
        JSON.stringify({ roomNumber: updatedData.roomNumber, roomStatus: updatedData.roomStatus })]);
      return { success: true };
    }
    return { success: false, message: '找不到此紀錄' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// 統計報表
// ============================================================
function getStatistics(filters) {
  try {
    const ss = getSpreadsheet();
    const anomalyData = ss.getSheetByName('異常紀錄').getDataRange().getValues();
    const inspData = ss.getSheetByName('巡房紀錄').getDataRange().getValues();
    const locData = ss.getSheetByName('據點清單').getDataRange().getValues();
    const visMap = getVisibilityMap();
    const settings = getSystemSettings();
    const locAreaMap = {};
    locData.slice(1).forEach(row => { locAreaMap[row[2]] = row[0]; });

    const fa = anomalyData.slice(1).filter(row => {
      if (!row[0]) return false;
      const locArea = locAreaMap[row[3]] || '';
      if (!hasAccess(filters.userScope, locArea, row[3])) return false;
      if (!canSeeInspector(filters.userScope, visMap[String(row[2]).trim()])) return false;
      const d = row[11] ? new Date(row[11]) : null;
      if (d) {
        if (filters.startDate) { const s = new Date(filters.startDate); s.setHours(0, 0, 0, 0); if (d < s) return false; }
        if (filters.endDate) { const e = new Date(filters.endDate); e.setHours(23, 59, 59, 999); if (d > e) return false; }
      }
      return true;
    });

    const fi = inspData.slice(1).filter(row => {
      if (!row[0]) return false;
      if (!hasAccess(filters.userScope, row[2], row[4])) return false;
      if (!canSeeInspector(filters.userScope, visMap[String(row[1]).trim()])) return false;
      const startDate = row[7] ? new Date(row[7]) : null;
      if (startDate) {
        if (filters.startDate) { const s = new Date(filters.startDate); s.setHours(0, 0, 0, 0); if (startDate < s) return false; }
        if (filters.endDate) { const e = new Date(filters.endDate); e.setHours(23, 59, 59, 999); if (startDate > e) return false; }
      }
      return true;
    });

    // 排名
    const rankMap = {};
    fi.forEach(row => {
      const inspector = String(row[1]).trim();
      if (!rankMap[inspector]) rankMap[inspector] = new Set();
      rankMap[inspector].add(row[4] + '-' + String(row[5]).toLowerCase());
    });
    const ranking = Object.entries(rankMap)
      .map(([name, rooms]) => ({ name, count: rooms.size }))
      .sort((a, b) => b.count - a.count).slice(0, 10);

    // 時效統計
    let avgEfficiency = '';
    if (settings['顯示時效統計'] === '是') {
      const resolved = fa.filter(row => isResolved(row[7], row[9]) && row[11] && row[12]);
      if (resolved.length > 0) {
        const totalHours = resolved.reduce((sum, row) => {
          return sum + (new Date(row[12]) - new Date(row[11])) / 3600000;
        }, 0);
        const avg = Math.round(totalHours / resolved.length);
        avgEfficiency = avg < 24 ? avg + '小時' : Math.round(avg / 24) + '天';
      }
    }

    const count = (arr, idx) => arr.reduce((acc, r) => { acc[r[idx]] = (acc[r[idx]] || 0) + 1; return acc; }, {});
    const sort = obj => Object.entries(obj).sort((a, b) => b[1] - a[1]);

    const locationDetails = {};
    fi.forEach(row => {
      const loc = row[4];
      if (!locationDetails[loc]) locationDetails[loc] = [];
      locationDetails[loc].push({
        inspector: row[1], roomNumber: row[5], roomStatus: row[6],
        startTime: row[7] ? Utilities.formatDate(new Date(row[7]), 'Asia/Taipei', 'yyyy-MM-dd HH:mm') : '',
        duration: row[9]
      });
    });

    const anomalyLocationDetails = {};
    fa.forEach(row => {
      const loc = row[3];
      if (!anomalyLocationDetails[loc]) anomalyLocationDetails[loc] = [];
      anomalyLocationDetails[loc].push({
        inspector: row[2], roomNumber: row[4], inspArea: row[5],
        severity: row[6], type: row[7], note: row[8], status: row[9],
        createdAt: row[11] ? Utilities.formatDate(new Date(row[11]), 'Asia/Taipei', 'yyyy-MM-dd HH:mm') : ''
      });
    });

    return {
      success: true,
      data: {
        totalInspections: fi.length, totalAnomalies: fa.length,
        ranking, avgEfficiency,
        locationStats: sort(count(fa, 3)),
        inspLocationStats: sort(Object.fromEntries(Object.entries(locationDetails).map(([k, v]) => [k, v.length]))),
        areaStats: sort(count(fa, 5)),
        typeStats: sort(count(fa, 7)),
        severityStats: sort(count(fa, 6)),
        locationDetails, anomalyLocationDetails
      }
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// 封存資料查詢
// ============================================================
function getArchiveData(filters) {
  try {
    const settings = getSystemSettings();
    const archiveId = settings['封存試算表ID'];
    if (!archiveId) return { success: false, message: '尚未設定封存試算表ID，請在系統設定分頁填入' };

    const archiveSS = SpreadsheetApp.openById(archiveId);
    const inspSheet = archiveSS.getSheetByName('巡房紀錄');
    const anomalySheet = archiveSS.getSheetByName('異常紀錄');
    if (!inspSheet) return { success: false, message: '封存試算表中找不到巡房紀錄分頁' };

    const inspData = inspSheet.getDataRange().getValues();
    const anomalyData = anomalySheet ? anomalySheet.getDataRange().getValues() : [];
    const pageSize = parseInt(filters.pageSize) || 20;

    const results = inspData.slice(1).filter(row => {
      if (!row[0]) return false;
      if (!hasAccess(filters.userScope, row[2], row[4])) return false;
      const startDate = row[7] ? new Date(row[7]) : null;
      if (startDate) {
        if (filters.startDate) { const s = new Date(filters.startDate); s.setHours(0, 0, 0, 0); if (startDate < s) return false; }
        if (filters.endDate) { const e = new Date(filters.endDate); e.setHours(23, 59, 59, 999); if (startDate > e) return false; }
      }
      if (filters.location && filters.location !== '全部' && row[4] !== filters.location) return false;
      return true;
    }).map(row => {
      const anomalies = anomalyData.slice(1).filter(a => a[1] === row[0]).map(a => ({
        inspArea: a[5], severity: a[6], type: a[7], note: a[8], status: a[9]
      }));
      return {
        recordId: row[0], inspector: row[1], location: row[4], roomNumber: row[5],
        roomStatus: row[6],
        startTime: row[7] ? Utilities.formatDate(new Date(row[7]), 'Asia/Taipei', 'yyyy-MM-dd HH:mm') : '',
        duration: row[9], generalNote: row[10],
        createdAt: row[11] ? Utilities.formatDate(new Date(row[11]), 'Asia/Taipei', 'yyyy-MM-dd HH:mm') : '',
        fillType: row[12] || '',
        anomalies
      };
    }).reverse();

    return { success: true, data: results.slice(0, pageSize), total: results.length, isArchive: true };
  } catch (e) {
    return { success: false, message: '封存查詢失敗：' + e.toString() };
  }
}

// ============================================================
// 定時任務：自動寄送報表
// ============================================================
function sendWeeklyReport() {
  const settings = getSystemSettings();
  const email = settings['通知Email'];
  if (!email) return;

  const today = new Date();
  const weekAgo = new Date(today - 7 * 24 * 3600000);
  const filters = {
    userScope: '全部',
    startDate: Utilities.formatDate(weekAgo, 'Asia/Taipei', 'yyyy-MM-dd'),
    endDate: Utilities.formatDate(today, 'Asia/Taipei', 'yyyy-MM-dd')
  };

  const stats = getStatistics(filters);
  if (!stats.success) return;

  const d = stats.data;
  const subject = '飯店巡房系統週報 - ' + Utilities.formatDate(today, 'Asia/Taipei', 'yyyy/MM/dd');
  const body = '本週巡房摘要\n\n'
    + '巡房總筆數：' + d.totalInspections + '\n'
    + '異常總件數：' + d.totalAnomalies + '\n\n'
    + '各館異常統計：\n'
    + (d.locationStats || []).map(([name, count]) => name + '：' + count + '件').join('\n');

  GmailApp.sendEmail(email, subject, body);
}
