const statusEl = document.getElementById('status');
const tableContainer = document.getElementById('tableContainer');
let excelData = [];

const REQUIRED_COLUMNS = ['评论序号', '年', '月', '日', '小时', '分钟'];
const ADMIN_EDIT_URL_PREFIX = 'https://www.teseu.net/wp-admin/comment.php?action=editcomment&c=';

document.getElementById('excelFile').addEventListener('change', function () {
  document.getElementById('fileName').textContent = this.files[0]?.name || '未选择文件';
});

function updateStatus(text) {
  const logLine = document.createElement('div');
  logLine.textContent = text;
  statusEl.appendChild(logLine);
  statusEl.scrollTop = statusEl.scrollHeight;
}

function renderTable(data) {
  if (!data.length) return;

  const keys = [...REQUIRED_COLUMNS, '评论链接'];
  const colWidths = {
    '评论序号': '12%',
    '年': '12%',
    '月': '10%',
    '日': '10%',
    '小时': '10%',
    '分钟': '10%',
    '评论链接': '36%'
  };

  let html = '<table><colgroup>';
  html += '<col style="width: 5%;">';

  keys.forEach((key) => {
    const width = colWidths[key] || '15%';
    html += `<col style="width: ${width};">`;
  });

  html += '</colgroup><thead><tr>';
  html += '<th>序号</th>';

  keys.forEach((key) => {
    html += `<th>${key}</th>`;
  });

  html += '</tr></thead><tbody>';

  data.forEach((row, index) => {
    html += `<tr><td>${index + 1}</td>`;
    keys.forEach((key) => {
      if (key === '评论链接') {
        const commentId = normalizeNumberText(row['评论序号']);
        const url = /^\d+$/.test(commentId) ? ADMIN_EDIT_URL_PREFIX + commentId : '';
        html += url
          ? `<td><a href="${url}" target="_blank" rel="noopener noreferrer">${url}</a></td>`
          : '<td></td>';
      } else {
        html += `<td>${row[key] ?? ''}</td>`;
      }
    });
    html += '</tr>';
  });

  html += '</tbody></table>';
  tableContainer.innerHTML = html;
}

function normalizeNumberText(value) {
  if (value == null) return '';
  const text = String(value).trim();
  if (!text) return '';
  if (/^\d+(\.0+)?$/.test(text)) {
    return String(Math.trunc(Number(text)));
  }
  return text;
}

function normalizeRows(rows) {
  return rows.map((row) => {
    const normalized = {};
    REQUIRED_COLUMNS.forEach((column) => {
      normalized[column] = normalizeNumberText(row[column]);
    });
    return normalized;
  });
}

function parseIntStrict(value) {
  if (!/^\d+$/.test(String(value))) return null;
  return parseInt(String(value), 10);
}

function validateRow(row) {
  const commentId = parseIntStrict(row['评论序号']);
  if (!commentId || commentId <= 0) {
    return { ok: false, message: '评论序号必须是大于 0 的整数' };
  }

  const year = parseIntStrict(row['年']);
  if (!year || year < 1900 || year > 2200) {
    return { ok: false, message: '年必须是 1900-2200 的整数' };
  }

  const month = parseIntStrict(row['月']);
  if (!month || month < 1 || month > 12) {
    return { ok: false, message: '月必须是 1-12 的整数' };
  }

  const day = parseIntStrict(row['日']);
  if (!day || day < 1 || day > 31) {
    return { ok: false, message: '日必须是 1-31 的整数' };
  }

  const hour = parseIntStrict(row['小时']);
  if (hour == null || hour < 0 || hour > 23) {
    return { ok: false, message: '小时必须是 0-23 的整数' };
  }

  const minute = parseIntStrict(row['分钟']);
  if (minute == null || minute < 0 || minute > 59) {
    return { ok: false, message: '分钟必须是 0-59 的整数' };
  }

  return {
    ok: true,
    payload: {
      commentId: String(commentId),
      year: String(year),
      month: String(month).padStart(2, '0'),
      day: String(day).padStart(2, '0'),
      hour: String(hour).padStart(2, '0'),
      minute: String(minute).padStart(2, '0')
    }
  };
}

document.getElementById('excelFile').addEventListener('change', async (e) => {
  tableContainer.innerHTML = '';
  updateStatus('开始上传 Excel 文件...');

  try {
    const file = e.target.files[0];
    if (!file) {
      updateStatus('未选中文件！');
      return;
    }

    updateStatus('读取文件：' + file.name);
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet);

    if (!json.length) {
      updateStatus('Excel 内容为空！');
      return;
    }

    const firstRowKeys = Object.keys(json[0]);
    const missingColumns = REQUIRED_COLUMNS.filter((column) => !firstRowKeys.includes(column));
    if (missingColumns.length) {
      updateStatus('Excel 缺少必要列：' + missingColumns.join('、'));
      updateStatus('需要的列为：' + REQUIRED_COLUMNS.join('、'));
      return;
    }

    excelData = normalizeRows(json);
    renderTable(excelData);

    updateStatus('读取成功，共 ' + excelData.length + ' 条数据');
    updateStatus('');
  } catch (err) {
    updateStatus('读取失败: ' + err.message);
  }
});

document.getElementById('run').addEventListener('click', async () => {
  if (!excelData.length) {
    updateStatus('请先上传 Excel 文件！');
    return;
  }

  const intervalSeconds = parseFloat(document.getElementById('intervalInput').value) || 0;
  const intervalMs = intervalSeconds * 1000 + 2000;

  for (let i = 0; i < excelData.length; i++) {
    const row = excelData[i];
    const seq = i + 1;

    const validation = validateRow(row);
    if (!validation.ok) {
      updateStatus(`序号${seq}：${validation.message}，已跳过`);
      continue;
    }

    const payload = validation.payload;
    const url = ADMIN_EDIT_URL_PREFIX + payload.commentId;
    let tabId = null;

    try {
      updateStatus(`序号${seq}：正在打开链接 ${url}`);

      const tab = await new Promise((resolve, reject) => {
        chrome.tabs.create({ url }, (createdTab) => {
          if (chrome.runtime.lastError) {
            reject(chrome.runtime.lastError);
          } else {
            resolve(createdTab);
          }
        });
      });

      tabId = tab.id;

      await new Promise((resolve) => {
        const listener = (updatedTabId, changeInfo) => {
          if (updatedTabId === tabId && changeInfo.status === 'complete') {
            chrome.tabs.onUpdated.removeListener(listener);
            resolve();
          }
        };
        chrome.tabs.onUpdated.addListener(listener);
      });

      updateStatus(`序号${seq}：正在设置评论时间并更新...`);

      const scriptResult = await chrome.scripting.executeScript({
        target: { tabId },
        func: fillCommentTimestamp,
        args: [payload]
      });

      const result = scriptResult?.[0]?.result;
      if (!result || !result.ok) {
        throw new Error(result?.message || '页面元素定位失败');
      }

      updateStatus(`序号${seq}：更新成功 ✔`);

      if (intervalMs > 0) {
        updateStatus(`等待 ${intervalSeconds} 秒后继续下一条...`);
        await new Promise((resolve) => setTimeout(resolve, intervalMs));
      }
    } catch (e) {
      updateStatus(`序号${seq}：操作失败 ❌ ${e.message}`);
    } finally {
      if (tabId != null) {
        await new Promise((resolve) => {
          chrome.tabs.remove(tabId, () => {
            resolve();
          });
        });
      }
    }
  }

  updateStatus('全部操作完成！');
});

async function fillCommentTimestamp(payload) {
  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  const setInputValue = (selector, value) => {
    const el = document.querySelector(selector);
    if (!el) return false;
    el.value = value;
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
    return true;
  };

  const setSelectValue = (selector, value) => {
    const el = document.querySelector(selector);
    if (!el) return false;
    el.value = value;
    el.dispatchEvent(new Event('change', { bubbles: true }));
    return true;
  };

  const editTimestampBtn = document.querySelector('a.edit-timestamp[href="#edit_timestamp"]') || document.querySelector('a.edit-timestamp');
  if (!editTimestampBtn) {
    return { ok: false, message: '未找到编辑日期和时间按钮' };
  }

  editTimestampBtn.click();
  await sleep(250);

  const requiredSetOk = [
    setInputValue('#aa', payload.year),
    setSelectValue('#mm', payload.month),
    setInputValue('#jj', payload.day),
    setInputValue('#hh', payload.hour),
    setInputValue('#mn', payload.minute)
  ].every(Boolean);

  if (!requiredSetOk) {
    return { ok: false, message: '未找到时间输入框(aa/mm/jj/hh/mn)' };
  }

  const saveTimestampBtn = document.querySelector('a.save-timestamp[href="#edit_timestamp"]') || document.querySelector('a.save-timestamp');
  if (!saveTimestampBtn) {
    return { ok: false, message: '未找到时间确定按钮' };
  }

  saveTimestampBtn.click();
  await sleep(300);

  const updateBtn = document.querySelector('#save') || document.querySelector('input[name="save"]');
  if (!updateBtn) {
    return { ok: false, message: '未找到更新按钮' };
  }

  updateBtn.click();
  return { ok: true };
}

