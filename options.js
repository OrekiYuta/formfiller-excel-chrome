const statusEl = document.getElementById('status');
const tableContainer = document.getElementById('tableContainer');
let excelData = [];

// 不清空原始提示，逐条追加日志
function updateStatus(text) {
  const logLine = document.createElement('div');
  logLine.textContent = text;
  statusEl.appendChild(logLine);
  statusEl.scrollTop = statusEl.scrollHeight;
}

// 渲染 Excel 表格，包含序号列
function renderTable(data) {
  if (!data.length) return;

  const keys = Object.keys(data[0]);
  let html = '<table><thead><tr>';

  // 添加序号表头
  html += '<th>序号</th>';

  // 添加其他表头
  keys.forEach(key => {
    html += `<th>${key}</th>`;
  });
  html += '</tr></thead><tbody>';

  // 添加表格内容，序号从1开始
  data.forEach((row, index) => {
    html += `<tr><td>${index + 1}</td>`;
    keys.forEach(key => {
      html += `<td>${row[key] ?? ''}</td>`;
    });
    html += '</tr>';
  });

  html += '</tbody></table>';
  tableContainer.innerHTML = html;
}

// 监听文件上传事件
document.getElementById('excelFile').addEventListener('change', async (e) => {
  tableContainer.innerHTML = ''; // 清空表格
  updateStatus("开始上传 Excel 文件...");
  try {
    const file = e.target.files[0];
    if (!file) {
      updateStatus("未选中文件！");
      return;
    }

    updateStatus("读取文件：" + file.name);
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array" });

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet);

    if (!json.length) {
      updateStatus("Excel 内容为空！");
      return;
    }
    updateStatus("读取成功，共 " + json.length + " 条数据");
    excelData = json;
    renderTable(excelData);

  } catch (err) {
    updateStatus("读取失败: " + err.message);
  }
});

// 点击开始按钮，注入脚本
document.getElementById('run').addEventListener('click', () => {
  if (!excelData.length) {
    updateStatus("请先上传 Excel 文件！");
    return;
  }

  const row = excelData[0];
  updateStatus("打开链接：" + row['独立站链接']);

  chrome.tabs.create({ url: row['独立站链接'] }, (tab) => {
    chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: fillForm,
      args: [row]
    }).then(() => {
      updateStatus("已注入脚本，正在填写表单...");
    }).catch(e => {
      updateStatus("注入脚本失败：" + e.message);
    });
  });
});

// 注入页面脚本，填写表单
function fillForm(row) {
  const setValue = (selector, value) => {
    const el = document.querySelector(selector);
    if (el) {
      el.value = value;
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
    }
  };
  setValue('#comment', row['评论内容']);
  setValue('#author', row['名字']);
  setValue('#email', row['Email']);
}
