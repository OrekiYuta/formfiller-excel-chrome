const statusEl = document.getElementById('status');
const tableContainer = document.getElementById('tableContainer');
let excelData = [];
const REQUIRED_COLUMNS = ['标题', '独立站链接', '评论内容', '名字', '邮箱'];


document.getElementById('excelFile').addEventListener('change', function () {
    document.getElementById('fileName').textContent = this.files[0]?.name || '未选择文件';
});

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

  const keys = REQUIRED_COLUMNS;

  // 每列宽度配置
 const colWidths = {
  '标题': '10%',
  '独立站链接': '25%',
  '评论内容': '40%',
  '名字': '10%',
  '邮箱': '15%',
};


  // 构建表格 HTML
  let html = '<table><colgroup>';
  html += '<col style="width: 5%;">'; // 序号列宽度

  // 为每列添加对应宽度
  keys.forEach(key => {
    const width = colWidths[key] || '15%';
    html += `<col style="width: ${width};">`;
  });
  html += '</colgroup><thead><tr>';

  // 表头：序号 + 数据列
  html += '<th>序号</th>';
  keys.forEach(key => {
    html += `<th>${key}</th>`;
  });
  html += '</tr></thead><tbody>';

  // 表体内容
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

    const firstRowKeys = Object.keys(json[0]);
    const missingColumns = REQUIRED_COLUMNS.filter((column) => !firstRowKeys.includes(column));
    if (missingColumns.length) {
      updateStatus("Excel 缺少必要列: " + missingColumns.join('、'));
      updateStatus("需要的列为: " + REQUIRED_COLUMNS.join('、'));
      return;
    }

    const normalizedRows = json.map((row) => {
      const normalized = {};
      REQUIRED_COLUMNS.forEach((column) => {
        normalized[column] = row[column] == null ? '' : String(row[column]).trim();
      });
      return normalized;
    });

    updateStatus("读取成功，共 " + json.length + " 条数据");
    updateStatus(" ");
    updateStatus(" ");

    excelData = normalizedRows;
    renderTable(excelData);
//    console.log("Excel数据:", excelData)

  } catch (err) {
    updateStatus("读取失败: " + err.message);
  }
});


// 点击开始按钮，注入脚本
document.getElementById('run').addEventListener('click', async () => {
  if (!excelData.length) {
    updateStatus("请先上传 Excel 文件！");
    return;
  }

  const intervalSeconds = parseFloat(document.getElementById('intervalInput').value) || 0;
  const intervalMs = intervalSeconds * 1000 + 2000; // 在用户输入的基础上加 2 秒

  for (let i = 0; i < excelData.length; i++) {
  const row = excelData[i];
  const seq = i + 1;
  try {
    const url = row['独立站链接'];
    if (!url) {
      updateStatus(`序号${seq}：缺少“独立站链接”，已跳过`);
      continue;
    }

    updateStatus(`序号${seq}：正在打开链接 ${url}`);

    const tab = await new Promise((resolve, reject) => {
      chrome.tabs.create({ url }, (tab) => {
        if (chrome.runtime.lastError) {
          reject(chrome.runtime.lastError);
        } else {
          resolve(tab);
        }
      });
    });

    await new Promise((resolve) => {
      const listener = (tabId, changeInfo) => {
        if (tabId === tab.id && changeInfo.status === 'complete') {
          chrome.tabs.onUpdated.removeListener(listener);
          resolve();
        }
      };
      chrome.tabs.onUpdated.addListener(listener);
    });

    updateStatus(`序号${seq}：正在填写数据...`);

    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: fillFormTeseu,
      args: [row]
    });

    updateStatus(`序号${seq}：完成填写并提交 ✔`);

    // 等待指定秒数再继续下一条
    if (intervalMs > 0) {
      updateStatus(`等待 ${intervalSeconds} 秒后继续下一条...`);
      await new Promise(resolve => setTimeout(resolve, intervalMs));
    }

    // 关闭当前标签页
    await new Promise((resolve) => {
      chrome.tabs.remove(tab.id, () => {
        if (chrome.runtime.lastError) {
          // 关闭失败也不阻止流程，打印错误日志
          updateStatus(`序号${seq}：关闭标签页失败 ${chrome.runtime.lastError.message}`);
          resolve();
        } else {
          resolve();
        }
      });
    });

    } catch (e) {
      updateStatus(`序号${seq}：操作失败 ❌  ${e.message}`);
    }


  }

  updateStatus("全部操作完成！🎉");
});


// 页面注入脚本，填写表单
function fillFormTeseu(row) {
  const setValue = (selector, value) => {
    const el = document.querySelector(selector);
    if (el) {
      el.value = value;
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
    }
  };

  setValue('input[name="wcpr_review_title"]', row['标题']);
  setValue('#wcpr-comment', row['评论内容']);
  setValue('#author', row['名字']);
  setValue('#email', row['邮箱']);

  // 选中 5 星评分
  const starsEl = document.querySelector('p.stars');
  if (starsEl) {
    starsEl.classList.add('selected'); // 父元素添加 selected 类
    // 找到 5 星的 a 标签，添加 active 类
    const star5 = starsEl.querySelector('a.star-5');
    if (star5) {
      star5.classList.add('active');
      star5.click(); // 触发点击事件，模拟用户选择
    }
  }

  const ratingSelect = document.querySelector('#wcpr-rating');
  if (ratingSelect) {
    ratingSelect.value = '5';
    ratingSelect.dispatchEvent(new Event('change', { bubbles: true }));
  }

  const submitBtn = document.querySelector('#submit') || document.querySelector('input[type="submit"]');
  if (submitBtn) {
    submitBtn.click();
  }
}

