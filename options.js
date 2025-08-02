const statusEl = document.getElementById('status');
const tableContainer = document.getElementById('tableContainer');
let excelData = [];


document.getElementById('excelFile').addEventListener('change', function () {
    const fileName = this.files[0]?.name || '未选择文件';
    document.getElementById('fileName').textContent = fileName;
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

  const keys = Object.keys(data[0]);

  // 每列宽度配置
 const colWidths = {
  '品名': '10%',
  '沃尔玛ITEM ID': '10%',
  '独立站链接': '20%',
  '评论内容': '40%',
  '名字': '5%',
  '邮箱': '10%',
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
    updateStatus("读取成功，共 " + json.length + " 条数据");
    updateStatus(" ");
    updateStatus(" ");

    excelData = json;
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
    updateStatus(`序号${seq}：正在打开链接 ${row['独立站链接']}`);

    const tab = await new Promise((resolve, reject) => {
      chrome.tabs.create({ url: row['独立站链接'] }, (tab) => {
        if (chrome.runtime.lastError) {
          reject(chrome.runtime.lastError);
        } else {
          resolve(tab);
        }
      });
    });

    updateStatus(`序号${seq}：正在填写数据...`);
    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: fillForm,
      args: [row]
    });

    updateStatus(`序号${seq}：完成填写并提交 ✔`);

    // 等待指定秒数再继续下一条
    if (i < excelData.length - 1 && intervalMs > 0) {
      updateStatus(`等待 ${intervalSeconds} 秒后继续下一条...`);
      await new Promise(resolve => setTimeout(resolve, intervalMs));
    }

    // 关闭当前标签页
    await new Promise((resolve, reject) => {
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

  const submitBtn = document.querySelector('#submit');
  if (submitBtn) {
    submitBtn.click();
  }
}


