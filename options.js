const statusEl = document.getElementById('status');
let excelData = [];

function updateStatus(text) {
  statusEl.textContent += text + "\n";
  statusEl.scrollTop = statusEl.scrollHeight;
}

// 读取Excel逻辑不变
document.getElementById('excelFile').addEventListener('change', async (e) => {
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
    } else {
      updateStatus("读取成功，共 " + json.length + " 条数据");
      excelData = json;
      console.log("Excel数据:", excelData);
    }
  } catch (err) {
    updateStatus("读取失败: " + err.message);
  }
});

// 运行时循环执行每条数据
document.getElementById('run').addEventListener('click', async () => {
  if (!excelData.length) {
    updateStatus("请先上传 Excel 文件！");
    return;
  }

  updateStatus(`开始处理共 ${excelData.length} 条数据...`);

  for (let i = 0; i < excelData.length; i++) {
    const row = excelData[i];
    updateStatus(`第 ${i+1} 条，打开链接：${row['独立站链接']}`);

    // 打开标签页并注入脚本填表
    const tab = await createTabAsync(row['独立站链接']);
    updateStatus(`第 ${i+1} 条，页面已打开，准备注入填表脚本...`);

    try {
      await executeScriptAsync(tab.id, fillForm, [row]);
      updateStatus(`第 ${i+1} 条，已完成填写。`);
    } catch(e) {
      updateStatus(`第 ${i+1} 条，注入脚本失败：${e.message}`);
    }

    // 等待 5 秒（或根据实际情况调整）
    await delay(5000);
  }

  updateStatus("所有数据处理完成！");
});

function createTabAsync(url) {
  return new Promise((resolve) => {
    chrome.tabs.create({ url }, (tab) => resolve(tab));
  });
}

function executeScriptAsync(tabId, func, args) {
  return chrome.scripting.executeScript({
    target: { tabId },
    func,
    args
  });
}

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// 注入到页面的填表脚本
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

  // 点击第5颗星星
  const star5 = document.querySelector('a.star-5');
  if (star5) {
    star5.click();
  } else {
    console.warn('找不到 5星 星星元素');
  }
}
