const statusEl = document.getElementById('status');
let excelData = [];

function updateStatus(text) {
  statusEl.textContent += text + "\n";
  statusEl.scrollTop = statusEl.scrollHeight; // 自动滚动到底部
}

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
