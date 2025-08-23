/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("Insert").addEventListener("click", insertUnicode);

  }
});

async function insertUnicode() {
  const input = document.getElementById("unicodeInput").value.trim().toUpperCase();

  if (!input) {
    showMessage("请输入 Unicode 码，例如 U+2103 或 2103");
    return;
  }

  // 处理输入（支持 U+2103 或 2103）
  const hex = input.startsWith("U+") ? input.slice(2) : input;
  let char;
  try {
    char = String.fromCodePoint(parseInt(hex, 16));
  } catch (e) {
    showMessage("请输入 Unicode 码，例如 U+2103 或 2103");
    return;
  }
  showMessage("");

  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("values");
    await context.sync();

    let current = range.values[0][0] || "";
    range.values = [[current + char]];
    await context.sync();
  });
}

function showMessage(msg, color="red") {
  const el = document.getElementById("message");
  el.style.color = color;
  el.textContent = msg;
}
