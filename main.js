const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const path = require("path");
const fs = require("fs");
const genMk = require("./handlers/genMkPc.js");
const genMb = require("./handlers/genMbPc.js");

function createWindow() {
  const win = new BrowserWindow({
    width: 900,
    height: 700,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
    },
  });

  win.loadFile("public/indexPc.html");
  // win.webContents.openDevTools(); // Mở để debug nếu cần
}

app.whenReady().then(createWindow);

// Lắng nghe sự kiện xử lý từ Giao diện
ipcMain.handle("process-excel", async (event, { filePath, mode }) => {
  try {
    const handler = mode === "beauty" ? genMk : genMb;
    const result = await handler(filePath);
    return { success: true, path: result };
  } catch (error) {
    return { success: false, message: error.message };
  }
});

// Lắng nghe yêu cầu chọn file
ipcMain.handle("open-file-dialog", async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    properties: ["openFile"],
    filters: [{ name: "Excel Files", extensions: ["xlsx", "xls"] }],
  });
  return canceled ? null : filePaths[0];
});
