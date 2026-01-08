const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const path = require("path");
const fs = require("fs");
const genMKT = require("./handlers/genMKTV1.js");

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 600,
    height: 750,
    title: "Sim Hải Đăng - Đối Soát TikTok",
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
    },
  });
  mainWindow.loadFile(path.join(__dirname, "public/indexPC.html"));
  // mainWindow.setMenu(null);
}

ipcMain.handle("process-excel", async (event, { file1Path, file2Path }) => {
  try {
    const fileObj1 = {
      path: file1Path,
      originalname: path.basename(file1Path),
    };
    const fileObj2 = {
      path: file2Path,
      originalname: path.basename(file2Path),
    };

    const resultFilePath = await genMKT(fileObj1, fileObj2);

    const { filePath } = await dialog.showSaveDialog({
      title: "Chọn nơi lưu file kết quả",
      defaultPath: path.join(
        app.getPath("desktop"),
        `Bao_Cao_Tong_Hop_${Date.now()}.xlsx`
      ),
      filters: [{ name: "Excel Files", extensions: ["xlsx"] }],
    });

    if (filePath) {
      fs.copyFileSync(resultFilePath, filePath);
      if (fs.existsSync(resultFilePath)) fs.unlinkSync(resultFilePath);
      return { success: true, message: "Đã xuất báo cáo thành công!" };
    }
    return { success: false, message: "Hủy bỏ lưu file." };
  } catch (error) {
    return { success: false, message: error.message };
  }
});

app.whenReady().then(createWindow);
app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
