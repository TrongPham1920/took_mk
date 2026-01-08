const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const path = require("path");
const fs = require("fs");
const genMKT = require("./handlers/genMKTPC.js");

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 600,
    height: 750,
    title: "Sim Hải Đăng - Đối Soát TikTok",
    icon: path.join(__dirname, "public", "image.png"),
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      webSecurity: false,
    },
  });

  mainWindow.loadFile("public/indexPc.html");

  // Mở DevTools đúng cách sau khi mainWindow đã được khởi tạo
  // mainWindow.webContents.openDevTools();

  mainWindow.on("closed", () => {
    mainWindow = null;
  });
}

// 1. Handler mở hộp thoại chọn file (Cực kỳ quan trọng để lấy path chuẩn)
ipcMain.handle("open-file-dialog", async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    properties: ["openFile"],
    filters: [{ name: "Excel Files", extensions: ["xlsx", "xls"] }],
  });
  return canceled ? null : filePaths[0];
});

// 2. Handler xử lý Logic Gộp Excel
ipcMain.handle("process-excel", async (event, { file1Path, file2Path }) => {
  try {
    // Kiểm tra dữ liệu đầu vào
    if (!file1Path || !file2Path) {
      throw new Error("Thiếu đường dẫn file đầu vào.");
    }

    const fileObj1 = {
      path: file1Path,
      originalname: path.basename(file1Path),
    };
    const fileObj2 = {
      path: file2Path,
      originalname: path.basename(file2Path),
    };

    // Gọi logic xử lý
    const resultFilePath = await genMKT(fileObj1, fileObj2);

    // Hộp thoại lưu file
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
      // Xóa file tạm sau khi copy thành công
      if (fs.existsSync(resultFilePath)) fs.unlinkSync(resultFilePath);
      return { success: true, message: "Đã xuất báo cáo thành công!" };
    }

    return { success: false, message: "Người dùng đã hủy lưu file." };
  } catch (error) {
    console.error("Lỗi xử lý Excel:", error);
    return { success: false, message: error.message };
  }
});

app.whenReady().then(createWindow);

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

app.on("activate", () => {
  if (mainWindow === null) createWindow();
});
