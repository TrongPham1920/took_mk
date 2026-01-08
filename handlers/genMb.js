const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const archiver = require("archiver");

/**
 * UTILS
 */
const normalizeKey = (key) => (key ? key.toString().trim().toLowerCase() : "");

const outputDir = path.join(__dirname, "..", "output");
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);

/**
 * HÀM DỌN DẸP TRIỆT ĐỂ THƯ MỤC OUTPUT
 */
const cleanOutputFolder = () => {
  try {
    const files = fs.readdirSync(outputDir);
    for (const file of files) {
      const fullPath = path.join(outputDir, file);
      // Xóa tất cả bao gồm file và thư mục con
      fs.rmSync(fullPath, { recursive: true, force: true });
    }
    console.log("🧹 Thư mục output đã được dọn dẹp trống.");
  } catch (e) {
    console.error("❌ Lỗi khi dọn dẹp thư mục output:", e);
  }
};

/**
 * XỬ LÝ CHÍNH
 */
module.exports = async (req, res) => {
  const timestamp = Date.now();
  const sessionDir = path.join(outputDir, `session_mk_${timestamp}`);
  if (!fs.existsSync(sessionDir)) fs.mkdirSync(sessionDir);

  const workbooks = new Map();
  const summary = {};
  let headers = [];
  const outFiles = [];

  try {
    const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(
      req.file.path,
      {
        entries: "emit",
        sharedStrings: "cache",
        worksheets: "emit",
      }
    );

    for await (const worksheet of workbookReader) {
      for await (const row of worksheet) {
        if (row.number === 1) {
          headers = row.values.map((v) => normalizeKey(v));
          continue;
        }

        const rawRow = {};
        headers.forEach((h, idx) => {
          if (h) rawRow[h] = row.values[idx];
        });

        const stbRaw = rawRow["stb"] || "";
        const stbClean = stbRaw.toString().replace(/\D/g, "");
        const giaBan = rawRow["giá bán lẻ"] || rawRow["giá bán"] || 0;
        const dangSo = rawRow["dạng số"]
          ? rawRow["dạng số"].toString().toUpperCase()
          : "UNKNOWN";

        if (!workbooks.has(dangSo)) {
          const cleanName = dangSo.replace(/[^a-zA-Z0-9]/g, "_");
          const filePath = path.join(
            sessionDir,
            `${cleanName}_${timestamp}.xlsx`
          );
          const workbookWriter = new ExcelJS.stream.xlsx.WorkbookWriter({
            filename: filePath,
          });
          const sheet = workbookWriter.addWorksheet("DATA");

          sheet.columns = [
            { header: "stb", key: "stb", width: 15 },
            { header: "stb gốc", key: "stb_goc", width: 15 },
            { header: "Giá bán", key: "gia_ban", width: 15 },
            { header: "Giá khuyến mãi", key: "gia_km", width: 15 },
            { header: "DẠNG SỐ", key: "dang_so", width: 20 },
          ];

          workbooks.set(dangSo, { writer: workbookWriter, sheet, filePath });
          outFiles.push(filePath);
        }

        const { sheet } = workbooks.get(dangSo);
        sheet
          .addRow({
            stb: stbClean,
            stb_goc: stbClean,
            gia_ban: giaBan,
            gia_km: 0,
            dang_so: dangSo,
          })
          .commit();

        summary[dangSo] = (summary[dangSo] || 0) + 1;
      }
    }

    for (const [name, entry] of workbooks) {
      await entry.writer.commit();
    }

    const summaryFile = path.join(sessionDir, `00_tong_hop_${timestamp}.xlsx`);
    const sWriter = new ExcelJS.stream.xlsx.WorkbookWriter({
      filename: summaryFile,
    });
    const sSheet = sWriter.addWorksheet("TONG HOP");
    sSheet.columns = [
      { header: "Dạng Số", key: "type" },
      { header: "Số lượng", key: "count" },
    ];
    Object.entries(summary).forEach(([k, v]) =>
      sSheet.addRow({ type: k, count: v }).commit()
    );
    await sWriter.commit();
    outFiles.push(summaryFile);

    const zipPath = path.join(outputDir, `ket_qua_mk_${timestamp}.zip`);
    const outputZip = fs.createWriteStream(zipPath);
    const archive = archiver("zip", { zlib: { level: 5 } });

    // --- LOGIC DỌN DẸP SAU KHI TẢI XONG ---
    res.on("finish", () => {
      // Đợi 5 giây để đảm bảo file đã được stream hoàn tất tới trình duyệt trước khi xóa
      setTimeout(() => {
        cleanOutputFolder();
      }, 5000);
    });

    outputZip.on("close", () => {
      res.download(zipPath);
    });

    archive.pipe(outputZip);
    outFiles.forEach((f) => archive.file(f, { name: path.basename(f) }));
    await archive.finalize();
  } catch (err) {
    console.error("Lỗi genMk:", err);
    throw err;
  }
};
