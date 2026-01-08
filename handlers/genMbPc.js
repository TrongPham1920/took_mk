const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

module.exports = async (inputPath) => {
  const timestamp = Date.now();
  // Lấy đường dẫn thư mục chứa file gốc để tạo folder kết quả tại đó
  const baseDir = path.dirname(inputPath);
  const sessionDir = path.join(baseDir, `Ket_Qua_Beauty_${timestamp}`);

  if (!fs.existsSync(sessionDir)) fs.mkdirSync(sessionDir);

  const workbooks = new Map();
  const summary = {};
  let headers = [];

  const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(inputPath, {
    entries: "emit",
    sharedStrings: "cache",
    worksheets: "emit",
  });

  for await (const worksheet of workbookReader) {
    for await (const row of worksheet) {
      if (row.number === 1) {
        headers = row.values.map((v) =>
          v ? v.toString().trim().toLowerCase() : ""
        );
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
        const filePath = path.join(sessionDir, `${cleanName}.xlsx`);
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
        workbooks.set(dangSo, { writer: workbookWriter, sheet });
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

  // Đóng tất cả file
  for (const [name, entry] of workbooks) {
    await entry.writer.commit();
  }

  // Tạo file tổng hợp
  const summaryFile = path.join(sessionDir, `00_tong_hop.xlsx`);
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

  return sessionDir; // Trả về đường dẫn thư mục kết quả
};
