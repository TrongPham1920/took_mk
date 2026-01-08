const ExcelJS = require("exceljs");
const path = require("path");
const readFileAuto = require("../utils/readFileAuto");

module.exports = async (fileObj1, fileObj2) => {
  const timestamp = Date.now();
  const outputFile = path.join(
    __dirname,
    "../output",
    `merged_${timestamp}.xlsx`
  );
  const outputWorkbook = new ExcelJS.Workbook();

  const processFile = async (fileObj) => {
    const result = await readFileAuto(fileObj);
    if (!result || result.data.length === 0) return;

    const fileName = path.parse(fileObj.originalname).name;
    // Tên sheet tổng = Tên file + Tên sheet đầu của file đó
    const newSheetName = `${fileName}_${result.sheetName}`
      .substring(0, 31)
      .replace(/[\\\/\?\*\[\]:]/g, "");

    const ws = outputWorkbook.addWorksheet(newSheetName);

    // Đổ dữ liệu vào sheet mới
    result.data.forEach((row) => {
      const r = ws.getRow(row.number);
      r.values = row.values;
      r.commit();
    });
  };

  // Chạy lần lượt hoặc song song
  await Promise.all([processFile(fileObj1), processFile(fileObj2)]);

  if (outputWorkbook.worksheets.length === 0) {
    throw new Error("Tất cả các file đều không có dữ liệu đọc được!");
  }

  await outputWorkbook.xlsx.writeFile(outputFile);
  return outputFile;
};
