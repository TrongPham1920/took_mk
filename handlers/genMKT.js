const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

module.exports = async (inputPath) => {
  const timestamp = Date.now();
  const baseDir = path.dirname(inputPath);
  const outputFile = path.join(baseDir, `Tiktok_nam92_${timestamp}.xlsx`);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputPath);

  // Lấy sheet đầu tiên làm dữ liệu nguồn
  const sourceSheet = workbook.worksheets[0];
  const data = [];

  // Lấy header (dòng 1)
  const headers = [];
  sourceSheet.getRow(1).eachCell((cell, colNumber) => {
    headers[colNumber] = cell.value
      ? cell.value.toString().trim().toLowerCase()
      : "";
  });

  // Đọc dữ liệu từ dòng 2
  sourceSheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const rowData = {};
    row.eachCell((cell, colNumber) => {
      const header = headers[colNumber];
      if (header) rowData[header] = cell.value;
    });
    data.push(rowData);
  });

  // Tạo sheet mới đặt tên là nam92
  const nam92Sheet = workbook.addWorksheet("nam92");

  // Định nghĩa cột cho sheet nam92
  nam92Sheet.columns = [
    { header: "Mã Vận Đơn", key: "order_id", width: 25 },
    { header: "ID Theo Dõi", key: "tracking_id", width: 25 },
    { header: "Ngày Đặt Hàng", key: "created_time", width: 20 },
    { header: "Địa Chỉ", key: "address", width: 10 },
    { header: "Sản Phẩm", key: "sku", width: 30 },
    { header: "Số Lượng Trước Hủy", key: "qty_before", width: 15 },
    { header: "Số Lượng Sau Hủy", key: "qty_after", width: 15 },
    { header: "Số Lượng Cuối Cùng", key: "qty_final", width: 15 },
    { header: "Sàn", key: "platform", width: 10 },
    { header: "Shop", key: "shop", width: 15 },
    { header: "Doanh Thu", key: "revenue", width: 15 },
    { header: "Ngày Quyết Toán", key: "settlement_date", width: 15 },
    { header: "Tình Trạng", key: "status", width: 20 },
  ];

  // Map dữ liệu vào sheet mới
  data.forEach((item) => {
    nam92Sheet.addRow({
      order_id: item["order id"] || "",
      tracking_id: item["tracking id"] || "",
      created_time: item["created time"] || "",
      address: "",
      sku: item["seller sku"] || item["sku id"] || "",
      qty_before: item["quantity"] || 0,
      qty_after: 0,
      qty_final: item["quantity"] || 0,
      platform: "TIKTOK",
      shop: "Sim Hải Đăng",
      revenue: item["order amount"] || 0,
      settlement_date: "",
      status: item["order status"] || "",
    });
  });

  // Ghi file (File này sẽ chứa cả sheet gốc và sheet nam92)
  await workbook.xlsx.writeFile(outputFile);

  return outputFile;
};
