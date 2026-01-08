const ExcelJS = require("exceljs");

async function generateExcel({
  start = 1,
  end = 100,
  output = "excel_100_rows.xlsx",
}) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("data");

  worksheet.columns = [
    { header: "name", key: "name", width: 25 },
    { header: "front", key: "front", width: 15 },
    { header: "back", key: "back", width: 15 },
    { header: "selfie", key: "selfie", width: 15 },
  ];

  const now = new Date();
  const pad = (n) => String(n).padStart(2, "0");

  const dateTime = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(
    now.getDate()
  )}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;

  for (let i = start; i <= end; i++) {
    worksheet.addRow({
      name: `${i}_${dateTime}`,
      front: `${i}a.jpg`,
      back: `${i}b.jpg`,
      selfie: `${i}c.jpg`,
    });
  }

  await workbook.xlsx.writeFile(output);
  console.log(`✅ Created file: ${output}`);
}

generateExcel({
  start: 1,
  end: 100,
  output: "excel_100_rows.xlsx",
});
