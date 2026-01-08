const XLSX = require("xlsx");
const fs = require("fs");

module.exports = async (fileObj) => {
  const buffer = fs.readFileSync(fileObj.path);

  const wb = XLSX.read(buffer, {
    type: "buffer",
    raw: false,
    cellDates: true,
  });

  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];

  if (!sheet) return { type: "empty", data: [] };

  // convert → array of array
  const data = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
  });

  if (!data.length) return { type: "empty", data: [] };

  // map row thẳng, không unshift
  const rows = data.map((row, i) => ({
    number: i + 1,
    values: row,
  }));

  console.log("HEADER:", rows[0].values.slice(0, 6));

  return {
    type: "xlsx",
    sheetName,
    data: rows,
  };
};
