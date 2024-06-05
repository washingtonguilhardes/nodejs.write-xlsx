import ExcelJS from "exceljs";
import color from "./color";
import { join } from "node:path";
const workbook = new ExcelJS.Workbook();

workbook.properties.date1904 = true;
workbook.creator = "Me";
workbook.lastModifiedBy = "Her";

workbook.created = new Date(1985, 8, 30);
workbook.modified = new Date();
workbook.lastPrinted = new Date(2016, 9, 27);

// create a sheet where the grid lines are hidden
// const sheet = workbook.addWorksheet('My Sheet', {views: [{showGridLines: false}]});

// // create a sheet with the first row and column frozen
// const sheet = workbook.addWorksheet('My Sheet', {views:[{state: 'frozen', xSplit: 1, ySplit:1}]});

// Create worksheets with headers and footers
const worksheet = workbook.addWorksheet("My Sheet", {
  headerFooter: { firstHeader: "Hello Exceljs", firstFooter: "Hello World" },
});

worksheet.columns = [
  { header: "Id", key: "id", width: 10 },
  { header: "Name", key: "name", width: 32 },
  { header: "Inventory", key: "inventory", width: 10 },
];
for (let i = 1; i <= 1000; i++) {
  const product = {
    id: i,
    name: `Product ${i}`,
    inventory: Math.floor(Math.random() * 101), // Random number between 0 and 100
  };

  const row = worksheet.addRow(product);

  let fgColor = "FFFFFFFF";
  let fontColor = "00000000";
  switch (true) {
    case product.inventory < 20:
      fontColor = "FFFFFFFF";
      fgColor = "FFED0000";
      break;
    case product.inventory < 30:
      fgColor = "FFEDBC00";
      fontColor = "FFFFFFFF";
      break;
    case product.inventory < 60:
      fgColor = "FF0089ED";
      fontColor = "FFFFFFFF";
      break;
    default:
      fgColor = "FF01B3AC";
      fontColor = "FFFFFFFF";
      break;
  }
  row.eachCell((cell) => {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: fgColor },
    };
    cell.font = {
      bold: true,
      color: { argb: fontColor },
    };
  });
}

workbook.xlsx.writeFile(
  join(__dirname, "..", "output", `test.${Date.now()}.xlsx`)
);
