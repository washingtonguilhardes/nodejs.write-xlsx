# Excel File Generator

This project generates an Excel file with styled cells using the `exceljs` library.

## Features

- Generates an Excel file with 1000 rows of product data.
- Each row represents a product with a unique ID, a name, and a random inventory amount between 0 and 100.
- The name cell of each row is styled with a background color and text color based on certain conditions.

## Code Snippet

```typescript
// Set cell styles based on certain conditions
let fgColor, fontColor;
switch (condition) {
  case 'condition1':
    fgColor = "FF01B3AC";
    fontColor = "FFFFFFFF";
    break;
  // Add more conditions here
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

// Write to file
workbook.xlsx.writeFile(
  join(__dirname, "..", "output", `test.${Date.now()}.xlsx`)
);
```

# How to Run

Clone the repository.
Install dependencies with `npm install`.
Run the script with `npm start`.
This will generate an Excel file in the `output` directory.

