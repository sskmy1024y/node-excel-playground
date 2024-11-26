import ExcelJS from 'exceljs';

const copyWorksheetTemplate = async (
  sourceSheet: ExcelJS.Worksheet,
  newSheet: ExcelJS.Worksheet
): Promise<void> => {
  // Copy page setup and print settings
  if (sourceSheet.pageSetup) {
    newSheet.pageSetup = { ...sourceSheet.pageSetup };
  }

  // Clear all border styles
  newSheet.properties.showGridLines = false;

  // Copy merged cells first
  sourceSheet.model.merges.forEach((mergeRange) => {
    newSheet.mergeCells(mergeRange);
  });

  // Copy rows and styles from the template worksheet to the new worksheet
  sourceSheet.eachRow({ includeEmpty: true }, (row, rowIndex) => {
    const newRow = newSheet.getRow(rowIndex);
    row.eachCell({ includeEmpty: true }, (cell, colIndex) => {
      const newCell = newRow.getCell(colIndex);

      // Copy cell value
      newCell.value = cell.value;

      // Copy cell style
      if (cell.style) {
        newCell.style = { ...cell.style };
      }
    });
    newRow.commit();
  });

  // Adjust column widths
  sourceSheet.columns.forEach((col, colIndex) => {
    if (col.width) {
      newSheet.getColumn(colIndex + 1).width = col.width;
    }
  });

  
}

const main = async () => {
  const templateWB = new ExcelJS.Workbook();
  await templateWB.xlsx.readFile("./src/template.xlsx");

  const templateWS = templateWB.getWorksheet("テンプレ");
  if (!templateWS) {
    throw new Error("テンプレートシートが見つかりません");
  }

  const options = {
    filename: "test.xlsx",
    useStyles: true,
    useSharedStrings: true
  };
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter(options);

  const worksheet = workbook.addWorksheet("利用者1");
  copyWorksheetTemplate(templateWS, worksheet);

  const worksheet2 = workbook.addWorksheet("利用者2");
  worksheet2.getCell("A1").value = "利用者2のデータ";

  worksheet.commit();
  worksheet2.commit();
  workbook.commit();
};

main();
