import ExcelJS from 'exceljs';

export const copyWorksheet = async (
  sourceSheet: ExcelJS.Worksheet,
  newSheet: ExcelJS.Worksheet
): Promise<void> => {
  // Copy page setup and print settings
  if (sourceSheet.pageSetup) {
    newSheet.pageSetup = { ...sourceSheet.pageSetup };
  }

  // Copy merged cells first
  sourceSheet.model.merges.forEach((mergeRange) => {
    newSheet.mergeCells(mergeRange);
  });

   // Adjust column widths
   sourceSheet.columns.forEach((col, colIndex) => {
    if (col.width) {
      newSheet.getColumn(colIndex + 1).width = col.width;
    }
  });

  // Copy rows and styles from the template worksheet to the new worksheet
  sourceSheet.eachRow({ includeEmpty: true }, (row, rowIndex) => {
    const newRow = newSheet.getRow(rowIndex);
    row.eachCell({ includeEmpty: true }, (cell, colIndex) => {
      const newCell = newRow.getCell(colIndex);
      
      // Copy cell value
      newCell.value = structuredClone(cell.value);
    
      // Copy cell style
      if (cell.style) {
        newCell.style = { ...cell.style };
      }
    });

    newRow.commit();
  });
}
