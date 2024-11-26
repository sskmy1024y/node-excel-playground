import ExcelJS from 'exceljs';
import { copyWorksheet } from './copyWorksheet';

// streaming I/Oを使うとリッチテキストの書き込みができない
// https://github.com/exceljs/exceljs/issues/409

const main = async () => {
  const templateWB = new ExcelJS.Workbook();
  await templateWB.xlsx.readFile("./src/template.xlsx");

  const templateWS = templateWB.getWorksheet("テンプレ");
  if (!templateWS) {
    throw new Error("テンプレートシートが見つかりません");
  }

  const workbook = new ExcelJS.Workbook();

  const array = [0]

  array.forEach((i) => {
    const worksheet = workbook.addWorksheet(`user${i}`, {
      views: [{ showGridLines: false }]
    });
    copyWorksheet(templateWS, worksheet);
});

  await workbook.xlsx.writeFile("./test.xlsx");
};

main();
