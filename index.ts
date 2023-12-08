import * as XLSX from 'xlsx';
import * as fs from 'fs';
import { format } from 'fast-csv';

const processSheet = (sheet: XLSX.WorkSheet): string[] => {
  const data: unknown[][] = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: '',
  }) as unknown[][];

  const getData = data
    .map((row: unknown[]) => {
      const cellValue = row[0];
      let phoneNumber = cellValue?.toString().trim();
      phoneNumber = phoneNumber?.replace('+972', '').replace(/\s|-/g, '');
      phoneNumber = phoneNumber?.replace(/\u2066/g, '');

      if (phoneNumber?.length === 9) {
        return phoneNumber;
      } else {
        return null;
      }
    })
    .filter(Boolean) as string[];

  const newData = [...new Set(getData)];

  return newData;
};

const saveToExcel = (processedData: string[], outputFilePath: string) => {
  const newWorkbook = XLSX.utils.book_new();
  const newWorksheet = XLSX.utils.aoa_to_sheet(
    processedData.map((phoneNumber) => [phoneNumber])
  );

  //
  XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'ProcessedNumbers');

  XLSX.writeFile(newWorkbook, outputFilePath);
};

const processExcelFile = (filePath: string, outputFilePath: string) => {
  const workbook = XLSX.readFile(filePath);
  let allProcessedNumbers: string[] = [];

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const processedNumbers = processSheet(sheet);
    allProcessedNumbers = allProcessedNumbers.concat(processedNumbers);
  });

  saveToExcel(allProcessedNumbers, outputFilePath);
};

const filePath = './target.xlsx';
const csvFilePath = './output.xlsx';
const countryCode = '+972';
processExcelFile(filePath, csvFilePath);
