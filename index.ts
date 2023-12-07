import * as XLSX from 'xlsx';
import * as fs from 'fs';
import { format } from 'fast-csv';

const processSheet = (sheet: XLSX.WorkSheet): string[] => {
  const data: unknown[][] = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: '',
  }) as unknown[][];

  return data
    .map((row: unknown[]) => {
      const cellValue = row[0];
      const phoneNumber = cellValue?.toString().trim();
      return phoneNumber?.replace('+972', '').replace(/\s|-/g, '');
    })
    .filter(Boolean) as string[]; // Filter out non-phone number entries
};

const processExcelFile = (filePath: string, csvFilePath: string) => {
  const workbook = XLSX.readFile(filePath);
  const stream = fs.createWriteStream(csvFilePath);
  const csvStream = format({ headers: true });

  csvStream.pipe(stream);

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const processedNumbers = processSheet(sheet);
    processedNumbers.forEach((number) =>
      csvStream.write({ PhoneNumber: number })
    );
  });

  csvStream.end();
};

// Example usage
const filePath = './target.xlsx';
const csvFilePath = './output.csv';
const countryCode = '+972';
processExcelFile(filePath, csvFilePath);
