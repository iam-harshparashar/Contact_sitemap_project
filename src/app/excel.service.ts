import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';
@Injectable({
  providedIn: 'root',
})
export class ExcelService {
  constructor() {}

// Function to parse an Excel file and return its data as an array of objects

  parseExcel(file: File): Promise<any[]> {
    return new Promise((resolve, reject) => {

    // Create a FileReader to read the contents of the Excel file

      const reader: FileReader = new FileReader();
      reader.onload = (e: any) => {

    // Convert the binary data of the file to Uint8Array

        const data: Uint8Array = new Uint8Array(e.target.result);

    // Use XLSX library to read the data from the Uint8Array and create a workbook

        const workbook: XLSX.WorkBook = XLSX.read(data, { type: 'array' });

    // Get the first worksheet from the workbook

        const worksheet: XLSX.WorkSheet =
          workbook.Sheets[workbook.SheetNames[0]];

    // Convert the worksheet data to an array of objects using sheet_to_json method
        const jsonData: any[] = XLSX.utils.sheet_to_json(worksheet, {
          raw: true,
        });

        // Convert the numeric representation of dates to formatted date strings
        jsonData.forEach((row) => {
          if (row.birthdate && typeof row.birthdate === 'number') {
            const dateValue = row.birthdate;
            const excelDate = new Date(
              Math.round((dateValue - 25569) * 86400 * 1000)
            );
            const formattedDate = excelDate.toLocaleDateString(); // Modify this to the desired date format
            row.birthdate = formattedDate;
          }
        });

        resolve(jsonData);
      };
      reader.onerror = (error) => reject(error);
      reader.readAsArrayBuffer(file);
    });
  }
}
