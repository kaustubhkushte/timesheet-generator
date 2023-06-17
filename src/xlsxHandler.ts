import * as XLSX from 'xlsx';



interface EmployeeData {
  employeeName: string;
  date: Date;
  task: string;
  hrs: number;
}

export function groupByEmployeeName(data: EmployeeData[]): Record<string, EmployeeData[]> {
  const groupedData: Record<string, EmployeeData[]> = {};

  data.forEach((row) => {
    const { employeeName, date, task, hrs } = row;

    if (!groupedData[employeeName]) {
      groupedData[employeeName] = [];
    }

    groupedData[employeeName].push({employeeName, date, task, hrs });
  });

  return groupedData;
}

export function createWorkbook(groupedData: Record<string, EmployeeData[]>): XLSX.WorkBook {
  const workbook: XLSX.WorkBook = XLSX.utils.book_new();

  for (const employeeName in groupedData) {
    const worksheetData = groupedData[employeeName];

    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(worksheetData, {
      header: ['Date', 'Task', 'Hrs'],
      skipHeader: true,
    });

    XLSX.utils.book_append_sheet(workbook, worksheet, employeeName);
  }

  return workbook;
}

export function readExcelFile(filePath: string): Promise<EmployeeData[]> {
  return new Promise((resolve, reject) => {
    try {
      const workbook: XLSX.WorkBook = XLSX.readFile(filePath);
      const worksheet: XLSX.WorkSheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData: EmployeeData[] = XLSX.utils.sheet_to_json(worksheet, {
        header: ['employeeName', 'date', 'task', 'hrs'],
        defval: undefined,
        raw: false,
        dateNF: 'yyyy-mm-dd',
      });

      resolve(jsonData);
    } catch (error) {
      reject(error);
    }
  });
}

export function writeExcelFile(filePath: string, workbook: XLSX.WorkBook): Promise<void> {
  return new Promise((resolve, reject) => {
    try {
      XLSX.writeFile(workbook, filePath);
      resolve();
    } catch (error) {
      reject(error);
    }
  });
}
