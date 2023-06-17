import * as XLSX from 'xlsx';
import { EmployeeData } from './models/EmployeeData';
import { TimesheetData } from './models/TimesheetData';


function evaluateDay(fromDate: string): string {  
  const [day, month, year] = fromDate.split('-').map(Number);
  const date = new Date(year, month - 1, day);
  const options:Intl.DateTimeFormatOptions = {
    weekday: 'long'
  };
  const dayOfWeek = date.toLocaleDateString('en-US', options);
  return dayOfWeek;
}

export function groupByEmployeeName(data: EmployeeData[]): Record<string, TimesheetData[]> {
  const groupedData: Record<string, TimesheetData[]> = {};

  data.forEach((row) => {
    const { employeeName, date, task, hrs } = row;

    if (!groupedData[employeeName]) {
      groupedData[employeeName] = [];
    }
    const day = evaluateDay(date);    
    groupedData[employeeName].push({ date, day, task, hrs });
  });

  return groupedData;
}

export function createWorkbook(groupedData: Record<string, TimesheetData[]>): XLSX.WorkBook {
  const workbook: XLSX.WorkBook = XLSX.utils.book_new();

  for (const employeeName in groupedData) {
    const worksheetData = groupedData[employeeName];
        
    
    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(worksheetData, {
      header: ['date','day', 'task', 'hrs'],
      skipHeader: false,
    });
    worksheet['!cols'] = [
      { width: 15 },
      { width: 20 },
      { width: 100 },
      { width: 50 }
    ];
    XLSX.utils.book_append_sheet(workbook, worksheet, employeeName);
  }

  return workbook;
}

export function readExcelFile(filePath: string): Promise<EmployeeData[]> {
  return new Promise((resolve, reject) => {
    try {
      const workbook: XLSX.WorkBook = XLSX.readFile(filePath);
      const worksheet: XLSX.WorkSheet = workbook.Sheets[workbook.SheetNames[0]];
      
      const jsonArray = XLSX.utils.sheet_to_json(worksheet, {        
        defval: undefined,
        raw: false,
        dateNF: 'yyyy-mm-dd',
      });

      const jsonData: EmployeeData[] = jsonArray.map((rowObject:any)=>{
        const employeeData:EmployeeData = {
          employeeName: rowObject.EmployeeName,
          date: rowObject.Date,
          task: rowObject.Description,
          hrs: rowObject.TotalWorkingHours
        } 
        return employeeData;
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
