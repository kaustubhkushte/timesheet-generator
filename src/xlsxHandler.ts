import * as XLSX from 'xlsx';
import { EmployeeData } from './models/EmployeeData';
import { TimesheetData } from './models/TimesheetData';
import { EmployeeTimesheetData } from './models/EmployeeTimesheetData';

function evaluateDay(fromDate: string): string {  
  const [day, month, year] = fromDate.split('-').map(Number);
  const date = new Date(year, month - 1, day);
  const options:Intl.DateTimeFormatOptions = {
    weekday: 'long'
  };
  const dayOfWeek = date.toLocaleDateString('en-US', options);
  return dayOfWeek;
}

export function groupByEmployeeName(data: EmployeeData[]): EmployeeTimesheetData[] {
  
  const groupedData: Record<string, TimesheetData[]> = {};
  
  data.forEach((row) => {
    const { employeeName, date, task, hrs }:EmployeeData = row;

    if (!groupedData[employeeName]) {
      groupedData[employeeName] = [];
    }
    const day = evaluateDay(date); 
      
    groupedData[employeeName].push({ date, day, task, hrs });
  });

  const employeeTimesheetArray:EmployeeTimesheetData[] = [];

  for (const employeeName in groupedData) {
    const timesheetData:TimesheetData[] = groupedData[employeeName];
    const employeeTimesheet:EmployeeTimesheetData = {
      employeeName:employeeName,
      data:timesheetData
    }
    employeeTimesheetArray.push(employeeTimesheet);
  }
  

  return employeeTimesheetArray;
}

function calculateSumOfHrs(data: TimesheetData[]): number {
  return data.reduce((sum, entry) => sum + Number(entry.hrs), 0);
}

export function createWorkbook(employeeTimesheetArray: EmployeeTimesheetData[]): XLSX.WorkBook {
  const workbook: XLSX.WorkBook = XLSX.utils.book_new();
  const employeeSummary:any[] = []
  employeeSummary.push(['Employee Name', 'Total Hrs', 'Working Days']);
  employeeTimesheetArray.forEach(employeeTimesheet => {
    const sum = calculateSumOfHrs(employeeTimesheet.data);
    const manDays = Math.ceil(sum / 8)
    const summary:any[] = []
    summary.push(employeeTimesheet.employeeName);
    summary.push(sum);
    summary.push(manDays);
    employeeSummary.push(summary)
  });

  const summaryWorkSheet: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(employeeSummary, {});


  XLSX.utils.book_append_sheet(workbook, summaryWorkSheet, "Summary");

  employeeTimesheetArray.forEach(employeeTimesheet => {
  
    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(employeeTimesheet.data, {
      header: ['date','day', 'task', 'hrs'],
      skipHeader: false,
    });
    worksheet['!cols'] = [
      { width: 15 },
      { width: 20 },
      { width: 100 },
      { width: 50 }
    ];
    XLSX.utils.book_append_sheet(workbook, worksheet, employeeTimesheet.employeeName);

});
  

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
