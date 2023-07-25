import * as XLSX from 'xlsx';
import { EmployeeData } from './models/EmployeeData';
import { TimesheetData } from './models/TimesheetData';
import { EmployeeTimesheetData } from './models/EmployeeTimesheetData';
import exp from 'constants';

function calculateStartEndDate(referenceDateStr:string){
  
// Split the date string
const dateParts = referenceDateStr.split('-');

// Extract day, month, and year from the date parts
const day = parseInt(dateParts[0], 10);
const month = parseInt(dateParts[1], 10) - 1; // Adjust month to 0-based index
const year = parseInt(dateParts[2], 10);

// Create the reference date object
const referenceDate = new Date(year, month, day);

// Get the year and month from the reference date
const referenceYear = referenceDate.getFullYear();
const referenceMonth = referenceDate.getMonth();

// Calculate the start date of the month
const startDate = new Date(referenceYear, referenceMonth, 1);
const startDateString = `${startDate.getFullYear()}-${('0' + (startDate.getMonth() + 1)).slice(-2)}-${('0' + startDate.getDate()).slice(-2)}`;

// Calculate the end date of the month
const endDate = new Date(referenceYear, referenceMonth + 1, 0);
const endDateString = `${endDate.getFullYear()}-${('0' + (endDate.getMonth() + 1)).slice(-2)}-${('0' + endDate.getDate()).slice(-2)}`;

// Print the start and end dates
console.log('Start Date:', startDateString);
console.log('End Date:', endDateString);
return {
  "startDate":startDate,
  "endDate":endDate
}
}

function evaluateDay(fromDate: string): string {  
  const [day, month, year] = fromDate.split('-').map(Number);
  const date = new Date(year, month - 1, day);
  const options:Intl.DateTimeFormatOptions = {
    weekday: 'long'
  };
  const dayOfWeek = date.toLocaleDateString('en-US', options);
  return dayOfWeek;
}

export function addMissingDaysData(timesheetData:TimesheetData[]){
  // Extract the unique dates from the CSV data
  const existingDates = [...new Set(timesheetData.map((data) => data.date))];
  const monthDates = calculateStartEndDate(existingDates[0]);
  const startDate = monthDates.startDate;
  const endDate = monthDates.endDate;
  // Get all dates for the month
  const allDates: string[] = [];
  for (let date = startDate; date <= endDate; date.setDate(date.getDate() + 1)) {
    const year = date.getFullYear();
  const month = ('0' + (date.getMonth() + 1)).slice(-2);
  const day = ('0' + date.getDate()).slice(-2);
  const dateString = `${day}-${month}-${year}`;
    
    allDates.push(dateString);
  }
  
  // Find the missing dates
  const missingDates = allDates.filter((date) => !existingDates.includes(date));

  // Generate default data for missing dates
  const blankData = missingDates.map((date) => ({
    date,
    day:evaluateDay(date),
    task: '',
    hrs: 0,
  }));

// Combine the existing data with the default data
const updatedData = [...timesheetData, ...blankData];


// Sort the data by date
updatedData.sort((a, b) => {
  const datePartsA = a.date.split('-');
  const datePartsB = b.date.split('-');

  // Convert the date parts to yyyy-mm-dd format for comparison
  const dateA = new Date(`${datePartsA[2]}-${datePartsA[1]}-${datePartsA[0]}`);
  const dateB = new Date(`${datePartsB[2]}-${datePartsB[1]}-${datePartsB[0]}`);

  return dateA.getTime() - dateB.getTime();
});

  return updatedData;
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

    const updatedData = addMissingDaysData(timesheetData);

    const employeeTimesheet:EmployeeTimesheetData = {
      employeeName:employeeName,
      data:updatedData
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
      { width: 30 },
      { width: 15 }
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
        raw: true,
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
