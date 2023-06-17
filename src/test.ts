import * as XLSXHandler from './xlsxHandler';
import { EmployeeData } from './xlsxHandler';

describe('XLSX Handler', () => {
  const sampleData: EmployeeData[] = [
    {
      employeeName: 'John Doe',
      date: new Date('2023-06-01'),
      task: 'Task 1',
      hrs: 5,
    },
    {
      employeeName: 'John Doe',
      date: new Date('2023-06-01'),
      task: 'Task 2',
      hrs: 3,
    },
    {
      employeeName: 'Jane Smith',
      date: new Date('2023-06-02'),
      task: 'Task 3',
      hrs: 8,
    },
  ];

  describe('groupByEmployeeName', () => {
    it('should group employee data by employee name', () => {
      const groupedData = XLSXHandler.groupByEmployeeName(sampleData);

      expect(groupedData).toEqual({
        'John Doe': [
          {
            employeeName: 'John Doe',
            date: new Date('2023-06-01'),
            task: 'Task 1',
            hrs: 5,
          },
          {
            employeeName: 'John Doe',
            date: new Date('2023-06-01'),
            task: 'Task 2',
            hrs: 3,
          },
        ],
        'Jane Smith': [
          {
            employeeName: 'Jane Smith',
            date: new Date('2023-06-02'),
            task: 'Task 3',
            hrs: 8,
          },
        ],
      });
    });
  });

  describe('createWorkbook', () => {
    it('should create a workbook with separate worksheets for each employee', () => {
      const groupedData = XLSXHandler.groupByEmployeeName(sampleData);
      const workbook = XLSXHandler.createWorkbook(groupedData);

      expect(workbook.SheetNames).toEqual(['John Doe', 'Jane Smith']);

      const johnDoeSheet = workbook.Sheets['John Doe'];
      const johnDoeData = XLSX.utils.sheet_to_json(johnDoeSheet);

      expect(johnDoeData).toEqual([
        { Date: '2023-06-01', Task: 'Task 1', Hrs: 5 },
        { Date: '2023-06-01', Task: 'Task 2', Hrs: 3 },
      ]);

      const janeSmithSheet = workbook.Sheets['Jane Smith'];
      const janeSmithData = XLSX.utils.sheet_to_json(janeSmithSheet);

      expect(janeSmithData).toEqual([{ Date: '2023-06-02', Task: 'Task 3', Hrs: 8 }]);
    });
  });

  describe('readExcelFile', () => {
    it('should read data from an Excel file', async () => {
      const filePath = 'test.xlsx';

      const jsonData = await XLSXHandler.readExcelFile(filePath);

      expect(jsonData).toEqual(sampleData);
    });
  });

  describe('writeExcelFile', () => {
    it('should write data to an Excel file', async () => {
      const filePath = 'test.xlsx';
      const workbook = XLSXHandler.createWorkbook(XLSXHandler.groupByEmployeeName(sampleData));

      await XLSXHandler.writeExcelFile(filePath, workbook);

      // Perform assertions or additional checks if necessary
    });
  });
});
