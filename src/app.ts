import express from 'express';
import ejs from 'ejs';
import multer from 'multer';
import * as XLSXHandler from './xlsxHandler';
import swaggerJSDoc from 'swagger-jsdoc';
import swaggerUi from 'swagger-ui-express';
import * as XLSX from 'xlsx';
import fs from 'fs';
interface EmployeeData {
  employeeName: string;
  date: Date;
  task: string;
  hrs: number;
}

const app = express();
const upload = multer({ dest: 'uploads/' });

app.set('view engine', 'ejs');

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.get('/', (_req, res) => {
  res.render('landing', { title: 'Timesheet Data Formatter' });
});

app.post('/generate', upload.single('xlsxFile'), async (req, res) => {
  if (!req.file) {
    res.status(400).send('No file uploaded');
    return;
  }
  const filePath = req.file.path;
  try {
    const jsonData: EmployeeData[] = await XLSXHandler.readExcelFile(filePath);
    const groupedData: Record<string, EmployeeData[]> = XLSXHandler.groupByEmployeeName(jsonData);
    const workbook: XLSX.WorkBook = XLSXHandler.createWorkbook(groupedData);
    const outputPath = 'converted.xlsx';
    await XLSXHandler.writeExcelFile(outputPath, workbook);
    res.download(outputPath, 'converted.xlsx', () => {
      // Clean up the uploaded file and converted file
      fs.unlinkSync(filePath);
      fs.unlinkSync(outputPath);
    });
  } catch (error) {
    console.error('Error occurred while generating the converted file:', error);
    res.status(500).send('Error occurred while generating the converted file');
  }
});

// Swagger configuration
const swaggerDefinition = {
  openapi: '3.0.0',
  info: {
    title: 'Timesheet Data Formatter API',
    version: '1.0.0',
    description: 'API to upload and convert Timesheet data from XLSX',
  },
  servers: [
    {
      url: 'http://localhost:3000',
      description: 'Local server',
    },
  ],
};

const options = {
  swaggerDefinition,
  apis: ['./index.ts'],
};

const swaggerSpec = swaggerJSDoc(options);
app.use('/api-docs', swaggerUi.serve, swaggerUi.setup(swaggerSpec));

/**
 * @swagger
 * /generate:
 *   post:
 *     summary: Upload and convert Timesheet data from XLSX
 *     requestBody:
 *       required: true
 *       content:
 *         multipart/form-data:
 *           schema:
 *             type: object
 *             properties:
 *               xlsxFile:
 *                 type: string
 *                 format: binary
 *     responses:
 *       '200':
 *         description: OK
 *         content:
 *           application/vnd.openxmlformats-officedocument.spreadsheetml.sheet:
 *             schema:
 *               type: string
 *               format: binary
 *       '500':
 *         description: Internal Server Error
 *         content:
 *           application/json:
 *             schema:
 *               $ref: '#/components/schemas/Error'
 */
app.listen(3000, () => {
  console.log('Server is running on port 3000');
});
