import { IExecuteFunctions, ILoadOptionsFunctions, INodeExecutionData } from 'n8n-workflow';
import { ExcelAI } from './ExcelAI.node';
import * as ExcelJS from 'exceljs';
import * as fs from 'fs/promises';

// Mock n8n-workflow
jest.mock('n8n-workflow', () => ({
	...jest.requireActual('n8n-workflow'),
	NodeOperationError: class extends Error {
		constructor(node: any, message: string) {
			super(message);
			this.name = 'NodeOperationError';
		}
	},
}));

// Mock fs/promises
jest.mock('fs/promises');

// Mock ExcelJS
jest.mock('exceljs', () => {
	const actual = jest.requireActual('exceljs');
	return {
		...actual,
		Workbook: jest.fn().mockImplementation(() => {
			const workbook = new actual.Workbook();
			const originalReadFile = workbook.xlsx.readFile.bind(workbook);
			const originalWriteFile = workbook.xlsx.writeFile.bind(workbook);
			
			workbook.xlsx.readFile = jest.fn().mockImplementation(async (path: string) => {
				// Use mocked buffer data instead of actual file read
				const mockBuffer = (global as any).__mockExcelBuffer__;
				if (mockBuffer) {
					await workbook.xlsx.load(mockBuffer);
					return workbook;
				}
				return originalReadFile(path);
			});
			
			workbook.xlsx.writeFile = jest.fn().mockResolvedValue(undefined);
			
			return workbook;
		}),
	};
});

describe('ExcelAI Node - Unit Tests', () => {
	let excelAI: ExcelAI;
	let mockContext: jest.Mocked<IExecuteFunctions>;
	let mockLoadOptionsContext: jest.Mocked<ILoadOptionsFunctions>;

	beforeEach(() => {
		excelAI = new ExcelAI();
		
		// Reset global mock buffer
		(global as any).__mockExcelBuffer__ = null;
		
		// Setup mock context
		mockContext = {
			getNode: jest.fn().mockReturnValue({ name: 'ExcelAI' }),
			getNodeParameter: jest.fn(),
			getInputData: jest.fn().mockReturnValue([{ json: {} }]),
			helpers: {
				assertBinaryData: jest.fn(),
				getBinaryDataBuffer: jest.fn() as any,
				prepareBinaryData: jest.fn().mockImplementation(async (buffer: Buffer, filename: string) => ({
					data: buffer.toString('base64'),
					mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
					fileName: filename,
				})),
			},
			getCredentials: jest.fn(),
			continueOnFail: jest.fn().mockReturnValue(false),
		} as any;

		// Setup mock load options context
		mockLoadOptionsContext = {
			getNode: jest.fn().mockReturnValue({ name: 'ExcelAI' }),
			getNodeParameter: jest.fn(),
			helpers: {
				request: jest.fn(),
			},
		} as any;

		// Mock fs
		(fs.readFile as jest.Mock).mockResolvedValue(Buffer.from(''));
		(fs.writeFile as jest.Mock).mockResolvedValue(undefined);
		(fs.access as jest.Mock).mockResolvedValue(undefined);
	});

	afterEach(() => {
		jest.clearAllMocks();
	});

	// ===== 1. Node Properties Tests =====
	describe('Node Properties', () => {
		test('should have correct node name', () => {
			expect(excelAI.description.name).toBe('excelAI');
			expect(excelAI.description.displayName).toBe('Excel AI');
		});

		test('should have correct node type', () => {
			expect(excelAI.description.group).toContain('transform');
		});

		test('should have correct version', () => {
			expect(excelAI.description.version).toBe(1);
		});

		test('should have AI Agent support', () => {
			expect(excelAI.description.usableAsTool).toBe(true);
		});

		test('should have correct inputs and outputs', () => {
			expect(excelAI.description.inputs).toEqual(['main']);
			expect(excelAI.description.outputs).toEqual(['main']);
		});

		test('should have all required properties defined', () => {
			const properties = excelAI.description.properties;
			expect(properties).toBeDefined();
			expect(Array.isArray(properties)).toBe(true);
			expect(properties.length).toBeGreaterThan(0);
		});

		test('should have Resource options', () => {
			const resourceProp = excelAI.description.properties.find((p: any) => p.name === 'resource');
			expect(resourceProp).toBeDefined();
			expect(resourceProp?.type).toBe('options');
			expect(resourceProp?.options).toContainEqual(
				expect.objectContaining({ value: 'row' })
			);
			expect(resourceProp?.options).toContainEqual(
				expect.objectContaining({ value: 'worksheet' })
			);
		});

		test('should have all Row operation options', () => {
			const operationProp = excelAI.description.properties.find((p: any) => p.name === 'operation');
			expect(operationProp).toBeDefined();
			const operationValues = operationProp?.options?.map((o: any) => o.value);
			expect(operationValues).toContain('readRows');
			expect(operationValues).toContain('appendRow');
			expect(operationValues).toContain('insertRow');
			expect(operationValues).toContain('updateRow');
			expect(operationValues).toContain('deleteRow');
			expect(operationValues).toContain('filterRows');
		});
	});

	// ===== 2. Load Options Methods Tests =====
	describe('Load Options Methods', () => {
		test('should have methods property', () => {
			expect(excelAI.methods).toBeDefined();
			expect(excelAI.methods?.loadOptions).toBeDefined();
		});

		test('getWorksheets should return worksheets list', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			testWorkbook.addWorksheet('Sheet1');
			testWorkbook.addWorksheet('Sheet2');
			testWorkbook.addWorksheet('DataSheet');
			
			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockLoadOptionsContext.getNodeParameter.mockImplementation((paramName: string) => {
				if (paramName === 'inputMode') return 'filePath';
				if (paramName === 'filePath') return '/test/file.xlsx';
				return undefined;
			});

			const getWorksheets = excelAI.methods?.loadOptions?.getWorksheets;
			expect(getWorksheets).toBeDefined();

			if (getWorksheets) {
				const result = await getWorksheets.call(mockLoadOptionsContext);
				expect(result).toBeDefined();
				expect(Array.isArray(result)).toBe(true);
				expect(result.length).toBe(3);
				expect(result).toContainEqual(expect.objectContaining({ value: 'Sheet1' }));
				expect(result).toContainEqual(expect.objectContaining({ value: 'Sheet2' }));
				expect(result).toContainEqual(expect.objectContaining({ value: 'DataSheet' }));
			}
		});

		test('getColumns should return columns list', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['Name', 'Age', 'Department', 'Email']);
			sheet.addRow(['John', 30, 'IT', 'john@example.com']);
			
			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockLoadOptionsContext.getNodeParameter.mockImplementation((paramName: string) => {
				if (paramName === 'inputMode') return 'filePath';
				if (paramName === 'filePath') return '/test/file.xlsx';
				if (paramName === 'sheetNameOptions') return 'TestSheet';
				if (paramName === 'sheetName') return 'TestSheet';
				return undefined;
			});

			const getColumns = excelAI.methods?.loadOptions?.getColumns;
			expect(getColumns).toBeDefined();

			if (getColumns) {
				const result = await getColumns.call(mockLoadOptionsContext);
				expect(result).toBeDefined();
				expect(Array.isArray(result)).toBe(true);
				expect(result.length).toBe(4);
				expect(result).toContainEqual(expect.objectContaining({ value: 'Name' }));
				expect(result).toContainEqual(expect.objectContaining({ value: 'Age' }));
			}
		});
	});

	// ===== 3. Row Operations Tests (File Path Mode) =====
	describe('Row Operations - File Path Mode', () => {
		test('Append Row - should append new row', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['Name', 'Age', 'Department']);
			sheet.addRow(['John', 30, 'IT']);
			
			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockContext.getNodeParameter.mockImplementation((paramName: string) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'appendRow',
					inputMode: 'filePath',
					filePath: '/test/employees.xlsx',
					sheetNameOptions: 'TestSheet',
					rowData: JSON.stringify({ Name: 'Jane', Age: 25, Department: 'Sales' }),
				};
				return params[paramName];
			});

			const result = await excelAI.execute.call(mockContext);

			expect(result).toBeDefined();
			expect(result[0][0].json).toHaveProperty('success', true);
			expect(result[0][0].json).toHaveProperty('operation', 'appendRow');
		});

		test('Read Rows - should read all rows', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['Product', 'Price', 'Stock']);
			sheet.addRow(['Laptop', 30000, 15]);
			sheet.addRow(['Mouse', 500, 100]);
			sheet.addRow(['Keyboard', 1500, 50]);
			
			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'readRows',
					inputMode: 'filePath',
					filePath: '/test/products.xlsx',
					sheetNameOptions: 'TestSheet',
					startRow: 2,
					endRow: 0,
				};
				return params[paramName] ?? fallback;
			});

			const result = await excelAI.execute.call(mockContext);

			expect(result).toBeDefined();
			expect(Array.isArray(result[0])).toBe(true);
			expect(result[0].length).toBe(3);
			expect(result[0][0].json).toHaveProperty('Product', 'Laptop');
			expect(result[0][1].json).toHaveProperty('Product', 'Mouse');
			expect(result[0][2].json).toHaveProperty('Product', 'Keyboard');
		});

		test('Update Row - should update specified row', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['EmployeeID', 'Name', 'Position']);
			sheet.addRow(['E001', 'John', 'Engineer']);
			
			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockContext.getNodeParameter.mockImplementation((paramName: string) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'updateRow',
					inputMode: 'filePath',
					filePath: '/test/employees.xlsx',
					sheetNameOptions: 'TestSheet',
					rowNumber: 2,
					updatedData: JSON.stringify({ Position: 'Senior Engineer', Name: 'John Doe' }),
				};
				return params[paramName];
			});

			const result = await excelAI.execute.call(mockContext);

			expect(result).toBeDefined();
			expect(result[0][0].json).toHaveProperty('success', true);
			expect(result[0][0].json).toHaveProperty('operation', 'updateRow');
			expect(result[0][0].json).toHaveProperty('rowNumber', 2);
		});

		test('Delete Row - should delete specified row', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['ID', 'Name']);
			sheet.addRow([1, 'ItemA']);
			sheet.addRow([2, 'ItemB']);
			
			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockContext.getNodeParameter.mockImplementation((paramName: string) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'deleteRow',
					inputMode: 'filePath',
					filePath: '/test/data.xlsx',
					sheetNameOptions: 'TestSheet',
					rowNumber: 2,
				};
				return params[paramName];
			});

			const result = await excelAI.execute.call(mockContext);

			expect(result).toBeDefined();
			expect(result[0][0].json).toHaveProperty('success', true);
			expect(result[0][0].json).toHaveProperty('operation', 'deleteRow');
		});

		test('Insert Row - should insert row at specified position', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['Order', 'Item']);
			sheet.addRow([1, 'First']);
			sheet.addRow([3, 'Third']);
			
			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockContext.getNodeParameter.mockImplementation((paramName: string) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'insertRow',
					inputMode: 'filePath',
					filePath: '/test/data.xlsx',
					sheetNameOptions: 'TestSheet',
					rowNumber: 3,
					rowData: JSON.stringify({ Order: 2, Item: 'Second' }),
				};
				return params[paramName];
			});

			const result = await excelAI.execute.call(mockContext);

			expect(result).toBeDefined();
			expect(result[0][0].json).toHaveProperty('success', true);
			expect(result[0][0].json).toHaveProperty('operation', 'insertRow');
		});
		test('Read Rows - should read from specific worksheet', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			
			// Create multiple worksheets
			const sheet1 = testWorkbook.addWorksheet('Sheet1');
			sheet1.addRow(['Name', 'Age']);
			sheet1.addRow(['Alice', 25]);
			sheet1.addRow(['Bob', 30]);
			
			const sheet2 = testWorkbook.addWorksheet('Sheet2');
			sheet2.addRow(['Product', 'Price']);
			sheet2.addRow(['ItemA', 100]);
			sheet2.addRow(['ItemB', 200]);
			
			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockContext.getNodeParameter.mockImplementation((paramName: string) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'readRows',
					inputMode: 'filePath',
					filePath: '/test/data.xlsx',
					sheetNameOptions: 'Sheet2',  // Read from Sheet2
					startRow: 2,
					endRow: 0,
				};
				return params[paramName];
			});

			const result = await excelAI.execute.call(mockContext);

			expect(result).toBeDefined();
			expect(Array.isArray(result[0])).toBe(true);
			expect(result[0].length).toBe(2);
			// Should read from Sheet2, not Sheet1
			expect(result[0][0].json).toHaveProperty('Product', 'ItemA');
			expect(result[0][0].json).toHaveProperty('Price', 100);
			expect(result[0][1].json).toHaveProperty('Product', 'ItemB');
			expect(result[0][1].json).toHaveProperty('Price', 200);
		});
	});

	// ===== 3.5. Filter Rows Tests (File Path Mode) =====
	describe('Filter Rows - File Path Mode', () => {
		test('Filter Rows - should filter by single equals condition', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['Name', 'Department', 'Age', 'Status']);
			sheet.addRow(['Alice', 'Engineering', 30, 'Active']);
			sheet.addRow(['Bob', 'Sales', 25, 'Active']);
			sheet.addRow(['Charlie', 'Engineering', 35, 'Inactive']);
			sheet.addRow(['David', 'Engineering', 28, 'Active']);

			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'filterRows',
					inputMode: 'filePath',
					filePath: '/test/employees.xlsx',
					sheetNameOptions: 'TestSheet',
					filterConditions: {
						conditions: [
							{ field: 'Department', operator: 'equals', value: 'Engineering' }
						]
					},
					conditionLogic: 'and',
				};
				return params[paramName] ?? fallback;
			});

			const result = await excelAI.execute.call(mockContext);

			expect(result).toBeDefined();
			expect(Array.isArray(result[0])).toBe(true);
			expect(result[0].length).toBe(3);
			expect(result[0][0].json).toHaveProperty('Name', 'Alice');
			expect(result[0][1].json).toHaveProperty('Name', 'Charlie');
			expect(result[0][2].json).toHaveProperty('Name', 'David');
		});

		test('Filter Rows - should filter by multiple AND conditions', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['Name', 'Department', 'Age', 'Status']);
			sheet.addRow(['Alice', 'Engineering', 30, 'Active']);
			sheet.addRow(['Bob', 'Sales', 25, 'Active']);
			sheet.addRow(['Charlie', 'Engineering', 35, 'Inactive']);
			sheet.addRow(['David', 'Engineering', 28, 'Active']);

			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'filterRows',
					inputMode: 'filePath',
					filePath: '/test/employees.xlsx',
					sheetNameOptions: 'TestSheet',
					filterConditions: {
						conditions: [
							{ field: 'Department', operator: 'equals', value: 'Engineering' },
							{ field: 'Status', operator: 'equals', value: 'Active' }
						]
					},
					conditionLogic: 'and',
				};
				return params[paramName] ?? fallback;
			});

			const result = await excelAI.execute.call(mockContext);

			expect(result).toBeDefined();
			expect(Array.isArray(result[0])).toBe(true);
			expect(result[0].length).toBe(2);
			expect(result[0][0].json).toHaveProperty('Name', 'Alice');
			expect(result[0][1].json).toHaveProperty('Name', 'David');
		});

		test('Filter Rows - should filter by multiple OR conditions', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['Name', 'Department', 'Salary']);
			sheet.addRow(['Alice', 'Engineering', 80000]);
			sheet.addRow(['Bob', 'Sales', 60000]);
			sheet.addRow(['Charlie', 'Marketing', 70000]);
			sheet.addRow(['David', 'Engineering', 75000]);

			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'filterRows',
					inputMode: 'filePath',
					filePath: '/test/employees.xlsx',
					sheetNameOptions: 'TestSheet',
					filterConditions: {
						conditions: [
							{ field: 'Department', operator: 'equals', value: 'Sales' },
							{ field: 'Department', operator: 'equals', value: 'Marketing' }
						]
					},
					conditionLogic: 'or',
				};
				return params[paramName] ?? fallback;
			});

			const result = await excelAI.execute.call(mockContext);

			expect(result).toBeDefined();
			expect(Array.isArray(result[0])).toBe(true);
			expect(result[0].length).toBe(2);
			expect(result[0][0].json).toHaveProperty('Name', 'Bob');
			expect(result[0][1].json).toHaveProperty('Name', 'Charlie');
		});

		test('Filter Rows - should filter using greater than operator', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['Product', 'Price', 'Stock']);
			sheet.addRow(['Laptop', 1000, 15]);
			sheet.addRow(['Mouse', 25, 100]);
			sheet.addRow(['Keyboard', 75, 50]);
			sheet.addRow(['Monitor', 300, 30]);

			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'filterRows',
					inputMode: 'filePath',
					filePath: '/test/products.xlsx',
					sheetNameOptions: 'TestSheet',
					filterConditions: {
						conditions: [
							{ field: 'Price', operator: 'greaterThan', value: '100' }
						]
					},
					conditionLogic: 'and',
				};
				return params[paramName] ?? fallback;
			});

			const result = await excelAI.execute.call(mockContext);

			expect(result).toBeDefined();
			expect(Array.isArray(result[0])).toBe(true);
			expect(result[0].length).toBe(2);
			expect(result[0][0].json).toHaveProperty('Product', 'Laptop');
			expect(result[0][1].json).toHaveProperty('Product', 'Monitor');
		});

		test('Filter Rows - should filter using less or equal operator', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['Employee', 'Age', 'Experience']);
			sheet.addRow(['Alice', 28, 5]);
			sheet.addRow(['Bob', 35, 10]);
			sheet.addRow(['Charlie', 25, 3]);
			sheet.addRow(['David', 30, 7]);

			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'filterRows',
					inputMode: 'filePath',
					filePath: '/test/employees.xlsx',
					sheetNameOptions: 'TestSheet',
					filterConditions: {
						conditions: [
							{ field: 'Age', operator: 'lessOrEqual', value: '30' }
						]
					},
					conditionLogic: 'and',
				};
				return params[paramName] ?? fallback;
			});

			const result = await excelAI.execute.call(mockContext);

			expect(result).toBeDefined();
			expect(Array.isArray(result[0])).toBe(true);
			expect(result[0].length).toBe(3);
			expect(result[0][0].json).toHaveProperty('Employee', 'Alice');
			expect(result[0][1].json).toHaveProperty('Employee', 'Charlie');
			expect(result[0][2].json).toHaveProperty('Employee', 'David');
		});

		test('Filter Rows - should filter using contains operator', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['Name', 'Email', 'Role']);
			sheet.addRow(['Alice Johnson', 'alice@company.com', 'Engineer']);
			sheet.addRow(['Bob Smith', 'bob@company.com', 'Designer']);
			sheet.addRow(['Charlie Davis', 'charlie@external.com', 'Consultant']);
			sheet.addRow(['David Brown', 'david@company.com', 'Manager']);

			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'filterRows',
					inputMode: 'filePath',
					filePath: '/test/employees.xlsx',
					sheetNameOptions: 'TestSheet',
					filterConditions: {
						conditions: [
							{ field: 'Email', operator: 'contains', value: '@company.com' }
						]
					},
					conditionLogic: 'and',
				};
				return params[paramName] ?? fallback;
			});

			const result = await excelAI.execute.call(mockContext);

			expect(result).toBeDefined();
			expect(Array.isArray(result[0])).toBe(true);
			expect(result[0].length).toBe(3);
			expect(result[0][0].json).toHaveProperty('Name', 'Alice Johnson');
			expect(result[0][1].json).toHaveProperty('Name', 'Bob Smith');
			expect(result[0][2].json).toHaveProperty('Name', 'David Brown');
		});

		test('Filter Rows - should filter using startsWith operator', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['Code', 'Description', 'Status']);
			sheet.addRow(['ENG-001', 'Engine Part', 'Active']);
			sheet.addRow(['ENG-002', 'Engine Assembly', 'Active']);
			sheet.addRow(['BRK-001', 'Brake System', 'Inactive']);
			sheet.addRow(['ENG-003', 'Engine Cover', 'Active']);

			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'filterRows',
					inputMode: 'filePath',
					filePath: '/test/parts.xlsx',
					sheetNameOptions: 'TestSheet',
					filterConditions: {
						conditions: [
							{ field: 'Code', operator: 'startsWith', value: 'ENG' }
						]
					},
					conditionLogic: 'and',
				};
				return params[paramName] ?? fallback;
			});

			const result = await excelAI.execute.call(mockContext);

			expect(result).toBeDefined();
			expect(Array.isArray(result[0])).toBe(true);
			expect(result[0].length).toBe(3);
			expect(result[0][0].json).toHaveProperty('Code', 'ENG-001');
			expect(result[0][1].json).toHaveProperty('Code', 'ENG-002');
			expect(result[0][2].json).toHaveProperty('Code', 'ENG-003');
		});

		test('Filter Rows - should filter using endsWith operator', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['Filename', 'Size', 'Type']);
			sheet.addRow(['document.pdf', 1024, 'PDF']);
			sheet.addRow(['image.png', 2048, 'Image']);
			sheet.addRow(['report.pdf', 4096, 'PDF']);
			sheet.addRow(['data.xlsx', 8192, 'Spreadsheet']);

			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'filterRows',
					inputMode: 'filePath',
					filePath: '/test/files.xlsx',
					sheetNameOptions: 'TestSheet',
					filterConditions: {
						conditions: [
							{ field: 'Filename', operator: 'endsWith', value: '.pdf' }
						]
					},
					conditionLogic: 'and',
				};
				return params[paramName] ?? fallback;
			});

			const result = await excelAI.execute.call(mockContext);

			expect(result).toBeDefined();
			expect(Array.isArray(result[0])).toBe(true);
			expect(result[0].length).toBe(2);
			expect(result[0][0].json).toHaveProperty('Filename', 'document.pdf');
			expect(result[0][1].json).toHaveProperty('Filename', 'report.pdf');
		});

		test('Filter Rows - should filter using isEmpty operator', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['Name', 'Phone', 'Notes']);
			sheet.addRow(['Alice', '123-456-7890', 'Has account']);
			sheet.addRow(['Bob', '', 'Needs follow-up']);
			sheet.addRow(['Charlie', null, '']);
			sheet.addRow(['David', '555-1234', 'VIP customer']);

			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'filterRows',
					inputMode: 'filePath',
					filePath: '/test/contacts.xlsx',
					sheetNameOptions: 'TestSheet',
					filterConditions: {
						conditions: [
							{ field: 'Phone', operator: 'isEmpty', value: '' }
						]
					},
					conditionLogic: 'and',
				};
				return params[paramName] ?? fallback;
			});

			const result = await excelAI.execute.call(mockContext);

			expect(result).toBeDefined();
			expect(Array.isArray(result[0])).toBe(true);
			expect(result[0].length).toBe(2);
			expect(result[0][0].json).toHaveProperty('Name', 'Bob');
			expect(result[0][1].json).toHaveProperty('Name', 'Charlie');
		});

		test('Filter Rows - should filter using isNotEmpty operator', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['Task', 'Assignee', 'DueDate']);
			sheet.addRow(['Task 1', 'Alice', '2024-01-01']);
			sheet.addRow(['Task 2', '', '2024-01-02']);
			sheet.addRow(['Task 3', 'Bob', '']);
			sheet.addRow(['Task 4', 'Charlie', '2024-01-03']);

			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'filterRows',
					inputMode: 'filePath',
					filePath: '/test/tasks.xlsx',
					sheetNameOptions: 'TestSheet',
					filterConditions: {
						conditions: [
							{ field: 'Assignee', operator: 'isNotEmpty', value: '' }
						]
					},
					conditionLogic: 'and',
				};
				return params[paramName] ?? fallback;
			});

			const result = await excelAI.execute.call(mockContext);

			expect(result).toBeDefined();
			expect(Array.isArray(result[0])).toBe(true);
			expect(result[0].length).toBe(3);
			expect(result[0][0].json).toHaveProperty('Task', 'Task 1');
			expect(result[0][1].json).toHaveProperty('Task', 'Task 3');
			expect(result[0][2].json).toHaveProperty('Task', 'Task 4');
		});

		test('Filter Rows - should filter using notEquals operator', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['Item', 'Status', 'Priority']);
			sheet.addRow(['Item 1', 'Completed', 'High']);
			sheet.addRow(['Item 2', 'Pending', 'Medium']);
			sheet.addRow(['Item 3', 'Completed', 'Low']);
			sheet.addRow(['Item 4', 'In Progress', 'High']);

			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'filterRows',
					inputMode: 'filePath',
					filePath: '/test/items.xlsx',
					sheetNameOptions: 'TestSheet',
					filterConditions: {
						conditions: [
							{ field: 'Status', operator: 'notEquals', value: 'Completed' }
						]
					},
					conditionLogic: 'and',
				};
				return params[paramName] ?? fallback;
			});

			const result = await excelAI.execute.call(mockContext);

			expect(result).toBeDefined();
			expect(Array.isArray(result[0])).toBe(true);
			expect(result[0].length).toBe(2);
			expect(result[0][0].json).toHaveProperty('Item', 'Item 2');
			expect(result[0][1].json).toHaveProperty('Item', 'Item 4');
		});

		test('Filter Rows - should filter using notContains operator', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['Title', 'Tags', 'Category']);
			sheet.addRow(['Article 1', 'tech, ai', 'Technology']);
			sheet.addRow(['Article 2', 'business, finance', 'Business']);
			sheet.addRow(['Article 3', 'tech, web', 'Technology']);
			sheet.addRow(['Article 4', 'sports, news', 'Sports']);

			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'filterRows',
					inputMode: 'filePath',
					filePath: '/test/articles.xlsx',
					sheetNameOptions: 'TestSheet',
					filterConditions: {
						conditions: [
							{ field: 'Tags', operator: 'notContains', value: 'tech' }
						]
					},
					conditionLogic: 'and',
				};
				return params[paramName] ?? fallback;
			});

			const result = await excelAI.execute.call(mockContext);

			expect(result).toBeDefined();
			expect(Array.isArray(result[0])).toBe(true);
			expect(result[0].length).toBe(2);
			expect(result[0][0].json).toHaveProperty('Title', 'Article 2');
			expect(result[0][1].json).toHaveProperty('Title', 'Article 4');
		});

		test('Filter Rows - should return all rows when no conditions', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['Name', 'Age']);
			sheet.addRow(['Alice', 30]);
			sheet.addRow(['Bob', 25]);
			sheet.addRow(['Charlie', 35]);

			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'filterRows',
					inputMode: 'filePath',
					filePath: '/test/data.xlsx',
					sheetNameOptions: 'TestSheet',
					filterConditions: { conditions: [] },
					conditionLogic: 'and',
				};
				return params[paramName] ?? fallback;
			});

			const result = await excelAI.execute.call(mockContext);

			expect(result).toBeDefined();
			expect(Array.isArray(result[0])).toBe(true);
			expect(result[0].length).toBe(3);
		});

		test('Filter Rows - should include row numbers in results', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['Product', 'Stock']);
			sheet.addRow(['Item A', 10]);
			sheet.addRow(['Item B', 0]);
			sheet.addRow(['Item C', 25]);
			sheet.addRow(['Item D', 5]);

			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'filterRows',
					inputMode: 'filePath',
					filePath: '/test/inventory.xlsx',
					sheetNameOptions: 'TestSheet',
					filterConditions: {
						conditions: [
							{ field: 'Stock', operator: 'greaterThan', value: '0' }
						]
					},
					conditionLogic: 'and',
				};
				return params[paramName] ?? fallback;
			});

			const result = await excelAI.execute.call(mockContext);

			expect(result).toBeDefined();
			expect(Array.isArray(result[0])).toBe(true);
			expect(result[0].length).toBe(3);
			expect(result[0][0].json).toHaveProperty('_rowNumber');
			expect(result[0][0].json._rowNumber).toBe(2);
			expect(result[0][1].json._rowNumber).toBe(4);
			expect(result[0][2].json._rowNumber).toBe(5);
		});

		test('Filter Rows - should not duplicate results with multiple input items', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('Employees');
			sheet.addRow(['Name', 'Department', 'Age', 'Salary']);
			sheet.addRow(['Alice', 'Engineering', 28, 75000]);
			sheet.addRow(['Bob', 'Sales', 35, 65000]);
			sheet.addRow(['Carol', 'Engineering', 42, 80000]);
			sheet.addRow(['David', 'Engineering', 31, 72000]);
			sheet.addRow(['Emma', 'Sales', 26, 58000]);

			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			// Simulate 5 input items (like in the screenshot)
			mockContext.getInputData.mockReturnValue([
				{ json: { id: 1 } },
				{ json: { id: 2 } },
				{ json: { id: 3 } },
				{ json: { id: 4 } },
				{ json: { id: 5 } }
			]);

			mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'filterRows',
					inputMode: 'filePath',
					filePath: '/data/test_employees.xlsx',
					sheetNameOptions: 'Employees',
					filterConditions: {
						conditions: [
							{ field: 'Department', operator: 'equals', value: 'Engineering' }
						]
					},
					conditionLogic: 'and',
				};
				return params[paramName] ?? fallback;
			});

			const result = await excelAI.execute.call(mockContext);

			// Should only return 3 filtered results (not 15 = 3 * 5)
			expect(result).toBeDefined();
			expect(Array.isArray(result[0])).toBe(true);
			expect(result[0].length).toBe(3);
			expect(result[0][0].json).toHaveProperty('Name', 'Alice');
			expect(result[0][0].json).toHaveProperty('Department', 'Engineering');
			expect(result[0][1].json).toHaveProperty('Name', 'Carol');
			expect(result[0][1].json).toHaveProperty('Department', 'Engineering');
			expect(result[0][2].json).toHaveProperty('Name', 'David');
			expect(result[0][2].json).toHaveProperty('Department', 'Engineering');
		});
	});

	// ===== 3.6. Filter Rows Tests (Binary Data Mode) =====
	describe('Filter Rows - Binary Data Mode', () => {
		test('Filter Rows - should filter with binary data input', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['Name', 'Department', 'Salary']);
			sheet.addRow(['Alice', 'IT', 75000]);
			sheet.addRow(['Bob', 'HR', 60000]);
			sheet.addRow(['Charlie', 'IT', 80000]);
			sheet.addRow(['David', 'Sales', 65000]);

			const buffer = await testWorkbook.xlsx.writeBuffer();
			(mockContext.helpers.getBinaryDataBuffer as jest.Mock).mockResolvedValue(buffer);

			mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'filterRows',
					inputMode: 'binaryData',
					binaryPropertyName: 'data',
					sheetName: 'TestSheet',
					filterConditionsBinary: {
						conditions: [
							{ field: 'Department', operator: 'equals', value: 'IT' }
						]
					},
					conditionLogic: 'and',
				};
				return params[paramName] ?? fallback;
			});

			const result = await excelAI.execute.call(mockContext);

			expect(result).toBeDefined();
			expect(Array.isArray(result[0])).toBe(true);
			expect(result[0].length).toBe(2);
			expect(result[0][0].json).toHaveProperty('Name', 'Alice');
			expect(result[0][1].json).toHaveProperty('Name', 'Charlie');
		});

		test('Filter Rows - should filter with complex conditions in binary mode', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['Product', 'Category', 'Price', 'InStock']);
			sheet.addRow(['Laptop', 'Electronics', 1200, 'Yes']);
			sheet.addRow(['Mouse', 'Electronics', 25, 'Yes']);
			sheet.addRow(['Chair', 'Furniture', 300, 'No']);
			sheet.addRow(['Monitor', 'Electronics', 350, 'Yes']);

			const buffer = await testWorkbook.xlsx.writeBuffer();
			(mockContext.helpers.getBinaryDataBuffer as jest.Mock).mockResolvedValue(buffer);

			mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'filterRows',
					inputMode: 'binaryData',
					binaryPropertyName: 'data',
					sheetName: 'TestSheet',
					filterConditionsBinary: {
						conditions: [
							{ field: 'Category', operator: 'equals', value: 'Electronics' },
							{ field: 'Price', operator: 'lessThan', value: '1000' }
						]
					},
					conditionLogic: 'and',
				};
				return params[paramName] ?? fallback;
			});

			const result = await excelAI.execute.call(mockContext);

			expect(result).toBeDefined();
			expect(Array.isArray(result[0])).toBe(true);
			expect(result[0].length).toBe(2);
			expect(result[0][0].json).toHaveProperty('Product', 'Mouse');
			expect(result[0][1].json).toHaveProperty('Product', 'Monitor');
		});
	});

	// ===== 3.7. Error Handling for Row Operations =====
	describe('Row Operations - Error Handling', () => {
		test('Append Row - invalid JSON should throw error', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['Name', 'Age']);

			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockContext.getNodeParameter.mockImplementation((paramName: string) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'appendRow',
					inputMode: 'filePath',
					filePath: '/test/employees.xlsx',
					sheetNameOptions: 'TestSheet',
					rowData: '{invalid-json}',
				};
				return params[paramName];
			});

			await expect(excelAI.execute.call(mockContext)).rejects.toThrow();
		});
	});

	// ===== 4. Worksheet Operations Tests =====
	describe('Worksheet Operations', () => {
	test('Create Worksheet - should create new worksheet', async () => {
		const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
		testWorkbook.addWorksheet('ExistingSheet');
		
		const buffer = await testWorkbook.xlsx.writeBuffer();
		(global as any).__mockExcelBuffer__ = buffer;

		mockContext.getNodeParameter.mockImplementation((paramName: string) => {
			const params: Record<string, any> = {
				resource: 'worksheet',
				worksheetOperation: 'createWorksheet',
				inputMode: 'filePath',
				filePath: '/test/workbook.xlsx',
				newSheetName: 'NewSheet',
				initialData: '[]',
			};
			return params[paramName];
		});

		const result = await excelAI.execute.call(mockContext);

		expect(result).toBeDefined();
		expect(result[0][0].json).toHaveProperty('success', true);
		expect(result[0][0].json).toHaveProperty('operation', 'createWorksheet');
	});

	test('Delete Worksheet - should delete specified worksheet', async () => {
		const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
		testWorkbook.addWorksheet('Sheet1');
		testWorkbook.addWorksheet('ToDelete');
		
		const buffer = await testWorkbook.xlsx.writeBuffer();
		(global as any).__mockExcelBuffer__ = buffer;

		mockContext.getNodeParameter.mockImplementation((paramName: string) => {
			const params: Record<string, any> = {
				resource: 'worksheet',
				worksheetOperation: 'deleteWorksheet',
				inputMode: 'filePath',
				filePath: '/test/workbook.xlsx',
				sheetNameOptions: 'ToDelete',
			};
			return params[paramName];
		});

		const result = await excelAI.execute.call(mockContext);

		expect(result).toBeDefined();
		expect(result[0][0].json).toHaveProperty('success', true);
		expect(result[0][0].json).toHaveProperty('operation', 'deleteWorksheet');
	});

	test('List Worksheets - should list all worksheets', async () => {
		const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
		testWorkbook.addWorksheet('Employees');
		testWorkbook.addWorksheet('Products');
		testWorkbook.addWorksheet('Sales');
		
		const buffer = await testWorkbook.xlsx.writeBuffer();
		(global as any).__mockExcelBuffer__ = buffer;

		mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
			const params: Record<string, any> = {
				resource: 'worksheet',
				worksheetOperation: 'listWorksheets',
				inputMode: 'filePath',
				filePath: '/test/workbook.xlsx',
				includeHidden: false,
			};
			return params[paramName] ?? fallback;
		});

		const result = await excelAI.execute.call(mockContext);

		expect(result).toBeDefined();
		expect(result[0][0].json).toHaveProperty('operation', 'listWorksheets');
		expect(result[0][0].json).toHaveProperty('worksheets');
		expect(Array.isArray(result[0][0].json.worksheets)).toBe(true);
	});

	test('Rename Worksheet - should rename worksheet successfully', async () => {
	const testWorkbook = new ExcelJS.Workbook();
	testWorkbook.addWorksheet('OldSheet');
	
	const buffer = await testWorkbook.xlsx.writeBuffer();
	(global as any).__mockExcelBuffer__ = buffer;
	(fs.readFile as jest.Mock).mockResolvedValue(buffer);
	(fs.writeFile as jest.Mock).mockResolvedValue(undefined);

	mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
		const params: Record<string, any> = {
			resource: 'worksheet',
			worksheetOperation: 'renameWorksheet',
			inputMode: 'filePath',
			filePath: '/test/file.xlsx',
			worksheetNameOptions: 'OldSheet',
			newSheetName: 'NewSheet',
			autoSave: true,
		};
		return params[paramName] ?? fallback;
	});

	const result = await excelAI.execute.call(mockContext);

	expect(result).toBeDefined();
	expect(result[0][0].json).toHaveProperty('success', true);
	expect(result[0][0].json).toHaveProperty('operation', 'renameWorksheet');
	expect(result[0][0].json).toHaveProperty('oldName', 'OldSheet');
	expect(result[0][0].json).toHaveProperty('newName', 'NewSheet');
	});

	test('Copy Worksheet - should copy worksheet successfully', async () => {
	const testWorkbook = new ExcelJS.Workbook();
	const sourceSheet = testWorkbook.addWorksheet('SourceSheet');
	sourceSheet.addRow(['Name', 'Age']);
	sourceSheet.addRow(['Alice', 30]);
	sourceSheet.addRow(['Bob', 25]);
	
	const buffer = await testWorkbook.xlsx.writeBuffer();
	(global as any).__mockExcelBuffer__ = buffer;
	(fs.readFile as jest.Mock).mockResolvedValue(buffer);
	(fs.writeFile as jest.Mock).mockResolvedValue(undefined);

	mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
		const params: Record<string, any> = {
			resource: 'worksheet',
			worksheetOperation: 'copyWorksheet',
			inputMode: 'filePath',
			filePath: '/test/file.xlsx',
			worksheetNameOptions: 'SourceSheet',
			newSheetName: 'CopiedSheet',
			autoSave: true,
		};
		return params[paramName] ?? fallback;
	});

	const result = await excelAI.execute.call(mockContext);

	expect(result).toBeDefined();
	expect(result[0][0].json).toHaveProperty('success', true);
	expect(result[0][0].json).toHaveProperty('operation', 'copyWorksheet');
	expect(result[0][0].json).toHaveProperty('sourceName', 'SourceSheet');
	expect(result[0][0].json).toHaveProperty('newName', 'CopiedSheet');
	expect(result[0][0].json).toHaveProperty('rowCount');
	});

	test('Get Worksheet Info - should return worksheet details', async () => {
	const testWorkbook = new ExcelJS.Workbook();
	const sheet = testWorkbook.addWorksheet('DataSheet');
	sheet.addRow(['UserID', 'Name', 'Email']);
	sheet.addRow([1, 'Alice', 'alice@example.com']);
	sheet.addRow([2, 'Bob', 'bob@example.com']);
	sheet.getColumn(1).width = 15;
	sheet.getColumn(2).width = 25;
	sheet.getColumn(3).width = 30;
	
	const buffer = await testWorkbook.xlsx.writeBuffer();
	(global as any).__mockExcelBuffer__ = buffer;
	(fs.readFile as jest.Mock).mockResolvedValue(buffer);

	mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
		const params: Record<string, any> = {
			resource: 'worksheet',
			worksheetOperation: 'getWorksheetInfo',
			inputMode: 'filePath',
			filePath: '/test/file.xlsx',
			worksheetNameOptions: 'DataSheet',
		};
		return params[paramName] ?? fallback;
	});

	const result = await excelAI.execute.call(mockContext);

	expect(result).toBeDefined();
	expect(result[0][0].json).toHaveProperty('operation', 'getWorksheetInfo');
	expect(result[0][0].json).toHaveProperty('sheetName', 'DataSheet');
	expect(result[0][0].json).toHaveProperty('rowCount');
	expect(result[0][0].json).toHaveProperty('columnCount');
	expect(result[0][0].json).toHaveProperty('columns');
	const columns = result[0][0].json.columns as any[];
	expect(Array.isArray(columns)).toBe(true);
	expect(columns.length).toBe(3);
	expect(columns[0]).toHaveProperty('index', 1);
	expect(columns[0]).toHaveProperty('letter', 'A');
	expect(columns[0]).toHaveProperty('header', 'UserID');
	expect(columns[0]).toHaveProperty('width', 15);
	});
	});

	// ===== 5. Error Handling Tests =====
	describe('Error Handling', () => {
	test('should handle worksheet not found error', async () => {
		const testWorkbook = new ExcelJS.Workbook();
		testWorkbook.addWorksheet('ExistingSheet');
		
		const buffer = await testWorkbook.xlsx.writeBuffer();
		(fs.readFile as jest.Mock).mockResolvedValue(buffer);

		mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
			const params: Record<string, any> = {
				resource: 'row',
				operation: 'readRows',
				inputMode: 'filePath',
				filePath: '/test/file.xlsx',
				sheetNameOptions: 'NonExistentSheet',
				startRow: 2,
				endRow: 0,
			};
			return params[paramName] ?? fallback;
		});

		await expect(excelAI.execute.call(mockContext)).rejects.toThrow();
	});

	test('should handle file read failure error', async () => {
		(fs.readFile as jest.Mock).mockRejectedValue(new Error('ENOENT: no such file or directory'));

		mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
			const params: Record<string, any> = {
				resource: 'row',
				operation: 'readRows',
				inputMode: 'filePath',
				filePath: '/nonexistent/file.xlsx',
				sheetNameOptions: 'Sheet1',
				startRow: 2,
				endRow: 0,
			};
			return params[paramName] ?? fallback;
		});

		await expect(excelAI.execute.call(mockContext)).rejects.toThrow();
	});

	test('Continue on Fail mode should return error instead of throwing', async () => {
		mockContext.continueOnFail.mockReturnValue(true);
		(fs.readFile as jest.Mock).mockRejectedValue(new Error('File not found'));

		mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
			const params: Record<string, any> = {
				resource: 'row',
				operation: 'readRows',
				inputMode: 'filePath',
				filePath: '/nonexistent/file.xlsx',
				sheetNameOptions: 'Sheet1',
				startRow: 2,
				endRow: 0,
			};
			return params[paramName] ?? fallback;
		});

		const result = await excelAI.execute.call(mockContext);

		expect(result).toBeDefined();
		expect(result[0][0].json).toHaveProperty('error');
	});
	});

	// ===== 6. Binary Data Mode Tests =====
	describe('Binary Data Mode', () => {
	test('should read Excel file from Binary Data', async () => {
		const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
		const sheet = testWorkbook.addWorksheet('TestSheet');
		sheet.addRow(['Name', 'Age']);
		sheet.addRow(['John', 30]);
		sheet.addRow(['Jane', 25]);
		
		const buffer = await testWorkbook.xlsx.writeBuffer();
		
		(mockContext.helpers.getBinaryDataBuffer as jest.Mock).mockResolvedValue(buffer);

		mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
			const params: Record<string, any> = {
				resource: 'row',
				operation: 'readRows',
				inputMode: 'binaryData',
				binaryPropertyName: 'data',
				sheetName: 'TestSheet',
				startRow: 2,
				endRow: 0,
			};
			return params[paramName] ?? fallback;
		});

		const result = await excelAI.execute.call(mockContext);

		expect(result).toBeDefined();
		expect(Array.isArray(result[0])).toBe(true);
		expect(result[0].length).toBe(2);
		expect(result[0][0].json).toHaveProperty('Name', 'John');
		expect(result[0][1].json).toHaveProperty('Name', 'Jane');
	});

	test('should append row in Binary Data mode', async () => {
		const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
		const sheet = testWorkbook.addWorksheet('TestSheet');
		sheet.addRow(['Product', 'Price']);
		sheet.addRow(['ProductA', 1000]);
		
		const buffer = await testWorkbook.xlsx.writeBuffer();
		(mockContext.helpers.getBinaryDataBuffer as jest.Mock).mockResolvedValue(buffer);

		mockContext.getNodeParameter.mockImplementation((paramName: string) => {
			const params: Record<string, any> = {
				resource: 'row',
				operation: 'appendRow',
				inputMode: 'binaryData',
				binaryPropertyName: 'data',
				sheetName: 'TestSheet',
				rowData: JSON.stringify({ Product: 'ProductB', Price: 2000 }),
			};
			return params[paramName];
		});

		const result = await excelAI.execute.call(mockContext);

		expect(result).toBeDefined();
		expect(result[0][0].json).toHaveProperty('success', true);
	});
	});
});






