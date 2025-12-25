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
			expect(operationValues).toContain('findRows');
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

		test('Find Rows - should return matching rows', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['Name', 'Role']);
			sheet.addRow(['Alice', 'Engineer']);
			sheet.addRow(['Bob', 'Designer']);
			sheet.addRow(['Alice', 'Lead']);

			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'findRows',
					inputMode: 'filePath',
					filePath: '/test/data.xlsx',
					sheetNameOptions: 'TestSheet',
					searchColumn: 'Name',
					searchValue: 'Alice',
					matchType: 'exact',
					returnRowNumbers: false,
				};
				return params[paramName] ?? fallback;
			});

			const result = await excelAI.execute.call(mockContext);

			expect(Array.isArray(result[0])).toBe(true);
			expect(result[0].length).toBe(2);
			expect(result[0][0].json).toHaveProperty('Name', 'Alice');
			expect(result[0][1].json).toHaveProperty('Name', 'Alice');
		});

		test('Find Rows - should return row numbers only', async () => {
			const testWorkbook = new (jest.requireActual('exceljs')).Workbook();
			const sheet = testWorkbook.addWorksheet('TestSheet');
			sheet.addRow(['Id', 'Status']);
			sheet.addRow([1, 'Open']);
			sheet.addRow([2, 'Closed']);
			sheet.addRow([3, 'Open']);

			const buffer = await testWorkbook.xlsx.writeBuffer();
			(global as any).__mockExcelBuffer__ = buffer;

			mockContext.getNodeParameter.mockImplementation((paramName: string, itemIndex: number, fallback?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					operation: 'findRows',
					inputMode: 'filePath',
					filePath: '/test/data.xlsx',
					sheetNameOptions: 'TestSheet',
					searchColumn: 'Status',
					searchValue: 'Open',
					matchType: 'exact',
					returnRowNumbers: true,
				};
				return params[paramName] ?? fallback;
			});

			const result = await excelAI.execute.call(mockContext);

			expect(result[0][0].json).toHaveProperty('rowNumbers');
			expect(result[0][0].json.rowNumbers).toEqual([2, 4]);
			expect(result[0][0].json).toHaveProperty('count', 2);
		});

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






