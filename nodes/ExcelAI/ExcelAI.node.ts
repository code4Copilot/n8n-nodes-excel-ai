import {
	IExecuteFunctions,
	ILoadOptionsFunctions,
	INodeExecutionData,
	INodePropertyOptions,
	INodeType,
	INodeTypeDescription,
	NodeOperationError,
} from 'n8n-workflow';
import * as ExcelJS from 'exceljs';
import * as fs from 'fs/promises';

export class ExcelAI implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Excel AI',
		name: 'excelAI',
		icon: 'file:excel.svg',
		group: ['transform'],
		version: 1,
		subtitle: '={{$parameter["resource"] + ": " + ($parameter["resource"] === "row" ? $parameter["operation"] : $parameter["worksheetOperation"])}}',
		description: 'Perform CRUD operations on Excel files with AI Agent support',
		defaults: {
			name: 'Excel AI',
		},
		inputs: ['main'],
		outputs: ['main'],
		
		// Enable AI Agent Integration
		usableAsTool: true,
		
		properties: [
			// Resource Selection
			{
				displayName: 'Resource',
				name: 'resource',
				type: 'options',
				noDataExpression: true,
				options: [
					{
						name: 'Row',
						value: 'row',
						description: 'Perform operations on rows',
					},
					{
						name: 'Worksheet',
						value: 'worksheet',
						description: 'Perform operations on worksheets',
					},
				],
				default: 'row',
			},

			// Input Mode
			{
				displayName: 'Input Mode',
				name: 'inputMode',
				type: 'options',
				options: [
					{
						name: 'File Path',
						value: 'filePath',
						description: 'Use file from file system',
					},
					{
						name: 'Binary Data',
						value: 'binaryData',
						description: 'Use file from binary data',
					},
				],
				default: 'filePath',
				description: 'How to provide the Excel file',
			},

			// File Path Input (AI-friendly)
			{
				displayName: 'File Path',
				name: 'filePath',
				type: 'string',
				displayOptions: {
					show: {
						inputMode: ['filePath'],
					},
				},
				default: '',
				required: true,
				placeholder: '/data/myfile.xlsx',
				description: 'Absolute path to the Excel file',
			},

			// Binary Property Name
			{
				displayName: 'Binary Property',
				name: 'binaryPropertyName',
				type: 'string',
				displayOptions: {
					show: {
						inputMode: ['binaryData'],
					},
				},
				default: 'data',
				required: true,
				description: 'Name of the binary property containing the Excel file',
			},

			// Row Operations
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				displayOptions: {
					show: {
						resource: ['row'],
					},
				},
				options: [
					{
						name: 'Read Rows',
						value: 'readRows',
						description: 'Read data from Excel file',
						action: 'Read rows from Excel',
					},
					{
						name: 'Append Row',
						value: 'appendRow',
						description: 'Add a new row at the end',
						action: 'Add row to Excel',
					},
					{
						name: 'Insert Row',
						value: 'insertRow',
						description: 'Insert a row at specific position',
						action: 'Insert row in Excel',
					},
					{
						name: 'Update Row',
						value: 'updateRow',
						description: 'Update an existing row',
						action: 'Update row in Excel',
					},
					{
						name: 'Delete Row',
						value: 'deleteRow',
						description: 'Delete a specific row',
						action: 'Delete row from Excel',
					},
					{
						name: 'Find Rows',
						value: 'findRows',
						description: 'Search for rows matching criteria',
						action: 'Find rows in Excel',
					},
				],
				default: 'readRows',
			},

			// Worksheet Operations
			{
				displayName: 'Worksheet Operation',
				name: 'worksheetOperation',
				type: 'options',
				displayOptions: {
					show: {
						resource: ['worksheet'],
					},
				},
				options: [
					{
						name: 'List Worksheets',
						value: 'listWorksheets',
						description: 'Get list of all worksheets',
						action: 'List worksheets in Excel',
					},
					{
						name: 'Create Worksheet',
						value: 'createWorksheet',
						description: 'Create a new worksheet',
						action: 'Create worksheet in Excel',
					},
					{
						name: 'Delete Worksheet',
						value: 'deleteWorksheet',
						description: 'Delete a worksheet',
						action: 'Delete worksheet from Excel',
					},
				{
					name: 'Rename Worksheet',
					value: 'renameWorksheet',
					description: 'Rename an existing worksheet',
					action: 'Rename worksheet in Excel',
				},
				{
					name: 'Copy Worksheet',
					value: 'copyWorksheet',
					description: 'Copy a worksheet to a new worksheet',
					action: 'Copy worksheet in Excel',
				},
				{
					name: 'Get Worksheet Info',
					value: 'getWorksheetInfo',
					description: 'Get detailed information about a worksheet including columns',
					action: 'Get worksheet info from Excel',
				},
			],
			default: 'listWorksheets',
		},

		// Sheet Name Options for File Path Mode (Row Operations)
		{
			displayName: 'Sheet Name',
			name: 'sheetNameOptions',
			type: 'options',
			displayOptions: {
				show: {
					inputMode: ['filePath'],
					resource: ['row'],
				},
			},
			typeOptions: {
				loadOptionsMethod: 'getWorksheets',
				loadOptionsDependsOn: ['filePath'],
			},
			default: '',
			required: false,
			description: 'Name of the worksheet to operate on. If not specified, the first worksheet will be used.',
		},


	// Sheet Name for Binary Mode
	{
		displayName: 'Sheet Name',
		name: 'sheetName',
		type: 'string',
		displayOptions: {
			show: {
				inputMode: ['binaryData'],
				resource: ['row'],
			},
		},
		default: '',
		required: false,
		description: 'Name of the worksheet to operate on. If not specified, the first worksheet will be used.',
	},

	// Read Rows Parameters
	{
				displayName: 'Start Row',
				name: 'startRow',
				type: 'number',
				displayOptions: {
					show: {
						resource: ['row'],
						operation: ['readRows'],
					},
				},
				default: 2,
				description: 'Row number to start reading from (1-based, row 1 is header)',
			},
			{
				displayName: 'End Row',
				name: 'endRow',
				type: 'number',
				displayOptions: {
					show: {
						resource: ['row'],
						operation: ['readRows'],
					},
				},
				default: 0,
				description: 'Row number to end reading (0 means read all rows)',
			},

			// Row Data (JSON) for Add/Insert
			{
				displayName: 'Row Data',
				name: 'rowData',
				type: 'json',
				displayOptions: {
					show: {
						resource: ['row'],
						operation: ['appendRow', 'insertRow'],
					},
				},
				default: '{}',
				required: true,
				placeholder: '{"Name": "John Doe", "Age": 30, "Email": "john@example.com"}',
				description: 'Row data as JSON object. Column names will be automatically mapped.',
			},

			// Row Number for Insert/Update/Delete
			{
				displayName: 'Row Number',
				name: 'rowNumber',
				type: 'number',
				displayOptions: {
					show: {
						resource: ['row'],
						operation: ['insertRow', 'updateRow', 'deleteRow'],
					},
				},
				default: 2,
				required: true,
				description: 'Target row number (1-based, row 1 is header)',
			},

			// Updated Data for Update
			{
				displayName: 'Updated Data',
				name: 'updatedData',
				type: 'json',
				displayOptions: {
					show: {
						resource: ['row'],
						operation: ['updateRow'],
					},
				},
				default: '{}',
				required: true,
				placeholder: '{"Email": "newemail@example.com", "Status": "Active"}',
				description: 'Data to update as JSON object. Only specified columns will be updated.',
			},

			// Find Rows Parameters
			{
				displayName: 'Search Column',
				name: 'searchColumn',
				type: 'options',
				typeOptions: {
					loadOptionsMethod: 'getColumns',
					loadOptionsDependsOn: ['filePath', 'sheetNameOptions'],
				},
				displayOptions: {
					show: {
						resource: ['row'],
						operation: ['findRows'],
					},
				},
				default: '',
				required: true,
				description: 'Column to search in',
			},
			{
				displayName: 'Search Value',
				name: 'searchValue',
				type: 'string',
				displayOptions: {
					show: {
						resource: ['row'],
						operation: ['findRows'],
					},
				},
				default: '',
				required: true,
				description: 'Value to search for',
			},
			{
				displayName: 'Match Type',
				name: 'matchType',
				type: 'options',
				options: [
					{
						name: 'Exact',
						value: 'exact',
						description: 'Exact match (case-insensitive)',
					},
					{
						name: 'Contains',
						value: 'contains',
						description: 'Contains substring (case-insensitive)',
					},
					{
						name: 'Starts With',
						value: 'startsWith',
						description: 'Starts with (case-insensitive)',
					},
					{
						name: 'Ends With',
						value: 'endsWith',
						description: 'Ends with (case-insensitive)',
					},
				],
				displayOptions: {
					show: {
						resource: ['row'],
						operation: ['findRows'],
					},
				},
				default: 'exact',
				description: 'How to match the search value',
			},
			{
				displayName: 'Return Row Numbers Only',
				name: 'returnRowNumbers',
				type: 'boolean',
				displayOptions: {
					show: {
						resource: ['row'],
						operation: ['findRows'],
					},
				},
				default: false,
				description: 'Whether to return only row numbers instead of full row data',
			},

			// Worksheet Parameters
			{
				displayName: 'Worksheet Name',
				name: 'worksheetNameOptions',
				type: 'options',
				typeOptions: {
					loadOptionsMethod: 'getWorksheets',
					loadOptionsDependsOn: ['filePath'],
				},
				displayOptions: {
					show: {
						resource: ['worksheet'],
						worksheetOperation: ['deleteWorksheet', 'renameWorksheet', 'copyWorksheet', 'getWorksheetInfo'],
						inputMode: ['filePath'],
					},
				},
				default: '',
				required: true,
				description: 'Name of the worksheet',
			},
			{
				displayName: 'New Sheet Name',
				name: 'newSheetName',
				type: 'string',
				displayOptions: {
					show: {
						resource: ['worksheet'],
						worksheetOperation: ['createWorksheet', 'renameWorksheet', 'copyWorksheet'],
					},
				},
				default: '',
				required: true,
				placeholder: 'NewSheet',
				description: 'Name for the new worksheet',
			},
			{
				displayName: 'Initial Data',
				name: 'initialData',
				type: 'json',
				displayOptions: {
					show: {
						resource: ['worksheet'],
						worksheetOperation: ['createWorksheet'],
					},
				},
				default: '[]',
				placeholder: '[{"Name": "John", "Age": 30}]',
				description: 'Initial data as array of objects (optional)',
			},
			{
				displayName: 'Include Hidden Sheets',
				name: 'includeHidden',
				type: 'boolean',
				displayOptions: {
					show: {
						resource: ['worksheet'],
						worksheetOperation: ['listWorksheets'],
					},
				},
				default: false,
				description: 'Whether to include hidden worksheets in the list',
			},

			// Auto Save Option
			{
				displayName: 'Auto Save',
				name: 'autoSave',
				type: 'boolean',
				displayOptions: {
					show: {
						inputMode: ['filePath'],
					},
				},
				default: true,
				description: 'Whether to automatically save changes to file',
			},
		],
	};

	methods = {
		loadOptions: {
			// Get available worksheets
			async getWorksheets(this: ILoadOptionsFunctions): Promise<INodePropertyOptions[]> {
				try {
					const inputMode = this.getNodeParameter('inputMode') as string;

					if (inputMode !== 'filePath') {
						return [
							{
								name: 'Sheet name only available in file path mode',
								value: '',
							},
						];
					}

					const filePath = this.getNodeParameter('filePath') as string;

					if (!filePath || filePath.trim() === '') {
						return [
							{
								name: 'Please specify file path first',
								value: '',
							},
						];
					}

					// Check file exists and is accessible
					await fs.access(filePath);
					
					const workbook = new ExcelJS.Workbook();
					await workbook.xlsx.readFile(filePath);

					const sheets = workbook.worksheets.map((sheet) => ({
						name: sheet.name,
						value: sheet.name,
					}));

					if (sheets.length === 0) {
						return [
							{
								name: 'No worksheets found in file',
								value: '',
							},
						];
					}

					return sheets;
				} catch (error) {
					const errorMessage = error instanceof Error ? error.message : 'Unknown error';
					return [
						{
							name: `Error: ${errorMessage}`,
							value: '',
						},
					];
				}
			},

			// Get available columns
			async getColumns(this: ILoadOptionsFunctions): Promise<INodePropertyOptions[]> {
				const inputMode = this.getNodeParameter('inputMode') as string;

				if (inputMode !== 'filePath') {
					return [
						{
							name: 'Column detection only available in file path mode',
							value: '',
						},
					];
				}

				const filePath = this.getNodeParameter('filePath') as string;
				const sheetName = this.getNodeParameter('sheetNameOptions', '') as string;

				if (!filePath || !sheetName) {
					return [
						{
							name: 'Please specify file path and sheet name first',
							value: '',
						},
					];
				}

				try {
					const workbook = new ExcelJS.Workbook();
					await workbook.xlsx.readFile(filePath);
					const worksheet = workbook.getWorksheet(sheetName);

					if (!worksheet) {
						return [
							{
								name: 'Sheet not found',
								value: '',
							},
						];
					}

					const headerRow = worksheet.getRow(1);
					const columns: INodePropertyOptions[] = [];

					headerRow.eachCell((cell) => {
						if (cell.value) {
							const columnName = cell.value.toString().trim();
							columns.push({
								name: columnName,
								value: columnName,
							});
						}
					});

					return columns.length > 0
						? columns
						: [{ name: 'No columns found', value: '' }];
				} catch (error) {
					const errorMessage = error instanceof Error ? error.message : 'Unknown error';
					return [
						{
							name: `Error: ${errorMessage}`,
							value: '',
						},
					];
				}
			},
		},
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnData: INodeExecutionData[] = [];

		for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
			try {
				const resource = this.getNodeParameter('resource', itemIndex) as string;
				const inputMode = this.getNodeParameter('inputMode', itemIndex) as string;

				// Load workbook
				let workbook: ExcelJS.Workbook;
				let filePath = '';
				let result: any;

				if (inputMode === 'filePath') {
					filePath = this.getNodeParameter('filePath', itemIndex) as string;
					workbook = new ExcelJS.Workbook();
					await workbook.xlsx.readFile(filePath);
				} else {
					const binaryPropertyName = this.getNodeParameter('binaryPropertyName', itemIndex) as string;
					await this.helpers.assertBinaryData(itemIndex, binaryPropertyName);
					const buffer = await this.helpers.getBinaryDataBuffer(itemIndex, binaryPropertyName);
					workbook = new ExcelJS.Workbook();
					await workbook.xlsx.load(buffer as any);
				}

				if (resource === 'row') {
					const operation = this.getNodeParameter('operation', itemIndex) as string;
					let sheetName = inputMode === 'filePath'
						? this.getNodeParameter('sheetNameOptions', itemIndex, '') as string
						: this.getNodeParameter('sheetName', itemIndex, '') as string;

					// If no sheet name specified, use the first worksheet
					if (!sheetName) {
						const firstWorksheet = workbook.worksheets[0];
						if (!firstWorksheet) {
							throw new NodeOperationError(
								this.getNode(),
								'No worksheets found in the workbook'
							);
						}
						sheetName = firstWorksheet.name;
					}

					const worksheet = workbook.getWorksheet(sheetName);
					if (!worksheet) {
						throw new NodeOperationError(
							this.getNode(),
							`Worksheet "${sheetName}" not found`
						);
					}

					const columnMap = ExcelAI.getColumnMapping(worksheet);

					switch (operation) {
						case 'readRows':
						result = await ExcelAI.handleReadRows(this, worksheet, itemIndex);
						break;

					case 'appendRow':
						result = await ExcelAI.handleAppendRow(this, worksheet, itemIndex, columnMap);
						if (inputMode === 'filePath' && this.getNodeParameter('autoSave', itemIndex, true)) {
							await workbook.xlsx.writeFile(filePath);
						}
						break;

					case 'insertRow':
						result = await ExcelAI.handleInsertRow(this, worksheet, itemIndex, columnMap);
						if (inputMode === 'filePath' && this.getNodeParameter('autoSave', itemIndex, true)) {
							await workbook.xlsx.writeFile(filePath);
						}
						break;

					case 'updateRow':
						result = await ExcelAI.handleUpdateRow(this, worksheet, itemIndex, columnMap);
						if (inputMode === 'filePath' && this.getNodeParameter('autoSave', itemIndex, true)) {
							await workbook.xlsx.writeFile(filePath);
						}
						break;

					case 'deleteRow':
						result = await ExcelAI.handleDeleteRow(this, worksheet, itemIndex);
						if (inputMode === 'filePath' && this.getNodeParameter('autoSave', itemIndex, true)) {
							await workbook.xlsx.writeFile(filePath);
						}
						break;

					case 'findRows':
						result = await ExcelAI.handleFindRows(this, worksheet, itemIndex, columnMap);
						break;

					default:
							throw new NodeOperationError(
								this.getNode(),
								`Unknown operation: ${operation}`
							);
					}
				} else if (resource === 'worksheet') {
					const worksheetOperation = this.getNodeParameter('worksheetOperation', itemIndex) as string;

					switch (worksheetOperation) {
						case 'listWorksheets':
							result = await ExcelAI.handleListWorksheets(this, workbook, itemIndex);
							break;

						case 'createWorksheet':
							result = await ExcelAI.handleCreateWorksheet(this, workbook, itemIndex);
							if (inputMode === 'filePath' && this.getNodeParameter('autoSave', itemIndex, true)) {
								await workbook.xlsx.writeFile(filePath);
							}
							break;

						case 'deleteWorksheet':
							result = await ExcelAI.handleDeleteWorksheet(this, workbook, itemIndex);
							if (inputMode === 'filePath' && this.getNodeParameter('autoSave', itemIndex, true)) {
								await workbook.xlsx.writeFile(filePath);
							}
							break;


					case 'renameWorksheet':
						result = await ExcelAI.handleRenameWorksheet(this, workbook, itemIndex);
						if (inputMode === 'filePath' && this.getNodeParameter('autoSave', itemIndex, true)) {
							await workbook.xlsx.writeFile(filePath);
						}
						break;

					case 'copyWorksheet':
						result = await ExcelAI.handleCopyWorksheet(this, workbook, itemIndex);
						if (inputMode === 'filePath' && this.getNodeParameter('autoSave', itemIndex, true)) {
							await workbook.xlsx.writeFile(filePath);
						}
						break;

					case 'getWorksheetInfo':
						result = await ExcelAI.handleGetWorksheetInfo(this, workbook, itemIndex);
						break;
						default:
							throw new NodeOperationError(
								this.getNode(),
								`Unknown worksheet operation: ${worksheetOperation}`
							);
					}
				}

				// Handle result output
				if (Array.isArray(result)) {
					result.forEach((item: any) => returnData.push({ json: item }));
				} else {
					returnData.push({ json: result });
				}

				// Return binary data if in binary mode
				if (inputMode === 'binaryData') {
					const buffer = await workbook.xlsx.writeBuffer();
					const binaryData = await this.helpers.prepareBinaryData(
						Buffer.from(buffer),
						'modified.xlsx',
						'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
					);
					returnData[returnData.length - 1].binary = { data: binaryData };
				}

			} catch (error) {
				if (this.continueOnFail()) {
					const errorMessage = error instanceof Error ? error.message : 'Unknown error';
					returnData.push({
						json: {
							error: errorMessage,
							resource: this.getNodeParameter('resource', itemIndex, ''),
							operation: this.getNodeParameter('operation', itemIndex, ''),
						},
					});
					continue;
				}
				throw error;
			}
		}

		return [returnData];
	}

	// ========== Helper Methods ==========

	private static getColumnMapping(worksheet: ExcelJS.Worksheet): Map<string, number> {
		const columnMap = new Map<string, number>();
		const headerRow = worksheet.getRow(1);

		headerRow.eachCell((cell, colNumber) => {
			if (cell.value) {
				const columnName = cell.value.toString().trim();
				columnMap.set(columnName, colNumber);
			}
		});

		return columnMap;
	}

	private static mapRowData(rowData: any, columnMap: Map<string, number>): any[] {
		const rowArray: any[] = new Array(columnMap.size);

		columnMap.forEach((colNumber, columnName) => {
			const value = rowData[columnName];
			rowArray[colNumber - 1] = value !== undefined ? value : '';
		});

		return rowArray;
	}

	private static parseJsonInput(context: IExecuteFunctions, input: string | object): any {
		if (typeof input === 'string') {
			try {
				return JSON.parse(input);
			} catch (error) {
				const errorMessage = error instanceof Error ? error.message : 'Unknown error';
				throw new NodeOperationError(
					context.getNode(),
					`Invalid JSON format: ${errorMessage}`
				);
			}
		}
		return input;
	}

	// ========== Row Operation Handlers ==========

	private static async handleReadRows(
		context: IExecuteFunctions,
		worksheet: ExcelJS.Worksheet,
		itemIndex: number
	): Promise<any[]> {
		const startRow = context.getNodeParameter('startRow', itemIndex, 2) as number;
		const endRow = context.getNodeParameter('endRow', itemIndex, 0) as number;

		const rows: any[] = [];
		const headerRow = worksheet.getRow(1);
		const headers: string[] = [];

		headerRow.eachCell((cell, colNumber) => {
			headers[colNumber - 1] = cell.value?.toString() || '';
		});

		const actualEndRow = endRow === 0 ? worksheet.rowCount : Math.min(endRow, worksheet.rowCount);

		for (let rowNum = startRow; rowNum <= actualEndRow; rowNum++) {
			const row = worksheet.getRow(rowNum);
			const rowData: any = { _rowNumber: rowNum };

			row.eachCell((cell, colNumber) => {
				const header = headers[colNumber - 1];
				if (header) {
					rowData[header] = cell.value;
				}
			});

			// Skip empty rows
			const hasData = Object.keys(rowData).some(key => key !== '_rowNumber' && rowData[key]);
			if (hasData) {
				rows.push(rowData);
			}
		}

		return rows;
	}

	private static async handleAppendRow(
		context: IExecuteFunctions,
		worksheet: ExcelJS.Worksheet,
		itemIndex: number,
		columnMap: Map<string, number>
	): Promise<any> {
		const rowDataInput = context.getNodeParameter('rowData', itemIndex) as string | object;
		const data = ExcelAI.parseJsonInput(context, rowDataInput);

		const rowArray = ExcelAI.mapRowData(data, columnMap);
		worksheet.addRow(rowArray);
		const rowNumber = worksheet.rowCount;

		return {
			success: true,
			operation: 'appendRow',
			rowNumber,
			data,
			message: `Row added successfully at row ${rowNumber}`,
		};
	}

	private static async handleInsertRow(
		context: IExecuteFunctions,
		worksheet: ExcelJS.Worksheet,
		itemIndex: number,
		columnMap: Map<string, number>
	): Promise<any> {
		const rowNumber = context.getNodeParameter('rowNumber', itemIndex) as number;
		const rowDataInput = context.getNodeParameter('rowData', itemIndex) as string | object;
		const data = ExcelAI.parseJsonInput(context, rowDataInput);

		if (rowNumber < 2) {
			throw new NodeOperationError(
				context.getNode(),
				'Cannot insert before header row (row 1)'
			);
		}

		const rowArray = ExcelAI.mapRowData(data, columnMap);
		worksheet.insertRow(rowNumber, rowArray);

		return {
			success: true,
			operation: 'insertRow',
			rowNumber,
			data,
			message: `Row inserted successfully at row ${rowNumber}`,
		};
	}

	private static async handleUpdateRow(
		context: IExecuteFunctions,
		worksheet: ExcelJS.Worksheet,
		itemIndex: number,
		columnMap: Map<string, number>
	): Promise<any> {
		const rowNumber = context.getNodeParameter('rowNumber', itemIndex) as number;
		const updatedDataInput = context.getNodeParameter('updatedData', itemIndex) as string | object;
		const data = ExcelAI.parseJsonInput(context, updatedDataInput);

		if (rowNumber < 2) {
			throw new NodeOperationError(
				context.getNode(),
				'Cannot update header row (row 1)'
			);
		}

		const row = worksheet.getRow(rowNumber);

		Object.entries(data).forEach(([columnName, value]) => {
			const colNumber = columnMap.get(columnName);
			if (colNumber) {
				row.getCell(colNumber).value = value as ExcelJS.CellValue;
			}
		});

		row.commit();

		return {
			success: true,
			operation: 'updateRow',
			rowNumber,
			updatedFields: Object.keys(data),
			message: `Row ${rowNumber} updated successfully`,
		};
	}

	private static async handleDeleteRow(
		context: IExecuteFunctions,
		worksheet: ExcelJS.Worksheet,
		itemIndex: number
	): Promise<any> {
		const rowNumber = context.getNodeParameter('rowNumber', itemIndex) as number;

		if (rowNumber < 2) {
			throw new NodeOperationError(
				context.getNode(),
				'Cannot delete header row (row 1)'
			);
		}

		worksheet.spliceRows(rowNumber, 1);

		return {
			success: true,
			operation: 'deleteRow',
			rowNumber,
			message: `Row ${rowNumber} deleted successfully`,
		};
	}

	private static async handleFindRows(
		context: IExecuteFunctions,
		worksheet: ExcelJS.Worksheet,
		itemIndex: number,
		columnMap: Map<string, number>
	): Promise<any[]> {
		const searchColumn = context.getNodeParameter('searchColumn', itemIndex) as string;
		const searchValue = context.getNodeParameter('searchValue', itemIndex) as string;
		const matchType = context.getNodeParameter('matchType', itemIndex, 'exact') as string;
		const returnRowNumbers = context.getNodeParameter('returnRowNumbers', itemIndex, false) as boolean;

		const colNumber = columnMap.get(searchColumn);
		if (!colNumber) {
			throw new NodeOperationError(
				context.getNode(),
				`Column "${searchColumn}" not found`
			);
		}

		const matchingRows: any[] = [];
		const headerRow = worksheet.getRow(1);
		const headers: string[] = [];

		headerRow.eachCell((cell, colNum) => {
			headers[colNum - 1] = cell.value?.toString() || '';
		});

		worksheet.eachRow((row, rowNumber) => {
			if (rowNumber === 1) return;

			const cellValue = row.getCell(colNumber).value?.toString() || '';
			let isMatch = false;

			const searchLower = searchValue.toLowerCase();
			const cellLower = cellValue.toLowerCase();

			switch (matchType) {
				case 'exact':
					isMatch = cellLower === searchLower;
					break;
				case 'contains':
					isMatch = cellLower.includes(searchLower);
					break;
				case 'startsWith':
					isMatch = cellLower.startsWith(searchLower);
					break;
				case 'endsWith':
					isMatch = cellLower.endsWith(searchLower);
					break;
			}

			if (isMatch) {
				if (returnRowNumbers) {
					matchingRows.push(rowNumber);
				} else {
					const rowData: any = { _rowNumber: rowNumber };
					row.eachCell((cell, colNum) => {
						const header = headers[colNum - 1];
						if (header) {
							rowData[header] = cell.value;
						}
					});
					matchingRows.push(rowData);
				}
			}
		});

		if (returnRowNumbers) {
			return [
				{
					operation: 'findRows',
					searchColumn,
					searchValue,
					matchType,
					rowNumbers: matchingRows,
					count: matchingRows.length,
				},
			];
		}

		// Return message if no rows found
		if (matchingRows.length === 0) {
			return [
				{
					operation: 'findRows',
					searchColumn,
					searchValue,
					matchType,
					count: 0,
					message: 'No matching rows found',
				},
			];
		}

		return matchingRows;
	}

	// ========== Worksheet Operation Handlers ==========

	private static async handleListWorksheets(
		context: IExecuteFunctions,
		workbook: ExcelJS.Workbook,
		itemIndex: number
	): Promise<any> {
		const includeHidden = context.getNodeParameter('includeHidden', itemIndex, false) as boolean;
		const worksheets: any[] = [];

		workbook.eachSheet((worksheet, sheetId) => {
			if (!includeHidden && worksheet.state !== 'visible') {
				return;
			}

			worksheets.push({
				id: sheetId,
				name: worksheet.name,
				rowCount: worksheet.rowCount,
				columnCount: worksheet.columnCount,
				state: worksheet.state,
			});
		});

		return {
			operation: 'listWorksheets',
			count: worksheets.length,
			worksheets,
		};
	}

	private static async handleCreateWorksheet(
		context: IExecuteFunctions,
		workbook: ExcelJS.Workbook,
		itemIndex: number
	): Promise<any> {
		const newSheetName = context.getNodeParameter('newSheetName', itemIndex) as string;
		const initialDataInput = context.getNodeParameter('initialData', itemIndex, '[]') as string | any[];
		const initialData = ExcelAI.parseJsonInput(context, initialDataInput);

		// Check if worksheet already exists
		if (workbook.getWorksheet(newSheetName)) {
			throw new NodeOperationError(
				context.getNode(),
				`Worksheet "${newSheetName}" already exists`
			);
		}

		const worksheet = workbook.addWorksheet(newSheetName);

		// Add initial data if provided
		if (Array.isArray(initialData) && initialData.length > 0) {
			const headers = Object.keys(initialData[0]);
			worksheet.addRow(headers);

			initialData.forEach((row) => {
				const rowArray = headers.map((header) => row[header] ?? '');
				worksheet.addRow(rowArray);
			});
		}

		return {
			success: true,
			operation: 'createWorksheet',
			sheetName: newSheetName,
			rowCount: worksheet.rowCount,
			message: `Worksheet "${newSheetName}" created successfully`,
		};
	}

	private static async handleDeleteWorksheet(
		context: IExecuteFunctions,
		workbook: ExcelJS.Workbook,
		itemIndex: number
	): Promise<any> {
		const worksheetName = context.getNodeParameter('worksheetNameOptions', itemIndex) as string;
		const worksheet = workbook.getWorksheet(worksheetName);

		if (!worksheet) {
			throw new NodeOperationError(
				context.getNode(),
				`Worksheet "${worksheetName}" not found`
			);
		}

		workbook.removeWorksheet(worksheet.id);

		return {
			success: true,
			operation: 'deleteWorksheet',
			sheetName: worksheetName,
			message: `Worksheet "${worksheetName}" deleted successfully`,
		};
	}

	private static async handleRenameWorksheet(
		context: IExecuteFunctions,
		workbook: ExcelJS.Workbook,
		itemIndex: number
	): Promise<any> {
		const worksheetName = context.getNodeParameter('worksheetNameOptions', itemIndex) as string;
		const newSheetName = context.getNodeParameter('newSheetName', itemIndex) as string;
		const worksheet = workbook.getWorksheet(worksheetName);

		if (!worksheet) {
			throw new NodeOperationError(
				context.getNode(),
				`Worksheet "${worksheetName}" not found`
			);
		}

		// Check if new name already exists
		if (workbook.getWorksheet(newSheetName)) {
			throw new NodeOperationError(
				context.getNode(),
				`Worksheet "${newSheetName}" already exists`
			);
		}

		worksheet.name = newSheetName;

		return {
			success: true,
			operation: 'renameWorksheet',
			oldName: worksheetName,
			newName: newSheetName,
			message: `Worksheet renamed from "${worksheetName}" to "${newSheetName}" successfully`,
		};
	}

	private static async handleCopyWorksheet(
		context: IExecuteFunctions,
		workbook: ExcelJS.Workbook,
		itemIndex: number
	): Promise<any> {
		const worksheetName = context.getNodeParameter('worksheetNameOptions', itemIndex) as string;
		const newSheetName = context.getNodeParameter('newSheetName', itemIndex) as string;
		const sourceWorksheet = workbook.getWorksheet(worksheetName);

		if (!sourceWorksheet) {
			throw new NodeOperationError(
				context.getNode(),
				`Worksheet "${worksheetName}" not found`
			);
		}

		// Check if new name already exists
		if (workbook.getWorksheet(newSheetName)) {
			throw new NodeOperationError(
				context.getNode(),
				`Worksheet "${newSheetName}" already exists`
			);
		}

		// Create new worksheet
		const newWorksheet = workbook.addWorksheet(newSheetName);

		// Copy all rows from source to new worksheet
		sourceWorksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
			const newRow = newWorksheet.getRow(rowNumber);
			row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
				const newCell = newRow.getCell(colNumber);
				newCell.value = cell.value;
				newCell.style = cell.style;
			});
			newRow.height = row.height;
			newRow.commit();
		});

		// Copy column widths
		sourceWorksheet.columns.forEach((column, index) => {
			if (column && newWorksheet.getColumn(index + 1)) {
				newWorksheet.getColumn(index + 1).width = column.width;
			}
		});

		return {
			success: true,
			operation: 'copyWorksheet',
			sourceName: worksheetName,
			newName: newSheetName,
			rowCount: newWorksheet.rowCount,
			message: `Worksheet "${worksheetName}" copied to "${newSheetName}" successfully`,
		};
	}

	private static async handleGetWorksheetInfo(
		context: IExecuteFunctions,
		workbook: ExcelJS.Workbook,
		itemIndex: number
	): Promise<any> {
		const worksheetName = context.getNodeParameter('worksheetNameOptions', itemIndex) as string;
		const worksheet = workbook.getWorksheet(worksheetName);

		if (!worksheet) {
			throw new NodeOperationError(
				context.getNode(),
				`Worksheet "${worksheetName}" not found`
			);
		}

		// Get column information
		const columns: any[] = [];
		const headerRow = worksheet.getRow(1);
		
		headerRow.eachCell((cell, colNumber) => {
			const column = worksheet.getColumn(colNumber);
			columns.push({
				index: colNumber,
				letter: ExcelAI.columnNumberToLetter(colNumber),
				header: cell.value?.toString() || '',
				width: column.width || 10,
			});
		});

		return {
			operation: 'getWorksheetInfo',
			sheetName: worksheetName,
			rowCount: worksheet.rowCount,
			columnCount: worksheet.columnCount,
			actualRowCount: worksheet.actualRowCount,
			actualColumnCount: worksheet.actualColumnCount,
			state: worksheet.state,
			columns,
		};
	}

	private static columnNumberToLetter(column: number): string {
		let letter = '';
		while (column > 0) {
			const remainder = (column - 1) % 26;
			letter = String.fromCharCode(65 + remainder) + letter;
			column = Math.floor((column - 1) / 26);
		}
		return letter;
	}
}
