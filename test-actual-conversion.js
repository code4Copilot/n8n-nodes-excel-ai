/**
 * 實際測試型態轉換功能
 * 這個測試會實際創建 Excel 檔案並驗證型態轉換
 */

const ExcelJS = require('exceljs');
const fs = require('fs/promises');
const path = require('path');

// 模擬 convertValue 函數（與節點中的實現相同）
function convertValue(value) {
    // Handle null, undefined, or already non-string values
    if (value === null || value === undefined) {
        return null;
    }

    // If not a string, return as-is
    if (typeof value !== 'string') {
        return value;
    }

    // Trim whitespace
    const trimmed = value.trim();

    // Empty string or "null" -> null
    if (trimmed === '' || trimmed.toLowerCase() === 'null') {
        return null;
    }

    // Boolean values (case-insensitive)
    if (trimmed.toLowerCase() === 'true') {
        return true;
    }
    if (trimmed.toLowerCase() === 'false') {
        return false;
    }

    // Try to parse as number
    if (/^-?\d+$/.test(trimmed)) {
        // Integer
        const num = parseInt(trimmed, 10);
        if (!isNaN(num)) {
            return num;
        }
    } else if (/^-?\d*\.\d+$/.test(trimmed)) {
        // Float
        const num = parseFloat(trimmed);
        if (!isNaN(num)) {
            return num;
        }
    }

    // Check for ISO 8601 date format
    if (/^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2}(\.\d{1,3})?)?(Z|[+-]\d{2}:?\d{2})?)?$/.test(trimmed)) {
        const date = new Date(trimmed);
        if (!isNaN(date.getTime())) {
            return date;
        }
    }

    // Return original string if no conversion applies
    return value;
}

async function testTypeConversion() {
    console.log('=== 開始測試型態轉換功能 ===\n');

    // 創建工作簿
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('TypeConversionTest');

    // 添加標題行
    worksheet.addRow(['Name', 'Age', 'Active', 'Balance', 'JoinDate', 'Notes']);

    // 測試資料 - 所有值都是字串
    const testData = {
        Name: 'John Doe',
        Age: '30',              // 字串 → 數字
        Active: 'true',         // 字串 → 布林
        Balance: '1234.56',     // 字串 → 浮點數
        JoinDate: '2024-01-15', // 字串 → 日期
        Notes: 'null'           // 字串 → null
    };

    console.log('原始輸入資料（全部為字串）：');
    console.log(JSON.stringify(testData, null, 2));
    console.log('');

    // 模擬節點的 mapRowData 行為
    const columnMap = new Map([
        ['Name', 1],
        ['Age', 2],
        ['Active', 3],
        ['Balance', 4],
        ['JoinDate', 5],
        ['Notes', 6]
    ]);

    const rowArray = [];
    columnMap.forEach((colNumber, columnName) => {
        const value = testData[columnName];
        const convertedValue = value !== undefined ? convertValue(value) : '';
        rowArray[colNumber - 1] = convertedValue;
        
        console.log(`${columnName}:`);
        console.log(`  輸入: ${JSON.stringify(value)} (type: ${typeof value})`);
        console.log(`  輸出: ${JSON.stringify(convertedValue)} (type: ${typeof convertedValue})`);
        console.log('');
    });

    // 添加轉換後的資料
    worksheet.addRow(rowArray);

    // 儲存檔案
    const outputPath = path.join(__dirname, 'test-type-conversion.xlsx');
    await workbook.xlsx.writeFile(outputPath);
    console.log(`✓ Excel 檔案已創建: ${outputPath}`);

    // 讀取檔案並驗證
    console.log('\n=== 驗證儲存的資料 ===\n');
    const readWorkbook = new ExcelJS.Workbook();
    await readWorkbook.xlsx.readFile(outputPath);
    const readWorksheet = readWorkbook.getWorksheet('TypeConversionTest');
    const dataRow = readWorksheet.getRow(2);

    console.log('從 Excel 讀取的資料：');
    dataRow.eachCell((cell, colNumber) => {
        const header = readWorksheet.getRow(1).getCell(colNumber).value;
        console.log(`${header}: ${JSON.stringify(cell.value)} (type: ${typeof cell.value})`);
    });

    console.log('\n=== 測試完成 ===');
    console.log('檔案位置:', outputPath);
}

// 運行測試
testTypeConversion().catch(console.error);
