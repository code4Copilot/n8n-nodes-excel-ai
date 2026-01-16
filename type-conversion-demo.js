/**
 * 自動型態轉換示範
 * 
 * 這個檔案展示了 ExcelAI 節點的自動型態轉換功能
 */

// 示範資料 - 所有值都是字串格式
const exampleRowData = {
  // 字串會保持字串
  "Name": "John Doe",
  "Department": "Engineering",
  
  // 字串數字會轉換為數值
  "Age": "30",                    // → 30 (數字)
  "Salary": "75000.50",           // → 75000.50 (浮點數)
  "YearsOfService": "5",          // → 5 (數字)
  
  // 字串布林值會轉換為布林
  "IsActive": "true",             // → true
  "HasBenefits": "TRUE",          // → true (不區分大小寫)
  "IsRemote": "false",            // → false
  
  // 日期字串會轉換為 Date 物件
  "HireDate": "2019-01-15",       // → Date 物件
  "LastReview": "2024-01-15T10:30:00Z",  // → Date 物件
  
  // 空值相關
  "MiddleName": "null",           // → null
  "Nickname": "",                 // → null
  "Suffix": "   ",                // → null (僅空白)
  
  // 負數和小數
  "VacationDays": "-2",           // → -2 (負數)
  "OvertimeHours": "12.5",        // → 12.5 (浮點數)
  
  // 不符合格式的保持字串
  "PhoneNumber": "N/A",           // → "N/A" (保持字串)
  "Bonus": "$1000",               // → "$1000" (保持字串)
  "Notes": "Excellent performer"  // → "Excellent performer" (保持字串)
};

console.log("=== 自動型態轉換示範 ===\n");
console.log("輸入資料（所有值都是字串）：");
console.log(JSON.stringify(exampleRowData, null, 2));

console.log("\n轉換後的結果：");
console.log("- Name: 字串 'John Doe'");
console.log("- Age: 數字 30");
console.log("- Salary: 數字 75000.50");
console.log("- IsActive: 布林值 true");
console.log("- HasBenefits: 布林值 true");
console.log("- IsRemote: 布林值 false");
console.log("- HireDate: Date 物件 (2019-01-15)");
console.log("- LastReview: Date 物件 (2024-01-15T10:30:00Z)");
console.log("- MiddleName: null");
console.log("- Nickname: null");
console.log("- Suffix: null");
console.log("- VacationDays: 數字 -2");
console.log("- OvertimeHours: 數字 12.5");
console.log("- PhoneNumber: 字串 'N/A'");
console.log("- Bonus: 字串 '$1000'");
console.log("- Notes: 字串 'Excellent performer'");

console.log("\n=== 在 n8n 中使用 ===\n");
console.log("1. 新增列 (Append Row):");
console.log("   在 'Row Data' 欄位輸入 JSON：");
console.log("   " + JSON.stringify({
  Name: "Alice",
  Age: "25",        // 會轉換為數字 25
  Active: "true"    // 會轉換為布林值 true
}, null, 2).replace(/\n/g, '\n   '));

console.log("\n2. 更新列 (Update Row):");
console.log("   在 'Updated Data' 欄位輸入 JSON：");
console.log("   " + JSON.stringify({
  Age: "26",           // 會轉換為數字 26
  Active: "FALSE",     // 會轉換為布林值 false (不區分大小寫)
  Balance: "1000.50"   // 會轉換為數字 1000.50
}, null, 2).replace(/\n/g, '\n   '));

console.log("\n3. 插入列 (Insert Row):");
console.log("   在 'Row Data' 欄位輸入 JSON：");
console.log("   " + JSON.stringify({
  Name: "Bob",
  Age: "",             // 空字串會轉換為 null
  Active: "true",      // 會轉換為布林值 true
  Notes: "null"        // 字串 "null" 會轉換為 null
}, null, 2).replace(/\n/g, '\n   '));

console.log("\n=== 與 AI Agent 整合 ===\n");
console.log("AI Agent 可以直接傳遞字串值，系統會自動轉換：");
console.log("- AI: '將員工年齡設為 30' → 系統自動轉換 '30' 為數字 30");
console.log("- AI: '設定狀態為 true' → 系統自動轉換 'true' 為布林值 true");
console.log("- AI: '清空備註欄位' → 系統將空字串轉換為 null");
console.log("- AI: '設定入職日期為 2024-01-15' → 系統轉換為 Date 物件");

console.log("\n完成！");
