# n8n-nodes-excel-ai

![n8n](https://img.shields.io/badge/n8n-Community%20Node-FF6D5A)
![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Version](https://img.shields.io/npm/v/n8n-nodes-excel-ai)

一個強大的 n8n 社群節點，用於對 Excel 檔案執行 CRUD（新增、讀取、更新、刪除）操作，並支援 **AI Agent 整合**，可與 n8n AI Agents 無縫整合，實現自然語言的 Excel 操作。

> **v1.0.12 新增：`includeRowNumber` 開關，讀取（readRows / filterRows）時可選擇是否在輸出中包含 `_rowNumber` 欄位。預設為 `true` 以維持向後相容；純讀取情境設為 `false` 可保持輸出乾淨。**

> **v1.0.11 新增：Clear Rows 操作，可一鍵清除所有資料列並保留標題列（可選），支援 AI Agent 呼叫。**

> **v1.0.10 改進：工作表下拉選單錯誤狀態改用 `__error__` 哨兵值，提供更清楚的 ⚠ 警告提示，避免以無效工作表名稱提交，提升操作體驗。**
## ✨ 功能特色

### 🤖 AI Agent 整合
**純值回傳邏輯**：所有 cell 讀取皆回傳可直接使用的純值（數字、字串、日期、布林、公式結果、超連結文字、RichText 純文字、錯誤字串等），不再回傳物件。

### 📊 完整的 CRUD 操作
- **讀取**：使用篩選器和分頁查詢資料
- **新增**：使用自動欄位映射新增列
- **更新**：部分更新現有列
- **刪除**：依行號刪除列
- **過濾**：使用進階條件和多種運算符過濾列
- **清除列**：清除所有資料列並保留標題列

### 🗂️ 工作表管理
- **列出工作表**：取得活頁簿中的所有工作表
- **建立工作表**：新增工作表並可選擇初始資料
- **刪除工作表**：從活頁簿中移除工作表
- **重新命名工作表**：重新命名現有工作表
- **複製工作表**：複製工作表，包含所有資料和格式
- **取得工作表資訊**：取得工作表的詳細資訊，包括欄位設定

### 🔄 彈性的輸入模式
- **檔案路徑**：使用檔案系統中的 Excel 檔案
- **二進位資料**：處理來自上一個工作流程步驟的檔案

## 📦 安裝

### 選項 1：npm（推薦）

```bash
# 導航到 n8n 自訂節點目錄
cd ~/.n8n/nodes

# 安裝套件
npm install n8n-nodes-excel-ai
```

#### 🔒 安全性：修復 form-data 漏洞

為了解決來自 `n8n-workflow` 的 `form-data` 安全漏洞，請在安裝目錄的 `package.json` 中加入：

```json
{
  "overrides": {
    "form-data": "^4.0.4"
  }
}
```

然後重新安裝：

```bash
npm install
npm audit
```

### 選項 2：Docker

在您的 `docker-compose.yml` 中新增：

```yaml
version: '3.7'

services:
  n8n:
    image: n8nio/n8n
    environment:
      - NODE_FUNCTION_ALLOW_EXTERNAL=n8n-nodes-excel-ai
      - N8N_CUSTOM_EXTENSIONS=/home/node/.n8n/custom
    volumes:
      - ~/.n8n/custom:/home/node/.n8n/custom
```

然後在容器內安裝：
```bash
docker exec -it <n8n-container> npm install n8n-nodes-excel-ai
```

### 選項 3：手動安裝
```bash
# 複製儲存庫
git clone https://github.com/code4Copilot/n8n-nodes-excel-ai.git
cd n8n-nodes-excel-ai

# 安裝相依套件
npm install

# 建置節點
npm run build

# 連結到 n8n
npm link
cd ~/.n8n/nodes
npm link n8n-nodes-excel-ai
```

## 🚀 快速開始

### 基本使用

#### 1. 從 Excel 讀取資料
```javascript
// 節點設定
資源：列
操作：讀取列
檔案路徑：/data/customers.xlsx
工作表名稱：Customers
起始行：2
結束行：0（讀取全部）

// 輸出
[
  { "_rowNumber": 2, "Name": "張三", "Email": "zhang@example.com" },
  { "_rowNumber": 3, "Name": "李四", "Email": "li@example.com" }
]
```

#### 2. 新增列
```javascript
// 節點設定
資源：列
操作：附加列
檔案路徑：/data/customers.xlsx
工作表名稱：Customers
行資料：{"Name": "王五", "Email": "wang@example.com", "Status": "Active"}

// 輸出
{
  "success": true,
  "operation": "appendRow",
  "rowNumber": 15,
  "message": "Row added successfully at row 15"
}
```

#### 3. 過濾列
```javascript
// 節點設定
資源：列
操作：過濾列
檔案路徑：/data/customers.xlsx
工作表名稱：Customers
過濾條件：
  - 欄位：Status
  - 運算符：equals
  - 值：Active
條件邏輯：and

// 輸出
[
  { "_rowNumber": 2, "Name": "張三", "Status": "Active" },
  { "_rowNumber": 5, "Name": "李四", "Status": "Active" }
]
```

### 🤖 AI Agent 使用範例

#### 範例 1：自然語言查詢

**使用者：** "顯示 customers.xlsx 檔案中的所有客戶"

**AI Agent 執行：**
```javascript
{
  "operation": "readRows",
  "filePath": "/data/customers.xlsx",
  "sheetName": "Customers",
  "startRow": 2,
  "endRow": 0
}
```

#### 範例 2：透過 AI 新增資料

**使用者：** "新增一個名為 Sarah Johnson 的客戶，電子郵件是 sarah@example.com"

**AI Agent 執行：**
```javascript
{
  "operation": "appendRow",
  "filePath": "/data/customers.xlsx",
  "sheetName": "Customers",
  "rowData": {
    "Name": "Sarah Johnson",
    "Email": "sarah@example.com"
  }
}
```

#### 範例 3：使用 AI 過濾

**使用者：** "找出波士頓的所有活躍客戶"

**AI Agent 執行：**
```javascript
{
  "operation": "filterRows",
  "filePath": "/data/customers.xlsx",
  "sheetName": "Customers",
  "filterConditions": {
    "conditions": [
      { "field": "Status", "operator": "equals", "value": "Active" },
      { "field": "City", "operator": "equals", "value": "Boston" }
    ]
  },
  "conditionLogic": "and"
}
```

## � 自動型態轉換

節點會在新增或更新列時自動將字串值轉換為適當的型態。這使得與 AI Agents 和手動輸入的配合更加容易。

### 支援的轉換

#### 數字
字串形式的數字會自動轉換為數值型態：
- `"123"` → `123`（整數）
- `"45.67"` → `45.67`（浮點數）
- `"-99"` → `-99`（負整數）
- `"-123.45"` → `-123.45`（負浮點數）

#### 布林值
字串形式的布林值會自動轉換（不區分大小寫）：
- `"true"` → `true`
- `"false"` → `false`
- `"TRUE"` → `true`
- `"False"` → `false`

#### 日期
ISO 8601 格式的日期字串會轉換為 Date 物件：
- `"2024-01-15"` → `Date 物件`
- `"2024-01-15T10:30:00Z"` → `Date 物件`
- `"2024-01-15T10:30:00.123Z"` → `Date 物件`

#### 空值
以下值會被轉換為 `null`：
- `"null"`（字串）→ `null`
- `""`（空字串）→ `null`
- `"   "`（僅有空白）→ `null`

#### 保持不變的值
- 普通字串保持為字串：`"Hello"` → `"Hello"`
- 已轉換的值會被保留：`123` → `123`、`true` → `true`
- 非標準格式會被保留：`"$100"` → `"$100"`、`"N/A"` → `"N/A"`

### 使用範例

**範例 1：使用型態轉換新增列**
```javascript
{
  "operation": "appendRow",
  "rowData": {
    "Name": "John Doe",      // 字串 → "John Doe"
    "Age": "30",             // 字串 → 30（數字）
    "Active": "true",        // 字串 → true（布林值）
    "JoinDate": "2024-01-15", // 字串 → Date 物件
    "Salary": "75000.50",    // 字串 → 75000.50（數字）
    "Notes": "null"          // 字串 → null
  }
}
```

**範例 2：使用型態轉換更新列**
```javascript
{
  "operation": "updateRow",
  "rowNumber": 5,
  "updatedData": {
    "Age": "35",             // 字串 → 35（數字）
    "Active": "FALSE",       // 字串 → false（不區分大小寫）
    "Balance": "-100.50"     // 字串 → -100.50（負數）
  }
}
```

**範例 3：AI Agent 整合**
AI 現在可以直接傳遞字串值，無需擔心型態問題：

**使用者**：「新增一位員工：Alice，年齡 25，狀態為 active」

**AI Agent**：
```javascript
{
  "operation": "appendRow",
  "rowData": {
    "Name": "Alice",
    "Age": "25",        // AI 傳遞字串，自動轉換為 25
    "Active": "true"    // AI 傳遞字串，自動轉換為 true
  }
}
```

## �📚 操作參考

### 列操作

#### 讀取列
- **用途**：從 Excel 檔案讀取資料
- **參數**：
  - `startRow`：起始行號（預設：2）
  - `endRow`：結束行號（0 = 所有列）
- **返回**：帶有 `_rowNumber` 欄位的列物件陣列

#### 附加列
- **用途**：在工作表末端新增列
- **智能空白列處理**：自動重用最後一列（如果為空白）
- **參數**：
  - `rowData`：包含欄位名稱和值的 JSON 物件
- **返回**：成功狀態、列號和 `wasEmptyRowReused` 標記

**智能行為：**
- ✅ 偵測最後一列是否為空白（所有儲存格為 null 或空字串）
- ✅ 重用空白列以保持 Excel 檔案整潔
- ✅ 僅在最後一列包含資料時才新增新列
- ✅ 重用空白列時返回 `wasEmptyRowReused: true`
- ✅ 訊息中標註 "(reused empty row)" 以便識別

**範例：**
```javascript
// 如果 Excel 的最後一列為空白，將被重用
{
  "operation": "appendRow",
  "rowData": {
    "Name": "Jane",
    "Age": 25,
    "Department": "Sales"
  }
}
```

**輸出（重用空白列時）：**
```javascript
{
  "success": true,
  "operation": "appendRow",
  "rowNumber": 3,
  "wasEmptyRowReused": true,
  "message": "Row added successfully at row 3 (reused empty row)"
}
```

#### 插入列
- **用途**：在特定位置插入列
- **參數**：
  - `rowNumber`：插入位置
  - `rowData`：包含欄位名稱和值的 JSON 物件
- **返回**：成功狀態和列號

#### 更新列
- **用途**：更新現有列
- **參數**：
  - `rowNumber`：要更新的列
  - `updatedData`：包含要更新欄位的 JSON 物件
- **返回**：成功狀態和已更新的欄位
- **欄位驗證**：
  - ✅ 自動跳過不存在的欄位，不會中斷執行
  - ✅ 返回 `updatedFields` 陣列列出成功更新的欄位
  - ✅ 返回 `skippedFields` 陣列列出被跳過的欄位（如果有）
  - ✅ 包含 `warning` 訊息說明哪些欄位被跳過

**更新列範例（含欄位驗證）：**
```javascript
{
  "operation": "updateRow",
  "rowNumber": 5,
  "updatedData": {
    "Status": "Completed",
    "InvalidField": "test",  // 此欄位不存在
    "Notes": "Updated"
  }
}
```

**輸出：**
```javascript
{
  "success": true,
  "operation": "updateRow",
  "rowNumber": 5,
  "updatedFields": ["Status", "Notes"],
  "skippedFields": ["InvalidField"],
  "warning": "The following fields were not found in the worksheet and were skipped: InvalidField",
  "message": "Row 5 updated successfully"
}
```

#### 刪除列
- **用途**：移除特定列
- **參數**：
  - `rowNumber`：要刪除的列（不能是 1 - 標題列）
- **返回**：成功狀態

#### 過濾列
- **用途**：使用多個條件和邏輯運算符過濾列
- **參數**：
  - `filterConditions`：過濾條件陣列，每個條件包含：
    - `field`：要過濾的欄位名稱
    - `operator`：equals | notEquals | contains | notContains | greaterThan | greaterOrEqual | lessThan | lessOrEqual | startsWith | endsWith | isEmpty | isNotEmpty
    - `value`：要比較的值（isEmpty/isNotEmpty 不需要）
  - `conditionLogic`：and | or - 如何組合多個條件
- **返回**：符合的列陣列，包含 _rowNumber 欄位
- **欄位驗證**：
  - ❌ 過濾條件使用不存在的欄位會立即拋出錯誤
  - ✅ 錯誤訊息會列出無效的欄位和所有可用欄位
  - ✅ 防止產生不正確的過濾結果
  - ✅ 同時支援 File Path 和 Binary Data 模式

**錯誤範例：**
```javascript
// 如果 "Category" 欄位不存在於工作表中
{
  "operation": "filterRows",
  "filterConditions": {
    "conditions": [
      { "field": "Category", "operator": "equals", "value": "Electronics" }
    ]
  }
}
```

**錯誤訊息：**
```
Filter condition error: The following field(s) do not exist in the worksheet: Category. 
Available fields are: Product, Price, Stock, Status
```

**過濾列範例：**

1. **單一條件 - 精確匹配：**
```javascript
{
  "operation": "filterRows",
  "filePath": "/data/employees.xlsx",
  "sheetName": "Staff",
  "filterConditions": {
    "conditions": [
      { "field": "Department", "operator": "equals", "value": "Engineering" }
    ]
  },
  "conditionLogic": "and"
}
```

2. **使用 AND 的多條件：**
```javascript
{
  "operation": "filterRows",
  "filePath": "/data/products.xlsx",
  "sheetName": "Inventory",
  "filterConditions": {
    "conditions": [
      { "field": "Category", "operator": "equals", "value": "Electronics" },
      { "field": "Price", "operator": "greaterThan", "value": "100" },
      { "field": "Stock", "operator": "greaterThan", "value": "0" }
    ]
  },
  "conditionLogic": "and"
}
```

3. **使用 OR 的多條件：**
```javascript
{
  "operation": "filterRows",
  "filterConditions": {
    "conditions": [
      { "field": "Priority", "operator": "equals", "value": "High" },
      { "field": "Priority", "operator": "equals", "value": "Urgent" }
    ]
  },
  "conditionLogic": "or"
}
```

4. **使用 Contains 的文字搜尋：**
```javascript
{
  "operation": "filterRows",
  "filterConditions": {
    "conditions": [
      { "field": "Email", "operator": "contains", "value": "@company.com" }
    ]
  },
  "conditionLogic": "and"
}
```

5. **檢查空白欄位：**
```javascript
{
  "operation": "filterRows",
  "filterConditions": {
    "conditions": [
      { "field": "Phone", "operator": "isEmpty" }
    ]
  },
  "conditionLogic": "and"
}
```

6. **範圍過濾：**
```javascript
{
  "operation": "filterRows",
  "filterConditions": {
    "conditions": [
      { "field": "Age", "operator": "greaterOrEqual", "value": "18" },
      { "field": "Age", "operator": "lessOrEqual", "value": "65" }
    ]
  },
  "conditionLogic": "and"
}
```

**可用運算符：**
- `equals` - 精確匹配
- `notEquals` - 不等於
- `contains` - 文字包含子字串
- `notContains` - 文字不包含子字串
- `greaterThan` - 數值大於
- `greaterOrEqual` - 數值大於或等於
- `lessThan` - 數值小於
- `lessOrEqual` - 數值小於或等於
- `startsWith` - 文字開始於
- `endsWith` - 文字結束於
- `isEmpty` - 欄位為空或 null
- `isNotEmpty` - 欄位有值

### 工作表操作

#### 列出工作表
- **用途**：取得活頁簿中的所有工作表
- **參數**：
  - `includeHidden`：包含隱藏工作表（預設：false）
- **返回**：工作表資訊陣列

#### 建立工作表
- **用途**：建立新工作表
- **參數**：
  - `newSheetName`：新工作表的名稱
  - `initialData`：可選的初始資料物件陣列
- **返回**：成功狀態和工作表資訊

#### 刪除工作表
- **用途**：從活頁簿移除工作表
- **參數**：
  - `worksheetName`：要刪除的工作表名稱
- **返回**：成功狀態

#### 重新命名工作表
- **用途**：重新命名現有工作表
- **參數**：
  - `worksheetName`：工作表的目前名稱
  - `newSheetName`：工作表的新名稱
- **返回**：包含舊名稱和新名稱的成功狀態

**範例：**
```javascript
{
  "worksheetOperation": "renameWorksheet",
  "filePath": "/data/reports.xlsx",
  "worksheetName": "Sheet1",
  "newSheetName": "Sales_2024",
  "autoSave": true
}
```

#### 複製工作表
- **用途**：複製工作表及其所有資料和格式
- **參數**：
  - `worksheetName`：要複製的工作表名稱
  - `newSheetName`：複製工作表的名稱
- **返回**：包含來源和新工作表名稱、列數的成功狀態

**範例：**
```javascript
{
  "worksheetOperation": "copyWorksheet",
  "filePath": "/data/templates.xlsx",
  "worksheetName": "Template_2024",
  "newSheetName": "Template_2025",
  "autoSave": true
}
```

**輸出：**
```javascript
{
  "success": true,
  "operation": "copyWorksheet",
  "sourceName": "Template_2024",
  "newName": "Template_2025",
  "rowCount": 50
}
```

#### 取得工作表資訊
- **用途**：取得工作表的詳細資訊
- **參數**：
  - `worksheetName`：工作表名稱
- **返回**：包含欄位在內的詳細工作表資訊

**範例：**
```javascript
{
  "worksheetOperation": "getWorksheetInfo",
  "filePath": "/data/database.xlsx",
  "worksheetName": "Users"
}
```

**輸出：**
```javascript
{
  "operation": "getWorksheetInfo",
  "sheetName": "Users",
  "rowCount": 150,
  "columnCount": 6,
  "actualRowCount": 151,
  "actualColumnCount": 6,
  "state": "visible",
  "columns": [
    {
      "index": 1,
      "letter": "A",
      "header": "UserID",
      "width": 15
    },
    {
      "index": 2,
      "letter": "B",
      "header": "Name",
      "width": 25
    },
    {
      "index": 3,
      "letter": "C",
      "header": "Email",
      "width": 30
    }
    // ... 更多欄位
  ]
}
```

## 🤖 與 AI Agents 一起使用

### 設定

此節點設計為與 n8n AI Agents 無縫協作。該節點配置了 `usableAsTool: true`，使其自動可供 AI Agents 使用。

### 啟用 AI 參數

1. 在節點設定中，尋找帶有 ✨ 星號圖示的參數
2. 點擊任何參數旁邊的 ✨ 圖示以啟用 AI 自動填寫
3. AI Agent 現在可以自動設定該參數的值

### AI Agent 範例

#### 範例 1：自然語言資料操作

**工作流程設定：**
```
AI Agent → Excel AI 節點
```

**使用者查詢：** "從 Excel 檔案中取得所有客戶並顯示紐約的客戶"

**AI Agent 動作：**
1. 使用 Excel AI 讀取所有列
2. 篩選紐約客戶的結果
3. 返回格式化的結果

#### 範例 2：多步驟操作

**使用者查詢：** "複製 2024 範本工作表以建立 2025 版本，然後新增一月資料"

**AI Agent 動作：**
1. 使用 `copyWorksheet` 操作複製工作表
2. 使用 `appendRow` 新增資料列
3. 確認完成

#### 範例 3：資料分析

**使用者查詢：** "顯示 Users 工作表的結構"

**AI Agent 動作：**
1. 使用 `getWorksheetInfo` 取得欄位詳情
2. 格式化並呈現結構
3. 根據欄位建議資料操作

### AI 整合最佳實踐

1. **清晰的檔案路徑**：使用檔案的絕對路徑
2. **描述性工作表名稱**：清楚命名工作表以便 AI 理解
3. **一致的欄位標題**：使用清晰、描述性的欄位名稱
4. **啟用 AI 參數**：允許 AI 控制操作和工作表選擇
5. **錯誤上下文**：AI 將自然地處理和解釋錯誤

## 🔧 進階功能

### 自動欄位映射

節點會自動從標題列（第 1 列）偵測欄位，並相應地映射您的 JSON 資料：

```javascript
// Excel 標題：Name | Email | Phone | Status

// 您的輸入
{
  "Name": "張三",
  "Email": "zhang@example.com",
  "Status": "Active"
}

// 自動映射到正確的欄位
// Phone 將保持空白
```

### 智能資料類型

- **字串**：自動處理
- **數字**：保留為數字類型
- **日期**：由 ExcelJS 處理
- **公式**：存在時保留
- **空儲存格**：返回為空字串

### 錯誤處理

```javascript
// 錯誤回應格式
{
  "error": "Column 'InvalidColumn' not found",
  "operation": "filterRows",
  "resource": "row"
}
```

在節點設定中啟用「失敗時繼續」以在工作流程中優雅地處理錯誤。

## ⚙️ 設定選項

### 檔案路徑 vs 二進位資料

**檔案路徑模式：**
- 最適合：伺服器端操作、排程工作流程
- 優點：直接檔案存取、自動儲存支援
- 缺點：需要檔案系統存取

**二進位資料模式：**
- 最適合：處理上傳的檔案、工作流程資料
- 優點：適用於任何檔案來源、可攜性
- 缺點：必須手動處理檔案儲存

### 自動儲存

啟用時（僅限檔案路徑模式）：
- 變更會自動儲存到原始檔案
- 停用以在儲存前預覽/驗證

## 💡 範例

### 範例 1：資料匯入工作流程

```
HTTP 請求（上傳）→ Excel CRUD（附加列）→ Slack（通知）
```

### 範例 2：資料驗證

```
排程觸發器 → Excel CRUD（讀取列）→ If（驗證）→ Excel CRUD（更新列）
```

### 範例 3：AI 驅動的資料輸入

```
AI Agent 聊天 → Excel CRUD（多個操作）→ 回應
```

### 範例 4：報表生成

```
Excel CRUD（讀取列）→ 彙總 → Excel CRUD（建立工作表）→ 電子郵件
```

## 🧪 測試

```bash
# 執行所有測試
npm test

# 執行覆蓋率測試
npm test -- --coverage

# 監視模式
npm test -- --watch
```

## 🛠️ 開發

```bash
# 複製並安裝
git clone https://github.com/code4Copilot/n8n-nodes-excel-ai.git
cd n8n-nodes-excel-ai
npm install

# 建置
npm run build

# 開發的監視模式
npm run dev

# 程式碼檢查
npm run lint
npm run lintfix
```

## 📝 變更日誌

### v1.0.3 (2026-01-05) - 目前版本
- 🔄 **重大變更**：將 `尋找列` 操作替換為更強大的 `過濾列`
- ✨ **過濾列功能**：
  - 支援 12 種進階運算符（equals、notEquals、contains、notContains、greaterThan、greaterOrEqual、lessThan、lessOrEqual、startsWith、endsWith、isEmpty、isNotEmpty）
  - 支援 AND/OR 邏輯的多重過濾條件
  - 結果自動追蹤行號
  - 支援檔案路徑和二進位資料模式
  - 支援複雜過濾場景（範圍、文字搜尋、空值檢查）
- 📝 更新文件，提供完整的過濾列範例
- 🧪 新增 14 個過濾列功能測試案例
- 📚 增強 AI Agent 範例，展示過濾列用法

### v1.0.2
- 🐛 錯誤修復和效能改進
- 📝 文件更新

### v1.0.1
- 🔧 小幅改進
- 📝 README 增強

### v1.0.0
- ✨ 新增完整的 AI Agent 整合（`usableAsTool: true`）
- ✨ 自動欄位偵測和映射
- ✨ 增強的 JSON 資料處理
- 📝 改進 AI 的參數描述
- 🐛 更好的錯誤訊息
- 📚 全面的 AI 使用文件
- ➕ 新增工作表操作（列表、建立、刪除、重新命名、複製、取得資訊）
- ➕ 二進位資料支援
- ➕ 自動儲存選項
- ➕ 插入列操作
- ➕ 尋找列操作（在 v1.0.3 中已棄用）

### v0.9.0
- 🎉 初始版本
- ✅ 基本 CRUD 操作（新增、讀取、更新、刪除）
- ✅ 檔案路徑支援
- ✅ 使用 ExcelJS 處理 Excel 檔案

## 🤝 貢獻

歡迎貢獻！請隨時提交 Pull Request。

1. Fork 儲存庫
2. 建立您的功能分支（`git checkout -b feature/AmazingFeature`）
3. 提交您的變更（`git commit -m 'Add some AmazingFeature'`）
4. 推送到分支（`git push origin feature/AmazingFeature`）
5. 開啟一個 Pull Request

## 📄 授權

此專案採用 MIT 授權 - 詳情請參閱 [LICENSE](LICENSE) 檔案。

```
MIT License

Copyright (c) 2024 Your Name

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## 🙏 致謝

- 為 [n8n](https://n8n.io) 工作流程自動化平台建立
- 使用 [ExcelJS](https://github.com/exceljs/exceljs) 進行 Excel 檔案處理
- 受 n8n 社群啟發

## 🆘 支援

- **問題**：[GitHub Issues](https://github.com/code4Copilot/n8n-nodes-excel-ai/issues)
- **討論**：[GitHub Discussions](https://github.com/code4Copilot/n8n-nodes-excel-ai/discussions)
- **n8n 社群**：[n8n 社群論壇](https://community.n8n.io)

## 💖 顯示您的支持

如果您覺得這個節點有用，請考慮：
- ⭐ 為儲存庫加星標
- 🐛 報告錯誤
- 💡 建議新功能
- 📝 改進文件
- 🔧 貢獻程式碼

---

**用 ❤️ 為 n8n 社群製作**
