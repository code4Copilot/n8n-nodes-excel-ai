# n8n-nodes-excel-ai

![n8n](https://img.shields.io/badge/n8n-Community%20Node-FF6D5A)
![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Version](https://img.shields.io/npm/v/n8n-nodes-excel-ai)

一個強大的 n8n 社群節點，用於對 Excel 檔案執行 CRUD（新增、讀取、更新、刪除）操作，並支援 **AI Agent 整合**，可與 n8n AI Agents 無縫整合，實現自然語言的 Excel 操作。

## ✨ 功能特色

### 🤖 AI Agent 整合
- **原生 AI 支援**：可作為 AI Agent 工具使用（`usableAsTool: true`）
- **自然語言**：AI 可使用對話式查詢與 Excel 檔案互動
- **自動欄位映射**：自動偵測並映射試算表的欄位
- **智能資料處理**：接受 JSON 資料並進行智能欄位映射

### 📊 完整的 CRUD 操作
- **讀取**：使用篩選器和分頁查詢資料
- **新增**：使用自動欄位映射新增列
- **更新**：部分更新現有列
- **刪除**：依行號刪除列
- **搜尋**：使用進階匹配尋找行（精確匹配、包含、開頭/結尾匹配）

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

#### 3. 尋找列
```javascript
// 節點設定
資源：列
操作：尋找列
檔案路徑：/data/customers.xlsx
工作表名稱：Customers
搜尋欄位：Status
搜尋值：Active
匹配類型：exact

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

#### 範例 3：使用 AI 搜尋

**使用者：** "找出波士頓的所有活躍客戶"

**AI Agent 執行：**
```javascript
{
  "operation": "findRows",
  "filePath": "/data/customers.xlsx",
  "sheetName": "Customers",
  "searchColumn": "Status",
  "searchValue": "Active",
  "matchType": "exact"
}
```

然後在後續操作中按城市篩選。

## 📚 操作參考

### 列操作

#### 讀取列
- **用途**：從 Excel 檔案讀取資料
- **參數**：
  - `startRow`：起始行號（預設：2）
  - `endRow`：結束行號（0 = 所有列）
- **返回**：帶有 `_rowNumber` 欄位的列物件陣列

#### 附加列
- **用途**：在工作表末端新增列
- **參數**：
  - `rowData`：包含欄位名稱和值的 JSON 物件
- **返回**：成功狀態和新列號

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

#### 刪除列
- **用途**：移除特定列
- **參數**：
  - `rowNumber`：要刪除的列（不能是 1 - 標題列）
- **返回**：成功狀態

#### 尋找列
- **用途**：搜尋符合條件的列
- **參數**：
  - `searchColumn`：要搜尋的欄位
  - `searchValue`：要搜尋的值
  - `matchType`：exact | contains | startsWith | endsWith
  - `returnRowNumbers`：只返回列號（預設：false）
- **返回**：符合列的陣列或列號

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
  "operation": "findRows",
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

### v1.0.0（最新）
- ✨ 新增完整的 AI Agent 整合（`usableAsTool: true`）
- ✨ 自動欄位偵測和映射
- ✨ 增強的 JSON 資料處理
- 📝 改進 AI 的參數描述
- 🐛 更好的錯誤訊息
- 📚 全面的 AI 使用文件
- 新增工作表操作
- 二進位資料支援
- 自動儲存選項
- 尋找列操作
- 帶匹配類型的進階搜尋
- 插入列操作
- 初始版本
- 基本 CRUD 操作
- 檔案路徑支援

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
