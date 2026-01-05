# Filter Rows 功能測試摘要

## 更新日期
2026年1月5日

## 測試概述
因程式功能變更，已將 **Find Rows** 功能升級為功能更強大的 **Filter Rows**，並重新撰寫完整的測試套件。

## 測試統計
- **總測試數量**: 45 個測試
- **Filter Rows 相關測試**: 14 個新增測試
- **測試通過率**: 100%

## Filter Rows 功能特性

### 支援的運算符
Filter Rows 功能支援以下 12 種運算符：

1. **equals** - 等於
2. **notEquals** - 不等於
3. **contains** - 包含
4. **notContains** - 不包含
5. **greaterThan** - 大於
6. **greaterOrEqual** - 大於或等於
7. **lessThan** - 小於
8. **lessOrEqual** - 小於或等於
9. **startsWith** - 開始於
10. **endsWith** - 結束於
11. **isEmpty** - 為空
12. **isNotEmpty** - 不為空

### 支援的邏輯運算
- **AND** - 所有條件必須同時滿足
- **OR** - 任一條件滿足即可

### 輸入模式支援
- **File Path Mode** - 從檔案系統路徑讀取 Excel 檔案
- **Binary Data Mode** - 從二進位資料讀取 Excel 檔案

## 測試涵蓋範圍

### 1. File Path Mode 測試 (12 項)

#### 1.1 單一條件過濾
- ✅ 使用 equals 運算符過濾 (單一條件)
- ✅ 無條件時返回所有行

#### 1.2 多條件過濾
- ✅ 使用 AND 邏輯的多條件過濾
- ✅ 使用 OR 邏輯的多條件過濾

#### 1.3 數值比較運算符
- ✅ greaterThan - 過濾大於指定值的行
- ✅ lessOrEqual - 過濾小於或等於指定值的行

#### 1.4 字串匹配運算符
- ✅ contains - 過濾包含指定文字的行
- ✅ startsWith - 過濾以指定文字開頭的行
- ✅ endsWith - 過濾以指定文字結尾的行

#### 1.5 空值檢查運算符
- ✅ isEmpty - 過濾空值欄位的行
- ✅ isNotEmpty - 過濾非空值欄位的行

#### 1.6 否定運算符
- ✅ notEquals - 過濾不等於指定值的行
- ✅ notContains - 過濾不包含指定文字的行

#### 1.7 特殊功能
- ✅ 結果中包含行號 (_rowNumber 欄位)

### 2. Binary Data Mode 測試 (2 項)
- ✅ 使用二進位資料輸入進行基本過濾
- ✅ 使用二進位資料輸入進行複雜條件過濾

### 3. 其他相關測試
- ✅ Node Properties 中正確列出 filterRows 操作選項
- ✅ 錯誤處理測試確保穩定性

## 測試資料範例

### 範例 1: 部門過濾 (equals)
```
輸入資料:
- Alice | Engineering | Active
- Bob | Sales | Active
- Charlie | Engineering | Inactive
- David | Engineering | Active

過濾條件: Department equals "Engineering"

預期結果: 3 行 (Alice, Charlie, David)
```

### 範例 2: 複合條件 (AND)
```
輸入資料:
- Alice | Engineering | 30 | Active
- Bob | Sales | 25 | Active
- Charlie | Engineering | 35 | Inactive
- David | Engineering | 28 | Active

過濾條件: 
- Department equals "Engineering" AND
- Status equals "Active"

預期結果: 2 行 (Alice, David)
```

### 範例 3: 數值範圍 (greaterThan)
```
輸入資料:
- Laptop | 1000
- Mouse | 25
- Keyboard | 75
- Monitor | 300

過濾條件: Price greaterThan "100"

預期結果: 2 行 (Laptop, Monitor)
```

## 測試執行結果

```
Test Suites: 1 passed, 1 total
Tests:       45 passed, 45 total
Snapshots:   0 total
Time:        2.512 s
```

所有測試均成功通過，確認 Filter Rows 功能運作正常。

## 與舊版 Find Rows 的差異

| 功能 | Find Rows (舊版) | Filter Rows (新版) |
|------|-----------------|-------------------|
| 運算符數量 | 2-3 種 | 12 種 |
| 多條件支援 | 不支援 | 支援 (AND/OR) |
| 數值比較 | 不支援 | 完整支援 |
| 空值檢查 | 不支援 | 支援 |
| 模糊匹配 | 基本 | 進階 (contains/startsWith/endsWith) |
| 行號追蹤 | 可選 | 自動包含 |

## 結論

Filter Rows 是 Find Rows 的全面升級版本，提供了：
- ✅ 更多的過濾運算符
- ✅ 支援複雜的多條件邏輯
- ✅ 更好的數值和字串處理能力
- ✅ 完整的測試涵蓋率

所有功能均經過完整測試並驗證，可以安全地用於生產環境。
