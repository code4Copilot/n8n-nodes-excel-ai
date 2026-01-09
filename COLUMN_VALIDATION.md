# 欄位驗證功能說明

## 功能摘要

此更新增強了 Excel AI 節點的欄位驗證功能，確保在操作 Excel 檔案時提供更清晰的錯誤訊息和警告資訊。

## 新增功能

### 1. UpdateRow 操作 - 返回 skippedFields 資訊

當使用 `updateRow` 操作更新資料時，如果提供的欄位名稱不存在於工作表中：

**行為**：
- ✅ 不會中斷執行
- ✅ 自動跳過不存在的欄位
- ✅ 只更新存在的欄位
- ✅ 在結果中返回 `skippedFields` 陣列
- ✅ 在結果中包含 `warning` 訊息說明哪些欄位被跳過

**返回結果範例**：
```json
{
  "success": true,
  "operation": "updateRow",
  "rowNumber": 2,
  "updatedFields": ["Age", "Email"],
  "skippedFields": ["InvalidColumn1", "InvalidColumn2"],
  "warning": "The following fields were not found in the worksheet and were skipped: InvalidColumn1, InvalidColumn2",
  "message": "Row 2 updated successfully"
}
```

**使用場景**：
- 使用 Expression 動態生成更新資料時
- 不確定工作表結構是否完全匹配時
- 需要容錯處理的批次更新操作

---

### 2. FilterRows 操作 - 嚴格驗證欄位存在性

當使用 `filterRows` 操作篩選資料時，如果過濾條件中的欄位不存在於工作表中：

**行為**：
- ❌ 立即拋出錯誤，停止執行
- ✅ 提供清晰的錯誤訊息，說明哪些欄位不存在
- ✅ 列出工作表中所有可用的欄位名稱

**錯誤訊息範例**：
```
Filter condition error: The following field(s) do not exist in the worksheet: InvalidField1, InvalidField2. 
Available fields are: Name, Age, Email
```

**設計理念**：
- 過濾不存在的欄位沒有意義，應該直接報錯
- 幫助使用者快速發現錯誤並修正
- 避免產生錯誤的篩選結果

**適用模式**：
- ✅ File Path 模式
- ✅ Binary Data 模式

---

## 錯誤處理統一格式

所有欄位驗證錯誤都使用 `NodeOperationError` 格式，確保：
1. 錯誤訊息清晰易懂
2. 與 n8n 平台的錯誤處理機制整合
3. 支援 `continueOnFail` 模式

---

## 測試覆蓋

新增了完整的單元測試覆蓋以下場景：

### UpdateRow 測試
- ✅ 部分欄位不存在時返回 skippedFields
- ✅ 所有欄位都存在時正常更新
- ✅ 所有欄位都不存在時返回完整 skippedFields

### FilterRows 測試
- ✅ 單一欄位不存在時拋出錯誤
- ✅ 多個欄位不存在時拋出錯誤並列出所有無效欄位
- ✅ Binary Data 模式下的欄位驗證
- ✅ 所有欄位存在時正常篩選
- ✅ 錯誤訊息包含可用欄位列表

### 錯誤處理測試
- ✅ continueOnFail 模式下返回錯誤資訊而非拋出異常

**測試結果**：55/55 測試通過 ✅

---

## 不影響的操作

以下操作不受此次更新影響，保持原有行為：

- ✅ `appendRow` - 添加新行
- ✅ `insertRow` - 插入行
- ✅ `readRows` - 讀取行
- ✅ `deleteRow` - 刪除行
- ✅ 所有 Worksheet 相關操作
- ✅ Binary Data 模式的其他操作

---

## 使用建議

### UpdateRow
- 適合需要容錯處理的場景
- 建議檢查返回結果中的 `skippedFields` 和 `warning`
- 如需嚴格驗證，可以在後續節點中檢查 `skippedFields` 是否為空

### FilterRows
- 使用前確認欄位名稱正確
- 發生錯誤時檢查錯誤訊息中的「Available fields」列表
- 在 Binary Data 模式下更需要注意欄位名稱拼寫

---

## 版本資訊

- **修改日期**：2026-01-09
- **影響版本**：1.0.5+
- **向後相容性**：✅ 完全相容，只新增功能不改變現有行為
