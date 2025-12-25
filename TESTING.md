# 測試指南

本文件說明如何執行 n8n-nodes-excel-ai 的單元測試。

## 前置準備

### 1. 安裝依賴

首先確保所有依賴都已安裝：

```bash
npm install
```

這會安裝以下測試相關套件：
- `jest`: 測試框架
- `ts-jest`: TypeScript 支援
- `@types/jest`: Jest 的 TypeScript 型別定義
- `@jest/globals`: Jest 全域變數型別定義

## 執行測試

### 基本測試命令

執行所有測試：

```bash
npm test
```

或使用 Jest 直接執行：

```bash
npx jest
```

### 監視模式

在開發過程中，可以使用監視模式自動執行測試：

```bash
npm run test:watch
```

這會在檔案變更時自動重新執行測試。

### 測試覆蓋率

生成測試覆蓋率報告：

```bash
npm run test:coverage
```

這會：
1. 執行所有測試
2. 生成覆蓋率報告
3. 在 `coverage/` 目錄下建立 HTML 報告
4. 在終端顯示摘要資訊

查看詳細的 HTML 報告：

```bash
# Windows
start coverage/lcov-report/index.html

# macOS
open coverage/lcov-report/index.html

# Linux
xdg-open coverage/lcov-report/index.html
```

### 執行特定測試

執行特定測試檔案：

```bash
npx jest ExcelAI.node.test.ts
```

執行包含特定名稱的測試：

```bash
npx jest -t "Append Row"
```

### 詳細輸出

獲得更詳細的測試輸出：

```bash
npx jest --verbose
```

## 測試結構

測試檔案位於：`nodes/ExcelAI/ExcelAI.node.test.ts`

測試涵蓋以下功能：

### 1. 節點屬性測試
- 驗證節點的基本屬性（名稱、版本、類別）
- 驗證所有必要的屬性都已定義
- 驗證可用的操作選項

### 2. 載入選項方法測試
- **getWorksheets**: 測試工作表名稱載入
- **getColumns**: 測試欄位清單載入

### 3. 資料列操作測試（檔案路徑模式）
- **Append Row**: 新增新列
- **Read Rows**: 讀取資料列
- **Update Row**: 更新資料
- **Delete Row**: 刪除資料
- **Find Rows**: 尋找資料（包含完整資料和行號兩種模式）
- **Insert Row**: 插入資料

### 4. 工作表操作測試
- **Create Worksheet**: 建立新工作表（含/不含初始資料）
- **Delete Worksheet**: 刪除工作表
- **List Worksheets**: 列出所有工作表

### 5. 錯誤處理測試
- 工作表不存在的錯誤處理
- 檔案讀取失敗的錯誤處理
- Continue on Fail 模式的測試

### 6. Binary Data 模式測試
- 測試從 Binary Data 讀取 Excel 檔案

## 測試覆蓋率目標

建議的覆蓋率目標：
- **語句覆蓋率（Statements）**: > 80%
- **分支覆蓋率（Branches）**: > 75%
- **函數覆蓋率（Functions）**: > 80%
- **行數覆蓋率（Lines）**: > 80%

## 常見問題

### Q: 測試失敗怎麼辦？

1. 檢查錯誤訊息，確認失敗的原因
2. 確認程式碼是否有語法錯誤
3. 檢查 mock 物件是否正確設定
4. 執行 `npm install` 重新安裝依賴

### Q: 如何偵錯測試？

在 VSCode 中：
1. 在測試檔案中設定中斷點
2. 按 F5 或使用偵錯面板
3. 選擇 "Jest" 設定（需要配置 launch.json）

或使用 Node 偵錯器：

```bash
node --inspect-brk node_modules/.bin/jest --runInBand
```

### Q: 如何跳過特定測試？

跳過特定測試：

```typescript
it.skip('should do something', () => {
  // 這個測試會被略過
});
```

只執行特定測試：

```typescript
it.only('should do something', () => {
  // 只會執行這個測試
});
```

## 持續整合（CI）

在 CI/CD 管線中執行測試：

```yaml
# GitHub Actions 範例
- name: Run tests
  run: npm test

- name: Generate coverage
  run: npm run test:coverage

- name: Upload coverage to Codecov
  uses: codecov/codecov-action@v3
```

## 撰寫新測試

當新增新功能時，請遵循以下步驟：

1. 在 `ExcelAI.node.test.ts` 中新增相應的測試區塊
2. 使用 `describe` 區塊組織相關的測試
3. 使用 `it` 或 `test` 撰寫個別測試案例
4. 確保 mock 物件正確設定
5. 使用描述性的測試名稱
6. 測試正常流程和錯誤流程

範例：

```typescript
describe('New Feature', () => {
  it('should work correctly with valid input', async () => {
    // Arrange
    mockExecuteFunctions.getNodeParameter = jest.fn((name) => {
      return { /* mock parameters */ };
    });

    // Act
    const result = await excelAI.execute.call(mockExecuteFunctions);

    // Assert
    expect(result[0]).toHaveLength(1);
    expect(result[0][0].json).toHaveProperty('success', true);
  });

  it('should handle errors gracefully', async () => {
    // Test error case
  });
});
```

## 參考資源

- [Jest 官方文件](https://jestjs.io/)
- [ts-jest 文件](https://kulshekhar.github.io/ts-jest/)
- [n8n 節點開發指南](https://docs.n8n.io/integrations/creating-nodes/)

---

如有任何問題，請參考專案 README.md 或聯繫維護者。
