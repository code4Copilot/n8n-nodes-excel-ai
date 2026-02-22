# 發布檢查清單 - Version 1.0.11

## ✅ 發布前檢查

### 1. 版本更新
- [x] package.json 版本更新為 1.0.11
- [x] CHANGELOG.md 新增 1.0.11 版本記錄
- [x] README.md 更新版本提示
- [x] README.zh-TW.md 更新版本提示

### 2. 文檔更新
- [x] README.md 更新（Clear Rows 功能說明）
- [x] README.zh-TW.md 更新（Clear Rows 功能說明）

### 3. 程式碼品質
- [x] TypeScript 編譯成功（無錯誤）
- [x] 所有測試通過（84/84）
- [x] 測試覆蓋率 100%

### 4. 建置驗證
- [ ] `npm run build` 成功
- [ ] dist/ 目錄已生成
- [ ] 圖示檔案已複製

### 5. 功能驗證
- [x] 新功能實作完成（Clear Rows 操作）
- [x] 向下相容性確認
- [x] 無破壞性變更

---

## 📦 發布步驟

### 步驟 1: 最終檢查
```bash
# 確認版本號
cat package.json | grep "version"

# 執行所有測試
npm test

# 執行建置
npm run build

# 檢查 lint
npm run lint
```

### 步驟 2: Git 提交
```bash
# 查看變更
git status

# 加入所有變更
git add .

# 提交變更
git commit -m "feat: release version 1.0.11

- Add Clear Rows operation (keep header option)
- Add tests for Clear Rows (84 tests total)
- Update documentation (README, README.zh-TW, CHANGELOG)
- All tests passing (84/84, 100% coverage)"

# 推送到遠端
git push origin main
```

### 步驟 3: 建立 Git Tag
```bash
# 創建版本標籤
git tag -a v1.0.11 -m "Version 1.0.11 - Clear Rows Operation

New Features:
- Clear Rows operation: delete all data rows while keeping header
- keepHeader option (default: true) for flexible clearing
- Returns clearedRows count in response
- Full AI Agent support

Testing:
- 84 tests passing (new Clear Rows tests added)
- 100% feature coverage

Documentation:
- Updated README (EN & ZH-TW)
- Complete CHANGELOG
"

# 推送標籤
git push origin v1.0.11
```

### 步驟 4: 發布到 npm
```bash
# 登入 npm（如果尚未登入）
npm login

# 確認 npm 帳號
npm whoami

# 執行發布（會自動執行 prepublishOnly）
npm publish

# 驗證發布成功
npm view n8n-nodes-excel-ai version
```

### 步驟 5: 創建 GitHub Release
1. 前往 GitHub Repository
2. 點擊 "Releases" → "Create a new release"
3. 選擇標籤: `v1.0.11`
4. 發布標題: `v1.0.11 - Clear Rows Operation`
5. 發布說明: 複製以下內容

```markdown
## 🎉 Version 1.0.11 - Clear Rows Operation

### 🆕 New Features

#### Clear Rows Operation
- New **Clear Rows** operation clears all data rows while optionally keeping the header row
- `keepHeader` option (default: `true`) preserves the first row when clearing
- Returns `clearedRows` count and descriptive `message` in the output
- Works with both File Path and Binary Data modes
- Full AI Agent support

### 📊 Testing

- ✅ New tests for Clear Rows operation
- ✅ **84 total tests** all passing (100% success rate)

### 🔄 Upgrade Notes

- **Fully backward compatible** - no breaking changes
- Existing workflows continue to work as before

### 📦 Installation

```bash
npm install n8n-nodes-excel-ai@1.0.11
```

**Full Changelog**: https://github.com/code4Copilot/n8n-nodes-excel-ai/compare/v1.0.10...v1.0.11
```

6. 點擊 "Publish release"

---

## 🔍 發布後驗證

### 驗證 npm 發布
```bash
# 檢查最新版本
npm view n8n-nodes-excel-ai version

# 檢查完整資訊
npm info n8n-nodes-excel-ai
```

### 驗證 GitHub Release
- [ ] GitHub Releases 頁面顯示 v1.0.11
- [ ] 標籤已正確創建
- [ ] 發布說明完整顯示

### 驗證安裝
```bash
npm install n8n-nodes-excel-ai@1.0.11
npm list n8n-nodes-excel-ai
```

---

## ✅ 發布完成確認

- [ ] npm 上可以查到 1.0.11 版本
- [ ] GitHub 有 v1.0.11 的 Release
- [ ] Git 有 v1.0.11 的 Tag
- [ ] 所有文檔都已更新
- [ ] 測試全部通過
- [ ] 可以成功安裝新版本
