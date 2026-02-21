# Version 1.0.10 Release Summary

**Release Date**: 2026-02-21  
**Version**: 1.0.10  
**Type**: Patch (UX Improvement)

## Overview

This release improves the worksheet dropdown selection experience in **File Path** mode. Error states now use a dedicated `__error__` sentinel value instead of an empty string, preventing accidental submissions with an invalid sheet name and providing clearer visual feedback with ⚠ warning messages.

## Changes

### Improved: Worksheet Dropdown UX ([ExcelAI.node.ts](nodes/ExcelAI/ExcelAI.node.ts))

| Scenario | Before | After |
|---|---|---|
| No file path entered | Returns `{ name: 'Please specify file path first', value: '' }` | Returns `{ name: '⚠Please specify file path first', value: '__error__' }` |
| No worksheets in file | Returns `{ name: 'No worksheets found in file', value: '' }` | Returns `{ name: '⚠No worksheets found in the specified file', value: '__error__' }` |
| File path error | Returns `{ name: 'Error: <message>', value: '' }` | Returns `{ name: '⚠File path error: Please enter a valid path, then click here to select a worksheet', value: '__error__' }` |

**Key improvements:**
- Error entries now use `value: '__error__'` so they are easily distinguishable from a valid empty selection
- `getColumns()` skips loading when sheet name is `__error__`
- Sheet name resolution converts `__error__` → `''` → falls back to first worksheet
- Updated worksheet field description: *"If error occurs, please reselect after correcting the file path."*

### Tests ([ExcelAI.node.test.ts](nodes/ExcelAI/ExcelAI.node.test.ts))

- Added tests for `__error__` sentinel value handling in sheet name resolution
- All unit tests pass with 100% success rate

## How to Upgrade

```bash
npm install n8n-nodes-excel-ai@1.0.10
```

Or in your n8n custom nodes directory:

```bash
npm update n8n-nodes-excel-ai
```

## Compatibility

- No breaking changes
- All existing workflows continue to work as before
- Requires n8n >= 1.0.0
