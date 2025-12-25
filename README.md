# n8n-nodes-excel-ai

![n8n](https://img.shields.io/badge/n8n-Community%20Node-FF6D5A)
![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Version](https://img.shields.io/npm/v/n8n-nodes-excel-ai)

A powerful n8n community node for performing CRUD (Create, Read, Update, Delete) operations on Excel files with **AI Agent support**. Works seamlessly with n8n AI Agents for natural language Excel operations.

## ✨ Features

### 🤖 AI Agent Integration
- **Native AI Support**: Works as AI Agent Tool (`usableAsTool: true`)
- **Natural Language**: AI can interact with Excel files using conversational queries
- **Auto Column Mapping**: Automatically detects and maps columns from your spreadsheets
- **Smart Data Handling**: Accepts JSON data with intelligent field mapping

### 📊 Complete CRUD Operations
- **Read**: Query data with filters and pagination
- **Create**: Add new rows with automatic column mapping
- **Update**: Modify existing rows with partial updates
- **Delete**: Remove rows by row number
- **Search**: Find rows with advanced matching (exact, contains, starts/ends with)

### 🗂️ Worksheet Management
- **List Worksheets**: Get all sheets in a workbook
- **Create Worksheet**: Add new sheets with optional initial data
- **Delete Worksheet**: Remove sheets from workbook
- **Rename Worksheet**: Rename existing worksheets
- **Copy Worksheet**: Duplicate worksheets with all data and formatting
- **Get Worksheet Info**: Retrieve detailed worksheet information including columns

### 🔄 Flexible Input Modes
- **File Path**: Work with Excel files from your file system
- **Binary Data**: Process files from previous workflow steps

## 📦 Installation

### Option 1: npm (Recommended)

```bash
# Navigate to n8n custom nodes directory
cd ~/.n8n/nodes

# Install the package
npm install n8n-nodes-excel-ai
```

### Option 2: Docker

Add to your `docker-compose.yml`:

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

Then install inside the container:

```bash
docker exec -it <n8n-container> npm install n8n-nodes-excel-ai
```

### Option 3: Manual Installation

```bash
# Clone the repository
git clone https://github.com/code4Copilot/n8n-nodes-excel-ai.git
cd n8n-nodes-excel-ai

# Install dependencies
npm install

# Build the node
npm run build

# Link to n8n
npm link
cd ~/.n8n/nodes
npm link n8n-nodes-excel-ai
```

## 🚀 Quick Start

### Basic Usage

#### 1. Read Data from Excel

```javascript
// Node Configuration
Resource: Row
Operation: Read Rows
File Path: /data/customers.xlsx
Sheet Name: Customers
Start Row: 2
End Row: 0 (read all)

// Output
[
  { "_rowNumber": 2, "Name": "John Doe", "Email": "john@example.com" },
  { "_rowNumber": 3, "Name": "Jane Smith", "Email": "jane@example.com" }
]
```

#### 2. Add New Row

```javascript
// Node Configuration
Resource: Row
Operation: Append Row
File Path: /data/customers.xlsx
Sheet Name: Customers
Row Data: {"Name": "Bob Wilson", "Email": "bob@example.com", "Status": "Active"}

// Output
{
  "success": true,
  "operation": "appendRow",
  "rowNumber": 15,
  "message": "Row added successfully at row 15"
}
```

#### 3. Find Rows

```javascript
// Node Configuration
Resource: Row
Operation: Find Rows
File Path: /data/customers.xlsx
Sheet Name: Customers
Search Column: Status
Search Value: Active
Match Type: exact

// Output
[
  { "_rowNumber": 2, "Name": "John Doe", "Status": "Active" },
  { "_rowNumber": 5, "Name": "Alice Johnson", "Status": "Active" }
]
```

### 🤖 AI Agent Usage Examples

#### Example 1: Natural Language Query

**User:** "Show me all customers from the customers.xlsx file"

**AI Agent Execution:**
```javascript
{
  "operation": "readRows",
  "filePath": "/data/customers.xlsx",
  "sheetName": "Customers",
  "startRow": 2,
  "endRow": 0
}
```

#### Example 2: Add Data via AI

**User:** "Add a new customer named Sarah Johnson with email sarah@example.com"

**AI Agent Execution:**
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

#### Example 3: Search with AI

**User:** "Find all active customers in Boston"

**AI Agent Execution:**
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

Then filter by City in subsequent operation.

## 📚 Operations Reference

### Row Operations

#### Read Rows
- **Purpose**: Read data from Excel file
- **Parameters**:
  - `startRow`: Starting row number (default: 2)
  - `endRow`: Ending row number (0 = all rows)
- **Returns**: Array of row objects with `_rowNumber` field

#### Append Row
- **Purpose**: Add new row at the end of the sheet
- **Parameters**:
  - `rowData`: JSON object with column names and values
- **Returns**: Success status and new row number

#### Insert Row
- **Purpose**: Insert row at specific position
- **Parameters**:
  - `rowNumber`: Position to insert
  - `rowData`: JSON object with column names and values
- **Returns**: Success status and row number

#### Update Row
- **Purpose**: Update existing row
- **Parameters**:
  - `rowNumber`: Row to update
  - `updatedData`: JSON object with fields to update
- **Returns**: Success status and updated fields

#### Delete Row
- **Purpose**: Remove specific row
- **Parameters**:
  - `rowNumber`: Row to delete (cannot be 1 - header row)
- **Returns**: Success status

#### Find Rows
- **Purpose**: Search for rows matching criteria
- **Parameters**:
  - `searchColumn`: Column to search in
  - `searchValue`: Value to search for
  - `matchType`: exact | contains | startsWith | endsWith
  - `returnRowNumbers`: Return only row numbers (default: false)
- **Returns**: Array of matching rows or row numbers

### Worksheet Operations

#### List Worksheets
- **Purpose**: Get all worksheets in workbook
- **Parameters**:
  - `includeHidden`: Include hidden sheets (default: false)
- **Returns**: Array of worksheet info

#### Create Worksheet
- **Purpose**: Create new worksheet
- **Parameters**:
  - `newSheetName`: Name for new sheet
  - `initialData`: Optional array of objects for initial data
- **Returns**: Success status and sheet info

#### Delete Worksheet
- **Purpose**: Remove worksheet from workbook
- **Parameters**:
  - `worksheetName`: Name of sheet to delete
- **Returns**: Success status

#### Rename Worksheet
- **Purpose**: Rename an existing worksheet
- **Parameters**:
  - `worksheetName`: Current name of the sheet
  - `newSheetName`: New name for the sheet
- **Returns**: Success status with old and new names

**Example:**
```javascript
{
  "worksheetOperation": "renameWorksheet",
  "filePath": "/data/reports.xlsx",
  "worksheetName": "Sheet1",
  "newSheetName": "Sales_2024",
  "autoSave": true
}
```

#### Copy Worksheet
- **Purpose**: Duplicate a worksheet with all data and formatting
- **Parameters**:
  - `worksheetName`: Name of sheet to copy
  - `newSheetName`: Name for the copied sheet
- **Returns**: Success status with source and new sheet names, row count

**Example:**
```javascript
{
  "worksheetOperation": "copyWorksheet",
  "filePath": "/data/templates.xlsx",
  "worksheetName": "Template_2024",
  "newSheetName": "Template_2025",
  "autoSave": true
}
```

**Output:**
```javascript
{
  "success": true,
  "operation": "copyWorksheet",
  "sourceName": "Template_2024",
  "newName": "Template_2025",
  "rowCount": 50
}
```

#### Get Worksheet Info
- **Purpose**: Retrieve detailed information about a worksheet
- **Parameters**:
  - `worksheetName`: Name of the sheet
- **Returns**: Detailed worksheet information including columns

**Example:**
```javascript
{
  "worksheetOperation": "getWorksheetInfo",
  "filePath": "/data/database.xlsx",
  "worksheetName": "Users"
}
```

**Output:**
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
    // ... more columns
  ]
}
```

## 🤖 Using with AI Agents

### Setup

This node is designed to work seamlessly with n8n AI Agents. The node is configured with `usableAsTool: true`, making it automatically available to AI Agents.

### Enabling AI Parameters

1. In the node configuration, look for parameters with a ✨ sparkle icon
2. Click the ✨ icon next to any parameter to enable AI auto-fill
3. The AI Agent can now automatically set values for that parameter

### AI Agent Examples

#### Example 1: Natural Language Data Operations

**Workflow Setup:**
```
AI Agent → Excel AI Node
```

**User Query:** "Get all customers from the Excel file and show me those in New York"

**AI Agent Actions:**
1. Uses Excel AI to read all rows
2. Filters results for New York customers
3. Returns formatted results

#### Example 2: Multi-Step Operations

**User Query:** "Copy the 2024 template sheet to create a 2025 version, then add January data"

**AI Agent Actions:**
1. Uses `copyWorksheet` operation to duplicate the sheet
2. Uses `appendRow` to add new data rows
3. Confirms completion

#### Example 3: Data Analysis

**User Query:** "Show me the structure of the Users worksheet"

**AI Agent Actions:**
1. Uses `getWorksheetInfo` to retrieve column details
2. Formats and presents the structure
3. Suggests data operations based on columns

### Best Practices for AI Integration

1. **Clear File Paths**: Use absolute paths for files
2. **Descriptive Sheet Names**: Name worksheets clearly for AI understanding
3. **Consistent Column Headers**: Use clear, descriptive column names
4. **Enable AI Parameters**: Allow AI to control operation and sheet selection
5. **Error Context**: AI will handle and explain errors naturally

## 🔧 Advanced Features

### Automatic Column Mapping

The node automatically detects columns from the header row (row 1) and maps your JSON data accordingly:

```javascript
// Excel Headers: Name | Email | Phone | Status

// Your Input
{
  "Name": "John Doe",
  "Email": "john@example.com",
  "Status": "Active"
}

// Automatically mapped to correct columns
// Phone will be left empty
```

### Smart Data Types

- **Strings**: Handled automatically
- **Numbers**: Preserved as numeric types
- **Dates**: Handled by ExcelJS
- **Formulas**: Preserved when present
- **Empty cells**: Returned as empty strings

### Error Handling

```javascript
// Error Response Format
{
  "error": "Column 'InvalidColumn' not found",
  "operation": "findRows",
  "resource": "row"
}
```

Enable "Continue on Fail" in node settings to handle errors gracefully in workflows.

## ⚙️ Configuration Options

### File Path vs Binary Data

**File Path Mode:**
- Best for: Server-side operations, scheduled workflows
- Pros: Direct file access, auto-save support
- Cons: Requires file system access

**Binary Data Mode:**
- Best for: Processing uploaded files, workflow data
- Pros: Works with any file source, portable
- Cons: Must handle file saving manually

### Auto Save

When enabled (File Path mode only):
- Changes are automatically saved to the original file
- Disable for preview/validation before saving

## 💡 Examples

### Example 1: Data Import Workflow

```
HTTP Request (Upload) → Excel CRUD (Append Row) → Slack (Notify)
```

### Example 2: Data Validation

```
Schedule Trigger → Excel CRUD (Read Rows) → If (Validate) → Excel CRUD (Update Row)
```

### Example 3: AI-Powered Data Entry

```
AI Agent Chat → Excel CRUD (Multiple Operations) → Response
```

### Example 4: Report Generation

```
Excel CRUD (Read Rows) → Aggregate → Excel CRUD (Create Worksheet) → Email
```

## ?�� Testing

```bash
# Run all tests
npm test

# Run with coverage
npm test -- --coverage

# Watch mode
npm test -- --watch
```

## 🛠️ Development

```bash
# Clone and install
git clone https://github.com/code4Copilot/n8n-nodes-excel-ai.git
cd n8n-nodes-excel-ai
npm install

# Build
npm run build

# Watch mode for development
npm run dev

# Lint
npm run lint
npm run lintfix
```

## 📝 Changelog

### v1.3.0 (Latest)
- ✨ Added full AI Agent integration (`usableAsTool: true`)
- ✨ Automatic column detection and mapping
- ✨ Enhanced JSON data handling
- 📝 Improved parameter descriptions for AI
- 🐛 Better error messages
- 📚 Comprehensive AI usage documentation

### v1.2.0
- Added worksheet operations
- Binary data support
- Auto-save option

### v1.1.0
- Find rows operation
- Advanced search with match types
- Insert row operation

### v1.0.0
- Initial release
- Basic CRUD operations
- File path support

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

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

## 🙏 Acknowledgments

- Built for the [n8n](https://n8n.io) workflow automation platform
- Uses [ExcelJS](https://github.com/exceljs/exceljs) for Excel file processing
- Inspired by the n8n community

## 🆘 Support

- **Issues**: [GitHub Issues](https://github.com/code4Copilot/n8n-nodes-excel-ai/issues)
- **Discussions**: [GitHub Discussions](https://github.com/code4Copilot/n8n-nodes-excel-ai/discussions)
- **n8n Community**: [n8n Community Forum](https://community.n8n.io)

## 💖 Show Your Support

If you find this node useful, please consider:
- ⭐ Starring the repository
- 🐛 Reporting bugs
- 💡 Suggesting new features
- 📝 Improving documentation
- 🔧 Contributing code

---

**Made with ❤️ for the n8n community**
