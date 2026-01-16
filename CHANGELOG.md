# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.7] - 2026-01-16

### Added
- **Automatic Type Conversion**: Intelligent automatic type conversion for field input values
  - **Numbers**: Converts string numbers to numeric values (e.g., "123" → 123, "45.67" → 45.67)
  - **Booleans**: Converts string booleans to boolean values (e.g., "true" → true, "FALSE" → false, case-insensitive)
  - **Dates**: Converts ISO 8601 date strings to Date objects (e.g., "2024-01-15" → Date, "2024-01-15T10:30:00Z" → Date)
  - **Null Values**: Converts "null", empty strings, and whitespace-only strings to null
  - **Smart Detection**: Preserves regular strings and already-converted values as-is
  - Works in all row operations: Append Row, Insert Row, and Update Row

### Improved
- Enhanced AI Agent integration with automatic type handling
- Users can now input values directly as strings without manual type conversion
- Better data consistency in Excel files with proper type storage
- Negative numbers and decimal values are properly handled

### Tests
- Added 7 comprehensive tests for automatic type conversion
- All 62 tests passing with 100% success rate
- Test coverage for numbers, booleans, dates, null values, and edge cases

### Documentation
- Added TYPE_CONVERSION_DEMO.md with detailed usage examples
- Added TYPE_CONVERSION_VERIFICATION_REPORT.md with complete test results
- Created demo scripts for type conversion functionality

## [1.0.6] - 2026-01-09

### Added
- **Column Validation for FilterRows**: Now validates that all filter condition fields exist in the worksheet before executing
  - Throws clear error message listing invalid fields and available fields
  - Works in both File Path and Binary Data modes
  - Prevents filtering on non-existent columns which would produce incorrect results

- **UpdateRow Enhanced Feedback**: Now tracks and reports skipped fields when updating rows
  - Returns `updatedFields` array with successfully updated column names
  - Returns `skippedFields` array when columns don't exist in the worksheet
  - Includes `warning` message explaining which fields were skipped
  - Continues execution gracefully instead of failing

### Improved
- Unified error message format using `NodeOperationError` for better consistency
- Enhanced error messages now show available fields to help users correct mistakes
- Better support for dynamic Expression values in column names

### Tests
- Added 10 new unit tests for column validation scenarios
- All 55 tests passing with 100% success rate
- Comprehensive coverage of error handling with `continueOnFail` mode

## [1.0.5] - 2026-01-06

### Added
- **Filter Rows Operation**: New advanced filtering functionality with comprehensive operator support
  - 12 operators: equals, notEquals, contains, notContains, greaterThan, greaterOrEqual, lessThan, lessOrEqual, startsWith, endsWith, isEmpty, isNotEmpty
  - Multiple condition support with AND/OR logic
  - Works with both File Path and Binary Data modes
  - Dynamic field selection via dropdown in File Path mode

### Fixed
- Fixed Filter Rows operation duplicating results when multiple input items are present
- Read-only operations (readRows, filterRows) now execute only once instead of repeating for each input item

### Tests
- Added 14 comprehensive tests for Filter Rows operation
- Tests cover all operators, condition logic, and both input modes
- Total test count increased to 45 tests with 100% pass rate

## [1.0.4] - 2026-01-06

### Fixed
- Fixed Filter Rows operation duplicating results when multiple input items are present
- Read-only operations (readRows, filterRows) now execute only once instead of repeating for each input item

### Added
- Added test case to verify Filter Rows does not duplicate results with multiple input items

## [1.0.3] - 2025-12-XX

### Added
- Initial release with full CRUD operations
- AI Agent integration support
- Read, Append, Insert, Update, Delete row operations
- Filter Rows with advanced conditions
- Worksheet management operations
- Support for both file path and binary data modes
- Comprehensive test coverage

## [1.0.2] - 2025-12-XX

### Changed
- Internal improvements

## [1.0.1] - 2025-12-XX

### Changed
- Documentation updates

## [1.0.0] - 2025-12-XX

### Added
- Initial stable release
