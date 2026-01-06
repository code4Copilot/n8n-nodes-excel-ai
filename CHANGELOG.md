# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
