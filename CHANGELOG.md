# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2025-11-29

### Added
- Initial release of AlSawife Factory application
- Invoice management system with Excel export functionality
- Client ledger management
- Product pricing based on JSON configuration
- Automatic price calculation based on product, thickness, and dimensions
- Support for complex pricing with range-based pricing

### Fixed
- Fixed SUM formulas in invoice summary table to correctly calculate totals
- Resolved attribute access errors in invoice view
- Fixed import issues with utility modules
- Corrected number input formatting in length field
- Fixed Excel export issues including proper "ش" prefix in product descriptions

### Changed
- Removed discount functionality as requested
- Improved UI with better Arabic text support
- Enhanced error handling and user feedback