# Changelog

All notable changes to MetaTable will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2025-11-17

### Added
- Initial release of MetaTable
- Recursive file scanning capability
- Export file metadata to Excel (.xlsx format)
- Modern WPF user interface with MSU branding
- Real-time progress tracking during scans
- Animated scanning indicator
- Cancellable scan operations
- Double-click to open folder/file locations
- Hover tooltips for all interactive elements
- Select Folder and Output File buttons disabled during scanning
- Read-only clickable path text boxes with visual feedback
- Rounded checkbox corners for modern appearance
- Custom instant tooltips that don't overlap elements

### Metadata Captured
- File name and extension
- Full file path
- File size
- Creation date and time
- Last modified date and time
- Last accessed date and time
- File attributes (Read-only, Hidden, System, Archive)

### Technical Details
- Built with .NET 9.0
- WPF (Windows Presentation Foundation)
- EPPlus for Excel export
- Self-contained deployment option
- Single-file executable support
