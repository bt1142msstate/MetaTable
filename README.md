# MetaTable

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET](https://img.shields.io/badge/.NET-9.0-512BD4)](https://dotnet.microsoft.com/)
[![Platform](https://img.shields.io/badge/Platform-Windows-0078D4)](https://www.microsoft.com/windows)
[![C#](https://img.shields.io/badge/C%23-12-239120)](https://docs.microsoft.com/en-us/dotnet/csharp/)
[![WPF](https://img.shields.io/badge/UI-WPF-512BD4)](https://docs.microsoft.com/en-us/dotnet/desktop/wpf/)
[![GitHub issues](https://img.shields.io/github/issues/bt1142msstate/MetaTable)](https://github.com/bt1142msstate/MetaTable/issues)
[![GitHub stars](https://img.shields.io/github/stars/bt1142msstate/MetaTable)](https://github.com/bt1142msstate/MetaTable/stargazers)

A modern WPF desktop application for Windows that scans files in a directory and exports comprehensive metadata to Excel.

![MetaTable Screenshot](screenshots/main-window.png)

> **Note:** The screenshot shows MSU (Mississippi State University) branding. The branding can be easily removed or customized by replacing the icon and color scheme in the source code.

## Features

- 🔍 **Recursive File Scanning** - Scans all files in selected folder and subfolders
- 📊 **Excel Export** - Exports detailed metadata to `.xlsx` format
- 🎨 **Modern UI** - Clean, user-friendly interface with smooth animations
- ⚡ **Real-time Progress** - Live progress tracking during file scanning
- 🛑 **Cancellable Operations** - Stop scanning at any time
- 📁 **Smart File Management** - Automatically opens output location after export

## Metadata Captured

MetaTable extracts the following information for each file:

- File name and extension
- Full file path
- File size
- Creation date and time
- Last modified date and time
- Last accessed date and time
- File attributes (Read-only, Hidden, System, Archive)

## System Requirements

- **OS**: Windows 10 (version 1809 or later) / Windows 11
- **Runtime**: .NET 9.0 Runtime (or use the self-contained executable)
- **Disk Space**: ~100 MB for installation
- **Memory**: 512 MB RAM minimum

## Installation

### Option 1: Self-Contained Executable (Recommended)
1. Download the latest release from the [Releases](https://github.com/bt1142msstate/MetaTable/releases) page
2. Extract the ZIP file to your desired location
3. Run `MetaTable.exe` - no .NET installation required!

### Option 2: Build from Source
See the [Building from Source](#building-from-source) section below.

## Usage

1. **Select Folder** - Click "📁 Select Folder to Scan" to choose the directory
2. **Choose Output** - Click "📄 Select Output Excel File" to specify where to save results
3. **Scan** - Click "🔍 Scan Files and Export Data to Excel" to start scanning
4. **View Results** - The Excel file will open automatically when complete

### Tips

- Double-click the folder path to open the folder in Explorer
- Double-click the output path to open the Excel file location
- Click the scan button during operation to cancel the scan
- Hover over any control to see helpful tooltips

## Building from Source

### Prerequisites
- [.NET 9.0 SDK](https://dotnet.microsoft.com/download/dotnet/9.0)
- [Visual Studio 2022](https://visualstudio.microsoft.com/) (optional, but recommended)

### Build Steps

```powershell
# Clone the repository
git clone https://github.com/bt1142msstate/MetaTable.git
cd MetaTable

# Restore dependencies
dotnet restore

# Build the project
dotnet build -c Release

# Run the application
dotnet run --project MetaTable/MetaTable.csproj
```

### Publishing a Release Build

Create a self-contained, single-file executable for distribution:

```powershell
dotnet publish MetaTable/MetaTable.csproj `
  -c Release `
  -r win-x64 `
  --self-contained true `
  -p:PublishSingleFile=true `
  -p:IncludeNativeLibrariesForSelfExtract=true `
  -o publish
```

The executable will be in the `publish` folder.

## Technology Stack

- **Framework**: [.NET 9.0](https://dotnet.microsoft.com/)
- **UI**: WPF (Windows Presentation Foundation)
- **Excel Export**: [EPPlus](https://www.nuget.org/packages/EPPlus/)
- **Language**: C# 12

## Project Structure

```
MetaTable/
├── MetaTable/              # Main application project
│   ├── MainWindow.xaml     # UI layout
│   ├── MainWindow.xaml.cs  # Application logic
│   ├── App.xaml            # Application resources
│   ├── Logo.ico            # Application icon
│   └── MetaTable.csproj    # Project file
├── screenshots/            # Application screenshots
├── Standalone/             # Batch launchers
├── LICENSE                 # MIT License
├── README.md               # This file
├── CONTRIBUTING.md         # Contribution guidelines
└── CHANGELOG.md            # Version history
```

## Customization

### Changing the Branding

To customize the MSU branding:

1. **Icon**: Replace `MetaTable/Logo.ico` with your own icon
2. **Colors**: Edit the color values in `MainWindow.xaml`:
   - Primary color: `#660000` (maroon)
   - Accent color: `#4A9EFF` (blue)
   - Background: `#1E1E1E` (dark gray)
3. **Window Title**: Change `Title` property in `MainWindow.xaml`

## Troubleshooting

### Application won't start
- Ensure you have .NET 9.0 Runtime installed (or use the self-contained version)
- Try running as administrator
- Check Windows Event Viewer for error details

### Excel file won't open
- Ensure you have Microsoft Excel or a compatible spreadsheet application installed
- Check that the output path is writable
- Verify the file isn't open in another application

### Scan is very slow
- Large directories with many files will take time
- Network drives are slower than local drives
- Consider scanning smaller subdirectories

## Frequently Asked Questions

**Q: Can I scan network drives?**  
A: Yes, but performance may be slower depending on network speed.

**Q: What's the maximum number of files it can handle?**  
A: There's no hard limit, but very large scans (100,000+ files) may take significant time and memory.

**Q: Can I customize what metadata is captured?**  
A: Yes! Modify the scanning logic in `MainWindow.xaml.cs` to add or remove metadata fields.

**Q: Does it work with Excel Online or Google Sheets?**  
A: The output is a standard .xlsx file that can be opened in Excel Online, Google Sheets, LibreOffice, etc.

## License

This project is open source and available under the [MIT License](LICENSE).

## Author

Created by Brandon Temple

## Contributing

Contributions, issues, and feature requests are welcome! See [CONTRIBUTING.md](CONTRIBUTING.md) for details.

Feel free to check the [issues page](https://github.com/bt1142msstate/MetaTable/issues) for open issues or to submit new ones.

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for a detailed version history.

## Acknowledgments

- Built with [WPF](https://docs.microsoft.com/en-us/dotnet/desktop/wpf/)
- Excel export powered by [EPPlus](https://www.eppluspoftware.com/)
- Inspired by the need for simple, efficient file metadata analysis

## Support

If you find this project helpful, please consider:
- ⭐ Starring the repository
- 🐛 Reporting bugs
- 💡 Suggesting new features
- 🔀 Contributing code

---

**Made with ❤️ for the file management community**

