# ExcelLoader UI Add-ons

Adds two optional UIs:
- **ExcelLoader.ExcelAddIn** (VSTO ribbon for Excel desktop on Windows)
- **ExcelLoader.Admin.Wpf** (WPF admin app to manage mappings and run loads)

Both assume the repository DB created from the main Starter Kit.
Set `EXCELLOADER_CLI` to the built CLI DLL path (e.g., `C:\path\to\ExcelLoader.Cli.dll`).

Build prerequisites
- Visual Studio 2022
- .NET 8 SDK
- Office desktop + VSTO tooling (for the add-in)
