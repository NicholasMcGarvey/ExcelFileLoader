# ExcelLoader Starter Kit

A reusable, data-driven pipeline to load Excel files into database schemas with per-FileGroup / per-Entity / per-LoadType mappings.

## Contents
- **database/**: Repository schema (mappings, overrides, run audit) and sample target DB objects
- **src/dotnet/**: .NET 8 projects (Core library, Excel parser, CLI)
- **src/powershell/ExcelLoader/**: PowerShell wrapper module
- **config/**: Example configuration
- **scripts/examples/**: Example registration and load scripts

## Quickstart (SQL Server + .NET 8)
1. Create the repository DB objects:
   ```sql
   :r .\database\RepositorySchema.sql
   :r .\database\SampleTargetSchema.sql
   :r .\database\SampleSeed.sql
   ```

2. Restore NuGet packages and build:
   ```bash
   dotnet build .\src\dotnet\ExcelLoader.sln -c Release
   ```

3. Dry run a sample file:
   ```bash
   dotnet run --project .\src\dotnet\ExcelLoader.Cli --      --connection "Server=localhost;Database=ExcelLoader;Trusted_Connection=True;TrustServerCertificate=True;"      --file "C:\Data\Collateral\LenderA_Daily_2025-09-29.xlsx"      --fileGroup COL --entity LenderA --loadType daily --dryRun
   ```

4. Execute a load:
   ```bash
   dotnet run --project .\src\dotnet\ExcelLoader.Cli --      --connection "Server=localhost;Database=ExcelLoader;Trusted_Connection=True;TrustServerCertificate=True;"      --file "C:\Data\Collateral\LenderA_Daily_2025-09-29.xlsx"      --fileGroup COL --entity LenderA --loadType daily
   ```

5. PowerShell module (optional UX):
   ```powershell
   Import-Module .\src\powershell\ExcelLoader\ExcelLoader.psd1 -Force
   Test-FileLoad `
     -Connection "Server=localhost;Database=ExcelLoader;Trusted_Connection=True;TrustServerCertificate=True;" `
     -Path "C:\Data\Collateral\LenderA_Daily_2025-09-29.xlsx" `
     -FileGroup COL -Entity LenderA -LoadType daily

   Invoke-FileLoad `
     -Connection "Server=localhost;Database=ExcelLoader;Trusted_Connection=True;TrustServerCertificate=True;" `
     -Path "C:\Data\Collateral\LenderA_Daily_2025-09-29.xlsx" `
     -FileGroup COL -Entity LenderA -LoadType daily
   ```

## Notes
- Starter kit supports multiple target tables per file via data-driven mappings.
- Transforms implemented: `trim`, `upper`, `lower`, `strip_nonnum`, `int`, `decimal(p,s)`, `bool`, `date(pattern)`, `money`, `coalesce(value)`.
- Excel parsing via ClosedXML (first matching or specified sheet). CSV is also supported by extension method.
- Uses `Microsoft.Data.SqlClient` and `Dapper` for DB operations.
- See comments marked **TODO** to extend validations (FK checks, rule engine) and to broaden transform set.
