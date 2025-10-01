
# Example: register collateral schema and mapping (already covered by SampleSeed.sql)
# This script simply shows how to run the CLI in dry run and actual mode.

$connection = "Server=localhost;Database=ExcelLoader;Trusted_Connection=True;TrustServerCertificate=True;"
$file = "C:\Data\Collateral\LenderA_Daily_2025-09-29.xlsx"

# Dry run
dotnet run --project ..\..\src\dotnet\ExcelLoader.Cli -- `
  --connection "$connection" `
  --file "$file" --fileGroup COL --entity LenderA --loadType daily --dryRun

# Execute load
dotnet run --project ..\..\src\dotnet\ExcelLoader.Cli -- `
  --connection "$connection" `
  --file "$file" --fileGroup COL --entity LenderA --loadType daily
