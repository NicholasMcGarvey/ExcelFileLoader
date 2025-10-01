
@{
    RootModule        = 'ExcelLoader.psm1'
    ModuleVersion     = '0.1.0'
    GUID              = 'c6a9d2f1-6d2e-4c0c-9a12-8c0e3b9b1c11'
    Author            = 'ExcelLoader'
    CompanyName       = 'ExcelLoader'
    CompatiblePSEditions = @('Desktop','Core')
    PowerShellVersion = '5.1'
    Description       = 'PowerShell wrapper for ExcelLoader CLI'
    FunctionsToExport = @('Test-FileLoad','Invoke-FileLoad')
    PrivateData       = @{
        PSData = @{
            Tags = @('excel','etl','loader')
        }
    }
}
