
function Invoke-ExcelLoaderCli {
    param(
        [Parameter(Mandatory)][string]$Connection,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$FileGroup,
        [string]$Entity,
        [string]$LoadType,
        [string]$Sheet,
        [switch]$DryRun
    )
    $cli = Join-Path $PSScriptRoot '..\..\dotnet\ExcelLoader.Cli\bin\Release\net8.0\ExcelLoader.Cli.dll'
    if (-not (Test-Path $cli)) {
        Write-Host "Building CLI (Release)..." -ForegroundColor Yellow
        dotnet build (Join-Path $PSScriptRoot '..\..\dotnet\ExcelLoader.sln') -c Release | Out-Null
    }
    $args = @("--connection", $Connection, "--file", $Path, "--fileGroup", $FileGroup)
    if ($Entity)   { $args += @("--entity", $Entity) }
    if ($LoadType) { $args += @("--loadType", $LoadType) }
    if ($Sheet)    { $args += @("--sheet", $Sheet) }
    if ($DryRun)   { $args += @("--dryRun") }
    dotnet $cli @args
}

function Test-FileLoad {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Connection,
        [Parameter(Mandatory)][Alias('File','FullName')][string]$Path,
        [Parameter(Mandatory)][string]$FileGroup,
        [string]$Entity,
        [string]$LoadType,
        [string]$Sheet
    )
    Invoke-ExcelLoaderCli -Connection $Connection -Path $Path -FileGroup $FileGroup -Entity $Entity -LoadType $LoadType -Sheet $Sheet -DryRun
}

function Invoke-FileLoad {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Connection,
        [Parameter(Mandatory)][Alias('File','FullName')][string]$Path,
        [Parameter(Mandatory)][string]$FileGroup,
        [string]$Entity,
        [string]$LoadType,
        [string]$Sheet
    )
    Invoke-ExcelLoaderCli -Connection $Connection -Path $Path -FileGroup $FileGroup -Entity $Entity -LoadType $LoadType -Sheet $Sheet
}
