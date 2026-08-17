[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$temporaryBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$temporaryRoot = Join-Path $temporaryBase ("dr9-db-build-" + [Guid]::NewGuid().ToString("N"))
$temporaryScripts = Join-Path $temporaryRoot "scripts"
$temporaryData = Join-Path $temporaryRoot "App_Data"
$cscriptPath = Join-Path $env:WINDIR "SysWOW64\cscript.exe"

New-Item -ItemType Directory -Path $temporaryScripts, $temporaryData | Out-Null
Copy-Item -LiteralPath (Join-Path $PSScriptRoot "create_database.vbs") -Destination $temporaryScripts
Copy-Item -LiteralPath (Join-Path $PSScriptRoot "smoke_test.vbs") -Destination $temporaryScripts

try {
    & $cscriptPath //nologo (Join-Path $temporaryScripts "create_database.vbs")
    if ($LASTEXITCODE -ne 0) {
        throw "create_database.vbs failed with exit code $LASTEXITCODE."
    }

    & $cscriptPath //nologo (Join-Path $temporaryScripts "smoke_test.vbs")
    if ($LASTEXITCODE -ne 0) {
        throw "smoke_test.vbs failed with exit code $LASTEXITCODE."
    }

    $databaseFile = Get-Item -LiteralPath (Join-Path $temporaryData "tasks.mdb")
    Write-Host ("PASS: clean database build (" + $databaseFile.Length + " bytes)")
} finally {
    $resolvedTemporaryRoot = [IO.Path]::GetFullPath($temporaryRoot)
    $safeToClean = $resolvedTemporaryRoot.StartsWith($temporaryBase, [StringComparison]::OrdinalIgnoreCase) -and
        ([IO.Path]::GetFileName($resolvedTemporaryRoot) -like "dr9-db-build-*")

    if ($safeToClean) {
        $temporaryFiles = @(
            (Join-Path $temporaryData "tasks.ldb"),
            (Join-Path $temporaryData "tasks.mdb"),
            (Join-Path $temporaryData "tasks.building.mdb"),
            (Join-Path $temporaryScripts "create_database.vbs"),
            (Join-Path $temporaryScripts "smoke_test.vbs")
        )
        foreach ($temporaryFile in $temporaryFiles) {
            if (Test-Path -LiteralPath $temporaryFile) {
                Remove-Item -LiteralPath $temporaryFile -Force
            }
        }

        foreach ($temporaryDirectory in @($temporaryData, $temporaryScripts, $resolvedTemporaryRoot)) {
            if (Test-Path -LiteralPath $temporaryDirectory) {
                Remove-Item -LiteralPath $temporaryDirectory -Force
            }
        }
    }
}
