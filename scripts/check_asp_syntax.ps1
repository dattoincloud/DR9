[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$temporaryBase = [IO.Path]::GetTempPath()
$temporaryName = "dr9-asp-check-" + [Guid]::NewGuid().ToString("N")
$temporaryRoot = Join-Path $temporaryBase $temporaryName
$cscriptPath = Join-Path $env:WINDIR "System32\cscript.exe"

function Get-AspBlocks {
    param([string]$Text)

    return [regex]::Matches($Text, "<%(?!@)(=)?([\s\S]*?)%>")
}

function Get-IncludeScript {
    param([string]$PageText)

    $result = New-Object System.Text.StringBuilder
    $includePattern = '<!--\s*#include\s+file="([^"]+)"\s*-->'

    foreach ($includeMatch in [regex]::Matches($PageText, $includePattern, "IgnoreCase")) {
        $relativePath = $includeMatch.Groups[1].Value.Replace("/", [IO.Path]::DirectorySeparatorChar)
        $includePath = Join-Path $projectRoot $relativePath

        if (-not (Test-Path -LiteralPath $includePath)) {
            throw "Missing include: $relativePath"
        }

        $includeText = Get-Content -LiteralPath $includePath -Raw -Encoding UTF8
        foreach ($block in Get-AspBlocks -Text $includeText) {
            [void]$result.AppendLine($block.Groups[2].Value)
        }
    }

    return $result.ToString()
}

function Get-PageScript {
    param([string]$PageText)

    $result = New-Object System.Text.StringBuilder

    foreach ($block in Get-AspBlocks -Text $PageText) {
        $isExpression = $block.Groups[1].Success
        $scriptText = $block.Groups[2].Value
        $scriptText = [regex]::Replace($scriptText, '(?im)^\s*Option\s+Explicit\s*$', '')

        if ($isExpression) {
            [void]$result.AppendLine("Call AspExpressionCheck(" + $scriptText.Trim() + ")")
        } else {
            [void]$result.AppendLine($scriptText)
        }
    }

    return $result.ToString()
}

New-Item -ItemType Directory -Path $temporaryRoot | Out-Null

try {
    $failures = 0
    $aspFiles = Get-ChildItem -LiteralPath $projectRoot -Filter "*.asp" -File | Sort-Object Name

    foreach ($aspFile in $aspFiles) {
        $pageText = Get-Content -LiteralPath $aspFile.FullName -Raw -Encoding UTF8
        $combined = New-Object System.Text.StringBuilder
        [void]$combined.AppendLine("Option Explicit")
        [void]$combined.AppendLine("Dim Server, Request, Response, Session")
        [void]$combined.AppendLine("Sub AspExpressionCheck(ByVal value)")
        [void]$combined.AppendLine("End Sub")
        [void]$combined.AppendLine((Get-IncludeScript -PageText $pageText))
        [void]$combined.AppendLine("WScript.Quit 0")
        [void]$combined.AppendLine((Get-PageScript -PageText $pageText))

        $checkPath = Join-Path $temporaryRoot ($aspFile.BaseName + ".vbs")
        [IO.File]::WriteAllText($checkPath, $combined.ToString(), [Text.Encoding]::Unicode)

        $savedErrorActionPreference = $ErrorActionPreference
        $ErrorActionPreference = "Continue"
        $output = & $cscriptPath //nologo $checkPath 2>&1
        $checkExitCode = $LASTEXITCODE
        $ErrorActionPreference = $savedErrorActionPreference

        if ($checkExitCode -eq 0) {
            Write-Host ("PASS: " + $aspFile.Name)
        } else {
            $failures += 1
            Write-Host ("FAIL: " + $aspFile.Name)
            $output | ForEach-Object { Write-Host $_ }
        }
    }

    if ($failures -gt 0) {
        throw "$failures ASP page(s) failed VBScript compilation."
    }
} finally {
    $resolvedTemporaryBase = [IO.Path]::GetFullPath($temporaryBase)
    $resolvedTemporaryRoot = [IO.Path]::GetFullPath($temporaryRoot)
    $isExpectedTemp = $resolvedTemporaryRoot.StartsWith($resolvedTemporaryBase, [StringComparison]::OrdinalIgnoreCase) -and
        ([IO.Path]::GetFileName($resolvedTemporaryRoot) -like "dr9-asp-check-*")

    if ($isExpectedTemp -and (Test-Path -LiteralPath $resolvedTemporaryRoot)) {
        Remove-Item -LiteralPath $resolvedTemporaryRoot -Recurse -Force
    }
}
