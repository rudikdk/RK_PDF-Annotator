param(
    [Parameter(Mandatory = $false)]
    [string]$Version = "1.0.0"
)

$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$addinRoot = Split-Path -Parent $scriptRoot
$repoRoot = Split-Path -Parent $addinRoot
$project = Join-Path $addinRoot "RKPdfAnnotator.csproj"
$publish = Join-Path $addinRoot "bin\x64\Release\net48\publish"
$dist = Join-Path $addinRoot "dist"
$packageRoot = Join-Path $dist "RKPdfAnnotator-v$Version"
$zipPath = Join-Path $dist "RKPdfAnnotator-v$Version.zip"

dotnet build $project -c Release /p:Platform=x64

if (Test-Path $packageRoot) {
    Remove-Item -LiteralPath $packageRoot -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $packageRoot | Out-Null

Copy-Item -LiteralPath (Join-Path $publish "RKPdfAnnotator-packed.xll") -Destination $packageRoot
Copy-Item -LiteralPath (Join-Path $publish "RKPdfAnnotator64-packed.xll") -Destination $packageRoot
Copy-Item -LiteralPath (Join-Path $addinRoot "docs\USER_GUIDE.html") -Destination $packageRoot
Copy-Item -LiteralPath (Join-Path $addinRoot "README.md") -Destination $packageRoot
Copy-Item -LiteralPath (Join-Path $addinRoot "CHANGELOG.md") -Destination $packageRoot
Copy-Item -LiteralPath (Join-Path $addinRoot "THIRD_PARTY_NOTICES.md") -Destination $packageRoot

if (Test-Path (Join-Path $repoRoot "LICENSE")) {
    Copy-Item -LiteralPath (Join-Path $repoRoot "LICENSE") -Destination $packageRoot
}

$checksums = Get-ChildItem -LiteralPath $packageRoot -File |
    Sort-Object Name |
    ForEach-Object {
        $hash = Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256
        "$($hash.Hash)  $($_.Name)"
    }
$checksums | Set-Content -LiteralPath (Join-Path $packageRoot "CHECKSUMS.txt") -Encoding ASCII

if (Test-Path $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}
Compress-Archive -Path (Join-Path $packageRoot "*") -DestinationPath $zipPath

Write-Host "Created $zipPath"
