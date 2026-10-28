param([string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$projectDirectory = $PSScriptRoot
$sourceDirectory = Join-Path $projectDirectory 'src'
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $projectDirectory 'dist'
}
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $compiler)) {
    throw '.NET Framework C# compiler not found.'
}
$outputFile = Join-Path $OutputDirectory 'MathTypeTextEditor.exe'
$sources = @(
    (Join-Path $sourceDirectory 'MathTypeEditor.cs'),
    (Join-Path $sourceDirectory 'FormulaText.cs'),
    (Join-Path $sourceDirectory 'WordBridge.cs'),
    (Join-Path $sourceDirectory 'EditorTests.cs')
)
& $compiler /nologo /target:winexe /platform:x86 /optimize+ /codepage:65001 /utf8output /reference:System.dll /reference:System.Core.dll /reference:System.Drawing.dll /reference:System.Windows.Forms.dll /reference:System.Xml.dll /reference:System.Xml.Linq.dll /reference:Microsoft.CSharp.dll ('/out:' + $outputFile) @sources
if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
Copy-Item -LiteralPath (Join-Path $projectDirectory 'docs\使用说明.txt') -Destination (Join-Path $OutputDirectory '使用说明.txt') -Force
Write-Output $outputFile
