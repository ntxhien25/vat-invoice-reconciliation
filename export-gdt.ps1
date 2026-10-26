# Public-package launcher. The existing Python/Playwright code performs the export.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][ValidatePattern('^\d{4}-\d{2}$')][string]$Month,
    [string]$TaxId,
    [ValidateSet('all','purchases','sales')][string]$Mode = 'all',
    [string]$Python = 'python',
    [switch]$Plan
)
$vatArguments = @((Join-Path $PSScriptRoot 'gdt_export.py'), '--month', $Month, '--mode', $Mode)
if ($Plan) {
    $vatArguments += '--plan'
} else {
    if ([string]::IsNullOrWhiteSpace($TaxId)) { throw 'Supply -TaxId, or use -Plan.' }
    $vatArguments += @('--tax-id', $TaxId, '--output', (Join-Path $PSScriptRoot 'exports'))
}
& $Python @vatArguments
if ($LASTEXITCODE -ne 0) { throw "Exporter stopped with exit code $LASTEXITCODE. Review its message and run report." }
