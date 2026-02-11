<#
.SYNOPSIS
    Check if client IPs in a CSV belong to registered subnets.

.PARAMETER ClientCsvPath
    Client list CSV (required column: IPAddress / IP)

.PARAMETER SubnetCsvPath
    Subnet list CSV (required column: Subnet in CIDR notation e.g. 192.168.1.0/24)

.PARAMETER OutputPath
    Output CSV path (optional, default: result.csv in same folder as ClientCsvPath)

.EXAMPLE
    .\Check-SubnetBatch.ps1 -ClientCsvPath ".\new_clients.csv" -SubnetCsvPath ".\subnets.csv"

.EXAMPLE
    .\Check-SubnetBatch.ps1 -ClientCsvPath ".\new_clients.csv" -SubnetCsvPath ".\subnets.csv" -OutputPath ".\check_result.csv"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ClientCsvPath,
    [Parameter(Mandatory)][string]$SubnetCsvPath,
    [string]$OutputPath
)

# --- Helper Functions ---
function ConvertTo-BinaryIP ([string]$IP) {
    $o = $IP.Split('.')
    return ([uint32]$o[0] -shl 24) + ([uint32]$o[1] -shl 16) + ([uint32]$o[2] -shl 8) + [uint32]$o[3]
}

function Test-IPInSubnet ([string]$IP, [string]$Subnet) {
    $parts = $Subnet.Split('/')
    $prefix = [int]$parts[1]
    $mask = if ($prefix -eq 0) { [uint32]0 } else { [uint32]([math]::Pow(2,32) - [math]::Pow(2,32-$prefix)) }
    return ((ConvertTo-BinaryIP $IP) -band $mask) -eq ((ConvertTo-BinaryIP $parts[0]) -band $mask)
}

function Find-Column ($Csv, [string[]]$Patterns) {
    foreach ($p in $Patterns) {
        $col = $Csv[0].PSObject.Properties.Name | Where-Object { $_ -match $p } | Select-Object -First 1
        if ($col) { return $col }
    }
    return $null
}

# --- Default OutputPath ---
if (-not $OutputPath) {
    $dir = Split-Path $ClientCsvPath -Parent
    if (-not $dir) { $dir = '.' }
    $OutputPath = Join-Path $dir 'check_result.csv'
}

# --- Load CSVs ---
$clientCsv = Import-Csv $ClientCsvPath -Encoding UTF8
$subnetCsv = Import-Csv $SubnetCsvPath -Encoding UTF8

$ipCol = Find-Column $clientCsv @('IPAddress', 'IP', 'address')
$snCol = Find-Column $subnetCsv @('subnet', 'network', 'cidr', 'prefix')

if (-not $ipCol) { Write-Error "IP column not found in client CSV. Columns: $($clientCsv[0].PSObject.Properties.Name -join ', ')"; exit 1 }
if (-not $snCol) { Write-Error "Subnet column not found in subnet CSV. Columns: $($subnetCsv[0].PSObject.Properties.Name -join ', ')"; exit 1 }

$siteCol = Find-Column $subnetCsv @('site')

$subnets = $subnetCsv | ForEach-Object {
    [PSCustomObject]@{
        Subnet = $_.$snCol.Trim()
        Site   = if ($siteCol) { $_.$siteCol.Trim() } else { '' }
    }
} | Where-Object { $_.Subnet -match '^\d+\.\d+\.\d+\.\d+/\d+$' }

Write-Host "Subnets: $($subnets.Count) / Clients: $($clientCsv.Count)" -ForegroundColor Cyan

# --- Match ---
$clientProps = $clientCsv[0].PSObject.Properties.Name
$okCount = 0
$ngCount = 0
$errCount = 0

$results = foreach ($client in $clientCsv) {
    $ip = $client.$ipCol.Trim()
    $matchedSubnet = ''
    $matchedSite   = ''
    $status = 'NG'

    if ($ip -match '^\d+\.\d+\.\d+\.\d+$') {
        foreach ($sn in $subnets) {
            if (Test-IPInSubnet $ip $sn.Subnet) {
                $matchedSubnet = $sn.Subnet
                $matchedSite   = $sn.Site
                $status = 'OK'
                break
            }
        }
    } else {
        $status = 'InvalidIP'
    }

    switch ($status) {
        'OK'        { $okCount++ }
        'NG'        { $ngCount++ }
        'InvalidIP' { $errCount++ }
    }

    $obj = [ordered]@{}
    foreach ($p in $clientProps) { $obj[$p] = $client.$p }
    $obj['MatchedSubnet'] = $matchedSubnet
    $obj['MatchedSite']   = $matchedSite
    $obj['Status']        = $status
    [PSCustomObject]$obj
}

# --- Output ---
$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

$ngPath = $OutputPath -replace '\.csv$', '_NG.csv'
$results | Where-Object { $_.Status -ne 'OK' } |
    Export-Csv -Path $ngPath -NoTypeInformation -Encoding UTF8

# --- Console Output ---
Write-Host ""
Write-Host "===== Result =====" -ForegroundColor Cyan

foreach ($r in $results) {
    $color = switch ($r.Status) { 'OK' { 'Green' } 'NG' { 'Yellow' } default { 'Red' } }
    $detail = if ($r.Status -eq 'OK') { "$($r.MatchedSubnet) ($($r.MatchedSite))" } else { $r.Status }
    Write-Host ("  {0,-15} {1,-4} {2}" -f $r.$ipCol, $r.Status, $detail) -ForegroundColor $color
}

Write-Host ""
Write-Host "  OK: $okCount / NG: $ngCount / InvalidIP: $errCount" -ForegroundColor Cyan
Write-Host ""
Write-Host "  All    : $OutputPath"
Write-Host "  NG only: $ngPath"
