<#
.SYNOPSIS
    Match client IPs against registered subnets and find unregistered ones.

.PARAMETER ClientCsvPath
    Client list CSV (required column: IPAddress / IP)

.PARAMETER SubnetCsvPath
    Subnet list CSV (required column: Subnet in CIDR notation e.g. 192.168.1.0/24)

.PARAMETER OutputPath
    Output CSV path

.EXAMPLE
    .\Match-SubnetSimple.ps1 -ClientCsvPath ".\clients.csv" -SubnetCsvPath ".\subnets.csv" -OutputPath ".\result.csv"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ClientCsvPath,
    [Parameter(Mandatory)][string]$SubnetCsvPath,
    [Parameter(Mandatory)][string]$OutputPath
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

$results = foreach ($client in $clientCsv) {
    $ip = $client.$ipCol.Trim()
    $matchedSubnet = ''
    $matchedSite   = ''
    $status = 'Unregistered'

    if ($ip -match '^\d+\.\d+\.\d+\.\d+$') {
        foreach ($sn in $subnets) {
            if (Test-IPInSubnet $ip $sn.Subnet) {
                $matchedSubnet = $sn.Subnet
                $matchedSite   = $sn.Site
                $status = 'Matched'
                break
            }
        }
    } else {
        $status = 'InvalidIP'
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

$unmatchedPath = $OutputPath -replace '\.csv$', '_unmatched.csv'
$results | Where-Object { $_.Status -eq 'Unregistered' } |
    Export-Csv -Path $unmatchedPath -NoTypeInformation -Encoding UTF8

# --- Summary ---
$matched   = ($results | Where-Object { $_.Status -eq 'Matched' }).Count
$unmatched = ($results | Where-Object { $_.Status -eq 'Unregistered' }).Count
$invalid   = ($results | Where-Object { $_.Status -eq 'InvalidIP' }).Count

Write-Host ""
Write-Host "===== Result =====" -ForegroundColor Cyan
Write-Host "  Matched      : $matched" -ForegroundColor Green
Write-Host "  Unregistered : $unmatched" -ForegroundColor Yellow
Write-Host "  InvalidIP    : $invalid" -ForegroundColor Red
Write-Host ""
Write-Host "  All          : $OutputPath"
Write-Host "  Unregistered : $unmatchedPath"
