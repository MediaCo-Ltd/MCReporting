#Requires -Version 5.1
<#
.SYNOPSIS
    Validates dependencies for MCReporting application migration.
.DESCRIPTION
    Reads validate_dependencies.json and performs all specified checks.
.EXAMPLE
    .\validate_dependencies.ps1
#>

$ErrorActionPreference = "Continue"
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$jsonPath = Join-Path $scriptPath "validate_dependencies.json"

if (-not (Test-Path $jsonPath)) {
    Write-Host "ERROR: validate_dependencies.json not found at $jsonPath" -ForegroundColor Red
    exit 1
}

$config = Get-Content $jsonPath | ConvertFrom-Json
$results = @()
$passed = 0
$failed = 0

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  $($config.application) Dependency Validation" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Network checks
Write-Host "Network Checks:" -ForegroundColor Yellow
foreach ($check in $config.checks.network) {
    $result = @{ Check = $check.description; Status = "UNKNOWN" }

    if ($check.type -eq "ping") {
        try {
            $ping = Test-Connection -ComputerName $check.target -Count 1 -Quiet -ErrorAction SilentlyContinue
            if ($ping) {
                $result.Status = "PASS"
                $passed++
                Write-Host "  [PASS] $($check.description) - $($check.target)" -ForegroundColor Green
            } else {
                $result.Status = "FAIL"
                $failed++
                Write-Host "  [FAIL] $($check.description) - $($check.target) not reachable" -ForegroundColor Red
            }
        } catch {
            $result.Status = "FAIL"
            $failed++
            Write-Host "  [FAIL] $($check.description) - Error: $_" -ForegroundColor Red
        }
    }
    elseif ($check.type -eq "port") {
        try {
            $tcp = New-Object System.Net.Sockets.TcpClient
            $connect = $tcp.BeginConnect($check.target, $check.port, $null, $null)
            $wait = $connect.AsyncWaitHandle.WaitOne(3000, $false)
            if ($wait -and $tcp.Connected) {
                $result.Status = "PASS"
                $passed++
                Write-Host "  [PASS] $($check.description) - $($check.target):$($check.port)" -ForegroundColor Green
            } else {
                $result.Status = "FAIL"
                $failed++
                Write-Host "  [FAIL] $($check.description) - $($check.target):$($check.port) not accessible" -ForegroundColor Red
            }
            $tcp.Close()
        } catch {
            $result.Status = "FAIL"
            $failed++
            Write-Host "  [FAIL] $($check.description) - Error: $_" -ForegroundColor Red
        }
    }
    $results += $result
}

# COM Object checks
Write-Host "`nCOM Object Checks:" -ForegroundColor Yellow
foreach ($check in $config.checks.com_objects) {
    $result = @{ Check = $check.description; Status = "UNKNOWN" }
    try {
        $obj = New-Object -ComObject $check.progid -ErrorAction Stop
        if ($null -ne $obj) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($obj) | Out-Null
        }
        $result.Status = "PASS"
        $passed++
        Write-Host "  [PASS] $($check.description) - $($check.progid)" -ForegroundColor Green
    } catch {
        $result.Status = "FAIL"
        $failed++
        Write-Host "  [FAIL] $($check.description) - $($check.progid) not registered" -ForegroundColor Red
    }
    $results += $result
}

# Folder checks
Write-Host "`nFolder Checks:" -ForegroundColor Yellow
foreach ($check in $config.checks.folders) {
    $result = @{ Check = $check.description; Status = "UNKNOWN" }
    if (Test-Path $check.path) {
        $result.Status = "PASS"
        $passed++
        Write-Host "  [PASS] $($check.description) - $($check.path)" -ForegroundColor Green
    } else {
        $result.Status = "FAIL"
        $failed++
        Write-Host "  [FAIL] $($check.description) - $($check.path) not found" -ForegroundColor Red
    }
    $results += $result
}

# IIS checks
Write-Host "`nIIS Checks:" -ForegroundColor Yellow
$iisAvailable = $false
try {
    Import-Module WebAdministration -ErrorAction Stop
    $iisAvailable = $true
} catch {
    Write-Host "  [WARN] WebAdministration module not available - skipping IIS checks" -ForegroundColor Yellow
}

if ($iisAvailable) {
    foreach ($check in $config.checks.iis) {
        $result = @{ Check = $check.description; Status = "UNKNOWN" }
        try {
            if ($check.type -eq "application_pool") {
                $pool = Get-IISAppPool -Name $check.name -ErrorAction SilentlyContinue
                if ($pool) {
                    $result.Status = "PASS"
                    $passed++
                    Write-Host "  [PASS] $($check.description)" -ForegroundColor Green
                } else {
                    $result.Status = "FAIL"
                    $failed++
                    Write-Host "  [FAIL] $($check.description) - App pool '$($check.name)' not found" -ForegroundColor Red
                }
            }
            elseif ($check.type -eq "website") {
                $site = Get-IISSite -Name $check.name -ErrorAction SilentlyContinue
                if ($site) {
                    $result.Status = "PASS"
                    $passed++
                    Write-Host "  [PASS] $($check.description)" -ForegroundColor Green
                } else {
                    $result.Status = "FAIL"
                    $failed++
                    Write-Host "  [FAIL] $($check.description) - Website '$($check.name)' not found" -ForegroundColor Red
                }
            }
        } catch {
            $result.Status = "FAIL"
            $failed++
            Write-Host "  [FAIL] $($check.description) - Error: $_" -ForegroundColor Red
        }
        $results += $result
    }
}

# SMTP checks
Write-Host "`nSMTP Checks:" -ForegroundColor Yellow
foreach ($check in $config.checks.smtp) {
    $result = @{ Check = $check.description; Status = "UNKNOWN" }
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $connect = $tcp.BeginConnect($check.host, 25, $null, $null)
        $wait = $connect.AsyncWaitHandle.WaitOne(5000, $false)
        if ($wait -and $tcp.Connected) {
            $result.Status = "PASS"
            $passed++
            Write-Host "  [PASS] $($check.description) - $($check.host):25" -ForegroundColor Green
        } else {
            $result.Status = "FAIL"
            $failed++
            Write-Host "  [FAIL] $($check.description) - $($check.host):25 not accessible" -ForegroundColor Red
        }
        $tcp.Close()
    } catch {
        $result.Status = "FAIL"
        $failed++
        Write-Host "  [FAIL] $($check.description) - Error: $_" -ForegroundColor Red
    }
    $results += $result
}

# Summary
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  Summary" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Passed: $passed" -ForegroundColor Green
Write-Host "  Failed: $failed" -ForegroundColor $(if ($failed -gt 0) { "Red" } else { "Green" })
Write-Host "  Total:  $($passed + $failed)" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

if ($failed -gt 0) {
    exit 1
} else {
    exit 0
}
