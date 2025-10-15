<#
.SYNOPSIS
    RVTools Excel file processor for consolidation and anonymization

.DESCRIPTION
    Processes RVTools Excel exports to consolidate multiple files and/or anonymize sensitive data
    (VM names, hosts, clusters, datacenters, IP addresses)

.PARAMETER Mode
    Operation mode: consolidate, anonymize, both, or deanonymize

.PARAMETER InputFile
    Specific input file to process

.PARAMETER Directory
    Directory to search for RVTools files (default: current directory)

.PARAMETER OutputFile
    Custom output filename (default: auto-generated with timestamp)

.PARAMETER MappingFile
    JSON mapping file for deanonymization

.EXAMPLE
    .\rvtools_processor.ps1 -Mode consolidate
    Consolidates all RVTools files in current directory

.EXAMPLE
    .\rvtools_processor.ps1 -Mode anonymize -InputFile "RVTools_Export.xlsx"
    Anonymizes a single file

.EXAMPLE
    .\rvtools_processor.ps1 -Mode both -Directory "C:\RVTools"
    Consolidates and anonymizes all files in specified directory

.NOTES
    Requires: ImportExcel module (auto-installs if missing)
    Compatible: PowerShell 5.1+ and PowerShell 7+
#>

param(
    [ValidateSet('consolidate','anonymize','both','deanonymize')]
    [string]$Mode = "both",
    [string]$InputFile,
    [string]$Directory = ".",
    [string]$OutputFile,
    [string]$MappingFile
)

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installing ImportExcel module..." -ForegroundColor Yellow
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel
$WarningPreference = 'SilentlyContinue'

function Find-RVToolsFiles {
    param([string]$Directory = ".")
    $files = @()
    Get-ChildItem -Path $Directory -Filter "*.xlsx" | Where-Object { -not $_.Name.StartsWith("~") } | ForEach-Object {
        try {
            $sheets = (Get-ExcelSheetInfo -Path $_.FullName).Name
            if ($sheets -match 'vInfo|vHost|vCluster') { $files += $_ }
        } catch {}
    }
    return $files
}

function Get-OutputFilename {
    param([string]$Mode)
    $timestamp = Get-Date -Format "yyyyMMdd_HHmm"
    switch ($Mode) {
        "both" { "RVTools_Consolidated_Anonymized_$timestamp.xlsx" }
        "consolidate" { "RVTools_Combined_$timestamp.xlsx" }
        "anonymize" { "RVTools_Anonymized_$timestamp.xlsx" }
        "deanonymize" { "RVTools_Deanonymized_$timestamp.xlsx" }
    }
}

function Merge-RVToolsFiles {
    param([array]$InputFiles, [string]$OutputFile)
    Write-Host "`nConsolidating $($InputFiles.Count) files..." -ForegroundColor Cyan
    $consolidatedSheets = @{}
    
    foreach ($file in $InputFiles) {
        Write-Host "Processing: $($file.Name)"
        try {
            $sheets = (Get-ExcelSheetInfo -Path $file.FullName).Name
            foreach ($sheet in $sheets) {
                $data = Import-Excel -Path $file.FullName -WorksheetName $sheet -WarningAction SilentlyContinue
                if ($consolidatedSheets.ContainsKey($sheet)) {
                    $consolidatedSheets[$sheet] += $data
                } else {
                    $consolidatedSheets[$sheet] = @($data)
                }
            }
        } catch {
            Write-Host "Error: $_" -ForegroundColor Red
        }
    }
    
    Write-Host "Creating: $OutputFile" -ForegroundColor Green
    foreach ($sheet in $consolidatedSheets.Keys) {
        $consolidatedSheets[$sheet] | Export-Excel -Path $OutputFile -WorksheetName $sheet -AutoSize
    }
    
    Write-Host "`nSummary: $($InputFiles.Count) files" -ForegroundColor Green
    foreach ($sheet in $consolidatedSheets.Keys) {
        Write-Host " - $sheet`: $($consolidatedSheets[$sheet].Count) rows"
    }
    return $OutputFile
}

function New-AnonymizationManager {
    return @{
        vmCounter = 1
        hostCounter = 1
        clusterCounter = 1
        dcCounter = 1
        ipMappings = @{}
        nameMappings = @{}
        reverseMappings = @{}
        vmIdMappings = @{}
        uniqueVMs = New-Object System.Collections.Generic.HashSet[string]
        uniqueHosts = New-Object System.Collections.Generic.HashSet[string]
        uniqueClusters = New-Object System.Collections.Generic.HashSet[string]
        uniqueDCs = New-Object System.Collections.Generic.HashSet[string]
    }
}

function Get-AnonymizedVMName {
    param($Manager, [string]$vmName, [string]$vmId, [bool]$count = $true)
    if ([string]::IsNullOrWhiteSpace($vmName)) { return $vmName }
    
    if ($Manager.nameMappings.ContainsKey($vmName)) {
        return $Manager.nameMappings[$vmName]
    }
    
    if ($count) { [void]$Manager.uniqueVMs.Add($vmName) }
    
    if (-not [string]::IsNullOrWhiteSpace($vmId)) {
        $anonName = $vmId
        $Manager.nameMappings[$vmName] = $anonName
        $Manager.reverseMappings[$anonName] = $vmName
        $Manager.vmIdMappings[$vmId] = $anonName
        return $anonName
    }
    
    $anonName = "VM-{0:D4}" -f $Manager.vmCounter
    $Manager.nameMappings[$vmName] = $anonName
    $Manager.reverseMappings[$anonName] = $vmName
    $Manager.vmCounter++
    return $anonName
}

function Get-AnonymizedHostName {
    param($Manager, [string]$hostName, [bool]$count = $true)
    if ([string]::IsNullOrWhiteSpace($hostName)) { return $hostName }
    
    if (-not $Manager.nameMappings.ContainsKey($hostName)) {
        if ($count) { [void]$Manager.uniqueHosts.Add($hostName) }
        $anonName = "HOST-{0:D4}" -f $Manager.hostCounter
        $Manager.nameMappings[$hostName] = $anonName
        $Manager.reverseMappings[$anonName] = $hostName
        $Manager.hostCounter++
    }
    return $Manager.nameMappings[$hostName]
}

function Get-AnonymizedClusterName {
    param($Manager, [string]$clusterName)
    if ([string]::IsNullOrWhiteSpace($clusterName)) { return $clusterName }
    
    if (-not $Manager.nameMappings.ContainsKey($clusterName)) {
        [void]$Manager.uniqueClusters.Add($clusterName)
        $anonName = "CLUSTER-{0:D4}" -f $Manager.clusterCounter
        $Manager.nameMappings[$clusterName] = $anonName
        $Manager.reverseMappings[$anonName] = $clusterName
        $Manager.clusterCounter++
    }
    return $Manager.nameMappings[$clusterName]
}

function Get-AnonymizedDCName {
    param($Manager, [string]$dcName)
    if ([string]::IsNullOrWhiteSpace($dcName)) { return $dcName }
    
    if (-not $Manager.nameMappings.ContainsKey($dcName)) {
        [void]$Manager.uniqueDCs.Add($dcName)
        $anonName = "DC-{0:D4}" -f $Manager.dcCounter
        $Manager.nameMappings[$dcName] = $anonName
        $Manager.reverseMappings[$anonName] = $dcName
        $Manager.dcCounter++
    }
    return $Manager.nameMappings[$dcName]
}

function Get-AnonymizedIP {
    param($Manager, [string]$ipStr)
    if ([string]::IsNullOrWhiteSpace($ipStr)) { return $ipStr }
    
    if ($ipStr -match '[,;]') {
        $sep = if ($ipStr -match ',') { ',' } else { ';' }
        $ips = $ipStr -split $sep | ForEach-Object { $_.Trim() }
        $anonIPs = $ips | ForEach-Object { Get-AnonymizedSingleIP -Manager $Manager -ipStr $_ } | Where-Object { $_ }
        return $anonIPs -join $sep
    }
    return Get-AnonymizedSingleIP -Manager $Manager -ipStr $ipStr
}

function Get-AnonymizedSingleIP {
    param($Manager, [string]$ipStr)
    if ([string]::IsNullOrWhiteSpace($ipStr)) { return $ipStr }
    if ($Manager.ipMappings.ContainsKey($ipStr)) { return $Manager.ipMappings[$ipStr] }
    
    try {
        $ip = [System.Net.IPAddress]::Parse($ipStr)
        if ($ip.AddressFamily -eq 'InterNetwork') {
            $bytes = $ip.GetAddressBytes()
            $hash = [System.Security.Cryptography.MD5]::Create().ComputeHash([System.Text.Encoding]::UTF8.GetBytes("$($bytes[0])$($bytes[1])$($bytes[2])"))
            $networkId = ([BitConverter]::ToUInt16($hash, 0) % 254) + 1
            $anonIP = "10.$networkId.$($bytes[2]).$($bytes[3])"
            $Manager.ipMappings[$ipStr] = $anonIP
            return $anonIP
        }
    } catch {}
    return $ipStr
}

function Invoke-Anonymization {
    param([string]$InputFile, [string]$OutputFile)
    Write-Host "`nAnonymizing: $InputFile" -ForegroundColor Cyan
    $mgr = New-AnonymizationManager
    $sheets = (Get-ExcelSheetInfo -Path $InputFile).Name
    
    foreach ($sheet in $sheets) {
        $data = Import-Excel -Path $InputFile -WorksheetName $sheet -WarningAction SilentlyContinue
        foreach ($row in $data) {
            if ($row.VM) {
                $vmId = if ($row.'VM ID') { $row.'VM ID' } else { $null }
                $row.VM = Get-AnonymizedVMName -Manager $mgr -vmName $row.VM -vmId $vmId
            }
            if ($row.Host) { $row.Host = Get-AnonymizedHostName -Manager $mgr -hostName $row.Host }
            if ($row.Cluster) { $row.Cluster = Get-AnonymizedClusterName -Manager $mgr -clusterName $row.Cluster }
            if ($row.Datacenter) { $row.Datacenter = Get-AnonymizedDCName -Manager $mgr -dcName $row.Datacenter }
            if ($row.'Primary IP Address') { $row.'Primary IP Address' = Get-AnonymizedIP -Manager $mgr -ipStr $row.'Primary IP Address' }
        }
        $data | Export-Excel -Path $OutputFile -WorksheetName $sheet -AutoSize
    }
    
    $mappingFile = $OutputFile -replace '\.xlsx$', '_mapping.json'
    @{
        nameMappings = $mgr.nameMappings
        reverseMappings = $mgr.reverseMappings
        ipMappings = $mgr.ipMappings
    } | ConvertTo-Json -Depth 10 | Set-Content -Path $mappingFile
    
    Write-Host "`nComplete!" -ForegroundColor Green
    Write-Host "Output: $OutputFile"
    Write-Host "Mapping: $mappingFile"
}

switch ($Mode) {
    "consolidate" {
        $files = if ($InputFile) { @(Get-Item $InputFile) } else { Find-RVToolsFiles -Directory $Directory }
        if ($files.Count -eq 0) { Write-Host "No RVTools files found" -ForegroundColor Red; exit }
        $output = if ($OutputFile) { $OutputFile } else { Get-OutputFilename -Mode $Mode }
        Merge-RVToolsFiles -InputFiles $files -OutputFile $output
    }
    "anonymize" {
        if (-not $InputFile) { Write-Host "Input file required" -ForegroundColor Red; exit }
        $output = if ($OutputFile) { $OutputFile } else { Get-OutputFilename -Mode $Mode }
        Invoke-Anonymization -InputFile $InputFile -OutputFile $output
    }
    "both" {
        $files = if ($InputFile) { @(Get-Item $InputFile) } else { Find-RVToolsFiles -Directory $Directory }
        if ($files.Count -eq 0) { Write-Host "No RVTools files found" -ForegroundColor Red; exit }
        $temp = "temp_$(Get-Date -Format 'yyyyMMddHHmm').xlsx"
        Merge-RVToolsFiles -InputFiles $files -OutputFile $temp
        $output = if ($OutputFile) { $OutputFile } else { Get-OutputFilename -Mode $Mode }
        Invoke-Anonymization -InputFile $temp -OutputFile $output
        Remove-Item $temp
    }
}
