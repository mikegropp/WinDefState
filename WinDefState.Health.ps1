#Requires -Version 5.1
<#!
.SYNOPSIS
Read-only Windows protection health inventory, independent of the state engine.
.DESCRIPTION
Returns a report object. Optional JSON or self-contained HTML export refuses to
overwrite an existing file. Does not elevate, change settings, or contact a service.
#>
[CmdletBinding()]
param(
    [Alias('OutputPath')][string]$HealthOutputPath,
    [Alias('Format')][ValidateSet('Json', 'Html')] [string]$HealthFormat = 'Json'
)

function Get-HealthProperty {
    param($Object, [string]$Name)
    if ($null -ne $Object -and $null -ne $Object.PSObject.Properties[$Name]) {
        $Object.$Name
    }
}

function Invoke-HealthProbe {
    param([scriptblock]$Read)
    $timer = [Diagnostics.Stopwatch]::StartNew()
    try {
        $value = & $Read
        if ($null -eq $value) { throw 'The provider returned no data.' }
        [pscustomobject]@{ Success = $true; Value = $value; Error = ''; DurationMs = $timer.ElapsedMilliseconds }
    } catch {
        [pscustomobject]@{ Success = $false; Value = $null; Error = $_.Exception.Message; DurationMs = $timer.ElapsedMilliseconds }
    }
}

function Get-WindowsHealthIdentity {
    param($OperatingSystem, $CurrentVersion, [datetime]$AsOf = (Get-Date))
    # ProductName can still say Windows 10 on Windows 11. Build + ProductType
    # identify the client family; Server 2025 shares a Windows 11 build number.
    $build = 0
    $null = [int]::TryParse([string](Get-HealthProperty $OperatingSystem 'BuildNumber'), [ref]$build)
    $productType = Get-HealthProperty $OperatingSystem 'ProductType'
    $edition = [string](Get-HealthProperty $CurrentVersion 'EditionID')
    $displayVersion = [string](Get-HealthProperty $CurrentVersion 'DisplayVersion')
    $family = 'Unknown Windows'
    if ($null -ne $productType -and [int]$productType -ne 1) {
        $family = 'Windows Server'
    } elseif ($null -ne $productType -and $build -ge 22000) {
        $family = 'Windows 11'
    } elseif ($null -ne $productType -and $build -ge 10240) {
        $family = 'Windows 10'
    }
    $release = $displayVersion
    $endDate = $null
    $note = 'Lifecycle not mapped. Check the exact edition and release with Microsoft.'
    $source = 'https://learn.microsoft.com/en-us/windows/release-health/windows11-release-information'
    $consumerEditions = @('Core', 'CoreN', 'CoreSingleLanguage', 'CoreCountrySpecific', 'Professional', 'ProfessionalN', 'ProfessionalEducation', 'ProfessionalEducationN', 'ProfessionalWorkstation', 'ProfessionalWorkstationN')
    $enterpriseEditions = @('Enterprise', 'EnterpriseN', 'EnterpriseMultiSession', 'ServerRdsh', 'Education', 'EducationN', 'IoTEnterprise')
    # Release lifecycle metadata reviewed 2026-09-23; deliberately exact builds.
    $releases = @{
        22000 = @('21H2', '2023-10-10', '2024-10-08')
        22621 = @('22H2', '2024-10-08', '2025-10-14')
        22631 = @('23H2', '2025-11-11', '2026-11-10')
        26100 = @('24H2', '2026-10-13', '2027-10-12')
        26200 = @('25H2', '2027-10-12', '2028-10-10')
        28000 = @('26H1', '2028-03-14', '2029-03-13')
    }
    $isLtsc = $edition -in @('EnterpriseS', 'EnterpriseSN', 'IoTEnterpriseS', 'IoTEnterpriseSK')
    if ($family -eq 'Windows 11' -and $isLtsc -and $build -eq 26100) {
        $release = 'LTSC 2024'
        $endDate = if ($edition -like 'IoT*') { '2034-10-10' } else { '2029-10-09' }
    } elseif ($family -eq 'Windows 11' -and -not $isLtsc -and $releases.ContainsKey($build)) {
        $release = $releases[$build][0]
        if ($edition -in $consumerEditions) { $endDate = $releases[$build][1] }
        elseif ($edition -in $enterpriseEditions) { $endDate = $releases[$build][2] }
    } elseif ($family -eq 'Windows 10') {
        $source = 'https://learn.microsoft.com/en-us/windows/release-health/release-information'
        if ($isLtsc) {
            switch ($build) {
                10240 { $release = 'LTSB 2015'; $endDate = '2025-10-14' }
                14393 { $release = 'LTSB 2016'; $endDate = '2026-10-13' }
                17763 { $release = 'LTSC 2019'; $endDate = '2029-01-09' }
                19044 {
                    $release = 'LTSC 2021'
                    $endDate = if ($edition -like 'IoT*') { '2032-01-13' } else { '2027-01-12' }
                }
            }
        } elseif ($build -eq 19045 -and ($edition -in $consumerEditions -or $edition -in $enterpriseEditions)) {
            $release = '22H2'
            $endDate = '2025-10-14'
        }
        $note = 'ESU enrollment and entitlement are not inferred. Verify update coverage separately.'
    }
    $lifecycle = 'Unknown'
    if ($null -ne $endDate) {
        $lastDay = [datetime]::ParseExact($endDate, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
        $lifecycle = if ($AsOf.Date -gt $lastDay) { 'Attention' } else { 'Healthy' }
        $note = if ($lifecycle -eq 'Healthy') { "Published servicing end: $endDate. This does not verify installed updates." } else { "Published servicing ended $endDate. Verify upgrade or applicable ESU coverage." }
        if ($family -eq 'Windows 10') { $note += ' ESU enrollment is not checked.' }
    }
    if ($family -eq 'Windows Server') { $note = 'Server detected; this client lifecycle matrix does not assess Server support.' }
    [pscustomobject]@{
        Family = $family; Edition = $edition; Release = $release; Build = $build
        Revision = Get-HealthProperty $CurrentVersion 'UBR'
        Architecture = Get-HealthProperty $OperatingSystem 'OSArchitecture'
        LifecycleStatus = $lifecycle; ServicingEnd = $endDate; LifecycleDetail = $note
        LifecycleSource = $source; LifecycleReviewedOn = '2026-09-23'
    }
}

function New-HealthCheck {
    param([string]$Id, [string]$Category, [string]$Name,
        [ValidateSet('Healthy', 'Attention', 'Unknown', 'Info')] [string]$Status,
        [string]$Value, [string]$Detail, [string]$Source)
    [pscustomobject]@{ Id = $Id; Category = $Category; Name = $Name; Status = $Status; Value = $Value; Detail = $Detail; Source = $Source }
}

function New-HealthBooleanCheck {
    param([string]$Id, [string]$Category, [string]$Name, $Probe, [string]$Property,
        [string]$Detail, [string]$Source)
    $value = if ($Probe.Success) { Get-HealthProperty $Probe.Value $Property } else { $null }
    if ($value -isnot [bool]) {
        $reason = if ($Probe.Success) { "Provider did not return a Boolean $Property value." } else { $Probe.Error }
        return New-HealthCheck $Id $Category $Name Unknown 'Unavailable' $reason $Source
    }
    $status = if ($value) { 'Healthy' } else { 'Attention' }
    $label = if ($value) { 'Enabled' } else { 'Disabled' }
    New-HealthCheck $Id $Category $Name $status $label $Detail $Source
}

function ConvertTo-WindowsHealthReport {
    param([Parameter(Mandatory)] [System.Collections.IDictionary]$Probes,
        [datetime]$AsOf = (Get-Date), [string]$ComputerName = $env:COMPUTERNAME)
    $identity = Get-WindowsHealthIdentity $Probes.OS.Value $Probes.Version.Value $AsOf
    $checks = New-Object 'System.Collections.Generic.List[object]'
    $checks.Add((New-HealthCheck 'os.lifecycle' Windows 'Windows servicing' $identity.LifecycleStatus "$($identity.Family) $($identity.Release) $($identity.Edition)" $identity.LifecycleDetail $identity.LifecycleSource))
    foreach ($providerName in @('OS', 'Version')) {
        if (-not $Probes[$providerName].Success) {
            $checks.Add((New-HealthCheck "os.$($providerName.ToLowerInvariant())" Windows "$providerName identification" Unknown Unavailable $Probes[$providerName].Error 'Windows inventory'))
        }
    }
    $defenderSource = 'Get-MpComputerStatus'
    $mode = if ($Probes.Defender.Success) { [string](Get-HealthProperty $Probes.Defender.Value 'AMRunningMode') } else { '' }
    $modeStatus = if ([string]::IsNullOrWhiteSpace($mode)) { 'Unknown' } else { 'Info' }
    $checks.Add((New-HealthCheck 'defender.mode' Defender 'Antivirus running mode' $modeStatus $mode 'Passive or EDR block mode needs context from the primary antivirus provider; this is not an antivirus compliance score.' $defenderSource))
    foreach ($field in @(
        @('realtime', 'Real-time protection', 'RealTimeProtectionEnabled'),
        @('behavior', 'Behavior monitoring', 'BehaviorMonitorEnabled'),
        @('tamper', 'Tamper protection', 'IsTamperProtected')
    )) {
        $check = New-HealthBooleanCheck "defender.$($field[0])" Defender $field[1] $Probes.Defender $field[2] 'Observed runtime state. Managed policy can differ from local preferences.' $defenderSource
        if ($check.Status -eq 'Attention' -and $mode -in @('Passive', 'SxS Passive Mode', 'EDR Block Mode') -and $field[0] -ne 'tamper') {
            $check.Status = 'Info'
            $check.Detail = "Defender mode: $mode. Verify protection with the primary antivirus provider."
        }
        $checks.Add($check)
    }
    $signature = Get-HealthProperty $Probes.Defender.Value 'AntivirusSignatureLastUpdated'
    $signatureTime = [datetime]::MinValue
    if ($Probes.Defender.Success -and $null -ne $signature -and [datetime]::TryParse([string]$signature, [ref]$signatureTime) -and $signatureTime.Year -ge 2000) {
        $age = ($AsOf.ToUniversalTime() - $signatureTime.ToUniversalTime()).TotalHours
        $status = if ($age -lt 0) { 'Unknown' } elseif ($age -gt 168) { 'Attention' } else { 'Healthy' }
        $checks.Add((New-HealthCheck 'defender.signatures' Defender 'Security intelligence age' $status ('{0:N1} hours' -f $age) 'Advisory threshold: 7 days. Future timestamps are unknown; check the system clock. This does not compare against the latest online definition.' $defenderSource))
    } else {
        $checks.Add((New-HealthCheck 'defender.signatures' Defender 'Security intelligence age' Unknown Unavailable 'No usable signature timestamp was returned.' $defenderSource))
    }
    $checks.Add((New-HealthBooleanCheck 'firmware.secureboot' Firmware 'Secure Boot' $Probes.SecureBoot Enabled 'Observed UEFI state. Access-denied and unsupported firmware are unknown, not disabled.' 'Confirm-SecureBootUEFI'))
    $checks.Add((New-HealthBooleanCheck 'firmware.tpm' Firmware 'TPM ready' $Probes.Tpm TpmReady 'TPM readiness does not certify TPM 2.0 or Windows 11 hardware eligibility.' 'Get-Tpm'))
    foreach ($service in @(@('credentialguard', 'Credential Guard', 1), @('memoryintegrity', 'Memory integrity (HVCI)', 2))) {
        $running = Get-HealthProperty $Probes.DeviceGuard.Value 'SecurityServicesRunning'
        $configured = Get-HealthProperty $Probes.DeviceGuard.Value 'SecurityServicesConfigured'
        $status = 'Unknown'; $value = 'Unavailable'; $detail = $Probes.DeviceGuard.Error
        if ($Probes.DeviceGuard.Success -and $null -ne $running -and @($running).Count -gt 0) {
            $active = @($running) -contains $service[2]
            $status = if ($active) { 'Healthy' } else { 'Attention' }
            $value = if ($active) { 'Running' } elseif (@($configured) -contains $service[2]) { 'Configured, not running' } else { 'Not running' }
            $detail = 'Runtime observation from Win32_DeviceGuard. Configured does not imply active; review hardware, policy and restart requirements.'
        } elseif ($Probes.DeviceGuard.Success) { $detail = 'The provider did not return SecurityServicesRunning.' }
        $checks.Add((New-HealthCheck "virtualization.$($service[0])" Virtualization $service[1] $status $value $detail 'Win32_DeviceGuard'))
    }
    $vbs = Get-HealthProperty $Probes.DeviceGuard.Value 'VirtualizationBasedSecurityStatus'
    $vbsStatus = 'Unknown'; $vbsValue = 'Unavailable'
    if ($Probes.DeviceGuard.Success -and $null -ne $vbs -and $vbs -in @(0, 1, 2)) {
        $vbsValue = @('Not enabled', 'Enabled, not running', 'Running')[[int]$vbs]
        $vbsStatus = if ($vbs -eq 2) { 'Healthy' } else { 'Attention' }
    }
    $checks.Add((New-HealthCheck 'virtualization.vbs' Virtualization 'Virtualization-based security' $vbsStatus $vbsValue 'Runtime state, not a registry-policy inference.' 'Win32_DeviceGuard'))
    foreach ($profileName in @('Domain', 'Private', 'Public')) {
        $profiles = @($Probes.Firewall.Value | Where-Object { [string](Get-HealthProperty $_ Name) -eq $profileName })
        $status = 'Unknown'; $value = 'Unavailable'; $detail = $Probes.Firewall.Error
        if ($Probes.Firewall.Success -and $profiles.Count -eq 1) {
            $enabled = [string](Get-HealthProperty $profiles[0] Enabled)
            if ($enabled -in @('True', '1')) { $status = 'Healthy'; $value = 'Enabled' }
            elseif ($enabled -in @('False', '0')) { $status = 'Attention'; $value = 'Disabled' }
            $detail = 'Effective profile configuration from ActiveStore; a profile need not be connected. Firewall rules are not assessed.'
        } elseif ($Probes.Firewall.Success) { $detail = 'The provider did not return exactly one matching profile.' }
        $checks.Add((New-HealthCheck "firewall.$($profileName.ToLowerInvariant())" Firewall "$profileName firewall" $status $value $detail 'Get-NetFirewallProfile -PolicyStore ActiveStore'))
    }
    $pending = @(Get-HealthProperty $Probes.Reboot.Value 'Reasons' | Where-Object { $null -ne $_ })
    $rebootStatus = 'Unknown'; $rebootValue = 'Unavailable'; $rebootDetail = $Probes.Reboot.Error
    if ($Probes.Reboot.Success) {
        $rebootStatus = if (@($pending).Count -gt 0) { 'Attention' } else { 'Info' }
        $rebootValue = if (@($pending).Count -gt 0) { 'Restart indicated' } else { 'No common markers' }
        $rebootDetail = 'Common CBS, Windows Update and pending file-rename markers only; not an exhaustive restart assessment. ' + (@($pending) -join ', ')
    }
    $checks.Add((New-HealthCheck 'windows.reboot' Windows 'Pending restart' $rebootStatus $rebootValue $rebootDetail 'Local registry'))
    [pscustomobject]@{
        SchemaVersion = 1; ReportType = 'WinDefState.Health'; ComputerName = $ComputerName
        CapturedAtUtc = $AsOf.ToUniversalTime().ToString('o'); Windows = $identity
        Summary = [pscustomobject]@{
            Healthy = @($checks | Where-Object Status -eq Healthy).Count
            Attention = @($checks | Where-Object Status -eq Attention).Count
            Unknown = @($checks | Where-Object Status -eq Unknown).Count
            Info = @($checks | Where-Object Status -eq Info).Count
        }
        Checks = @($checks.ToArray())
        Providers = @($Probes.Keys | Sort-Object | ForEach-Object {
            [pscustomobject]@{ Name = $_; Success = $Probes[$_].Success; DurationMs = $Probes[$_].DurationMs; Error = $Probes[$_].Error }
        })
    }
}

function Get-WindowsHealthReport {
    [CmdletBinding()]
    param()
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Health capture requires Windows.' }
    if (-not [Environment]::Is64BitProcess -and [Environment]::Is64BitOperatingSystem) { throw 'Run health capture in native 64-bit Windows PowerShell.' }
    $probes = [ordered]@{}
    $probes.OS = Invoke-HealthProbe { Get-CimInstance Win32_OperatingSystem -OperationTimeoutSec 8 -ErrorAction Stop }
    $probes.Version = Invoke-HealthProbe { Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop }
    $probes.Defender = Invoke-HealthProbe { Get-MpComputerStatus -ErrorAction Stop }
    $probes.SecureBoot = Invoke-HealthProbe { [pscustomobject]@{ Enabled = Confirm-SecureBootUEFI -ErrorAction Stop } }
    $probes.Tpm = Invoke-HealthProbe { Get-Tpm -ErrorAction Stop }
    $probes.DeviceGuard = Invoke-HealthProbe { Get-CimInstance -Namespace root\Microsoft\Windows\DeviceGuard -ClassName Win32_DeviceGuard -OperationTimeoutSec 8 -ErrorAction Stop }
    $probes.Firewall = Invoke-HealthProbe { Get-NetFirewallProfile -PolicyStore ActiveStore -ErrorAction Stop }
    $probes.Reboot = Invoke-HealthProbe {
        $reasons = @(
            foreach ($key in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending', 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')) {
                if (Test-Path -LiteralPath $key -ErrorAction Stop) { Split-Path $key -Leaf }
            }
            $sessionManager = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -ErrorAction Stop
            if (@(Get-HealthProperty $sessionManager PendingFileRenameOperations).Count -gt 0) { 'PendingFileRenameOperations' }
        )
        [pscustomobject]@{ Reasons = $reasons }
    }
    ConvertTo-WindowsHealthReport -Probes $probes
}

function ConvertTo-HealthHtml {
    param([Parameter(Mandatory)]$Report)
    $rows = foreach ($check in $Report.Checks) {
        $cells = foreach ($property in @('Status', 'Category', 'Name', 'Value', 'Detail', 'Source')) {
            '<td>' + [Net.WebUtility]::HtmlEncode([string]$check.$property) + '</td>'
        }
        '<tr>' + ($cells -join '') + '</tr>'
    }
    $hostLabel = [Net.WebUtility]::HtmlEncode([string]$Report.ComputerName)
    $captured = [Net.WebUtility]::HtmlEncode([string]$Report.CapturedAtUtc)
    $osLabel = [Net.WebUtility]::HtmlEncode(('{0} {1} / {2} / build {3}.{4}' -f $Report.Windows.Family, $Report.Windows.Release, $Report.Windows.Edition, $Report.Windows.Build, $Report.Windows.Revision))
    $cards = foreach ($status in @('Healthy', 'Attention', 'Unknown', 'Info')) {
        '<div class="card"><strong>' + [int]$Report.Summary.$status + '</strong><span>' + $status + '</span></div>'
    }
    @"
<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>WinDefState health - $hostLabel</title><style>
:root{color-scheme:light}*{box-sizing:border-box}body{font:15px/1.55 system-ui,Segoe UI,sans-serif;margin:0;background:#f3f6fa;color:#172c3e}main{max-width:1440px;margin:auto;padding:40px 28px}header{border-left:5px solid #147b74;padding-left:22px}h1{font-size:34px;margin:6px 0}p{color:#425870}small{letter-spacing:2px;color:#147b74;font-weight:700}.cards{display:flex;gap:16px;flex-wrap:wrap;margin:28px 0}.card{flex:1;min-width:140px;background:white;border:1px solid #d4dfe8;border-radius:12px;padding:18px}.card strong{display:block;font-size:32px}.card span{color:#425870}label{display:block;font-weight:600}input,select{font:inherit;padding:10px;border:1px solid #8ca0b3;border-radius:6px;margin:6px 12px 18px 0}input{width:min(420px,100%)}.table{overflow:auto;border:1px solid #d4dfe8;border-radius:10px}table{border-collapse:collapse;width:100%;background:white;text-align:left}th{background:#e7edf4;font-size:12px;text-transform:uppercase;letter-spacing:1px}th,td{padding:14px;border-bottom:1px solid #e1e7ef;vertical-align:top}td:first-child{font-weight:700}tr[data-status=Attention] td:first-child{color:#8c4700}tr[data-status=Healthy] td:first-child{color:#08705c}tr[data-status=Unknown] td:first-child{color:#5d4bb0}footer{margin-top:20px;color:#425870}@media print{input,select,label{display:none}main{padding:0}.table{overflow:visible}tr{break-inside:avoid}}
</style></head><body><main><header><small>WinDefState</small><h1>Security health</h1><p>$hostLabel &middot; $osLabel<br>Captured $captured</p></header>
<div class="cards">$($cards -join '')</div><label for="search">Find a check</label><input id="search" type="search" placeholder="Search name, status or evidence"><select id="status" aria-label="Filter by status"><option>All statuses</option><option>Attention</option><option>Unknown</option><option>Healthy</option><option>Info</option></select><span id="count" role="status"></span>
<div class="table"><table><thead><tr><th>Status</th><th>Category</th><th>Check</th><th>Observed value</th><th>Context</th><th>Source</th></tr></thead><tbody>$($rows -join '')</tbody></table></div>
<footer>Read-only inventory. Unknown means a check could not be established. This report is not a restore snapshot, compliance certification, or proof that a device is fully protected. Lifecycle metadata reviewed $([Net.WebUtility]::HtmlEncode([string]$Report.Windows.LifecycleReviewedOn)).</footer>
</main><script>
const rows=Array.from(document.querySelectorAll('tbody tr')),search=document.querySelector('#search'),status=document.querySelector('#status');
rows.forEach(row=>row.dataset.status=row.cells[0].textContent);
function filter(){const tokens=search.value.toLowerCase().trim().split(/\s+/).filter(Boolean);let shown=0;rows.forEach(row=>{row.hidden=!tokens.every(t=>row.textContent.toLowerCase().includes(t))||(status.selectedIndex>0&&row.dataset.status!==status.value);if(!row.hidden)shown++});document.querySelector('#count').textContent=shown+' of '+rows.length+' checks'}
search.addEventListener('input',filter);status.addEventListener('change',filter);filter();
</script></body></html>
"@
}

function Export-WindowsHealthReport {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Report, [Parameter(Mandatory)][string]$Path,
        [ValidateSet('Json', 'Html')][string]$Format = 'Json')
    $content = if ($Format -eq 'Html') { ConvertTo-HealthHtml $Report } else { $Report | ConvertTo-Json -Depth 12 }
    $fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    # CreateNew is race-safe and will never clobber a snapshot or existing report.
    $stream = [IO.File]::Open($fullPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    try {
        $bytes = [Text.UTF8Encoding]::new($false).GetBytes($content)
        $stream.Write($bytes, 0, $bytes.Length)
    } finally { $stream.Dispose() }
    $fullPath
}

if ($MyInvocation.InvocationName -ne '.') {
    $ErrorActionPreference = 'Stop'
    $report = Get-WindowsHealthReport
    if (-not [string]::IsNullOrWhiteSpace($HealthOutputPath)) {
        $null = Export-WindowsHealthReport -Report $report -Path $HealthOutputPath -Format $HealthFormat
    }
    $report
}
